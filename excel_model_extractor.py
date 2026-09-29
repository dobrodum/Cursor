from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Update these two paths before running.
input_dir = "/path/to/input"
output_dir = "/path/to/output"


EMPIRICAL_COLUMNS = [
    "model",
    "ticker",
    "model_period",
    "model_date",
    "method",
    "parameter_name",
    "parameter_value",
    "num_quarters_used",
    "last_quarter_used",
    "forecast_value",
    "actual_value",
    "forecast_max",
    "forecast_min",
    "range_width",
    "avg_penetration_pct",
    "quarterly_sales",
    "reported_sales",
    "growth_rate_pct",
    "sales_captured_in_db_pct",
    "source_file",
]

REGRESSION_COLUMNS = [
    "model",
    "ticker",
    "model_period",
    "model_date",
    "method",
    "parameter_name",
    "parameter_value",
    "num_quarters_used",
    "forecast_value",
    "actual_value",
    "forecast_max",
    "forecast_min",
    "range_width",
    "intercept",
    "slope",
    "source_file",
]

DAY_BY_PHASE = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_BY_NAME = {name.lower(): idx for idx, name in enumerate(calendar.month_abbr) if name}
MONTH_BY_NAME.update({name.lower(): idx for idx, name in enumerate(calendar.month_name) if name})


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    lowered = text.lower()
    if lowered in {"na", "n/a", "none", "#n/a"}:
        return None
    is_percent = text.endswith("%")
    if is_percent:
        text = text[:-1].strip()
    try:
        number = float(text)
        return number / 100.0 if is_percent else number
    except ValueError:
        return None


def rounded_signature(values: Iterable[Optional[float]]) -> Tuple[Optional[float], ...]:
    sig: List[Optional[float]] = []
    for value in values:
        if value is None:
            sig.append(None)
        else:
            sig.append(round(value, 10))
    return tuple(sig)


def matrix_value(values: List[List[Any]], top_row: int, left_col: int, row: int, col: int) -> Any:
    r_idx = row - top_row
    c_idx = col - left_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(values):
        return None
    row_values = values[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def parse_month(month_text: str) -> Optional[int]:
    cleaned = re.sub(r"[^A-Za-z]", "", month_text).lower()
    if not cleaned:
        return None
    if cleaned in MONTH_BY_NAME:
        return MONTH_BY_NAME[cleaned]
    if cleaned[:3] in MONTH_BY_NAME:
        return MONTH_BY_NAME[cleaned[:3]]
    return None


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 and parts[1] else "UNKNOWN"

    period_segment = parts[2] if len(parts) >= 3 else stem
    period_token = period_segment.split("_")[0].strip()
    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, flags=re.IGNORECASE)

    if not match:
        model_period = "Unknown_Period"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return {
            "model": model,
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
        }

    phase = match.group(1).title()
    month_text = match.group(2)
    year = int(match.group(3))
    month = parse_month(month_text)
    if month is None:
        model_period = f"{phase}{month_text}_{year}"
        model_date = ""
    else:
        month_label = date(year, month, 1).strftime("%b")
        model_period = f"{phase}{month_label}_{year}"
        model_date = date(year, month, DAY_BY_PHASE[phase]).isoformat()
    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def output_path_for_run(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
    except Exception:
        pass

    try:
        wb.saved = True
        wb.close()
    except Exception as exc:
        print(f"  Warning: could not close workbook safely: {exc}")


def scan_sheet(sheet: xw.Sheet) -> Tuple[Any, List[List[Any]], int, int, List[Tuple[int, int, str]], Optional[Tuple[int, int]]]:
    used = sheet.used_range
    values = as_2d(used.value)
    top_row = used.row
    left_col = used.column
    labels: List[Tuple[int, int, str]] = []
    anchor: Optional[Tuple[int, int]] = None

    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            if not isinstance(value, str):
                continue
            text = value.strip().lower()
            if not text:
                continue
            row = top_row + r_idx
            col = left_col + c_idx
            labels.append((row, col, text))
            if anchor is None and text == "max":
                anchor = (row, col)
    return used, values, top_row, left_col, labels, anchor


def find_row_by_keywords(
    labels: Sequence[Tuple[int, int, str]], keywords: Sequence[str], near_col: Optional[int] = None
) -> Optional[int]:
    candidates: List[Tuple[int, int]] = []
    for row, col, text in labels:
        if all(keyword in text for keyword in keywords):
            candidates.append((row, col))
    if not candidates:
        return None
    if near_col is None:
        return candidates[0][0]
    candidates.sort(key=lambda item: abs(item[1] - near_col))
    return candidates[0][0]


def find_min_row(labels: Sequence[Tuple[int, int, str]], anchor_col: int, anchor_row: int) -> Optional[int]:
    exact_col_matches = [row for row, col, text in labels if text == "min" and col == anchor_col]
    if exact_col_matches:
        return min(exact_col_matches, key=lambda row: abs(row - anchor_row))
    nearby_matches = [(row, col) for row, col, text in labels if text == "min"]
    if not nearby_matches:
        return None
    nearby_matches.sort(key=lambda item: (abs(item[1] - anchor_col), abs(item[0] - anchor_row)))
    return nearby_matches[0][0]


def value_from_sheet_or_cache(
    sheet: xw.Sheet, values: List[List[Any]], top_row: int, left_col: int, row: int, col: int
) -> Any:
    cached = matrix_value(values, top_row, left_col, row, col)
    if cached is not None:
        return cached
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def extract_empirical_rows(
    wb: xw.Book, sheet: xw.Sheet, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    used, values, top_row, left_col, labels, anchor = scan_sheet(sheet)
    if anchor is None:
        print("  Empirical Model: skipped (no 'max' anchor found)")
        return rows

    anchor_row, anchor_col = anchor
    min_row = find_min_row(labels, anchor_col=anchor_col, anchor_row=anchor_row) or (anchor_row + 1)

    penetration_row = find_row_by_keywords(labels, ["penetration"], near_col=anchor_col) or (anchor_row - 5)
    quarterly_sales_row = find_row_by_keywords(labels, ["quarterly", "sales"], near_col=anchor_col) or (anchor_row - 4)
    reported_sales_row = find_row_by_keywords(labels, ["reported", "sales"], near_col=anchor_col) or (anchor_row - 3)
    growth_rate_row = find_row_by_keywords(labels, ["growth", "rate"], near_col=anchor_col) or (anchor_row - 2)
    captured_db_row = (
        find_row_by_keywords(labels, ["captured", "db"], near_col=anchor_col)
        or find_row_by_keywords(labels, ["captured", "database"], near_col=anchor_col)
        or (anchor_row - 1)
    )
    quarter_label_row = find_row_by_keywords(labels, ["quarter"], near_col=anchor_col) or (penetration_row - 1)

    temp_row = used.last_cell.row + 3
    temp_col = used.last_cell.column + 2
    avg_pen_cell = sheet.range((temp_row, temp_col))
    max_pen_cell = sheet.range((temp_row, temp_col + 1))
    min_pen_cell = sheet.range((temp_row, temp_col + 2))

    n_quarters = 10
    end_col = anchor_col - 1

    for n in range(1, n_quarters + 1):
        start_col = end_col - n + 1
        if start_col < 1:
            continue

        avg_pen_cell.formula2 = f"=AVERAGE(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
        max_pen_cell.formula2 = f"=MAX(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
        min_pen_cell.formula2 = f"=MIN(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
        wb.app.calculate()

        avg_penetration_pct = to_float(avg_pen_cell.value)
        max_pen = to_float(max_pen_cell.value)
        min_pen = to_float(min_pen_cell.value)

        quarterly_sales = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, quarterly_sales_row, end_col))
        reported_sales = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, reported_sales_row, end_col))
        growth_rate_pct = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, growth_rate_row, end_col))
        sales_captured_pct = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, captured_db_row, end_col))
        last_quarter_used = value_from_sheet_or_cache(sheet, values, top_row, left_col, quarter_label_row, end_col)

        forecast_value = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, anchor_row - 1, anchor_col + n))
        forecast_max = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, anchor_row, anchor_col + n))
        forecast_min = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, min_row, anchor_col + n))

        if forecast_value is None and quarterly_sales is not None and avg_penetration_pct not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration_pct
        if forecast_max is None and quarterly_sales is not None and min_pen not in (None, 0):
            forecast_max = quarterly_sales / min_pen
        if forecast_min is None and quarterly_sales is not None and max_pen not in (None, 0):
            forecast_min = quarterly_sales / max_pen

        if all(
            item is None
            for item in (
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
                forecast_value,
                forecast_max,
                forecast_min,
            )
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    avg_pen_cell.value = None
    max_pen_cell.value = None
    min_pen_cell.value = None
    return rows


def extract_regression_rows(
    wb: xw.Book, sheet: xw.Sheet, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    used, values, top_row, left_col, labels, anchor = scan_sheet(sheet)
    if anchor is None:
        print("  Regression Model: skipped (no 'max' anchor found)")
        return rows

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print("  Regression Model: skipped (anchor offsets out of bounds)")
        return rows

    data_end_row = anchor_row - 1
    points: List[Tuple[int, float, float]] = []
    for row in range(top_row, data_end_row + 1):
        x_val = to_float(matrix_value(values, top_row, left_col, row, x_col))
        y_val = to_float(matrix_value(values, top_row, left_col, row, y_col))
        if x_val is None or y_val is None:
            continue
        points.append((row, x_val, y_val))

    if len(points) < 2:
        print("  Regression Model: skipped (not enough numeric points)")
        return rows

    actual_sales_row = find_row_by_keywords(labels, ["reported", "sales"], near_col=anchor_col)
    actual_value = None
    if actual_sales_row is not None:
        actual_value = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, actual_sales_row, anchor_col + 1))

    temp_row = used.last_cell.row + 3
    temp_col = used.last_cell.column + 2
    intercept_cell = sheet.range((temp_row, temp_col))
    slope_cell = sheet.range((temp_row, temp_col + 1))
    max_cell = sheet.range((temp_row, temp_col + 2))
    min_cell = sheet.range((temp_row, temp_col + 3))

    previous_signature: Optional[Tuple[Optional[float], ...]] = None
    max_quarters = min(10, len(points))
    for n in range(2, max_quarters + 1):
        subset = points[-n:]
        start_row = subset[0][0]
        end_row = subset[-1][0]

        y_range = f"R{start_row}C{y_col}:R{end_row}C{y_col}"
        x_range = f"R{start_row}C{x_col}:R{end_row}C{x_col}"

        intercept_cell.formula2 = f"=INTERCEPT({y_range},{x_range})"
        slope_cell.formula2 = f"=SLOPE({y_range},{x_range})"
        max_cell.formula2 = f"=MAX({y_range})"
        min_cell.formula2 = f"=MIN({y_range})"
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        next_x = subset[-1][1] + 1.0
        forecast_value = None
        if intercept is not None and slope is not None:
            forecast_value = intercept + (slope * next_x)

        sheet_forecast = to_float(value_from_sheet_or_cache(sheet, values, top_row, left_col, anchor_row - 1, anchor_col + n - 1))
        if sheet_forecast is not None:
            forecast_value = sheet_forecast

        signature = rounded_signature([forecast_value, forecast_max, forecast_min, intercept, slope])
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    intercept_cell.value = None
    slope_cell.value = None
    max_cell.value = None
    min_cell.value = None
    return rows


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    output_file: Path, empirical_rows: Sequence[Dict[str, Any]], regression_rows: Sequence[Dict[str, Any]]
) -> None:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_file)


def process_workbook(
    app: xw.App, file_path: Path, metadata: Dict[str, str]
) -> Tuple[bool, List[Dict[str, Any]], List[Dict[str, Any]]]:
    workbook_opened = False
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    wb: Optional[xw.Book] = None

    try:
        wb = app.books.open(str(file_path), update_links=False)
        workbook_opened = True
        sheet_by_name = {sheet.name.strip().lower(): sheet for sheet in wb.sheets}

        empirical_sheet = sheet_by_name.get("empirical model")
        regression_sheet = sheet_by_name.get("regression model")

        if empirical_sheet is None:
            print("  Empirical Model: skipped (sheet not found)")
        else:
            empirical_rows.extend(extract_empirical_rows(wb, empirical_sheet, metadata, file_path.name))

        if regression_sheet is None:
            print("  Regression Model: skipped (sheet not found)")
        else:
            regression_rows.extend(extract_regression_rows(wb, regression_sheet, metadata, file_path.name))
    except Exception as exc:
        print(f"  Failed while processing workbook: {exc}")
    finally:
        if wb is not None:
            close_source_workbook(wb)

    return workbook_opened, empirical_rows, regression_rows


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file = output_path_for_run(input_path, output_path)
    input_folder_prefix = f"{input_path.name}_PARAM"

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.enable_events = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda item: item.name.lower()):
            if not file_path.is_file():
                print(f"Skipping {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipping {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipping {file_path.name}: not an .xlsx file")
                continue
            if file_path.stem.startswith(input_folder_prefix):
                print(f"Skipping {file_path.name}: prior output workbook")
                continue

            print(f"Processing {file_path.name}")
            metadata = parse_file_labels(file_path)
            opened, emp_rows, reg_rows = process_workbook(app, file_path, metadata)
            if opened:
                processed_files += 1
                empirical_rows.extend(emp_rows)
                regression_rows.extend(reg_rows)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
