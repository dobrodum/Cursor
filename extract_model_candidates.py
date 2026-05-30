#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable paths
# ----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


N_QUARTERS = 10

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

MONTH_LOOKUP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

PERIOD_DAY_LOOKUP = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def as_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    if values is None:
        return []
    return [[values]]


def safe_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def compact_signature(values: Sequence[Optional[float]]) -> Tuple[Optional[float], ...]:
    packed: List[Optional[float]] = []
    for value in values:
        if value is None:
            packed.append(None)
        else:
            packed.append(round(value, 8))
    return tuple(packed)


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = parts[-2] if len(parts) >= 2 else "UNKNOWN"
    period_token = parts[-1] if parts else stem
    period_token = re.sub(r"[_\-\s]*send$", "", period_token, flags=re.IGNORECASE)

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    if period_match:
        phase = period_match.group(1).title()
        month_name = period_match.group(2).lower()
        year = int(period_match.group(3))
        month_num = MONTH_LOOKUP.get(month_name)

        if month_num:
            month_abbrev = date(year, month_num, 1).strftime("%b")
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = date(year, month_num, PERIOD_DAY_LOOKUP[phase]).isoformat()
        else:
            model_period = period_token
            model_date = ""
    else:
        model_period = period_token
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder_name: str, out_dir: Path) -> Path:
    base_name = f"{input_folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_workbook_safe(wb: xw.Book) -> None:
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
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    best: Optional[Tuple[Tuple[int, int], Tuple[int, int]]] = None

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            abs_row = start_row + r_idx
            abs_col = start_col + c_idx

            below_is_min = 0
            if r_idx + 1 < len(values):
                row_below = values[r_idx + 1]
                if c_idx < len(row_below) and normalize_text(row_below[c_idx]) == "min":
                    below_is_min = 1
            score = (below_is_min, abs_row)

            if best is None or score > best[0]:
                best = (score, (abs_row, abs_col))

    return best[1] if best else None


def build_label_index(values: List[List[Any]], start_row: int, start_col: int) -> Dict[str, Tuple[int, int]]:
    labels: Dict[str, Tuple[int, int]] = {}
    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            text = normalize_text(value)
            if not text:
                continue
            labels.setdefault(text, (start_row + r_idx, start_col + c_idx))
    return labels


def find_label_position(
    labels: Dict[str, Tuple[int, int]],
    search_terms: Iterable[str],
) -> Optional[Tuple[int, int]]:
    normalized_terms = [normalize_text(term) for term in search_terms if normalize_text(term)]
    for term in normalized_terms:
        if term in labels:
            return labels[term]

    for term in normalized_terms:
        for label, position in labels.items():
            if term in label:
                return position
    return None


def read_column_values(sheet: xw.Sheet, start_row: int, end_row: int, col: int) -> List[Any]:
    if end_row < start_row:
        return []
    values = sheet.range((start_row, col), (end_row, col)).value
    if isinstance(values, list):
        return values
    return [values]


def map_rows_to_values(start_row: int, column_values: Sequence[Any]) -> Dict[int, Any]:
    return {start_row + idx: value for idx, value in enumerate(column_values)}


def locate_min_row(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> int:
    for offset in range(1, 8):
        row = anchor_row + offset
        if normalize_text(sheet.range((row, anchor_col)).value) == "min":
            return row
    return anchor_row + 1


def locate_avg_pen_target(
    sheet: xw.Sheet,
    labels: Dict[str, Tuple[int, int]],
    anchor_row: int,
    anchor_col: int,
) -> xw.Range:
    label_pos = find_label_position(labels, ["avg penetration", "avg_penetration_pct", "avg penetration pct"])
    if label_pos:
        return sheet.range((label_pos[0], label_pos[1] + 1))
    return sheet.range((anchor_row, anchor_col + 6))


def detect_empirical_columns(
    values: List[List[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    anchor_col: int,
) -> Dict[str, int]:
    search_terms = {
        "quarter": ("quarter", "qtr"),
        "avg_penetration_pct": ("penetration",),
        "quarterly_sales": ("quarterly sales", "quarter sales", "sales"),
        "reported_sales": ("reported sales", "reported"),
        "growth_rate_pct": ("growth",),
        "sales_captured_in_db_pct": ("captured", "db"),
    }

    best_row_idx: Optional[int] = None
    best_score = -1
    scan_start = max(start_row, anchor_row - 40)
    scan_end = anchor_row - 1

    for abs_row in range(scan_start, scan_end + 1):
        idx = abs_row - start_row
        if idx < 0 or idx >= len(values):
            continue
        row = values[idx]
        score = 0
        for cell in row:
            text = normalize_text(cell)
            if not text:
                continue
            if any(term in text for terms in search_terms.values() for term in terms):
                score += 1
        if score > best_score:
            best_score = score
            best_row_idx = idx

    col_map: Dict[str, int] = {}
    if best_row_idx is not None and best_score > 0:
        row = values[best_row_idx]
        for c_idx, cell in enumerate(row):
            text = normalize_text(cell)
            if not text:
                continue
            abs_col = start_col + c_idx
            for key, terms in search_terms.items():
                if key in col_map:
                    continue
                if any(term in text for term in terms):
                    col_map[key] = abs_col

    # Anchor-based fallbacks for robustness if headers are missing
    col_map.setdefault("quarter", max(1, anchor_col - 13))
    col_map.setdefault("avg_penetration_pct", max(1, anchor_col - 10))
    col_map.setdefault("quarterly_sales", max(1, anchor_col - 9))
    col_map.setdefault("reported_sales", max(1, anchor_col - 8))
    col_map.setdefault("growth_rate_pct", max(1, anchor_col - 7))
    col_map.setdefault("sales_captured_in_db_pct", max(1, anchor_col - 6))
    return col_map


def process_empirical_sheet(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped file: {source_file} (missing sheet: Empirical Model)")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipped file: {source_file} (Empirical Model missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    values = as_2d(used.value)
    start_row = used.row
    start_col = used.column
    labels = build_label_index(values, start_row, start_col)

    min_row = locate_min_row(sheet, anchor_row, anchor_col)
    max_value_cell = sheet.range((anchor_row, anchor_col + 1))
    min_value_cell = sheet.range((min_row, anchor_col + 1))

    col_map = detect_empirical_columns(values, start_row, start_col, anchor_row, anchor_col)

    data_start = start_row
    data_end = anchor_row - 1
    if data_end < data_start:
        return []

    penetration_vals = read_column_values(sheet, data_start, data_end, col_map["avg_penetration_pct"])
    penetration_by_row = map_rows_to_values(data_start, penetration_vals)
    valid_rows = [row for row, value in penetration_by_row.items() if safe_float(value) is not None]
    if not valid_rows:
        return []

    valid_rows.sort()
    max_quarters = min(N_QUARTERS, len(valid_rows))

    quarter_vals = map_rows_to_values(data_start, read_column_values(sheet, data_start, data_end, col_map["quarter"]))
    quarterly_sales_vals = map_rows_to_values(
        data_start, read_column_values(sheet, data_start, data_end, col_map["quarterly_sales"])
    )
    reported_sales_vals = map_rows_to_values(
        data_start, read_column_values(sheet, data_start, data_end, col_map["reported_sales"])
    )
    growth_vals = map_rows_to_values(data_start, read_column_values(sheet, data_start, data_end, col_map["growth_rate_pct"]))
    captured_vals = map_rows_to_values(
        data_start, read_column_values(sheet, data_start, data_end, col_map["sales_captured_in_db_pct"])
    )

    est_total_pos = find_label_position(labels, ["estimated total sold", "est total sold", "estimated total"])
    est_total_cell = sheet.range((est_total_pos[0], est_total_pos[1] + 1)) if est_total_pos else None

    avg_pen_target = locate_avg_pen_target(sheet, labels, anchor_row, anchor_col)
    original_formula = avg_pen_target.formula2
    original_value = avg_pen_target.value

    rows: List[Dict[str, Any]] = []
    try:
        for quarters_used in range(1, max_quarters + 1):
            window_rows = valid_rows[-quarters_used:]
            window_start = window_rows[0]
            window_end = window_rows[-1]

            avg_formula = (
                f"=AVERAGE(R{window_start}C{col_map['avg_penetration_pct']}:"
                f"R{window_end}C{col_map['avg_penetration_pct']})"
            )
            avg_pen_target.formula2 = avg_formula
            wb.app.calculate()

            avg_penetration_pct = safe_float(avg_pen_target.value)
            forecast_max = safe_float(max_value_cell.value)
            forecast_min = safe_float(min_value_cell.value)

            quarter_label = quarter_vals.get(window_start)
            quarterly_sales = safe_float(quarterly_sales_vals.get(window_end))
            reported_sales = safe_float(reported_sales_vals.get(window_end))
            growth_rate_pct = safe_float(growth_vals.get(window_end))
            sales_captured_in_db_pct = safe_float(captured_vals.get(window_end))

            forecast_value = safe_float(est_total_cell.value) if est_total_cell else None
            if forecast_value is None and avg_penetration_pct not in (None, 0) and quarterly_sales is not None:
                penetration_decimal = avg_penetration_pct / 100.0 if abs(avg_penetration_pct) > 1 else avg_penetration_pct
                if penetration_decimal:
                    forecast_value = quarterly_sales / penetration_decimal
            if forecast_value is None and forecast_max is not None and forecast_min is not None:
                forecast_value = (forecast_max + forecast_min) / 2.0

            range_width: Optional[float] = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            rows.append(
                {
                    "model": file_label.model,
                    "ticker": file_label.ticker,
                    "model_period": file_label.model_period,
                    "model_date": file_label.model_date,
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_penetration_pct,
                    "num_quarters_used": quarters_used,
                    "last_quarter_used": quarter_label,
                    "forecast_value": forecast_value,
                    "actual_value": reported_sales,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "avg_penetration_pct": avg_penetration_pct,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate_pct,
                    "sales_captured_in_db_pct": sales_captured_in_db_pct,
                    "source_file": source_file,
                }
            )
    finally:
        try:
            if original_formula:
                avg_pen_target.formula2 = original_formula
            else:
                avg_pen_target.value = original_value
            wb.app.calculate()
        except Exception:
            pass

    return rows


def process_regression_sheet(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped file: {source_file} (missing sheet: Regression Model)")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipped file: {source_file} (Regression Model missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    used = sheet.used_range
    values = as_2d(used.value)
    labels = build_label_index(values, used.row, used.column)

    min_row = locate_min_row(sheet, anchor_row, anchor_col)
    max_value_cell = sheet.range((anchor_row, anchor_col + 1))
    min_value_cell = sheet.range((min_row, anchor_col + 1))

    data_start = used.row
    data_end = anchor_row - 1
    if data_end < data_start:
        return []

    x_vals = read_column_values(sheet, data_start, data_end, x_col)
    y_vals = read_column_values(sheet, data_start, data_end, y_col)
    x_by_row = map_rows_to_values(data_start, x_vals)
    y_by_row = map_rows_to_values(data_start, y_vals)

    valid_rows = [
        row
        for row in range(data_start, data_end + 1)
        if safe_float(x_by_row.get(row)) is not None and safe_float(y_by_row.get(row)) is not None
    ]
    if not valid_rows:
        return []

    valid_rows.sort()
    max_quarters = min(N_QUARTERS, len(valid_rows))

    intercept_pos = find_label_position(labels, ["intercept"])
    slope_pos = find_label_position(labels, ["slope"])
    forecast_pos = find_label_position(labels, ["tot fcst w/o sa", "tot fcst without sa", "tot fcst wo sa"])
    actual_pos = find_label_position(labels, ["actual value", "reported sales", "actual"])

    intercept_cell = (
        sheet.range((intercept_pos[0], intercept_pos[1] + 1))
        if intercept_pos
        else sheet.range((anchor_row, anchor_col + 6))
    )
    slope_cell = (
        sheet.range((slope_pos[0], slope_pos[1] + 1))
        if slope_pos
        else sheet.range((anchor_row + 1, anchor_col + 6))
    )
    forecast_cell = sheet.range((forecast_pos[0], forecast_pos[1] + 1)) if forecast_pos else None
    actual_cell = sheet.range((actual_pos[0], actual_pos[1] + 1)) if actual_pos else None

    original_intercept_formula = intercept_cell.formula2
    original_intercept_value = intercept_cell.value
    original_slope_formula = slope_cell.formula2
    original_slope_value = slope_cell.value

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    try:
        for quarters_used in range(1, max_quarters + 1):
            window_rows = valid_rows[-quarters_used:]
            window_start = window_rows[0]
            window_end = window_rows[-1]

            intercept_cell.formula2 = (
                f"=INTERCEPT(R{window_start}C{y_col}:R{window_end}C{y_col},"
                f"R{window_start}C{x_col}:R{window_end}C{x_col})"
            )
            slope_cell.formula2 = (
                f"=SLOPE(R{window_start}C{y_col}:R{window_end}C{y_col},"
                f"R{window_start}C{x_col}:R{window_end}C{x_col})"
            )
            wb.app.calculate()

            intercept = safe_float(intercept_cell.value)
            slope = safe_float(slope_cell.value)
            forecast_total_without_sa = safe_float(forecast_cell.value) if forecast_cell else None
            if forecast_total_without_sa is None and intercept is not None and slope is not None:
                x_latest = safe_float(x_by_row.get(window_end))
                if x_latest is not None:
                    forecast_total_without_sa = intercept + slope * x_latest

            actual_value = safe_float(actual_cell.value) if actual_cell else None
            forecast_max = safe_float(max_value_cell.value)
            forecast_min = safe_float(min_value_cell.value)

            range_width: Optional[float] = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            signature = compact_signature(
                [forecast_total_without_sa, forecast_max, forecast_min, intercept, slope]
            )
            if previous_signature is not None and signature == previous_signature:
                continue
            previous_signature = signature

            rows.append(
                {
                    "model": file_label.model,
                    "ticker": file_label.ticker,
                    "model_period": file_label.model_period,
                    "model_date": file_label.model_date,
                    "method": "regression",
                    "parameter_name": "num_quarters_used",
                    "parameter_value": quarters_used,
                    "num_quarters_used": quarters_used,
                    "forecast_value": forecast_total_without_sa,
                    "actual_value": actual_value,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "intercept": intercept,
                    "slope": slope,
                    "source_file": source_file,
                }
            )
    finally:
        try:
            if original_intercept_formula:
                intercept_cell.formula2 = original_intercept_formula
            else:
                intercept_cell.value = original_intercept_value

            if original_slope_formula:
                slope_cell.formula2 = original_slope_formula
            else:
                slope_cell.value = original_slope_value
            wb.app.calculate()
        except Exception:
            pass

    return rows


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for item in rows:
        ws.append([item.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        max_scan_row = min(ws.max_row, 5000)
        for row in range(2, max_scan_row + 1):
            value = ws.cell(row=row, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def get_input_files(folder: Path) -> List[Path]:
    files: List[Path] = []
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {file_path.name} (not .xlsx)")
            continue
        files.append(file_path)
    return files


def main() -> None:
    source_folder = Path(input_dir)
    destination_folder = Path(output_dir)

    if not source_folder.exists():
        raise FileNotFoundError(f"Input folder does not exist: {source_folder}")

    destination_folder.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(source_folder.name, destination_folder)

    source_files = get_input_files(source_folder)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing file: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path.name)

                empirical_rows.extend(process_empirical_sheet(wb, label, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, label, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
