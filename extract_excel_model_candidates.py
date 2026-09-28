#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw

# Configure these before running.
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text.endswith("%"):
            text = text[:-1]
        try:
            float(text)
            return True
        except ValueError:
            return False
    return False


def as_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        pct = text.endswith("%")
        if pct:
            text = text[:-1]
        try:
            number = float(text)
            return number / 100.0 if pct else number
        except ValueError:
            return None
    return None


def parse_file_metadata(file_name: str) -> Optional[Dict[str, str]]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].upper()
    period_token = re.split(r"[_\s]", parts[2])[0]
    token_match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$", period_token, flags=re.IGNORECASE)
    if not token_match:
        return None

    period_bucket_raw, month_abbrev_raw, year_str = token_match.groups()
    period_bucket = period_bucket_raw.title()
    month_abbrev = month_abbrev_raw.title()
    year = int(year_str)

    try:
        month_num = datetime.strptime(month_abbrev, "%b").month
    except ValueError:
        return None

    day_by_bucket = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_by_bucket[period_bucket]
    model_date = date(year, month_num, model_day).isoformat()
    model_period = f"{period_bucket}{month_abbrev}_{year}"
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(out_dir: Path, input_folder_name: str) -> Path:
    base_name = f"{input_folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def safe_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def find_text_anchor(sheet: xw.Sheet, text: str) -> Optional[Tuple[int, int]]:
    try:
        found = sheet.api.UsedRange.Find(What=text, LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    try:
        used = sheet.used_range
        values = used.value
        start_row = int(used.row)
        start_col = int(used.column)
    except Exception:
        return None

    if values is None:
        return None
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    target = normalize_text(text)
    for row_idx, row_values in enumerate(values):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == target:
                return start_row + row_idx, start_col + col_idx
    return None


def header_entries_around_anchor(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    left_span: int = 20,
    right_span: int = 12,
) -> List[Tuple[int, str]]:
    start_col = max(1, anchor_col - left_span)
    end_col = anchor_col + right_span
    row_candidates = [anchor_row - 1, anchor_row]
    entries: List[Tuple[int, str]] = []
    for row in row_candidates:
        if row < 1:
            continue
        values = sheet.range((row, start_col), (row, end_col)).value
        if not isinstance(values, list):
            values = [values]
        for idx, value in enumerate(values):
            key = normalize_text(value)
            if key:
                entries.append((start_col + idx, key))
    return entries


def find_column(
    entries: Sequence[Tuple[int, str]],
    keyword_groups: Sequence[Sequence[str]],
    fallback_col: int,
) -> int:
    for keywords in keyword_groups:
        for col, key in entries:
            if all(keyword in key for keyword in keywords):
                return col
    return fallback_col


def find_last_numeric_row(sheet: xw.Sheet, start_row: int, required_cols: Sequence[int]) -> Optional[int]:
    for row in range(start_row, 0, -1):
        values = [safe_cell_value(sheet, row, col) for col in required_cols]
        if all(is_number(value) for value in values):
            return row
    return None


def find_first_numeric_row(sheet: xw.Sheet, last_row: int, required_cols: Sequence[int]) -> int:
    row = last_row
    while row >= 1:
        values = [safe_cell_value(sheet, row, col) for col in required_cols]
        if all(is_number(value) for value in values):
            row -= 1
            continue
        break
    return row + 1


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    cell.api.FormulaR1C1 = formula_r1c1


def close_source_workbook(wb: Optional[xw.Book]) -> None:
    if wb is None:
        return
    close_attempts = [
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ]
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception:
            continue


def extract_empirical_rows(wb: xw.Book, meta: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if EMPIRICAL_SHEET_NAME not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    max_anchor = find_text_anchor(sheet, "max")
    if max_anchor is None:
        return []
    anchor_row, anchor_col = max_anchor

    min_anchor = find_text_anchor(sheet, "min")
    min_col = min_anchor[1] if min_anchor and abs(min_anchor[0] - anchor_row) <= 2 else anchor_col + 1

    headers = header_entries_around_anchor(sheet, anchor_row, anchor_col)
    num_quarters_col = find_column(headers, (("num", "quarter"),), anchor_col - 12)
    last_quarter_col = find_column(headers, (("last", "quarter"),), anchor_col - 13)
    forecast_col = find_column(headers, (("estimated", "total", "sold"), ("forecast",)), anchor_col - 5)
    actual_col = find_column(headers, (("actual",), ("reported", "sales")), anchor_col - 4)
    quarterly_sales_col = find_column(headers, (("quarterly", "sales"),), anchor_col - 7)
    reported_sales_col = find_column(headers, (("reported", "sales"),), actual_col)
    growth_rate_col = find_column(headers, (("growth", "rate"),), anchor_col - 3)
    captured_col = find_column(headers, (("captured", "db"), ("captured",)), anchor_col - 2)
    avg_pen_table_col = find_column(headers, (("avg", "penetration"),), anchor_col - 8)

    penetration_col = anchor_col - 11
    quarter_label_col = last_quarter_col if last_quarter_col > 0 else anchor_col - 12

    last_hist_row = find_last_numeric_row(sheet, anchor_row - 1, [penetration_col])
    if last_hist_row is None:
        return []
    first_hist_row = find_first_numeric_row(sheet, last_hist_row, [penetration_col])

    temp_row_start = anchor_row + 40
    temp_col = anchor_col + 8
    windows: List[Tuple[int, int]] = []
    for n_quarters in range(1, 11):
        start_row = max(first_hist_row, last_hist_row - n_quarters + 1)
        windows.append((n_quarters, start_row))
        formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{last_hist_row}C{penetration_col})"
        set_formula2_r1c1(sheet.cells(temp_row_start + n_quarters - 1, temp_col), formula)

    wb.app.calculate()

    avg_values = sheet.range((temp_row_start, temp_col), (temp_row_start + 9, temp_col)).value
    if not isinstance(avg_values, list):
        avg_values = [avg_values]
    sheet.range((temp_row_start, temp_col), (temp_row_start + 9, temp_col)).value = None

    rows: List[Dict[str, Any]] = []
    for idx, (n_quarters, start_row) in enumerate(windows):
        candidate_row = anchor_row + n_quarters

        avg_penetration_pct = as_float(avg_values[idx]) if idx < len(avg_values) else None
        if avg_penetration_pct is None:
            avg_penetration_pct = as_float(safe_cell_value(sheet, candidate_row, avg_pen_table_col))

        num_quarters_used = as_float(safe_cell_value(sheet, candidate_row, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = float(n_quarters)

        last_quarter_used = safe_cell_value(sheet, candidate_row, last_quarter_col)
        if last_quarter_used in (None, ""):
            last_quarter_used = safe_cell_value(sheet, start_row, quarter_label_col)

        quarterly_sales = as_float(safe_cell_value(sheet, candidate_row, quarterly_sales_col))
        if quarterly_sales is None:
            quarterly_sales = as_float(safe_cell_value(sheet, last_hist_row, quarterly_sales_col))

        reported_sales = as_float(safe_cell_value(sheet, candidate_row, reported_sales_col))
        if reported_sales is None:
            reported_sales = as_float(safe_cell_value(sheet, candidate_row, actual_col))

        forecast_value = as_float(safe_cell_value(sheet, candidate_row, forecast_col))
        if forecast_value is None and avg_penetration_pct is not None and quarterly_sales is not None:
            forecast_value = avg_penetration_pct * quarterly_sales

        actual_value = as_float(safe_cell_value(sheet, candidate_row, actual_col))
        if actual_value is None:
            actual_value = reported_sales

        forecast_max = as_float(safe_cell_value(sheet, candidate_row, anchor_col))
        forecast_min = as_float(safe_cell_value(sheet, candidate_row, min_col))
        growth_rate_pct = as_float(safe_cell_value(sheet, candidate_row, growth_rate_col))
        sales_captured_in_db_pct = as_float(safe_cell_value(sheet, candidate_row, captured_col))

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": int(num_quarters_used),
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
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
    return rows


def row_signature(values: Sequence[Optional[float]]) -> Tuple[Optional[float], ...]:
    signature: List[Optional[float]] = []
    for value in values:
        if value is None:
            signature.append(None)
        else:
            signature.append(round(float(value), 10))
    return tuple(signature)


def extract_regression_rows(wb: xw.Book, meta: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if REGRESSION_SHEET_NAME not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets[REGRESSION_SHEET_NAME]
    max_anchor = find_text_anchor(sheet, "max")
    if max_anchor is None:
        return []
    anchor_row, anchor_col = max_anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    min_anchor = find_text_anchor(sheet, "min")
    min_col = min_anchor[1] if min_anchor and abs(min_anchor[0] - anchor_row) <= 2 else anchor_col + 1

    headers = header_entries_around_anchor(sheet, anchor_row, anchor_col)
    num_quarters_col = find_column(headers, (("num", "quarter"),), anchor_col - 12)
    forecast_col = find_column(
        headers,
        (
            ("tot", "fcst", "sa"),
            ("forecast", "without", "sa"),
            ("forecast",),
        ),
        anchor_col - 5,
    )
    actual_col = find_column(headers, (("actual",),), anchor_col - 4)

    last_hist_row = find_last_numeric_row(sheet, anchor_row - 1, [x_col, y_col])
    if last_hist_row is None:
        return []
    first_hist_row = find_first_numeric_row(sheet, last_hist_row, [x_col, y_col])

    temp_row_start = anchor_row + 40
    intercept_col = anchor_col + 8
    slope_col = anchor_col + 9
    windows: List[Tuple[int, int]] = []
    for n_quarters in range(1, 11):
        start_row = max(first_hist_row, last_hist_row - n_quarters + 1)
        windows.append((n_quarters, start_row))
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{last_hist_row}C{y_col},"
            f"R{start_row}C{x_col}:R{last_hist_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{last_hist_row}C{y_col},"
            f"R{start_row}C{x_col}:R{last_hist_row}C{x_col})"
        )
        set_formula2_r1c1(sheet.cells(temp_row_start + n_quarters - 1, intercept_col), intercept_formula)
        set_formula2_r1c1(sheet.cells(temp_row_start + n_quarters - 1, slope_col), slope_formula)

    wb.app.calculate()

    intercept_values = sheet.range(
        (temp_row_start, intercept_col),
        (temp_row_start + 9, intercept_col),
    ).value
    slope_values = sheet.range(
        (temp_row_start, slope_col),
        (temp_row_start + 9, slope_col),
    ).value
    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]
    sheet.range((temp_row_start, intercept_col), (temp_row_start + 9, slope_col)).value = None

    latest_x = as_float(safe_cell_value(sheet, last_hist_row, x_col))
    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for idx, (n_quarters, _start_row) in enumerate(windows):
        candidate_row = anchor_row + n_quarters

        num_quarters_used = as_float(safe_cell_value(sheet, candidate_row, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = float(n_quarters)

        intercept = as_float(intercept_values[idx]) if idx < len(intercept_values) else None
        slope = as_float(slope_values[idx]) if idx < len(slope_values) else None

        forecast_total_without_sa = as_float(safe_cell_value(sheet, candidate_row, forecast_col))
        if forecast_total_without_sa is None and intercept is not None and slope is not None and latest_x is not None:
            forecast_total_without_sa = intercept + (slope * latest_x)

        actual_value = as_float(safe_cell_value(sheet, candidate_row, actual_col))
        forecast_max = as_float(safe_cell_value(sheet, candidate_row, anchor_col))
        forecast_min = as_float(safe_cell_value(sheet, candidate_row, min_col))
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        current_signature = row_signature(
            [num_quarters_used, forecast_total_without_sa, forecast_max, forecast_min, intercept, slope]
        )
        if previous_signature == current_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": int(num_quarters_used),
                "num_quarters_used": int(num_quarters_used),
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
    return rows


def autosize_columns(sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
    for col_idx in range(1, sheet.max_column + 1):
        column_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        sheet.column_dimensions[column_letter].width = min(max(12, max_len + 2), 48)


def write_rows_to_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    sheet = wb.create_sheet(title=sheet_name)
    sheet.append(list(headers))
    for item in rows:
        sheet.append([item.get(header) for header in headers])

    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"

    if sheet.max_row >= 1:
        sheet.auto_filter.ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"
    autosize_columns(sheet)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = openpyxl.Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_rows_to_sheet(workbook, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_rows_to_sheet(workbook, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    workbook.save(output_path)


def run() -> None:
    source_dir = Path(input_dir).expanduser()
    target_dir = Path(output_dir).expanduser()

    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {source_dir}")

    target_dir.mkdir(parents=True, exist_ok=True)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(source_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (reason: temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (reason: not .xlsx)")
                continue

            meta = parse_file_metadata(file_path.name)
            if meta is None:
                print(f"skipped file: {file_path.name} (reason: unrecognized file label format)")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (reason: processing error: {exc})")
            finally:
                close_source_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_path = next_output_path(target_dir, source_dir.name)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
