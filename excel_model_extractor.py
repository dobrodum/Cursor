#!/usr/bin/env python3
"""Extract empirical and regression candidate rows from model workbooks.

This script scans `input_dir` for .xlsx files, opens each source workbook once
with xlwings, processes both model sheets while the workbook is open, and writes
all extracted rows to a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------------------------- USER CONFIGURATION -------------------------- #
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")
# ----------------------------------------------------------------------- #


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

MONTH_MAP = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    try:
        text = str(value).strip().replace(",", "")
        if not text:
            return None
        out = float(text)
        if math.isnan(out) or math.isinf(out):
            return None
        return out
    except (TypeError, ValueError):
        return None


def display_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return value


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator is None or denominator == 0:
        return None
    return numerator / denominator


def parse_file_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].strip().upper()
    elif parts:
        ticker = normalize_text(parts[0]).upper()[:6]

    period_source = parts[2] if len(parts) >= 3 else stem
    period_match = re.search(
        r"(?i)(early|mid|late)\s*([a-z]{3,9})\s*([12][0-9]{3})",
        period_source,
    )

    model_period = ""
    model_date = ""

    if period_match:
        period_word = period_match.group(1).lower()
        month_token = period_match.group(2)
        year = int(period_match.group(3))

        month_key = month_token[:3].lower()
        month = MONTH_MAP.get(month_key)
        day = PERIOD_DAY_MAP.get(period_word, 15)

        period_title = period_word.capitalize()
        month_title = month_token[:1].upper() + month_token[1:3].lower()
        model_period = f"{period_title}{month_title}_{year}"

        if month is not None:
            model_date = f"{year:04d}-{month:02d}-{day:02d}"
    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or stem

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def resolve_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"

    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    version = 1
    while True:
        candidate = output_folder / f"{base_name}.{version}.xlsx"
        if not candidate.exists():
            return candidate
        version += 1


def iter_source_files(input_folder: Path) -> Iterable[Path]:
    if not input_folder.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_folder}")
    if not input_folder.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {input_folder}")

    param_name_prefix = f"{input_folder.name}_param"
    for path in sorted(input_folder.iterdir()):
        if not path.is_file():
            print(f"SKIPPED: {path.name} (not a file)")
            continue
        if path.name.startswith("~"):
            print(f"SKIPPED: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"SKIPPED: {path.name} (not an .xlsx file)")
            continue
        if path.stem.lower().startswith(param_name_prefix):
            print(f"SKIPPED: {path.name} (looks like prior output workbook)")
            continue
        yield path


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    workbook.api.Close(SaveChanges=False)


def read_used_range(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int, int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    row_count = max(int(used.rows.count), 1)
    col_count = max(int(used.columns.count), 1)
    end_row = start_row + row_count - 1
    end_col = start_col + col_count - 1

    values = used.value
    if values is None:
        return [], start_row, start_col, end_row, end_col

    if not isinstance(values, list):
        matrix = [[values]]
    elif values and not isinstance(values[0], list):
        matrix = [values]
    else:
        matrix = values
    return matrix, start_row, start_col, end_row, end_col


def get_matrix_value(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int,
) -> Any:
    r_idx = row - start_row
    c_idx = col - start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(matrix):
        return None
    row_values = matrix[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def build_label_index(
    matrix: Sequence[Sequence[Any]], start_row: int, start_col: int
) -> Dict[str, List[Tuple[int, int]]]:
    label_index: Dict[str, List[Tuple[int, int]]] = {}
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str):
                key = normalize_text(value)
                if key:
                    label_index.setdefault(key, []).append(
                        (start_row + r_idx, start_col + c_idx)
                    )
    return label_index


def find_anchor_max(
    matrix: Sequence[Sequence[Any]], start_row: int, start_col: int
) -> Optional[Tuple[int, int]]:
    hits: List[Tuple[int, int]] = []
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == "max":
                hits.append((start_row + r_idx, start_col + c_idx))
    if not hits:
        return None
    # Usually the model summary anchor is the lower "max"; pick the latest one.
    hits.sort(key=lambda rc: (rc[0], rc[1]), reverse=True)
    return hits[0]


def extract_numeric_near_label(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    label_index: Dict[str, List[Tuple[int, int]]],
    labels: Sequence[str],
) -> Optional[float]:
    normalized_labels = [normalize_text(label) for label in labels]
    for key in normalized_labels:
        for row, col in label_index.get(key, []):
            # Rightward scan first.
            for delta in range(1, 9):
                val = to_float(get_matrix_value(matrix, start_row, start_col, row, col + delta))
                if val is not None:
                    return val
            # Then downward scan.
            for delta in range(1, 9):
                val = to_float(get_matrix_value(matrix, start_row, start_col, row + delta, col))
                if val is not None:
                    return val
    return None


def extract_any_near_label(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    label_index: Dict[str, List[Tuple[int, int]]],
    labels: Sequence[str],
) -> Any:
    normalized_labels = [normalize_text(label) for label in labels]
    for key in normalized_labels:
        for row, col in label_index.get(key, []):
            for delta in range(1, 9):
                val = get_matrix_value(matrix, start_row, start_col, row, col + delta)
                if val not in (None, ""):
                    return val
            for delta in range(1, 9):
                val = get_matrix_value(matrix, start_row, start_col, row + delta, col)
                if val not in (None, ""):
                    return val
    return None


def find_numeric_rows_in_column(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    col: int,
    row_start: int,
    row_end: int,
) -> List[int]:
    rows: List[int] = []
    for row in range(row_start, row_end + 1):
        val = to_float(get_matrix_value(matrix, start_row, start_col, row, col))
        if val is not None:
            rows.append(row)
    return rows


def pick_penetration_column(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    anchor_col: int,
) -> int:
    search_top = max(start_row, anchor_row - 60)
    search_bottom = anchor_row - 1
    if search_bottom < search_top:
        return max(start_col, anchor_col - 9)

    best_col = max(start_col, anchor_col - 9)
    best_score = -1
    left_bound = max(start_col, anchor_col - 20)
    right_bound = max(start_col, anchor_col)

    for col in range(left_bound, right_bound + 1):
        score = 0
        for row in range(search_top, search_bottom + 1):
            val = to_float(get_matrix_value(matrix, start_row, start_col, row, col))
            if val is None:
                continue
            if 0 <= val <= 2.5:
                score += 2
            else:
                score += 1
        if score > best_score:
            best_score = score
            best_col = col
    return best_col


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    # Prefer Formula2R1C1 to keep formulas in R1C1 notation.
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass
    cell.api.FormulaR1C1 = formula


def excel_round_signature(values: Sequence[Optional[float]]) -> Tuple[Optional[float], ...]:
    rounded: List[Optional[float]] = []
    for value in values:
        if value is None:
            rounded.append(None)
        else:
            rounded.append(round(value, 10))
    return tuple(rounded)


def process_empirical_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    matrix, start_row, start_col, end_row, end_col = read_used_range(sheet)
    if not matrix:
        return []

    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        print(f"SKIPPED empirical in {source_file}: could not find 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    data_end_row = anchor_row - 1
    if data_end_row < start_row:
        print(f"SKIPPED empirical in {source_file}: invalid data range")
        return []

    labels = build_label_index(matrix, start_row, start_col)
    penetration_col = pick_penetration_column(
        matrix, start_row, start_col, anchor_row, anchor_col
    )
    numeric_pen_rows = find_numeric_rows_in_column(
        matrix, start_row, start_col, penetration_col, start_row, data_end_row
    )
    if len(numeric_pen_rows) < 2:
        print(f"SKIPPED empirical in {source_file}: not enough penetration rows")
        return []

    quarterly_sales = extract_numeric_near_label(
        matrix,
        start_row,
        start_col,
        labels,
        labels=["quarterly sales", "quarterly_sales", "current quarter sales", "q sales"],
    )
    reported_sales = extract_numeric_near_label(
        matrix,
        start_row,
        start_col,
        labels,
        labels=["reported sales", "reported_sales", "reported revenue", "actual sales"],
    )
    growth_rate_pct = extract_numeric_near_label(
        matrix,
        start_row,
        start_col,
        labels,
        labels=["growth rate", "growth_rate", "growth %", "growth pct"],
    )
    sales_captured_in_db_pct = extract_numeric_near_label(
        matrix,
        start_row,
        start_col,
        labels,
        labels=["sales captured in db", "sales_captured_in_db", "captured in db"],
    )

    if sales_captured_in_db_pct is None:
        sales_captured_in_db_pct = safe_div(quarterly_sales, reported_sales)

    last_quarter_used = extract_any_near_label(
        matrix,
        start_row,
        start_col,
        labels,
        labels=["last quarter used", "latest quarter", "quarter"],
    )
    if last_quarter_used in (None, ""):
        quarter_guess = get_matrix_value(
            matrix, start_row, start_col, numeric_pen_rows[-1], penetration_col - 1
        )
        if quarter_guess not in (None, ""):
            last_quarter_used = quarter_guess

    scratch_start_row = max(end_row + 3, anchor_row + 3)
    scratch_col_avg = end_col + 3
    scratch_col_max = end_col + 4
    scratch_col_min = end_col + 5

    row_configs: List[Tuple[int, int, int]] = []
    for n in range(1, N_QUARTERS + 1):
        if len(numeric_pen_rows) < n:
            continue
        selected_rows = numeric_pen_rows[-n:]
        row_configs.append((n, selected_rows[0], selected_rows[-1]))

    if not row_configs:
        return []

    for idx, (_, start_n, end_n) in enumerate(row_configs):
        row = scratch_start_row + idx
        avg_cell = sheet.range((row, scratch_col_avg))
        max_cell = sheet.range((row, scratch_col_max))
        min_cell = sheet.range((row, scratch_col_min))

        set_formula2_r1c1(
            avg_cell,
            f"=AVERAGE(R{start_n}C{penetration_col}:R{end_n}C{penetration_col})",
        )
        set_formula2_r1c1(
            max_cell,
            f"=MAX(R{start_n}C{penetration_col}:R{end_n}C{penetration_col})",
        )
        set_formula2_r1c1(
            min_cell,
            f"=MIN(R{start_n}C{penetration_col}:R{end_n}C{penetration_col})",
        )

    workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx, (n_quarters, _, _) in enumerate(row_configs):
        row = scratch_start_row + idx
        avg_pen = to_float(sheet.range((row, scratch_col_avg)).value)
        pen_max = to_float(sheet.range((row, scratch_col_max)).value)
        pen_min = to_float(sheet.range((row, scratch_col_min)).value)

        forecast_value = None
        forecast_max = None
        forecast_min = None

        if quarterly_sales is not None and avg_pen not in (None, 0):
            forecast_value = quarterly_sales / avg_pen

        if quarterly_sales is not None and pen_min not in (None, 0):
            forecast_max = quarterly_sales / pen_min
        if quarterly_sales is not None and pen_max not in (None, 0):
            forecast_min = quarterly_sales / pen_max

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
                "parameter_value": avg_pen,
                "num_quarters_used": n_quarters,
                "last_quarter_used": display_value(last_quarter_used),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    # Clear temporary formula cells written for calculations.
    sheet.range(
        (scratch_start_row, scratch_col_avg),
        (scratch_start_row + len(row_configs) - 1, scratch_col_min),
    ).clear_contents()

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    matrix, start_row, start_col, end_row, end_col = read_used_range(sheet)
    if not matrix:
        return []

    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        print(f"SKIPPED regression in {source_file}: could not find 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_end_row = anchor_row - 1
    if data_end_row < start_row:
        print(f"SKIPPED regression in {source_file}: invalid data range")
        return []

    series: List[Tuple[int, float, float]] = []
    for row in range(start_row, data_end_row + 1):
        x_val = to_float(get_matrix_value(matrix, start_row, start_col, row, x_col))
        y_val = to_float(get_matrix_value(matrix, start_row, start_col, row, y_col))
        if x_val is not None and y_val is not None:
            series.append((row, x_val, y_val))

    if len(series) < 2:
        print(f"SKIPPED regression in {source_file}: not enough x/y points")
        return []

    if len(series) > N_QUARTERS:
        series = series[-N_QUARTERS:]

    scratch_start_row = max(end_row + 3, anchor_row + 3)
    scratch_col_intercept = end_col + 3
    scratch_col_slope = end_col + 4

    row_configs: List[Tuple[int, int, int]] = []
    for n in range(2, len(series) + 1):
        start_n = series[-n][0]
        end_n = series[-1][0]
        row_configs.append((n, start_n, end_n))

    for idx, (_, start_n, end_n) in enumerate(row_configs):
        row = scratch_start_row + idx
        intercept_cell = sheet.range((row, scratch_col_intercept))
        slope_cell = sheet.range((row, scratch_col_slope))

        set_formula2_r1c1(
            intercept_cell,
            f"=INTERCEPT(R{start_n}C{y_col}:R{end_n}C{y_col},R{start_n}C{x_col}:R{end_n}C{x_col})",
        )
        set_formula2_r1c1(
            slope_cell,
            f"=SLOPE(R{start_n}C{y_col}:R{end_n}C{y_col},R{start_n}C{x_col}:R{end_n}C{x_col})",
        )

    workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for idx, (n_quarters, _, _) in enumerate(row_configs):
        subset = series[-n_quarters:]
        row = scratch_start_row + idx
        intercept = to_float(sheet.range((row, scratch_col_intercept)).value)
        slope = to_float(sheet.range((row, scratch_col_slope)).value)

        forecast_value = None
        if intercept is not None and slope is not None:
            next_x = subset[-1][1] + 1.0
            forecast_value = intercept + (slope * next_x)

        actual_candidate = to_float(
            get_matrix_value(matrix, start_row, start_col, subset[-1][0] + 1, y_col)
        )
        actual_value = actual_candidate if actual_candidate is not None else ""

        forecast_max = None
        forecast_min = None
        if intercept is not None and slope is not None and forecast_value is not None:
            residuals = [y - (intercept + slope * x) for _, x, y in subset]
            if len(subset) > 2:
                stdev = math.sqrt(
                    sum(res * res for res in residuals) / (len(subset) - 2)
                )
            else:
                stdev = 0.0
            forecast_max = forecast_value + (1.96 * stdev)
            forecast_min = forecast_value - (1.96 * stdev)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = excel_round_signature([intercept, slope, forecast_value])
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
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
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

    # Clear temporary formula cells written for calculations.
    sheet.range(
        (scratch_start_row, scratch_col_intercept),
        (scratch_start_row + len(row_configs) - 1, scratch_col_slope),
    ).clear_contents()

    return rows


def write_output_sheet(
    worksheet,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([display_value(row.get(column, "")) for column in columns])

    header_font = Font(bold=True)
    for col_index in range(1, len(columns) + 1):
        worksheet.cell(row=1, column=col_index).font = header_font

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_index, col_name in enumerate(columns, start=1):
        max_width = len(col_name) + 2
        for row_index in range(2, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=row_index, column=col_index).value
            length = len(str(cell_value)) if cell_value is not None else 0
            if length + 2 > max_width:
                max_width = length + 2
        worksheet.column_dimensions[get_column_letter(col_index)].width = min(
            max(max_width, 12), 48
        )


def save_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    output_wb = Workbook()
    empirical_ws = output_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = output_wb.create_sheet("regression_candidates")

    write_output_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    output_wb.save(output_path)


def main() -> None:
    output_path = resolve_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False
    app.calculation = "manual"

    try:
        for file_path in iter_source_files(input_dir):
            metadata = parse_file_label(file_path)
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)

                empirical_sheet = None
                regression_sheet = None
                for sht in workbook.sheets:
                    if sht.name.strip().lower() == "empirical model":
                        empirical_sheet = sht
                    elif sht.name.strip().lower() == "regression model":
                        regression_sheet = sht

                file_empirical_rows: List[Dict[str, Any]] = []
                file_regression_rows: List[Dict[str, Any]] = []

                if empirical_sheet is None:
                    print(f"SKIPPED empirical in {file_path.name}: sheet not found")
                else:
                    file_empirical_rows = process_empirical_sheet(
                        workbook, empirical_sheet, metadata, file_path.name
                    )

                if regression_sheet is None:
                    print(f"SKIPPED regression in {file_path.name}: sheet not found")
                else:
                    file_regression_rows = process_regression_sheet(
                        workbook, regression_sheet, metadata, file_path.name
                    )

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_files += 1
                print(
                    f"PROCESSED: {file_path.name} | "
                    f"empirical_rows={len(file_empirical_rows)} | "
                    f"regression_rows={len(file_regression_rows)}"
                )

            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    save_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {processed_files}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
