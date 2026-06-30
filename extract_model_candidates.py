#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from Excel workbooks.

This script:
1) Scans input_dir for .xlsx files (excluding temporary "~" files),
2) Opens each workbook once in a single hidden Excel app,
3) Processes both "Empirical Model" and "Regression Model" while open,
4) Writes one combined output workbook with:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------- Required top-level I/O variables --------
input_dir = "./input"
output_dir = "./output"


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10
SCRATCH_COL = 16384  # XFD, used for temporary formulas
SCRATCH_ROW_BASE = 1

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

MONTH_TO_NUMBER = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

PERIOD_TO_DAY = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_RE = re.compile(
    r"(?P<phase>Early|Mid|Late)\s*(?P<month>[A-Za-z]{3,9})\s*(?P<year>\d{4})",
    flags=re.IGNORECASE,
)


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    try:
        if isinstance(value, bool):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def diff(a: Any, b: Any) -> Optional[float]:
    left = to_float(a)
    right = to_float(b)
    if left is None or right is None:
        return None
    return left - right


def first_non_none(*values: Any) -> Any:
    for value in values:
        if value is not None and value != "":
            return value
    return None


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 else ""
    period_segment = parts[2] if len(parts) > 2 else ""
    period_token = period_segment.split("_")[0]

    model_period = ""
    model_date = ""

    match = PERIOD_RE.search(period_token)
    if match:
        phase = match.group("phase").title()
        month_token = match.group("month").title()[:3]
        year = match.group("year")
        month_num = MONTH_TO_NUMBER.get(month_token)
        if month_num is not None:
            model_period = f"{phase}{month_token}_{year}"
            day = PERIOD_TO_DAY[phase]
            model_date = f"{year}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or stem)
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_label_positions(
    matrix: List[List[Any]], start_row: int, start_col: int
) -> Dict[str, List[Tuple[int, int]]]:
    positions: Dict[str, List[Tuple[int, int]]] = {}
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if not isinstance(value, str):
                continue
            label = normalize_label(value)
            if not label:
                continue
            positions.setdefault(label, []).append((start_row + r_idx, start_col + c_idx))
    return positions


def find_label_position(
    label_positions: Dict[str, List[Tuple[int, int]]],
    *,
    exact: Optional[str] = None,
    contains: Optional[Sequence[str]] = None,
) -> Optional[Tuple[int, int]]:
    if exact:
        exact_key = normalize_label(exact)
        if exact_key in label_positions:
            return label_positions[exact_key][0]

    if contains:
        needles = [normalize_label(needle) for needle in contains]
        for label, positions in label_positions.items():
            if any(needle in label for needle in needles):
                return positions[0]
    return None


def matrix_value(
    matrix: List[List[Any]], start_row: int, start_col: int, row: int, col: int
) -> Any:
    r_idx = row - start_row
    c_idx = col - start_col
    if r_idx < 0 or c_idx < 0 or r_idx >= len(matrix):
        return None
    if c_idx >= len(matrix[r_idx]):
        return None
    return matrix[r_idx][c_idx]


def collect_numeric_rows(
    matrix: List[List[Any]],
    start_row: int,
    start_col: int,
    *,
    target_col: int,
    top_row: int,
    bottom_row: int,
) -> List[Tuple[int, float]]:
    rows: List[Tuple[int, float]] = []
    for row in range(top_row, bottom_row + 1):
        numeric_value = to_float(matrix_value(matrix, start_row, start_col, row, target_col))
        if numeric_value is not None:
            rows.append((row, numeric_value))
    return rows


def set_formula2(range_obj: xw.Range, formula: str) -> None:
    try:
        range_obj.formula2 = formula
    except Exception:
        # Safe fallback if formula2 is not supported by host Excel build.
        range_obj.formula = formula


def safe_close_source_workbook(wb: Optional[xw.Book]) -> None:
    if wb is None:
        return
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception as exc:
        print(f"Warning: failed to close workbook safely: {exc}")


def get_sheet_or_none(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[name]
    except Exception:
        return None


def locate_value_cell(
    label_positions: Dict[str, List[Tuple[int, int]]],
    *,
    contains: Sequence[str],
    default: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    label_pos = find_label_position(label_positions, contains=contains)
    if label_pos is None:
        return default
    return (label_pos[0], label_pos[1] + 1)


def read_cell(sheet: xw.Sheet, cell_pos: Optional[Tuple[int, int]]) -> Any:
    if cell_pos is None:
        return None
    return sheet.range(cell_pos).value


def process_empirical_sheet(
    wb: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_or_none(wb, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        print(f"Skipped empirical for {source_file}: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    used = sheet.used_range
    matrix = to_2d(used.value)
    start_row = used.row
    start_col = used.column
    labels = build_label_positions(matrix, start_row, start_col)

    anchor = find_label_position(labels, exact="max") or find_label_position(
        labels, contains=("max",)
    )
    if anchor is None:
        print(f"Skipped empirical for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    top_row = start_row
    bottom_row = max(start_row, anchor_row - 1)

    # Anchor-based offsets from `max` cell for historical quarter data.
    quarterly_sales_col = anchor_col - 11
    reported_sales_col = anchor_col - 7
    penetration_col = anchor_col - 9

    penetration_rows = collect_numeric_rows(
        matrix,
        start_row,
        start_col,
        target_col=penetration_col,
        top_row=top_row,
        bottom_row=bottom_row,
    )

    if not penetration_rows:
        print(f"Skipped empirical for {source_file}: no penetration history found")
        return []

    quarter_limit = min(N_QUARTERS, len(penetration_rows))

    avg_pen_cell = locate_value_cell(
        labels,
        contains=("avg penetration", "average penetration"),
    )
    forecast_value_cell = locate_value_cell(
        labels,
        contains=("estimated total sold", "est total sold", "forecast"),
        default=(anchor_row - 2, anchor_col + 1),
    )
    actual_value_cell = locate_value_cell(
        labels,
        contains=("reported sales", "actual sales", "actual"),
        default=(anchor_row - 1, anchor_col + 1),
    )
    forecast_max_cell = (anchor_row, anchor_col + 1)
    forecast_min_cell = locate_value_cell(
        labels,
        contains=("min",),
        default=(anchor_row + 1, anchor_col + 1),
    )
    growth_rate_cell = locate_value_cell(labels, contains=("growth rate",))
    captured_pct_cell = locate_value_cell(
        labels,
        contains=("captured in db", "captured in database", "sales captured"),
    )

    rows: List[Dict[str, Any]] = []
    scratch_avg = sheet.range((SCRATCH_ROW_BASE, SCRATCH_COL))

    for n_quarters in range(1, quarter_limit + 1):
        start_hist_row = penetration_rows[-n_quarters][0]
        end_hist_row = penetration_rows[-1][0]
        avg_formula = (
            f"=AVERAGE(R{start_hist_row}C{penetration_col}:"
            f"R{end_hist_row}C{penetration_col})"
        )

        formula_target = sheet.range(avg_pen_cell) if avg_pen_cell else scratch_avg
        set_formula2(formula_target, avg_formula)
        wb.app.calculate()

        avg_penetration_pct = formula_target.value
        forecast_value = read_cell(sheet, forecast_value_cell)
        actual_value = read_cell(sheet, actual_value_cell)
        forecast_max = read_cell(sheet, forecast_max_cell)
        forecast_min = read_cell(sheet, forecast_min_cell)
        growth_rate_pct = read_cell(sheet, growth_rate_cell)
        sales_captured_pct = read_cell(sheet, captured_pct_cell)

        last_quarter_used = matrix_value(
            matrix, start_row, start_col, end_hist_row, quarterly_sales_col - 1
        )
        quarterly_sales = matrix_value(
            matrix, start_row, start_col, end_hist_row, quarterly_sales_col
        )
        reported_sales = matrix_value(
            matrix, start_row, start_col, end_hist_row, reported_sales_col
        )

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": diff(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": first_non_none(reported_sales, actual_value),
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def rows_match_regression(prev_row: Dict[str, Any], new_row: Dict[str, Any]) -> bool:
    compare_keys = ("intercept", "slope", "forecast_value", "forecast_max", "forecast_min")
    for key in compare_keys:
        prev_val = to_float(prev_row.get(key))
        new_val = to_float(new_row.get(key))
        if prev_val is None or new_val is None:
            if prev_row.get(key) != new_row.get(key):
                return False
        elif abs(prev_val - new_val) > 1e-10:
            return False
    return True


def process_regression_sheet(
    wb: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_or_none(wb, REGRESSION_SHEET_NAME)
    if sheet is None:
        print(f"Skipped regression for {source_file}: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    used = sheet.used_range
    matrix = to_2d(used.value)
    start_row = used.row
    start_col = used.column
    labels = build_label_positions(matrix, start_row, start_col)

    anchor = find_label_position(labels, exact="max") or find_label_position(
        labels, contains=("max",)
    )
    if anchor is None:
        print(f"Skipped regression for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    top_row = start_row
    bottom_row = max(start_row, anchor_row - 1)
    x_rows = collect_numeric_rows(
        matrix, start_row, start_col, target_col=x_col, top_row=top_row, bottom_row=bottom_row
    )
    y_rows = collect_numeric_rows(
        matrix, start_row, start_col, target_col=y_col, top_row=top_row, bottom_row=bottom_row
    )

    y_lookup = {row: value for row, value in y_rows}
    paired_rows = [(row, x_val, y_lookup[row]) for row, x_val in x_rows if row in y_lookup]

    if not paired_rows:
        print(f"Skipped regression for {source_file}: no paired x/y history found")
        return []

    quarter_limit = min(N_QUARTERS, len(paired_rows))

    forecast_value_cell = locate_value_cell(
        labels,
        contains=("tot fcst w/o sa", "tot fcst without sa", "forecast w/o sa"),
    )
    forecast_max_cell = (anchor_row, anchor_col + 1)
    forecast_min_cell = locate_value_cell(
        labels,
        contains=("min",),
        default=(anchor_row + 1, anchor_col + 1),
    )
    actual_value_cell = locate_value_cell(
        labels,
        contains=("actual", "reported sales"),
    )

    intercept_cell = sheet.range((SCRATCH_ROW_BASE, SCRATCH_COL))
    slope_cell = sheet.range((SCRATCH_ROW_BASE + 1, SCRATCH_COL))

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, quarter_limit + 1):
        start_hist_row = paired_rows[-n_quarters][0]
        end_hist_row = paired_rows[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        forecast_value = read_cell(sheet, forecast_value_cell)
        if forecast_value is None:
            next_x = paired_rows[-1][1] + 1
            if to_float(intercept) is not None and to_float(slope) is not None:
                forecast_value = to_float(intercept) + to_float(slope) * next_x

        forecast_max = read_cell(sheet, forecast_max_cell)
        forecast_min = read_cell(sheet, forecast_min_cell)
        actual_value = read_cell(sheet, actual_value_cell)

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": diff(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if n_quarters == quarter_limit and rows and rows_match_regression(rows[-1], row):
            continue
        rows.append(row)

    return rows


def write_rows_to_sheet(
    ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]
) -> None:
    ws.append(list(headers))
    for item in rows:
        ws.append([item.get(header, "") for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_length = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_length:
                max_length = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_length + 2), 48)


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    ws_reg = wb.create_sheet(title="regression_candidates")

    write_rows_to_sheet(ws_emp, EMPIRICAL_HEADERS, empirical_rows)
    write_rows_to_sheet(ws_reg, REGRESSION_HEADERS, regression_rows)

    wb.save(output_file)


def collect_source_files(input_path: Path, output_prefix: str) -> List[Path]:
    source_files: List[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        if file_path.stem.startswith(output_prefix):
            print(f"Skipped {file_path.name}: appears to be prior output")
            continue
        source_files.append(file_path)
    return source_files


def run() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"input_dir does not exist or is not a directory: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_prefix = f"{input_path.name}_PARAM"
    source_files = collect_source_files(input_path, output_prefix=output_prefix)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in source_files:
            print(f"Processing: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                safe_close_source_workbook(wb)
    finally:
        app.quit()

    output_file = build_output_path(input_path, output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
