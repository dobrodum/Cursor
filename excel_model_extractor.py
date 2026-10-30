#!/usr/bin/env python3
"""Extract empirical/regression model candidates from all XLSX files in a folder.

This script opens each source workbook exactly once with xlwings, processes both:
  - Empirical Model
  - Regression Model
and writes one combined output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from datetime import date
import math
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------------------------
# User-configurable paths
# -------------------------
input_dir = Path("./input")
output_dir = Path("./output")

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

MONTH_TO_NUM = {
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

HEADER_SYNONYMS: Dict[str, Sequence[str]] = {
    "num_quarters_used": (
        "num quarters used",
        "number of quarters used",
        "num qtrs",
        "num_qtrs",
        "# qtrs",
        "# quarters",
        "n qtrs",
        "n quarters",
        "quarters used",
    ),
    "last_quarter_used": (
        "last quarter used",
        "last quarter",
        "latest quarter",
    ),
    "forecast_value": (
        "estimated total sold",
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "tot fcst without sa",
        "forecast total without sa",
        "total forecast",
        "forecast",
    ),
    "forecast_total_without_sa": (
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "tot fcst without sa",
        "forecast total without sa",
    ),
    "actual_value": (
        "actual value",
        "actual sales",
        "actual",
        "reported sales",
    ),
    "forecast_max": (
        "max",
        "forecast max",
        "upper",
    ),
    "forecast_min": (
        "min",
        "forecast min",
        "lower",
    ),
    "avg_penetration_pct": (
        "avg penetration %",
        "average penetration %",
        "avg penetration",
        "penetration %",
    ),
    "quarterly_sales": (
        "quarterly sales",
        "quarter sales",
    ),
    "reported_sales": (
        "reported sales",
    ),
    "growth_rate_pct": (
        "growth rate %",
        "growth %",
        "growth rate",
    ),
    "sales_captured_in_db_pct": (
        "sales captured in db %",
        "captured in db %",
        "sales captured in db",
    ),
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def safe_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("%", "")
    text = re.sub(r"^\$", "", text)
    try:
        return float(text)
    except ValueError:
        return None


def diff_values(lhs: Any, rhs: Any) -> Optional[float]:
    left = safe_float(lhs)
    right = safe_float(rhs)
    if left is None or right is None:
        return None
    return left - right


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1]
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,6}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0].strip()
    period_match = re.search(
        r"(?i)\b(early|mid|late)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(\d{4})\b",
        period_source,
    )

    model_period = ""
    model_date = ""
    if period_match:
        timing_key = period_match.group(1).lower()
        timing = timing_key.capitalize()
        month_abbr = period_match.group(2).lower()
        month_text = month_abbr.capitalize()
        year = int(period_match.group(3))
        day = {"early": 5, "mid": 15, "late": 25}[timing_key]
        month_num = MONTH_TO_NUM[month_abbr]
        model_period = f"{timing}{month_text}_{year}"
        model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def output_file_path(input_folder: Path, destination_dir: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    primary = destination_dir / f"{base_name}.xlsx"
    if not primary.exists():
        return primary

    suffix = 1
    while True:
        candidate = destination_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_workbook_no_save(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def set_formula2_r1c1(target_range: xw.Range, formula_r1c1: str) -> None:
    try:
        target_range.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    try:
        target_range.api.FormulaR1C1 = formula_r1c1
    except Exception:
        target_range.formula2 = formula_r1c1


def read_used_matrix(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    matrix = as_2d(used.value)
    return start_row, start_col, matrix


def find_max_anchor(
    start_row: int,
    start_col: int,
    matrix: List[List[Any]],
) -> Optional[Tuple[int, int]]:
    if not matrix:
        return None

    # Primary lookup: exact "max" label.
    for r_idx, row in enumerate(matrix):
        for c_idx, cell in enumerate(row):
            if normalize_text(cell) == "max":
                return start_row + r_idx, start_col + c_idx

    # Secondary lookup: any header containing token "max".
    for r_idx, row in enumerate(matrix):
        for c_idx, cell in enumerate(row):
            norm = normalize_text(cell)
            if norm.startswith("max ") or norm.endswith(" max") or " max " in f" {norm} ":
                return start_row + r_idx, start_col + c_idx
    return None


def locate_headers_near_anchor(
    start_row: int,
    start_col: int,
    matrix: List[List[Any]],
    anchor_row: int,
    anchor_col: int,
) -> Dict[str, int]:
    header_cols: Dict[str, int] = {}
    best_distance: Dict[str, int] = {}

    for r_idx, row in enumerate(matrix):
        abs_row = start_row + r_idx
        if abs(abs_row - anchor_row) > 2:
            continue
        for c_idx, cell in enumerate(row):
            abs_col = start_col + c_idx
            if abs(abs_col - anchor_col) > 40:
                continue
            normalized = normalize_text(cell)
            if not normalized:
                continue

            for key, synonyms in HEADER_SYNONYMS.items():
                if any(
                    normalized == normalize_text(synonym)
                    or normalize_text(synonym) in normalized
                    for synonym in synonyms
                ):
                    distance = abs(abs_col - anchor_col) + abs(abs_row - anchor_row)
                    if key not in best_distance or distance < best_distance[key]:
                        best_distance[key] = distance
                        header_cols[key] = abs_col
                    break

    return header_cols


def get_cell_value(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def r1c1_ref(row_offset: int, col_offset: int) -> str:
    row_ref = "R" if row_offset == 0 else f"R[{row_offset}]"
    col_ref = "C" if col_offset == 0 else f"C[{col_offset}]"
    return f"{row_ref}{col_ref}"


def discover_data_rows(sheet: xw.Sheet, anchor_row: int, anchor_col: int, n_rows: int) -> List[int]:
    block_top = anchor_row + 1
    block_bottom = anchor_row + 120
    left_col = max(1, anchor_col - 14)
    right_col = anchor_col + 14

    block_values = as_2d(sheet.range((block_top, left_col), (block_bottom, right_col)).value)
    rows: List[int] = []
    for idx, row_vals in enumerate(block_values):
        if any(value not in (None, "") for value in row_vals):
            rows.append(block_top + idx)
        if len(rows) >= n_rows:
            break

    next_row = anchor_row + 1
    while len(rows) < n_rows:
        if next_row not in rows:
            rows.append(next_row)
        next_row += 1

    rows.sort()
    return rows[:n_rows]


def rounded_key(value: Any) -> Any:
    numeric = safe_float(value)
    if numeric is not None:
        return round(numeric, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def extract_empirical_candidates(
    workbook: xw.Book,
    model_meta: Dict[str, str],
    source_file_name: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped file: {source_file_name} (reason: missing sheet 'Empirical Model')")
        return []

    used_row, used_col, matrix = read_used_matrix(sheet)
    anchor = find_max_anchor(used_row, used_col, matrix)
    if anchor is None:
        print(f"Skipped file: {source_file_name} (reason: no 'max' anchor in Empirical Model)")
        return []
    anchor_row, anchor_col = anchor

    header_cols = locate_headers_near_anchor(used_row, used_col, matrix, anchor_row, anchor_col)
    rows = discover_data_rows(sheet, anchor_row, anchor_col, N_QUARTERS)

    forecast_max_col = header_cols.get("forecast_max", anchor_col)
    forecast_min_col = header_cols.get("forecast_min", anchor_col + 1)
    forecast_col = header_cols.get("forecast_value", anchor_col - 1)
    actual_col = header_cols.get("actual_value", header_cols.get("reported_sales"))
    num_quarters_col = header_cols.get("num_quarters_used")
    last_quarter_col = header_cols.get("last_quarter_used")
    penetration_col = header_cols.get("avg_penetration_pct")
    quarterly_sales_col = header_cols.get("quarterly_sales")
    reported_sales_col = header_cols.get("reported_sales")
    growth_col = header_cols.get("growth_rate_pct")
    captured_col = header_cols.get("sales_captured_in_db_pct")

    scratch_col = max(sheet.used_range.last_cell.column + 3, anchor_col + 16)
    for idx, row in enumerate(rows, start=1):
        target = sheet.range((row, scratch_col))

        if penetration_col is not None:
            pen_offset = penetration_col - scratch_col
            if idx == 1:
                formula = f'=IFERROR({r1c1_ref(0, pen_offset)},"")'
            else:
                formula = (
                    f'=IFERROR(AVERAGE({r1c1_ref(-(idx - 1), pen_offset)}:'
                    f'{r1c1_ref(0, pen_offset)}),"")'
                )
        elif quarterly_sales_col is not None and reported_sales_col is not None:
            q_offset = quarterly_sales_col - scratch_col
            r_offset = reported_sales_col - scratch_col
            formula = f'=IFERROR({r1c1_ref(0, q_offset)}/{r1c1_ref(0, r_offset)},"")'
        else:
            formula = '=""'

        set_formula2_r1c1(target, formula)

    # One calculate pass after all formula updates.
    workbook.app.calculate()

    extracted: List[Dict[str, Any]] = []
    for idx, row in enumerate(rows, start=1):
        num_quarters = get_cell_value(sheet, row, num_quarters_col) if num_quarters_col else idx
        if num_quarters in (None, ""):
            num_quarters = idx

        forecast_value = get_cell_value(sheet, row, forecast_col)
        actual_value = get_cell_value(sheet, row, actual_col)
        forecast_max = get_cell_value(sheet, row, forecast_max_col)
        forecast_min = get_cell_value(sheet, row, forecast_min_col)
        avg_penetration = get_cell_value(sheet, row, scratch_col)
        quarterly_sales = get_cell_value(sheet, row, quarterly_sales_col)
        reported_sales = get_cell_value(sheet, row, reported_sales_col)
        growth_rate = get_cell_value(sheet, row, growth_col)
        sales_captured = get_cell_value(sheet, row, captured_col)

        if all(
            value in (None, "")
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        extracted.append(
            {
                "model": model_meta["model"],
                "ticker": model_meta["ticker"],
                "model_period": model_meta["model_period"],
                "model_date": model_meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": get_cell_value(sheet, row, last_quarter_col),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": diff_values(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file_name,
            }
        )

    if rows:
        sheet.range((rows[0], scratch_col), (rows[-1], scratch_col)).clear_contents()

    return extracted


def extract_regression_candidates(
    workbook: xw.Book,
    model_meta: Dict[str, str],
    source_file_name: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped file: {source_file_name} (reason: missing sheet 'Regression Model')")
        return []

    used_row, used_col, matrix = read_used_matrix(sheet)
    anchor = find_max_anchor(used_row, used_col, matrix)
    if anchor is None:
        print(f"Skipped file: {source_file_name} (reason: no 'max' anchor in Regression Model)")
        return []
    anchor_row, anchor_col = anchor

    header_cols = locate_headers_near_anchor(used_row, used_col, matrix, anchor_row, anchor_col)
    rows = discover_data_rows(sheet, anchor_row, anchor_col, N_QUARTERS)

    # Required by spec.
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    num_quarters_col = header_cols.get("num_quarters_used")
    forecast_total_col = header_cols.get("forecast_total_without_sa", header_cols.get("forecast_value", anchor_col - 1))
    actual_col = header_cols.get("actual_value")
    forecast_max_col = header_cols.get("forecast_max", anchor_col)
    forecast_min_col = header_cols.get("forecast_min", anchor_col + 1)

    scratch_col = max(sheet.used_range.last_cell.column + 4, anchor_col + 18)
    intercept_col = scratch_col
    slope_col = scratch_col + 1

    for idx, row in enumerate(rows, start=1):
        y_intercept_offset = y_col - intercept_col
        x_intercept_offset = x_col - intercept_col
        y_slope_offset = y_col - slope_col
        x_slope_offset = x_col - slope_col

        if idx == 1:
            y_intercept_range = f"{r1c1_ref(0, y_intercept_offset)}:{r1c1_ref(0, y_intercept_offset)}"
            x_intercept_range = f"{r1c1_ref(0, x_intercept_offset)}:{r1c1_ref(0, x_intercept_offset)}"
            y_slope_range = f"{r1c1_ref(0, y_slope_offset)}:{r1c1_ref(0, y_slope_offset)}"
            x_slope_range = f"{r1c1_ref(0, x_slope_offset)}:{r1c1_ref(0, x_slope_offset)}"
        else:
            y_intercept_range = (
                f"{r1c1_ref(-(idx - 1), y_intercept_offset)}:{r1c1_ref(0, y_intercept_offset)}"
            )
            x_intercept_range = (
                f"{r1c1_ref(-(idx - 1), x_intercept_offset)}:{r1c1_ref(0, x_intercept_offset)}"
            )
            y_slope_range = f"{r1c1_ref(-(idx - 1), y_slope_offset)}:{r1c1_ref(0, y_slope_offset)}"
            x_slope_range = f"{r1c1_ref(-(idx - 1), x_slope_offset)}:{r1c1_ref(0, x_slope_offset)}"

        intercept_formula = f'=IFERROR(INTERCEPT({y_intercept_range},{x_intercept_range}),"")'
        slope_formula = f'=IFERROR(SLOPE({y_slope_range},{x_slope_range}),"")'

        set_formula2_r1c1(sheet.range((row, intercept_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((row, slope_col)), slope_formula)

    # One calculate pass after all formula updates.
    workbook.app.calculate()

    extracted: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Any, ...]] = None

    for idx, row in enumerate(rows, start=1):
        num_quarters = get_cell_value(sheet, row, num_quarters_col) if num_quarters_col else idx
        if num_quarters in (None, ""):
            num_quarters = idx

        forecast_value = get_cell_value(sheet, row, forecast_total_col)
        actual_value = get_cell_value(sheet, row, actual_col)
        forecast_max = get_cell_value(sheet, row, forecast_max_col)
        forecast_min = get_cell_value(sheet, row, forecast_min_col)
        intercept = get_cell_value(sheet, row, intercept_col)
        slope = get_cell_value(sheet, row, slope_col)

        if all(
            value in (None, "")
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        current_key = (
            rounded_key(num_quarters),
            rounded_key(forecast_value),
            rounded_key(forecast_max),
            rounded_key(forecast_min),
            rounded_key(intercept),
            rounded_key(slope),
        )
        is_last = idx == len(rows)
        if is_last and previous_key is not None and current_key == previous_key:
            continue

        extracted.append(
            {
                "model": model_meta["model"],
                "ticker": model_meta["ticker"],
                "model_period": model_meta["model_period"],
                "model_date": model_meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": diff_values(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file_name,
            }
        )
        previous_key = current_key

    if rows:
        sheet.range((rows[0], intercept_col), (rows[-1], slope_col)).clear_contents()

    return extracted


def write_sheet(
    worksheet,
    columns: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for header_cell in worksheet[1]:
        header_cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_index, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for value, in worksheet.iter_rows(
            min_row=2, max_row=worksheet.max_row, min_col=col_index, max_col=col_index, values_only=True
        ):
            if value is None:
                continue
            if isinstance(value, date):
                text = value.isoformat()
            else:
                text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        worksheet.column_dimensions[get_column_letter(col_index)].width = min(48, max(12, max_len + 2))


def write_output_workbook(
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
    destination_path: Path,
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = workbook.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(destination_path)


def collect_source_files(folder: Path) -> List[Path]:
    files: List[Path] = []
    for file_path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if file_path.name.startswith("~"):
            print(f"Skipped file: {file_path.name} (reason: temp file)")
            continue
        if not file_path.is_file():
            print(f"Skipped file: {file_path.name} (reason: not a file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {file_path.name} (reason: not .xlsx)")
            continue
        files.append(file_path)
    return files


def main() -> None:
    if not input_dir.exists():
        print(f"Skipped file: {input_dir} (reason: input directory does not exist)")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    source_files = collect_source_files(input_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    if source_files:
        app: Optional[xw.App] = None
        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            for file_path in source_files:
                workbook: Optional[xw.Book] = None
                try:
                    workbook = app.books.open(str(file_path), update_links=False)
                    model_meta = parse_file_label(file_path.name)
                    empirical_rows.extend(
                        extract_empirical_candidates(workbook, model_meta=model_meta, source_file_name=file_path.name)
                    )
                    regression_rows.extend(
                        extract_regression_candidates(workbook, model_meta=model_meta, source_file_name=file_path.name)
                    )
                    files_processed += 1
                    print(f"Processed file: {file_path.name}")
                except Exception as exc:
                    print(f"Skipped file: {file_path.name} (reason: {exc})")
                finally:
                    if workbook is not None:
                        close_workbook_no_save(workbook)
        finally:
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    app.kill()
    else:
        print(f"Skipped file: {input_dir} (reason: no .xlsx files found)")

    output_path = output_file_path(input_dir, output_dir)
    write_output_workbook(empirical_rows, regression_rows, output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
