#!/usr/bin/env python3
"""
Batch extraction script for empirical and regression model candidates.

This script reads all eligible .xlsx workbooks from input_dir, opens each source
workbook exactly once, extracts both model outputs while the workbook is open,
then closes it without saving. It writes a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import datetime as dt
import math
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
# Update these two paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"
N_QUARTERS = 10

FILE_LABEL_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
    r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})(?:_|$)",
    re.IGNORECASE,
)

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        numeric = float(value)
        if math.isnan(numeric) or math.isinf(numeric):
            return None
        return numeric
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        is_percent = text.endswith("%")
        if is_percent:
            text = text[:-1]
        text = text.replace(",", "")
        try:
            numeric = float(text)
        except ValueError:
            return None
        return numeric / 100.0 if is_percent else numeric
    return None


def almost_equal(left: Optional[float], right: Optional[float], tol: float = 1e-9) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return abs(left - right) <= tol


def as_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first_item = values[0]
        if isinstance(first_item, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def parse_file_metadata(file_path: Path) -> Optional[Dict[str, str]]:
    """
    Parse filename segments like:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    match = FILE_LABEL_PATTERN.search(file_path.stem)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    phase = match.group("phase").capitalize()
    month_token = match.group("month")[:3].title()
    year = int(match.group("year"))

    try:
        month_number = dt.datetime.strptime(month_token, "%b").month
    except ValueError:
        return None

    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_lookup[phase]
    model_period = f"{phase}{month_token}_{year}"
    model_date = dt.date(year, month_number, model_day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_folder: Path, destination_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM.xlsx"
    output_path = destination_folder / base_name
    if not output_path.exists():
        return output_path

    suffix = 1
    while True:
        candidate = destination_folder / f"{input_folder.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def should_skip_file(file_path: Path, input_folder_name: str) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.match(
        rf"^{re.escape(input_folder_name)}_PARAM(?:\.\d+)?\.xlsx$",
        file_path.name,
        flags=re.IGNORECASE,
    ):
        return "prior output workbook"
    return None


def collect_source_files(input_folder: Path) -> List[Path]:
    source_files: List[Path] = []
    for file_path in sorted(input_folder.iterdir()):
        skip_reason = should_skip_file(file_path, input_folder.name)
        if skip_reason:
            print(f"Skipped file: {file_path.name} ({skip_reason})")
            continue
        source_files.append(file_path)
    return source_files


def get_sheet_by_name(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = normalize_text(sheet_name)
    for sheet in workbook.sheets:
        if normalize_text(sheet.name) == target:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = as_2d(used.value)
    start_row = used.row
    start_col = used.column
    target = anchor_label.strip().lower()

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return start_row + r_idx, start_col + c_idx
    return None


def build_header_map(
    sheet: xw.Sheet,
    header_row: int,
    left_col: int,
    right_col: int,
) -> Dict[str, int]:
    header_map: Dict[str, int] = {}
    for col in range(left_col, right_col + 1):
        label = normalize_text(sheet.cells(header_row, col).value)
        if label and label not in header_map:
            header_map[label] = col
    return header_map


def find_column(header_map: Dict[str, int], aliases: Sequence[str]) -> Optional[int]:
    normalized_aliases = [normalize_text(alias) for alias in aliases if alias]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for alias in normalized_aliases:
        for header, col in header_map.items():
            if alias in header:
                return col
    return None


def get_cell_value(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def get_numeric_history_rows(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    end_row: int,
) -> List[int]:
    rows: List[int] = []
    row = end_row
    while row >= 1:
        x_val = to_float(sheet.cells(row, x_col).value)
        y_val = to_float(sheet.cells(row, y_col).value)
        if x_val is not None and y_val is not None:
            rows.append(row)
            row -= 1
            continue

        if rows:
            break
        row -= 1

    rows.reverse()
    return rows


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def row_is_duplicate_for_regression(current_row: Dict[str, Any], previous_row: Dict[str, Any]) -> bool:
    numeric_keys = [
        "parameter_value",
        "forecast_value",
        "actual_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    for key in numeric_keys:
        if not almost_equal(to_float(current_row.get(key)), to_float(previous_row.get(key))):
            return False
    return True


def process_empirical_sheet(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_by_name(workbook, EMPIRICAL_MODEL_SHEET)
    if sheet is None:
        print(f"Skipped file: {source_file} ({EMPIRICAL_MODEL_SHEET} sheet not found)")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped file: {source_file} ({EMPIRICAL_MODEL_SHEET} max anchor not found)")
        return rows

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    left_col = used.column
    right_col = used.last_cell.column
    header_map = build_header_map(sheet, anchor_row, left_col, right_col)

    num_quarters_col = find_column(
        header_map,
        ["num quarters used", "num quarters", "quarters used", "n quarters", "# quarters"],
    )
    last_quarter_col = find_column(header_map, ["last quarter used", "last quarter"])
    forecast_value_col = find_column(
        header_map,
        ["estimated total sold", "forecast value", "forecast", "total forecast"],
    )
    actual_value_col = find_column(header_map, ["actual value", "reported sales", "actual"])
    max_col = find_column(header_map, ["forecast max", "max"]) or anchor_col
    min_col = find_column(header_map, ["forecast min", "min"]) or (anchor_col + 1)
    avg_pen_col = find_column(
        header_map,
        ["avg penetration pct", "avg penetration", "average penetration"],
    )
    quarterly_sales_col = find_column(
        header_map,
        ["quarterly sales", "quarter sales", "sales in db", "sales db"],
    )
    reported_sales_col = find_column(header_map, ["reported sales"])
    growth_rate_col = find_column(header_map, ["growth rate pct", "growth rate"])
    captured_col = find_column(
        header_map,
        ["sales captured in db pct", "captured in db pct", "sales captured"],
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = anchor_row - 1
    history_rows: List[int] = []
    if x_col >= 1 and y_col >= 1 and history_end_row >= 1:
        history_rows = get_numeric_history_rows(sheet, x_col, y_col, history_end_row)

    max_quarters = min(N_QUARTERS, len(history_rows)) if history_rows else N_QUARTERS
    helper_col = max(right_col + 2, anchor_col + 2)
    avg_pen_cell = sheet.cells(anchor_row, helper_col)

    latest_x = to_float(sheet.cells(history_rows[-1], x_col).value) if history_rows else None
    latest_y = to_float(sheet.cells(history_rows[-1], y_col).value) if history_rows else None

    for n_quarters in range(1, max_quarters + 1):
        candidate_row = anchor_row + n_quarters
        num_quarters_value = to_float(get_cell_value(sheet, candidate_row, num_quarters_col))
        num_quarters_used = int(num_quarters_value) if num_quarters_value is not None else n_quarters

        start_history_row = history_rows[-n_quarters] if history_rows and n_quarters <= len(history_rows) else None
        end_history_row = history_rows[-1] if history_rows else None

        avg_penetration = to_float(get_cell_value(sheet, candidate_row, avg_pen_col))
        if avg_penetration is None and start_history_row and end_history_row:
            avg_formula = (
                "=AVERAGE("
                f"IFERROR(R{start_history_row}C{y_col}:R{end_history_row}C{y_col}/"
                f"R{start_history_row}C{x_col}:R{end_history_row}C{x_col},\"\"))"
            )
            set_formula2(avg_pen_cell, avg_formula)
            workbook.app.calculate()
            avg_penetration = to_float(avg_pen_cell.value)

        quarterly_sales = to_float(get_cell_value(sheet, candidate_row, quarterly_sales_col))
        if quarterly_sales is None:
            quarterly_sales = latest_x

        reported_sales = to_float(get_cell_value(sheet, candidate_row, reported_sales_col))
        if reported_sales is None:
            reported_sales = latest_y

        forecast_value = to_float(get_cell_value(sheet, candidate_row, forecast_value_col))
        if forecast_value is None and quarterly_sales is not None and avg_penetration not in (None, 0.0):
            forecast_value = quarterly_sales / avg_penetration

        actual_value = to_float(get_cell_value(sheet, candidate_row, actual_value_col))
        if actual_value is None:
            actual_value = reported_sales

        forecast_max = to_float(get_cell_value(sheet, candidate_row, max_col))
        forecast_min = to_float(get_cell_value(sheet, candidate_row, min_col))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        growth_rate = to_float(get_cell_value(sheet, candidate_row, growth_rate_col))
        captured_pct = to_float(get_cell_value(sheet, candidate_row, captured_col))
        last_quarter_used = get_cell_value(sheet, candidate_row, last_quarter_col)
        if (last_quarter_used is None or str(last_quarter_used).strip() == "") and start_history_row:
            last_quarter_used = sheet.cells(start_history_row, max(1, x_col - 1)).value

        has_data = any(
            value is not None
            for value in [
                avg_penetration,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            ]
        )
        if not has_data and not history_rows:
            break

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": captured_pct,
            "source_file": source_file,
        }
        rows.append(row)

    avg_pen_cell.value = None
    return rows


def process_regression_sheet(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_by_name(workbook, REGRESSION_MODEL_SHEET)
    if sheet is None:
        print(f"Skipped file: {source_file} ({REGRESSION_MODEL_SHEET} sheet not found)")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped file: {source_file} ({REGRESSION_MODEL_SHEET} max anchor not found)")
        return rows

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    left_col = used.column
    right_col = used.last_cell.column
    header_map = build_header_map(sheet, anchor_row, left_col, right_col)

    num_quarters_col = find_column(
        header_map,
        ["num quarters used", "num quarters", "quarters used", "n quarters", "# quarters"],
    )
    forecast_value_col = find_column(
        header_map,
        ["tot fcst w o sa", "tot fcst w/o sa", "tot fcst without sa", "forecast value", "forecast"],
    )
    actual_value_col = find_column(header_map, ["actual value", "reported sales", "actual"])
    max_col = find_column(header_map, ["forecast max", "max"]) or anchor_col
    min_col = find_column(header_map, ["forecast min", "min"]) or (anchor_col + 1)
    intercept_col = find_column(header_map, ["intercept"])
    slope_col = find_column(header_map, ["slope"])

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = anchor_row - 1
    history_rows: List[int] = []
    if x_col >= 1 and y_col >= 1 and history_end_row >= 1:
        history_rows = get_numeric_history_rows(sheet, x_col, y_col, history_end_row)

    max_quarters = min(N_QUARTERS, len(history_rows)) if history_rows else N_QUARTERS
    helper_col = max(right_col + 2, anchor_col + 2)
    intercept_cell = sheet.cells(anchor_row, helper_col)
    slope_cell = sheet.cells(anchor_row, helper_col + 1)

    latest_x = to_float(sheet.cells(history_rows[-1], x_col).value) if history_rows else None

    for n_quarters in range(1, max_quarters + 1):
        candidate_row = anchor_row + n_quarters
        num_quarters_value = to_float(get_cell_value(sheet, candidate_row, num_quarters_col))
        num_quarters_used = int(num_quarters_value) if num_quarters_value is not None else n_quarters

        intercept_value = to_float(get_cell_value(sheet, candidate_row, intercept_col))
        slope_value = to_float(get_cell_value(sheet, candidate_row, slope_col))

        start_history_row = history_rows[-n_quarters] if history_rows and n_quarters <= len(history_rows) else None
        end_history_row = history_rows[-1] if history_rows else None
        if (intercept_value is None or slope_value is None) and start_history_row and end_history_row:
            intercept_formula = (
                f"=INTERCEPT(R{start_history_row}C{y_col}:R{end_history_row}C{y_col},"
                f"R{start_history_row}C{x_col}:R{end_history_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{start_history_row}C{y_col}:R{end_history_row}C{y_col},"
                f"R{start_history_row}C{x_col}:R{end_history_row}C{x_col})"
            )
            set_formula2(intercept_cell, intercept_formula)
            set_formula2(slope_cell, slope_formula)
            workbook.app.calculate()
            intercept_value = to_float(intercept_cell.value)
            slope_value = to_float(slope_cell.value)

        forecast_value = to_float(get_cell_value(sheet, candidate_row, forecast_value_col))
        if (
            forecast_value is None
            and intercept_value is not None
            and slope_value is not None
            and latest_x is not None
        ):
            forecast_value = intercept_value + (slope_value * latest_x)

        actual_value = to_float(get_cell_value(sheet, candidate_row, actual_value_col))
        forecast_max = to_float(get_cell_value(sheet, candidate_row, max_col))
        forecast_min = to_float(get_cell_value(sheet, candidate_row, min_col))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        has_data = any(
            value is not None
            for value in [
                intercept_value,
                slope_value,
                forecast_value,
                forecast_max,
                forecast_min,
            ]
        )
        if not has_data and not history_rows:
            break

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        # Prevent duplicate final row if the workbook emits a repeated terminal line.
        if n_quarters == max_quarters and rows and row_is_duplicate_for_regression(row, rows[-1]):
            continue
        rows.append(row)

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def clean_for_excel(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    return value


def write_output_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(sheet_name)
    ws.append(list(columns))

    for row in rows:
        ws.append([clean_for_excel(row.get(col)) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(ws.max_row, 1)}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_output_sheet(workbook, EMPIRICAL_OUTPUT_SHEET, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(workbook, REGRESSION_OUTPUT_SHEET, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def run_extraction() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = build_output_path(input_dir.resolve(), output_dir.resolve())
    source_files = collect_source_files(input_dir.resolve())

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.api.EnableEvents = False
        except Exception:
            pass
        try:
            # Excel constant xlCalculationManual = -4135
            app.api.Calculation = -4135
        except Exception:
            pass

        for file_path in source_files:
            metadata = parse_file_metadata(file_path)
            if metadata is None:
                print(f"Skipped file: {file_path.name} (filename format is not recognized)")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, metadata, file_path.name))
                files_processed += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_source_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run_extraction()
