#!/usr/bin/env python3
"""
Build empirical/regression parameter candidates from a folder of .xlsx files.

This script opens each source workbook only once, processes both model sheets
while it is open, and writes a single output workbook with two sheets:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# Configure paths here
# -----------------------------
input_dir = "input"
output_dir = "output"


EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "" or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def round_key(value: Optional[float], decimals: int = 8) -> Optional[float]:
    if value is None:
        return None
    return round(value, decimals)


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def parse_file_labels(file_name: str) -> Optional[FileLabels]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].upper()
    period_token = re.sub(r"(?i)_send$", "", parts[2]).strip()
    match = re.search(r"(?i)^(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})$", period_token)
    if not match:
        return None

    timing = match.group(1).lower()
    month_raw = match.group(2)
    year = int(match.group(3))
    month_key = month_raw[:3].lower()
    month = MONTH_MAP.get(month_key)
    if month is None:
        return None

    month_short = month_raw[:3].title()
    model_period = f"{timing.title()}{month_short}_{year}"
    model_date = date(year, month, DAY_MAP[timing]).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def candidate_output_path(output_path: Path, input_path: Path) -> Path:
    input_folder_name = input_path.name
    base_name = f"{input_folder_name}_PARAM.xlsx"
    first_candidate = output_path / base_name
    if not first_candidate.exists():
        return first_candidate

    index = 1
    while True:
        candidate = output_path / f"{input_folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def get_sheet_by_name(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = normalize_2d(used.value)
    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == "max":
                return used.row + row_idx, used.column + col_idx
    return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    # Prefer Formula2 R1C1; fallback to older formula APIs if needed.
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

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
        workbook.api.Close(SaveChanges=False)
        return
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


def collect_numeric_rows(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    x_col: int,
    y_col: int,
) -> List[int]:
    rows: List[int] = []
    if end_row < start_row:
        return rows

    x_values = normalize_2d(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_values = normalize_2d(sheet.range((start_row, y_col), (end_row, y_col)).value)

    for idx in range(min(len(x_values), len(y_values))):
        x_val = to_float(x_values[idx][0] if x_values[idx] else None)
        y_val = to_float(y_values[idx][0] if y_values[idx] else None)
        if x_val is not None and y_val is not None:
            rows.append(start_row + idx)
    return rows


def build_empirical_rows(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical for {source_file}: 'Empirical Model' sheet not found")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipped empirical for {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    quarter_col = anchor_col - 12
    reported_col = anchor_col - 6
    growth_col = anchor_col - 5

    start_row = max(2, anchor_row - 450)
    end_row = max(start_row, anchor_row - 1)
    numeric_rows = collect_numeric_rows(sheet, start_row, end_row, x_col, y_col)
    if not numeric_rows:
        print(f"Skipped empirical for {source_file}: no numeric history rows")
        return []

    numeric_rows = numeric_rows[-10:]
    max_n = min(10, len(numeric_rows))

    scratch_row = max(sheet.used_range.last_cell.row + 2, anchor_row + 2)
    avg_cell = sheet.cells(scratch_row, anchor_col + 1)
    forecast_cell = sheet.cells(scratch_row, anchor_col + 2)
    max_cell = sheet.cells(scratch_row, anchor_col + 3)
    min_cell = sheet.cells(scratch_row, anchor_col + 4)

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, max_n + 1):
        subset_rows = numeric_rows[-n_quarters:]
        subset_start = subset_rows[0]
        subset_end = subset_rows[-1]
        latest_row = subset_end

        reported_value = to_float(sheet.cells(latest_row, reported_col).value)
        reported_ref_col = reported_col if reported_value is not None else y_col

        avg_formula = f"=AVERAGE(R{subset_start}C{x_col}:R{subset_end}C{x_col})"
        forecast_formula = (
            f"=IFERROR(R{latest_row}C{reported_ref_col}/R{avg_cell.row}C{avg_cell.column},NA())"
        )
        max_formula = (
            f"=IFERROR(R{latest_row}C{reported_ref_col}/MIN(R{subset_start}C{x_col}:R{subset_end}C{x_col}),NA())"
        )
        min_formula = (
            f"=IFERROR(R{latest_row}C{reported_ref_col}/MAX(R{subset_start}C{x_col}:R{subset_end}C{x_col}),NA())"
        )

        set_formula2_r1c1(avg_cell, avg_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)
        set_formula2_r1c1(max_cell, max_formula)
        set_formula2_r1c1(min_cell, min_formula)
        workbook.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        forecast_value = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        quarterly_sales = to_float(sheet.cells(latest_row, y_col).value)
        reported_sales = to_float(sheet.cells(latest_row, reported_col).value)
        if reported_sales is None:
            reported_sales = quarterly_sales

        growth_rate = to_float(sheet.cells(latest_row, growth_col).value)
        sales_captured = to_float(sheet.cells(latest_row, x_col).value)
        last_quarter_used = sheet.cells(latest_row, quarter_col).value

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )
    return rows


def build_regression_rows(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Regression Model")
    if sheet is None:
        print(f"Skipped regression for {source_file}: 'Regression Model' sheet not found")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipped regression for {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    start_row = max(2, anchor_row - 450)
    end_row = max(start_row, anchor_row - 1)
    numeric_rows = collect_numeric_rows(sheet, start_row, end_row, x_col, y_col)
    if len(numeric_rows) < 2:
        print(f"Skipped regression for {source_file}: not enough numeric history rows")
        return []

    numeric_rows = numeric_rows[-10:]
    max_n = min(10, len(numeric_rows))
    if max_n < 2:
        return []

    scratch_row = max(sheet.used_range.last_cell.row + 2, anchor_row + 2)
    intercept_cell = sheet.cells(scratch_row, anchor_col + 1)
    slope_cell = sheet.cells(scratch_row, anchor_col + 2)
    forecast_cell = sheet.cells(scratch_row, anchor_col + 3)
    max_cell = sheet.cells(scratch_row, anchor_col + 4)
    min_cell = sheet.cells(scratch_row, anchor_col + 5)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for n_quarters in range(2, max_n + 1):
        subset_rows = numeric_rows[-n_quarters:]
        subset_start = subset_rows[0]
        subset_end = subset_rows[-1]

        predictor_row = subset_end + 1
        predictor_x = to_float(sheet.cells(predictor_row, x_col).value)
        if predictor_x is None:
            predictor_row = subset_end
            predictor_x = to_float(sheet.cells(predictor_row, x_col).value)

        intercept_formula = (
            f"=INTERCEPT(R{subset_start}C{y_col}:R{subset_end}C{y_col},"
            f"R{subset_start}C{x_col}:R{subset_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{subset_start}C{y_col}:R{subset_end}C{y_col},"
            f"R{subset_start}C{x_col}:R{subset_end}C{x_col})"
        )
        forecast_formula = (
            f"=R{intercept_cell.row}C{intercept_cell.column}"
            f"+R{slope_cell.row}C{slope_cell.column}*R{predictor_row}C{x_col}"
        )
        max_formula = f"=MAX(R{subset_start}C{y_col}:R{subset_end}C{y_col})"
        min_formula = f"=MIN(R{subset_start}C{y_col}:R{subset_end}C{y_col})"

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)
        set_formula2_r1c1(max_cell, max_formula)
        set_formula2_r1c1(min_cell, min_formula)
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )
        actual_value = to_float(sheet.cells(predictor_row, y_col).value)

        signature = (
            round_key(forecast),
            round_key(intercept),
            round_key(slope),
            round_key(forecast_max),
            round_key(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast,
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


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    sheet = workbook.create_sheet(sheet_name)
    sheet.append(list(columns))
    header_font = Font(bold=True)
    for cell in sheet[1]:
        cell.font = header_font

    widths = [len(col) + 2 for col in columns]
    for row in rows:
        values = [row.get(col) for col in columns]
        sheet.append(values)
        for idx, value in enumerate(values):
            text = "" if value is None else str(value)
            widths[idx] = min(60, max(widths[idx], len(text) + 2))

    sheet.freeze_panes = "A2"
    last_col = get_column_letter(len(columns))
    sheet.auto_filter.ref = f"A1:{last_col}{max(1, sheet.max_row)}"

    for idx, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(idx)].width = width


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_file)


def process_workbooks(input_path: Path) -> Tuple[int, List[Dict[str, Any]], List[Dict[str, Any]]]:
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            labels = parse_file_labels(file_path.name)
            if labels is None:
                print(f"Skipped {file_path.name}: filename pattern not recognized")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_empirical = build_empirical_rows(workbook, labels, file_path.name)
                file_regression = build_regression_rows(workbook, labels, file_path.name)
                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                processed_files += 1
                print(
                    f"Processed {file_path.name}: "
                    f"empirical_rows={len(file_empirical)}, regression_rows={len(file_regression)}"
                )
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    safe_close_source_workbook(workbook)
    finally:
        app.quit()

    return processed_files, empirical_rows, regression_rows


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    processed_files, empirical_rows, regression_rows = process_workbooks(input_path)
    output_file = candidate_output_path(output_path, input_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output written to: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
