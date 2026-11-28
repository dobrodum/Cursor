#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from .xlsx workbooks.

The script opens each source workbook once, processes both the "Empirical Model"
and "Regression Model" sheets while the workbook is open, and writes one output
workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# =========================
# User-configurable inputs
# =========================
input_dir = Path(r"./input")
output_dir = Path(r"./output")

N_QUARTERS = 10

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

# Anchor-relative offsets for row-level values.
# These preserve an anchor-based approach and are easy to adjust if model layout changes.
EMPIRICAL_OFFSETS = {
    "row_start": 1,  # first candidate row starts below the "max" anchor row
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "forecast_value": -1,  # estimated total sold
    "actual_value": 2,  # reported sales
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -4,
    "reported_sales": 2,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -2,
    # Most recent penetration input value in the row; earlier values are to the left.
    "penetration_tail_col": -12,
}

REGRESSION_OFFSETS = {
    "row_start": 1,  # first candidate row starts below the "max" anchor row
    "num_quarters_used": -10,
    "forecast_value": -1,  # TOT FCST w/o SA
    "actual_value": 2,  # optional in source; can be blank
    "forecast_max": 0,
    "forecast_min": 1,
}

TIMING_DAY = {"early": 5, "mid": 15, "late": 25}

MONTH_MAP = {
    "jan": (1, "Jan"),
    "january": (1, "Jan"),
    "feb": (2, "Feb"),
    "february": (2, "Feb"),
    "mar": (3, "Mar"),
    "march": (3, "Mar"),
    "apr": (4, "Apr"),
    "april": (4, "Apr"),
    "may": (5, "May"),
    "jun": (6, "Jun"),
    "june": (6, "Jun"),
    "jul": (7, "Jul"),
    "july": (7, "Jul"),
    "aug": (8, "Aug"),
    "august": (8, "Aug"),
    "sep": (9, "Sep"),
    "sept": (9, "Sep"),
    "september": (9, "Sep"),
    "oct": (10, "Oct"),
    "october": (10, "Oct"),
    "nov": (11, "Nov"),
    "november": (11, "Nov"),
    "dec": (12, "Dec"),
    "december": (12, "Dec"),
}


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetSnapshot:
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    values: list[list[Any]]


def to_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], (list, tuple)):
        return [list(values)]
    return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]


def get_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = to_matrix(used.value)

    if not values:
        return SheetSnapshot(
            first_row=used.row,
            first_col=used.column,
            last_row=used.row,
            last_col=used.column,
            values=[],
        )

    max_cols = max((len(row) for row in values), default=0)
    row_count = len(values)

    return SheetSnapshot(
        first_row=used.row,
        first_col=used.column,
        last_row=used.row + row_count - 1,
        last_col=used.column + max_cols - 1,
        values=values,
    )


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> Optional[tuple[int, int]]:
    target = anchor_text.strip().lower()
    for r_idx, row in enumerate(snapshot.values):
        for c_idx, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip().lower() == target:
                return snapshot.first_row + r_idx, snapshot.first_col + c_idx
    return None


def value_at(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def to_float(value: Any) -> Optional[float]:
    if value is None:
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

    pct = text.endswith("%")
    clean = text.replace(",", "").replace("%", "")
    try:
        number = float(clean)
    except ValueError:
        return None
    return number / 100.0 if pct else number


def to_int(value: Any, fallback: Optional[int] = None) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return fallback
    return int(round(number))


def output_value(value: Any) -> Any:
    if value is None:
        return ""
    return value


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    formula = formula_r1c1 if formula_r1c1.startswith("=") else f"={formula_r1c1}"
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass

    # Fallback if Formula2R1C1 is unavailable in a given Excel build.
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass

    cell.api.FormulaR1C1 = formula


def close_workbook_no_save(wb: xw.Book) -> None:
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

    app = wb.app
    previous_alerts = app.display_alerts
    app.display_alerts = False
    try:
        wb.close()
    finally:
        app.display_alerts = previous_alerts


def parse_labels(file_path: Path) -> FileLabels:
    # Expected example: MedMiner_Model - AORT - MidJan2026_Send.xlsx
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("filename does not contain expected ' - ' sections")

    ticker = parts[1].upper()
    period_token = parts[2].split("_", 1)[0]
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token, flags=re.IGNORECASE)
    if not match:
        raise ValueError(f"cannot parse model period token '{period_token}'")

    timing_raw, month_raw, year_text = match.groups()
    timing_key = timing_raw.lower()
    if timing_key not in TIMING_DAY:
        raise ValueError(f"unsupported timing token '{timing_raw}'")

    month_key = month_raw.lower()
    if month_key not in MONTH_MAP:
        raise ValueError(f"unsupported month token '{month_raw}'")

    month_num, month_abbrev = MONTH_MAP[month_key]
    year = int(year_text)
    day = TIMING_DAY[timing_key]

    model_period = f"{timing_raw.capitalize()}{month_abbrev}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def file_skip_reason(file_path: Path, output_base_name: str) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary workbook"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if file_path.stem.startswith(output_base_name):
        return "looks like prior output"
    return None


def compute_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def extract_empirical_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical for {source_file}: missing 'Empirical Model' sheet")
        return []

    snapshot = get_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped empirical for {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + EMPIRICAL_OFFSETS["row_start"]
    scratch_col = snapshot.last_col + 4

    scratch_cells: list[xw.Range] = []
    for quarter_count in range(1, N_QUARTERS + 1):
        row = start_row + (quarter_count - 1)

        pen_tail_col = anchor_col + EMPIRICAL_OFFSETS["penetration_tail_col"]
        pen_start_col = pen_tail_col - (quarter_count - 1)
        if pen_start_col < 1 or pen_tail_col < 1:
            continue

        avg_pen_cell = sheet.range((row, scratch_col))
        avg_pen_formula = (
            f'=IFERROR(AVERAGE(R{row}C{pen_start_col}:R{row}C{pen_tail_col}),"")'
        )
        set_formula2_r1c1(avg_pen_cell, avg_pen_formula)
        scratch_cells.append(avg_pen_cell)

    if scratch_cells:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for quarter_count in range(1, N_QUARTERS + 1):
        row = start_row + (quarter_count - 1)

        num_quarters_used = to_int(
            value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"]),
            fallback=quarter_count,
        )
        last_quarter_used = value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"])
        forecast_value = to_float(value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_value"]))
        actual_value = to_float(value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["actual_value"]))
        forecast_max = to_float(value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_max"]))
        forecast_min = to_float(value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_min"]))
        avg_penetration_pct = to_float(value_at(sheet, row, scratch_col))
        quarterly_sales = to_float(value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]))
        reported_sales = to_float(value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["reported_sales"]))
        growth_rate_pct = to_float(value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]))
        sales_captured_in_db_pct = to_float(
            value_at(sheet, row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"])
        )

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": output_value(last_quarter_used),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    for cell in scratch_cells:
        cell.value = None

    return rows


def rows_match(lhs: dict[str, Any], rhs: dict[str, Any]) -> bool:
    keys = ("num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    for key in keys:
        left_value = lhs.get(key)
        right_value = rhs.get(key)

        if left_value is None and right_value is None:
            continue
        if isinstance(left_value, (int, float)) and isinstance(right_value, (int, float)):
            if abs(float(left_value) - float(right_value)) > 1e-9:
                return False
            continue
        if left_value != right_value:
            return False
    return True


def extract_regression_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression for {source_file}: missing 'Regression Model' sheet")
        return []

    snapshot = get_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped regression for {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + REGRESSION_OFFSETS["row_start"]
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    if y_col < 1 or x_col < 1:
        print(f"Skipped regression for {source_file}: invalid x/y source columns from anchor")
        return []

    scratch_intercept_col = snapshot.last_col + 4
    scratch_slope_col = scratch_intercept_col + 1
    scratch_forecast_col = scratch_intercept_col + 2

    history_end_row = anchor_row - 1
    formula_rows: list[tuple[int, int]] = []
    for quarter_count in range(1, N_QUARTERS + 1):
        history_start_row = history_end_row - quarter_count + 1
        if history_start_row < 1:
            continue

        row = start_row + (quarter_count - 1)
        intercept_cell = sheet.range((row, scratch_intercept_col))
        slope_cell = sheet.range((row, scratch_slope_col))
        forecast_cell = sheet.range((row, scratch_forecast_col))

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{history_start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{history_start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        forecast_formula = (
            f'=IFERROR(R{row}C{scratch_intercept_col}+R{row}C{scratch_slope_col}*R{anchor_row}C{x_col},"")'
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)

        formula_rows.append((quarter_count, row))

    if formula_rows:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_row: Optional[dict[str, Any]] = None

    for quarter_count, row in formula_rows:
        num_quarters_used = to_int(
            value_at(sheet, row, anchor_col + REGRESSION_OFFSETS["num_quarters_used"]),
            fallback=quarter_count,
        )
        forecast_value = to_float(value_at(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_value"]))
        if forecast_value is None:
            forecast_value = to_float(value_at(sheet, row, scratch_forecast_col))

        actual_value = value_at(sheet, row, anchor_col + REGRESSION_OFFSETS["actual_value"])
        if actual_value == "":
            actual_value = None
        actual_value = to_float(actual_value) if actual_value is not None else None

        forecast_max = to_float(value_at(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_max"]))
        forecast_min = to_float(value_at(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_min"]))
        intercept = to_float(value_at(sheet, row, scratch_intercept_col))
        slope = to_float(value_at(sheet, row, scratch_slope_col))

        if all(value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row_payload = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        # Suppress duplicated trailing row by comparing against prior calculated row.
        if previous_row is not None and rows_match(row_payload, previous_row):
            continue

        rows.append(row_payload)
        previous_row = row_payload

    for _, row in formula_rows:
        sheet.range((row, scratch_intercept_col), (row, scratch_forecast_col)).value = [[None, None, None]]

    return rows


def choose_output_path(input_folder: Path, destination_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = destination_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = destination_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def write_sheet(
    workbook: Workbook,
    title: str,
    headers: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=title)
    ws.append(headers)

    for record in rows:
        ws.append([output_value(record.get(column)) for column in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(headers, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            value_len = len(str(value)) if value not in (None, "") else 0
            if value_len > max_len:
                max_len = value_len

        ws.column_dimensions[get_column_letter(col_idx)].width = min(60, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(
        workbook=workbook,
        title="empirical_candidates",
        headers=EMPIRICAL_HEADERS,
        rows=empirical_rows,
    )
    write_sheet(
        workbook=workbook,
        title="regression_candidates",
        headers=REGRESSION_HEADERS,
        rows=regression_rows,
    )

    workbook.save(output_path)


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {in_dir}")
    if not in_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = choose_output_path(in_dir, out_dir)
    output_base_name = f"{in_dir.name}_PARAM"

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
            skip_reason = file_skip_reason(file_path, output_base_name)
            if skip_reason is not None:
                print(f"Skipped: {file_path.name} ({skip_reason})")
                continue

            try:
                labels = parse_labels(file_path)
            except ValueError as exc:
                print(f"Skipped: {file_path.name} ({exc})")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_empirical_rows = extract_empirical_rows(wb, labels, file_path.name)
                file_regression_rows = extract_regression_rows(wb, labels, file_path.name)
                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_count += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_no_save(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output written: {output_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
