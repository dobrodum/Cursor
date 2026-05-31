#!/usr/bin/env python3
"""Extract empirical and regression candidates from model workbooks.

This script scans all .xlsx files in ``input_dir`` and writes one output workbook
with two tabs:
    - empirical_candidates
    - regression_candidates

Design goals:
    - Open each source workbook exactly once.
    - Process both source model sheets while the workbook is open.
    - Never save or modify source files permanently.
    - Keep runtime low by minimizing workbook opens and recalculations.
"""

from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
N_QUARTERS = 10
EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"

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

# Offsets are measured relative to the column containing the "max" anchor.
# Data rows begin immediately below the anchor row.
EMPIRICAL_OFFSETS = {
    "row_start": 1,
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "avg_penetration_source": -4,
    "actual_value": -2,
    "forecast_value": -1,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_helper": 25,
}

REGRESSION_OFFSETS = {
    "row_start": 1,
    "num_quarters_used": -10,
    "actual_value": -2,  # optional; may be blank
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept_helper": 25,
    "slope_helper": 26,
}

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
PERIOD_REGEX = re.compile(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", re.IGNORECASE)


def normalize_value(value: Any) -> Any:
    """Normalize Excel values for consistent downstream handling."""
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value


def is_blank(value: Any) -> bool:
    value = normalize_value(value)
    return value is None


def to_int(value: Any) -> Optional[int]:
    value = normalize_value(value)
    if value is None:
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def to_float(value: Any) -> Optional[float]:
    value = normalize_value(value)
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def numeric_diff(a: Any, b: Any) -> Optional[float]:
    af = to_float(a)
    bf = to_float(b)
    if af is None or bf is None:
        return None
    return af - bf


def safe_close_workbook(wb: Any) -> None:
    """Close workbook without saving, with safe fallbacks across platforms."""
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
    except Exception:
        pass


def set_formula2_r1c1(cell: Any, formula_r1c1: str) -> None:
    """Set formula using Formula2 with R1C1-style references."""
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

    # Last-resort fallback if Formula2 isn't available in this Excel runtime.
    cell.formula = formula_r1c1


def find_anchor_cell(sheet: Any, anchor_text: str = "max") -> Tuple[int, int]:
    """Find (row, col) of the first anchor cell matching anchor_text."""
    anchor_text = anchor_text.strip().lower()

    try:
        found = sheet.api.Cells.Find(What=anchor_text, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    used = sheet.used_range
    values = used.value
    start_row = int(used.row)
    start_col = int(used.column)

    if not isinstance(values, list):
        grid = [[values]]
    elif values and not isinstance(values[0], list):
        grid = [values]
    else:
        grid = values or []

    for r_idx, row_vals in enumerate(grid):
        for c_idx, cell_val in enumerate(row_vals):
            if isinstance(cell_val, str) and cell_val.strip().lower() == anchor_text:
                return start_row + r_idx, start_col + c_idx

    raise ValueError(f'Anchor "{anchor_text}" not found in sheet "{sheet.name}"')


def parse_file_label(file_path: Path) -> Dict[str, str]:
    """Parse ticker/model period/model date from source filename."""
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.split(r"[_\s]", parts[1])[0].strip().upper()
    if not ticker:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    model_period = ""
    model_date = ""
    period_match = PERIOD_REGEX.search(stem)
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.capitalize()
        month_abbrev = month_raw[:3].title()
        model_period = f"{phase}{month_abbrev}_{year_raw}"

        month_num = MONTH_MAP.get(month_abbrev.lower())
        day_num = PERIOD_DAY_MAP.get(phase.lower())
        if month_num and day_num:
            model_date = date(int(year_raw), month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or model_period or stem)

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_path: Path, output_path: Path) -> Path:
    """Build output path, adding .1/.2/etc if needed."""
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def read_anchor_offset(sheet: Any, row: int, anchor_col: int, col_offset: int) -> Any:
    col = anchor_col + col_offset
    if col < 1:
        return None
    return normalize_value(sheet.range((row, col)).value)


def extract_empirical_rows(wb: Any, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    """Extract empirical candidate rows from one open workbook."""
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET]
    except Exception:
        print(f"  - {EMPIRICAL_SHEET!r} not found; empirical rows skipped.")
        return []

    anchor_row, anchor_col = find_anchor_cell(sheet, "max")
    prepared_rows: List[Tuple[int, int, Any]] = []

    for i in range(N_QUARTERS):
        row = anchor_row + EMPIRICAL_OFFSETS["row_start"] + i
        num_quarters_used = to_int(
            read_anchor_offset(sheet, row, anchor_col, EMPIRICAL_OFFSETS["num_quarters_used"])
        )
        if not num_quarters_used or num_quarters_used < 1:
            num_quarters_used = i + 1

        helper_col = anchor_col + EMPIRICAL_OFFSETS["avg_penetration_helper"]
        helper_cell = sheet.range((row, helper_col))
        relative_penetration_col = (
            EMPIRICAL_OFFSETS["avg_penetration_source"] - EMPIRICAL_OFFSETS["avg_penetration_helper"]
        )
        lookback = max(num_quarters_used - 1, 0)

        # R1C1 + Formula2 to avoid column-letter conversion.
        avg_pen_formula = (
            f'=IFERROR(AVERAGE(R[-{lookback}]C[{relative_penetration_col}]:RC[{relative_penetration_col}]),"")'
        )
        set_formula2_r1c1(helper_cell, avg_pen_formula)
        prepared_rows.append((row, num_quarters_used, helper_cell))

    if prepared_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for row, num_quarters_used, helper_cell in prepared_rows:
        last_quarter_used = read_anchor_offset(
            sheet, row, anchor_col, EMPIRICAL_OFFSETS["last_quarter_used"]
        )
        forecast_value = read_anchor_offset(sheet, row, anchor_col, EMPIRICAL_OFFSETS["forecast_value"])
        actual_value = read_anchor_offset(sheet, row, anchor_col, EMPIRICAL_OFFSETS["actual_value"])
        forecast_max = read_anchor_offset(sheet, row, anchor_col, EMPIRICAL_OFFSETS["forecast_max"])
        forecast_min = read_anchor_offset(sheet, row, anchor_col, EMPIRICAL_OFFSETS["forecast_min"])
        quarterly_sales = read_anchor_offset(
            sheet, row, anchor_col, EMPIRICAL_OFFSETS["quarterly_sales"]
        )
        reported_sales = read_anchor_offset(
            sheet, row, anchor_col, EMPIRICAL_OFFSETS["reported_sales"]
        )
        growth_rate_pct = read_anchor_offset(
            sheet, row, anchor_col, EMPIRICAL_OFFSETS["growth_rate_pct"]
        )
        sales_captured_in_db_pct = read_anchor_offset(
            sheet, row, anchor_col, EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
        )
        avg_penetration_pct = normalize_value(helper_cell.value)

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
                sales_captured_in_db_pct,
                avg_penetration_pct,
            )
        ):
            continue

        row_data = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": numeric_diff(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row_data)

    for _, _, helper_cell in prepared_rows:
        helper_cell.value = None

    return rows


def numbers_close(a: Any, b: Any, tol: float = 1e-9) -> bool:
    af = to_float(a)
    bf = to_float(b)
    if af is None and bf is None:
        return True
    if af is None or bf is None:
        return False
    return abs(af - bf) <= tol


def is_duplicate_regression_row(previous: Dict[str, Any], current: Dict[str, Any]) -> bool:
    fields = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    return all(numbers_close(previous.get(field), current.get(field)) for field in fields)


def extract_regression_rows(wb: Any, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    """Extract regression candidate rows from one open workbook."""
    try:
        sheet = wb.sheets[REGRESSION_SHEET]
    except Exception:
        print(f"  - {REGRESSION_SHEET!r} not found; regression rows skipped.")
        return []

    anchor_row, anchor_col = find_anchor_cell(sheet, "max")
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    series_end_row = max(anchor_row - 1, 1)

    prepared_rows: List[Tuple[int, int, Any, Any]] = []
    for i in range(N_QUARTERS):
        row = anchor_row + REGRESSION_OFFSETS["row_start"] + i
        num_quarters_used = to_int(
            read_anchor_offset(sheet, row, anchor_col, REGRESSION_OFFSETS["num_quarters_used"])
        )
        if not num_quarters_used or num_quarters_used < 1:
            num_quarters_used = i + 1
        num_quarters_used = min(num_quarters_used, series_end_row)

        start_row = max(series_end_row - num_quarters_used + 1, 1)
        intercept_cell = sheet.range((row, anchor_col + REGRESSION_OFFSETS["intercept_helper"]))
        slope_cell = sheet.range((row, anchor_col + REGRESSION_OFFSETS["slope_helper"]))

        # R1C1 + Formula2 keeps formulas fast and avoids A1 column conversion.
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{series_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{series_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{series_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{series_end_row}C{x_col}),"")'
        )
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        prepared_rows.append((row, num_quarters_used, intercept_cell, slope_cell))

    if prepared_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for row, num_quarters_used, intercept_cell, slope_cell in prepared_rows:
        forecast_value = read_anchor_offset(sheet, row, anchor_col, REGRESSION_OFFSETS["forecast_value"])
        actual_value = read_anchor_offset(sheet, row, anchor_col, REGRESSION_OFFSETS["actual_value"])
        forecast_max = read_anchor_offset(sheet, row, anchor_col, REGRESSION_OFFSETS["forecast_max"])
        forecast_min = read_anchor_offset(sheet, row, anchor_col, REGRESSION_OFFSETS["forecast_min"])
        intercept = normalize_value(intercept_cell.value)
        slope = normalize_value(slope_cell.value)

        if all(is_blank(v) for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row_data = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": numeric_diff(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(rows[-1], row_data):
            continue
        rows.append(row_data)

    for _, _, intercept_cell, slope_cell in prepared_rows:
        intercept_cell.value = None
        slope_cell.value = None

    return rows


def write_sheet(ws: Any, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def write_output_workbook(
    output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    empirical_ws = wb.create_sheet(EMPIRICAL_OUTPUT_SHEET)
    regression_ws = wb.create_sheet(REGRESSION_OUTPUT_SHEET)

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    in_path = input_dir.expanduser().resolve()
    out_path = output_dir.expanduser().resolve()

    if not in_path.exists() or not in_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_path}")

    out_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(in_path, out_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(in_path.iterdir()):
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not an .xlsx file)")
                continue

            print(f"Processing file: {file_path.name}")
            wb = None
            try:
                # Source workbook safety requirement.
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_label(file_path)

                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    # Never save source workbooks.
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
