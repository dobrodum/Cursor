#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from all .xlsx files in input_dir.
"""

from __future__ import annotations

import datetime as dt
import math
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --------------------------
# User-configurable folders
# --------------------------
input_dir = "./input"
output_dir = "./output"

# --------------------------
# Extraction configuration
# --------------------------
N_QUARTERS = 10

# Offsets are relative to the "max" anchor column.
# These are kept centralized so they can be adjusted for workbook template variants.
EMPIRICAL_OFFSETS = {
    "result_start_row": 1,  # first candidate row is one row below anchor
    "forecast_value_col": -1,  # estimated total sold
    "actual_value_col": -2,  # reported sales
    "last_quarter_used_col": -8,
    "quarterly_sales_col": -5,
    "reported_sales_col": -4,
    "growth_rate_pct_col": -3,
    "sales_captured_in_db_pct_col": -7,
    "penetration_history_col": -9,  # source col for avg penetration formula
    "temp_avg_formula_col": 24,  # temporary write col for formula2
}

REGRESSION_OFFSETS = {
    "result_start_row": 1,  # first candidate row is one row below anchor
    "forecast_max_col": 0,
    "forecast_min_col": 1,
    "actual_value_col": -2,  # may be blank in many files
    "temp_intercept_formula_col": 24,  # temporary write col for formula2
    "temp_slope_formula_col": 25,  # temporary write col for formula2
}

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


def parse_file_label(file_path: Path) -> Dict[str, Any]:
    """
    Parse metadata from names like:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[-2] if len(parts) >= 2 else "UNKNOWN"
    period_raw = parts[-1] if parts else ""
    period_token = re.sub(r"(?i)_send.*$", "", period_raw).strip()

    model_period = period_token
    model_date = None

    match = re.search(r"(?i)^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", period_token)
    if match:
        phase = match.group(1).title()
        month_token = match.group(2).title()
        year = int(match.group(3))

        month_num = month_to_number(month_token)
        if month_num is not None:
            month_abbr = dt.date(2000, month_num, 1).strftime("%b")
            model_period = f"{phase}{month_abbr}_{year}"
            day = {"Early": 5, "Mid": 15, "Late": 25}[phase]
            model_date = dt.date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def month_to_number(month_token: str) -> Optional[int]:
    token = month_token[:3].title()
    try:
        return dt.datetime.strptime(token, "%b").month
    except ValueError:
        return None


def next_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    input_folder_name = input_path.resolve().name
    base = f"{input_folder_name}_PARAM"

    candidate = output_path / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def normalize_grid(values: Any) -> List[List[Any]]:
    if values is None:
        return []

    if not isinstance(values, (list, tuple)):
        return [[values]]

    if values and not isinstance(values[0], (list, tuple)):
        return [list(values)]

    grid: List[List[Any]] = []
    for row in values:
        if isinstance(row, (list, tuple)):
            grid.append(list(row))
        else:
            grid.append([row])
    return grid


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    grid = normalize_grid(used.value)
    target = anchor_text.strip().lower()

    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target:
                return used.row + r_idx, used.column + c_idx
    return None


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    """
    Set formula via formula2; fallback paths support older engines.
    """
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

    try:
        cell.api.FormulaR1C1 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_workbook(wb: Optional[xw.Book]) -> None:
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
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(numeric) or math.isinf(numeric):
        return None
    return numeric


def compute_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def maybe_iso(value: Any) -> Any:
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat()
    return value


def extract_empirical_rows(
    wb: xw.Book, metadata: Dict[str, Any], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"SKIPPED: {source_file} (missing sheet: Empirical Model)")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIPPED: {source_file} (no 'max' anchor in Empirical Model)")
        return []

    anchor_row, anchor_col = anchor
    result_start_row = anchor_row + EMPIRICAL_OFFSETS["result_start_row"]
    temp_avg_col = anchor_col + EMPIRICAL_OFFSETS["temp_avg_formula_col"]
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["penetration_history_col"]

    # Write all average-penetration formulas first, then calculate once.
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        result_row = result_start_row + idx

        source_end_row = anchor_row - 1
        source_start_row = max(1, source_end_row - n_quarters + 1)
        formula = (
            f"=AVERAGE(R{source_start_row}C{penetration_col}:"
            f"R{source_end_row}C{penetration_col})"
        )
        set_formula2(sheet.cells(result_row, temp_avg_col), formula)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        result_row = result_start_row + idx

        avg_penetration = sheet.cells(result_row, temp_avg_col).value
        forecast_max = sheet.cells(result_row, anchor_col).value
        forecast_min = sheet.cells(result_row, anchor_col + 1).value
        forecast_value = sheet.cells(
            result_row, anchor_col + EMPIRICAL_OFFSETS["forecast_value_col"]
        ).value
        actual_value = sheet.cells(
            result_row, anchor_col + EMPIRICAL_OFFSETS["actual_value_col"]
        ).value
        last_quarter_used = sheet.cells(
            result_row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used_col"]
        ).value
        quarterly_sales = sheet.cells(
            result_row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"]
        ).value
        reported_sales = sheet.cells(
            result_row, anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"]
        ).value
        growth_rate_pct = sheet.cells(
            result_row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct_col"]
        ).value
        sales_captured = sheet.cells(
            result_row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct_col"]
        ).value

        # Fallbacks preserve robustness across minor layout variants.
        if forecast_value is None:
            avg_num = to_float(avg_penetration)
            sales_num = to_float(quarterly_sales)
            if avg_num is not None and sales_num is not None:
                forecast_value = avg_num * sales_num

        if actual_value is None:
            actual_value = reported_sales

        if (
            avg_penetration is None
            and forecast_value is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": maybe_iso(last_quarter_used),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book, metadata: Dict[str, Any], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"SKIPPED: {source_file} (missing sheet: Regression Model)")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIPPED: {source_file} (no 'max' anchor in Regression Model)")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    end_row = anchor_row - 1

    result_start_row = anchor_row + REGRESSION_OFFSETS["result_start_row"]
    intercept_col = anchor_col + REGRESSION_OFFSETS["temp_intercept_formula_col"]
    slope_col = anchor_col + REGRESSION_OFFSETS["temp_slope_formula_col"]

    # Write all INTERCEPT/SLOPE formulas first, then calculate once.
    valid_n_values: List[int] = []
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        start_row = end_row - n_quarters + 1
        if start_row < 1:
            continue

        out_row = result_start_row + idx
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2(sheet.cells(out_row, intercept_col), intercept_formula)
        set_formula2(sheet.cells(out_row, slope_col), slope_formula)
        valid_n_values.append(n_quarters)

    wb.app.calculate()

    x_next = sheet.cells(anchor_row, x_col).value
    if to_float(x_next) is None:
        x_last = sheet.cells(end_row, x_col).value
        x_last_num = to_float(x_last)
        if x_last_num is not None:
            x_next = x_last_num + 1

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n_quarters in valid_n_values:
        idx = n_quarters - 1
        out_row = result_start_row + idx

        intercept = sheet.cells(out_row, intercept_col).value
        slope = sheet.cells(out_row, slope_col).value

        intercept_num = to_float(intercept)
        slope_num = to_float(slope)
        x_next_num = to_float(x_next)

        forecast_value = None
        if (
            intercept_num is not None
            and slope_num is not None
            and x_next_num is not None
        ):
            forecast_value = intercept_num + (slope_num * x_next_num)

        forecast_max = sheet.cells(
            out_row, anchor_col + REGRESSION_OFFSETS["forecast_max_col"]
        ).value
        forecast_min = sheet.cells(
            out_row, anchor_col + REGRESSION_OFFSETS["forecast_min_col"]
        ).value
        actual_value = sheet.cells(
            out_row, anchor_col + REGRESSION_OFFSETS["actual_value_col"]
        ).value

        if (
            intercept is None
            and slope is None
            and forecast_value is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        signature = (
            rounded_signature_value(forecast_value),
            rounded_signature_value(forecast_max),
            rounded_signature_value(forecast_min),
            rounded_signature_value(intercept),
            rounded_signature_value(slope),
        )
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
                "range_width": compute_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def rounded_signature_value(value: Any) -> Any:
    numeric = to_float(value)
    if numeric is None:
        return value
    return round(numeric, 10)


def write_sheet(
    ws: Any, columns: List[str], rows: List[Dict[str, Any]], max_col_width: int = 42
) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_index, column_name in enumerate(columns, start=1):
        width = len(column_name)
        for row_index in range(2, ws.max_row + 1):
            value = ws.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            width = max(width, len(str(value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(width + 2, max_col_width)


def write_output_workbook(
    output_file: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()

    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = workbook.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input folder not found or not a directory: {input_path}")

    output_file = next_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"SKIPPED: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"SKIPPED: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED: {file_path.name} (not .xlsx)")
                continue

            print(f"PROCESSING: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                # Open source workbook once, process both model sheets, close without save.
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_label(file_path)

                empirical = extract_empirical_rows(wb, metadata, file_path.name)
                regression = extract_regression_rows(wb, metadata, file_path.name)

                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_files += 1
                print(
                    f"PROCESSED: {file_path.name} "
                    f"(empirical_rows={len(empirical)}, regression_rows={len(regression)})"
                )
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (error: {exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_file}")
    print(f"FILES PROCESSED: {processed_files}")
    print(f"EMPIRICAL ROWS: {len(empirical_rows)}")
    print(f"REGRESSION ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
