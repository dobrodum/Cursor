#!/usr/bin/env python3
"""
Batch extract empirical/regression candidates from .xlsx model files.

Runtime priorities:
- one hidden Excel app for the full run
- open each workbook once
- process Empirical Model + Regression Model while workbook is open
- close workbook without saving source changes
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


# --------- user-editable paths ---------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")
# ---------------------------------------

N_QUARTERS = 10

# Anchor-based offsets (relative to "max" anchor column)
RESULT_ROW_START_OFFSET = 1
FORECAST_VALUE_COL_OFFSET = -1
ACTUAL_VALUE_COL_OFFSET = -2
FORECAST_MAX_COL_OFFSET = 0
FORECAST_MIN_COL_OFFSET = 1
SCRATCH_COL_OFFSET = 8

# Empirical-specific offsets (relative to anchor column)
EMPIRICAL_QUARTERLY_SALES_COL_OFFSET = -11
EMPIRICAL_REPORTED_SALES_COL_OFFSET = -7
EMPIRICAL_LAST_QUARTER_LABEL_COL_OFFSET = -12

# Regression-specific offsets (required by prompt)
REGRESSION_Y_COL_OFFSET = -7
REGRESSION_X_COL_OFFSET = -11

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

PERIOD_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_RE = re.compile(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", re.IGNORECASE)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def has_sheet(wb: xw.Book, sheet_name: str) -> bool:
    return any(s.name == sheet_name for s in wb.sheets)


def normalize_used_range_values(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[tuple[int, int]]:
    used = sheet.used_range
    values = normalize_used_range_values(used.value)
    if not values:
        return None

    target = anchor_text.strip().lower()
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return used.row + r_idx, used.column + c_idx
    return None


def read_formula2(cell: xw.Range) -> Optional[str]:
    try:
        return cell.formula2
    except Exception:
        try:
            return cell.api.Formula2
        except Exception:
            return None


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass
    cell.formula = formula


def restore_temporary_cells(temp_cells: list[tuple[xw.Range, Optional[str], Any]]) -> None:
    for cell, previous_formula2, previous_value in temp_cells:
        try:
            if previous_formula2:
                set_formula2_r1c1(cell, previous_formula2)
            else:
                cell.value = previous_value
        except Exception:
            # Best-effort restore only; workbook is still closed without saving.
            pass


def close_workbook_safely(wb: Optional[xw.Book]) -> None:
    if wb is None:
        return
    close_attempts = (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.close(),
    )
    for close_call in close_attempts:
        try:
            close_call()
            return
        except TypeError:
            continue
        except Exception:
            continue


def parse_file_metadata(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"
    period_chunk = parts[2] if len(parts) >= 3 else ""
    period_token = period_chunk.split("_")[0]

    model_period = "unknown_period"
    model_date = ""

    match = PERIOD_RE.search(period_token)
    if match:
        period_label = match.group(1).title()
        month_text = match.group(2)
        year_text = match.group(3)
        month_abbrev = month_text[:3].title()
        month_number = MONTH_MAP.get(month_abbrev.lower())
        if month_number:
            model_period = f"{period_label}{month_abbrev}_{year_text}"
            model_date = date(
                int(year_text),
                month_number,
                PERIOD_DAY_MAP.get(period_label, 15),
            ).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def make_unique_output_path(input_root: Path, output_root: Path) -> Path:
    base_name = f"{input_root.name}_PARAM"
    output_path = output_root / f"{base_name}.xlsx"

    counter = 1
    while output_path.exists():
        output_path = output_root / f"{base_name}.{counter}.xlsx"
        counter += 1
    return output_path


def calc_avg_penetration_python(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    quarterly_col: int,
    reported_col: int,
) -> Optional[float]:
    ratios: list[float] = []
    for row in range(start_row, end_row + 1):
        quarterly = to_float(sheet.cells(row, quarterly_col).value)
        reported = to_float(sheet.cells(row, reported_col).value)
        ratio = safe_div(quarterly, reported)
        if ratio is not None:
            ratios.append(ratio)
    if not ratios:
        return None
    return sum(ratios) / len(ratios)


def extract_empirical_candidates(
    wb: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet_name = "Empirical Model"
    if not has_sheet(wb, sheet_name):
        print(f"Skipped: {source_file} ({sheet_name} sheet not found)")
        return []

    sheet = wb.sheets[sheet_name]
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped: {source_file} (max anchor not found on {sheet_name})")
        return []

    anchor_row, anchor_col = anchor
    hist_end_row = anchor_row - 1
    if hist_end_row <= 0:
        return []

    result_start_row = anchor_row + RESULT_ROW_START_OFFSET
    quarterly_col = anchor_col + EMPIRICAL_QUARTERLY_SALES_COL_OFFSET
    reported_col = anchor_col + EMPIRICAL_REPORTED_SALES_COL_OFFSET
    last_quarter_col = anchor_col + EMPIRICAL_LAST_QUARTER_LABEL_COL_OFFSET
    scratch_col = anchor_col + SCRATCH_COL_OFFSET

    rows: list[dict[str, Any]] = []
    temp_cells: list[tuple[xw.Range, Optional[str], Any]] = []

    try:
        for n_quarters in range(1, N_QUARTERS + 1):
            hist_start_row = hist_end_row - n_quarters + 1
            if hist_start_row <= 0:
                break

            result_row = result_start_row + n_quarters - 1

            avg_pen_cell = sheet.cells(result_row, scratch_col)
            temp_cells.append((avg_pen_cell, read_formula2(avg_pen_cell), avg_pen_cell.value))
            avg_pen_formula = (
                f"=LET(q,R{hist_start_row}C{quarterly_col}:R{hist_end_row}C{quarterly_col},"
                f"r,R{hist_start_row}C{reported_col}:R{hist_end_row}C{reported_col},"
                "AVERAGE(IFERROR(q/r,\"\")))"
            )
            set_formula2_r1c1(avg_pen_cell, avg_pen_formula)
            wb.app.calculate()

            avg_penetration_pct = to_float(avg_pen_cell.value)
            if avg_penetration_pct is None:
                avg_penetration_pct = calc_avg_penetration_python(
                    sheet, hist_start_row, hist_end_row, quarterly_col, reported_col
                )

            quarterly_sales = to_float(sheet.cells(hist_end_row, quarterly_col).value)
            reported_sales = to_float(sheet.cells(hist_end_row, reported_col).value)
            sales_captured_in_db_pct = safe_div(quarterly_sales, reported_sales)
            estimated_total_sold = safe_div(quarterly_sales, avg_penetration_pct)

            forecast_value = to_float(sheet.cells(result_row, anchor_col + FORECAST_VALUE_COL_OFFSET).value)
            if forecast_value is None:
                forecast_value = estimated_total_sold

            actual_value = to_float(sheet.cells(result_row, anchor_col + ACTUAL_VALUE_COL_OFFSET).value)
            if actual_value is None:
                actual_value = reported_sales

            forecast_max = to_float(sheet.cells(result_row, anchor_col + FORECAST_MAX_COL_OFFSET).value)
            forecast_min = to_float(sheet.cells(result_row, anchor_col + FORECAST_MIN_COL_OFFSET).value)
            if forecast_max is None:
                forecast_max = forecast_value
            if forecast_min is None:
                forecast_min = forecast_value

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            previous_quarterly_sales = None
            if hist_end_row - 1 >= hist_start_row:
                previous_quarterly_sales = to_float(sheet.cells(hist_end_row - 1, quarterly_col).value)
            growth_rate_pct = None
            if quarterly_sales is not None and previous_quarterly_sales not in (None, 0):
                growth_rate_pct = (quarterly_sales - previous_quarterly_sales) / previous_quarterly_sales

            last_quarter_used = sheet.cells(hist_end_row, last_quarter_col).value
            if last_quarter_used is None:
                last_quarter_used = f"row_{hist_end_row}"

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
        restore_temporary_cells(temp_cells)

    return rows


def build_regression_signature(
    forecast_value: Optional[float],
    forecast_max: Optional[float],
    forecast_min: Optional[float],
    intercept: Optional[float],
    slope: Optional[float],
) -> tuple[Any, ...]:
    def normalize_number(value: Optional[float]) -> Optional[float]:
        if value is None:
            return None
        return round(value, 10)

    return (
        normalize_number(forecast_value),
        normalize_number(forecast_max),
        normalize_number(forecast_min),
        normalize_number(intercept),
        normalize_number(slope),
    )


def extract_regression_candidates(
    wb: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet_name = "Regression Model"
    if not has_sheet(wb, sheet_name):
        print(f"Skipped: {source_file} ({sheet_name} sheet not found)")
        return []

    sheet = wb.sheets[sheet_name]
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped: {source_file} (max anchor not found on {sheet_name})")
        return []

    anchor_row, anchor_col = anchor
    hist_end_row = anchor_row - 1
    if hist_end_row <= 0:
        return []

    result_start_row = anchor_row + RESULT_ROW_START_OFFSET
    x_col = anchor_col + REGRESSION_X_COL_OFFSET
    y_col = anchor_col + REGRESSION_Y_COL_OFFSET
    intercept_scratch_col = anchor_col + SCRATCH_COL_OFFSET
    slope_scratch_col = intercept_scratch_col + 1

    rows: list[dict[str, Any]] = []
    temp_cells: list[tuple[xw.Range, Optional[str], Any]] = []
    previous_signature: Optional[tuple[Any, ...]] = None

    try:
        for n_quarters in range(1, N_QUARTERS + 1):
            hist_start_row = hist_end_row - n_quarters + 1
            if hist_start_row <= 0:
                break

            result_row = result_start_row + n_quarters - 1

            intercept_cell = sheet.cells(result_row, intercept_scratch_col)
            slope_cell = sheet.cells(result_row, slope_scratch_col)
            temp_cells.append((intercept_cell, read_formula2(intercept_cell), intercept_cell.value))
            temp_cells.append((slope_cell, read_formula2(slope_cell), slope_cell.value))

            intercept_formula = (
                f"=INTERCEPT(R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
                f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
                f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col})"
            )

            set_formula2_r1c1(intercept_cell, intercept_formula)
            set_formula2_r1c1(slope_cell, slope_formula)
            wb.app.calculate()

            intercept = to_float(intercept_cell.value)
            slope = to_float(slope_cell.value)

            forecast_value = to_float(sheet.cells(result_row, anchor_col + FORECAST_VALUE_COL_OFFSET).value)
            if forecast_value is None and intercept is not None and slope is not None:
                x_next = to_float(sheet.cells(hist_end_row + 1, x_col).value)
                if x_next is None:
                    x_next = to_float(sheet.cells(hist_end_row, x_col).value)
                if x_next is not None:
                    forecast_value = intercept + slope * x_next

            forecast_max = to_float(sheet.cells(result_row, anchor_col + FORECAST_MAX_COL_OFFSET).value)
            forecast_min = to_float(sheet.cells(result_row, anchor_col + FORECAST_MIN_COL_OFFSET).value)
            if forecast_max is None:
                forecast_max = forecast_value
            if forecast_min is None:
                forecast_min = forecast_value

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            actual_raw = sheet.cells(result_row, anchor_col + ACTUAL_VALUE_COL_OFFSET).value
            actual_numeric = to_float(actual_raw)
            actual_value = actual_numeric if actual_numeric is not None else ""

            signature = build_regression_signature(
                forecast_value=forecast_value,
                forecast_max=forecast_max,
                forecast_min=forecast_min,
                intercept=intercept,
                slope=slope,
            )
            if previous_signature == signature:
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
    finally:
        restore_temporary_cells(temp_cells)

    return rows


def excel_safe_value(value: Any) -> Any:
    if value is None:
        return ""
    return value


def write_sheet(ws: Worksheet, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)

    for row in rows:
        ws.append([excel_safe_value(row.get(col)) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for (value,) in ws.iter_rows(
            min_row=2,
            max_row=ws.max_row,
            min_col=col_idx,
            max_col=col_idx,
            values_only=True,
        ):
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(max_len + 2, 48))


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb_out = Workbook()
    ws_empirical = wb_out.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb_out.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb_out.save(output_path)


def should_skip_file(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary Excel file"
    if file_path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    if re.search(r"_PARAM(\.\d+)?\.xlsx$", file_path.name, re.IGNORECASE):
        return "existing output workbook"
    return None


def main() -> None:
    input_root = Path(input_dir).expanduser().resolve()
    output_root = Path(output_dir).expanduser().resolve()

    if not input_root.exists() or not input_root.is_dir():
        raise SystemExit(f"Input folder does not exist or is not a directory: {input_root}")

    output_root.mkdir(parents=True, exist_ok=True)
    output_path = make_unique_output_path(input_root, output_root)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.enable_events = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_root.iterdir(), key=lambda p: p.name.lower()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped: {file_path.name} ({skip_reason})")
                continue

            print(f"Processing: {file_path.name}")
            metadata = parse_file_metadata(file_path.name)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_candidates(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_candidates(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                close_workbook_safely(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
