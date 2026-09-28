#!/usr/bin/env python3
"""
Extract empirical and regression model candidate rows from Excel model files.

Requirements implemented:
- Open each source workbook once with xlwings and process both model sheets while open.
- Never save or modify source files permanently (temporary formula writes only).
- Write one output workbook with two sheets:
    - empirical_candidates
    - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# --------- User-configurable paths ---------
input_dir = "/path/to/input"
output_dir = "/path/to/output"
# ------------------------------------------


N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"


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


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class EmpiricalOffsets:
    first_result_row_offset: int = 1
    num_quarters_col_offset: int = -9
    last_quarter_col_offset: int = -8
    forecast_value_col_offset: int = -7
    actual_value_col_offset: int = -6
    quarterly_sales_col_offset: int = -5
    reported_sales_col_offset: int = -4
    growth_rate_col_offset: int = -3
    sales_captured_col_offset: int = -2
    forecast_max_col_offset: int = 0
    forecast_min_col_offset: int = 1
    penetration_source_col_offset: int = -10
    helper_formula_col_offset: int = 3
    penetration_end_row_offset: int = -1


@dataclass(frozen=True)
class RegressionOffsets:
    first_result_row_offset: int = 1
    num_quarters_col_offset: int = -9
    forecast_total_col_offset: int = -7
    forecast_max_col_offset: int = 0
    forecast_min_col_offset: int = 1
    helper_intercept_col_offset: int = 3
    helper_slope_col_offset: int = 4
    history_end_row_offset: int = -1


EMPIRICAL_OFFSETS = EmpiricalOffsets()
REGRESSION_OFFSETS = RegressionOffsets()


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


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "UNKNOWN"
    model_date = ""
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase_key = phase_raw.lower()
        month_key = month_raw[:3].lower()
        year = int(year_raw)
        month_num = MONTH_MAP.get(month_key)
        day_num = DAY_MAP.get(phase_key)
        if month_num and day_num:
            phase = phase_raw[:1].upper() + phase_raw[1:].lower()
            month = month_raw[:1].upper() + month_raw[1:3].lower()
            model_period = f"{phase}{month}_{year}"
            model_date = date(year, month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    base_path = out_dir / f"{base_name}.xlsx"
    if not base_path.exists():
        return base_path

    version = 1
    while True:
        candidate = out_dir / f"{base_name}.{version}.xlsx"
        if not candidate.exists():
            return candidate
        version += 1


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor_max(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    grid = to_2d(used.value)
    if not grid:
        return None

    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return used.row + r_idx, used.column + c_idx
    return None


def get_sheet_case_insensitive(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def read_cell_or_default(sheet: xw.Sheet, row: int, col: int, default: Any) -> Any:
    value = read_cell(sheet, row, col)
    return default if value is None else value


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback for older Excel APIs that do not expose formula2.
        cell.formula = formula_r1c1


def subtract_if_numeric(a: Any, b: Any) -> Any:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return a - b
    return None


def normalize_for_signature(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 12)
    return value


def close_workbook_safely(workbook: xw.Book) -> None:
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
        workbook.saved = True
        workbook.close()
    except Exception:
        print(f"Warning: failed to close workbook safely: {workbook.name}")


def extract_empirical_candidates(
    workbook: xw.Book, sheet: xw.Sheet, labels: FileLabels, source_file: str
) -> tuple[list[dict[str, Any]], str | None]:
    anchor = find_anchor_max(sheet)
    if anchor is None:
        return [], "max anchor not found"

    anchor_row, anchor_col = anchor
    rows: list[dict[str, Any]] = []

    penetration_col = anchor_col + EMPIRICAL_OFFSETS.penetration_source_col_offset
    penetration_end_row = anchor_row + EMPIRICAL_OFFSETS.penetration_end_row_offset

    for n_quarters in range(1, N_QUARTERS + 1):
        result_row = (
            anchor_row
            + EMPIRICAL_OFFSETS.first_result_row_offset
            + (n_quarters - 1)
        )

        avg_penetration_pct = None
        if penetration_col >= 1 and penetration_end_row >= 1:
            penetration_start_row = max(1, penetration_end_row - n_quarters + 1)
            if penetration_start_row <= penetration_end_row:
                helper_col = anchor_col + EMPIRICAL_OFFSETS.helper_formula_col_offset
                helper_cell = sheet.cells(result_row, helper_col)
                avg_formula = (
                    f"=AVERAGE(R{penetration_start_row}C{penetration_col}:"
                    f"R{penetration_end_row}C{penetration_col})"
                )
                set_formula2(helper_cell, avg_formula)
                workbook.app.calculate()
                avg_penetration_pct = helper_cell.value

        num_quarters_used = read_cell_or_default(
            sheet,
            result_row,
            anchor_col + EMPIRICAL_OFFSETS.num_quarters_col_offset,
            n_quarters,
        )
        last_quarter_used = read_cell(
            sheet, result_row, anchor_col + EMPIRICAL_OFFSETS.last_quarter_col_offset
        )
        forecast_value = read_cell(
            sheet,
            result_row,
            anchor_col + EMPIRICAL_OFFSETS.forecast_value_col_offset,
        )
        actual_value = read_cell(
            sheet,
            result_row,
            anchor_col + EMPIRICAL_OFFSETS.actual_value_col_offset,
        )
        forecast_max = read_cell(
            sheet, result_row, anchor_col + EMPIRICAL_OFFSETS.forecast_max_col_offset
        )
        forecast_min = read_cell(
            sheet, result_row, anchor_col + EMPIRICAL_OFFSETS.forecast_min_col_offset
        )
        quarterly_sales = read_cell(
            sheet,
            result_row,
            anchor_col + EMPIRICAL_OFFSETS.quarterly_sales_col_offset,
        )
        reported_sales = read_cell(
            sheet,
            result_row,
            anchor_col + EMPIRICAL_OFFSETS.reported_sales_col_offset,
        )
        growth_rate_pct = read_cell(
            sheet, result_row, anchor_col + EMPIRICAL_OFFSETS.growth_rate_col_offset
        )
        sales_captured_in_db_pct = read_cell(
            sheet, result_row, anchor_col + EMPIRICAL_OFFSETS.sales_captured_col_offset
        )

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
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
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows, None


def extract_regression_candidates(
    workbook: xw.Book, sheet: xw.Sheet, labels: FileLabels, source_file: str
) -> tuple[list[dict[str, Any]], str | None]:
    anchor = find_anchor_max(sheet)
    if anchor is None:
        return [], "max anchor not found"

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        return [], "anchor-based x/y columns are out of bounds"

    history_end_row = anchor_row + REGRESSION_OFFSETS.history_end_row_offset
    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters in range(1, N_QUARTERS + 1):
        history_start_row = history_end_row - n_quarters + 1
        if history_start_row < 1:
            break

        result_row = (
            anchor_row
            + REGRESSION_OFFSETS.first_result_row_offset
            + (n_quarters - 1)
        )

        intercept_col = anchor_col + REGRESSION_OFFSETS.helper_intercept_col_offset
        slope_col = anchor_col + REGRESSION_OFFSETS.helper_slope_col_offset
        intercept_cell = sheet.cells(result_row, intercept_col)
        slope_cell = sheet.cells(result_row, slope_col)

        intercept_formula = (
            f"=INTERCEPT(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        workbook.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        num_quarters_used = read_cell_or_default(
            sheet,
            result_row,
            anchor_col + REGRESSION_OFFSETS.num_quarters_col_offset,
            n_quarters,
        )
        forecast_total_without_sa = read_cell(
            sheet,
            result_row,
            anchor_col + REGRESSION_OFFSETS.forecast_total_col_offset,
        )
        if (
            forecast_total_without_sa is None
            and isinstance(intercept, (int, float))
            and isinstance(slope, (int, float))
        ):
            x_forecast = read_cell(sheet, history_end_row + 1, x_col)
            if x_forecast is None:
                x_forecast = read_cell(sheet, history_end_row, x_col)
            if isinstance(x_forecast, (int, float)):
                forecast_total_without_sa = intercept + slope * x_forecast

        forecast_max = read_cell(
            sheet, result_row, anchor_col + REGRESSION_OFFSETS.forecast_max_col_offset
        )
        forecast_min = read_cell(
            sheet, result_row, anchor_col + REGRESSION_OFFSETS.forecast_min_col_offset
        )

        if all(
            value is None
            for value in (
                intercept,
                slope,
                forecast_total_without_sa,
                forecast_max,
                forecast_min,
            )
        ):
            continue

        signature = (
            normalize_for_signature(num_quarters_used),
            normalize_for_signature(intercept),
            normalize_for_signature(slope),
            normalize_for_signature(forecast_total_without_sa),
            normalize_for_signature(forecast_max),
            normalize_for_signature(forecast_min),
        )
        if signature == previous_signature:
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
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_total_without_sa,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows, None


def write_sheet(
    workbook: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    ws = workbook.create_sheet(sheet_name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    write_sheet(output_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_wb.save(output_path)


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        print(f"Input directory not found: {in_dir}")
        return

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue

            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path.name)
                source_file = file_path.name

                empirical_sheet = get_sheet_case_insensitive(workbook, EMPIRICAL_SHEET_NAME)
                regression_sheet = get_sheet_case_insensitive(workbook, REGRESSION_SHEET_NAME)
                if empirical_sheet is None and regression_sheet is None:
                    print(
                        f"Skipped {file_path.name}: missing sheets "
                        f"'{EMPIRICAL_SHEET_NAME}' and '{REGRESSION_SHEET_NAME}'"
                    )
                    continue

                processed_file_count += 1
                print(f"Processed {file_path.name}")

                if empirical_sheet is None:
                    print(f"Skipped empirical in {file_path.name}: sheet not found")
                else:
                    extracted, reason = extract_empirical_candidates(
                        workbook, empirical_sheet, labels, source_file
                    )
                    empirical_rows.extend(extracted)
                    if reason:
                        print(f"Skipped empirical in {file_path.name}: {reason}")

                if regression_sheet is None:
                    print(f"Skipped regression in {file_path.name}: sheet not found")
                else:
                    extracted, reason = extract_regression_candidates(
                        workbook, regression_sheet, labels, source_file
                    )
                    regression_rows.extend(extracted)
                    if reason:
                        print(f"Skipped regression in {file_path.name}: {reason}")

            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_file_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
