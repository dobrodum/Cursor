#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----- User-editable paths -----
input_dir = Path("./input")
output_dir = Path("./output")

# ----- Extraction configuration -----
N_QUARTERS = 10

# Empirical Model offsets are relative to the "max" anchor cell.
EMP_NUM_QTR_INPUT_OFFSET = (-3, 1)
EMP_HELPER_FORMULA_OFFSET = (-2, 4)
EMP_FORECAST_VALUE_OFFSET = (-1, 1)
EMP_FORECAST_MAX_OFFSET = (0, 1)
EMP_FORECAST_MIN_OFFSET = (1, 1)
EMP_ACTUAL_VALUE_OFFSET = (2, 1)
EMP_LAST_QUARTER_COL_OFFSET = -12
EMP_QUARTERLY_SALES_COL_OFFSET = -11
EMP_REPORTED_SALES_COL_OFFSET = -10
EMP_GROWTH_RATE_COL_OFFSET = -9
EMP_CAPTURED_IN_DB_COL_OFFSET = -8
EMP_PENETRATION_COL_OFFSET = -7

# Regression Model offsets are relative to the "max" anchor cell.
REG_NUM_QTR_INPUT_OFFSET = (-3, 1)
REG_INTERCEPT_HELPER_OFFSET = (-2, 4)
REG_SLOPE_HELPER_OFFSET = (-1, 4)
REG_FORECAST_VALUE_OFFSET = (-1, 1)  # TOT FCST w/o SA
REG_FORECAST_MAX_OFFSET = (0, 1)
REG_FORECAST_MIN_OFFSET = (1, 1)
REG_ACTUAL_VALUE_OFFSET = (2, 1)  # optional; may be blank

EARLY_MID_LATE_DAY = {"Early": 5, "Mid": 15, "Late": 25}
FILE_PERIOD_RE = re.compile(
    r"^(?P<prefix>.+?)\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
    r"(?P<period>(Early|Mid|Late)([A-Za-z]{3})(\d{4}))_Send$",
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


@dataclass(frozen=True)
class ModelLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class AnchorCell:
    row: int
    col: int


def parse_file_labels(file_path: Path) -> ModelLabels:
    """Parse ticker/model period/date from filename."""
    match = FILE_PERIOD_RE.match(file_path.stem)
    if not match:
        raise ValueError("filename format does not match expected convention")

    ticker = match.group("ticker").upper()
    period_token = match.group("period")
    period_match = re.match(
        r"^(?P<bucket>Early|Mid|Late)(?P<mon>[A-Za-z]{3})(?P<year>\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if not period_match:
        raise ValueError("unable to parse period token")

    bucket = period_match.group("bucket").title()
    mon = period_match.group("mon").title()
    year = int(period_match.group("year"))
    month_number = datetime.strptime(mon, "%b").month
    day = EARLY_MID_LATE_DAY[bucket]

    model_period = f"{bucket}{mon}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    """Build the next available output path."""
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{in_dir.name}_PARAM"
    initial = out_dir / f"{base}.xlsx"
    if not initial.exists():
        return initial

    idx = 1
    while True:
        candidate = out_dir / f"{base}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def scalar_to_rows(value: Any) -> list[list[Any]]:
    """Convert xlwings UsedRange values to a 2D list."""
    if value is None:
        return []
    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def find_anchor(sheet: xw.Sheet, token: str = "max") -> AnchorCell | None:
    """Find the first case-insensitive token in used range."""
    used = sheet.used_range
    values = scalar_to_rows(used.value)
    base_row = used.row
    base_col = used.column

    token_norm = token.strip().lower()
    for r_idx, row_vals in enumerate(values):
        for c_idx, cell_val in enumerate(row_vals):
            if isinstance(cell_val, str) and cell_val.strip().lower() == token_norm:
                return AnchorCell(base_row + r_idx, base_col + c_idx)
    return None


def safe_close_source_workbook(wb: xw.Book) -> None:
    """Close a source workbook without saving, with fallbacks."""
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
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


def set_formula2(cell: xw.Range, formula: str) -> None:
    """Set Formula2 if available, else fallback to formula."""
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def add_offset(anchor: AnchorCell, offset: tuple[int, int]) -> tuple[int, int]:
    """Convert anchor+offset into absolute row/col."""
    return anchor.row + offset[0], anchor.col + offset[1]


def to_float(value: Any) -> float | None:
    """Best-effort float conversion."""
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def is_number(value: Any) -> bool:
    return to_float(value) is not None


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def contiguous_numeric_depth(sheet: xw.Sheet, col: int, end_row: int, limit: int) -> int:
    """Count contiguous numeric cells upward ending at end_row."""
    depth = 0
    for row in range(end_row, max(0, end_row - limit), -1):
        if is_number(read_cell(sheet, row, col)):
            depth += 1
        else:
            break
    return depth


def process_empirical_sheet(
    wb: xw.Book,
    labels: ModelLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    """Extract empirical candidate rows from one workbook."""
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  skipped empirical extraction: sheet 'Empirical Model' missing")
        return []

    anchor = find_anchor(sheet, token="max")
    if anchor is None:
        print("  skipped empirical extraction: 'max' anchor not found")
        return []

    num_qtr_input_row, num_qtr_input_col = add_offset(anchor, EMP_NUM_QTR_INPUT_OFFSET)
    helper_row, helper_col = add_offset(anchor, EMP_HELPER_FORMULA_OFFSET)
    forecast_value_row, forecast_value_col = add_offset(anchor, EMP_FORECAST_VALUE_OFFSET)
    forecast_max_row, forecast_max_col = add_offset(anchor, EMP_FORECAST_MAX_OFFSET)
    forecast_min_row, forecast_min_col = add_offset(anchor, EMP_FORECAST_MIN_OFFSET)
    actual_value_row, actual_value_col = add_offset(anchor, EMP_ACTUAL_VALUE_OFFSET)

    data_end_row = anchor.row - 1
    penetration_col = anchor.col + EMP_PENETRATION_COL_OFFSET
    available_depth = contiguous_numeric_depth(sheet, penetration_col, data_end_row, N_QUARTERS * 3)
    iterations = min(N_QUARTERS, available_depth if available_depth > 0 else N_QUARTERS)

    num_qtr_input_cell = sheet.cells(num_qtr_input_row, num_qtr_input_col)
    helper_cell = sheet.cells(helper_row, helper_col)
    rows: list[dict[str, Any]] = []

    for n_quarters in range(1, iterations + 1):
        start_row = data_end_row - n_quarters + 1
        if start_row < 1:
            continue

        # Trigger workbook formulas for this quarter count and compute avg penetration.
        num_qtr_input_cell.value = n_quarters
        avg_pen_formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{data_end_row}C{penetration_col})"
        )
        set_formula2(helper_cell, avg_pen_formula)
        wb.app.calculate()

        avg_penetration_pct = to_float(helper_cell.value)
        forecast_value = to_float(read_cell(sheet, forecast_value_row, forecast_value_col))
        forecast_max = to_float(read_cell(sheet, forecast_max_row, forecast_max_col))
        forecast_min = to_float(read_cell(sheet, forecast_min_row, forecast_min_col))

        support_row = data_end_row
        last_quarter_used = read_cell(sheet, support_row, anchor.col + EMP_LAST_QUARTER_COL_OFFSET)
        quarterly_sales = to_float(read_cell(sheet, support_row, anchor.col + EMP_QUARTERLY_SALES_COL_OFFSET))
        reported_sales = to_float(read_cell(sheet, support_row, anchor.col + EMP_REPORTED_SALES_COL_OFFSET))
        growth_rate_pct = to_float(read_cell(sheet, support_row, anchor.col + EMP_GROWTH_RATE_COL_OFFSET))
        sales_captured_in_db_pct = to_float(
            read_cell(sheet, support_row, anchor.col + EMP_CAPTURED_IN_DB_COL_OFFSET)
        )

        workbook_actual = to_float(read_cell(sheet, actual_value_row, actual_value_col))
        actual_value = reported_sales if reported_sales is not None else workbook_actual
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
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

    helper_cell.value = None
    return rows


def process_regression_sheet(
    wb: xw.Book,
    labels: ModelLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    """Extract regression candidate rows from one workbook."""
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  skipped regression extraction: sheet 'Regression Model' missing")
        return []

    anchor = find_anchor(sheet, token="max")
    if anchor is None:
        print("  skipped regression extraction: 'max' anchor not found")
        return []

    y_col = anchor.col - 7
    x_col = anchor.col - 11

    num_qtr_input_row, num_qtr_input_col = add_offset(anchor, REG_NUM_QTR_INPUT_OFFSET)
    intercept_row, intercept_col = add_offset(anchor, REG_INTERCEPT_HELPER_OFFSET)
    slope_row, slope_col = add_offset(anchor, REG_SLOPE_HELPER_OFFSET)
    forecast_value_row, forecast_value_col = add_offset(anchor, REG_FORECAST_VALUE_OFFSET)
    forecast_max_row, forecast_max_col = add_offset(anchor, REG_FORECAST_MAX_OFFSET)
    forecast_min_row, forecast_min_col = add_offset(anchor, REG_FORECAST_MIN_OFFSET)
    actual_value_row, actual_value_col = add_offset(anchor, REG_ACTUAL_VALUE_OFFSET)

    data_end_row = anchor.row - 1
    available_depth = min(
        contiguous_numeric_depth(sheet, y_col, data_end_row, N_QUARTERS * 3),
        contiguous_numeric_depth(sheet, x_col, data_end_row, N_QUARTERS * 3),
    )
    iterations = min(N_QUARTERS, available_depth if available_depth > 0 else N_QUARTERS)

    num_qtr_input_cell = sheet.cells(num_qtr_input_row, num_qtr_input_col)
    intercept_cell = sheet.cells(intercept_row, intercept_col)
    slope_cell = sheet.cells(slope_row, slope_col)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters in range(1, iterations + 1):
        start_row = data_end_row - n_quarters + 1
        if start_row < 1:
            continue

        num_qtr_input_cell.value = n_quarters
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = to_float(read_cell(sheet, forecast_value_row, forecast_value_col))
        forecast_max = to_float(read_cell(sheet, forecast_max_row, forecast_max_col))
        forecast_min = to_float(read_cell(sheet, forecast_min_row, forecast_min_col))
        actual_value = to_float(read_cell(sheet, actual_value_row, actual_value_col))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            intercept,
            slope,
            forecast_value,
            forecast_max,
            forecast_min,
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

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: Iterable[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)
    for col_idx in range(1, len(columns) + 1):
        ws.cell(1, col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for cell in ws[get_column_letter(col_idx)]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def iter_input_files(path: Path) -> Iterable[Path]:
    for item in sorted(path.iterdir()):
        yield item


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input directory does not exist: {input_dir}")

    output_path = next_output_path(input_dir, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_input_files(input_dir):
            if file_path.is_dir():
                print(f"skipped file: {file_path.name} (is a directory)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            try:
                labels = parse_file_labels(file_path)
            except Exception as exc:
                print(f"skipped file: {file_path.name} ({exc})")
                continue

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, labels, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, labels, file_path.name))
                files_processed += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} ({exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"output path: {output_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
