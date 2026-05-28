#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw

# ----------------------
# User-configurable I/O
# ----------------------
input_dir = Path("input")
output_dir = Path("output")

EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"
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

# Anchor-based offsets relative to the "max" cell in each sheet.
EMPIRICAL_OFFSETS = {
    "data_start_row": 1,
    "num_quarters_used": -5,
    "last_quarter_used": -4,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -7,
    "reported_sales": -2,
    "growth_rate_pct": 2,
    "sales_captured_in_db_pct": 3,
}
EMPIRICAL_TEMP_AVG_PEN_COL_OFFSET = 8

REGRESSION_OFFSETS = {
    "data_start_row": 1,
    "num_quarters_used": -4,
    "forecast_value": -1,  # TOT FCST w/o SA
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}
REGRESSION_TEMP_INTERCEPT_COL_OFFSET = 8
REGRESSION_TEMP_SLOPE_COL_OFFSET = 9

MONTH_TO_NUMBER = {
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
PERIOD_TO_DAY = {"early": 5, "mid": 15, "late": 25}
PERIOD_TOKEN_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})", re.IGNORECASE)


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


def to_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> float | None:
    if value is None or value == "" or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def canonical_numeric(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value, 10)


def safe_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    # .formula2 keeps all formulas in R1C1 without column-letter conversion.
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    used_values = to_2d(used.value)
    target = anchor_text.strip().lower()
    start_row = used.row
    start_col = used.column

    for row_index, row_values in enumerate(used_values):
        for col_index, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return start_row + row_index, start_col + col_index
    return None


def get_sheet_by_name(book: xw.Book, sheet_name: str) -> xw.Sheet | None:
    target = sheet_name.strip().lower()
    for sheet in book.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def close_source_book(book: xw.Book) -> None:
    try:
        book.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        book.close(False)
        return
    except Exception:
        pass

    try:
        book.api.Close(SaveChanges=False)
    except Exception:
        try:
            book.close()
        except Exception:
            pass


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]
    if len(parts) < 3:
        raise ValueError("filename does not match '... - TICKER - PeriodYear_Send.xlsx'")

    ticker = parts[-2].upper()
    period_token_source = parts[-1]
    period_match = PERIOD_TOKEN_RE.search(period_token_source)
    if not period_match:
        raise ValueError("missing period token like EarlyJan2026, MidJan2026, LateJan2026")

    period_label = period_match.group(1).capitalize()
    month_abbrev = period_match.group(2).title()
    year_text = period_match.group(3)

    month_number = MONTH_TO_NUMBER.get(month_abbrev.lower())
    if month_number is None:
        raise ValueError(f"unsupported month token: {month_abbrev}")

    day = PERIOD_TO_DAY[period_label.lower()]
    model_period = f"{period_label}{month_abbrev}_{year_text}"
    model_date = date(int(year_text), month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_name,
    )


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    base_path = out_dir / base_name
    if not base_path.exists():
        return base_path

    suffix = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def should_skip_file(file_path: Path) -> str | None:
    if file_path.is_dir():
        return "is a directory"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    return None


def process_empirical_sheet(book: xw.Book, labels: FileLabels) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(book, EMPIRICAL_MODEL_SHEET)
    if sheet is None:
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + EMPIRICAL_OFFSETS["data_start_row"]
    temp_avg_pen_col = anchor_col + EMPIRICAL_TEMP_AVG_PEN_COL_OFFSET

    formula_rows: list[int] = []
    sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]
    reported_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales"]

    for idx in range(N_QUARTERS):
        row_idx = start_row + idx
        n_quarters = idx + 1
        history_start_row = max(1, row_idx - n_quarters + 1)
        avg_pen_formula = (
            f'=IFERROR(AVERAGE(R{history_start_row}C{sales_col}:R{row_idx}C{sales_col}'
            f"/R{history_start_row}C{reported_col}:R{row_idx}C{reported_col}),\"\")"
        )
        set_formula2(sheet.cells(row_idx, temp_avg_pen_col), avg_pen_formula)
        formula_rows.append(row_idx)

    if formula_rows:
        book.app.calculate()

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_idx = start_row + idx
        num_quarters = to_int(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"]))
        if num_quarters is None:
            num_quarters = idx + 1

        last_quarter_used = safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"])
        forecast_value = to_float(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_value"]))
        actual_value = to_float(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["actual_value"]))
        forecast_max = to_float(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_max"]))
        forecast_min = to_float(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_min"]))
        avg_penetration_pct = to_float(safe_cell_value(sheet, row_idx, temp_avg_pen_col))
        quarterly_sales = to_float(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]))
        reported_sales = to_float(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["reported_sales"]))
        growth_rate_pct = to_float(safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]))
        sales_captured_pct = to_float(
            safe_cell_value(sheet, row_idx, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"])
        )

        if not any(
            value is not None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
                avg_penetration_pct,
            )
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
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
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": labels.source_file,
            }
        )

    for row_idx in formula_rows:
        sheet.cells(row_idx, temp_avg_pen_col).value = None

    return rows


def process_regression_sheet(book: xw.Book, labels: FileLabels) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(book, REGRESSION_MODEL_SHEET)
    if sheet is None:
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + REGRESSION_OFFSETS["data_start_row"]
    x_col = anchor_col - 11
    y_col = anchor_col - 7

    temp_intercept_col = anchor_col + REGRESSION_TEMP_INTERCEPT_COL_OFFSET
    temp_slope_col = anchor_col + REGRESSION_TEMP_SLOPE_COL_OFFSET

    formula_specs: list[tuple[int, int]] = []
    for idx in range(N_QUARTERS):
        row_idx = start_row + idx
        num_quarters = to_int(safe_cell_value(sheet, row_idx, anchor_col + REGRESSION_OFFSETS["num_quarters_used"]))
        if num_quarters is None or num_quarters < 1:
            num_quarters = idx + 1

        history_end_row = anchor_row - 1
        history_start_row = max(1, history_end_row - num_quarters + 1)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{history_start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{history_start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )

        set_formula2(sheet.cells(row_idx, temp_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row_idx, temp_slope_col), slope_formula)
        formula_specs.append((row_idx, num_quarters))

    if formula_specs:
        book.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for row_idx, num_quarters in formula_specs:
        intercept = to_float(safe_cell_value(sheet, row_idx, temp_intercept_col))
        slope = to_float(safe_cell_value(sheet, row_idx, temp_slope_col))
        forecast_value = to_float(safe_cell_value(sheet, row_idx, anchor_col + REGRESSION_OFFSETS["forecast_value"]))

        if forecast_value is None and intercept is not None and slope is not None:
            x_value = to_float(safe_cell_value(sheet, row_idx, x_col))
            if x_value is not None:
                forecast_value = intercept + slope * x_value

        actual_value = to_float(safe_cell_value(sheet, row_idx, anchor_col + REGRESSION_OFFSETS["actual_value"]))
        forecast_max = to_float(safe_cell_value(sheet, row_idx, anchor_col + REGRESSION_OFFSETS["forecast_max"]))
        forecast_min = to_float(safe_cell_value(sheet, row_idx, anchor_col + REGRESSION_OFFSETS["forecast_min"]))

        if not any(value is not None for value in (forecast_value, actual_value, forecast_max, forecast_min, intercept, slope)):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        current_signature = (
            num_quarters,
            canonical_numeric(forecast_value),
            canonical_numeric(forecast_max),
            canonical_numeric(forecast_min),
            canonical_numeric(intercept),
            canonical_numeric(slope),
        )
        if previous_signature == current_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": labels.source_file,
            }
        )

    for row_idx, _ in formula_specs:
        sheet.cells(row_idx, temp_intercept_col).value = None
        sheet.cells(row_idx, temp_slope_col).value = None

    return rows


def write_sheet(workbook: Workbook, sheet_name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
    sheet = workbook.create_sheet(sheet_name)
    sheet.append(headers)
    for row in rows:
        sheet.append([row.get(header) for header in headers])

    header_font = Font(bold=True)
    for cell in sheet[1]:
        cell.font = header_font

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is not None:
                max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, EMPIRICAL_OUTPUT_SHEET, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(workbook, REGRESSION_OUTPUT_SHEET, REGRESSION_HEADERS, regression_rows)

    workbook.save(output_path)


def main() -> int:
    if not input_dir.exists():
        print(f"Input directory does not exist: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    source_paths = sorted(input_dir.iterdir())
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

        for file_path in source_paths:
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            try:
                labels = parse_file_labels(file_path.name)
            except ValueError as exc:
                print(f"Skipped {file_path.name}: {exc}")
                continue

            book: xw.Book | None = None
            try:
                book = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(book, labels))
                regression_rows.extend(process_regression_sheet(book, labels))
                files_processed += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing failed ({exc})")
            finally:
                if book is not None:
                    close_source_book(book)
    finally:
        app.quit()

    output_path = build_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
