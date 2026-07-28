#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths for your environment.
input_dir = Path("./input")
output_dir = Path("./output")

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
EMPIRICAL_N_QUARTERS = 10

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

# Anchor-based offsets relative to the "max" cell.
EMPIRICAL_OFFSETS = {
    "last_quarter_used": -9,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "avg_penetration_pct": -4,
    "forecast_value": -3,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_OFFSETS = {
    "forecast_value": -1,  # TOT FCST w/o SA
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}

MONTH_LOOKUP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

DAY_LOOKUP = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass(frozen=True)
class ParsedFileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def float_signature(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value, 10)


def parse_model_labels(file_name: str) -> ParsedFileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,10}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_match = re.search(
        r"\b(Early|Mid|Late)\s*([A-Za-z]+)\s*[_-]?(\d{4})\b",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "UNKNOWN"
    model_date = ""
    if period_match:
        period_part = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))

        month_text = f"{month_token[:1].upper()}{month_token[1:].lower()}"
        month_num = MONTH_LOOKUP.get(month_token.lower())
        if month_num is None and len(month_token) >= 3:
            month_num = MONTH_LOOKUP.get(month_token[:3].lower())

        model_period = f"{period_part}{month_text}_{year}"
        if month_num is not None:
            model_date = datetime(year, month_num, DAY_LOOKUP[period_part]).date().isoformat()

    model = f"{ticker}_{model_period}" if model_period != "UNKNOWN" else ticker
    return ParsedFileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def list_source_files(folder: Path) -> list[Path]:
    paths: list[Path] = []
    for file_path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            print(f"SKIPPED: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"SKIPPED: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"SKIPPED: {file_path.name} (not .xlsx)")
            continue
        paths.append(file_path)
    return paths


def safe_get_sheet(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    workbook.api.Close(SaveChanges=False)


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    # Prefer Formula2R1C1 so we stay in R1C1 mode and avoid column-letter conversion.
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    # xlwings fallback if Formula2R1C1 is unavailable in this environment.
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used_range = sheet.used_range
    values = used_range.value
    if values is None:
        return None

    if isinstance(values, list):
        if values and not isinstance(values[0], list):
            matrix = [values]
        else:
            matrix = values
    else:
        matrix = [[values]]

    for row_idx, row_values in enumerate(matrix):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == anchor_text:
                return used_range.row + row_idx, used_range.column + col_idx
    return None


def read_with_anchor_offset(sheet: xw.Sheet, row: int, anchor_col: int, offset: int) -> Any:
    return sheet.range((row, anchor_col + offset)).value


def process_empirical_sheet(
    workbook: xw.Book,
    parsed: ParsedFileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = safe_get_sheet(workbook, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    rows: list[dict[str, Any]] = []

    avg_col = anchor_col + EMPIRICAL_OFFSETS["avg_penetration_pct"]
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
    relative_penetration_col = penetration_col - avg_col

    for idx in range(EMPIRICAL_N_QUARTERS):
        target_row = start_row + idx
        avg_cell = sheet.range((target_row, avg_col))
        if idx == 0:
            formula = f'=IFERROR(RC[{relative_penetration_col}], "")'
        else:
            formula = (
                f'=IFERROR(AVERAGE(R[-{idx}]C[{relative_penetration_col}]'
                f':RC[{relative_penetration_col}]), "")'
            )
        set_formula2_r1c1(avg_cell, formula)

    workbook.app.calculate()

    for idx in range(EMPIRICAL_N_QUARTERS):
        row_num = start_row + idx
        num_quarters_used = idx + 1

        last_quarter_used = read_with_anchor_offset(
            sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["last_quarter_used"]
        )
        quarterly_sales = to_float(
            read_with_anchor_offset(sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["quarterly_sales"])
        )
        reported_sales = to_float(
            read_with_anchor_offset(sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["reported_sales"])
        )
        growth_rate_pct = to_float(
            read_with_anchor_offset(sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["growth_rate_pct"])
        )
        sales_captured_in_db_pct = to_float(
            read_with_anchor_offset(
                sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
            )
        )
        avg_penetration_pct = to_float(
            read_with_anchor_offset(sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["avg_penetration_pct"])
        )
        forecast_value = to_float(
            read_with_anchor_offset(sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["forecast_value"])
        )
        forecast_max = to_float(
            read_with_anchor_offset(sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["forecast_max"])
        )
        forecast_min = to_float(
            read_with_anchor_offset(sheet, row_num, anchor_col, EMPIRICAL_OFFSETS["forecast_min"])
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

        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )

        rows.append(
            {
                "model": parsed.model,
                "ticker": parsed.ticker,
                "model_period": parsed.model_period,
                "model_date": parsed.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
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

    return rows


def collect_regression_points(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    anchor_row: int,
    max_points: int = 64,
) -> list[tuple[int, float, float]]:
    points: list[tuple[int, float, float]] = []
    started = False

    for row in range(anchor_row - 1, 0, -1):
        x_value = to_float(sheet.range((row, x_col)).value)
        y_value = to_float(sheet.range((row, y_col)).value)

        if x_value is None or y_value is None:
            if started:
                break
            continue

        started = True
        points.append((row, x_value, y_value))

        if len(points) >= max_points:
            break

    points.reverse()
    return points


def process_regression_sheet(
    workbook: xw.Book,
    parsed: ParsedFileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = safe_get_sheet(workbook, REGRESSION_SHEET_NAME)
    if sheet is None:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    points = collect_regression_points(sheet, x_col=x_col, y_col=y_col, anchor_row=anchor_row)
    if len(points) < 2:
        return []

    max_quarters = min(10, len(points))

    temp_row_start = anchor_row + 2
    temp_intercept_col = anchor_col + 6
    temp_slope_col = anchor_col + 7
    temp_forecast_col = anchor_col + 8

    calc_rows: list[tuple[int, int]] = []

    for num_quarters_used in range(2, max_quarters + 1):
        start_data_row = points[-num_quarters_used][0]
        end_data_row = points[-1][0]
        calc_row = temp_row_start + (num_quarters_used - 2)

        intercept_cell = sheet.range((calc_row, temp_intercept_col))
        slope_cell = sheet.range((calc_row, temp_slope_col))
        forecast_cell = sheet.range((calc_row, temp_forecast_col))

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},'
            f'R{start_data_row}C{x_col}:R{end_data_row}C{x_col}), "")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},'
            f'R{start_data_row}C{x_col}:R{end_data_row}C{x_col}), "")'
        )

        next_x_value = to_float(sheet.range((end_data_row + 1, x_col)).value)
        if next_x_value is None:
            next_x_value = points[-1][1] + 1.0

        forecast_formula = (
            f'=IFERROR(RC[{temp_intercept_col - temp_forecast_col}] + '
            f'RC[{temp_slope_col - temp_forecast_col}] * {next_x_value}, "")'
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)
        calc_rows.append((num_quarters_used, calc_row))

    workbook.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[float | None, ...] | None = None

    for num_quarters_used, calc_row in calc_rows:
        intercept = to_float(sheet.range((calc_row, temp_intercept_col)).value)
        slope = to_float(sheet.range((calc_row, temp_slope_col)).value)
        forecast_value = to_float(sheet.range((calc_row, temp_forecast_col)).value)

        source_row = anchor_row + (num_quarters_used - 1)
        source_forecast = to_float(
            read_with_anchor_offset(sheet, source_row, anchor_col, REGRESSION_OFFSETS["forecast_value"])
        )
        if source_forecast is not None:
            forecast_value = source_forecast

        actual_value = read_with_anchor_offset(
            sheet, source_row, anchor_col, REGRESSION_OFFSETS["actual_value"]
        )
        if actual_value is None:
            actual_value = ""

        forecast_max = to_float(
            read_with_anchor_offset(sheet, source_row, anchor_col, REGRESSION_OFFSETS["forecast_max"])
        )
        forecast_min = to_float(
            read_with_anchor_offset(sheet, source_row, anchor_col, REGRESSION_OFFSETS["forecast_min"])
        )
        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )

        signature = (
            float_signature(forecast_value),
            float_signature(forecast_max),
            float_signature(forecast_min),
            float_signature(intercept),
            float_signature(slope),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        if all(value in ("", None) for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        rows.append(
            {
                "model": parsed.model,
                "ticker": parsed.ticker,
                "model_period": parsed.model_period,
                "model_date": parsed.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
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

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    headers: list[str],
    rows: list[dict[str, Any]],
) -> None:
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.append(headers)

    for row_data in rows:
        sheet.append([row_data.get(header, "") for header in headers])

    header_font = Font(bold=True)
    for cell in sheet[1]:
        cell.font = header_font

    sheet.freeze_panes = "A2"

    last_col = get_column_letter(sheet.max_column)
    last_row = max(sheet.max_row, 1)
    sheet.auto_filter.ref = f"A1:{last_col}{last_row}"

    for col_idx, column in enumerate(sheet.columns, start=1):
        max_len = 0
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_len:
                max_len = len(value)
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    workbook.save(output_file)


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a directory: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = unique_output_path(input_dir, output_dir)

    source_files = list_source_files(input_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in source_files:
            print(f"PROCESSING: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                parsed = parse_model_labels(file_path.name)

                empirical_rows.extend(
                    process_empirical_sheet(workbook=workbook, parsed=parsed, source_file=file_path.name)
                )
                regression_rows.extend(
                    process_regression_sheet(workbook=workbook, parsed=parsed, source_file=file_path.name)
                )

                files_processed += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_file}")
    print(f"FILES_PROCESSED: {files_processed}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
