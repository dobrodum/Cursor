#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import math
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ====== User inputs ======
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")
# =========================

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
ANCHOR_LABEL = "max"
N_QUARTERS = 10

# Anchor-based offsets (relative to the "max" anchor column).
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "quarterly_sales": -7,
    "reported_sales": -6,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -4,
    "forecast_value": -3,  # estimated total sold
    "actual_value": -2,  # reported sales
    "forecast_max": 0,
    "forecast_min": 1,
}
EMPIRICAL_PENETRATION_COL_OFFSET = -4

REGRESSION_OFFSETS = {
    "num_quarters_used": -8,
    "actual_value": -2,
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
}

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

DAY_BY_PERIOD_PREFIX = {"early": 5, "mid": 15, "late": 25}
FILE_PERIOD_RE = re.compile(
    r"^(?P<prefix>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    first_item = values[0]
    if isinstance(first_item, (list, tuple)):
        return [list(row) if isinstance(row, tuple) else row for row in values]
    return [list(values) if isinstance(values, tuple) else values]


def clean_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return value


def as_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def numeric_subtract(a: Any, b: Any) -> float | str:
    a_num = as_float(a)
    b_num = as_float(b)
    if a_num is None or b_num is None:
        return ""
    return a_num - b_num


def get_sheet_by_name(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    target_name = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target_name:
            return sheet
    return None


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    anchor_word = ANCHOR_LABEL.lower()
    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == anchor_word:
                return used.row + row_idx, used.column + col_idx
    return None


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()

    period_token = ""
    if len(parts) >= 3 and parts[2]:
        period_token = parts[2].split("_")[0].strip()

    model_period = ""
    model_date = ""
    period_match = FILE_PERIOD_RE.match(period_token)
    if period_match:
        prefix = period_match.group("prefix").title()
        month = period_match.group("month").title()
        year = int(period_match.group("year"))
        model_period = f"{prefix}{month}_{year}"
        try:
            month_num = datetime.strptime(month, "%b").month
            day = DAY_BY_PERIOD_PREFIX[prefix.lower()]
            model_date = date(year, month_num, day).isoformat()
        except ValueError:
            model_date = ""
    elif period_token:
        model_period = period_token

    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    target_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def signature_value(value: Any) -> Any:
    num = as_float(value)
    if num is not None:
        return round(num, 10)
    return clean_value(value)


def close_workbook_without_saving(workbook: xw.Book, source_name: str) -> None:
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

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception as exc:
        print(f"  Warning: unable to safely close {source_name}: {exc}")


def extract_empirical_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    meta: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  Skipped empirical in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    history_end_row = anchor_row - 1
    penetration_col = anchor_col + EMPIRICAL_PENETRATION_COL_OFFSET
    helper_row = anchor_row + N_QUARTERS + 5
    helper_col = anchor_col + 6

    rows: list[dict[str, Any]] = []
    for i in range(N_QUARTERS):
        data_row = anchor_row + 1 + i
        n_raw = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"])
        n_value = as_float(n_raw)
        num_quarters_used = int(n_value) if n_value is not None else (i + 1)
        start_row = max(1, history_end_row - num_quarters_used + 1)

        if history_end_row >= start_row and penetration_col > 0:
            avg_formula_r1c1 = (
                f"=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{history_end_row}C{penetration_col}),\"\")"
            )
        else:
            avg_formula_r1c1 = "=\"\""

        helper_cell = sheet.cells(helper_row, helper_col)
        helper_cell.formula2 = avg_formula_r1c1
        workbook.app.calculate()
        avg_penetration_pct = helper_cell.value

        forecast_value = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["forecast_value"])
        forecast_max = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["forecast_max"])
        forecast_min = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["forecast_min"])
        last_quarter_used = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"])
        quarterly_sales = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"])
        reported_sales = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["reported_sales"])
        actual_value = reported_sales
        growth_rate_pct = read_cell(sheet, data_row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"])
        sales_captured_in_db_pct = read_cell(
            sheet,
            data_row,
            anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
        )
        range_width = numeric_subtract(forecast_max, forecast_min)

        has_payload = any(
            clean_value(value) != ""
            for value in (forecast_value, actual_value, forecast_max, forecast_min, avg_penetration_pct)
        )
        if not has_payload:
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": clean_value(avg_penetration_pct),
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": clean_value(last_quarter_used),
                "forecast_value": clean_value(forecast_value),
                "actual_value": clean_value(actual_value),
                "forecast_max": clean_value(forecast_max),
                "forecast_min": clean_value(forecast_min),
                "range_width": clean_value(range_width),
                "avg_penetration_pct": clean_value(avg_penetration_pct),
                "quarterly_sales": clean_value(quarterly_sales),
                "reported_sales": clean_value(reported_sales),
                "growth_rate_pct": clean_value(growth_rate_pct),
                "sales_captured_in_db_pct": clean_value(sales_captured_in_db_pct),
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    meta: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  Skipped regression in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = anchor_row - 1

    helper_row = anchor_row + N_QUARTERS + 5
    intercept_cell = sheet.cells(helper_row, anchor_col + 6)
    slope_cell = sheet.cells(helper_row, anchor_col + 7)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for i in range(N_QUARTERS):
        data_row = anchor_row + 1 + i
        n_raw = read_cell(sheet, data_row, anchor_col + REGRESSION_OFFSETS["num_quarters_used"])
        n_value = as_float(n_raw)
        num_quarters_used = int(n_value) if n_value is not None else (i + 1)
        start_row = max(1, history_end_row - num_quarters_used + 1)

        if history_end_row >= start_row and x_col > 0 and y_col > 0:
            intercept_formula_r1c1 = (
                f"=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{history_end_row}C{x_col}),\"\")"
            )
            slope_formula_r1c1 = (
                f"=IFERROR(SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{history_end_row}C{x_col}),\"\")"
            )
        else:
            intercept_formula_r1c1 = "=\"\""
            slope_formula_r1c1 = "=\"\""

        intercept_cell.formula2 = intercept_formula_r1c1
        slope_cell.formula2 = slope_formula_r1c1
        workbook.app.calculate()
        intercept = intercept_cell.value
        slope = slope_cell.value

        forecast_value = read_cell(sheet, data_row, anchor_col + REGRESSION_OFFSETS["forecast_value"])
        actual_value = read_cell(sheet, data_row, anchor_col + REGRESSION_OFFSETS["actual_value"])
        forecast_max = read_cell(sheet, data_row, anchor_col + REGRESSION_OFFSETS["forecast_max"])
        forecast_min = read_cell(sheet, data_row, anchor_col + REGRESSION_OFFSETS["forecast_min"])
        range_width = numeric_subtract(forecast_max, forecast_min)

        row_signature = (
            signature_value(num_quarters_used),
            signature_value(intercept),
            signature_value(slope),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
        )
        if i == N_QUARTERS - 1 and previous_signature == row_signature:
            continue
        previous_signature = row_signature

        has_payload = any(
            clean_value(value) != ""
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        )
        if not has_payload:
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": clean_value(forecast_value),
                "actual_value": clean_value(actual_value),
                "forecast_max": clean_value(forecast_max),
                "forecast_min": clean_value(forecast_min),
                "range_width": clean_value(range_width),
                "intercept": clean_value(intercept),
                "slope": clean_value(slope),
                "source_file": source_file,
            }
        )

    return rows


def set_column_widths(sheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = clean_value(row.get(header, ""))
            cell_text = str(value)
            if len(cell_text) > max_len:
                max_len = len(cell_text)
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_sheet(sheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    sheet.append(headers)
    for row in rows:
        sheet.append([clean_value(row.get(column, "")) for column in headers])

    for header_cell in sheet[1]:
        header_cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    set_column_widths(sheet, headers, rows)


def write_output_workbook(
    path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_output_sheet(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)
    write_output_sheet(regression_sheet, REGRESSION_HEADERS, regression_rows)
    workbook.save(path)


def main() -> None:
    source_dir = Path(input_dir)
    target_dir = Path(output_dir)
    if not source_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {source_dir}")

    output_path = build_output_path(source_dir, target_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(source_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                files_processed += 1
                metadata = parse_file_metadata(file_path.name)

                empirical_sheet = get_sheet_by_name(workbook, EMPIRICAL_SHEET_NAME)
                if empirical_sheet is None:
                    print(f"  Skipped empirical in {file_path.name}: sheet '{EMPIRICAL_SHEET_NAME}' not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(workbook, empirical_sheet, metadata, file_path.name)
                    )

                regression_sheet = get_sheet_by_name(workbook, REGRESSION_SHEET_NAME)
                if regression_sheet is None:
                    print(
                        f"  Skipped regression in {file_path.name}: sheet '{REGRESSION_SHEET_NAME}' not found"
                    )
                else:
                    regression_rows.extend(
                        extract_regression_rows(workbook, regression_sheet, metadata, file_path.name)
                    )
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_workbook_without_saving(workbook, file_path.name)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
