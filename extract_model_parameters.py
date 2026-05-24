#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# Configure paths here
# -----------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")

N_QUARTERS = 10

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

# Anchor-based offsets (relative to "max" anchor cell).
EMPIRICAL_OFFSETS = {
    "quarter_col": -11,
    "penetration_col": -10,
    "quarterly_sales_col": -9,
    "reported_sales_col": -8,
    "growth_rate_col": -7,
    "captured_pct_col": -6,
    "scratch_start_row": 2,
    "scratch_col": 14,
}

REGRESSION_OFFSETS = {
    "y_col": -7,   # Required by spec.
    "x_col": -11,  # Required by spec.
    "scratch_start_row": 2,
    "scratch_col": 18,
}

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_NAME = {
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


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def _to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    is_percent = text.endswith("%")
    cleaned = text.replace(",", "").replace("%", "")
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = "-" + cleaned[1:-1]
    try:
        number = float(cleaned)
    except ValueError:
        return None
    return number / 100.0 if is_percent else number


def rounded_signature(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 10)


def parse_filename_metadata(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"\s+", "", parts[1]).upper()
    else:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = re.search(r"(?i)(early|mid|late)([a-z]+)(\d{4})", stem)
    if not period_match:
        model_period = "unknown_period"
        model_date = ""
        return ModelMeta(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date=model_date,
        )

    period_word = period_match.group(1).lower()
    month_word = period_match.group(2).lower()
    year = int(period_match.group(3))

    month_num = MONTH_BY_NAME.get(month_word)
    if month_num is None:
        month_num = MONTH_BY_NAME.get(month_word[:3])
    if month_num is None:
        model_period = "unknown_period"
        model_date = ""
        return ModelMeta(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date=model_date,
        )

    month_abbrev = datetime(year, month_num, 1).strftime("%b")
    period_title = period_word.capitalize()
    model_period = f"{period_title}{month_abbrev}_{year}"
    model_date = date(year, month_num, DAY_BY_PERIOD[period_word]).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(source_dir: Path, destination_dir: Path) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    base_stem = f"{source_dir.name}_PARAM"
    base_path = destination_dir / f"{base_stem}.xlsx"
    if not base_path.exists():
        return base_path

    index = 1
    while True:
        candidate = destination_dir / f"{base_stem}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def list_source_files(source_dir: Path) -> list[Path]:
    files: list[Path] = []
    if not source_dir.exists():
        print(f"Skipped all files: input_dir does not exist: {source_dir}")
        return files

    for file_path in sorted(source_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file: {file_path.name} (temp file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {file_path.name} (not .xlsx)")
            continue
        files.append(file_path)
    return files


def close_workbook_safe(workbook: xw.Book) -> None:
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
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def set_formula2(cell: xw.Range, formula: str) -> None:
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

    try:
        cell.api.FormulaR1C1 = formula
        return
    except Exception:
        # Last fallback keeps script usable if Formula2* and FormulaR1C1 are unavailable.
        cell.formula = formula


def find_max_anchor(sheet: xw.Sheet) -> Optional[tuple[int, int]]:
    used_range = sheet.used_range
    values_2d = _to_2d(used_range.value)
    for r_idx, row in enumerate(values_2d):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return used_range.row + r_idx, used_range.column + c_idx
    return None


def read_max_min_from_anchor(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> tuple[Optional[float], Optional[float]]:
    max_value = to_float(sheet.range((anchor_row, anchor_col + 1)).value)
    min_value = to_float(sheet.range((anchor_row + 1, anchor_col + 1)).value)

    if max_value is None:
        for offset in range(1, 6):
            candidate = to_float(sheet.range((anchor_row, anchor_col + offset)).value)
            if candidate is not None:
                max_value = candidate
                break

    if min_value is None:
        for row_offset in range(1, 8):
            label = normalize_text(sheet.range((anchor_row + row_offset, anchor_col)).value)
            if label == "min":
                min_value = to_float(sheet.range((anchor_row + row_offset, anchor_col + 1)).value)
                break
    return max_value, min_value


def collect_contiguous_numeric_rows(
    sheet: xw.Sheet,
    start_row: int,
    required_cols: Iterable[int],
    max_scan_rows: int = 500,
) -> list[int]:
    rows: list[int] = []
    row = max(start_row, 1)
    scanned = 0
    cols = list(required_cols)
    while row >= 1 and scanned < max_scan_rows:
        scanned += 1
        values = [to_float(sheet.range((row, col)).value) for col in cols]
        if all(value is not None for value in values):
            rows.append(row)
            row -= 1
            continue
        if rows:
            break
        row -= 1
    rows.reverse()
    return rows


def process_empirical_sheet(workbook: xw.Book, meta: ModelMeta, source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Empirical Model"
    sheet_names = {sheet.name for sheet in workbook.sheets}
    if sheet_name not in sheet_names:
        print(f"Skipped empirical extraction: {source_file} (sheet '{sheet_name}' not found)")
        return []

    sheet = workbook.sheets[sheet_name]
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipped empirical extraction: {source_file} (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    quarter_col = anchor_col + EMPIRICAL_OFFSETS["quarter_col"]
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["penetration_col"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"]
    growth_rate_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_col"]
    captured_pct_col = anchor_col + EMPIRICAL_OFFSETS["captured_pct_col"]

    data_rows = collect_contiguous_numeric_rows(
        sheet,
        start_row=anchor_row - 1,
        required_cols=[penetration_col, quarterly_sales_col],
    )
    if not data_rows:
        print(f"Skipped empirical extraction: {source_file} (no numeric empirical rows found)")
        return []

    forecast_max, forecast_min = read_max_min_from_anchor(sheet, anchor_row, anchor_col)
    range_width = None
    if forecast_max is not None and forecast_min is not None:
        range_width = forecast_max - forecast_min

    output_rows: list[dict[str, Any]] = []
    max_quarters = min(N_QUARTERS, len(data_rows))
    scratch_col = anchor_col + EMPIRICAL_OFFSETS["scratch_col"]
    scratch_row_start = anchor_row + EMPIRICAL_OFFSETS["scratch_start_row"]

    for n_quarters in range(1, max_quarters + 1):
        first_row = data_rows[-n_quarters]
        last_row = data_rows[-1]

        calc_row = scratch_row_start + (n_quarters - 1)
        avg_pen_cell = sheet.range((calc_row, scratch_col))
        forecast_cell = sheet.range((calc_row, scratch_col + 1))

        avg_formula = (
            f"=AVERAGE(R{first_row}C{penetration_col}:R{last_row}C{penetration_col})"
        )
        forecast_formula = (
            f'=IFERROR(R{last_row}C{quarterly_sales_col}/R{calc_row}C{scratch_col},"")'
        )
        set_formula2(avg_pen_cell, avg_formula)
        set_formula2(forecast_cell, forecast_formula)

    workbook.app.calculate()

    for n_quarters in range(1, max_quarters + 1):
        first_row = data_rows[-n_quarters]
        last_row = data_rows[-1]

        calc_row = scratch_row_start + (n_quarters - 1)
        avg_pen_cell = sheet.range((calc_row, scratch_col))
        forecast_cell = sheet.range((calc_row, scratch_col + 1))

        avg_penetration = to_float(avg_pen_cell.value)
        forecast_value = to_float(forecast_cell.value)
        quarterly_sales = to_float(sheet.range((last_row, quarterly_sales_col)).value)
        reported_sales = to_float(sheet.range((last_row, reported_sales_col)).value)
        growth_rate = to_float(sheet.range((last_row, growth_rate_col)).value)
        captured_pct = to_float(sheet.range((last_row, captured_pct_col)).value)
        last_quarter_used = sheet.range((first_row, quarter_col)).value

        if avg_penetration is None and forecast_value is None:
            continue

        output_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
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
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return output_rows


def process_regression_sheet(workbook: xw.Book, meta: ModelMeta, source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Regression Model"
    sheet_names = {sheet.name for sheet in workbook.sheets}
    if sheet_name not in sheet_names:
        print(f"Skipped regression extraction: {source_file} (sheet '{sheet_name}' not found)")
        return []

    sheet = workbook.sheets[sheet_name]
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipped regression extraction: {source_file} (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col + REGRESSION_OFFSETS["y_col"]
    x_col = anchor_col + REGRESSION_OFFSETS["x_col"]

    data_rows = collect_contiguous_numeric_rows(
        sheet,
        start_row=anchor_row - 1,
        required_cols=[x_col, y_col],
    )
    if len(data_rows) < 2:
        print(f"Skipped regression extraction: {source_file} (insufficient x/y rows)")
        return []

    forecast_max, forecast_min = read_max_min_from_anchor(sheet, anchor_row, anchor_col)
    range_width = None
    if forecast_max is not None and forecast_min is not None:
        range_width = forecast_max - forecast_min

    output_rows: list[dict[str, Any]] = []
    max_quarters = min(N_QUARTERS, len(data_rows))
    scratch_col = anchor_col + REGRESSION_OFFSETS["scratch_col"]
    scratch_row_start = anchor_row + REGRESSION_OFFSETS["scratch_start_row"]
    previous_signature: Optional[tuple[Optional[float], ...]] = None

    for n_quarters in range(2, max_quarters + 1):
        first_row = data_rows[-n_quarters]
        last_row = data_rows[-1]

        calc_row = scratch_row_start + (n_quarters - 2)
        intercept_cell = sheet.range((calc_row, scratch_col))
        slope_cell = sheet.range((calc_row, scratch_col + 1))
        forecast_cell = sheet.range((calc_row, scratch_col + 2))

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        forecast_formula = (
            f"=R{calc_row}C{scratch_col}+R{calc_row}C{scratch_col + 1}*R{last_row}C{x_col}"
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        set_formula2(forecast_cell, forecast_formula)

    workbook.app.calculate()

    for n_quarters in range(2, max_quarters + 1):
        calc_row = scratch_row_start + (n_quarters - 2)
        intercept_cell = sheet.range((calc_row, scratch_col))
        slope_cell = sheet.range((calc_row, scratch_col + 1))
        forecast_cell = sheet.range((calc_row, scratch_col + 2))

        intercept_value = to_float(intercept_cell.value)
        slope_value = to_float(slope_cell.value)
        forecast_value = to_float(forecast_cell.value)

        current_signature = (
            rounded_signature(intercept_value),
            rounded_signature(slope_value),
            rounded_signature(forecast_value),
            rounded_signature(forecast_max),
            rounded_signature(forecast_min),
        )

        # Prevent duplicate trailing row when the final iteration yields same values.
        if current_signature == previous_signature:
            continue
        previous_signature = current_signature

        output_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return output_rows


def write_output_sheet(worksheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet.append(headers)
    for col_idx, _ in enumerate(headers, start=1):
        worksheet.cell(row=1, column=col_idx).font = Font(bold=True)

    for row_data in rows:
        worksheet.append([row_data.get(column_name, "") for column_name in headers])

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def save_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_output_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def run() -> None:
    files = list_source_files(input_dir)
    output_path = build_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    if files:
        app: Optional[xw.App] = None
        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            try:
                app.calculation = "manual"
            except Exception:
                pass

            for file_path in files:
                print(f"Processing file: {file_path.name}")
                workbook: Optional[xw.Book] = None
                try:
                    workbook = app.books.open(str(file_path), update_links=False)
                    meta = parse_filename_metadata(file_path)
                    empirical_rows.extend(process_empirical_sheet(workbook, meta, file_path.name))
                    regression_rows.extend(process_regression_sheet(workbook, meta, file_path.name))
                    processed_files += 1
                except Exception as exc:
                    print(f"Skipped file: {file_path.name} (processing error: {exc})")
                finally:
                    if workbook is not None:
                        close_workbook_safe(workbook)
        finally:
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    pass

    save_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
