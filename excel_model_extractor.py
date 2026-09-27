#!/usr/bin/env python3
"""
Extract empirical/regression candidate rows from Excel model workbooks.

This script opens each source workbook only once, processes both target model
worksheets while it is open, and writes one combined output workbook.
"""

from __future__ import annotations

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
input_dir = "/workspace/input"
output_dir = "/workspace/output"


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


EMPIRICAL_N_QUARTERS = 10
REGRESSION_N_QUARTERS = 10
ANCHOR_SCAN_WINDOW = 30


EMPIRICAL_ALIASES = {
    "num_quarters_used": ["num quarters used", "quarters used", "num quarters", "n quarters"],
    "last_quarter_used": ["last quarter used", "last qtr used", "last quarter", "last qtr"],
    "forecast_value": [
        "estimated total sold",
        "est total sold",
        "forecast value",
        "forecast",
        "total forecast",
    ],
    "actual_value": ["actual value", "actual sales", "actual", "reported sales"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "avg_penetration_pct": [
        "avg penetration pct",
        "average penetration pct",
        "avg penetration",
        "average penetration",
    ],
    "quarterly_sales": ["quarterly sales", "qtr sales"],
    "reported_sales": ["reported sales", "reported"],
    "growth_rate_pct": ["growth rate pct", "growth rate", "growth %"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db",
    ],
}

EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -11,
    "last_quarter_used": -10,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -3,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
}


REGRESSION_ALIASES = {
    "num_quarters_used": ["num quarters used", "quarters used", "num quarters", "n quarters"],
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "total forecast without sa",
        "forecast total without sa",
        "forecast",
    ],
    "actual_value": ["actual value", "actual sales", "actual", "reported sales"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "intercept": ["intercept"],
    "slope": ["slope"],
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -11,
    "forecast_value": -2,
    "actual_value": None,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": -4,
    "slope": -3,
}


FILENAME_PATTERN = re.compile(
    r"Model\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<timing>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})",
    re.IGNORECASE,
)

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

TIMING_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = re.sub(r"[_\-/()%]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    is_percent = "%" in text
    text = text.replace("%", "")
    try:
        number = float(text)
    except ValueError:
        return None
    if is_percent:
        return number / 100.0
    return number


def to_quarters(value: Any, fallback: int) -> int:
    number = to_number(value)
    if number is None:
        return fallback
    rounded = round(number)
    if abs(number - rounded) < 1e-9:
        return int(rounded)
    return int(number)


def numeric_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def comparison_value(value: Any) -> Any:
    number = to_number(value)
    if number is None:
        return value
    return round(number, 10)


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    match = FILENAME_PATTERN.search(stem)
    if not match:
        return {
            "model": stem,
            "ticker": "",
            "model_period": "",
            "model_date": "",
        }

    ticker = match.group("ticker").upper()
    timing_raw = match.group("timing").lower()
    month_raw = match.group("month")
    year_raw = match.group("year")

    month_number = MONTH_LOOKUP.get(month_raw.lower())
    if month_number is None:
        month_number = MONTH_LOOKUP.get(month_raw[:3].lower())

    month_abbrev = month_raw[:3].title()
    timing_label = timing_raw.title()
    model_period = f"{timing_label}{month_abbrev}_{year_raw}"

    model_date = ""
    if month_number is not None:
        day = TIMING_DAY[timing_raw]
        model_date = date(int(year_raw), month_number, day).isoformat()

    return {
        "model": f"{ticker}_{model_period}",
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.name
    base_name = f"{folder_name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def safe_close_source_workbook(workbook: xw.Book) -> None:
    close_attempts = [
        lambda: workbook.close(save=False),
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.api.Close(False),
    ]
    for close_attempt in close_attempts:
        try:
            close_attempt()
            return
        except TypeError:
            continue
        except Exception:
            continue


def ensure_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def get_sheet(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    base_row = used.row
    base_col = used.column
    candidates: List[Tuple[int, int, int]] = []

    for row_index, row_values in enumerate(values):
        for col_index, raw_value in enumerate(row_values):
            if normalize_text(raw_value) != "max":
                continue
            has_min_to_right = False
            if col_index + 1 < len(row_values):
                has_min_to_right = normalize_text(row_values[col_index + 1]) == "min"
            absolute_row = base_row + row_index
            absolute_col = base_col + col_index
            priority = 0 if has_min_to_right else 1
            candidates.append((priority, absolute_row, absolute_col))

    if not candidates:
        return None

    candidates.sort()
    _, anchor_row, anchor_col = candidates[0]
    return anchor_row, anchor_col


def infer_offsets(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    aliases: Dict[str, List[str]],
    defaults: Dict[str, Optional[int]],
) -> Dict[str, Optional[int]]:
    start_col = max(1, anchor_col - ANCHOR_SCAN_WINDOW)
    end_col = anchor_col + ANCHOR_SCAN_WINDOW
    header_values = sheet.range((anchor_row, start_col), (anchor_row, end_col)).value
    if not isinstance(header_values, list):
        header_values = [header_values]

    offsets: Dict[str, Optional[int]] = dict(defaults)
    for local_col_index, raw_header in enumerate(header_values):
        header = normalize_text(raw_header)
        if not header:
            continue
        absolute_col = start_col + local_col_index
        offset = absolute_col - anchor_col
        for field_name, field_aliases in aliases.items():
            if field_name == "forecast_max":
                continue
            if any(alias in header for alias in field_aliases):
                current = offsets.get(field_name)
                if current is None or abs(offset) < abs(current):
                    offsets[field_name] = offset

    offsets["forecast_max"] = 0
    if offsets.get("forecast_min") is None:
        offsets["forecast_min"] = 1
    return offsets


def read_offset_cell(sheet: xw.Sheet, row: int, anchor_col: int, offset: Optional[int]) -> Any:
    if offset is None:
        return None
    absolute_col = anchor_col + offset
    if absolute_col < 1:
        return None
    return sheet.range((row, absolute_col)).value


def find_last_numeric_row(sheet: xw.Sheet, column: int, max_row: int) -> Optional[int]:
    if column < 1 or max_row < 1:
        return None
    values = sheet.range((1, column), (max_row, column)).value
    if not isinstance(values, list):
        values = [values]
    for idx in range(len(values) - 1, -1, -1):
        if to_number(values[idx]) is not None:
            return idx + 1
    return None


def find_last_contiguous_numeric_block(
    sheet: xw.Sheet, column: int, max_row: int
) -> Optional[Tuple[int, int]]:
    end_row = find_last_numeric_row(sheet, column, max_row)
    if end_row is None:
        return None

    values = sheet.range((1, column), (end_row, column)).value
    if not isinstance(values, list):
        values = [values]

    start_row = end_row
    cursor = end_row - 1
    while cursor >= 1:
        if to_number(values[cursor - 1]) is None:
            break
        start_row = cursor
        cursor -= 1
    return start_row, end_row


def process_empirical_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print("  empirical skipped: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    offsets = infer_offsets(sheet, anchor_row, anchor_col, EMPIRICAL_ALIASES, EMPIRICAL_DEFAULT_OFFSETS)

    helper_avg_col = sheet.used_range.last_cell.column + 3
    penetration_offset = offsets.get("sales_captured_in_db_pct")
    if penetration_offset is None:
        penetration_offset = offsets.get("avg_penetration_pct")
    penetration_col = anchor_col + (penetration_offset or 0)

    formulas_written = False
    data_end_row = find_last_numeric_row(sheet, penetration_col, anchor_row - 1)
    if data_end_row is not None and penetration_col > 0:
        for n_quarters in range(1, EMPIRICAL_N_QUARTERS + 1):
            start_row = data_end_row - n_quarters + 1
            if start_row < 1:
                continue
            target_row = anchor_row + n_quarters
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:'
                f"R{data_end_row}C{penetration_col}),\"\")"
            )
            sheet.range((target_row, helper_avg_col)).formula2 = formula
            formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, EMPIRICAL_N_QUARTERS + 1):
        row_idx = anchor_row + n_quarters
        num_quarters_raw = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("num_quarters_used"))
        num_quarters_used = to_quarters(num_quarters_raw, n_quarters)

        last_quarter_used = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("last_quarter_used"))
        forecast_value = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("forecast_value"))
        forecast_max = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("forecast_max"))
        forecast_min = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("forecast_min"))
        avg_penetration_pct = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("avg_penetration_pct"))

        if (avg_penetration_pct is None or avg_penetration_pct == "") and formulas_written:
            avg_penetration_pct = sheet.range((row_idx, helper_avg_col)).value

        reported_sales = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("reported_sales"))
        actual_value = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("actual_value"))
        if actual_value in (None, ""):
            actual_value = reported_sales

        quarterly_sales = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("quarterly_sales"))
        growth_rate_pct = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("growth_rate_pct"))
        sales_captured_in_db_pct = read_offset_cell(
            sheet, row_idx, anchor_col, offsets.get("sales_captured_in_db_pct")
        )

        if forecast_value in (None, ""):
            max_num = to_number(forecast_max)
            min_num = to_number(forecast_min)
            if max_num is not None and min_num is not None:
                forecast_value = (max_num + min_num) / 2.0

        if all(
            value in (None, "")
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
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print("  regression skipped: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    offsets = infer_offsets(sheet, anchor_row, anchor_col, REGRESSION_ALIASES, REGRESSION_DEFAULT_OFFSETS)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print("  regression skipped: anchor offsets produce invalid x/y columns")
        return []

    numeric_block = find_last_contiguous_numeric_block(sheet, y_col, anchor_row - 1)
    if numeric_block is None:
        print("  regression skipped: no numeric y-series found")
        return []

    data_start_row, data_end_row = numeric_block
    helper_intercept_col = sheet.used_range.last_cell.column + 6
    helper_slope_col = helper_intercept_col + 1

    formulas_written = False
    for n_quarters in range(1, REGRESSION_N_QUARTERS + 1):
        start_row = data_end_row - n_quarters + 1
        if start_row < data_start_row:
            continue
        target_row = anchor_row + n_quarters
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},'
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col}),\"\")"
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},'
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col}),\"\")"
        )
        sheet.range((target_row, helper_intercept_col)).formula2 = intercept_formula
        sheet.range((target_row, helper_slope_col)).formula2 = slope_formula
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    next_x_value = to_number(sheet.range((data_end_row + 1, x_col)).value)
    if next_x_value is None:
        last_x = to_number(sheet.range((data_end_row, x_col)).value)
        if last_x is not None:
            next_x_value = last_x + 1

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for n_quarters in range(1, REGRESSION_N_QUARTERS + 1):
        row_idx = anchor_row + n_quarters

        num_quarters_raw = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("num_quarters_used"))
        num_quarters_used = to_quarters(num_quarters_raw, n_quarters)

        intercept = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("intercept"))
        slope = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("slope"))
        if intercept in (None, "") and formulas_written:
            intercept = sheet.range((row_idx, helper_intercept_col)).value
        if slope in (None, "") and formulas_written:
            slope = sheet.range((row_idx, helper_slope_col)).value

        forecast_value = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("forecast_value"))
        if forecast_value in (None, ""):
            intercept_num = to_number(intercept)
            slope_num = to_number(slope)
            if intercept_num is not None and slope_num is not None and next_x_value is not None:
                forecast_value = intercept_num + slope_num * next_x_value

        actual_value = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("actual_value"))
        forecast_max = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("forecast_max"))
        forecast_min = read_offset_cell(sheet, row_idx, anchor_col, offsets.get("forecast_min"))

        if forecast_max in (None, "") and forecast_value not in (None, ""):
            forecast_max = forecast_value
        if forecast_min in (None, "") and forecast_value not in (None, ""):
            forecast_min = forecast_value

        if all(
            value in (None, "")
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        signature = (
            comparison_value(forecast_value),
            comparison_value(forecast_max),
            comparison_value(forecast_min),
            comparison_value(intercept),
            comparison_value(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            continue

        previous_signature = signature
        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(columns)
    for row_data in rows:
        ws.append([row_data.get(column_name) for column_name in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col_letter}{max(ws.max_row, 1)}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            if isinstance(value, float):
                text = f"{value:.8g}"
            else:
                text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_file: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        print(f"input_dir not found or not a directory: {input_path}")
        return

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    try:
        app = xw.App(visible=False, add_book=False)
    except Exception as exc:
        print(f"failed to start hidden Excel app: {exc}")
        return

    app.display_alerts = False
    app.screen_updating = False

    try:
        source_files = sorted(input_path.iterdir(), key=lambda path: path.name.lower())
        for file_path in source_files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"processing {file_path.name}")
            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"skipped {file_path.name}: failed to open ({exc})")
                continue

            try:
                labels = parse_file_label(file_path.name)

                file_empirical_rows: List[Dict[str, Any]] = []
                file_regression_rows: List[Dict[str, Any]] = []

                empirical_sheet = get_sheet(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print("  empirical skipped: missing sheet 'Empirical Model'")
                else:
                    file_empirical_rows = process_empirical_sheet(
                        workbook=workbook,
                        sheet=empirical_sheet,
                        labels=labels,
                        source_file=file_path.name,
                    )

                regression_sheet = get_sheet(workbook, "Regression Model")
                if regression_sheet is None:
                    print("  regression skipped: missing sheet 'Regression Model'")
                else:
                    file_regression_rows = process_regression_sheet(
                        workbook=workbook,
                        sheet=regression_sheet,
                        labels=labels,
                        source_file=file_path.name,
                    )

                if not file_empirical_rows and not file_regression_rows:
                    print(f"skipped {file_path.name}: no candidate rows extracted")
                    continue

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_files += 1
                print(
                    f"processed {file_path.name}: "
                    f"empirical_rows={len(file_empirical_rows)}, "
                    f"regression_rows={len(file_regression_rows)}"
                )
            except Exception as exc:
                print(f"skipped {file_path.name}: extraction error ({exc})")
            finally:
                safe_close_source_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)
    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
