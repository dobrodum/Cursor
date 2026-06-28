#!/usr/bin/env python3
"""
Extract empirical and regression candidates from .xlsx workbooks.

This script:
- Opens each source workbook once.
- Processes both "Empirical Model" and "Regression Model" while open.
- Closes source workbooks without saving.
- Writes one consolidated output workbook with two sheets:
  empirical_candidates and regression_candidates.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import re

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Inputs / outputs
input_dir = Path("input")
output_dir = Path("output")

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

MONTH_TO_NUMBER = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    if text.startswith("#"):  # Excel error values
        return None
    is_pct = text.endswith("%")
    if is_pct:
        text = text[:-1]
    text = text.replace(",", "").replace("$", "")
    try:
        number = float(text)
    except ValueError:
        return None
    if is_pct:
        return number / 100.0
    return number


def rounded_or_none(value: Optional[float], places: int = 8) -> Optional[float]:
    if value is None:
        return None
    return round(value, places)


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def safe_subtract(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return a - b


def close_workbook_safe(workbook: Any) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    fallbacks = (
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.close(),
    )
    for closer in fallbacks:
        try:
            closer()
            return
        except Exception:
            continue


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem
    pattern = re.compile(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<part>Early|Mid|Late)\s*(?P<month>[A-Za-z]{3})\s*(?P<year>\d{4})",
        flags=re.IGNORECASE,
    )
    match = pattern.search(stem)

    ticker = "UNKNOWN"
    model_period = "UNKNOWN_PERIOD"
    model_date = ""

    if match:
        ticker = match.group("ticker").upper()
        part = match.group("part").title()
        month = match.group("month").title()
        year = match.group("year")
        month_num = MONTH_TO_NUMBER.get(month)
        day = PERIOD_DAY.get(part)
        model_period = f"{part}{month}_{year}"
        if month_num is not None and day is not None:
            model_date = f"{year}-{month_num:02d}-{day:02d}"
    else:
        split_parts = [p.strip() for p in stem.split("-") if p.strip()]
        if len(split_parts) >= 2:
            ticker = re.sub(r"[^A-Za-z0-9]", "", split_parts[1]).upper() or ticker
        if len(split_parts) >= 3:
            cleaned_period = re.sub(r"[^A-Za-z0-9_]", "", split_parts[2])
            if cleaned_period:
                model_period = cleaned_period

    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_output_path(input_folder: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_stem = f"{input_folder.name}_PARAM"
    output_path = out_dir / f"{base_stem}.xlsx"
    if not output_path.exists():
        return output_path

    index = 1
    while True:
        candidate = out_dir / f"{base_stem}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def is_max_label(value: Any) -> bool:
    return isinstance(value, str) and value.strip().lower() == "max"


def is_min_label(value: Any) -> bool:
    return isinstance(value, str) and value.strip().lower() == "min"


def matrix_get(matrix: Sequence[Sequence[Any]], row_idx: int, col_idx: int) -> Any:
    if row_idx < 0 or col_idx < 0 or row_idx >= len(matrix):
        return None
    row = matrix[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


def find_max_anchor(sheet: Any) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return None

    base_row = used.row
    base_col = used.column
    best: Optional[Tuple[int, int, int, int]] = None  # score, row, col, tie_break

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if not is_max_label(value):
                continue

            right = matrix_get(values, r_idx, c_idx + 1)
            below = matrix_get(values, r_idx + 1, c_idx)
            below_right = matrix_get(values, r_idx + 1, c_idx + 1)

            score = 0
            if to_float(right) is not None:
                score += 2
            if is_min_label(below):
                score += 2
            if to_float(below_right) is not None:
                score += 1

            abs_row = base_row + r_idx
            abs_col = base_col + c_idx
            candidate = (score, abs_row, abs_col, abs_row * 10000 + abs_col)
            if best is None or candidate > best:
                best = candidate

    if best is None:
        return None
    return best[1], best[2]


def get_column_values(sheet: Any, col_idx: int, last_row: int) -> List[Any]:
    if col_idx < 1 or last_row < 1:
        return []
    values = sheet.range((1, col_idx), (last_row, col_idx)).value
    if isinstance(values, list):
        return values
    return [values]


def candidate_data_rows(
    x_values: Sequence[Any], y_values: Sequence[Any], row_cap: int, max_rows: int = N_QUARTERS
) -> List[int]:
    rows: List[int] = []
    for row_idx, (x_val, y_val) in enumerate(zip(x_values, y_values), start=1):
        if row_idx > row_cap:
            break
        if to_float(x_val) is None or to_float(y_val) is None:
            continue
        rows.append(row_idx)
    return rows[-max_rows:]


def get_sheet_if_exists(workbook: Any, sheet_name: str) -> Optional[Any]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def get_row_label(sheet: Any, row_idx: int, candidate_cols: Sequence[int]) -> str:
    for col_idx in candidate_cols:
        if col_idx < 1:
            continue
        try:
            value = sheet.cells(row_idx, col_idx).value
        except Exception:
            continue
        if value is None:
            continue
        text = str(value).strip()
        if text and text.lower() != "none":
            return text
    return ""


def process_empirical_sheet(workbook: Any, meta: FileMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, "Empirical Model")
    if sheet is None:
        print(f"  empirical: skipped ({source_file}) - sheet missing")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  empirical: skipped ({source_file}) - max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    quarterly_col = anchor_col - 11
    reported_col = anchor_col - 7
    if quarterly_col < 1 or reported_col < 1:
        print(f"  empirical: skipped ({source_file}) - invalid anchor offsets")
        return []

    last_row = max(sheet.used_range.last_cell.row, anchor_row)
    quarterly_values = get_column_values(sheet, quarterly_col, last_row)
    reported_values = get_column_values(sheet, reported_col, last_row)
    rows = candidate_data_rows(quarterly_values, reported_values, row_cap=anchor_row, max_rows=N_QUARTERS)

    if not rows:
        print(f"  empirical: skipped ({source_file}) - no data rows found")
        return []

    latest_row = rows[-1]
    latest_quarterly = to_float(quarterly_values[latest_row - 1])
    latest_reported = to_float(reported_values[latest_row - 1])
    previous_reported = to_float(reported_values[rows[-2] - 1]) if len(rows) >= 2 else None

    scratch_col = sheet.used_range.last_cell.column + 4
    scratch_row_start = anchor_row + 2
    calc_rows: List[Tuple[int, int, int]] = []  # formula_row, n_quarters, start_row

    for idx, n_quarters in enumerate(range(1, len(rows) + 1)):
        start_row = rows[-n_quarters]
        formula_row = scratch_row_start + idx
        avg_cell = sheet.cells(formula_row, scratch_col)
        forecast_cell = sheet.cells(formula_row, scratch_col + 1)

        avg_cell.formula2 = (
            f"=SUM(R{start_row}C{quarterly_col}:R{latest_row}C{quarterly_col})/"
            f"SUM(R{start_row}C{reported_col}:R{latest_row}C{reported_col})"
        )
        forecast_cell.formula2 = f"=R{latest_row}C{quarterly_col}/R{formula_row}C{scratch_col}"
        calc_rows.append((formula_row, n_quarters, start_row))

    workbook.app.calculate()

    output_rows: List[Dict[str, Any]] = []
    for formula_row, n_quarters, start_row in calc_rows:
        avg_penetration = to_float(sheet.cells(formula_row, scratch_col).value)
        forecast_value = to_float(sheet.cells(formula_row, scratch_col + 1).value)

        subset_rows = rows[-n_quarters:]
        penetration_values: List[float] = []
        for row_idx in subset_rows:
            q_val = to_float(quarterly_values[row_idx - 1])
            r_val = to_float(reported_values[row_idx - 1])
            ratio = safe_div(q_val, r_val)
            if ratio is not None:
                penetration_values.append(ratio)

        positive_penetrations = [p for p in penetration_values if p > 0]
        min_pen = min(positive_penetrations) if positive_penetrations else None
        max_pen = max(positive_penetrations) if positive_penetrations else None
        forecast_max = safe_div(latest_quarterly, min_pen)
        forecast_min = safe_div(latest_quarterly, max_pen)
        range_width = safe_subtract(forecast_max, forecast_min)

        growth_rate_pct = None
        if latest_reported is not None and previous_reported not in (None, 0):
            growth_rate_pct = (latest_reported - previous_reported) / previous_reported

        sales_captured_pct = safe_div(latest_quarterly, latest_reported)
        last_quarter_used = get_row_label(
            sheet,
            latest_row,
            candidate_cols=[quarterly_col - 1, quarterly_col - 2, reported_col - 1],
        )

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
                "actual_value": latest_reported,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": latest_quarterly,
                "reported_sales": latest_reported,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    if calc_rows:
        end_formula_row = scratch_row_start + len(calc_rows) - 1
        sheet.range((scratch_row_start, scratch_col), (end_formula_row, scratch_col + 1)).value = None

    return output_rows


def process_regression_sheet(workbook: Any, meta: FileMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, "Regression Model")
    if sheet is None:
        print(f"  regression: skipped ({source_file}) - sheet missing")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  regression: skipped ({source_file}) - max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print(f"  regression: skipped ({source_file}) - invalid anchor offsets")
        return []

    last_row = max(sheet.used_range.last_cell.row, anchor_row)
    x_values = get_column_values(sheet, x_col, last_row)
    y_values = get_column_values(sheet, y_col, last_row)
    rows = candidate_data_rows(x_values, y_values, row_cap=anchor_row, max_rows=N_QUARTERS)

    if len(rows) < 2:
        print(f"  regression: skipped ({source_file}) - insufficient data rows")
        return []

    latest_row = rows[-1]
    latest_x = to_float(x_values[latest_row - 1])
    latest_actual = to_float(y_values[latest_row - 1])

    scratch_col = sheet.used_range.last_cell.column + 6
    scratch_row_start = anchor_row + 2
    calc_rows: List[Tuple[int, int]] = []  # formula_row, n_quarters

    for idx, n_quarters in enumerate(range(2, len(rows) + 1)):
        start_row = rows[-n_quarters]
        formula_row = scratch_row_start + idx
        intercept_col = scratch_col
        slope_col = scratch_col + 1
        forecast_col = scratch_col + 2
        max_col = scratch_col + 3
        min_col = scratch_col + 4

        sheet.cells(formula_row, intercept_col).formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )
        sheet.cells(formula_row, slope_col).formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )
        sheet.cells(formula_row, forecast_col).formula2 = (
            f"=R{latest_row}C{x_col}*R{formula_row}C{slope_col}+R{formula_row}C{intercept_col}"
        )
        sheet.cells(formula_row, max_col).formula2 = (
            f"=MAX(R{start_row}C{x_col}:R{latest_row}C{x_col})*R{formula_row}C{slope_col}"
            f"+R{formula_row}C{intercept_col}"
        )
        sheet.cells(formula_row, min_col).formula2 = (
            f"=MIN(R{start_row}C{x_col}:R{latest_row}C{x_col})*R{formula_row}C{slope_col}"
            f"+R{formula_row}C{intercept_col}"
        )
        calc_rows.append((formula_row, n_quarters))

    workbook.app.calculate()

    output_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None
    for formula_row, n_quarters in calc_rows:
        intercept = to_float(sheet.cells(formula_row, scratch_col).value)
        slope = to_float(sheet.cells(formula_row, scratch_col + 1).value)
        forecast_value = to_float(sheet.cells(formula_row, scratch_col + 2).value)
        forecast_max = to_float(sheet.cells(formula_row, scratch_col + 3).value)
        forecast_min = to_float(sheet.cells(formula_row, scratch_col + 4).value)
        range_width = safe_subtract(forecast_max, forecast_min)

        signature = (
            rounded_or_none(intercept),
            rounded_or_none(slope),
            rounded_or_none(forecast_value),
            rounded_or_none(forecast_max),
            rounded_or_none(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

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
                "actual_value": latest_actual if latest_actual is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if calc_rows:
        end_formula_row = scratch_row_start + len(calc_rows) - 1
        sheet.range((scratch_row_start, scratch_col), (end_formula_row, scratch_col + 4)).value = None

    # Prevent duplicate final row by comparing to previous row.
    if len(output_rows) >= 2:
        last = output_rows[-1]
        prev = output_rows[-2]
        sig_last = (
            rounded_or_none(to_float(last["intercept"])),
            rounded_or_none(to_float(last["slope"])),
            rounded_or_none(to_float(last["forecast_value"])),
            rounded_or_none(to_float(last["forecast_max"])),
            rounded_or_none(to_float(last["forecast_min"])),
        )
        sig_prev = (
            rounded_or_none(to_float(prev["intercept"])),
            rounded_or_none(to_float(prev["slope"])),
            rounded_or_none(to_float(prev["forecast_value"])),
            rounded_or_none(to_float(prev["forecast_max"])),
            rounded_or_none(to_float(prev["forecast_min"])),
        )
        if sig_last == sig_prev:
            output_rows.pop()

    # Read once so linters cannot flag it as unused in future edits.
    _ = latest_x
    return output_rows


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

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
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def iter_source_files(in_dir: Path) -> Sequence[Path]:
    if not in_dir.exists():
        return []
    return sorted(path for path in in_dir.iterdir() if path.is_file())


def main() -> None:
    files_processed = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    source_files = iter_source_files(input_dir)
    if not source_files:
        print(f"No source files found in: {input_dir.resolve()}")

    app: Optional[Any] = None
    original_calculation: Optional[str] = None

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            original_calculation = app.calculation
            app.calculation = "manual"
        except Exception:
            original_calculation = None

        for file_path in source_files:
            if file_path.name.startswith("~"):
                print(f"SKIP {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIP {file_path.name}: not .xlsx")
                continue

            print(f"PROCESS {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path)
                empirical_rows.extend(process_empirical_sheet(workbook, meta, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, meta, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"SKIP {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_safe(workbook)
    finally:
        if app is not None:
            if original_calculation is not None:
                try:
                    app.calculation = original_calculation
                except Exception:
                    pass
            app.quit()

    output_path = get_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path.resolve()}")
    print(f"files processed: {files_processed}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
