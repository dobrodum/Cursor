#!/usr/bin/env python3
"""Extract empirical and regression candidates from model workbooks."""

from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running the script.
input_dir = Path("./input")
output_dir = Path("./output")

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
MAX_QUARTERS = 10

# Empirical offsets are relative to the "max" anchor cell.
# These keep extraction anchor-based and easy to tune if layouts drift.
EMPIRICAL_OFFSETS = {
    "quarter_label_col": -11,
    "quarterly_sales_col": -7,
    "reported_sales_col": -6,
    "growth_rate_col": -5,
    "capture_pct_col": -4,
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

DAY_BY_PERIOD = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_TO_NUM = {
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
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(values: Any) -> List[List[Any]]:
    """Normalize xlwings values into a 2D list."""
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and isinstance(values[0], list):
        return values
    return [values]


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and not math.isnan(value)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    return number


def safe_subtract(left: Optional[float], right: Optional[float]) -> Optional[float]:
    if left is None or right is None:
        return None
    return left - right


def parse_file_labels(file_name: str) -> Optional[FileLabels]:
    """Parse ticker/model period/date from source file name."""
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].strip()
    period_token = re.sub(r"_?send.*$", "", parts[2], flags=re.IGNORECASE).strip()
    period_token = re.sub(r"\s+", "", period_token)

    match = re.fullmatch(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, flags=re.IGNORECASE)
    if not match:
        return None

    period_prefix, month_token, year_str = match.groups()
    period_prefix = period_prefix.capitalize()
    month_num = MONTH_TO_NUM.get(month_token.lower())
    if month_num is None:
        return None

    day = DAY_BY_PERIOD[period_prefix]
    model_date = date(int(year_str), month_num, day).isoformat()
    month_abbrev = calendar.month_abbr[month_num]
    model_period = f"{period_prefix}{month_abbrev}_{year_str}"
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_path: Path, output_path: Path) -> Path:
    """Build output path with conflict-safe suffixing."""
    base_stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_stem}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_path / f"{base_stem}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def close_workbook_safely(workbook: xw.Book) -> None:
    """Close workbook without saving, with compatibility fallback."""
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
        workbook.close()
    except Exception:
        pass


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    """Write an R1C1 formula using formula2 with a robust fallback."""
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

    cell.api.FormulaR1C1 = formula_r1c1


def find_max_anchor(sheet: xw.Sheet) -> Tuple[int, int]:
    """Find the best 'max' anchor cell in a sheet."""
    used_range = sheet.used_range
    used_values = normalize_2d(used_range.value)
    if not used_values:
        raise ValueError(f"{sheet.name}: used range is empty")

    top_row = used_range.row
    left_col = used_range.column
    candidates: List[Tuple[int, int, int]] = []

    for r_idx, row_vals in enumerate(used_values):
        for c_idx, value in enumerate(row_vals):
            if isinstance(value, str) and value.strip().lower() == "max":
                row_num = top_row + r_idx
                col_num = left_col + c_idx
                score = row_num * 10

                right_val = row_vals[c_idx + 1] if c_idx + 1 < len(row_vals) else None
                left_val = row_vals[c_idx - 1] if c_idx - 1 >= 0 else None

                if isinstance(right_val, str) and right_val.strip().lower() == "min":
                    score += 7
                if isinstance(left_val, str) and any(
                    key in left_val.strip().lower() for key in ("forecast", "fcst", "total")
                ):
                    score += 3
                candidates.append((score, row_num, col_num))

    if not candidates:
        raise ValueError(f"{sheet.name}: could not locate a 'max' anchor")

    _, anchor_row, anchor_col = max(candidates, key=lambda item: item[0])
    return anchor_row, anchor_col


def pull_rows(
    sheet: xw.Sheet,
    row_start: int,
    row_end: int,
    columns: Sequence[int],
) -> List[Tuple[int, Dict[int, Any]]]:
    """Read a block once and return row-wise column dictionaries."""
    if row_end < row_start:
        return []

    min_col = min(columns)
    max_col = max(columns)
    data = normalize_2d(sheet.range((row_start, min_col), (row_end, max_col)).value)

    rows: List[Tuple[int, Dict[int, Any]]] = []
    for row_offset, values in enumerate(data):
        absolute_row = row_start + row_offset
        row_map = {col: values[col - min_col] if (col - min_col) < len(values) else None for col in columns}
        rows.append((absolute_row, row_map))
    return rows


def extract_empirical_rows(workbook: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    """Extract empirical candidates from one workbook."""
    try:
        sheet = workbook.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"Skipping {source_file}: missing sheet '{EMPIRICAL_SHEET_NAME}'")
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)

    quarter_col = max(1, anchor_col + EMPIRICAL_OFFSETS["quarter_label_col"])
    quarterly_sales_col = max(1, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"])
    reported_sales_col = max(1, anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"])
    growth_rate_col = max(1, anchor_col + EMPIRICAL_OFFSETS["growth_rate_col"])
    capture_pct_col = max(1, anchor_col + EMPIRICAL_OFFSETS["capture_pct_col"])

    scan_start = max(1, anchor_row - 300)
    scan_end = anchor_row - 1
    row_block = pull_rows(
        sheet,
        row_start=scan_start,
        row_end=scan_end,
        columns=[quarter_col, quarterly_sales_col, reported_sales_col, growth_rate_col, capture_pct_col],
    )

    numeric_rows = [
        (row_idx, row_map)
        for row_idx, row_map in row_block
        if is_number(row_map.get(quarterly_sales_col)) and is_number(row_map.get(capture_pct_col))
    ]
    if not numeric_rows:
        print(f"Skipping empirical extraction for {source_file}: no numeric rows above anchor")
        return []

    scratch_row = max(anchor_row + 2, sheet.used_range.last_cell.row + 2)
    scratch_col = max(anchor_col + 2, sheet.used_range.last_cell.column + 2)

    avg_cell = sheet.range((scratch_row, scratch_col))
    forecast_cell = sheet.range((scratch_row, scratch_col + 1))
    forecast_max_cell = sheet.range((scratch_row, scratch_col + 2))
    forecast_min_cell = sheet.range((scratch_row, scratch_col + 3))

    latest_row, latest_values = numeric_rows[-1]
    max_windows = min(MAX_QUARTERS, len(numeric_rows))
    rows: List[Dict[str, Any]] = []

    for n_quarters in range(1, max_windows + 1):
        start_row, start_values = numeric_rows[-n_quarters]
        set_formula2_r1c1(avg_cell, f"=AVERAGE(R{start_row}C{capture_pct_col}:R{latest_row}C{capture_pct_col})")
        set_formula2_r1c1(
            forecast_cell,
            f'=IFERROR(R{latest_row}C{quarterly_sales_col}/R{scratch_row}C{scratch_col},"")',
        )
        set_formula2_r1c1(
            forecast_max_cell,
            f'=IFERROR(R{latest_row}C{quarterly_sales_col}/MIN(R{start_row}C{capture_pct_col}:R{latest_row}C{capture_pct_col}),"")',
        )
        set_formula2_r1c1(
            forecast_min_cell,
            f'=IFERROR(R{latest_row}C{quarterly_sales_col}/MAX(R{start_row}C{capture_pct_col}:R{latest_row}C{capture_pct_col}),"")',
        )
        workbook.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        forecast_value = to_float(forecast_cell.value)
        forecast_max = to_float(forecast_max_cell.value)
        forecast_min = to_float(forecast_min_cell.value)

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n_quarters,
            "last_quarter_used": start_values.get(quarter_col),
            "forecast_value": forecast_value,
            "actual_value": latest_values.get(reported_sales_col),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": latest_values.get(quarterly_sales_col),
            "reported_sales": latest_values.get(reported_sales_col),
            "growth_rate_pct": latest_values.get(growth_rate_col),
            "sales_captured_in_db_pct": latest_values.get(capture_pct_col),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def extract_regression_rows(workbook: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    """Extract regression candidates from one workbook."""
    try:
        sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"Skipping {source_file}: missing sheet '{REGRESSION_SHEET_NAME}'")
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    scan_start = max(1, anchor_row - 300)
    scan_end = anchor_row - 1
    row_block = pull_rows(sheet, row_start=scan_start, row_end=scan_end, columns=[x_col, y_col])
    numeric_rows = [
        (row_idx, row_map)
        for row_idx, row_map in row_block
        if is_number(row_map.get(x_col)) and is_number(row_map.get(y_col))
    ]
    if len(numeric_rows) < 2:
        print(f"Skipping regression extraction for {source_file}: insufficient numeric rows above anchor")
        return []

    latest_row, latest_values = numeric_rows[-1]

    scratch_row = max(anchor_row + 2, sheet.used_range.last_cell.row + 2)
    scratch_col = max(anchor_col + 2, sheet.used_range.last_cell.column + 2)

    intercept_cell = sheet.range((scratch_row, scratch_col))
    slope_cell = sheet.range((scratch_row, scratch_col + 1))
    forecast_cell = sheet.range((scratch_row, scratch_col + 2))
    forecast_max_cell = sheet.range((scratch_row, scratch_col + 3))
    forecast_min_cell = sheet.range((scratch_row, scratch_col + 4))

    max_windows = min(MAX_QUARTERS, len(numeric_rows))
    rows: List[Dict[str, Any]] = []
    last_signature: Optional[Tuple[Optional[float], Optional[float], Optional[float], Optional[float], Optional[float]]] = None

    for n_quarters in range(1, max_windows + 1):
        start_row, _ = numeric_rows[-n_quarters]

        if n_quarters == 1:
            set_formula2_r1c1(intercept_cell, f"=R{latest_row}C{y_col}")
            set_formula2_r1c1(slope_cell, "=0")
            set_formula2_r1c1(forecast_cell, f"=R{latest_row}C{y_col}")
            set_formula2_r1c1(forecast_max_cell, f"=R{latest_row}C{y_col}")
            set_formula2_r1c1(forecast_min_cell, f"=R{latest_row}C{y_col}")
        else:
            set_formula2_r1c1(
                intercept_cell,
                f"=INTERCEPT(R{start_row}C{y_col}:R{latest_row}C{y_col},R{start_row}C{x_col}:R{latest_row}C{x_col})",
            )
            set_formula2_r1c1(
                slope_cell,
                f"=SLOPE(R{start_row}C{y_col}:R{latest_row}C{y_col},R{start_row}C{x_col}:R{latest_row}C{x_col})",
            )
            set_formula2_r1c1(
                forecast_cell,
                f"=R{scratch_row}C{scratch_col}+R{scratch_row}C{scratch_col + 1}*R{latest_row}C{x_col}",
            )
            set_formula2_r1c1(forecast_max_cell, f"=MAX(R{start_row}C{y_col}:R{latest_row}C{y_col})")
            set_formula2_r1c1(forecast_min_cell, f"=MIN(R{start_row}C{y_col}:R{latest_row}C{y_col})")
        workbook.app.calculate()

        intercept_value = to_float(intercept_cell.value)
        slope_value = to_float(slope_cell.value)
        forecast_value = to_float(forecast_cell.value)
        forecast_max = to_float(forecast_max_cell.value)
        forecast_min = to_float(forecast_min_cell.value)

        signature = (intercept_value, slope_value, forecast_value, forecast_max, forecast_min)
        if last_signature is not None and signature == last_signature:
            continue
        last_signature = signature

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": latest_values.get(y_col),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(
    workbook: Workbook,
    title: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    """Write one output sheet with basic formatting."""
    ws = workbook.create_sheet(title=title)
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_index, header in enumerate(headers, start=1):
        max_length = len(header)
        for row_index in range(2, ws.max_row + 1):
            value = ws.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max_length + 2, 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    """Write both candidate sheets to one workbook."""
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    wb.save(output_path)


def main() -> int:
    in_path = input_dir.expanduser().resolve()
    out_path = output_dir.expanduser().resolve()

    if not in_path.exists() or not in_path.is_dir():
        print(f"Input directory does not exist: {in_path}")
        return 1

    out_path.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(in_path, out_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipping {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipping {file_path.name}: not an .xlsx file")
                continue

            labels = parse_file_labels(file_path.name)
            if labels is None:
                print(f"Skipping {file_path.name}: filename does not match expected pattern")
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, labels, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
