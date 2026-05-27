#!/usr/bin/env python3
"""Extract empirical and regression candidates from model workbooks.

This script:
1) Opens each source workbook once with xlwings.
2) Processes both "Empirical Model" and "Regression Model" while the workbook is open.
3) Writes one consolidated output workbook with sheets:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = "./input"
output_dir = "./output"

# ---------------------------------------------------------------------------
# Extraction settings
# ---------------------------------------------------------------------------
N_QUARTERS = 10

# Anchor-based column offsets (relative to the "max" anchor cell).
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "quarterly_sales": -10,
    "reported_sales": -9,
    "growth_rate_pct": -8,
    "sales_captured_in_db_pct": -7,
    "avg_penetration_pct": -6,
    "forecast_value": -1,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": 2,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -12,
    "intercept": -5,
    "slope": -4,
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": 2,
}

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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", re.IGNORECASE)


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int_if_close(value: Any) -> Any:
    numeric = to_float(value)
    if numeric is None:
        return value
    if abs(numeric - round(numeric)) < 1e-9:
        return int(round(numeric))
    return numeric


def parse_file_labels(file_path: Path) -> Dict[str, Optional[str]]:
    """Parse ticker/model_period/model_date/model from source filename."""
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""

    period_raw = ""
    if len(parts) >= 3:
        period_raw = parts[2].split("_")[0].strip()

    model_period = ""
    model_date = ""
    match = PERIOD_PATTERN.search(period_raw)
    if match:
        phase = match.group(1).title()
        month_token = match.group(2)
        year = int(match.group(3))

        month_num = month_token_to_number(month_token)
        if month_num is not None:
            month_abbrev = calendar.month_abbr[month_num]
            day = PERIOD_DAY_MAP[phase.lower()]
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = date(year, month_num, day).isoformat()

    if not model_period and period_raw:
        model_period = period_raw

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or stem
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def month_token_to_number(token: str) -> Optional[int]:
    token = token.strip().lower()
    if not token:
        return None

    month_names = {name.lower(): idx for idx, name in enumerate(calendar.month_name) if name}
    month_abbrs = {name.lower(): idx for idx, name in enumerate(calendar.month_abbr) if name}

    if token in month_names:
        return month_names[token]
    if token in month_abbrs:
        return month_abbrs[token]

    if len(token) >= 3 and token[:3] in month_abbrs:
        return month_abbrs[token[:3]]
    return None


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int, int]]:
    """Return (anchor_row, anchor_col, last_used_col)."""
    used = sheet.used_range
    values = normalize_2d(used.value)
    first_row = used.row
    first_col = used.column
    last_used_col = used.last_cell.column

    target = anchor_text.strip().lower()
    for row_idx, row in enumerate(values):
        for col_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target:
                return first_row + row_idx, first_col + col_idx, last_used_col
    return None


def read_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def set_formula2(cell: xw.Range, formula: str) -> None:
    """Set formula using .formula2 with fallback."""
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_without_saving(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def next_output_file(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def column_map(anchor_col: int, offsets: Dict[str, int]) -> Dict[str, int]:
    return {name: anchor_col + offset for name, offset in offsets.items()}


def empirical_rows_from_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, Optional[str]],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped empirical extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col, last_used_col = anchor
    cols = column_map(anchor_col, EMPIRICAL_OFFSETS)
    temp_avg_col = max(last_used_col + 2, anchor_col + 8)
    sales_capture_col = cols["sales_captured_in_db_pct"]
    temp_rows: List[int] = []

    # Use R1C1 formula2 for average penetration on each n_quarters scenario.
    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters
        hist_start = anchor_row - n_quarters
        hist_end = anchor_row - 1
        if hist_start < 1 or sales_capture_col < 1:
            continue
        formula = (
            f'=IFERROR(AVERAGE(R{hist_start}C{sales_capture_col}:'
            f'R{hist_end}C{sales_capture_col}),"")'
        )
        set_formula2(sheet.cells(out_row, temp_avg_col), formula)
        temp_rows.append(out_row)

    if temp_rows:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        row = anchor_row + n_quarters
        num_quarters_used = read_cell_value(sheet, row, cols["num_quarters_used"])
        last_quarter_used = read_cell_value(sheet, row, cols["last_quarter_used"])
        forecast_value = to_float(read_cell_value(sheet, row, cols["forecast_value"]))
        forecast_max = to_float(read_cell_value(sheet, row, cols["forecast_max"]))
        forecast_min = to_float(read_cell_value(sheet, row, cols["forecast_min"]))
        actual_value_from_sheet = to_float(read_cell_value(sheet, row, cols["actual_value"]))
        reported_sales = to_float(read_cell_value(sheet, row, cols["reported_sales"]))
        quarterly_sales = to_float(read_cell_value(sheet, row, cols["quarterly_sales"]))
        growth_rate_pct = to_float(read_cell_value(sheet, row, cols["growth_rate_pct"]))
        sales_captured_in_db_pct = to_float(
            read_cell_value(sheet, row, cols["sales_captured_in_db_pct"])
        )
        avg_penetration_pct = to_float(read_cell_value(sheet, row, cols["avg_penetration_pct"]))
        if avg_penetration_pct is None:
            avg_penetration_pct = to_float(read_cell_value(sheet, row, temp_avg_col))

        actual_value = reported_sales if reported_sales is not None else actual_value_from_sheet

        if num_quarters_used is None:
            num_quarters_used = n_quarters
        num_quarters_used = to_int_if_close(num_quarters_used)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        # Skip empty rows.
        if all(
            value is None
            for value in [
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                reported_sales,
                quarterly_sales,
            ]
        ):
            continue

        row_data = {
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
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row_data)

    # Remove temporary formulas written for calculation.
    for row in temp_rows:
        sheet.cells(row, temp_avg_col).value = None

    return rows


def regression_rows_from_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, Optional[str]],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped regression extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col, last_used_col = anchor
    cols = column_map(anchor_col, REGRESSION_OFFSETS)

    # Existing regression logic uses y_col and x_col from max anchor.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    temp_intercept_col = max(last_used_col + 2, anchor_col + 8)
    temp_slope_col = temp_intercept_col + 1
    temp_rows: List[int] = []

    # Use formula2 with R1C1 to calculate INTERCEPT and SLOPE.
    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters
        hist_start = anchor_row - n_quarters
        hist_end = anchor_row - 1
        if hist_start < 1 or x_col < 1 or y_col < 1:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{hist_start}C{y_col}:R{hist_end}C{y_col},'
            f'R{hist_start}C{x_col}:R{hist_end}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{hist_start}C{y_col}:R{hist_end}C{y_col},'
            f'R{hist_start}C{x_col}:R{hist_end}C{x_col}),"")'
        )
        set_formula2(sheet.cells(out_row, temp_intercept_col), intercept_formula)
        set_formula2(sheet.cells(out_row, temp_slope_col), slope_formula)
        temp_rows.append(out_row)

    if temp_rows:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        row = anchor_row + n_quarters
        num_quarters_used = read_cell_value(sheet, row, cols["num_quarters_used"])
        if num_quarters_used is None:
            num_quarters_used = n_quarters
        num_quarters_used = to_int_if_close(num_quarters_used)

        intercept = to_float(read_cell_value(sheet, row, cols["intercept"]))
        slope = to_float(read_cell_value(sheet, row, cols["slope"]))
        if intercept is None:
            intercept = to_float(read_cell_value(sheet, row, temp_intercept_col))
        if slope is None:
            slope = to_float(read_cell_value(sheet, row, temp_slope_col))

        forecast_value = to_float(read_cell_value(sheet, row, cols["forecast_value"]))
        forecast_max = to_float(read_cell_value(sheet, row, cols["forecast_max"]))
        forecast_min = to_float(read_cell_value(sheet, row, cols["forecast_min"]))
        actual_value = to_float(read_cell_value(sheet, row, cols["actual_value"]))

        if all(
            value is None
            for value in [forecast_value, forecast_max, forecast_min, intercept, slope]
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row_data = {
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
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row_data)

    # Prevent duplicated final row by comparing to the previous row.
    if len(rows) >= 2:
        prev = rows[-2]
        curr = rows[-1]
        comparable_fields = ["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
        if all(to_float(curr[field]) == to_float(prev[field]) for field in comparable_fields):
            rows.pop()

    for row in temp_rows:
        sheet.cells(row, temp_intercept_col).value = None
        sheet.cells(row, temp_slope_col).value = None

    return rows


def write_sheet(sheet, headers: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    sheet.append(list(headers))
    for row in rows:
        sheet.append([row.get(header) for header in headers])

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        print(f"Input directory does not exist: {input_path}")
        return

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_file(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)

                sheet_names = {sheet.name for sheet in workbook.sheets}
                if "Empirical Model" in sheet_names:
                    empirical_rows.extend(
                        empirical_rows_from_sheet(workbook, workbook.sheets["Empirical Model"], labels, file_path.name)
                    )
                else:
                    print(f"Skipped empirical extraction in {file_path.name}: missing sheet 'Empirical Model'")

                if "Regression Model" in sheet_names:
                    regression_rows.extend(
                        regression_rows_from_sheet(
                            workbook, workbook.sheets["Regression Model"], labels, file_path.name
                        )
                    )
                else:
                    print(f"Skipped regression extraction in {file_path.name}: missing sheet 'Regression Model'")

                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_without_saving(workbook)
    finally:
        app.quit()

    out_wb = Workbook()
    empirical_sheet = out_wb.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = out_wb.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_file)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
