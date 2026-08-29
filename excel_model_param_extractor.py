#!/usr/bin/env python3
"""Extract empirical/regression candidates from all XLSX files in a folder.

This script is optimized for runtime:
- One hidden Excel app for the full run
- Each source workbook opened exactly once
- Empirical and regression models processed while workbook is open
- Source workbooks always closed without saving
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ===========================
# User-configurable locations
# ===========================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

# ===========================
# Model extraction parameters
# ===========================
N_QUARTERS = 10

# Offsets are relative to the "max" anchor cell in each model sheet.
# If your workbook layout differs, only these offsets should need adjustment.
EMPIRICAL_ROW_OFFSET = 1
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "avg_penetration_direct": -4,  # fallback if helper formula is blank
    "actual_value": -2,
    "forecast_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_ROW_OFFSET = 1
REGRESSION_OFFSETS = {
    "num_quarters_used": -12,
    "actual_value": -2,  # optional, may be blank in many workbooks
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
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


def choose_output_path(src_input_dir: Path, dest_output_dir: Path) -> Path:
    dest_output_dir.mkdir(parents=True, exist_ok=True)
    input_folder_name = src_input_dir.name or "input"
    stem = f"{input_folder_name}_PARAM"

    first_choice = dest_output_dir / f"{stem}.xlsx"
    if not first_choice.exists():
        return first_choice

    idx = 1
    while True:
        candidate = dest_output_dir / f"{stem}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].upper()

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if period_match:
        period_key = period_match.group(1).title()
        month_key = period_match.group(2).title()
        year = period_match.group(3)
        model_period = f"{period_key}{month_key}_{year}"
        model_date = date(
            int(year),
            MONTH_TO_NUMBER[month_key],
            PERIOD_DAY[period_key],
        ).isoformat()

    model = ""
    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    elif model_period:
        model = model_period

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


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


def normalize_1d(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        return values
    return [values]


def as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
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


def as_int_if_whole(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if abs(value - round(value)) < 1e-9:
        return int(round(value))
    return value


def find_anchor_max(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return None

    start_row = used.row
    start_col = used.column

    for row_idx, row_values in enumerate(matrix):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return start_row + row_idx, start_col + col_idx
    return None


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def read_numeric(sheet: xw.Sheet, row: int, col: int) -> Optional[float]:
    return as_float(read_cell(sheet, row, col))


def helper_column(sheet: xw.Sheet) -> int:
    used = sheet.used_range
    return used.column + used.columns.count + 2


def close_workbook_without_saving(wb: xw.Book) -> None:
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

    # Last-resort COM close fallback.
    try:
        wb.api.Close(False)
    except Exception:
        pass


def row_has_data(values: Sequence[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return True
    return False


def process_empirical_sheet(
    wb: xw.Book, labels: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sht.name for sht in wb.sheets]:
        print(f"  - skipped Empirical Model: sheet not found in {source_file}")
        return []

    sheet = wb.sheets["Empirical Model"]
    anchor = find_anchor_max(sheet)
    if not anchor:
        print(f"  - skipped Empirical Model: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col = anchor
    first_row = anchor_row + EMPIRICAL_ROW_OFFSET
    helper_col = helper_column(sheet)

    # Use R1C1 formula2 once per candidate row, then calculate once.
    for idx in range(N_QUARTERS):
        row = first_row + idx
        n = idx + 1
        q_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]
        r_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales"]
        q_rel = q_sales_col - helper_col
        r_rel = r_sales_col - helper_col
        sheet.range((row, helper_col)).formula2 = (
            f'=IFERROR(AVERAGE(RC[{q_rel}]:R[{n - 1}]C[{q_rel}])/'
            f'AVERAGE(RC[{r_rel}]:R[{n - 1}]C[{r_rel}]), "")'
        )

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row = first_row + idx
        num_quarters_used = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"]
        )
        if num_quarters_used is None:
            num_quarters_used = float(idx + 1)

        last_quarter_used = read_cell(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"]
        )
        quarterly_sales = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]
        )
        reported_sales = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["reported_sales"]
        )
        growth_rate_pct = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]
        )
        sales_captured_pct = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
        )

        avg_penetration_pct = as_float(read_cell(sheet, row, helper_col))
        if avg_penetration_pct is None:
            avg_penetration_pct = read_numeric(
                sheet, row, anchor_col + EMPIRICAL_OFFSETS["avg_penetration_direct"]
            )

        forecast_value = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_value"]
        )
        actual_value = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["actual_value"]
        )
        forecast_max = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_max"]
        )
        forecast_min = read_numeric(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_min"]
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if not row_has_data(
            [
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            ]
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
                "num_quarters_used": as_int_if_whole(num_quarters_used),
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,  # estimated total sold
                "actual_value": actual_value,  # reported sales
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    wb: xw.Book, labels: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sht.name for sht in wb.sheets]:
        print(f"  - skipped Regression Model: sheet not found in {source_file}")
        return []

    sheet = wb.sheets["Regression Model"]
    anchor = find_anchor_max(sheet)
    if not anchor:
        print(f"  - skipped Regression Model: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    used = sheet.used_range
    last_row = used.row + used.rows.count - 1
    first_data_row = anchor_row + REGRESSION_ROW_OFFSET

    x_vals = normalize_1d(sheet.range((first_data_row, x_col), (last_row, x_col)).value)
    y_vals = normalize_1d(sheet.range((first_data_row, y_col), (last_row, y_col)).value)

    paired_rows: List[Tuple[int, float, float]] = []
    for idx, (x_val, y_val) in enumerate(zip(x_vals, y_vals)):
        x_num = as_float(x_val)
        y_num = as_float(y_val)
        if x_num is None or y_num is None:
            continue
        paired_rows.append((first_data_row + idx, x_num, y_num))

    if len(paired_rows) < 2:
        print(f"  - skipped Regression Model: insufficient X/Y data in {source_file}")
        return []

    helper_col = helper_column(sheet)
    helper_row = first_data_row
    formula_cells: List[Tuple[int, Any, Any]] = []
    max_candidates = min(N_QUARTERS, len(paired_rows))

    for n in range(1, max_candidates + 1):
        sample_rows = paired_rows[-n:]
        start_row = sample_rows[0][0]
        end_row = sample_rows[-1][0]

        intercept_cell = sheet.range((helper_row + (n - 1), helper_col))
        slope_cell = sheet.range((helper_row + (n - 1), helper_col + 1))

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
        )
        formula_cells.append((n, intercept_cell, slope_cell))

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for n, intercept_cell, slope_cell in formula_cells:
        intercept = as_float(intercept_cell.value)
        slope = as_float(slope_cell.value)

        row = first_data_row + (n - 1)
        num_quarters_used = read_numeric(
            sheet, row, anchor_col + REGRESSION_OFFSETS["num_quarters_used"]
        )
        if num_quarters_used is None:
            num_quarters_used = float(n)

        forecast_value = read_numeric(
            sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_value"]
        )
        actual_value = read_numeric(
            sheet, row, anchor_col + REGRESSION_OFFSETS["actual_value"]
        )
        forecast_max = read_numeric(
            sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_max"]
        )
        forecast_min = read_numeric(
            sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_min"]
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if not row_has_data(
            [num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope]
        ):
            continue

        signature = (
            round(intercept, 8) if intercept is not None else None,
            round(slope, 8) if slope is not None else None,
            round(forecast_value, 8) if forecast_value is not None else None,
            round(forecast_max, 8) if forecast_max is not None else None,
            round(forecast_min, 8) if forecast_min is not None else None,
        )
        if previous_signature == signature:
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
                "parameter_value": as_int_if_whole(num_quarters_used),
                "num_quarters_used": as_int_if_whole(num_quarters_used),
                "forecast_value": forecast_value,  # TOT FCST w/o SA
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


def write_rows(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row_data in rows:
        ws.append([row_data.get(col) for col in columns])


def apply_sheet_formatting(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        values = [str(col_name)]
        for row in rows:
            value = row.get(col_name)
            if value is None:
                continue
            values.append(str(value))
        max_len = max(len(v) for v in values) if values else len(col_name)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 52)


def build_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    write_rows(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)
    apply_sheet_formatting(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = wb.create_sheet("regression_candidates")
    write_rows(ws_reg, REGRESSION_COLUMNS, regression_rows)
    apply_sheet_formatting(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_candidate_files(src_input_dir: Path) -> Iterable[Path]:
    output_name_pattern = re.compile(
        rf"^{re.escape(src_input_dir.name)}_PARAM(?:\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )
    for path in sorted(src_input_dir.iterdir()):
        if not path.is_file():
            print(f"Skipped {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue
        if output_name_pattern.match(path.name):
            print(f"Skipped {path.name}: prior output workbook")
            continue
        yield path


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    if not input_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {input_dir}")

    output_path = choose_output_path(input_dir, output_dir)
    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_candidate_files(input_dir):
            print(f"Processing {file_path.name}")
            labels = parse_file_labels(file_path)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    process_empirical_sheet(wb=wb, labels=labels, source_file=file_path.name)
                )
                regression_rows.extend(
                    process_regression_sheet(wb=wb, labels=labels, source_file=file_path.name)
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_workbook_without_saving(wb)
    finally:
        app.quit()

    build_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
