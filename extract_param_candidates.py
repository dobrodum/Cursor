#!/usr/bin/env python3
"""
Fast Excel model extractor.

Processes every eligible .xlsx workbook in input_dir once, extracts both
empirical and regression candidate rows while the workbook is open, and writes
a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import calendar
import datetime as dt
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font


# ----------------------------
# Configure paths here
# ----------------------------
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"


N_QUARTERS = 10
EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"


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


MONTH_MAP = {
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


# Default offsets from the "max" anchor column. Header matching (when found)
# overrides these for robustness across small workbook layout shifts.
EMP_DEFAULTS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -6,
    "quarterly_sales": -5,
    "reported_sales": -2,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
}

REG_DEFAULTS = {
    "forecast_value": -1,  # TOT FCST w/o SA
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"\s+", " ", re.sub(r"[^a-zA-Z0-9]+", " ", value).strip().lower())


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        pct = text.endswith("%")
        if pct:
            text = text[:-1].strip()
        try:
            val = float(text)
            return val / 100.0 if pct else val
        except ValueError:
            return None
    return None


def range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_number(max_value)
    min_num = to_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})",
        stem,
        re.IGNORECASE,
    )

    if not match:
        ticker = "UNKNOWN"
        model_period = "UNKNOWN"
        model_date = ""
        model = ticker
        return {
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
            "model": model,
        }

    ticker = match.group("ticker").upper()
    period_token = match.group("period")
    period_match = re.match(
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        period_token,
        re.IGNORECASE,
    )
    if not period_match:
        model_period = period_token
        model_date = ""
    else:
        phase = period_match.group("phase").capitalize()
        phase_key = phase.lower()
        month_text = period_match.group("month").lower()
        year = int(period_match.group("year"))
        month_num = MONTH_MAP.get(month_text, 1)
        month_abbrev = calendar.month_abbr[month_num]
        day = DAY_BY_PERIOD.get(phase_key, 15)
        model_period = f"{phase}{month_abbrev}_{year}"
        model_date = dt.date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{base}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    fallback_attempts = [
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ]
    for closer in fallback_attempts:
        try:
            closer()
            return
        except Exception:
            continue


def get_sheet_if_exists(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def get_matrix(sheet: xw.Sheet) -> Tuple[xw.Range, List[List[Any]]]:
    used = sheet.used_range
    data = used.value
    if data is None:
        return used, []
    if not isinstance(data, list):
        return used, [[data]]
    if data and not isinstance(data[0], list):
        return used, [data]
    return used, data


def matrix_value(used: xw.Range, matrix: List[List[Any]], row: int, col: int) -> Any:
    r_idx = row - used.row
    c_idx = col - used.column
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(matrix):
        return None
    row_vals = matrix[r_idx]
    if c_idx >= len(row_vals):
        return None
    return row_vals[c_idx]


def find_max_anchor(used: xw.Range, matrix: List[List[Any]]) -> Optional[Tuple[int, int]]:
    for r_idx, row_vals in enumerate(matrix):
        for c_idx, cell_val in enumerate(row_vals):
            if isinstance(cell_val, str) and cell_val.strip().lower() == "max":
                return used.row + r_idx, used.column + c_idx
    return None


def build_header_lookup(
    used: xw.Range, matrix: List[List[Any]], header_row: int
) -> Dict[str, int]:
    idx = header_row - used.row
    if idx < 0 or idx >= len(matrix):
        return {}
    headers = {}
    for c_idx, value in enumerate(matrix[idx]):
        label = normalize_label(value)
        if label:
            headers[label] = used.column + c_idx
    return headers


def first_matching_col(
    header_lookup: Dict[str, int], aliases: List[str], default_col: int
) -> int:
    normalized_aliases = [normalize_label(a) for a in aliases]
    for alias in normalized_aliases:
        for label, col in header_lookup.items():
            if alias and alias in label:
                return col
    return default_col


def set_formula_value(cell: xw.Range, formula_r1c1: str, wb: xw.Book) -> Any:
    cell.formula2 = formula_r1c1
    wb.app.calculate()
    return cell.value


def extract_empirical_rows(
    wb: xw.Book, file_meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, EMPIRICAL_SHEET)
    if sheet is None:
        print(f"Skipped {source_file}: missing '{EMPIRICAL_SHEET}'")
        return []

    used, matrix = get_matrix(sheet)
    anchor = find_max_anchor(used, matrix)
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_lookup = build_header_lookup(used, matrix, anchor_row)

    col_num_q = first_matching_col(
        header_lookup,
        ["num quarters used", "num_q", "# quarters", "quarters used"],
        anchor_col + EMP_DEFAULTS["num_quarters_used"],
    )
    col_last_q = first_matching_col(
        header_lookup,
        ["last quarter used", "last quarter"],
        anchor_col + EMP_DEFAULTS["last_quarter_used"],
    )
    col_forecast = first_matching_col(
        header_lookup,
        ["estimated total sold", "forecast value", "tot fcst", "forecast"],
        anchor_col + EMP_DEFAULTS["forecast_value"],
    )
    col_actual = first_matching_col(
        header_lookup,
        ["reported sales", "actual value", "actual"],
        anchor_col + EMP_DEFAULTS["actual_value"],
    )
    col_max = anchor_col
    col_min = first_matching_col(
        header_lookup, ["min", "forecast min"], anchor_col + EMP_DEFAULTS["forecast_min"]
    )
    col_pen = first_matching_col(
        header_lookup,
        ["penetration", "avg penetration", "penetration pct"],
        anchor_col + EMP_DEFAULTS["avg_penetration_pct"],
    )
    col_q_sales = first_matching_col(
        header_lookup, ["quarterly sales", "quarter sales"], anchor_col + EMP_DEFAULTS["quarterly_sales"]
    )
    col_reported = first_matching_col(
        header_lookup, ["reported sales", "reported"], anchor_col + EMP_DEFAULTS["reported_sales"]
    )
    col_growth = first_matching_col(
        header_lookup, ["growth rate", "growth"], anchor_col + EMP_DEFAULTS["growth_rate_pct"]
    )
    col_capture = first_matching_col(
        header_lookup,
        ["sales captured in db", "captured in db", "captured"],
        anchor_col + EMP_DEFAULTS["sales_captured_in_db_pct"],
    )

    rows: List[Dict[str, Any]] = []
    data_start_row = anchor_row + 1
    helper_col = anchor_col + 3

    for i in range(N_QUARTERS):
        current_row = data_start_row + i
        num_quarters_used = matrix_value(used, matrix, current_row, col_num_q)
        if num_quarters_used is None:
            num_quarters_used = i + 1

        # Temporary formula write (R1C1 + formula2) to trigger calculation.
        # We average penetration values from data_start_row through current_row.
        start_rel = data_start_row - current_row
        pen_rel = col_pen - helper_col
        helper_cell = sheet.range((current_row, helper_col))
        avg_penetration = set_formula_value(
            helper_cell,
            f"=AVERAGE(R[{start_rel}]C[{pen_rel}]:RC[{pen_rel}])",
            wb,
        )

        forecast_value = matrix_value(used, matrix, current_row, col_forecast)
        actual_value = matrix_value(used, matrix, current_row, col_actual)
        forecast_max = matrix_value(used, matrix, current_row, col_max)
        forecast_min = matrix_value(used, matrix, current_row, col_min)
        quarterly_sales = matrix_value(used, matrix, current_row, col_q_sales)
        reported_sales = matrix_value(used, matrix, current_row, col_reported)
        growth_rate_pct = matrix_value(used, matrix, current_row, col_growth)
        sales_captured = matrix_value(used, matrix, current_row, col_capture)
        last_quarter_used = matrix_value(used, matrix, current_row, col_last_q)

        row = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": source_file,
        }

        # Skip fully blank candidate lines.
        signal_values = [
            row["forecast_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["avg_penetration_pct"],
            row["reported_sales"],
        ]
        if all(v in (None, "") for v in signal_values):
            continue

        rows.append(row)

    return rows


def extract_regression_rows(
    wb: xw.Book, file_meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, REGRESSION_SHEET)
    if sheet is None:
        print(f"Skipped {source_file}: missing '{REGRESSION_SHEET}'")
        return []

    used, matrix = get_matrix(sheet)
    anchor = find_max_anchor(used, matrix)
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_lookup = build_header_lookup(used, matrix, anchor_row)

    # Required by spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    col_forecast = first_matching_col(
        header_lookup,
        ["tot fcst w o sa", "tot fcst wo sa", "tot fcst", "forecast total", "forecast"],
        anchor_col + REG_DEFAULTS["forecast_value"],
    )
    col_actual = first_matching_col(
        header_lookup, ["actual value", "actual", "reported sales"], anchor_col + REG_DEFAULTS["actual_value"]
    )
    col_max = anchor_col
    col_min = first_matching_col(
        header_lookup, ["min", "forecast min"], anchor_col + REG_DEFAULTS["forecast_min"]
    )

    data_start_row = anchor_row + 1
    helper_intercept_col = anchor_col + 3
    helper_slope_col = anchor_col + 4

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(N_QUARTERS):
        num_quarters_used = i + 1
        row_end = data_start_row + i

        intercept_cell = sheet.range((row_end, helper_intercept_col))
        slope_cell = sheet.range((row_end, helper_slope_col))

        start_rel_i = data_start_row - row_end
        y_rel_i = y_col - helper_intercept_col
        x_rel_i = x_col - helper_intercept_col
        intercept_formula = (
            f"=INTERCEPT(R[{start_rel_i}]C[{y_rel_i}]:RC[{y_rel_i}],"
            f"R[{start_rel_i}]C[{x_rel_i}]:RC[{x_rel_i}])"
        )
        intercept_cell.formula2 = intercept_formula

        start_rel_s = data_start_row - row_end
        y_rel_s = y_col - helper_slope_col
        x_rel_s = x_col - helper_slope_col
        slope_formula = (
            f"=SLOPE(R[{start_rel_s}]C[{y_rel_s}]:RC[{y_rel_s}],"
            f"R[{start_rel_s}]C[{x_rel_s}]:RC[{x_rel_s}])"
        )
        slope_cell.formula2 = slope_formula

        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        forecast_value = matrix_value(used, matrix, row_end, col_forecast)
        actual_value = matrix_value(used, matrix, row_end, col_actual)
        forecast_max = matrix_value(used, matrix, row_end, col_max)
        forecast_min = matrix_value(used, matrix, row_end, col_min)

        row_signature = (
            round(to_number(intercept), 12) if to_number(intercept) is not None else None,
            round(to_number(slope), 12) if to_number(slope) is not None else None,
            round(to_number(forecast_value), 12) if to_number(forecast_value) is not None else None,
            round(to_number(forecast_max), 12) if to_number(forecast_max) is not None else None,
            round(to_number(forecast_min), 12) if to_number(forecast_min) is not None else None,
        )

        # Prevent duplicate terminal row.
        if previous_signature is not None and row_signature == previous_signature:
            continue
        previous_signature = row_signature

        row = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        signal_values = [
            row["forecast_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["intercept"],
            row["slope"],
        ]
        if all(v in (None, "") for v in signal_values):
            continue

        rows.append(row)

    return rows


def write_output_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_cells in ws.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        adjusted = min(max(12, max_len + 2), 48)
        ws.column_dimensions[col_cells[0].column_letter].width = adjusted


def run() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_path}")

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processed {file_path.name}")

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_label(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    output_file = build_output_path(input_path, output_path)
    out_wb = Workbook()
    ws_emp = out_wb.active
    ws_emp.title = "empirical_candidates"
    write_output_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = out_wb.create_sheet("regression_candidates")
    write_output_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_file)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
