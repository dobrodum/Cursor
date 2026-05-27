#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks.

This script processes every `.xlsx` workbook in `input_dir`, extracts
candidate rows from:
  - Empirical Model
  - Regression Model

and writes a single output workbook containing:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ===== User-configurable paths =====
input_dir = Path("./input")
output_dir = Path("./output")

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

N_QUARTERS = 10
SCAN_ROW_LIMIT = 30

EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_min": 1,
    "avg_penetration_pct": -6,
    "quarterly_sales": -5,
    "reported_sales": -4,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -2,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -10,
    "forecast_value": -1,
    "forecast_min": 1,
}

HEADER_ALIASES = {
    "num_quarters_used": ["num quarters", "quarters used", "n quarters", "quarters"],
    "last_quarter_used": ["last quarter used", "last quarter"],
    "forecast_value": [
        "estimated total sold",
        "forecast value",
        "tot fcst w/o sa",
        "total forecast w o sa",
        "tot fcst without sa",
        "forecast",
    ],
    "actual_value": ["reported sales", "actual value", "actual"],
    "forecast_min": ["min", "minimum"],
    "avg_penetration_pct": ["avg penetration pct", "average penetration", "avg penetration"],
    "quarterly_sales": ["quarterly sales", "quarter sales"],
    "reported_sales": ["reported sales"],
    "growth_rate_pct": ["growth rate pct", "growth rate", "growth"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db",
    ],
}

PERIOD_RE = re.compile(r"(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})", re.I)
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
BUCKET_DAY = {"early": 5, "mid": 15, "late": 25}


def is_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return not (isinstance(value, float) and math.isnan(value))
    return False


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def normalize_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def normalize_matrix(values: Any, nrows: int, ncols: int) -> List[List[Any]]:
    if nrows <= 0 or ncols <= 0:
        return []
    if nrows == 1 and ncols == 1:
        return [[values]]
    if nrows == 1:
        if isinstance(values, list):
            return [values]
        return [[values]]
    if ncols == 1:
        if isinstance(values, list):
            return [[v] for v in values]
        return [[values]]
    if isinstance(values, list):
        return values
    return [[values]]


def normalize_vector(values: Any, expected_len: int) -> List[Any]:
    if expected_len <= 0:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flattened = [row[0] if row else None for row in values]
            if len(flattened) >= expected_len:
                return flattened[:expected_len]
            return flattened + [None] * (expected_len - len(flattened))
        if len(values) >= expected_len:
            return values[:expected_len]
        return values + [None] * (expected_len - len(values))
    return [values] + [None] * (expected_len - 1)


def to_int(value: Any) -> Optional[int]:
    if is_number(value):
        return int(value)
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.isdigit():
            return int(stripped)
    return None


def to_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    left = to_float(a)
    right = to_float(b)
    if left is None or right is None:
        return None
    return left - right


def safe_close_workbook(workbook: xw.Book) -> None:
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


def parse_month_number(month_token: str) -> int:
    month_key = month_token[:3].lower()
    if month_key not in MONTH_MAP:
        raise ValueError(f"Unsupported month token: {month_token}")
    return MONTH_MAP[month_key]


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 and parts[1] else "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_token = re.split(r"[_\s]", period_source)[0]
    period_match = PERIOD_RE.search(period_token)

    if period_match:
        bucket_raw = period_match.group("bucket").title()
        month_raw = period_match.group("month")
        month_token = month_raw[:3].title()
        year = int(period_match.group("year"))
        day = BUCKET_DAY[bucket_raw.lower()]
        month_number = parse_month_number(month_token)
        model_period = f"{bucket_raw}{month_token}_{year}"
        model_date = date(year, month_number, day).isoformat()
    else:
        model_period = "UNKNOWN_0000"
        model_date = ""

    return {
        "model": f"{ticker}_{model_period}",
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    input_folder_name = in_dir.resolve().name
    base_name = f"{input_folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def find_anchor(sheet: xw.Sheet, anchor_label: str = "max") -> Dict[str, int]:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    row_count = used.rows.count
    col_count = used.columns.count
    last_row = first_row + row_count - 1
    last_col = first_col + col_count - 1

    used_values = normalize_matrix(used.value, row_count, col_count)
    needle = anchor_label.strip().lower()

    for r_idx, row in enumerate(used_values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == needle:
                return {
                    "anchor_row": first_row + r_idx,
                    "anchor_col": first_col + c_idx,
                    "first_row": first_row,
                    "first_col": first_col,
                    "last_row": last_row,
                    "last_col": last_col,
                }
    raise ValueError(f"Could not find '{anchor_label}' anchor in sheet '{sheet.name}'")


def build_header_map(sheet: xw.Sheet, header_row: int, first_col: int, last_col: int) -> Dict[str, int]:
    ncols = max(0, last_col - first_col + 1)
    if ncols == 0:
        return {}
    row_values = normalize_matrix(
        sheet.range((header_row, first_col), (header_row, last_col)).value,
        1,
        ncols,
    )[0]
    mapping: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        if isinstance(value, str):
            normalized = normalize_header(value)
            if normalized:
                mapping[normalized] = first_col + idx
    return mapping


def resolve_column(
    header_map: Dict[str, int],
    logical_name: str,
    anchor_col: int,
    fallback_offsets: Dict[str, int],
) -> Optional[int]:
    aliases = HEADER_ALIASES.get(logical_name, [])

    for alias in aliases:
        normalized_alias = normalize_header(alias)
        if normalized_alias in header_map:
            return header_map[normalized_alias]

    for alias in aliases:
        normalized_alias = normalize_header(alias)
        for header_text, header_col in header_map.items():
            if normalized_alias and normalized_alias in header_text:
                return header_col

    offset = fallback_offsets.get(logical_name)
    if offset is None:
        return None
    col = anchor_col + offset
    if col < 1:
        return None
    return col


def read_matrix(
    sheet: xw.Sheet,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> List[List[Any]]:
    if end_row < start_row or end_col < start_col:
        return []
    nrows = end_row - start_row + 1
    ncols = end_col - start_col + 1
    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    return normalize_matrix(values, nrows, ncols)


def get_matrix_value(matrix: List[List[Any]], row_idx: int, abs_col: Optional[int], min_col: int) -> Any:
    if abs_col is None or row_idx < 0 or row_idx >= len(matrix):
        return None
    offset = abs_col - min_col
    if offset < 0 or (matrix[row_idx] and offset >= len(matrix[row_idx])):
        return None
    if not matrix[row_idx]:
        return None
    return matrix[row_idx][offset]


def try_formula2(range_obj: xw.Range, formula: Any) -> None:
    try:
        range_obj.formula2 = formula
    except Exception:
        range_obj.formula = formula


def extract_empirical_rows(
    workbook: xw.Book,
    file_path: Path,
    labels: Dict[str, str],
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"  - skipped empirical extraction: missing 'Empirical Model'")
        return []

    anchor_ctx = find_anchor(sheet, "max")
    anchor_row = anchor_ctx["anchor_row"]
    anchor_col = anchor_ctx["anchor_col"]
    first_col = anchor_ctx["first_col"]
    last_col = anchor_ctx["last_col"]

    header_map = build_header_map(sheet, anchor_row, first_col, last_col)

    num_quarters_col = resolve_column(
        header_map, "num_quarters_used", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    last_quarter_col = resolve_column(
        header_map, "last_quarter_used", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    forecast_value_col = resolve_column(
        header_map, "forecast_value", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    actual_value_col = resolve_column(
        header_map, "actual_value", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    forecast_max_col = anchor_col
    forecast_min_col = resolve_column(
        header_map, "forecast_min", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    avg_penetration_col = resolve_column(
        header_map, "avg_penetration_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    quarterly_sales_col = resolve_column(
        header_map, "quarterly_sales", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    reported_sales_col = resolve_column(
        header_map, "reported_sales", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    growth_rate_col = resolve_column(
        header_map, "growth_rate_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    captured_col = resolve_column(
        header_map, "sales_captured_in_db_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )

    tracked_cols = [
        col
        for col in [
            num_quarters_col,
            last_quarter_col,
            forecast_value_col,
            actual_value_col,
            forecast_max_col,
            forecast_min_col,
            avg_penetration_col,
            quarterly_sales_col,
            reported_sales_col,
            growth_rate_col,
            captured_col,
        ]
        if col is not None and col >= 1
    ]
    if not tracked_cols:
        return []

    row_start = anchor_row + 1
    row_end = row_start + N_QUARTERS - 1
    min_col = min(tracked_cols)
    max_col = max(tracked_cols)
    matrix = read_matrix(sheet, row_start, min_col, row_end, max_col)

    avg_pen_formula_values: List[Any] = [None] * N_QUARTERS
    if quarterly_sales_col is not None and reported_sales_col is not None:
        helper_col = max(last_col, max_col) + 2
        helper_range = sheet.range((row_start, helper_col), (row_end, helper_col))
        quarterly_offset = quarterly_sales_col - helper_col
        reported_offset = reported_sales_col - helper_col
        formula = f'=IFERROR(RC[{quarterly_offset}]/RC[{reported_offset}],"")'
        try_formula2(helper_range, formula)
        workbook.app.calculate()
        avg_pen_formula_values = normalize_vector(helper_range.value, N_QUARTERS)
        helper_range.clear_contents()

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters_used = get_matrix_value(matrix, idx, num_quarters_col, min_col)
        last_quarter_used = get_matrix_value(matrix, idx, last_quarter_col, min_col)
        forecast_value = get_matrix_value(matrix, idx, forecast_value_col, min_col)
        actual_value = get_matrix_value(matrix, idx, actual_value_col, min_col)
        forecast_max = get_matrix_value(matrix, idx, forecast_max_col, min_col)
        forecast_min = get_matrix_value(matrix, idx, forecast_min_col, min_col)
        quarterly_sales = get_matrix_value(matrix, idx, quarterly_sales_col, min_col)
        reported_sales = get_matrix_value(matrix, idx, reported_sales_col, min_col)
        growth_rate_pct = get_matrix_value(matrix, idx, growth_rate_col, min_col)
        sales_captured_pct = get_matrix_value(matrix, idx, captured_col, min_col)

        avg_penetration_raw = get_matrix_value(matrix, idx, avg_penetration_col, min_col)
        avg_penetration_pct = (
            avg_penetration_raw if not is_blank(avg_penetration_raw) else avg_pen_formula_values[idx]
        )

        major_values = [forecast_value, actual_value, forecast_max, forecast_min, avg_penetration_pct]
        if all(is_blank(value) for value in major_values):
            continue

        final_num_quarters = to_int(num_quarters_used)
        if final_num_quarters is None:
            final_num_quarters = idx + 1

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": final_num_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales if not is_blank(reported_sales) else actual_value,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": file_path.name,
        }
        rows.append(row)
    return rows


def rows_are_equal_for_dedupe(left: Dict[str, Any], right: Dict[str, Any]) -> bool:
    fields = ["num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for field in fields:
        lval = left.get(field)
        rval = right.get(field)
        if is_number(lval) and is_number(rval):
            if abs(float(lval) - float(rval)) > 1e-9:
                return False
        else:
            if ("" if lval is None else str(lval)) != ("" if rval is None else str(rval)):
                return False
    return True


def extract_regression_rows(
    workbook: xw.Book,
    file_path: Path,
    labels: Dict[str, str],
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"  - skipped regression extraction: missing 'Regression Model'")
        return []

    anchor_ctx = find_anchor(sheet, "max")
    anchor_row = anchor_ctx["anchor_row"]
    anchor_col = anchor_ctx["anchor_col"]
    first_col = anchor_ctx["first_col"]
    last_col = anchor_ctx["last_col"]
    first_row = anchor_ctx["first_row"]
    last_row = anchor_ctx["last_row"]

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"  - skipped regression extraction: invalid x/y anchor offsets")
        return []

    header_map = build_header_map(sheet, anchor_row, first_col, last_col)
    num_quarters_col = resolve_column(
        header_map, "num_quarters_used", anchor_col, REGRESSION_FALLBACK_OFFSETS
    )
    forecast_value_col = resolve_column(
        header_map, "forecast_value", anchor_col, REGRESSION_FALLBACK_OFFSETS
    )
    forecast_max_col = anchor_col
    forecast_min_col = resolve_column(
        header_map, "forecast_min", anchor_col, REGRESSION_FALLBACK_OFFSETS
    )

    tracked_cols = [
        col
        for col in [num_quarters_col, forecast_value_col, forecast_max_col, forecast_min_col]
        if col is not None and col >= 1
    ]
    if not tracked_cols:
        return []

    row_start = anchor_row + 1
    row_end_scan = row_start + SCAN_ROW_LIMIT - 1
    min_col = min(tracked_cols)
    max_col = max(tracked_cols)
    scan_matrix = read_matrix(sheet, row_start, min_col, row_end_scan, max_col)

    candidate_rows: List[Dict[str, Any]] = []
    started = False
    blank_streak = 0
    for idx, _ in enumerate(scan_matrix):
        row_idx = row_start + idx
        num_q = get_matrix_value(scan_matrix, idx, num_quarters_col, min_col)
        forecast_value = get_matrix_value(scan_matrix, idx, forecast_value_col, min_col)
        forecast_max = get_matrix_value(scan_matrix, idx, forecast_max_col, min_col)
        forecast_min = get_matrix_value(scan_matrix, idx, forecast_min_col, min_col)
        has_data = not all(is_blank(v) for v in [num_q, forecast_value, forecast_max, forecast_min])

        if has_data:
            started = True
            blank_streak = 0
            candidate_rows.append(
                {
                    "row_idx": row_idx,
                    "num_quarters_used": num_q,
                    "forecast_value": forecast_value,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                }
            )
        elif started:
            blank_streak += 1
            if blank_streak >= 3:
                break

    if not candidate_rows:
        return []

    x_values = normalize_vector(
        sheet.range((first_row, x_col), (last_row, x_col)).value,
        last_row - first_row + 1,
    )
    y_values = normalize_vector(
        sheet.range((first_row, y_col), (last_row, y_col)).value,
        last_row - first_row + 1,
    )
    xy_rows = []
    for idx, (xv, yv) in enumerate(zip(x_values, y_values)):
        if is_number(xv) and is_number(yv):
            xy_rows.append(first_row + idx)

    helper_base_col = max(last_col, max_col) + 2
    intercept_col = helper_base_col
    slope_col = helper_base_col + 1
    helper_row_start = min(row["row_idx"] for row in candidate_rows)
    helper_row_end = max(row["row_idx"] for row in candidate_rows)

    intercept_formulas: List[List[str]] = []
    slope_formulas: List[List[str]] = []
    formula_for_row: Dict[int, Dict[str, str]] = {}

    for candidate in candidate_rows:
        num_q = to_int(candidate["num_quarters_used"])
        if num_q is None:
            num_q = len(formula_for_row) + 1
        if num_q < 2 or len(xy_rows) < num_q:
            formula_for_row[candidate["row_idx"]] = {"intercept": "", "slope": ""}
            continue
        start_data_row = xy_rows[-num_q]
        end_data_row = xy_rows[-1]
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},'
            f'R{start_data_row}C{x_col}:R{end_data_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},'
            f'R{start_data_row}C{x_col}:R{end_data_row}C{x_col}),"")'
        )
        formula_for_row[candidate["row_idx"]] = {
            "intercept": intercept_formula,
            "slope": slope_formula,
        }

    for row_idx in range(helper_row_start, helper_row_end + 1):
        formulas = formula_for_row.get(row_idx, {"intercept": "", "slope": ""})
        intercept_formulas.append([formulas["intercept"]])
        slope_formulas.append([formulas["slope"]])

    intercept_range = sheet.range((helper_row_start, intercept_col), (helper_row_end, intercept_col))
    slope_range = sheet.range((helper_row_start, slope_col), (helper_row_end, slope_col))
    try_formula2(intercept_range, intercept_formulas)
    try_formula2(slope_range, slope_formulas)
    workbook.app.calculate()

    intercept_values = normalize_vector(intercept_range.value, helper_row_end - helper_row_start + 1)
    slope_values = normalize_vector(slope_range.value, helper_row_end - helper_row_start + 1)
    intercept_range.clear_contents()
    slope_range.clear_contents()

    output_rows: List[Dict[str, Any]] = []
    for candidate in candidate_rows:
        row_idx = candidate["row_idx"]
        helper_offset = row_idx - helper_row_start
        intercept = intercept_values[helper_offset] if 0 <= helper_offset < len(intercept_values) else None
        slope = slope_values[helper_offset] if 0 <= helper_offset < len(slope_values) else None
        num_quarters_used = to_int(candidate["num_quarters_used"])
        if num_quarters_used is None:
            num_quarters_used = len(output_rows) + 1

        current = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": candidate["forecast_value"],
            "actual_value": "",
            "forecast_max": candidate["forecast_max"],
            "forecast_min": candidate["forecast_min"],
            "range_width": safe_subtract(candidate["forecast_max"], candidate["forecast_min"]),
            "intercept": intercept,
            "slope": slope,
            "source_file": file_path.name,
        }

        if output_rows and rows_are_equal_for_dedupe(output_rows[-1], current):
            continue
        output_rows.append(current)

    return output_rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    if sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
    else:
        ws = workbook.create_sheet(sheet_name)

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
            value_text = "" if value is None else str(value)
            if len(value_text) > max_len:
                max_len = len(value_text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 50)


def write_output_workbook(
    out_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(out_path)


def iter_source_files(in_dir: Path) -> List[Path]:
    files: List[Path] = []
    for file_path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        files.append(file_path)
    return files


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(in_dir, out_dir)
    source_files = iter_source_files(in_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"Processing {file_path.name}")
            workbook = None
            try:
                labels = parse_file_labels(file_path)
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, file_path, labels))
                regression_rows.extend(extract_regression_rows(workbook, file_path, labels))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
