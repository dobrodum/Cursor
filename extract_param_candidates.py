#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from all .xlsx files in a folder.

Runtime priorities:
- Use one hidden Excel app for the full run.
- Open each source workbook once, process both sheets, then close without saving.
- Use anchor-based offsets and R1C1 formulas to avoid heavy coordinate conversion.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =============================
# Configure paths here
# =============================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

PERIOD_DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def as_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, (list, tuple)):
        return [list(row) for row in values]  # type: ignore[arg-type]
    return [list(values)]  # type: ignore[arg-type]


def safe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def clean_number(value: Any) -> Any:
    number = safe_float(value)
    if number is None:
        return value if value not in ("", None) else None
    rounded = round(number)
    if abs(number - rounded) < 1e-9:
        return int(rounded)
    return number


def parse_file_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"
    period_raw = parts[2] if len(parts) >= 3 else ""
    period_token = re.split(r"[_\s-]+", period_raw)[0] if period_raw else ""

    match = re.search(r"(?i)(Early|Mid|Late)([A-Za-z]{3})(\d{4})", period_token)
    if not match:
        model_period = period_token or "UnknownPeriod"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return {
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
            "model": model,
        }

    window = match.group(1).title()
    month_text = match.group(2).title()
    year = int(match.group(3))

    month_num = MONTH_MAP.get(month_text.lower())
    day = PERIOD_DAY_MAP[window.lower()]

    model_period = f"{window}{month_text}_{year}"
    model_date = date(year, month_num, day).isoformat() if month_num else ""
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def get_sheet_case_insensitive(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    target = name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def sheet_snapshot(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    top = used.row
    left = used.column
    matrix = as_matrix(used.value)
    return top, left, matrix


def find_max_anchor(top: int, left: int, matrix: Sequence[Sequence[Any]]) -> Optional[Tuple[int, int]]:
    for r_idx, row in enumerate(matrix):
        row_num = top + r_idx
        for c_idx, value in enumerate(row):
            col_num = left + c_idx
            norm = normalize_label(value)
            if norm == "max":
                return row_num, col_num
    return None


def header_band_map(
    top: int,
    left: int,
    matrix: Sequence[Sequence[Any]],
    anchor_row: int,
    band: int = 3,
) -> Dict[str, int]:
    labels: Dict[str, int] = {}
    for r_idx, row in enumerate(matrix):
        row_num = top + r_idx
        if abs(row_num - anchor_row) > band:
            continue
        for c_idx, value in enumerate(row):
            if not isinstance(value, str):
                continue
            norm = normalize_label(value)
            if norm:
                labels[norm] = left + c_idx
    return labels


def map_empirical_columns(anchor_col: int, headers: Dict[str, int]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for norm, col in headers.items():
        if norm == "max":
            out["forecast_max"] = col
        elif norm == "min":
            out["forecast_min"] = col
        elif "estimatedtotalsold" in norm or ("estimated" in norm and "sold" in norm):
            out["forecast_value"] = col
        elif "reportedsales" in norm or ("reported" in norm and "sales" in norm):
            out["actual_value"] = col
        elif "avgpenetration" in norm:
            out["avg_penetration_pct"] = col
        elif "quarterlysales" in norm:
            out["quarterly_sales"] = col
        elif "growthrate" in norm:
            out["growth_rate_pct"] = col
        elif "salescapturedindb" in norm:
            out["sales_captured_in_db_pct"] = col
        elif "numquartersused" in norm:
            out["num_quarters_used"] = col
        elif "lastquarterused" in norm:
            out["last_quarter_used"] = col
        elif "penetration" in norm and "avg" not in norm:
            out["penetration_source"] = col

    # Anchor-based defaults if headers are not found.
    out.setdefault("forecast_max", anchor_col + 0)
    out.setdefault("forecast_min", anchor_col + 1)
    out.setdefault("forecast_value", anchor_col - 2)
    out.setdefault("actual_value", anchor_col - 1)
    return out


def map_regression_columns(anchor_col: int, headers: Dict[str, int]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for norm, col in headers.items():
        if norm == "max":
            out["forecast_max"] = col
        elif norm == "min":
            out["forecast_min"] = col
        elif "totfcstwosa" in norm or ("totfcst" in norm and "wosa" in norm):
            out["forecast_value"] = col
        elif "actualvalue" in norm or "actualsales" in norm:
            out["actual_value"] = col
        elif "numquartersused" in norm:
            out["num_quarters_used"] = col

    # Anchor-based defaults if headers are not found.
    out.setdefault("forecast_max", anchor_col + 0)
    out.setdefault("forecast_min", anchor_col + 1)
    out.setdefault("forecast_value", anchor_col - 2)
    return out


def matrix_value(top: int, left: int, matrix: Sequence[Sequence[Any]], row: int, col: int) -> Any:
    r_idx = row - top
    c_idx = col - left
    if r_idx < 0 or c_idx < 0 or r_idx >= len(matrix):
        return None
    row_vals = matrix[r_idx]
    if c_idx >= len(row_vals):
        return None
    return row_vals[c_idx]


def candidate_rows_from_col(
    top: int,
    left: int,
    matrix: Sequence[Sequence[Any]],
    col: int,
    anchor_row: int,
    limit: int = N_QUARTERS,
) -> List[int]:
    c_idx = col - left
    if c_idx < 0:
        return []

    down: List[int] = []
    up: List[int] = []
    for r_idx, row_vals in enumerate(matrix):
        if c_idx >= len(row_vals):
            continue
        row_num = top + r_idx
        num = safe_float(row_vals[c_idx])
        if num is None:
            continue
        if row_num > anchor_row:
            down.append(row_num)
        elif row_num < anchor_row:
            up.append(row_num)

    if len(down) >= len(up):
        return down[:limit]
    # Near-anchor first on the upward side.
    return list(reversed(up))[:limit]


def numeric_series(
    top: int,
    left: int,
    matrix: Sequence[Sequence[Any]],
    col: int,
) -> List[Tuple[int, float]]:
    c_idx = col - left
    if c_idx < 0:
        return []

    out: List[Tuple[int, float]] = []
    for r_idx, row_vals in enumerate(matrix):
        if c_idx >= len(row_vals):
            continue
        value = safe_float(row_vals[c_idx])
        if value is not None:
            out.append((top + r_idx, value))
    return out


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def safe_close_workbook(wb: xw.Book) -> None:
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

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        pass


def extract_empirical_rows(wb: xw.Book, meta: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"skipped empirical: {source_file} (missing 'Empirical Model' sheet)")
        return []

    top, left, matrix = sheet_snapshot(sheet)
    if not matrix:
        print(f"skipped empirical: {source_file} (empty 'Empirical Model' sheet)")
        return []

    anchor = find_max_anchor(top, left, matrix)
    if anchor is None:
        print(f"skipped empirical: {source_file} (no 'max' anchor in 'Empirical Model')")
        return []
    anchor_row, anchor_col = anchor

    headers = header_band_map(top, left, matrix, anchor_row)
    cols = map_empirical_columns(anchor_col, headers)

    row_source_col = cols.get("num_quarters_used", cols["forecast_max"])
    candidate_rows = candidate_rows_from_col(top, left, matrix, row_source_col, anchor_row, N_QUARTERS)
    if not candidate_rows:
        candidate_rows = [anchor_row + i for i in range(1, N_QUARTERS + 1)]

    used_width = max((len(row) for row in matrix), default=1)
    scratch_col = left + used_width + 3
    scratch_row = anchor_row
    scratch_cell = sheet.range((scratch_row, scratch_col))

    penetration_col = cols.get("penetration_source")
    penetration_points = numeric_series(top, left, matrix, penetration_col) if penetration_col else []

    rows: List[Dict[str, Any]] = []
    for i in range(1, N_QUARTERS + 1):
        row_num = candidate_rows[i - 1] if i - 1 < len(candidate_rows) else anchor_row + i
        num_quarters_raw = matrix_value(top, left, matrix, row_num, cols.get("num_quarters_used", -1))
        num_quarters = safe_float(num_quarters_raw)
        num_quarters = int(num_quarters) if num_quarters is not None else i

        avg_penetration = safe_float(matrix_value(top, left, matrix, row_num, cols.get("avg_penetration_pct", -1)))
        if penetration_points:
            take_n = min(max(num_quarters, 1), len(penetration_points))
            start_row = penetration_points[-take_n][0]
            end_row = penetration_points[-1][0]
            formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
            set_formula2(scratch_cell, formula)
            wb.app.calculate()
            avg_penetration = safe_float(scratch_cell.value)

        forecast_value = clean_number(matrix_value(top, left, matrix, row_num, cols["forecast_value"]))
        actual_value = clean_number(matrix_value(top, left, matrix, row_num, cols["actual_value"]))
        forecast_max = clean_number(matrix_value(top, left, matrix, row_num, cols["forecast_max"]))
        forecast_min = clean_number(matrix_value(top, left, matrix, row_num, cols["forecast_min"]))

        max_num = safe_float(forecast_max)
        min_num = safe_float(forecast_min)
        range_width = max_num - min_num if max_num is not None and min_num is not None else None

        quarterly_sales = clean_number(matrix_value(top, left, matrix, row_num, cols.get("quarterly_sales", -1)))
        reported_sales = actual_value
        growth_rate = clean_number(matrix_value(top, left, matrix, row_num, cols.get("growth_rate_pct", -1)))
        sales_captured = clean_number(
            matrix_value(top, left, matrix, row_num, cols.get("sales_captured_in_db_pct", -1))
        )
        last_quarter_used = matrix_value(top, left, matrix, row_num, cols.get("last_quarter_used", -1))
        if isinstance(last_quarter_used, str):
            last_quarter_used = last_quarter_used.strip()

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": clean_number(avg_penetration),
            "num_quarters_used": num_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": clean_number(range_width),
            "avg_penetration_pct": clean_number(avg_penetration),
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": source_file,
        }

        has_signal = any(
            row.get(key) is not None
            for key in ("parameter_value", "forecast_value", "actual_value", "forecast_max", "forecast_min")
        )
        if has_signal:
            rows.append(row)

    scratch_cell.value = None
    return rows


def extract_regression_rows(wb: xw.Book, meta: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"skipped regression: {source_file} (missing 'Regression Model' sheet)")
        return []

    top, left, matrix = sheet_snapshot(sheet)
    if not matrix:
        print(f"skipped regression: {source_file} (empty 'Regression Model' sheet)")
        return []

    anchor = find_max_anchor(top, left, matrix)
    if anchor is None:
        print(f"skipped regression: {source_file} (no 'max' anchor in 'Regression Model')")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    headers = header_band_map(top, left, matrix, anchor_row)
    cols = map_regression_columns(anchor_col, headers)

    row_source_col = cols.get("num_quarters_used", cols["forecast_max"])
    candidate_rows = candidate_rows_from_col(top, left, matrix, row_source_col, anchor_row, N_QUARTERS)

    x_points = {row: val for row, val in numeric_series(top, left, matrix, x_col)}
    y_points = {row: val for row, val in numeric_series(top, left, matrix, y_col)}
    paired_rows = sorted(set(x_points) & set(y_points))

    if len(paired_rows) < 2:
        print(f"skipped regression: {source_file} (not enough paired x/y points)")
        return []

    used_width = max((len(row) for row in matrix), default=1)
    scratch_col = left + used_width + 3
    scratch_row = anchor_row
    intercept_cell = sheet.range((scratch_row, scratch_col))
    slope_cell = sheet.range((scratch_row, scratch_col + 1))

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    max_n = min(N_QUARTERS, len(paired_rows))

    for n in range(1, max_n + 1):
        active_rows = paired_rows[-n:]
        start_row = active_rows[0]
        end_row = active_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = safe_float(intercept_cell.value)
        slope = safe_float(slope_cell.value)

        row_num = candidate_rows[n - 1] if n - 1 < len(candidate_rows) else None
        num_quarters_raw = matrix_value(top, left, matrix, row_num, cols.get("num_quarters_used", -1)) if row_num else None
        num_quarters_val = safe_float(num_quarters_raw)
        num_quarters_used = int(num_quarters_val) if num_quarters_val is not None else n

        forecast_value = (
            clean_number(matrix_value(top, left, matrix, row_num, cols["forecast_value"])) if row_num else None
        )
        if forecast_value is None and intercept is not None and slope is not None:
            x_current = x_points[end_row]
            forecast_value = clean_number(intercept + slope * x_current)

        actual_value = (
            clean_number(matrix_value(top, left, matrix, row_num, cols.get("actual_value", -1))) if row_num else None
        )
        forecast_max = clean_number(matrix_value(top, left, matrix, row_num, cols["forecast_max"])) if row_num else None
        forecast_min = clean_number(matrix_value(top, left, matrix, row_num, cols["forecast_min"])) if row_num else None

        max_num = safe_float(forecast_max)
        min_num = safe_float(forecast_min)
        range_width = max_num - min_num if max_num is not None and min_num is not None else None

        signature = (
            num_quarters_used,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            clean_number(forecast_value),
            clean_number(forecast_max),
            clean_number(forecast_min),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": clean_number(range_width),
            "intercept": clean_number(intercept),
            "slope": clean_number(slope),
            "source_file": source_file,
        }

        has_signal = any(
            row.get(key) is not None
            for key in ("forecast_value", "intercept", "slope", "forecast_max", "forecast_min")
        )
        if has_signal:
            rows.append(row)

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_sheet(ws, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for col_idx in range(1, len(headers) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row_data in rows:
        ws.append([row_data.get(header) for header in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, ws.max_row)}"

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    ws_reg = wb.create_sheet("regression_candidates")

    write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    in_dir = Path(input_dir).expanduser()
    out_dir = Path(output_dir).expanduser()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = build_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                print(f"skipped file: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"processed file: {file_path.name}")
            metadata = parse_file_label(file_path)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
