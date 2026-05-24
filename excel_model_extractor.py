#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from .xlsx workbooks.

Workflow:
1. Open each source workbook once with xlwings.
2. While workbook is open, process both:
   - Empirical Model
   - Regression Model
3. Close source workbook without saving changes.
4. Write one output workbook containing:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# User-configurable directories.
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")

N_QUARTERS_MAX = 10
DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}

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


@dataclass
class SheetCache:
    values: List[List[Any]]
    top: int
    left: int
    last_row: int
    last_col: int

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.top
        c_idx = col - self.left
        if r_idx < 0 or c_idx < 0:
            return None
        if r_idx >= len(self.values):
            return None
        row_values = self.values[r_idx]
        if c_idx >= len(row_values):
            return None
        return row_values[c_idx]


def normalize_2d(values: Any, n_rows: int, n_cols: int) -> List[List[Any]]:
    if n_rows == 1 and n_cols == 1:
        return [[values]]
    if n_rows == 1:
        if isinstance(values, list):
            return [values]
        return [[values]]
    if n_cols == 1:
        if isinstance(values, list):
            return [[item] for item in values]
        return [[values]]
    if isinstance(values, list):
        return values
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def key_text(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1].strip()
            try:
                return float(cleaned) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    rounded = round(number)
    if abs(number - rounded) > 1e-9:
        return None
    return int(rounded)


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def flatten_excel_vector(value: Any) -> List[Any]:
    if isinstance(value, list):
        if value and isinstance(value[0], list):
            return [row[0] if row else None for row in value]
        return value
    return [value]


def canonical_float(value: Any) -> Optional[float]:
    num = to_float(value)
    if num is None:
        return None
    return round(num, 10)


def month_to_number(token: str) -> Optional[int]:
    token = token.strip()
    for candidate in (token, token[:3]):
        for fmt in ("%b", "%B"):
            try:
                return datetime.strptime(candidate, fmt).month
            except ValueError:
                pass
    return None


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) >= 2 else ""
    period_raw = parts[2] if len(parts) >= 3 else ""
    period_raw = re.sub(r"[_-]?send.*$", "", period_raw, flags=re.IGNORECASE).strip(" _-")

    model_period = ""
    model_date = ""

    match = re.search(
        r"(early|mid|late)\s*([a-zA-Z]{3,9})\s*[_-]?(\d{4})",
        period_raw,
        flags=re.IGNORECASE,
    )
    if not match:
        match = re.search(
            r"(early|mid|late)([a-zA-Z]{3,9})(\d{4})",
            period_raw,
            flags=re.IGNORECASE,
        )

    if match:
        period_key = match.group(1).lower()
        month_token = match.group(2)
        year = int(match.group(3))
        month_num = month_to_number(month_token)
        if month_num is not None:
            month_label = datetime(year, month_num, 1).strftime("%b")
            model_period = f"{period_key.capitalize()}{month_label}_{year}"
            model_date = f"{year:04d}-{month_num:02d}-{DAY_BY_PERIOD[period_key]:02d}"

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def make_unique_output_path(input_path: Path, target_output_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = target_output_dir / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def safe_close_source_workbook(wb: xw.Book) -> None:
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
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        pass


def set_formula2(target_range: xw.Range, formula_r1c1: str) -> None:
    try:
        target_range.formula2 = formula_r1c1
    except Exception:
        target_range.formula = formula_r1c1


def read_sheet_cache(ws: xw.Sheet) -> SheetCache:
    used = ws.used_range
    top = used.row
    left = used.column
    n_rows = used.rows.count
    n_cols = used.columns.count
    values = normalize_2d(used.value, n_rows, n_cols)
    return SheetCache(
        values=values,
        top=top,
        left=left,
        last_row=top + n_rows - 1,
        last_col=left + n_cols - 1,
    )


def get_sheet_by_name(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    wanted = normalize_text(target_name)
    for ws in wb.sheets:
        if normalize_text(ws.name) == wanted:
            return ws
    return None


def find_anchor_max(cache: SheetCache) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int]] = []
    for r_idx, row_values in enumerate(cache.values):
        abs_row = cache.top + r_idx
        for c_idx, value in enumerate(row_values):
            if key_text(value) == "max":
                abs_col = cache.left + c_idx
                candidates.append((abs_row, abs_col))
    if not candidates:
        return None

    def score(candidate: Tuple[int, int]) -> int:
        row, col = candidate
        s = 0
        for delta in range(1, 5):
            if key_text(cache.get(row, col + delta)) == "min":
                s += 5
                break
        numeric_below = 0
        for delta_row in range(1, 15):
            if to_float(cache.get(row + delta_row, col)) is not None:
                numeric_below += 1
        s += numeric_below
        return s

    candidates.sort(key=score, reverse=True)
    return candidates[0]


def find_column_near_anchor(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    keyword_sets: Sequence[Sequence[str]],
    row_window: int = 4,
    col_window: int = 50,
) -> Optional[int]:
    best: Optional[Tuple[int, int]] = None
    row_start = max(cache.top, anchor_row - row_window)
    row_end = min(cache.last_row, anchor_row + row_window)
    col_start = max(cache.left, anchor_col - col_window)
    col_end = min(cache.last_col, anchor_col + col_window)

    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            text = key_text(cache.get(row, col))
            if not text:
                continue
            for tokens in keyword_sets:
                if all(token in text for token in tokens):
                    distance = abs(anchor_row - row) + abs(anchor_col - col)
                    if best is None or distance < best[0]:
                        best = (distance, col)
                    break
    return best[1] if best else None


def find_column_anywhere(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    keyword_sets: Sequence[Sequence[str]],
) -> Optional[int]:
    best: Optional[Tuple[int, int]] = None
    for r_idx, row_values in enumerate(cache.values):
        abs_row = cache.top + r_idx
        for c_idx, value in enumerate(row_values):
            text = key_text(value)
            if not text:
                continue
            for tokens in keyword_sets:
                if all(token in text for token in tokens):
                    abs_col = cache.left + c_idx
                    distance = abs(anchor_row - abs_row) + abs(anchor_col - abs_col)
                    if best is None or distance < best[0]:
                        best = (distance, abs_col)
                    break
    return best[1] if best else None


def infer_candidate_rows(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    max_candidates: int = N_QUARTERS_MAX,
) -> Tuple[List[int], List[int]]:
    best_rows: List[int] = []
    best_quarters: List[int] = []

    col_start = max(cache.left, anchor_col - 20)
    col_end = min(cache.last_col, anchor_col + 6)

    for col in range(col_start, col_end + 1):
        rows: List[int] = []
        quarters: List[int] = []
        for row in range(anchor_row + 1, min(cache.last_row, anchor_row + 80) + 1):
            n = to_int(cache.get(row, col))
            if n is None or n <= 0:
                if len(rows) >= max_candidates:
                    break
                continue
            rows.append(row)
            quarters.append(n)
            if len(rows) >= max_candidates:
                break
        if len(rows) > len(best_rows):
            best_rows = rows[:max_candidates]
            best_quarters = quarters[:max_candidates]

    if len(best_rows) >= 3:
        return best_rows, best_quarters

    rows = [anchor_row + 1 + i for i in range(max_candidates)]
    quarters = [i + 1 for i in range(max_candidates)]
    return rows, quarters


def collect_numeric_rows(cache: SheetCache, col: int, row_end: int) -> List[int]:
    rows: List[int] = []
    for row in range(cache.top, min(cache.last_row, row_end) + 1):
        if to_float(cache.get(row, col)) is not None:
            rows.append(row)
    return rows


def collect_paired_numeric_rows(cache: SheetCache, x_col: int, y_col: int, row_end: int) -> List[int]:
    rows: List[int] = []
    for row in range(cache.top, min(cache.last_row, row_end) + 1):
        x_val = to_float(cache.get(row, x_col))
        y_val = to_float(cache.get(row, y_col))
        if x_val is not None and y_val is not None:
            rows.append(row)
    return rows


def extract_empirical_rows(wb: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    ws = get_sheet_by_name(wb, "Empirical Model")
    if ws is None:
        print(f"Skipped empirical for {source_file}: sheet missing")
        return []

    cache = read_sheet_cache(ws)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"Skipped empirical for {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor

    min_col = find_column_near_anchor(cache, anchor_row, anchor_col, [("min",)], row_window=2, col_window=6)
    if min_col is None:
        min_col = anchor_col + 1

    num_quarters_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("num", "quarter"), ("quarters", "used"), ("n", "quarter")],
        row_window=4,
        col_window=25,
    )
    last_quarter_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("last", "quarter")],
        row_window=5,
        col_window=35,
    )
    forecast_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [
            ("estimated", "total", "sold"),
            ("est", "total", "sold"),
            ("forecast", "value"),
            ("tot", "fcst"),
        ],
        row_window=5,
        col_window=40,
    )
    if forecast_col is None:
        forecast_col = anchor_col - 1

    actual_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("reported", "sales"), ("actual", "sales"), ("actual", "value")],
        row_window=6,
        col_window=45,
    )
    avg_pen_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("avg", "penetration"), ("average", "penetration"), ("penetration", "pct")],
        row_window=6,
        col_window=45,
    )
    quarterly_sales_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("quarterly", "sales"), ("qtr", "sales"), ("quarter", "sales")],
        row_window=8,
        col_window=55,
    )
    reported_sales_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("reported", "sales")],
        row_window=8,
        col_window=55,
    )
    growth_rate_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("growth", "rate"), ("growth", "pct"), ("growth",)],
        row_window=8,
        col_window=55,
    )
    captured_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("captured", "db"), ("sales", "captured"), ("captured", "pct"), ("captured",)],
        row_window=8,
        col_window=55,
    )

    candidate_rows, candidate_quarters = infer_candidate_rows(cache, anchor_row, anchor_col, N_QUARTERS_MAX)
    if num_quarters_col is not None:
        updated_quarters: List[int] = []
        for idx, row in enumerate(candidate_rows):
            value = to_int(cache.get(row, num_quarters_col))
            updated_quarters.append(value if value is not None else candidate_quarters[idx])
        candidate_quarters = updated_quarters

    # R1C1 helper formula for avg penetration.
    penetration_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("penetration",), ("pen", "pct")],
        row_window=15,
        col_window=70,
    )
    if penetration_col is None:
        penetration_col = find_column_anywhere(
            cache,
            anchor_row,
            anchor_col,
            [("penetration",), ("pen", "pct")],
        )

    avg_by_index: Dict[int, Any] = {}
    if penetration_col is not None:
        history_rows = collect_numeric_rows(cache, penetration_col, row_end=anchor_row - 1)
        if history_rows:
            helper_start_row = cache.last_row + 2
            helper_col = cache.last_col + 2
            formula_indices: List[int] = []
            for idx, n in enumerate(candidate_quarters):
                if n is None or n <= 0 or n > len(history_rows):
                    continue
                start_row = history_rows[-n]
                end_row = history_rows[-1]
                formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
                set_formula2(ws.range((helper_start_row + idx, helper_col)), formula)
                formula_indices.append(idx)

            if formula_indices:
                wb.app.calculate()
                helper_values = ws.range(
                    (helper_start_row, helper_col),
                    (helper_start_row + len(candidate_rows) - 1, helper_col),
                ).value
                helper_values = flatten_excel_vector(helper_values)
                for idx in formula_indices:
                    avg_by_index[idx] = helper_values[idx]

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(candidate_rows):
        n_quarters = candidate_quarters[idx] if idx < len(candidate_quarters) else (idx + 1)
        if n_quarters is None:
            n_quarters = idx + 1

        forecast_max = cache.get(row, anchor_col)
        forecast_min = cache.get(row, min_col)
        forecast_value = cache.get(row, forecast_col)
        actual_value = cache.get(row, actual_col) if actual_col is not None else None
        avg_pen = avg_by_index.get(idx)
        if avg_pen is None and avg_pen_col is not None:
            avg_pen = cache.get(row, avg_pen_col)

        last_quarter = cache.get(row, last_quarter_col) if last_quarter_col is not None else None
        quarterly_sales = cache.get(row, quarterly_sales_col) if quarterly_sales_col is not None else None
        reported_sales = cache.get(row, reported_sales_col) if reported_sales_col is not None else actual_value
        growth_rate = cache.get(row, growth_rate_col) if growth_rate_col is not None else None
        captured_pct = cache.get(row, captured_col) if captured_col is not None else None
        range_width = calc_range_width(forecast_max, forecast_min)

        if all(
            value in (None, "")
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_pen,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(wb: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    ws = get_sheet_by_name(wb, "Regression Model")
    if ws is None:
        print(f"Skipped regression for {source_file}: sheet missing")
        return []

    cache = read_sheet_cache(ws)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"Skipped regression for {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    min_col = find_column_near_anchor(cache, anchor_row, anchor_col, [("min",)], row_window=2, col_window=6)
    if min_col is None:
        min_col = anchor_col + 1

    num_quarters_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("num", "quarter"), ("quarters", "used"), ("n", "quarter")],
        row_window=4,
        col_window=25,
    )
    forecast_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [
            ("tot", "fcst", "w", "o", "sa"),
            ("tot", "fcst", "without", "sa"),
            ("forecast", "without", "sa"),
        ],
        row_window=8,
        col_window=55,
    )
    if forecast_col is None:
        forecast_col = anchor_col - 1

    actual_col = find_column_near_anchor(
        cache,
        anchor_row,
        anchor_col,
        [("actual", "sales"), ("reported", "sales"), ("actual", "value")],
        row_window=8,
        col_window=55,
    )

    candidate_rows, candidate_quarters = infer_candidate_rows(cache, anchor_row, anchor_col, N_QUARTERS_MAX)
    if num_quarters_col is not None:
        updated_quarters: List[int] = []
        for idx, row in enumerate(candidate_rows):
            value = to_int(cache.get(row, num_quarters_col))
            updated_quarters.append(value if value is not None else candidate_quarters[idx])
        candidate_quarters = updated_quarters

    # Regression anchor offsets (required).
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    pair_rows = collect_paired_numeric_rows(cache, x_col=x_col, y_col=y_col, row_end=anchor_row - 1)

    intercept_by_index: Dict[int, Any] = {}
    slope_by_index: Dict[int, Any] = {}
    if pair_rows:
        helper_start_row = cache.last_row + 2
        helper_intercept_col = cache.last_col + 2
        helper_slope_col = cache.last_col + 3
        formula_indices: List[int] = []

        for idx, n in enumerate(candidate_quarters):
            if n is None or n <= 0 or n > len(pair_rows):
                continue
            start_row = pair_rows[-n]
            end_row = pair_rows[-1]
            intercept_formula = (
                f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            )
            set_formula2(ws.range((helper_start_row + idx, helper_intercept_col)), intercept_formula)
            set_formula2(ws.range((helper_start_row + idx, helper_slope_col)), slope_formula)
            formula_indices.append(idx)

        if formula_indices:
            wb.app.calculate()
            intercept_values = ws.range(
                (helper_start_row, helper_intercept_col),
                (helper_start_row + len(candidate_rows) - 1, helper_intercept_col),
            ).value
            slope_values = ws.range(
                (helper_start_row, helper_slope_col),
                (helper_start_row + len(candidate_rows) - 1, helper_slope_col),
            ).value
            intercept_values = flatten_excel_vector(intercept_values)
            slope_values = flatten_excel_vector(slope_values)
            for idx in formula_indices:
                intercept_by_index[idx] = intercept_values[idx]
                slope_by_index[idx] = slope_values[idx]

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Optional[float], ...]] = None
    final_index = len(candidate_rows) - 1

    for idx, row in enumerate(candidate_rows):
        n_quarters = candidate_quarters[idx] if idx < len(candidate_quarters) else (idx + 1)
        if n_quarters is None:
            n_quarters = idx + 1

        forecast_max = cache.get(row, anchor_col)
        forecast_min = cache.get(row, min_col)
        forecast_value = cache.get(row, forecast_col)
        actual_value = cache.get(row, actual_col) if actual_col is not None else ""
        intercept = intercept_by_index.get(idx)
        slope = slope_by_index.get(idx)
        range_width = calc_range_width(forecast_max, forecast_min)

        if all(value in (None, "") for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        signature = (
            canonical_float(forecast_value),
            canonical_float(forecast_max),
            canonical_float(forecast_min),
            canonical_float(intercept),
            canonical_float(slope),
        )
        # Prevent duplicate final row.
        if idx == final_index and prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
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
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def source_files_from_input(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            print(f"Skipped file: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files_from_input(input_path):
            print(f"Processing file: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_label(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                files_processed += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    final_output_path = make_unique_output_path(input_path, output_path)
    write_output_workbook(final_output_path, empirical_rows, regression_rows)

    print(f"Output path: {final_output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
