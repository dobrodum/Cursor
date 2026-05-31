#!/usr/bin/env python3
"""Extract empirical/regression parameter candidates from Excel model workbooks.

This script opens each source workbook exactly once, processes both
"Empirical Model" and "Regression Model" while it is open, and writes one
combined output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import calendar
import re
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


# ---------------------------------------------------------------------------
# Output schema
# ---------------------------------------------------------------------------
EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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


# ---------------------------------------------------------------------------
# Anchor-driven offsets (primary reads). Label lookups are fallback.
# ---------------------------------------------------------------------------
EMPIRICAL_OFFSETS: Dict[str, Tuple[int, int]] = {
    "forecast_value": (-1, 1),  # estimated total sold
    "actual_value": (2, 1),  # reported sales
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
    "quarterly_sales": (3, 1),
    "growth_rate_pct": (4, 1),
    "sales_captured_in_db_pct": (5, 1),
}
EMPIRICAL_NUM_QUARTERS_INPUT_OFFSET = (-2, 1)

REGRESSION_OFFSETS: Dict[str, Tuple[int, int]] = {
    "forecast_value": (-1, 1),  # TOT FCST w/o SA
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
    "actual_value": (2, 1),
}


LABEL_VALUE_OFFSETS: Tuple[Tuple[int, int], ...] = (
    (0, 1),
    (1, 0),
    (0, -1),
    (-1, 0),
    (1, 1),
    (1, -1),
    (-1, 1),
    (-1, -1),
    (0, 2),
    (2, 0),
)


@dataclass(frozen=True)
class FileModelInfo:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetCache:
    start_row: int
    start_col: int
    row_count: int
    col_count: int
    values: List[List[Any]]


MONTH_MAP = {
    month.lower(): month_index
    for month_index, month in enumerate(calendar.month_abbr)
    if month
}
PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def safe_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    return None


def values_equal(a: Any, b: Any, tol: float = 1e-10) -> bool:
    if a is None and b is None:
        return True
    if is_number(a) and is_number(b):
        return abs(float(a) - float(b)) <= tol
    return a == b


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if len(values) == 0:
        return []
    first = values[0]
    if isinstance(first, (list, tuple)):
        return [list(row) for row in values]
    return [list(values)]


def build_sheet_cache(sheet: xw.Sheet) -> SheetCache:
    used = sheet.used_range
    matrix = ensure_2d(used.value)
    return SheetCache(
        start_row=used.row,
        start_col=used.column,
        row_count=len(matrix),
        col_count=len(matrix[0]) if matrix else 0,
        values=matrix,
    )


def cache_value(cache: SheetCache, row: int, col: int) -> Any:
    row_i = row - cache.start_row
    col_i = col - cache.start_col
    if row_i < 0 or col_i < 0:
        return None
    if row_i >= cache.row_count or col_i >= cache.col_count:
        return None
    return cache.values[row_i][col_i]


def find_max_anchor(cache: SheetCache) -> Optional[Tuple[int, int]]:
    fallback: Optional[Tuple[int, int]] = None
    for row_i, row_values in enumerate(cache.values):
        for col_i, value in enumerate(row_values):
            if not isinstance(value, str):
                continue
            label = normalize_label(value)
            abs_row = cache.start_row + row_i
            abs_col = cache.start_col + col_i
            if label == "max":
                return abs_row, abs_col
            if fallback is None and label.endswith(" max"):
                fallback = (abs_row, abs_col)
    return fallback


def build_label_index(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    row_window: int = 40,
    col_window: int = 25,
) -> Dict[str, List[Tuple[int, int]]]:
    labels: Dict[str, List[Tuple[int, int]]] = {}
    for row_i, row_values in enumerate(cache.values):
        abs_row = cache.start_row + row_i
        if abs(abs_row - anchor_row) > row_window:
            continue
        for col_i, value in enumerate(row_values):
            abs_col = cache.start_col + col_i
            if abs(abs_col - anchor_col) > col_window:
                continue
            if not isinstance(value, str):
                continue
            label = normalize_label(value)
            if not label:
                continue
            labels.setdefault(label, []).append((abs_row, abs_col))
    return labels


def find_label_cell(
    label_index: Dict[str, List[Tuple[int, int]]],
    anchor_row: int,
    anchor_col: int,
    tokens: Sequence[str],
) -> Optional[Tuple[int, int]]:
    best_cell: Optional[Tuple[int, int]] = None
    best_distance: Optional[int] = None
    normalized_tokens = [normalize_label(token) for token in tokens if token]
    if not normalized_tokens:
        return None
    for label, cells in label_index.items():
        if not any(token in label for token in normalized_tokens):
            continue
        for row, col in cells:
            distance = abs(anchor_row - row) + abs(anchor_col - col)
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_cell = (row, col)
    return best_cell


def value_ok(value: Any, numeric: bool) -> bool:
    if numeric:
        return is_number(value)
    return value not in (None, "")


def resolve_metric_cell(
    sheet: xw.Sheet,
    cache: SheetCache,
    label_index: Dict[str, List[Tuple[int, int]]],
    anchor_row: int,
    anchor_col: int,
    offset: Tuple[int, int],
    label_tokens: Sequence[str],
    numeric: bool,
) -> Optional[Tuple[int, int]]:
    # Primary path: anchor-based offsets.
    primary = (anchor_row + offset[0], anchor_col + offset[1])
    primary_value = cache_value(cache, primary[0], primary[1])
    if value_ok(primary_value, numeric):
        return primary

    # Fallback path: nearest matching label + adjacent value cell.
    label_cell = find_label_cell(label_index, anchor_row, anchor_col, label_tokens)
    if not label_cell:
        return primary
    for d_row, d_col in LABEL_VALUE_OFFSETS:
        row = label_cell[0] + d_row
        col = label_cell[1] + d_col
        candidate_value = cache_value(cache, row, col)
        if candidate_value is None:
            candidate_value = sheet.cells(row, col).value
        if value_ok(candidate_value, numeric):
            return row, col
    return primary


def read_coord(sheet: xw.Sheet, coord: Optional[Tuple[int, int]]) -> Any:
    if coord is None:
        return None
    return sheet.cells(coord[0], coord[1]).value


def resolve_optional_input_cell(
    cache: SheetCache,
    label_index: Dict[str, List[Tuple[int, int]]],
    anchor_row: int,
    anchor_col: int,
    default_offset: Tuple[int, int],
    label_tokens: Sequence[str],
) -> Optional[Tuple[int, int]]:
    fallback = (anchor_row + default_offset[0], anchor_col + default_offset[1])
    label_cell = find_label_cell(label_index, anchor_row, anchor_col, label_tokens)
    if not label_cell:
        return fallback
    for d_row, d_col in LABEL_VALUE_OFFSETS:
        row = label_cell[0] + d_row
        col = label_cell[1] + d_col
        value = cache_value(cache, row, col)
        if value is None:
            continue
        if is_number(value) or value in ("", None):
            return row, col
    return fallback


def longest_numeric_segment(
    cache: SheetCache,
    row: int,
    from_col: int,
) -> Optional[Tuple[int, int]]:
    start_col = max(from_col, cache.start_col)
    end_col = cache.start_col + cache.col_count - 1
    best: Optional[Tuple[int, int]] = None
    current_start: Optional[int] = None

    for col in range(start_col, end_col + 2):
        value = cache_value(cache, row, col) if col <= end_col else None
        if is_number(value):
            if current_start is None:
                current_start = col
            continue
        if current_start is not None:
            segment = (current_start, col - 1)
            if best is None or (segment[1] - segment[0]) > (best[1] - best[0]):
                best = segment
            current_start = None
    return best


def find_penetration_series(
    cache: SheetCache,
    label_index: Dict[str, List[Tuple[int, int]]],
    anchor_row: int,
    anchor_col: int,
) -> Optional[Tuple[int, int, int]]:
    candidates: List[Tuple[int, int]] = []
    for label, cells in label_index.items():
        if "penetration" in label:
            candidates.extend(cells)
    candidates.sort(key=lambda c: abs(c[0] - anchor_row) + abs(c[1] - anchor_col))

    for row, col in candidates:
        segment = longest_numeric_segment(cache, row, col + 1)
        if segment and (segment[1] - segment[0] + 1) >= 2:
            return row, segment[0], segment[1]

    # Fallback: nearest row with a long numeric streak.
    row_min = max(cache.start_row, anchor_row - 25)
    row_max = min(cache.start_row + cache.row_count - 1, anchor_row + 10)
    best: Optional[Tuple[int, int, int]] = None
    best_distance: Optional[int] = None
    for row in range(row_min, row_max + 1):
        segment = longest_numeric_segment(cache, row, cache.start_col)
        if not segment:
            continue
        length = segment[1] - segment[0] + 1
        if length < 4:
            continue
        distance = abs(row - anchor_row)
        if best is None or distance < (best_distance or 10**9):
            best = (row, segment[0], segment[1])
            best_distance = distance
    return best


def get_trailing_valid_rows(
    cache: SheetCache,
    x_col: int,
    y_col: int,
    anchor_row: int,
) -> List[int]:
    row_start = cache.start_row
    row_end = cache.start_row + cache.row_count - 1
    valid_rows = [
        row
        for row in range(row_start, row_end + 1)
        if row < anchor_row
        and is_number(cache_value(cache, row, x_col))
        and is_number(cache_value(cache, row, y_col))
    ]
    if not valid_rows:
        return []
    valid_rows.sort()
    trailing = [valid_rows[-1]]
    for row in reversed(valid_rows[:-1]):
        if trailing[0] - row == 1:
            trailing.insert(0, row)
        else:
            break
    return trailing


def parse_filename_info(file_name: str) -> Optional[FileModelInfo]:
    pattern = re.compile(
        r"Model\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
        r"(?P<period>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(file_name)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    period = match.group("period").title()
    month_abbr = match.group("month").title()
    year = int(match.group("year"))

    month_number = MONTH_MAP.get(month_abbr.lower())
    day = PERIOD_DAY_MAP.get(period.lower())
    if month_number is None or day is None:
        return None

    model_period = f"{period}{month_abbr}_{year}"
    model_date = f"{year:04d}-{month_number:02d}-{day:02d}"
    model = f"{ticker}_{model_period}"
    return FileModelInfo(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


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
        try:
            wb.close()
        except Exception:
            pass


def get_sheet_if_exists(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def extract_empirical_rows(
    wb: xw.Book,
    model_info: FileModelInfo,
    source_file_name: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, "Empirical Model")
    if sheet is None:
        print(f"  - skipped empirical: missing sheet 'Empirical Model'")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        print("  - skipped empirical: missing 'max' anchor")
        return []
    anchor_row, anchor_col = anchor
    label_index = build_label_index(cache, anchor_row, anchor_col)

    # Find the historical penetration row once and use it for all n_quarters.
    series = find_penetration_series(cache, label_index, anchor_row, anchor_col)
    if not series:
        print("  - skipped empirical: unable to locate penetration history")
        return []
    penetration_row, series_start_col, series_end_col = series
    available_periods = series_end_col - series_start_col + 1

    # Resolve frequently-read cells once (anchor offsets first, labels as fallback).
    metric_coords = {
        "forecast_value": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_value"],
            ["estimated total sold", "est total sold", "forecast total"],
            numeric=False,
        ),
        "actual_value": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            EMPIRICAL_OFFSETS["actual_value"],
            ["reported sales", "actual sales"],
            numeric=False,
        ),
        "forecast_max": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_max"],
            ["max", "forecast max"],
            numeric=True,
        ),
        "forecast_min": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_min"],
            ["min", "forecast min"],
            numeric=True,
        ),
        "quarterly_sales": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            EMPIRICAL_OFFSETS["quarterly_sales"],
            ["quarterly sales"],
            numeric=False,
        ),
        "growth_rate_pct": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            EMPIRICAL_OFFSETS["growth_rate_pct"],
            ["growth rate"],
            numeric=False,
        ),
        "sales_captured_in_db_pct": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
            ["sales captured in db", "captured in db"],
            numeric=False,
        ),
    }

    num_quarters_input = resolve_optional_input_cell(
        cache,
        label_index,
        anchor_row,
        anchor_col,
        EMPIRICAL_NUM_QUARTERS_INPUT_OFFSET,
        ["num quarters", "quarters used", "n quarters"],
    )

    # Scratch cells to keep formulas localized.
    scratch_col = cache.start_col + cache.col_count + 2
    avg_penetration_cell = sheet.cells(anchor_row, scratch_col)

    empirical_rows: List[Dict[str, Any]] = []
    max_quarters_to_test = 10
    for n_quarters in range(1, max_quarters_to_test + 1):
        if n_quarters > available_periods:
            break

        if num_quarters_input is not None:
            sheet.cells(num_quarters_input[0], num_quarters_input[1]).value = n_quarters

        start_col = series_end_col - n_quarters + 1
        avg_penetration_cell.formula2 = (
            f"=AVERAGE(R{penetration_row}C{start_col}:R{penetration_row}C{series_end_col})"
        )
        wb.app.calculate()

        avg_penetration = avg_penetration_cell.value
        forecast_value = read_coord(sheet, metric_coords["forecast_value"])
        actual_value = read_coord(sheet, metric_coords["actual_value"])
        forecast_max = read_coord(sheet, metric_coords["forecast_max"])
        forecast_min = read_coord(sheet, metric_coords["forecast_min"])
        quarterly_sales = read_coord(sheet, metric_coords["quarterly_sales"])
        growth_rate_pct = read_coord(sheet, metric_coords["growth_rate_pct"])
        sales_captured_pct = read_coord(sheet, metric_coords["sales_captured_in_db_pct"])

        # Fallback forecast calculation when the workbook output cell is blank.
        if forecast_value in (None, "") and is_number(quarterly_sales) and is_number(avg_penetration):
            avg_float = float(avg_penetration)
            if avg_float != 0:
                forecast_value = float(quarterly_sales) / avg_float

        forecast_max_float = safe_float(forecast_max)
        forecast_min_float = safe_float(forecast_min)
        range_width = (
            forecast_max_float - forecast_min_float
            if forecast_max_float is not None and forecast_min_float is not None
            else None
        )

        last_quarter_used = cache_value(cache, penetration_row - 1, series_end_col)
        reported_sales = actual_value

        empirical_rows.append(
            {
                "model": model_info.model,
                "ticker": model_info.ticker,
                "model_period": model_info.model_period,
                "model_date": model_info.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file_name,
            }
        )

    return empirical_rows


def extract_regression_rows(
    wb: xw.Book,
    model_info: FileModelInfo,
    source_file_name: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, "Regression Model")
    if sheet is None:
        print("  - skipped regression: missing sheet 'Regression Model'")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        print("  - skipped regression: missing 'max' anchor")
        return []
    anchor_row, anchor_col = anchor
    label_index = build_label_index(cache, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_rows = get_trailing_valid_rows(cache, x_col=x_col, y_col=y_col, anchor_row=anchor_row)
    if len(data_rows) < 2:
        print("  - skipped regression: not enough trailing x/y observations")
        return []

    metric_coords = {
        "forecast_value": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            REGRESSION_OFFSETS["forecast_value"],
            ["tot fcst w o sa", "tot fcst without sa", "forecast total"],
            numeric=False,
        ),
        "forecast_max": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            REGRESSION_OFFSETS["forecast_max"],
            ["max", "forecast max"],
            numeric=True,
        ),
        "forecast_min": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            REGRESSION_OFFSETS["forecast_min"],
            ["min", "forecast min"],
            numeric=True,
        ),
        "actual_value": resolve_metric_cell(
            sheet,
            cache,
            label_index,
            anchor_row,
            anchor_col,
            REGRESSION_OFFSETS["actual_value"],
            ["actual sales", "reported sales"],
            numeric=False,
        ),
    }

    scratch_col = cache.start_col + cache.col_count + 2
    intercept_cell = sheet.cells(anchor_row, scratch_col)
    slope_cell = sheet.cells(anchor_row + 1, scratch_col)

    data_end_row = data_rows[-1]
    max_quarters_to_test = min(10, len(data_rows))
    regression_rows: List[Dict[str, Any]] = []

    for n_quarters in range(2, max_quarters_to_test + 1):
        start_row = data_end_row - n_quarters + 1
        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        forecast_value = read_coord(sheet, metric_coords["forecast_value"])
        forecast_max = read_coord(sheet, metric_coords["forecast_max"])
        forecast_min = read_coord(sheet, metric_coords["forecast_min"])
        actual_value = read_coord(sheet, metric_coords["actual_value"])

        if forecast_value in (None, "") and is_number(intercept) and is_number(slope):
            x_target = cache_value(cache, data_end_row + 1, x_col)
            if not is_number(x_target):
                x_target = cache_value(cache, data_end_row, x_col)
            if is_number(x_target):
                forecast_value = float(intercept) + float(slope) * float(x_target)

        forecast_max_float = safe_float(forecast_max)
        forecast_min_float = safe_float(forecast_min)
        range_width = (
            forecast_max_float - forecast_min_float
            if forecast_max_float is not None and forecast_min_float is not None
            else None
        )

        row = {
            "model": model_info.model,
            "ticker": model_info.ticker,
            "model_period": model_info.model_period,
            "model_date": model_info.model_date,
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
            "source_file": source_file_name,
        }

        # Prevent the duplicate terminal row when calculations repeat.
        if regression_rows:
            prev = regression_rows[-1]
            duplicate = all(
                values_equal(row.get(key), prev.get(key))
                for key in ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
            )
            if duplicate:
                continue

        regression_rows.append(row)

    return regression_rows


def write_sheet(
    ws,
    columns: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column, None) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_index in range(2, ws.max_row + 1):
            value = ws.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max(max_length + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def list_source_files(folder: Path) -> Tuple[List[Path], List[Tuple[str, str]]]:
    if not folder.exists():
        raise FileNotFoundError(f"Input directory does not exist: {folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {folder}")

    process_files: List[Path] = []
    skipped: List[Tuple[str, str]] = []
    for item in sorted(folder.iterdir()):
        if not item.is_file():
            continue
        if item.name.startswith("~"):
            skipped.append((item.name, "temp file"))
            continue
        if item.suffix.lower() != ".xlsx":
            skipped.append((item.name, "not .xlsx"))
            continue
        process_files.append(item)
    return process_files, skipped


def extract_all(input_folder: Path, output_folder: Path) -> Path:
    source_files, skipped_files = list_source_files(input_folder)
    for file_name, reason in skipped_files:
        print(f"skipped file: {file_name} ({reason})")

    output_path = next_output_path(input_folder, output_folder)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    try:
        try:
            app.display_alerts = False
        except Exception:
            pass
        try:
            app.screen_updating = False
        except Exception:
            pass

        for file_path in source_files:
            model_info = parse_filename_info(file_path.name)
            if model_info is None:
                print(f"skipped file: {file_path.name} (unable to parse ticker/model period)")
                continue

            source_wb: Optional[xw.Book] = None
            try:
                # Open source workbook exactly once and process both model sheets.
                source_wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=source_wb,
                        model_info=model_info,
                        source_file_name=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=source_wb,
                        model_info=model_info,
                        source_file_name=file_path.name,
                    )
                )
                processed_count += 1
                print(f"processed file: {file_path.name}")
            except Exception as error:
                print(f"skipped file: {file_path.name} (processing error: {error})")
                traceback.print_exc()
            finally:
                if source_wb is not None:
                    # Never save source workbooks.
                    safe_close_workbook(source_wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_count}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")
    return output_path


def main() -> int:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()
    extract_all(in_dir, out_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
