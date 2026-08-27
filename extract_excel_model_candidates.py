#!/usr/bin/env python3
"""
Extract empirical and regression candidates from .xlsx model files.

The script opens each source workbook only once, processes both model sheets
while it is open, and writes a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import math
import re
import statistics
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# Configure input/output paths
# -----------------------------
input_dir = Path("input")
output_dir = Path("output")


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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

FILE_PARSE_RE = re.compile(
    r"""
    ^.*?-\s*
    (?P<ticker>[A-Za-z0-9]+)\s*-\s*
    (?P<period_tag>Early|Mid|Late)
    (?P<month>[A-Za-z]{3,9})
    [_\-\s]*
    (?P<year>\d{4})
    """,
    re.IGNORECASE | re.VERBOSE,
)


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and not math.isnan(value)


def to_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    return None


def normalized_ratio(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    # Handle percent-as-whole-number values (e.g. 37 for 37%)
    if abs(value) > 1:
        return value / 100.0
    return value


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator is None:
        return None
    denominator = normalized_ratio(denominator)
    if denominator in (None, 0):
        return None
    return numerator / denominator


def safe_pct(part: Optional[float], whole: Optional[float]) -> Optional[float]:
    if part is None or whole in (None, 0):
        return None
    return part / whole


def find_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_folder / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def parse_model_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    match = FILE_PARSE_RE.search(stem)
    if not match:
        parts = [p.strip() for p in stem.split("-")]
        fallback_ticker = parts[1].upper() if len(parts) > 1 else "UNKNOWN"
        return {
            "model": stem,
            "ticker": fallback_ticker,
            "model_period": "unknown",
            "model_date": "",
        }

    ticker = match.group("ticker").upper()
    period_tag = match.group("period_tag").title()
    month_key = match.group("month")[:3].lower()
    year = int(match.group("year"))

    month_number = MONTH_MAP.get(month_key)
    day = PERIOD_DAY_MAP[period_tag.lower()]
    if month_number is None:
        model_date = ""
        model_period = f"{period_tag}{match.group('month')}_{year}"
    else:
        model_date = date(year, month_number, day).isoformat()
        model_period = f"{period_tag}{match.group('month')[:3].title()}_{year}"

    return {
        "model": f"{ticker}_{model_period}",
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def get_used_grid(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    return values, used.row, used.column


def grid_get(
    grid: Sequence[Sequence[Any]], base_row: int, base_col: int, abs_row: int, abs_col: int
) -> Any:
    r = abs_row - base_row
    c = abs_col - base_col
    if r < 0 or c < 0:
        return None
    if r >= len(grid):
        return None
    if c >= len(grid[r]):
        return None
    return grid[r][c]


def find_anchor_max(
    grid: Sequence[Sequence[Any]], base_row: int, base_col: int
) -> Optional[Tuple[int, int]]:
    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return base_row + r_idx, base_col + c_idx
    return None


def scan_string_cells(
    grid: Sequence[Sequence[Any]], base_row: int, base_col: int
) -> List[Tuple[int, int, str]]:
    cells: List[Tuple[int, int, str]] = []
    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row):
            if isinstance(value, str):
                text = value.strip().lower()
                if text:
                    cells.append((base_row + r_idx, base_col + c_idx, text))
    return cells


def find_col_by_keywords(
    string_cells: Sequence[Tuple[int, int, str]],
    keywords: Sequence[str],
    anchor_row: int,
    anchor_col: int,
    max_col_distance: int = 40,
) -> Optional[int]:
    best: Optional[Tuple[int, int]] = None  # (score, abs_col)
    keyset = [k.lower() for k in keywords]
    for row, col, text in string_cells:
        if not all(k in text for k in keyset):
            continue
        if abs(col - anchor_col) > max_col_distance:
            continue
        # Score near the anchor row/col first.
        score = abs(row - anchor_row) * 5 + abs(col - anchor_col)
        if best is None or score < best[0]:
            best = (score, col)
    return None if best is None else best[1]


def collect_numeric_rows(
    grid: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    row_start: int,
    row_end: int,
    required_cols: Sequence[int],
) -> List[int]:
    rows: List[int] = []
    for abs_row in range(row_start, row_end + 1):
        ok = True
        for col in required_cols:
            if not is_number(grid_get(grid, base_row, base_col, abs_row, col)):
                ok = False
                break
        if ok:
            rows.append(abs_row)
    return rows


def close_source_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        # Continue to fallback variants below.
        pass

    for closer in (lambda: wb.close(False), wb.close):
        try:
            closer()
            return
        except Exception:
            continue


def extract_empirical_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    grid, base_row, base_col = get_used_grid(sheet)
    if not grid:
        return []

    anchor = find_anchor_max(grid, base_row, base_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    string_cells = scan_string_cells(grid, base_row, base_col)

    quarter_col = find_col_by_keywords(string_cells, ["quarter"], anchor_row, anchor_col)
    quarterly_sales_col = find_col_by_keywords(
        string_cells, ["quarterly", "sales"], anchor_row, anchor_col
    )
    reported_sales_col = find_col_by_keywords(
        string_cells, ["reported", "sales"], anchor_row, anchor_col
    )
    penetration_col = find_col_by_keywords(
        string_cells, ["penetration"], anchor_row, anchor_col
    )
    growth_col = find_col_by_keywords(string_cells, ["growth"], anchor_row, anchor_col)
    captured_col = find_col_by_keywords(
        string_cells, ["captured", "db"], anchor_row, anchor_col
    )

    # Fallbacks remain anchor-relative so logic stays layout-resilient.
    if penetration_col is None:
        penetration_col = anchor_col - 11
    if quarterly_sales_col is None:
        quarterly_sales_col = anchor_col - 7
    if reported_sales_col is None:
        reported_sales_col = anchor_col - 6

    row_start = base_row
    row_end = anchor_row - 1
    numeric_rows = collect_numeric_rows(
        grid,
        base_row,
        base_col,
        row_start=row_start,
        row_end=row_end,
        required_cols=[penetration_col, quarterly_sales_col, reported_sales_col],
    )
    if not numeric_rows:
        return []

    n_quarters = 10
    max_n = min(n_quarters, len(numeric_rows))

    scratch_col = base_col + max(len(r) for r in grid) + 2
    avg_cell = sheet.range((anchor_row, scratch_col))
    min_cell = sheet.range((anchor_row + 1, scratch_col))
    max_cell = sheet.range((anchor_row + 2, scratch_col))

    rows: List[Dict[str, Any]] = []
    for n in range(1, max_n + 1):
        row_subset = numeric_rows[-n:]
        start_row = row_subset[0]
        end_row = row_subset[-1]

        avg_cell.formula2 = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        )
        min_cell.formula2 = f"=MIN(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        max_cell.formula2 = f"=MAX(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        wb.app.calculate()

        avg_penetration_pct = to_float(avg_cell.value)
        min_penetration_pct = to_float(min_cell.value)
        max_penetration_pct = to_float(max_cell.value)

        quarterly_sales = to_float(
            grid_get(grid, base_row, base_col, end_row, quarterly_sales_col)
        )
        reported_sales = to_float(
            grid_get(grid, base_row, base_col, end_row, reported_sales_col)
        )

        forecast_value = safe_div(quarterly_sales, avg_penetration_pct)
        forecast_max = safe_div(quarterly_sales, min_penetration_pct)
        forecast_min = safe_div(quarterly_sales, max_penetration_pct)
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if len(row_subset) >= 2:
            prior_sales = to_float(
                grid_get(grid, base_row, base_col, row_subset[-2], quarterly_sales_col)
            )
        else:
            prior_sales = None
        growth_rate_pct = (
            ((quarterly_sales / prior_sales) - 1)
            if quarterly_sales is not None and prior_sales not in (None, 0)
            else None
        )

        sales_captured_in_db_pct = safe_pct(quarterly_sales, reported_sales)

        quarter_value = (
            grid_get(grid, base_row, base_col, start_row, quarter_col)
            if quarter_col is not None
            else None
        )
        if growth_col is not None:
            sheet_growth = to_float(grid_get(grid, base_row, base_col, end_row, growth_col))
            if sheet_growth is not None:
                growth_rate_pct = sheet_growth
        if captured_col is not None:
            sheet_captured = to_float(grid_get(grid, base_row, base_col, end_row, captured_col))
            if sheet_captured is not None:
                sales_captured_in_db_pct = sheet_captured

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n,
                "last_quarter_used": quarter_value,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
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
        )

    # Clean up temporary formulas.
    avg_cell.value = None
    min_cell.value = None
    max_cell.value = None
    wb.app.calculate()

    return rows


def almost_equal(a: Any, b: Any, tol: float = 1e-9) -> bool:
    fa = to_float(a)
    fb = to_float(b)
    if fa is None or fb is None:
        return a == b
    return abs(fa - fb) <= tol


def extract_regression_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    grid, base_row, base_col = get_used_grid(sheet)
    if not grid:
        return []

    anchor = find_anchor_max(grid, base_row, base_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    row_start = base_row
    row_end = anchor_row - 1
    numeric_rows = collect_numeric_rows(
        grid,
        base_row,
        base_col,
        row_start=row_start,
        row_end=row_end,
        required_cols=[x_col, y_col],
    )
    if len(numeric_rows) < 2:
        return []

    max_n = min(10, len(numeric_rows))

    scratch_col = base_col + max(len(r) for r in grid) + 2
    intercept_cell = sheet.range((anchor_row, scratch_col))
    slope_cell = sheet.range((anchor_row + 1, scratch_col))

    rows: List[Dict[str, Any]] = []
    prev_row: Optional[Dict[str, Any]] = None

    for n in range(2, max_n + 1):
        row_subset = numeric_rows[-n:]
        start_row = row_subset[0]
        end_row = row_subset[-1]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        wb.app.calculate()

        intercept_value = to_float(intercept_cell.value)
        slope_value = to_float(slope_cell.value)

        x_values = [
            to_float(grid_get(grid, base_row, base_col, r, x_col))
            for r in row_subset
        ]
        y_values = [
            to_float(grid_get(grid, base_row, base_col, r, y_col))
            for r in row_subset
        ]

        if intercept_value is None or slope_value is None:
            continue

        x_latest = x_values[-1]
        forecast_total_without_sa = (
            intercept_value + slope_value * x_latest if x_latest is not None else None
        )

        residuals: List[float] = []
        for xv, yv in zip(x_values, y_values):
            if xv is None or yv is None:
                continue
            residuals.append(yv - (intercept_value + slope_value * xv))

        residual_std = statistics.pstdev(residuals) if len(residuals) >= 2 else 0.0
        forecast_max = (
            forecast_total_without_sa + residual_std
            if forecast_total_without_sa is not None
            else None
        )
        forecast_min = (
            forecast_total_without_sa - residual_std
            if forecast_total_without_sa is not None
            else None
        )
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        current = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n,
            "num_quarters_used": n,
            "forecast_value": forecast_total_without_sa,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        if prev_row is not None:
            duplicate = all(
                [
                    almost_equal(current["forecast_value"], prev_row["forecast_value"]),
                    almost_equal(current["forecast_max"], prev_row["forecast_max"]),
                    almost_equal(current["forecast_min"], prev_row["forecast_min"]),
                    almost_equal(current["intercept"], prev_row["intercept"]),
                    almost_equal(current["slope"], prev_row["slope"]),
                ]
            )
            if duplicate:
                continue

        rows.append(current)
        prev_row = current

    intercept_cell.value = None
    slope_cell.value = None
    wb.app.calculate()

    return rows


def autosize_columns(ws) -> None:
    for idx, column in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
        max_len = 0
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_len:
                max_len = len(value)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    for sheet_name, columns, rows in (
        ("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows),
        ("regression_candidates", REGRESSION_COLUMNS, regression_rows),
    ):
        ws = wb.create_sheet(sheet_name)
        ws.append(columns)
        for row in rows:
            ws.append([row.get(col) for col in columns])

        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autosize_columns(ws)

    wb.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"SKIPPED: {path.name} (temp file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"SKIPPED: {path.name} (not .xlsx)")
            continue
        yield path


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = find_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_input_files(input_dir):
            print(f"PROCESSING: {file_path.name}")
            labels = parse_model_labels(file_path)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_rows(wb, labels=labels, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_rows(wb, labels=labels, source_file=file_path.name)
                )
                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook_safely(wb)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {processed_files}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
