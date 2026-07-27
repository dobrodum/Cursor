#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from .xlsx files.

The script opens each source workbook once, processes both model sheets while it
is open, then closes without saving any source changes.
"""

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

# ---- User-configurable paths -------------------------------------------------
input_dir = Path("input")
output_dir = Path("output")
# ------------------------------------------------------------------------------

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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
MONTH_MAP = {name[:3].lower(): idx for idx, name in enumerate(calendar.month_name) if name}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class UsedGrid:
    values: List[List[Any]]
    top_row: int
    left_col: int

    @property
    def bottom_row(self) -> int:
        return self.top_row + len(self.values) - 1 if self.values else self.top_row

    @property
    def right_col(self) -> int:
        return self.left_col + len(self.values[0]) - 1 if self.values and self.values[0] else self.left_col

    def value_at(self, abs_row: int, abs_col: int) -> Any:
        row_idx = abs_row - self.top_row
        col_idx = abs_col - self.left_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        row = self.values[row_idx]
        if col_idx >= len(row):
            return None
        return row[col_idx]


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


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "").replace("%", "")
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_column_values(values: Any, expected_count: int) -> List[Any]:
    if expected_count <= 0:
        return []
    if expected_count == 1:
        return [values]
    if not isinstance(values, list):
        return [values]
    normalized: List[Any] = []
    for item in values:
        if isinstance(item, list):
            normalized.append(item[0] if item else None)
        else:
            normalized.append(item)
    return normalized


def find_label_cell(grid: UsedGrid, keywords: Sequence[str]) -> Optional[Tuple[int, int]]:
    keyset = [k.lower() for k in keywords]
    for r_idx, row in enumerate(grid.values):
        for c_idx, value in enumerate(row):
            if not isinstance(value, str):
                continue
            text = value.strip().lower()
            if any(key in text for key in keyset):
                return (grid.top_row + r_idx, grid.left_col + c_idx)
    return None


def find_anchor_cell(grid: UsedGrid, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    target = anchor_text.strip().lower()
    for r_idx, row in enumerate(grid.values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target:
                return (grid.top_row + r_idx, grid.left_col + c_idx)
    return None


def first_numeric_to_right(grid: UsedGrid, abs_row: int, abs_col: int) -> Optional[float]:
    row_idx = abs_row - grid.top_row
    if row_idx < 0 or row_idx >= len(grid.values):
        return None
    row_values = grid.values[row_idx]
    start_col_idx = max(abs_col - grid.left_col + 1, 0)
    for col_idx in range(start_col_idx, len(row_values)):
        numeric_value = to_float(row_values[col_idx])
        if numeric_value is not None:
            return numeric_value
    return None


def numeric_series_from_row(grid: UsedGrid, abs_row: int, start_abs_col: int) -> List[Tuple[int, float]]:
    row_idx = abs_row - grid.top_row
    if row_idx < 0 or row_idx >= len(grid.values):
        return []
    row_values = grid.values[row_idx]
    start_idx = max(start_abs_col - grid.left_col, 0)
    series: List[Tuple[int, float]] = []
    for col_idx in range(start_idx, len(row_values)):
        numeric_value = to_float(row_values[col_idx])
        if numeric_value is not None:
            abs_col = grid.left_col + col_idx
            series.append((abs_col, numeric_value))
    return series


def metric_value_by_keywords(grid: UsedGrid, keywords: Sequence[str]) -> Optional[float]:
    label_cell = find_label_cell(grid, keywords)
    if not label_cell:
        return None
    return first_numeric_to_right(grid, label_cell[0], label_cell[1])


def safe_percent_to_decimal(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if abs(value) > 1:
        return value / 100.0
    return value


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    # Example: MedMiner_Model - AORT - MidJan2026_Send
    pattern = re.compile(
        r".*-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)([A-Za-z]{3,9})\s*([0-9]{4})(?:_Send)?$",
        flags=re.IGNORECASE,
    )
    match = pattern.match(stem)
    if match:
        ticker = match.group(1).upper()
        period_prefix = match.group(2).title()
        month_token = match.group(3)[:3].lower()
        year = int(match.group(4))
        month_number = MONTH_MAP.get(month_token)
        period_day = PERIOD_DAY_MAP[period_prefix.lower()]

        model_period = f"{period_prefix}{month_token.title()}_{year}"
        model_date = ""
        if month_number:
            model_date = date(year, month_number, period_day).isoformat()
        model = f"{ticker}_{model_period}"
        return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    # Fallback parser keeps script resilient when naming is not exact.
    parts = [p.strip() for p in stem.split("-") if p.strip()]
    ticker = "UNKNOWN"
    if len(parts) >= 2:
        ticker = parts[1].split()[0].upper()
    model_period = "unknown_period"
    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date="")


def get_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    input_folder_name = in_dir.name
    base_name = f"{input_folder_name}_PARAM.xlsx"
    base_path = out_dir / base_name
    if not base_path.exists():
        return base_path

    index = 1
    while True:
        candidate = out_dir / f"{input_folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def get_used_grid(sheet: xw.Sheet) -> UsedGrid:
    used = sheet.used_range
    values_2d = ensure_2d(used.value)
    return UsedGrid(values=values_2d, top_row=used.row, left_col=used.column)


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    # Prefer formula2 for dynamic-array aware modern Excel engines.
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def close_workbook_no_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    # Last fallback. With display alerts off, this still avoids source edits.
    try:
        wb.close()
    except Exception as exc:
        print(f"Warning: failed to close workbook cleanly: {exc}")


def extract_empirical_candidates(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        return rows

    sheet = wb.sheets["Empirical Model"]
    grid = get_used_grid(sheet)
    anchor = find_anchor_cell(grid, "max")
    if not anchor:
        return rows
    anchor_row, anchor_col = anchor

    penetration_label = find_label_cell(grid, ("penetration", "avg penetration"))
    if penetration_label:
        penetration_row, penetration_label_col = penetration_label
    else:
        penetration_row = anchor_row - 1
        penetration_label_col = max(grid.left_col, anchor_col - 24)

    penetration_series = numeric_series_from_row(grid, penetration_row, penetration_label_col + 1)
    if not penetration_series:
        return rows

    n_quarters = min(N_QUARTERS, len(penetration_series))
    recent_series = penetration_series[-n_quarters:]
    series_last_col = recent_series[-1][0]

    # Use anchor-based offsets for temporary formula writes.
    scratch_col = anchor_col + 8
    scratch_start_row = anchor_row + 1
    for idx, quarter_count in enumerate(range(1, n_quarters + 1)):
        window_first_col = recent_series[-quarter_count][0]
        avg_formula = (
            f"=AVERAGE(R{penetration_row}C{window_first_col}:R{penetration_row}C{series_last_col})"
        )
        set_formula2(sheet.cells(scratch_start_row + idx, scratch_col), avg_formula)

    wb.app.calculate()
    avg_values_raw = sheet.range(
        (scratch_start_row, scratch_col),
        (scratch_start_row + n_quarters - 1, scratch_col),
    ).value
    avg_values = normalize_column_values(avg_values_raw, n_quarters)

    quarterly_sales = metric_value_by_keywords(grid, ("quarterly sales", "qtr sales"))
    reported_sales = metric_value_by_keywords(grid, ("reported sales", "actual sales"))
    growth_rate_pct = metric_value_by_keywords(grid, ("growth rate", "growth %"))
    sales_captured_in_db_pct = metric_value_by_keywords(
        grid, ("sales captured in db", "captured in db", "captured %")
    )

    quarter_label_row = penetration_row - 1
    quarter_labels = {col: grid.value_at(quarter_label_row, col) for col, _ in recent_series}

    for idx in range(n_quarters):
        num_quarters_used = idx + 1
        avg_penetration_pct = to_float(avg_values[idx])
        avg_pen_decimal = safe_percent_to_decimal(avg_penetration_pct)

        model_row = anchor_row + num_quarters_used
        forecast_value = to_float(sheet.cells(model_row, anchor_col - 1).value)
        forecast_max = to_float(sheet.cells(model_row, anchor_col).value)
        forecast_min = to_float(sheet.cells(model_row, anchor_col + 1).value)

        if forecast_value is None and quarterly_sales is not None and avg_pen_decimal not in (None, 0):
            forecast_value = quarterly_sales / avg_pen_decimal

        if forecast_max is None and forecast_value is not None and growth_rate_pct is not None:
            growth_decimal = abs(safe_percent_to_decimal(growth_rate_pct) or 0.0)
            forecast_max = forecast_value * (1 + growth_decimal)
        if forecast_min is None and forecast_value is not None and growth_rate_pct is not None:
            growth_decimal = abs(safe_percent_to_decimal(growth_rate_pct) or 0.0)
            forecast_min = forecast_value * (1 - growth_decimal)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        first_window_col = recent_series[-num_quarters_used][0]
        last_quarter_used = quarter_labels.get(first_window_col)

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
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

    return rows


def rows_equal_for_regression_dedup(row_a: Dict[str, Any], row_b: Dict[str, Any]) -> bool:
    compare_fields = [
        "num_quarters_used",
        "intercept",
        "slope",
        "forecast_value",
        "forecast_max",
        "forecast_min",
    ]
    for field in compare_fields:
        a_val = row_a.get(field)
        b_val = row_b.get(field)
        if isinstance(a_val, (int, float)) and isinstance(b_val, (int, float)):
            if not math.isclose(float(a_val), float(b_val), rel_tol=0, abs_tol=1e-9):
                return False
        else:
            if a_val != b_val:
                return False
    return True


def extract_regression_candidates(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if "Regression Model" not in [s.name for s in wb.sheets]:
        return rows

    sheet = wb.sheets["Regression Model"]
    grid = get_used_grid(sheet)
    anchor = find_anchor_cell(grid, "max")
    if not anchor:
        return rows
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    data_points: List[Tuple[int, float, float]] = []
    for abs_row in range(anchor_row + 1, grid.bottom_row + 1):
        x_val = to_float(grid.value_at(abs_row, x_col))
        y_val = to_float(grid.value_at(abs_row, y_col))
        if x_val is not None and y_val is not None:
            data_points.append((abs_row, x_val, y_val))

    if not data_points:
        return rows

    n_quarters = min(N_QUARTERS, len(data_points))
    scratch_col = anchor_col + 8
    scratch_start_row = anchor_row + 1

    for idx, quarter_count in enumerate(range(1, n_quarters + 1)):
        subset = data_points[-quarter_count:]
        first_row = subset[0][0]
        last_row = subset[-1][0]
        formula_row = scratch_start_row + idx

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )

        set_formula2(sheet.cells(formula_row, scratch_col), intercept_formula)
        set_formula2(sheet.cells(formula_row, scratch_col + 1), slope_formula)

    wb.app.calculate()
    intercept_raw = sheet.range(
        (scratch_start_row, scratch_col),
        (scratch_start_row + n_quarters - 1, scratch_col),
    ).value
    slope_raw = sheet.range(
        (scratch_start_row, scratch_col + 1),
        (scratch_start_row + n_quarters - 1, scratch_col + 1),
    ).value
    intercept_values = normalize_column_values(intercept_raw, n_quarters)
    slope_values = normalize_column_values(slope_raw, n_quarters)

    for idx in range(n_quarters):
        num_quarters_used = idx + 1
        intercept = to_float(intercept_values[idx])
        slope = to_float(slope_values[idx])

        model_row = anchor_row + num_quarters_used
        forecast_total_without_sa = to_float(sheet.cells(model_row, anchor_col - 1).value)
        forecast_max = to_float(sheet.cells(model_row, anchor_col).value)
        forecast_min = to_float(sheet.cells(model_row, anchor_col + 1).value)

        if forecast_total_without_sa is None and intercept is not None and slope is not None:
            x_latest = data_points[-1][1]
            forecast_total_without_sa = intercept + slope * x_latest

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_total_without_sa,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if len(rows) >= 2 and rows_equal_for_regression_dedup(rows[-1], rows[-2]):
        rows.pop()

    return rows


def style_output_sheet(ws, headers: Sequence[str]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    for col_idx, _ in enumerate(headers, start=1):
        max_len = 10
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)) + 2)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len, 48)


def write_output_workbook(
    destination: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_empirical = wb.create_sheet("empirical_candidates")
    ws_empirical.append(EMPIRICAL_COLUMNS)
    for row_data in empirical_rows:
        ws_empirical.append([row_data.get(col) for col in EMPIRICAL_COLUMNS])
    style_output_sheet(ws_empirical, EMPIRICAL_COLUMNS)

    ws_regression = wb.create_sheet("regression_candidates")
    ws_regression.append(REGRESSION_COLUMNS)
    for row_data in regression_rows:
        ws_regression.append([row_data.get(col) for col in REGRESSION_COLUMNS])
    style_output_sheet(ws_regression, REGRESSION_COLUMNS)

    wb.save(destination)


def process_workbooks() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a folder: {input_dir}")

    output_path = get_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calc_mode = app.calculation
    app.calculation = "manual"

    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(extract_empirical_candidates(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_candidates(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_workbook_no_save(wb)
    finally:
        app.calculation = original_calc_mode
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output written to: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
