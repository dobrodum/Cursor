#!/usr/bin/env python3
"""Extract empirical and regression parameter candidates from Excel model files.

This script processes every source workbook once, while extracting both the
"Empirical Model" and "Regression Model" sheets during that single open/close
cycle. Source workbooks are never saved or modified on disk.
"""

from __future__ import annotations

import itertools
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = r"/path/to/input/folder"
output_dir = r"/path/to/output/folder"

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

DAY_BY_PREFIX = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(values: Any) -> list[list[Any]]:
    """Normalize an Excel used-range payload into a 2D list."""
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> float | None:
    """Convert workbook cell values to float where possible."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1]
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def subtract(a: float | None, b: float | None) -> float | None:
    if a is None or b is None:
        return None
    return a - b


def multiply(a: float | None, b: float | None) -> float | None:
    if a is None or b is None:
        return None
    return a * b


def model_key(*values: float | None) -> tuple[Any, ...]:
    """Stable tuple key for duplicate filtering with floating point tolerance."""
    key: list[Any] = []
    for value in values:
        if value is None:
            key.append(None)
        else:
            key.append(round(value, 10))
    return tuple(key)


class SheetCache:
    """One-pass snapshot of an Excel sheet used range for fast lookups."""

    def __init__(self, base_row: int, base_col: int, values: list[list[Any]]):
        self.base_row = base_row
        self.base_col = base_col
        self.values = values
        self.height = len(values)
        self.width = max((len(row) for row in values), default=0)

    @classmethod
    def from_sheet(cls, sheet: xw.Sheet) -> "SheetCache":
        used = sheet.used_range
        return cls(base_row=used.row, base_col=used.column, values=normalize_2d(used.value))

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.base_row
        c_idx = col - self.base_col
        if r_idx < 0 or c_idx < 0 or r_idx >= self.height:
            return None
        row_data = self.values[r_idx]
        if c_idx >= len(row_data):
            return None
        return row_data[c_idx]

    def find_anchor(self, anchor_text: str = "max") -> tuple[int, int] | None:
        target = anchor_text.strip().lower()
        for r_idx, row in enumerate(self.values):
            for c_idx, value in enumerate(row):
                if isinstance(value, str) and value.strip().lower() == target:
                    return (self.base_row + r_idx, self.base_col + c_idx)
        return None

    def nearest_header_column(
        self,
        anchor_row: int,
        anchor_col: int,
        contains_all: tuple[str, ...],
    ) -> int | None:
        """Find nearest text cell containing all requested fragments."""
        best_col: int | None = None
        best_distance: int | None = None
        for r_idx, row in enumerate(self.values):
            for c_idx, value in enumerate(row):
                if not isinstance(value, str):
                    continue
                lowered = value.strip().lower()
                if not lowered:
                    continue
                if not all(fragment in lowered for fragment in contains_all):
                    continue
                abs_row = self.base_row + r_idx
                abs_col = self.base_col + c_idx
                distance = abs(abs_row - anchor_row) + abs(abs_col - anchor_col)
                if best_distance is None or distance < best_distance:
                    best_distance = distance
                    best_col = abs_col
        return best_col

    def has_numeric_window(self, col: int, start_row: int, end_row: int) -> bool:
        numeric_count = 0
        total = max(0, end_row - start_row + 1)
        if total == 0:
            return False
        for row in range(start_row, end_row + 1):
            if to_float(self.get(row, col)) is not None:
                numeric_count += 1
        return numeric_count >= total


def parse_file_label(file_path: Path) -> ModelMetadata | None:
    """Parse ticker/model period/date from file names like:
    MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        return None

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", stem, flags=re.IGNORECASE)
    if period_match is None:
        return None

    prefix = period_match.group(1).lower()
    month_token = period_match.group(2).title()
    year_text = period_match.group(3)

    if prefix not in DAY_BY_PREFIX:
        return None

    try:
        month_num = datetime.strptime(month_token, "%b").month
    except ValueError:
        return None

    model_period = f"{prefix.title()}{month_token}_{year_text}"
    model_date = date(int(year_text), month_num, DAY_BY_PREFIX[prefix]).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(input_path: Path, output_path: Path) -> Path:
    input_folder_name = input_path.name
    base_name = f"{input_folder_name}_PARAM.xlsx"
    base_path = output_path / base_name
    if not base_path.exists():
        return base_path

    for idx in itertools.count(1):
        candidate = output_path / f"{input_folder_name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate

    raise RuntimeError("Unable to determine a unique output filename.")


def close_source_workbook(wb: xw.Book | None) -> None:
    """Close source workbook without saving, with safe fallbacks."""
    if wb is None:
        return

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

    try:
        wb.close()
    except Exception:
        pass


def choose_column(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    default_offset: int,
    header_fragments: tuple[str, ...],
    data_start_row: int,
    data_end_row: int,
) -> int:
    default_col = anchor_col + default_offset
    if cache.has_numeric_window(default_col, data_start_row, data_end_row):
        return default_col
    header_col = cache.nearest_header_column(anchor_row, anchor_col, header_fragments)
    if header_col is not None:
        return header_col
    return default_col


def extract_empirical_candidates(
    wb: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if EMPIRICAL_SHEET_NAME not in [sheet.name for sheet in wb.sheets]:
        return rows

    sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    cache = SheetCache.from_sheet(sheet)
    anchor = cache.find_anchor("max")
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor
    latest_row = anchor_row - 1
    earliest_row = max(cache.base_row, latest_row - (N_QUARTERS - 1))

    quarter_col = choose_column(
        cache,
        anchor_row,
        anchor_col,
        default_offset=-11,
        header_fragments=("quarter",),
        data_start_row=earliest_row,
        data_end_row=latest_row,
    )
    penetration_col = choose_column(
        cache,
        anchor_row,
        anchor_col,
        default_offset=-7,
        header_fragments=("penetration",),
        data_start_row=earliest_row,
        data_end_row=latest_row,
    )
    quarterly_sales_col = choose_column(
        cache,
        anchor_row,
        anchor_col,
        default_offset=-5,
        header_fragments=("quarterly", "sales"),
        data_start_row=earliest_row,
        data_end_row=latest_row,
    )
    reported_sales_col = choose_column(
        cache,
        anchor_row,
        anchor_col,
        default_offset=-4,
        header_fragments=("reported", "sales"),
        data_start_row=earliest_row,
        data_end_row=latest_row,
    )
    growth_rate_col = choose_column(
        cache,
        anchor_row,
        anchor_col,
        default_offset=-3,
        header_fragments=("growth",),
        data_start_row=earliest_row,
        data_end_row=latest_row,
    )
    captured_in_db_col = choose_column(
        cache,
        anchor_row,
        anchor_col,
        default_offset=-2,
        header_fragments=("captured", "db"),
        data_start_row=earliest_row,
        data_end_row=latest_row,
    )

    avg_formula_cell = sheet.cells(anchor_row + 2, anchor_col + 2)

    for num_quarters_used in range(1, N_QUARTERS + 1):
        start_row = latest_row - num_quarters_used + 1
        if start_row < cache.base_row:
            break

        penetration_values: list[float] = []
        for row in range(start_row, latest_row + 1):
            value = to_float(cache.get(row, penetration_col))
            if value is None:
                penetration_values = []
                break
            penetration_values.append(value)
        if len(penetration_values) != num_quarters_used:
            continue

        avg_formula_cell.formula2 = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{latest_row}C{penetration_col})"
        )
        wb.app.calculate()
        avg_penetration_pct = to_float(avg_formula_cell.value)
        if avg_penetration_pct is None:
            avg_penetration_pct = sum(penetration_values) / len(penetration_values)

        quarterly_sales = to_float(cache.get(latest_row, quarterly_sales_col))
        reported_sales = to_float(cache.get(latest_row, reported_sales_col))
        growth_rate_pct = to_float(cache.get(latest_row, growth_rate_col))
        sales_captured_in_db_pct = to_float(cache.get(latest_row, captured_in_db_col))

        forecast_value = multiply(avg_penetration_pct, quarterly_sales)
        forecast_max = multiply(max(penetration_values), quarterly_sales)
        forecast_min = multiply(min(penetration_values), quarterly_sales)
        range_width = subtract(forecast_max, forecast_min)
        last_quarter_used = cache.get(start_row, quarter_col)

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


def extract_regression_candidates(
    wb: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if REGRESSION_SHEET_NAME not in [sheet.name for sheet in wb.sheets]:
        return rows

    sheet = wb.sheets[REGRESSION_SHEET_NAME]
    cache = SheetCache.from_sheet(sheet)
    anchor = cache.find_anchor("max")
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor
    latest_row = anchor_row - 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    actual_col = cache.nearest_header_column(anchor_row, anchor_col, ("actual",))

    intercept_cell = sheet.cells(anchor_row + 2, anchor_col + 2)
    slope_cell = sheet.cells(anchor_row + 3, anchor_col + 2)

    previous_calc_key: tuple[Any, ...] | None = None

    for num_quarters_used in range(2, N_QUARTERS + 1):
        start_row = latest_row - num_quarters_used + 1
        if start_row < cache.base_row:
            break

        x_values: list[float] = []
        y_values: list[float] = []
        valid_window = True
        for row in range(start_row, latest_row + 1):
            x_value = to_float(cache.get(row, x_col))
            y_value = to_float(cache.get(row, y_col))
            if x_value is None or y_value is None:
                valid_window = False
                break
            x_values.append(x_value)
            y_values.append(y_value)
        if not valid_window:
            continue

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        forecast_total_without_sa = intercept + slope * x_values[-1]
        forecast_max = to_float(cache.get(anchor_row, anchor_col + 1))
        forecast_min = to_float(cache.get(anchor_row + 1, anchor_col + 1))

        if forecast_max is None:
            forecast_max = forecast_total_without_sa
        if forecast_min is None:
            forecast_min = forecast_total_without_sa

        actual_value = None
        if actual_col is not None:
            actual_value = to_float(cache.get(latest_row, actual_col))

        calc_key = model_key(
            float(num_quarters_used),
            intercept,
            slope,
            forecast_total_without_sa,
            forecast_max,
            forecast_min,
        )
        if calc_key == previous_calc_key:
            continue
        previous_calc_key = calc_key

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
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_empirical = wb.create_sheet("empirical_candidates")
    ws_empirical.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_empirical.append([row.get(column_name) for column_name in EMPIRICAL_COLUMNS])
    style_output_sheet(ws_empirical)

    ws_regression = wb.create_sheet("regression_candidates")
    ws_regression.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_regression.append([row.get(column_name) for column_name in REGRESSION_COLUMNS])
    style_output_sheet(ws_regression)

    wb.save(output_path)


def style_output_sheet(ws: Any) -> None:
    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def collect_input_files(input_path: Path) -> list[Path]:
    files: list[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped: {file_path.name} (temporary workbook)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {file_path.name} (not .xlsx)")
            continue
        if "_param" in file_path.stem.lower():
            print(f"skipped: {file_path.name} (looks like prior output)")
            continue
        files.append(file_path)
    return files


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    source_files = collect_input_files(input_path)
    output_file = next_output_path(input_path, output_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files:
            metadata = parse_file_label(file_path)
            if metadata is None:
                print(f"skipped: {file_path.name} (filename format did not match expected pattern)")
                continue

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_candidates(wb, metadata, file_path.name))
                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing error: {exc})")
            finally:
                close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output_path: {output_file}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
