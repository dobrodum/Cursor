#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import math
import re
import statistics
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = r"./input"
output_dir = r"./output"


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        if text.endswith("%"):
            return float(text[:-1]) / 100.0
        return float(text)
    except ValueError:
        return None


def normalize_pct(value: float | None) -> float | None:
    if value is None:
        return None
    if abs(value) > 1 and abs(value) <= 100:
        return value / 100.0
    return value


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], (list, tuple)):
        matrix: list[list[Any]] = []
        for row in values:
            if isinstance(row, tuple):
                matrix.append(list(row))
            elif isinstance(row, list):
                matrix.append(row)
            else:
                matrix.append([row])
        return matrix
    return [values]


@dataclass
class SheetCache:
    sheet: xw.Sheet
    top_row: int
    left_col: int
    values: list[list[Any]]

    @classmethod
    def from_sheet(cls, sheet: xw.Sheet) -> "SheetCache":
        used = sheet.used_range
        return cls(sheet=sheet, top_row=used.row, left_col=used.column, values=to_2d(used.value))

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        return len(self.values[0]) if self.values else 0

    @property
    def bottom_row(self) -> int:
        return self.top_row + max(0, self.row_count - 1)

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.top_row
        c_idx = col - self.left_col
        if r_idx < 0 or c_idx < 0:
            return None
        if r_idx >= self.row_count or c_idx >= self.col_count:
            return None
        return self.values[r_idx][c_idx]

    def row_values(self, row: int) -> list[Any]:
        r_idx = row - self.top_row
        if r_idx < 0 or r_idx >= self.row_count:
            return []
        return self.values[r_idx]


def find_sheet_case_insensitive(wb: xw.Book, expected_name: str) -> xw.Sheet | None:
    expected = expected_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == expected:
            return sheet
    return None


def find_anchor_max(cache: SheetCache) -> tuple[int, int] | None:
    target = "max"
    for r_idx, row_values in enumerate(cache.values):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                return cache.top_row + r_idx, cache.left_col + c_idx
    return None


def find_column(
    cache: SheetCache,
    search_rows: list[int],
    aliases: list[str],
    anchor_col: int,
    default_offset: int | None = None,
) -> int | None:
    alias_norm = [normalize_text(a) for a in aliases]
    for row in search_rows:
        row_values = cache.row_values(row)
        if not row_values:
            continue
        for c_idx, value in enumerate(row_values):
            text = normalize_text(value)
            if not text:
                continue
            if any(alias in text for alias in alias_norm):
                return cache.left_col + c_idx

    if default_offset is None:
        return None
    return anchor_col + default_offset


def collect_numeric_rows(cache: SheetCache, col: int | None, start_row: int) -> list[int]:
    if col is None:
        return []
    rows: list[int] = []
    for row in range(start_row, cache.bottom_row + 1):
        if to_float(cache.get(row, col)) is not None:
            rows.append(row)
    return rows


def latest_numeric(cache: SheetCache, col: int | None, rows: list[int]) -> float | None:
    if col is None:
        return None
    for row in reversed(rows):
        numeric = to_float(cache.get(row, col))
        if numeric is not None:
            return numeric
    return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    cell.formula = formula_r1c1


def safe_close_workbook(wb: xw.Book) -> None:
    close_attempts = [
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.close(),
    ]

    last_error: Exception | None = None
    for attempt in close_attempts:
        try:
            attempt()
            return
        except Exception as exc:
            last_error = exc

    if last_error is not None:
        raise last_error


def parse_file_metadata(file_name: str) -> dict[str, str] | None:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker: str | None = None
    period_token: str | None = None

    if len(parts) >= 3:
        ticker = parts[1]
        period_token = parts[2].split("_")[0]
    else:
        match = re.search(r"-\s*([A-Za-z0-9]+)\s*-\s*([A-Za-z]+[A-Za-z]+\d{4})", stem)
        if match:
            ticker = match.group(1)
            period_token = match.group(2)

    if not ticker or not period_token:
        return None

    period_match = re.search(r"(?i)(early|mid|late)\s*([A-Za-z]+)\s*(\d{4})", period_token)
    if not period_match:
        return None

    phase_raw = period_match.group(1).lower()
    month_raw = period_match.group(2).lower()
    year = int(period_match.group(3))

    month_num = MONTH_MAP.get(month_raw) or MONTH_MAP.get(month_raw[:3])
    if month_num is None:
        return None

    day = PHASE_DAY[phase_raw]
    model_date = date(year, month_num, day).isoformat()
    phase_title = phase_raw.capitalize()
    month_abbrev = date(year, month_num, 1).strftime("%b")
    model_period = f"{phase_title}{month_abbrev}_{year}"
    model_name = f"{ticker}_{model_period}"

    return {
        "model": model_name,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    input_folder_name = in_dir.name
    base = out_dir / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    idx = 1
    while True:
        candidate = out_dir / f"{input_folder_name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def almost_equal(a: float | None, b: float | None, tol: float = 1e-9) -> bool:
    if a is None or b is None:
        return a is b
    scale = max(1.0, abs(a), abs(b))
    return abs(a - b) <= tol * scale


def is_duplicate_regression_row(current: dict[str, Any], previous: dict[str, Any]) -> bool:
    fields = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    for field in fields:
        curr = current.get(field)
        prev = previous.get(field)
        if isinstance(curr, (int, float)) or isinstance(prev, (int, float)):
            if not almost_equal(to_float(curr), to_float(prev)):
                return False
        else:
            if curr != prev:
                return False
    return True


def extract_empirical_candidates(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        print(f"Skipping empirical extraction for {source_file}: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    cache = SheetCache.from_sheet(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"Skipping empirical extraction for {source_file}: anchor cell 'max' not found")
        return []

    anchor_row, anchor_col = anchor
    header_rows = [anchor_row - 1, anchor_row, anchor_row + 1]

    col_penetration = find_column(
        cache,
        header_rows,
        aliases=["penetration", "pen %", "avg penetration"],
        anchor_col=anchor_col,
        default_offset=-8,
    )
    col_quarterly_sales = find_column(
        cache,
        header_rows,
        aliases=["quarterly sales", "qtr sales", "quarter sales"],
        anchor_col=anchor_col,
        default_offset=-6,
    )
    col_reported_sales = find_column(
        cache,
        header_rows,
        aliases=["reported sales", "actual sales", "reported"],
        anchor_col=anchor_col,
        default_offset=-5,
    )
    col_growth_rate = find_column(
        cache,
        header_rows,
        aliases=["growth rate", "growth %"],
        anchor_col=anchor_col,
        default_offset=-4,
    )
    col_sales_captured = find_column(
        cache,
        header_rows,
        aliases=["captured in db", "sales captured"],
        anchor_col=anchor_col,
        default_offset=-3,
    )
    col_last_quarter = find_column(
        cache,
        header_rows,
        aliases=["last quarter", "last qtr", "quarter"],
        anchor_col=anchor_col,
        default_offset=-10,
    )
    col_estimated_total = find_column(
        cache,
        header_rows,
        aliases=["estimated total sold", "tot fcst", "forecast value"],
        anchor_col=anchor_col,
        default_offset=-1,
    )
    col_min = find_column(
        cache,
        header_rows,
        aliases=["min"],
        anchor_col=anchor_col,
        default_offset=1,
    )

    data_rows = collect_numeric_rows(cache, col_penetration, anchor_row + 1)
    if not data_rows:
        print(f"Skipping empirical extraction for {source_file}: no penetration history rows found")
        return []

    data_rows = data_rows[-N_QUARTERS:]
    max_n = min(N_QUARTERS, len(data_rows))
    if max_n == 0:
        return []

    avg_cell = sheet.range((anchor_row, anchor_col + 20))
    min_cell = sheet.range((anchor_row, anchor_col + 21))
    max_cell = sheet.range((anchor_row, anchor_col + 22))

    rows: list[dict[str, Any]] = []
    for num_quarters in range(1, max_n + 1):
        window_rows = data_rows[-num_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        set_formula2_r1c1(avg_cell, f"=AVERAGE(R{start_row}C{col_penetration}:R{end_row}C{col_penetration})")
        set_formula2_r1c1(min_cell, f"=MIN(R{start_row}C{col_penetration}:R{end_row}C{col_penetration})")
        set_formula2_r1c1(max_cell, f"=MAX(R{start_row}C{col_penetration}:R{end_row}C{col_penetration})")
        wb.app.calculate()

        avg_penetration_pct = normalize_pct(to_float(avg_cell.value))
        min_penetration = normalize_pct(to_float(min_cell.value))
        max_penetration = normalize_pct(to_float(max_cell.value))

        quarterly_sales = latest_numeric(cache, col_quarterly_sales, window_rows)
        reported_sales = latest_numeric(cache, col_reported_sales, window_rows)
        growth_rate_pct = normalize_pct(latest_numeric(cache, col_growth_rate, window_rows))
        sales_captured_in_db_pct = normalize_pct(latest_numeric(cache, col_sales_captured, window_rows))
        forecast_value = latest_numeric(cache, col_estimated_total, window_rows)

        if forecast_value is None and quarterly_sales is not None and avg_penetration_pct:
            if avg_penetration_pct != 0:
                forecast_value = quarterly_sales / avg_penetration_pct

        forecast_max = latest_numeric(cache, anchor_col, window_rows)
        forecast_min = latest_numeric(cache, col_min, window_rows)

        if forecast_max is None and quarterly_sales is not None and min_penetration and min_penetration != 0:
            forecast_max = quarterly_sales / min_penetration
        if forecast_min is None and quarterly_sales is not None and max_penetration and max_penetration != 0:
            forecast_min = quarterly_sales / max_penetration

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        last_quarter_used = cache.get(window_rows[0], col_last_quarter) if col_last_quarter is not None else None

        rows.append(
            {
                **metadata,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
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
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, REGRESSION_SHEET_NAME)
    if sheet is None:
        print(f"Skipping regression extraction for {source_file}: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    cache = SheetCache.from_sheet(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"Skipping regression extraction for {source_file}: anchor cell 'max' not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    header_rows = [anchor_row - 1, anchor_row, anchor_row + 1]

    col_forecast_total_wo_sa = find_column(
        cache,
        header_rows,
        aliases=["tot fcst w/o sa", "total fcst w/o sa", "tot fcst without sa", "forecast w/o sa"],
        anchor_col=anchor_col,
        default_offset=-1,
    )
    col_actual_value = find_column(
        cache,
        header_rows,
        aliases=["actual value", "actual sales", "actual"],
        anchor_col=anchor_col,
        default_offset=None,
    )
    col_min = find_column(
        cache,
        header_rows,
        aliases=["min"],
        anchor_col=anchor_col,
        default_offset=1,
    )

    data_rows: list[int] = []
    for row in range(anchor_row + 1, cache.bottom_row + 1):
        if to_float(cache.get(row, x_col)) is not None and to_float(cache.get(row, y_col)) is not None:
            data_rows.append(row)

    if len(data_rows) < 2:
        print(f"Skipping regression extraction for {source_file}: not enough x/y rows found")
        return []

    max_n = min(N_QUARTERS, len(data_rows))
    if max_n < 2:
        return []

    intercept_cell = sheet.range((anchor_row, anchor_col + 20))
    slope_cell = sheet.range((anchor_row, anchor_col + 21))
    forecast_cell = sheet.range((anchor_row, anchor_col + 22))

    rows: list[dict[str, Any]] = []
    for num_quarters in range(2, max_n + 1):
        window_rows = data_rows[-num_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]
        next_x_row = end_row + 1 if to_float(cache.get(end_row + 1, x_col)) is not None else end_row

        set_formula2_r1c1(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            forecast_cell,
            f"=R{anchor_row}C{anchor_col + 20}+R{anchor_row}C{anchor_col + 21}*R{next_x_row}C{x_col}",
        )
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = to_float(cache.get(end_row, col_forecast_total_wo_sa))
        if forecast_value is None:
            forecast_value = to_float(forecast_cell.value)

        actual_value = to_float(cache.get(end_row, col_actual_value)) if col_actual_value is not None else None
        forecast_max = to_float(cache.get(end_row, anchor_col))
        forecast_min = to_float(cache.get(end_row, col_min))

        if forecast_value is not None and (forecast_max is None or forecast_min is None):
            x_vals = [to_float(cache.get(r, x_col)) for r in window_rows]
            y_vals = [to_float(cache.get(r, y_col)) for r in window_rows]
            if intercept is not None and slope is not None:
                residuals = [
                    y - (intercept + slope * x)
                    for x, y in zip(x_vals, y_vals)
                    if x is not None and y is not None
                ]
                sigma = statistics.pstdev(residuals) if len(residuals) > 1 else 0.0
            else:
                sigma = 0.0

            if forecast_max is None:
                forecast_max = forecast_value + (2.0 * sigma)
            if forecast_min is None:
                forecast_min = forecast_value - (2.0 * sigma)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row = {
            **metadata,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(row, rows[-1]):
            continue
        rows.append(row)

    return rows


def write_rows_with_formatting(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            max_len = max(max_len, len(text))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(50, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    write_rows_with_formatting(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    write_rows_with_formatting(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    out_file = next_output_path(input_path, output_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app: xw.App | None = None
    original_calculation = None

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            original_calculation = app.calculation
            app.calculation = "manual"
        except Exception:
            original_calculation = None

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            metadata = parse_file_metadata(file_path.name)
            if metadata is None:
                print(f"Skipped {file_path.name}: filename does not match expected pattern")
                continue

            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_candidates(wb, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing failed ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            try:
                if original_calculation is not None:
                    app.calculation = original_calculation
            except Exception:
                pass
            app.quit()

    write_output_workbook(out_file, empirical_rows, regression_rows)

    print(f"Output path: {out_file}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
