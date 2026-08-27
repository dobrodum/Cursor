from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
MONTH_MAP = {abbr.lower(): idx for idx, abbr in enumerate(calendar.month_abbr) if abbr}

# Empirical fallback offsets from "max" anchor, used when header lookup misses.
EMPIRICAL_OFFSETS = {
    "quarter": -12,
    "captured_pct": -9,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate": -6,
}


@dataclass
class SheetSnapshot:
    top_row: int
    left_col: int
    values: list[list[Any]]

    @property
    def bottom_row(self) -> int:
        return self.top_row + len(self.values) - 1

    @property
    def right_col(self) -> int:
        return self.left_col + len(self.values[0]) - 1 if self.values else self.left_col


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [[item] for item in values]


def build_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = ensure_2d(used.value)
    return SheetSnapshot(top_row=used.row, left_col=used.column, values=values)


def find_anchor_max(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    for r_offset, row in enumerate(snapshot.values):
        for c_offset, cell in enumerate(row):
            if normalize_text(cell) == "max":
                return snapshot.top_row + r_offset, snapshot.left_col + c_offset
    return None


def find_column_by_keywords(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    keyword_groups: Iterable[tuple[str, ...]],
    row_window: int = 8,
    col_window: int = 24,
) -> int | None:
    min_row = max(snapshot.top_row, anchor_row - row_window)
    max_row = min(snapshot.bottom_row, anchor_row + row_window)
    min_col = max(snapshot.left_col, anchor_col - col_window)
    max_col = min(snapshot.right_col, anchor_col + 4)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            r_idx = row - snapshot.top_row
            c_idx = col - snapshot.left_col
            token = normalize_text(snapshot.values[r_idx][c_idx])
            if not token:
                continue
            for group in keyword_groups:
                if all(word in token for word in group):
                    return col
    return None


def read_column_segment(
    sheet: xw.Sheet,
    col: int,
    start_row: int,
    end_row: int,
) -> list[Any]:
    if end_row < start_row:
        return []
    values = sheet.range((start_row, col), (end_row, col)).value
    if isinstance(values, list):
        return [v[0] if isinstance(v, list) else v for v in values]
    return [values]


def collect_numeric_rows(
    sheet: xw.Sheet,
    col: int,
    start_row: int,
    end_row: int,
) -> list[int]:
    rows: list[int] = []
    values = read_column_segment(sheet, col, start_row, end_row)
    for i, value in enumerate(values, start=start_row):
        if to_float(value) is not None:
            rows.append(i)
    return rows


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Compatibility fallback for older Excel engines.
        cell.formula = formula_r1c1


def parse_model_metadata(file_name: str) -> dict[str, str] | None:
    stem = Path(file_name).stem
    pattern = re.compile(r"-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)([A-Za-z]+)(\d{4})", re.IGNORECASE)
    match = pattern.search(stem)
    if not match:
        return None

    ticker = match.group(1).upper()
    period_word = match.group(2).capitalize()
    month_token = match.group(3)[:3].lower()
    year = int(match.group(4))
    month = MONTH_MAP.get(month_token)
    if month is None:
        return None

    day = DAY_BY_PERIOD[period_word.lower()]
    model_date = date(year, month, day).isoformat()
    month_abbr = calendar.month_abbr[month]
    model_period = f"{period_word}{month_abbr}_{year}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def find_output_path(base_name: str, folder: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    base_path = folder / f"{base_name}_PARAM.xlsx"
    if not base_path.exists():
        return base_path
    counter = 1
    while True:
        candidate = folder / f"{base_name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


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
        wb.api.Close(False)


def as_ratio(value: float | None) -> float | None:
    if value is None:
        return None
    if value > 1.0:
        return value / 100.0
    return value


def calc_range_width(v_max: float | None, v_min: float | None) -> float | None:
    if v_max is None or v_min is None:
        return None
    return v_max - v_min


def maybe_round(value: float | None, digits: int = 8) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def process_empirical_sheet(
    wb: xw.Book,
    meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    snapshot = build_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    quarter_col = (
        find_column_by_keywords(snapshot, anchor_row, anchor_col, [("quarter",), ("period",)])
        or anchor_col + EMPIRICAL_OFFSETS["quarter"]
    )
    captured_col = (
        find_column_by_keywords(
            snapshot,
            anchor_row,
            anchor_col,
            [("captured",), ("penetration",), ("db",)],
        )
        or anchor_col + EMPIRICAL_OFFSETS["captured_pct"]
    )
    quarterly_sales_col = (
        find_column_by_keywords(snapshot, anchor_row, anchor_col, [("quarterly", "sales"), ("sales", "db")])
        or anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]
    )
    reported_sales_col = (
        find_column_by_keywords(snapshot, anchor_row, anchor_col, [("reported", "sales"), ("actual", "sales")])
        or anchor_col + EMPIRICAL_OFFSETS["reported_sales"]
    )
    growth_rate_col = (
        find_column_by_keywords(snapshot, anchor_row, anchor_col, [("growth",), ("growth", "rate")])
        or anchor_col + EMPIRICAL_OFFSETS["growth_rate"]
    )

    lookback_start = max(1, anchor_row - 240)
    numeric_rows = collect_numeric_rows(sheet, captured_col, lookback_start, anchor_row - 1)
    if not numeric_rows:
        return []

    numeric_rows = numeric_rows[-10:]
    max_n = len(numeric_rows)

    temp_avg_cell = sheet.cells(anchor_row + 2, anchor_col + 2)
    rows: list[dict[str, Any]] = []
    captured_values_cache = {
        r: as_ratio(to_float(sheet.cells(r, captured_col).value)) for r in numeric_rows
    }

    for n_quarters in range(1, max_n + 1):
        start_row = numeric_rows[-n_quarters]
        end_row = numeric_rows[-1]
        set_formula2(
            temp_avg_cell,
            f"=AVERAGE(R{start_row}C{captured_col}:R{end_row}C{captured_col})",
        )
        wb.app.calculate()

        avg_penetration_raw = to_float(temp_avg_cell.value)
        avg_penetration_pct = as_ratio(avg_penetration_raw)
        if avg_penetration_pct is None or avg_penetration_pct == 0:
            continue

        quarterly_sales = to_float(sheet.cells(end_row, quarterly_sales_col).value)
        reported_sales = to_float(sheet.cells(end_row, reported_sales_col).value)
        growth_rate_pct = as_ratio(to_float(sheet.cells(end_row, growth_rate_col).value))
        sales_captured_in_db_pct = captured_values_cache.get(end_row)
        quarter_label = sheet.cells(end_row, quarter_col).value

        forecast_value = None
        if quarterly_sales is not None:
            forecast_value = quarterly_sales / avg_penetration_pct

        window_values = [captured_values_cache[r] for r in numeric_rows[-n_quarters:] if captured_values_cache[r] is not None]
        stdev = None
        if len(window_values) >= 2:
            mean = sum(window_values) / len(window_values)
            variance = sum((val - mean) ** 2 for val in window_values) / (len(window_values) - 1)
            stdev = math.sqrt(variance)

        forecast_max = None
        forecast_min = None
        if quarterly_sales is not None and avg_penetration_pct is not None:
            hi_pen = avg_penetration_pct - (stdev or 0.0)
            lo_pen = avg_penetration_pct + (stdev or 0.0)
            if hi_pen and hi_pen > 0:
                forecast_max = quarterly_sales / hi_pen
            if lo_pen and lo_pen > 0:
                forecast_min = quarterly_sales / lo_pen

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": maybe_round(avg_penetration_pct),
            "num_quarters_used": n_quarters,
            "last_quarter_used": quarter_label if quarter_label is not None else f"row_{end_row}",
            "forecast_value": maybe_round(forecast_value),
            "actual_value": maybe_round(reported_sales),
            "forecast_max": maybe_round(forecast_max),
            "forecast_min": maybe_round(forecast_min),
            "range_width": maybe_round(calc_range_width(forecast_max, forecast_min)),
            "avg_penetration_pct": maybe_round(avg_penetration_pct),
            "quarterly_sales": maybe_round(quarterly_sales),
            "reported_sales": maybe_round(reported_sales),
            "growth_rate_pct": maybe_round(growth_rate_pct),
            "sales_captured_in_db_pct": maybe_round(sales_captured_in_db_pct),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def approx_equal(a: float | None, b: float | None, tol: float = 1e-10) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def process_regression_sheet(
    wb: xw.Book,
    meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    snapshot = build_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    lookback_start = max(1, anchor_row - 240)
    xy_matrix = ensure_2d(sheet.range((lookback_start, x_col), (anchor_row - 1, y_col)).value)
    valid_rows: list[int] = []
    for idx, row_vals in enumerate(xy_matrix, start=lookback_start):
        x_val = to_float(row_vals[0]) if len(row_vals) >= 1 else None
        y_idx = y_col - x_col
        y_val = to_float(row_vals[y_idx]) if len(row_vals) > y_idx else None
        if x_val is not None and y_val is not None:
            valid_rows.append(idx)

    if not valid_rows:
        return []
    valid_rows = valid_rows[-10:]

    # Optional actual value, if a label exists in the sheet.
    actual_col = find_column_by_keywords(snapshot, anchor_row, anchor_col, [("actual",), ("reported", "sales")])
    actual_value = to_float(sheet.cells(valid_rows[-1], actual_col).value) if actual_col else None

    temp_intercept = sheet.cells(anchor_row + 2, anchor_col + 2)
    temp_slope = sheet.cells(anchor_row + 2, anchor_col + 3)
    temp_forecast = sheet.cells(anchor_row + 2, anchor_col + 4)

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, len(valid_rows) + 1):
        start_row = valid_rows[-n_quarters]
        end_row = valid_rows[-1]

        set_formula2(
            temp_intercept,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(
            temp_slope,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        next_x = to_float(sheet.cells(end_row, x_col).value)
        if next_x is None:
            continue
        set_formula2(temp_forecast, f"=RC[-2]+RC[-1]*{next_x}")
        wb.app.calculate()

        intercept = to_float(temp_intercept.value)
        slope = to_float(temp_slope.value)
        forecast_total_without_sa = to_float(temp_forecast.value)

        forecast_max = None
        forecast_min = None
        if intercept is not None and slope is not None and forecast_total_without_sa is not None:
            y_vals_raw = read_column_segment(sheet, y_col, start_row, end_row)
            x_vals_raw = read_column_segment(sheet, x_col, start_row, end_row)
            residuals: list[float] = []
            for xv_raw, yv_raw in zip(x_vals_raw, y_vals_raw):
                xv = to_float(xv_raw)
                yv = to_float(yv_raw)
                if xv is None or yv is None:
                    continue
                residuals.append(yv - (intercept + slope * xv))
            if len(residuals) >= 2:
                mean = sum(residuals) / len(residuals)
                variance = sum((r - mean) ** 2 for r in residuals) / (len(residuals) - 1)
                sigma = math.sqrt(variance)
                forecast_max = forecast_total_without_sa + sigma
                forecast_min = forecast_total_without_sa - sigma

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": maybe_round(forecast_total_without_sa),
            "actual_value": maybe_round(actual_value),
            "forecast_max": maybe_round(forecast_max),
            "forecast_min": maybe_round(forecast_min),
            "range_width": maybe_round(calc_range_width(forecast_max, forecast_min)),
            "intercept": maybe_round(intercept),
            "slope": maybe_round(slope),
            "source_file": source_file,
        }

        if rows:
            prev = rows[-1]
            is_duplicate = (
                prev["num_quarters_used"] == row["num_quarters_used"]
                or (
                    approx_equal(prev["intercept"], row["intercept"])
                    and approx_equal(prev["slope"], row["slope"])
                    and approx_equal(prev["forecast_value"], row["forecast_value"])
                    and approx_equal(prev["forecast_max"], row["forecast_max"])
                    and approx_equal(prev["forecast_min"], row["forecast_min"])
                )
            )
            if is_duplicate:
                continue

        rows.append(row)

    return rows


def write_sheet(ws, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for item in rows:
        ws.append([item.get(col, "") for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def is_processable_xlsx(path: Path) -> tuple[bool, str]:
    if path.is_dir():
        return False, "is a directory"
    if path.name.startswith("~"):
        return False, "temporary Excel file"
    if path.suffix.lower() != ".xlsx":
        return False, "not an .xlsx file"
    return True, ""


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir not found: {input_dir}")

    output_path = find_output_path(input_dir.name, output_dir)
    files = sorted(input_dir.iterdir())

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files:
            ok, reason = is_processable_xlsx(file_path)
            if not ok:
                print(f"SKIP: {file_path.name} ({reason})")
                continue

            metadata = parse_model_metadata(file_path.name)
            if metadata is None:
                print(f"SKIP: {file_path.name} (cannot parse ticker/model period from filename)")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_empirical = process_empirical_sheet(wb, metadata, file_path.name)
                file_regression = process_regression_sheet(wb, metadata, file_path.name)
                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                processed_files += 1
                print(
                    f"PROCESSED: {file_path.name} "
                    f"(empirical_rows={len(file_empirical)}, regression_rows={len(file_regression)})"
                )
            except Exception as exc:
                print(f"SKIP: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        app.quit()

    wb_out = Workbook()
    default_sheet = wb_out.active
    wb_out.remove(default_sheet)

    ws_empirical = wb_out.create_sheet("empirical_candidates")
    ws_regression = wb_out.create_sheet("regression_candidates")
    write_sheet(ws_empirical, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_HEADERS, regression_rows)
    wb_out.save(output_path)

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {processed_files}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
