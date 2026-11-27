#!/usr/bin/env python3
"""Extract empirical/regression parameter candidates from .xlsx model files.

Requirements covered:
- Process all .xlsx files in input_dir once each (skip temp files).
- Open each workbook one time, process both sheets, close without saving.
- Use anchor-based extraction from `max` and R1C1 formulas via .formula2.
- Write one output workbook with empirical_candidates and regression_candidates.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------
# User-configurable paths
# -----------------------
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

FILE_LABEL_RE = re.compile(
    r"\s-\s(?P<ticker>[A-Za-z0-9]+)\s-\s(?P<period>(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4}))",
    re.IGNORECASE,
)


def log(message: str) -> None:
    print(message, flush=True)


def to_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
        try:
            return float(text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def pct_normalize(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if value > 1.0 and value <= 100.0:
        return value / 100.0
    return value


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def next_output_path(input_folder: Path, out_dir: Path) -> Path:
    base_name = f"{input_folder.resolve().name}_PARAM"
    first = out_dir / f"{base_name}.xlsx"
    if not first.exists():
        return first

    index = 1
    while True:
        candidate = out_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def parse_file_labels(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    match = FILE_LABEL_RE.search(stem)
    if not match:
        return {
            "ticker": "UNKNOWN",
            "model_period": "UNKNOWN",
            "model_date": "",
            "model": stem,
        }

    ticker = match.group("ticker").upper()
    bucket_raw = match.group("bucket")
    month_raw = match.group("month")
    year_raw = match.group("year")

    bucket_norm = bucket_raw.capitalize()
    month_norm = month_raw.capitalize()
    model_period = f"{bucket_norm}{month_norm}_{year_raw}"

    month_num = MONTH_MAP.get(month_raw.lower())
    day_num = PERIOD_DAY_MAP.get(bucket_raw.lower())
    if month_num is None or day_num is None:
        model_date = ""
    else:
        model_date = date(int(year_raw), month_num, day_num).isoformat()

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def safe_close_workbook(wb: Any) -> None:
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
        # Last-resort fallback; workbook may already be closed.
        pass


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback for older Excel builds.
        cell.formula = formula


def find_anchor(sheet: Any, target: str = "max") -> Optional[tuple[int, int, int, int]]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    values = to_2d(used.value)
    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            if str(value or "").strip().lower() == target:
                return start_row + r_idx, start_col + c_idx, start_col, start_col + len(row_values) - 1
    return None


def headers_from_anchor_row(sheet: Any, row: int, start_col: int, end_col: int) -> list[tuple[str, int]]:
    row_values = sheet.range((row, start_col), (row, end_col)).value
    row_values = row_values if isinstance(row_values, list) else [row_values]
    headers: list[tuple[str, int]] = []
    for idx, raw in enumerate(row_values):
        header = normalize_header(raw)
        if header:
            headers.append((header, start_col + idx))
    return headers


def detect_column(headers: list[tuple[str, int]], patterns: Iterable[tuple[str, ...]]) -> Optional[int]:
    for pattern in patterns:
        for header, col in headers:
            if all(token in header for token in pattern):
                return col
    return None


def collect_numeric_rows(
    sheet: Any, col: int, end_row: int, max_lookback: int = 240
) -> list[tuple[int, float]]:
    if col < 1 or end_row < 1:
        return []

    start_row = max(1, end_row - max_lookback + 1)
    values = sheet.range((start_row, col), (end_row, col)).value
    if not isinstance(values, list):
        values = [values]

    rows: list[tuple[int, float]] = []
    for idx, raw in enumerate(values):
        number = to_float(raw)
        if number is not None:
            rows.append((start_row + idx, number))
    return rows


def collect_xy_rows(
    sheet: Any, x_col: int, y_col: int, end_row: int, max_lookback: int = 240
) -> list[tuple[int, float, float]]:
    if min(x_col, y_col) < 1 or end_row < 1:
        return []

    start_row = max(1, end_row - max_lookback + 1)
    x_values = sheet.range((start_row, x_col), (end_row, x_col)).value
    y_values = sheet.range((start_row, y_col), (end_row, y_col)).value

    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]

    rows: list[tuple[int, float, float]] = []
    for idx, (x_raw, y_raw) in enumerate(zip(x_values, y_values)):
        x_val = to_float(x_raw)
        y_val = to_float(y_raw)
        if x_val is not None and y_val is not None:
            rows.append((start_row + idx, x_val, y_val))
    return rows


def process_empirical_sheet(
    wb: Any, model_meta: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet_name = "Empirical Model"
    try:
        ws = wb.sheets[sheet_name]
    except Exception:
        log(f"  skipped empirical extraction ({sheet_name} missing)")
        return rows

    anchor = find_anchor(ws, "max")
    if not anchor:
        log("  skipped empirical extraction (max anchor not found)")
        return rows

    anchor_row, anchor_col, used_start_col, used_end_col = anchor
    headers = headers_from_anchor_row(ws, anchor_row, used_start_col, used_end_col)

    num_quarters_col = detect_column(headers, [("num", "quarter"), ("quarters", "used")]) or (anchor_col - 6)
    last_quarter_col = detect_column(headers, [("last", "quarter")]) or (anchor_col - 5)
    quarterly_sales_col = detect_column(headers, [("quarterly", "sales"), ("db", "sales")]) or (anchor_col - 4)
    reported_sales_col = detect_column(headers, [("reported", "sales"), ("actual", "sales")]) or (anchor_col - 3)
    growth_col = detect_column(headers, [("growth",), ("growth", "rate")]) or (anchor_col - 2)
    captured_col = detect_column(
        headers, [("captured", "db"), ("sales", "captured"), ("penetration",)]
    ) or (anchor_col - 1)

    max_col = anchor_col
    min_col = detect_column(headers, [("min",)]) or (anchor_col + 1)
    actual_col = detect_column(headers, [("actual",), ("reported", "sales")]) or reported_sales_col
    forecast_col = detect_column(headers, [("forecast",), ("estimated", "total", "sold")]) or (anchor_col - 2)

    penetration_col = captured_col
    penetration_points = collect_numeric_rows(ws, penetration_col, anchor_row - 1)
    penetration_points = [(r, pct_normalize(v) or 0.0) for r, v in penetration_points if (pct_normalize(v) or 0) > 0]

    if not penetration_points:
        log("  skipped empirical extraction (no usable penetration history)")
        return rows

    history_rows = [r for r, _ in penetration_points]
    latest_history_row = history_rows[-1]

    latest_quarterly_sales = to_float(ws.range((latest_history_row, quarterly_sales_col)).value)
    latest_reported_sales = to_float(ws.range((latest_history_row, reported_sales_col)).value)
    latest_growth = pct_normalize(to_float(ws.range((latest_history_row, growth_col)).value))
    latest_captured = pct_normalize(to_float(ws.range((latest_history_row, captured_col)).value))
    latest_quarter_label = ws.range((latest_history_row, last_quarter_col)).value

    temp_col = used_end_col + 2
    avg_cell = ws.range((anchor_row, temp_col))
    max_pen_cell = ws.range((anchor_row, temp_col + 1))
    min_pen_cell = ws.range((anchor_row, temp_col + 2))

    n_quarters = 10
    max_available = min(n_quarters, len(history_rows))
    for n in range(1, max_available + 1):
        start_row = history_rows[-n]
        end_row = history_rows[-1]

        set_formula2(avg_cell, f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})")
        set_formula2(max_pen_cell, f"=MAX(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})")
        set_formula2(min_pen_cell, f"=MIN(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})")
        wb.app.calculate()

        avg_pen = pct_normalize(to_float(avg_cell.value))
        max_pen = pct_normalize(to_float(max_pen_cell.value))
        min_pen = pct_normalize(to_float(min_pen_cell.value))

        forecast_value = safe_div(latest_quarterly_sales, avg_pen)
        forecast_max = safe_div(latest_quarterly_sales, min_pen)
        forecast_min = safe_div(latest_quarterly_sales, max_pen)
        range_width = (
            (forecast_max - forecast_min) if forecast_max is not None and forecast_min is not None else None
        )

        table_row = anchor_row + n
        table_actual = to_float(ws.range((table_row, actual_col)).value)
        table_forecast = to_float(ws.range((table_row, forecast_col)).value)
        table_max = to_float(ws.range((table_row, max_col)).value)
        table_min = to_float(ws.range((table_row, min_col)).value)

        rows.append(
            {
                "model": model_meta["model"],
                "ticker": model_meta["ticker"],
                "model_period": model_meta["model_period"],
                "model_date": model_meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": int(n),
                "last_quarter_used": latest_quarter_label,
                "forecast_value": forecast_value if forecast_value is not None else table_forecast,
                "actual_value": latest_reported_sales if latest_reported_sales is not None else table_actual,
                "forecast_max": forecast_max if forecast_max is not None else table_max,
                "forecast_min": forecast_min if forecast_min is not None else table_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": latest_quarterly_sales,
                "reported_sales": latest_reported_sales,
                "growth_rate_pct": latest_growth,
                "sales_captured_in_db_pct": latest_captured,
                "source_file": source_file,
            }
        )

    ws.range((anchor_row, temp_col), (anchor_row, temp_col + 2)).value = None
    return rows


def process_regression_sheet(
    wb: Any, model_meta: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet_name = "Regression Model"
    try:
        ws = wb.sheets[sheet_name]
    except Exception:
        log(f"  skipped regression extraction ({sheet_name} missing)")
        return rows

    anchor = find_anchor(ws, "max")
    if not anchor:
        log("  skipped regression extraction (max anchor not found)")
        return rows

    anchor_row, anchor_col, used_start_col, used_end_col = anchor
    headers = headers_from_anchor_row(ws, anchor_row, used_start_col, used_end_col)

    # Required anchor offsets from the prompt.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    actual_col = detect_column(headers, [("actual",), ("reported", "sales")]) or -1

    xy_rows = collect_xy_rows(ws, x_col, y_col, anchor_row - 1)
    if len(xy_rows) < 2:
        log("  skipped regression extraction (not enough x/y history)")
        return rows

    latest_x = xy_rows[-1][1]

    temp_col = used_end_col + 2
    intercept_cell = ws.range((anchor_row, temp_col))
    slope_cell = ws.range((anchor_row, temp_col + 1))
    max_cell = ws.range((anchor_row, temp_col + 2))
    min_cell = ws.range((anchor_row, temp_col + 3))

    prev_signature: Optional[tuple[Any, ...]] = None
    max_quarters = min(10, len(xy_rows))
    for n in range(2, max_quarters + 1):
        start_row = xy_rows[-n][0]
        end_row = xy_rows[-1][0]

        set_formula2(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(max_cell, f"=MAX(R{start_row}C{y_col}:R{end_row}C{y_col})")
        set_formula2(min_cell, f"=MIN(R{start_row}C{y_col}:R{end_row}C{y_col})")
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)

        forecast_total_without_sa = (
            (intercept + slope * latest_x) if intercept is not None and slope is not None else None
        )
        range_width = (
            (forecast_max - forecast_min) if forecast_max is not None and forecast_min is not None else None
        )

        actual_value = None
        if actual_col >= 1:
            actual_value = to_float(ws.range((anchor_row + n, actual_col)).value)

        signature = (
            round(forecast_total_without_sa, 10) if forecast_total_without_sa is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": model_meta["model"],
                "ticker": model_meta["ticker"],
                "model_period": model_meta["model_period"],
                "model_date": model_meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": int(n),
                "num_quarters_used": int(n),
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    ws.range((anchor_row, temp_col), (anchor_row, temp_col + 3)).value = None
    return rows


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def write_sheet(
    wb: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    # Formatting requirements.
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[col_letter]:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 60)


def list_source_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            log(f"skipped: {path.name} (not a file)")
            continue
        if path.name.startswith("~"):
            log(f"skipped: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            log(f"skipped: {path.name} (not .xlsx)")
            continue
        files.append(path)
    return files


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    source_files = list_source_files(input_dir)
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in source_files:
            log(f"processing: {file_path.name}")
            wb = None
            try:
                model_meta = parse_file_labels(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(process_empirical_sheet(wb, model_meta, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, model_meta, file_path.name))

                processed_count += 1
            except Exception as exc:
                log(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    log(f"output_path: {output_path}")
    log(f"files_processed: {processed_count}")
    log(f"empirical_rows: {len(empirical_rows)}")
    log(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
