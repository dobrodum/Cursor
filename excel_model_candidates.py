#!/usr/bin/env python3
"""Extract empirical and regression candidates from .xlsx model files.

This script opens each source workbook exactly once with xlwings, processes
both model sheets while the workbook is open, and writes a single output
workbook containing:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import datetime as dt
import math
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ------------------------
# User-editable directories
# ------------------------
input_dir = r"./input"
output_dir = r"./output"

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


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_subtract(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return a - b


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].strip() if len(parts) > 1 and parts[1].strip() else "UNKNOWN"

    period_segment = parts[2] if len(parts) > 2 else stem
    period_segment = re.split(r"_", period_segment)[0].strip()

    match = re.search(
        r"(?i)\b(early|mid|late)\s*[-_ ]*([a-z]{3})\s*[-_ ]*(\d{4})\b",
        period_segment,
    )
    if not match:
        match = re.search(
            r"(?i)\b(early|mid|late)\s*[-_ ]*([a-z]{3})\s*[-_ ]*(\d{4})\b",
            stem,
        )

    if not match:
        model_period = "UnknownPeriod"
        model_date = ""
    else:
        timing, month_abbrev, year = match.groups()
        timing_title = timing.lower().capitalize()
        month_title = month_abbrev.lower().capitalize()
        month_map = {
            "Jan": 1,
            "Feb": 2,
            "Mar": 3,
            "Apr": 4,
            "May": 5,
            "Jun": 6,
            "Jul": 7,
            "Aug": 8,
            "Sep": 9,
            "Oct": 10,
            "Nov": 11,
            "Dec": 12,
        }
        month_num = month_map.get(month_title, 1)
        day_map = {"early": 5, "mid": 15, "late": 25}
        day_num = day_map[timing.lower()]
        model_period = f"{timing_title}{month_title}_{year}"
        model_date = dt.date(int(year), month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.resolve().name
    output_path.mkdir(parents=True, exist_ok=True)

    base = output_path / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    idx = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


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


def build_sheet_cache(
    sheet: xw.Sheet,
) -> Tuple[int, int, List[List[Any]], Dict[str, Tuple[int, int]]]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    values = ensure_2d(used.options(ndim=2).value)

    labels: Dict[str, Tuple[int, int]] = {}
    for r_offset, row_vals in enumerate(values):
        row_number = start_row + r_offset
        for c_offset, cell_value in enumerate(row_vals):
            if not isinstance(cell_value, str):
                continue
            norm = normalize_label(cell_value)
            if not norm:
                continue
            if norm not in labels:
                labels[norm] = (row_number, start_col + c_offset)
    return start_row, start_col, values, labels


def find_anchor_max(
    start_row: int,
    start_col: int,
    values: List[List[Any]],
    labels: Dict[str, Tuple[int, int]],
) -> Optional[Tuple[int, int]]:
    if "max" in labels:
        return labels["max"]
    for r_offset, row_vals in enumerate(values):
        for c_offset, cell_value in enumerate(row_vals):
            if normalize_label(cell_value) == "max":
                return start_row + r_offset, start_col + c_offset
    return None


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


def set_r1c1_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    """Set formula with Formula2 in R1C1 mode; keep conservative fallbacks."""
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.FormulaR1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    cell.formula = formula_r1c1


def get_sheet_by_name(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    for sheet in wb.sheets:
        if normalize_label(sheet.name) == normalize_label(name):
            return sheet
    return None


def find_label_value(
    sheet: xw.Sheet,
    labels: Dict[str, Tuple[int, int]],
    *candidates: str,
    col_offset: int = 1,
    row_offset: int = 0,
) -> Any:
    for candidate in candidates:
        norm = normalize_label(candidate)
        if norm in labels:
            row, col = labels[norm]
            return sheet.cells(row + row_offset, col + col_offset).value
    return None


def collect_numeric_series(
    sheet: xw.Sheet,
    data_col: int,
    stop_row_exclusive: int,
) -> List[Tuple[int, float]]:
    if data_col < 1 or stop_row_exclusive <= 1:
        return []
    values = sheet.range((1, data_col), (stop_row_exclusive - 1, data_col)).value
    values_1d = values if isinstance(values, list) else [values]

    series: List[Tuple[int, float]] = []
    for idx, value in enumerate(values_1d, start=1):
        num = to_float(value)
        if num is not None:
            series.append((idx, num))
    return series


def collect_xy_series(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    stop_row_exclusive: int,
) -> List[Tuple[int, float, float]]:
    if x_col < 1 or y_col < 1 or stop_row_exclusive <= 1:
        return []

    x_vals = sheet.range((1, x_col), (stop_row_exclusive - 1, x_col)).value
    y_vals = sheet.range((1, y_col), (stop_row_exclusive - 1, y_col)).value

    x_1d = x_vals if isinstance(x_vals, list) else [x_vals]
    y_1d = y_vals if isinstance(y_vals, list) else [y_vals]

    size = min(len(x_1d), len(y_1d))
    series: List[Tuple[int, float, float]] = []
    for idx in range(size):
        x_num = to_float(x_1d[idx])
        y_num = to_float(y_1d[idx])
        if x_num is None or y_num is None:
            continue
        series.append((idx + 1, x_num, y_num))
    return series


def process_empirical_sheet(
    wb: xw.Book,
    file_path: Path,
    model_info: Dict[str, str],
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    sheet = get_sheet_by_name(wb, "Empirical Model")
    if sheet is None:
        print(f"  - skipped empirical: missing sheet 'Empirical Model'")
        return rows

    start_row, start_col, values, labels = build_sheet_cache(sheet)
    anchor = find_anchor_max(start_row, start_col, values, labels)
    if anchor is None:
        print(f"  - skipped empirical: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor

    # Anchor-based offsets around the max/min block.
    forecast_max = to_float(sheet.cells(anchor_row, anchor_col + 1).value)
    forecast_min = to_float(sheet.cells(anchor_row + 1, anchor_col + 1).value)

    # Fallback to labeled values if direct offsets are empty.
    if forecast_max is None:
        forecast_max = to_float(find_label_value(sheet, labels, "max"))
    if forecast_min is None:
        forecast_min = to_float(find_label_value(sheet, labels, "min"))

    quarterly_sales = to_float(
        find_label_value(
            sheet,
            labels,
            "quarterly sales",
            "quarterly_sales",
            "quarterly sale",
        )
    )
    reported_sales = to_float(
        find_label_value(
            sheet,
            labels,
            "reported sales",
            "reported_sale",
            "actual sales",
            "actual value",
        )
    )
    growth_rate_pct = to_float(
        find_label_value(sheet, labels, "growth rate", "growth_rate_pct", "growth %")
    )
    sales_captured_in_db_pct = to_float(
        find_label_value(
            sheet,
            labels,
            "sales captured in db %",
            "sales captured in db",
            "sales_captured_in_db_pct",
        )
    )

    # Use anchor offsets and only scan one candidate history column.
    penetration_col = anchor_col - 5
    for label_name, (_, col_number) in labels.items():
        if "penetration" in label_name:
            penetration_col = col_number
            break

    penetration_series = collect_numeric_series(sheet, penetration_col, anchor_row)
    if not penetration_series:
        print(
            f"  - skipped empirical: no numeric penetration history found "
            f"(col={penetration_col})"
        )
        return rows

    quarter_col = max(1, penetration_col - 1)
    helper_avg_cell = sheet.cells(anchor_row + 2, anchor_col + 2)

    max_loops = min(N_QUARTERS, len(penetration_series))
    for n_used in range(1, max_loops + 1):
        subset = penetration_series[-n_used:]
        start_data_row = subset[0][0]
        end_data_row = subset[-1][0]

        avg_formula = (
            f"=AVERAGE(R{start_data_row}C{penetration_col}:"
            f"R{end_data_row}C{penetration_col})"
        )
        set_r1c1_formula2(helper_avg_cell, avg_formula)
        wb.app.calculate()
        avg_penetration_pct = to_float(helper_avg_cell.value)

        # Empirical forecast value = estimated total sold.
        estimated_total_sold = None
        direct_estimated = to_float(
            find_label_value(
                sheet,
                labels,
                "estimated total sold",
                "est total sold",
                "forecast total sold",
            )
        )
        if direct_estimated is not None:
            estimated_total_sold = direct_estimated
        elif reported_sales is not None and avg_penetration_pct not in (None, 0):
            estimated_total_sold = reported_sales / avg_penetration_pct

        last_quarter_used = sheet.cells(end_data_row, quarter_col).value

        row = {
            "model": model_info["model"],
            "ticker": model_info["ticker"],
            "model_period": model_info["model_period"],
            "model_date": model_info["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": n_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": estimated_total_sold,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": file_path.name,
        }
        rows.append(row)

    helper_avg_cell.value = None
    return rows


def process_regression_sheet(
    wb: xw.Book,
    file_path: Path,
    model_info: Dict[str, str],
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    sheet = get_sheet_by_name(wb, "Regression Model")
    if sheet is None:
        print(f"  - skipped regression: missing sheet 'Regression Model'")
        return rows

    start_row, start_col, values, labels = build_sheet_cache(sheet)
    anchor = find_anchor_max(start_row, start_col, values, labels)
    if anchor is None:
        print(f"  - skipped regression: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor

    # Required anchor-based offsets.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    xy_series = collect_xy_series(sheet, x_col, y_col, anchor_row)
    if not xy_series:
        print(
            f"  - skipped regression: no numeric X/Y history found "
            f"(x_col={x_col}, y_col={y_col})"
        )
        return rows

    forecast_max = to_float(sheet.cells(anchor_row, anchor_col + 1).value)
    forecast_min = to_float(sheet.cells(anchor_row + 1, anchor_col + 1).value)
    if forecast_max is None:
        forecast_max = to_float(find_label_value(sheet, labels, "max"))
    if forecast_min is None:
        forecast_min = to_float(find_label_value(sheet, labels, "min"))

    actual_value = to_float(
        find_label_value(
            sheet,
            labels,
            "reported sales",
            "actual sales",
            "actual value",
        )
    )

    helper_intercept = sheet.cells(anchor_row + 2, anchor_col + 2)
    helper_slope = sheet.cells(anchor_row + 3, anchor_col + 2)
    helper_forecast = sheet.cells(anchor_row + 4, anchor_col + 2)

    prev_signature: Optional[Tuple[Any, ...]] = None
    max_loops = min(N_QUARTERS, len(xy_series))
    for n_used in range(1, max_loops + 1):
        subset = xy_series[-n_used:]
        start_data_row = subset[0][0]
        end_data_row = subset[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )
        set_r1c1_formula2(helper_intercept, intercept_formula)
        set_r1c1_formula2(helper_slope, slope_formula)

        # Use latest X in subset as forecast point for TOT FCST w/o SA.
        forecast_formula = (
            f"=R[-2]C+R[-1]C*R{end_data_row}C{x_col}"
        )
        set_r1c1_formula2(helper_forecast, forecast_formula)

        wb.app.calculate()

        intercept = to_float(helper_intercept.value)
        slope = to_float(helper_slope.value)
        forecast_total_without_sa = to_float(helper_forecast.value)

        # Keep workbook-provided TOT FCST w/o SA if available.
        labeled_forecast = to_float(
            find_label_value(
                sheet,
                labels,
                "tot fcst w/o sa",
                "total forecast w/o sa",
                "tot fcst without sa",
            )
        )
        if labeled_forecast is not None:
            forecast_total_without_sa = labeled_forecast

        signature = (
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_total_without_sa, 10)
            if forecast_total_without_sa is not None
            else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        row = {
            "model": model_info["model"],
            "ticker": model_info["ticker"],
            "model_period": model_info["model_period"],
            "model_date": model_info["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_used,
            "num_quarters_used": n_used,
            "forecast_value": forecast_total_without_sa,
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": file_path.name,
        }
        rows.append(row)

    helper_intercept.value = None
    helper_slope.value = None
    helper_forecast.value = None
    return rows


def apply_sheet_formatting(ws) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = min(max(12, max_len + 2), 48)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    ws_empirical.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_empirical.append([row.get(col) for col in EMPIRICAL_COLUMNS])

    ws_regression.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_regression.append([row.get(col) for col in REGRESSION_COLUMNS])

    apply_sheet_formatting(ws_empirical)
    apply_sheet_formatting(ws_regression)
    wb.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")

    output_file = build_output_path(input_path, output_path)
    source_files = list(iter_input_files(input_path))

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                model_info = parse_file_labels(file_path)

                empirical_rows.extend(process_empirical_sheet(wb, file_path, model_info))
                regression_rows.extend(
                    process_regression_sheet(wb, file_path, model_info)
                )
                processed_files += 1
            except Exception as exc:
                print(f"  - skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
