#!/usr/bin/env python3
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
from openpyxl.worksheet.worksheet import Worksheet

# Update these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

MAX_QUARTERS = 10
EPS = 1e-12

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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_NUM = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class XYPoint:
    row: int
    x: float
    y: float
    label: str


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split("-")]

    ticker = "UNKNOWN"
    if len(parts) > 1 and parts[1]:
        ticker = parts[1].upper()

    period_source = parts[2].strip() if len(parts) > 2 else stem
    period_source = period_source.split("_")[0]
    period_source = period_source.replace(" ", "")

    match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_source, flags=re.IGNORECASE)

    model_period = "Unknown_0000"
    model_date = ""
    if match:
        period_tag = match.group(1).title()
        month_abbr = match.group(2)[:3].title()
        year = int(match.group(3))
        model_period = f"{period_tag}{month_abbr}_{year}"

        month_num = MONTH_NUM.get(month_abbr.lower())
        day = PERIOD_DAY[period_tag.lower()]
        if month_num:
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None

    if isinstance(value, (int, float)):
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            return None
        return number

    text = str(value).strip().replace(",", "")
    if not text:
        return None

    percent = text.endswith("%")
    if percent:
        text = text[:-1].strip()

    try:
        number = float(text)
    except ValueError:
        return None

    return number / 100.0 if percent else number


def to_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]
        return [list(values)]
    return [[values]]


def matrix_get(matrix: List[List[Any]], start_row: int, start_col: int, row: int, col: int) -> Any:
    row_idx = row - start_row
    col_idx = col - start_col
    if row_idx < 0 or col_idx < 0 or row_idx >= len(matrix):
        return None
    row_values = matrix[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor(start_row: int, start_col: int, matrix: List[List[Any]], target: str = "max") -> Optional[Tuple[int, int]]:
    expected = target.lower()
    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            text = normalize_text(value)
            if text == expected or text.startswith(f"{expected}:") or text.startswith(f"{expected} "):
                return start_row + row_idx, start_col + col_idx
    return None


def guess_row_label(
    matrix: List[List[Any]],
    start_row: int,
    start_col: int,
    row: int,
    preferred_cols: Sequence[int],
) -> str:
    for col in preferred_cols:
        if col < 1:
            continue
        raw = matrix_get(matrix, start_row, start_col, row, col)
        if raw is None:
            continue
        text = str(raw).strip()
        if text:
            return text
    return f"R{row}"


def collect_xy_series(
    matrix: List[List[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    x_col: int,
    y_col: int,
) -> List[XYPoint]:
    first_row = max(start_row, anchor_row - 240)
    series: List[XYPoint] = []

    for row in range(first_row, anchor_row):
        x_val = to_float(matrix_get(matrix, start_row, start_col, row, x_col))
        y_val = to_float(matrix_get(matrix, start_row, start_col, row, y_col))
        if x_val is None or y_val is None:
            continue
        label = guess_row_label(matrix, start_row, start_col, row, [x_col - 1, y_col - 1, start_col])
        series.append(XYPoint(row=row, x=x_val, y=y_val, label=label))

    return series


def set_formula2_r1c1(cell: xw.main.Range, formula_r1c1: str) -> None:
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
    except Exception:
        cell.formula = formula_r1c1


def close_source_workbook(wb: xw.main.Book) -> None:
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


def safe_divide(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator is None or abs(denominator) < EPS:
        return None
    return numerator / denominator


def maybe_round(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 10)


def get_sheet_if_exists(wb: xw.main.Book, sheet_name: str) -> Optional[xw.main.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def process_empirical_sheet(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    used = sheet.used_range
    start_row, start_col = used.row, used.column
    matrix = to_matrix(used.value)
    anchor = find_anchor(start_row, start_col, matrix, target="max")

    if anchor is None:
        print(f"  skipped empirical in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    series = collect_xy_series(matrix, start_row, start_col, anchor_row, x_col, y_col)
    if not series:
        print(f"  skipped empirical in {source_file}: no numeric x/y series found near anchor")
        return []

    max_quarters = min(MAX_QUARTERS, len(series))
    avg_cell = sheet.range((anchor_row + 2, anchor_col + 3))
    anchor_max = to_float(sheet.range((anchor_row, anchor_col + 1)).value)
    anchor_min = to_float(sheet.range((anchor_row + 1, anchor_col + 1)).value)

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, max_quarters + 1):
        window = series[-n_quarters:]
        start = window[0].row
        end = window[-1].row

        avg_formula = f"=SUM(R{start}C{x_col}:R{end}C{x_col})/SUM(R{start}C{y_col}:R{end}C{y_col})"
        set_formula2_r1c1(avg_cell, avg_formula)
        wb.app.calculate()

        avg_penetration_ratio = to_float(avg_cell.value)
        quarterly_sales = window[-1].x
        reported_sales = window[-1].y
        forecast_value = safe_divide(quarterly_sales, avg_penetration_ratio)

        penetration_values: List[float] = []
        for point in window:
            penetration = safe_divide(point.x, point.y)
            if penetration is not None:
                penetration_values.append(penetration)

        forecast_max = None
        forecast_min = None
        if penetration_values:
            min_pen = min(penetration_values)
            max_pen = max(penetration_values)
            forecast_max = safe_divide(quarterly_sales, min_pen)
            forecast_min = safe_divide(quarterly_sales, max_pen)

        if forecast_max is None:
            forecast_max = anchor_max
        if forecast_min is None:
            forecast_min = anchor_min

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        growth_rate_pct = None
        if len(window) >= 2 and abs(window[-2].x) > EPS:
            growth_rate_pct = ((window[-1].x / window[-2].x) - 1.0) * 100.0

        sales_captured_pct = None
        captured_ratio = safe_divide(quarterly_sales, reported_sales)
        if captured_ratio is not None:
            sales_captured_pct = captured_ratio * 100.0

        avg_penetration_pct = None
        if avg_penetration_ratio is not None:
            avg_penetration_pct = avg_penetration_ratio * 100.0

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": window[0].label,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    used = sheet.used_range
    start_row, start_col = used.row, used.column
    matrix = to_matrix(used.value)
    anchor = find_anchor(start_row, start_col, matrix, target="max")

    if anchor is None:
        print(f"  skipped regression in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    series = collect_xy_series(matrix, start_row, start_col, anchor_row, x_col, y_col)
    if not series:
        print(f"  skipped regression in {source_file}: no numeric x/y series found near anchor")
        return []

    max_quarters = min(MAX_QUARTERS, len(series))
    intercept_cell = sheet.range((anchor_row + 2, anchor_col + 6))
    slope_cell = sheet.range((anchor_row + 3, anchor_col + 6))
    anchor_max = to_float(sheet.range((anchor_row, anchor_col + 1)).value)
    anchor_min = to_float(sheet.range((anchor_row + 1, anchor_col + 1)).value)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for n_quarters in range(1, max_quarters + 1):
        window = series[-n_quarters:]
        start = window[0].row
        end = window[-1].row

        intercept_formula = f"=INTERCEPT(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        slope_formula = f"=SLOPE(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        latest_x = window[-1].x

        forecast_total_without_sa = None
        if intercept is not None and slope is not None:
            forecast_total_without_sa = intercept + slope * latest_x

        range_width = None
        if anchor_max is not None and anchor_min is not None:
            range_width = anchor_max - anchor_min

        signature = (
            maybe_round(forecast_total_without_sa),
            maybe_round(anchor_max),
            maybe_round(anchor_min),
            maybe_round(intercept),
            maybe_round(slope),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": "",
                "forecast_max": anchor_max,
                "forecast_min": anchor_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def set_column_widths(ws: Worksheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    for idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row in rows:
            value = row.get(column, "")
            if value is None:
                text = ""
            elif isinstance(value, float):
                text = f"{value:.6g}"
            else:
                text = str(value)
            max_len = max(max_len, len(text))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 52)


def write_sheet(ws: Worksheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    set_column_widths(ws, columns, rows)


def write_output_workbook(output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]) -> None:
    wb_out = Workbook()
    ws_emp = wb_out.active
    ws_emp.title = "empirical_candidates"
    write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = wb_out.create_sheet(title="regression_candidates")
    write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb_out.save(output_path)


def discover_files(in_dir: Path) -> List[Path]:
    files: List[Path] = []
    if not in_dir.exists():
        print(f"skipped input_dir: not found ({in_dir})")
        return files

    for file_path in sorted(in_dir.iterdir()):
        if not file_path.is_file():
            print(f"skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped {file_path.name}: not .xlsx")
            continue
        files.append(file_path)
    return files


def main() -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    files = discover_files(input_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.main.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        try:
            app.display_alerts = False
            app.screen_updating = False
        except Exception:
            pass

        for file_path in files:
            print(f"processing {file_path.name}")
            wb: Optional[xw.main.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path.name)

                empirical_sheet = get_sheet_if_exists(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"  skipped empirical in {file_path.name}: sheet not found")
                else:
                    empirical_rows.extend(process_empirical_sheet(wb, empirical_sheet, label, file_path.name))

                regression_sheet = get_sheet_if_exists(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"  skipped regression in {file_path.name}: sheet not found")
                else:
                    regression_rows.extend(process_regression_sheet(wb, regression_sheet, label, file_path.name))

                processed_files += 1
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_path = next_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
