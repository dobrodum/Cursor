#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from source Excel workbooks.

This script:
1) scans all .xlsx files in input_dir (skipping temp files),
2) opens each workbook once in one hidden Excel app,
3) processes both "Empirical Model" and "Regression Model" while open,
4) writes one output workbook with:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# User-configurable inputs
# =========================
input_dir = "/path/to/input"
output_dir = "/path/to/output"


# =========================
# Constants
# =========================
N_QUARTERS = 10
EARLY_MID_LATE_DAY = {"Early": 5, "Mid": 15, "Late": 25}
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


@dataclass(frozen=True)
class ModelLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    has_percent = "%" in text
    text = text.replace(",", "").replace("%", "")
    try:
        number = float(text)
    except ValueError:
        return None
    if has_percent:
        return number / 100.0
    return number


def ratio_from_percent_like(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if value > 1:
        return value / 100.0
    return value


def safe_divide(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def round_signature(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 10)


def parse_model_label(file_name: str) -> ModelLabel:
    stem = Path(file_name).stem
    ticker = ""
    period_token = ""

    parts = [part.strip() for part in stem.split("-")]
    if len(parts) >= 3:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        period_match = re.search(r"(Early|Mid|Late)[A-Za-z]{3,9}\d{4}", parts[2], flags=re.IGNORECASE)
        if period_match:
            period_token = period_match.group(0)

    if not ticker:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    if not period_token:
        period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", stem, flags=re.IGNORECASE)
        if period_match:
            period_token = period_match.group(0)

    model_period = ""
    model_date = ""

    if period_token:
        period_parts = re.match(
            r"(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
            period_token,
            flags=re.IGNORECASE,
        )
        if period_parts:
            bucket = period_parts.group("bucket").title()
            month_text = period_parts.group("month")
            year_text = period_parts.group("year")
            month_key = month_text[:3].lower()
            month_num = MONTH_MAP.get(month_key)
            if month_num is not None:
                month_abbrev = month_key.title()
                model_period = f"{bucket}{month_abbrev}_{year_text}"
                day = EARLY_MID_LATE_DAY[bucket]
                model_date = date(int(year_text), month_num, day).isoformat()

    model = f"{ticker}_{model_period}".strip("_")
    return ModelLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    input_folder_name = input_path.name
    base = f"{input_folder_name}_PARAM"
    candidate = output_path / f"{base}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_path / f"{base}.{counter}.xlsx"
        counter += 1
    return candidate


def list_source_files(input_path: Path) -> List[Path]:
    source_files: List[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary workbook")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not .xlsx")
            continue
        source_files.append(file_path)
    return source_files


def read_used_matrix(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int]:
    used = sheet.used_range
    matrix = ensure_2d(used.value)
    return matrix, used.row, used.column


def matrix_value(matrix: List[List[Any]], top_row: int, left_col: int, row: int, col: int) -> Any:
    row_idx = row - top_row
    col_idx = col - left_col
    if row_idx < 0 or col_idx < 0 or row_idx >= len(matrix):
        return None
    row_values = matrix[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor_max(matrix: List[List[Any]], top_row: int, left_col: int) -> Optional[Tuple[int, int]]:
    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return top_row + row_idx, left_col + col_idx
    return None


def contiguous_history_rows(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    anchor_row: int,
    x_col: int,
    y_col: int,
) -> List[int]:
    rows: List[int] = []
    for row in range(anchor_row - 1, top_row - 1, -1):
        x_value = to_float(matrix_value(matrix, top_row, left_col, row, x_col))
        y_value = to_float(matrix_value(matrix, top_row, left_col, row, y_col))
        if x_value is None or y_value is None:
            if rows:
                break
            continue
        rows.append(row)
    rows.sort()
    return rows


def snapshot_cell(cell: xw.Range) -> Tuple[Any, Any, Any]:
    formula2 = None
    try:
        formula2 = cell.formula2
    except Exception:
        formula2 = None
    formula = None
    try:
        formula = cell.formula
    except Exception:
        formula = None
    value = cell.value
    return formula2, formula, value


def restore_cell(cell: xw.Range, snapshot: Tuple[Any, Any, Any]) -> None:
    formula2, formula, value = snapshot
    try:
        if formula2 not in (None, ""):
            cell.formula2 = formula2
            return
    except Exception:
        pass
    try:
        if formula not in (None, ""):
            cell.formula = formula
            return
    except Exception:
        pass
    cell.value = value


def set_r1c1_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    # Fallback when Formula2 assignment differs across Excel engine versions.
    cell.api.Formula2R1C1 = formula_r1c1


def value_from_nearby_label(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    anchor_row: int,
    anchor_col: int,
    labels: Sequence[str],
) -> Any:
    labels_norm = [normalize_text(label) for label in labels]
    best_position: Optional[Tuple[int, int]] = None
    best_distance: Optional[int] = None

    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            text = normalize_text(value)
            if not text:
                continue
            if not any(label in text for label in labels_norm):
                continue
            row = top_row + row_idx
            col = left_col + col_idx
            distance = abs(row - anchor_row) + abs(col - anchor_col)
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_position = (row, col)

    if best_position is None:
        return None

    row, col = best_position
    return matrix_value(matrix, top_row, left_col, row, col + 1)


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


def extract_empirical_candidates(wb: xw.Book, meta: ModelLabel, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped {source_file} empirical extraction: missing 'Empirical Model' sheet")
        return []

    matrix, top_row, left_col = read_used_matrix(sheet)
    if not matrix:
        print(f"Skipped {source_file} empirical extraction: empty 'Empirical Model' sheet")
        return []

    anchor = find_anchor_max(matrix, top_row, left_col)
    if anchor is None:
        print(f"Skipped {source_file} empirical extraction: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    penetration_col = anchor_col - 11
    sales_col = anchor_col - 7
    quarter_label_col = penetration_col - 1

    history_rows = contiguous_history_rows(matrix, top_row, left_col, anchor_row, penetration_col, sales_col)
    if len(history_rows) < 2:
        print(f"Skipped {source_file} empirical extraction: insufficient numeric history")
        return []

    reported_sales_fixed = to_float(
        value_from_nearby_label(
            matrix,
            top_row,
            left_col,
            anchor_row,
            anchor_col,
            labels=("reported sales", "actual sales"),
        )
    )
    growth_rate_fixed = to_float(
        value_from_nearby_label(
            matrix,
            top_row,
            left_col,
            anchor_row,
            anchor_col,
            labels=("growth rate", "growth %"),
        )
    )
    sales_captured_fixed = to_float(
        value_from_nearby_label(
            matrix,
            top_row,
            left_col,
            anchor_row,
            anchor_col,
            labels=("sales captured in db", "captured in db"),
        )
    )

    scratch_avg = sheet.cells(anchor_row, anchor_col + 3)
    scratch_snapshot = snapshot_cell(scratch_avg)

    rows: List[Dict[str, Any]] = []
    try:
        for n_quarters in range(1, N_QUARTERS + 1):
            if len(history_rows) < n_quarters:
                continue
            sample_rows = history_rows[-n_quarters:]
            start_row = sample_rows[0]
            end_row = sample_rows[-1]

            avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
            set_r1c1_formula2(scratch_avg, avg_formula)
            wb.app.calculate()

            avg_penetration_raw = to_float(scratch_avg.value)
            avg_penetration_ratio = ratio_from_percent_like(avg_penetration_raw)
            avg_penetration_pct = None if avg_penetration_ratio is None else avg_penetration_ratio * 100.0

            quarterly_sales = to_float(matrix_value(matrix, top_row, left_col, end_row, sales_col))
            reported_sales = reported_sales_fixed
            if reported_sales is None:
                reported_sales = to_float(matrix_value(matrix, top_row, left_col, anchor_row, sales_col))

            forecast_value = safe_divide(quarterly_sales, avg_penetration_ratio)

            penetration_values: List[float] = []
            for row in sample_rows:
                p = ratio_from_percent_like(to_float(matrix_value(matrix, top_row, left_col, row, penetration_col)))
                if p is not None:
                    penetration_values.append(p)

            forecast_max = None
            forecast_min = None
            if quarterly_sales is not None and penetration_values:
                min_pen = min(penetration_values)
                max_pen = max(penetration_values)
                forecast_max = safe_divide(quarterly_sales, min_pen)
                forecast_min = safe_divide(quarterly_sales, max_pen)

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            last_quarter_used = matrix_value(matrix, top_row, left_col, end_row, quarter_label_col)

            growth_rate_pct = growth_rate_fixed
            if growth_rate_pct is None and len(sample_rows) >= 2:
                previous_sales = to_float(matrix_value(matrix, top_row, left_col, sample_rows[-2], sales_col))
                growth_ratio = safe_divide(
                    None if quarterly_sales is None or previous_sales is None else quarterly_sales - previous_sales,
                    previous_sales,
                )
                growth_rate_pct = None if growth_ratio is None else growth_ratio * 100.0
            elif growth_rate_pct is not None and abs(growth_rate_pct) <= 1:
                growth_rate_pct = growth_rate_pct * 100.0

            sales_captured_in_db_pct = sales_captured_fixed
            if sales_captured_in_db_pct is None:
                sales_captured_in_db_pct = avg_penetration_pct
            elif abs(sales_captured_in_db_pct) <= 1:
                sales_captured_in_db_pct = sales_captured_in_db_pct * 100.0

            row = {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
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
            rows.append(row)
    finally:
        restore_cell(scratch_avg, scratch_snapshot)

    return rows


def extract_regression_candidates(wb: xw.Book, meta: ModelLabel, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped {source_file} regression extraction: missing 'Regression Model' sheet")
        return []

    matrix, top_row, left_col = read_used_matrix(sheet)
    if not matrix:
        print(f"Skipped {source_file} regression extraction: empty 'Regression Model' sheet")
        return []

    anchor = find_anchor_max(matrix, top_row, left_col)
    if anchor is None:
        print(f"Skipped {source_file} regression extraction: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    history_rows = contiguous_history_rows(matrix, top_row, left_col, anchor_row, x_col, y_col)
    if len(history_rows) < 2:
        print(f"Skipped {source_file} regression extraction: insufficient numeric history")
        return []

    actual_value = to_float(
        value_from_nearby_label(
            matrix,
            top_row,
            left_col,
            anchor_row,
            anchor_col,
            labels=("actual value", "actual sales", "reported sales"),
        )
    )
    if actual_value is None:
        actual_value = to_float(matrix_value(matrix, top_row, left_col, anchor_row, y_col))

    scratch_intercept = sheet.cells(anchor_row, anchor_col + 3)
    scratch_slope = sheet.cells(anchor_row + 1, anchor_col + 3)
    intercept_snapshot = snapshot_cell(scratch_intercept)
    slope_snapshot = snapshot_cell(scratch_slope)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    try:
        for n_quarters in range(1, N_QUARTERS + 1):
            if len(history_rows) < n_quarters:
                continue
            sample_rows = history_rows[-n_quarters:]
            start_row = sample_rows[0]
            end_row = sample_rows[-1]

            intercept_formula = (
                f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
            )
            slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
            set_r1c1_formula2(scratch_intercept, intercept_formula)
            set_r1c1_formula2(scratch_slope, slope_formula)
            wb.app.calculate()

            intercept = to_float(scratch_intercept.value)
            slope = to_float(scratch_slope.value)

            x_target = to_float(matrix_value(matrix, top_row, left_col, anchor_row, x_col))
            if x_target is None:
                x_target = to_float(matrix_value(matrix, top_row, left_col, end_row, x_col))

            forecast_total_without_sa = None
            if intercept is not None and slope is not None and x_target is not None:
                forecast_total_without_sa = intercept + slope * x_target

            y_values = [
                to_float(matrix_value(matrix, top_row, left_col, row, y_col))
                for row in sample_rows
                if to_float(matrix_value(matrix, top_row, left_col, row, y_col)) is not None
            ]
            forecast_max = max(y_values) if y_values else None
            forecast_min = min(y_values) if y_values else None
            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            signature = (
                round_signature(forecast_total_without_sa),
                round_signature(intercept),
                round_signature(slope),
                round_signature(forecast_max),
                round_signature(forecast_min),
            )
            if signature == previous_signature:
                # Guard against duplicate trailing rows from repeated calculations.
                continue
            previous_signature = signature

            row = {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
            rows.append(row)
    finally:
        restore_cell(scratch_intercept, intercept_snapshot)
        restore_cell(scratch_slope, slope_snapshot)

    return rows


def write_rows(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(ws.cell(row=1, column=col_idx).value or ""))
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    ws_empirical = workbook.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = workbook.create_sheet("regression_candidates")

    write_rows(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows(ws_regression, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_file = unique_output_path(input_path, output_path)
    source_files = list_source_files(input_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"Processing {file_path.name}")
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open ({exc})")
                continue

            try:
                label = parse_model_label(file_path.name)
                emp_rows = extract_empirical_candidates(wb, label, file_path.name)
                reg_rows = extract_regression_candidates(wb, label, file_path.name)
                empirical_rows.extend(emp_rows)
                regression_rows.extend(reg_rows)
                files_processed += 1
                print(
                    f"Processed {file_path.name}: "
                    f"empirical_rows={len(emp_rows)}, regression_rows={len(reg_rows)}"
                )
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing failure ({exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
