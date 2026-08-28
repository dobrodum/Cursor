#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks.

This script scans an input folder for .xlsx files, opens each workbook once,
extracts both Empirical Model and Regression Model candidates, and writes one
output workbook with two sheets:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------
# User-configurable paths
# ---------------------------
input_dir = Path("./input")
output_dir = Path("./output")


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

MONTH_TO_INT = {
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class ParsedLabel:
    model: str
    ticker: str
    model_period: str
    model_date: Optional[str]


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def to_col(values: Any) -> List[Any]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [values]
    if values and isinstance(values[0], list):
        return [row[0] if row else None for row in values]
    return values


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        pct = text.endswith("%")
        if pct:
            text = text[:-1].strip()
        try:
            num = float(text)
        except ValueError:
            return None
        return num / 100.0 if pct else num
    return None


def normalize_output(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def parse_file_label(file_path: Path) -> ParsedLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].strip().upper() if len(parts) > 1 and parts[1].strip() else "UNKNOWN"

    period_source = parts[2] if len(parts) > 2 else stem
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?\s*(20\d{2})",
        period_source,
        flags=re.IGNORECASE,
    )

    model_period = "Unknown_0000"
    model_date: Optional[str] = None

    if period_match:
        period_word = period_match.group(1).capitalize()
        month_token = period_match.group(2)[:3].lower()
        year = int(period_match.group(3))
        month = MONTH_TO_INT.get(month_token)
        day = PERIOD_DAY[period_word.lower()]
        if month:
            model_period = f"{period_word}{month_token.capitalize()}_{year}"
            model_date = date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}"
    return ParsedLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def safe_get_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
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
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    # Preferred path: Formula2R1C1 via API.
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    # Fallback for environments exposing formula2 directly.
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    # Last fallback still keeps R1C1 addressing.
    cell.api.FormulaR1C1 = formula_r1c1


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = to_2d(used.value)
    if not matrix:
        return None

    start_row = used.row
    start_col = used.column

    candidates: List[Tuple[int, int, int]] = []  # score, row, col
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_label(value) != "max":
                continue
            score = 0

            # Prefer anchors with a nearby "min" label.
            right_label = ""
            if c_idx + 1 < len(row):
                right_label = normalize_label(row[c_idx + 1])
            if right_label == "min":
                score += 3

            below_value: Any = None
            if r_idx + 1 < len(matrix) and c_idx < len(matrix[r_idx + 1]):
                below_value = matrix[r_idx + 1][c_idx]
            if to_number(below_value) is not None:
                score += 1

            candidates.append((score, start_row + r_idx, start_col + c_idx))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    return (candidates[0][1], candidates[0][2])


def window_values(
    sheet: xw.Sheet,
    center_row: int,
    center_col: int,
    row_radius: int = 80,
    col_radius: int = 45,
) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    max_row = used.last_cell.row
    max_col = used.last_cell.column

    top = max(1, center_row - row_radius)
    left = max(1, center_col - col_radius)
    bottom = min(max_row, center_row + row_radius)
    right = min(max_col, center_col + col_radius)

    matrix = to_2d(sheet.range((top, left), (bottom, right)).value)
    return top, left, matrix


def within(matrix: List[List[Any]], r: int, c: int) -> bool:
    return 0 <= r < len(matrix) and 0 <= c < len(matrix[r])


def find_value_cells_near_anchor(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    labels: Dict[str, Sequence[str]],
) -> Dict[str, Tuple[int, int]]:
    top, left, matrix = window_values(sheet, anchor_row, anchor_col)
    found: Dict[str, Tuple[int, int]] = {}

    for key, variants in labels.items():
        target_norms = [normalize_label(v) for v in variants]
        cell_ref: Optional[Tuple[int, int]] = None

        for r_idx, row in enumerate(matrix):
            if cell_ref:
                break
            for c_idx, value in enumerate(row):
                label_norm = normalize_label(value)
                if not label_norm:
                    continue
                if not any(target in label_norm for target in target_norms):
                    continue

                # Check nearby cells for a likely value cell.
                for dr, dc in ((0, 1), (1, 0), (0, -1), (-1, 0)):
                    nr, nc = r_idx + dr, c_idx + dc
                    if not within(matrix, nr, nc):
                        continue
                    neighbor = matrix[nr][nc]
                    if neighbor is None or neighbor == "":
                        continue
                    cell_ref = (top + nr, left + nc)
                    break

                # Fallback if no clear value found: pick right-adjacent cell.
                if not cell_ref:
                    abs_row = top + r_idx
                    abs_col = left + c_idx + 1
                    if abs_col >= 1:
                        cell_ref = (abs_row, abs_col)
                if cell_ref:
                    break

        if cell_ref:
            found[key] = cell_ref

    return found


def find_min_value_cell(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> Tuple[int, int]:
    top, left, matrix = window_values(sheet, anchor_row, anchor_col, row_radius=5, col_radius=10)
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_label(value) != "min":
                continue
            nr, nc = r_idx + 1, c_idx
            if within(matrix, nr, nc):
                return (top + nr, left + nc)
    return (anchor_row + 1, anchor_col + 1)


def read_cell(sheet: xw.Sheet, ref: Optional[Tuple[int, int]]) -> Any:
    if ref is None:
        return None
    row, col = ref
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def collect_numeric_series(
    sheet: xw.Sheet,
    col: int,
    end_row: int,
    lookback: int = 120,
) -> List[Tuple[int, float]]:
    if col < 1 or end_row < 1:
        return []
    start_row = max(1, end_row - lookback + 1)
    values = to_col(sheet.range((start_row, col), (end_row, col)).value)
    series: List[Tuple[int, float]] = []
    for offset, value in enumerate(values):
        row = start_row + offset
        num = to_number(value)
        if num is not None:
            series.append((row, num))
    return series


def calculate_if_needed(app: xw.App) -> None:
    try:
        app.calculate()
    except Exception:
        app.api.Calculate()


def calc_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def extract_empirical_rows(
    wb: xw.Book,
    label: ParsedLabel,
    source_file: str,
) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    sheet = safe_get_sheet(wb, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        return [], "Empirical Model sheet missing"

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return [], "max anchor not found in Empirical Model"

    anchor_row, anchor_col = anchor
    named_cells = find_value_cells_near_anchor(
        sheet,
        anchor_row,
        anchor_col,
        {
            "avg_penetration_pct": ("avg penetration", "average penetration"),
            "estimated_total_sold": ("estimated total sold", "est total sold"),
            "reported_sales": ("reported sales", "actual sales"),
            "quarterly_sales": ("quarterly sales",),
            "growth_rate_pct": ("growth rate",),
            "sales_captured_in_db_pct": ("sales captured in db", "captured in db"),
            "last_quarter_used": ("last quarter",),
        },
    )

    avg_pen_ref = named_cells.get("avg_penetration_pct")
    if avg_pen_ref is None:
        avg_pen_ref = (anchor_row + 1, max(1, anchor_col - 6))

    avg_pen_cell = sheet.cells(avg_pen_ref[0], avg_pen_ref[1])
    original_avg_formula = None
    try:
        original_avg_formula = avg_pen_cell.formula
    except Exception:
        original_avg_formula = None

    forecast_max_ref = (anchor_row + 1, anchor_col)
    forecast_min_ref = find_min_value_cell(sheet, anchor_row, anchor_col)

    history = collect_numeric_series(sheet, avg_pen_ref[1], avg_pen_ref[0] - 1, lookback=160)
    if not history:
        return [], "No numeric penetration history found"

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        subset = history[-min(n_quarters, len(history)) :]
        if not subset:
            break
        subset_start_row = subset[0][0]
        subset_end_row = subset[-1][0]

        # R1C1 + formula2 path (with API fallback) for fast recalculation.
        avg_formula = (
            f"=AVERAGE(R{subset_start_row}C{avg_pen_ref[1]}:R{subset_end_row}C{avg_pen_ref[1]})"
        )
        set_formula2_r1c1(avg_pen_cell, avg_formula)
        calculate_if_needed(wb.app)

        avg_penetration = to_number(avg_pen_cell.value)
        forecast_value = to_number(read_cell(sheet, named_cells.get("estimated_total_sold")))
        reported_sales = read_cell(sheet, named_cells.get("reported_sales"))
        quarterly_sales = read_cell(sheet, named_cells.get("quarterly_sales"))
        growth_rate_pct = read_cell(sheet, named_cells.get("growth_rate_pct"))
        sales_captured_pct = read_cell(sheet, named_cells.get("sales_captured_in_db_pct"))
        last_quarter_used = read_cell(sheet, named_cells.get("last_quarter_used"))

        if last_quarter_used in (None, ""):
            # Fallback: quarter label is often adjacent to the penetration history.
            last_quarter_used = read_cell(sheet, (subset_end_row, max(1, avg_pen_ref[1] - 1)))

        forecast_max = to_number(read_cell(sheet, forecast_max_ref))
        forecast_min = to_number(read_cell(sheet, forecast_min_ref))

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n_quarters,
            "last_quarter_used": normalize_output(last_quarter_used),
            "forecast_value": forecast_value,
            "actual_value": normalize_output(reported_sales),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": normalize_output(quarterly_sales),
            "reported_sales": normalize_output(reported_sales),
            "growth_rate_pct": normalize_output(growth_rate_pct),
            "sales_captured_in_db_pct": normalize_output(sales_captured_pct),
            "source_file": source_file,
        }
        rows.append(row)

    if original_avg_formula:
        try:
            avg_pen_cell.formula = original_avg_formula
            calculate_if_needed(wb.app)
        except Exception:
            pass

    return rows, None


def collect_xy_pairs(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    end_row: int,
    lookback: int = 180,
) -> List[Tuple[int, float, float]]:
    start_row = max(1, end_row - lookback + 1)
    x_vals = to_col(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_vals = to_col(sheet.range((start_row, y_col), (end_row, y_col)).value)

    pairs: List[Tuple[int, float, float]] = []
    for idx, (x_val, y_val) in enumerate(zip(x_vals, y_vals)):
        row = start_row + idx
        x_num = to_number(x_val)
        y_num = to_number(y_val)
        if x_num is None or y_num is None:
            continue
        pairs.append((row, x_num, y_num))
    return pairs


def regression_rows_duplicate(previous: Dict[str, Any], current: Dict[str, Any]) -> bool:
    keys = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    for key in keys:
        left = to_number(previous.get(key))
        right = to_number(current.get(key))
        if left is None and right is None:
            continue
        if left is None or right is None:
            return False
        if not math.isclose(left, right, rel_tol=1e-9, abs_tol=1e-9):
            return False
    return True


def extract_regression_rows(
    wb: xw.Book,
    label: ParsedLabel,
    source_file: str,
) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    sheet = safe_get_sheet(wb, REGRESSION_SHEET_NAME)
    if sheet is None:
        return [], "Regression Model sheet missing"

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return [], "max anchor not found in Regression Model"

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        return [], "Regression anchor offsets out of bounds"

    named_cells = find_value_cells_near_anchor(
        sheet,
        anchor_row,
        anchor_col,
        {
            "tot_fcst_wo_sa": ("tot fcst w o sa", "tot fcst wo sa", "total forecast without sa"),
            "actual_value": ("actual sales", "reported sales"),
        },
    )

    forecast_max_ref = (anchor_row + 1, anchor_col)
    forecast_min_ref = find_min_value_cell(sheet, anchor_row, anchor_col)
    actual_ref = named_cells.get("actual_value")
    tot_fcst_ref = named_cells.get("tot_fcst_wo_sa")

    pairs = collect_xy_pairs(sheet, x_col, y_col, end_row=anchor_row - 1)
    if len(pairs) < 2:
        return [], "Insufficient x/y history for regression"

    temp_col = anchor_col + 8
    intercept_cell = sheet.cells(anchor_row + 2, temp_col)
    slope_cell = sheet.cells(anchor_row + 3, temp_col)
    forecast_cell = sheet.cells(anchor_row + 4, temp_col)

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        subset = pairs[-min(n_quarters, len(pairs)) :]
        if len(subset) < 2:
            continue

        start_row = subset[0][0]
        end_row = subset[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)

        next_x = subset[-1][1] + 1.0
        forecast_formula = (
            f"=R{intercept_cell.row}C{intercept_cell.column}"
            f"+R{slope_cell.row}C{slope_cell.column}*{next_x}"
        )
        set_formula2_r1c1(forecast_cell, forecast_formula)
        calculate_if_needed(wb.app)

        intercept = to_number(intercept_cell.value)
        slope = to_number(slope_cell.value)
        forecast_total_without_sa = to_number(read_cell(sheet, tot_fcst_ref))
        if forecast_total_without_sa is None:
            forecast_total_without_sa = to_number(forecast_cell.value)

        forecast_max = to_number(read_cell(sheet, forecast_max_ref))
        forecast_min = to_number(read_cell(sheet, forecast_min_ref))
        if forecast_max is None:
            forecast_max = max(point[2] for point in subset)
        if forecast_min is None:
            forecast_min = min(point[2] for point in subset)

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_total_without_sa,
            "actual_value": normalize_output(read_cell(sheet, actual_ref)),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and regression_rows_duplicate(rows[-1], row):
            continue
        rows.append(row)

    return rows, None


def output_file_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_sheet(
    ws: Any,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    for col_idx, column_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=column_name)
        cell.font = Font(bold=True)

    for row_idx, row in enumerate(rows, start=2):
        for col_idx, column_name in enumerate(columns, start=1):
            value = normalize_output(row.get(column_name))
            ws.cell(row=row_idx, column=col_idx, value=value)

    ws.freeze_panes = "A2"
    bottom_row = max(2, len(rows) + 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{bottom_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, len(rows) + 2):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()

    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = workbook.create_sheet(title="regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def iter_candidate_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        yield path


def main() -> None:
    src_dir = Path(input_dir).expanduser().resolve()
    dst_dir = Path(output_dir).expanduser().resolve()

    if not src_dir.exists():
        print(f"Input folder does not exist: {src_dir}")
        return

    files_to_process: List[Path] = []
    for file_path in iter_candidate_files(src_dir):
        if file_path.name.startswith("~"):
            print(f"Skipped: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped: {file_path.name} (not .xlsx)")
            continue
        files_to_process.append(file_path)

    output_path = output_file_path(src_dir, dst_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files_to_process:
            print(f"Processing: {file_path.name}")
            parsed = parse_file_label(file_path)

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (open failed: {exc})")
                continue

            processed_files += 1
            try:
                emp_rows, emp_note = extract_empirical_rows(wb, parsed, file_path.name)
                if emp_note:
                    print(f"  Empirical: {emp_note}")
                empirical_rows.extend(emp_rows)

                reg_rows, reg_note = extract_regression_rows(wb, parsed, file_path.name)
                if reg_note:
                    print(f"  Regression: {reg_note}")
                regression_rows.extend(reg_rows)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (extraction error: {exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
