#!/usr/bin/env python3
"""
Extract empirical and regression candidate rows from model workbooks.

Runtime-focused design:
- One hidden Excel app for the entire run.
- Each source workbook is opened exactly once.
- Both model sheets are processed while the workbook is open.
- Source workbooks are always closed without saving.
"""

from __future__ import annotations

import datetime as dt
import re
from itertools import count
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10

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

PHASE_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


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

    is_pct = text.endswith("%")
    if is_pct:
        text = text[:-1]

    text = text.replace(",", "").replace("$", "")

    try:
        number = float(text)
    except ValueError:
        return None

    if is_pct:
        return number / 100.0
    return number


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def flatten_column_values(values: Any, expected_len: int) -> List[Any]:
    matrix = to_2d(values)
    flattened: List[Any] = []
    for row in matrix:
        if isinstance(row, list) and row:
            flattened.append(row[0])
        else:
            flattened.append(row)
    if len(flattened) < expected_len:
        flattened.extend([None] * (expected_len - len(flattened)))
    return flattened[:expected_len]


def round_or_none(value: Optional[float], places: int = 10) -> Optional[float]:
    if value is None:
        return None
    return round(value, places)


def safe_diff(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return a - b


def load_sheet_matrix(sheet: xw.Sheet) -> Dict[str, Any]:
    used = sheet.used_range
    values = to_2d(used.value)
    n_cols = len(values[0]) if values else 0
    return {
        "row0": used.row,
        "col0": used.column,
        "rows": len(values),
        "cols": n_cols,
        "values": values,
    }


def matrix_get(matrix: Dict[str, Any], row: int, col: int) -> Any:
    r = row - matrix["row0"]
    c = col - matrix["col0"]
    if r < 0 or c < 0:
        return None
    if r >= matrix["rows"] or c >= matrix["cols"]:
        return None
    return matrix["values"][r][c]


def find_max_anchor(matrix: Dict[str, Any]) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int]] = []
    for r_idx, row in enumerate(matrix["values"]):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                abs_row = matrix["row0"] + r_idx
                abs_col = matrix["col0"] + c_idx
                candidates.append((abs_row, abs_col))

    if not candidates:
        return None

    def candidate_score(pos: Tuple[int, int]) -> Tuple[int, int]:
        row, col = pos
        score = 0
        for delta in range(1, 6):
            if normalize_text(matrix_get(matrix, row, col + delta)) == "min":
                score += (6 - delta)
        # Prefer left-most anchors if scores tie.
        return (score, -col)

    return max(candidates, key=candidate_score)


def find_col_by_patterns(
    matrix: Dict[str, Any],
    row: int,
    anchor_col: int,
    pattern_sets: Sequence[Sequence[str]],
    window: int = 60,
) -> Optional[int]:
    if matrix["rows"] == 0 or matrix["cols"] == 0:
        return None

    min_col = max(matrix["col0"], anchor_col - window)
    max_col = min(matrix["col0"] + matrix["cols"] - 1, anchor_col + window)

    matches: List[Tuple[int, int]] = []
    for col in range(min_col, max_col + 1):
        text = normalize_text(matrix_get(matrix, row, col))
        if not text:
            continue
        for patterns in pattern_sets:
            if all(part in text for part in patterns):
                matches.append((abs(col - anchor_col), col))
                break

    if not matches:
        return None
    matches.sort()
    return matches[0][1]


def history_bounds(anchor_row: int, n_quarters: int) -> Tuple[int, int]:
    end_row = anchor_row - 1
    start_row = max(1, end_row - n_quarters + 1)
    return start_row, end_row


def column_values(matrix: Dict[str, Any], start_row: int, end_row: int, col: int) -> List[float]:
    values: List[float] = []
    for row in range(start_row, end_row + 1):
        number = to_float(matrix_get(matrix, row, col))
        if number is not None:
            values.append(number)
    return values


def column_sum(matrix: Dict[str, Any], start_row: int, end_row: int, col: int) -> Optional[float]:
    values = column_values(matrix, start_row, end_row, col)
    if not values:
        return None
    return sum(values)


def first_last_numeric(
    matrix: Dict[str, Any], start_row: int, end_row: int, col: int
) -> Tuple[Optional[float], Optional[float]]:
    values = column_values(matrix, start_row, end_row, col)
    if not values:
        return None, None
    return values[0], values[-1]


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass

    # Fallbacks for environments where formula2 R1C1 assignment is COM-only.
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass

    cell.api.FormulaR1C1 = formula


def calculate_formulas(
    sheet: xw.Sheet,
    formulas: Sequence[str],
    temp_col: int,
    temp_row_start: int = 2,
) -> List[Any]:
    if not formulas:
        return []

    for idx, formula in enumerate(formulas):
        row = temp_row_start + idx
        set_formula2_r1c1(sheet.range((row, temp_col)), formula)

    sheet.book.app.calculate()

    temp_end_row = temp_row_start + len(formulas) - 1
    values = sheet.range((temp_row_start, temp_col), (temp_end_row, temp_col)).value
    sheet.range((temp_row_start, temp_col), (temp_end_row, temp_col)).clear_contents()
    return flatten_column_values(values, expected_len=len(formulas))


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
    except Exception:
        # Last resort; still avoid save calls.
        try:
            wb.close()
        except Exception:
            pass


def parse_file_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = parts[1].upper() if len(parts) >= 2 else ""

    period_match = re.search(
        r"(Early|Mid|Late)[ _-]*([A-Za-z]{3,9})[ _-]*(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase = period_match.group(1).title()
        month_text = period_match.group(2).strip()
        year = int(period_match.group(3))

        month_key = month_text[:3].lower()
        month_num = MONTH_MAP.get(month_key)
        if month_num is not None:
            day = PHASE_DAY_MAP[phase.lower()]
            model_period = f"{phase}{month_text[:3].title()}_{year}"
            model_date = dt.date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or stem
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def find_next_output_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM.xlsx"
    candidate = output_folder / base_name
    if not candidate.exists():
        return candidate

    for idx in count(1):
        candidate = output_folder / f"{input_folder.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate

    raise RuntimeError("Could not resolve an output file path.")


def list_source_files(folder: Path) -> Tuple[List[Path], List[Tuple[Path, str]]]:
    source_files: List[Path] = []
    skipped: List[Tuple[Path, str]] = []

    if not folder.exists():
        return source_files, skipped

    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue

        if path.name.startswith("~"):
            skipped.append((path, "temp file"))
            continue

        if path.suffix.lower() != ".xlsx":
            skipped.append((path, "not .xlsx"))
            continue

        source_files.append(path)

    return source_files, skipped


def process_empirical_sheet(
    sheet: xw.Sheet,
    label: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    matrix = load_sheet_matrix(sheet)
    anchor = find_max_anchor(matrix)
    if anchor is None:
        print(f"  - {EMPIRICAL_SHEET_NAME}: skipped (no 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7

    min_col = find_col_by_patterns(
        matrix, anchor_row, anchor_col, pattern_sets=[["min"]]
    ) or (anchor_col + 1)
    forecast_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["estimated", "sold"], ["tot", "fcst"], ["forecast"]],
    ) or (anchor_col - 1)
    actual_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["reported", "sales"], ["actual"]],
    ) or (anchor_col - 2)
    num_quarters_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["num", "quarter"], ["n", "quarter"]],
    ) or (anchor_col - 3)
    last_quarter_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["last", "quarter"]],
    ) or (anchor_col - 4)
    growth_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["growth"]],
    ) or (anchor_col - 5)
    captured_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["captured"], ["db"]],
    ) or (anchor_col - 6)

    quarterly_sales_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["quarterly", "sales"]],
    ) or x_col
    reported_sales_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        pattern_sets=[["reported", "sales"]],
    ) or y_col

    formulas: List[str] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row, end_row = history_bounds(anchor_row, n_quarters)
        if x_col < 1 or y_col < 1 or start_row > end_row:
            formulas.append('=""')
            continue
        # R1C1 formula in formula2 to compute avg penetration quickly.
        formulas.append(
            f'=IFERROR(SUM(R{start_row}C{y_col}:R{end_row}C{y_col})/'
            f'SUM(R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )

    avg_pen_values = calculate_formulas(sheet, formulas, temp_col=16384, temp_row_start=2)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        n_default = idx + 1
        row = anchor_row + n_default

        num_quarters_val = to_float(matrix_get(matrix, row, num_quarters_col))
        num_quarters_used = int(round(num_quarters_val)) if num_quarters_val is not None else n_default
        num_quarters_used = max(1, num_quarters_used)

        start_row, end_row = history_bounds(anchor_row, num_quarters_used)
        avg_penetration_pct = to_float(avg_pen_values[idx])

        forecast_value = to_float(matrix_get(matrix, row, forecast_col))
        actual_value = to_float(matrix_get(matrix, row, actual_col))
        forecast_max = to_float(matrix_get(matrix, row, anchor_col))
        forecast_min = to_float(matrix_get(matrix, row, min_col))

        quarterly_sales = column_sum(matrix, start_row, end_row, quarterly_sales_col)
        reported_sales = column_sum(matrix, start_row, end_row, reported_sales_col)

        if forecast_value is None and avg_penetration_pct is not None and quarterly_sales is not None:
            forecast_value = avg_penetration_pct * quarterly_sales
        if actual_value is None:
            actual_value = reported_sales
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * 1.05
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value * 0.95

        growth_rate_pct = to_float(matrix_get(matrix, row, growth_col))
        if growth_rate_pct is None:
            first_q, last_q = first_last_numeric(matrix, start_row, end_row, quarterly_sales_col)
            if first_q not in (None, 0) and last_q is not None:
                growth_rate_pct = (last_q / first_q) - 1.0

        sales_captured_in_db_pct = to_float(matrix_get(matrix, row, captured_col))
        if (
            sales_captured_in_db_pct is None
            and quarterly_sales not in (None, 0)
            and reported_sales is not None
        ):
            sales_captured_in_db_pct = reported_sales / quarterly_sales

        last_quarter_used = matrix_get(matrix, row, last_quarter_col)
        if last_quarter_used in (None, "") and (quarterly_sales_col - 1) >= 1:
            last_quarter_used = matrix_get(matrix, end_row, quarterly_sales_col - 1)

        meaningful = any(
            value is not None
            for value in (
                avg_penetration_pct,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            )
        )
        if not meaningful:
            continue

        rows.append(
            {
                "model": label["model"],
                "ticker": label["ticker"],
                "model_period": label["model_period"],
                "model_date": label["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_diff(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    sheet: xw.Sheet,
    label: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    matrix = load_sheet_matrix(sheet)
    anchor = find_max_anchor(matrix)
    if anchor is None:
        print(f"  - {REGRESSION_SHEET_NAME}: skipped (no 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    min_col = find_col_by_patterns(matrix, anchor_row, anchor_col, [["min"]]) or (anchor_col + 1)
    forecast_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        [["tot", "fcst", "w/o"], ["tot", "fcst"], ["forecast"]],
    ) or (anchor_col - 1)
    actual_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        [["actual"], ["reported", "sales"]],
    )
    num_col = find_col_by_patterns(
        matrix,
        anchor_row,
        anchor_col,
        [["num", "quarter"], ["n", "quarter"]],
    ) or (anchor_col - 3)

    intercept_formulas: List[str] = []
    slope_formulas: List[str] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row, end_row = history_bounds(anchor_row, n_quarters)
        if x_col < 1 or y_col < 1 or start_row > end_row:
            intercept_formulas.append('=""')
            slope_formulas.append('=""')
            continue

        intercept_formulas.append(
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_formulas.append(
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )

    intercept_vals = calculate_formulas(sheet, intercept_formulas, temp_col=16383, temp_row_start=2)
    slope_vals = calculate_formulas(sheet, slope_formulas, temp_col=16384, temp_row_start=2)
    latest_x = to_float(matrix_get(matrix, anchor_row - 1, x_col))

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(N_QUARTERS):
        n_default = idx + 1
        row = anchor_row + n_default

        num_q = to_float(matrix_get(matrix, row, num_col))
        num_quarters_used = int(round(num_q)) if num_q is not None else n_default
        num_quarters_used = max(1, num_quarters_used)

        intercept = to_float(intercept_vals[idx])
        slope = to_float(slope_vals[idx])

        forecast_value = to_float(matrix_get(matrix, row, forecast_col))
        if forecast_value is None and None not in (intercept, slope, latest_x):
            forecast_value = intercept + slope * latest_x

        forecast_max = to_float(matrix_get(matrix, row, anchor_col))
        forecast_min = to_float(matrix_get(matrix, row, min_col))

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * 1.05
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value * 0.95

        actual_value = to_float(matrix_get(matrix, row, actual_col)) if actual_col else None

        meaningful = any(
            value is not None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        )
        if not meaningful:
            continue

        signature = (
            num_quarters_used,
            round_or_none(forecast_value),
            round_or_none(forecast_max),
            round_or_none(forecast_min),
            round_or_none(intercept),
            round_or_none(slope),
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": label["model"],
                "ticker": label["ticker"],
                "model_period": label["model_period"],
                "model_date": label["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_diff(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_rows_to_sheet(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    max_widths = [len(header) for header in headers]
    for values in ws.iter_rows(min_row=2, max_col=len(headers), max_row=ws.max_row, values_only=True):
        for idx, value in enumerate(values):
            if value is None:
                continue
            text = f"{value:.6g}" if isinstance(value, float) else str(value)
            max_widths[idx] = min(60, max(max_widths[idx], len(text)))

    for idx, width in enumerate(max_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = max(12, width + 2)


def save_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_rows_to_sheet(ws_empirical, EMPIRICAL_HEADERS, empirical_rows)
    write_rows_to_sheet(ws_regression, REGRESSION_HEADERS, regression_rows)
    wb.save(output_path)


def main() -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    source_files, skipped_files = list_source_files(input_dir)

    for path, reason in skipped_files:
        print(f"Skipped file: {path.name} ({reason})")

    if not input_dir.exists():
        print(f"Input directory does not exist: {input_dir.resolve()}")

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

        for file_path in source_files:
            print(f"Processing file: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path)

                sheet_names = {sheet.name for sheet in wb.sheets}
                if EMPIRICAL_SHEET_NAME in sheet_names:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            sheet=wb.sheets[EMPIRICAL_SHEET_NAME],
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"  - {EMPIRICAL_SHEET_NAME}: skipped (sheet not found)")

                if REGRESSION_SHEET_NAME in sheet_names:
                    regression_rows.extend(
                        process_regression_sheet(
                            sheet=wb.sheets[REGRESSION_SHEET_NAME],
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"  - {REGRESSION_SHEET_NAME}: skipped (sheet not found)")

                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_path = find_next_output_path(input_dir=input_dir, output_folder=output_dir)
    save_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path.resolve()}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
