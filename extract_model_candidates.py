#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from .xlsx workbooks.

Runtime-focused behavior:
- one hidden Excel app for the whole run
- each source workbook is opened once
- both model sheets are processed while workbook is open
- source workbooks are always closed without saving
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover
    raise SystemExit("xlwings is required to run this script.") from exc

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configure inputs/outputs here
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
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

PERIOD_DAY_MAP = {"EARLY": 5, "MID": 15, "LATE": 25}
MONTH_NUM_MAP = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], (list, tuple)):
        normalized: List[List[Any]] = []
        for row in values:
            if isinstance(row, (list, tuple)):
                normalized.append(list(row))
            else:
                normalized.append([row])
        return normalized
    return [list(values)]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    normalized = re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()
    return re.sub(r"\s+", " ", normalized)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int_with_default(value: Any, default: int) -> int:
    parsed = to_float(value)
    if parsed is None:
        return default
    try:
        return int(parsed)
    except (TypeError, ValueError):
        return default


def subtract_if_numeric(left: Any, right: Any) -> Any:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def matrix_value(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    row: int,
    col: int,
) -> Any:
    row_idx = row - base_row
    col_idx = col - base_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(matrix):
        return None
    row_values = matrix[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def safe_cell_read(sheet: xw.main.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1 or row < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def collect_text_cells(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
) -> List[Tuple[int, int, str]]:
    text_cells: List[Tuple[int, int, str]] = []
    for row_offset, row_values in enumerate(matrix):
        row_num = base_row + row_offset
        for col_offset, value in enumerate(row_values):
            normalized = normalize_text(value)
            if not normalized:
                continue
            text_cells.append((row_num, base_col + col_offset, normalized))
    return text_cells


def find_anchor(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    target: str = "max",
) -> Optional[Tuple[int, int]]:
    target_normalized = normalize_text(target)
    for row_offset, row_values in enumerate(matrix):
        for col_offset, value in enumerate(row_values):
            normalized = normalize_text(value)
            if normalized == target_normalized or normalized.startswith(f"{target_normalized} "):
                return base_row + row_offset, base_col + col_offset
    return None


def find_col_by_keywords(
    text_cells: Iterable[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
    keyword_groups: Sequence[Sequence[str]],
    fallback_col: int,
    max_row_distance: int = 10,
) -> int:
    best_col: Optional[int] = None
    best_score: Optional[int] = None

    for row_num, col_num, text in text_cells:
        if abs(row_num - anchor_row) > max_row_distance:
            continue
        for group in keyword_groups:
            if all(keyword in text for keyword in group):
                score = abs(row_num - anchor_row) * 100 + abs(col_num - anchor_col)
                if best_score is None or score < best_score:
                    best_score = score
                    best_col = col_num
                break

    return best_col if best_col is not None else fallback_col


def collect_numeric_rows(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    start_row: int,
    end_row: int,
    required_cols: Sequence[int],
) -> List[int]:
    rows: List[int] = []
    if end_row < start_row:
        return rows

    for row_num in range(start_row, end_row + 1):
        valid = True
        for col_num in required_cols:
            if col_num < 1:
                valid = False
                break
            if to_float(matrix_value(matrix, base_row, base_col, row_num, col_num)) is None:
                valid = False
                break
        if valid:
            rows.append(row_num)

    return rows


def set_formula2_r1c1(cell: xw.main.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def close_book_without_saving(wb: xw.Book) -> None:
    """Close a source workbook without persisting any temporary formula writes."""
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
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def unique_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    root_name = f"{source_input_dir.name}_PARAM"
    candidate = target_output_dir / f"{root_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_output_dir / f"{root_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def parse_filename_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    token_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )
    if not token_match:
        raise ValueError("missing Early/Mid/Late + month + year token in filename")

    period_name = token_match.group(1).upper()
    month_name = token_match.group(2).upper()
    year = int(token_match.group(3))

    model_period = f"{period_name.title()}{month_name.title()}_{year}"
    model_date = date(year, MONTH_NUM_MAP[month_name], PERIOD_DAY_MAP[period_name]).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.main.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def infer_empirical_columns(
    text_cells: Sequence[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
) -> Dict[str, int]:
    return {
        "max_col": anchor_col,
        "min_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("min",),),
            fallback_col=anchor_col + 1,
        ),
        "num_quarters_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("num", "quarter"), ("quarters", "used")),
            fallback_col=anchor_col - 9,
        ),
        "last_quarter_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("last", "quarter"),),
            fallback_col=anchor_col - 10,
        ),
        "forecast_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(
                ("estimated", "total"),
                ("forecast", "value"),
                ("total", "sold"),
                ("tot", "fcst"),
            ),
            fallback_col=anchor_col - 6,
        ),
        "actual_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("reported", "sales"), ("actual", "sales"), ("actual",)),
            fallback_col=anchor_col - 5,
        ),
        "quarterly_sales_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("quarterly", "sales"), ("quarter", "sales"), ("qtr", "sales")),
            fallback_col=anchor_col - 7,
        ),
        "reported_sales_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("reported", "sales"),),
            fallback_col=anchor_col - 5,
        ),
        "growth_rate_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("growth", "rate"), ("growth", "pct"), ("growth",)),
            fallback_col=anchor_col - 4,
        ),
        "sales_captured_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(
                ("sales", "captured"),
                ("captured", "db"),
                ("penetration", "pct"),
                ("penetration",),
            ),
            fallback_col=anchor_col - 3,
        ),
        "quarter_label_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("quarter",),),
            fallback_col=anchor_col - 11,
        ),
    }


def infer_regression_columns(
    text_cells: Sequence[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
) -> Dict[str, int]:
    return {
        "max_col": anchor_col,
        "min_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("min",),),
            fallback_col=anchor_col + 1,
        ),
        "num_quarters_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("num", "quarter"), ("quarters", "used")),
            fallback_col=anchor_col - 9,
        ),
        "forecast_total_wo_sa_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(
                ("tot", "fcst", "w", "o", "sa"),
                ("tot", "fcst", "wo", "sa"),
                ("tot", "fcst", "without", "sa"),
                ("forecast", "without", "sa"),
                ("tot", "fcst"),
            ),
            fallback_col=anchor_col - 4,
        ),
        "actual_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("actual",), ("reported", "sales")),
            fallback_col=anchor_col - 3,
        ),
        # Explicit required offsets from max anchor:
        "x_col": anchor_col - 11,
        "y_col": anchor_col - 7,
    }


def extract_empirical_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_case_insensitive(wb, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        print(f"Skipped {source_file}: missing '{EMPIRICAL_SHEET_NAME}' sheet")
        return rows

    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        print(f"Skipped {source_file}: '{EMPIRICAL_SHEET_NAME}' sheet is empty")
        return rows

    base_row = used.row
    base_col = used.column
    max_width = max(len(row_values) for row_values in matrix) if matrix else 1

    anchor = find_anchor(matrix, base_row, base_col, target="max")
    if anchor is None:
        print(f"Skipped {source_file}: '{EMPIRICAL_SHEET_NAME}' max anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    text_cells = collect_text_cells(matrix, base_row, base_col)
    cols = infer_empirical_columns(text_cells, anchor_row, anchor_col)

    history_rows = collect_numeric_rows(
        matrix=matrix,
        base_row=base_row,
        base_col=base_col,
        start_row=base_row,
        end_row=anchor_row - 1,
        required_cols=[cols["sales_captured_col"]],
    )
    if not history_rows:
        history_rows = collect_numeric_rows(
            matrix=matrix,
            base_row=base_row,
            base_col=base_col,
            start_row=base_row,
            end_row=anchor_row - 1,
            required_cols=[cols["reported_sales_col"]],
        )
    if not history_rows:
        print(f"Skipped {source_file}: insufficient empirical history rows")
        return rows

    helper_avg_col = max(base_col + max_width + 2, anchor_col + 2)
    helper_forecast_col = helper_avg_col + 1
    formulas_written = False

    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters
        hist_end = history_rows[-1]
        hist_start = history_rows[max(0, len(history_rows) - n_quarters)]

        avg_formula = (
            f'=IFERROR(AVERAGE(R{hist_start}C{cols["sales_captured_col"]}:'
            f'R{hist_end}C{cols["sales_captured_col"]}),"")'
        )
        forecast_formula = (
            f'=IFERROR(R{out_row}C{cols["quarterly_sales_col"]}/'
            f'R{out_row}C{helper_avg_col},"")'
        )

        set_formula2_r1c1(sheet.range((out_row, helper_avg_col)), avg_formula)
        set_formula2_r1c1(sheet.range((out_row, helper_forecast_col)), forecast_formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    last_history_row = history_rows[-1]
    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters

        avg_pen = safe_cell_read(sheet, out_row, helper_avg_col)
        forecast_value = safe_cell_read(sheet, out_row, cols["forecast_col"])
        if forecast_value in (None, ""):
            forecast_value = safe_cell_read(sheet, out_row, helper_forecast_col)

        actual_value = safe_cell_read(sheet, out_row, cols["actual_col"])
        if actual_value in (None, ""):
            actual_value = safe_cell_read(sheet, last_history_row, cols["actual_col"])

        forecast_max = safe_cell_read(sheet, out_row, cols["max_col"])
        forecast_min = safe_cell_read(sheet, out_row, cols["min_col"])

        quarterly_sales = safe_cell_read(sheet, out_row, cols["quarterly_sales_col"])
        if quarterly_sales in (None, ""):
            quarterly_sales = safe_cell_read(sheet, last_history_row, cols["quarterly_sales_col"])

        reported_sales = safe_cell_read(sheet, out_row, cols["reported_sales_col"])
        if reported_sales in (None, ""):
            reported_sales = safe_cell_read(sheet, last_history_row, cols["reported_sales_col"])

        growth_rate = safe_cell_read(sheet, out_row, cols["growth_rate_col"])
        if growth_rate in (None, ""):
            growth_rate = safe_cell_read(sheet, last_history_row, cols["growth_rate_col"])

        sales_captured = safe_cell_read(sheet, out_row, cols["sales_captured_col"])
        if sales_captured in (None, ""):
            sales_captured = safe_cell_read(sheet, last_history_row, cols["sales_captured_col"])

        last_quarter_used = safe_cell_read(sheet, out_row, cols["last_quarter_col"])
        if last_quarter_used in (None, ""):
            last_quarter_used = safe_cell_read(sheet, last_history_row, cols["quarter_label_col"])

        num_quarters_raw = safe_cell_read(sheet, out_row, cols["num_quarters_col"])
        num_quarters_used = to_int_with_default(num_quarters_raw, n_quarters)

        if all(
            value in (None, "")
            for value in (avg_pen, forecast_value, actual_value, forecast_max, forecast_min)
        ):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def regression_signature(values: Sequence[Any]) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        parsed = to_float(value)
        if parsed is not None:
            signature.append(round(parsed, 10))
        else:
            signature.append("" if value in (None, "") else str(value))
    return tuple(signature)


def extract_regression_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_case_insensitive(wb, REGRESSION_SHEET_NAME)
    if sheet is None:
        print(f"Skipped {source_file}: missing '{REGRESSION_SHEET_NAME}' sheet")
        return rows

    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        print(f"Skipped {source_file}: '{REGRESSION_SHEET_NAME}' sheet is empty")
        return rows

    base_row = used.row
    base_col = used.column
    max_width = max(len(row_values) for row_values in matrix) if matrix else 1

    anchor = find_anchor(matrix, base_row, base_col, target="max")
    if anchor is None:
        print(f"Skipped {source_file}: '{REGRESSION_SHEET_NAME}' max anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    text_cells = collect_text_cells(matrix, base_row, base_col)
    cols = infer_regression_columns(text_cells, anchor_row, anchor_col)

    history_rows = collect_numeric_rows(
        matrix=matrix,
        base_row=base_row,
        base_col=base_col,
        start_row=base_row,
        end_row=anchor_row - 1,
        required_cols=[cols["x_col"], cols["y_col"]],
    )
    if len(history_rows) < 2:
        print(f"Skipped {source_file}: insufficient regression history rows")
        return rows

    helper_intercept_col = max(base_col + max_width + 2, anchor_col + 2)
    helper_slope_col = helper_intercept_col + 1
    formulas_written = False

    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters
        hist_end = history_rows[-1]
        hist_start = history_rows[max(0, len(history_rows) - n_quarters)]

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{hist_start}C{cols["y_col"]}:R{hist_end}C{cols["y_col"]},'
            f'R{hist_start}C{cols["x_col"]}:R{hist_end}C{cols["x_col"]}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{hist_start}C{cols["y_col"]}:R{hist_end}C{cols["y_col"]},'
            f'R{hist_start}C{cols["x_col"]}:R{hist_end}C{cols["x_col"]}),"")'
        )

        set_formula2_r1c1(sheet.range((out_row, helper_intercept_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((out_row, helper_slope_col)), slope_formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    previous_signature: Optional[Tuple[Any, ...]] = None
    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters

        num_quarters_raw = safe_cell_read(sheet, out_row, cols["num_quarters_col"])
        num_quarters_used = to_int_with_default(num_quarters_raw, n_quarters)

        intercept = safe_cell_read(sheet, out_row, helper_intercept_col)
        slope = safe_cell_read(sheet, out_row, helper_slope_col)
        forecast_total_wo_sa = safe_cell_read(sheet, out_row, cols["forecast_total_wo_sa_col"])
        forecast_max = safe_cell_read(sheet, out_row, cols["max_col"])
        forecast_min = safe_cell_read(sheet, out_row, cols["min_col"])

        actual_value = safe_cell_read(sheet, out_row, cols["actual_col"])
        if actual_value is None:
            actual_value = ""

        signature = regression_signature(
            (
                num_quarters_used,
                intercept,
                slope,
                forecast_total_wo_sa,
                forecast_max,
                forecast_min,
            )
        )
        # Prevent duplicate final rows by skipping repeated calculated signatures.
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        if all(value in (None, "") for value in (intercept, slope, forecast_total_wo_sa, forecast_max)):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_total_wo_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def set_column_widths(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 48)


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(columns))
    last_row = max(ws.max_row, 1)
    ws.auto_filter.ref = f"A1:{last_col}{last_row}"
    set_column_widths(ws)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def should_skip_file(file_path: Path, input_folder_name: str) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"

    output_pattern = re.compile(
        rf"^{re.escape(input_folder_name)}_PARAM(?:\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )
    if output_pattern.match(file_path.name):
        return "generated PARAM output file"
    return None


def main() -> None:
    source_input = input_dir.expanduser().resolve()
    target_output = output_dir.expanduser().resolve()

    if not source_input.exists() or not source_input.is_dir():
        raise SystemExit(f"Input directory does not exist or is not a directory: {source_input}")

    target_output.mkdir(parents=True, exist_ok=True)
    output_path = unique_output_path(source_input, target_output)

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

        for file_path in sorted(source_input.iterdir(), key=lambda path: path.name.lower()):
            skip_reason = should_skip_file(file_path, source_input.name)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            print(f"Processing {file_path.name}")

            try:
                metadata = parse_filename_metadata(file_path)
            except Exception as exc:
                print(f"Skipped {file_path.name}: filename parse failed ({exc})")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: workbook open failed ({exc})")
                continue

            try:
                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
            except Exception as exc:
                print(f"Skipped empirical extraction for {file_path.name}: {exc}")

            try:
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
            except Exception as exc:
                print(f"Skipped regression extraction for {file_path.name}: {exc}")
            finally:
                close_book_without_saving(wb)

            files_processed += 1
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
