#!/usr/bin/env python3
"""
Build empirical/regression model-candidate rows from .xlsx source workbooks.

Design goals:
- fast batch runtime
- open each source workbook only once
- process both target sheets while workbook is open
- never save or modify source files on disk
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


# ============================================================================
# User-configurable paths
# ============================================================================
input_dir = Path("./input")
output_dir = Path("./output")


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

PERIOD_DAY = {"EARLY": 5, "MID": 15, "LATE": 25}
MONTH_NUM = {
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


def parse_filename_metadata(file_path: Path) -> Dict[str, str]:
    """Parse model labels from file names like `... - AORT - MidJan2026_...xlsx`."""
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]
    ticker = parts[1].strip() if len(parts) >= 2 else ""

    token_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if token_match:
        cadence = token_match.group(1).title()
        month = token_match.group(2).title()
        year = token_match.group(3)
        model_period = f"{cadence}{month}_{year}"
        model_date = date(
            int(year),
            MONTH_NUM[month.upper()],
            PERIOD_DAY[cadence.upper()],
        ).isoformat()

    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    elif model_period:
        model = model_period
    else:
        model = stem

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []

    first = values[0]
    if isinstance(first, (list, tuple)):
        out: List[List[Any]] = []
        for row in values:
            if isinstance(row, (list, tuple)):
                out.append(list(row))
            else:
                out.append([row])
        return out
    return [list(values)]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


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


def collect_text_cells(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
) -> List[Tuple[int, int, str]]:
    text_cells: List[Tuple[int, int, str]] = []
    for r_offset, row_vals in enumerate(matrix):
        row_no = base_row + r_offset
        for c_offset, val in enumerate(row_vals):
            text = normalize_text(val)
            if text:
                text_cells.append((row_no, base_col + c_offset, text))
    return text_cells


def find_anchor(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    target: str = "max",
) -> Optional[Tuple[int, int]]:
    target_norm = normalize_text(target)
    for r_offset, row_vals in enumerate(matrix):
        for c_offset, val in enumerate(row_vals):
            if normalize_text(val) == target_norm:
                return base_row + r_offset, base_col + c_offset
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

    for row_no, col_no, text in text_cells:
        if abs(row_no - anchor_row) > max_row_distance:
            continue
        for group in keyword_groups:
            if all(token in text for token in group):
                score = abs(row_no - anchor_row) * 100 + abs(col_no - anchor_col)
                if best_score is None or score < best_score:
                    best_score = score
                    best_col = col_no
                break

    return best_col if best_col is not None else fallback_col


def safe_cell_read(sheet: xw.main.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1 or row < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


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

    for row_no in range(start_row, end_row + 1):
        ok = True
        for col_no in required_cols:
            if to_float(matrix_value(matrix, base_row, base_col, row_no, col_no)) is None:
                ok = False
                break
        if ok:
            rows.append(row_no)
    return rows


def subtract_if_numeric(a: Any, b: Any) -> Any:
    af = to_float(a)
    bf = to_float(b)
    if af is None or bf is None:
        return ""
    return af - bf


def set_formula2(cell: xw.main.Range, formula_r1c1: str) -> None:
    """Set R1C1 formula, preferring .formula2."""
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def close_book_without_saving(wb: xw.Book) -> None:
    """Close source workbook with no save, including fallback paths."""
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(SaveChanges=False)
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
            keyword_groups=(("quarterly", "sales"), ("captured", "sales"), ("db", "sales")),
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
            keyword_groups=(("growth", "rate"), ("growth", "pct")),
            fallback_col=anchor_col - 4,
        ),
        "sales_captured_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(
                ("sales", "captured"),
                ("captured", "db"),
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
) -> Dict[str, Optional[int]]:
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
                ("tot", "fcst"),
                ("forecast", "without", "sa"),
                ("total", "without", "sa"),
            ),
            fallback_col=anchor_col - 4,
        ),
        "actual_col": find_col_by_keywords(
            text_cells,
            anchor_row,
            anchor_col,
            keyword_groups=(("actual", "sales"), ("reported", "sales")),
            fallback_col=anchor_col - 3,
        ),
        # Required explicit offsets:
        "x_col": anchor_col - 11,
        "y_col": anchor_col - 7,
    }


def get_sheet_by_name(wb: xw.Book, target_name: str) -> Optional[xw.main.Sheet]:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def extract_empirical_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_by_name(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped {source_file}: Empirical Model sheet missing")
        return rows

    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        print(f"Skipped {source_file}: Empirical Model sheet empty")
        return rows

    base_row = used.row
    base_col = used.column
    max_width = max(len(r) for r in matrix) if matrix else 1

    anchor = find_anchor(matrix, base_row, base_col, target="max")
    if anchor is None:
        print(f"Skipped {source_file}: Empirical Model max anchor not found")
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
        return rows

    helper_col = max(base_col + max_width + 2, anchor_col + 2)
    helper_forecast_col = helper_col + 1
    formulas_written = False

    # Required: n_quarters = 10 with R1C1 formula writes via formula2.
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
            f'R{out_row}C{helper_col},"")'
        )

        set_formula2(sheet.range((out_row, helper_col)), avg_formula)
        set_formula2(sheet.range((out_row, helper_forecast_col)), forecast_formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    last_history_row = history_rows[-1]
    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters

        avg_pen = safe_cell_read(sheet, out_row, helper_col)
        forecast_value = safe_cell_read(sheet, out_row, cols["forecast_col"])
        if forecast_value in (None, ""):
            forecast_value = safe_cell_read(sheet, out_row, helper_forecast_col)

        forecast_max = safe_cell_read(sheet, out_row, cols["max_col"])
        forecast_min = safe_cell_read(sheet, out_row, cols["min_col"])

        quarterly_sales = safe_cell_read(sheet, out_row, cols["quarterly_sales_col"])
        if quarterly_sales in (None, ""):
            quarterly_sales = safe_cell_read(
                sheet, last_history_row, cols["quarterly_sales_col"]
            )

        reported_sales = safe_cell_read(sheet, out_row, cols["reported_sales_col"])
        if reported_sales in (None, ""):
            reported_sales = safe_cell_read(
                sheet, last_history_row, cols["reported_sales_col"]
            )

        growth_rate = safe_cell_read(sheet, out_row, cols["growth_rate_col"])
        if growth_rate in (None, ""):
            growth_rate = safe_cell_read(sheet, last_history_row, cols["growth_rate_col"])

        sales_captured = safe_cell_read(sheet, out_row, cols["sales_captured_col"])
        if sales_captured in (None, ""):
            sales_captured = safe_cell_read(
                sheet, last_history_row, cols["sales_captured_col"]
            )

        last_quarter_used = safe_cell_read(sheet, out_row, cols["last_quarter_col"])
        if last_quarter_used in (None, ""):
            last_quarter_used = safe_cell_read(
                sheet, last_history_row, cols["quarter_label_col"]
            )

        num_quarters_used = safe_cell_read(sheet, out_row, cols["num_quarters_col"])
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        actual_value = reported_sales

        if all(
            v in (None, "")
            for v in (avg_pen, forecast_value, actual_value, forecast_max, forecast_min)
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


def _signature_for_regression_row(values: Sequence[Any]) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        number = to_float(value)
        if number is not None:
            signature.append(round(number, 10))
        else:
            signature.append("" if value in (None, "") else str(value))
    return tuple(signature)


def extract_regression_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_by_name(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped {source_file}: Regression Model sheet missing")
        return rows

    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        print(f"Skipped {source_file}: Regression Model sheet empty")
        return rows

    base_row = used.row
    base_col = used.column
    max_width = max(len(r) for r in matrix) if matrix else 1

    anchor = find_anchor(matrix, base_row, base_col, target="max")
    if anchor is None:
        print(f"Skipped {source_file}: Regression Model max anchor not found")
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
    if not history_rows:
        return rows

    helper_intercept_col = max(base_col + max_width + 2, anchor_col + 2)
    helper_slope_col = helper_intercept_col + 1
    formulas_written = False

    # Required: y_col = anchor_col - 7, x_col = anchor_col - 11.
    # Required: n_quarters = 10 and R1C1 formula2 for INTERCEPT/SLOPE.
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

        set_formula2(sheet.range((out_row, helper_intercept_col)), intercept_formula)
        set_formula2(sheet.range((out_row, helper_slope_col)), slope_formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    previous_signature: Optional[Tuple[Any, ...]] = None
    for n_quarters in range(1, N_QUARTERS + 1):
        out_row = anchor_row + n_quarters

        num_quarters_used = safe_cell_read(sheet, out_row, cols["num_quarters_col"])
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        intercept = safe_cell_read(sheet, out_row, helper_intercept_col)
        slope = safe_cell_read(sheet, out_row, helper_slope_col)
        forecast_total_wo_sa = safe_cell_read(sheet, out_row, cols["forecast_total_wo_sa_col"])
        forecast_max = safe_cell_read(sheet, out_row, cols["max_col"])
        forecast_min = safe_cell_read(sheet, out_row, cols["min_col"])
        actual_value = safe_cell_read(sheet, out_row, cols["actual_col"])
        if actual_value in (None,):
            actual_value = ""

        signature = _signature_for_regression_row(
            (
                num_quarters_used,
                intercept,
                slope,
                forecast_total_wo_sa,
                forecast_max,
                forecast_min,
            )
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        if all(v in (None, "") for v in (intercept, slope, forecast_total_wo_sa, forecast_max)):
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
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 44)


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
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

    output_name_pattern = re.compile(
        rf"^{re.escape(input_folder_name)}_PARAM(?:\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )
    if output_name_pattern.match(file_path.name):
        return "output workbook pattern"
    return None


def main() -> None:
    source_input = input_dir.expanduser().resolve()
    target_output = output_dir.expanduser().resolve()

    if not source_input.exists():
        raise SystemExit(f"Input directory does not exist: {source_input}")
    if not source_input.is_dir():
        raise SystemExit(f"Input path is not a directory: {source_input}")

    target_output.mkdir(parents=True, exist_ok=True)
    output_path = unique_output_path(source_input, target_output)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(source_input.iterdir()):
            skip_reason = should_skip_file(file_path, source_input.name)
            if skip_reason:
                print(f"Skipped: {file_path.name} ({skip_reason})")
                continue

            print(f"Processed: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path)

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_book_without_saving(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
