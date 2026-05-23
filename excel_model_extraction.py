#!/usr/bin/env python3
"""Extract empirical and regression candidate rows from Excel model files."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError:  # pragma: no cover - allows importing file without Excel deps
    xw = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover - allows importing file without writer deps
    Workbook = None
    Font = None
    get_column_letter = None


# -------------------------
# User-configurable paths
# -------------------------
input_dir = "./input"
output_dir = "./output"


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

MONTH_TO_NUMBER = {
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


def to_matrix(values: Any) -> List[List[Any]]:
    """Normalize xlwings used_range values into a 2D list."""
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]

    if values and not isinstance(values[0], (list, tuple)):
        return [list(values)]

    rows = [list(row) for row in values]
    if not rows:
        return rows
    max_cols = max(len(row) for row in rows)
    return [row + [None] * (max_cols - len(row)) for row in rows]


def as_number(value: Any) -> Optional[float]:
    """Convert cell values to float when possible."""
    if value is None:
        return None
    if isinstance(value, bool):
        return float(int(value))
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
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
    return None


def normalize_text(value: Any) -> str:
    return str(value).strip().lower() if value is not None else ""


def get_matrix_value(
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
    if col_idx >= len(matrix[row_idx]):
        return None
    return matrix[row_idx][col_idx]


def find_anchor_cell(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    label: str = "max",
) -> Optional[Tuple[int, int]]:
    target = label.strip().lower()
    for r_idx, row_values in enumerate(matrix):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                return base_row + r_idx, base_col + c_idx
    return None


def find_nearest_label_cell(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    pattern: str,
    anchor_row: int,
    anchor_col: int,
) -> Optional[Tuple[int, int]]:
    regex = re.compile(pattern, re.IGNORECASE)
    best: Optional[Tuple[int, int]] = None
    best_score: Optional[int] = None

    for r_idx, row_values in enumerate(matrix):
        abs_row = base_row + r_idx
        for c_idx, cell_value in enumerate(row_values):
            if not isinstance(cell_value, str):
                continue
            if not regex.search(cell_value.strip()):
                continue
            abs_col = base_col + c_idx
            score = abs(abs_col - anchor_col) * 5 + abs(abs_row - anchor_row)
            if best_score is None or score < best_score:
                best_score = score
                best = (abs_row, abs_col)

    return best


def infer_candidate_columns(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    anchor_row: int,
    anchor_col: int,
    count: int = N_QUARTERS,
) -> List[int]:
    """Infer candidate output columns from numeric values near the max anchor row."""
    row_candidates: List[int] = []
    source_rows = [anchor_row, anchor_row - 1, anchor_row + 1]

    for source_row in source_rows:
        row_idx = source_row - base_row
        if row_idx < 0 or row_idx >= len(matrix):
            continue
        for c_idx, value in enumerate(matrix[row_idx]):
            abs_col = base_col + c_idx
            if abs_col <= anchor_col:
                continue
            if as_number(value) is not None:
                row_candidates.append(abs_col)

    unique_right = sorted(set(row_candidates))
    if len(unique_right) >= count:
        return unique_right[:count]

    columns = list(unique_right)
    next_col = anchor_col + 1
    while len(columns) < count:
        if next_col not in columns:
            columns.append(next_col)
        next_col += 1
    return columns


def parse_model_metadata(file_name: str) -> Dict[str, str]:
    """Parse ticker/model period/date from file name."""
    stem = Path(file_name).stem
    pieces = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(pieces) >= 2:
        ticker = pieces[1].upper()
    else:
        match = re.search(r"\b([A-Z]{1,8})\b", stem)
        if match:
            ticker = match.group(1).upper()

    period_source = ""
    if len(pieces) >= 3:
        period_source = pieces[2].split("_")[0].strip()

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        period_source or stem,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase = period_match.group(1).title()
        month = period_match.group(2).title()
        year = int(period_match.group(3))
        model_period = f"{phase}{month}_{year}"
        model_day = PERIOD_DAY[phase.lower()]
        month_num = MONTH_TO_NUMBER[month.lower()]
        model_date = datetime(year, month_num, model_day).strftime("%Y-%m-%d")

    model = "_".join(part for part in (ticker, model_period) if part)
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "source_file": file_name,
    }


def choose_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1

    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1

    return candidate


def safe_close_workbook(workbook: Any) -> None:
    """Close source workbook without saving, with fallbacks."""
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass


def set_formula2(cell: Any, formula_r1c1: str) -> None:
    """Prefer formula2, but gracefully fall back to formula when needed."""
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def subtract_if_numeric(left: Any, right: Any) -> Any:
    left_num = as_number(left)
    right_num = as_number(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def int_or_default(value: Any, default: int) -> int:
    numeric_value = as_number(value)
    if numeric_value is None:
        return default
    try:
        return max(1, int(round(numeric_value)))
    except Exception:
        return default


def find_last_numeric_row(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    row_limit: int,
    x_col: int,
    y_col: int,
) -> int:
    for row in range(row_limit, base_row - 1, -1):
        x_val = as_number(get_matrix_value(matrix, base_row, base_col, row, x_col))
        y_val = as_number(get_matrix_value(matrix, base_row, base_col, row, y_col))
        if x_val is not None and y_val is not None:
            return row
    return row_limit


def near_equal(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    a_num = as_number(a)
    b_num = as_number(b)
    if a_num is None or b_num is None:
        return normalize_text(a) == normalize_text(b)
    scale = max(1.0, abs(a_num), abs(b_num))
    return abs(a_num - b_num) <= tolerance * scale


def is_duplicate_regression_row(previous: Dict[str, Any], current: Dict[str, Any]) -> bool:
    keys = ["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    return all(near_equal(previous.get(key), current.get(key)) for key in keys)


def process_empirical_sheet(workbook: Any, metadata: Dict[str, str]) -> List[Dict[str, Any]]:
    sheet_name = "Empirical Model"
    try:
        sheet = workbook.sheets[sheet_name]
    except Exception:
        print(f"Skipped empirical extraction: sheet '{sheet_name}' not found.")
        return []

    used = sheet.used_range
    matrix = to_matrix(used.value)
    if not matrix:
        print("Skipped empirical extraction: sheet is empty.")
        return []

    base_row, base_col = used.row, used.column
    anchor = find_anchor_cell(matrix, base_row, base_col, label="max")
    if anchor is None:
        print("Skipped empirical extraction: max anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    candidate_cols = infer_candidate_columns(
        matrix=matrix,
        base_row=base_row,
        base_col=base_col,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        count=N_QUARTERS,
    )

    min_cell = find_nearest_label_cell(
        matrix, base_row, base_col, r"\bmin\b|minimum", anchor_row, anchor_col
    )
    forecast_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"estimated.*total.*sold|forecast",
        anchor_row,
        anchor_col,
    )
    actual_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"reported.*sales|actual",
        anchor_row,
        anchor_col,
    )
    num_quarters_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"num.*quarters|n.?quarters|quarters.*used",
        anchor_row,
        anchor_col,
    )
    last_quarter_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"last.*quarter",
        anchor_row,
        anchor_col,
    )
    avg_pen_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"avg.*penetration|average.*penetration|penetration.*avg",
        anchor_row,
        anchor_col,
    )
    quarterly_sales_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"quarterly.*sales",
        anchor_row,
        anchor_col,
    )
    growth_rate_cell = find_nearest_label_cell(
        matrix, base_row, base_col, r"growth", anchor_row, anchor_col
    )
    captured_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"sales.*captured.*db|captured.*db",
        anchor_row,
        anchor_col,
    )

    min_row = min_cell[0] if min_cell else anchor_row + 1
    forecast_row = forecast_cell[0] if forecast_cell else anchor_row - 1
    actual_row = actual_cell[0] if actual_cell else anchor_row - 2
    num_quarters_row = num_quarters_cell[0] if num_quarters_cell else anchor_row - 3
    last_quarter_row = last_quarter_cell[0] if last_quarter_cell else anchor_row - 4
    avg_pen_row = avg_pen_cell[0] if avg_pen_cell else None
    quarterly_sales_row = quarterly_sales_cell[0] if quarterly_sales_cell else None
    growth_rate_row = growth_rate_cell[0] if growth_rate_cell else None
    captured_row = captured_cell[0] if captured_cell else None

    helper_col = max(candidate_cols) + 2
    helper_start_row = anchor_row + 1
    avg_source_row = avg_pen_row or captured_row or quarterly_sales_row

    pending_rows: List[Dict[str, Any]] = []
    helper_cells: List[Any] = []

    for index, col in enumerate(candidate_cols, start=1):
        num_quarters_value = get_matrix_value(matrix, base_row, base_col, num_quarters_row, col)
        num_quarters_used = int_or_default(num_quarters_value, index)

        helper_cell = None
        if avg_source_row is not None:
            start_index = max(0, index - num_quarters_used)
            start_col = candidate_cols[start_index]
            helper_cell = sheet.cells(helper_start_row + index - 1, helper_col)
            avg_formula = (
                f"=AVERAGE(R{avg_source_row}C{start_col}:R{avg_source_row}C{col})"
            )
            set_formula2(helper_cell, avg_formula)
            helper_cells.append(helper_cell)

        pending_rows.append(
            {
                "column": col,
                "num_quarters_used": num_quarters_used,
                "helper_cell": helper_cell,
            }
        )

    if helper_cells:
        workbook.app.calculate()

    output_rows: List[Dict[str, Any]] = []
    for entry in pending_rows:
        col = int(entry["column"])
        num_quarters_used = int(entry["num_quarters_used"])
        helper_cell = entry["helper_cell"]

        forecast_value = get_matrix_value(matrix, base_row, base_col, forecast_row, col)
        actual_value = get_matrix_value(matrix, base_row, base_col, actual_row, col)
        forecast_max = get_matrix_value(matrix, base_row, base_col, anchor_row, col)
        forecast_min = get_matrix_value(matrix, base_row, base_col, min_row, col)
        last_quarter_used = get_matrix_value(
            matrix, base_row, base_col, last_quarter_row, col
        )
        quarterly_sales = (
            get_matrix_value(matrix, base_row, base_col, quarterly_sales_row, col)
            if quarterly_sales_row is not None
            else ""
        )
        growth_rate_pct = (
            get_matrix_value(matrix, base_row, base_col, growth_rate_row, col)
            if growth_rate_row is not None
            else ""
        )
        sales_captured_in_db_pct = (
            get_matrix_value(matrix, base_row, base_col, captured_row, col)
            if captured_row is not None
            else ""
        )
        avg_penetration_pct = helper_cell.value if helper_cell is not None else ""
        if avg_penetration_pct in (None, "") and avg_pen_row is not None:
            avg_penetration_pct = get_matrix_value(
                matrix, base_row, base_col, avg_pen_row, col
            )

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_if_numeric(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": actual_value,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": metadata["source_file"],
        }
        output_rows.append(row)

    return output_rows


def process_regression_sheet(workbook: Any, metadata: Dict[str, str]) -> List[Dict[str, Any]]:
    sheet_name = "Regression Model"
    try:
        sheet = workbook.sheets[sheet_name]
    except Exception:
        print(f"Skipped regression extraction: sheet '{sheet_name}' not found.")
        return []

    used = sheet.used_range
    matrix = to_matrix(used.value)
    if not matrix:
        print("Skipped regression extraction: sheet is empty.")
        return []

    base_row, base_col = used.row, used.column
    anchor = find_anchor_cell(matrix, base_row, base_col, label="max")
    if anchor is None:
        print("Skipped regression extraction: max anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    candidate_cols = infer_candidate_columns(
        matrix=matrix,
        base_row=base_row,
        base_col=base_col,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        count=N_QUARTERS,
    )

    min_cell = find_nearest_label_cell(
        matrix, base_row, base_col, r"\bmin\b|minimum", anchor_row, anchor_col
    )
    forecast_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"tot.*fcst.*w/?o.*sa|total.*forecast.*without.*sa",
        anchor_row,
        anchor_col,
    )
    num_quarters_cell = find_nearest_label_cell(
        matrix,
        base_row,
        base_col,
        r"num.*quarters|n.?quarters|quarters.*used",
        anchor_row,
        anchor_col,
    )
    actual_cell = find_nearest_label_cell(
        matrix, base_row, base_col, r"actual|reported", anchor_row, anchor_col
    )

    min_row = min_cell[0] if min_cell else anchor_row + 1
    forecast_row = forecast_cell[0] if forecast_cell else anchor_row - 1
    num_quarters_row = num_quarters_cell[0] if num_quarters_cell else anchor_row - 2
    actual_row = actual_cell[0] if actual_cell else None

    history_end_row = find_last_numeric_row(
        matrix=matrix,
        base_row=base_row,
        base_col=base_col,
        row_limit=max(base_row, anchor_row - 1),
        x_col=x_col,
        y_col=y_col,
    )
    latest_x = as_number(get_matrix_value(matrix, base_row, base_col, history_end_row, x_col))

    helper_intercept_col = max(candidate_cols) + 2
    helper_slope_col = helper_intercept_col + 1
    helper_start_row = anchor_row + 1

    pending_rows: List[Dict[str, Any]] = []
    for index, col in enumerate(candidate_cols, start=1):
        num_quarters_value = get_matrix_value(matrix, base_row, base_col, num_quarters_row, col)
        num_quarters_used = int_or_default(num_quarters_value, index)
        regression_window = max(2, num_quarters_used)
        start_row = max(base_row, history_end_row - regression_window + 1)

        intercept_cell = sheet.cells(helper_start_row + index - 1, helper_intercept_col)
        slope_cell = sheet.cells(helper_start_row + index - 1, helper_slope_col)

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)

        pending_rows.append(
            {
                "column": col,
                "num_quarters_used": num_quarters_used,
                "intercept_cell": intercept_cell,
                "slope_cell": slope_cell,
            }
        )

    if pending_rows:
        workbook.app.calculate()

    output_rows: List[Dict[str, Any]] = []
    for entry in pending_rows:
        col = int(entry["column"])
        num_quarters_used = int(entry["num_quarters_used"])
        intercept_value = entry["intercept_cell"].value
        slope_value = entry["slope_cell"].value

        forecast_value = get_matrix_value(matrix, base_row, base_col, forecast_row, col)
        if forecast_value in (None, "") and latest_x is not None:
            intercept_num = as_number(intercept_value)
            slope_num = as_number(slope_value)
            if intercept_num is not None and slope_num is not None:
                forecast_value = intercept_num + slope_num * latest_x

        actual_value: Any = ""
        if actual_row is not None:
            actual_value = get_matrix_value(matrix, base_row, base_col, actual_row, col)

        forecast_max = get_matrix_value(matrix, base_row, base_col, anchor_row, col)
        forecast_min = get_matrix_value(matrix, base_row, base_col, min_row, col)

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_if_numeric(forecast_max, forecast_min),
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": metadata["source_file"],
        }

        if output_rows and is_duplicate_regression_row(output_rows[-1], row):
            continue
        output_rows.append(row)

    return output_rows


def format_output_sheet(sheet: Any, columns: Sequence[str]) -> None:
    if Font is None or get_column_letter is None:
        raise RuntimeError("openpyxl is required to format output sheets.")

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_length = len(col_name)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(
            60, max(12, max_length + 2)
        )


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(col, "") for col in EMPIRICAL_COLUMNS])
    format_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS)

    regression_sheet = workbook.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(col, "") for col in REGRESSION_COLUMNS])
    format_output_sheet(regression_sheet, REGRESSION_COLUMNS)

    workbook.save(output_file)


def collect_source_files(input_path: Path) -> List[Path]:
    if not input_path.exists() or not input_path.is_dir():
        print(f"Input directory missing or invalid: {input_path}")
        return []

    source_files: List[Path] = []
    for item in sorted(input_path.iterdir(), key=lambda path: path.name.lower()):
        if not item.is_file():
            print(f"Skipped file: {item.name} (reason: not a file)")
            continue
        if item.name.startswith("~"):
            print(f"Skipped file: {item.name} (reason: temporary file)")
            continue
        if item.suffix.lower() != ".xlsx":
            print(f"Skipped file: {item.name} (reason: not .xlsx)")
            continue
        source_files.append(item)

    return source_files


def run_extraction() -> None:
    if xw is None:
        raise RuntimeError("xlwings is required to run this script.")
    if Workbook is None:
        raise RuntimeError("openpyxl is required to run this script.")

    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    source_files = collect_source_files(input_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing file: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path.name)

                empirical_rows.extend(process_empirical_sheet(workbook, metadata))
                regression_rows.extend(process_regression_sheet(workbook, metadata))
                processed_count += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (reason: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    output_file = choose_output_path(input_path, output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run_extraction()
