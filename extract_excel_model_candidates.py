#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable input/output
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

N_QUARTERS = 10
EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"

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

MONTH_LOOKUP = {
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

DAY_LOOKUP = {"early": 5, "mid": 15, "late": 25}


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def read_range_2d(
    sheet: xw.Sheet,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> List[List[Any]]:
    if end_row < start_row or end_col < start_col:
        return []

    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    n_rows = end_row - start_row + 1
    n_cols = end_col - start_col + 1

    if n_rows == 1 and n_cols == 1:
        return [[values]]

    if n_rows == 1:
        if isinstance(values, (list, tuple)):
            return [list(values)]
        return [[values]]

    if n_cols == 1:
        if isinstance(values, (list, tuple)):
            return [[item] for item in values]
        return [[values]]

    if isinstance(values, (list, tuple)):
        rows: List[List[Any]] = []
        for row in values:
            if isinstance(row, (list, tuple)):
                rows.append(list(row))
            else:
                rows.append([row])
        return rows

    return [[values]]


def normalize_header(value: Any) -> str:
    if is_blank(value):
        return ""
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Tuple[int, int]:
    used = sheet.used_range
    values = read_range_2d(
        sheet,
        used.row,
        used.column,
        used.last_cell.row,
        used.last_cell.column,
    )
    anchor_norm = normalize_header(anchor_text)
    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_header(cell_value) == anchor_norm:
                return used.row + r_idx, used.column + c_idx
    raise ValueError(f'Anchor "{anchor_text}" not found')


def headers_near_anchor(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int, window: int = 30
) -> List[Tuple[str, int]]:
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window
    raw_headers = read_range_2d(sheet, anchor_row, start_col, anchor_row, end_col)
    if not raw_headers:
        return []
    results: List[Tuple[str, int]] = []
    for idx, raw_value in enumerate(raw_headers[0]):
        normalized = normalize_header(raw_value)
        if normalized:
            results.append((normalized, start_col + idx))
    return results


def resolve_column(
    headers: Sequence[Tuple[str, int]],
    aliases: Sequence[str],
    anchor_col: int,
    default_offset: Optional[int] = None,
) -> Optional[int]:
    for alias in aliases:
        alias_norm = normalize_header(alias)
        if not alias_norm:
            continue
        for header, header_col in headers:
            if header == alias_norm or alias_norm in header or header in alias_norm:
                return header_col
    if default_offset is None:
        return None
    fallback_col = anchor_col + default_offset
    return fallback_col if fallback_col >= 1 else None


def value_from_row_block(
    row_values: Sequence[Any], min_col: int, target_col: Optional[int]
) -> Any:
    if target_col is None:
        return None
    idx = target_col - min_col
    if idx < 0 or idx >= len(row_values):
        return None
    return row_values[idx]


def subtract_values(left: Any, right: Any) -> Optional[float]:
    left_n = to_float(left)
    right_n = to_float(right)
    if left_n is None or right_n is None:
        return None
    return left_n - right_n


def parse_model_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem

    ticker = ""
    period_token = ""
    match = re.search(
        r"-\s*([A-Za-z0-9]+)\s*-\s*((?:Early|Mid|Late)[A-Za-z]{3}\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if match:
        ticker = match.group(1).upper()
        period_token = match.group(2)
    else:
        parts = [piece.strip() for piece in stem.split("-")]
        if len(parts) >= 2:
            ticker = parts[1].split("_")[0].strip().upper()
        if len(parts) >= 3:
            period_token = re.split(r"[_\s]", parts[2])[0].strip()

    model_period = period_token
    model_date = ""
    period_match = re.match(
        r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if period_match:
        period_type = period_match.group(1).title()
        month_abbrev = period_match.group(2).title()
        year_value = int(period_match.group(3))

        model_period = f"{period_type}{month_abbrev}_{year_value}"
        month_value = MONTH_LOOKUP.get(month_abbrev.lower())
        if month_value:
            day_value = DAY_LOOKUP[period_type.lower()]
            model_date = date(year_value, month_value, day_value).isoformat()

    model = ticker
    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif model_period:
        model = model_period
    else:
        model = stem

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(input_path: Path, destination_dir: Path) -> Path:
    stem = f"{input_path.name}_PARAM"
    candidate = destination_dir / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        versioned = destination_dir / f"{stem}.{counter}.xlsx"
        if not versioned.exists():
            return versioned
        counter += 1


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


def get_numeric_column_history(
    sheet: xw.Sheet, col_index: Optional[int], end_row: int
) -> List[Tuple[int, float]]:
    if col_index is None or end_row < 1:
        return []
    values = read_range_2d(sheet, 1, col_index, end_row, col_index)
    output: List[Tuple[int, float]] = []
    for row_idx, row_values in enumerate(values, start=1):
        number = to_float(row_values[0] if row_values else None)
        if number is not None:
            output.append((row_idx, number))
    return output


def get_column_values(
    sheet: xw.Sheet, col_index: Optional[int], end_row: int
) -> List[Any]:
    if col_index is None or end_row < 1:
        return []
    values = read_range_2d(sheet, 1, col_index, end_row, col_index)
    return [row[0] if row else None for row in values]


def get_numeric_xy_history(
    sheet: xw.Sheet, x_col: int, y_col: int, end_row: int
) -> List[Tuple[int, float, float]]:
    if end_row < 1:
        return []
    min_col = min(x_col, y_col)
    max_col = max(x_col, y_col)
    values = read_range_2d(sheet, 1, min_col, end_row, max_col)
    output: List[Tuple[int, float, float]] = []
    for row_idx, row_values in enumerate(values, start=1):
        x_value = to_float(value_from_row_block(row_values, min_col, x_col))
        y_value = to_float(value_from_row_block(row_values, min_col, y_col))
        if x_value is not None and y_value is not None:
            output.append((row_idx, x_value, y_value))
    return output


def process_empirical_sheet(
    wb: xw.Book, labels: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET]
    except Exception:
        print(f"[skipped sheet] {source_file}: missing '{EMPIRICAL_SHEET}'")
        return []

    try:
        anchor_row, anchor_col = find_anchor(sheet, "max")
    except ValueError as exc:
        print(f"[skipped sheet] {source_file}: {exc}")
        return []

    headers = headers_near_anchor(sheet, anchor_row, anchor_col)
    col_map = {
        "num_quarters_used": resolve_column(
            headers,
            ["num quarters used", "quarters used", "num_quarters_used", "n_quarters"],
            anchor_col,
            -9,
        ),
        "last_quarter_used": resolve_column(
            headers,
            ["last quarter used", "last quarter", "last_quarter_used"],
            anchor_col,
            -8,
        ),
        "forecast_value": resolve_column(
            headers,
            ["estimated total sold", "forecast value", "tot fcst", "forecast"],
            anchor_col,
            -1,
        ),
        "actual_value": resolve_column(
            headers,
            ["actual value", "reported sales", "actual"],
            anchor_col,
            -2,
        ),
        "forecast_max": resolve_column(headers, ["max", "forecast max"], anchor_col, 0),
        "forecast_min": resolve_column(headers, ["min", "forecast min"], anchor_col, 1),
        "avg_penetration_pct": resolve_column(
            headers,
            ["avg penetration", "avg penetration pct", "avg_penetration_pct"],
            anchor_col,
            -7,
        ),
        "quarterly_sales": resolve_column(
            headers,
            ["quarterly sales", "quarter sales", "quarterly_sales"],
            anchor_col,
            -6,
        ),
        "reported_sales": resolve_column(
            headers, ["reported sales", "reported_sales"], anchor_col, -5
        ),
        "growth_rate_pct": resolve_column(
            headers, ["growth rate", "growth rate pct", "growth_rate_pct"], anchor_col, -4
        ),
        "sales_captured_in_db_pct": resolve_column(
            headers,
            ["sales captured in db", "sales_captured_in_db_pct", "db capture"],
            anchor_col,
            -3,
        ),
    }

    penetration_source_col = resolve_column(
        headers,
        ["penetration", "penetration pct", "penetration rate"],
        anchor_col,
        -11,
    )
    quarter_label_col = resolve_column(
        headers,
        ["quarter", "quarter label", "quarter used"],
        anchor_col,
        -12,
    )

    candidate_start = anchor_row + 1
    candidate_end = candidate_start + N_QUARTERS - 1

    used_columns = [col for col in col_map.values() if col is not None]
    min_col = min(used_columns + [anchor_col])
    max_col = max(used_columns + [anchor_col + 1])
    candidate_block = read_range_2d(sheet, candidate_start, min_col, candidate_end, max_col)

    penetration_history = get_numeric_column_history(
        sheet, penetration_source_col, anchor_row - 1
    )
    quarter_labels = get_column_values(sheet, quarter_label_col, anchor_row - 1)

    temp_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 20)
    formula_rows: List[int] = []

    for idx in range(N_QUARTERS):
        row_number = candidate_start + idx
        row_values = candidate_block[idx] if idx < len(candidate_block) else []
        row_quarters = to_float(
            value_from_row_block(row_values, min_col, col_map["num_quarters_used"])
        )
        n_used = int(row_quarters) if row_quarters and row_quarters > 0 else idx + 1
        if penetration_history and n_used <= len(penetration_history):
            start_hist_row = penetration_history[-n_used][0]
            end_hist_row = penetration_history[-1][0]
            sheet.cells[row_number, temp_col].formula2 = (
                f"=AVERAGE(R{start_hist_row}C{penetration_source_col}:"
                f"R{end_hist_row}C{penetration_source_col})"
            )
            formula_rows.append(row_number)

    avg_formula_values: Dict[int, Any] = {}
    if formula_rows:
        wb.app.calculate()
        formula_block = read_range_2d(sheet, candidate_start, temp_col, candidate_end, temp_col)
        for idx, formula_row in enumerate(range(candidate_start, candidate_end + 1)):
            value = formula_block[idx][0] if idx < len(formula_block) and formula_block[idx] else None
            avg_formula_values[formula_row] = value

    rows: List[Dict[str, Any]] = []
    blank_streak = 0
    for idx in range(N_QUARTERS):
        row_number = candidate_start + idx
        row_values = candidate_block[idx] if idx < len(candidate_block) else []

        raw_num_quarters = value_from_row_block(
            row_values, min_col, col_map["num_quarters_used"]
        )
        num_quarters_float = to_float(raw_num_quarters)
        num_quarters_used = int(num_quarters_float) if num_quarters_float else idx + 1

        avg_penetration_sheet = value_from_row_block(
            row_values, min_col, col_map["avg_penetration_pct"]
        )
        avg_penetration = (
            avg_penetration_sheet
            if not is_blank(avg_penetration_sheet)
            else avg_formula_values.get(row_number)
        )

        forecast_value = value_from_row_block(row_values, min_col, col_map["forecast_value"])
        actual_value = value_from_row_block(row_values, min_col, col_map["actual_value"])
        forecast_max = value_from_row_block(row_values, min_col, col_map["forecast_max"])
        forecast_min = value_from_row_block(row_values, min_col, col_map["forecast_min"])

        if all(
            is_blank(value)
            for value in [avg_penetration, forecast_value, forecast_max, forecast_min, actual_value]
        ):
            blank_streak += 1
            if blank_streak >= 3:
                break
            continue
        blank_streak = 0

        last_quarter_used = value_from_row_block(
            row_values, min_col, col_map["last_quarter_used"]
        )
        if is_blank(last_quarter_used) and quarter_labels:
            if num_quarters_used <= len(quarter_labels):
                last_quarter_used = quarter_labels[-1]

        row_payload = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_values(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": value_from_row_block(
                row_values, min_col, col_map["quarterly_sales"]
            ),
            "reported_sales": value_from_row_block(
                row_values, min_col, col_map["reported_sales"]
            ),
            "growth_rate_pct": value_from_row_block(
                row_values, min_col, col_map["growth_rate_pct"]
            ),
            "sales_captured_in_db_pct": value_from_row_block(
                row_values, min_col, col_map["sales_captured_in_db_pct"]
            ),
            "source_file": source_file,
        }
        rows.append(row_payload)

    return rows


def is_duplicate_regression_row(
    previous_row: Dict[str, Any], current_row: Dict[str, Any]
) -> bool:
    numeric_fields = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    for key in numeric_fields:
        prev_value = to_float(previous_row.get(key))
        curr_value = to_float(current_row.get(key))
        if prev_value is None and curr_value is None:
            continue
        if prev_value is None or curr_value is None:
            return False
        if abs(prev_value - curr_value) > 1e-10:
            return False
    return True


def process_regression_sheet(
    wb: xw.Book, labels: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SHEET]
    except Exception:
        print(f"[skipped sheet] {source_file}: missing '{REGRESSION_SHEET}'")
        return []

    try:
        anchor_row, anchor_col = find_anchor(sheet, "max")
    except ValueError as exc:
        print(f"[skipped sheet] {source_file}: {exc}")
        return []

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print(
            f"[skipped sheet] {source_file}: invalid regression offsets from max anchor "
            f"(anchor_col={anchor_col})"
        )
        return []

    headers = headers_near_anchor(sheet, anchor_row, anchor_col)
    col_map = {
        "num_quarters_used": resolve_column(
            headers,
            ["num quarters used", "quarters used", "num_quarters_used", "n_quarters"],
            anchor_col,
            -9,
        ),
        "forecast_value": resolve_column(
            headers,
            [
                "tot fcst w/o sa",
                "tot fcst wo sa",
                "forecast total without sa",
                "forecast value",
            ],
            anchor_col,
            -1,
        ),
        "actual_value": resolve_column(
            headers,
            ["actual value", "reported sales", "actual"],
            anchor_col,
            None,
        ),
        "forecast_max": resolve_column(headers, ["max", "forecast max"], anchor_col, 0),
        "forecast_min": resolve_column(headers, ["min", "forecast min"], anchor_col, 1),
    }

    candidate_start = anchor_row + 1
    candidate_end = candidate_start + N_QUARTERS - 1

    used_columns = [col for col in col_map.values() if col is not None]
    min_col = min(used_columns + [anchor_col])
    max_col = max(used_columns + [anchor_col + 1])
    candidate_block = read_range_2d(sheet, candidate_start, min_col, candidate_end, max_col)

    xy_history = get_numeric_xy_history(sheet, x_col, y_col, anchor_row - 1)

    temp_intercept_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 25)
    temp_slope_col = temp_intercept_col + 1
    formula_meta: Dict[int, Tuple[int, float]] = {}

    for idx in range(N_QUARTERS):
        row_number = candidate_start + idx
        row_values = candidate_block[idx] if idx < len(candidate_block) else []
        raw_num_quarters = value_from_row_block(
            row_values, min_col, col_map["num_quarters_used"]
        )
        num_quarters_float = to_float(raw_num_quarters)
        num_quarters_used = int(num_quarters_float) if num_quarters_float else idx + 1

        if len(xy_history) < max(num_quarters_used, 2):
            continue

        history_slice = xy_history[-num_quarters_used:]
        start_hist_row = history_slice[0][0]
        end_hist_row = history_slice[-1][0]
        next_x_value = history_slice[-1][1] + 1

        intercept_formula = (
            f"=INTERCEPT(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
        )

        sheet.cells[row_number, temp_intercept_col].formula2 = intercept_formula
        sheet.cells[row_number, temp_slope_col].formula2 = slope_formula
        formula_meta[row_number] = (num_quarters_used, next_x_value)

    intercept_values: Dict[int, Any] = {}
    slope_values: Dict[int, Any] = {}
    if formula_meta:
        wb.app.calculate()
        intercept_block = read_range_2d(
            sheet,
            candidate_start,
            temp_intercept_col,
            candidate_end,
            temp_intercept_col,
        )
        slope_block = read_range_2d(
            sheet,
            candidate_start,
            temp_slope_col,
            candidate_end,
            temp_slope_col,
        )
        for idx, row_number in enumerate(range(candidate_start, candidate_end + 1)):
            if idx < len(intercept_block):
                intercept_values[row_number] = intercept_block[idx][0]
            if idx < len(slope_block):
                slope_values[row_number] = slope_block[idx][0]

    rows: List[Dict[str, Any]] = []
    blank_streak = 0
    for idx in range(N_QUARTERS):
        row_number = candidate_start + idx
        row_values = candidate_block[idx] if idx < len(candidate_block) else []

        raw_num_quarters = value_from_row_block(
            row_values, min_col, col_map["num_quarters_used"]
        )
        num_quarters_float = to_float(raw_num_quarters)
        num_quarters_used = int(num_quarters_float) if num_quarters_float else idx + 1

        intercept = intercept_values.get(row_number)
        slope = slope_values.get(row_number)

        forecast_value = value_from_row_block(row_values, min_col, col_map["forecast_value"])
        if is_blank(forecast_value):
            intercept_num = to_float(intercept)
            slope_num = to_float(slope)
            meta = formula_meta.get(row_number)
            if intercept_num is not None and slope_num is not None and meta:
                forecast_value = intercept_num + slope_num * meta[1]

        forecast_max = value_from_row_block(row_values, min_col, col_map["forecast_max"])
        forecast_min = value_from_row_block(row_values, min_col, col_map["forecast_min"])
        actual_value = value_from_row_block(row_values, min_col, col_map["actual_value"])

        if all(
            is_blank(value)
            for value in [forecast_value, forecast_max, forecast_min, intercept, slope]
        ):
            blank_streak += 1
            if blank_streak >= 3:
                break
            continue
        blank_streak = 0

        row_payload = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if not is_blank(actual_value) else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_values(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if idx == N_QUARTERS - 1 and rows:
            if is_duplicate_regression_row(rows[-1], row_payload):
                continue

        rows.append(row_payload)

    return rows


def write_sheet(
    sheet,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    sheet.append(list(columns))
    for row_data in rows:
        sheet.append([row_data.get(column, "") for column in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    if sheet.max_row >= 1 and sheet.max_column >= 1:
        sheet.auto_filter.ref = sheet.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(
            max(max_len + 2, 12), 48
        )


def main() -> None:
    source_dir = Path(input_dir)
    destination_dir = Path(output_dir)

    if not source_dir.exists() or not source_dir.is_dir():
        print(f"[error] input_dir does not exist or is not a directory: {source_dir}")
        return

    destination_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(source_dir, destination_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass
    try:
        app.api.EnableEvents = False
    except Exception:
        pass

    try:
        for file_path in sorted(source_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"[skipped] {file_path.name}: temporary Excel file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"[skipped] {file_path.name}: not an .xlsx file")
                continue

            if file_path.resolve() == output_path.resolve():
                print(f"[skipped] {file_path.name}: output target file")
                continue

            print(f"[processing] {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_model_labels(file_path.name)
                empirical_rows.extend(process_empirical_sheet(wb, labels, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"[skipped] {file_path.name}: processing failed ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    empirical_sheet = output_wb.create_sheet("empirical_candidates")
    regression_sheet = output_wb.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    output_wb.save(output_path)

    print(f"output path: {output_path}")
    print(f"files processed: {processed_files}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
