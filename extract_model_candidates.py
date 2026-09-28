from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# User-configurable paths
# -----------------------------------------------------------------------------
input_dir = "/path/to/input/folder"
output_dir = "/path/to/output/folder"

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
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

EMPIRICAL_HEADER_ALIASES = {
    "num_quarters_used": ["num_quarters_used", "num quarters used", "quarters used", "n quarters"],
    "last_quarter_used": ["last_quarter_used", "last quarter used", "last quarter", "last qtr"],
    "forecast_value": [
        "forecast_value",
        "estimated total sold",
        "est total sold",
        "total forecast",
        "forecast",
    ],
    "actual_value": ["actual_value", "actual", "reported sales", "reported"],
    "forecast_min": ["forecast_min", "min"],
    "forecast_max": ["forecast_max", "max"],
    "avg_penetration_pct": ["avg_penetration_pct", "avg penetration", "average penetration", "penetration pct"],
    "quarterly_sales": ["quarterly_sales", "quarterly sales", "quarter sales"],
    "reported_sales": ["reported_sales", "reported sales"],
    "growth_rate_pct": ["growth_rate_pct", "growth rate", "growth pct"],
    "sales_captured_in_db_pct": [
        "sales_captured_in_db_pct",
        "sales captured in db pct",
        "captured in db",
    ],
    "penetration_source": ["penetration", "penetration pct", "quarter penetration"],
}

REGRESSION_HEADER_ALIASES = {
    "num_quarters_used": ["num_quarters_used", "num quarters used", "quarters used", "n quarters"],
    "forecast_value": [
        "forecast_value",
        "tot fcst w/o sa",
        "total forecast without sa",
        "total fcst w/o sa",
    ],
    "actual_value": ["actual_value", "actual", "reported sales", "reported"],
    "forecast_min": ["forecast_min", "min"],
    "forecast_max": ["forecast_max", "max"],
}

# Defaults are anchor-based offsets from the "max" cell.
EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -11,
    "last_quarter_used": -10,
    "quarterly_sales": -9,
    "reported_sales": -8,
    "growth_rate_pct": -7,
    "sales_captured_in_db_pct": -6,
    "avg_penetration_pct": -5,
    "forecast_value": -4,
    "actual_value": -3,
    "forecast_min": 1,
    "forecast_max": 0,
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -4,
    "actual_value": -3,
    "forecast_min": 1,
    "forecast_max": 0,
}


def normalize_token(value: Any) -> str:
    if value is None:
        return ""
    token = str(value).strip().lower()
    token = token.replace("%", " pct ")
    token = re.sub(r"[^a-z0-9]+", "", token)
    return token


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> Optional[int]:
    num = to_float(value)
    if num is None:
        return None
    return int(round(num))


def first_non_blank(*values: Any) -> Any:
    for value in values:
        if not is_blank(value):
            return value
    return None


def flatten_column(values: Any) -> List[Any]:
    if not isinstance(values, list):
        return [values]
    if values and isinstance(values[0], list):
        return [row[0] if row else None for row in values]
    return values


def parse_filename_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    pattern = re.compile(
        r"-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([0-9]{4})",
        re.IGNORECASE,
    )
    match = pattern.search(stem)

    ticker = ""
    model_period = ""
    model_date = ""
    model = ""

    if match:
        ticker = match.group(1).upper()
        timing = match.group(2).title()
        month_token = match.group(3).lower()
        year = int(match.group(4))
        month_lookup = {
            "jan": 1,
            "january": 1,
            "feb": 2,
            "february": 2,
            "mar": 3,
            "march": 3,
            "apr": 4,
            "april": 4,
            "may": 5,
            "jun": 6,
            "june": 6,
            "jul": 7,
            "july": 7,
            "aug": 8,
            "august": 8,
            "sep": 9,
            "sept": 9,
            "september": 9,
            "oct": 10,
            "october": 10,
            "nov": 11,
            "november": 11,
            "dec": 12,
            "december": 12,
        }
        month_num = month_lookup.get(month_token[:3], month_lookup.get(month_token))

        if month_num:
            month_abbrev = calendar.month_abbr[month_num]
            model_period = f"{timing}{month_abbrev}_{year}"
            day_map = {"Early": 5, "Mid": 15, "Late": 25}
            model_date = date(year, month_num, day_map[timing]).isoformat()
            model = f"{ticker}_{model_period}"

    if not ticker:
        fallback = re.split(r"\s*-\s*", stem)
        ticker = fallback[1].strip().upper() if len(fallback) > 1 else "UNKNOWN"
    if not model_period:
        model_period = "UNKNOWN_PERIOD"
    if not model_date:
        model_date = ""
    if not model:
        model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def close_source_workbook(wb: xw.Book) -> None:
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


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def find_anchor(sheet: xw.main.Sheet, token: str = "max") -> Optional[Tuple[int, int]]:
    try:
        found = sheet.api.Cells.Find(What=token, LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    used = sheet.used_range
    used_values = used.value
    start_row = used.row
    start_col = used.column

    if not isinstance(used_values, list):
        used_values = [[used_values]]
    elif used_values and not isinstance(used_values[0], list):
        used_values = [used_values]

    target = normalize_token(token)
    for r_idx, row_values in enumerate(used_values):
        for c_idx, value in enumerate(row_values):
            if normalize_token(value) == target:
                return start_row + r_idx, start_col + c_idx
    return None


def collect_header_offsets(
    sheet: xw.main.Sheet, anchor_row: int, anchor_col: int, span: int = 35
) -> Dict[str, int]:
    left = max(1, anchor_col - span)
    right = anchor_col + span
    row_values = sheet.range((anchor_row, left), (anchor_row, right)).value

    if not isinstance(row_values, list):
        row_values = [row_values]
    if row_values and isinstance(row_values[0], list):
        row_values = row_values[0]

    offsets: Dict[str, int] = {}
    for index, value in enumerate(row_values):
        token = normalize_token(value)
        if token and token not in offsets:
            offsets[token] = (left + index) - anchor_col
    return offsets


def resolve_col_from_anchor(
    anchor_col: int,
    header_offsets: Dict[str, int],
    aliases: Iterable[str],
    fallback_offset: int,
) -> int:
    for alias in aliases:
        token = normalize_token(alias)
        if token in header_offsets:
            return anchor_col + header_offsets[token]
    return anchor_col + fallback_offset


def collect_numeric_history(
    sheet: xw.main.Sheet,
    col: int,
    anchor_row: int,
    lookback_rows: int = 250,
) -> List[Tuple[int, float]]:
    top = max(1, anchor_row - lookback_rows)
    if anchor_row - 1 < top:
        return []
    raw_values = sheet.range((top, col), (anchor_row - 1, col)).value
    values = flatten_column(raw_values)
    history: List[Tuple[int, float]] = []
    for idx, value in enumerate(values, start=top):
        num = to_float(value)
        if num is not None:
            history.append((idx, num))
    return history


def collect_xy_history(
    sheet: xw.main.Sheet,
    x_col: int,
    y_col: int,
    anchor_row: int,
    lookback_rows: int = 250,
) -> List[Tuple[int, float, float]]:
    top = max(1, anchor_row - lookback_rows)
    if anchor_row - 1 < top:
        return []

    x_raw = sheet.range((top, x_col), (anchor_row - 1, x_col)).value
    y_raw = sheet.range((top, y_col), (anchor_row - 1, y_col)).value
    x_vals = flatten_column(x_raw)
    y_vals = flatten_column(y_raw)

    history: List[Tuple[int, float, float]] = []
    for row_idx, (x_val, y_val) in enumerate(zip(x_vals, y_vals), start=top):
        x_num = to_float(x_val)
        y_num = to_float(y_val)
        if x_num is not None and y_num is not None:
            history.append((row_idx, x_num, y_num))
    return history


def make_empirical_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        return []

    anchor = find_anchor(sheet, token="max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    header_offsets = collect_header_offsets(sheet, anchor_row, anchor_col)
    cols = {
        "num_quarters_used": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["num_quarters_used"],
            EMPIRICAL_DEFAULT_OFFSETS["num_quarters_used"],
        ),
        "last_quarter_used": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["last_quarter_used"],
            EMPIRICAL_DEFAULT_OFFSETS["last_quarter_used"],
        ),
        "forecast_value": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["forecast_value"],
            EMPIRICAL_DEFAULT_OFFSETS["forecast_value"],
        ),
        "actual_value": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["actual_value"],
            EMPIRICAL_DEFAULT_OFFSETS["actual_value"],
        ),
        "forecast_max": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["forecast_max"],
            EMPIRICAL_DEFAULT_OFFSETS["forecast_max"],
        ),
        "forecast_min": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["forecast_min"],
            EMPIRICAL_DEFAULT_OFFSETS["forecast_min"],
        ),
        "avg_penetration_pct": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["avg_penetration_pct"],
            EMPIRICAL_DEFAULT_OFFSETS["avg_penetration_pct"],
        ),
        "quarterly_sales": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["quarterly_sales"],
            EMPIRICAL_DEFAULT_OFFSETS["quarterly_sales"],
        ),
        "reported_sales": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["reported_sales"],
            EMPIRICAL_DEFAULT_OFFSETS["reported_sales"],
        ),
        "growth_rate_pct": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["growth_rate_pct"],
            EMPIRICAL_DEFAULT_OFFSETS["growth_rate_pct"],
        ),
        "sales_captured_in_db_pct": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            EMPIRICAL_HEADER_ALIASES["sales_captured_in_db_pct"],
            EMPIRICAL_DEFAULT_OFFSETS["sales_captured_in_db_pct"],
        ),
    }

    penetration_source_col = resolve_col_from_anchor(
        anchor_col,
        header_offsets,
        EMPIRICAL_HEADER_ALIASES["penetration_source"],
        EMPIRICAL_DEFAULT_OFFSETS["avg_penetration_pct"],
    )
    penetration_history = collect_numeric_history(sheet, penetration_source_col, anchor_row)

    helper_col = max(anchor_col + 25, sheet.used_range.last_cell.column + 2)
    helper_rows: Dict[int, int] = {}
    pending_calc = False

    interim_rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row = anchor_row + 1 + idx
        num_q_value = sheet.range((row, cols["num_quarters_used"])).value
        num_quarters_used = to_int(num_q_value) or (idx + 1)

        forecast_max = sheet.range((row, cols["forecast_max"])).value
        forecast_min = sheet.range((row, cols["forecast_min"])).value
        forecast_value = sheet.range((row, cols["forecast_value"])).value
        actual_value = sheet.range((row, cols["actual_value"])).value
        quarterly_sales = sheet.range((row, cols["quarterly_sales"])).value
        reported_sales = sheet.range((row, cols["reported_sales"])).value
        growth_rate_pct = sheet.range((row, cols["growth_rate_pct"])).value
        sales_captured = sheet.range((row, cols["sales_captured_in_db_pct"])).value
        avg_penetration = sheet.range((row, cols["avg_penetration_pct"])).value
        last_quarter = sheet.range((row, cols["last_quarter_used"])).value

        if all(
            is_blank(v)
            for v in [
                forecast_max,
                forecast_min,
                forecast_value,
                actual_value,
                quarterly_sales,
                reported_sales,
                avg_penetration,
            ]
        ):
            continue

        # Existing empirical workflow behavior: compute avg penetration with an R1C1
        # formula for the selected n_quarters when the sheet does not already provide it.
        if is_blank(avg_penetration) and len(penetration_history) >= num_quarters_used:
            start_row = penetration_history[-num_quarters_used][0]
            end_row = penetration_history[-1][0]
            formula = f"=AVERAGE(R{start_row}C{penetration_source_col}:R{end_row}C{penetration_source_col})"
            helper_row = anchor_row + 1 + idx
            set_formula2(sheet.range((helper_row, helper_col)), formula)
            helper_rows[idx] = helper_row
            pending_calc = True

        interim_rows.append(
            {
                "idx": idx,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": first_non_blank(actual_value, reported_sales),
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
            }
        )

    if pending_calc:
        wb.app.calculate()
        for item in interim_rows:
            if item["idx"] in helper_rows and is_blank(item["avg_penetration_pct"]):
                helper_value = sheet.range((helper_rows[item["idx"]], helper_col)).value
                item["avg_penetration_pct"] = helper_value

    output_rows: List[Dict[str, Any]] = []
    for item in interim_rows:
        forecast_max_num = to_float(item["forecast_max"])
        forecast_min_num = to_float(item["forecast_min"])
        range_width = (
            forecast_max_num - forecast_min_num
            if forecast_max_num is not None and forecast_min_num is not None
            else None
        )

        output_rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": item["avg_penetration_pct"],
                "num_quarters_used": item["num_quarters_used"],
                "last_quarter_used": item["last_quarter_used"],
                "forecast_value": item["forecast_value"],
                "actual_value": item["actual_value"],
                "forecast_max": item["forecast_max"],
                "forecast_min": item["forecast_min"],
                "range_width": range_width,
                "avg_penetration_pct": item["avg_penetration_pct"],
                "quarterly_sales": item["quarterly_sales"],
                "reported_sales": item["reported_sales"],
                "growth_rate_pct": item["growth_rate_pct"],
                "sales_captured_in_db_pct": item["sales_captured_in_db_pct"],
                "source_file": source_file,
            }
        )

    return output_rows


def make_regression_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        return []

    anchor = find_anchor(sheet, token="max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    header_offsets = collect_header_offsets(sheet, anchor_row, anchor_col)
    cols = {
        "num_quarters_used": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            REGRESSION_HEADER_ALIASES["num_quarters_used"],
            REGRESSION_DEFAULT_OFFSETS["num_quarters_used"],
        ),
        "forecast_value": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            REGRESSION_HEADER_ALIASES["forecast_value"],
            REGRESSION_DEFAULT_OFFSETS["forecast_value"],
        ),
        "actual_value": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            REGRESSION_HEADER_ALIASES["actual_value"],
            REGRESSION_DEFAULT_OFFSETS["actual_value"],
        ),
        "forecast_max": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            REGRESSION_HEADER_ALIASES["forecast_max"],
            REGRESSION_DEFAULT_OFFSETS["forecast_max"],
        ),
        "forecast_min": resolve_col_from_anchor(
            anchor_col,
            header_offsets,
            REGRESSION_HEADER_ALIASES["forecast_min"],
            REGRESSION_DEFAULT_OFFSETS["forecast_min"],
        ),
    }

    # Existing regression workflow specification.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    xy_history = collect_xy_history(sheet, x_col, y_col, anchor_row)
    if len(xy_history) < 2:
        return []

    helper_col = max(anchor_col + 25, sheet.used_range.last_cell.column + 2)
    helper_rows: Dict[int, int] = {}
    max_rows = min(N_QUARTERS, len(xy_history))

    for n_quarters in range(1, max_rows + 1):
        start_row = xy_history[-n_quarters][0]
        end_row = xy_history[-1][0]
        formula_row = anchor_row + n_quarters

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2(sheet.range((formula_row, helper_col)), intercept_formula)
        set_formula2(sheet.range((formula_row, helper_col + 1)), slope_formula)
        helper_rows[n_quarters] = formula_row

    wb.app.calculate()

    x_next = xy_history[-1][1] + 1.0
    output_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n_quarters in range(1, max_rows + 1):
        row = anchor_row + n_quarters
        helper_row = helper_rows[n_quarters]

        intercept = sheet.range((helper_row, helper_col)).value
        slope = sheet.range((helper_row, helper_col + 1)).value
        intercept_num = to_float(intercept)
        slope_num = to_float(slope)

        forecast_from_regression = (
            (intercept_num + (slope_num * x_next))
            if intercept_num is not None and slope_num is not None
            else None
        )
        sheet_forecast_value = sheet.range((row, cols["forecast_value"])).value
        forecast_value = first_non_blank(sheet_forecast_value, forecast_from_regression)

        actual_value = sheet.range((row, cols["actual_value"])).value
        forecast_max = sheet.range((row, cols["forecast_max"])).value
        forecast_min = sheet.range((row, cols["forecast_min"])).value
        num_q_raw = sheet.range((row, cols["num_quarters_used"])).value
        num_quarters_used = to_int(num_q_raw) or n_quarters

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        signature = (
            num_quarters_used,
            round(intercept_num, 10) if intercept_num is not None else None,
            round(slope_num, 10) if slope_num is not None else None,
            round(to_float(forecast_value), 10) if to_float(forecast_value) is not None else None,
            round(max_num, 10) if max_num is not None else None,
            round(min_num, 10) if min_num is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        output_rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return output_rows


def style_sheet(ws, columns: List[str], row_count: int) -> None:
    for col_idx in range(1, len(columns) + 1):
        header_cell = ws.cell(row=1, column=col_idx)
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, row_count + 1)}"

    for col_idx, col_name in enumerate(columns, start=1):
        values = [col_name]
        for row_idx in range(2, row_count + 2):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is not None:
                values.append(str(value))
        max_len = max(len(v) for v in values) if values else 10
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_file: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()

    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_empirical.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_empirical.append([row.get(col) for col in EMPIRICAL_COLUMNS])
    style_sheet(ws_empirical, EMPIRICAL_COLUMNS, len(empirical_rows))

    ws_regression = wb.create_sheet("regression_candidates")
    ws_regression.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_regression.append([row.get(col) for col in REGRESSION_COLUMNS])
    style_sheet(ws_regression, REGRESSION_COLUMNS, len(regression_rows))

    wb.save(output_file)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        print(f"Skipped: input directory does not exist -> {input_path}")
        return

    output_file = next_output_path(input_path, output_path)
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
        files = sorted(input_path.iterdir())
        for file_path in files:
            if not file_path.is_file():
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not .xlsx)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary file)")
                continue

            print(f"Processing: {file_path.name}")
            wb = None
            try:
                labels = parse_filename_labels(file_path)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(make_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(make_regression_rows(wb, labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
