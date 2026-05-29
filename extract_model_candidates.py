from __future__ import annotations

import datetime as dt
import math
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# =========================
# User-configurable paths
# =========================
input_dir = Path("./input")
output_dir = Path("./output")


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
    "sept": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

WINDOW_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if values and not isinstance(values[0], (list, tuple)):
        return [list(values)]
    return [list(row) for row in values]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        raw = value.strip().replace(",", "")
        if not raw:
            return None
        if raw.endswith("%"):
            raw = raw[:-1].strip()
            try:
                return float(raw) / 100.0
            except ValueError:
                return None
        try:
            return float(raw)
        except ValueError:
            return None
    return None


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def coalesce(*values: Any) -> Any:
    for value in values:
        if value is not None and value != "":
            return value
    return None


def format_quarter_label(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        # Excel serial date fallback handling (broadly safe for modern sheets)
        if value > 20000:
            try:
                excel_epoch = dt.datetime(1899, 12, 30)
                guessed = excel_epoch + dt.timedelta(days=float(value))
                return guessed.date().isoformat()
            except Exception:
                pass
    return str(value).strip()


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    ticker = "UNKNOWN"
    period_token = ""

    # Expected pattern: "<prefix> - TICKER - MidJan2026_Send"
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3,9}\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if match:
        ticker = match.group("ticker").upper()
        period_token = match.group("period")
    else:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()
        period_match = re.search(r"(Early|Mid|Late)[A-Za-z]{3,9}\d{4}", stem, flags=re.IGNORECASE)
        if period_match:
            period_token = period_match.group(0)

    window = "Mid"
    month_token = "Jan"
    year_token = "1900"
    period_parse = re.match(
        r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if period_parse:
        window = period_parse.group(1).capitalize()
        month_token = period_parse.group(2).title()
        year_token = period_parse.group(3)

    month_key = month_token.lower()[:4] if month_token.lower().startswith("sept") else month_token.lower()[:3]
    month_num = MONTH_MAP.get(month_key, 1)
    month_short = dt.date(2000, month_num, 1).strftime("%b")
    day = WINDOW_DAY_MAP.get(window, 15)
    year_int = int(year_token)
    model_period = f"{window}{month_short}_{year_int}"
    model_date = dt.date(year_int, month_num, day).isoformat()

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    index = 1
    while True:
        candidate = out_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def get_sheet_by_name(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def build_text_index(
    values_2d: List[List[Any]],
    start_row: int,
    start_col: int,
) -> List[Tuple[int, int, str]]:
    index: List[Tuple[int, int, str]] = []
    for r_idx, row in enumerate(values_2d):
        for c_idx, value in enumerate(row):
            if isinstance(value, str):
                cleaned = value.strip()
                if cleaned:
                    index.append((start_row + r_idx, start_col + c_idx, cleaned.lower()))
    return index


def find_anchor_max(
    values_2d: List[List[Any]],
    start_row: int,
    start_col: int,
) -> Tuple[int, int]:
    candidates: List[Tuple[int, int]] = []
    for r_idx, row in enumerate(values_2d):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                candidates.append((start_row + r_idx, start_col + c_idx))

    if not candidates:
        raise ValueError("Could not find 'max' anchor cell.")

    def score(candidate: Tuple[int, int]) -> int:
        row, col = candidate
        local_row = row - start_row
        local_col = col - start_col
        s = 0
        neighbors = [
            (local_row + 1, local_col),
            (local_row - 1, local_col),
            (local_row, local_col + 1),
            (local_row, local_col - 1),
        ]
        for n_row, n_col in neighbors:
            if 0 <= n_row < len(values_2d) and 0 <= n_col < len(values_2d[n_row]):
                n_val = values_2d[n_row][n_col]
                if isinstance(n_val, str) and n_val.strip().lower() == "min":
                    s += 3
                if to_float(n_val) is not None:
                    s += 1
        return s

    return max(candidates, key=score)


def find_nearest_text_cell(
    text_index: Sequence[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
    required_tokens: Sequence[str],
    row_radius: int = 120,
    col_radius: int = 120,
) -> Optional[Tuple[int, int]]:
    lowered = [token.lower() for token in required_tokens]
    matches: List[Tuple[int, int, int]] = []
    for row, col, text in text_index:
        if abs(row - anchor_row) > row_radius or abs(col - anchor_col) > col_radius:
            continue
        if all(token in text for token in lowered):
            distance = abs(row - anchor_row) + abs(col - anchor_col)
            matches.append((distance, row, col))
    if not matches:
        return None
    matches.sort(key=lambda item: item[0])
    _, row, col = matches[0]
    return row, col


def read_adjacent_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    # Typical label/value layouts are right/down from label.
    for d_row, d_col in ((0, 1), (1, 0), (0, 2), (1, 1), (0, -1)):
        test_val = sheet.range((row + d_row, col + d_col)).value
        if test_val not in (None, ""):
            return test_val
    return None


def find_numeric_series_row(
    values_2d: List[List[Any]],
    start_row: int,
    text_index: Sequence[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
    tokens: Sequence[str],
) -> Optional[Tuple[int, List[int]]]:
    label = find_nearest_text_cell(text_index, anchor_row, anchor_col, tokens)
    if label is None:
        return None
    row, _ = label
    local_row = row - start_row
    if not (0 <= local_row < len(values_2d)):
        return None

    numeric_cols: List[int] = []
    for c_idx, value in enumerate(values_2d[local_row]):
        if to_float(value) is not None:
            numeric_cols.append(c_idx)

    if len(numeric_cols) < 2:
        return None
    return row, numeric_cols


def r1c1_average_formula(row: int, first_col: int, last_col: int) -> str:
    return f"=AVERAGE(R{row}C{first_col}:R{row}C{last_col})"


def r1c1_intercept_formula(first_row: int, last_row: int, x_col: int, y_col: int) -> str:
    return (
        f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
        f"R{first_row}C{x_col}:R{last_row}C{x_col})"
    )


def r1c1_slope_formula(first_row: int, last_row: int, x_col: int, y_col: int) -> str:
    return (
        f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
        f"R{first_row}C{x_col}:R{last_row}C{x_col})"
    )


def excel_calculate(app: xw.App) -> None:
    app.calculate()


def safe_close_workbook(wb: Optional[xw.Book]) -> None:
    if wb is None:
        return
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
        try:
            wb.api.Close(False)
        except Exception:
            pass


def extract_empirical_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Empirical Model")
    if sheet is None:
        return []

    used = sheet.used_range
    values_2d = normalize_2d(used.value)
    if not values_2d:
        return []
    start_row = used.row
    start_col = used.column

    anchor_row, anchor_col = find_anchor_max(values_2d, start_row, start_col)
    text_index = build_text_index(values_2d, start_row, start_col)

    helper_col = max(used.last_cell.column + 5, anchor_col + 5)
    helper_row = max(used.last_cell.row + 5, anchor_row + 5)

    avg_pen_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["avg", "penetration"])
    avg_pen_cell = None
    if avg_pen_label:
        avg_pen_cell = sheet.range((avg_pen_label[0], avg_pen_label[1] + 1))
    else:
        avg_pen_cell = sheet.range((helper_row, helper_col))

    original_avg_formula = avg_pen_cell.formula2
    original_avg_value = avg_pen_cell.value

    max_cell = sheet.range((anchor_row, anchor_col + 1))
    min_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["min"])
    min_cell = sheet.range((min_label[0], min_label[1] + 1)) if min_label else sheet.range((anchor_row + 1, anchor_col + 1))

    estimated_total_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["estimated", "total", "sold"])
    forecast_cell = (
        sheet.range((estimated_total_label[0], estimated_total_label[1] + 1))
        if estimated_total_label
        else sheet.range((anchor_row - 2, anchor_col + 1))
    )

    quarterly_sales_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["quarterly", "sales"])
    reported_sales_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["reported", "sales"])
    growth_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["growth"])
    captured_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["captured", "db"])

    quarterly_sales_cell = (
        sheet.range((quarterly_sales_label[0], quarterly_sales_label[1] + 1))
        if quarterly_sales_label
        else None
    )
    reported_sales_cell = (
        sheet.range((reported_sales_label[0], reported_sales_label[1] + 1))
        if reported_sales_label
        else None
    )
    growth_cell = sheet.range((growth_label[0], growth_label[1] + 1)) if growth_label else None
    captured_cell = sheet.range((captured_label[0], captured_label[1] + 1)) if captured_label else None

    penetration_series = find_numeric_series_row(
        values_2d, start_row, text_index, anchor_row, anchor_col, ["penetration"]
    )
    quarter_row = penetration_series[0] - 1 if penetration_series else None
    pen_cols = penetration_series[1] if penetration_series else []

    rows: List[Dict[str, Any]] = []
    n_quarters = 10
    try:
        for used_n in range(1, n_quarters + 1):
            avg_pen = to_float(avg_pen_cell.value)
            if penetration_series and len(pen_cols) >= used_n:
                first_local_col = pen_cols[-used_n]
                last_local_col = pen_cols[-1]
                first_col = start_col + first_local_col
                last_col = start_col + last_local_col
                formula = r1c1_average_formula(penetration_series[0], first_col, last_col)
                avg_pen_cell.formula2 = formula
                excel_calculate(wb.app)
                avg_pen = to_float(avg_pen_cell.value)

            forecast_max = to_float(coalesce(max_cell.value, read_adjacent_value(sheet, anchor_row, anchor_col)))
            forecast_min = to_float(
                coalesce(min_cell.value, read_adjacent_value(sheet, min_label[0], min_label[1]) if min_label else None)
            )
            forecast_value = to_float(
                coalesce(
                    forecast_cell.value,
                    safe_div(to_float(quarterly_sales_cell.value if quarterly_sales_cell else None), avg_pen),
                )
            )
            quarterly_sales = to_float(quarterly_sales_cell.value if quarterly_sales_cell else None)
            reported_sales = to_float(reported_sales_cell.value if reported_sales_cell else None)
            growth_rate = to_float(growth_cell.value if growth_cell else None)
            captured_pct = to_float(captured_cell.value if captured_cell else None)
            if captured_pct is None:
                captured_pct = safe_div(quarterly_sales, reported_sales)

            last_quarter_used = ""
            if penetration_series and quarter_row is not None and len(pen_cols) >= used_n:
                q_col = start_col + pen_cols[-used_n]
                quarter_val = sheet.range((quarter_row, q_col)).value
                last_quarter_used = format_quarter_label(quarter_val)

            if growth_rate is None and quarterly_sales is not None and penetration_series and len(pen_cols) >= (used_n + 1):
                prev_col = start_col + pen_cols[-(used_n + 1)]
                prev_val = to_float(sheet.range((quarterly_sales_label[0], prev_col)).value) if quarterly_sales_label else None
                if prev_val not in (None, 0):
                    growth_rate = (quarterly_sales - prev_val) / prev_val

            actual_value = reported_sales
            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            rows.append(
                {
                    "model": metadata["model"],
                    "ticker": metadata["ticker"],
                    "model_period": metadata["model_period"],
                    "model_date": metadata["model_date"],
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_pen,
                    "num_quarters_used": used_n,
                    "last_quarter_used": last_quarter_used,
                    "forecast_value": forecast_value,
                    "actual_value": actual_value,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "avg_penetration_pct": avg_pen,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate,
                    "sales_captured_in_db_pct": captured_pct,
                    "source_file": source_file,
                }
            )
    finally:
        # Restore any changed formula/value so the source workbook session remains safe.
        try:
            if original_avg_formula:
                avg_pen_cell.formula2 = original_avg_formula
            else:
                avg_pen_cell.value = original_avg_value
        except Exception:
            pass

    return rows


def regression_data_rows(
    values_2d: List[List[Any]],
    start_row: int,
    end_row: int,
    start_col: int,
    x_col: int,
    y_col: int,
) -> List[int]:
    valid_rows: List[int] = []
    if end_row < start_row:
        return valid_rows

    x_local = x_col - start_col
    y_local = y_col - start_col
    if x_local < 0 or y_local < 0:
        return valid_rows

    for row in range(start_row, end_row + 1):
        local_row = row - start_row
        if not (0 <= local_row < len(values_2d)):
            continue
        row_vals = values_2d[local_row]
        if x_local >= len(row_vals) or y_local >= len(row_vals):
            continue
        x_val = to_float(row_vals[x_local])
        y_val = to_float(row_vals[y_local])
        if x_val is not None and y_val is not None:
            valid_rows.append(row)
    return valid_rows


def extract_regression_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Regression Model")
    if sheet is None:
        return []

    used = sheet.used_range
    values_2d = normalize_2d(used.value)
    if not values_2d:
        return []

    start_row = used.row
    start_col = used.column
    anchor_row, anchor_col = find_anchor_max(values_2d, start_row, start_col)
    text_index = build_text_index(values_2d, start_row, start_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    helper_col = max(used.last_cell.column + 8, anchor_col + 8)
    helper_row = max(used.last_cell.row + 8, anchor_row + 8)
    intercept_cell = sheet.range((helper_row, helper_col))
    slope_cell = sheet.range((helper_row + 1, helper_col))

    original_intercept_formula = intercept_cell.formula2
    original_intercept_value = intercept_cell.value
    original_slope_formula = slope_cell.formula2
    original_slope_value = slope_cell.value

    max_cell = sheet.range((anchor_row, anchor_col + 1))
    min_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["min"])
    min_cell = sheet.range((min_label[0], min_label[1] + 1)) if min_label else sheet.range((anchor_row + 1, anchor_col + 1))
    fcst_label = find_nearest_text_cell(text_index, anchor_row, anchor_col, ["tot", "fcst", "w/o", "sa"])
    forecast_cell = sheet.range((fcst_label[0], fcst_label[1] + 1)) if fcst_label else None

    rows = regression_data_rows(values_2d, start_row, anchor_row - 1, start_col, x_col, y_col)
    if len(rows) < 2:
        return []

    max_n = min(10, len(rows))
    collected: List[Dict[str, Any]] = []

    try:
        prev_key: Optional[Tuple[Any, ...]] = None
        for used_n in range(2, max_n + 1):
            first_row = rows[-used_n]
            last_row = rows[-1]

            intercept_cell.formula2 = r1c1_intercept_formula(first_row, last_row, x_col, y_col)
            slope_cell.formula2 = r1c1_slope_formula(first_row, last_row, x_col, y_col)
            excel_calculate(wb.app)

            intercept = to_float(intercept_cell.value)
            slope = to_float(slope_cell.value)

            next_x = to_float(sheet.range((last_row + 1, x_col)).value)
            if next_x is None:
                next_x = to_float(sheet.range((last_row, x_col)).value)

            forecast_value = to_float(forecast_cell.value if forecast_cell else None)
            if forecast_value is None and intercept is not None and slope is not None and next_x is not None:
                forecast_value = intercept + slope * next_x

            forecast_max = to_float(max_cell.value)
            forecast_min = to_float(min_cell.value)

            if forecast_max is None and forecast_value is not None:
                forecast_max = forecast_value
            if forecast_min is None and forecast_value is not None:
                forecast_min = forecast_value

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            row_payload = {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": used_n,
                "num_quarters_used": used_n,
                "forecast_value": forecast_value,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }

            dedupe_key = (
                round(intercept, 12) if intercept is not None else None,
                round(slope, 12) if slope is not None else None,
                round(forecast_value, 12) if forecast_value is not None else None,
                round(forecast_max, 12) if forecast_max is not None else None,
                round(forecast_min, 12) if forecast_min is not None else None,
            )
            if prev_key is not None and dedupe_key == prev_key:
                continue
            prev_key = dedupe_key
            collected.append(row_payload)
    finally:
        try:
            if original_intercept_formula:
                intercept_cell.formula2 = original_intercept_formula
            else:
                intercept_cell.value = original_intercept_value
            if original_slope_formula:
                slope_cell.formula2 = original_slope_formula
            else:
                slope_cell.value = original_slope_value
        except Exception:
            pass

    return collected


def autosize_columns(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[letter]:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[letter].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    emp_ws = wb.active
    emp_ws.title = "empirical_candidates"
    reg_ws = wb.create_sheet("regression_candidates")

    def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
        ws.append(list(columns))
        for col_idx in range(1, len(columns) + 1):
            ws.cell(row=1, column=col_idx).font = Font(bold=True)

        for row in rows:
            ws.append([row.get(column) for column in columns])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autosize_columns(ws)

    write_sheet(emp_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(reg_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def iter_candidate_files(folder: Path) -> Iterable[Tuple[Optional[Path], str]]:
    if not folder.exists():
        raise FileNotFoundError(f"Input folder does not exist: {folder}")
    for item in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not item.is_file():
            yield None, f"skipped {item.name}: not a file"
            continue
        if item.name.startswith("~"):
            yield None, f"skipped {item.name}: temporary file"
            continue
        if item.suffix.lower() != ".xlsx":
            yield None, f"skipped {item.name}: not .xlsx"
            continue
        yield item, ""


def run() -> None:
    app = None
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    output_path = build_output_path(input_dir, output_dir)
    print(f"output path: {output_path}")

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path, skip_reason in iter_candidate_files(input_dir):
            if file_path is None:
                print(skip_reason)
                continue

            print(f"processing file: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_metadata(file_path.name)
                file_empirical: List[Dict[str, Any]] = []
                file_regression: List[Dict[str, Any]] = []
                try:
                    file_empirical = extract_empirical_rows(wb, meta, file_path.name)
                except Exception as exc:
                    print(f"skipped empirical in {file_path.name}: {exc}")
                try:
                    file_regression = extract_regression_rows(wb, meta, file_path.name)
                except Exception as exc:
                    print(f"skipped regression in {file_path.name}: {exc}")

                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                processed_files += 1
            except Exception as exc:
                print(f"skipped {file_path.name}: processing error: {exc}")
            finally:
                safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"files processed: {processed_files}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")
    print(f"output path: {output_path}")


if __name__ == "__main__":
    run()
