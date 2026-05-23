from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import openpyxl
import xlwings as xw
from openpyxl.styles import Font


# Configure these two paths before running.
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
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

LABEL_PATTERNS = {
    "reported_sales": ["reported sales"],
    "quarterly_sales": ["quarterly sales"],
    "growth_rate_pct": ["growth rate"],
    "sales_captured_in_db_pct": ["sales captured in db", "captured in db"],
    "forecast_value": ["estimated total sold", "tot fcst w/o sa", "tot fcst wo sa", "total forecast without sa"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "actual_value": ["actual value"],
}


@dataclass
class FileLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        is_pct = "%" in raw
        cleaned = raw.replace(",", "").replace("%", "")
        cleaned = cleaned.strip("()")
        try:
            numeric = float(cleaned)
        except ValueError:
            return None
        if raw.startswith("(") and raw.endswith(")"):
            numeric = -numeric
        if is_pct:
            numeric /= 100.0
        return numeric
    return None


def safe_round(value: Optional[float], digits: int = 6) -> Optional[float]:
    if value is None:
        return None
    return round(value, digits)


def as_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def parse_label_from_filename(file_path: Path) -> FileLabel:
    stem = file_path.stem
    pattern = re.compile(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        ticker_guess = stem.split("-")[0].strip().replace(" ", "_")
        return FileLabel(
            ticker=ticker_guess,
            model_period="unknown_period",
            model_date="",
            model=f"{ticker_guess}_unknown_period",
        )

    ticker = match.group("ticker").upper()
    period_token = match.group("period")
    period_match = re.match(r"(?P<phase>Early|Mid|Late)(?P<mon>[A-Za-z]{3})(?P<yr>\d{4})", period_token, flags=re.IGNORECASE)

    if not period_match:
        return FileLabel(
            ticker=ticker,
            model_period=period_token,
            model_date="",
            model=f"{ticker}_{period_token}",
        )

    phase = period_match.group("phase")
    mon = period_match.group("mon")
    yr = int(period_match.group("yr"))
    month_num = MONTH_MAP.get(mon.lower())
    day = PERIOD_DAY_MAP.get(phase.lower(), 15)

    model_period = f"{phase}{mon}_{yr}"
    model_date = ""
    if month_num:
        model_date = date(yr, month_num, day).isoformat()

    return FileLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=f"{ticker}_{model_period}",
    )


def next_output_path(input_folder_name: str, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder_name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def build_sheet_snapshot(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int]:
    used = sheet.used_range
    values = as_matrix(used.value)
    return values, used.row, used.column


def build_value_map(matrix: Sequence[Sequence[Any]], top_row: int, left_col: int) -> Dict[Tuple[int, int], Any]:
    value_map: Dict[Tuple[int, int], Any] = {}
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            value_map[(top_row + r_idx, left_col + c_idx)] = value
    return value_map


def find_anchor_max_cell(matrix: Sequence[Sequence[Any]], top_row: int, left_col: int) -> Optional[Tuple[int, int]]:
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return top_row + r_idx, left_col + c_idx
    return None


def nearest_label_value(
    value_map: Dict[Tuple[int, int], Any],
    anchor: Tuple[int, int],
    patterns: Sequence[str],
) -> Optional[float]:
    matches: List[Tuple[int, int, int]] = []
    anchor_row, anchor_col = anchor
    for (row, col), value in value_map.items():
        normalized = normalize_text(value)
        if not normalized:
            continue
        for pattern in patterns:
            if pattern in normalized:
                distance = abs(row - anchor_row) + abs(col - anchor_col)
                matches.append((distance, row, col))
                break
    if not matches:
        return None

    matches.sort(key=lambda item: item[0])
    _, label_row, label_col = matches[0]
    right_value = to_number(value_map.get((label_row, label_col + 1)))
    if right_value is not None:
        return right_value
    below_value = to_number(value_map.get((label_row + 1, label_col)))
    if below_value is not None:
        return below_value
    return to_number(value_map.get((label_row + 1, label_col + 1)))


def row_numeric_values(
    value_map: Dict[Tuple[int, int], Any],
    row: int,
    max_col: int,
) -> List[Tuple[int, float]]:
    values: List[Tuple[int, float]] = []
    for (r, c), value in value_map.items():
        if r != row or c >= max_col:
            continue
        numeric = to_number(value)
        if numeric is None:
            continue
        values.append((c, numeric))
    values.sort(key=lambda item: item[0])
    return values


def find_empirical_rows(
    value_map: Dict[Tuple[int, int], Any],
    top_row: int,
    bottom_row: int,
    anchor_row: int,
    anchor_col: int,
) -> Tuple[Optional[int], Optional[int]]:
    penetration_candidates: List[Tuple[int, int, int]] = []
    sales_candidates: List[Tuple[int, int, int]] = []

    search_start = max(top_row, anchor_row - 50)
    search_end = min(bottom_row, anchor_row + 5)

    for row in range(search_start, search_end + 1):
        numeric = row_numeric_values(value_map, row, anchor_col)
        if len(numeric) < 6:
            continue
        vals = [n for _, n in numeric]
        pct_like = [n for n in vals if 0 <= n <= 2]
        big_vals = [n for n in vals if n > 2]
        proximity = abs(anchor_row - row)

        if len(pct_like) >= 6:
            penetration_candidates.append((proximity, -len(pct_like), row))
        if len(big_vals) >= 6:
            sales_candidates.append((proximity, -len(big_vals), row))

    penetration_row = None
    if penetration_candidates:
        penetration_candidates.sort()
        penetration_row = penetration_candidates[0][2]

    sales_row = None
    if sales_candidates:
        sales_candidates.sort()
        sales_row = sales_candidates[0][2]

    return penetration_row, sales_row


def set_formula2_r1c1(target: xw.Range, formula_r1c1: str) -> None:
    try:
        target.formula2 = formula_r1c1
    except Exception:
        target.api.Formula2R1C1 = formula_r1c1


def calc_once(app: xw.App) -> None:
    app.calculate()


def close_source_workbook(wb: xw.Book) -> None:
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

    wb.api.Close(SaveChanges=False)


def get_quarter_label(
    value_map: Dict[Tuple[int, int], Any],
    penetration_row: int,
    quarter_col: int,
) -> str:
    label = value_map.get((penetration_row - 1, quarter_col))
    if label is None:
        label = value_map.get((penetration_row - 2, quarter_col))
    if label is None:
        return ""
    return str(label).strip()


def extract_empirical_candidates(
    wb: xw.Book,
    sheet: xw.Sheet,
    file_label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    matrix, top_row, left_col = build_sheet_snapshot(sheet)
    if not matrix:
        return rows

    bottom_row = top_row + len(matrix) - 1
    value_map = build_value_map(matrix, top_row, left_col)
    anchor = find_anchor_max_cell(matrix, top_row, left_col)
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor
    penetration_row, sales_row = find_empirical_rows(value_map, top_row, bottom_row, anchor_row, anchor_col)
    if penetration_row is None:
        return rows

    penetration_numeric = row_numeric_values(value_map, penetration_row, anchor_col)
    if len(penetration_numeric) < 2:
        return rows

    penetration_numeric = penetration_numeric[-10:]
    quarter_cols = [col for col, _ in penetration_numeric]
    penetration_values = [val for _, val in penetration_numeric]

    sales_series: List[float] = []
    if sales_row is not None:
        sales_map = dict(row_numeric_values(value_map, sales_row, anchor_col))
        sales_series = [sales_map.get(col) for col in quarter_cols]
        sales_series = [v for v in sales_series if v is not None]

    reported_sales_label = nearest_label_value(value_map, anchor, LABEL_PATTERNS["reported_sales"])
    forecast_max_label = nearest_label_value(value_map, anchor, LABEL_PATTERNS["forecast_max"])
    forecast_min_label = nearest_label_value(value_map, anchor, LABEL_PATTERNS["forecast_min"])
    quarterly_sales_label = nearest_label_value(value_map, anchor, LABEL_PATTERNS["quarterly_sales"])
    growth_rate_label = nearest_label_value(value_map, anchor, LABEL_PATTERNS["growth_rate_pct"])
    sales_captured_label = nearest_label_value(value_map, anchor, LABEL_PATTERNS["sales_captured_in_db_pct"])

    latest_sales = sales_series[-1] if sales_series else quarterly_sales_label
    previous_sales = sales_series[-2] if len(sales_series) >= 2 else None
    if reported_sales_label is None:
        reported_sales_label = latest_sales

    scratch_avg = sheet.cells(anchor_row + 2, anchor_col + 2)

    max_n = min(10, len(penetration_values))
    for n in range(1, max_n + 1):
        start_col = quarter_cols[-n]
        end_col = quarter_cols[-1]
        set_formula2_r1c1(
            scratch_avg,
            f"=AVERAGE(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})",
        )
        calc_once(wb.app)
        avg_pen = to_number(scratch_avg.value)
        if avg_pen is None:
            window = penetration_values[-n:]
            avg_pen = sum(window) / len(window)

        reported_sales = reported_sales_label
        if reported_sales is None:
            continue

        window_pen = [v for v in penetration_values[-n:] if v and v > 0]
        if not window_pen:
            continue

        forecast_value = reported_sales / avg_pen if avg_pen and avg_pen > 0 else None
        computed_max = reported_sales / min(window_pen)
        computed_min = reported_sales / max(window_pen)
        forecast_max = forecast_max_label if forecast_max_label is not None else computed_max
        forecast_min = forecast_min_label if forecast_min_label is not None else computed_min
        range_width = forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None

        growth_rate = growth_rate_label
        if growth_rate is None and latest_sales is not None and previous_sales not in (None, 0):
            growth_rate = (latest_sales - previous_sales) / previous_sales

        quarter_label = get_quarter_label(value_map, penetration_row, start_col)
        sales_captured = sales_captured_label if sales_captured_label is not None else avg_pen

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": safe_round(avg_pen),
                "num_quarters_used": n,
                "last_quarter_used": quarter_label,
                "forecast_value": safe_round(forecast_value),
                "actual_value": safe_round(reported_sales),
                "forecast_max": safe_round(forecast_max),
                "forecast_min": safe_round(forecast_min),
                "range_width": safe_round(range_width),
                "avg_penetration_pct": safe_round(avg_pen),
                "quarterly_sales": safe_round(latest_sales),
                "reported_sales": safe_round(reported_sales),
                "growth_rate_pct": safe_round(growth_rate),
                "sales_captured_in_db_pct": safe_round(sales_captured),
                "source_file": source_file,
            }
        )

    scratch_avg.clear_contents()
    return rows


def manual_linear_regression(points: Sequence[Tuple[float, float]]) -> Tuple[Optional[float], Optional[float]]:
    n = len(points)
    if n < 2:
        return None, None
    sum_x = sum(x for x, _ in points)
    sum_y = sum(y for _, y in points)
    sum_xy = sum(x * y for x, y in points)
    sum_x2 = sum(x * x for x, _ in points)

    denom = (n * sum_x2) - (sum_x ** 2)
    if denom == 0:
        return None, None
    slope = ((n * sum_xy) - (sum_x * sum_y)) / denom
    intercept = (sum_y - slope * sum_x) / n
    return intercept, slope


def close_enough(a: Optional[float], b: Optional[float], tol: float = 1e-9) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def extract_regression_candidates(
    wb: xw.Book,
    sheet: xw.Sheet,
    file_label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    matrix, top_row, left_col = build_sheet_snapshot(sheet)
    if not matrix:
        return rows

    value_map = build_value_map(matrix, top_row, left_col)
    anchor = find_anchor_max_cell(matrix, top_row, left_col)
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7

    data_points: List[Tuple[int, float, float]] = []
    for row in range(top_row, anchor_row):
        x = to_number(value_map.get((row, x_col)))
        y = to_number(value_map.get((row, y_col)))
        if x is None or y is None:
            continue
        data_points.append((row, x, y))

    if len(data_points) < 2:
        return rows

    data_points = data_points[-10:]

    actual_value = nearest_label_value(value_map, anchor, LABEL_PATTERNS["actual_value"])
    static_tot_fcst = nearest_label_value(value_map, anchor, LABEL_PATTERNS["forecast_value"])
    static_max = nearest_label_value(value_map, anchor, LABEL_PATTERNS["forecast_max"])
    static_min = nearest_label_value(value_map, anchor, LABEL_PATTERNS["forecast_min"])

    scratch_intercept = sheet.cells(anchor_row + 2, anchor_col + 2)
    scratch_slope = sheet.cells(anchor_row + 2, anchor_col + 3)

    previous_signature: Optional[Tuple[Optional[float], Optional[float], Optional[float], Optional[float], Optional[float]]] = None

    max_n = min(10, len(data_points))
    for n in range(2, max_n + 1):
        window = data_points[-n:]
        row_start = window[0][0]
        row_end = window[-1][0]

        set_formula2_r1c1(
            scratch_intercept,
            f"=INTERCEPT(R{row_start}C{y_col}:R{row_end}C{y_col},R{row_start}C{x_col}:R{row_end}C{x_col})",
        )
        set_formula2_r1c1(
            scratch_slope,
            f"=SLOPE(R{row_start}C{y_col}:R{row_end}C{y_col},R{row_start}C{x_col}:R{row_end}C{x_col})",
        )
        calc_once(wb.app)

        intercept = to_number(scratch_intercept.value)
        slope = to_number(scratch_slope.value)

        if intercept is None or slope is None:
            manual_points = [(x, y) for _, x, y in window]
            intercept, slope = manual_linear_regression(manual_points)
            if intercept is None or slope is None:
                continue

        latest_x = window[-1][1]
        predicted = [intercept + slope * x for _, x, _ in window]
        forecast_value = intercept + slope * latest_x
        forecast_max = static_max if static_max is not None else max(predicted)
        forecast_min = static_min if static_min is not None else min(predicted)
        if static_tot_fcst is not None:
            forecast_value = static_tot_fcst

        signature = (
            safe_round(intercept, 10),
            safe_round(slope, 10),
            safe_round(forecast_value, 10),
            safe_round(forecast_max, 10),
            safe_round(forecast_min, 10),
        )
        if previous_signature is not None and all(close_enough(cur, prev) for cur, prev in zip(signature, previous_signature)):
            continue
        previous_signature = signature

        range_width = forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": safe_round(forecast_value),
                "actual_value": safe_round(actual_value),
                "forecast_max": safe_round(forecast_max),
                "forecast_min": safe_round(forecast_min),
                "range_width": safe_round(range_width),
                "intercept": safe_round(intercept),
                "slope": safe_round(slope),
                "source_file": source_file,
            }
        )

    scratch_intercept.clear_contents()
    scratch_slope.clear_contents()
    return rows


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_table(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_table(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def write_table(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = min(48, max(12, max_len + 2))


def process_workbooks() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    files = sorted(
        path
        for path in input_dir.iterdir()
        if path.is_file() and path.suffix.lower() == ".xlsx"
    )

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = "manual"
        try:
            app.api.ReferenceStyle = -4150  # xlR1C1
        except Exception:
            pass

        for file_path in files:
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temp file)")
                continue

            print(f"Processing file: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (open failed: {exc})")
                continue

            try:
                metadata = parse_label_from_filename(file_path)

                if "Empirical Model" in [s.name for s in wb.sheets]:
                    empirical_sheet = wb.sheets["Empirical Model"]
                    empirical_rows.extend(
                        extract_empirical_candidates(
                            wb=wb,
                            sheet=empirical_sheet,
                            file_label=metadata,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"Skipped Empirical Model in {file_path.name} (sheet missing)")

                if "Regression Model" in [s.name for s in wb.sheets]:
                    regression_sheet = wb.sheets["Regression Model"]
                    regression_rows.extend(
                        extract_regression_candidates(
                            wb=wb,
                            sheet=regression_sheet,
                            file_label=metadata,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"Skipped Regression Model in {file_path.name} (sheet missing)")

                processed_count += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    output_path = next_output_path(input_dir.resolve().name, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
