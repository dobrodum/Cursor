from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ===== User-configurable paths =====
input_dir = Path("input")
output_dir = Path("output")


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

N_QUARTERS = 10
ROW_SCAN_LIMIT = 500
HEADER_SCAN_WIDTH = 30

PERIOD_DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

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


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetCache:
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    values: List[List[Any]]

    def value(self, row: int, col: int) -> Any:
        if row < self.first_row or col < self.first_col:
            return None
        if row > self.last_row or col > self.last_col:
            return None
        return self.values[row - self.first_row][col - self.first_col]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def parse_file_label(file_name: str) -> Optional[FileLabel]:
    # Example: MedMiner_Model - AORT - MidJan2026_Send.xlsx
    pattern = re.compile(
        r"\s-\s(?P<ticker>[A-Za-z0-9]+)\s-\s(?P<period>(?P<timing>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4}))",
        re.IGNORECASE,
    )
    match = pattern.search(file_name)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    timing = match.group("timing").title()
    month_token = match.group("month")[:3].lower()
    year = int(match.group("year"))

    month_number = MONTH_MAP.get(month_token)
    day = PERIOD_DAY_MAP.get(timing.lower())
    if month_number is None or day is None:
        return None

    model_period = f"{timing}{calendar.month_abbr[month_number]}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def ensure_2d(values: Any, n_rows: int, n_cols: int) -> List[List[Any]]:
    if n_rows <= 0 or n_cols <= 0:
        return [[]]

    if n_rows == 1 and n_cols == 1:
        return [[values]]

    if values is None:
        return [[None for _ in range(n_cols)] for _ in range(n_rows)]

    if n_rows == 1:
        if isinstance(values, list):
            return [values[:n_cols] + [None] * max(0, n_cols - len(values))]
        return [[values] + [None] * (n_cols - 1)]

    if n_cols == 1:
        if isinstance(values, list):
            return [[item] for item in values[:n_rows]]
        return [[values]] + [[None] for _ in range(n_rows - 1)]

    if isinstance(values, list) and values and isinstance(values[0], list):
        matrix: List[List[Any]] = []
        for row in values[:n_rows]:
            row_list = row[:n_cols] + [None] * max(0, n_cols - len(row))
            matrix.append(row_list)
        if len(matrix) < n_rows:
            matrix.extend([[None for _ in range(n_cols)] for _ in range(n_rows - len(matrix))])
        return matrix

    if isinstance(values, list):
        # Fallback: flatten row-major into expected shape.
        flat = values[: n_rows * n_cols] + [None] * max(0, n_rows * n_cols - len(values))
        matrix = []
        for start in range(0, len(flat), n_cols):
            matrix.append(flat[start : start + n_cols])
        return matrix

    return [[values] + [None] * (n_cols - 1)] + [[None for _ in range(n_cols)] for _ in range(n_rows - 1)]


def build_sheet_cache(sheet: xw.main.Sheet) -> Optional[SheetCache]:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    last_row = used.last_cell.row
    last_col = used.last_cell.column
    if first_row is None or first_col is None or last_row is None or last_col is None:
        return None
    values = sheet.range((first_row, first_col), (last_row, last_col)).value
    n_rows = last_row - first_row + 1
    n_cols = last_col - first_col + 1
    matrix = ensure_2d(values, n_rows, n_cols)
    return SheetCache(
        first_row=first_row,
        first_col=first_col,
        last_row=last_row,
        last_col=last_col,
        values=matrix,
    )


def find_anchor(cache: SheetCache, target: str = "max") -> Optional[Tuple[int, int]]:
    target_norm = normalize_label(target)
    for r_idx, row in enumerate(cache.values):
        for c_idx, value in enumerate(row):
            if normalize_label(value) == target_norm:
                return cache.first_row + r_idx, cache.first_col + c_idx
    return None


def find_sheet_case_insensitive(wb: xw.main.Book, target_name: str) -> Optional[xw.main.Sheet]:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_header_columns(cache: SheetCache, header_row: int, anchor_col: int) -> Dict[str, int]:
    aliases: Dict[str, Sequence[str]] = {
        "quarter_label": ("quarter", "qtr", "period"),
        "num_quarters_used": ("num quarters", "quarters used", "num qtrs", "# quarters", "n quarters"),
        "last_quarter_used": ("last quarter", "last qtr"),
        "quarterly_sales": ("quarterly sales", "quarter sales", "db sales"),
        "reported_sales": ("reported sales", "actual sales"),
        "growth_rate_pct": ("growth rate", "growth %"),
        "sales_captured_in_db_pct": ("captured in db", "sales captured", "penetration"),
        "forecast_value": ("estimated total sold", "tot fcst", "forecast"),
        "forecast_min": ("min",),
        "forecast_max": ("max",),
    }

    out: Dict[str, int] = {}
    min_col = max(cache.first_col, anchor_col - HEADER_SCAN_WIDTH)
    max_col = min(cache.last_col, anchor_col + HEADER_SCAN_WIDTH)
    for col in range(min_col, max_col + 1):
        label = normalize_label(cache.value(header_row, col))
        if not label:
            continue
        for key, options in aliases.items():
            if key in out:
                continue
            if any(opt in label for opt in options):
                out[key] = col
    if "forecast_max" not in out:
        out["forecast_max"] = anchor_col
    return out


def collect_numeric_rows(
    cache: SheetCache,
    start_row: int,
    col_a: int,
    col_b: Optional[int] = None,
    max_rows: int = ROW_SCAN_LIMIT,
) -> List[int]:
    rows: List[int] = []
    blank_run = 0
    row = start_row
    hard_stop = min(cache.last_row, start_row + max_rows)
    while row <= hard_stop:
        a = to_float(cache.value(row, col_a))
        b = to_float(cache.value(row, col_b)) if col_b is not None else 0.0
        is_valid = a is not None and (col_b is None or b is not None)
        if is_valid:
            rows.append(row)
            blank_run = 0
        else:
            blank_run += 1
            if blank_run >= 8 and rows:
                break
        row += 1
    return rows


def choose_empirical_series_columns(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    headers: Dict[str, int],
) -> Tuple[int, int]:
    candidates: List[Tuple[int, int]] = []
    q_col = headers.get("quarterly_sales")
    r_col = headers.get("reported_sales")
    if q_col and r_col:
        candidates.append((q_col, r_col))

    fallback_pairs = [
        (anchor_col - 7, anchor_col - 11),
        (anchor_col - 11, anchor_col - 7),
        (anchor_col - 8, anchor_col - 12),
        (anchor_col - 9, anchor_col - 13),
        (anchor_col - 6, anchor_col - 10),
    ]
    candidates.extend(fallback_pairs)

    best = candidates[0]
    best_count = -1
    for first_col, second_col in candidates:
        if first_col < cache.first_col or second_col < cache.first_col:
            continue
        count = len(collect_numeric_rows(cache, anchor_row + 1, first_col, second_col))
        if count > best_count:
            best = (first_col, second_col)
            best_count = count
    return best


def safe_close_without_save(wb: xw.main.Book) -> None:
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
        # Final fallback: try default close without save arguments.
        try:
            wb.close()
        except Exception:
            pass


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def residual_stddev(x_values: Sequence[float], y_values: Sequence[float], intercept: float, slope: float) -> float:
    residuals: List[float] = []
    for x_value, y_value in zip(x_values, y_values):
        residuals.append(y_value - (intercept + slope * x_value))
    if len(residuals) < 2:
        return 0.0
    mean = sum(residuals) / len(residuals)
    variance = sum((value - mean) ** 2 for value in residuals) / (len(residuals) - 1)
    return math.sqrt(variance)


def to_cell_output(value: Any) -> Any:
    if value is None:
        return ""
    return value


def extract_empirical_rows(
    wb: xw.main.Book,
    label: FileLabel,
    file_name: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        return rows

    cache = build_sheet_cache(sheet)
    if cache is None:
        return rows

    anchor = find_anchor(cache, target="max")
    if anchor is None:
        return rows
    anchor_row, anchor_col = anchor

    headers = find_header_columns(cache, anchor_row, anchor_col)
    quarterly_col, reported_col = choose_empirical_series_columns(cache, anchor_row, anchor_col, headers)
    quarter_rows = collect_numeric_rows(cache, anchor_row + 1, quarterly_col, reported_col)
    if not quarter_rows:
        return rows

    quarter_label_col = headers.get("quarter_label")
    formula_col = max(cache.last_col + 2, anchor_col + 2)
    formula_row = max(cache.last_row + 2, anchor_row + 2)
    avg_pen_cell = sheet.range((formula_row, formula_col))
    max_pen_cell = sheet.range((formula_row, formula_col + 1))
    min_pen_cell = sheet.range((formula_row, formula_col + 2))

    max_quarters = min(N_QUARTERS, len(quarter_rows))
    for num_quarters in range(1, max_quarters + 1):
        window_rows = quarter_rows[-num_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        avg_formula = (
            f'=IFERROR(AVERAGE((R{start_row}C{quarterly_col}:R{end_row}C{quarterly_col})/'
            f'(R{start_row}C{reported_col}:R{end_row}C{reported_col})),"")'
        )
        max_formula = (
            f'=IFERROR(MAX((R{start_row}C{quarterly_col}:R{end_row}C{quarterly_col})/'
            f'(R{start_row}C{reported_col}:R{end_row}C{reported_col})),"")'
        )
        min_formula = (
            f'=IFERROR(MIN((R{start_row}C{quarterly_col}:R{end_row}C{quarterly_col})/'
            f'(R{start_row}C{reported_col}:R{end_row}C{reported_col})),"")'
        )

        set_formula2(avg_pen_cell, avg_formula)
        set_formula2(max_pen_cell, max_formula)
        set_formula2(min_pen_cell, min_formula)
        wb.app.calculate()

        avg_pen = to_float(avg_pen_cell.value)
        max_pen = to_float(max_pen_cell.value)
        min_pen = to_float(min_pen_cell.value)

        quarterly_sales = to_float(cache.value(end_row, quarterly_col))
        reported_sales = to_float(cache.value(end_row, reported_col))
        first_quarter_sales = to_float(cache.value(start_row, quarterly_col))

        forecast_value = None
        if quarterly_sales is not None and avg_pen not in (None, 0):
            forecast_value = quarterly_sales / avg_pen

        forecast_max = None
        if quarterly_sales is not None and min_pen not in (None, 0):
            forecast_max = quarterly_sales / min_pen

        forecast_min = None
        if quarterly_sales is not None and max_pen not in (None, 0):
            forecast_min = quarterly_sales / max_pen

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        growth_rate_pct = None
        if (
            first_quarter_sales is not None
            and quarterly_sales is not None
            and first_quarter_sales not in (0.0, -0.0)
        ):
            growth_rate_pct = (quarterly_sales / first_quarter_sales) - 1.0

        last_quarter_raw = cache.value(end_row, quarter_label_col) if quarter_label_col else end_row
        last_quarter_used = str(last_quarter_raw) if last_quarter_raw is not None else ""

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": num_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,  # estimated total sold
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": avg_pen,
            "source_file": file_name,
        }
        rows.append(row)

    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    label: FileLabel,
    file_name: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        return rows

    cache = build_sheet_cache(sheet)
    if cache is None:
        return rows

    anchor = find_anchor(cache, target="max")
    if anchor is None:
        return rows
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < cache.first_col or x_col < cache.first_col:
        return rows

    data_rows = collect_numeric_rows(cache, anchor_row + 1, x_col, y_col)
    if len(data_rows) < 2:
        return rows

    formula_col = max(cache.last_col + 2, anchor_col + 2)
    formula_row = max(cache.last_row + 2, anchor_row + 2)
    intercept_cell = sheet.range((formula_row, formula_col))
    slope_cell = sheet.range((formula_row, formula_col + 1))

    max_quarters = min(N_QUARTERS, len(data_rows))
    previous_signature: Optional[Tuple[Optional[float], ...]] = None
    for num_quarters in range(2, max_quarters + 1):
        window_rows = data_rows[-num_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        x_values: List[float] = []
        y_values: List[float] = []
        for row in window_rows:
            x_value = to_float(cache.value(row, x_col))
            y_value = to_float(cache.value(row, y_col))
            if x_value is None or y_value is None:
                continue
            x_values.append(x_value)
            y_values.append(y_value)

        if len(x_values) < 2 or len(y_values) < 2:
            continue

        forecast_x = to_float(cache.value(end_row + 1, x_col))
        if forecast_x is None:
            forecast_x = x_values[-1]

        forecast_total_without_sa = intercept + slope * forecast_x
        error_band = residual_stddev(x_values, y_values, intercept, slope)
        forecast_max = forecast_total_without_sa + error_band
        forecast_min = forecast_total_without_sa - error_band
        range_width = forecast_max - forecast_min

        signature: Tuple[Optional[float], ...] = (
            round(forecast_total_without_sa, 8),
            round(forecast_max, 8),
            round(forecast_min, 8),
            round(intercept, 8),
            round(slope, 8),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_total_without_sa,  # TOT FCST w/o SA
            "actual_value": "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": file_name,
        }
        rows.append(row)

    return rows


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([to_cell_output(row.get(column)) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row in rows:
            value = row.get(column_name)
            if value is None:
                continue
            length = len(str(value))
            if length > max_len:
                max_len = length
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def next_output_path(input_folder: Path, out_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def process_single_workbook(
    app: xw.App,
    file_path: Path,
    label: FileLabel,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    wb = app.books.open(str(file_path), update_links=False)
    try:
        empirical_rows = extract_empirical_rows(wb, label, file_path.name)
        regression_rows = extract_regression_rows(wb, label, file_path.name)
        return empirical_rows, regression_rows
    finally:
        safe_close_without_save(wb)


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    all_entries = sorted(in_dir.iterdir(), key=lambda path: path.name.lower())
    source_files = [path for path in all_entries if path.is_file()]
    output_path = next_output_path(in_dir, out_dir)

    processed_files = 0
    skipped_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in source_files:
            if file_path.suffix.lower() != ".xlsx":
                skipped_files += 1
                print(f"[SKIP] {file_path.name} -> not an .xlsx file")
                continue
            if file_path.name.startswith("~"):
                skipped_files += 1
                print(f"[SKIP] {file_path.name} -> temporary file")
                continue

            label = parse_file_label(file_path.name)
            if label is None:
                skipped_files += 1
                print(f"[SKIP] {file_path.name} -> file name format not recognized")
                continue

            print(f"[PROCESS] {file_path.name}")
            try:
                empirical_part, regression_part = process_single_workbook(app, file_path, label)
                empirical_rows.extend(empirical_part)
                regression_rows.extend(regression_part)
                processed_files += 1
            except Exception as exc:
                skipped_files += 1
                print(f"[SKIP] {file_path.name} -> processing error: {exc}")
    finally:
        try:
            app.quit()
        except Exception:
            pass

    output_wb = Workbook()
    empirical_ws = output_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = output_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    output_wb.save(output_path)

    print(f"[OUTPUT] {output_path}")
    print(f"[SUMMARY] files_processed={processed_files}")
    print(f"[SUMMARY] files_skipped={skipped_files}")
    print(f"[SUMMARY] empirical_rows={len(empirical_rows)}")
    print(f"[SUMMARY] regression_rows={len(regression_rows)}")


if __name__ == "__main__":
    main()
