from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable paths
# ----------------------------
input_dir = Path("./input")
output_dir = Path("./output")

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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
MONTH_ABBR = {
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


def log(msg: str) -> None:
    print(msg)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    pct = text.endswith("%")
    if pct:
        text = text[:-1]

    text = text.replace(",", "")
    try:
        parsed = float(text)
    except ValueError:
        return None
    return parsed / 100.0 if pct else parsed


def clean_value(value: Any) -> Any:
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            return int(round(value))
        return value
    return value


def safe_sheet_name(sheet_name: str) -> str:
    return sheet_name.strip().lower()


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split("-")]

    ticker = ""
    period_chunk = stem
    if len(parts) >= 3:
        ticker = re.sub(r"\s+", "", parts[1])
        period_chunk = parts[2]
    elif len(parts) >= 2:
        ticker = re.sub(r"\s+", "", parts[1])

    ticker = ticker or "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})", period_chunk, flags=re.IGNORECASE
    )

    model_period = "UNKNOWN"
    model_date = ""

    if period_match:
        period_word = period_match.group(1).title()
        month_word = period_match.group(2).title()
        year_word = period_match.group(3)

        model_period = f"{period_word}{month_word}_{year_word}"

        month_num = MONTH_ABBR.get(month_word.lower())
        day_num = DAY_BY_PERIOD.get(period_word.lower())
        if month_num and day_num:
            model_date = date(int(year_word), month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def get_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    input_folder_name = source_input_dir.resolve().name
    target_output_dir.mkdir(parents=True, exist_ok=True)

    base_path = target_output_dir / f"{input_folder_name}_PARAM.xlsx"
    if not base_path.exists():
        return base_path

    i = 1
    while True:
        candidate = target_output_dir / f"{input_folder_name}_PARAM.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return [[]]
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return [[]]

    first = values[0]
    if isinstance(first, list):
        return values
    return [values]


class SheetCache:
    def __init__(self, sheet: xw.main.Sheet):
        self.sheet = sheet
        used = sheet.used_range
        self.start_row = used.row
        self.start_col = used.column
        self.values = ensure_2d(used.value)
        self.n_rows = len(self.values)
        self.n_cols = max((len(r) for r in self.values), default=0)
        self.text_cells: List[Tuple[int, int, str, str]] = []
        self._build_text_index()

    def _build_text_index(self) -> None:
        for r_idx, row in enumerate(self.values):
            for c_idx, value in enumerate(row):
                if isinstance(value, str) and value.strip():
                    abs_row = self.start_row + r_idx
                    abs_col = self.start_col + c_idx
                    text = value.strip()
                    norm = normalize_text(text)
                    self.text_cells.append((abs_row, abs_col, text, norm))

    def get_value(self, row: int, col: int) -> Any:
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if 0 <= r_idx < self.n_rows and 0 <= c_idx < len(self.values[r_idx]):
            return self.values[r_idx][c_idx]
        return self.sheet.range((row, col)).value


def find_anchor_max(cache: SheetCache) -> Optional[Tuple[int, int]]:
    for row, col, _raw, norm in cache.text_cells:
        if norm == "max":
            return row, col
    return None


def find_label_cell_near(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    labels: Iterable[str],
    row_window: int = 80,
    col_window: int = 40,
) -> Optional[Tuple[int, int, str]]:
    label_norms = [normalize_text(lbl) for lbl in labels if lbl]
    if not label_norms:
        return None

    best: Optional[Tuple[int, int, str, int]] = None
    for row, col, raw, norm in cache.text_cells:
        if abs(row - anchor_row) > row_window or abs(col - anchor_col) > col_window:
            continue
        if not any(lbl in norm for lbl in label_norms):
            continue
        score = abs(row - anchor_row) * 3 + abs(col - anchor_col)
        if best is None or score < best[3]:
            best = (row, col, raw, score)

    if best is None:
        return None
    return best[0], best[1], best[2]


def read_numeric_adjacent(
    cache: SheetCache,
    row: int,
    col: int,
    preferred_offsets: Optional[List[Tuple[int, int]]] = None,
) -> Optional[float]:
    offsets = preferred_offsets or [
        (0, 1),
        (0, -1),
        (1, 0),
        (-1, 0),
        (1, 1),
        (1, -1),
        (-1, 1),
        (-1, -1),
        (0, 2),
        (2, 0),
    ]

    for r_off, c_off in offsets:
        value = cache.get_value(row + r_off, col + c_off)
        num = to_float(value)
        if num is not None:
            return num
    return None


def find_numeric_from_labels(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    labels: Iterable[str],
) -> Optional[float]:
    loc = find_label_cell_near(cache, anchor_row, anchor_col, labels)
    if not loc:
        return None
    row, col, _ = loc
    return read_numeric_adjacent(cache, row, col)


def find_value_cell_near_label(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    labels: Iterable[str],
) -> Optional[Tuple[int, int]]:
    loc = find_label_cell_near(cache, anchor_row, anchor_col, labels)
    if not loc:
        return None
    row, col, _ = loc
    for r_off, c_off in [(0, 1), (0, -1), (1, 0), (-1, 0)]:
        target_row = row + r_off
        target_col = col + c_off
        value = cache.get_value(target_row, target_col)
        if value not in (None, ""):
            return target_row, target_col
    return None


def collect_numeric_series(
    cache: SheetCache,
    col: int,
    max_row: int,
    lookback_rows: int = 120,
) -> List[Tuple[int, float]]:
    start_row = max(cache.start_row, max_row - lookback_rows)
    values: List[Tuple[int, float]] = []
    for row in range(start_row, max_row + 1):
        num = to_float(cache.get_value(row, col))
        if num is not None:
            values.append((row, num))
    return values


def collect_text_series(
    cache: SheetCache,
    col: int,
    max_row: int,
    lookback_rows: int = 120,
) -> Dict[int, str]:
    start_row = max(cache.start_row, max_row - lookback_rows)
    out: Dict[int, str] = {}
    for row in range(start_row, max_row + 1):
        value = cache.get_value(row, col)
        if isinstance(value, str) and value.strip():
            out[row] = value.strip()
    return out


def set_formula2_r1c1(cell: xw.main.Range, formula_r1c1: str) -> None:
    api_cell = cell.api
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        api_cell.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        api_cell.FormulaR1C1 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def parse_sheet_name_case_insensitive(wb: xw.main.Book, name: str) -> Optional[xw.main.Sheet]:
    target = safe_sheet_name(name)
    for sheet in wb.sheets:
        if safe_sheet_name(sheet.name) == target:
            return sheet
    return None


def extract_empirical_rows(
    wb: xw.main.Book,
    file_labels: Dict[str, str],
    file_name: str,
) -> List[Dict[str, Any]]:
    sheet = parse_sheet_name_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        return []

    cache = SheetCache(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    quarter_col = anchor_col - 12
    penetration_col = anchor_col - 8
    quarterly_sales_col = anchor_col - 11
    reported_sales_col = anchor_col - 7
    growth_rate_col = anchor_col - 9
    sales_captured_col = anchor_col - 10

    # Label-based overrides (still anchored around max to avoid broad scanning).
    loc = find_label_cell_near(cache, anchor_row, anchor_col, ["avg penetration", "penetration"])
    if loc:
        penetration_col = loc[1]
    loc = find_label_cell_near(cache, anchor_row, anchor_col, ["quarterly sales", "qtr sales"])
    if loc:
        quarterly_sales_col = loc[1]
    loc = find_label_cell_near(cache, anchor_row, anchor_col, ["reported sales"])
    if loc:
        reported_sales_col = loc[1]
    loc = find_label_cell_near(cache, anchor_row, anchor_col, ["growth rate"])
    if loc:
        growth_rate_col = loc[1]
    loc = find_label_cell_near(cache, anchor_row, anchor_col, ["sales captured in db", "captured in db"])
    if loc:
        sales_captured_col = loc[1]

    penetration_series = collect_numeric_series(cache, penetration_col, anchor_row - 1)
    quarterly_series = dict(collect_numeric_series(cache, quarterly_sales_col, anchor_row - 1))
    reported_series = dict(collect_numeric_series(cache, reported_sales_col, anchor_row - 1))
    growth_series = dict(collect_numeric_series(cache, growth_rate_col, anchor_row - 1))
    captured_series = dict(collect_numeric_series(cache, sales_captured_col, anchor_row - 1))
    quarter_labels = collect_text_series(cache, quarter_col, anchor_row - 1)

    if not penetration_series:
        return []

    scratch_cell = sheet.range((anchor_row + 2, anchor_col + 2))
    max_cell = sheet.range((anchor_row, anchor_col + 1))
    num_quarters_cell = find_value_cell_near_label(
        cache,
        anchor_row,
        anchor_col,
        labels=["num quarters", "quarters used", "number of quarters"],
    )

    min_cell_loc = find_value_cell_near_label(cache, anchor_row, anchor_col, ["min"])
    forecast_cell_loc = find_value_cell_near_label(
        cache, anchor_row, anchor_col, ["estimated total sold", "est total sold", "tot fcst"]
    )
    actual_cell_loc = find_value_cell_near_label(cache, anchor_row, anchor_col, ["reported sales", "actual sales"])

    rows: List[Dict[str, Any]] = []
    n_quarter_cap = min(10, len(penetration_series))
    for n_quarters in range(1, n_quarter_cap + 1):
        subset = penetration_series[-n_quarters:]
        start_row = subset[0][0]
        end_row = subset[-1][0]

        if num_quarters_cell is not None:
            sheet.range(num_quarters_cell).value = n_quarters
        set_formula2_r1c1(scratch_cell, f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})")
        wb.app.calculate()
        avg_penetration_pct = to_float(scratch_cell.value)

        forecast_max = to_float(max_cell.value)
        if forecast_max is None:
            forecast_max = find_numeric_from_labels(cache, anchor_row, anchor_col, ["max"])

        forecast_min = None
        if min_cell_loc:
            forecast_min = to_float(sheet.range(min_cell_loc).value)
        if forecast_min is None:
            forecast_min = find_numeric_from_labels(cache, anchor_row, anchor_col, ["min"])

        forecast_value = None
        if forecast_cell_loc:
            forecast_value = to_float(sheet.range(forecast_cell_loc).value)
        if forecast_value is None:
            forecast_value = find_numeric_from_labels(
                cache, anchor_row, anchor_col, ["estimated total sold", "est total sold", "tot fcst"]
            )

        actual_value = None
        if actual_cell_loc:
            actual_value = to_float(sheet.range(actual_cell_loc).value)
        if actual_value is None:
            actual_value = find_numeric_from_labels(cache, anchor_row, anchor_col, ["reported sales", "actual sales"])

        last_row = end_row
        row_data = {
            "model": file_labels["model"],
            "ticker": file_labels["ticker"],
            "model_period": file_labels["model_period"],
            "model_date": file_labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": n_quarters,
            "last_quarter_used": quarter_labels.get(last_row, ""),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": (
                (forecast_max - forecast_min)
                if forecast_max is not None and forecast_min is not None
                else None
            ),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_series.get(last_row),
            "reported_sales": reported_series.get(last_row),
            "growth_rate_pct": growth_series.get(last_row),
            "sales_captured_in_db_pct": captured_series.get(last_row),
            "source_file": file_name,
        }
        rows.append(row_data)

    # Clear temporary formula to avoid leaving visible artifacts in source workbook.
    scratch_cell.value = None
    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    file_labels: Dict[str, str],
    file_name: str,
) -> List[Dict[str, Any]]:
    sheet = parse_sheet_name_case_insensitive(wb, "Regression Model")
    if sheet is None:
        return []

    cache = SheetCache(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    x_series = dict(collect_numeric_series(cache, x_col, anchor_row - 1))
    y_series = dict(collect_numeric_series(cache, y_col, anchor_row - 1))
    candidate_rows = sorted(set(x_series.keys()) & set(y_series.keys()))
    pairs = [(row, x_series[row], y_series[row]) for row in candidate_rows]
    if not pairs:
        return []

    intercept_cell = sheet.range((anchor_row + 2, anchor_col + 2))
    slope_cell = sheet.range((anchor_row + 3, anchor_col + 2))
    max_cell = sheet.range((anchor_row, anchor_col + 1))
    num_quarters_cell = find_value_cell_near_label(
        cache,
        anchor_row,
        anchor_col,
        labels=["num quarters", "quarters used", "number of quarters"],
    )
    min_cell_loc = find_value_cell_near_label(cache, anchor_row, anchor_col, ["min"])
    forecast_cell_loc = find_value_cell_near_label(
        cache, anchor_row, anchor_col, ["tot fcst w/o sa", "tot fcst wo sa", "total forecast without sa"]
    )
    actual_cell_loc = find_value_cell_near_label(cache, anchor_row, anchor_col, ["actual sales", "reported sales"])

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None
    n_quarter_cap = min(10, len(pairs))
    for n_quarters in range(1, n_quarter_cap + 1):
        subset = pairs[-n_quarters:]
        start_row = subset[0][0]
        end_row = subset[-1][0]

        if num_quarters_cell is not None:
            sheet.range(num_quarters_cell).value = n_quarters

        set_formula2_r1c1(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        forecast_max = to_float(max_cell.value)
        if forecast_max is None:
            forecast_max = find_numeric_from_labels(cache, anchor_row, anchor_col, ["max"])

        forecast_min = None
        if min_cell_loc:
            forecast_min = to_float(sheet.range(min_cell_loc).value)
        if forecast_min is None:
            forecast_min = find_numeric_from_labels(cache, anchor_row, anchor_col, ["min"])

        forecast_value = None
        if forecast_cell_loc:
            forecast_value = to_float(sheet.range(forecast_cell_loc).value)
        if forecast_value is None:
            forecast_value = find_numeric_from_labels(
                cache, anchor_row, anchor_col, ["tot fcst w/o sa", "tot fcst wo sa", "total forecast without sa"]
            )

        actual_value = None
        if actual_cell_loc:
            actual_value = to_float(sheet.range(actual_cell_loc).value)

        row_data = {
            "model": file_labels["model"],
            "ticker": file_labels["ticker"],
            "model_period": file_labels["model_period"],
            "model_date": file_labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": (
                (forecast_max - forecast_min)
                if forecast_max is not None and forecast_min is not None
                else None
            ),
            "intercept": intercept,
            "slope": slope,
            "source_file": file_name,
        }

        signature = (
            row_data["forecast_value"],
            row_data["forecast_max"],
            row_data["forecast_min"],
            row_data["intercept"],
            row_data["slope"],
        )
        if prev_signature is not None and signature == prev_signature:
            # Avoid duplicate terminal row as requested.
            continue
        prev_signature = signature
        rows.append(row_data)

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def close_source_workbook_safely(wb: xw.main.Book) -> None:
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


def auto_format_sheet(ws) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].font = Font(bold=True)

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(42, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    ws_empirical.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        ws_empirical.append([clean_value(row.get(col)) for col in EMPIRICAL_HEADERS])

    ws_regression.append(REGRESSION_HEADERS)
    for row in regression_rows:
        ws_regression.append([clean_value(row.get(col)) for col in REGRESSION_HEADERS])

    auto_format_sheet(ws_empirical)
    auto_format_sheet(ws_regression)
    wb.save(output_path)


def is_valid_input_file(path: Path, input_folder_name: str) -> Tuple[bool, str]:
    if not path.is_file():
        return False, "not a file"
    if path.name.startswith("~"):
        return False, "temp workbook"
    if path.suffix.lower() != ".xlsx":
        return False, "not .xlsx"
    if path.name.startswith(f"{input_folder_name}_PARAM"):
        return False, "output workbook pattern"
    return True, ""


def run() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir not found: {input_dir}")

    output_path = get_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    files_processed = 0
    input_folder_name = input_dir.resolve().name
    source_files = sorted(input_dir.iterdir())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    prior_calc = None
    try:
        try:
            prior_calc = app.calculation
            app.calculation = "manual"
        except Exception:
            prior_calc = None

        for file_path in source_files:
            ok, reason = is_valid_input_file(file_path, input_folder_name)
            if not ok:
                log(f"Skipped: {file_path.name} ({reason})")
                continue

            log(f"Processing: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)

                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                files_processed += 1
            except Exception as exc:
                log(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook_safely(wb)
    finally:
        try:
            if prior_calc is not None:
                app.calculation = prior_calc
        except Exception:
            pass
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    log(f"Output path: {output_path}")
    log(f"Files processed: {files_processed}")
    log(f"Empirical rows: {len(empirical_rows)}")
    log(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
