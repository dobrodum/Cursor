#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --------- Configure these two paths ---------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")
# ---------------------------------------------

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

TIMING_DAY = {"early": 5, "mid": 15, "late": 25}

MONTH_NUM = {
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


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any, default: int) -> int:
    f = to_float(value)
    if f is None:
        return default
    return max(1, int(round(f)))


def has_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str) and value.strip() == "":
        return False
    return True


def safe_diff(a: Any, b: Any) -> Optional[float]:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def set_formula2(cell: Any, formula_r1c1: str) -> bool:
    try:
        cell.formula2 = formula_r1c1
        return True
    except Exception:
        try:
            cell.formula = formula_r1c1
            return True
        except Exception:
            return False


def safe_close_workbook(wb: Any) -> None:
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
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_file_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    ticker = ""
    model_period = ""
    model_date = ""
    model = stem

    parts = [part.strip() for part in stem.split(" - ")]
    period_token = ""
    if len(parts) >= 3:
        ticker = parts[1]
        period_token = parts[2].split("_")[0].strip()
    elif len(parts) == 2:
        ticker = parts[1]

    match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, flags=re.IGNORECASE)
    if match:
        timing_raw = match.group(1)
        month_raw = match.group(2)
        year_raw = match.group(3)

        timing = timing_raw[0].upper() + timing_raw[1:].lower()
        month_key = month_raw[:3].lower()
        month_num = MONTH_NUM.get(month_key)

        if month_num is not None:
            month_abbr = month_key.title()
            year_int = int(year_raw)
            day = TIMING_DAY[timing.lower()]
            model_period = f"{timing}{month_abbr}_{year_raw}"
            model_date = date(year_int, month_num, day).isoformat()
            if ticker:
                model = f"{ticker}_{model_period}"

    if not ticker and len(parts) >= 2:
        ticker = parts[1]

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


class SheetCache:
    def __init__(self, sheet: Any) -> None:
        used = sheet.used_range
        self.sheet = sheet
        self.start_row = used.row
        self.start_col = used.column
        self.values = ensure_2d(used.value)
        self.row_count = len(self.values)
        self.col_count = max((len(row) for row in self.values), default=0)
        self.end_row = self.start_row + self.row_count - 1
        self.end_col = self.start_col + self.col_count - 1

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= self.row_count:
            return None
        row_values = self.values[row_idx]
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]

    def header_norm_map(self, row: int) -> Dict[int, str]:
        return {
            col: normalize_header(self.get(row, col))
            for col in range(self.start_col, self.end_col + 1)
        }


def find_anchor_max(cache: SheetCache) -> Optional[Tuple[int, int]]:
    for row in range(cache.start_row, cache.end_row + 1):
        for col in range(cache.start_col, cache.end_col + 1):
            value = cache.get(row, col)
            if isinstance(value, str) and value.strip().lower() == "max":
                return row, col
    return None


def find_col_by_keywords(header_map: Dict[int, str], phrase_options: Sequence[Sequence[str]]) -> Optional[int]:
    for tokens in phrase_options:
        for col, normalized in header_map.items():
            if all(token in normalized for token in tokens):
                return col
    return None


def value_or_blank(value: Any) -> Any:
    if value is None:
        return ""
    return value


def extract_empirical_rows(
    wb: Any,
    model_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped {source_file}: missing sheet 'Empirical Model'")
        return []

    cache = SheetCache(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"Skipped {source_file}: no 'max' anchor on 'Empirical Model'")
        return []

    anchor_row, anchor_col = anchor
    header_map = cache.header_norm_map(anchor_row)

    num_quarters_col = find_col_by_keywords(
        header_map,
        [
            ("num", "quarter"),
            ("n", "quarter"),
            ("quarter", "used"),
        ],
    )
    last_quarter_col = find_col_by_keywords(
        header_map,
        [
            ("last", "quarter"),
            ("latest", "quarter"),
        ],
    )
    forecast_col = find_col_by_keywords(
        header_map,
        [
            ("estimated", "total"),
            ("forecast", "value"),
            ("forecast", "total"),
            ("tot", "fcst"),
        ],
    )
    reported_sales_col = find_col_by_keywords(
        header_map,
        [
            ("reported", "sales"),
            ("actual", "sales"),
        ],
    )
    avg_pen_col = find_col_by_keywords(
        header_map,
        [
            ("avg", "penetration"),
            ("average", "penetration"),
        ],
    )
    quarterly_sales_col = find_col_by_keywords(
        header_map,
        [
            ("quarterly", "sales"),
            ("qtr", "sales"),
        ],
    )
    growth_col = find_col_by_keywords(
        header_map,
        [
            ("growth", "rate"),
            ("growth", "pct"),
        ],
    )
    captured_col = find_col_by_keywords(
        header_map,
        [
            ("sales", "captured", "db"),
            ("captured", "db"),
            ("penetration", "pct"),
        ],
    )
    min_col = find_col_by_keywords(
        header_map,
        [
            ("min",),
        ],
    ) or (anchor_col + 1)

    history_penetration_col = captured_col if captured_col is not None else max(cache.start_col, anchor_col - 5)
    history_rows = [
        row
        for row in range(cache.start_row, anchor_row)
        if to_float(cache.get(row, history_penetration_col)) is not None
    ]

    helper_avg_col = cache.end_col + 2
    helper_fcst_col = cache.end_col + 3
    formula_rows: List[Tuple[int, int]] = []
    wrote_formulas = False

    for idx in range(10):
        row = anchor_row + 1 + idx
        quarters_val = cache.get(row, num_quarters_col) if num_quarters_col is not None else None
        n_quarters = to_int(quarters_val, idx + 1)
        formula_rows.append((row, n_quarters))

        if not history_rows:
            continue

        lookback = min(n_quarters, len(history_rows))
        used_rows = history_rows[-lookback:]
        start_row = used_rows[0]
        end_row = used_rows[-1]

        avg_formula = f'=IFERROR(AVERAGE(R{start_row}C{history_penetration_col}:R{end_row}C{history_penetration_col}),"")'
        wrote_formulas = set_formula2(sheet.cells(row, helper_avg_col), avg_formula) or wrote_formulas

        if reported_sales_col is not None:
            fcst_formula = f'=IFERROR(R{row}C{reported_sales_col}/R{row}C{helper_avg_col},"")'
            wrote_formulas = set_formula2(sheet.cells(row, helper_fcst_col), fcst_formula) or wrote_formulas

    if wrote_formulas:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    blank_streak = 0

    for row, n_quarters in formula_rows:
        avg_pen = cache.get(row, avg_pen_col) if avg_pen_col is not None else None
        if not has_value(avg_pen):
            avg_pen = sheet.cells(row, helper_avg_col).value

        forecast_value = cache.get(row, forecast_col) if forecast_col is not None else None
        if not has_value(forecast_value):
            forecast_value = sheet.cells(row, helper_fcst_col).value

        reported_sales = cache.get(row, reported_sales_col) if reported_sales_col is not None else None
        quarterly_sales = cache.get(row, quarterly_sales_col) if quarterly_sales_col is not None else None
        growth_rate = cache.get(row, growth_col) if growth_col is not None else None
        sales_captured = cache.get(row, captured_col) if captured_col is not None else None
        last_quarter = cache.get(row, last_quarter_col) if last_quarter_col is not None else None
        forecast_max = cache.get(row, anchor_col)
        forecast_min = cache.get(row, min_col)

        if not any(
            has_value(v)
            for v in (
                avg_pen,
                forecast_value,
                reported_sales,
                forecast_max,
                forecast_min,
                quarterly_sales,
                sales_captured,
            )
        ):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0

        range_width = safe_diff(forecast_max, forecast_min)

        row_payload = {
            "model": model_meta["model"],
            "ticker": model_meta["ticker"],
            "model_period": model_meta["model_period"],
            "model_date": model_meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": value_or_blank(avg_pen),
            "num_quarters_used": n_quarters,
            "last_quarter_used": value_or_blank(last_quarter),
            "forecast_value": value_or_blank(forecast_value),
            "actual_value": value_or_blank(reported_sales),
            "forecast_max": value_or_blank(forecast_max),
            "forecast_min": value_or_blank(forecast_min),
            "range_width": value_or_blank(range_width),
            "avg_penetration_pct": value_or_blank(avg_pen),
            "quarterly_sales": value_or_blank(quarterly_sales),
            "reported_sales": value_or_blank(reported_sales),
            "growth_rate_pct": value_or_blank(growth_rate),
            "sales_captured_in_db_pct": value_or_blank(sales_captured),
            "source_file": source_file,
        }
        rows.append(row_payload)

    return rows


def extract_regression_rows(
    wb: Any,
    model_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped {source_file}: missing sheet 'Regression Model'")
        return []

    cache = SheetCache(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"Skipped {source_file}: no 'max' anchor on 'Regression Model'")
        return []

    anchor_row, anchor_col = anchor
    header_map = cache.header_norm_map(anchor_row)

    num_quarters_col = find_col_by_keywords(
        header_map,
        [
            ("num", "quarter"),
            ("n", "quarter"),
            ("quarter", "used"),
        ],
    )
    forecast_col = find_col_by_keywords(
        header_map,
        [
            ("tot", "fcst", "sa"),
            ("forecast", "without", "sa"),
            ("forecast", "value"),
            ("forecast", "total"),
        ],
    )
    actual_col = find_col_by_keywords(
        header_map,
        [
            ("actual", "sales"),
            ("reported", "sales"),
        ],
    )
    min_col = find_col_by_keywords(
        header_map,
        [
            ("min",),
        ],
    ) or (anchor_col + 1)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    history_rows = [
        row
        for row in range(cache.start_row, anchor_row)
        if to_float(cache.get(row, x_col)) is not None and to_float(cache.get(row, y_col)) is not None
    ]

    helper_intercept_col = cache.end_col + 2
    helper_slope_col = cache.end_col + 3
    helper_fcst_col = cache.end_col + 4

    formula_rows: List[Tuple[int, int]] = []
    wrote_formulas = False

    for idx in range(10):
        row = anchor_row + 1 + idx
        quarters_val = cache.get(row, num_quarters_col) if num_quarters_col is not None else None
        n_quarters = to_int(quarters_val, idx + 1)
        formula_rows.append((row, n_quarters))

        if len(history_rows) < 2:
            continue

        lookback = min(max(n_quarters, 2), len(history_rows))
        used_rows = history_rows[-lookback:]
        start_row = used_rows[0]
        end_row = used_rows[-1]

        intercept_formula = f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        slope_formula = f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        wrote_formulas = set_formula2(sheet.cells(row, helper_intercept_col), intercept_formula) or wrote_formulas
        wrote_formulas = set_formula2(sheet.cells(row, helper_slope_col), slope_formula) or wrote_formulas

        if forecast_col is None:
            x_pred = to_float(cache.get(row, x_col))
            if x_pred is None:
                x_pred = to_float(cache.get(history_rows[-1], x_col))
            if x_pred is not None:
                fcst_formula = f'=IFERROR(R{row}C{helper_intercept_col}+R{row}C{helper_slope_col}*{x_pred},"")'
                wrote_formulas = set_formula2(sheet.cells(row, helper_fcst_col), fcst_formula) or wrote_formulas

    if wrote_formulas:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None
    blank_streak = 0

    for row, n_quarters in formula_rows:
        forecast_value = cache.get(row, forecast_col) if forecast_col is not None else None
        if not has_value(forecast_value):
            forecast_value = sheet.cells(row, helper_fcst_col).value

        actual_value = cache.get(row, actual_col) if actual_col is not None else ""
        forecast_max = cache.get(row, anchor_col)
        forecast_min = cache.get(row, min_col)
        intercept = sheet.cells(row, helper_intercept_col).value
        slope = sheet.cells(row, helper_slope_col).value

        if not any(
            has_value(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
                actual_value,
            )
        ):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0

        signature = (
            n_quarters,
            round(to_float(forecast_value) or 0.0, 10),
            round(to_float(forecast_max) or 0.0, 10),
            round(to_float(forecast_min) or 0.0, 10),
            round(to_float(intercept) or 0.0, 10),
            round(to_float(slope) or 0.0, 10),
        )
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        range_width = safe_diff(forecast_max, forecast_min)
        row_payload = {
            "model": model_meta["model"],
            "ticker": model_meta["ticker"],
            "model_period": model_meta["model_period"],
            "model_date": model_meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": value_or_blank(forecast_value),
            "actual_value": value_or_blank(actual_value),
            "forecast_max": value_or_blank(forecast_max),
            "forecast_min": value_or_blank(forecast_min),
            "range_width": value_or_blank(range_width),
            "intercept": value_or_blank(intercept),
            "slope": value_or_blank(slope),
            "source_file": source_file,
        }
        rows.append(row_payload)

    return rows


def write_sheet(wb: Workbook, title: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws = wb.create_sheet(title=title)
    ws.append(list(columns))
    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 60)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary lock file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    output_path = unique_output_path(input_dir, output_dir)

    app: Optional[Any] = None
    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in iter_input_files(input_dir):
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open workbook ({exc})")
                continue

            file_meta = parse_file_label(file_path)
            try:
                empirical_rows.extend(extract_empirical_rows(wb, file_meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, file_meta, file_path.name))
                processed_files += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    output_book = Workbook()
    output_book.remove(output_book.active)
    write_sheet(output_book, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_book, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_book.save(output_path)

    print(f"Output file: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
