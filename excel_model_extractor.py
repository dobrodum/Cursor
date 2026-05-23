#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError:  # pragma: no cover - environment-specific dependency
    xw = None


# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("./input")
output_dir = Path("./output")


N_QUARTERS = 10

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


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(int(value))
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    is_percent = text.endswith("%")
    if is_percent:
        text = text[:-1].strip()
    try:
        number = float(text)
        return number / 100.0 if is_percent else number
    except ValueError:
        return None


def coerce_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def parse_model_meta(file_name: str) -> ModelMeta:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]

    ticker = "UNKNOWN"
    if len(parts) >= 2:
        ticker = parts[1].split("_")[0].strip().upper()
    else:
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = re.search(
        r"\b(Early|Mid|Late)\s*[-_ ]*\s*([A-Za-z]{3,9})\s*[-_ ]*\s*(20\d{2})\b",
        stem,
        flags=re.IGNORECASE,
    )

    if period_match:
        bucket_raw, month_raw, year_raw = period_match.groups()
        bucket = bucket_raw.title()
        month_abbrev = month_raw[:3].title()
        month_map = {
            "Jan": 1,
            "Feb": 2,
            "Mar": 3,
            "Apr": 4,
            "May": 5,
            "Jun": 6,
            "Jul": 7,
            "Aug": 8,
            "Sep": 9,
            "Oct": 10,
            "Nov": 11,
            "Dec": 12,
        }
        month_num = month_map.get(month_abbrev, 1)
        year = int(year_raw)
        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        model_day = day_map[bucket]
        model_period = f"{bucket}{month_abbrev}_{year}"
        model_date = date(year, month_num, model_day).isoformat()
    else:
        model_period = "UnknownPeriod"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, out_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = out_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_safe(wb: Any) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def get_sheet_ci(wb: Any, sheet_name: str) -> Optional[Any]:
    wanted = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == wanted:
            return sheet
    return None


def read_used_grid(sheet: Any) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    values = coerce_2d(used.value)
    return start_row, start_col, values


def find_anchor_max(start_row: int, start_col: int, grid: List[List[Any]]) -> Optional[Tuple[int, int]]:
    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def read_row_segment(sheet: Any, row: int, left_col: int, right_col: int) -> List[Any]:
    if right_col < left_col:
        return []
    values = sheet.range((row, left_col), (row, right_col)).value
    if not isinstance(values, list):
        return [values]
    return values


def build_header_candidates(
    sheet: Any,
    anchor_row: int,
    anchor_col: int,
    row_window: Sequence[int] = (-1, 0, 1),
    col_span: int = 30,
) -> List[Tuple[str, int, int]]:
    left_col = max(1, anchor_col - col_span)
    right_col = anchor_col + col_span
    headers: List[Tuple[str, int, int]] = []
    for offset in row_window:
        row = anchor_row + offset
        if row < 1:
            continue
        row_values = read_row_segment(sheet, row, left_col, right_col)
        for idx, value in enumerate(row_values):
            norm = normalize_text(value)
            if norm:
                headers.append((norm, row, left_col + idx))
    return headers


def resolve_column(
    headers: Sequence[Tuple[str, int, int]],
    anchor_row: int,
    anchor_col: int,
    keyword_groups: Sequence[Sequence[str]],
    fallback_offset: int,
) -> int:
    best: Optional[Tuple[int, int, int]] = None
    # Tuple shape for comparison: (row_distance, col_distance, col)
    for norm_text, hdr_row, hdr_col in headers:
        for group in keyword_groups:
            if all(token in norm_text for token in group):
                candidate = (abs(hdr_row - anchor_row), abs(hdr_col - anchor_col), hdr_col)
                if best is None or candidate < best:
                    best = candidate
    if best is not None:
        return best[2]
    return anchor_col + fallback_offset


def clamp_col(col: int) -> int:
    return max(1, col)


def first_data_row(sheet: Any, anchor_row: int, max_col: int, min_col: int) -> int:
    for candidate in range(anchor_row + 1, anchor_row + 5):
        max_value = to_float(sheet.range((candidate, max_col)).value)
        min_value = to_float(sheet.range((candidate, min_col)).value)
        if max_value is not None or min_value is not None:
            return candidate
    return anchor_row + 1


def safe_diff(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def clean_numeric(value: Any) -> Any:
    number = to_float(value)
    if number is None:
        return value if value not in ("", None) else None
    if abs(number - round(number)) < 1e-12:
        return int(round(number))
    return number


def almost_equal(a: Any, b: Any, tolerance: float = 1e-10) -> bool:
    if a is None and b is None:
        return True
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return a == b
    return abs(a_num - b_num) <= tolerance


def extract_empirical_candidates(wb: Any, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_ci(wb, "Empirical Model")
    if sheet is None:
        return []

    start_row, start_col, grid = read_used_grid(sheet)
    anchor = find_anchor_max(start_row, start_col, grid)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    headers = build_header_candidates(sheet, anchor_row, anchor_col)
    max_col = clamp_col(anchor_col)
    min_col = clamp_col(
        resolve_column(headers, anchor_row, anchor_col, (("min",),), fallback_offset=1)
    )
    num_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("num", "quarter"), ("quarters", "used"), ("n", "quarter")),
            fallback_offset=-10,
        )
    )
    last_q_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("last", "quarter"), ("quarter", "used")),
            fallback_offset=-9,
        )
    )
    forecast_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("estimated", "total", "sold"), ("est", "total", "sold"), ("forecast", "value")),
            fallback_offset=-4,
        )
    )
    actual_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("reported", "sales"), ("actual", "sales"), ("actual", "value")),
            fallback_offset=-3,
        )
    )
    avg_pen_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("avg", "penetration"), ("average", "penetration"), ("penetration", "pct")),
            fallback_offset=-5,
        )
    )
    quarterly_sales_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("quarterly", "sales"), ("quarter", "sales"), ("db", "sales")),
            fallback_offset=-8,
        )
    )
    reported_sales_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("reported", "sales"),),
            fallback_offset=-3,
        )
    )
    growth_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (("growth", "rate"), ("growth", "pct")),
            fallback_offset=-7,
        )
    )
    captured_col = clamp_col(
        resolve_column(
            headers,
            anchor_row,
            anchor_col,
            (
                ("sales", "captured", "db"),
                ("captured", "db"),
                ("captured", "pct"),
            ),
            fallback_offset=-6,
        )
    )

    data_start = first_data_row(sheet, anchor_row, max_col, min_col)
    helper_avg_col = anchor_col + 35
    staged_rows: List[Dict[str, Any]] = []
    formulas_written = False

    for idx in range(N_QUARTERS):
        row = data_start + idx
        read_cols = [
            num_col,
            last_q_col,
            forecast_col,
            actual_col,
            max_col,
            min_col,
            avg_pen_col,
            quarterly_sales_col,
            reported_sales_col,
            growth_col,
            captured_col,
        ]
        left_col = min(read_cols)
        right_col = max(read_cols)
        row_values = read_row_segment(sheet, row, left_col, right_col)

        def value_at(col: int) -> Any:
            return row_values[col - left_col]

        raw_forecast = value_at(forecast_col)
        raw_max = value_at(max_col)
        raw_min = value_at(min_col)
        raw_q_sales = value_at(quarterly_sales_col)
        raw_reported = value_at(reported_sales_col)
        raw_actual = value_at(actual_col)
        raw_captured = value_at(captured_col)
        raw_avg = value_at(avg_pen_col)

        all_empty = all(
            value in ("", None)
            for value in (
                raw_forecast,
                raw_max,
                raw_min,
                raw_q_sales,
                raw_reported,
                raw_actual,
                raw_captured,
                raw_avg,
            )
        )
        if all_empty:
            continue

        helper_formula = None
        if captured_col:
            helper_formula = (
                f'=IFERROR(AVERAGE(R{data_start}C{captured_col}:R{row}C{captured_col}),"")'
            )
        elif quarterly_sales_col and reported_sales_col:
            helper_formula = (
                f'=IFERROR('
                f'SUM(R{data_start}C{quarterly_sales_col}:R{row}C{quarterly_sales_col})/'
                f'SUM(R{data_start}C{reported_sales_col}:R{row}C{reported_sales_col}),'
                f'"")'
            )
        elif avg_pen_col:
            helper_formula = (
                f'=IFERROR(AVERAGE(R{data_start}C{avg_pen_col}:R{row}C{avg_pen_col}),"")'
            )

        if helper_formula:
            set_formula2(sheet.range((row, helper_avg_col)), helper_formula)
            formulas_written = True

        staged_rows.append(
            {
                "idx": idx,
                "row": row,
                "num_quarters_used": value_at(num_col),
                "last_quarter_used": value_at(last_q_col),
                "forecast_value_raw": raw_forecast,
                "actual_value_raw": raw_actual,
                "forecast_max_raw": raw_max,
                "forecast_min_raw": raw_min,
                "avg_pen_raw": raw_avg,
                "quarterly_sales_raw": raw_q_sales,
                "reported_sales_raw": raw_reported,
                "growth_rate_raw": value_at(growth_col),
                "captured_pct_raw": raw_captured,
                "avg_helper_row": row,
            }
        )

    if formulas_written:
        wb.app.calculate()

    results: List[Dict[str, Any]] = []
    for staged in staged_rows:
        row = staged["row"]
        helper_avg = (
            sheet.range((staged["avg_helper_row"], helper_avg_col)).value if formulas_written else None
        )
        avg_penetration = helper_avg if helper_avg not in ("", None) else staged["avg_pen_raw"]
        if avg_penetration in ("", None):
            avg_penetration = staged["captured_pct_raw"]

        forecast_value = staged["forecast_value_raw"]
        if forecast_value in ("", None):
            q_sales = to_float(staged["quarterly_sales_raw"])
            avg_pen_num = to_float(avg_penetration)
            if q_sales is not None and avg_pen_num not in (None, 0):
                forecast_value = q_sales / avg_pen_num

        actual_value = staged["actual_value_raw"]
        if actual_value in ("", None):
            actual_value = staged["reported_sales_raw"]

        forecast_max = staged["forecast_max_raw"]
        forecast_min = staged["forecast_min_raw"]

        num_q = staged["num_quarters_used"]
        num_q_num = to_float(num_q)
        if num_q_num is None:
            num_q = staged["idx"] + 1
        else:
            num_q = int(num_q_num) if abs(num_q_num - round(num_q_num)) < 1e-12 else num_q_num

        sales_captured_pct = staged["captured_pct_raw"]
        if sales_captured_pct in ("", None):
            q_sales = to_float(staged["quarterly_sales_raw"])
            reported = to_float(staged["reported_sales_raw"])
            if q_sales is not None and reported not in (None, 0):
                sales_captured_pct = q_sales / reported

        result = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": clean_numeric(avg_penetration),
            "num_quarters_used": clean_numeric(num_q),
            "last_quarter_used": staged["last_quarter_used"],
            "forecast_value": clean_numeric(forecast_value),
            "actual_value": clean_numeric(actual_value),
            "forecast_max": clean_numeric(forecast_max),
            "forecast_min": clean_numeric(forecast_min),
            "range_width": clean_numeric(safe_diff(forecast_max, forecast_min)),
            "avg_penetration_pct": clean_numeric(avg_penetration),
            "quarterly_sales": clean_numeric(staged["quarterly_sales_raw"]),
            "reported_sales": clean_numeric(staged["reported_sales_raw"]),
            "growth_rate_pct": clean_numeric(staged["growth_rate_raw"]),
            "sales_captured_in_db_pct": clean_numeric(sales_captured_pct),
            "source_file": source_file,
        }
        results.append(result)

    return results


def extract_regression_candidates(wb: Any, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_ci(wb, "Regression Model")
    if sheet is None:
        return []

    start_row, start_col, grid = read_used_grid(sheet)
    anchor = find_anchor_max(start_row, start_col, grid)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    if y_col < 1 or x_col < 1 or anchor_row <= 1:
        return []

    lookback = 600
    start_data_row = max(1, anchor_row - lookback)
    if start_data_row > anchor_row - 1:
        return []
    left = min(x_col, y_col)
    right = max(x_col, y_col)
    block = coerce_2d(sheet.range((start_data_row, left), (anchor_row - 1, right)).value)

    contiguous_rows: List[int] = []
    for idx in range(len(block) - 1, -1, -1):
        row_num = start_data_row + idx
        row_values = block[idx]
        x_value = to_float(row_values[x_col - left])
        y_value = to_float(row_values[y_col - left])
        if x_value is not None and y_value is not None:
            contiguous_rows.append(row_num)
        elif contiguous_rows:
            break
    contiguous_rows.reverse()

    if len(contiguous_rows) < 2:
        return []

    helper_base_col = anchor_col + 28
    helper_rows: List[Tuple[int, int]] = []
    formulas_written = False
    max_window = min(N_QUARTERS, len(contiguous_rows))

    for window in range(2, max_window + 1):
        helper_row = anchor_row + (window - 1)
        data_start = contiguous_rows[-window]
        data_end = contiguous_rows[-1]

        intercept_cell = sheet.range((helper_row, helper_base_col))
        slope_cell = sheet.range((helper_row, helper_base_col + 1))
        forecast_cell = sheet.range((helper_row, helper_base_col + 2))
        max_cell = sheet.range((helper_row, helper_base_col + 3))
        min_cell = sheet.range((helper_row, helper_base_col + 4))
        num_q_cell = sheet.range((helper_row, helper_base_col + 5))

        set_formula2(
            intercept_cell,
            (
                f'=IFERROR(INTERCEPT('
                f'R{data_start}C{y_col}:R{data_end}C{y_col},'
                f'R{data_start}C{x_col}:R{data_end}C{x_col}'
                f'),"")'
            ),
        )
        set_formula2(
            slope_cell,
            (
                f'=IFERROR(SLOPE('
                f'R{data_start}C{y_col}:R{data_end}C{y_col},'
                f'R{data_start}C{x_col}:R{data_end}C{x_col}'
                f'),"")'
            ),
        )
        set_formula2(
            forecast_cell,
            (
                f'=IFERROR('
                f'RC[-2]+RC[-1]*(MAX(R{data_start}C{x_col}:R{data_end}C{x_col})+1),'
                f'"")'
            ),
        )
        set_formula2(max_cell, f'=IFERROR(MAX(R{data_start}C{y_col}:R{data_end}C{y_col}),"")')
        set_formula2(min_cell, f'=IFERROR(MIN(R{data_start}C{y_col}:R{data_end}C{y_col}),"")')
        num_q_cell.value = window

        formulas_written = True
        helper_rows.append((helper_row, window))

    if formulas_written:
        wb.app.calculate()

    regression_rows: List[Dict[str, Any]] = []
    for helper_row, window in helper_rows:
        intercept = sheet.range((helper_row, helper_base_col)).value
        slope = sheet.range((helper_row, helper_base_col + 1)).value
        forecast_value = sheet.range((helper_row, helper_base_col + 2)).value
        forecast_max = sheet.range((helper_row, helper_base_col + 3)).value
        forecast_min = sheet.range((helper_row, helper_base_col + 4)).value

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": window,
            "num_quarters_used": window,
            "forecast_value": clean_numeric(forecast_value),
            "actual_value": None,
            "forecast_max": clean_numeric(forecast_max),
            "forecast_min": clean_numeric(forecast_min),
            "range_width": clean_numeric(safe_diff(forecast_max, forecast_min)),
            "intercept": clean_numeric(intercept),
            "slope": clean_numeric(slope),
            "source_file": source_file,
        }
        regression_rows.append(row)

    if len(regression_rows) >= 2:
        final_row = regression_rows[-1]
        prev_row = regression_rows[-2]
        duplicate = (
            almost_equal(final_row.get("forecast_value"), prev_row.get("forecast_value"))
            and almost_equal(final_row.get("forecast_max"), prev_row.get("forecast_max"))
            and almost_equal(final_row.get("forecast_min"), prev_row.get("forecast_min"))
            and almost_equal(final_row.get("intercept"), prev_row.get("intercept"))
            and almost_equal(final_row.get("slope"), prev_row.get("slope"))
        )
        if duplicate:
            regression_rows.pop()

    return regression_rows


def write_sheet(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = row.get(header)
            if value is None:
                continue
            text = str(value)
            max_len = max(max_len, len(text))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path, empirical_rows: Sequence[Dict[str, Any]], regression_rows: Sequence[Dict[str, Any]]
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_empirical = wb.create_sheet("empirical_candidates")
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_HEADERS, regression_rows)

    wb.save(output_path)


def iter_source_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            print(f"SKIP: {path.name} (not a file)")
            continue
        if path.name.startswith("~"):
            print(f"SKIP: {path.name} (temp file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"SKIP: {path.name} (not .xlsx)")
            continue
        stem = path.stem
        if stem.endswith("_PARAM") or re.search(r"_PARAM\.\d+$", stem):
            print(f"SKIP: {path.name} (existing parameter output)")
            continue
        yield path


def process_all_files() -> int:
    if xw is None:
        print("ERROR: xlwings is not installed. Install with: pip install xlwings")
        return 1

    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists() or not in_dir.is_dir():
        print(f"ERROR: input_dir does not exist or is not a directory: {in_dir}")
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_source_files(in_dir):
            print(f"PROCESS: {file_path.name}")
            meta = parse_model_meta(file_path.name)
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_candidates(wb, meta, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"SKIP: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(process_all_files())
