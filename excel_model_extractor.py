#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

SOURCE_EMPIRICAL_SHEET = "Empirical Model"
SOURCE_REGRESSION_SHEET = "Regression Model"
OUTPUT_EMPIRICAL_SHEET = "empirical_candidates"
OUTPUT_REGRESSION_SHEET = "regression_candidates"
N_QUARTERS = 10
MAX_EXCEL_COL = 16384

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[\s_/\-\\]+", " ", text)


def ensure_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], tuple):
            values = [list(row) for row in values]
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def ensure_column(values: Any, expected_len: int) -> List[Any]:
    if isinstance(values, tuple):
        values = list(values)
    if isinstance(values, list):
        if values and isinstance(values[0], (list, tuple)):
            flattened = [row[0] if row else None for row in values]
        else:
            flattened = list(values)
    else:
        flattened = [values]
    if len(flattened) < expected_len:
        flattened.extend([None] * (expected_len - len(flattened)))
    return flattened[:expected_len]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    is_pct = text.endswith("%")
    if is_pct:
        text = text[:-1]
    try:
        out = float(text)
    except ValueError:
        return None
    if is_pct:
        out /= 100.0
    return out


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def text_or_blank(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def safe_subtract(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return a - b


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        # Final fallback for older Excel APIs.
        cell.formula = formula_r1c1


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


def find_anchor_cell(
    values: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    anchor_text: str = "max",
) -> Optional[Tuple[int, int, int]]:
    target = normalize_text(anchor_text)
    for row_idx, row in enumerate(values):
        for col_idx, cell_value in enumerate(row):
            if normalize_text(cell_value) == target:
                return start_row + row_idx, start_col + col_idx, row_idx
    return None


def build_header_lookup(row_values: Sequence[Any], start_col: int) -> Dict[str, int]:
    lookup: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        key = normalize_text(value)
        if key and key not in lookup:
            lookup[key] = start_col + idx
    return lookup


def find_column(
    header_lookup: Dict[str, int],
    candidates: Iterable[Sequence[str] | str],
    default: Optional[int] = None,
) -> Optional[int]:
    for candidate in candidates:
        if isinstance(candidate, str):
            needle = normalize_text(candidate)
            for header, col in header_lookup.items():
                if needle and needle in header:
                    return col
        else:
            words = [normalize_text(word) for word in candidate if normalize_text(word)]
            for header, col in header_lookup.items():
                if words and all(word in header for word in words):
                    return col
    return default


def read_column_values(sheet: xw.Sheet, start_row: int, end_row: int, col: Optional[int]) -> List[Any]:
    row_count = end_row - start_row + 1
    if col is None or col < 1:
        return [None] * row_count
    values = sheet.range((start_row, col), (end_row, col)).value
    return ensure_column(values, row_count)


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    left_col: int,
    right_col: int,
) -> List[List[Any]]:
    expected_rows = end_row - start_row + 1
    expected_cols = right_col - left_col + 1
    values = ensure_matrix(sheet.range((start_row, left_col), (end_row, right_col)).value)
    if len(values) < expected_rows:
        values.extend([[None] * expected_cols for _ in range(expected_rows - len(values))])
    fixed_rows: List[List[Any]] = []
    for row in values[:expected_rows]:
        row_list = list(row)
        if len(row_list) < expected_cols:
            row_list.extend([None] * (expected_cols - len(row_list)))
        fixed_rows.append(row_list[:expected_cols])
    return fixed_rows


def parse_model_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_raw = parts[2] if len(parts) >= 3 else ""
    period_raw = re.sub(r"[_\-]?send.*$", "", period_raw, flags=re.IGNORECASE).strip()

    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", period_raw, flags=re.IGNORECASE)
    model_period = ""
    model_date = ""
    if period_match:
        phase = period_match.group(1).title()
        month_text = period_match.group(2)
        year = int(period_match.group(3))
        month_abbrev = month_text[:3].title()
        month_num = MONTH_MAP.get(month_abbrev.lower())
        model_period = f"{phase}{month_abbrev}_{year}"
        if month_num:
            model_date = date(year, month_num, PERIOD_DAY_MAP[phase.lower()]).isoformat()
    else:
        fallback = re.sub(r"\s+", "", period_raw)
        model_period = fallback if fallback else "unknown_period"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def choose_temp_cols(used_end_col: int, anchor_col: int) -> Tuple[int, int]:
    base = max(used_end_col + 2, anchor_col + 5)
    base = min(max(base, 1), MAX_EXCEL_COL - 1)
    return base, base + 1


def extract_empirical_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[SOURCE_EMPIRICAL_SHEET]
    except Exception:
        print(f"  Skipped empirical extraction for {source_file}: sheet '{SOURCE_EMPIRICAL_SHEET}' not found")
        return []

    used = sheet.used_range
    values = ensure_matrix(used.value)
    if not values:
        print(f"  Skipped empirical extraction for {source_file}: sheet is empty")
        return []

    used_start_row = used.row
    used_start_col = used.column
    used_end_col = used_start_col + max(len(row) for row in values) - 1

    anchor = find_anchor_cell(values, used_start_row, used_start_col, "max")
    if anchor is None:
        print(f"  Skipped empirical extraction for {source_file}: max anchor not found")
        return []
    anchor_row, anchor_col, anchor_row_idx = anchor

    header_lookup = build_header_lookup(values[anchor_row_idx], used_start_col)

    max_col = anchor_col
    min_col = find_column(header_lookup, ["min"], default=anchor_col + 1)
    forecast_col = find_column(
        header_lookup,
        [
            "estimated total sold",
            "forecast value",
            ("total", "forecast"),
            ("tot", "fcst"),
        ],
        default=anchor_col - 1,
    )
    reported_sales_col = find_column(
        header_lookup,
        [
            "reported sales",
            ("actual", "sales"),
            "actual value",
            "actual",
        ],
        default=anchor_col - 2,
    )
    num_quarters_col = find_column(
        header_lookup,
        [
            ("num", "quarters", "used"),
            ("quarters", "used"),
            ("num", "quarters"),
        ],
        default=anchor_col - 12,
    )
    last_quarter_col = find_column(
        header_lookup,
        [
            ("last", "quarter", "used"),
            ("last", "quarter"),
        ],
        default=anchor_col - 13,
    )
    quarterly_sales_col = find_column(
        header_lookup,
        ["quarterly sales", ("quarterly", "sales"), ("quarter", "sales")],
        default=anchor_col - 5,
    )
    growth_rate_col = find_column(
        header_lookup,
        ["growth rate", ("growth", "rate")],
        default=anchor_col - 4,
    )
    captured_pct_col = find_column(
        header_lookup,
        [
            ("sales", "captured", "db"),
            ("captured", "db"),
            ("captured", "pct"),
        ],
        default=anchor_col - 3,
    )

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1

    num_quarters_values = read_column_values(sheet, first_data_row, last_data_row, num_quarters_col)

    temp_avg_col, _ = choose_temp_cols(used_end_col, anchor_col)
    penetration_last_col = anchor_col - 2
    penetration_first_col = max(1, anchor_col - 11)

    quarter_counts: List[int] = []
    formulas_written = False
    for idx in range(N_QUARTERS):
        row_num = first_data_row + idx
        row_quarters = to_int(num_quarters_values[idx])
        if row_quarters is None:
            row_quarters = idx + 1
        row_quarters = max(1, min(N_QUARTERS, row_quarters))
        quarter_counts.append(row_quarters)

        avg_start_col = max(penetration_first_col, penetration_last_col - row_quarters + 1)
        if avg_start_col > penetration_last_col:
            continue

        start_offset = avg_start_col - temp_avg_col
        end_offset = penetration_last_col - temp_avg_col
        formula = f"=AVERAGE(RC[{start_offset}]:RC[{end_offset}])"
        set_formula2_r1c1(sheet.range((row_num, temp_avg_col)), formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()
    avg_pen_values = read_column_values(sheet, first_data_row, last_data_row, temp_avg_col)
    if formulas_written:
        sheet.range((first_data_row, temp_avg_col), (last_data_row, temp_avg_col)).clear_contents()

    columns = [max_col, min_col, forecast_col, reported_sales_col, num_quarters_col]
    if last_quarter_col is not None:
        columns.append(last_quarter_col)
    if quarterly_sales_col is not None:
        columns.append(quarterly_sales_col)
    if growth_rate_col is not None:
        columns.append(growth_rate_col)
    if captured_pct_col is not None:
        columns.append(captured_pct_col)
    valid_cols = sorted({col for col in columns if col is not None and col >= 1})
    if not valid_cols:
        return []

    left_col = min(valid_cols)
    right_col = max(valid_cols)
    block = read_block(sheet, first_data_row, last_data_row, left_col, right_col)

    def block_value(row_idx: int, col: Optional[int]) -> Any:
        if col is None or col < left_col or col > right_col:
            return None
        return block[row_idx][col - left_col]

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters_used = to_int(block_value(idx, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = quarter_counts[idx]

        forecast_max = to_float(block_value(idx, max_col))
        forecast_min = to_float(block_value(idx, min_col))
        forecast_value = to_float(block_value(idx, forecast_col))
        reported_sales = to_float(block_value(idx, reported_sales_col))
        avg_penetration_pct = to_float(avg_pen_values[idx])

        if all(
            value is None
            for value in [forecast_max, forecast_min, forecast_value, reported_sales, avg_penetration_pct]
        ):
            continue

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": text_or_blank(block_value(idx, last_quarter_col)),
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": to_float(block_value(idx, quarterly_sales_col)),
            "reported_sales": reported_sales,
            "growth_rate_pct": to_float(block_value(idx, growth_rate_col)),
            "sales_captured_in_db_pct": to_float(block_value(idx, captured_pct_col)),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def nearly_equal(a: Any, b: Any, tol: float = 1e-9) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is not None and b_num is not None:
        return abs(a_num - b_num) <= tol
    return str(a) == str(b)


def dedupe_trailing_regression_row(rows: List[Dict[str, Any]]) -> None:
    if len(rows) < 2:
        return
    keys = ["num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    last = rows[-1]
    prev = rows[-2]
    if all(nearly_equal(last.get(key), prev.get(key)) for key in keys):
        rows.pop()


def extract_regression_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[SOURCE_REGRESSION_SHEET]
    except Exception:
        print(f"  Skipped regression extraction for {source_file}: sheet '{SOURCE_REGRESSION_SHEET}' not found")
        return []

    used = sheet.used_range
    values = ensure_matrix(used.value)
    if not values:
        print(f"  Skipped regression extraction for {source_file}: sheet is empty")
        return []

    used_start_row = used.row
    used_start_col = used.column
    used_end_col = used_start_col + max(len(row) for row in values) - 1

    anchor = find_anchor_cell(values, used_start_row, used_start_col, "max")
    if anchor is None:
        print(f"  Skipped regression extraction for {source_file}: max anchor not found")
        return []
    anchor_row, anchor_col, anchor_row_idx = anchor

    header_lookup = build_header_lookup(values[anchor_row_idx], used_start_col)

    max_col = anchor_col
    min_col = find_column(header_lookup, ["min"], default=anchor_col + 1)
    forecast_col = find_column(
        header_lookup,
        [
            "tot fcst w/o sa",
            "tot fcst without sa",
            "forecast total without sa",
            ("total", "forecast", "without", "sa"),
            ("total", "forecast"),
        ],
        default=anchor_col - 1,
    )
    num_quarters_col = find_column(
        header_lookup,
        [
            ("num", "quarters", "used"),
            ("quarters", "used"),
            ("num", "quarters"),
        ],
        default=anchor_col - 12,
    )
    actual_col = find_column(
        header_lookup,
        ["actual value", "actual", ("reported", "sales")],
        default=None,
    )
    intercept_existing_col = find_column(header_lookup, ["intercept"], default=None)
    slope_existing_col = find_column(header_lookup, ["slope"], default=None)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1
    num_quarters_values = read_column_values(sheet, first_data_row, last_data_row, num_quarters_col)

    temp_intercept_col, temp_slope_col = choose_temp_cols(used_end_col, anchor_col)
    formulas_written = False
    quarter_counts: List[int] = []

    for idx in range(N_QUARTERS):
        row_num = first_data_row + idx
        row_quarters = to_int(num_quarters_values[idx])
        if row_quarters is None:
            row_quarters = idx + 1
        row_quarters = max(2, min(N_QUARTERS, row_quarters))
        quarter_counts.append(row_quarters)

        if x_col < 1 or y_col < 1:
            continue

        series_start = max(first_data_row, row_num - row_quarters + 1)
        if row_num - series_start + 1 < 2:
            continue

        intercept_formula = (
            f"=INTERCEPT(R{series_start}C{y_col}:R{row_num}C{y_col},"
            f"R{series_start}C{x_col}:R{row_num}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{series_start}C{y_col}:R{row_num}C{y_col},"
            f"R{series_start}C{x_col}:R{row_num}C{x_col})"
        )
        set_formula2_r1c1(sheet.range((row_num, temp_intercept_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((row_num, temp_slope_col)), slope_formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()
    intercept_values = read_column_values(sheet, first_data_row, last_data_row, temp_intercept_col)
    slope_values = read_column_values(sheet, first_data_row, last_data_row, temp_slope_col)
    if formulas_written:
        sheet.range((first_data_row, temp_intercept_col), (last_data_row, temp_slope_col)).clear_contents()

    columns = [max_col, min_col, forecast_col, num_quarters_col]
    if actual_col is not None:
        columns.append(actual_col)
    if intercept_existing_col is not None:
        columns.append(intercept_existing_col)
    if slope_existing_col is not None:
        columns.append(slope_existing_col)
    valid_cols = sorted({col for col in columns if col is not None and col >= 1})
    if not valid_cols:
        return []

    left_col = min(valid_cols)
    right_col = max(valid_cols)
    block = read_block(sheet, first_data_row, last_data_row, left_col, right_col)

    def block_value(row_idx: int, col: Optional[int]) -> Any:
        if col is None or col < left_col or col > right_col:
            return None
        return block[row_idx][col - left_col]

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters_used = to_int(block_value(idx, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = quarter_counts[idx]

        forecast_max = to_float(block_value(idx, max_col))
        forecast_min = to_float(block_value(idx, min_col))
        forecast_value = to_float(block_value(idx, forecast_col))
        actual_value = to_float(block_value(idx, actual_col))

        intercept = to_float(intercept_values[idx])
        slope = to_float(slope_values[idx])
        if intercept is None:
            intercept = to_float(block_value(idx, intercept_existing_col))
        if slope is None:
            slope = to_float(block_value(idx, slope_existing_col))

        if all(
            value is None
            for value in [forecast_max, forecast_min, forecast_value, intercept, slope]
        ):
            continue

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
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    dedupe_trailing_regression_row(rows)
    return rows


def output_workbook_path(in_dir: Path, out_dir: Path) -> Path:
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


def value_length(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, float):
        text = f"{value:.6g}"
    else:
        text = str(value)
    return len(text)


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            max_len = max(max_len, value_length(ws.cell(row=row_idx, column=col_idx).value))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet(OUTPUT_EMPIRICAL_SHEET)
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet(OUTPUT_REGRESSION_SHEET)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_workbook_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(input_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processed file: {file_path.name}")
            metadata = parse_model_metadata(file_path.name)

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_count += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
