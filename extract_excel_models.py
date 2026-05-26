#!/usr/bin/env python3
"""
Extract candidate rows from Excel model workbooks.

Workflow summary:
- Process every `.xlsx` file from `input_dir` (excluding temp files like `~$...`).
- Open each workbook exactly once with xlwings, process both source sheets while open:
  - Empirical Model
  - Regression Model
- Never save source workbooks.
- Write one combined output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------------------------
# User-configurable paths
# -------------------------
input_dir = "./input"
output_dir = "./output"

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

WINDOW_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def excel_error_to_none(value: Any) -> Any:
    if isinstance(value, str) and value.strip().startswith("#"):
        return None
    return value


def to_float(value: Any) -> Optional[float]:
    value = excel_error_to_none(value)
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> Optional[int]:
    parsed = to_float(value)
    if parsed is None:
        return None
    return int(round(parsed))


def subtract(a: Any, b: Any) -> Optional[float]:
    fa = to_float(a)
    fb = to_float(b)
    if fa is None or fb is None:
        return None
    return fa - fb


def find_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.name
    base_name = f"{folder_name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[-2] if len(parts) >= 2 else ""
    period_token = parts[-1] if parts else stem
    period_token = period_token.split("_")[0]

    # Expected token format: EarlyJan2026 / MidJan2026 / LateJan2026
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$", period_token)
    if not match:
        # fallback when naming is imperfect
        model_period = period_token
        model_date = ""
    else:
        window_raw, month_raw, year_raw = match.groups()
        month_num = MONTH_MAP.get(month_raw.lower())
        day = WINDOW_DAY_MAP.get(window_raw.lower())
        year = int(year_raw)

        if month_num is None or day is None:
            model_period = period_token
            model_date = ""
        else:
            model_period = f"{window_raw}{month_raw.title()}_{year}"
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or stem)
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


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

    try:
        wb.api.Close(SaveChanges=False)
    except Exception as exc:
        print(f"WARN: workbook close fallback failed: {exc}")


def get_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def find_anchor(sheet: xw.Sheet, label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    target = label.strip().lower()
    start_row = used.row
    start_col = used.column

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return (start_row + r_idx, start_col + c_idx)
    return None


def header_map_for_row(sheet: xw.Sheet, row_idx: int) -> Dict[int, str]:
    last_col = sheet.used_range.last_cell.column
    values = sheet.range((row_idx, 1), (row_idx, last_col)).value
    if not isinstance(values, list):
        values = [values]
    return {col_idx: normalize_text(value) for col_idx, value in enumerate(values, start=1)}


def first_matching_col(
    headers: Dict[int, str], phrases: Iterable[str], default_col: int
) -> int:
    normalized_phrases = [p.lower() for p in phrases]
    for col_idx, text in headers.items():
        if not text:
            continue
        for phrase in normalized_phrases:
            if phrase in text:
                return col_idx
    return default_col


def get_block_value(
    block: List[List[Any]], row_idx: int, col_idx: int, start_row: int, start_col: int
) -> Any:
    r = row_idx - start_row
    c = col_idx - start_col
    if r < 0 or c < 0:
        return None
    if r >= len(block):
        return None
    row = block[r]
    if c >= len(row):
        return None
    return row[c]


def flatten_vertical(values: Any) -> List[Any]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [values]
    if not values:
        return []
    if isinstance(values[0], list):
        return [row[0] if row else None for row in values]
    return values


def extract_empirical_rows(
    wb: xw.Book, meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = header_map_for_row(sheet, anchor_row)

    # Anchor-based columns (header match first, offset fallback second).
    num_quarters_col = first_matching_col(headers, ["num quarters", "quarters used"], anchor_col - 12)
    last_quarter_col = first_matching_col(headers, ["last quarter", "last qtr"], anchor_col - 11)
    quarterly_sales_col = first_matching_col(headers, ["quarterly sales", "quarter sales"], anchor_col - 10)
    growth_rate_col = first_matching_col(headers, ["growth rate", "growth %"], anchor_col - 9)
    sales_captured_col = first_matching_col(
        headers,
        ["sales captured in db", "captured in db", "sales captured", "penetration"],
        anchor_col - 8,
    )
    forecast_col = first_matching_col(
        headers,
        ["estimated total sold", "forecast value", "forecast", "tot fcst"],
        anchor_col - 2,
    )
    reported_sales_col = first_matching_col(
        headers,
        ["reported sales", "actual sales", "actual value", "actual"],
        anchor_col - 1,
    )
    max_col = anchor_col
    min_col = first_matching_col(headers, ["min"], anchor_col + 1)

    n_quarters = 10
    start_row = anchor_row + 1
    end_row = start_row + n_quarters - 1

    data_cols = sorted(
        {
            num_quarters_col,
            last_quarter_col,
            quarterly_sales_col,
            growth_rate_col,
            sales_captured_col,
            forecast_col,
            reported_sales_col,
            max_col,
            min_col,
        }
    )
    data_cols = [c for c in data_cols if c > 0]
    if not data_cols:
        return []

    block_start_col = min(data_cols)
    block_end_col = max(data_cols)
    block_values = to_2d(sheet.range((start_row, block_start_col), (end_row, block_end_col)).value)

    # Write average penetration formulas in temporary helper cells (R1C1 + formula2).
    scratch_col = sheet.used_range.last_cell.column + 10
    for idx in range(n_quarters):
        row_idx = start_row + idx
        lookback = idx + 1
        src_offset = sales_captured_col - scratch_col
        if lookback == 1:
            formula = f"=RC[{src_offset}]"
        else:
            formula = f"=AVERAGE(R[-{lookback - 1}]C[{src_offset}]:RC[{src_offset}])"
        sheet.range((row_idx, scratch_col)).formula2 = formula

    wb.app.calculate()
    avg_pen_values = flatten_vertical(sheet.range((start_row, scratch_col), (end_row, scratch_col)).value)

    rows: List[Dict[str, Any]] = []
    for idx in range(n_quarters):
        row_idx = start_row + idx

        forecast_value = get_block_value(block_values, row_idx, forecast_col, start_row, block_start_col)
        reported_sales = get_block_value(
            block_values, row_idx, reported_sales_col, start_row, block_start_col
        )
        forecast_max = get_block_value(block_values, row_idx, max_col, start_row, block_start_col)
        forecast_min = get_block_value(block_values, row_idx, min_col, start_row, block_start_col)
        sales_captured = get_block_value(
            block_values, row_idx, sales_captured_col, start_row, block_start_col
        )
        quarterly_sales = get_block_value(
            block_values, row_idx, quarterly_sales_col, start_row, block_start_col
        )

        if all(value in (None, "") for value in [forecast_value, reported_sales, forecast_max, forecast_min]):
            continue

        num_quarters_used = get_block_value(
            block_values, row_idx, num_quarters_col, start_row, block_start_col
        )
        last_quarter_used = get_block_value(
            block_values, row_idx, last_quarter_col, start_row, block_start_col
        )
        avg_penetration = avg_pen_values[idx] if idx < len(avg_pen_values) else None
        growth_rate = get_block_value(block_values, row_idx, growth_rate_col, start_row, block_start_col)

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": excel_error_to_none(avg_penetration),
            "num_quarters_used": to_int(num_quarters_used) or (idx + 1),
            "last_quarter_used": excel_error_to_none(last_quarter_used),
            "forecast_value": excel_error_to_none(forecast_value),
            "actual_value": excel_error_to_none(reported_sales),
            "forecast_max": excel_error_to_none(forecast_max),
            "forecast_min": excel_error_to_none(forecast_min),
            "range_width": subtract(forecast_max, forecast_min),
            "avg_penetration_pct": excel_error_to_none(avg_penetration),
            "quarterly_sales": excel_error_to_none(quarterly_sales),
            "reported_sales": excel_error_to_none(reported_sales),
            "growth_rate_pct": excel_error_to_none(growth_rate),
            "sales_captured_in_db_pct": excel_error_to_none(sales_captured),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def extract_regression_rows(
    wb: xw.Book, meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = header_map_for_row(sheet, anchor_row)

    # Required by specification.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = first_matching_col(headers, ["num quarters", "quarters used"], anchor_col - 12)
    forecast_col = first_matching_col(
        headers,
        ["tot fcst w/o sa", "total fcst w/o sa", "forecast total without sa", "tot fcst"],
        anchor_col - 1,
    )
    actual_col = first_matching_col(headers, ["actual", "reported"], -1)
    max_col = anchor_col
    min_col = first_matching_col(headers, ["min"], anchor_col + 1)

    n_quarters = 10
    start_row = anchor_row + 1
    end_row = start_row + n_quarters - 1

    data_cols = sorted({num_quarters_col, forecast_col, actual_col, max_col, min_col})
    data_cols = [c for c in data_cols if c > 0]
    if not data_cols:
        return []

    block_start_col = min(data_cols)
    block_end_col = max(data_cols)
    block_values = to_2d(sheet.range((start_row, block_start_col), (end_row, block_end_col)).value)

    # Temporary helper formulas for INTERCEPT and SLOPE using R1C1 with formula2.
    scratch_intercept_col = sheet.used_range.last_cell.column + 10
    scratch_slope_col = scratch_intercept_col + 1

    for idx in range(n_quarters):
        row_idx = start_row + idx
        lookback = idx + 1

        y_offset_i = y_col - scratch_intercept_col
        x_offset_i = x_col - scratch_intercept_col
        if lookback == 1:
            intercept_formula = f"=INTERCEPT(RC[{y_offset_i}],RC[{x_offset_i}])"
        else:
            intercept_formula = (
                f"=INTERCEPT(R[-{lookback - 1}]C[{y_offset_i}]:RC[{y_offset_i}],"
                f"R[-{lookback - 1}]C[{x_offset_i}]:RC[{x_offset_i}])"
            )
        sheet.range((row_idx, scratch_intercept_col)).formula2 = intercept_formula

        y_offset_s = y_col - scratch_slope_col
        x_offset_s = x_col - scratch_slope_col
        if lookback == 1:
            slope_formula = f"=SLOPE(RC[{y_offset_s}],RC[{x_offset_s}])"
        else:
            slope_formula = (
                f"=SLOPE(R[-{lookback - 1}]C[{y_offset_s}]:RC[{y_offset_s}],"
                f"R[-{lookback - 1}]C[{x_offset_s}]:RC[{x_offset_s}])"
            )
        sheet.range((row_idx, scratch_slope_col)).formula2 = slope_formula

    wb.app.calculate()
    intercept_values = flatten_vertical(
        sheet.range((start_row, scratch_intercept_col), (end_row, scratch_intercept_col)).value
    )
    slope_values = flatten_vertical(
        sheet.range((start_row, scratch_slope_col), (end_row, scratch_slope_col)).value
    )

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(n_quarters):
        row_idx = start_row + idx
        num_quarters_used = get_block_value(
            block_values, row_idx, num_quarters_col, start_row, block_start_col
        )
        forecast_value = get_block_value(block_values, row_idx, forecast_col, start_row, block_start_col)
        actual_value = (
            get_block_value(block_values, row_idx, actual_col, start_row, block_start_col)
            if actual_col > 0
            else ""
        )
        forecast_max = get_block_value(block_values, row_idx, max_col, start_row, block_start_col)
        forecast_min = get_block_value(block_values, row_idx, min_col, start_row, block_start_col)
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None

        if all(value in (None, "") for value in [forecast_value, forecast_max, forecast_min]):
            continue

        signature = (
            to_float(forecast_value),
            to_float(forecast_max),
            to_float(forecast_min),
            to_float(intercept),
            to_float(slope),
        )
        if idx == n_quarters - 1 and previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": to_int(num_quarters_used) or (idx + 1),
            "num_quarters_used": to_int(num_quarters_used) or (idx + 1),
            "forecast_value": excel_error_to_none(forecast_value),
            "actual_value": excel_error_to_none(actual_value),
            "forecast_max": excel_error_to_none(forecast_max),
            "forecast_min": excel_error_to_none(forecast_min),
            "range_width": subtract(forecast_max, forecast_min),
            "intercept": excel_error_to_none(intercept),
            "slope": excel_error_to_none(slope),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header, "") if row.get(header, "") is not None else "" for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    max_rows = ws.max_row
    for col_idx, header in enumerate(headers, start=1):
        letter = get_column_letter(col_idx)
        max_len = len(header)
        for row_idx in range(2, max_rows + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[letter].width = min(max(12, max_len + 2), 42)


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


def should_skip_file(path: Path, input_folder_name: str) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    if path.name.startswith("~"):
        return "temporary file"
    if re.match(rf"^{re.escape(input_folder_name)}_PARAM(\.\d+)?\.xlsx$", path.name):
        return "existing PARAM output"
    return None


def main() -> None:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()
    out_path.mkdir(parents=True, exist_ok=True)

    if not in_path.exists():
        raise FileNotFoundError(f"Input directory does not exist: {in_path}")

    output_file = find_output_path(in_path, out_path)
    input_folder_name = in_path.name

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_path.iterdir()):
            skip_reason = should_skip_file(file_path, input_folder_name)
            if skip_reason is not None:
                print(f"SKIP: {file_path.name} ({skip_reason})")
                continue

            print(f"PROCESS: {file_path.name}")
            meta = parse_file_metadata(file_path.name)
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"SKIP: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)

        write_output_workbook(output_file, empirical_rows, regression_rows)
    finally:
        app.quit()

    print(f"OUTPUT: {output_file}")
    print(f"FILES_PROCESSED: {processed_files}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
