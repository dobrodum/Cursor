#!/usr/bin/env python3
"""
Extract empirical and regression parameter candidates from Excel model workbooks.

This script:
1) opens each source workbook once (hidden Excel app),
2) processes both "Empirical Model" and "Regression Model" while open,
3) closes source workbook without saving, and
4) writes one consolidated output workbook with two sheets.
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

# ----------------------------
# User-configurable directories
# ----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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

MONTH_LOOKUP = {
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if is_number(value):
        return float(value)
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def subtract_numbers(a: Any, b: Any) -> Optional[float]:
    af = to_float(a)
    bf = to_float(b)
    if af is None or bf is None:
        return None
    return af - bf


def parse_filename_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].strip() if len(parts) >= 2 and parts[1].strip() else "UNKNOWN"

    period_token = parts[2].strip() if len(parts) >= 3 else ""
    # Example: MidJan2026_Send -> MidJan2026
    period_token = period_token.split("_")[0].strip()

    match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", period_token, re.IGNORECASE)

    model_period = "UNKNOWN_PERIOD"
    model_date = ""

    if match:
        period_part = match.group(1).lower()
        month_token = re.sub(r"[^A-Za-z]", "", match.group(2)).lower()
        year = int(match.group(3))
        day = DAY_BY_PERIOD.get(period_part, 15)
        month_num = MONTH_LOOKUP.get(month_token)

        if month_num:
            month_short = date(year, month_num, 1).strftime("%b")
            model_period = f"{period_part.title()}{month_short}_{year}"
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    first_candidate = output_folder / f"{base_name}.xlsx"
    if not first_candidate.exists():
        return first_candidate

    i = 1
    while True:
        candidate = output_folder / f"{base_name}.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def should_process_file(file_path: Path, input_folder_name: str) -> Tuple[bool, str]:
    if not file_path.is_file():
        return False, "not a file"
    if file_path.suffix.lower() != ".xlsx":
        return False, "not an .xlsx file"
    if file_path.name.startswith("~"):
        return False, "temporary lock file"
    if re.match(rf"^{re.escape(input_folder_name)}_PARAM(\.\d+)?\.xlsx$", file_path.name, re.IGNORECASE):
        return False, "looks like a prior output workbook"
    return True, ""


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor_max(sheet: xw.Sheet) -> Optional[Tuple[int, int, int, int]]:
    used = sheet.used_range
    values_2d = as_2d(used.value)
    if not values_2d:
        return None

    start_row = used.row
    start_col = used.column
    end_row = start_row + len(values_2d) - 1
    end_col = start_col + len(values_2d[0]) - 1

    for r_idx, row in enumerate(values_2d):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return start_row + r_idx, start_col + c_idx, end_row, end_col
    return None


def find_column_by_keywords(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    keyword_options: Sequence[str],
    fallback_offset: int,
    search_radius: int = 35,
) -> int:
    left = max(1, anchor_col - search_radius)
    right = anchor_col + search_radius

    candidate_rows: List[int] = []
    for row in (anchor_row, anchor_row - 1, anchor_row + 1):
        if row > 0:
            candidate_rows.append(row)

    best_col = None
    best_distance = None

    for row in candidate_rows:
        row_values = sheet.range((row, left), (row, right)).value
        if not isinstance(row_values, list):
            row_values = [row_values]

        for idx, cell_val in enumerate(row_values):
            cell_text = normalize_text(cell_val)
            if not cell_text:
                continue
            if any(keyword in cell_text for keyword in keyword_options):
                col = left + idx
                distance = abs(col - anchor_col)
                if best_col is None or distance < (best_distance or 10**9):
                    best_col = col
                    best_distance = distance

    if best_col is not None:
        return best_col
    return max(1, anchor_col + fallback_offset)


def column_values(sheet: xw.Sheet, col: int, start_row: int, end_row: int) -> List[Any]:
    if end_row < start_row:
        return []
    vals = sheet.range((start_row, col), (end_row, col)).value
    if not isinstance(vals, list):
        return [vals]
    return vals


def numeric_rows_for_column(sheet: xw.Sheet, col: int, end_row: int) -> List[int]:
    rows: List[int] = []
    for row_idx, value in enumerate(column_values(sheet, col, 1, end_row), start=1):
        if is_number(value):
            rows.append(row_idx)
    return rows


def set_r1c1_formula(cell: xw.Range, formula: str) -> None:
    # Prefer Formula2R1C1 for dynamic-array aware Excel versions.
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass
    try:
        cell.api.FormulaR1C1 = formula
        return
    except Exception:
        pass
    cell.formula = formula


def safe_close_workbook(wb: xw.Book) -> None:
    # Never save source workbooks.
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
        wb.api.Close(False)
    except Exception as exc:
        print(f"warning: unable to close workbook cleanly: {exc}")


def process_empirical_sheet(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"  skipped empirical: sheet 'Empirical Model' not found in {source_file}")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"  skipped empirical: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col, _, end_col = anchor
    start_row = anchor_row + 1

    min_col = find_column_by_keywords(sheet, anchor_row, anchor_col, ["min"], fallback_offset=1)
    forecast_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["estimated total sold", "tot fcst", "forecast"],
        fallback_offset=-1,
    )
    actual_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["actual", "reported sales", "reported"],
        fallback_offset=2,
    )
    reported_sales_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["reported sales", "reported"],
        fallback_offset=2,
    )
    quarterly_sales_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["quarterly sales", "qtr sales", "quarter sales"],
        fallback_offset=-4,
    )
    growth_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["growth", "growth rate"],
        fallback_offset=-2,
    )
    captured_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["captured", "captured in db", "db pct"],
        fallback_offset=-3,
    )
    last_quarter_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["last quarter", "quarter used"],
        fallback_offset=-8,
    )
    num_quarters_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["num quarters", "quarters used"],
        fallback_offset=-10,
    )
    avg_pen_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["avg penetration", "average penetration", "penetration pct"],
        fallback_offset=-5,
    )
    penetration_source_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["penetration"],
        fallback_offset=avg_pen_col - anchor_col,
    )

    penetration_rows = numeric_rows_for_column(sheet, penetration_source_col, anchor_row - 1)
    helper_col = max(end_col + 2, anchor_col + 2)
    formula_rows = min(N_QUARTERS, len(penetration_rows))
    wrote_formulas = False

    if formula_rows > 0:
        for i in range(1, formula_rows + 1):
            target_row = start_row + i - 1
            src_start = penetration_rows[-i]
            src_end = penetration_rows[-1]
            avg_formula = f"=AVERAGE(R{src_start}C{penetration_source_col}:R{src_end}C{penetration_source_col})"
            set_r1c1_formula(sheet.range((target_row, helper_col)), avg_formula)
        wb.app.calculate()
        wrote_formulas = True

    rows: List[Dict[str, Any]] = []
    for i in range(1, N_QUARTERS + 1):
        row_num = start_row + i - 1
        num_quarters_raw = sheet.range((row_num, num_quarters_col)).value
        num_quarters_used = int(num_quarters_raw) if is_number(num_quarters_raw) else i

        avg_penetration = (
            sheet.range((row_num, helper_col)).value
            if wrote_formulas and i <= formula_rows
            else sheet.range((row_num, avg_pen_col)).value
        )

        forecast_max = sheet.range((row_num, anchor_col)).value
        forecast_min = sheet.range((row_num, min_col)).value
        forecast_value = sheet.range((row_num, forecast_col)).value
        actual_value = sheet.range((row_num, actual_col)).value
        quarterly_sales = sheet.range((row_num, quarterly_sales_col)).value
        reported_sales = sheet.range((row_num, reported_sales_col)).value
        growth_rate_pct = sheet.range((row_num, growth_col)).value
        captured_pct = sheet.range((row_num, captured_col)).value
        last_quarter_used = sheet.range((row_num, last_quarter_col)).value

        if all(v in (None, "") for v in (forecast_max, forecast_min, forecast_value, avg_penetration, actual_value)):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_numbers(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    if wrote_formulas:
        sheet.range((start_row, helper_col), (start_row + formula_rows - 1, helper_col)).clear_contents()

    return rows


def signature_value(value: Any) -> Any:
    if is_number(value):
        return round(float(value), 8)
    return value if value not in ("",) else None


def process_regression_sheet(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"  skipped regression: sheet 'Regression Model' not found in {source_file}")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"  skipped regression: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col, _, end_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    if y_col < 1 or x_col < 1:
        print(f"  skipped regression: anchor offsets invalid in {source_file}")
        return []

    x_vals = column_values(sheet, x_col, 1, anchor_row - 1)
    y_vals = column_values(sheet, y_col, 1, anchor_row - 1)
    numeric_xy_rows = [
        row_idx
        for row_idx, (x_val, y_val) in enumerate(zip(x_vals, y_vals), start=1)
        if is_number(x_val) and is_number(y_val)
    ]

    if len(numeric_xy_rows) < 2:
        print(f"  skipped regression: not enough numeric x/y history in {source_file}")
        return []

    min_col = find_column_by_keywords(sheet, anchor_row, anchor_col, ["min"], fallback_offset=1)
    forecast_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["tot fcst w/o sa", "tot fcst", "forecast"],
        fallback_offset=-1,
    )
    actual_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["actual", "reported sales", "reported"],
        fallback_offset=2,
    )
    num_quarters_col = find_column_by_keywords(
        sheet,
        anchor_row,
        anchor_col,
        ["num quarters", "quarters used"],
        fallback_offset=-10,
    )

    max_rows = min(N_QUARTERS, len(numeric_xy_rows))
    start_row = anchor_row + 1

    helper_intercept_col = max(end_col + 2, anchor_col + 2)
    helper_slope_col = helper_intercept_col + 1

    for i in range(1, max_rows + 1):
        target_row = start_row + i - 1
        src_start = numeric_xy_rows[-i]
        src_end = numeric_xy_rows[-1]
        intercept_formula = f"=INTERCEPT(R{src_start}C{y_col}:R{src_end}C{y_col},R{src_start}C{x_col}:R{src_end}C{x_col})"
        slope_formula = f"=SLOPE(R{src_start}C{y_col}:R{src_end}C{y_col},R{src_start}C{x_col}:R{src_end}C{x_col})"
        set_r1c1_formula(sheet.range((target_row, helper_intercept_col)), intercept_formula)
        set_r1c1_formula(sheet.range((target_row, helper_slope_col)), slope_formula)

    wb.app.calculate()

    last_x = to_float(sheet.range((numeric_xy_rows[-1], x_col)).value)
    next_x = to_float(sheet.range((numeric_xy_rows[-1] + 1, x_col)).value)
    if next_x is None and last_x is not None:
        next_x = last_x + 1.0

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(1, max_rows + 1):
        row_num = start_row + i - 1

        num_quarters_raw = sheet.range((row_num, num_quarters_col)).value
        num_quarters_used = int(num_quarters_raw) if is_number(num_quarters_raw) else i
        intercept = sheet.range((row_num, helper_intercept_col)).value
        slope = sheet.range((row_num, helper_slope_col)).value
        forecast_value = sheet.range((row_num, forecast_col)).value
        forecast_max = sheet.range((row_num, anchor_col)).value
        forecast_min = sheet.range((row_num, min_col)).value
        actual_value = sheet.range((row_num, actual_col)).value

        if forecast_value in (None, "") and is_number(intercept) and is_number(slope) and next_x is not None:
            forecast_value = float(intercept) + float(slope) * next_x

        if all(v in (None, "") for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        current_signature = (
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if current_signature == previous_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
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
                "range_width": subtract_numbers(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    sheet.range((start_row, helper_intercept_col), (start_row + max_rows - 1, helper_slope_col)).clear_contents()
    return rows


def autosize_columns(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            val = row.get(header)
            if val is None:
                continue
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_sheet(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(header) for header in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"
    autosize_columns(ws, headers, rows)


def write_output_workbook(
    destination: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(destination)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        yield path


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a directory: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file_path = next_output_path(input_path, output_path)
    input_folder_name = input_path.name

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    original_calc_mode = None
    try:
        try:
            original_calc_mode = app.calculation
            app.calculation = "manual"
        except Exception:
            original_calc_mode = None

        for file_path in iter_input_files(input_path):
            should_process, reason = should_process_file(file_path, input_folder_name)
            if not should_process:
                print(f"Skipped: {file_path.name} ({reason})")
                continue

            print(f"Processed: {file_path.name}")
            wb = None
            try:
                metadata = parse_filename_metadata(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if original_calc_mode is not None:
            try:
                app.calculation = original_calc_mode
            except Exception:
                pass
        app.quit()

    write_output_workbook(output_file_path, empirical_rows, regression_rows)

    print(f"Output path: {output_file_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
