#!/usr/bin/env python3
"""
Extract empirical/regression parameter candidates from Excel model workbooks.

This script processes all .xlsx files in input_dir and writes one combined
output workbook with:
  - empirical_candidates
  - regression_candidates

Each source workbook is opened once, both model sheets are processed while the
workbook is open, and then it is closed without saving.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import xlwings as xw


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = "./input"
output_dir = "./output"


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
N_QUARTERS = 10
EMPIRICAL_SOURCE_SHEET = "Empirical Model"
REGRESSION_SOURCE_SHEET = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"

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
DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileLabel:
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


def is_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    return isinstance(value, (int, float))


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if is_number(value):
        return float(value)
    try:
        text = str(value).strip().replace(",", "")
        if text == "":
            return None
        return float(text)
    except (TypeError, ValueError):
        return None


def to_int(value: Any, default: int) -> int:
    num = to_float(value)
    if num is None:
        return default
    return int(round(num))


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def list_input_files(in_dir: Path, input_folder_name: str) -> Tuple[List[Path], List[Tuple[Path, str]]]:
    candidates: List[Path] = []
    skipped: List[Tuple[Path, str]] = []
    output_prefix = f"{input_folder_name}_param"

    for file_path in sorted(in_dir.iterdir()):
        if not file_path.is_file():
            continue

        lower_name = file_path.name.lower()
        if lower_name.startswith("~"):
            skipped.append((file_path, "temp file"))
            continue
        if file_path.suffix.lower() != ".xlsx":
            skipped.append((file_path, "not .xlsx"))
            continue
        if lower_name.startswith(output_prefix):
            skipped.append((file_path, "looks like a prior output workbook"))
            continue
        candidates.append(file_path)

    return candidates, skipped


def build_unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    input_folder_name = in_dir.name
    base = out_dir / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    index = 1
    while True:
        candidate = out_dir / f"{input_folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def parse_file_label(file_path: Path) -> FileLabel:
    """
    Parse filenames like:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx

    Required outputs:
      ticker = AORT
      model_period = MidJan_2026
      model_date = 2026-01-15
      model = AORT_MidJan_2026
    """
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = "UNKNOWN"
    period_token = "MidJan2000"

    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"\s+", "", parts[1]).upper()

    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token, flags=re.IGNORECASE)
    if not match:
        return FileLabel(
            model=f"{ticker}_{period_token}",
            ticker=ticker,
            model_period=period_token,
            model_date="",
        )

    period_part = match.group(1).title()
    month_part = match.group(2)
    year_part = int(match.group(3))

    month_key = month_part[:3].lower()
    month_num = MONTH_MAP.get(month_key)
    if month_num is None:
        return FileLabel(
            model=f"{ticker}_{period_token}",
            ticker=ticker,
            model_period=period_token,
            model_date="",
        )

    month_abbrev = month_key.title()
    model_period = f"{period_part}{month_abbrev}_{year_part}"
    day = DAY_BY_PERIOD[period_part.lower()]
    model_date = date(year_part, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def ensure_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def find_anchor_max(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = ensure_matrix(used.value)
    if not matrix:
        return None

    start_row = used.row
    start_col = used.column

    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return start_row + r_idx, start_col + c_idx

    return None


def get_row_headers_near_anchor(sheet: xw.Sheet, header_row: int, anchor_col: int, span: int = 24) -> Dict[int, str]:
    first_col = max(1, anchor_col - span)
    last_col = anchor_col + span
    values = sheet.range((header_row, first_col), (header_row, last_col)).value
    if not isinstance(values, list):
        values = [values]
    return {first_col + idx: normalize_text(value) for idx, value in enumerate(values)}


def find_col_by_keywords(
    header_map: Dict[int, str],
    keywords: Sequence[str],
    fallback_col: Optional[int] = None,
) -> Optional[int]:
    for col, header in header_map.items():
        if not header:
            continue
        for keyword in keywords:
            if keyword in header:
                return col
    return fallback_col


def set_formula2_r1c1(cell: xw.Range, formula: str) -> bool:
    """
    Set an R1C1 formula with Formula2 where possible.
    """
    try:
        cell.api.Formula2R1C1 = formula
        return True
    except Exception:
        pass

    try:
        cell.api.Formula2 = formula
        return True
    except Exception:
        pass

    try:
        cell.formula2 = formula
        return True
    except Exception:
        return False


def close_source_workbook_safely(wb: xw.Book) -> None:
    # Preferred close path.
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # COM fallback for platforms/backends that do not support save kwarg.
    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    # Last resort; xlwings close defaults to no-save behavior.
    try:
        wb.close()
    except Exception:
        pass


def process_empirical_sheet(wb: xw.Book, file_label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    if EMPIRICAL_SOURCE_SHEET not in [sheet.name for sheet in wb.sheets]:
        print(f"  skipped empirical sheet: {EMPIRICAL_SOURCE_SHEET!r} not found")
        return []

    sheet = wb.sheets[EMPIRICAL_SOURCE_SHEET]
    anchor = find_anchor_max(sheet)
    if anchor is None:
        print("  skipped empirical sheet: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = get_row_headers_near_anchor(sheet, anchor_row, anchor_col)

    max_col = anchor_col
    min_col = find_col_by_keywords(header_map, ["min"], fallback_col=anchor_col - 1)
    forecast_col = find_col_by_keywords(
        header_map,
        ["estimated total sold", "tot fcst", "forecast", "fcst"],
        fallback_col=anchor_col - 2,
    )
    actual_col = find_col_by_keywords(
        header_map,
        ["reported sales", "actual sales", "actual"],
        fallback_col=anchor_col - 3,
    )
    num_quarters_col = find_col_by_keywords(
        header_map,
        ["num quarters", "quarters used", "num qtrs", "n quarters", "n qtrs"],
        fallback_col=anchor_col - 6,
    )
    last_quarter_col = find_col_by_keywords(
        header_map,
        ["last quarter", "last qtr", "latest quarter"],
        fallback_col=anchor_col - 5,
    )
    avg_pen_col = find_col_by_keywords(
        header_map,
        ["avg penetration", "average penetration", "avg pen"],
        fallback_col=anchor_col - 4,
    )
    penetration_history_col = find_col_by_keywords(
        header_map,
        ["penetration", "pen %"],
        fallback_col=anchor_col - 4,
    )
    quarterly_sales_col = find_col_by_keywords(
        header_map,
        ["quarterly sales", "qtr sales", "sales in db", "db sales"],
        fallback_col=anchor_col - 7,
    )
    growth_col = find_col_by_keywords(
        header_map,
        ["growth rate", "growth %"],
        fallback_col=anchor_col - 8,
    )
    captured_col = find_col_by_keywords(
        header_map,
        ["captured in db", "sales captured", "capture %", "db pct"],
        fallback_col=anchor_col - 9,
    )

    start_row = anchor_row + 1
    rows = list(range(start_row, start_row + N_QUARTERS))

    # Build average-penetration formulas in one batch, then calculate once.
    scratch_start_row = max(sheet.used_range.last_cell.row + 2, anchor_row + N_QUARTERS + 2)
    scratch_col = max_col + 2
    avg_cells: List[xw.Range] = []
    formula_updates = False

    history_rows: List[int] = []
    if penetration_history_col is not None and anchor_row > 1:
        pen_values = sheet.range((1, penetration_history_col), (anchor_row - 1, penetration_history_col)).value
        if not isinstance(pen_values, list):
            pen_values = [pen_values]
        for idx, value in enumerate(pen_values, start=1):
            if is_number(value):
                history_rows.append(idx)

    for idx, row in enumerate(rows):
        scratch_row = scratch_start_row + idx
        scratch_cell = sheet.range((scratch_row, scratch_col))
        avg_cells.append(scratch_cell)

        n_quarters_guess = to_int(
            sheet.range((row, num_quarters_col)).value if num_quarters_col is not None else None,
            default=idx + 1,
        )
        if history_rows:
            n = max(1, min(n_quarters_guess, len(history_rows)))
            h_start = history_rows[-n]
            h_end = history_rows[-1]
            formula = f'=IFERROR(AVERAGE(R{h_start}C{penetration_history_col}:R{h_end}C{penetration_history_col}),"")'
            formula_updates = set_formula2_r1c1(scratch_cell, formula) or formula_updates

    if formula_updates:
        wb.app.calculate()

    empirical_rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(rows):
        num_quarters_used = to_int(
            sheet.range((row, num_quarters_col)).value if num_quarters_col is not None else None,
            default=idx + 1,
        )

        last_quarter_used = sheet.range((row, last_quarter_col)).value if last_quarter_col is not None else None
        forecast_max = to_float(sheet.range((row, max_col)).value)
        forecast_min = to_float(sheet.range((row, min_col)).value if min_col is not None else None)
        forecast_value = to_float(sheet.range((row, forecast_col)).value if forecast_col is not None else None)
        actual_value = to_float(sheet.range((row, actual_col)).value if actual_col is not None else None)

        avg_penetration_pct = to_float(avg_cells[idx].value)
        if avg_penetration_pct is None and avg_pen_col is not None:
            avg_penetration_pct = to_float(sheet.range((row, avg_pen_col)).value)

        quarterly_sales = to_float(sheet.range((row, quarterly_sales_col)).value if quarterly_sales_col is not None else None)
        reported_sales = to_float(sheet.range((row, actual_col)).value if actual_col is not None else None)
        growth_rate_pct = to_float(sheet.range((row, growth_col)).value if growth_col is not None else None)
        sales_captured_in_db_pct = to_float(sheet.range((row, captured_col)).value if captured_col is not None else None)

        if forecast_value is None:
            # Useful fallback when table stores DB sales + penetration but not explicit forecast.
            forecast_value = safe_div(quarterly_sales, avg_penetration_pct)
        if reported_sales is None:
            reported_sales = actual_value
        if actual_value is None:
            actual_value = reported_sales
        if sales_captured_in_db_pct is None:
            sales_captured_in_db_pct = safe_div(quarterly_sales, forecast_value)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        # Skip fully empty candidate rows.
        row_signal = [forecast_max, forecast_min, forecast_value, actual_value, avg_penetration_pct]
        if all(value is None for value in row_signal):
            continue

        empirical_rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return empirical_rows


def process_regression_sheet(wb: xw.Book, file_label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    if REGRESSION_SOURCE_SHEET not in [sheet.name for sheet in wb.sheets]:
        print(f"  skipped regression sheet: {REGRESSION_SOURCE_SHEET!r} not found")
        return []

    sheet = wb.sheets[REGRESSION_SOURCE_SHEET]
    anchor = find_anchor_max(sheet)
    if anchor is None:
        print("  skipped regression sheet: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor

    # Required by spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    last_row = max(sheet.used_range.last_cell.row, anchor_row)
    x_values = sheet.range((1, x_col), (last_row, x_col)).value
    y_values = sheet.range((1, y_col), (last_row, y_col)).value
    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]

    xy_rows: List[Tuple[int, float, float]] = []
    for row_idx, (x_val, y_val) in enumerate(zip(x_values, y_values), start=1):
        if is_number(x_val) and is_number(y_val):
            xy_rows.append((row_idx, float(x_val), float(y_val)))

    if len(xy_rows) < 2:
        print("  skipped regression sheet: insufficient numeric x/y history")
        return []

    max_n = min(N_QUARTERS, len(xy_rows))
    header_map = get_row_headers_near_anchor(sheet, anchor_row, anchor_col)
    actual_col = find_col_by_keywords(header_map, ["actual sales", "actual", "reported sales"], fallback_col=None)

    forecast_x = to_float(sheet.range((anchor_row, x_col)).value)
    if forecast_x is None:
        forecast_x = xy_rows[-1][1]

    # Scratch area to hold temporary formulas for one-pass calculation.
    scratch_start_row = max(sheet.used_range.last_cell.row + 2, anchor_row + 2)
    intercept_col = anchor_col + 2
    slope_col = anchor_col + 3
    steyx_col = anchor_col + 4
    forecast_col = anchor_col + 5
    max_col = anchor_col + 6
    min_col = anchor_col + 7

    formula_updates = False
    for i in range(max_n):
        n = i + 1
        scratch_row = scratch_start_row + i
        start_row = xy_rows[-n][0]
        end_row = xy_rows[-1][0]

        sheet.range((scratch_row, anchor_col + 1)).value = n

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        steyx_formula = (
            f'=IFERROR(STEYX(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        forecast_formula = (
            f"=IFERROR({forecast_x}*R{scratch_row}C{slope_col}+R{scratch_row}C{intercept_col},\"\")"
        )
        max_formula = f"=IFERROR(R{scratch_row}C{forecast_col}+R{scratch_row}C{steyx_col},\"\")"
        min_formula = f"=IFERROR(R{scratch_row}C{forecast_col}-R{scratch_row}C{steyx_col},\"\")"

        formula_updates = set_formula2_r1c1(sheet.range((scratch_row, intercept_col)), intercept_formula) or formula_updates
        formula_updates = set_formula2_r1c1(sheet.range((scratch_row, slope_col)), slope_formula) or formula_updates
        formula_updates = set_formula2_r1c1(sheet.range((scratch_row, steyx_col)), steyx_formula) or formula_updates
        formula_updates = set_formula2_r1c1(sheet.range((scratch_row, forecast_col)), forecast_formula) or formula_updates
        formula_updates = set_formula2_r1c1(sheet.range((scratch_row, max_col)), max_formula) or formula_updates
        formula_updates = set_formula2_r1c1(sheet.range((scratch_row, min_col)), min_formula) or formula_updates

    if formula_updates:
        wb.app.calculate()

    regression_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(max_n):
        n = i + 1
        scratch_row = scratch_start_row + i

        intercept = to_float(sheet.range((scratch_row, intercept_col)).value)
        slope = to_float(sheet.range((scratch_row, slope_col)).value)
        forecast_total_without_sa = to_float(sheet.range((scratch_row, forecast_col)).value)
        forecast_max = to_float(sheet.range((scratch_row, max_col)).value)
        forecast_min = to_float(sheet.range((scratch_row, min_col)).value)
        actual_value = to_float(sheet.range((anchor_row + i + 1, actual_col)).value) if actual_col else None

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            round(forecast_total_without_sa, 10) if forecast_total_without_sa is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )

        # Prevent duplicate trailing rows when the last calculation repeats.
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        if all(value is None for value in [intercept, slope, forecast_total_without_sa, forecast_max, forecast_min]):
            continue

        regression_rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return regression_rows


def write_output_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(columns))

    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) if row.get(col) is not None else "" for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Reasonable column sizing without expensive full-sheet formatting scans.
    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = build_unique_output_path(in_dir, out_dir)

    files_to_process, skipped_files = list_input_files(in_dir, in_dir.name)
    for file_path, reason in skipped_files:
        print(f"skipped file: {file_path.name} ({reason})")

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in files_to_process:
            wb = None
            try:
                print(f"processed file: {file_path.name}")
                wb = app.books.open(str(file_path), update_links=False)
                file_label = parse_file_label(file_path)

                empirical_rows.extend(
                    process_empirical_sheet(
                        wb=wb,
                        file_label=file_label,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        wb=wb,
                        file_label=file_label,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook_safely(wb)
    finally:
        app.quit()

    output_wb = Workbook()
    default_ws = output_wb.active
    output_wb.remove(default_ws)

    write_output_sheet(output_wb, EMPIRICAL_OUTPUT_SHEET, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(output_wb, REGRESSION_OUTPUT_SHEET, REGRESSION_COLUMNS, regression_rows)
    output_wb.save(output_path)

    print(f"output path: {output_path}")
    print(f"files processed: {files_processed}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
