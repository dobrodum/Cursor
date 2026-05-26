#!/usr/bin/env python3
"""
Extract empirical and regression candidate parameters from source Excel models.

This script:
- opens each source workbook exactly once
- processes "Empirical Model" and "Regression Model" while the book is open
- writes a single output workbook with two sheets:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import datetime as dt
import math
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - dependency/runtime guard
    raise SystemExit("xlwings is required. Install it with: pip install xlwings") from exc


# =========================
# User-configured directories
# =========================
input_dir = r"./input"
output_dir = r"./output"


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


# Anchor-based empirical offsets (relative to the "max" anchor column).
EMPIRICAL_OFFSETS = {
    "quarter_col": -11,
    "quarterly_sales_col": -7,
    "reported_sales_col": -6,
    "growth_rate_col": -5,
    "sales_captured_col": -4,
    "penetration_col": -3,
}


REGRESSION_X_OFFSET = -11
REGRESSION_Y_OFFSET = -7
MAX_QUARTERS = 10


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_1d(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        return values
    return [values]


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def numeric_or_none(value: Any) -> Optional[float]:
    return float(value) if is_number(value) else None


def find_anchor_info(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Dict[str, int]]:
    used = sheet.used_range
    used_values = normalize_2d(used.value)
    if not used_values:
        return None

    top_row = used.row
    left_col = used.column
    bottom_row = top_row + len(used_values) - 1

    anchor_text = anchor_text.strip().lower()
    for r_idx, row_vals in enumerate(used_values):
        for c_idx, raw_value in enumerate(row_vals):
            if isinstance(raw_value, str) and raw_value.strip().lower() == anchor_text:
                return {
                    "anchor_row": top_row + r_idx,
                    "anchor_col": left_col + c_idx,
                    "used_top_row": top_row,
                    "used_bottom_row": bottom_row,
                }
    return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    """
    Write an R1C1 formula, preferring xlwings .formula2 and falling back safely.
    """
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        pass

    # xlwings can require COM API for Formula2R1C1 depending on platform/version.
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        cell.api.FormulaR1C1 = formula_r1c1


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


def month_from_token(month_token: str) -> Optional[int]:
    months = {
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
    return months.get(month_token.lower())


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    """
    Example:
        MedMiner_Model - AORT - MidJan2026_Send.xlsx
    -> ticker=AORT, model_period=MidJan_2026, model_date=2026-01-15
    """
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 and parts[1] else "UNKNOWN"
    period_chunk = parts[2] if len(parts) >= 3 else ""
    period_token = period_chunk.split("_")[0].strip()

    period_match = re.fullmatch(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    if not period_match:
        # Fallback: locate period token anywhere in the filename stem.
        period_match = re.search(
            r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )

    if period_match:
        period_part = period_match.group(1).capitalize()
        month_part = period_match.group(2).capitalize()
        year_part = period_match.group(3)
    else:
        period_part = "Mid"
        month_part = "Jan"
        year_part = "1900"

    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
    month_number = month_from_token(month_part) or 1
    model_day = day_lookup.get(period_part, 15)
    model_date = dt.date(int(year_part), month_number, model_day).isoformat()
    model_period = f"{period_part}{month_part}_{year_part}"
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    version = 1
    while True:
        candidate = output_path / f"{base_name}.{version}.xlsx"
        if not candidate.exists():
            return candidate
        version += 1


def read_column(sheet: xw.Sheet, start_row: int, end_row: int, col: int) -> List[Any]:
    if end_row < start_row or col < 1:
        return []
    values = sheet.range((start_row, col), (end_row, col)).value
    return normalize_1d(values)


def build_recent_rows(
    top_row: int,
    primary_values: Sequence[Any],
    secondary_values: Sequence[Any],
    max_rows: int = MAX_QUARTERS,
) -> List[int]:
    rows: List[int] = []
    max_index = min(len(primary_values), len(secondary_values)) - 1
    for idx in range(max_index, -1, -1):
        if is_number(primary_values[idx]) and is_number(secondary_values[idx]):
            rows.append(top_row + idx)
            if len(rows) >= max_rows:
                break
    rows.reverse()
    return rows


def process_empirical_sheet(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped {source_file} empirical extraction: missing 'Empirical Model' sheet")
        return rows

    anchor = find_anchor_info(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped {source_file} empirical extraction: 'max' anchor not found")
        return rows

    anchor_row = anchor["anchor_row"]
    anchor_col = anchor["anchor_col"]
    used_top_row = anchor["used_top_row"]
    used_bottom_row = anchor["used_bottom_row"]
    data_bottom_row = anchor_row - 1
    if data_bottom_row < used_top_row:
        print(f"Skipped {source_file} empirical extraction: no data rows above anchor")
        return rows

    quarter_col = anchor_col + EMPIRICAL_OFFSETS["quarter_col"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"]
    growth_rate_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_col"]
    sales_captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_col"]
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["penetration_col"]

    if min(
        quarter_col,
        quarterly_sales_col,
        reported_sales_col,
        growth_rate_col,
        sales_captured_col,
        penetration_col,
    ) < 1:
        print(f"Skipped {source_file} empirical extraction: anchor offsets out of bounds")
        return rows

    quarter_values = read_column(sheet, used_top_row, data_bottom_row, quarter_col)
    quarterly_sales_values = read_column(sheet, used_top_row, data_bottom_row, quarterly_sales_col)
    reported_sales_values = read_column(sheet, used_top_row, data_bottom_row, reported_sales_col)
    growth_rate_values = read_column(sheet, used_top_row, data_bottom_row, growth_rate_col)
    sales_captured_values = read_column(sheet, used_top_row, data_bottom_row, sales_captured_col)
    penetration_values = read_column(sheet, used_top_row, data_bottom_row, penetration_col)

    candidate_rows = build_recent_rows(
        used_top_row, primary_values=penetration_values, secondary_values=quarterly_sales_values, max_rows=MAX_QUARTERS
    )
    if not candidate_rows:
        print(f"Skipped {source_file} empirical extraction: no valid trailing quarter rows found")
        return rows

    row_to_index = {used_top_row + idx: idx for idx in range(data_bottom_row - used_top_row + 1)}

    scratch_row = used_bottom_row + 2
    scratch_col = max(anchor_col + 2, 1)
    avg_pen_cell = sheet.range((scratch_row, scratch_col))
    max_pen_cell = sheet.range((scratch_row, scratch_col + 1))
    min_pen_cell = sheet.range((scratch_row, scratch_col + 2))
    forecast_cell = sheet.range((scratch_row, scratch_col + 3))
    forecast_max_cell = sheet.range((scratch_row, scratch_col + 4))
    forecast_min_cell = sheet.range((scratch_row, scratch_col + 5))

    max_n = min(MAX_QUARTERS, len(candidate_rows))
    for n_quarters in range(1, max_n + 1):
        window_rows = candidate_rows[-n_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        set_formula2_r1c1(
            avg_pen_cell,
            f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col}),"")',
        )
        set_formula2_r1c1(
            max_pen_cell,
            f'=IFERROR(MAX(R{start_row}C{penetration_col}:R{end_row}C{penetration_col}),"")',
        )
        set_formula2_r1c1(
            min_pen_cell,
            f'=IFERROR(MIN(R{start_row}C{penetration_col}:R{end_row}C{penetration_col}),"")',
        )
        set_formula2_r1c1(
            forecast_cell,
            f'=IFERROR(R{end_row}C{quarterly_sales_col}/R{scratch_row}C{scratch_col},"")',
        )
        set_formula2_r1c1(
            forecast_max_cell,
            f'=IFERROR(R{end_row}C{quarterly_sales_col}/R{scratch_row}C{scratch_col + 2},"")',
        )
        set_formula2_r1c1(
            forecast_min_cell,
            f'=IFERROR(R{end_row}C{quarterly_sales_col}/R{scratch_row}C{scratch_col + 1},"")',
        )

        wb.app.calculate()

        scratch_values = normalize_1d(
            sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 5)).value
        )
        avg_penetration = numeric_or_none(scratch_values[0] if len(scratch_values) > 0 else None)
        forecast_value = numeric_or_none(scratch_values[3] if len(scratch_values) > 3 else None)
        forecast_max = numeric_or_none(scratch_values[4] if len(scratch_values) > 4 else None)
        forecast_min = numeric_or_none(scratch_values[5] if len(scratch_values) > 5 else None)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        latest_idx = row_to_index[end_row]
        first_idx = row_to_index[start_row]

        quarterly_sales = numeric_or_none(quarterly_sales_values[latest_idx] if latest_idx < len(quarterly_sales_values) else None)
        reported_sales = numeric_or_none(reported_sales_values[latest_idx] if latest_idx < len(reported_sales_values) else None)
        growth_rate = numeric_or_none(growth_rate_values[latest_idx] if latest_idx < len(growth_rate_values) else None)
        sales_captured_pct = numeric_or_none(
            sales_captured_values[latest_idx] if latest_idx < len(sales_captured_values) else None
        )
        last_quarter_used = quarter_values[first_idx] if first_idx < len(quarter_values) else None

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def nearly_equal(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    if is_number(a) and is_number(b):
        return abs(float(a) - float(b)) <= tolerance
    return a == b


def process_regression_sheet(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped {source_file} regression extraction: missing 'Regression Model' sheet")
        return rows

    anchor = find_anchor_info(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped {source_file} regression extraction: 'max' anchor not found")
        return rows

    anchor_row = anchor["anchor_row"]
    anchor_col = anchor["anchor_col"]
    used_top_row = anchor["used_top_row"]
    used_bottom_row = anchor["used_bottom_row"]
    data_bottom_row = anchor_row - 1
    if data_bottom_row < used_top_row:
        print(f"Skipped {source_file} regression extraction: no data rows above anchor")
        return rows

    y_col = anchor_col + REGRESSION_Y_OFFSET
    x_col = anchor_col + REGRESSION_X_OFFSET
    if min(y_col, x_col) < 1:
        print(f"Skipped {source_file} regression extraction: anchor offsets out of bounds")
        return rows

    y_values = read_column(sheet, used_top_row, data_bottom_row, y_col)
    x_values = read_column(sheet, used_top_row, data_bottom_row, x_col)
    candidate_rows = build_recent_rows(
        used_top_row,
        primary_values=y_values,
        secondary_values=x_values,
        max_rows=MAX_QUARTERS,
    )
    if len(candidate_rows) < 2:
        print(f"Skipped {source_file} regression extraction: insufficient valid rows for regression")
        return rows

    scratch_row = used_bottom_row + 3
    scratch_col = max(anchor_col + 2, 1)
    intercept_cell = sheet.range((scratch_row, scratch_col))
    slope_cell = sheet.range((scratch_row, scratch_col + 1))
    forecast_cell = sheet.range((scratch_row, scratch_col + 2))
    max_cell = sheet.range((scratch_row, scratch_col + 3))
    min_cell = sheet.range((scratch_row, scratch_col + 4))

    previous_calculated: Optional[Tuple[Any, Any, Any, Any, Any]] = None
    max_n = min(MAX_QUARTERS, len(candidate_rows))
    for n_quarters in range(2, max_n + 1):
        window_rows = candidate_rows[-n_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        set_formula2_r1c1(
            intercept_cell,
            (
                f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
            ),
        )
        set_formula2_r1c1(
            slope_cell,
            (
                f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
            ),
        )
        set_formula2_r1c1(
            forecast_cell,
            (
                f'=IFERROR(R{scratch_row}C{scratch_col}+R{scratch_row}C{scratch_col + 1}*'
                f'(MAX(R{start_row}C{x_col}:R{end_row}C{x_col})+1),"")'
            ),
        )
        set_formula2_r1c1(
            max_cell,
            f'=IFERROR(MAX(R{start_row}C{y_col}:R{end_row}C{y_col}),"")',
        )
        set_formula2_r1c1(
            min_cell,
            f'=IFERROR(MIN(R{start_row}C{y_col}:R{end_row}C{y_col}),"")',
        )

        wb.app.calculate()

        calculated = normalize_1d(
            sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 4)).value
        )
        intercept = numeric_or_none(calculated[0] if len(calculated) > 0 else None)
        slope = numeric_or_none(calculated[1] if len(calculated) > 1 else None)
        forecast_value = numeric_or_none(calculated[2] if len(calculated) > 2 else None)
        forecast_max = numeric_or_none(calculated[3] if len(calculated) > 3 else None)
        forecast_min = numeric_or_none(calculated[4] if len(calculated) > 4 else None)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        current_signature = (intercept, slope, forecast_value, forecast_max, forecast_min)
        if previous_calculated and all(
            nearly_equal(curr, prev) for curr, prev in zip(current_signature, previous_calculated)
        ):
            continue
        previous_calculated = current_signature

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(
    workbook: openpyxl.Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(column) for column in headers])

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
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 44)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    output_wb = openpyxl.Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    write_sheet(output_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(output_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)

    output_wb.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(input_path=input_path, output_path=output_path)
    output_name_pattern = re.compile(
        rf"^{re.escape(input_path.name)}_PARAM(?:\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        # Some versions/platforms may not expose this setter.
        pass

    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue
            if output_name_pattern.match(file_path.name):
                print(f"Skipped {file_path.name}: output workbook pattern")
                continue

            print(f"Processed {file_path.name}")
            metadata = parse_file_metadata(file_path.name)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed while extracting ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)
    print(f"Output workbook: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
