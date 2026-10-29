#!/usr/bin/env python3
"""Batch extract empirical/regression model candidates from .xlsx workbooks."""

from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path(r"./input")
output_dir = Path(r"./output")

# ---------------------------------------------------------------------------
# Extraction constants
# ---------------------------------------------------------------------------
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
MAX_QUARTERS = 10

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

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}
PERIOD_RE = re.compile(
    r"(?P<phase>Early|Mid|Late)[\s_-]*(?P<month>[A-Za-z]{3,9})[\s_-]*(?P<year>\d{4})",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class EmpiricalOffsets:
    data_end_row_offset: int = -2
    quarter_label_col_offset: int = -13
    quarterly_sales_col_offset: int = -11
    growth_rate_col_offset: int = -9
    sales_captured_col_offset: int = -8
    reported_sales_col_offset: int = -7
    penetration_pct_col_offset: int = -6
    max_value_col_offset: int = 1
    min_row_offset: int = 1
    helper_col_offset: int = 6


@dataclass(frozen=True)
class RegressionOffsets:
    data_end_row_offset: int = -2
    forecast_row_offset: int = 1
    helper_col_offset: int = 6


EMP_OFFSETS = EmpiricalOffsets()
REG_OFFSETS = RegressionOffsets()


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    try:
        parsed = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    if math.isnan(parsed) or math.isinf(parsed):
        return None
    return parsed


def numeric_equal(a: Optional[float], b: Optional[float], tol: float = 1e-9) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def parse_month(month_token: str) -> Optional[int]:
    compact = month_token.strip().lower()
    month_by_name = {
        name.lower(): idx for idx, name in enumerate(calendar.month_name) if name
    }
    month_by_abbr = {
        abbr.lower(): idx for idx, abbr in enumerate(calendar.month_abbr) if abbr
    }
    if compact in month_by_name:
        return month_by_name[compact]
    if compact in month_by_abbr:
        return month_by_abbr[compact]
    compact_three = compact[:3]
    return month_by_abbr.get(compact_three)


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_", 1)[0].strip()

    model_period = period_source
    model_date = ""

    match = PERIOD_RE.search(period_source)
    if match:
        phase = match.group("phase").title()
        month_token = match.group("month")
        year = int(match.group("year"))
        month = parse_month(month_token)
        if month is not None:
            month_abbr = calendar.month_abbr[month]
            day = PHASE_DAY.get(phase.lower(), 15)
            model_period = f"{phase}{month_abbr}_{year}"
            model_date = date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{idx}.xlsx"
        idx += 1
    return candidate


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    try:
        found = sheet.api.Cells.Find(
            What=anchor_text,
            LookIn=-4163,  # xlValues
            LookAt=1,      # xlWhole
            SearchOrder=1,  # xlByRows
            SearchDirection=1,  # xlNext
            MatchCase=False,
        )
    except Exception:
        found = None

    if found is None:
        return None
    return int(found.Row), int(found.Column)


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        # Fallback in case Formula2R1C1 is unavailable in this Excel build.
        cell.formula2 = formula_r1c1


def clear_formula(cell: xw.Range) -> None:
    try:
        cell.api.Formula2 = ""
    except Exception:
        cell.value = None


def safe_close_no_save(wb: xw.Book) -> None:
    close_attempts = (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    )
    last_error: Optional[Exception] = None
    for attempt in close_attempts:
        try:
            attempt()
            return
        except Exception as exc:  # pragma: no cover - COM variation
            last_error = exc
    if last_error is not None:
        raise last_error


def get_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def extract_empirical_rows(
    wb: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        print(f"SKIP {source_file}: missing sheet '{EMPIRICAL_SHEET_NAME}'")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIP {source_file}: no 'max' anchor found in '{EMPIRICAL_SHEET_NAME}'")
        return []

    anchor_row, anchor_col = anchor
    data_end_row = anchor_row + EMP_OFFSETS.data_end_row_offset
    if data_end_row < 1:
        print(f"SKIP {source_file}: invalid empirical anchor position")
        return []

    quarter_col = anchor_col + EMP_OFFSETS.quarter_label_col_offset
    quarterly_sales_col = anchor_col + EMP_OFFSETS.quarterly_sales_col_offset
    growth_rate_col = anchor_col + EMP_OFFSETS.growth_rate_col_offset
    captured_col = anchor_col + EMP_OFFSETS.sales_captured_col_offset
    reported_sales_col = anchor_col + EMP_OFFSETS.reported_sales_col_offset
    penetration_col = anchor_col + EMP_OFFSETS.penetration_pct_col_offset
    target_row = data_end_row + 1

    helper_col = anchor_col + EMP_OFFSETS.helper_col_offset
    avg_cell = sheet.cells(anchor_row, helper_col)
    forecast_cell = sheet.cells(anchor_row + 1, helper_col)

    static_max = to_number(
        sheet.cells(anchor_row, anchor_col + EMP_OFFSETS.max_value_col_offset).value
    )
    static_min = to_number(
        sheet.cells(
            anchor_row + EMP_OFFSETS.min_row_offset,
            anchor_col + EMP_OFFSETS.max_value_col_offset,
        ).value
    )

    rows: List[Dict[str, Any]] = []
    try:
        for n_quarters in range(MAX_QUARTERS, 0, -1):
            data_start_row = data_end_row - n_quarters + 1
            if data_start_row < 1:
                continue

            avg_formula = (
                f'=IFERROR(AVERAGE(R{data_start_row}C{penetration_col}:'
                f'R{data_end_row}C{penetration_col}),"")'
            )
            forecast_formula = (
                f'=IFERROR(R{target_row}C{reported_sales_col}/'
                f'R{anchor_row}C{helper_col},"")'
            )

            set_formula2_r1c1(avg_cell, avg_formula)
            set_formula2_r1c1(forecast_cell, forecast_formula)
            wb.app.calculate()

            avg_penetration = to_number(avg_cell.value)
            forecast_value = to_number(forecast_cell.value)
            quarterly_sales = to_number(sheet.cells(target_row, quarterly_sales_col).value)
            reported_sales = to_number(sheet.cells(target_row, reported_sales_col).value)
            growth_rate_pct = to_number(sheet.cells(target_row, growth_rate_col).value)
            sales_captured_pct = to_number(sheet.cells(target_row, captured_col).value)
            last_quarter_used = sheet.cells(data_start_row, quarter_col).value

            range_width = (
                static_max - static_min
                if static_max is not None and static_min is not None
                else None
            )

            rows.append(
                {
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
                    "forecast_max": static_max,
                    "forecast_min": static_min,
                    "range_width": range_width,
                    "avg_penetration_pct": avg_penetration,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate_pct,
                    "sales_captured_in_db_pct": sales_captured_pct,
                    "source_file": source_file,
                }
            )
    finally:
        clear_formula(avg_cell)
        clear_formula(forecast_cell)
    return rows


def extract_regression_rows(
    wb: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, REGRESSION_SHEET_NAME)
    if sheet is None:
        print(f"SKIP {source_file}: missing sheet '{REGRESSION_SHEET_NAME}'")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIP {source_file}: no 'max' anchor found in '{REGRESSION_SHEET_NAME}'")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_end_row = anchor_row + REG_OFFSETS.data_end_row_offset
    forecast_row = data_end_row + REG_OFFSETS.forecast_row_offset
    if data_end_row < 1 or forecast_row < 1:
        print(f"SKIP {source_file}: invalid regression anchor position")
        return []

    helper_col = anchor_col + REG_OFFSETS.helper_col_offset
    intercept_cell = sheet.cells(anchor_row, helper_col)
    slope_cell = sheet.cells(anchor_row + 1, helper_col)
    forecast_cell = sheet.cells(anchor_row + 2, helper_col)
    max_cell = sheet.cells(anchor_row + 3, helper_col)
    min_cell = sheet.cells(anchor_row + 4, helper_col)

    rows: List[Dict[str, Any]] = []
    previous_calc: Optional[Tuple[Optional[float], ...]] = None

    try:
        for n_quarters in range(MAX_QUARTERS, 0, -1):
            data_start_row = data_end_row - n_quarters + 1
            if data_start_row < 1:
                continue

            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
            )
            forecast_formula = (
                f'=IFERROR(R{anchor_row}C{helper_col}+'
                f'R{anchor_row + 1}C{helper_col}*R{forecast_row}C{x_col},"")'
            )
            max_formula = f'=IFERROR(MAX(R{data_start_row}C{y_col}:R{data_end_row}C{y_col}),"")'
            min_formula = f'=IFERROR(MIN(R{data_start_row}C{y_col}:R{data_end_row}C{y_col}),"")'

            set_formula2_r1c1(intercept_cell, intercept_formula)
            set_formula2_r1c1(slope_cell, slope_formula)
            set_formula2_r1c1(forecast_cell, forecast_formula)
            set_formula2_r1c1(max_cell, max_formula)
            set_formula2_r1c1(min_cell, min_formula)
            wb.app.calculate()

            intercept = to_number(intercept_cell.value)
            slope = to_number(slope_cell.value)
            forecast_total_wo_sa = to_number(forecast_cell.value)
            forecast_max = to_number(max_cell.value)
            forecast_min = to_number(min_cell.value)
            actual_value = to_number(sheet.cells(forecast_row, y_col).value)
            range_width = (
                forecast_max - forecast_min
                if forecast_max is not None and forecast_min is not None
                else None
            )

            current_calc = (
                intercept,
                slope,
                forecast_total_wo_sa,
                forecast_max,
                forecast_min,
                actual_value,
            )
            if previous_calc is not None and all(
                numeric_equal(previous_calc[idx], current_calc[idx])
                for idx in range(len(current_calc))
            ):
                continue
            previous_calc = current_calc

            rows.append(
                {
                    "model": metadata["model"],
                    "ticker": metadata["ticker"],
                    "model_period": metadata["model_period"],
                    "model_date": metadata["model_date"],
                    "method": "regression",
                    "parameter_name": "num_quarters_used",
                    "parameter_value": n_quarters,
                    "num_quarters_used": n_quarters,
                    "forecast_value": forecast_total_wo_sa,
                    "actual_value": actual_value,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "intercept": intercept,
                    "slope": slope,
                    "source_file": source_file,
                }
            )
    finally:
        clear_formula(intercept_cell)
        clear_formula(slope_cell)
        clear_formula(forecast_cell)
        clear_formula(max_cell)
        clear_formula(min_cell)
    return rows


def apply_output_formatting(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    ws.freeze_panes = "A2"
    last_col = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col}{max(ws.max_row, 1)}"
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx in range(1, ws.max_column + 1):
        max_len = len(str(ws.cell(row=1, column=col_idx).value or ""))
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 45)


def write_sheet(
    wb_out: openpyxl.Workbook, sheet_name: str, columns: List[str], rows: Iterable[Dict[str, Any]]
) -> None:
    ws = wb_out.create_sheet(title=sheet_name)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])
    apply_output_formatting(ws)


def save_output_workbook(
    output_file: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb_out = openpyxl.Workbook()
    default_sheet = wb_out.active
    wb_out.remove(default_sheet)

    write_sheet(wb_out, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb_out, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    wb_out.save(output_file)
    wb_out.close()


def process_workbooks() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input folder does not exist: {input_path}")

    output_file = unique_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.api.EnableEvents = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"SKIP {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIP {file_path.name}: not .xlsx")
                continue

            print(f"PROCESS {file_path.name}")
            metadata = parse_file_metadata(file_path.name)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"SKIP {file_path.name}: processing error -> {exc}")
            finally:
                if wb is not None:
                    try:
                        safe_close_no_save(wb)
                    except Exception as close_exc:
                        print(f"SKIP {file_path.name}: close error -> {close_exc}")
    finally:
        try:
            app.calculation = "automatic"
        except Exception:
            pass
        app.quit()

    save_output_workbook(output_file, empirical_rows, regression_rows)
    print(f"OUTPUT {output_file}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
