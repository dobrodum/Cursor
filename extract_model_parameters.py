#!/usr/bin/env python3
"""
Extract empirical/regression model candidates from .xlsx files into one output workbook.

Requirements implemented:
- One hidden Excel app for the full run.
- Open each source workbook once, process both model sheets, then close without saving.
- Anchor-based logic from the "max" cell and R1C1 Formula2 writes for calculations.
- Output workbook with two formatted sheets:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------
# User-configurable locations
# ---------------------------
input_dir = Path(r"/path/to/input")
output_dir = Path(r"/path/to/output")


# ---------------------------
# Constants
# ---------------------------
N_QUARTERS = 10

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

PERIOD_RE = re.compile(
    r"(?i)\b(early|mid|late)\s*([A-Za-z]{3,9})\s*(\d{4})\b"
)


@dataclass(frozen=True)
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class XYPoint:
    row: int
    x: float
    y: float
    quarter_label: Any


def normalize_grid(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_column(values: Any) -> List[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def safe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            if text.endswith("%"):
                return float(text[:-1]) / 100.0
            return float(text)
        except ValueError:
            return None
    return None


def format_quarter_label(value: Any, fallback_row: int) -> str:
    if value is None:
        return f"row_{fallback_row}"
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text if text else f"row_{fallback_row}"


def parse_file_meta(filename: str) -> FileMeta:
    stem = Path(filename).stem
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]

    ticker = "UNKNOWN"
    period_source = stem
    if len(parts) >= 3:
        ticker = parts[1].upper()
        period_source = parts[2]
    elif len(parts) >= 2:
        ticker = parts[-2].upper()
        period_source = parts[-1]

    # Trim known suffixes like "_Send"
    period_source = period_source.split("_")[0]

    match = PERIOD_RE.search(period_source) or PERIOD_RE.search(stem)
    if not match:
        model_period = "UnknownPeriod"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    timing_raw, month_raw, year_raw = match.groups()
    timing_key = timing_raw.lower()
    month_key = month_raw.lower()
    month_num = MONTH_LOOKUP.get(month_key)
    if month_num is None and len(month_key) >= 3:
        month_num = MONTH_LOOKUP.get(month_key[:3])

    if month_num is None:
        model_period = "UnknownPeriod"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    year = int(year_raw)
    day = TIMING_DAY[timing_key]
    month_abbrev = date(year, month_num, 1).strftime("%b")
    timing_fmt = timing_raw.capitalize()

    model_period = f"{timing_fmt}{month_abbrev}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_path: Path, out_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    # Prefer Formula2R1C1 (modern dynamic formulas), fallback to FormulaR1C1.
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass
    try:
        cell.api.FormulaR1C1 = formula
    except Exception:
        # Last fallback through xlwings wrapper if API names differ on platform.
        cell.formula2 = formula


def close_workbook_safely(wb: xw.Book) -> None:
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
        # Intentionally swallow close errors after best-effort safe close attempts.
        pass


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    needle = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == needle:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, target: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = normalize_grid(used.value)
    if not values:
        return None

    target_norm = target.strip().lower()
    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target_norm:
                return (used.row + r_idx, used.column + c_idx)
    return None


def extract_xy_points(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    x_col: int,
    y_col: int,
    quarter_col: int,
) -> List[XYPoint]:
    if end_row < start_row:
        return []

    x_vals = normalize_column(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_vals = normalize_column(sheet.range((start_row, y_col), (end_row, y_col)).value)
    q_vals = normalize_column(sheet.range((start_row, quarter_col), (end_row, quarter_col)).value)

    points: List[XYPoint] = []
    for idx, (x_raw, y_raw, q_raw) in enumerate(zip(x_vals, y_vals, q_vals)):
        x_num = safe_float(x_raw)
        y_num = safe_float(y_raw)
        if x_num is None or y_num is None:
            continue
        points.append(
            XYPoint(
                row=start_row + idx,
                x=x_num,
                y=y_num,
                quarter_label=q_raw,
            )
        )
    return points


def process_empirical_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    meta: FileMeta,
    source_file: str,
) -> List[dict]:
    anchor = find_anchor_cell(sheet, target="max")
    if anchor is None:
        print(f"  - empirical skipped ({source_file}): no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    # Anchor-based offsets:
    # - reported-sales-like series (y) at anchor_col - 7
    # - quarterly-sales-like series (x) at anchor_col - 11
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = max(1, x_col - 1)

    used = sheet.used_range
    start_row = used.row
    end_row = anchor_row - 1
    points = extract_xy_points(
        sheet=sheet,
        start_row=start_row,
        end_row=end_row,
        x_col=x_col,
        y_col=y_col,
        quarter_col=quarter_col,
    )
    if not points:
        print(f"  - empirical skipped ({source_file}): no usable x/y points")
        return []

    max_quarters = min(N_QUARTERS, len(points))
    rows: List[dict] = []

    # Scratch cells also placed by anchor-based offsets; formulas are temporary.
    scratch_col = anchor_col + 2
    avg_cell = sheet.range((anchor_row + 2, scratch_col))
    forecast_cell = sheet.range((anchor_row + 3, scratch_col))
    stdev_cell = sheet.range((anchor_row + 4, scratch_col))
    forecast_max_cell = sheet.range((anchor_row + 5, scratch_col))
    forecast_min_cell = sheet.range((anchor_row + 6, scratch_col))

    for n in range(1, max_quarters + 1):
        subset = points[-n:]
        first_row = subset[0].row
        last_row = subset[-1].row

        avg_formula = (
            f'=AVERAGE(IFERROR(R{first_row}C{x_col}:R{last_row}C{x_col}/'
            f'R{first_row}C{y_col}:R{last_row}C{y_col},""))'
        )
        forecast_formula = (
            f"=IFERROR(R{last_row}C{x_col}/R{anchor_row + 2}C{scratch_col},\"\")"
        )
        stdev_formula = (
            f'=STDEV.P(IFERROR(R{first_row}C{x_col}:R{last_row}C{x_col}/'
            f'R{first_row}C{y_col}:R{last_row}C{y_col},""))'
        )
        max_formula = f"=R{anchor_row + 3}C{scratch_col}*(1+R{anchor_row + 4}C{scratch_col})"
        min_formula = (
            f"=MAX(0,R{anchor_row + 3}C{scratch_col}*(1-R{anchor_row + 4}C{scratch_col}))"
        )

        set_formula2_r1c1(avg_cell, avg_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)
        set_formula2_r1c1(stdev_cell, stdev_formula)
        set_formula2_r1c1(forecast_max_cell, max_formula)
        set_formula2_r1c1(forecast_min_cell, min_formula)

        wb.app.calculate()

        avg_penetration = safe_float(avg_cell.value)
        forecast_value = safe_float(forecast_cell.value)
        forecast_max = safe_float(forecast_max_cell.value)
        forecast_min = safe_float(forecast_min_cell.value)

        last_point = subset[-1]
        quarterly_sales = last_point.x
        reported_sales = last_point.y
        actual_value = reported_sales

        growth_rate_pct: Optional[float] = None
        if len(subset) >= 2 and subset[-2].x != 0:
            growth_rate_pct = (subset[-1].x / subset[-2].x) - 1.0

        sales_captured_in_db_pct: Optional[float] = None
        if reported_sales != 0:
            sales_captured_in_db_pct = quarterly_sales / reported_sales

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * 1.10
        if forecast_min is None and forecast_value is not None:
            forecast_min = max(0.0, forecast_value * 0.90)

        range_width: Optional[float] = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n,
                "last_quarter_used": format_quarter_label(last_point.quarter_label, last_point.row),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def _round_sig(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 10)


def process_regression_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    meta: FileMeta,
    source_file: str,
) -> List[dict]:
    anchor = find_anchor_cell(sheet, target="max")
    if anchor is None:
        print(f"  - regression skipped ({source_file}): no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    # Required anchor-based offsets from prompt.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = max(1, x_col - 1)

    used = sheet.used_range
    start_row = used.row
    end_row = anchor_row - 1
    points = extract_xy_points(
        sheet=sheet,
        start_row=start_row,
        end_row=end_row,
        x_col=x_col,
        y_col=y_col,
        quarter_col=quarter_col,
    )
    if not points:
        print(f"  - regression skipped ({source_file}): no usable x/y points")
        return []

    max_quarters = min(N_QUARTERS, len(points))
    rows: List[dict] = []

    scratch_col = anchor_col + 2
    intercept_cell = sheet.range((anchor_row + 10, scratch_col))
    slope_cell = sheet.range((anchor_row + 11, scratch_col))
    forecast_cell = sheet.range((anchor_row + 12, scratch_col))
    max_cell = sheet.range((anchor_row + 13, scratch_col))
    min_cell = sheet.range((anchor_row + 14, scratch_col))

    prev_signature: Optional[Tuple[Any, ...]] = None

    for n in range(1, max_quarters + 1):
        subset = points[-n:]
        first_row = subset[0].row
        last_row = subset[-1].row

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        forecast_formula = (
            f"=R{anchor_row + 10}C{scratch_col}+"
            f"R{anchor_row + 11}C{scratch_col}*R{last_row}C{x_col}"
        )
        max_formula = f"=MAX(R{first_row}C{y_col}:R{last_row}C{y_col})"
        min_formula = f"=MIN(R{first_row}C{y_col}:R{last_row}C{y_col})"

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)
        set_formula2_r1c1(max_cell, max_formula)
        set_formula2_r1c1(min_cell, min_formula)

        wb.app.calculate()

        intercept_value = safe_float(intercept_cell.value)
        slope_value = safe_float(slope_cell.value)
        forecast_value = safe_float(forecast_cell.value)
        forecast_max = safe_float(max_cell.value)
        forecast_min = safe_float(min_cell.value)

        range_width: Optional[float] = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            n,
            _round_sig(intercept_value),
            _round_sig(slope_value),
            _round_sig(forecast_value),
            _round_sig(forecast_max),
            _round_sig(forecast_min),
        )
        # Prevent duplicate final row.
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,  # TOT FCST w/o SA
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws, columns: Sequence[str], rows: Sequence[dict]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    # Header formatting
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Freeze top row and apply filter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"

    # Reasonable column widths
    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for col_cells in ws.iter_cols(
            min_col=col_idx,
            max_col=col_idx,
            min_row=1,
            max_row=ws.max_row,
        ):
            for cell in col_cells:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def write_output_workbook(
    path: Path,
    empirical_rows: Sequence[dict],
    regression_rows: Sequence[dict],
) -> None:
    wb = Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = wb.create_sheet("regression_candidates")
    write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)
    wb.save(path)


def iter_source_files(input_path: Path) -> Iterable[Path]:
    for path in sorted(input_path.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"SKIPPED {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"SKIPPED {path.name}: not an .xlsx file")
            continue
        yield path


def main() -> None:
    input_path = Path(input_dir).expanduser()
    out_path = Path(output_dir).expanduser()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    out_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(input_path, out_path)

    source_files = list(iter_source_files(input_path))
    empirical_rows: List[dict] = []
    regression_rows: List[dict] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in source_files:
            print(f"PROCESSING {file_path.name}")
            wb = app.books.open(str(file_path), update_links=False)
            try:
                meta = parse_file_meta(file_path.name)

                emp_sheet = get_sheet_case_insensitive(wb, EMPIRICAL_SHEET_NAME)
                reg_sheet = get_sheet_case_insensitive(wb, REGRESSION_SHEET_NAME)
                if emp_sheet is None and reg_sheet is None:
                    print(f"SKIPPED {file_path.name}: missing both model sheets")
                    continue

                if emp_sheet is None:
                    print(f"  - empirical skipped ({file_path.name}): missing '{EMPIRICAL_SHEET_NAME}' sheet")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            wb=wb,
                            sheet=emp_sheet,
                            meta=meta,
                            source_file=file_path.name,
                        )
                    )

                if reg_sheet is None:
                    print(f"  - regression skipped ({file_path.name}): missing '{REGRESSION_SHEET_NAME}' sheet")
                else:
                    regression_rows.extend(
                        process_regression_sheet(
                            wb=wb,
                            sheet=reg_sheet,
                            meta=meta,
                            source_file=file_path.name,
                        )
                    )

                processed_count += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: {exc}")
            finally:
                close_workbook_safely(wb)
    finally:
        app.quit()

    write_output_workbook(
        path=output_file,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"OUTPUT: {output_file}")
    print(f"FILES PROCESSED: {processed_count}")
    print(f"EMPIRICAL ROWS: {len(empirical_rows)}")
    print(f"REGRESSION ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
