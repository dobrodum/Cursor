#!/usr/bin/env python3
"""Extract empirical/regression model candidates from .xlsx workbooks.

This script opens each source workbook once, processes both target sheets while
the workbook is open, and writes one consolidated output workbook.
"""

from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = Path("./input")
output_dir = Path("./output")


N_QUARTERS = 10

EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"

EMPIRICAL_COLUMNS: Sequence[str] = (
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
)

REGRESSION_COLUMNS: Sequence[str] = (
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
)

# Offsets from the "max" anchor column for empirical rows.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -13,
    "last_quarter_used": -12,
    "forecast_value": -1,  # estimated total sold
    "actual_value": +2,  # reported sales
    "forecast_max": 0,
    "forecast_min": +1,
    "quarterly_sales": -7,
    "reported_sales": -6,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -4,
}

# Offsets from the "max" anchor column for regression rows.
REGRESSION_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -1,  # TOT FCST w/o SA
    "actual_value": +2,  # optional if present
    "forecast_max": 0,
    "forecast_min": +1,
}

# Helper column offsets from anchor column for temporary formula writes.
EMPIRICAL_AVG_PEN_HELPER_OFFSET = +8
REGRESSION_INTERCEPT_HELPER_OFFSET = +8
REGRESSION_SLOPE_HELPER_OFFSET = +9


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


MONTH_INDEX = {
    **{m.lower(): i for i, m in enumerate(calendar.month_abbr) if m},
    **{m.lower(): i for i, m in enumerate(calendar.month_name) if m},
}

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}


def parse_model_metadata(file_name: str) -> ModelMeta:
    """Parse model metadata from source filename."""
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split("-")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1]

    period_token = ""
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    # Fallback: search period token anywhere in filename stem.
    if not period_token:
        match = re.search(r"(Early|Mid|Late)[A-Za-z]{3,9}\d{4}", stem, flags=re.IGNORECASE)
        if match:
            period_token = match.group(0)

    period_match = re.match(
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase_raw = period_match.group("phase")
        month_raw = period_match.group("month")
        year = int(period_match.group("year"))

        phase = phase_raw.title()
        month_key = month_raw.lower()
        month_num = MONTH_INDEX.get(month_key)
        if month_num is None and len(month_key) >= 3:
            month_num = MONTH_INDEX.get(month_key[:3])

        if month_num:
            month_abbr = calendar.month_abbr[month_num]
            day = DAY_BY_PHASE[phase.lower()]
            model_period = f"{phase}{month_abbr}_{year}"
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}".strip("_")
    return ModelMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_unique_output_path(input_folder: Path, out_dir: Path) -> Path:
    """Return a unique output path as {input_folder}_PARAM(.n).xlsx."""
    folder_name = input_folder.resolve().name
    base_name = f"{folder_name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_number(value: Any) -> Optional[float]:
    """Convert cell value to float where possible."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "").replace("%", "")
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    """Safely subtract numeric values."""
    left_num = safe_number(left)
    right_num = safe_number(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def is_effectively_empty(values: Iterable[Any]) -> bool:
    """True if all values are blank/None."""
    for value in values:
        if value not in (None, ""):
            return False
    return True


def normalize_matrix(values: Any) -> List[List[Any]]:
    """Normalize xlwings values into a 2D list."""
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
    """Find the first cell containing 'max' (case-insensitive)."""
    used = sheet.used_range
    values = normalize_matrix(used.options(ndim=2).value)
    start_row = used.row
    start_col = used.column

    for r_idx, row_vals in enumerate(values):
        for c_idx, value in enumerate(row_vals):
            if isinstance(value, str) and value.strip().lower() == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    """Write formula with formula2, fallback to formula."""
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_source_workbook(wb: xw.Book) -> None:
    """Close source workbook without saving, with safe fallback."""
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)  # type: ignore[attr-defined]
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception as exc:
        print(f"warning: failed to close workbook safely: {exc}")


def get_sheet_if_present(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    """Return sheet by name or None."""
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    min_col: int,
    max_col: int,
) -> List[List[Any]]:
    """Read a rectangular block once for performance."""
    values = sheet.range((start_row, min_col), (end_row, max_col)).options(ndim=2).value
    return normalize_matrix(values)


def block_value(
    block: Sequence[Sequence[Any]],
    row: int,
    col: int,
    start_row: int,
    min_col: int,
) -> Any:
    """Read a value from a pre-fetched block."""
    return block[row - start_row][col - min_col]


def extract_empirical_rows(
    wb: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    """Extract empirical candidates from one workbook."""
    sheet = get_sheet_if_present(wb, EMPIRICAL_SHEET)
    if sheet is None:
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"skipped empirical in {source_file}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
    avg_helper_col = anchor_col + EMPIRICAL_AVG_PEN_HELPER_OFFSET

    # Write one average-penetration formula per quarter count (R1C1/formula2).
    for n in range(1, N_QUARTERS + 1):
        row_idx = start_row + n - 1
        formula = f"=AVERAGE(R{start_row}C{captured_col}:R{row_idx}C{captured_col})"
        set_formula2(sheet.cells(row_idx, avg_helper_col), formula)

    # Recalculate once after formula updates.
    wb.app.calculate()

    cols_needed = {
        anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"],
        anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"],
        anchor_col + EMPIRICAL_OFFSETS["forecast_value"],
        anchor_col + EMPIRICAL_OFFSETS["actual_value"],
        anchor_col + EMPIRICAL_OFFSETS["forecast_max"],
        anchor_col + EMPIRICAL_OFFSETS["forecast_min"],
        anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"],
        anchor_col + EMPIRICAL_OFFSETS["reported_sales"],
        anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"],
        anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
        avg_helper_col,
    }
    min_col = min(cols_needed)
    max_col = max(cols_needed)
    block = read_block(sheet, start_row, end_row, min_col, max_col)

    rows: List[Dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_idx = start_row + n - 1

        num_quarters_used = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"],
            start_row,
            min_col,
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n

        last_quarter_used = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"],
            start_row,
            min_col,
        )
        forecast_value = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["forecast_value"],
            start_row,
            min_col,
        )
        actual_value = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["actual_value"],
            start_row,
            min_col,
        )
        forecast_max = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["forecast_max"],
            start_row,
            min_col,
        )
        forecast_min = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["forecast_min"],
            start_row,
            min_col,
        )
        avg_penetration_pct = block_value(
            block,
            row_idx,
            avg_helper_col,
            start_row,
            min_col,
        )
        quarterly_sales = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"],
            start_row,
            min_col,
        )
        reported_sales = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["reported_sales"],
            start_row,
            min_col,
        )
        growth_rate_pct = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"],
            start_row,
            min_col,
        )
        sales_captured_in_db_pct = block_value(
            block,
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
            start_row,
            min_col,
        )

        if is_effectively_empty(
            (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def regression_signature(row: Dict[str, Any]) -> Tuple[Any, ...]:
    """Hashable comparable tuple for de-duplication."""
    def r(value: Any) -> Any:
        numeric = safe_number(value)
        if numeric is None:
            return value
        return round(numeric, 10)

    return (
        r(row.get("num_quarters_used")),
        r(row.get("forecast_value")),
        r(row.get("forecast_max")),
        r(row.get("forecast_min")),
        r(row.get("intercept")),
        r(row.get("slope")),
    )


def extract_regression_rows(
    wb: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    """Extract regression candidates from one workbook."""
    sheet = get_sheet_if_present(wb, REGRESSION_SHEET)
    if sheet is None:
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"skipped regression in {source_file}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    intercept_col = anchor_col + REGRESSION_INTERCEPT_HELPER_OFFSET
    slope_col = anchor_col + REGRESSION_SLOPE_HELPER_OFFSET

    # Write INTERCEPT and SLOPE formulas for quarter windows.
    for n in range(1, N_QUARTERS + 1):
        row_idx = start_row + n - 1
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{row_idx}C{y_col},"
            f"R{start_row}C{x_col}:R{row_idx}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{row_idx}C{y_col},"
            f"R{start_row}C{x_col}:R{row_idx}C{x_col})"
        )
        set_formula2(sheet.cells(row_idx, intercept_col), intercept_formula)
        set_formula2(sheet.cells(row_idx, slope_col), slope_formula)

    # Recalculate once after formula updates.
    wb.app.calculate()

    cols_needed = {
        anchor_col + REGRESSION_OFFSETS["num_quarters_used"],
        anchor_col + REGRESSION_OFFSETS["forecast_value"],
        anchor_col + REGRESSION_OFFSETS["actual_value"],
        anchor_col + REGRESSION_OFFSETS["forecast_max"],
        anchor_col + REGRESSION_OFFSETS["forecast_min"],
        intercept_col,
        slope_col,
    }
    min_col = min(cols_needed)
    max_col = max(cols_needed)
    block = read_block(sheet, start_row, end_row, min_col, max_col)

    rows: List[Dict[str, Any]] = []
    previous_sig: Optional[Tuple[Any, ...]] = None

    for n in range(1, N_QUARTERS + 1):
        row_idx = start_row + n - 1

        num_quarters_used = block_value(
            block,
            row_idx,
            anchor_col + REGRESSION_OFFSETS["num_quarters_used"],
            start_row,
            min_col,
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n

        forecast_value = block_value(
            block,
            row_idx,
            anchor_col + REGRESSION_OFFSETS["forecast_value"],
            start_row,
            min_col,
        )
        actual_value = block_value(
            block,
            row_idx,
            anchor_col + REGRESSION_OFFSETS["actual_value"],
            start_row,
            min_col,
        )
        forecast_max = block_value(
            block,
            row_idx,
            anchor_col + REGRESSION_OFFSETS["forecast_max"],
            start_row,
            min_col,
        )
        forecast_min = block_value(
            block,
            row_idx,
            anchor_col + REGRESSION_OFFSETS["forecast_min"],
            start_row,
            min_col,
        )
        intercept = block_value(block, row_idx, intercept_col, start_row, min_col)
        slope = block_value(block, row_idx, slope_col, start_row, min_col)

        if is_effectively_empty((forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value not in (None, "") else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        sig = regression_signature(row)
        if previous_sig is not None and sig == previous_sig:
            continue

        rows.append(row)
        previous_sig = sig

    return rows


def write_output_sheet(
    wb: Workbook,
    title: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    """Write one sheet with formatting."""
    ws = wb.create_sheet(title=title)
    ws.append(list(columns))
    for item in rows:
        ws.append([item.get(col) for col in columns])

    # Header formatting.
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Reasonable column widths.
    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 50)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    """Write both output sheets into one workbook."""
    wb = Workbook()
    # Remove default empty sheet, then write both targets.
    wb.remove(wb.active)
    write_output_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def iter_source_files(folder: Path) -> Iterable[Path]:
    """Yield valid .xlsx files and print skip reasons for others."""
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        name = file_path.name
        if name.startswith("~"):
            print(f"skipped: {name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {name} (not .xlsx)")
            continue
        yield file_path


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        print(f"input directory does not exist: {in_dir}")
        return

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_unique_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calculation = app.calculation
    app.calculation = "manual"

    try:
        for file_path in iter_source_files(in_dir):
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                meta = parse_model_metadata(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))

                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.calculation = original_calculation
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
