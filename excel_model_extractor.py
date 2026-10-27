#!/usr/bin/env python3
"""
Extract empirical/regression candidate rows from all .xlsx files in an input
folder into one output workbook.

Performance notes:
- Uses one hidden Excel app for the full run.
- Opens each source workbook once, processes both model sheets, then closes.
- Never saves source workbooks.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path(r"/path/to/input")
output_dir = Path(r"/path/to/output")


# ---------------------------------------------------------------------------
# Output schemas (exact column order)
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Extraction controls / layout assumptions
# ---------------------------------------------------------------------------
N_QUARTERS = 10

# Empirical layout is anchored off the cell containing "max" in Empirical Model.
# These offsets keep all references anchor-based and avoid repeated sheet scans.
EMPIRICAL_OFFSETS = {
    # where model output values appear relative to anchor
    "output_row_offset": 1,
    "forecast_value_col_offset": -2,      # estimated total sold
    "actual_value_col_offset": -1,        # reported sales
    "reported_sales_col_offset": -1,
    # supporting columns (same row as historical quarter rows)
    "last_quarter_col_offset": -8,
    "quarterly_sales_col_offset": -7,
    "growth_rate_col_offset": -6,
    "sales_captured_col_offset": -5,
    # quarter-count input cell for empirical calculation
    "num_quarters_input_row_offset": -1,
    "num_quarters_input_col_offset": -9,
    # temporary formula target for avg penetration
    "avg_pen_formula_row_offset": -1,
    "avg_pen_formula_col_offset": -4,
    # historical series end row (usually one row above anchor/header row)
    "history_end_row_offset": -1,
}

# Regression layout anchored off "max" in Regression Model.
REGRESSION_OFFSETS = {
    "output_row_offset": 1,
    "forecast_value_col_offset": -2,      # TOT FCST w/o SA
    "actual_value_col_offset": -1,        # optional/blank
    "num_quarters_input_row_offset": -1,
    "num_quarters_input_col_offset": -9,
    "history_end_row_offset": -1,
}


DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
MONTHS = {
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


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def next_output_path(src_input_dir: Path, dst_output_dir: Path) -> Path:
    base = f"{src_input_dir.name}_PARAM.xlsx"
    candidate = dst_output_dir / base
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = dst_output_dir / f"{src_input_dir.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_file_labels(file_path: Path) -> Optional[FileLabels]:
    """
    Expected style example:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[-2].strip().upper()
    period_token = re.sub(r"_send$", "", parts[-1], flags=re.IGNORECASE).strip()

    match = re.match(
        r"^(?P<timing>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    timing_raw = match.group("timing")
    month_raw = match.group("month")
    year = int(match.group("year"))

    month_key = month_raw.lower()
    if month_key not in MONTHS:
        return None
    month_num = MONTHS[month_key]

    timing = timing_raw[0].upper() + timing_raw[1:].lower()
    month_abbr = date(2000, month_num, 1).strftime("%b")
    model_period = f"{timing}{month_abbr}_{year}"
    day = DAY_BY_PERIOD[timing.lower()]
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def normalize_2d(values: Any) -> List[List[Any]]:
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
    """
    Finds anchor cell containing exact text 'max' (case-insensitive) in used range.
    Prefers a 'max' with adjacent 'min' to the right if available.
    """
    used = sheet.used_range
    values = normalize_2d(used.options(ndim=2).value)
    if not values:
        return None

    max_hits: List[Tuple[int, int]] = []
    min_hits: set[Tuple[int, int]] = set()

    for r_idx, row in enumerate(values):
        for c_idx, val in enumerate(row):
            if isinstance(val, str):
                text = val.strip().lower()
                abs_row = used.row + r_idx
                abs_col = used.column + c_idx
                if text == "max":
                    max_hits.append((abs_row, abs_col))
                elif text == "min":
                    min_hits.add((abs_row, abs_col))

    if not max_hits:
        return None

    for row, col in max_hits:
        if (row, col + 1) in min_hits:
            return row, col
    return max_hits[0]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        raw = value.strip().replace(",", "")
        if raw == "":
            return None
        if raw.endswith("%"):
            raw_num = raw[:-1].strip()
            if raw_num == "":
                return None
            try:
                return float(raw_num) / 100.0
            except ValueError:
                return None
        try:
            return float(raw)
        except ValueError:
            return None
    return None


def subtract_or_none(a: Any, b: Any) -> Optional[float]:
    a_val = to_float(a)
    b_val = to_float(b)
    if a_val is None or b_val is None:
        return None
    return a_val - b_val


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # formula2 may not exist on older Excel builds; fall back safely.
        cell.formula = formula


def close_workbook_no_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    # Last fallback; alerts are disabled at app level to avoid save prompts.
    try:
        wb.close()
    except Exception:
        pass


def process_empirical_sheet(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        return []

    ws = wb.sheets["Empirical Model"]
    anchor = find_anchor_max(ws)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    output_row = anchor_row + EMPIRICAL_OFFSETS["output_row_offset"]
    history_end_row = anchor_row + EMPIRICAL_OFFSETS["history_end_row_offset"]

    num_q_cell = ws.range(
        (
            anchor_row + EMPIRICAL_OFFSETS["num_quarters_input_row_offset"],
            anchor_col + EMPIRICAL_OFFSETS["num_quarters_input_col_offset"],
        )
    )
    avg_pen_formula_cell = ws.range(
        (
            anchor_row + EMPIRICAL_OFFSETS["avg_pen_formula_row_offset"],
            anchor_col + EMPIRICAL_OFFSETS["avg_pen_formula_col_offset"],
        )
    )

    sales_captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_col_offset"]

    rows: List[Dict[str, Any]] = []
    for num_quarters in range(1, N_QUARTERS + 1):
        history_start_row = history_end_row - num_quarters + 1
        if history_start_row < 1:
            break

        num_q_cell.value = num_quarters
        avg_formula = (
            f'=IFERROR(AVERAGE('
            f'R{history_start_row}C{sales_captured_col}:'
            f'R{history_end_row}C{sales_captured_col}'
            f'),"")'
        )
        set_formula2(avg_pen_formula_cell, avg_formula)
        wb.app.calculate()

        avg_penetration_pct = avg_pen_formula_cell.value
        forecast_max = ws.range((output_row, anchor_col)).value
        forecast_min = ws.range((output_row, anchor_col + 1)).value
        forecast_value = ws.range(
            (output_row, anchor_col + EMPIRICAL_OFFSETS["forecast_value_col_offset"])
        ).value
        actual_value = ws.range(
            (output_row, anchor_col + EMPIRICAL_OFFSETS["actual_value_col_offset"])
        ).value

        last_quarter_used = ws.range(
            (history_start_row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_col_offset"])
        ).value
        quarterly_sales = ws.range(
            (
                history_start_row,
                anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col_offset"],
            )
        ).value
        growth_rate_pct = ws.range(
            (history_start_row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_col_offset"])
        ).value
        sales_captured_in_db_pct = ws.range((history_start_row, sales_captured_col)).value
        reported_sales = ws.range(
            (output_row, anchor_col + EMPIRICAL_OFFSETS["reported_sales_col_offset"])
        ).value

        # Skip empty calculation rows.
        if (
            forecast_value in (None, "")
            and forecast_max in (None, "")
            and forecast_min in (None, "")
            and avg_penetration_pct in (None, "")
        ):
            continue

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_or_none(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def signature_value(value: Any) -> Any:
    num = to_float(value)
    if num is not None:
        return round(num, 10)
    return value


def process_regression_sheet(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        return []

    ws = wb.sheets["Regression Model"]
    anchor = find_anchor_max(ws)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    output_row = anchor_row + REGRESSION_OFFSETS["output_row_offset"]
    history_end_row = anchor_row + REGRESSION_OFFSETS["history_end_row_offset"]

    # Required by prompt
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_q_cell = ws.range(
        (
            anchor_row + REGRESSION_OFFSETS["num_quarters_input_row_offset"],
            anchor_col + REGRESSION_OFFSETS["num_quarters_input_col_offset"],
        )
    )

    used_last_row = ws.used_range.last_cell.row
    used_last_col = ws.used_range.last_cell.column
    temp_row = used_last_row + 2
    temp_col = used_last_col + 2
    intercept_cell = ws.range((temp_row, temp_col))
    slope_cell = ws.range((temp_row, temp_col + 1))

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for num_quarters in range(1, N_QUARTERS + 1):
        start_row = history_end_row - num_quarters + 1
        if start_row < 1:
            break

        num_q_cell.value = num_quarters

        intercept_formula = (
            f'=IFERROR(INTERCEPT('
            f'R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}'
            f'),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE('
            f'R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}'
            f'),"")'
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value

        x_last = ws.range((history_end_row, x_col)).value
        forecast_calc: Optional[float] = None
        i_val = to_float(intercept)
        s_val = to_float(slope)
        x_last_val = to_float(x_last)
        if i_val is not None and s_val is not None and x_last_val is not None:
            forecast_calc = i_val + s_val * (x_last_val + 1.0)

        forecast_value = ws.range(
            (output_row, anchor_col + REGRESSION_OFFSETS["forecast_value_col_offset"])
        ).value
        if forecast_value in (None, ""):
            forecast_value = forecast_calc

        actual_value = ws.range(
            (output_row, anchor_col + REGRESSION_OFFSETS["actual_value_col_offset"])
        ).value
        forecast_max = ws.range((output_row, anchor_col)).value
        forecast_min = ws.range((output_row, anchor_col + 1)).value

        current_signature = (
            signature_value(intercept),
            signature_value(slope),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
        )
        # Prevent duplicate terminal row.
        if prev_signature is not None and current_signature == prev_signature:
            continue
        prev_signature = current_signature

        # Skip empty rows.
        if (
            forecast_value in (None, "")
            and intercept in (None, "")
            and slope in (None, "")
            and forecast_max in (None, "")
            and forecast_min in (None, "")
        ):
            continue

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_or_none(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: List[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def save_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def iter_candidate_files(src_dir: Path) -> Iterable[Path]:
    for item in sorted(src_dir.iterdir()):
        if not item.is_file():
            print(f"skipped: {item.name} (not a file)")
            continue
        if item.name.startswith("~"):
            print(f"skipped: {item.name} (temporary Excel file)")
            continue
        if item.suffix.lower() != ".xlsx":
            print(f"skipped: {item.name} (not .xlsx)")
            continue
        yield item


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir not found or not a directory: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in iter_candidate_files(input_dir):
            labels = parse_file_labels(file_path)
            if labels is None:
                print(f"skipped: {file_path.name} (filename did not match expected pattern)")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                file_empirical_rows = process_empirical_sheet(
                    wb=wb, labels=labels, source_file=file_path.name
                )
                file_regression_rows = process_regression_sheet(
                    wb=wb, labels=labels, source_file=file_path.name
                )

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_no_save(wb)
    finally:
        app.quit()

    save_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"output: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
