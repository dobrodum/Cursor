#!/usr/bin/env python3
"""
Extract empirical and regression model candidate rows from Excel workbooks.

The script processes every .xlsx file in input_dir, skips temporary files, opens
each workbook once, extracts both model sheets while the workbook is open, and
writes one consolidated output workbook.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configure paths here
# ---------------------------------------------------------------------------
input_dir = "input"
output_dir = "output"


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


# Anchor-relative offsets chosen to avoid repeated sheet scans.
EMPIRICAL_N_QUARTERS = 10
EMPIRICAL_FORECAST_VALUE_OFFSET = (0, -1)  # estimated total sold
EMPIRICAL_ACTUAL_VALUE_OFFSET = (0, -2)  # reported sales
EMPIRICAL_FORECAST_MAX_OFFSET = (0, 0)
EMPIRICAL_FORECAST_MIN_OFFSET = (0, 1)
EMPIRICAL_LAST_QUARTER_OFFSET = (-1, -11)
EMPIRICAL_QUARTERLY_SALES_OFFSET = (-1, -7)
EMPIRICAL_REPORTED_SALES_OFFSET = (-1, -6)
EMPIRICAL_GROWTH_RATE_OFFSET = (-1, -5)
EMPIRICAL_DB_CAPTURED_OFFSET = (-1, -4)
EMPIRICAL_PENETRATION_COL_OFFSET = -4

REGRESSION_N_QUARTERS = 10
REGRESSION_FORECAST_TOTAL_OFFSET = (0, -1)  # TOT FCST w/o SA
REGRESSION_ACTUAL_VALUE_OFFSET = (0, -2)  # optional
REGRESSION_FORECAST_MAX_OFFSET = (0, 0)
REGRESSION_FORECAST_MIN_OFFSET = (0, 1)

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


def parse_file_label(file_name: str) -> Optional[FileLabel]:
    """
    Parse ticker/model metadata from source filename.

    Expected example:
    MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = Path(file_name).stem
    pattern = re.compile(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
        r"(?P<period>Early|Mid|Late)\s*"
        r"(?P<month>[A-Za-z]{3})\s*"
        r"(?P<year>\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    period_part = match.group("period").title()
    month_abbrev = match.group("month").title()
    year = int(match.group("year"))

    try:
        month = datetime.strptime(month_abbrev, "%b").month
    except ValueError:
        return None

    day = DAY_BY_PERIOD[period_part.lower()]
    try:
        model_date = date(year, month, day).isoformat()
    except ValueError:
        return None

    model_period = f"{period_part}{month_abbrev}_{year}"
    model = f"{ticker}_{model_period}"
    return FileLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    input_folder_name = in_dir.resolve().name
    base_name = f"{input_folder_name}_PARAM.xlsx"
    output_path = out_dir / base_name
    index = 1
    while output_path.exists():
        output_path = out_dir / f"{input_folder_name}_PARAM.{index}.xlsx"
        index += 1
    return output_path


def safe_close_workbook(wb: xw.Book) -> None:
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
        wb.api.Close(SaveChanges=False)  # type: ignore[attr-defined]
    except Exception:
        # Last-resort best effort close without persisting source changes.
        pass


def find_anchor_cell(ws: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    """
    Find the first cell equal to anchor_text (case-insensitive) in used range.
    Returns absolute (row, col) if found.
    """
    used = ws.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        candidate = str(values).strip().lower()
        if candidate == anchor_text.lower():
            return (used.row, used.column)
        return None

    first_row = values[0] if values else None
    if first_row is None:
        return None
    if not isinstance(first_row, list):
        values = [values]

    for row_idx, row in enumerate(values):
        if not isinstance(row, list):
            row = [row]
        for col_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == anchor_text.lower():
                return (used.row + row_idx, used.column + col_idx)
    return None


def _cell_value(ws: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return ws.range((row, col)).value


def _to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_f = _to_float(max_value)
    min_f = _to_float(min_value)
    if max_f is None or min_f is None:
        return None
    return max_f - min_f


def process_empirical_sheet(
    wb: xw.Book,
    ws: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(ws, anchor_text="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    penetration_col = anchor_col + EMPIRICAL_PENETRATION_COL_OFFSET

    # Write all temp formulas first, then calculate once.
    temp_col = anchor_col + 8
    temp_rows: List[int] = []
    for n in range(1, EMPIRICAL_N_QUARTERS + 1):
        temp_row = anchor_row + 40 + n
        start_row = anchor_row - n
        end_row = anchor_row - 1
        formula = (
            f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:'
            f'R{end_row}C{penetration_col}),"")'
        )
        ws.range((temp_row, temp_col)).formula2 = formula
        temp_rows.append(temp_row)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n in range(1, EMPIRICAL_N_QUARTERS + 1):
        row = anchor_row + n
        end_row = anchor_row - 1

        forecast_value = _cell_value(
            ws,
            row + EMPIRICAL_FORECAST_VALUE_OFFSET[0],
            anchor_col + EMPIRICAL_FORECAST_VALUE_OFFSET[1],
        )
        actual_value = _cell_value(
            ws,
            row + EMPIRICAL_ACTUAL_VALUE_OFFSET[0],
            anchor_col + EMPIRICAL_ACTUAL_VALUE_OFFSET[1],
        )
        forecast_max = _cell_value(
            ws,
            row + EMPIRICAL_FORECAST_MAX_OFFSET[0],
            anchor_col + EMPIRICAL_FORECAST_MAX_OFFSET[1],
        )
        forecast_min = _cell_value(
            ws,
            row + EMPIRICAL_FORECAST_MIN_OFFSET[0],
            anchor_col + EMPIRICAL_FORECAST_MIN_OFFSET[1],
        )
        avg_penetration = _cell_value(ws, temp_rows[n - 1], temp_col)

        last_quarter_used = _cell_value(
            ws,
            end_row + EMPIRICAL_LAST_QUARTER_OFFSET[0],
            anchor_col + EMPIRICAL_LAST_QUARTER_OFFSET[1],
        )
        quarterly_sales = _cell_value(
            ws,
            end_row + EMPIRICAL_QUARTERLY_SALES_OFFSET[0],
            anchor_col + EMPIRICAL_QUARTERLY_SALES_OFFSET[1],
        )
        reported_sales = _cell_value(
            ws,
            end_row + EMPIRICAL_REPORTED_SALES_OFFSET[0],
            anchor_col + EMPIRICAL_REPORTED_SALES_OFFSET[1],
        )
        growth_rate_pct = _cell_value(
            ws,
            end_row + EMPIRICAL_GROWTH_RATE_OFFSET[0],
            anchor_col + EMPIRICAL_GROWTH_RATE_OFFSET[1],
        )
        db_captured_pct = _cell_value(
            ws,
            end_row + EMPIRICAL_DB_CAPTURED_OFFSET[0],
            anchor_col + EMPIRICAL_DB_CAPTURED_OFFSET[1],
        )

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": _range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": db_captured_pct,
                "source_file": source_file,
            }
        )

    # Optional cleanup of temporary formulas.
    ws.range((temp_rows[0], temp_col), (temp_rows[-1], temp_col)).clear_contents()
    return rows


def process_regression_sheet(
    wb: xw.Book,
    ws: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(ws, anchor_text="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    temp_intercept_col = anchor_col + 8
    temp_slope_col = anchor_col + 9
    temp_rows: List[int] = []

    for n in range(1, REGRESSION_N_QUARTERS + 1):
        temp_row = anchor_row + 60 + n
        start_row = anchor_row - n
        end_row = anchor_row - 1
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        ws.range((temp_row, temp_intercept_col)).formula2 = intercept_formula
        ws.range((temp_row, temp_slope_col)).formula2 = slope_formula
        temp_rows.append(temp_row)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n in range(1, REGRESSION_N_QUARTERS + 1):
        row = anchor_row + n

        intercept = _cell_value(ws, temp_rows[n - 1], temp_intercept_col)
        slope = _cell_value(ws, temp_rows[n - 1], temp_slope_col)
        forecast_value = _cell_value(
            ws,
            row + REGRESSION_FORECAST_TOTAL_OFFSET[0],
            anchor_col + REGRESSION_FORECAST_TOTAL_OFFSET[1],
        )
        actual_value = _cell_value(
            ws,
            row + REGRESSION_ACTUAL_VALUE_OFFSET[0],
            anchor_col + REGRESSION_ACTUAL_VALUE_OFFSET[1],
        )
        forecast_max = _cell_value(
            ws,
            row + REGRESSION_FORECAST_MAX_OFFSET[0],
            anchor_col + REGRESSION_FORECAST_MAX_OFFSET[1],
        )
        forecast_min = _cell_value(
            ws,
            row + REGRESSION_FORECAST_MIN_OFFSET[0],
            anchor_col + REGRESSION_FORECAST_MIN_OFFSET[1],
        )

        signature = (
            round(_to_float(intercept) or 0.0, 10),
            round(_to_float(slope) or 0.0, 10),
            round(_to_float(forecast_value) or 0.0, 10),
            round(_to_float(forecast_max) or 0.0, 10),
            round(_to_float(forecast_min) or 0.0, 10),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": _range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    ws.range((temp_rows[0], temp_intercept_col), (temp_rows[-1], temp_slope_col)).clear_contents()
    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    bold_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold_font

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{max(1, ws.max_row)}"

    for idx, header in enumerate(headers, start=1):
        col_values = [header]
        for value in ws.iter_cols(min_col=idx, max_col=idx, min_row=2, values_only=True):
            col_values.extend([v for v in value if v is not None])
        width = max(len(str(v)) for v in col_values) + 2 if col_values else 12
        ws.column_dimensions[get_column_letter(idx)].width = min(max(width, 12), 48)


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = unique_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
    except Exception:
        pass

    try:
        if not in_dir.exists():
            print(f"Skipped input folder: {in_dir} (missing directory)")
            workbooks: List[Path] = []
        else:
            workbooks = sorted(in_dir.iterdir(), key=lambda p: p.name.lower())

        for file_path in workbooks:
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (not a regular file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            label = parse_file_label(file_path.name)
            if label is None:
                print(f"Skipped file: {file_path.name} (filename parse failed)")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                try:
                    empirical_ws = wb.sheets["Empirical Model"]
                    empirical_rows.extend(
                        process_empirical_sheet(wb, empirical_ws, label, file_path.name)
                    )
                except KeyError:
                    print(
                        f"Skipped sheet: Empirical Model in {file_path.name} (sheet missing)"
                    )

                try:
                    regression_ws = wb.sheets["Regression Model"]
                    regression_rows.extend(
                        process_regression_sheet(wb, regression_ws, label, file_path.name)
                    )
                except KeyError:
                    print(
                        f"Skipped sheet: Regression Model in {file_path.name} (sheet missing)"
                    )

                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (open/process error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    out_wb = Workbook()
    # Remove default auto-created sheet.
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    out_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
