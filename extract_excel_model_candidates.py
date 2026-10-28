from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Optional

import openpyxl
from openpyxl.styles import Font
import xlwings as xw


# ==============================
# User-configurable input/output
# ==============================
input_dir = Path("./input")
output_dir = Path("./output")


# ==============================
# Extraction settings
# ==============================
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
ANCHOR_TEXT = "max"
N_QUARTERS = 10

# Anchor-based offsets from the "max" cell for Empirical Model.
# These match the expected workbook template; adjust if your template changes.
EMPIRICAL_LAST_QUARTER_COL_OFFSET = -12
EMPIRICAL_PENETRATION_COL_OFFSET = -8
EMPIRICAL_QUARTERLY_SALES_COL_OFFSET = -7
EMPIRICAL_REPORTED_SALES_COL_OFFSET = -6
EMPIRICAL_GROWTH_RATE_COL_OFFSET = -5
EMPIRICAL_CAPTURED_PCT_COL_OFFSET = -4
EMPIRICAL_FORECAST_MAX_COL_OFFSET = 1
EMPIRICAL_FORECAST_MIN_ROW_OFFSET = 1
EMPIRICAL_FORECAST_MIN_COL_OFFSET = 1

# Helper cells are written in unused columns to avoid touching model inputs.
HELPER_COL_GAP = 6
AVG_PENETRATION_HELPER_ROW_OFFSET = 0
INTERCEPT_HELPER_ROW_OFFSET = 0
SLOPE_HELPER_ROW_OFFSET = 1


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


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        cleaned = str(value).replace(",", "").strip()
        if cleaned.endswith("%"):
            return float(cleaned[:-1]) / 100.0
        return float(cleaned)
    except Exception:
        return None


def to_iso_date(value: date) -> str:
    return value.strftime("%Y-%m-%d")


def parse_month(month_token: str) -> Optional[int]:
    month_token = month_token.strip()
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(month_token, fmt).month
        except ValueError:
            continue

    # Fallback: first three letters (e.g. "January" -> "Jan").
    short = month_token[:3]
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(short, fmt).month
        except ValueError:
            continue
    return None


def parse_file_label(file_path: Path) -> Optional[FileLabel]:
    """
    Expected stem format example:
      MedMiner_Model - AORT - MidJan2026_Send
    """
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1]
    period_chunk = parts[2].split("_")[0].strip()
    period_match = re.fullmatch(
        r"(?i)(early|mid|late)([a-z]+)(\d{4})",
        period_chunk,
    )
    if not period_match:
        return None

    period_bucket = period_match.group(1).title()
    month_token = period_match.group(2).title()
    year = int(period_match.group(3))
    month_num = parse_month(month_token)
    if month_num is None:
        return None

    month_abbrev = datetime(year, month_num, 1).strftime("%b")
    day_by_bucket = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_by_bucket[period_bucket]
    model_dt = date(year, month_num, model_day)

    model_period = f"{period_bucket}{month_abbrev}_{year}"
    model = f"{ticker}_{model_period}"
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=to_iso_date(model_dt),
    )


def next_output_path(input_folder_name: str, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base = out_dir / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    i = 1
    while True:
        candidate = out_dir / f"{input_folder_name}_PARAM.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def close_source_workbook_safely(wb: xw.Book) -> None:
    # xlwings versions differ: close(save=False) is preferred when supported.
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        # Fall through to API-level close fallback.
        pass

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        # Final fallback: close() with no args (still no explicit save call).
        wb.close()


def find_anchor(sheet: xw.Sheet, anchor_text: str = ANCHOR_TEXT) -> tuple[int, int]:
    # Fast COM find first (avoids scanning all cells in Python).
    try:
        found = sheet.api.UsedRange.Find(
            What=anchor_text,
            LookAt=1,  # xlWhole
            MatchCase=False,
        )
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    # Fallback: one pass over used_range values.
    used = sheet.used_range
    values = used.value
    row0 = int(used.row)
    col0 = int(used.column)

    if values is None:
        raise ValueError(f'Anchor "{anchor_text}" not found in sheet "{sheet.name}"')

    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    anchor_norm = anchor_text.strip().lower()
    for r_idx, row_vals in enumerate(values):
        for c_idx, cell_val in enumerate(row_vals):
            if normalize_text(cell_val) == anchor_norm:
                return row0 + r_idx, col0 + c_idx

    raise ValueError(f'Anchor "{anchor_text}" not found in sheet "{sheet.name}"')


def set_formula2_r1c1(target_cell: xw.Range, formula_r1c1: str) -> None:
    # Prefer Formula2R1C1 (Excel dynamic array engine), fallback to FormulaR1C1.
    try:
        target_cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        target_cell.api.FormulaR1C1 = formula_r1c1


def find_last_numeric_row(sheet: xw.Sheet, col: int, start_row: int, stop_row: int = 2) -> int:
    for row in range(start_row, stop_row - 1, -1):
        if to_float(sheet.cells(row, col).value) is not None:
            return row
    return stop_row


def extract_empirical_rows(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    if EMPIRICAL_SHEET_NAME not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    anchor_row, anchor_col = find_anchor(sheet, ANCHOR_TEXT)

    last_quarter_col = anchor_col + EMPIRICAL_LAST_QUARTER_COL_OFFSET
    penetration_col = anchor_col + EMPIRICAL_PENETRATION_COL_OFFSET
    quarterly_sales_col = anchor_col + EMPIRICAL_QUARTERLY_SALES_COL_OFFSET
    reported_sales_col = anchor_col + EMPIRICAL_REPORTED_SALES_COL_OFFSET
    growth_col = anchor_col + EMPIRICAL_GROWTH_RATE_COL_OFFSET
    captured_col = anchor_col + EMPIRICAL_CAPTURED_PCT_COL_OFFSET

    data_end_row = find_last_numeric_row(
        sheet=sheet,
        col=quarterly_sales_col,
        start_row=anchor_row - 1,
        stop_row=2,
    )
    data_start_row = max(2, data_end_row - (N_QUARTERS * 3))

    helper_col = sheet.used_range.last_cell.column + HELPER_COL_GAP
    helper_cell = sheet.cells(anchor_row + AVG_PENETRATION_HELPER_ROW_OFFSET, helper_col)

    forecast_max_val = to_float(
        sheet.cells(anchor_row, anchor_col + EMPIRICAL_FORECAST_MAX_COL_OFFSET).value
    )
    forecast_min_val = to_float(
        sheet.cells(
            anchor_row + EMPIRICAL_FORECAST_MIN_ROW_OFFSET,
            anchor_col + EMPIRICAL_FORECAST_MIN_COL_OFFSET,
        ).value
    )

    rows: list[dict[str, Any]] = []

    for n in range(1, N_QUARTERS + 1):
        start_row = max(data_start_row, data_end_row - n + 1)
        end_row = data_end_row
        if start_row > end_row:
            continue

        avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        set_formula2_r1c1(helper_cell, avg_formula)
        wb.app.calculate()
        avg_penetration = to_float(helper_cell.value)
        if avg_penetration is None or avg_penetration == 0:
            continue

        last_quarter_used = sheet.cells(end_row, last_quarter_col).value
        quarterly_sales = to_float(sheet.cells(end_row, quarterly_sales_col).value)
        reported_sales = to_float(sheet.cells(end_row, reported_sales_col).value)
        growth_rate = to_float(sheet.cells(end_row, growth_col).value)
        captured_pct = to_float(sheet.cells(end_row, captured_col).value)

        estimated_total_sold = None
        if quarterly_sales is not None and avg_penetration != 0:
            estimated_total_sold = quarterly_sales / avg_penetration

        # If the workbook does not expose explicit max/min next to the anchor,
        # derive them from windowed penetration bounds.
        dynamic_max = forecast_max_val
        dynamic_min = forecast_min_val
        if (dynamic_max is None or dynamic_min is None) and quarterly_sales is not None:
            pen_vals = []
            for row in range(start_row, end_row + 1):
                v = to_float(sheet.cells(row, penetration_col).value)
                if v is not None and v > 0:
                    pen_vals.append(v)
            if pen_vals:
                p_min = min(pen_vals)
                p_max = max(pen_vals)
                if p_min > 0:
                    dynamic_max = quarterly_sales / p_min
                if p_max > 0:
                    dynamic_min = quarterly_sales / p_max

        range_width = (
            (dynamic_max - dynamic_min)
            if dynamic_max is not None and dynamic_min is not None
            else None
        )

        row = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n,
            "last_quarter_used": last_quarter_used,
            "forecast_value": estimated_total_sold,
            "actual_value": reported_sales,
            "forecast_max": dynamic_max,
            "forecast_min": dynamic_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": captured_pct,
            "source_file": source_file,
        }
        rows.append(row)

    # Clear helper formula after extraction (no source save).
    helper_cell.value = None
    return rows


def extract_regression_rows(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    if REGRESSION_SHEET_NAME not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets[REGRESSION_SHEET_NAME]
    anchor_row, anchor_col = find_anchor(sheet, ANCHOR_TEXT)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    data_end_row = find_last_numeric_row(
        sheet=sheet,
        col=y_col,
        start_row=anchor_row - 1,
        stop_row=2,
    )
    data_start_row = max(2, data_end_row - (N_QUARTERS * 3))

    helper_col = sheet.used_range.last_cell.column + HELPER_COL_GAP
    intercept_cell = sheet.cells(anchor_row + INTERCEPT_HELPER_ROW_OFFSET, helper_col)
    slope_cell = sheet.cells(anchor_row + SLOPE_HELPER_ROW_OFFSET, helper_col)

    anchor_max = to_float(sheet.cells(anchor_row, anchor_col + 1).value)
    anchor_min = to_float(sheet.cells(anchor_row + 1, anchor_col + 1).value)

    rows: list[dict[str, Any]] = []
    previous_signature: Optional[tuple[Any, ...]] = None

    for n in range(1, N_QUARTERS + 1):
        start_row = max(data_start_row, data_end_row - n + 1)
        end_row = data_end_row
        if start_row > end_row:
            continue

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        next_x = to_float(sheet.cells(end_row + 1, x_col).value)
        if next_x is None:
            current_x = to_float(sheet.cells(end_row, x_col).value)
            next_x = (current_x + 1) if current_x is not None else float(n + 1)

        forecast_total_without_sa = intercept + (slope * next_x)

        # Prefer workbook-provided max/min anchored next to "max"/"min";
        # fallback to observed y-window bounds when absent.
        forecast_max = anchor_max
        forecast_min = anchor_min
        if forecast_max is None or forecast_min is None:
            y_vals = []
            for row in range(start_row, end_row + 1):
                y_val = to_float(sheet.cells(row, y_col).value)
                if y_val is not None:
                    y_vals.append(y_val)
            if y_vals:
                forecast_max = max(y_vals)
                forecast_min = min(y_vals)

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            round(intercept, 10),
            round(slope, 10),
            round(forecast_total_without_sa, 10),
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            # Prevent duplicate final row when incremental windows stop changing.
            continue
        previous_signature = signature

        row = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n,
            "num_quarters_used": n,
            "forecast_value": forecast_total_without_sa,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def autosize_worksheet_columns(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        max_len = 0
        for cell in col_cells:
            cell_val = "" if cell.value is None else str(cell.value)
            if len(cell_val) > max_len:
                max_len = len(cell_val)
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    ws_emp = wb_out.create_sheet("empirical_candidates")
    ws_reg = wb_out.create_sheet("regression_candidates")

    ws_emp.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_emp.append([row.get(c) for c in EMPIRICAL_COLUMNS])

    ws_reg.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_reg.append([row.get(c) for c in REGRESSION_COLUMNS])

    for ws in (ws_emp, ws_reg):
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autosize_worksheet_columns(ws)

    wb_out.save(output_path)


def process_all_workbooks() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_path = next_output_path(input_dir.resolve().name, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    processed_files = 0

    excel_app = xw.App(visible=False, add_book=False)
    excel_app.display_alerts = False
    excel_app.screen_updating = False

    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"SKIP: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"SKIP: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIP: {file_path.name} (not .xlsx)")
                continue

            label = parse_file_label(file_path)
            if label is None:
                print(f"SKIP: {file_path.name} (cannot parse ticker/model period from filename)")
                continue

            print(f"PROCESS: {file_path.name}")
            wb = None
            try:
                wb = excel_app.books.open(str(file_path), update_links=False)
                try:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            file_label=label,
                            source_file=file_path.name,
                        )
                    )
                except Exception as exc:
                    print(f"SKIP: {file_path.name} (empirical extraction error: {exc})")

                try:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            file_label=label,
                            source_file=file_path.name,
                        )
                    )
                except Exception as exc:
                    print(f"SKIP: {file_path.name} (regression extraction error: {exc})")

                processed_files += 1
            except Exception as exc:
                print(f"SKIP: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook_safely(wb)
    finally:
        excel_app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_all_workbooks()
