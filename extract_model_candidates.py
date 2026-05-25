#!/usr/bin/env python3
"""Extract empirical and regression candidate rows from Excel model workbooks.

This script opens each source workbook once, processes both model sheets while
it is open, and writes a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import math
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency
    raise SystemExit(
        "xlwings is required. Install it with: pip install xlwings"
    ) from exc

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -------------------------------
# User-configurable paths
# -------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


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


# Anchor-based offsets relative to the "max" anchor cell.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "avg_penetration_pct": -7,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
    "forecast_value": -2,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -9,
    "forecast_value": -2,  # TOT FCST w/o SA
    "actual_value": -1,  # optional, may be blank
    "forecast_max": 0,
    "forecast_min": 1,
}

N_QUARTERS = 10
METHOD_EMPIRICAL = "empirical"
METHOD_REGRESSION = "regression"


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: Optional[date]


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return False


def clean_value(value: Any) -> Any:
    if is_blank(value):
        return None
    return value


def numeric(value: Any) -> Optional[float]:
    value = clean_value(value)
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).strip().replace(",", "")
        if text == "":
            return None
        return float(text)
    except Exception:
        return None


def parse_month_token(token: str) -> int:
    token = token.strip()
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token, fmt).month
        except ValueError:
            continue

    token_3 = token[:3]
    try:
        return datetime.strptime(token_3, "%b").month
    except ValueError:
        pass

    month_map = {
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
    key = token.lower()[:3]
    if key not in month_map:
        raise ValueError(f"Unrecognized month token: {token}")
    return month_map[key]


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].strip().upper() if len(parts) >= 2 else "UNKNOWN"
    period_chunk = parts[2].strip() if len(parts) >= 3 else ""
    period_chunk = period_chunk.split("_")[0]

    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_chunk, re.IGNORECASE)
    model_period = period_chunk or "UnknownPeriod"
    model_date_obj: Optional[date] = None

    if match:
        cadence = match.group(1).title()
        month_token = match.group(2).title()
        year = int(match.group(3))

        try:
            month_num = parse_month_token(month_token)
            month_abbr = datetime(year, month_num, 1).strftime("%b")
            model_period = f"{cadence}{month_abbr}_{year}"

            day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
            model_date_obj = date(year, month_num, day_lookup[cadence])
        except ValueError:
            # Keep fallback values when month token is malformed.
            pass

    model = f"{ticker}_{model_period}"
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date_obj,
    )


def model_date_string(value: Optional[date]) -> Optional[str]:
    if value is None:
        return None
    return value.isoformat()


def output_path_for_run(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    input_folder_name = input_folder.resolve().name
    base = out_folder / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    index = 1
    while True:
        candidate = out_folder / f"{input_folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        matrix = [[values]]
    elif values and not isinstance(values[0], list):
        matrix = [values]
    else:
        matrix = values

    for r_idx, row in enumerate(matrix):
        for c_idx, raw in enumerate(row):
            if isinstance(raw, str) and raw.strip().lower() == "max":
                return used.row + r_idx, used.column + c_idx
    return None


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
        # Final fallback: do not rethrow; we still quit app in finally.
        pass


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def calculate_app(wb: xw.Book) -> None:
    try:
        wb.app.calculate()
    except Exception:
        wb.app.api.Calculate()


def get_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    return clean_value(sheet.cells(row, col).value)


def find_last_non_empty_row(sheet: xw.Sheet, col: int, start_row: int) -> Optional[int]:
    for row in range(start_row, 0, -1):
        if not is_blank(sheet.cells(row, col).value):
            return row
    return None


def contiguous_block_start(sheet: xw.Sheet, col: int, end_row: int) -> int:
    row = end_row
    while row > 1 and not is_blank(sheet.cells(row - 1, col).value):
        row -= 1
    return row


def r1c1_ref(row: int, col: int) -> str:
    return f"R{row}C{col}"


def empirical_rows_from_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    rows: List[Dict[str, Any]] = []

    pen_col = anchor_col + EMPIRICAL_OFFSETS["avg_penetration_pct"]
    temp_formula_col = anchor_col + 6

    history_end = find_last_non_empty_row(sheet, pen_col, anchor_row - 1)
    history_start = None
    if history_end is not None:
        history_start = contiguous_block_start(sheet, pen_col, history_end)

    # Write all average-penetration formulas first, then calculate once.
    if history_end is not None and history_start is not None:
        for n in range(1, N_QUARTERS + 1):
            row = anchor_row + n
            calc_start = max(history_start, history_end - n + 1)
            formula = f"=AVERAGE({r1c1_ref(calc_start, pen_col)}:{r1c1_ref(history_end, pen_col)})"
            set_formula2(sheet.cells(row, temp_formula_col), formula)

        calculate_app(wb)

    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n

        num_quarters = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"])
        if is_blank(num_quarters):
            num_quarters = n

        last_quarter_used = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"])
        forecast_value = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_value"])
        forecast_max = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_max"])
        forecast_min = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_min"])
        quarterly_sales = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"])
        reported_sales = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["reported_sales"])
        growth_rate_pct = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"])
        sales_captured = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"])

        avg_penetration = get_value(sheet, row, temp_formula_col)
        if is_blank(avg_penetration):
            avg_penetration = get_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["avg_penetration_pct"])

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        max_num = numeric(forecast_max)
        min_num = numeric(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": model_date_string(label.model_date),
                "method": METHOD_EMPIRICAL,
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    # Clear temporary formula cells so no transient edits remain in memory.
    for n in range(1, N_QUARTERS + 1):
        sheet.cells(anchor_row + n, temp_formula_col).value = None

    return rows


def regression_rows_from_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    y_end = find_last_non_empty_row(sheet, y_col, anchor_row - 1)
    x_end = find_last_non_empty_row(sheet, x_col, anchor_row - 1)
    if y_end is None or x_end is None:
        return []

    history_end = min(y_end, x_end)
    history_start = max(contiguous_block_start(sheet, y_col, history_end), contiguous_block_start(sheet, x_col, history_end))
    available_points = history_end - history_start + 1
    if available_points < 2:
        return []

    max_n = min(N_QUARTERS, available_points)
    temp_intercept_col = anchor_col + 6
    temp_slope_col = anchor_col + 7
    temp_forecast_col = anchor_col + 8
    next_x_row = history_end + 1

    # Write all formulas first, then calculate once.
    for n in range(2, max_n + 1):
        row = anchor_row + n - 1
        reg_start = history_end - n + 1

        intercept_formula = (
            f"=INTERCEPT({r1c1_ref(reg_start, y_col)}:{r1c1_ref(history_end, y_col)},"
            f"{r1c1_ref(reg_start, x_col)}:{r1c1_ref(history_end, x_col)})"
        )
        slope_formula = (
            f"=SLOPE({r1c1_ref(reg_start, y_col)}:{r1c1_ref(history_end, y_col)},"
            f"{r1c1_ref(reg_start, x_col)}:{r1c1_ref(history_end, x_col)})"
        )
        forecast_formula = (
            f"={r1c1_ref(row, temp_intercept_col)}+"
            f"{r1c1_ref(row, temp_slope_col)}*{r1c1_ref(next_x_row, x_col)}"
        )

        set_formula2(sheet.cells(row, temp_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row, temp_slope_col), slope_formula)
        set_formula2(sheet.cells(row, temp_forecast_col), forecast_formula)

    calculate_app(wb)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, Any, Any, Any, Any]] = None

    for n in range(2, max_n + 1):
        row = anchor_row + n - 1

        num_quarters = get_value(sheet, row, anchor_col + REGRESSION_OFFSETS["num_quarters_used"])
        if is_blank(num_quarters):
            num_quarters = n

        intercept = get_value(sheet, row, temp_intercept_col)
        slope = get_value(sheet, row, temp_slope_col)
        forecast_value = get_value(sheet, row, temp_forecast_col)
        actual_value = get_value(sheet, row, anchor_col + REGRESSION_OFFSETS["actual_value"])
        forecast_max = get_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_max"])
        forecast_min = get_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_min"])

        if all(is_blank(v) for v in (intercept, slope, forecast_value, forecast_max, forecast_min)):
            continue

        max_num = numeric(forecast_max)
        min_num = numeric(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        signature = (
            round(numeric(intercept), 10) if numeric(intercept) is not None else intercept,
            round(numeric(slope), 10) if numeric(slope) is not None else slope,
            round(numeric(forecast_value), 10) if numeric(forecast_value) is not None else forecast_value,
            round(max_num, 10) if max_num is not None else forecast_max,
            round(min_num, 10) if min_num is not None else forecast_min,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": model_date_string(label.model_date),
                "method": METHOD_REGRESSION,
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    # Clear temporary formula cells so no transient edits remain in memory.
    for n in range(2, max_n + 1):
        row = anchor_row + n - 1
        sheet.cells(row, temp_intercept_col).value = None
        sheet.cells(row, temp_slope_col).value = None
        sheet.cells(row, temp_forecast_col).value = None

    return rows


def autosize_columns(ws) -> None:
    for col_idx, column_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
        max_len = 0
        for cell in column_cells:
            value = cell.value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        width = min(max(12, max_len + 2), 48)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_sheet(ws, columns: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def process_workbook(
    app: xw.App,
    file_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> bool:
    label = parse_file_label(file_path)
    wb: Optional[xw.Book] = None
    try:
        wb = app.books.open(str(file_path), update_links=False)

        if "Empirical Model" in [s.name for s in wb.sheets]:
            sheet_emp = wb.sheets["Empirical Model"]
            empirical_rows.extend(
                empirical_rows_from_sheet(
                    wb=wb,
                    sheet=sheet_emp,
                    label=label,
                    source_file=file_path.name,
                )
            )

        if "Regression Model" in [s.name for s in wb.sheets]:
            sheet_reg = wb.sheets["Regression Model"]
            regression_rows.extend(
                regression_rows_from_sheet(
                    wb=wb,
                    sheet=sheet_reg,
                    label=label,
                    source_file=file_path.name,
                )
            )

        return True
    finally:
        if wb is not None:
            close_workbook_safely(wb)


def is_source_xlsx(file_path: Path) -> Tuple[bool, str]:
    if not file_path.is_file():
        return False, "not a file"
    if file_path.name.startswith("~"):
        return False, "temp file"
    if file_path.suffix.lower() != ".xlsx":
        return False, "not an .xlsx file"
    return True, ""


def main() -> None:
    if not input_dir.exists():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_path_for_run(input_folder=input_dir, out_folder=output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = "manual"

        for file_path in sorted(input_dir.iterdir()):
            is_valid, reason = is_source_xlsx(file_path)
            if not is_valid:
                print(f"Skipped: {file_path.name} ({reason})")
                continue

            try:
                ok = process_workbook(
                    app=app,
                    file_path=file_path,
                    empirical_rows=empirical_rows,
                    regression_rows=regression_rows,
                )
                if ok:
                    processed_count += 1
                    print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")

        out_wb = Workbook()
        ws_emp = out_wb.active
        ws_emp.title = "empirical_candidates"
        write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

        ws_reg = out_wb.create_sheet("regression_candidates")
        write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)

        out_wb.save(out_path)

    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    print(f"Output path: {out_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
