from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

# -----------------------------
# Constants
# -----------------------------
EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"
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

# Anchor-based offsets from the "max" cell.
EMPIRICAL_OFFSETS = {
    "quarterly_sales": -11,
    "reported_sales": -10,
    "growth_rate_pct": -9,
    "sales_captured_in_db_pct": -8,
    "num_quarters_used": -6,
    "last_quarter_used": -5,
    "forecast_value": -4,
    "actual_value": -3,
    "forecast_max": 0,
    "forecast_min": 1,
}
EMPIRICAL_HELPER_COL_OFFSET = 8

REGRESSION_OFFSETS = {
    "num_quarters_used": -8,
    "forecast_value": -2,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
}
REGRESSION_HELPER_INTERCEPT_COL_OFFSET = 6
REGRESSION_HELPER_SLOPE_COL_OFFSET = 7

PERIOD_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
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


def to_2d(value: Any) -> list[list[Any]]:
    if value is None:
        return []
    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        pct = cleaned.endswith("%")
        if pct:
            cleaned = cleaned[:-1]
        try:
            number = float(cleaned)
        except ValueError:
            return None
        return number / 100.0 if pct else number
    return None


def to_int(value: Any) -> int | None:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def normalize_value(value: Any) -> Any:
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    return value


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_without_save(wb: xw.main.Book) -> None:
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
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        pass


def parse_file_labels(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 and parts[1] else "UNKNOWN"
    trailing = parts[2] if len(parts) > 2 else ""
    token = trailing.split("_")[0].strip() if trailing else ""

    model_period = token if token else "UNKNOWN_PERIOD"
    model_date = ""

    match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        token,
        flags=re.IGNORECASE,
    )
    if match:
        period_part = match.group(1).title()
        month_part = match.group(2).title()
        year = int(match.group(3))
        day = PERIOD_DAY_MAP[period_part]
        month = MONTH_MAP[month_part.lower()]
        model_period = f"{period_part}{month_part}_{year}"
        model_date = dt.date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def find_anchor_max_cell(sheet: xw.main.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def get_sheet(wb: xw.main.Book, sheet_name: str) -> xw.main.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def extract_empirical_rows(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor_max_cell(sheet)
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: no 'max' anchor found.")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    helper_col = anchor_col + EMPIRICAL_HELPER_COL_OFFSET

    left_col = anchor_col + min(EMPIRICAL_OFFSETS.values())
    right_col = anchor_col + max(EMPIRICAL_OFFSETS.values())
    block = to_2d(sheet.range((start_row, left_col), (end_row, right_col)).value)

    if not block:
        return []

    rows: list[dict[str, Any]] = []
    helper_cells: list[tuple[int, int]] = []

    for idx, row_values in enumerate(block):
        excel_row = start_row + idx

        def get_offset_value(offset: int) -> Any:
            return row_values[(anchor_col + offset) - left_col]

        num_q = to_int(get_offset_value(EMPIRICAL_OFFSETS["num_quarters_used"])) or (idx + 1)
        sales_captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
        avg_start_row = max(start_row, excel_row - num_q + 1)

        helper_formula = (
            f'=IFERROR(AVERAGE(R{avg_start_row}C{sales_captured_col}:'
            f"R{excel_row}C{sales_captured_col}),\"\")"
        )
        set_formula2(sheet.range((excel_row, helper_col)), helper_formula)
        helper_cells.append((excel_row, helper_col))

    if helper_cells:
        wb.app.calculate()

    for idx, row_values in enumerate(block):
        excel_row = start_row + idx

        def get_offset_value(offset: int) -> Any:
            return row_values[(anchor_col + offset) - left_col]

        num_q = to_int(get_offset_value(EMPIRICAL_OFFSETS["num_quarters_used"])) or (idx + 1)
        last_quarter = normalize_value(get_offset_value(EMPIRICAL_OFFSETS["last_quarter_used"]))
        forecast_value = to_float(get_offset_value(EMPIRICAL_OFFSETS["forecast_value"]))
        actual_value = to_float(get_offset_value(EMPIRICAL_OFFSETS["actual_value"]))
        forecast_max = to_float(get_offset_value(EMPIRICAL_OFFSETS["forecast_max"]))
        forecast_min = to_float(get_offset_value(EMPIRICAL_OFFSETS["forecast_min"]))
        quarterly_sales = to_float(get_offset_value(EMPIRICAL_OFFSETS["quarterly_sales"]))
        reported_sales = to_float(get_offset_value(EMPIRICAL_OFFSETS["reported_sales"]))
        growth_rate = to_float(get_offset_value(EMPIRICAL_OFFSETS["growth_rate_pct"]))
        sales_captured = to_float(get_offset_value(EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]))
        avg_penetration = to_float(sheet.range((excel_row, helper_col)).value)

        if all(
            is_blank(v)
            for v in (
                num_q,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                **labels,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_q,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor_max_cell(sheet)
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: no 'max' anchor found.")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    last_data_row = anchor_row - 1

    left_col = anchor_col + min(REGRESSION_OFFSETS.values())
    right_col = anchor_col + max(REGRESSION_OFFSETS.values())
    block = to_2d(sheet.range((start_row, left_col), (end_row, right_col)).value)

    if not block:
        return []

    helper_intercept_col = anchor_col + REGRESSION_HELPER_INTERCEPT_COL_OFFSET
    helper_slope_col = anchor_col + REGRESSION_HELPER_SLOPE_COL_OFFSET
    pending_rows: list[dict[str, Any]] = []

    for idx, row_values in enumerate(block):
        excel_row = start_row + idx

        def get_offset_value(offset: int) -> Any:
            return row_values[(anchor_col + offset) - left_col]

        num_q = to_int(get_offset_value(REGRESSION_OFFSETS["num_quarters_used"])) or (idx + 1)
        data_start_row = last_data_row - num_q + 1
        if data_start_row < 1:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{last_data_row}C{y_col},'
            f"R{data_start_row}C{x_col}:R{last_data_row}C{x_col}),\"\")"
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{last_data_row}C{y_col},'
            f"R{data_start_row}C{x_col}:R{last_data_row}C{x_col}),\"\")"
        )

        set_formula2(sheet.range((excel_row, helper_intercept_col)), intercept_formula)
        set_formula2(sheet.range((excel_row, helper_slope_col)), slope_formula)

        pending_rows.append(
            {
                "excel_row": excel_row,
                "row_values": row_values,
                "num_q": num_q,
            }
        )

    if pending_rows:
        wb.app.calculate()

    output_rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for pending in pending_rows:
        excel_row = pending["excel_row"]
        row_values = pending["row_values"]
        num_q = pending["num_q"]

        def get_offset_value(offset: int) -> Any:
            return row_values[(anchor_col + offset) - left_col]

        forecast_value = to_float(get_offset_value(REGRESSION_OFFSETS["forecast_value"]))
        forecast_max = to_float(get_offset_value(REGRESSION_OFFSETS["forecast_max"]))
        forecast_min = to_float(get_offset_value(REGRESSION_OFFSETS["forecast_min"]))
        intercept = to_float(sheet.range((excel_row, helper_intercept_col)).value)
        slope = to_float(sheet.range((excel_row, helper_slope_col)).value)

        if forecast_value is None and intercept is not None and slope is not None:
            next_x = to_float(sheet.range((last_data_row + 1, x_col)).value)
            if next_x is not None:
                forecast_value = intercept + (slope * next_x)

        if all(is_blank(v) for v in (num_q, forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            num_q,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        output_rows.append(
            {
                **labels,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_q,
                "num_quarters_used": num_q,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return output_rows


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    base = output_folder / f"{input_folder.name}_PARAM.xlsx"
    if not base.exists():
        return base

    suffix = 1
    while True:
        candidate = output_folder / f"{input_folder.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_sheet(
    ws,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row in rows:
            text = "" if row.get(column) is None else str(row.get(column))
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 52)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def list_source_files(folder: Path) -> list[Path]:
    source_files: list[Path] = []
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary workbook.")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file.")
            continue
        source_files.append(file_path)
    return source_files


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    source_files = list_source_files(input_dir)

    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app: xw.main.App | None = None
    if source_files:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = "manual"

    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            wb: xw.main.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False) if app else None
                if wb is None:
                    print(f"Skipped {file_path.name}: workbook failed to open.")
                    continue

                labels = parse_file_labels(file_path)
                empirical_sheet = get_sheet(wb, EMPIRICAL_MODEL_SHEET)
                regression_sheet = get_sheet(wb, REGRESSION_MODEL_SHEET)

                if empirical_sheet is None:
                    print(f"Skipped empirical extraction in {file_path.name}: sheet not found.")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                if regression_sheet is None:
                    print(f"Skipped regression extraction in {file_path.name}: sheet not found.")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error: {exc}")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
