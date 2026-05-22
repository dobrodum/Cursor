from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------- user-configurable paths ----------
input_dir = Path(r"/path/to/input")
output_dir = Path(r"/path/to/output")
# -------------------------------------------

N_QUARTERS = 10

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


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(value: Any) -> list[list[Any]]:
    if isinstance(value, list):
        if value and isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def parse_month(month_token: str) -> tuple[int | None, str]:
    month_map = {
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
    token = month_token.strip().lower()
    month_num = month_map.get(token) or month_map.get(token[:3])
    if month_num is None:
        return None, month_token.title()
    month_abbr = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ][month_num - 1]
    return month_num, month_abbr


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""
    period_part = parts[2] if len(parts) >= 3 else ""
    period_part = period_part.split("_")[0].strip()

    model_period = ""
    model_date = ""
    match = re.match(r"(?i)^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_part)
    if match:
        phase = match.group(1).title()
        month_token = match.group(2)
        year = int(match.group(3))
        month_num, month_abbr = parse_month(month_token)
        model_period = f"{phase}{month_abbr}_{year}"
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        if month_num is not None:
            model_date = date(year, month_num, day_lookup[phase]).isoformat()
    if not model_period:
        model_period = period_part or "unknown_period"

    model = f"{ticker}_{model_period}" if ticker else model_period
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def to_float(value: Any) -> float | None:
    if value in ("", None):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    if value in ("", None):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def find_anchor(sheet: xw.Sheet, label: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = normalize_2d(used.value)
    target = label.strip().lower()
    for r_idx, row in enumerate(values):
        for c_idx, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip().lower() == target:
                return used.row + r_idx, used.column + c_idx
    return None


def contiguous_non_blank_rows(
    sheet: xw.Sheet,
    start_row: int,
    col_idx: int,
    max_scan: int = 400,
) -> int:
    values = sheet.range((start_row, col_idx), (start_row + max_scan - 1, col_idx)).value
    values = values if isinstance(values, list) else [values]
    count = 0
    for val in values:
        if val in ("", None):
            break
        count += 1
    return count


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
    except TypeError:
        try:
            wb.api.Close(SaveChanges=False)
        except Exception:
            wb.close()
    except Exception:
        try:
            wb.api.Close(SaveChanges=False)
        except Exception:
            pass


def output_path_for_run(out_dir: Path, input_folder_name: str) -> Path:
    base = out_dir / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base
    idx = 1
    while True:
        candidate = out_dir / f"{input_folder_name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def process_empirical_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"skipped empirical model in {source_file}: missing 'Empirical Model' sheet")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped empirical model in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1

    # Anchor-based offsets from the "max" column.
    num_quarters_col = anchor_col - 11
    last_quarter_col = anchor_col - 10
    quarterly_sales_col = anchor_col - 8
    growth_rate_col = anchor_col - 6
    sales_captured_col = anchor_col - 5
    forecast_total_col = anchor_col - 3
    reported_sales_col = anchor_col - 2
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1

    available_rows = contiguous_non_blank_rows(sheet, data_start_row, forecast_max_col)
    loop_count = min(N_QUARTERS, max(available_rows, 0))
    if loop_count == 0:
        return []

    avg_pen_calc_col = anchor_col + 6
    forecast_calc_col = anchor_col + 7

    for idx in range(loop_count):
        row = data_start_row + idx
        start_row = data_start_row
        # Use formula2/R1C1 for rolling average penetration; values are read after one calculate call.
        sheet.range((row, avg_pen_calc_col)).formula2 = (
            f'=IFERROR(AVERAGE(R{start_row}C{sales_captured_col}:R{row}C{sales_captured_col}),"")'
        )
        sheet.range((row, forecast_calc_col)).formula2 = (
            f'=IFERROR(R{row}C{quarterly_sales_col}/R{row}C{avg_pen_calc_col},"")'
        )

    wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for idx in range(loop_count):
        row = data_start_row + idx
        forecast_max = to_float(sheet.range((row, forecast_max_col)).value)
        forecast_min = to_float(sheet.range((row, forecast_min_col)).value)
        if forecast_max is None and forecast_min is None:
            continue

        num_quarters_used = to_int(sheet.range((row, num_quarters_col)).value) or (idx + 1)
        avg_penetration_pct = to_float(sheet.range((row, avg_pen_calc_col)).value)
        forecast_value = to_float(sheet.range((row, forecast_total_col)).value)
        if forecast_value is None:
            forecast_value = to_float(sheet.range((row, forecast_calc_col)).value)

        actual_value = to_float(sheet.range((row, reported_sales_col)).value)
        quarterly_sales = to_float(sheet.range((row, quarterly_sales_col)).value)
        reported_sales = to_float(sheet.range((row, reported_sales_col)).value)
        growth_rate_pct = to_float(sheet.range((row, growth_rate_col)).value)
        sales_captured_in_db_pct = to_float(sheet.range((row, sales_captured_col)).value)
        last_quarter_used = sheet.range((row, last_quarter_col)).value

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": (
                    (forecast_max - forecast_min)
                    if forecast_max is not None and forecast_min is not None
                    else None
                ),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"skipped regression model in {source_file}: missing 'Regression Model' sheet")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped regression model in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1

    # Required anchor-based source columns.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = x_col
    forecast_total_col = anchor_col - 3  # TOT FCST w/o SA
    actual_value_col = anchor_col - 2
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1
    intercept_col = anchor_col + 4
    slope_col = anchor_col + 5

    available_rows = contiguous_non_blank_rows(sheet, data_start_row, y_col)
    loop_count = min(N_QUARTERS, max(available_rows, 0))
    if loop_count == 0:
        return []

    for idx in range(loop_count):
        row = data_start_row + idx
        start_row = data_start_row
        sheet.range((row, intercept_col)).formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        sheet.range((row, slope_col)).formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )

    wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for idx in range(loop_count):
        row = data_start_row + idx
        num_quarters_used = to_int(sheet.range((row, num_quarters_col)).value) or (idx + 1)
        forecast_value = to_float(sheet.range((row, forecast_total_col)).value)
        actual_value = to_float(sheet.range((row, actual_value_col)).value)
        forecast_max = to_float(sheet.range((row, forecast_max_col)).value)
        forecast_min = to_float(sheet.range((row, forecast_min_col)).value)
        intercept = to_float(sheet.range((row, intercept_col)).value)
        slope = to_float(sheet.range((row, slope_col)).value)

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        signature = (
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )
        if idx == loop_count - 1 and previous_signature == signature:
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
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": (
                    (forecast_max - forecast_min)
                    if forecast_max is not None and forecast_min is not None
                    else None
                ),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def fit_column_widths(worksheet) -> None:
    for col_idx in range(1, worksheet.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[letter].width = min(max(12, max_len + 2), 48)


def write_sheet(workbook: Workbook, name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws = workbook.create_sheet(name)
    ws.append(headers)
    for record in rows:
        ws.append([record.get(col, "") if record.get(col, None) is not None else "" for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    fit_column_widths(ws)


def run() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_path_for_run(output_dir, input_dir.name)

    files_to_process: list[Path] = []
    for file_path in sorted(input_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped {file_path.name}: not an .xlsx file")
            continue
        files_to_process.append(file_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in files_to_process:
            print(f"processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path.name)

                empirical_rows.extend(process_empirical_sheet(wb, label, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, label, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    out_wb.save(out_path)

    print(f"output path: {out_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
