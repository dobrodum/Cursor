#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these paths before running.
input_dir = "/path/to/input"
output_dir = "/path/to/output"

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

PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", flags=re.IGNORECASE)


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        percent = text.endswith("%")
        cleaned = text.replace(",", "").replace("%", "")
        try:
            parsed = float(cleaned)
        except ValueError:
            return None
        return parsed / 100.0 if percent else parsed
    return None


def safe_subtract(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left - right


def parse_month(month_token: str) -> tuple[int, str]:
    token = re.sub(r"[^A-Za-z]", "", month_token).lower()
    if not token:
        raise ValueError(f"Could not parse month token '{month_token}'.")

    full_to_num = {name.lower(): i for i, name in enumerate(calendar.month_name) if name}
    abbr_to_num = {name.lower(): i for i, name in enumerate(calendar.month_abbr) if name}
    if token in full_to_num:
        month_num = full_to_num[token]
    elif token in abbr_to_num:
        month_num = abbr_to_num[token]
    elif token[:3] in abbr_to_num:
        month_num = abbr_to_num[token[:3]]
    else:
        raise ValueError(f"Unsupported month token '{month_token}'.")
    return month_num, calendar.month_abbr[month_num]


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem

    dash_parts = [part.strip() for part in stem.replace(" - ", "-").split("-") if part.strip()]
    ticker = ""
    if len(dash_parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", dash_parts[1]).upper()

    period_match = PERIOD_PATTERN.search(stem)
    if not period_match:
        fallback_model = ticker or stem
        return FileLabels(
            model=fallback_model,
            ticker=ticker,
            model_period="",
            model_date="",
        )

    period_word_raw, month_token, year_str = period_match.groups()
    period_word = period_word_raw.capitalize()
    month_num, month_abbr = parse_month(month_token)
    day = PERIOD_DAY[period_word]
    model_period = f"{period_word}{month_abbr}_{year_str}"
    model_date = date(int(year_str), month_num, day).isoformat()
    model = f"{ticker}_{model_period}" if ticker else model_period

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(source_dir: Path, target_dir: Path) -> Path:
    folder_name = source_dir.resolve().name
    base_name = f"{folder_name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def try_get_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def close_source_workbook(wb: xw.Book) -> None:
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

    try:
        wb.close()
    except Exception:
        pass


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    cell.api.Formula2R1C1 = formula_r1c1


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def pick_first_numeric(sheet: xw.Sheet, addresses: list[tuple[int, int]]) -> float | None:
    for row, col in addresses:
        value = to_number(read_cell(sheet, row, col))
        if value is not None:
            return value
    return None


def used_bounds(sheet: xw.Sheet) -> tuple[int, int, int, int]:
    used = sheet.used_range
    first_row = int(used.row)
    first_col = int(used.column)
    last_row = first_row + int(used.rows.count) - 1
    last_col = first_col + int(used.columns.count) - 1
    return first_row, first_col, last_row, last_col


def find_anchor(sheet: xw.Sheet, keyword: str = "max") -> tuple[int, int] | None:
    try:
        found = sheet.api.Cells.Find(
            What=keyword,
            After=sheet.api.Cells(1, 1),
            LookIn=-4163,  # xlValues
            LookAt=1,  # xlWhole
            SearchOrder=1,  # xlByRows
            SearchDirection=1,  # xlNext
            MatchCase=False,
        )
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    try:
        found = sheet.api.Cells.Find(What=keyword)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    try:
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
        for r_idx, row_values in enumerate(matrix):
            for c_idx, value in enumerate(row_values):
                if isinstance(value, str) and value.strip().lower() == keyword.lower():
                    return int(used.row + r_idx), int(used.column + c_idx)
    except Exception:
        return None
    return None


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, keyword="max")
    if anchor is None:
        print(f"Skipped empirical in {source_file}: could not find 'max' anchor.")
        return []

    anchor_row, anchor_col = anchor
    _, _, last_used_row, last_used_col = used_bounds(sheet)

    data_end_row = anchor_row - 1
    data_start_row = max(1, data_end_row - (N_QUARTERS - 1))
    history_len = data_end_row - data_start_row + 1
    if history_len <= 0:
        print(f"Skipped empirical in {source_file}: insufficient data around anchor.")
        return []

    quarter_label_col = anchor_col - 13
    penetration_col = anchor_col - 9
    quarterly_sales_col = anchor_col - 8
    reported_sales_col = anchor_col - 7
    growth_rate_col = anchor_col - 6
    sales_captured_col = anchor_col - 5
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1

    scratch_row = last_used_row + 3
    scratch_col = last_used_col + 2

    n_cell = sheet.cells(scratch_row, scratch_col)
    avg_pen_cell = sheet.cells(scratch_row, scratch_col + 1)
    forecast_cell = sheet.cells(scratch_row, scratch_col + 2)

    set_formula2_r1c1(
        avg_pen_cell,
        (
            "=LET("
            f"n,MAX(1,MIN(R{scratch_row}C{scratch_col},{history_len})),"
            f"rng,R{data_start_row}C{penetration_col}:R{data_end_row}C{penetration_col},"
            "AVERAGE(INDEX(rng,ROWS(rng)-n+1):INDEX(rng,ROWS(rng)))"
            ")"
        ),
    )
    set_formula2_r1c1(
        forecast_cell,
        f'=IFERROR(R{scratch_row}C{scratch_col + 1}*R{data_end_row}C{quarterly_sales_col},"")',
    )

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        n_cell.value = n_quarters
        wb.app.calculate()

        lookback = min(n_quarters, history_len)
        last_quarter_row = data_end_row - lookback + 1

        avg_penetration = to_number(avg_pen_cell.value)
        quarterly_sales = to_number(read_cell(sheet, data_end_row, quarterly_sales_col))
        reported_sales = to_number(read_cell(sheet, data_end_row, reported_sales_col))
        growth_rate = to_number(read_cell(sheet, data_end_row, growth_rate_col))
        sales_captured = to_number(read_cell(sheet, data_end_row, sales_captured_col))
        forecast_value = to_number(forecast_cell.value)

        forecast_max = pick_first_numeric(
            sheet,
            [
                (anchor_row + n_quarters, forecast_max_col),
                (anchor_row + 1, forecast_max_col),
                (data_end_row, forecast_max_col),
            ],
        )
        forecast_min = pick_first_numeric(
            sheet,
            [
                (anchor_row + n_quarters, forecast_min_col),
                (anchor_row + 1, forecast_min_col),
                (data_end_row, forecast_min_col),
            ],
        )

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * 1.1
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value * 0.9

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n_quarters,
            "last_quarter_used": read_cell(sheet, last_quarter_row, quarter_label_col),
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": source_file,
        }
        rows.append(row)

    sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 2)).clear_contents()
    return rows


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, keyword="max")
    if anchor is None:
        print(f"Skipped regression in {source_file}: could not find 'max' anchor.")
        return []

    anchor_row, anchor_col = anchor
    _, _, last_used_row, last_used_col = used_bounds(sheet)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1

    data_end_row = anchor_row - 1
    data_start_row = max(1, data_end_row - (N_QUARTERS - 1))
    history_len = data_end_row - data_start_row + 1
    if history_len < 2:
        print(f"Skipped regression in {source_file}: fewer than 2 rows for regression.")
        return []

    scratch_row = last_used_row + 6
    scratch_col = last_used_col + 2

    n_cell = sheet.cells(scratch_row, scratch_col)
    intercept_cell = sheet.cells(scratch_row, scratch_col + 1)
    slope_cell = sheet.cells(scratch_row, scratch_col + 2)
    forecast_cell = sheet.cells(scratch_row, scratch_col + 3)

    set_formula2_r1c1(
        intercept_cell,
        (
            "=LET("
            f"n,MAX(2,MIN(R{scratch_row}C{scratch_col},{history_len})),"
            f"x,R{data_start_row}C{x_col}:R{data_end_row}C{x_col},"
            f"y,R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            "INTERCEPT("
            "INDEX(y,ROWS(y)-n+1):INDEX(y,ROWS(y)),"
            "INDEX(x,ROWS(x)-n+1):INDEX(x,ROWS(x))"
            ")"
            ")"
        ),
    )
    set_formula2_r1c1(
        slope_cell,
        (
            "=LET("
            f"n,MAX(2,MIN(R{scratch_row}C{scratch_col},{history_len})),"
            f"x,R{data_start_row}C{x_col}:R{data_end_row}C{x_col},"
            f"y,R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            "SLOPE("
            "INDEX(y,ROWS(y)-n+1):INDEX(y,ROWS(y)),"
            "INDEX(x,ROWS(x)-n+1):INDEX(x,ROWS(x))"
            ")"
            ")"
        ),
    )
    set_formula2_r1c1(
        forecast_cell,
        (
            f'=IFERROR(R{scratch_row}C{scratch_col + 1}+'
            f"R{scratch_row}C{scratch_col + 2}*R{data_end_row + 1}C{x_col},\"\")"
        ),
    )

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters in range(2, N_QUARTERS + 1):
        n_cell.value = n_quarters
        wb.app.calculate()

        intercept = to_number(intercept_cell.value)
        slope = to_number(slope_cell.value)
        forecast_value = to_number(forecast_cell.value)
        actual_value = to_number(read_cell(sheet, data_end_row + 1, y_col))

        forecast_max = pick_first_numeric(
            sheet,
            [
                (anchor_row + n_quarters, forecast_max_col),
                (anchor_row + 1, forecast_max_col),
                (data_end_row, forecast_max_col),
            ],
        )
        forecast_min = pick_first_numeric(
            sheet,
            [
                (anchor_row + n_quarters, forecast_min_col),
                (anchor_row + 1, forecast_min_col),
                (data_end_row, forecast_min_col),
            ],
        )

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * 1.1
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value * 0.9

        signature = (
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 3)).clear_contents()
    return rows


def apply_output_format(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[letter].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    emp_ws = out_wb.create_sheet("empirical_candidates")
    emp_ws.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        emp_ws.append([row.get(column) for column in EMPIRICAL_HEADERS])
    apply_output_format(emp_ws)

    reg_ws = out_wb.create_sheet("regression_candidates")
    reg_ws.append(REGRESSION_HEADERS)
    for row in regression_rows:
        reg_ws.append([row.get(column) for column in REGRESSION_HEADERS])
    apply_output_format(reg_ws)

    out_wb.save(output_path)


def process_all_files(source_dir: Path, target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(source_dir, target_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(source_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb = None
            try:
                labels = parse_file_labels(file_path)
                wb = app.books.open(str(file_path), update_links=False)

                empirical_sheet = try_get_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical in {file_path.name}: sheet 'Empirical Model' not found.")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = try_get_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression in {file_path.name}: sheet 'Regression Model' not found.")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        try:
            app.display_alerts = True
            app.screen_updating = True
        except Exception:
            pass
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    target_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {source_dir}")
    if not source_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {source_dir}")

    process_all_files(source_dir, target_dir)


if __name__ == "__main__":
    main()
