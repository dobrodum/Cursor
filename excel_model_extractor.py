#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable paths
# ----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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

# Anchor-based offsets used by legacy models.
EMPIRICAL_OFFSETS = {
    "quarter_label_col": -12,
    "quarterly_sales_col": -11,
    "reported_sales_col": -10,
    "penetration_pct_col": -7,
    "growth_rate_pct_col": -5,
    "sales_captured_col": -4,
    "forecast_max_col": 1,
    "forecast_min_col": 2,
}

MONTH_TO_NUMBER = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

PERIOD_PATTERN = re.compile(
    r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
    flags=re.IGNORECASE,
)

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not (
        isinstance(value, float) and math.isnan(value)
    )


def to_float(value: Any) -> float | None:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text.endswith("%"):
            text = text[:-1]
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def clean_output_value(value: Any) -> Any:
    if value is None:
        return ""
    return value


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    period_token = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    if not ticker:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = PERIOD_PATTERN.search(period_token) or PERIOD_PATTERN.search(stem)
    if not period_match:
        model_period = ""
        model_date = ""
    else:
        period_word, month_word, year = period_match.groups()
        period_word = period_word.title()
        month_key = month_word[:3].title()
        month_number = MONTH_TO_NUMBER.get(month_key)
        if month_number is None:
            model_period = ""
            model_date = ""
        else:
            model_period = f"{period_word}{month_key}_{year}"
            day = PERIOD_DAY[period_word.lower()]
            model_date = date(int(year), month_number, day).isoformat()

    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    else:
        model = stem

    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def make_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_folder / f"{base}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values_2d = normalize_2d(used.value)
    if not values_2d:
        return None

    base_row = used.row
    base_col = used.column
    target = anchor_text.lower()

    for row_index, row in enumerate(values_2d):
        for col_index, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return base_row + row_index, base_col + col_index
    return None


def first_numeric(values: Iterable[Any]) -> float | None:
    for value in values:
        number = to_float(value)
        if number is not None:
            return number
    return None


def collect_data_rows(
    sheet: xw.Sheet,
    start_row: int,
    required_cols: list[int],
    max_lookback: int = 200,
) -> list[int]:
    valid_cols = [col for col in required_cols if col > 0]
    if not valid_cols:
        return []

    rows: list[int] = []
    row = max(1, start_row)
    empty_streak = 0
    scanned = 0

    while row >= 1 and scanned < max_lookback:
        values = [to_float(sheet.cells(row, col).value) for col in valid_cols]
        if all(value is not None for value in values):
            rows.append(row)
            empty_streak = 0
        elif rows:
            empty_streak += 1
            if empty_streak >= 2:
                break
        row -= 1
        scanned += 1

    rows.reverse()
    return rows


def close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception as exc:
        print(f"warning: could not close workbook cleanly: {exc}")


def get_sheet(wb: xw.Book, name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[name]
    except Exception:
        return None


def extract_empirical_rows(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped Empirical Model in {source_file}: sheet not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped Empirical Model in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    quarter_label_col = max(1, anchor_col + EMPIRICAL_OFFSETS["quarter_label_col"])
    quarterly_sales_col = max(1, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"])
    reported_sales_col = max(1, anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"])
    penetration_col = max(1, anchor_col + EMPIRICAL_OFFSETS["penetration_pct_col"])
    growth_rate_col = max(1, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct_col"])
    sales_captured_col = max(1, anchor_col + EMPIRICAL_OFFSETS["sales_captured_col"])
    forecast_max_col = max(1, anchor_col + EMPIRICAL_OFFSETS["forecast_max_col"])
    forecast_min_col = max(1, anchor_col + EMPIRICAL_OFFSETS["forecast_min_col"])

    data_rows = collect_data_rows(
        sheet=sheet,
        start_row=anchor_row - 1,
        required_cols=[penetration_col],
    )
    use_penetration_column = True

    if not data_rows:
        data_rows = collect_data_rows(
            sheet=sheet,
            start_row=anchor_row - 1,
            required_cols=[quarterly_sales_col, reported_sales_col],
        )
        use_penetration_column = False

    if not data_rows:
        print(f"Skipped Empirical Model in {source_file}: no quarter history found")
        return []

    data_rows = data_rows[-N_QUARTERS:]
    quarter_count = min(N_QUARTERS, len(data_rows))

    helper_col = sheet.used_range.last_cell.column + 3
    helper_start_row = anchor_row + 2
    calc_specs: list[dict[str, int]] = []

    for idx, n_used in enumerate(range(1, quarter_count + 1)):
        selected_rows = data_rows[-n_used:]
        start = selected_rows[0]
        end = selected_rows[-1]
        helper_row = helper_start_row + idx

        avg_cell = sheet.cells(helper_row, helper_col)
        forecast_cell = sheet.cells(helper_row, helper_col + 1)

        if use_penetration_column:
            avg_cell.formula2 = (
                f'=IFERROR(AVERAGE(R{start}C{penetration_col}:R{end}C{penetration_col}),"")'
            )
        else:
            avg_cell.formula2 = (
                f'=IFERROR(SUM(R{start}C{quarterly_sales_col}:R{end}C{quarterly_sales_col})/'
                f'SUM(R{start}C{reported_sales_col}:R{end}C{reported_sales_col}),"")'
            )

        forecast_cell.formula2 = (
            f'=IFERROR(R{end}C{reported_sales_col}/R{helper_row}C{helper_col},"")'
        )

        calc_specs.append({"n_used": n_used, "end_row": end, "helper_row": helper_row})

    if calc_specs:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    table_start_row = anchor_row - quarter_count + 1

    for idx, spec in enumerate(calc_specs):
        n_used = spec["n_used"]
        end_row = spec["end_row"]
        helper_row = spec["helper_row"]
        table_row = table_start_row + idx

        avg_penetration = to_float(sheet.cells(helper_row, helper_col).value)
        forecast_value = to_float(sheet.cells(helper_row, helper_col + 1).value)

        reported_sales = to_float(sheet.cells(end_row, reported_sales_col).value)
        quarterly_sales = to_float(sheet.cells(end_row, quarterly_sales_col).value)
        growth_rate_pct = to_float(sheet.cells(end_row, growth_rate_col).value)
        sales_captured_pct = to_float(sheet.cells(end_row, sales_captured_col).value)
        if sales_captured_pct is None:
            sales_captured_pct = avg_penetration

        forecast_max = first_numeric(
            [
                sheet.cells(table_row, forecast_max_col).value,
                sheet.cells(anchor_row, forecast_max_col).value,
            ]
        )
        forecast_min = first_numeric(
            [
                sheet.cells(table_row, forecast_min_col).value,
                sheet.cells(anchor_row, forecast_min_col).value,
                sheet.cells(anchor_row + 1, forecast_max_col).value,
            ]
        )

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        last_quarter_used = sheet.cells(end_row, quarter_label_col).value
        if hasattr(last_quarter_used, "strftime"):
            last_quarter_used = last_quarter_used.strftime("%Y-%m-%d")

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_used,
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
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    # Clear temporary helper formulas so the workbook state stays clean while open.
    if calc_specs:
        clear_bottom_row = helper_start_row + len(calc_specs) - 1
        sheet.range((helper_start_row, helper_col), (clear_bottom_row, helper_col + 1)).value = None

    return rows


def rows_match_for_duplicate_check(
    previous_row: dict[str, Any],
    current_row: dict[str, Any],
) -> bool:
    keys = ["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    tolerance = 1e-10

    for key in keys:
        previous = to_float(previous_row.get(key))
        current = to_float(current_row.get(key))
        if previous is None and current is None:
            continue
        if previous is None or current is None:
            return False
        if abs(previous - current) > tolerance:
            return False
    return True


def extract_regression_rows(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped Regression Model in {source_file}: sheet not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped Regression Model in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print(f"Skipped Regression Model in {source_file}: invalid anchor offsets")
        return []

    data_rows = collect_data_rows(
        sheet=sheet,
        start_row=anchor_row - 1,
        required_cols=[x_col, y_col],
    )
    if not data_rows:
        print(f"Skipped Regression Model in {source_file}: no x/y history found")
        return []

    data_rows = data_rows[-N_QUARTERS:]
    quarter_count = min(N_QUARTERS, len(data_rows))

    helper_col = sheet.used_range.last_cell.column + 3
    helper_start_row = anchor_row + 2
    calc_specs: list[dict[str, int]] = []

    for idx, n_used in enumerate(range(1, quarter_count + 1)):
        selected_rows = data_rows[-n_used:]
        start = selected_rows[0]
        end = selected_rows[-1]
        helper_row = helper_start_row + idx

        intercept_cell = sheet.cells(helper_row, helper_col)
        slope_cell = sheet.cells(helper_row, helper_col + 1)
        forecast_cell = sheet.cells(helper_row, helper_col + 2)

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{start}C{y_col}:R{end}C{y_col},'
            f'R{start}C{x_col}:R{end}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{start}C{y_col}:R{end}C{y_col},'
            f'R{start}C{x_col}:R{end}C{x_col}),"")'
        )
        forecast_cell.formula2 = (
            f'=IFERROR(R{helper_row}C{helper_col}+R{helper_row}C{helper_col + 1}*'
            f'(R{end}C{x_col}+1),"")'
        )

        calc_specs.append({"n_used": n_used, "end_row": end, "helper_row": helper_row})

    if calc_specs:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    table_start_row = anchor_row - quarter_count + 1

    for idx, spec in enumerate(calc_specs):
        n_used = spec["n_used"]
        end_row = spec["end_row"]
        helper_row = spec["helper_row"]
        table_row = table_start_row + idx

        intercept = to_float(sheet.cells(helper_row, helper_col).value)
        slope = to_float(sheet.cells(helper_row, helper_col + 1).value)
        forecast_total_without_sa = to_float(sheet.cells(helper_row, helper_col + 2).value)

        forecast_max = first_numeric(
            [
                sheet.cells(table_row, anchor_col + 1).value,
                sheet.cells(anchor_row, anchor_col + 1).value,
            ]
        )
        forecast_min = first_numeric(
            [
                sheet.cells(table_row, anchor_col + 2).value,
                sheet.cells(anchor_row, anchor_col + 2).value,
                sheet.cells(anchor_row + 1, anchor_col + 1).value,
            ]
        )

        actual_value = first_numeric(
            [
                sheet.cells(table_row, y_col).value,
                sheet.cells(end_row, y_col).value,
            ]
        )

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_used,
            "num_quarters_used": n_used,
            "forecast_value": forecast_total_without_sa,
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        # The source models sometimes repeat the final candidate row;
        # skip exact duplicates to keep output clean.
        if rows and rows_match_for_duplicate_check(rows[-1], row):
            continue
        rows.append(row)

    if calc_specs:
        clear_bottom_row = helper_start_row + len(calc_specs) - 1
        sheet.range((helper_start_row, helper_col), (clear_bottom_row, helper_col + 2)).value = None

    return rows


def write_candidate_sheet(
    wb: Workbook,
    title: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=title)
    ws.append(columns)
    for row in rows:
        ws.append([clean_output_value(row.get(column, "")) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 14), 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_candidate_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_candidate_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def process_workbooks() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_dir}")

    output_path = make_output_path(input_dir, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(
                    extract_empirical_rows(wb=wb, metadata=metadata, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_rows(wb=wb, metadata=metadata, source_file=file_path.name)
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
