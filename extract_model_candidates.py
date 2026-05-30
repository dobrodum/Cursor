#!/usr/bin/env python3
"""Extract empirical and regression model candidates from .xlsx files."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-editable paths
input_dir = Path("./input")
output_dir = Path("./output")

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

# Fallback offsets are relative to the "max" anchor column.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -5,
    "quarterly_sales": -9,
    "reported_sales": -8,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -4,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": 2,
    "slope": 3,
}

EMPIRICAL_HEADER_ALIASES = {
    "num_quarters_used": ["num quarters", "quarters used", "n quarters", "n_qtrs"],
    "last_quarter_used": ["last quarter", "last qtr"],
    "forecast_value": [
        "estimated total sold",
        "total sold",
        "forecast value",
        "total forecast",
        "tot fcst",
    ],
    "actual_value": ["actual value", "actual sales", "reported sales"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "avg_penetration_pct": ["avg penetration", "average penetration"],
    "quarterly_sales": ["quarterly sales", "q sales"],
    "reported_sales": ["reported sales"],
    "growth_rate_pct": ["growth rate"],
    "sales_captured_in_db_pct": ["sales captured in db", "captured in db", "penetration"],
}

REGRESSION_HEADER_ALIASES = {
    "num_quarters_used": ["num quarters", "quarters used", "n quarters", "n_qtrs"],
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "forecast total without sa",
        "total forecast without sa",
    ],
    "actual_value": ["actual value", "actual sales", "reported sales"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "intercept": ["intercept"],
    "slope": ["slope"],
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def coerce_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, tuple):
        values = list(values)

    if not isinstance(values, list):
        return [[values]]

    if not values:
        return [[]]

    first = values[0]
    if isinstance(first, tuple):
        values = [list(row) for row in values]
        first = values[0] if values else []

    if isinstance(first, list):
        return values

    return [values]


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def to_int(value: Any) -> int | None:
    if is_blank(value):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def to_float(value: Any) -> float | None:
    if is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used_range = sheet.used_range
    values = coerce_2d(used_range.value)
    if not values:
        return None

    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return used_range.row + row_idx, used_range.column + col_idx
    return None


def discover_offsets(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    aliases: dict[str, list[str]],
    fallback_offsets: dict[str, int],
    search_left: int = 30,
    search_right: int = 10,
) -> dict[str, int]:
    start_row = max(1, anchor_row - 1)
    end_row = anchor_row + 1
    start_col = max(1, anchor_col - search_left)
    end_col = anchor_col + search_right

    grid = coerce_2d(sheet.range((start_row, start_col), (end_row, end_col)).value)
    offsets: dict[str, int] = {}

    for field_name, alias_list in aliases.items():
        best_col = None
        best_score = None
        for r_idx, row_values in enumerate(grid):
            for c_idx, raw_value in enumerate(row_values):
                normalized = normalize_text(raw_value)
                if not normalized:
                    continue
                if any(alias in normalized for alias in alias_list):
                    col = start_col + c_idx
                    row = start_row + r_idx
                    score = abs(col - anchor_col) * 5 + abs(row - anchor_row)
                    if best_score is None or score < best_score:
                        best_score = score
                        best_col = col
        if best_col is not None:
            offsets[field_name] = best_col - anchor_col

    for field_name, fallback in fallback_offsets.items():
        offsets.setdefault(field_name, fallback)

    offsets["forecast_max"] = 0
    return offsets


def get_offset_value(sheet: xw.Sheet, row: int, anchor_col: int, offset: int | None) -> Any:
    if offset is None:
        return None
    col = anchor_col + offset
    if col < 1 or row < 1:
        return None
    return sheet.cells(row, col).value


def parse_file_label(file_path: Path) -> dict[str, str] | None:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[-2]
    period_token = re.sub(r"_send$", "", parts[-1], flags=re.IGNORECASE)

    period_match = re.match(
        r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if not period_match:
        return None

    window = period_match.group(1).title()
    month_abbrev = period_match.group(2).title()
    year = int(period_match.group(3))

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_map[window]

    month_lookup = {
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
    month = month_lookup.get(month_abbrev)
    if month is None:
        return None

    model_period = f"{window}{month_abbrev}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def calculate_empirical_avg_penetration(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    n_quarters: int,
    penetration_col: int,
) -> Any:
    end_row = anchor_row - 1
    start_row = end_row - n_quarters + 1
    if start_row < 1 or penetration_col < 1:
        return None

    helper_row = anchor_row + n_quarters
    helper_col = anchor_col + 8
    helper_cell = sheet.cells(helper_row, helper_col)
    formula = (
        f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col}),"")'
    )

    helper_cell.formula2 = formula
    sheet.book.app.calculate()
    value = helper_cell.value
    helper_cell.value = None
    return value


def calculate_regression_coefficients(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    n_quarters: int,
) -> tuple[Any, Any]:
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    end_row = anchor_row - 1
    start_row = end_row - n_quarters + 1
    if start_row < 1 or x_col < 1 or y_col < 1:
        return None, None

    helper_row = anchor_row + n_quarters
    intercept_cell = sheet.cells(helper_row, anchor_col + 8)
    slope_cell = sheet.cells(helper_row, anchor_col + 9)

    intercept_formula = (
        f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
        f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
    )
    slope_formula = (
        f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
        f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
    )

    intercept_cell.formula2 = intercept_formula
    slope_cell.formula2 = slope_formula
    sheet.book.app.calculate()
    intercept = intercept_cell.value
    slope = slope_cell.value
    intercept_cell.value = None
    slope_cell.value = None
    return intercept, slope


def calc_range_width(forecast_max: Any, forecast_min: Any) -> float | None:
    max_value = to_float(forecast_max)
    min_value = to_float(forecast_min)
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def process_empirical_sheet(
    sheet: xw.Sheet,
    model_meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  - skipped empirical extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    offsets = discover_offsets(
        sheet,
        anchor_row,
        anchor_col,
        EMPIRICAL_HEADER_ALIASES,
        EMPIRICAL_FALLBACK_OFFSETS,
    )

    rows: list[dict[str, Any]] = []
    empty_streak = 0

    for idx in range(1, N_QUARTERS + 1):
        row = anchor_row + idx
        num_quarters_used = to_int(get_offset_value(sheet, row, anchor_col, offsets["num_quarters_used"]))
        if num_quarters_used is None:
            num_quarters_used = idx

        last_quarter_used = get_offset_value(sheet, row, anchor_col, offsets.get("last_quarter_used"))
        forecast_value = get_offset_value(sheet, row, anchor_col, offsets.get("forecast_value"))
        actual_value = get_offset_value(sheet, row, anchor_col, offsets.get("actual_value"))
        forecast_max = get_offset_value(sheet, row, anchor_col, offsets.get("forecast_max"))
        forecast_min = get_offset_value(sheet, row, anchor_col, offsets.get("forecast_min"))
        avg_penetration_pct = get_offset_value(sheet, row, anchor_col, offsets.get("avg_penetration_pct"))
        quarterly_sales = get_offset_value(sheet, row, anchor_col, offsets.get("quarterly_sales"))
        reported_sales = get_offset_value(sheet, row, anchor_col, offsets.get("reported_sales"))
        growth_rate_pct = get_offset_value(sheet, row, anchor_col, offsets.get("growth_rate_pct"))
        sales_captured_in_db_pct = get_offset_value(
            sheet, row, anchor_col, offsets.get("sales_captured_in_db_pct")
        )

        if is_blank(avg_penetration_pct):
            penetration_col = anchor_col + offsets.get("sales_captured_in_db_pct", 0)
            avg_penetration_pct = calculate_empirical_avg_penetration(
                sheet,
                anchor_row,
                anchor_col,
                num_quarters_used,
                penetration_col,
            )

        if all(
            is_blank(value)
            for value in [forecast_value, actual_value, forecast_max, forecast_min, avg_penetration_pct]
        ):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        row_out = {
            "model": model_meta["model"],
            "ticker": model_meta["ticker"],
            "model_period": model_meta["model_period"],
            "model_date": model_meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row_out)

    return rows


def normalize_compare_value(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def is_duplicate_regression_row(previous: dict[str, Any], current: dict[str, Any]) -> bool:
    key_fields = ["num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for field in key_fields:
        if normalize_compare_value(previous.get(field)) != normalize_compare_value(current.get(field)):
            return False
    return True


def process_regression_sheet(
    sheet: xw.Sheet,
    model_meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print("  - skipped regression extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    offsets = discover_offsets(
        sheet,
        anchor_row,
        anchor_col,
        REGRESSION_HEADER_ALIASES,
        REGRESSION_FALLBACK_OFFSETS,
    )

    rows: list[dict[str, Any]] = []
    empty_streak = 0

    for idx in range(1, N_QUARTERS + 1):
        row = anchor_row + idx
        num_quarters_used = to_int(get_offset_value(sheet, row, anchor_col, offsets["num_quarters_used"]))
        if num_quarters_used is None:
            num_quarters_used = idx

        forecast_value = get_offset_value(sheet, row, anchor_col, offsets.get("forecast_value"))
        actual_value = get_offset_value(sheet, row, anchor_col, offsets.get("actual_value"))
        forecast_max = get_offset_value(sheet, row, anchor_col, offsets.get("forecast_max"))
        forecast_min = get_offset_value(sheet, row, anchor_col, offsets.get("forecast_min"))
        intercept = get_offset_value(sheet, row, anchor_col, offsets.get("intercept"))
        slope = get_offset_value(sheet, row, anchor_col, offsets.get("slope"))

        if is_blank(intercept) or is_blank(slope):
            intercept_calc, slope_calc = calculate_regression_coefficients(
                sheet,
                anchor_row,
                anchor_col,
                num_quarters_used,
            )
            if is_blank(intercept):
                intercept = intercept_calc
            if is_blank(slope):
                slope = slope_calc

        if all(is_blank(value) for value in [forecast_value, forecast_max, forecast_min, intercept, slope]):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        row_out = {
            "model": model_meta["model"],
            "ticker": model_meta["ticker"],
            "model_period": model_meta["model_period"],
            "model_date": model_meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(rows[-1], row_out):
            continue

        rows.append(row_out)

    return rows


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def get_sheet(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def next_output_path(source_input_dir: Path, destination_dir: Path) -> Path:
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = destination_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = destination_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb_out = Workbook()
    default_ws = wb_out.active
    wb_out.remove(default_ws)

    for sheet_name, columns, rows in [
        ("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows),
        ("regression_candidates", REGRESSION_COLUMNS, regression_rows),
    ]:
        ws = wb_out.create_sheet(sheet_name)
        ws.append(columns)
        for cell in ws[1]:
            cell.font = Font(bold=True)

        for row_data in rows:
            ws.append([row_data.get(column) for column in columns])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for index, column_name in enumerate(columns, start=1):
            max_length = len(column_name)
            for row_data in rows:
                value = row_data.get(column_name)
                if value is None:
                    continue
                max_length = max(max_length, len(str(value)))
            width = min(max(12, max_length + 2), 48)
            ws.column_dimensions[get_column_letter(index)].width = width

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(output_path)


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    destination_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {source_dir}")

    destination_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(source_dir, destination_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(source_dir.iterdir()):
            if not file_path.is_file():
                continue

            file_name = file_path.name
            if file_name.startswith("~"):
                print(f"SKIPPED: {file_name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED: {file_name} (not .xlsx)")
                continue

            model_meta = parse_file_label(file_path)
            if model_meta is None:
                print(f"SKIPPED: {file_name} (filename pattern not recognized)")
                continue

            print(f"PROCESSING: {file_name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                processed_files += 1

                empirical_sheet = get_sheet(workbook, EMPIRICAL_SHEET_NAME)
                if empirical_sheet is None:
                    print(f"  - skipped empirical extraction: sheet '{EMPIRICAL_SHEET_NAME}' not found")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            empirical_sheet,
                            model_meta,
                            file_name,
                        )
                    )

                regression_sheet = get_sheet(workbook, REGRESSION_SHEET_NAME)
                if regression_sheet is None:
                    print(f"  - skipped regression extraction: sheet '{REGRESSION_SHEET_NAME}' not found")
                else:
                    regression_rows.extend(
                        process_regression_sheet(
                            regression_sheet,
                            model_meta,
                            file_name,
                        )
                    )
            except Exception as exc:
                print(f"SKIPPED: {file_name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_path}")
    print(f"FILES PROCESSED: {processed_files}")
    print(f"EMPIRICAL ROWS: {len(empirical_rows)}")
    print(f"REGRESSION ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
