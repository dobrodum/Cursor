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

# --------- Configure folders here ---------
input_dir = Path("input")
output_dir = Path("output")
# -----------------------------------------

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

PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", re.IGNORECASE)
DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = re.sub(r"[^A-Za-z0-9]+", " ", str(value).strip().lower())
    return " ".join(text.split())


def month_number_from_token(token: str) -> int:
    clean = re.sub(r"[^A-Za-z]", "", token).lower()
    month_lookup: dict[str, int] = {}
    for month_index in range(1, 13):
        month_lookup[calendar.month_abbr[month_index].lower()] = month_index
        month_lookup[calendar.month_name[month_index].lower()] = month_index
    month_lookup["sept"] = 9

    if clean in month_lookup:
        return month_lookup[clean]
    short = clean[:3]
    if short in month_lookup:
        return month_lookup[short]
    raise ValueError(f"Unknown month token: {token}")


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError(f"Filename does not match expected pattern: {file_path.name}")

    ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        raise ValueError(f"Could not parse ticker from filename: {file_path.name}")

    period_match = PERIOD_PATTERN.search(stem)
    if not period_match:
        raise ValueError(f"Could not parse model period from filename: {file_path.name}")

    period_word = period_match.group(1).lower()
    month_token = period_match.group(2)
    year = int(period_match.group(3))

    month_number = month_number_from_token(month_token)
    month_abbr = calendar.month_abbr[month_number]
    model_period = f"{period_word.title()}{month_abbr}_{year}"
    model_day = DAY_BY_PERIOD[period_word]
    model_date = date(year, month_number, model_day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    primary = output_folder / f"{base_name}.xlsx"
    if not primary.exists():
        return primary

    suffix = 1
    while True:
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_workbook_without_saving(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        # Older backends may not support the save kwarg.
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def to_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def clean_number(value: Any) -> float | int | None:
    number = to_number(value)
    if number is None:
        return None
    if abs(number - round(number)) < 1e-12:
        return int(round(number))
    return number


def number_or_blank(value: Any) -> float | int | str:
    parsed = clean_number(value)
    return "" if parsed is None else parsed


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = normalize_2d(used.value)
    start_row = used.row
    start_col = used.column

    needle = anchor_text.strip().lower()
    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == needle:
                return start_row + row_offset, start_col + col_offset
    return None


def build_header_map(sheet: xw.Sheet, header_row: int) -> tuple[dict[str, int], int]:
    used = sheet.used_range
    first_col = used.column
    last_col = first_col + used.columns.count - 1
    header_map: dict[str, int] = {}

    for col_idx in range(first_col, last_col + 1):
        key = normalize_label(sheet.range((header_row, col_idx)).value)
        if key and key not in header_map:
            header_map[key] = col_idx
    return header_map, last_col


def select_col(header_map: dict[str, int], candidates: list[str], default: int | None = None) -> int | None:
    normalized_candidates = [normalize_label(candidate) for candidate in candidates]
    for candidate in normalized_candidates:
        if candidate in header_map:
            return header_map[candidate]

    for candidate in normalized_candidates:
        for key, col_idx in header_map.items():
            if candidate in key or key in candidate:
                return col_idx
    return default


def values_equal(left: Any, right: Any) -> bool:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is not None and right_num is not None:
        return abs(left_num - right_num) <= 1e-9
    return left == right


def extract_empirical_rows(workbook: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Empirical Model"
    if sheet_name not in {sheet.name for sheet in workbook.sheets}:
        print(f"Skipped empirical extraction for {source_file}: sheet '{sheet_name}' not found")
        return []

    sheet = workbook.sheets[sheet_name]
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map, last_used_col = build_header_map(sheet, anchor_row)

    num_quarters_col = select_col(
        header_map, ["num quarters used", "num quarters", "quarters used", "n quarters"], anchor_col - 4
    )
    last_quarter_col = select_col(header_map, ["last quarter used", "last quarter"], anchor_col - 3)
    forecast_value_col = select_col(
        header_map, ["estimated total sold", "forecast value", "total forecast", "tot fcst"], anchor_col - 1
    )
    actual_value_col = select_col(
        header_map, ["reported sales", "actual value", "actual sales", "actual"], anchor_col - 2
    )
    forecast_max_col = select_col(header_map, ["max", "forecast max"], anchor_col)
    forecast_min_col = select_col(header_map, ["min", "forecast min"], anchor_col + 1)
    penetration_col = select_col(
        header_map,
        ["avg penetration pct", "penetration pct", "penetration", "average penetration"],
        anchor_col - 6,
    )
    quarterly_sales_col = select_col(header_map, ["quarterly sales", "sales"], anchor_col - 5)
    growth_rate_col = select_col(header_map, ["growth rate pct", "growth rate", "growth"], anchor_col + 2)
    sales_captured_col = select_col(
        header_map,
        ["sales captured in db pct", "sales captured in db", "captured in db"],
        anchor_col + 3,
    )

    helper_col = last_used_col + 3
    rows_with_formula: list[dict[str, Any]] = []

    for step in range(1, 11):
        row_idx = anchor_row + step
        num_quarters = int(to_number(sheet.range((row_idx, num_quarters_col)).value) or step)
        if num_quarters < 1:
            num_quarters = step

        forecast_value = clean_number(sheet.range((row_idx, forecast_value_col)).value)
        actual_value = clean_number(sheet.range((row_idx, actual_value_col)).value)
        forecast_max = clean_number(sheet.range((row_idx, forecast_max_col)).value)
        forecast_min = clean_number(sheet.range((row_idx, forecast_min_col)).value)

        if all(value is None for value in (forecast_value, forecast_max, forecast_min, actual_value)):
            continue

        start_row = max(anchor_row + 1, row_idx - num_quarters + 1)
        helper_cell = sheet.range((row_idx, helper_col))
        helper_cell.formula2 = (
            f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{row_idx}C{penetration_col}),"")'
        )

        rows_with_formula.append(
            {
                "row_idx": row_idx,
                "num_quarters": num_quarters,
                "last_quarter_used": sheet.range((row_idx, last_quarter_col)).value,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "quarterly_sales": clean_number(sheet.range((row_idx, quarterly_sales_col)).value),
                "growth_rate_pct": clean_number(sheet.range((row_idx, growth_rate_col)).value),
                "sales_captured_in_db_pct": clean_number(sheet.range((row_idx, sales_captured_col)).value),
                "helper_cell": helper_cell,
            }
        )

    if rows_with_formula:
        workbook.app.calculate()

    output_rows: list[dict[str, Any]] = []
    for item in rows_with_formula:
        avg_penetration = clean_number(item["helper_cell"].value)
        item["helper_cell"].value = None

        forecast_max = item["forecast_max"]
        forecast_min = item["forecast_min"]
        range_width: float | int | None = None
        if forecast_max is not None and forecast_min is not None:
            range_width = clean_number(float(forecast_max) - float(forecast_min))

        output_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": number_or_blank(avg_penetration),
                "num_quarters_used": item["num_quarters"],
                "last_quarter_used": item["last_quarter_used"] or "",
                "forecast_value": number_or_blank(item["forecast_value"]),
                "actual_value": number_or_blank(item["actual_value"]),
                "forecast_max": number_or_blank(forecast_max),
                "forecast_min": number_or_blank(forecast_min),
                "range_width": number_or_blank(range_width),
                "avg_penetration_pct": number_or_blank(avg_penetration),
                "quarterly_sales": number_or_blank(item["quarterly_sales"]),
                "reported_sales": number_or_blank(item["actual_value"]),
                "growth_rate_pct": number_or_blank(item["growth_rate_pct"]),
                "sales_captured_in_db_pct": number_or_blank(item["sales_captured_in_db_pct"]),
                "source_file": source_file,
            }
        )

    return output_rows


def extract_regression_rows(workbook: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Regression Model"
    if sheet_name not in {sheet.name for sheet in workbook.sheets}:
        print(f"Skipped regression extraction for {source_file}: sheet '{sheet_name}' not found")
        return []

    sheet = workbook.sheets[sheet_name]
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map, last_used_col = build_header_map(sheet, anchor_row)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = select_col(
        header_map, ["num quarters used", "num quarters", "quarters used", "n quarters"], anchor_col - 12
    )
    forecast_value_col = select_col(
        header_map,
        ["tot fcst w o sa", "tot fcst w/o sa", "forecast total without sa", "total forecast without sa"],
        anchor_col - 1,
    )
    actual_value_col = select_col(header_map, ["actual value", "actual sales", "reported sales"], None)
    forecast_max_col = select_col(header_map, ["max", "forecast max"], anchor_col)
    forecast_min_col = select_col(header_map, ["min", "forecast min"], anchor_col + 1)

    helper_intercept_col = last_used_col + 3
    helper_slope_col = last_used_col + 4
    rows_with_formula: list[dict[str, Any]] = []

    for step in range(1, 11):
        row_idx = anchor_row + step
        num_quarters = int(to_number(sheet.range((row_idx, num_quarters_col)).value) or step)
        if num_quarters < 1:
            num_quarters = step

        forecast_value = clean_number(sheet.range((row_idx, forecast_value_col)).value)
        forecast_max = clean_number(sheet.range((row_idx, forecast_max_col)).value)
        forecast_min = clean_number(sheet.range((row_idx, forecast_min_col)).value)

        if all(value is None for value in (forecast_value, forecast_max, forecast_min)):
            continue

        actual_value = (
            clean_number(sheet.range((row_idx, actual_value_col)).value) if actual_value_col is not None else None
        )
        start_row = max(anchor_row + 1, row_idx - num_quarters + 1)

        intercept_cell = sheet.range((row_idx, helper_intercept_col))
        slope_cell = sheet.range((row_idx, helper_slope_col))
        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row_idx}C{y_col},'
            f'R{start_row}C{x_col}:R{row_idx}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{row_idx}C{y_col},'
            f'R{start_row}C{x_col}:R{row_idx}C{x_col}),"")'
        )

        rows_with_formula.append(
            {
                "num_quarters": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "intercept_cell": intercept_cell,
                "slope_cell": slope_cell,
            }
        )

    if rows_with_formula:
        workbook.app.calculate()

    output_rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for item in rows_with_formula:
        intercept_value = clean_number(item["intercept_cell"].value)
        slope_value = clean_number(item["slope_cell"].value)
        item["intercept_cell"].value = None
        item["slope_cell"].value = None

        forecast_max = item["forecast_max"]
        forecast_min = item["forecast_min"]
        range_width: float | int | None = None
        if forecast_max is not None and forecast_min is not None:
            range_width = clean_number(float(forecast_max) - float(forecast_min))

        current_signature = (
            item["num_quarters"],
            intercept_value,
            slope_value,
            item["forecast_value"],
            forecast_max,
            forecast_min,
        )
        if previous_signature is not None and all(
            values_equal(current_signature[idx], previous_signature[idx]) for idx in range(len(current_signature))
        ):
            continue
        previous_signature = current_signature

        output_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": item["num_quarters"],
                "num_quarters_used": item["num_quarters"],
                "forecast_value": number_or_blank(item["forecast_value"]),
                "actual_value": number_or_blank(item["actual_value"]),
                "forecast_max": number_or_blank(forecast_max),
                "forecast_min": number_or_blank(forecast_min),
                "range_width": number_or_blank(range_width),
                "intercept": number_or_blank(intercept_value),
                "slope": number_or_blank(slope_value),
                "source_file": source_file,
            }
        )

    return output_rows


def write_sheet(worksheet, columns: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet.append(columns)
    for row in rows:
        worksheet.append([row.get(column, "") for column in columns])

    for header_cell in worksheet[1]:
        header_cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        longest = len(column_name)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is not None:
                longest = max(longest, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(longest + 2, 48)


def save_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    output_wb = Workbook()
    empirical_ws = output_wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = output_wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    output_wb.save(output_path)


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input directory not found: {input_dir.resolve()}")
        return

    output_path = build_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    try:
        app = xw.App(visible=False, add_book=False)
    except Exception as exc:
        print(f"Failed to start Excel app: {exc}")
        return

    for attr_name, attr_value in [
        ("display_alerts", False),
        ("screen_updating", False),
        ("enable_events", False),
    ]:
        try:
            setattr(app, attr_name, attr_value)
        except Exception:
            pass

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"Skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not .xlsx)")
                continue

            try:
                file_label = parse_file_label(file_path)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (filename parse error: {exc})")
                continue

            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, file_label, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, file_label, file_path.name))
                processed_files += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_without_saving(workbook)
    finally:
        app.quit()

    save_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output workbook: {output_path.resolve()}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
