#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Required user-configurable paths.
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

# Offsets are relative to the "max" anchor cell in each model sheet.
EMPIRICAL_ROW_OFFSET = 1
EMPIRICAL_OFFSETS = {
    "last_quarter_col": -12,
    "forecast_value_col": -2,  # estimated total sold
    "actual_value_col": -3,  # reported sales
    "forecast_max_col": 1,
    "forecast_min_col": 2,
    "avg_penetration_source_col": -4,  # quarter-level penetration values
    "quarterly_sales_col": -7,
    "reported_sales_col": -8,
    "growth_rate_col": -6,
    "sales_captured_col": -5,
}

REGRESSION_ROW_OFFSET = 1
REGRESSION_OFFSETS = {
    "num_quarters_col": -1,
    "forecast_value_col": -2,  # TOT FCST w/o SA
    "actual_value_col": -3,  # optional
    "forecast_max_col": 1,
    "forecast_min_col": 2,
}


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_month(month_token: str) -> int | None:
    if not month_token:
        return None
    normalized = month_token.strip()[:3].title()
    try:
        return datetime.strptime(normalized, "%b").month
    except ValueError:
        return None


def parse_model_metadata(file_path: Path) -> ModelMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"

    model_period = "Unknown"
    model_date = ""
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()
        match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token, flags=re.IGNORECASE)
        if match:
            phase_raw, month_raw, year_raw = match.groups()
            phase = phase_raw.capitalize()
            year = int(year_raw)
            month = parse_month(month_raw)
            if month is not None:
                month_abbr = datetime(year=year, month=month, day=1).strftime("%b")
                model_period = f"{phase}{month_abbr}_{year}"
                day_map = {"Early": 5, "Mid": 15, "Late": 25}
                model_date = date(year=year, month=month, day=day_map[phase]).isoformat()

    model = f"{ticker}_{model_period}" if model_period != "Unknown" else ticker
    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def normalize_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def find_anchor_cell(sheet: xw.Sheet, label: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return None
    start_row = used.row
    start_col = used.column
    target = label.strip().lower()
    for row_idx, row_values in enumerate(matrix):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return start_row + row_idx, start_col + col_idx
    return None


def coerce_output_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
    return value


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "").replace("%", "")
        if stripped == "":
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def safe_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def range_width(max_value: Any, min_value: Any) -> Any:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return ""
    return max_num - min_num


def set_formula2(range_obj: xw.Range, formula: str) -> None:
    try:
        range_obj.formula2 = formula
    except Exception:
        # Compatibility fallback for Excel builds without Formula2.
        range_obj.formula = formula


def close_workbook_safely(workbook: xw.Book) -> None:
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
        workbook.api.Close(SaveChanges=False)
    except Exception as exc:
        print(f"warning: failed to close workbook safely ({exc})")


def workbook_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def process_empirical_sheet(workbook: xw.Book, metadata: ModelMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Empirical Model"
    try:
        sheet = workbook.sheets[sheet_name]
    except Exception:
        print(f"skipped file: {source_file} (missing sheet: {sheet_name})")
        return []

    anchor = find_anchor_cell(sheet, label="max")
    if anchor is None:
        print(f"skipped file: {source_file} (missing 'max' anchor in {sheet_name})")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + EMPIRICAL_ROW_OFFSET
    temp_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 15)
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["avg_penetration_source_col"]
    formula_rows: list[tuple[int, int]] = []

    for quarter_count in range(1, N_QUARTERS + 1):
        row = start_row + quarter_count - 1
        first_row = row - quarter_count + 1
        if first_row < 1 or penetration_col < 1:
            continue
        formula = f"=AVERAGE(R{first_row}C{penetration_col}:R{row}C{penetration_col})"
        set_formula2(sheet.range((row, temp_col)), formula)
        formula_rows.append((quarter_count, row))

    if formula_rows:
        workbook.app.calculate()

    rows: list[dict[str, Any]] = []
    for quarter_count, row in formula_rows:
        forecast_max = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_max_col"])
        )
        forecast_min = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_min_col"])
        )
        avg_penetration = coerce_output_value(sheet.range((row, temp_col)).value)
        forecast_value = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_value_col"])
        )
        actual_value = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["actual_value_col"])
        )
        quarterly_sales = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"])
        )
        reported_sales = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"])
        )
        growth_rate = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_col"])
        )
        sales_captured = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_col"])
        )
        last_quarter_used = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_col"])
        )

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": quarter_count,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def regression_rows_match(current: dict[str, Any], previous: dict[str, Any]) -> bool:
    keys = ["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for key in keys:
        current_num = to_float(current.get(key))
        previous_num = to_float(previous.get(key))
        if current_num is None and previous_num is None:
            continue
        if current_num is None or previous_num is None:
            return False
        if not math.isclose(current_num, previous_num, rel_tol=1e-9, abs_tol=1e-9):
            return False
    return True


def process_regression_sheet(workbook: xw.Book, metadata: ModelMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Regression Model"
    try:
        sheet = workbook.sheets[sheet_name]
    except Exception:
        print(f"skipped file: {source_file} (missing sheet: {sheet_name})")
        return []

    anchor = find_anchor_cell(sheet, label="max")
    if anchor is None:
        print(f"skipped file: {source_file} (missing 'max' anchor in {sheet_name})")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"skipped file: {source_file} (invalid regression x/y offsets from anchor)")
        return []

    start_row = anchor_row + REGRESSION_ROW_OFFSET
    temp_intercept_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 15)
    temp_slope_col = temp_intercept_col + 1
    formula_rows: list[tuple[int, int]] = []

    for quarter_count in range(1, N_QUARTERS + 1):
        row = start_row + quarter_count - 1
        first_row = row - quarter_count + 1
        if first_row < 1:
            continue

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{row}C{y_col},R{first_row}C{x_col}:R{row}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{first_row}C{y_col}:R{row}C{y_col},R{first_row}C{x_col}:R{row}C{x_col})"
        set_formula2(sheet.range((row, temp_intercept_col)), intercept_formula)
        set_formula2(sheet.range((row, temp_slope_col)), slope_formula)
        formula_rows.append((quarter_count, row))

    if formula_rows:
        workbook.app.calculate()

    rows: list[dict[str, Any]] = []
    for idx, (quarter_count, row) in enumerate(formula_rows):
        num_quarters_raw = safe_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["num_quarters_col"])
        num_quarters_used = int(to_float(num_quarters_raw) or quarter_count)
        forecast_value = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_value_col"])
        )
        actual_value = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["actual_value_col"])
        )
        forecast_max = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_max_col"])
        )
        forecast_min = coerce_output_value(
            safe_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_min_col"])
        )
        intercept = coerce_output_value(sheet.range((row, temp_intercept_col)).value)
        slope = coerce_output_value(sheet.range((row, temp_slope_col)).value)

        row_payload = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        # Drop an accidental duplicate final row if calculations match the prior row.
        if idx == len(formula_rows) - 1 and rows and regression_rows_match(row_payload, rows[-1]):
            continue
        rows.append(row_payload)

    return rows


def write_sheet(sheet, columns: list[str], rows: list[dict[str, Any]]) -> None:
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column, "") for column in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_file: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_file)


def should_skip_file(file_path: Path) -> tuple[bool, str]:
    if not file_path.is_file():
        return True, "not a regular file"
    if file_path.name.startswith("~"):
        return True, "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return True, "not an .xlsx file"
    if re.search(r"_PARAM(\.\d+)?\.xlsx$", file_path.name, flags=re.IGNORECASE):
        return True, "prior parameter output workbook"
    return False, ""


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    final_output_file = workbook_output_path(input_path=input_path, output_path=output_path)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_path.iterdir()):
            skip, reason = should_skip_file(file_path)
            if skip:
                print(f"skipped file: {file_path.name} ({reason})")
                continue

            source_workbook = None
            try:
                source_workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path)
                empirical_rows.extend(
                    process_empirical_sheet(
                        workbook=source_workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        workbook=source_workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if source_workbook is not None:
                    close_workbook_safely(source_workbook)
    finally:
        app.quit()

    write_output_workbook(
        output_file=final_output_file,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"output path: {final_output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
