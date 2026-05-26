#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("input")
output_dir = Path("output")

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

# Offsets are relative to the "max" anchor cell.
EMPIRICAL_COLUMN_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "sales_captured_in_db_pct": -7,
    "quarterly_sales": -6,
    "growth_rate_pct": -5,
    "avg_penetration_pct": -4,
    "forecast_value": -3,  # estimated total sold
    "actual_value": -2,  # reported sales
    "forecast_min": -1,
    "forecast_max": 0,  # "max" anchor column
}

REGRESSION_COLUMN_OFFSETS = {
    "num_quarters_used": -9,
    "forecast_value": -3,  # TOT FCST w/o SA
    "actual_value": -2,
    "forecast_min": -1,
    "forecast_max": 0,  # "max" anchor column
}

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}
MONTHS = {
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


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], tuple):
        return [list(row) if isinstance(row, tuple) else [row] for row in values]
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_column(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [values]
    output: List[Any] = []
    for value in values:
        if isinstance(value, (list, tuple)):
            output.append(value[0] if value else None)
        else:
            output.append(value)
    return output


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return " ".join(value.strip().lower().split())


def to_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        try:
            return float(cleaned.rstrip("%"))
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def difference(a: Any, b: Any) -> Optional[float]:
    left = to_float(a)
    right = to_float(b)
    if left is None or right is None:
        return None
    return left - right


def values_equal(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    left = to_float(a)
    right = to_float(b)
    if left is not None and right is not None:
        scale = max(1.0, abs(left), abs(right))
        return abs(left - right) <= tolerance * scale
    return a == b


def parse_filename_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""
    period_token = parts[2] if len(parts) >= 3 else ""
    period_token = period_token.split("_")[0]

    period_match = re.search(
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )
    if not period_match:
        fallback_period = period_token if period_token else "unknown_period"
        model = f"{ticker}_{fallback_period}" if ticker else fallback_period
        return {
            "model": model,
            "ticker": ticker,
            "model_period": fallback_period,
            "model_date": "",
        }

    phase = period_match.group("phase").title()
    month_token = period_match.group("month")[:3].lower()
    year = int(period_match.group("year"))
    month_number = MONTHS.get(month_token)
    if month_number is None:
        fallback_period = f"{phase}{period_match.group('month')}_{year}"
        model = f"{ticker}_{fallback_period}" if ticker else fallback_period
        return {
            "model": model,
            "ticker": ticker,
            "model_period": fallback_period,
            "model_date": "",
        }

    model_period = f"{phase}{month_token.title()}_{year}"
    model_day = PHASE_TO_DAY[phase.lower()]
    model_date = date(year, month_number, model_day).isoformat()
    model = f"{ticker}_{model_period}" if ticker else model_period

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_folder / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


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
        workbook.api.Close(SaveChanges=False)
    except TypeError:
        workbook.api.Close(False)


def find_anchor_cell(sheet: xw.Sheet, anchor_label: str = "max") -> Tuple[Any, int, int]:
    used = sheet.used_range
    values = normalize_2d(used.value)
    start_row = used.row
    start_col = used.column
    target = anchor_label.strip().lower()
    for row_index, row_values in enumerate(values):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for column_index, cell_value in enumerate(row_values):
            if normalize_label(cell_value) == target:
                return used, start_row + row_index, start_col + column_index
    raise ValueError(f'Anchor "{anchor_label}" not found in sheet "{sheet.name}"')


def set_formula2_r1c1(cell: xw.Range, formula: str) -> bool:
    try:
        cell.api.Formula2R1C1 = formula
        return True
    except Exception:
        pass
    try:
        cell.formula2 = formula
        return True
    except Exception:
        pass
    try:
        cell.formula = formula
        return True
    except Exception:
        return False


def get_block_value(row_values: Sequence[Any], col_offset: int, min_offset: int) -> Any:
    index = col_offset - min_offset
    if index < 0 or index >= len(row_values):
        return None
    return row_values[index]


def numeric_rows_for_column(
    sheet: xw.Sheet, column: int, first_row: int, last_row: int
) -> List[int]:
    if first_row > last_row:
        return []
    values = normalize_column(sheet.range((first_row, column), (last_row, column)).value)
    rows: List[int] = []
    for idx, value in enumerate(values):
        if to_float(value) is not None:
            rows.append(first_row + idx)
    return rows


def paired_numeric_rows(
    sheet: xw.Sheet, x_col: int, y_col: int, first_row: int, last_row: int
) -> List[int]:
    if first_row > last_row:
        return []
    x_values = normalize_column(sheet.range((first_row, x_col), (last_row, x_col)).value)
    y_values = normalize_column(sheet.range((first_row, y_col), (last_row, y_col)).value)
    rows: List[int] = []
    for idx, (x_value, y_value) in enumerate(zip(x_values, y_values)):
        if to_float(x_value) is not None and to_float(y_value) is not None:
            rows.append(first_row + idx)
    return rows


def next_x_value(sheet: xw.Sheet, x_col: int, rows: List[int]) -> Optional[float]:
    if not rows:
        return None
    last_x = to_float(sheet.range((rows[-1], x_col)).value)
    if last_x is None:
        return None
    if len(rows) >= 2:
        previous_x = to_float(sheet.range((rows[-2], x_col)).value)
        if previous_x is not None:
            step = last_x - previous_x
            if step != 0:
                return last_x + step
    return last_x + 1.0


def extract_empirical_candidates(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = workbook.sheets["Empirical Model"]
    used, anchor_row, anchor_col = find_anchor_cell(sheet, anchor_label="max")

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    min_offset = min(EMPIRICAL_COLUMN_OFFSETS.values())
    max_offset = max(EMPIRICAL_COLUMN_OFFSETS.values())

    row_block = normalize_2d(
        sheet.range((start_row, anchor_col + min_offset), (end_row, anchor_col + max_offset)).value
    )

    scratch_row = used.row + used.rows.count + 2
    scratch_col = used.column + used.columns.count + 2
    avg_penetration_formula_cell = sheet.range((scratch_row, scratch_col))

    penetration_column = anchor_col + EMPIRICAL_COLUMN_OFFSETS["avg_penetration_pct"]
    penetration_rows = numeric_rows_for_column(sheet, penetration_column, 1, anchor_row - 1)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_values = row_block[idx] if idx < len(row_block) else []
        default_quarters = idx + 1

        quarters_value = get_block_value(
            row_values, EMPIRICAL_COLUMN_OFFSETS["num_quarters_used"], min_offset
        )
        num_quarters_used = to_int(quarters_value) or default_quarters

        avg_penetration_value = get_block_value(
            row_values, EMPIRICAL_COLUMN_OFFSETS["avg_penetration_pct"], min_offset
        )
        if len(penetration_rows) >= num_quarters_used:
            first_data_row = penetration_rows[-num_quarters_used]
            last_data_row = penetration_rows[-1]
            avg_formula = (
                f"=AVERAGE(R{first_data_row}C{penetration_column}:"
                f"R{last_data_row}C{penetration_column})"
            )
            if set_formula2_r1c1(avg_penetration_formula_cell, avg_formula):
                workbook.app.calculate()
                calculated_avg = avg_penetration_formula_cell.value
                if calculated_avg not in (None, ""):
                    avg_penetration_value = calculated_avg

        forecast_max = get_block_value(
            row_values, EMPIRICAL_COLUMN_OFFSETS["forecast_max"], min_offset
        )
        forecast_min = get_block_value(
            row_values, EMPIRICAL_COLUMN_OFFSETS["forecast_min"], min_offset
        )
        forecast_value = get_block_value(
            row_values, EMPIRICAL_COLUMN_OFFSETS["forecast_value"], min_offset
        )
        actual_value = get_block_value(
            row_values, EMPIRICAL_COLUMN_OFFSETS["actual_value"], min_offset
        )
        if (
            forecast_max in (None, "")
            and forecast_min in (None, "")
            and forecast_value in (None, "")
            and actual_value in (None, "")
        ):
            continue

        output_row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_value,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": get_block_value(
                row_values, EMPIRICAL_COLUMN_OFFSETS["last_quarter_used"], min_offset
            ),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": difference(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_value,
            "quarterly_sales": get_block_value(
                row_values, EMPIRICAL_COLUMN_OFFSETS["quarterly_sales"], min_offset
            ),
            "reported_sales": actual_value,
            "growth_rate_pct": get_block_value(
                row_values, EMPIRICAL_COLUMN_OFFSETS["growth_rate_pct"], min_offset
            ),
            "sales_captured_in_db_pct": get_block_value(
                row_values, EMPIRICAL_COLUMN_OFFSETS["sales_captured_in_db_pct"], min_offset
            ),
            "source_file": source_file,
        }
        rows.append(output_row)

    return rows


def extract_regression_candidates(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = workbook.sheets["Regression Model"]
    used, anchor_row, anchor_col = find_anchor_cell(sheet, anchor_label="max")

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    min_offset = min(REGRESSION_COLUMN_OFFSETS.values())
    max_offset = max(REGRESSION_COLUMN_OFFSETS.values())
    row_block = normalize_2d(
        sheet.range((start_row, anchor_col + min_offset), (end_row, anchor_col + max_offset)).value
    )

    data_rows = paired_numeric_rows(sheet, x_col, y_col, 1, anchor_row - 1)
    scratch_row = used.row + used.rows.count + 2
    scratch_col = used.column + used.columns.count + 2
    intercept_cell = sheet.range((scratch_row, scratch_col))
    slope_cell = sheet.range((scratch_row, scratch_col + 1))

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_values = row_block[idx] if idx < len(row_block) else []
        default_quarters = idx + 1
        quarters_value = get_block_value(
            row_values, REGRESSION_COLUMN_OFFSETS["num_quarters_used"], min_offset
        )
        num_quarters_used = to_int(quarters_value) or default_quarters
        if len(data_rows) < num_quarters_used:
            continue

        first_data_row = data_rows[-num_quarters_used]
        last_data_row = data_rows[-1]
        intercept_formula = (
            f"=INTERCEPT(R{first_data_row}C{y_col}:R{last_data_row}C{y_col},"
            f"R{first_data_row}C{x_col}:R{last_data_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_data_row}C{y_col}:R{last_data_row}C{y_col},"
            f"R{first_data_row}C{x_col}:R{last_data_row}C{x_col})"
        )

        formulas_set = False
        formulas_set = set_formula2_r1c1(intercept_cell, intercept_formula) or formulas_set
        formulas_set = set_formula2_r1c1(slope_cell, slope_formula) or formulas_set
        if formulas_set:
            workbook.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value

        forecast_value = get_block_value(
            row_values, REGRESSION_COLUMN_OFFSETS["forecast_value"], min_offset
        )
        if to_float(forecast_value) is None:
            x_forecast = next_x_value(sheet, x_col, data_rows)
            intercept_value = to_float(intercept)
            slope_value = to_float(slope)
            if (
                x_forecast is not None
                and intercept_value is not None
                and slope_value is not None
            ):
                forecast_value = intercept_value + slope_value * x_forecast

        forecast_max = get_block_value(
            row_values, REGRESSION_COLUMN_OFFSETS["forecast_max"], min_offset
        )
        forecast_min = get_block_value(
            row_values, REGRESSION_COLUMN_OFFSETS["forecast_min"], min_offset
        )
        actual_value = get_block_value(
            row_values, REGRESSION_COLUMN_OFFSETS["actual_value"], min_offset
        )
        if actual_value in (None, ""):
            actual_value = ""

        output_row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": difference(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows:
            prior_row = rows[-1]
            duplicate = all(
                values_equal(output_row[key], prior_row[key])
                for key in ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
            )
            if duplicate:
                continue

        rows.append(output_row)

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    worksheet = workbook.create_sheet(title=sheet_name)
    worksheet.append(list(headers))
    for row in rows:
        worksheet.append([row.get(header, "") for header in headers])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_index in range(1, worksheet.max_column + 1):
        max_length = 0
        for row_index in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=row_index, column=column_index).value
            text = "" if value is None else str(value)
            max_length = max(max_length, len(text))
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(12, max_length + 2),
            60,
        )


def build_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    workbook.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    source_paths = sorted(input_dir.iterdir(), key=lambda path: path.name.lower())
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_paths:
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (reason: not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (reason: temporary Excel file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (reason: not an .xlsx file)")
                continue

            print(f"Processing file: {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path)
                empirical_rows.extend(
                    extract_empirical_candidates(workbook, metadata, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(workbook, metadata, source_file=file_path.name)
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (reason: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    build_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
