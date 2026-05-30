#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", re.IGNORECASE)
PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def as_number(value: Any) -> Optional[float]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1]
        try:
            return float(text)
        except ValueError:
            return None
    return None


def as_int(value: Any, default: int) -> int:
    numeric_value = as_number(value)
    if numeric_value is None:
        return default
    return int(round(numeric_value))


def subtract_values(left: Any, right: Any) -> Optional[float]:
    left_num = as_number(left)
    right_num = as_number(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def compare_value(value: Any) -> Any:
    numeric_value = as_number(value)
    if numeric_value is not None:
        return round(numeric_value, 10)
    return value


def parse_model_meta(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("filename does not match '<prefix> - <ticker> - <period>_...xlsx'")

    ticker = re.sub(r"[^A-Za-z0-9]+", "", parts[1]).upper()
    if not ticker:
        raise ValueError("ticker token is empty")

    period_token = parts[2].split("_")[0].strip()
    period_match = PERIOD_PATTERN.search(period_token)
    if not period_match:
        raise ValueError("period token does not match Early/Mid/Late + Mon + YYYY")

    phase = period_match.group(1).title()
    month_abbr = period_match.group(2).title()
    year = int(period_match.group(3))
    phase_day_key = phase.lower()
    if phase_day_key not in PERIOD_DAY_MAP:
        raise ValueError(f"unsupported period phase: {phase}")

    try:
        month_number = datetime.strptime(month_abbr, "%b").month
    except ValueError as exc:
        raise ValueError(f"unsupported month abbreviation: {month_abbr}") from exc

    model_period = f"{phase}{month_abbr}_{year}"
    model_date = date(year, month_number, PERIOD_DAY_MAP[phase_day_key]).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def choose_output_path(input_folder: Path, destination_folder: Path) -> Path:
    destination_folder.mkdir(parents=True, exist_ok=True)
    folder_name = input_folder.resolve().name
    base_name = f"{folder_name}_PARAM"

    candidate = destination_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = destination_folder / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def list_source_files(folder: Path) -> List[Path]:
    files: List[Path] = []
    if not folder.exists():
        print(f"Skipped input folder: {folder} does not exist")
        return files

    for path in sorted(folder.iterdir()):
        if not path.is_file():
            print(f"Skipped file {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"Skipped file {path.name}: temporary Excel file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped file {path.name}: not an .xlsx file")
            continue
        files.append(path)
    return files


def safe_close_no_save(wb: xw.Book) -> None:
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
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def used_grid(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return used.row, used.column, []
    if not isinstance(values, list):
        return used.row, used.column, [[values]]
    if values and not isinstance(values[0], list):
        return used.row, used.column, [values]
    return used.row, used.column, values


def find_anchor_max(top_row: int, left_col: int, values: List[List[Any]]) -> Optional[Tuple[int, int]]:
    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return top_row + row_offset, left_col + col_offset
    return None


def header_map_for_row(
    top_row: int, left_col: int, values: List[List[Any]], header_row: int
) -> Dict[str, int]:
    row_index = header_row - top_row
    if row_index < 0 or row_index >= len(values):
        return {}
    mapping: Dict[str, int] = {}
    for col_offset, cell_value in enumerate(values[row_index]):
        key = normalize_label(cell_value)
        if key and key not in mapping:
            mapping[key] = left_col + col_offset
    return mapping


def pick_column(
    headers: Dict[str, int], aliases: List[str], anchor_col: int, fallback_offset: int
) -> int:
    for alias in aliases:
        key = normalize_label(alias)
        if key in headers:
            return headers[key]
    return max(1, anchor_col + fallback_offset)


def extract_empirical_candidates(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped file {source_file}: missing sheet 'Empirical Model'")
        return []

    top_row, left_col, values = used_grid(sheet)
    anchor = find_anchor_max(top_row, left_col, values)
    if not anchor:
        print(f"Skipped file {source_file}: could not find 'max' anchor in Empirical Model")
        return []

    anchor_row, anchor_col = anchor
    headers = header_map_for_row(top_row, left_col, values, anchor_row)
    widest_row = max((len(row) for row in values), default=1)
    helper_col = max(anchor_col + 6, left_col + widest_row + 2)

    data_start_row = anchor_row + 1
    num_quarters_col = pick_column(
        headers,
        ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
        anchor_col,
        -10,
    )
    last_quarter_col = pick_column(
        headers,
        ["last_quarter_used", "last_qtr_used", "last_quarter"],
        anchor_col,
        -9,
    )
    forecast_value_col = pick_column(
        headers,
        ["estimated_total_sold", "forecast_value", "tot_fcst_w_o_sa", "tot_fcst_wo_sa"],
        anchor_col,
        -1,
    )
    actual_value_col = pick_column(
        headers,
        ["reported_sales", "actual_sales", "actual_value"],
        anchor_col,
        -2,
    )
    forecast_min_col = pick_column(headers, ["min", "forecast_min"], anchor_col, 1)
    penetration_source_col = pick_column(
        headers,
        ["penetration_pct", "penetration", "avg_penetration_pct", "sales_captured_in_db_pct"],
        anchor_col,
        -4,
    )
    quarterly_sales_col = pick_column(headers, ["quarterly_sales"], anchor_col, -5)
    reported_sales_col = pick_column(headers, ["reported_sales"], anchor_col, -2)
    growth_rate_col = pick_column(headers, ["growth_rate_pct", "growth_rate"], anchor_col, -6)
    sales_captured_col = pick_column(
        headers,
        ["sales_captured_in_db_pct", "sales_captured_pct"],
        anchor_col,
        -7,
    )

    for n_quarters in range(1, N_QUARTERS + 1):
        row = data_start_row + n_quarters - 1
        start_row = max(data_start_row, row - n_quarters + 1)
        avg_formula = f"=AVERAGE(R{start_row}C{penetration_source_col}:R{row}C{penetration_source_col})"
        set_formula2(sheet.cells(row, helper_col), avg_formula)

    wb.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        row = data_start_row + n_quarters - 1
        num_quarters_used = as_int(sheet.cells(row, num_quarters_col).value, n_quarters)
        last_quarter_used = sheet.cells(row, last_quarter_col).value
        forecast_value = sheet.cells(row, forecast_value_col).value
        actual_value = sheet.cells(row, actual_value_col).value
        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, forecast_min_col).value
        avg_penetration_pct = sheet.cells(row, helper_col).value
        quarterly_sales = sheet.cells(row, quarterly_sales_col).value
        reported_sales = sheet.cells(row, reported_sales_col).value
        growth_rate_pct = sheet.cells(row, growth_rate_col).value
        sales_captured_in_db_pct = sheet.cells(row, sales_captured_col).value
        range_width = subtract_values(forecast_max, forecast_min)

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        extracted_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return extracted_rows


def extract_regression_candidates(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped file {source_file}: missing sheet 'Regression Model'")
        return []

    top_row, left_col, values = used_grid(sheet)
    anchor = find_anchor_max(top_row, left_col, values)
    if not anchor:
        print(f"Skipped file {source_file}: could not find 'max' anchor in Regression Model")
        return []

    anchor_row, anchor_col = anchor
    headers = header_map_for_row(top_row, left_col, values, anchor_row)
    widest_row = max((len(row) for row in values), default=1)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_start_row = anchor_row + 1
    helper_intercept_col = max(anchor_col + 6, left_col + widest_row + 2)
    helper_slope_col = helper_intercept_col + 1

    num_quarters_col = pick_column(
        headers,
        ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
        anchor_col,
        -10,
    )
    forecast_value_col = pick_column(
        headers,
        ["tot_fcst_w_o_sa", "tot_fcst_wo_sa", "forecast_total_without_sa", "forecast_value"],
        anchor_col,
        -1,
    )
    actual_value_col = pick_column(
        headers,
        ["actual_value", "actual_sales", "reported_sales"],
        anchor_col,
        -2,
    )
    forecast_min_col = pick_column(headers, ["min", "forecast_min"], anchor_col, 1)

    for n_quarters in range(1, N_QUARTERS + 1):
        row = data_start_row + n_quarters - 1
        start_row = max(data_start_row, row - n_quarters + 1)
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"
        )
        set_formula2(sheet.cells(row, helper_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row, helper_slope_col), slope_formula)

    wb.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Any, ...]] = None

    for n_quarters in range(1, N_QUARTERS + 1):
        row = data_start_row + n_quarters - 1
        num_quarters_used = as_int(sheet.cells(row, num_quarters_col).value, n_quarters)
        forecast_value = sheet.cells(row, forecast_value_col).value
        actual_value = sheet.cells(row, actual_value_col).value
        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, forecast_min_col).value
        intercept = sheet.cells(row, helper_intercept_col).value
        slope = sheet.cells(row, helper_slope_col).value
        range_width = subtract_values(forecast_max, forecast_min)

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        ):
            continue

        dedupe_key = (
            compare_value(num_quarters_used),
            compare_value(forecast_value),
            compare_value(forecast_max),
            compare_value(forecast_min),
            compare_value(intercept),
            compare_value(slope),
        )
        if dedupe_key == previous_key:
            continue
        previous_key = dedupe_key

        extracted_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
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

    return extracted_rows


def format_output_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row_data in rows:
        ws.append([row_data.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(ws.max_row, 1)}"

    for col_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_index in range(2, ws.max_row + 1):
            value = ws.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max_length + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    format_output_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    format_output_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def main() -> None:
    source_files = list_source_files(input_dir)
    output_path = choose_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        if source_files:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            for file_path in source_files:
                try:
                    meta = parse_model_meta(file_path)
                except ValueError as exc:
                    print(f"Skipped file {file_path.name}: {exc}")
                    continue

                wb: Optional[xw.Book] = None
                try:
                    wb = app.books.open(str(file_path), update_links=False)
                    empirical_rows.extend(extract_empirical_candidates(wb, meta, file_path.name))
                    regression_rows.extend(extract_regression_candidates(wb, meta, file_path.name))
                    processed_files += 1
                    print(f"Processed file: {file_path.name}")
                except Exception as exc:
                    print(f"Skipped file {file_path.name}: {exc}")
                finally:
                    if wb is not None:
                        safe_close_no_save(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
