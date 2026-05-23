#!/usr/bin/env python3
"""Extract empirical/regression model candidates from Excel workbooks."""

from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure folders here.
input_dir = Path("input")
output_dir = Path("output")

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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_NAME_TO_NUMBER = {
    name.lower(): number
    for number, name in enumerate(calendar.month_name)
    if number
}
MONTH_ABBR_TO_NUMBER = {
    abbr.lower(): number
    for number, abbr in enumerate(calendar.month_abbr)
    if number
}


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"\s+", " ", value.strip().lower())


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def safe_div(numerator: Any, denominator: Any) -> float | None:
    num = to_float(numerator)
    den = to_float(denominator)
    if num is None or den in (None, 0):
        return None
    return num / den


def round_for_signature(value: Any) -> Any:
    numeric = to_float(value)
    if numeric is None:
        return None
    return round(numeric, 10)


def parse_filename_metadata(file_name: str) -> ModelMeta:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_match = re.search(
        r"(?i)(early|mid|late)\s*([A-Za-z]{3,9})\s*[_\-\s]?(\d{4})",
        period_source,
    )

    model_period = "unknown_period"
    model_date = ""
    if period_match:
        phase = period_match.group(1).lower()
        month_text = period_match.group(2).lower()
        year = int(period_match.group(3))

        month_number = MONTH_ABBR_TO_NUMBER.get(month_text[:3])
        if month_number is None:
            month_number = MONTH_NAME_TO_NUMBER.get(month_text)

        if month_number:
            day = PHASE_TO_DAY[phase]
            phase_token = phase.capitalize()
            month_token = calendar.month_abbr[month_number]
            model_period = f"{phase_token}{month_token}_{year}"
            model_date = date(year, month_number, day).isoformat()
        else:
            model_period = re.sub(r"\W+", "_", period_match.group(0)).strip("_")

    model = f"{ticker}_{model_period}"
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_used_range_matrix(sheet: xw.Sheet) -> tuple[list[list[Any]], int, int, int, int]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return [], used.row, used.column, used.last_cell.row, used.last_cell.column
    if not isinstance(values, list):
        matrix = [[values]]
    elif values and not isinstance(values[0], list):
        matrix = [values]
    else:
        matrix = values
    return matrix, used.row, used.column, used.last_cell.row, used.last_cell.column


def build_label_positions(
    matrix: list[list[Any]], base_row: int, base_col: int
) -> list[tuple[int, int, str]]:
    labels: list[tuple[int, int, str]] = []
    for row_offset, row in enumerate(matrix):
        for col_offset, value in enumerate(row):
            normalized = normalize_label(value)
            if normalized:
                labels.append((base_row + row_offset, base_col + col_offset, normalized))
    return labels


def find_anchor_max(labels: Sequence[tuple[int, int, str]]) -> tuple[int, int] | None:
    exact = [(row, col) for row, col, label in labels if label == "max"]
    if exact:
        return max(exact, key=lambda item: (item[0], item[1]))

    broad = [(row, col) for row, col, label in labels if re.search(r"\bmax\b", label)]
    if broad:
        return max(broad, key=lambda item: (item[0], item[1]))
    return None


def find_nearest_label(
    labels: Sequence[tuple[int, int, str]],
    keywords: Iterable[str],
    anchor: tuple[int, int] | None,
) -> tuple[int, int, str] | None:
    normalized_keywords = [normalize_label(keyword) for keyword in keywords]
    matches = [
        (row, col, label)
        for row, col, label in labels
        if any(keyword in label for keyword in normalized_keywords)
    ]
    if not matches:
        return None
    if anchor is None:
        return matches[0]
    return min(matches, key=lambda item: abs(item[0] - anchor[0]) + abs(item[1] - anchor[1]))


def value_right_of(sheet: xw.Sheet, row: int, col: int, max_steps: int = 6) -> Any:
    for step in range(1, max_steps + 1):
        value = sheet.range((row, col + step)).value
        if value not in (None, ""):
            return value
    return None


def numeric_value_right_of(sheet: xw.Sheet, row: int, col: int, max_steps: int = 6) -> float | None:
    for step in range(1, max_steps + 1):
        numeric = to_float(sheet.range((row, col + step)).value)
        if numeric is not None:
            return numeric
    return None


def get_sheet_case_insensitive(workbook: xw.Book, name: str) -> xw.Sheet | None:
    target = normalize_label(name)
    for sheet in workbook.sheets:
        if normalize_label(sheet.name) == target:
            return sheet
    return None


def find_num_quarters_input_cell(
    labels: Sequence[tuple[int, int, str]],
    anchor: tuple[int, int] | None,
) -> tuple[int, int] | None:
    match = find_nearest_label(labels, ["num quarters", "num_quarters"], anchor)
    if not match:
        return None
    row, col, _ = match
    return row, col + 1


def get_label_numeric_value(
    sheet: xw.Sheet,
    labels: Sequence[tuple[int, int, str]],
    keywords: Sequence[str],
    anchor: tuple[int, int] | None,
) -> float | None:
    match = find_nearest_label(labels, keywords, anchor)
    if not match:
        return None
    row, col, _ = match
    return numeric_value_right_of(sheet, row, col)


def get_label_value(
    sheet: xw.Sheet,
    labels: Sequence[tuple[int, int, str]],
    keywords: Sequence[str],
    anchor: tuple[int, int] | None,
) -> Any:
    match = find_nearest_label(labels, keywords, anchor)
    if not match:
        return None
    row, col, _ = match
    return value_right_of(sheet, row, col)


def collect_numeric_pairs(
    sheet: xw.Sheet, x_col: int, y_col: int, first_row: int, last_row: int
) -> list[tuple[int, float, float]]:
    if x_col <= 0 or y_col <= 0 or first_row > last_row:
        return []
    x_values = sheet.range((first_row, x_col), (last_row, x_col)).value
    y_values = sheet.range((first_row, y_col), (last_row, y_col)).value
    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]
    pairs: list[tuple[int, float, float]] = []
    for idx, (x_value, y_value) in enumerate(zip(x_values, y_values)):
        x_numeric = to_float(x_value)
        y_numeric = to_float(y_value)
        if x_numeric is None or y_numeric is None:
            continue
        pairs.append((first_row + idx, x_numeric, y_numeric))
    return pairs


def set_r1c1_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        try:
            cell.api.FormulaR1C1 = formula_r1c1
        except Exception:
            cell.formula2 = formula_r1c1


def close_workbook_safe(workbook: xw.Book | None) -> None:
    if workbook is None:
        return
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
    except Exception:
        # Final fallback: best effort close.
        try:
            workbook.close()
        except Exception:
            pass


def process_empirical_sheet(workbook: xw.Book, source_file: str, meta: ModelMeta) -> list[dict[str, Any]]:
    sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
    if sheet is None:
        return []

    matrix, base_row, base_col, _, last_col = get_used_range_matrix(sheet)
    labels = build_label_positions(matrix, base_row, base_col)
    anchor = find_anchor_max(labels)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    quarterly_col = max(1, anchor_col - 11)
    reported_col = max(1, anchor_col - 7)
    history = collect_numeric_pairs(sheet, quarterly_col, reported_col, base_row, max(base_row, anchor_row - 1))
    if not history:
        history = collect_numeric_pairs(sheet, reported_col, quarterly_col, base_row, max(base_row, anchor_row - 1))
        if history:
            quarterly_col, reported_col = reported_col, quarterly_col
    if not history:
        return []

    num_quarters_input = find_num_quarters_input_cell(labels, anchor)
    forecast_label_keywords = ["estimated total sold", "forecast"]
    actual_label_keywords = ["reported sales", "actual sales", "actual"]
    min_label = find_nearest_label(labels, ["min"], anchor)
    last_quarter_label = find_nearest_label(labels, ["last quarter"], anchor)

    scratch_col = last_col + 3
    scratch_row = max(anchor_row, base_row + 1)
    scratch_cell = sheet.range((scratch_row, scratch_col))
    rows: list[dict[str, Any]] = []

    n_quarters = 10
    available = min(n_quarters, len(history))
    for count in range(1, available + 1):
        segment = history[-count:]
        start_row = segment[0][0]
        end_row = segment[-1][0]

        if num_quarters_input is not None:
            sheet.range(num_quarters_input).value = count

        avg_formula = (
            f"=AVERAGE("
            f"R{start_row}C{quarterly_col}:R{end_row}C{quarterly_col}/"
            f"R{start_row}C{reported_col}:R{end_row}C{reported_col})"
        )
        set_r1c1_formula2(scratch_cell, avg_formula)
        workbook.app.calculate()

        avg_penetration = to_float(scratch_cell.value)
        quarterly_sales = segment[-1][1]
        reported_sales = segment[-1][2]

        forecast_value = get_label_numeric_value(sheet, labels, forecast_label_keywords, anchor)
        actual_value = get_label_numeric_value(sheet, labels, actual_label_keywords, anchor)
        forecast_max = numeric_value_right_of(sheet, anchor_row, anchor_col)
        forecast_min = numeric_value_right_of(sheet, min_label[0], min_label[1]) if min_label else None
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )
        growth_rate_pct = get_label_numeric_value(sheet, labels, ["growth rate"], anchor)
        sales_captured_pct = get_label_numeric_value(
            sheet,
            labels,
            ["sales captured in db", "captured in db"],
            anchor,
        )
        if sales_captured_pct is None:
            sales_captured_pct = safe_div(quarterly_sales, reported_sales)

        last_quarter_used: Any = None
        if last_quarter_label:
            last_quarter_used = value_right_of(sheet, last_quarter_label[0], last_quarter_label[1])
        if last_quarter_used in (None, ""):
            quarter_col = max(1, quarterly_col - 1)
            last_quarter_used = sheet.range((end_row, quarter_col)).value

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": count,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
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

    try:
        scratch_cell.clear_contents()
    except Exception:
        pass
    return rows


def process_regression_sheet(workbook: xw.Book, source_file: str, meta: ModelMeta) -> list[dict[str, Any]]:
    sheet = get_sheet_case_insensitive(workbook, "Regression Model")
    if sheet is None:
        return []

    matrix, base_row, base_col, _, last_col = get_used_range_matrix(sheet)
    labels = build_label_positions(matrix, base_row, base_col)
    anchor = find_anchor_max(labels)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)
    history = collect_numeric_pairs(sheet, x_col, y_col, base_row, max(base_row, anchor_row - 1))
    if len(history) < 2:
        return []

    num_quarters_input = find_num_quarters_input_cell(labels, anchor)
    max_label_value = numeric_value_right_of(sheet, anchor_row, anchor_col)
    min_label = find_nearest_label(labels, ["min"], anchor)
    forecast_keywords = ["tot fcst w/o sa", "tot fcst without sa", "forecast"]
    actual_keywords = ["actual", "reported sales"]

    scratch_col = last_col + 3
    intercept_cell = sheet.range((anchor_row, scratch_col))
    slope_cell = sheet.range((anchor_row + 1, scratch_col))

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    max_quarters = min(10, len(history))
    for count in range(2, max_quarters + 1):
        segment = history[-count:]
        start_row = segment[0][0]
        end_row = segment[-1][0]

        if num_quarters_input is not None:
            sheet.range(num_quarters_input).value = count

        intercept_formula = (
            f"=INTERCEPT("
            f"R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE("
            f"R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_r1c1_formula2(intercept_cell, intercept_formula)
        set_r1c1_formula2(slope_cell, slope_formula)
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = get_label_numeric_value(sheet, labels, forecast_keywords, anchor)
        actual_value = get_label_numeric_value(sheet, labels, actual_keywords, anchor)
        forecast_min = numeric_value_right_of(sheet, min_label[0], min_label[1]) if min_label else None
        forecast_max = max_label_value
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            round_for_signature(forecast_value),
            round_for_signature(forecast_max),
            round_for_signature(forecast_min),
            round_for_signature(intercept),
            round_for_signature(slope),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": count,
                "num_quarters_used": count,
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

    try:
        intercept_cell.clear_contents()
        slope_cell.clear_contents()
    except Exception:
        pass
    return rows


def get_unique_output_path(input_folder: Path, destination_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM.xlsx"
    candidate = destination_folder / base_name
    suffix = 1
    while candidate.exists():
        candidate = destination_folder / f"{input_folder.name}_PARAM.{suffix}.xlsx"
        suffix += 1
    return candidate


def apply_standard_sheet_formatting(worksheet: Any) -> None:
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"

    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    max_preview_rows = min(worksheet.max_row, 2000)
    for col_idx in range(1, worksheet.max_column + 1):
        max_len = 0
        for row_idx in range(1, max_preview_rows + 1):
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)


def write_output_workbook(
    destination: Path,
    empirical_rows: Sequence[dict[str, Any]],
    regression_rows: Sequence[dict[str, Any]],
) -> None:
    out_wb = Workbook()

    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    empirical_ws.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_ws.append([row.get(header) for header in EMPIRICAL_HEADERS])
    apply_standard_sheet_formatting(empirical_ws)

    regression_ws = out_wb.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_ws.append([row.get(header) for header in REGRESSION_HEADERS])
    apply_standard_sheet_formatting(regression_ws)

    out_wb.save(destination)


def main() -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_unique_output_path(input_dir, output_dir)

    if not input_dir.exists():
        print(f"skipped: {input_dir} (input_dir does not exist)")
        print(f"output path: {output_path}")
        print("number of files processed: 0")
        print("number of empirical rows: 0")
        print("number of regression rows: 0")
        return

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue

            print(f"processing: {file_path.name}")
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path.name)
                empirical_rows.extend(process_empirical_sheet(workbook, file_path.name, metadata))
                regression_rows.extend(process_regression_sheet(workbook, file_path.name, metadata))
                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                close_workbook_safe(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
