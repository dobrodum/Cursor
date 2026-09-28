#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------- user inputs --------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")
# -----------------------------

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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    text = value.strip().lower()
    text = re.sub(r"[%/()\-]+", " ", text)
    text = re.sub(r"[_\s]+", " ", text).strip()
    return text


def to_rows(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text == "":
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def clean_cell_value(value: Any) -> Any:
    number = to_number(value)
    if number is not None:
        if number.is_integer():
            return int(number)
        return number
    if value is None:
        return ""
    return value


def subtract_if_numbers(left: Any, right: Any) -> Any:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def build_header_map(sheet: xw.Sheet, header_row: int, first_col: int, last_col: int) -> dict[str, int]:
    if last_col < first_col:
        return {}
    row_values = sheet.range((header_row, first_col), (header_row, last_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]

    header_map: dict[str, int] = {}
    for idx, raw in enumerate(row_values):
        norm = normalize_text(raw)
        if norm:
            header_map[norm] = first_col + idx
    return header_map


def find_header_column(header_map: dict[str, int], aliases: Iterable[str]) -> Optional[int]:
    normalized_aliases = [normalize_text(alias) for alias in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for alias in normalized_aliases:
        for header, col in header_map.items():
            if alias and alias in header:
                return col
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[tuple[int, int]]:
    used = sheet.used_range
    values = to_rows(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    target = normalize_text(anchor_text)

    for row_idx, row_vals in enumerate(values):
        for col_idx, value in enumerate(row_vals):
            if normalize_text(value) == target:
                return start_row + row_idx, start_col + col_idx
    return None


def safe_set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def get_sheet_bounds(sheet: xw.Sheet) -> tuple[int, int, int, int]:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    last_row = first_row + used.rows.count - 1
    last_col = first_col + used.columns.count - 1
    return first_row, last_row, first_col, last_col


def parse_file_label(file_name: str) -> Optional[FileLabel]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    if len(parts) < 3:
        return None

    ticker = parts[-2].upper()
    tail = re.sub(r"(?i)_send$", "", parts[-1]).strip()
    match = re.search(r"(?i)(early|mid|late)([a-z]{3,})(\d{4})", tail)
    if not match:
        return None

    phase = match.group(1).lower()
    month_token = match.group(2)[:3].title()
    year = int(match.group(3))
    try:
        month_num = datetime.strptime(month_token, "%b").month
    except ValueError:
        return None

    day = PHASE_TO_DAY[phase]
    model_period = f"{phase.title()}{month_token}_{year}"
    model_date = f"{year:04d}-{month_num:02d}-{day:02d}"
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    index = 1
    while candidate.exists():
        candidate = out_dir / f"{base}.{index}.xlsx"
        index += 1
    return candidate


def safe_close_workbook(wb: xw.Book) -> None:
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
    except Exception:
        pass


def extract_empirical_candidates(
    wb: xw.Book,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    first_row, last_row, first_col, last_col = get_sheet_bounds(sheet)
    header_map = build_header_map(sheet, anchor_row, first_col, last_col)

    min_col = find_header_column(header_map, ["min"]) or (anchor_col + 1)
    forecast_col = find_header_column(
        header_map,
        ["estimated total sold", "forecast value", "forecast", "tot fcst w/o sa", "tot fcst"],
    )
    actual_col = find_header_column(header_map, ["reported sales", "actual value", "actual sales", "actual"])
    num_quarters_col = find_header_column(
        header_map,
        ["num quarters used", "quarters used", "num quarters", "n quarters"],
    )
    last_quarter_col = find_header_column(header_map, ["last quarter used", "last quarter"])
    avg_penetration_col = find_header_column(
        header_map,
        ["avg penetration pct", "avg penetration", "average penetration", "penetration avg"],
    )
    quarterly_sales_col = find_header_column(header_map, ["quarterly sales"])
    reported_sales_col = find_header_column(header_map, ["reported sales"])
    growth_rate_col = find_header_column(header_map, ["growth rate pct", "growth rate"])
    captured_col = find_header_column(
        header_map,
        ["sales captured in db pct", "sales captured in db", "captured in db"],
    )
    penetration_series_col = find_header_column(
        header_map,
        ["penetration pct", "penetration", "sales penetration"],
    )

    # Fall back to anchor offsets when no labeled columns are found.
    if forecast_col is None:
        forecast_col = anchor_col - 2
    if actual_col is None:
        actual_col = anchor_col - 3
    if num_quarters_col is None:
        num_quarters_col = anchor_col - 5
    if last_quarter_col is None:
        last_quarter_col = anchor_col - 6

    rows: list[dict[str, Any]] = []
    temp_avg_cell = sheet.range((anchor_row, anchor_col + 8))
    history_end_row = anchor_row - 1

    for offset in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + offset
        if row_idx > last_row:
            break

        num_quarters_used = clean_cell_value(sheet.range((row_idx, num_quarters_col)).value) if num_quarters_col else offset
        if num_quarters_used == "":
            num_quarters_used = offset

        avg_penetration = ""
        if penetration_series_col and history_end_row >= first_row:
            start_row = max(history_end_row - int(offset) + 1, first_row)
            avg_formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{penetration_series_col}:'
                f"R{history_end_row}C{penetration_series_col}),\"\")"
            )
            safe_set_formula2(temp_avg_cell, avg_formula)
            wb.app.calculate()
            avg_penetration = clean_cell_value(temp_avg_cell.value)
        elif avg_penetration_col:
            avg_penetration = clean_cell_value(sheet.range((row_idx, avg_penetration_col)).value)

        forecast_value = clean_cell_value(sheet.range((row_idx, forecast_col)).value)
        actual_value = clean_cell_value(sheet.range((row_idx, actual_col)).value)
        forecast_max = clean_cell_value(sheet.range((row_idx, anchor_col)).value)
        forecast_min = clean_cell_value(sheet.range((row_idx, min_col)).value)
        range_width = subtract_if_numbers(forecast_max, forecast_min)
        last_quarter_used = clean_cell_value(sheet.range((row_idx, last_quarter_col)).value) if last_quarter_col else ""
        quarterly_sales = clean_cell_value(sheet.range((row_idx, quarterly_sales_col)).value) if quarterly_sales_col else ""
        reported_sales = clean_cell_value(sheet.range((row_idx, reported_sales_col)).value) if reported_sales_col else actual_value
        growth_rate = clean_cell_value(sheet.range((row_idx, growth_rate_col)).value) if growth_rate_col else ""
        sales_captured = clean_cell_value(sheet.range((row_idx, captured_col)).value) if captured_col else ""

        has_content = any(
            value not in ("", None)
            for value in (forecast_value, actual_value, forecast_max, forecast_min, avg_penetration)
        )
        if not has_content:
            continue

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
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


def extract_regression_candidates(
    wb: xw.Book,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    first_row, last_row, first_col, last_col = get_sheet_bounds(sheet)
    header_map = build_header_map(sheet, anchor_row, first_col, last_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    min_col = find_header_column(header_map, ["min"]) or (anchor_col + 1)
    num_quarters_col = find_header_column(
        header_map,
        ["num quarters used", "quarters used", "num quarters", "n quarters"],
    )
    forecast_total_col = find_header_column(
        header_map,
        ["tot fcst w/o sa", "forecast total without sa", "forecast value", "forecast total", "tot fcst"],
    )
    actual_col = find_header_column(header_map, ["actual value", "actual sales", "reported sales", "actual"])

    if num_quarters_col is None:
        num_quarters_col = anchor_col - 5
    if forecast_total_col is None:
        forecast_total_col = anchor_col - 2

    temp_intercept_cell = sheet.range((anchor_row, anchor_col + 8))
    temp_slope_cell = sheet.range((anchor_row, anchor_col + 9))
    history_end_row = anchor_row - 1

    rows: list[dict[str, Any]] = []
    previous_signature: Optional[tuple[Any, ...]] = None

    for offset in range(1, N_QUARTERS + 1):
        if history_end_row < first_row:
            break
        start_row = max(history_end_row - offset + 1, first_row)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )

        safe_set_formula2(temp_intercept_cell, intercept_formula)
        safe_set_formula2(temp_slope_cell, slope_formula)
        wb.app.calculate()

        intercept = clean_cell_value(temp_intercept_cell.value)
        slope = clean_cell_value(temp_slope_cell.value)

        row_idx = anchor_row + offset
        if row_idx > last_row:
            break

        num_quarters_used = clean_cell_value(sheet.range((row_idx, num_quarters_col)).value) if num_quarters_col else offset
        if num_quarters_used == "":
            num_quarters_used = offset

        forecast_value = clean_cell_value(sheet.range((row_idx, forecast_total_col)).value)
        forecast_max = clean_cell_value(sheet.range((row_idx, anchor_col)).value)
        forecast_min = clean_cell_value(sheet.range((row_idx, min_col)).value)
        actual_value = clean_cell_value(sheet.range((row_idx, actual_col)).value) if actual_col else ""
        range_width = subtract_if_numbers(forecast_max, forecast_min)

        signature = (num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope)
        if signature == previous_signature:
            continue
        previous_signature = signature

        has_content = any(
            value not in ("", None)
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        )
        if not has_content:
            continue

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
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 45)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_source_files(directory: Path) -> Iterable[Path]:
    for file_path in sorted(directory.iterdir()):
        if not file_path.is_file():
            continue
        yield file_path


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False
    app.calculation = "manual"

    try:
        for file_path in iter_source_files(input_dir):
            file_name = file_path.name

            if file_name.startswith("~"):
                print(f"Skipped {file_name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_name}: not an .xlsx file")
                continue

            label = parse_file_label(file_name)
            if label is None:
                print(f"Skipped {file_name}: filename does not match expected model label pattern")
                continue

            print(f"Processing {file_name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(wb, label, file_name))
                regression_rows.extend(extract_regression_candidates(wb, label, file_name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_name}: processing error: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
