#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---- user inputs ----
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"


EMPIRICAL_COLUMNS: List[str] = [
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


REGRESSION_COLUMNS: List[str] = [
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
class ModelLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetAnchor:
    row: int
    col: int
    header_map: Dict[str, int]
    last_col: int


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9%]+", "_", text)
    return text.strip("_")


def coerce_matrix(values: object) -> List[List[object]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: object) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def subtract(a: object, b: object) -> Optional[float]:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def is_blank(value: object) -> bool:
    return value in (None, "")


def find_sheet_anchor(sheet: xw.Sheet, token: str = "max") -> Optional[SheetAnchor]:
    used = sheet.used_range
    values = coerce_matrix(used.options(ndim=2).value)
    if not values:
        return None

    first_row = used.row
    first_col = used.column
    last_col = used.last_cell.column
    token_key = normalize_header(token)

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_header(cell_value) == token_key:
                anchor_row = first_row + r_idx
                anchor_col = first_col + c_idx
                header_map: Dict[str, int] = {}
                for header_c_idx, header_value in enumerate(row_values):
                    key = normalize_header(header_value)
                    if key:
                        header_map[key] = first_col + header_c_idx
                return SheetAnchor(
                    row=anchor_row,
                    col=anchor_col,
                    header_map=header_map,
                    last_col=last_col,
                )
    return None


def resolve_col(
    header_map: Dict[str, int],
    aliases: Sequence[str],
    fallback_col: int,
) -> int:
    for alias in aliases:
        key = normalize_header(alias)
        if key in header_map:
            return header_map[key]
    return fallback_col


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def parse_filename_labels(file_path: Path) -> ModelLabel:
    stem = file_path.stem
    ticker = "UNKNOWN"
    model_period = "Unknown_0000"
    model_date = ""

    ticker_match = re.search(r"\s-\s([A-Za-z0-9]+)\s-\s", stem)
    if ticker_match:
        ticker = ticker_match.group(1).upper()

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.capitalize()
        month_abbr = month_raw.capitalize()
        year = int(year_raw)
        day = {"Early": 5, "Mid": 15, "Late": 25}[phase]
        month = {
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
        }[month_abbr]
        model_period = f"{phase}{month_abbr}_{year}"
        model_date = date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}"
    return ModelLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def safe_close_workbook(wb: xw.Book) -> None:
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


def block_lookup(
    block_row: Sequence[object],
    first_col: int,
    col_index: int,
) -> object:
    return block_row[col_index - first_col]


def process_empirical_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    label: ModelLabel,
    source_file: str,
) -> List[Dict[str, object]]:
    anchor = find_sheet_anchor(sheet, token="max")
    if anchor is None:
        return []

    n_quarters = 10
    start_row = anchor.row + 1
    end_row = start_row + n_quarters - 1

    col_num_q = resolve_col(
        anchor.header_map,
        ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
        anchor.col - 9,
    )
    col_last_q = resolve_col(
        anchor.header_map,
        ["last_quarter_used", "last_quarter"],
        anchor.col - 8,
    )
    col_forecast = resolve_col(
        anchor.header_map,
        ["estimated_total_sold", "forecast_value", "forecast"],
        anchor.col - 1,
    )
    col_actual = resolve_col(
        anchor.header_map,
        ["reported_sales", "actual_value", "actual"],
        anchor.col - 2,
    )
    col_max = resolve_col(anchor.header_map, ["forecast_max", "max"], anchor.col)
    col_min = resolve_col(anchor.header_map, ["forecast_min", "min"], anchor.col + 1)
    col_quarterly_sales = resolve_col(
        anchor.header_map,
        ["quarterly_sales", "quarter_sales", "sales_quarter"],
        anchor.col - 3,
    )
    col_reported_sales = resolve_col(
        anchor.header_map,
        ["reported_sales", "sales_reported"],
        anchor.col - 2,
    )
    col_growth_rate = resolve_col(
        anchor.header_map,
        ["growth_rate_pct", "growth_rate"],
        anchor.col - 4,
    )
    col_sales_captured = resolve_col(
        anchor.header_map,
        ["sales_captured_in_db_pct", "sales_captured_pct", "captured_pct"],
        anchor.col - 5,
    )
    col_penetration_history = resolve_col(
        anchor.header_map,
        ["penetration_pct", "avg_penetration_pct", "penetration"],
        anchor.col - 6,
    )

    scratch_avg_col = max(anchor.last_col, anchor.col) + 3
    history_end_row = anchor.row - 1
    for idx in range(n_quarters):
        row_num = start_row + idx
        quarters_used = idx + 1
        history_start_row = history_end_row - quarters_used + 1
        if history_start_row < 1 or history_end_row < history_start_row:
            formula = '=""'
        else:
            formula = (
                f'=IFERROR(AVERAGE(R{history_start_row}C{col_penetration_history}:'
                f'R{history_end_row}C{col_penetration_history}),"")'
            )
        set_formula2(sheet.range((row_num, scratch_avg_col)), formula)

    wb.app.calculate()

    cols_to_read = sorted(
        {
            col_num_q,
            col_last_q,
            col_forecast,
            col_actual,
            col_max,
            col_min,
            col_quarterly_sales,
            col_reported_sales,
            col_growth_rate,
            col_sales_captured,
            scratch_avg_col,
        }
    )
    first_col = cols_to_read[0]
    last_col = cols_to_read[-1]
    block = coerce_matrix(
        sheet.range((start_row, first_col), (end_row, last_col))
        .options(ndim=2)
        .value
    )

    rows: List[Dict[str, object]] = []
    for idx, block_row in enumerate(block):
        quarters_used_val = block_lookup(block_row, first_col, col_num_q)
        quarters_used_num = to_float(quarters_used_val)
        quarters_used = int(quarters_used_num) if quarters_used_num is not None else idx + 1
        last_quarter_used = block_lookup(block_row, first_col, col_last_q)
        forecast_value = block_lookup(block_row, first_col, col_forecast)
        actual_value = block_lookup(block_row, first_col, col_actual)
        forecast_max = block_lookup(block_row, first_col, col_max)
        forecast_min = block_lookup(block_row, first_col, col_min)
        avg_penetration_pct = block_lookup(block_row, first_col, scratch_avg_col)
        quarterly_sales = block_lookup(block_row, first_col, col_quarterly_sales)
        reported_sales = block_lookup(block_row, first_col, col_reported_sales)
        growth_rate_pct = block_lookup(block_row, first_col, col_growth_rate)
        sales_captured_in_db_pct = block_lookup(block_row, first_col, col_sales_captured)

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def process_regression_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    label: ModelLabel,
    source_file: str,
) -> List[Dict[str, object]]:
    anchor = find_sheet_anchor(sheet, token="max")
    if anchor is None:
        return []

    n_quarters = 10
    start_row = anchor.row + 1
    end_row = start_row + n_quarters - 1

    # Required by spec.
    y_col = anchor.col - 7
    x_col = anchor.col - 11

    col_num_q = resolve_col(
        anchor.header_map,
        ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
        anchor.col - 9,
    )
    col_forecast = resolve_col(
        anchor.header_map,
        ["tot_fcst_w_o_sa", "tot_fcst_wo_sa", "forecast_value", "forecast"],
        anchor.col - 1,
    )
    col_actual = resolve_col(
        anchor.header_map,
        ["actual_value", "actual", "reported_sales"],
        anchor.col - 2,
    )
    col_max = resolve_col(anchor.header_map, ["forecast_max", "max"], anchor.col)
    col_min = resolve_col(anchor.header_map, ["forecast_min", "min"], anchor.col + 1)

    scratch_intercept_col = max(anchor.last_col, anchor.col) + 3
    scratch_slope_col = scratch_intercept_col + 1
    history_end_row = anchor.row - 1

    for idx in range(n_quarters):
        row_num = start_row + idx
        quarters_used = idx + 1
        history_start_row = history_end_row - quarters_used + 1
        if history_start_row < 1 or history_end_row < history_start_row:
            intercept_formula = '=""'
            slope_formula = '=""'
        else:
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},'
                f'R{history_start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},'
                f'R{history_start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
            )

        set_formula2(sheet.range((row_num, scratch_intercept_col)), intercept_formula)
        set_formula2(sheet.range((row_num, scratch_slope_col)), slope_formula)

    wb.app.calculate()

    cols_to_read = sorted(
        {
            col_num_q,
            col_forecast,
            col_actual,
            col_max,
            col_min,
            scratch_intercept_col,
            scratch_slope_col,
        }
    )
    first_col = cols_to_read[0]
    last_col = cols_to_read[-1]
    block = coerce_matrix(
        sheet.range((start_row, first_col), (end_row, last_col))
        .options(ndim=2)
        .value
    )

    rows: List[Dict[str, object]] = []
    previous_signature: Optional[Tuple[object, ...]] = None
    for idx, block_row in enumerate(block):
        quarters_used_val = block_lookup(block_row, first_col, col_num_q)
        quarters_used_num = to_float(quarters_used_val)
        quarters_used = int(quarters_used_num) if quarters_used_num is not None else idx + 1
        forecast_value = block_lookup(block_row, first_col, col_forecast)
        actual_value = block_lookup(block_row, first_col, col_actual)
        forecast_max = block_lookup(block_row, first_col, col_max)
        forecast_min = block_lookup(block_row, first_col, col_min)
        intercept_val = block_lookup(block_row, first_col, scratch_intercept_col)
        slope_val = block_lookup(block_row, first_col, scratch_slope_col)

        if all(
            is_blank(v)
            for v in (forecast_value, forecast_max, forecast_min, intercept_val, slope_val)
        ):
            continue

        signature = (
            quarters_used,
            intercept_val,
            slope_val,
            forecast_value,
            forecast_max,
            forecast_min,
        )
        if idx == n_quarters - 1 and previous_signature == signature:
            continue
        previous_signature = signature

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": quarters_used,
            "num_quarters_used": quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract(forecast_max, forecast_min),
            "intercept": intercept_val,
            "slope": slope_val,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(
    ws,
    columns: Sequence[str],
    rows: Iterable[Dict[str, object]],
) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    bold_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for r_idx in range(2, ws.max_row + 1):
            value = ws.cell(r_idx, idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def build_output_path(input_dir_path: Path, output_dir_path: Path) -> Path:
    input_folder_name = input_dir_path.resolve().name
    stem = f"{input_folder_name}_PARAM"

    candidate = output_dir_path / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_dir_path / f"{stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def main() -> None:
    input_dir_path = Path(input_dir).expanduser().resolve()
    output_dir_path = Path(output_dir).expanduser().resolve()

    if not input_dir_path.exists() or not input_dir_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_dir_path}")

    output_dir_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_dir_path, output_dir_path)

    empirical_rows: List[Dict[str, object]] = []
    regression_rows: List[Dict[str, object]] = []

    all_items = sorted(input_dir_path.iterdir(), key=lambda p: p.name.lower())
    source_files = [p for p in all_items if p.is_file()]

    processed_files = 0
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"Processing file: {file_path.name}")
            label = parse_filename_labels(file_path)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                try:
                    empirical_sheet = wb.sheets["Empirical Model"]
                except Exception:
                    empirical_sheet = None
                try:
                    regression_sheet = wb.sheets["Regression Model"]
                except Exception:
                    regression_sheet = None

                if empirical_sheet is None and regression_sheet is None:
                    print(f"Skipped file: {file_path.name} (missing both required sheets)")
                    continue

                if empirical_sheet is not None:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            wb=wb,
                            sheet=empirical_sheet,
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"Skipped empirical extraction for {file_path.name} (sheet missing)")

                if regression_sheet is not None:
                    regression_rows.extend(
                        process_regression_sheet(
                            wb=wb,
                            sheet=regression_sheet,
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"Skipped regression extraction for {file_path.name} (sheet missing)")

                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_file)

    print(f"Output path: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
