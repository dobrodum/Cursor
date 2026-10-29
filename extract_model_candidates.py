#!/usr/bin/env python3
"""
Extract empirical and regression candidates from all .xlsx files in input_dir.
"""

from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


input_dir = Path("./input")
output_dir = Path("./output")

N_QUARTERS = 10
MODEL_PHASE_DAYS = {"early": 5, "mid": 15, "late": 25}

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

MONTH_LOOKUP: dict[str, int] = {}
for month_idx in range(1, 13):
    MONTH_LOOKUP[calendar.month_abbr[month_idx].lower()] = month_idx
    MONTH_LOOKUP[calendar.month_name[month_idx].lower()] = month_idx


def normalize_key(value: Any) -> str:
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def is_empty(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        stripped = value.strip()
        return stripped == "" or stripped.startswith("#")
    return False


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def numeric_subtract(left: Any, right: Any) -> float | None:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def coerce_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def read_sheet_snapshot(sheet: xw.Sheet) -> tuple[list[list[Any]], int, int, int, int]:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    row_count = used.rows.count
    col_count = used.columns.count
    last_row = first_row + row_count - 1
    last_col = first_col + col_count - 1
    matrix = coerce_matrix(used.value)
    return matrix, first_row, first_col, last_row, last_col


def find_max_anchor(
    matrix: list[list[Any]], first_row: int, first_col: int
) -> tuple[int, int, int]:
    matches: list[tuple[int, int]] = []
    for r_idx, row in enumerate(matrix):
        for c_idx, cell in enumerate(row):
            if normalize_key(cell) == "max":
                matches.append((r_idx, c_idx))

    if not matches:
        raise ValueError("Could not find 'max' anchor cell.")

    for r_idx, c_idx in matches:
        right = None
        left = None
        if c_idx + 1 < len(matrix[r_idx]):
            right = normalize_key(matrix[r_idx][c_idx + 1])
        if c_idx - 1 >= 0:
            left = normalize_key(matrix[r_idx][c_idx - 1])
        if right == "min" or left == "min":
            return first_row + r_idx, first_col + c_idx, r_idx

    chosen_r, chosen_c = matches[0]
    return first_row + chosen_r, first_col + chosen_c, chosen_r


def build_header_map(
    matrix: list[list[Any]], header_row_idx: int, first_col: int
) -> dict[str, int]:
    if header_row_idx < 0 or header_row_idx >= len(matrix):
        return {}
    header_map: dict[str, int] = {}
    for rel_col, value in enumerate(matrix[header_row_idx]):
        key = normalize_key(value)
        if key and key not in header_map:
            header_map[key] = first_col + rel_col
    return header_map


def resolve_column(
    header_map: dict[str, int], aliases: list[str], default_col: int | None
) -> int | None:
    normalized_aliases = [normalize_key(alias) for alias in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for alias in normalized_aliases:
        alias_tokens = set(alias.split("_"))
        for key, col in header_map.items():
            key_tokens = set(key.split("_"))
            if alias in key or key in alias:
                return col
            if alias_tokens and len(alias_tokens.intersection(key_tokens)) >= 2:
                return col

    return default_col


def safe_cell_value(sheet: xw.Sheet, row: int, col: int | None) -> Any:
    if col is None or row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def parse_month(month_token: str) -> int:
    month_key = re.sub(r"[^A-Za-z]", "", month_token).lower()
    if not month_key:
        return 1
    if month_key in MONTH_LOOKUP:
        return MONTH_LOOKUP[month_key]
    short_key = month_key[:3]
    return MONTH_LOOKUP.get(short_key, 1)


def parse_file_label(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = parts[1].upper() if len(parts) >= 2 else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "unknown_period"
    model_date = ""

    if period_match:
        phase, month_token, year_text = period_match.groups()
        phase_key = phase.lower()
        month_num = parse_month(month_token)
        day_num = MODEL_PHASE_DAYS.get(phase_key, 15)
        year_num = int(year_text)
        model_period = f"{phase.capitalize()}{calendar.month_abbr[month_num]}_{year_num}"
        model_date = date(year_num, month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def close_workbook_safe(wb: xw.Book) -> None:
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


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_stem = f"{in_dir.name}_PARAM"
    primary = out_dir / f"{base_stem}.xlsx"
    if not primary.exists():
        return primary

    suffix = 1
    while True:
        candidate = out_dir / f"{base_stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def empirical_rows_from_sheet(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except KeyError:
        print(f"skipped empirical extraction for {source_file}: missing 'Empirical Model' sheet")
        return []

    matrix, first_row, first_col, _last_row, last_col = read_sheet_snapshot(sheet)
    if not matrix:
        return []

    anchor_row, anchor_col, header_idx = find_max_anchor(matrix, first_row, first_col)
    header_map = build_header_map(matrix, header_idx, first_col)

    num_quarters_col = resolve_column(
        header_map, ["num_quarters_used", "num quarters used", "quarters used"], anchor_col - 2
    )
    last_quarter_col = resolve_column(
        header_map, ["last_quarter_used", "last quarter used", "last quarter"], anchor_col - 3
    )
    forecast_value_col = resolve_column(
        header_map,
        [
            "estimated_total_sold",
            "forecast_value",
            "tot_fcst",
            "total_forecast",
            "forecast_total",
        ],
        anchor_col - 4,
    )
    actual_value_col = resolve_column(
        header_map, ["actual_value", "reported_sales", "actual sales", "reported sales"], anchor_col - 5
    )
    forecast_max_col = resolve_column(header_map, ["forecast_max", "max"], anchor_col)
    forecast_min_col = resolve_column(header_map, ["forecast_min", "min"], anchor_col + 1)
    avg_penetration_col = resolve_column(
        header_map, ["avg_penetration_pct", "avg penetration", "average penetration"], anchor_col - 6
    )
    quarterly_sales_col = resolve_column(
        header_map, ["quarterly_sales", "quarterly sales", "quarter sales"], anchor_col - 8
    )
    reported_sales_col = resolve_column(
        header_map, ["reported_sales", "reported sales", "actual sales"], actual_value_col
    )
    growth_rate_col = resolve_column(
        header_map, ["growth_rate_pct", "growth rate", "growth pct"], anchor_col - 9
    )
    sales_captured_col = resolve_column(
        header_map,
        [
            "sales_captured_in_db_pct",
            "sales captured in db",
            "captured in db",
            "penetration pct",
        ],
        avg_penetration_col,
    )

    start_row = anchor_row + 1
    helper_col = max(last_col + 2, anchor_col + 2)
    penetration_source_col = sales_captured_col if sales_captured_col is not None else avg_penetration_col

    formulas_written = False
    if penetration_source_col is not None:
        for idx in range(N_QUARTERS):
            n_quarters = idx + 1
            row_idx = start_row + idx
            offset = penetration_source_col - helper_col
            if n_quarters == 1:
                formula = f"=RC[{offset}]"
            else:
                formula = f"=AVERAGE(R[-{n_quarters - 1}]C[{offset}]:RC[{offset}])"
            set_formula2(sheet.cells(row_idx, helper_col), formula)
            formulas_written = True

    if formulas_written:
        wb.app.calculate()
        helper_values_raw = sheet.range(
            (start_row, helper_col), (start_row + N_QUARTERS - 1, helper_col)
        ).value
        if isinstance(helper_values_raw, list):
            helper_values = helper_values_raw
        else:
            helper_values = [helper_values_raw]
    else:
        helper_values = [None] * N_QUARTERS

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_idx = start_row + idx
        num_quarters_used = safe_cell_value(sheet, row_idx, num_quarters_col)
        if is_empty(num_quarters_used):
            num_quarters_used = idx + 1

        avg_pen = helper_values[idx] if idx < len(helper_values) else None
        if is_empty(avg_pen):
            avg_pen = safe_cell_value(sheet, row_idx, avg_penetration_col)

        forecast_value = safe_cell_value(sheet, row_idx, forecast_value_col)
        actual_value = safe_cell_value(sheet, row_idx, actual_value_col)
        forecast_max = safe_cell_value(sheet, row_idx, forecast_max_col)
        forecast_min = safe_cell_value(sheet, row_idx, forecast_min_col)
        reported_sales = safe_cell_value(sheet, row_idx, reported_sales_col)
        quarterly_sales = safe_cell_value(sheet, row_idx, quarterly_sales_col)
        growth_rate_pct = safe_cell_value(sheet, row_idx, growth_rate_col)
        sales_captured_in_db_pct = safe_cell_value(sheet, row_idx, sales_captured_col)
        last_quarter_used = safe_cell_value(sheet, row_idx, last_quarter_col)

        if all(
            is_empty(value)
            for value in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_pen,
                reported_sales,
            ]
        ):
            continue

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": numeric_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def regression_rows_from_sheet(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except KeyError:
        print(f"skipped regression extraction for {source_file}: missing 'Regression Model' sheet")
        return []

    matrix, first_row, first_col, _last_row, last_col = read_sheet_snapshot(sheet)
    if not matrix:
        return []

    anchor_row, anchor_col, header_idx = find_max_anchor(matrix, first_row, first_col)
    header_map = build_header_map(matrix, header_idx, first_col)

    num_quarters_col = resolve_column(
        header_map, ["num_quarters_used", "num quarters used", "quarters used"], anchor_col - 2
    )
    forecast_value_col = resolve_column(
        header_map,
        ["tot_fcst_w_o_sa", "tot fcst w/o sa", "tot fcst wo sa", "forecast_total_without_sa"],
        anchor_col - 1,
    )
    actual_value_col = resolve_column(
        header_map, ["actual_value", "reported_sales", "actual sales"], None
    )
    forecast_max_col = resolve_column(header_map, ["forecast_max", "max"], anchor_col)
    forecast_min_col = resolve_column(header_map, ["forecast_min", "min"], anchor_col + 1)

    start_row = anchor_row + 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    intercept_col = max(last_col + 2, anchor_col + 2)
    slope_col = intercept_col + 1

    formulas_written = False
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        row_idx = start_row + idx
        y_offset = y_col - intercept_col
        x_offset = x_col - intercept_col

        if n_quarters == 1:
            intercept_formula = (
                f'=IFERROR(INTERCEPT(RC[{y_offset}]:RC[{y_offset}],RC[{x_offset}]:RC[{x_offset}]),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(RC[{y_offset}]:RC[{y_offset}],RC[{x_offset}]:RC[{x_offset}]),"")'
            )
        else:
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R[-{n_quarters - 1}]C[{y_offset}]:RC[{y_offset}],'
                f'R[-{n_quarters - 1}]C[{x_offset}]:RC[{x_offset}]),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R[-{n_quarters - 1}]C[{y_offset}]:RC[{y_offset}],'
                f'R[-{n_quarters - 1}]C[{x_offset}]:RC[{x_offset}]),"")'
            )

        set_formula2(sheet.cells(row_idx, intercept_col), intercept_formula)
        set_formula2(sheet.cells(row_idx, slope_col), slope_formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()
        intercept_values_raw = sheet.range(
            (start_row, intercept_col), (start_row + N_QUARTERS - 1, intercept_col)
        ).value
        slope_values_raw = sheet.range((start_row, slope_col), (start_row + N_QUARTERS - 1, slope_col)).value
        intercept_values = intercept_values_raw if isinstance(intercept_values_raw, list) else [intercept_values_raw]
        slope_values = slope_values_raw if isinstance(slope_values_raw, list) else [slope_values_raw]
    else:
        intercept_values = [None] * N_QUARTERS
        slope_values = [None] * N_QUARTERS

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx in range(N_QUARTERS):
        row_idx = start_row + idx
        num_quarters_used = safe_cell_value(sheet, row_idx, num_quarters_col)
        if is_empty(num_quarters_used):
            num_quarters_used = idx + 1

        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None
        forecast_value = safe_cell_value(sheet, row_idx, forecast_value_col)
        actual_value = safe_cell_value(sheet, row_idx, actual_value_col)
        forecast_max = safe_cell_value(sheet, row_idx, forecast_max_col)
        forecast_min = safe_cell_value(sheet, row_idx, forecast_min_col)

        if all(
            is_empty(value)
            for value in [forecast_value, intercept, slope, forecast_max, forecast_min]
        ):
            continue

        signature = (
            to_float(num_quarters_used) or num_quarters_used,
            to_float(intercept),
            to_float(slope),
            to_float(forecast_value),
            to_float(forecast_max),
            to_float(forecast_min),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row = {
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
            "range_width": numeric_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(
    workbook: Workbook, sheet_name: str, headers: list[str], rows: list[dict[str, Any]]
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(headers)

    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        print(f"skipped all files: input directory does not exist -> {input_dir}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir, output_dir)

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
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary Excel file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not an .xlsx file)")
                continue

            wb = None
            try:
                metadata = parse_file_label(file_path)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    empirical_rows_from_sheet(wb, metadata, file_path.name)
                )
                regression_rows.extend(
                    regression_rows_from_sheet(wb, metadata, file_path.name)
                )
                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
