#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime guard
    raise SystemExit(
        "xlwings is required. Install it with `pip install xlwings`."
    ) from exc


# ---------------------------------------------------------------------------
# Configure paths here
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


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

N_QUARTERS = 10

# Anchor-based fallback offsets for Empirical Model (relative to "max" anchor)
EMPIRICAL_PENETRATION_COL_OFFSET = -7
EMPIRICAL_QUARTER_COL_OFFSET = -11

# Temp formula cells (relative to anchor) to avoid disturbing model output cells
TEMP_FORMULA_COL_OFFSET = 6
TEMP_FORMULA_ROW_GAP = 1


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            try:
                return float(cleaned[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_output(value: Any) -> Any:
    return "" if value is None else value


def month_to_number(month_token: str) -> int | None:
    token = month_token.strip().lower()
    month_map = {
        "jan": 1,
        "january": 1,
        "feb": 2,
        "february": 2,
        "mar": 3,
        "march": 3,
        "apr": 4,
        "april": 4,
        "may": 5,
        "jun": 6,
        "june": 6,
        "jul": 7,
        "july": 7,
        "aug": 8,
        "august": 8,
        "sep": 9,
        "sept": 9,
        "september": 9,
        "oct": 10,
        "october": 10,
        "nov": 11,
        "november": 11,
        "dec": 12,
        "december": 12,
    }
    return month_map.get(token)


def parse_file_metadata(file_name: str) -> dict[str, str]:
    day_map = {"early": 5, "mid": 15, "late": 25}
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1].strip().upper() if len(parts) >= 2 else ""

    period_token = parts[2].strip() if len(parts) >= 3 else ""
    period_token = re.sub(r"(_|-)?send.*$", "", period_token, flags=re.IGNORECASE)
    period_token = period_token.replace(" ", "")

    model_period = ""
    model_date = ""

    match = re.search(r"(early|mid|late)([A-Za-z]+)(\d{4})", period_token, re.IGNORECASE)
    if match:
        phase_raw, month_raw, year_raw = match.groups()
        phase = phase_raw.capitalize()
        month_clean = month_raw.capitalize()
        month_num = month_to_number(month_clean)
        model_period = f"{phase}{month_clean}_{year_raw}"
        if month_num is not None:
            model_date = str(date(int(year_raw), month_num, day_map[phase_raw.lower()]))

    if not model_period:
        model_period = period_token.replace("-", "_")

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or model_period
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    first = out_dir / f"{base_name}.xlsx"
    if not first.exists():
        return first

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def safe_close_workbook(wb: Any) -> None:
    # Primary path requested by the spec
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Fallback for engines that do not support close(save=False)
    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    # Last-resort close attempt with Excel alerts disabled
    try:
        wb.close()
    except Exception:
        pass


def get_sheet_case_insensitive(wb: Any, target_name: str) -> Any | None:
    normalized_target = normalize_text(target_name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == normalized_target:
            return sheet
    return None


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def scan_sheet_once(sheet: Any) -> tuple[int, int, list[list[Any]]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    return used.row, used.column, values


def find_anchor_max(sheet: Any) -> tuple[int, int] | None:
    row0, col0, values = scan_sheet_once(sheet)
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if normalize_text(cell_value) == "max":
                return row0 + r_idx, col0 + c_idx
    return None


def find_label_positions(values: list[list[Any]], row0: int, col0: int) -> dict[str, list[tuple[int, int]]]:
    positions: dict[str, list[tuple[int, int]]] = {}
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip():
                key = normalize_text(cell_value)
                positions.setdefault(key, []).append((row0 + r_idx, col0 + c_idx))
    return positions


def find_label_cell(
    label_positions: dict[str, list[tuple[int, int]]],
    required_tokens: Iterable[str],
) -> tuple[int, int] | None:
    tokens = [normalize_text(t) for t in required_tokens if t]
    for label_text, coords in label_positions.items():
        if all(token in label_text for token in tokens):
            return coords[0]
    return None


def read_cell_number(sheet: Any, row: int, col: int) -> float | None:
    return to_number(sheet.range((row, col)).value)


def read_value_right_of_label(
    sheet: Any,
    label_positions: dict[str, list[tuple[int, int]]],
    tokens: Iterable[str],
) -> Any:
    cell = find_label_cell(label_positions, tokens)
    if cell is None:
        return None
    row, col = cell
    return sheet.range((row, col + 1)).value


def rc_relative(from_row: int, from_col: int, to_row: int, to_col: int) -> str:
    dr = to_row - from_row
    dc = to_col - from_col
    r_part = "R" if dr == 0 else f"R[{dr}]"
    c_part = "C" if dc == 0 else f"C[{dc}]"
    return f"{r_part}{c_part}"


def set_r1c1_formula2(cell: Any, formula_r1c1: str) -> None:
    # Requested path: R1C1 with .formula2
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    # Safe fallback
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        cell.api.FormulaR1C1 = formula_r1c1


def empirical_rows_from_workbook(wb: Any, file_name: str, meta: dict[str, str]) -> list[dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"skipped sheet: {file_name} (missing 'Empirical Model')")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"skipped sheet: {file_name} (missing 'max' anchor in Empirical Model)")
        return []

    anchor_row, anchor_col = anchor
    row0, col0, values = scan_sheet_once(sheet)
    label_positions = find_label_positions(values, row0, col0)

    # Static/supporting values (read once)
    quarterly_sales = to_number(
        read_value_right_of_label(sheet, label_positions, ("quarterly", "sales"))
    )
    reported_sales = to_number(
        read_value_right_of_label(sheet, label_positions, ("reported", "sales"))
    )
    growth_rate_pct = to_number(
        read_value_right_of_label(sheet, label_positions, ("growth", "rate"))
    )
    sales_captured_pct = to_number(
        read_value_right_of_label(sheet, label_positions, ("captured", "db"))
    )

    max_label_cell = find_label_cell(label_positions, ("max",))
    min_label_cell = find_label_cell(label_positions, ("min",))
    forecast_max_static = None
    forecast_min_static = None
    if max_label_cell:
        forecast_max_static = to_number(sheet.range((max_label_cell[0], max_label_cell[1] + 1)).value)
    if min_label_cell:
        forecast_min_static = to_number(sheet.range((min_label_cell[0], min_label_cell[1] + 1)).value)

    penetration_col = anchor_col + EMPIRICAL_PENETRATION_COL_OFFSET
    quarter_col = anchor_col + EMPIRICAL_QUARTER_COL_OFFSET
    data_end_row = anchor_row - 1

    # Find start row with data in penetration column (single scan)
    numeric_rows: list[int] = []
    for r in range(row0, data_end_row + 1):
        val = read_cell_number(sheet, r, penetration_col)
        if val is not None:
            numeric_rows.append(r)
    if not numeric_rows:
        return []

    first_data_row = min(numeric_rows)
    last_data_row = max(numeric_rows)

    avg_pen_cell = sheet.range((anchor_row, anchor_col + TEMP_FORMULA_COL_OFFSET))

    results: list[dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        start_row = max(first_data_row, last_data_row - n + 1)
        num_quarters = last_data_row - start_row + 1
        if num_quarters <= 0:
            continue

        avg_formula = (
            f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{last_data_row}C{penetration_col}),"")'
        )
        set_r1c1_formula2(avg_pen_cell, avg_formula)
        wb.app.calculate()

        avg_penetration_pct = to_number(avg_pen_cell.value)
        if avg_penetration_pct is None:
            continue

        last_quarter_used = sheet.range((last_data_row, quarter_col)).value
        forecast_value = (
            avg_penetration_pct * quarterly_sales
            if quarterly_sales is not None
            else None
        )
        actual_value = reported_sales

        forecast_max = forecast_max_static
        forecast_min = forecast_min_static
        if forecast_max is None and forecast_value is not None:
            if growth_rate_pct is not None:
                forecast_max = forecast_value * (1.0 + growth_rate_pct)
            else:
                forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            if growth_rate_pct is not None:
                forecast_min = forecast_value * (1.0 - growth_rate_pct)
            else:
                forecast_min = forecast_value

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        results.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": to_output(last_quarter_used),
                "forecast_value": to_output(forecast_value),
                "actual_value": to_output(actual_value),
                "forecast_max": to_output(forecast_max),
                "forecast_min": to_output(forecast_min),
                "range_width": to_output(range_width),
                "avg_penetration_pct": to_output(avg_penetration_pct),
                "quarterly_sales": to_output(quarterly_sales),
                "reported_sales": to_output(reported_sales),
                "growth_rate_pct": to_output(growth_rate_pct),
                "sales_captured_in_db_pct": to_output(sales_captured_pct),
                "source_file": file_name,
            }
        )

    return results


def regression_rows_from_workbook(wb: Any, file_name: str, meta: dict[str, str]) -> list[dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"skipped sheet: {file_name} (missing 'Regression Model')")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"skipped sheet: {file_name} (missing 'max' anchor in Regression Model)")
        return []

    anchor_row, anchor_col = anchor

    # Required anchor-based offsets
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    row0, col0, values = scan_sheet_once(sheet)
    label_positions = find_label_positions(values, row0, col0)

    actual_value = to_number(
        read_value_right_of_label(sheet, label_positions, ("actual",))
    )

    data_end_row = anchor_row - 1
    candidate_rows: list[int] = []
    for r in range(row0, data_end_row + 1):
        y_val = read_cell_number(sheet, r, y_col)
        x_val = read_cell_number(sheet, r, x_col)
        if y_val is not None and x_val is not None:
            candidate_rows.append(r)
    if len(candidate_rows) < 2:
        return []

    first_data_row = min(candidate_rows)
    last_data_row = max(candidate_rows)

    intercept_cell = sheet.range((anchor_row, anchor_col + TEMP_FORMULA_COL_OFFSET))
    slope_cell = sheet.range((anchor_row + TEMP_FORMULA_ROW_GAP, anchor_col + TEMP_FORMULA_COL_OFFSET))
    forecast_cell = sheet.range((anchor_row + (2 * TEMP_FORMULA_ROW_GAP), anchor_col + TEMP_FORMULA_COL_OFFSET))
    max_cell = sheet.range((anchor_row + (3 * TEMP_FORMULA_ROW_GAP), anchor_col + TEMP_FORMULA_COL_OFFSET))
    min_cell = sheet.range((anchor_row + (4 * TEMP_FORMULA_ROW_GAP), anchor_col + TEMP_FORMULA_COL_OFFSET))

    results: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n in range(2, N_QUARTERS + 1):
        start_row = max(first_data_row, last_data_row - n + 1)
        num_quarters_used = last_data_row - start_row + 1
        if num_quarters_used < 2:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{last_data_row}C{y_col},'
            f'R{start_row}C{x_col}:R{last_data_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{last_data_row}C{y_col},'
            f'R{start_row}C{x_col}:R{last_data_row}C{x_col}),"")'
        )
        intercept_ref = rc_relative(
            forecast_cell.row, forecast_cell.column, intercept_cell.row, intercept_cell.column
        )
        slope_ref = rc_relative(
            forecast_cell.row, forecast_cell.column, slope_cell.row, slope_cell.column
        )
        forecast_formula = (
            f'=IFERROR({intercept_ref}+({slope_ref}*R{last_data_row}C{x_col}),"")'
        )
        max_formula = f'=IFERROR(MAX(R{start_row}C{y_col}:R{last_data_row}C{y_col}),"")'
        min_formula = f'=IFERROR(MIN(R{start_row}C{y_col}:R{last_data_row}C{y_col}),"")'

        set_r1c1_formula2(intercept_cell, intercept_formula)
        set_r1c1_formula2(slope_cell, slope_formula)
        set_r1c1_formula2(forecast_cell, forecast_formula)
        set_r1c1_formula2(max_cell, max_formula)
        set_r1c1_formula2(min_cell, min_formula)
        wb.app.calculate()

        intercept = to_number(intercept_cell.value)
        slope = to_number(slope_cell.value)
        forecast_total_without_sa = to_number(forecast_cell.value)
        forecast_max = to_number(max_cell.value)
        forecast_min = to_number(min_cell.value)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            num_quarters_used,
            intercept,
            slope,
            forecast_total_without_sa,
            forecast_max,
            forecast_min,
        )
        if signature == previous_signature:
            # Prevent duplicate final row.
            continue
        previous_signature = signature

        results.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": to_output(forecast_total_without_sa),
                "actual_value": to_output(actual_value),
                "forecast_max": to_output(forecast_max),
                "forecast_min": to_output(forecast_min),
                "range_width": to_output(range_width),
                "intercept": to_output(intercept),
                "slope": to_output(slope),
                "source_file": file_name,
            }
        )

    return results


def write_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([to_output(row.get(col)) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for i, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row in rows:
            value_str = str(to_output(row.get(column_name)))
            if len(value_str) > max_len:
                max_len = len(value_str)
        ws.column_dimensions[get_column_letter(i)].width = min(max_len + 2, 60)


def write_output_workbook(
    path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(path)


def iter_input_files(folder: Path) -> list[Path]:
    return sorted([p for p in folder.iterdir() if p.is_file()])


def main() -> int:
    if not input_dir.exists():
        print(f"input directory not found: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    files = iter_input_files(input_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False

    original_calc_mode = None
    try:
        try:
            original_calc_mode = app.api.Calculation
            # Excel constant: xlCalculationManual = -4135
            app.api.Calculation = -4135
        except Exception:
            original_calc_mode = None

        for file_path in files:
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue

            print(f"processing: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)

                empirical_rows.extend(
                    empirical_rows_from_workbook(wb, file_path.name, metadata)
                )
                regression_rows.extend(
                    regression_rows_from_workbook(wb, file_path.name, metadata)
                )
                processed_files += 1
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if original_calc_mode is not None:
            try:
                app.api.Calculation = original_calc_mode
            except Exception:
                pass
        app.quit()

    output_path = unique_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
