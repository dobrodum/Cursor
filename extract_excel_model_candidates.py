from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from statistics import stdev
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------------------------
# Configure these two paths
# -------------------------
input_dir = r"/path/to/input_folder"
output_dir = r"/path/to/output_folder"

# -------------------------
# Constants
# -------------------------
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

# Empirical anchor-based offsets (relative to "max" anchor column).
# If your template changes, update these offsets only.
EMP_COL_OFFSETS = {
    "last_quarter_used": -12,
    "quarterly_sales": -11,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
}

MODEL_PERIOD_RE = re.compile(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", re.IGNORECASE)
DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "").replace("%", "")
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def parse_file_label(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""

    period_source = ""
    if len(parts) >= 3:
        # Example: MidJan2026_Send -> MidJan2026
        period_source = re.split(r"[_\s-]+", parts[2])[0]

    match = MODEL_PERIOD_RE.search(period_source) or MODEL_PERIOD_RE.search(stem)
    model_period = ""
    model_date = ""
    if match:
        phase = match.group(1).title()
        month_text = match.group(2)[:3].title()
        year = int(match.group(3))
        try:
            month_num = datetime.strptime(month_text, "%b").month
            day = DAY_MAP[phase.lower()]
            model_period = f"{phase}{month_text}_{year}"
            model_date = f"{year:04d}-{month_num:02d}-{day:02d}"
        except ValueError:
            model_period = ""
            model_date = ""

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or stem)
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_folder / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_max_anchor(sheet: xw.Sheet) -> tuple[int | None, int | None]:
    used = sheet.used_range
    grid = to_2d(used.value)
    if not grid:
        return None, None

    start_row = used.row
    start_col = used.column
    for r_idx, row in enumerate(grid):
        for c_idx, cell in enumerate(row):
            if normalize_text(cell) == "max":
                return start_row + r_idx, start_col + c_idx
    return None, None


def close_workbook_safely(wb: xw.Book) -> None:
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

    # Last resort (still no save call).
    try:
        wb.close()
    except Exception:
        return


def set_formula2_r1c1(target_range: xw.Range, formula_r1c1: str) -> None:
    # Preferred path for R1C1 + Formula2.
    try:
        target_range.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    # Fallback if Formula2R1C1 isn't exposed.
    target_range.formula2 = formula_r1c1


def read_block(
    sheet: xw.Sheet, start_row: int, end_row: int, columns: list[int]
) -> tuple[dict[tuple[int, int], Any], int, int]:
    min_col = min(columns)
    max_col = max(columns)
    block_values = to_2d(sheet.range((start_row, min_col), (end_row, max_col)).value)
    cache: dict[tuple[int, int], Any] = {}
    for r_offset, row_values in enumerate(block_values):
        absolute_row = start_row + r_offset
        for c in columns:
            cache[(absolute_row, c)] = row_values[c - min_col]
    return cache, min_col, max_col


def cached_value(cache: dict[tuple[int, int], Any], row: int, col: int) -> Any:
    return cache.get((row, col))


def extract_empirical_candidates(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: sheet 'Empirical Model' not found")
        return rows

    anchor_row, anchor_col = find_max_anchor(sheet)
    if anchor_row is None or anchor_col is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return rows

    quarter_col = anchor_col + EMP_COL_OFFSETS["last_quarter_used"]
    quarterly_sales_col = anchor_col + EMP_COL_OFFSETS["quarterly_sales"]
    reported_sales_col = anchor_col + EMP_COL_OFFSETS["reported_sales"]
    growth_rate_col = anchor_col + EMP_COL_OFFSETS["growth_rate_pct"]
    captured_col = anchor_col + EMP_COL_OFFSETS["sales_captured_in_db_pct"]

    if min(quarter_col, quarterly_sales_col, reported_sales_col, growth_rate_col, captured_col) < 1:
        print(f"Skipped empirical extraction for {source_file}: invalid anchor offsets")
        return rows

    data_end_row = anchor_row - 1
    data_start_row = max(1, data_end_row - 80)
    if data_start_row > data_end_row:
        print(f"Skipped empirical extraction for {source_file}: no rows above empirical anchor")
        return rows

    col_list = [quarter_col, quarterly_sales_col, reported_sales_col, growth_rate_col, captured_col]
    cache, _, max_col = read_block(sheet, data_start_row, data_end_row, col_list)

    valid_rows: list[int] = []
    for r in range(data_start_row, data_end_row + 1):
        q_sales = to_float(cached_value(cache, r, quarterly_sales_col))
        rep_sales = to_float(cached_value(cache, r, reported_sales_col))
        if q_sales is not None and rep_sales is not None and rep_sales != 0:
            valid_rows.append(r)

    if not valid_rows:
        print(f"Skipped empirical extraction for {source_file}: no numeric empirical history found")
        return rows

    temp_col = max(max_col + 2, anchor_col + 2)
    temp_avg_cell = sheet.range((anchor_row + 2, temp_col))
    temp_fcst_cell = sheet.range((anchor_row + 3, temp_col))

    max_iterations = min(N_QUARTERS, len(valid_rows))
    for n_quarters in range(1, max_iterations + 1):
        used_rows = valid_rows[-n_quarters:]
        first_row = used_rows[0]
        last_row = used_rows[-1]

        avg_pen_formula = (
            "=IFERROR("
            f"SUM(R{first_row}C{quarterly_sales_col}:R{last_row}C{quarterly_sales_col})/"
            f"SUM(R{first_row}C{reported_sales_col}:R{last_row}C{reported_sales_col}),0)"
        )
        forecast_formula = (
            "=IFERROR("
            f"R{last_row}C{quarterly_sales_col}/R{temp_avg_cell.row}C{temp_avg_cell.column},0)"
        )

        set_formula2_r1c1(temp_avg_cell, avg_pen_formula)
        set_formula2_r1c1(temp_fcst_cell, forecast_formula)
        wb.app.calculate()

        avg_penetration_pct = to_float(temp_avg_cell.value)
        forecast_value = to_float(temp_fcst_cell.value)

        sample_forecasts: list[float] = []
        if avg_penetration_pct and avg_penetration_pct != 0:
            for r in used_rows:
                q_val = to_float(cached_value(cache, r, quarterly_sales_col))
                if q_val is not None:
                    sample_forecasts.append(q_val / avg_penetration_pct)

        forecast_max = max(sample_forecasts) if sample_forecasts else None
        forecast_min = min(sample_forecasts) if sample_forecasts else None
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        quarterly_sales = to_float(cached_value(cache, last_row, quarterly_sales_col))
        reported_sales = to_float(cached_value(cache, last_row, reported_sales_col))
        growth_rate = to_float(cached_value(cache, last_row, growth_rate_col))
        sales_captured = to_float(cached_value(cache, last_row, captured_col))
        last_quarter_used = cached_value(cache, last_row, quarter_col)

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: sheet 'Regression Model' not found")
        return rows

    anchor_row, anchor_col = find_max_anchor(sheet)
    if anchor_row is None or anchor_col is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return rows

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if min(y_col, x_col) < 1:
        print(f"Skipped regression extraction for {source_file}: invalid anchor offsets")
        return rows

    data_end_row = anchor_row - 1
    data_start_row = max(1, data_end_row - 120)
    if data_start_row > data_end_row:
        print(f"Skipped regression extraction for {source_file}: no rows above regression anchor")
        return rows

    cache, _, max_col = read_block(sheet, data_start_row, data_end_row, [x_col, y_col])

    valid_rows: list[int] = []
    for r in range(data_start_row, data_end_row + 1):
        x_val = to_float(cached_value(cache, r, x_col))
        y_val = to_float(cached_value(cache, r, y_col))
        if x_val is not None and y_val is not None:
            valid_rows.append(r)

    if len(valid_rows) < 2:
        print(f"Skipped regression extraction for {source_file}: not enough numeric regression history")
        return rows

    temp_col = max(max_col + 2, anchor_col + 2)
    temp_intercept = sheet.range((anchor_row + 2, temp_col))
    temp_slope = sheet.range((anchor_row + 3, temp_col))

    previous_signature: tuple[Any, ...] | None = None

    max_iterations = min(N_QUARTERS, len(valid_rows))
    for n_quarters in range(1, max_iterations + 1):
        used_rows = valid_rows[-n_quarters:]
        if len(used_rows) < 2:
            continue

        first_row = used_rows[0]
        last_row = used_rows[-1]

        intercept_formula = (
            "=INTERCEPT("
            f"R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            "=SLOPE("
            f"R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )

        set_formula2_r1c1(temp_intercept, intercept_formula)
        set_formula2_r1c1(temp_slope, slope_formula)
        wb.app.calculate()

        intercept = to_float(temp_intercept.value)
        slope = to_float(temp_slope.value)
        if intercept is None or slope is None:
            continue

        x_values: list[float] = []
        y_values: list[float] = []
        for r in used_rows:
            x_val = to_float(cached_value(cache, r, x_col))
            y_val = to_float(cached_value(cache, r, y_col))
            if x_val is None or y_val is None:
                continue
            x_values.append(x_val)
            y_values.append(y_val)

        if len(x_values) < 2 or len(y_values) < 2:
            continue

        x_next = to_float(sheet.cells(last_row + 1, x_col).value)
        if x_next is None:
            x_next = x_values[-1] + 1

        forecast_total_without_sa = intercept + (slope * x_next)

        residuals: list[float] = []
        for x_val, y_val in zip(x_values, y_values):
            residuals.append(y_val - (intercept + slope * x_val))

        margin = 0.0
        if len(residuals) > 1:
            try:
                margin = 1.96 * stdev(residuals)
            except Exception:
                margin = 0.0

        forecast_max = forecast_total_without_sa + margin
        forecast_min = forecast_total_without_sa - margin
        range_width = forecast_max - forecast_min

        actual_next = to_float(sheet.cells(last_row + 1, y_col).value)

        signature = (
            round(intercept, 8),
            round(slope, 8),
            round(forecast_total_without_sa, 8),
            round(forecast_max, 8),
            round(forecast_min, 8),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_next,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    wb: Workbook, sheet_name: str, headers: list[str], records: list[dict[str, Any]]
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(headers)
    for record in records:
        ws.append([record.get(h) for h in headers])

    # Header styling and filters.
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Reasonable width calculation.
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            display = f"{value:.6g}" if isinstance(value, float) else str(value)
            max_len = max(max_len, len(display))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    # Remove the default empty sheet so only the two required sheets exist.
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    wb.save(output_path)


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary Excel file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"Processing file: {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_label(file_path.name)

                empirical_rows.extend(
                    extract_empirical_candidates(wb, metadata, file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(wb, metadata, file_path.name)
                )

                files_processed += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
