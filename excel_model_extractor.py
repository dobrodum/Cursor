from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


# ---------------------------------------------------------------------------
# Extraction configuration
# ---------------------------------------------------------------------------
N_QUARTERS = 10

# Anchor-based offsets for source model tables.
EMPIRICAL_PENETRATION_COL_OFFSET = -11
EMPIRICAL_SALES_COL_OFFSET = -7
EMPIRICAL_REPORTED_SALES_COL_OFFSET = -6

REGRESSION_X_COL_OFFSET = -11
REGRESSION_Y_COL_OFFSET = -7

# Temporary cells for formula2 writes (relative to the found "max" anchor).
TEMP_FORMULA_ROW_OFFSET = 2
TEMP_FORMULA_COL_OFFSET = 18


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


MONTH_MAP = {
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

PERIOD_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_RE = re.compile(
    r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
    re.IGNORECASE,
)


def log(message: str) -> None:
    print(message, flush=True)


def is_number(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return False


def to_float(value: Any) -> float | None:
    if not is_number(value):
        return None
    return float(value)


def to_percent_display(value: float | None) -> float | None:
    if value is None:
        return None
    return value * 100.0 if abs(value) <= 1 else value


def to_ratio(value: float | None) -> float | None:
    if value is None:
        return None
    if abs(value) > 1:
        return value / 100.0
    return value


def safe_divide(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    try:
        return numerator / denominator
    except ZeroDivisionError:
        return None


def safe_subtract(a: float | None, b: float | None) -> float | None:
    if a is None or b is None:
        return None
    return a - b


def rounded_signature(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value, 10)


def serialize_cell_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def parse_file_metadata(file_path: Path) -> dict[str, str]:
    # Expected style: MedMiner_Model - AORT - MidJan2026_Send.xlsx
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 else ""
    period_chunk = parts[2] if len(parts) > 2 else ""
    period_token = period_chunk.split("_")[0]

    model_period = period_token or "unknown_period"
    model_date = ""

    match = PERIOD_RE.search(period_token)
    if match:
        period_label = match.group(1).title()
        month_abbrev = match.group(2).title()
        year = int(match.group(3))
        month_number = MONTH_MAP[month_abbrev.lower()]
        day = PERIOD_DAY_MAP[period_label]
        model_period = f"{period_label}{month_abbrev}_{year}"
        model_date = date(year, month_number, day).isoformat()

    ticker = ticker or "UNKNOWN"
    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def collect_input_files(in_dir: Path) -> tuple[list[Path], list[tuple[str, str]]]:
    valid: list[Path] = []
    skipped: list[tuple[str, str]] = []

    for path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            skipped.append((path.name, "temporary Excel file"))
            continue
        if path.suffix.lower() != ".xlsx":
            skipped.append((path.name, "not an .xlsx file"))
            continue
        valid.append(path)

    return valid, skipped


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int]:
    used = sheet.used_range
    values = normalize_2d(used.options(ndim=2).value)
    if not values:
        raise ValueError(f"Sheet '{sheet.name}' is empty.")

    row_base = used.row
    col_base = used.column
    best: tuple[int, int] | None = None
    best_score = (-1, -1)

    for r_idx, row_values in enumerate(values):
        for c_idx, raw in enumerate(row_values):
            if not isinstance(raw, str):
                continue
            if raw.strip().lower() != "max":
                continue

            score_has_min_neighbor = 0
            if c_idx + 1 < len(row_values):
                neighbor = row_values[c_idx + 1]
                if isinstance(neighbor, str) and neighbor.strip().lower() == "min":
                    score_has_min_neighbor = 1

            score = (score_has_min_neighbor, r_idx)
            if score > best_score:
                best_score = score
                best = (row_base + r_idx, col_base + c_idx)

    if best is None:
        raise ValueError(f'No "max" anchor found on sheet {sheet.name!r}.')
    return best


def trailing_numeric_block(
    sheet: xw.Sheet, column: int, end_row: int
) -> list[tuple[int, float]]:
    if end_row < 1 or column < 1:
        return []

    values = sheet.range((1, column), (end_row, column)).value
    if not isinstance(values, list):
        values = [values]

    idx = len(values) - 1
    while idx >= 0 and not is_number(values[idx]):
        idx -= 1
    if idx < 0:
        return []

    block: list[tuple[int, float]] = []
    while idx >= 0 and is_number(values[idx]):
        block.append((idx + 1, float(values[idx])))
        idx -= 1

    block.reverse()
    return block


def select_common_rows(
    series_a: list[tuple[int, float]], series_b: list[tuple[int, float]]
) -> list[tuple[int, float, float]]:
    map_a = {row: val for row, val in series_a}
    map_b = {row: val for row, val in series_b}
    rows = sorted(set(map_a).intersection(map_b))
    return [(row, map_a[row], map_b[row]) for row in rows]


def close_workbook_without_saving(wb: xw.Book) -> None:
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

    # Final fallback if platform-specific APIs differ.
    try:
        wb.close()
    except Exception:
        pass


def extract_empirical_rows(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    anchor_row, anchor_col = find_max_anchor(sheet)

    penetration_col = anchor_col + EMPIRICAL_PENETRATION_COL_OFFSET
    sales_col = anchor_col + EMPIRICAL_SALES_COL_OFFSET
    reported_sales_col = anchor_col + EMPIRICAL_REPORTED_SALES_COL_OFFSET

    penetration_series = trailing_numeric_block(sheet, penetration_col, anchor_row - 1)
    sales_series = trailing_numeric_block(sheet, sales_col, anchor_row - 1)
    reported_series = trailing_numeric_block(sheet, reported_sales_col, anchor_row - 1)

    aligned = select_common_rows(penetration_series, sales_series)
    if not aligned:
        return []

    penetration_by_row = {row: penetration for row, penetration, _ in aligned}
    sales_by_row = {row: sales for row, _, sales in aligned}
    reported_by_row = {row: value for row, value in reported_series}
    aligned_rows = [row for row, _, _ in aligned]

    max_n = min(N_QUARTERS, len(aligned_rows))
    if max_n < 1:
        return []

    temp_row = anchor_row + TEMP_FORMULA_ROW_OFFSET
    temp_col = anchor_col + TEMP_FORMULA_COL_OFFSET
    avg_penetration_cell = sheet.cells(temp_row, temp_col)

    rows: list[dict[str, Any]] = []
    for n_used in range(1, max_n + 1):
        window_rows = aligned_rows[-n_used:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        avg_penetration_cell.formula2 = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        )
        wb.app.calculate()

        avg_pen = to_float(avg_penetration_cell.value)
        avg_pen_ratio = to_ratio(avg_pen)

        penetration_window = [penetration_by_row[r] for r in window_rows]
        max_pen = max(penetration_window) if penetration_window else None
        min_pen = min(penetration_window) if penetration_window else None
        max_pen_ratio = to_ratio(max_pen)
        min_pen_ratio = to_ratio(min_pen)

        quarterly_sales = sales_by_row[end_row]
        reported_sales = reported_by_row.get(end_row, quarterly_sales)

        forecast_value = safe_divide(quarterly_sales, avg_pen_ratio)
        forecast_max = safe_divide(quarterly_sales, min_pen_ratio)
        forecast_min = safe_divide(quarterly_sales, max_pen_ratio)
        range_width = safe_subtract(forecast_max, forecast_min)

        prev_sales = sales_by_row.get(window_rows[-2]) if len(window_rows) > 1 else None
        growth_rate_pct = None
        if prev_sales not in (None, 0):
            growth_rate_pct = ((quarterly_sales - prev_sales) / prev_sales) * 100.0

        last_quarter_used = serialize_cell_value(
            sheet.cells(end_row, penetration_col - 1).value
        )
        if last_quarter_used in (None, ""):
            last_quarter_used = end_row

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": to_percent_display(avg_pen),
                "num_quarters_used": n_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,  # estimated total sold
                "actual_value": reported_sales,  # reported sales
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": to_percent_display(avg_pen),
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": to_percent_display(
                    penetration_by_row[end_row]
                ),
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    anchor_row, anchor_col = find_max_anchor(sheet)

    y_col = anchor_col + REGRESSION_Y_COL_OFFSET
    x_col = anchor_col + REGRESSION_X_COL_OFFSET

    y_series = trailing_numeric_block(sheet, y_col, anchor_row - 1)
    x_series = trailing_numeric_block(sheet, x_col, anchor_row - 1)
    aligned = select_common_rows(x_series, y_series)
    # aligned tuple = (row, x_value, y_value)
    if len(aligned) < 2:
        return []

    max_n = min(N_QUARTERS, len(aligned))
    if max_n < 2:
        return []

    temp_row = anchor_row + TEMP_FORMULA_ROW_OFFSET
    temp_col = anchor_col + TEMP_FORMULA_COL_OFFSET
    intercept_cell = sheet.cells(temp_row, temp_col)
    slope_cell = sheet.cells(temp_row + 1, temp_col)

    next_x_value = to_float(sheet.cells(anchor_row, x_col).value)
    if next_x_value is None:
        next_x_value = aligned[-1][1]

    actual_value = to_float(sheet.cells(anchor_row, y_col).value)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for n_used in range(2, max_n + 1):
        window = aligned[-n_used:]
        start_row = window[0][0]
        end_row = window[-1][0]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        predicted = [intercept + (slope * x_val) for _, x_val, _ in window]
        forecast_total_without_sa = intercept + (slope * next_x_value)
        forecast_max = max(predicted) if predicted else None
        forecast_min = min(predicted) if predicted else None
        range_width = safe_subtract(forecast_max, forecast_min)

        signature = (
            rounded_signature(intercept),
            rounded_signature(slope),
            rounded_signature(forecast_total_without_sa),
            rounded_signature(forecast_max),
            rounded_signature(forecast_min),
        )
        if signature == previous_signature:
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
                "parameter_value": n_used,
                "num_quarters_used": n_used,
                "forecast_value": forecast_total_without_sa,  # TOT FCST w/o SA
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


def write_sheet(
    ws: Any, columns: list[str], rows: list[dict[str, Any]], width_cap: int = 48
) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([serialize_cell_value(row.get(col)) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 1)}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_length = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, width_cap)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def run() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)

    files, skipped = collect_input_files(in_dir)
    for file_name, reason in skipped:
        log(f"Skipped {file_name}: {reason}")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files:
            log(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                # Requirement: open workbook this way and only once.
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)

                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_count += 1
            except Exception as exc:
                log(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_workbook_without_saving(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_path = build_output_path(in_dir, out_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    log(f"Output path: {output_path}")
    log(f"Number of files processed: {processed_count}")
    log(f"Number of empirical rows: {len(empirical_rows)}")
    log(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
