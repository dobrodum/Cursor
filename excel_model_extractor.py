#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Sequence

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw

# --------------------------- User configuration ---------------------------- #
input_dir = Path("./input")
output_dir = Path("./output")
# --------------------------------------------------------------------------- #

N_QUARTERS = 10
MAX_HISTORY_ROWS = 200
MIN_COLUMN_WIDTH = 12
MAX_COLUMN_WIDTH = 48

EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"

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
}

MODEL_PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}

# Relative to the "max" anchor cell.
EMPIRICAL_OFFSETS = {
    "quarter_label_col": -12,
    "sales_captured_in_db_pct_col": -11,
    "growth_rate_pct_col": -10,
    "reported_sales_col": -9,
    "quarterly_sales_col": -8,
}

REGRESSION_X_COL_OFFSET = -11
REGRESSION_Y_COL_OFFSET = -7


@dataclass(frozen=True)
class ModelLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetSnapshot:
    start_row: int
    start_col: int
    values: list[list[Any]]

    def get_value(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        row_values = self.values[row_idx]
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]


@dataclass(frozen=True)
class EmpiricalPoint:
    row: int
    quarter_label: str
    sales_captured_in_db_pct: float
    growth_rate_pct: float | None
    reported_sales: float | None
    quarterly_sales: float


@dataclass(frozen=True)
class RegressionPoint:
    row: int
    x_value: float
    y_value: float


def coerce_2d(values: Any, rows: int, cols: int) -> list[list[Any]]:
    if values is None:
        return []

    if rows == 1 and cols == 1:
        return [[values]]

    if rows == 1:
        if isinstance(values, Sequence) and not isinstance(values, (str, bytes)):
            return [list(values)]
        return [[values]]

    if cols == 1:
        if isinstance(values, Sequence) and not isinstance(values, (str, bytes)):
            return [[item] for item in values]
        return [[values]]

    if isinstance(values, Sequence) and not isinstance(values, (str, bytes)):
        return [list(row) if isinstance(row, Sequence) and not isinstance(row, (str, bytes)) else [row] for row in values]

    return [[values]]


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        is_pct = cleaned.endswith("%")
        if is_pct:
            cleaned = cleaned[:-1]
        try:
            number = float(cleaned)
        except ValueError:
            return None
        return number / 100 if is_pct else number
    return None


def normalize_pct(value: float | None) -> float | None:
    if value is None:
        return None
    if abs(value) > 1:
        return value / 100
    return value


def safe_divide(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def safe_subtract(a_value: float | None, b_value: float | None) -> float | None:
    if a_value is None or b_value is None:
        return None
    return a_value - b_value


def to_text(value: Any, default: str = "") -> str:
    if value is None:
        return default
    text = str(value).strip()
    return text if text else default


def round_or_none(value: float | None, digits: int = 10) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def close_workbook_safely(workbook: Any) -> None:
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
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception as exc:  # pragma: no cover - best effort fallback.
        print(f"Warning: workbook close fallback failed: {exc}")


def read_sheet_snapshot(sheet: Any) -> SheetSnapshot:
    used_range = sheet.used_range
    start_row = used_range.row
    start_col = used_range.column
    rows = used_range.rows.count
    cols = used_range.columns.count
    values = coerce_2d(used_range.value, rows, cols)
    return SheetSnapshot(start_row=start_row, start_col=start_col, values=values)


def find_anchor(snapshot: SheetSnapshot, token: str = "max") -> tuple[int, int] | None:
    token_lc = token.strip().lower()
    for row_offset, row_values in enumerate(snapshot.values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == token_lc:
                return snapshot.start_row + row_offset, snapshot.start_col + col_offset
    return None


def parse_model_label(file_path: Path) -> ModelLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"
    period_source = parts[2] if len(parts) >= 3 else stem
    period_token = period_source.split("_")[0].strip()

    period_match = re.search(r"(?i)(Early|Mid|Late)([A-Za-z]{3})(\d{4})", period_token)
    if not period_match:
        model_period = period_token or "unknown"
        model_date = ""
    else:
        cadence = period_match.group(1).title()
        month_abbr = period_match.group(2).title()
        year_text = period_match.group(3)
        model_period = f"{cadence}{month_abbr}_{year_text}"

        month_num = MONTH_MAP.get(month_abbr)
        day_num = MODEL_PERIOD_DAY.get(cadence)
        if month_num is None or day_num is None:
            model_date = ""
        else:
            model_date = date(int(year_text), month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return ModelLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(source_dir: Path, target_dir: Path) -> Path:
    base_name = f"{source_dir.name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def is_generated_output_file(file_path: Path, input_folder_name: str) -> bool:
    pattern = rf"^{re.escape(input_folder_name)}_PARAM(?:\.\d+)?\.xlsx$"
    return bool(re.match(pattern, file_path.name, flags=re.IGNORECASE))


def collect_empirical_points(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> list[EmpiricalPoint]:
    quarter_col = anchor_col + EMPIRICAL_OFFSETS["quarter_label_col"]
    sales_captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct_col"]
    growth_rate_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct_col"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"]

    scan_start = max(snapshot.start_row, anchor_row - MAX_HISTORY_ROWS)
    points: list[EmpiricalPoint] = []

    for row in range(scan_start, anchor_row):
        sales_captured = normalize_pct(to_float(snapshot.get_value(row, sales_captured_col)))
        quarterly_sales = to_float(snapshot.get_value(row, quarterly_sales_col))
        if sales_captured is None or quarterly_sales is None:
            continue

        points.append(
            EmpiricalPoint(
                row=row,
                quarter_label=to_text(snapshot.get_value(row, quarter_col), default=f"R{row}"),
                sales_captured_in_db_pct=sales_captured,
                growth_rate_pct=normalize_pct(to_float(snapshot.get_value(row, growth_rate_col))),
                reported_sales=to_float(snapshot.get_value(row, reported_sales_col)),
                quarterly_sales=quarterly_sales,
            )
        )

    return points


def collect_regression_points(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> list[RegressionPoint]:
    x_col = anchor_col + REGRESSION_X_COL_OFFSET
    y_col = anchor_col + REGRESSION_Y_COL_OFFSET
    scan_start = max(snapshot.start_row, anchor_row - MAX_HISTORY_ROWS)

    points: list[RegressionPoint] = []
    for row in range(scan_start, anchor_row):
        x_value = to_float(snapshot.get_value(row, x_col))
        y_value = to_float(snapshot.get_value(row, y_col))
        if x_value is None or y_value is None:
            continue
        points.append(RegressionPoint(row=row, x_value=x_value, y_value=y_value))

    return points


def clear_cells(*cells: Any) -> None:
    for cell in cells:
        try:
            cell.value = None
        except Exception:
            continue


def extract_empirical_candidates(workbook: Any, label: ModelLabel, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets[EMPIRICAL_MODEL_SHEET]
    except Exception:
        print(f"Skipped {source_file}: missing sheet '{EMPIRICAL_MODEL_SHEET}'")
        return []

    snapshot = read_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, token="max")
    if anchor is None:
        print(f"Skipped {source_file}: no 'max' anchor found on '{EMPIRICAL_MODEL_SHEET}'")
        return []

    anchor_row, anchor_col = anchor
    empirical_points = collect_empirical_points(snapshot, anchor_row, anchor_col)
    if not empirical_points:
        print(f"Skipped {source_file}: no empirical history points near anchor")
        return []

    penetration_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct_col"]
    helper_col = anchor_col + 6
    avg_pen_cell = sheet.cells(anchor_row, helper_col)
    min_pen_cell = sheet.cells(anchor_row, helper_col + 1)
    max_pen_cell = sheet.cells(anchor_row, helper_col + 2)

    rows: list[dict[str, Any]] = []
    max_iterations = min(N_QUARTERS, len(empirical_points))

    for n_quarters in range(1, max_iterations + 1):
        selected = empirical_points[-n_quarters:]
        first_row = selected[0].row
        last_row = selected[-1].row
        latest_point = selected[-1]

        avg_pen_cell.formula2 = f"=AVERAGE(R{first_row}C{penetration_col}:R{last_row}C{penetration_col})"
        min_pen_cell.formula2 = f"=MIN(R{first_row}C{penetration_col}:R{last_row}C{penetration_col})"
        max_pen_cell.formula2 = f"=MAX(R{first_row}C{penetration_col}:R{last_row}C{penetration_col})"
        workbook.app.calculate()

        avg_penetration = normalize_pct(to_float(avg_pen_cell.value))
        min_penetration = normalize_pct(to_float(min_pen_cell.value))
        max_penetration = normalize_pct(to_float(max_pen_cell.value))

        forecast_value = safe_divide(latest_point.quarterly_sales, avg_penetration)
        forecast_max = safe_divide(latest_point.quarterly_sales, min_penetration)
        forecast_min = safe_divide(latest_point.quarterly_sales, max_penetration)

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": latest_point.quarter_label,
                "forecast_value": forecast_value,
                "actual_value": latest_point.reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": latest_point.quarterly_sales,
                "reported_sales": latest_point.reported_sales,
                "growth_rate_pct": latest_point.growth_rate_pct,
                "sales_captured_in_db_pct": latest_point.sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    clear_cells(avg_pen_cell, min_pen_cell, max_pen_cell)
    return rows


def extract_regression_candidates(workbook: Any, label: ModelLabel, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets[REGRESSION_MODEL_SHEET]
    except Exception:
        print(f"Skipped {source_file}: missing sheet '{REGRESSION_MODEL_SHEET}'")
        return []

    snapshot = read_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, token="max")
    if anchor is None:
        print(f"Skipped {source_file}: no 'max' anchor found on '{REGRESSION_MODEL_SHEET}'")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col + REGRESSION_X_COL_OFFSET
    y_col = anchor_col + REGRESSION_Y_COL_OFFSET
    points = collect_regression_points(snapshot, anchor_row, anchor_col)
    if not points:
        print(f"Skipped {source_file}: no regression history points near anchor")
        return []

    helper_col = anchor_col + 6
    intercept_cell = sheet.cells(anchor_row, helper_col)
    slope_cell = sheet.cells(anchor_row, helper_col + 1)
    forecast_cell = sheet.cells(anchor_row, helper_col + 2)
    max_cell = sheet.cells(anchor_row, helper_col + 3)
    min_cell = sheet.cells(anchor_row, helper_col + 4)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[float | None, ...] | None = None
    total_points = len(points)

    for n_quarters in range(1, N_QUARTERS + 1):
        effective_quarters = min(n_quarters, total_points)
        selected = points[-effective_quarters:]
        first_row = selected[0].row
        last_row = selected[-1].row
        latest_x = selected[-1].x_value

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_cell.formula2 = f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        forecast_cell.formula2 = f"=R{anchor_row}C{helper_col}+R{anchor_row}C{helper_col + 1}*{latest_x}"
        max_cell.formula2 = f"=MAX(R{first_row}C{y_col}:R{last_row}C{y_col})"
        min_cell.formula2 = f"=MIN(R{first_row}C{y_col}:R{last_row}C{y_col})"
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_total_without_sa = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)

        signature = (
            round_or_none(intercept),
            round_or_none(slope),
            round_or_none(forecast_total_without_sa),
            round_or_none(forecast_max),
            round_or_none(forecast_min),
        )

        if signature == previous_signature:
            if effective_quarters == total_points:
                break
            continue
        previous_signature = signature

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": effective_quarters,
                "num_quarters_used": effective_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    clear_cells(intercept_cell, slope_cell, forecast_cell, max_cell, min_cell)
    return rows


def write_sheet(
    worksheet: Any,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    worksheet.append(columns)
    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for header_cell in worksheet[1]:
        header_cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for index, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for column_cells in worksheet.iter_cols(min_col=index, max_col=index, min_row=2, max_row=worksheet.max_row):
            for cell in column_cells:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
        column_letter = get_column_letter(index)
        worksheet.column_dimensions[column_letter].width = max(MIN_COLUMN_WIDTH, min(MAX_COLUMN_WIDTH, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def run() -> None:
    source_dir = input_dir.expanduser().resolve()
    target_dir = output_dir.expanduser().resolve()

    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a directory: {source_dir}")

    target_dir.mkdir(parents=True, exist_ok=True)
    output_path = build_output_path(source_dir, target_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    original_calc_mode: str | None = None
    try:
        app.display_alerts = False
        app.screen_updating = False
        try:
            original_calc_mode = app.calculation
            app.calculation = "manual"
        except Exception:
            original_calc_mode = None

        for file_path in sorted(source_dir.iterdir(), key=lambda item: item.name.lower()):
            if not file_path.is_file():
                continue

            file_name = file_path.name
            if file_name.startswith("~"):
                print(f"Skipped {file_name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_name}: not an .xlsx file")
                continue
            if is_generated_output_file(file_path, source_dir.name):
                print(f"Skipped {file_name}: generated output file pattern")
                continue

            print(f"Processing file: {file_name}")
            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_name}: failed to open workbook ({exc})")
                continue

            try:
                label = parse_model_label(file_path)
                empirical_rows.extend(extract_empirical_candidates(workbook, label, file_name))
                regression_rows.extend(extract_regression_candidates(workbook, label, file_name))
                processed_file_count += 1
            except Exception as exc:
                print(f"Skipped {file_name}: extraction error ({exc})")
            finally:
                close_workbook_safely(workbook)
    finally:
        if original_calc_mode is not None:
            try:
                app.calculation = original_calc_mode
            except Exception:
                pass
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_file_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
