#!/usr/bin/env python3
"""Extract empirical and regression candidates from model workbooks."""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-editable paths.
input_dir = Path("input")
output_dir = Path("output")

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

FILENAME_PATTERN = re.compile(
    r".* - (?P<ticker>[A-Za-z0-9]+) - (?P<window>Early|Mid|Late)"
    r"(?P<month>[A-Za-z]{3})(?P<year>\d{4})_Send\.xlsx$",
    re.IGNORECASE,
)

DAY_BY_WINDOW = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SeriesData:
    points: list[tuple[int, int, float]]
    orientation: str


def ensure_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def is_numeric(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> Optional[float]:
    if is_numeric(value):
        return float(value)
    return None


def normalize_label(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def safe_signature_number(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 10)


def parse_model_metadata(file_name: str) -> Optional[ModelMetadata]:
    match = FILENAME_PATTERN.match(file_name)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    window = match.group("window").title()
    month_token = match.group("month").title()
    year = int(match.group("year"))

    try:
        month = dt.datetime.strptime(month_token, "%b").month
    except ValueError:
        return None

    day = DAY_BY_WINDOW.get(window.lower())
    if day is None:
        return None

    model_period = f"{window}{month_token}_{year}"
    model_date = dt.date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def resolve_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_folder / f"{base}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def safe_close_workbook(wb: Any) -> None:
    if wb is None:
        return

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
    except Exception as exc:
        print(f"Warning: workbook close fallback failed: {exc}")


def get_sheet_if_exists(wb: Any, sheet_name: str) -> Optional[Any]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def find_anchor_cell(sheet: Any, anchor_text: str = "max") -> Optional[tuple[int, int]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    candidates: list[tuple[int, int, int]] = []

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if not isinstance(value, str):
                continue
            if normalize_label(value) != normalize_label(anchor_text):
                continue

            row_num = start_row + r_idx
            col_num = start_col + c_idx
            score = 0

            right_value = sheet.cells(row_num, col_num + 1).value
            if is_numeric(right_value):
                score += 2

            below_label = sheet.cells(row_num + 1, col_num).value
            if isinstance(below_label, str) and "min" in normalize_label(below_label):
                score += 2

            candidates.append((score, row_num, col_num))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (item[0], item[1], item[2]), reverse=True)
    _, row_num, col_num = candidates[0]
    return row_num, col_num


def build_label_map(
    sheet: Any,
    anchor_row: int,
    anchor_col: int,
    row_radius: int = 90,
    col_radius: int = 22,
) -> dict[str, list[tuple[int, int]]]:
    top = max(1, anchor_row - row_radius)
    bottom = anchor_row + row_radius
    left = max(1, anchor_col - col_radius)
    right = anchor_col + col_radius

    values = ensure_2d(sheet.range((top, left), (bottom, right)).value)
    label_map: dict[str, list[tuple[int, int]]] = {}

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if not isinstance(value, str):
                continue
            key = normalize_label(value)
            if not key:
                continue
            label_map.setdefault(key, []).append((top + r_idx, left + c_idx))

    return label_map


def find_label_cell(label_map: dict[str, list[tuple[int, int]]], *phrases: str) -> Optional[tuple[int, int]]:
    normalized = [normalize_label(phrase) for phrase in phrases if phrase]
    for phrase in normalized:
        if phrase in label_map:
            return label_map[phrase][0]

    for key, cells in label_map.items():
        for phrase in normalized:
            if phrase and phrase in key:
                return cells[0]
    return None


def pick_value_cell_next_to_label(sheet: Any, label_cell: tuple[int, int]) -> tuple[int, int]:
    row_num, col_num = label_cell
    for candidate_col in (col_num + 1, col_num + 2, col_num - 1):
        if candidate_col < 1:
            continue
        value = sheet.cells(row_num, candidate_col).value
        if value not in (None, ""):
            return row_num, candidate_col
    return row_num, col_num + 1


def choose_metric_cell(
    sheet: Any,
    label_map: dict[str, list[tuple[int, int]]],
    anchor_row: int,
    anchor_col: int,
    label_hints: tuple[str, ...],
    fallback_offset: tuple[int, int],
) -> tuple[int, int]:
    label_cell = find_label_cell(label_map, *label_hints)
    if label_cell:
        return pick_value_cell_next_to_label(sheet, label_cell)
    return anchor_row + fallback_offset[0], anchor_col + fallback_offset[1]


def collect_numeric_cells_horizontal(
    sheet: Any,
    row_num: int,
    start_col: int,
    end_col: int,
) -> list[tuple[int, int, float]]:
    if end_col < start_col:
        return []

    values = sheet.range((row_num, start_col), (row_num, end_col)).value
    if not isinstance(values, list):
        values = [values]

    points: list[tuple[int, int, float]] = []
    for idx, value in enumerate(values):
        if is_numeric(value):
            points.append((row_num, start_col + idx, float(value)))
    return points


def collect_numeric_cells_vertical(
    sheet: Any,
    col_num: int,
    start_row: int,
    end_row: int,
) -> list[tuple[int, int, float]]:
    if end_row < start_row:
        return []

    values = ensure_2d(sheet.range((start_row, col_num), (end_row, col_num)).value)
    points: list[tuple[int, int, float]] = []
    for idx, row in enumerate(values):
        if not row:
            continue
        value = row[0]
        if is_numeric(value):
            points.append((start_row + idx, col_num, float(value)))
    return points


def detect_numeric_series_from_label(
    sheet: Any,
    label_map: dict[str, list[tuple[int, int]]],
    phrases: tuple[str, ...],
    anchor_col: int,
) -> Optional[SeriesData]:
    label_cell = find_label_cell(label_map, *phrases)
    if not label_cell:
        return None

    label_row, label_col = label_cell
    max_h_col = max(label_col + 80, anchor_col + 2)
    horizontal = collect_numeric_cells_horizontal(sheet, label_row, label_col + 1, max_h_col)
    vertical = collect_numeric_cells_vertical(sheet, label_col + 1, label_row + 1, label_row + 80)

    if len(horizontal) >= len(vertical) and horizontal:
        return SeriesData(points=horizontal, orientation="horizontal")
    if vertical:
        return SeriesData(points=vertical, orientation="vertical")
    return None


def derive_ratio_series(quarterly: Optional[SeriesData], reported: Optional[SeriesData]) -> Optional[SeriesData]:
    if quarterly is None or reported is None:
        return None
    if quarterly.orientation != reported.orientation:
        return None

    ratios: list[tuple[int, int, float]] = []
    if quarterly.orientation == "horizontal":
        q_map = {col_num: (row_num, value) for row_num, col_num, value in quarterly.points}
        r_map = {col_num: (row_num, value) for row_num, col_num, value in reported.points}
        common_cols = sorted(set(q_map) & set(r_map))
        for col_num in common_cols:
            q_row, q_value = q_map[col_num]
            _, r_value = r_map[col_num]
            if r_value == 0:
                continue
            ratios.append((q_row, col_num, (q_value / r_value) * 100.0))
    else:
        q_map = {row_num: (col_num, value) for row_num, col_num, value in quarterly.points}
        r_map = {row_num: (col_num, value) for row_num, col_num, value in reported.points}
        common_rows = sorted(set(q_map) & set(r_map))
        for row_num in common_rows:
            q_col, q_value = q_map[row_num]
            _, r_value = r_map[row_num]
            if r_value == 0:
                continue
            ratios.append((row_num, q_col, (q_value / r_value) * 100.0))

    if not ratios:
        return None
    return SeriesData(points=ratios, orientation=quarterly.orientation)


def build_average_formula(selected_points: list[tuple[int, int, float]]) -> str:
    coords = [(row_num, col_num) for row_num, col_num, _ in selected_points]
    rows = [row_num for row_num, _ in coords]
    cols = [col_num for _, col_num in coords]

    if len(set(rows)) == 1:
        row_num = rows[0]
        sorted_cols = sorted(cols)
        contiguous = sorted_cols == list(range(sorted_cols[0], sorted_cols[-1] + 1))
        if contiguous:
            return f"=AVERAGE(R{row_num}C{sorted_cols[0]}:R{row_num}C{sorted_cols[-1]})"

    if len(set(cols)) == 1:
        col_num = cols[0]
        sorted_rows = sorted(rows)
        contiguous = sorted_rows == list(range(sorted_rows[0], sorted_rows[-1] + 1))
        if contiguous:
            return f"=AVERAGE(R{sorted_rows[0]}C{col_num}:R{sorted_rows[-1]}C{col_num})"

    refs = ",".join(f"R{row_num}C{col_num}" for row_num, col_num in coords)
    return f"=AVERAGE({refs})"


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def infer_last_quarter_label(sheet: Any, row_num: int, col_num: int, orientation: str) -> str:
    candidates: list[Any]
    if orientation == "horizontal":
        candidates = [
            sheet.cells(max(1, row_num - 1), col_num).value,
            sheet.cells(max(1, row_num - 2), col_num).value,
        ]
    else:
        candidates = [
            sheet.cells(row_num, max(1, col_num - 1)).value,
            sheet.cells(row_num, max(1, col_num - 2)).value,
        ]

    for candidate in candidates:
        if isinstance(candidate, str) and candidate.strip():
            return candidate.strip()
    return ""


def process_empirical_sheet(wb: Any, meta: ModelMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical in {source_file}: 'Empirical Model' sheet not found")
        return []

    anchor = find_anchor_cell(sheet, anchor_text="max")
    if anchor is None:
        print(f"Skipped empirical in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    label_map = build_label_map(sheet, anchor_row, anchor_col)

    forecast_max_cell = choose_metric_cell(
        sheet, label_map, anchor_row, anchor_col, ("max",), (0, 1)
    )
    forecast_min_cell = choose_metric_cell(
        sheet, label_map, anchor_row, anchor_col, ("min",), (1, 1)
    )
    forecast_value_cell = choose_metric_cell(
        sheet,
        label_map,
        anchor_row,
        anchor_col,
        ("estimated total sold", "est total sold", "forecast total", "forecast value"),
        (-2, 1),
    )
    actual_value_cell = choose_metric_cell(
        sheet,
        label_map,
        anchor_row,
        anchor_col,
        ("reported sales", "actual sales", "actual value"),
        (-1, 1),
    )
    growth_rate_cell = choose_metric_cell(
        sheet,
        label_map,
        anchor_row,
        anchor_col,
        ("growth rate", "growth pct", "growth %"),
        (2, 1),
    )
    db_capture_cell = choose_metric_cell(
        sheet,
        label_map,
        anchor_row,
        anchor_col,
        ("sales captured in db", "captured in db", "db capture"),
        (3, 1),
    )

    avg_pen_label = find_label_cell(label_map, "avg penetration", "average penetration", "penetration")
    if avg_pen_label:
        avg_pen_cell = pick_value_cell_next_to_label(sheet, avg_pen_label)
    else:
        avg_pen_cell = (anchor_row - 4, anchor_col + 1)

    quarterly_series = detect_numeric_series_from_label(
        sheet, label_map, ("quarterly sales", "quarter sales"), anchor_col
    )
    reported_series = detect_numeric_series_from_label(
        sheet, label_map, ("reported sales", "actual sales", "sales"), anchor_col
    )
    penetration_series = detect_numeric_series_from_label(
        sheet,
        label_map,
        ("penetration pct", "penetration %", "penetration", "avg penetration"),
        anchor_col,
    )

    if penetration_series is None or len(penetration_series.points) < 2:
        penetration_series = derive_ratio_series(quarterly_series, reported_series)

    if penetration_series is None or not penetration_series.points:
        # Keep one row if a penetration history was not found.
        loop_count = 1
    else:
        loop_count = min(N_QUARTERS, len(penetration_series.points))

    rows: list[dict[str, Any]] = []
    for n in range(1, loop_count + 1):
        selected_points: list[tuple[int, int, float]] = []
        if penetration_series is not None and len(penetration_series.points) >= n:
            selected_points = penetration_series.points[-n:]
            formula = build_average_formula(selected_points)
            set_formula2(sheet.cells(avg_pen_cell[0], avg_pen_cell[1]), formula)
            wb.app.calculate()

        avg_penetration = to_float(sheet.cells(avg_pen_cell[0], avg_pen_cell[1]).value)
        forecast_value = to_float(sheet.cells(forecast_value_cell[0], forecast_value_cell[1]).value)
        actual_value = to_float(sheet.cells(actual_value_cell[0], actual_value_cell[1]).value)
        forecast_max = to_float(sheet.cells(forecast_max_cell[0], forecast_max_cell[1]).value)
        forecast_min = to_float(sheet.cells(forecast_min_cell[0], forecast_min_cell[1]).value)
        growth_rate = to_float(sheet.cells(growth_rate_cell[0], growth_rate_cell[1]).value)
        sales_captured = to_float(sheet.cells(db_capture_cell[0], db_capture_cell[1]).value)

        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min
        else:
            range_width = None

        last_quarter_used = ""
        quarterly_sales: Optional[float] = None
        reported_sales: Optional[float] = None

        if quarterly_series is not None and len(quarterly_series.points) >= n:
            q_row, q_col, quarterly_sales = quarterly_series.points[-n]
            last_quarter_used = infer_last_quarter_label(
                sheet, q_row, q_col, quarterly_series.orientation
            )

        if reported_series is not None and len(reported_series.points) >= n:
            _, _, reported_sales = reported_series.points[-n]

        if quarterly_sales is not None and reported_sales not in (None, 0):
            sales_captured = (quarterly_sales / reported_sales) * 100.0

        if reported_series is not None and len(reported_series.points) >= n + 1 and reported_sales is not None:
            previous_reported = reported_series.points[-(n + 1)][2]
            if previous_reported != 0:
                growth_rate = ((reported_sales - previous_reported) / previous_reported) * 100.0

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n,
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
        rows.append(row)

    return rows


def process_regression_sheet(wb: Any, meta: ModelMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped regression in {source_file}: 'Regression Model' sheet not found")
        return []

    anchor = find_anchor_cell(sheet, anchor_text="max")
    if anchor is None:
        print(f"Skipped regression in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    label_map = build_label_map(sheet, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    used = sheet.used_range
    data_rows: list[int] = []
    for row_num in range(used.row, anchor_row):
        x_value = sheet.cells(row_num, x_col).value
        y_value = sheet.cells(row_num, y_col).value
        if is_numeric(x_value) and is_numeric(y_value):
            data_rows.append(row_num)

    if len(data_rows) < 2:
        print(f"Skipped regression in {source_file}: insufficient x/y pairs")
        return []

    intercept_label = find_label_cell(label_map, "intercept")
    slope_label = find_label_cell(label_map, "slope")
    if intercept_label:
        intercept_cell = pick_value_cell_next_to_label(sheet, intercept_label)
    else:
        intercept_cell = (anchor_row + 2, anchor_col + 1)
    if slope_label:
        slope_cell = pick_value_cell_next_to_label(sheet, slope_label)
    else:
        slope_cell = (anchor_row + 3, anchor_col + 1)

    forecast_value_cell = choose_metric_cell(
        sheet,
        label_map,
        anchor_row,
        anchor_col,
        ("tot fcst w o sa", "tot fcst wo sa", "tot fcst w/out sa", "forecast total"),
        (-2, 1),
    )
    actual_value_cell = choose_metric_cell(
        sheet,
        label_map,
        anchor_row,
        anchor_col,
        ("actual value", "reported sales", "actual sales"),
        (-1, 1),
    )
    forecast_max_cell = choose_metric_cell(
        sheet, label_map, anchor_row, anchor_col, ("max",), (0, 1)
    )
    forecast_min_cell = choose_metric_cell(
        sheet, label_map, anchor_row, anchor_col, ("min",), (1, 1)
    )

    max_n = min(N_QUARTERS, len(data_rows))
    rows: list[dict[str, Any]] = []

    for n in range(2, max_n + 1):
        window_rows = data_rows[-n:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2(sheet.cells(intercept_cell[0], intercept_cell[1]), intercept_formula)
        set_formula2(sheet.cells(slope_cell[0], slope_cell[1]), slope_formula)
        wb.app.calculate()

        intercept = to_float(sheet.cells(intercept_cell[0], intercept_cell[1]).value)
        slope = to_float(sheet.cells(slope_cell[0], slope_cell[1]).value)
        forecast_value = to_float(sheet.cells(forecast_value_cell[0], forecast_value_cell[1]).value)
        actual_value = to_float(sheet.cells(actual_value_cell[0], actual_value_cell[1]).value)
        forecast_max = to_float(sheet.cells(forecast_max_cell[0], forecast_max_cell[1]).value)
        forecast_min = to_float(sheet.cells(forecast_min_cell[0], forecast_min_cell[1]).value)

        if forecast_value is None and intercept is not None and slope is not None:
            latest_x = to_float(sheet.cells(window_rows[-1], x_col).value)
            if latest_x is not None:
                forecast_value = intercept + slope * (latest_x + 1.0)

        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min
        else:
            range_width = None

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n,
            "num_quarters_used": n,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows:
            prev = rows[-1]
            prev_sig = (
                safe_signature_number(prev["forecast_value"]),
                safe_signature_number(prev["forecast_max"]),
                safe_signature_number(prev["forecast_min"]),
                safe_signature_number(prev["intercept"]),
                safe_signature_number(prev["slope"]),
            )
            new_sig = (
                safe_signature_number(row["forecast_value"]),
                safe_signature_number(row["forecast_max"]),
                safe_signature_number(row["forecast_min"]),
                safe_signature_number(row["intercept"]),
                safe_signature_number(row["slope"]),
            )
            if new_sig == prev_sig:
                continue

        rows.append(row)

    return rows


def write_sheet_with_formatting(
    wb: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)

    for row_dict in rows:
        ws.append([row_dict.get(col_name) for col_name in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_dict in rows:
            value = row_dict.get(col_name)
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 50)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet_with_formatting(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet_with_formatting(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def main() -> None:
    if not input_dir.exists():
        print(f"Input directory does not exist: {input_dir}")
        return

    output_path = resolve_output_path(input_dir, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = None
    original_calculation = None

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.enable_events = False
        except Exception:
            pass

        try:
            original_calculation = app.calculation
            app.calculation = "manual"
        except Exception:
            original_calculation = None

        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (not a regular file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            metadata = parse_model_metadata(file_path.name)
            if metadata is None:
                print(f"Skipped file: {file_path.name} (filename format mismatch)")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, metadata, file_path.name))
                files_processed += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                safe_close_workbook(workbook)
    finally:
        if app is not None:
            if original_calculation is not None:
                try:
                    app.calculation = original_calculation
                except Exception:
                    pass
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
