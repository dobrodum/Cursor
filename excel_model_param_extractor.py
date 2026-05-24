#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -------- User-configurable paths --------
input_dir = Path("input")
output_dir = Path("output")


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


PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}
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


@dataclass
class SheetSnapshot:
    values: list[list[Any]]
    start_row: int
    start_col: int


def as_matrix(value: Any) -> list[list[Any]]:
    if value is None:
        return []
    if isinstance(value, tuple):
        value = list(value)
    if not isinstance(value, list):
        return [[value]]
    if not value:
        return []
    first = value[0]
    if isinstance(first, tuple):
        return [list(row) for row in value]
    if isinstance(first, list):
        return [list(row) for row in value]
    return [list(value)]


def normalized_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        if math.isnan(number):
            return None
        return number
    text = str(value).strip().replace(",", "")
    if text in {"", "-", "--", "n/a", "na"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def stringify(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text


def parse_file_labels(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    period_token = ""
    if len(parts) >= 2:
        ticker = parts[1]
    if len(parts) >= 3:
        period_token = parts[2]

    period_token = re.sub(r"(?i)_?send.*$", "", period_token).strip()

    model_period = ""
    model_date = ""
    period_match = re.match(
        r"(?i)^(Early|Mid|Late)[\s_]*([A-Za-z]{3,9})[\s_]*(\d{4})$",
        period_token,
    )
    if period_match:
        part_of_month = period_match.group(1).title()
        month_name = period_match.group(2)[:3].title()
        year = int(period_match.group(3))
        month = MONTH_MAP.get(month_name.lower())
        day = PERIOD_DAY[part_of_month.lower()]
        if month:
            model_period = f"{part_of_month}{month_name}_{year}"
            model_date = date(year, month, day).isoformat()

    if not ticker and parts:
        ticker = parts[0]
    if not model_period:
        model_period = period_token.replace(" ", "_") if period_token else ""
    model = f"{ticker}_{model_period}".strip("_")

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    idx = 1
    while candidate.exists():
        candidate = out_dir / f"{in_dir.name}_PARAM.{idx}.xlsx"
        idx += 1
    return candidate


def get_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    return SheetSnapshot(
        values=as_matrix(used.value),
        start_row=used.row,
        start_col=used.column,
    )


def matrix_value(snapshot: SheetSnapshot, abs_row: int, abs_col: int) -> Any:
    rel_row = abs_row - snapshot.start_row
    rel_col = abs_col - snapshot.start_col
    if rel_row < 0 or rel_col < 0 or rel_row >= len(snapshot.values):
        return None
    row = snapshot.values[rel_row]
    if rel_col >= len(row):
        return None
    return row[rel_col]


def find_anchor(snapshot: SheetSnapshot, label: str) -> tuple[int, int] | None:
    target = normalized_text(label)
    for rel_row, row in enumerate(snapshot.values):
        for rel_col, value in enumerate(row):
            if normalized_text(value) == target:
                return snapshot.start_row + rel_row, snapshot.start_col + rel_col
    return None


def find_nearby_label(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    label: str,
    radius: int = 8,
) -> tuple[int, int] | None:
    target = normalized_text(label)
    for row in range(anchor_row - radius, anchor_row + radius + 1):
        for col in range(anchor_col - radius, anchor_col + radius + 1):
            if normalized_text(matrix_value(snapshot, row, col)) == target:
                return row, col
    return None


def nearest_numeric(snapshot: SheetSnapshot, row: int, col: int) -> float | None:
    checks = [
        (0, 1),
        (0, -1),
        (1, 0),
        (-1, 0),
        (1, 1),
        (1, -1),
        (-1, 1),
        (-1, -1),
        (0, 2),
        (2, 0),
    ]
    for d_row, d_col in checks:
        value = to_float(matrix_value(snapshot, row + d_row, col + d_col))
        if value is not None:
            return value
    return None


def collect_xy_series(
    snapshot: SheetSnapshot,
    anchor_row: int,
    x_col: int,
    y_col: int,
    quarter_col: int,
) -> list[dict[str, Any]]:
    start_row = max(snapshot.start_row, anchor_row - 160)
    points: list[dict[str, Any]] = []
    for row in range(start_row, anchor_row):
        x_val = to_float(matrix_value(snapshot, row, x_col))
        y_val = to_float(matrix_value(snapshot, row, y_col))
        if x_val is None or y_val is None:
            continue
        points.append(
            {
                "row": row,
                "x": x_val,
                "y": y_val,
                "quarter": matrix_value(snapshot, row, quarter_col),
            }
        )

    if not points:
        return []

    # Keep the last contiguous block nearest the anchor.
    contiguous = [points[-1]]
    prev_row = points[-1]["row"]
    for point in reversed(points[:-1]):
        if prev_row - point["row"] <= 2:
            contiguous.append(point)
            prev_row = point["row"]
        else:
            break
    contiguous.reverse()
    return contiguous


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
        wb.close()


def calc_range_width(max_value: float | None, min_value: float | None) -> float | None:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def extract_empirical_rows(
    sheet: xw.Sheet,
    meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = get_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    quarter_col = x_col - 1
    series = collect_xy_series(snapshot, anchor_row, x_col, y_col, quarter_col)
    if not series:
        return []

    max_value = nearest_numeric(snapshot, anchor_row, anchor_col)
    min_anchor = find_nearby_label(snapshot, anchor_row, anchor_col, "min")
    min_value = (
        nearest_numeric(snapshot, min_anchor[0], min_anchor[1])
        if min_anchor
        else nearest_numeric(snapshot, anchor_row + 1, anchor_col)
    )
    range_width = calc_range_width(max_value, min_value)

    rows: list[dict[str, Any]] = []
    n_quarters = 10

    temp_col = anchor_col + 30
    avg_cell = sheet.range((anchor_row, temp_col))
    growth_cell = sheet.range((anchor_row + 1, temp_col))

    for n in range(1, n_quarters + 1):
        if len(series) < n:
            continue

        window = series[-n:]
        start_row = window[0]["row"]
        end_row = window[-1]["row"]
        last_point = window[-1]

        # R1C1 formula2 usage for empirical average penetration.
        avg_cell.formula2 = (
            f'=IFERROR(AVERAGE(R{start_row}C{x_col}:R{end_row}C{x_col}/'
            f'R{start_row}C{y_col}:R{end_row}C{y_col}),"")'
        )
        growth_cell.formula2 = (
            f'=IFERROR((R{end_row}C{x_col}/R{start_row}C{x_col})-1,"")'
        )
        sheet.book.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        growth_rate = to_float(growth_cell.value)
        quarterly_sales = last_point["x"]
        reported_sales = last_point["y"]

        forecast_value = None
        if avg_penetration not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration

        sales_captured = None
        if reported_sales not in (None, 0):
            sales_captured = quarterly_sales / reported_sales

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n,
                "last_quarter_used": stringify(last_point["quarter"]),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": max_value,
                "forecast_min": min_value,
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


def extract_regression_rows(
    sheet: xw.Sheet,
    meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = get_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = x_col - 1
    series = collect_xy_series(snapshot, anchor_row, x_col, y_col, quarter_col)
    if len(series) < 2:
        return []

    max_value = nearest_numeric(snapshot, anchor_row, anchor_col)
    min_anchor = find_nearby_label(snapshot, anchor_row, anchor_col, "min")
    min_value = (
        nearest_numeric(snapshot, min_anchor[0], min_anchor[1])
        if min_anchor
        else nearest_numeric(snapshot, anchor_row + 1, anchor_col)
    )
    range_width = calc_range_width(max_value, min_value)

    rows: list[dict[str, Any]] = []
    previous_key: tuple[Any, ...] | None = None

    temp_col = anchor_col + 30
    intercept_cell = sheet.range((anchor_row, temp_col))
    slope_cell = sheet.range((anchor_row + 1, temp_col))
    forecast_cell = sheet.range((anchor_row + 2, temp_col))

    for n in range(2, 11):
        if len(series) < n:
            continue

        window = series[-n:]
        start_row = window[0]["row"]
        end_row = window[-1]["row"]

        # R1C1 formula2 usage for INTERCEPT and SLOPE.
        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        forecast_cell.formula2 = (
            f'=IFERROR(R{anchor_row}C{temp_col}+R{anchor_row + 1}C{temp_col}*'
            f'R{end_row}C{x_col},"")'
        )
        sheet.book.app.calculate()

        intercept_value = to_float(intercept_cell.value)
        slope_value = to_float(slope_cell.value)
        forecast_value = to_float(forecast_cell.value)

        dedupe_key = (
            round(intercept_value, 10) if intercept_value is not None else None,
            round(slope_value, 10) if slope_value is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
        )
        if dedupe_key == previous_key:
            continue
        previous_key = dedupe_key

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": max_value,
                "forecast_min": min_value,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    workbook: Workbook,
    name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_column = get_column_letter(len(columns))
    if rows:
        ws.auto_filter.ref = f"A1:{last_column}{len(rows) + 1}"
    else:
        ws.auto_filter.ref = f"A1:{last_column}1"

    for idx, header in enumerate(columns, start=1):
        max_width = len(header)
        for row in rows:
            text = stringify(row.get(header, ""))
            if len(text) > max_width:
                max_width = len(text)
        ws.column_dimensions[get_column_letter(idx)].width = min(max_width + 2, 40)


def write_output_workbook(
    path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(path)


def iter_source_files(in_dir: Path) -> list[Path]:
    candidates = sorted(in_dir.iterdir())
    files: list[Path] = []
    for path in candidates:
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped file: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {path.name} (not .xlsx)")
            continue
        if re.search(r"(?i)_param(?:\.\d+)?\.xlsx$", path.name):
            print(f"Skipped file: {path.name} (output artifact)")
            continue
        files.append(path)
    return files


def process_workbook(
    app: xw.App,
    file_path: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    meta = parse_file_labels(file_path.name)
    wb = app.books.open(str(file_path), update_links=False)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    try:
        try:
            empirical_sheet = wb.sheets["Empirical Model"]
            empirical_rows = extract_empirical_rows(empirical_sheet, meta, file_path.name)
        except Exception:
            print(f"Skipped empirical sheet in {file_path.name} (missing or unreadable)")

        try:
            regression_sheet = wb.sheets["Regression Model"]
            regression_rows = extract_regression_rows(regression_sheet, meta, file_path.name)
        except Exception:
            print(f"Skipped regression sheet in {file_path.name} (missing or unreadable)")
    finally:
        safe_close_workbook(wb)

    return empirical_rows, regression_rows


def main() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")

    source_files = iter_source_files(in_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in source_files:
            print(f"Processing file: {file_path.name}")
            try:
                emp_rows, reg_rows = process_workbook(app, file_path)
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
                continue

            empirical_rows.extend(emp_rows)
            regression_rows.extend(reg_rows)
            processed_count += 1
            print(f"Processed file: {file_path.name}")
    finally:
        try:
            app.quit()
        except Exception:
            app.kill()

    output_path = unique_output_path(in_dir, out_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
