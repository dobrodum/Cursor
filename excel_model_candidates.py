#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from .xlsx files.

This script opens each source workbook once, processes both model sheets while the
workbook is open, and writes a single output workbook with two tabs:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import calendar
import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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

DAY_BY_MODEL_PERIOD = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?(\d{4})", re.IGNORECASE)


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


@dataclass
class SheetSnapshot:
    top_row: int
    left_col: int
    values: list[list[Any]]
    labels: dict[str, list[tuple[int, int]]]

    @property
    def bottom_row(self) -> int:
        return self.top_row + len(self.values) - 1

    @property
    def right_col(self) -> int:
        width = len(self.values[0]) if self.values else 0
        return self.left_col + width - 1

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.top_row
        col_idx = col - self.left_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        if col_idx >= len(self.values[row_idx]):
            return None
        return self.values[row_idx][col_idx]


def normalize_label(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip().lower())


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        is_percent = cleaned.endswith("%")
        if is_percent:
            cleaned = cleaned[:-1]
        try:
            parsed = float(cleaned)
        except ValueError:
            return None
        return parsed / 100.0 if is_percent else parsed
    return None


def parse_month(month_token: str) -> Optional[int]:
    token = month_token.strip().title()
    for fmt in ("%b", "%B"):
        try:
            return dt.datetime.strptime(token, fmt).month
        except ValueError:
            continue

    short = token[:3].title()
    for idx, abbr in enumerate(calendar.month_abbr):
        if idx > 0 and abbr.lower() == short.lower():
            return idx
    return None


def parse_file_labels(file_path: Path) -> Optional[FileLabels]:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]
    ticker = None

    if len(parts) >= 2:
        raw_ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if raw_ticker:
            ticker = raw_ticker

    if ticker is None:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = PERIOD_RE.search(stem)
    if ticker is None or not period_match:
        return None

    period_bucket = period_match.group(1).title()
    month_number = parse_month(period_match.group(2))
    year = int(period_match.group(3))
    if month_number is None:
        return None

    day = DAY_BY_MODEL_PERIOD[period_bucket]
    model_period = f"{period_bucket}{calendar.month_abbr[month_number]}_{year}"
    model_date = dt.date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    version = 1
    while True:
        candidate = output_path / f"{base_name}.{version}.xlsx"
        if not candidate.exists():
            return candidate
        version += 1


def get_sheet(workbook: xw.Book, name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[name]
    except Exception:
        return None


def build_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = used.options(ndim=2).value

    if values is None:
        matrix = [[None]]
    elif isinstance(values, list):
        if values and isinstance(values[0], list):
            matrix = values
        else:
            matrix = [values]
    else:
        matrix = [[values]]

    labels: dict[str, list[tuple[int, int]]] = {}
    for row_offset, row_values in enumerate(matrix):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str):
                key = normalize_label(cell_value)
                if key:
                    labels.setdefault(key, []).append((used.row + row_offset, used.column + col_offset))

    return SheetSnapshot(
        top_row=used.row,
        left_col=used.column,
        values=matrix,
        labels=labels,
    )


def find_anchor_max(snapshot: SheetSnapshot) -> Optional[tuple[int, int]]:
    exact = snapshot.labels.get("max", [])
    if exact:
        # Use the deepest "max" label, usually nearest the active output section.
        return sorted(exact, key=lambda item: (item[0], item[1]))[-1]

    loose_matches: list[tuple[int, int]] = []
    for key, coords in snapshot.labels.items():
        if key.startswith("max"):
            loose_matches.extend(coords)
    if loose_matches:
        return sorted(loose_matches, key=lambda item: (item[0], item[1]))[-1]
    return None


def find_label_cell(
    snapshot: SheetSnapshot,
    patterns: list[str],
    anchor: Optional[tuple[int, int]] = None,
    row_radius: int = 80,
    col_radius: int = 30,
) -> Optional[tuple[int, int]]:
    regexes = [re.compile(pattern, re.IGNORECASE) for pattern in patterns]
    best: Optional[tuple[tuple[int, int, int], tuple[int, int]]] = None

    for key, coords in snapshot.labels.items():
        if not any(regex.search(key) for regex in regexes):
            continue
        for row, col in coords:
            if anchor is None:
                rank = (0, row, col)
            else:
                if abs(row - anchor[0]) > row_radius or abs(col - anchor[1]) > col_radius:
                    continue
                distance = abs(row - anchor[0]) + abs(col - anchor[1])
                rank = (distance, row, col)
            if best is None or rank < best[0]:
                best = (rank, (row, col))

    return None if best is None else best[1]


def adjacent_value_coord(snapshot: SheetSnapshot, label_coord: tuple[int, int]) -> tuple[int, int]:
    row, col = label_coord
    for dr, dc in ((0, 1), (1, 0), (0, -1), (-1, 0)):
        value = snapshot.get(row + dr, col + dc)
        if value not in (None, ""):
            return (row + dr, col + dc)
    return (row, col + 1)


def resolve_value_coord(
    snapshot: SheetSnapshot,
    anchor: tuple[int, int],
    label_patterns: list[str],
    fallback_offset: tuple[int, int],
) -> tuple[int, int]:
    label_cell = find_label_cell(snapshot, label_patterns, anchor=anchor)
    if label_cell is not None:
        return adjacent_value_coord(snapshot, label_cell)
    return (anchor[0] + fallback_offset[0], anchor[1] + fallback_offset[1])


def read_cell(sheet: xw.Sheet, coord: tuple[int, int]) -> Any:
    return sheet.cells(coord[0], coord[1]).value


def read_numeric(sheet: xw.Sheet, coord: tuple[int, int]) -> Optional[float]:
    return to_float(read_cell(sheet, coord))


def trailing_contiguous(numbers: list[int]) -> list[int]:
    if not numbers:
        return []
    sorted_numbers = sorted(numbers)
    block = [sorted_numbers[-1]]
    for number in reversed(sorted_numbers[:-1]):
        if block[-1] - number == 1:
            block.append(number)
        else:
            break
    return sorted(block)


def infer_empirical_history(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
) -> Optional[tuple[int, list[int]]]:
    min_row = max(snapshot.top_row, anchor_row - 60)
    max_row = min(snapshot.bottom_row, anchor_row)
    min_col = max(snapshot.left_col, anchor_col - 50)
    max_col = max(snapshot.left_col, anchor_col - 1)
    best: Optional[tuple[int, int, list[int]]] = None

    for row in range(min_row, max_row + 1):
        numeric_cols: list[int] = []
        pct_like = 0
        for col in range(min_col, max_col + 1):
            value = to_float(snapshot.get(row, col))
            if value is None:
                continue
            numeric_cols.append(col)
            if 0.0 <= value <= 1.5:
                pct_like += 1
        if len(numeric_cols) < 4:
            continue
        score = pct_like * 2 + len(numeric_cols)
        if best is None or score > best[0]:
            best = (score, row, numeric_cols)

    if best is None:
        return None

    _, history_row, history_cols = best
    contiguous = trailing_contiguous(history_cols)
    if len(contiguous) >= 4:
        history_cols = contiguous
    return history_row, sorted(history_cols)


def infer_regression_data_rows(
    snapshot: SheetSnapshot,
    x_col: int,
    y_col: int,
    anchor_row: int,
) -> list[int]:
    min_row = max(snapshot.top_row, anchor_row - 250)
    max_row = min(snapshot.bottom_row, anchor_row - 1)
    all_rows: list[int] = []

    for row in range(min_row, max_row + 1):
        x_val = to_float(snapshot.get(row, x_col))
        y_val = to_float(snapshot.get(row, y_col))
        if x_val is not None and y_val is not None:
            all_rows.append(row)

    if not all_rows:
        return []
    return trailing_contiguous(all_rows)


def close_workbook_safely(workbook: xw.Book) -> None:
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
    except Exception as exc:
        print(f"warning: failed safe close: {exc}")


def extract_empirical_candidates(workbook: xw.Book, labels: FileLabels) -> list[dict[str, Any]]:
    sheet = get_sheet(workbook, "Empirical Model")
    if sheet is None:
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    history = infer_empirical_history(snapshot, anchor_row, anchor_col)
    if history is None:
        return []

    history_row, history_cols = history
    if not history_cols:
        return []

    # Anchor-based fallbacks if labels are missing.
    avg_pen_coord = resolve_value_coord(snapshot, anchor, [r"avg penetration"], (-2, 1))
    forecast_value_coord = resolve_value_coord(
        snapshot,
        anchor,
        [r"estimated total sold", r"forecast"],
        (-1, 1),
    )
    reported_sales_coord = resolve_value_coord(snapshot, anchor, [r"reported sales"], (2, 1))
    quarterly_sales_coord = resolve_value_coord(snapshot, anchor, [r"quarterly sales"], (1, -6))
    growth_rate_coord = resolve_value_coord(snapshot, anchor, [r"growth rate"], (3, 1))
    captured_coord = resolve_value_coord(
        snapshot,
        anchor,
        [r"sales captured", r"captured in db"],
        (4, 1),
    )
    last_quarter_coord = resolve_value_coord(snapshot, anchor, [r"last quarter"], (0, -2))

    forecast_max_coord = adjacent_value_coord(snapshot, anchor)
    min_label = find_label_cell(snapshot, [r"^min$"], anchor=anchor, row_radius=8, col_radius=8)
    forecast_min_coord = (
        adjacent_value_coord(snapshot, min_label) if min_label is not None else (anchor_row + 1, anchor_col + 1)
    )

    n_quarters = 10
    rows: list[dict[str, Any]] = []
    quarter_header_row = max(snapshot.top_row, history_row - 1)

    for quarter_step in range(1, n_quarters + 1):
        quarters_used = min(quarter_step, len(history_cols))
        start_col = history_cols[-quarters_used]
        end_col = history_cols[-1]

        avg_formula = f"=AVERAGE(R{history_row}C{start_col}:R{history_row}C{end_col})"
        sheet.cells(avg_pen_coord[0], avg_pen_coord[1]).formula2 = avg_formula
        workbook.app.calculate()

        avg_penetration_pct = read_numeric(sheet, avg_pen_coord)
        forecast_value = read_numeric(sheet, forecast_value_coord)
        reported_sales = read_numeric(sheet, reported_sales_coord)
        quarterly_sales = read_numeric(sheet, quarterly_sales_coord)
        growth_rate_pct = read_numeric(sheet, growth_rate_coord)
        sales_captured_in_db_pct = read_numeric(sheet, captured_coord)
        forecast_max = read_numeric(sheet, forecast_max_coord)
        forecast_min = read_numeric(sheet, forecast_min_coord)

        if forecast_value is None and quarterly_sales is not None and avg_penetration_pct not in (None, 0.0):
            forecast_value = quarterly_sales / avg_penetration_pct

        actual_value = reported_sales
        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )

        last_quarter_used = read_cell(sheet, last_quarter_coord)
        if last_quarter_used in (None, ""):
            last_quarter_used = snapshot.get(quarter_header_row, start_col)

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": labels.source_file,
            }
        )

    return rows


def extract_regression_candidates(workbook: xw.Book, labels: FileLabels) -> list[dict[str, Any]]:
    sheet = get_sheet(workbook, "Regression Model")
    if sheet is None:
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_rows = infer_regression_data_rows(snapshot, x_col, y_col, anchor_row)
    if len(data_rows) < 2:
        return []

    forecast_total_coord = resolve_value_coord(
        snapshot,
        anchor,
        [r"tot fcst.*w.?o.*sa", r"forecast.*without.*sa"],
        (-1, 1),
    )
    actual_value_coord = resolve_value_coord(snapshot, anchor, [r"reported sales", r"actual"], (2, 1))
    forecast_max_coord = adjacent_value_coord(snapshot, anchor)
    min_label = find_label_cell(snapshot, [r"^min$"], anchor=anchor, row_radius=8, col_radius=8)
    forecast_min_coord = (
        adjacent_value_coord(snapshot, min_label) if min_label is not None else (anchor_row + 1, anchor_col + 1)
    )

    # Use temporary scratch cells to avoid touching source formulas.
    scratch_col = snapshot.right_col + 3
    intercept_coord = (anchor_row, scratch_col)
    slope_coord = (anchor_row + 1, scratch_col)

    rows: list[dict[str, Any]] = []
    prev_signature: Optional[tuple[Any, ...]] = None
    max_quarters = min(10, len(data_rows))

    for quarters_used in range(2, max_quarters + 1):
        start_row = data_rows[-quarters_used]
        end_row = data_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        sheet.cells(intercept_coord[0], intercept_coord[1]).formula2 = intercept_formula
        sheet.cells(slope_coord[0], slope_coord[1]).formula2 = slope_formula
        workbook.app.calculate()

        intercept = read_numeric(sheet, intercept_coord)
        slope = read_numeric(sheet, slope_coord)
        forecast_value = read_numeric(sheet, forecast_total_coord)
        forecast_max = read_numeric(sheet, forecast_max_coord)
        forecast_min = read_numeric(sheet, forecast_min_coord)
        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )

        if forecast_value is None and intercept is not None and slope is not None:
            next_x = to_float(snapshot.get(end_row + 1, x_col))
            if next_x is None:
                next_x = to_float(snapshot.get(end_row, x_col))
            if next_x is not None:
                forecast_value = intercept + (slope * next_x)

        actual_raw = read_cell(sheet, actual_value_coord)
        actual_numeric = to_float(actual_raw)
        actual_value = actual_numeric if actual_numeric is not None else (actual_raw or None)

        signature = (
            quarters_used,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )

        # Skip duplicate tail rows caused by equivalent calculation windows.
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": quarters_used,
                "num_quarters_used": quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": labels.source_file,
            }
        )

    return rows


def write_sheet(workbook: Workbook, title: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
    sheet = workbook.create_sheet(title=title)
    sheet.append(headers)
    for row_data in rows:
        sheet.append([row_data.get(header) for header in headers])

    for header_cell in sheet[1]:
        header_cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    last_row = max(1, sheet.max_row)
    last_col = len(headers)
    sheet.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_row}"

    for col_idx, header in enumerate(headers, start=1):
        width = len(header)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(width + 2, 48)


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    workbook.save(output_file)


def run() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file = build_output_path(input_path, output_path)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            labels = parse_file_labels(file_path)
            if labels is None:
                print(f"skipped file: {file_path.name} (filename parse failed)")
                continue

            print(f"processed file: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(workbook, labels))
                regression_rows.extend(extract_regression_candidates(workbook, labels))
                processed_files += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (workbook error: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
