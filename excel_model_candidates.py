#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable folders
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

MONTHS = {
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

MODEL_DAY = {"early": 5, "mid": 15, "late": 25}

PERIOD_PATTERN = re.compile(
    r"(?i)\b(early|mid|late)\s*([a-z]{3,9})\s*[_-]?\s*(\d{4})\b"
)


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetGrid:
    values: list[list[Any]]
    first_row: int
    first_col: int
    nrows: int
    ncols: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_2d_matrix(values: Any, nrows: int, ncols: int) -> list[list[Any]]:
    if nrows <= 0 or ncols <= 0:
        return []
    if values is None:
        matrix: list[list[Any]] = [[None] * ncols for _ in range(nrows)]
    elif isinstance(values, (list, tuple)):
        outer = list(values)
        if outer and isinstance(outer[0], (list, tuple)):
            matrix = [list(row) for row in outer]
        else:
            matrix = [list(outer)]
    else:
        matrix = [[values]]

    if len(matrix) < nrows:
        matrix.extend([[None] * ncols for _ in range(nrows - len(matrix))])
    elif len(matrix) > nrows:
        matrix = matrix[:nrows]

    for i, row in enumerate(matrix):
        if len(row) < ncols:
            row.extend([None] * (ncols - len(row)))
        elif len(row) > ncols:
            matrix[i] = row[:ncols]
    return matrix


def read_sheet_grid(sheet: xw.Sheet) -> SheetGrid:
    used = sheet.used_range
    nrows = used.rows.count
    ncols = used.columns.count
    values = to_2d_matrix(used.value, nrows, ncols)
    return SheetGrid(values=values, first_row=used.row, first_col=used.column, nrows=nrows, ncols=ncols)


def cell_value(grid: SheetGrid, row: int, col: int) -> Any:
    if row < grid.first_row or col < grid.first_col:
        return None
    row_idx = row - grid.first_row
    col_idx = col - grid.first_col
    if row_idx >= grid.nrows or col_idx >= grid.ncols:
        return None
    return grid.values[row_idx][col_idx]


def find_max_anchor(grid: SheetGrid) -> tuple[int, int] | None:
    fallback: tuple[int, int] | None = None
    for r_offset, row in enumerate(grid.values):
        for c_offset, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            row_abs = grid.first_row + r_offset
            col_abs = grid.first_col + c_offset
            if fallback is None:
                fallback = (row_abs, col_abs)
            right_cell = normalize_text(cell_value(grid, row_abs, col_abs + 1))
            if right_cell == "min":
                return (row_abs, col_abs)
    return fallback


def header_entries(grid: SheetGrid, header_row: int) -> list[tuple[int, str]]:
    entries: list[tuple[int, str]] = []
    for col in range(grid.first_col, grid.first_col + grid.ncols):
        label = normalize_text(cell_value(grid, header_row, col))
        if label:
            entries.append((col, label))
    return entries


def pick_header_col(headers: list[tuple[int, str]], patterns: list[tuple[str, ...]]) -> int | None:
    for pattern in patterns:
        for col, header in headers:
            if all(token in header for token in pattern):
                return col
    return None


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    num = to_float(value)
    if num is None:
        return None
    return int(round(num))


def text_or_none(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_source = parts[2] if len(parts) >= 3 else stem
    period_token = period_source.split("_")[0]

    model_period = period_token
    model_date = ""

    match = PERIOD_PATTERN.search(period_token)
    if match:
        timing = match.group(1).lower()
        month_token = match.group(2).lower()[:3]
        year = int(match.group(3))
        month_num = MONTHS.get(month_token)
        day = MODEL_DAY.get(timing)
        if month_num is not None and day is not None:
            model_period = f"{timing.title()}{month_token.title()}_{year}"
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker else model_period
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    if not candidate.exists():
        return candidate
    idx = 1
    while True:
        candidate = out_dir / f"{base}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def safe_close_source_workbook(wb: xw.Book) -> None:
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


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        try:
            cell.api.Formula2R1C1 = formula
        except Exception:
            cell.formula = formula


def calc_range_width(max_value: float | None, min_value: float | None) -> float | None:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def estimate_total_sold(quarterly_sales: float | None, avg_penetration_pct: float | None) -> float | None:
    if quarterly_sales is None or avg_penetration_pct is None:
        return None
    ratio = avg_penetration_pct / 100.0 if avg_penetration_pct > 1 else avg_penetration_pct
    if ratio <= 0:
        return None
    return quarterly_sales / ratio


def extract_empirical_rows(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    grid = read_sheet_grid(sheet)
    anchor = find_max_anchor(grid)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = header_entries(grid, anchor_row)
    last_row = grid.first_row + grid.nrows - 1
    data_start_row = anchor_row + 1

    num_quarters_col = pick_header_col(headers, [("num", "quarter"), ("quarters", "used")])
    last_quarter_col = pick_header_col(headers, [("last", "quarter")])
    avg_pen_col = pick_header_col(headers, [("avg", "penetration"), ("average", "penetration")])
    quarterly_sales_col = pick_header_col(headers, [("quarterly", "sales"), ("sales", "db")])
    reported_sales_col = pick_header_col(headers, [("reported", "sales"), ("actual", "sales")])
    growth_rate_col = pick_header_col(headers, [("growth", "rate")])
    captured_pct_col = pick_header_col(headers, [("sales", "captured"), ("captured", "db")])
    forecast_value_col = pick_header_col(
        headers,
        [("estimated", "total", "sold"), ("est", "total", "sold"), ("forecast", "value")],
    )
    forecast_min_col = pick_header_col(headers, [("min",)])

    max_col = anchor_col
    min_col = forecast_min_col if forecast_min_col is not None else anchor_col + 1

    penetration_source_col = captured_pct_col or avg_pen_col
    penetration_rows: list[int] = []
    if penetration_source_col is not None:
        for row in range(data_start_row, last_row + 1):
            if to_float(cell_value(grid, row, penetration_source_col)) is not None:
                penetration_rows.append(row)

    scratch_col = max(anchor_col + 20, grid.first_col + grid.ncols + 1)
    scratch_cell = sheet.cells(anchor_row, scratch_col)

    extracted: list[dict[str, Any]] = []
    for idx in range(10):
        row = data_start_row + idx

        num_quarters_used = to_int(cell_value(grid, row, num_quarters_col)) if num_quarters_col else None
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        last_quarter_used = text_or_none(cell_value(grid, row, last_quarter_col)) if last_quarter_col else None
        forecast_max = to_float(cell_value(grid, row, max_col))
        forecast_min = to_float(cell_value(grid, row, min_col))

        quarterly_sales = (
            to_float(cell_value(grid, row, quarterly_sales_col)) if quarterly_sales_col is not None else None
        )
        reported_sales = (
            to_float(cell_value(grid, row, reported_sales_col)) if reported_sales_col is not None else None
        )
        growth_rate_pct = to_float(cell_value(grid, row, growth_rate_col)) if growth_rate_col is not None else None
        sales_captured_in_db_pct = (
            to_float(cell_value(grid, row, captured_pct_col)) if captured_pct_col is not None else None
        )
        avg_penetration_pct = to_float(cell_value(grid, row, avg_pen_col)) if avg_pen_col is not None else None

        if penetration_source_col is not None and len(penetration_rows) >= num_quarters_used:
            start_row = penetration_rows[-num_quarters_used]
            end_row = penetration_rows[-1]
            avg_formula = (
                f"=AVERAGE(R{start_row}C{penetration_source_col}:R{end_row}C{penetration_source_col})"
            )
            set_formula2_r1c1(scratch_cell, avg_formula)
            wb.app.calculate()
            computed_avg = to_float(scratch_cell.value)
            if computed_avg is not None:
                avg_penetration_pct = computed_avg

        forecast_value = (
            to_float(cell_value(grid, row, forecast_value_col)) if forecast_value_col is not None else None
        )
        if forecast_value is None:
            forecast_value = estimate_total_sold(quarterly_sales, avg_penetration_pct)

        actual_value = reported_sales
        if sales_captured_in_db_pct is None and quarterly_sales is not None and reported_sales not in (None, 0):
            sales_captured_in_db_pct = quarterly_sales / reported_sales

        range_width = calc_range_width(forecast_max, forecast_min)

        has_payload = any(
            value is not None
            for value in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            ]
        )
        if not has_payload:
            continue

        extracted.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
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
                "source_file": source_file,
            }
        )

    try:
        scratch_cell.value = None
    except Exception:
        pass
    return extracted


def _regression_signature(row: dict[str, Any]) -> tuple[Any, ...]:
    def rounded(value: Any) -> Any:
        if isinstance(value, float):
            return round(value, 10)
        return value

    return (
        rounded(row.get("forecast_value")),
        rounded(row.get("forecast_max")),
        rounded(row.get("forecast_min")),
        rounded(row.get("intercept")),
        rounded(row.get("slope")),
    )


def extract_regression_rows(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    grid = read_sheet_grid(sheet)
    anchor = find_max_anchor(grid)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = header_entries(grid, anchor_row)
    last_row = grid.first_row + grid.nrows - 1
    data_start_row = anchor_row + 1

    x_col = anchor_col - 11
    y_col = anchor_col - 7

    points: list[tuple[int, float, float]] = []
    for row in range(data_start_row, last_row + 1):
        x_value = to_float(cell_value(grid, row, x_col))
        y_value = to_float(cell_value(grid, row, y_col))
        if x_value is None or y_value is None:
            continue
        points.append((row, x_value, y_value))
    if not points:
        return []

    num_quarters_col = pick_header_col(headers, [("num", "quarter"), ("quarters", "used")])
    forecast_value_col = pick_header_col(
        headers,
        [("tot", "fcst", "wo", "sa"), ("tot", "forecast", "without", "sa"), ("forecast", "value")],
    )
    actual_value_col = pick_header_col(headers, [("actual", "sales"), ("reported", "sales"), ("actual",)])
    forecast_min_col = pick_header_col(headers, [("min",)])
    max_col = anchor_col
    min_col = forecast_min_col if forecast_min_col is not None else anchor_col + 1

    scratch_col = max(anchor_col + 20, grid.first_col + grid.ncols + 1)
    intercept_cell = sheet.cells(anchor_row, scratch_col)
    slope_cell = sheet.cells(anchor_row, scratch_col + 1)

    extracted: list[dict[str, Any]] = []
    prior_signature: tuple[Any, ...] | None = None

    for idx in range(min(10, len(points))):
        n_quarters = idx + 1
        start_row = points[-n_quarters][0]
        end_row = points[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        row = data_start_row + idx
        num_quarters_used = to_int(cell_value(grid, row, num_quarters_col)) if num_quarters_col else None
        if num_quarters_used is None:
            num_quarters_used = n_quarters

        forecast_value = (
            to_float(cell_value(grid, row, forecast_value_col)) if forecast_value_col is not None else None
        )
        if forecast_value is None and intercept is not None and slope is not None:
            x_latest = points[-1][1]
            forecast_value = intercept + slope * x_latest

        actual_value = to_float(cell_value(grid, row, actual_value_col)) if actual_value_col is not None else None
        forecast_max = to_float(cell_value(grid, row, max_col))
        forecast_min = to_float(cell_value(grid, row, min_col))
        range_width = calc_range_width(forecast_max, forecast_min)

        candidate = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
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

        signature = _regression_signature(candidate)
        if signature == prior_signature:
            continue
        prior_signature = signature
        extracted.append(candidate)

    try:
        intercept_cell.value = None
        slope_cell.value = None
    except Exception:
        pass
    return extracted


def write_rows(sheet: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(col) for col in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    last_col = get_column_letter(len(columns))
    sheet.auto_filter.ref = f"A1:{last_col}{sheet.max_row}"

    for idx, col_name in enumerate(columns, start=1):
        width = len(col_name)
        for row in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row, column=idx).value
            if value is None:
                continue
            if isinstance(value, float):
                value_text = f"{value:.8g}"
            else:
                value_text = str(value)
            width = max(width, len(value_text))
        sheet.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 12), 44)


def collect_source_files(in_dir: Path) -> list[Path]:
    files: list[Path] = []
    for entry in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
        if not entry.is_file():
            print(f"Skipped file: {entry.name} (not a file)")
            continue
        if entry.name.startswith("~"):
            print(f"Skipped file: {entry.name} (temporary file)")
            continue
        if entry.suffix.lower() != ".xlsx":
            print(f"Skipped file: {entry.name} (not .xlsx)")
            continue
        files.append(entry)
    return files


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise SystemExit(f"Input folder does not exist: {in_dir}")

    source_files = collect_source_files(in_dir)
    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files:
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                processed_count += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    empirical_sheet = out_wb.create_sheet("empirical_candidates")
    regression_sheet = out_wb.create_sheet("regression_candidates")
    write_rows(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
