#!/usr/bin/env python3
"""Extract empirical/regression candidate rows from model workbooks."""

from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Iterable

import xlwings as xw

# =========================
# Runtime configuration
# =========================
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetGrid:
    top_row: int
    left_col: int
    values: list[list[Any]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        if not self.values:
            return 0
        return len(self.values[0])

    @property
    def bottom_row(self) -> int:
        return self.top_row + self.row_count - 1

    @property
    def right_col(self) -> int:
        return self.left_col + self.col_count - 1


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def as_output_value(value: Any) -> Any:
    if value is None:
        return ""
    return value


def parse_month(month_token: str) -> int | None:
    token = month_token.strip().lower()
    if not token:
        return None
    for month_num in range(1, 13):
        if token.startswith(calendar.month_name[month_num].lower()):
            return month_num
        if token.startswith(calendar.month_abbr[month_num].lower()):
            return month_num
    return None


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_token = parts[2] if len(parts) >= 3 else (parts[-1] if parts else "")
    period_token = period_token.split("_")[0].strip()

    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, re.IGNORECASE)
    model_period = period_token or "unknown_period"
    model_date = ""

    if match:
        period_name = match.group(1).title()
        month_token = match.group(2)
        year = int(match.group(3))
        month_num = parse_month(month_token)
        if month_num is not None:
            month_abbr = datetime(year, month_num, 1).strftime("%b")
            model_period = f"{period_name}{month_abbr}_{year}"
            model_date = date(year, month_num, DAY_BY_PERIOD[period_name.lower()]).isoformat()
        else:
            model_period = f"{period_name}{month_token}_{year}"

    if not ticker and len(parts) >= 1:
        ticker = parts[0].replace(" ", "_")

    model = f"{ticker}_{model_period}" if ticker else model_period
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Fallback for environments where workbook.close(save=False) is unsupported.
    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def get_sheet(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    wanted = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == wanted:
            return sheet
    return None


def capture_sheet_grid(sheet: xw.Sheet) -> SheetGrid:
    used = sheet.used_range
    top_row = used.row
    left_col = used.column
    row_count = used.rows.count
    col_count = used.columns.count
    raw_values = used.value

    if row_count == 1 and col_count == 1:
        values = [[raw_values]]
    elif row_count == 1:
        values = [list(raw_values or [])]
    elif col_count == 1:
        values = [[item] for item in (raw_values or [])]
    else:
        values = [list(row) for row in (raw_values or [])]

    if not values:
        values = [[None]]
    return SheetGrid(top_row=top_row, left_col=left_col, values=values)


def grid_value(grid: SheetGrid, row: int, col: int) -> Any:
    if row < grid.top_row or col < grid.left_col:
        return None
    row_index = row - grid.top_row
    col_index = col - grid.left_col
    if row_index >= grid.row_count or col_index >= grid.col_count:
        return None
    return grid.values[row_index][col_index]


def iter_grid_cells(grid: SheetGrid) -> Iterable[tuple[int, int, Any]]:
    for row_index, row in enumerate(grid.values):
        for col_index, value in enumerate(row):
            yield grid.top_row + row_index, grid.left_col + col_index, value


def find_max_anchor(grid: SheetGrid) -> tuple[int, int] | None:
    candidates: list[tuple[float, int, int]] = []
    for row, col, value in iter_grid_cells(grid):
        if normalize_text(value) != "max":
            continue
        score = 0.0
        if to_float(grid_value(grid, row, col + 1)) is not None:
            score += 3.0
        if normalize_text(grid_value(grid, row + 1, col)) == "min":
            score += 1.0
        score += row / 10000.0
        candidates.append((score, row, col))

    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1], candidates[0][2]


def set_formula2(target: xw.Range, formula: str) -> None:
    target.formula2 = formula


def find_column_by_keywords(
    grid: SheetGrid,
    anchor_col: int,
    keyword_groups: list[tuple[str, ...]],
    fallback_col: int,
) -> int:
    hits: list[tuple[int, int, int]] = []
    for row, col, value in iter_grid_cells(grid):
        text = normalize_text(value)
        if not text:
            continue
        if any(all(keyword in text for keyword in group) for group in keyword_groups):
            hits.append((abs(col - anchor_col), row, col))
    if hits:
        hits.sort()
        return hits[0][2]
    return fallback_col


def find_numeric_value_by_label(
    grid: SheetGrid,
    anchor_row: int,
    anchor_col: int,
    labels: list[str],
) -> float | None:
    candidates: list[tuple[int, int, int]] = []
    labels_lower = [label.lower() for label in labels]
    for row, col, value in iter_grid_cells(grid):
        text = normalize_text(value)
        if not text:
            continue
        if any(label in text for label in labels_lower):
            distance = abs(row - anchor_row) + abs(col - anchor_col)
            candidates.append((distance, row, col))

    candidates.sort()
    for _, row, col in candidates:
        for row_offset, col_offset in ((0, 1), (1, 0), (0, 2), (2, 0), (0, -1)):
            numeric_value = to_float(grid_value(grid, row + row_offset, col + col_offset))
            if numeric_value is not None:
                return numeric_value
    return None


def stringify_quarter(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def numeric_sum(values: Iterable[Any]) -> float | None:
    total = 0.0
    count = 0
    for value in values:
        num = to_float(value)
        if num is None:
            continue
        total += num
        count += 1
    if count == 0:
        return None
    return total


def normalized_dedupe_value(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value, 10)


def extract_empirical_rows(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(workbook, "Empirical Model")
    if sheet is None:
        return []

    grid = capture_sheet_grid(sheet)
    anchor = find_max_anchor(grid)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    quarter_col = find_column_by_keywords(
        grid,
        anchor_col,
        keyword_groups=[("quarter",), ("period",)],
        fallback_col=anchor_col - 11,
    )
    penetration_col = find_column_by_keywords(
        grid,
        anchor_col,
        keyword_groups=[("penetration",)],
        fallback_col=anchor_col - 5,
    )
    quarterly_sales_col = find_column_by_keywords(
        grid,
        anchor_col,
        keyword_groups=[("quarterly", "sales"), ("db", "sales")],
        fallback_col=anchor_col - 7,
    )
    reported_sales_col = find_column_by_keywords(
        grid,
        anchor_col,
        keyword_groups=[("reported", "sales"), ("actual", "sales"), ("total", "sold")],
        fallback_col=anchor_col - 6,
    )

    candidate_rows: list[int] = []
    row_upper_bound = min(anchor_row - 1, grid.bottom_row)
    for row in range(grid.top_row, row_upper_bound + 1):
        penetration_value = to_float(grid_value(grid, row, penetration_col))
        if penetration_value is None:
            continue
        quarterly_value = to_float(grid_value(grid, row, quarterly_sales_col))
        reported_value = to_float(grid_value(grid, row, reported_sales_col))
        if quarterly_value is None and reported_value is None:
            continue
        candidate_rows.append(row)

    if not candidate_rows:
        return []

    max_quarters = min(N_QUARTERS, len(candidate_rows))
    forecast_max_static = to_float(grid_value(grid, anchor_row, anchor_col + 1))
    forecast_min_static = to_float(grid_value(grid, anchor_row + 1, anchor_col + 1))
    growth_rate_pct = find_numeric_value_by_label(
        grid,
        anchor_row,
        anchor_col,
        labels=["growth rate", "growth %"],
    )
    sales_captured_pct = find_numeric_value_by_label(
        grid,
        anchor_row,
        anchor_col,
        labels=["sales captured in db", "captured in db", "captured"],
    )

    temp_col = max(grid.right_col + 2, anchor_col + 5)
    avg_pen_cell = sheet.range((anchor_row, temp_col))
    forecast_cell = sheet.range((anchor_row + 1, temp_col))

    output_rows: list[dict[str, Any]] = []
    for num_quarters in range(1, max_quarters + 1):
        selected_rows = candidate_rows[-num_quarters:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        avg_formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        )
        forecast_formula = (
            f"=IF(R{anchor_row}C{temp_col}=0,\"\","
            f"SUM(R{start_row}C{quarterly_sales_col}:R{end_row}C{quarterly_sales_col})/"
            f"R{anchor_row}C{temp_col})"
        )

        set_formula2(avg_pen_cell, avg_formula)
        set_formula2(forecast_cell, forecast_formula)
        workbook.app.calculate()

        avg_penetration_pct = to_float(avg_pen_cell.value)
        forecast_value = to_float(forecast_cell.value)
        quarterly_sales = numeric_sum(
            grid_value(grid, row, quarterly_sales_col) for row in selected_rows
        )
        reported_sales = numeric_sum(
            grid_value(grid, row, reported_sales_col) for row in selected_rows
        )

        forecast_max = forecast_max_static if forecast_max_static is not None else forecast_value
        forecast_min = forecast_min_static if forecast_min_static is not None else forecast_value
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        output_rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": stringify_quarter(grid_value(grid, end_row, quarter_col)),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    avg_pen_cell.value = None
    forecast_cell.value = None
    return output_rows


def extract_regression_rows(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(workbook, "Regression Model")
    if sheet is None:
        return []

    grid = capture_sheet_grid(sheet)
    anchor = find_max_anchor(grid)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    regression_rows: list[int] = []
    row_upper_bound = min(anchor_row - 1, grid.bottom_row)
    for row in range(grid.top_row, row_upper_bound + 1):
        x_value = to_float(grid_value(grid, row, x_col))
        y_value = to_float(grid_value(grid, row, y_col))
        if x_value is None or y_value is None:
            continue
        regression_rows.append(row)

    if len(regression_rows) < 2:
        return []

    max_quarters = min(N_QUARTERS, len(regression_rows))
    output_rows: list[dict[str, Any]] = []
    previous_signature: tuple[float | None, ...] | None = None

    forecast_max_static = to_float(grid_value(grid, anchor_row, anchor_col + 1))
    forecast_min_static = to_float(grid_value(grid, anchor_row + 1, anchor_col + 1))
    actual_value = find_numeric_value_by_label(
        grid,
        anchor_row,
        anchor_col,
        labels=["reported sales", "actual sales"],
    )

    target_x = to_float(grid_value(grid, anchor_row, x_col))
    if target_x is None:
        last_x = to_float(grid_value(grid, regression_rows[-1], x_col))
        target_x = last_x + 1 if last_x is not None else None

    temp_col = max(grid.right_col + 2, anchor_col + 5)
    intercept_cell = sheet.range((anchor_row, temp_col))
    slope_cell = sheet.range((anchor_row + 1, temp_col))
    forecast_cell = sheet.range((anchor_row + 2, temp_col))

    for num_quarters in range(2, max_quarters + 1):
        selected_rows = regression_rows[-num_quarters:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        if target_x is not None:
            set_formula2(
                forecast_cell,
                f"=R{anchor_row}C{temp_col}+R{anchor_row + 1}C{temp_col}*{target_x}",
            )
        else:
            forecast_cell.value = None

        workbook.app.calculate()

        intercept_value = to_float(intercept_cell.value)
        slope_value = to_float(slope_cell.value)
        forecast_total_without_sa = to_float(forecast_cell.value)
        forecast_max = (
            forecast_max_static if forecast_max_static is not None else forecast_total_without_sa
        )
        forecast_min = (
            forecast_min_static if forecast_min_static is not None else forecast_total_without_sa
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            normalized_dedupe_value(intercept_value),
            normalized_dedupe_value(slope_value),
            normalized_dedupe_value(forecast_total_without_sa),
            normalized_dedupe_value(forecast_max),
            normalized_dedupe_value(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        output_rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    intercept_cell.value = None
    slope_cell.value = None
    forecast_cell.value = None
    return output_rows


def apply_sheet_formatting(sheet: xw.Sheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    max_row = max(1, len(rows) + 1)
    max_col = len(headers)

    header_range = sheet.range((1, 1), (1, max_col))
    header_range.api.Font.Bold = True

    data_range = sheet.range((1, 1), (max_row, max_col))
    data_range.api.AutoFilter()

    # Freeze the top row.
    sheet.activate()
    window = sheet.book.app.api.ActiveWindow
    window.SplitRow = 1
    window.SplitColumn = 0
    window.FreezePanes = True

    for col_index, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            text = str(as_output_value(row.get(header, "")))
            max_len = max(max_len, len(text))
        sheet.range((1, col_index)).column_width = min(max(12, max_len + 2), 60)


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = app.books.add()
    try:
        workbook.sheets[0].name = "empirical_candidates"
        workbook.sheets.add("regression_candidates", after=workbook.sheets[0])
        while len(workbook.sheets) > 2:
            workbook.sheets[-1].delete()

        empirical_sheet = workbook.sheets["empirical_candidates"]
        regression_sheet = workbook.sheets["regression_candidates"]

        empirical_sheet.range((1, 1)).value = EMPIRICAL_COLUMNS
        if empirical_rows:
            empirical_sheet.range((2, 1)).value = [
                [as_output_value(row.get(column)) for column in EMPIRICAL_COLUMNS]
                for row in empirical_rows
            ]

        regression_sheet.range((1, 1)).value = REGRESSION_COLUMNS
        if regression_rows:
            regression_sheet.range((2, 1)).value = [
                [as_output_value(row.get(column)) for column in REGRESSION_COLUMNS]
                for row in regression_rows
            ]

        apply_sheet_formatting(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
        apply_sheet_formatting(regression_sheet, REGRESSION_COLUMNS, regression_rows)

        workbook.save(str(output_path))
    finally:
        safe_close_workbook(workbook)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    source_files = sorted(input_dir.iterdir(), key=lambda item: item.name.lower())
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"SKIPPED: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED: {file_path.name} (not an .xlsx file)")
                continue

            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)

                empirical_result = extract_empirical_rows(workbook, metadata, file_path.name)
                regression_result = extract_regression_rows(workbook, metadata, file_path.name)
                empirical_rows.extend(empirical_result)
                regression_rows.extend(regression_result)
                processed_files += 1
                print(f"PROCESSED: {file_path.name}")
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)

        if app is None:
            raise RuntimeError("Excel application failed to start.")
        write_output_workbook(app, output_path, empirical_rows, regression_rows)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {processed_files}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
