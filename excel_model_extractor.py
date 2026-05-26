#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from itertools import count
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure paths for your environment.
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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

MONTH_NUMBERS = {
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass
class UsedGrid:
    values: list[list[Any]]
    first_row: int
    first_col: int
    row_count: int
    col_count: int

    @property
    def last_row(self) -> int:
        return self.first_row + self.row_count - 1

    @property
    def last_col(self) -> int:
        return self.first_col + self.col_count - 1


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return " ".join(value.strip().lower().split())
    return str(value).strip().lower()


def coerce_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    text = text.replace(",", "")
    text = text.replace("%", "")
    try:
        number = float(text)
        return -number if negative else number
    except ValueError:
        return None


def parse_model_metadata(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    pattern = re.compile(
        r"\s-\s(?P<ticker>[A-Za-z0-9]+)\s-\s(?P<period>(Early|Mid|Late)[A-Za-z]+(?P<year>\d{4}))",
        re.IGNORECASE,
    )
    match = pattern.search(stem)

    ticker = ""
    model_period = ""
    model_date = ""

    if match:
        ticker = match.group("ticker").upper()
        period_token = match.group("period")
        token_match = re.match(
            r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})",
            period_token,
            flags=re.IGNORECASE,
        )
        if token_match:
            phase = token_match.group("phase").title()
            month_token = token_match.group("month")[:3].lower()
            year = int(token_match.group("year"))
            month_num = MONTH_NUMBERS.get(month_token)
            model_period = f"{phase}{month_token.title()}_{year}"
            if month_num:
                day = DAY_BY_PERIOD[phase.lower()]
                model_date = date(year, month_num, day).isoformat()

    if not ticker:
        parts = [part.strip() for part in stem.split("-")]
        if len(parts) >= 2:
            ticker = parts[1].upper().replace(" ", "")
    if not ticker:
        ticker = "UNKNOWN"

    if not model_period:
        model_period = "UNKNOWN_PERIOD"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate
    for idx in count(1):
        candidate = output_path / f"{stem}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
    raise RuntimeError("Unable to find a free output filename.")


def collect_input_files(input_path: Path) -> tuple[list[Path], list[tuple[str, str]]]:
    files_to_process: list[Path] = []
    skipped: list[tuple[str, str]] = []

    for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            skipped.append((file_path.name, "not a file"))
            continue
        if file_path.name.startswith("~"):
            skipped.append((file_path.name, "temporary Excel file"))
            continue
        if file_path.suffix.lower() != ".xlsx":
            skipped.append((file_path.name, "not an .xlsx file"))
            continue
        files_to_process.append(file_path)

    return files_to_process, skipped


def get_used_grid(sheet: xw.Sheet) -> UsedGrid:
    used = sheet.used_range
    first_row = int(used.row)
    first_col = int(used.column)
    row_count = int(used.rows.count)
    col_count = int(used.columns.count)
    values = used.options(ndim=2).value

    if row_count == 1 and col_count == 1:
        values_2d = [[values]]
    elif row_count == 1:
        values_2d = [values]
    elif col_count == 1:
        values_2d = [[value] for value in values]
    else:
        values_2d = values

    return UsedGrid(
        values=values_2d,
        first_row=first_row,
        first_col=first_col,
        row_count=row_count,
        col_count=col_count,
    )


def find_anchor(grid: UsedGrid, anchor_label: str = "max") -> tuple[int, int] | None:
    target = normalize_label(anchor_label)
    for row_idx, row_values in enumerate(grid.values, start=grid.first_row):
        for col_offset, cell_value in enumerate(row_values):
            if normalize_label(cell_value) == target:
                return row_idx, grid.first_col + col_offset
    return None


def collect_header_candidates(grid: UsedGrid, anchor_row: int) -> dict[str, int]:
    headers: dict[str, int] = {}
    start = max(grid.first_row, anchor_row - 2)
    end = min(grid.last_row, anchor_row + 2)

    for row in range(start, end + 1):
        row_values = grid.values[row - grid.first_row]
        for col_offset, cell_value in enumerate(row_values):
            normalized = normalize_label(cell_value)
            if not normalized:
                continue
            if any(ch.isalpha() for ch in normalized):
                col = grid.first_col + col_offset
                headers.setdefault(normalized, col)
    return headers


def find_col_by_keywords(
    headers: dict[str, int],
    include_keywords: tuple[str, ...],
    exclude_keywords: tuple[str, ...] = (),
) -> int | None:
    for label, col in headers.items():
        if all(keyword in label for keyword in include_keywords) and not any(
            keyword in label for keyword in exclude_keywords
        ):
            return col
    return None


def get_cell_value(sheet: xw.Sheet, row: int, col: int | None) -> Any:
    if col is None or row <= 0 or col <= 0:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def get_numeric_cell(sheet: xw.Sheet, row: int, col: int | None) -> float | None:
    return coerce_number(get_cell_value(sheet, row, col))


def set_r1c1_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_workbook(workbook: xw.Book) -> None:
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
        workbook.app.display_alerts = False
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        workbook.api.Close(False)


def infer_empirical_columns(anchor_col: int, headers: dict[str, int]) -> dict[str, int | None]:
    return {
        "num_quarters_used": find_col_by_keywords(headers, ("num", "quarter"))
        or find_col_by_keywords(headers, ("quarters", "used"))
        or (anchor_col - 8),
        "last_quarter_used": find_col_by_keywords(headers, ("last", "quarter"))
        or (anchor_col - 7),
        "forecast_value": find_col_by_keywords(headers, ("estimated", "sold"))
        or find_col_by_keywords(headers, ("forecast", "value"))
        or find_col_by_keywords(headers, ("total", "sold"))
        or (anchor_col - 1),
        "forecast_max": anchor_col,
        "forecast_min": find_col_by_keywords(headers, ("min",)) or (anchor_col + 1),
        "avg_penetration_pct": find_col_by_keywords(headers, ("avg", "penetration"))
        or find_col_by_keywords(headers, ("average", "penetration"))
        or (anchor_col - 6),
        "quarterly_sales": find_col_by_keywords(headers, ("quarterly", "sales"))
        or find_col_by_keywords(headers, ("quarter", "sales"))
        or (anchor_col - 4),
        "reported_sales": find_col_by_keywords(headers, ("reported", "sales")) or (anchor_col - 2),
        "growth_rate_pct": find_col_by_keywords(headers, ("growth", "rate"))
        or find_col_by_keywords(headers, ("growth",))
        or (anchor_col - 3),
        "sales_captured_in_db_pct": find_col_by_keywords(headers, ("captured", "db"))
        or find_col_by_keywords(headers, ("captured",))
        or (anchor_col - 5),
    }


def infer_regression_columns(anchor_col: int, headers: dict[str, int]) -> dict[str, int | None]:
    return {
        "num_quarters_used": find_col_by_keywords(headers, ("num", "quarter"))
        or find_col_by_keywords(headers, ("quarters", "used"))
        or (anchor_col - 6),
        "forecast_value": find_col_by_keywords(headers, ("tot", "fcst"))
        or find_col_by_keywords(headers, ("forecast",))
        or (anchor_col - 1),
        "actual_value": find_col_by_keywords(headers, ("actual", "value"))
        or find_col_by_keywords(headers, ("actual",)),
        "forecast_max": anchor_col,
        "forecast_min": find_col_by_keywords(headers, ("min",)) or (anchor_col + 1),
    }


def extract_empirical_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    source_file: str,
    metadata: dict[str, str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    grid = get_used_grid(sheet)
    anchor = find_anchor(grid, "max")
    if anchor is None:
        print(f"  - Skipped empirical extraction: no 'max' anchor in '{sheet.name}'.")
        return rows

    anchor_row, anchor_col = anchor
    headers = collect_header_candidates(grid, anchor_row)
    cols = infer_empirical_columns(anchor_col, headers)

    temp_col = grid.last_col + 2
    temp_row = anchor_row
    temp_cell = sheet.cells(temp_row, temp_col)

    for n_quarters in range(1, N_QUARTERS + 1):
        table_row = anchor_row + n_quarters
        num_quarters_used = get_numeric_cell(sheet, table_row, cols["num_quarters_used"])
        if num_quarters_used is None:
            num_quarters_used = float(n_quarters)

        avg_penetration_pct = get_numeric_cell(sheet, table_row, cols["avg_penetration_pct"])
        source_avg_col = cols["sales_captured_in_db_pct"] or cols["avg_penetration_pct"]
        if source_avg_col is not None:
            start_row = anchor_row + 1
            end_row = table_row
            formula = f"=AVERAGE(R{start_row}C{int(source_avg_col)}:R{end_row}C{int(source_avg_col)})"
            set_r1c1_formula2(temp_cell, formula)
            workbook.app.calculate()
            formula_avg = coerce_number(temp_cell.value)
            if formula_avg is not None:
                avg_penetration_pct = formula_avg

        forecast_value = get_numeric_cell(sheet, table_row, cols["forecast_value"])
        forecast_max = get_numeric_cell(sheet, table_row, cols["forecast_max"])
        forecast_min = get_numeric_cell(sheet, table_row, cols["forecast_min"])
        reported_sales = get_numeric_cell(sheet, table_row, cols["reported_sales"])
        quarterly_sales = get_numeric_cell(sheet, table_row, cols["quarterly_sales"])
        growth_rate_pct = get_numeric_cell(sheet, table_row, cols["growth_rate_pct"])
        sales_captured_in_db_pct = get_numeric_cell(sheet, table_row, cols["sales_captured_in_db_pct"])
        last_quarter_used = get_cell_value(sheet, table_row, cols["last_quarter_used"])

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                reported_sales,
                avg_penetration_pct,
            )
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": int(num_quarters_used),
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
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

    temp_cell.value = None
    return rows


def collect_xy_points(sheet: xw.Sheet, x_col: int, y_col: int, grid: UsedGrid, anchor_row: int) -> list[tuple[int, float, float]]:
    before_anchor: list[tuple[int, float, float]] = []
    after_anchor: list[tuple[int, float, float]] = []

    for row in range(grid.first_row, grid.last_row + 1):
        x_value = get_numeric_cell(sheet, row, x_col)
        y_value = get_numeric_cell(sheet, row, y_col)
        if x_value is None or y_value is None:
            continue
        point = (row, x_value, y_value)
        if row < anchor_row:
            before_anchor.append(point)
        elif row > anchor_row:
            after_anchor.append(point)

    if len(before_anchor) >= 2:
        return before_anchor
    if len(after_anchor) >= 2:
        return after_anchor

    return sorted(before_anchor + after_anchor, key=lambda item: item[0])


def row_signature(values: tuple[Any, ...]) -> tuple[Any, ...]:
    signature: list[Any] = []
    for value in values:
        if isinstance(value, float):
            signature.append(round(value, 8))
        else:
            signature.append(value)
    return tuple(signature)


def extract_regression_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    source_file: str,
    metadata: dict[str, str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    grid = get_used_grid(sheet)
    anchor = find_anchor(grid, "max")
    if anchor is None:
        print(f"  - Skipped regression extraction: no 'max' anchor in '{sheet.name}'.")
        return rows

    anchor_row, anchor_col = anchor
    headers = collect_header_candidates(grid, anchor_row)
    cols = infer_regression_columns(anchor_col, headers)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    points = collect_xy_points(sheet, x_col, y_col, grid, anchor_row)

    if len(points) < 2:
        print(f"  - Skipped regression extraction: not enough x/y points in '{sheet.name}'.")
        return rows

    temp_intercept = sheet.cells(anchor_row, grid.last_col + 2)
    temp_slope = sheet.cells(anchor_row, grid.last_col + 3)

    max_rows = min(N_QUARTERS, len(points))
    prev_signature: tuple[Any, ...] | None = None

    for n_quarters in range(2, max_rows + 1):
        subset = points[-n_quarters:]
        start_row = subset[0][0]
        end_row = subset[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_r1c1_formula2(temp_intercept, intercept_formula)
        set_r1c1_formula2(temp_slope, slope_formula)
        workbook.app.calculate()

        intercept_value = coerce_number(temp_intercept.value)
        slope_value = coerce_number(temp_slope.value)

        table_row = anchor_row + n_quarters
        num_quarters_used = get_numeric_cell(sheet, table_row, cols["num_quarters_used"])
        if num_quarters_used is None:
            num_quarters_used = float(n_quarters)

        forecast_value = get_numeric_cell(sheet, table_row, cols["forecast_value"])
        if forecast_value is None and intercept_value is not None and slope_value is not None:
            forecast_x = get_numeric_cell(sheet, table_row, x_col)
            if forecast_x is None:
                forecast_x = subset[-1][1]
            forecast_value = intercept_value + slope_value * forecast_x

        forecast_max = get_numeric_cell(sheet, table_row, cols["forecast_max"])
        forecast_min = get_numeric_cell(sheet, table_row, cols["forecast_min"])
        actual_value = get_numeric_cell(sheet, table_row, cols["actual_value"])
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = row_signature(
            (
                int(num_quarters_used),
                forecast_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        )
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": int(num_quarters_used),
                "num_quarters_used": int(num_quarters_used),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    temp_intercept.value = None
    temp_slope.value = None
    return rows


def style_sheet(worksheet) -> None:
    header_font = Font(bold=True)
    for col in range(1, worksheet.max_column + 1):
        worksheet.cell(row=1, column=col).font = header_font
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col in range(1, worksheet.max_column + 1):
        max_len = len(str(worksheet.cell(row=1, column=col).value or ""))
        for row in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row, column=col).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    target_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(col) for col in EMPIRICAL_COLUMNS])
    style_sheet(empirical_sheet)

    regression_sheet = workbook.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(col) for col in REGRESSION_COLUMNS])
    style_sheet(regression_sheet)

    workbook.save(target_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    source_files, skipped = collect_input_files(input_path)
    for filename, reason in skipped:
        print(f"Skipped: {filename} ({reason})")

    if not source_files:
        raise RuntimeError("No .xlsx source files found in input_dir.")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files:
            print(f"Processing: {file_path.name}")
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path)

                if EMPIRICAL_SHEET_NAME in [sheet.name for sheet in workbook.sheets]:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            workbook=workbook,
                            sheet=workbook.sheets[EMPIRICAL_SHEET_NAME],
                            source_file=file_path.name,
                            metadata=metadata,
                        )
                    )
                else:
                    print(f"  - Skipped empirical extraction: missing '{EMPIRICAL_SHEET_NAME}' sheet.")

                if REGRESSION_SHEET_NAME in [sheet.name for sheet in workbook.sheets]:
                    regression_rows.extend(
                        extract_regression_rows(
                            workbook=workbook,
                            sheet=workbook.sheets[REGRESSION_SHEET_NAME],
                            source_file=file_path.name,
                            metadata=metadata,
                        )
                    )
                else:
                    print(f"  - Skipped regression extraction: missing '{REGRESSION_SHEET_NAME}' sheet.")

                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            app.quit()

    final_output_path = build_output_path(input_path, output_path)
    write_output_workbook(
        target_path=final_output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {final_output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
