from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import xlwings as xw


# =========================
# User-configurable paths
# =========================
input_dir = Path("./input")
output_dir = Path("./output")


# =========================
# Constants
# =========================
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10

# Anchor-based offsets (relative to the cell containing "max")
# These defaults match the expected model layout.
EMPIRICAL_ROW_START_OFFSET = 1
EMPIRICAL_COL_OFFSETS = {
    "quarterly_sales": -10,
    "reported_sales": -9,
    "growth_rate_pct": -8,
    "sales_captured_in_db_pct": -7,
    "last_quarter_used": -6,
    "forecast_value": -2,  # estimated total sold
    "actual_value": -3,  # reported sales
    "forecast_min": -1,
    "forecast_max": 0,
}
EMPIRICAL_TEMP_AVG_COL_OFFSET = 2

REGRESSION_ROW_START_OFFSET = 1
REGRESSION_COL_OFFSETS = {
    "num_quarters_used": -10,
    "forecast_total_without_sa": -2,  # TOT FCST w/o SA
    "actual_value": -3,  # optional
    "forecast_min": -1,
    "forecast_max": 0,
}
REGRESSION_INTERCEPT_TEMP_COL_OFFSET = 2
REGRESSION_SLOPE_TEMP_COL_OFFSET = 3

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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

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


@dataclass(frozen=True)
class FileMeta:
    ticker: str
    model_period: str
    model_date: str
    model: str


def parse_file_meta(file_name: str) -> FileMeta:
    stem = Path(file_name).stem

    ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-\s*", stem)
    if not ticker_match:
        raise ValueError("could not parse ticker from file name")
    ticker = ticker_match.group(1).upper()

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if not period_match:
        raise ValueError("could not parse model period from file name")

    period_raw = period_match.group(1)
    month_raw = period_match.group(2)
    year_raw = period_match.group(3)

    period_norm = period_raw.capitalize()
    month_norm = month_raw.capitalize()
    year_int = int(year_raw)

    day = PERIOD_DAY_MAP[period_raw.lower()]
    month_int = MONTH_MAP[month_raw.lower()]

    model_period = f"{period_norm}{month_norm}_{year_int}"
    model_date = date(year_int, month_int, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileMeta(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    fv = to_float(value)
    if fv is None:
        return None
    return int(round(fv))


def safe_subtract(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left - right


def normalize_for_dupe(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 10)
    return value


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.resolve().name
    out_dir.mkdir(parents=True, exist_ok=True)
    base_path = out_dir / f"{folder_name}_PARAM.xlsx"
    if not base_path.exists():
        return base_path

    index = 1
    while True:
        candidate = out_dir / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    name_lookup = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if str(sheet.name).strip().lower() == name_lookup:
            return sheet
    return None


def as_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if len(values) == 0:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def find_anchor_max(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return None

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return used.row + r_idx, used.column + c_idx
    return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        # Primary path requested in the requirements.
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    # Last-resort fallback for older Excel APIs.
    cell.api.FormulaR1C1 = formula_r1c1


def cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def close_workbook_no_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for fallback in (
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ):
        try:
            fallback()
            return
        except Exception:
            continue

    raise RuntimeError(f"unable to close workbook safely: {wb.name}")


def extract_empirical_rows(
    wb: xw.Book, sheet: xw.Sheet, meta: FileMeta, source_file: str
) -> list[dict[str, Any]]:
    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"  skipped {source_file}: no 'max' anchor in {EMPIRICAL_SHEET_NAME}")
        return []

    anchor_row, anchor_col = anchor
    rows: list[dict[str, Any]] = []
    temp_cells: list[xw.Range] = []

    sales_col = anchor_col + EMPIRICAL_COL_OFFSETS["sales_captured_in_db_pct"]

    for idx in range(N_QUARTERS):
        num_quarters = idx + 1
        row_idx = anchor_row + EMPIRICAL_ROW_START_OFFSET + idx
        temp_cell = sheet.range((row_idx, anchor_col + EMPIRICAL_TEMP_AVG_COL_OFFSET))
        hist_start = row_idx - num_quarters + 1

        formula_r1c1 = f"=AVERAGE(R{hist_start}C{sales_col}:R{row_idx}C{sales_col})"
        set_formula2_r1c1(temp_cell, formula_r1c1)
        temp_cells.append(temp_cell)

    if temp_cells:
        wb.app.calculate()

    for idx in range(N_QUARTERS):
        num_quarters = idx + 1
        row_idx = anchor_row + EMPIRICAL_ROW_START_OFFSET + idx

        forecast_value = to_float(
            cell_value(sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["forecast_value"])
        )
        actual_value = to_float(
            cell_value(sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["actual_value"])
        )
        forecast_max = to_float(
            cell_value(sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["forecast_max"])
        )
        forecast_min = to_float(
            cell_value(sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["forecast_min"])
        )
        avg_penetration_pct = to_float(temp_cells[idx].value)
        quarterly_sales = to_float(
            cell_value(sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["quarterly_sales"])
        )
        reported_sales = to_float(
            cell_value(sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["reported_sales"])
        )
        growth_rate_pct = to_float(
            cell_value(sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["growth_rate_pct"])
        )
        sales_captured_in_db_pct = to_float(
            cell_value(
                sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["sales_captured_in_db_pct"]
            )
        )
        last_quarter_used = cell_value(
            sheet, row_idx, anchor_col + EMPIRICAL_COL_OFFSETS["last_quarter_used"]
        )

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    # Optional cleanup: formulas are temporary and source workbooks are never saved.
    for temp_cell in temp_cells:
        temp_cell.value = None

    return rows


def extract_regression_rows(
    wb: xw.Book, sheet: xw.Sheet, meta: FileMeta, source_file: str
) -> list[dict[str, Any]]:
    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"  skipped {source_file}: no 'max' anchor in {REGRESSION_SHEET_NAME}")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    rows: list[dict[str, Any]] = []
    temp_intercepts: list[xw.Range] = []
    temp_slopes: list[xw.Range] = []

    for idx in range(N_QUARTERS):
        num_quarters = idx + 1
        row_idx = anchor_row + REGRESSION_ROW_START_OFFSET + idx
        hist_end = row_idx - 1
        hist_start = hist_end - num_quarters + 1

        intercept_cell = sheet.range((row_idx, anchor_col + REGRESSION_INTERCEPT_TEMP_COL_OFFSET))
        slope_cell = sheet.range((row_idx, anchor_col + REGRESSION_SLOPE_TEMP_COL_OFFSET))

        intercept_formula = (
            f"=INTERCEPT(R{hist_start}C{y_col}:R{hist_end}C{y_col},"
            f"R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{hist_start}C{y_col}:R{hist_end}C{y_col},"
            f"R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        temp_intercepts.append(intercept_cell)
        temp_slopes.append(slope_cell)

    if temp_intercepts or temp_slopes:
        wb.app.calculate()

    previous_key: tuple[Any, ...] | None = None

    for idx in range(N_QUARTERS):
        fallback_num_quarters = idx + 1
        row_idx = anchor_row + REGRESSION_ROW_START_OFFSET + idx

        num_quarters_used = to_int(
            cell_value(sheet, row_idx, anchor_col + REGRESSION_COL_OFFSETS["num_quarters_used"])
        )
        if num_quarters_used is None:
            num_quarters_used = fallback_num_quarters

        intercept = to_float(temp_intercepts[idx].value)
        slope = to_float(temp_slopes[idx].value)

        forecast_value = to_float(
            cell_value(
                sheet,
                row_idx,
                anchor_col + REGRESSION_COL_OFFSETS["forecast_total_without_sa"],
            )
        )
        actual_value = to_float(
            cell_value(sheet, row_idx, anchor_col + REGRESSION_COL_OFFSETS["actual_value"])
        )
        forecast_max = to_float(
            cell_value(sheet, row_idx, anchor_col + REGRESSION_COL_OFFSETS["forecast_max"])
        )
        forecast_min = to_float(
            cell_value(sheet, row_idx, anchor_col + REGRESSION_COL_OFFSETS["forecast_min"])
        )

        key = tuple(
            normalize_for_dupe(v)
            for v in (
                num_quarters_used,
                intercept,
                slope,
                forecast_value,
                forecast_max,
                forecast_min,
            )
        )
        if previous_key is not None and key == previous_key:
            continue
        previous_key = key

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    # Optional cleanup: formulas are temporary and source workbooks are never saved.
    for temp_cell in temp_intercepts + temp_slopes:
        temp_cell.value = None

    return rows


def write_sheet(
    app: xw.App,
    sheet: xw.Sheet,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    sheet.clear_contents()
    sheet.range((1, 1)).value = columns

    if rows:
        matrix = [[row.get(col, "") for col in columns] for row in rows]
        sheet.range((2, 1)).value = matrix

    max_row = max(1, len(rows) + 1)
    max_col = len(columns)

    header_range = sheet.range((1, 1), (1, max_col))
    header_range.api.Font.Bold = True

    table_range = sheet.range((1, 1), (max_row, max_col))
    table_range.api.AutoFilter()

    for col_idx, col_name in enumerate(columns, start=1):
        width = min(max(len(col_name) + 2, 12), 42)
        sheet.range((1, col_idx)).column_width = width
    sheet.autofit("c")

    try:
        sheet.activate()
        app.api.ActiveWindow.SplitRow = 1
        app.api.ActiveWindow.SplitColumn = 0
        app.api.ActiveWindow.FreezePanes = True
    except Exception:
        # Freeze panes depends on an active window; skip if unavailable.
        pass


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = app.books.add()
    try:
        empirical_sheet = out_wb.sheets[0]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet = out_wb.sheets.add("regression_candidates", after=empirical_sheet)

        # Keep only the two required output tabs.
        while len(out_wb.sheets) > 2:
            out_wb.sheets[len(out_wb.sheets) - 1].delete()

        write_sheet(app, empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
        write_sheet(app, regression_sheet, REGRESSION_COLUMNS, regression_rows)

        out_wb.save(str(output_path))
    finally:
        close_workbook_no_save(out_wb)


def process_file(app: xw.App, file_path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    meta = parse_file_meta(file_path.name)
    wb = app.books.open(str(file_path), update_links=False)
    try:
        empirical_rows: list[dict[str, Any]] = []
        regression_rows: list[dict[str, Any]] = []

        empirical_sheet = get_sheet_case_insensitive(wb, EMPIRICAL_SHEET_NAME)
        if empirical_sheet is not None:
            empirical_rows = extract_empirical_rows(wb, empirical_sheet, meta, file_path.name)
        else:
            print(f"  skipped {file_path.name}: missing sheet '{EMPIRICAL_SHEET_NAME}'")

        regression_sheet = get_sheet_case_insensitive(wb, REGRESSION_SHEET_NAME)
        if regression_sheet is not None:
            regression_rows = extract_regression_rows(wb, regression_sheet, meta, file_path.name)
        else:
            print(f"  skipped {file_path.name}: missing sheet '{REGRESSION_SHEET_NAME}'")

        return empirical_rows, regression_rows
    finally:
        close_workbook_no_save(wb)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    if not input_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a folder: {input_dir}")

    output_path = next_output_path(input_dir, output_dir)

    all_empirical_rows: list[dict[str, Any]] = []
    all_regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    files = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in files:
            if not file_path.is_file():
                continue

            file_name = file_path.name
            file_name_lower = file_name.lower()

            if file_name.startswith("~"):
                print(f"Skipped: {file_name} (temporary file)")
                continue
            if not file_name_lower.endswith(".xlsx"):
                print(f"Skipped: {file_name} (not .xlsx)")
                continue

            try:
                empirical_rows, regression_rows = process_file(app, file_path)
            except Exception as exc:
                print(f"Skipped: {file_name} ({exc})")
                continue

            all_empirical_rows.extend(empirical_rows)
            all_regression_rows.extend(regression_rows)
            processed_count += 1
            print(f"Processed file: {file_name}")

        write_output_workbook(app, output_path, all_empirical_rows, all_regression_rows)
    finally:
        app.quit()

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(all_empirical_rows)}")
    print(f"Number of regression rows: {len(all_regression_rows)}")


if __name__ == "__main__":
    main()
