#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# Configure folders here
# -----------------------------
input_dir = Path("./input").resolve()
output_dir = Path("./output").resolve()


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

DAY_BY_BUCKET = {"early": 5, "mid": 15, "late": 25}
MONTH_TO_NUM = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

FILENAME_PATTERN = re.compile(
    r".*?-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<bucket>Early|Mid|Late)\s*"
    r"(?P<month>[A-Za-z]+)\s*(?P<year>\d{4}).*",
    flags=re.IGNORECASE,
)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def as_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).strip().replace(",", "")
    if not raw:
        return None
    if raw.endswith("%"):
        try:
            return float(raw[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(raw)
    except ValueError:
        return None


def subtract_or_none(left: Any, right: Any) -> float | None:
    left_float = as_float(left)
    right_float = as_float(right)
    if left_float is None or right_float is None:
        return None
    return left_float - right_float


def parse_file_label(file_path: Path) -> dict[str, str] | None:
    match = FILENAME_PATTERN.match(file_path.stem)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    bucket_raw = match.group("bucket")
    month_raw = match.group("month")
    year_raw = match.group("year")

    month_num = MONTH_TO_NUM.get(month_raw.lower())
    if month_num is None:
        return None

    bucket_key = bucket_raw.lower()
    day = DAY_BY_BUCKET[bucket_key]
    month_abbrev = date(2000, month_num, 1).strftime("%b")
    model_period = f"{bucket_raw.title()}{month_abbrev}_{year_raw}"
    model_date = date(int(year_raw), month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def make_output_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    version = 1
    while candidate.exists():
        candidate = output_folder / f"{base_name}.{version}.xlsx"
        version += 1
    return candidate


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
        # Last resort: workbook will be closed when app quits.
        pass


def find_anchor_cell(ws: xw.Sheet, anchor_value: str = "max") -> tuple[int, int] | None:
    used_range = ws.used_range
    values = used_range.value
    if values is None:
        return None

    start_row = used_range.row
    start_col = used_range.column
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    target = normalize_text(anchor_value)
    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                return (start_row + row_idx, start_col + col_idx)
    return None


def get_used_bounds(ws: xw.Sheet) -> tuple[int, int, int, int]:
    used = ws.used_range
    first_row = used.row
    first_col = used.column
    values = used.value
    if values is None:
        return first_row, first_col, first_row, first_col
    if not isinstance(values, list):
        row_count = 1
        col_count = 1
    elif values and not isinstance(values[0], list):
        row_count = 1
        col_count = len(values)
    else:
        row_count = len(values)
        col_count = max((len(r) for r in values), default=1)
    last_row = first_row + row_count - 1
    last_col = first_col + col_count - 1
    return first_row, first_col, last_row, last_col


def build_header_index(
    ws: xw.Sheet, header_row: int, first_col: int, last_col: int
) -> dict[str, int]:
    header_values = ws.range((header_row, first_col), (header_row, last_col)).value
    if header_values is None:
        return {}
    if not isinstance(header_values, list):
        header_values = [header_values]

    index: dict[str, int] = {}
    for idx, value in enumerate(header_values):
        normalized = normalize_text(value)
        if normalized:
            index[normalized] = first_col + idx
    return index


def lookup_column(header_index: dict[str, int], aliases: Iterable[str]) -> int | None:
    for alias in aliases:
        normalized_alias = normalize_text(alias)
        if normalized_alias in header_index:
            return header_index[normalized_alias]
    return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def get_cell(ws: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return ws.range((row, col)).value


def is_meaningful_row(values: list[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return True
    return False


def process_empirical_sheet(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        ws = wb.sheets["Empirical Model"]
    except Exception:
        return []

    anchor = find_anchor_cell(ws, "max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    first_row, first_col, last_row, last_col = get_used_bounds(ws)
    header_index = build_header_index(ws, anchor_row, first_col, last_col)

    # Anchor-based column defaults from the "max" column.
    num_quarters_col = lookup_column(
        header_index, ["num quarters used", "num quarters", "quarters used", "n quarters"]
    ) or (anchor_col - 10)
    last_quarter_col = lookup_column(
        header_index, ["last quarter used", "last quarter", "last qtr"]
    ) or (anchor_col - 9)
    forecast_value_col = lookup_column(
        header_index, ["estimated total sold", "forecast value", "forecast"]
    ) or (anchor_col - 4)
    actual_value_col = lookup_column(
        header_index, ["reported sales", "actual value", "actual"]
    ) or (anchor_col - 3)
    forecast_max_col = lookup_column(header_index, ["max", "forecast max"]) or anchor_col
    forecast_min_col = lookup_column(header_index, ["min", "forecast min"]) or (anchor_col + 1)
    quarterly_sales_col = lookup_column(
        header_index, ["quarterly sales", "quarter sales"]
    ) or (anchor_col - 8)
    reported_sales_col = lookup_column(header_index, ["reported sales"]) or actual_value_col
    growth_rate_col = lookup_column(
        header_index, ["growth rate pct", "growth rate", "growth %", "growth"]
    ) or (anchor_col - 2)
    sales_captured_col = lookup_column(
        header_index,
        [
            "sales captured in db pct",
            "sales captured in db",
            "captured in db pct",
            "penetration",
            "penetration pct",
        ],
    ) or (anchor_col - 1)

    n_quarters = 10
    data_start_row = anchor_row + 1
    data_end_row = min(last_row, data_start_row + n_quarters - 1)
    if data_end_row < data_start_row:
        return []

    temp_col = max(last_col + 2, anchor_col + 2)
    formula_rows = data_end_row - data_start_row + 1

    # Write all formulas first, calculate once, then read all results.
    for i in range(formula_rows):
        row = data_start_row + i
        formula = (
            f"=AVERAGE(R{data_start_row}C{sales_captured_col}:R{row}C{sales_captured_col})"
        )
        set_formula2(ws.range((row, temp_col)), formula)

    wb.app.calculate()
    avg_pen_values = ws.range(
        (data_start_row, temp_col), (data_end_row, temp_col)
    ).value
    if not isinstance(avg_pen_values, list):
        avg_pen_values = [avg_pen_values]

    rows: list[dict[str, Any]] = []
    blank_streak = 0
    for idx in range(formula_rows):
        row = data_start_row + idx
        num_quarters_used = get_cell(ws, row, num_quarters_col)
        last_quarter_used = get_cell(ws, row, last_quarter_col)
        forecast_value = get_cell(ws, row, forecast_value_col)
        actual_value = get_cell(ws, row, actual_value_col)
        forecast_max = get_cell(ws, row, forecast_max_col)
        forecast_min = get_cell(ws, row, forecast_min_col)
        quarterly_sales = get_cell(ws, row, quarterly_sales_col)
        reported_sales = get_cell(ws, row, reported_sales_col)
        growth_rate_pct = get_cell(ws, row, growth_rate_col)
        sales_captured_in_db_pct = get_cell(ws, row, sales_captured_col)
        avg_penetration_pct = avg_pen_values[idx] if idx < len(avg_pen_values) else None

        row_values = [
            num_quarters_used,
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            quarterly_sales,
            reported_sales,
            growth_rate_pct,
            sales_captured_in_db_pct,
            avg_penetration_pct,
        ]
        if not is_meaningful_row(row_values):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0

        if as_float(num_quarters_used) is None:
            num_quarters_used = idx + 1

        if forecast_value is None and as_float(actual_value) is not None and as_float(avg_penetration_pct):
            forecast_value = as_float(actual_value) / as_float(avg_penetration_pct)

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_or_none(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    ws.range((data_start_row, temp_col), (data_end_row, temp_col)).clear_contents()
    return rows


def process_regression_sheet(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        ws = wb.sheets["Regression Model"]
    except Exception:
        return []

    anchor = find_anchor_cell(ws, "max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    first_row, first_col, last_row, last_col = get_used_bounds(ws)
    header_index = build_header_index(ws, anchor_row, first_col, last_col)

    num_quarters_col = lookup_column(
        header_index, ["num quarters used", "num quarters", "quarters used", "n quarters"]
    ) or (anchor_col - 12)
    forecast_value_col = lookup_column(
        header_index, ["tot fcst w/o sa", "total fcst w/o sa", "forecast value", "forecast"]
    ) or (anchor_col - 1)
    actual_value_col = lookup_column(header_index, ["actual value", "actual", "reported sales"])
    forecast_max_col = lookup_column(header_index, ["max", "forecast max"]) or anchor_col
    forecast_min_col = lookup_column(header_index, ["min", "forecast min"]) or (anchor_col + 1)

    # Required anchor-based regression columns.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    n_quarters = 10
    data_start_row = anchor_row + 1
    data_end_row = min(last_row, data_start_row + n_quarters - 1)
    if data_end_row < data_start_row:
        return []

    temp_intercept_col = max(last_col + 2, anchor_col + 3)
    temp_slope_col = temp_intercept_col + 1
    formula_rows = data_end_row - data_start_row + 1

    for i in range(formula_rows):
        row = data_start_row + i
        start_row = data_start_row
        formula_intercept = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},"
            f"R{start_row}C{x_col}:R{row}C{x_col})"
        )
        formula_slope = (
            f"=SLOPE(R{start_row}C{y_col}:R{row}C{y_col},"
            f"R{start_row}C{x_col}:R{row}C{x_col})"
        )
        set_formula2(ws.range((row, temp_intercept_col)), formula_intercept)
        set_formula2(ws.range((row, temp_slope_col)), formula_slope)

    wb.app.calculate()
    intercept_values = ws.range(
        (data_start_row, temp_intercept_col), (data_end_row, temp_intercept_col)
    ).value
    slope_values = ws.range((data_start_row, temp_slope_col), (data_end_row, temp_slope_col)).value
    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    blank_streak = 0

    for idx in range(formula_rows):
        row = data_start_row + idx
        num_quarters_used = get_cell(ws, row, num_quarters_col)
        forecast_value = get_cell(ws, row, forecast_value_col)
        actual_value = get_cell(ws, row, actual_value_col) if actual_value_col else None
        forecast_max = get_cell(ws, row, forecast_max_col)
        forecast_min = get_cell(ws, row, forecast_min_col)
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None

        row_values = [
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        ]
        if not is_meaningful_row(row_values):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0

        if as_float(num_quarters_used) is None:
            num_quarters_used = idx + 1

        signature = (
            as_float(num_quarters_used),
            as_float(forecast_value),
            as_float(forecast_max),
            as_float(forecast_min),
            as_float(intercept),
            as_float(slope),
        )
        # Avoid duplicate terminal row generated by workbook logic.
        if previous_signature is not None and signature == previous_signature:
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
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_or_none(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    ws.range((data_start_row, temp_intercept_col), (data_end_row, temp_slope_col)).clear_contents()
    return rows


def write_sheet(
    ws: Any, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 44)


def should_skip_file(file_path: Path, input_folder_name: str) -> tuple[bool, str]:
    if not file_path.is_file():
        return True, "not a file"
    if file_path.name.startswith("~"):
        return True, "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return True, "not .xlsx"
    lower_name = file_path.name.lower()
    if lower_name.startswith(f"{input_folder_name.lower()}_param"):
        return True, "looks like a generated output workbook"
    return False, ""


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = make_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_dir.iterdir()):
            skip, reason = should_skip_file(file_path, input_dir.name)
            if skip:
                print(f"SKIPPED: {file_path.name} ({reason})")
                continue

            metadata = parse_file_label(file_path)
            if metadata is None:
                print(f"SKIPPED: {file_path.name} (unable to parse ticker/period/date)")
                continue

            print(f"PROCESSING: {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    empirical_sheet = out_wb.create_sheet("empirical_candidates")
    regression_sheet = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)

    print(f"OUTPUT: {output_path}")
    print(f"FILES PROCESSED: {files_processed}")
    print(f"EMPIRICAL ROWS: {len(empirical_rows)}")
    print(f"REGRESSION ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
