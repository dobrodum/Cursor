#!/usr/bin/env python3
"""Extract empirical/regression parameter candidates from Excel model workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------
# User-configurable paths
# ---------------------------
input_dir = Path("./input")
output_dir = Path("./output")

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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}

FILENAME_PERIOD_RE = re.compile(
    r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class FileModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except ValueError:
        return None


def as_int(value: Any) -> int | None:
    parsed = as_float(value)
    if parsed is None:
        return None
    return int(round(parsed))


def parse_metadata_from_filename(file_path: Path) -> FileModelMetadata | None:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    if len(parts) < 3:
        return None

    ticker = parts[1].upper()
    period_match = FILENAME_PERIOD_RE.search(parts[2])
    if not period_match:
        return None

    period_bucket, month_token, year_token = period_match.groups()
    period_key = period_bucket.lower()
    month_key = month_token.lower()
    year_num = int(year_token)

    model_period = f"{period_bucket.title()}{month_token.title()}_{year_num}"
    model_date = date(year_num, MONTH_MAP[month_key], DAY_MAP[period_key]).isoformat()
    model = f"{ticker}_{model_period}"
    return FileModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    matrix = to_2d(used.value)
    if not matrix:
        return None

    top_row = used.row
    left_col = used.column
    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return top_row + row_idx, left_col + col_idx
    return None


def read_anchor_header_row(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    scan_width: int = 30,
) -> tuple[int, list[Any]]:
    left = max(1, anchor_col - scan_width)
    right = anchor_col + scan_width
    header_values = to_2d(sheet.range((anchor_row, left), (anchor_row, right)).value)
    if not header_values:
        return left, []
    return left, header_values[0]


def find_header_column(
    header_left: int,
    header_row: list[Any],
    required_tokens: Iterable[str],
    default_col: int,
) -> int:
    tokens = [token.lower() for token in required_tokens]
    for idx, raw_value in enumerate(header_row):
        text = normalize_text(raw_value)
        if text and all(token in text for token in tokens):
            return header_left + idx
    return default_col


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    file_stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{file_stem}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_path / f"{file_stem}.{counter}.xlsx"
        counter += 1
    return candidate


def normalize_ratio(value: float | None) -> float | None:
    if value is None:
        return None
    if value > 1:
        return value / 100.0
    return value


def clamp_column(col_num: int) -> int:
    return max(1, int(col_num))


def extract_empirical_candidates(
    workbook: xw.Book,
    metadata: FileModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"skipped empirical extraction for {source_file}: missing 'Empirical Model' sheet")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"skipped empirical extraction for {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1
    header_left, header_row = read_anchor_header_row(sheet, anchor_row, anchor_col)

    col_map = {
        "num_quarters_used": anchor_col - 10,
        "last_quarter_used": anchor_col - 11,
        "quarterly_sales": anchor_col - 9,
        "reported_sales": anchor_col - 8,
        "growth_rate_pct": anchor_col - 7,
        "sales_captured_in_db_pct": anchor_col - 6,
        "penetration": anchor_col - 5,
        "forecast_value": anchor_col - 4,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }

    col_map["num_quarters_used"] = find_header_column(
        header_left,
        header_row,
        ("quarter", "used"),
        col_map["num_quarters_used"],
    )
    col_map["last_quarter_used"] = find_header_column(
        header_left,
        header_row,
        ("last", "quarter"),
        col_map["last_quarter_used"],
    )
    col_map["quarterly_sales"] = find_header_column(
        header_left,
        header_row,
        ("quarterly", "sales"),
        col_map["quarterly_sales"],
    )
    col_map["reported_sales"] = find_header_column(
        header_left,
        header_row,
        ("reported", "sales"),
        col_map["reported_sales"],
    )
    col_map["growth_rate_pct"] = find_header_column(
        header_left,
        header_row,
        ("growth", "rate"),
        col_map["growth_rate_pct"],
    )
    col_map["sales_captured_in_db_pct"] = find_header_column(
        header_left,
        header_row,
        ("captured", "db"),
        col_map["sales_captured_in_db_pct"],
    )
    col_map["penetration"] = find_header_column(
        header_left,
        header_row,
        ("penetration",),
        col_map["penetration"],
    )
    col_map["forecast_value"] = find_header_column(
        header_left,
        header_row,
        ("estimated", "sold"),
        col_map["forecast_value"],
    )
    col_map["forecast_max"] = find_header_column(
        header_left,
        header_row,
        ("max",),
        col_map["forecast_max"],
    )
    col_map["forecast_min"] = find_header_column(
        header_left,
        header_row,
        ("min",),
        col_map["forecast_min"],
    )
    col_map = {key: clamp_column(value) for key, value in col_map.items()}

    helper_col = max(col_map.values()) + 3
    helper_range = sheet.range((data_start_row, helper_col), (data_end_row, helper_col))
    for row_idx in range(N_QUARTERS):
        row_num = data_start_row + row_idx
        window_start = max(data_start_row, row_num - row_idx)
        helper_formula = (
            f'=IFERROR(AVERAGE(R{window_start}C{col_map["penetration"]}:'
            f'R{row_num}C{col_map["penetration"]}),"")'
        )
        sheet.cells(row_num, helper_col).formula2 = helper_formula

    workbook.app.calculate()
    avg_pen_values = helper_range.value
    helper_range.value = None

    if not isinstance(avg_pen_values, list):
        avg_pen_values = [avg_pen_values]

    block_left = min(col_map.values())
    block_right = max(col_map.values())
    row_block = to_2d(sheet.range((data_start_row, block_left), (data_end_row, block_right)).value)

    def row_cell(row_idx: int, col_num: int) -> Any:
        return row_block[row_idx][col_num - block_left]

    rows: list[dict[str, Any]] = []
    for idx in range(min(N_QUARTERS, len(row_block))):
        avg_penetration = as_float(avg_pen_values[idx] if idx < len(avg_pen_values) else None)
        num_quarters = as_int(row_cell(idx, col_map["num_quarters_used"])) or (idx + 1)
        last_quarter_used = row_cell(idx, col_map["last_quarter_used"])
        quarterly_sales = as_float(row_cell(idx, col_map["quarterly_sales"]))
        reported_sales = as_float(row_cell(idx, col_map["reported_sales"]))
        growth_rate_pct = as_float(row_cell(idx, col_map["growth_rate_pct"]))
        sales_captured_pct = as_float(row_cell(idx, col_map["sales_captured_in_db_pct"]))
        forecast_value = as_float(row_cell(idx, col_map["forecast_value"]))
        forecast_max = as_float(row_cell(idx, col_map["forecast_max"]))
        forecast_min = as_float(row_cell(idx, col_map["forecast_min"]))

        if forecast_value is None and quarterly_sales is not None and avg_penetration is not None:
            ratio = normalize_ratio(avg_penetration)
            if ratio not in (None, 0):
                forecast_value = quarterly_sales / ratio

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        if (
            avg_penetration is None
            and quarterly_sales is None
            and reported_sales is None
            and forecast_max is None
            and forecast_min is None
            and forecast_value is None
        ):
            continue

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    workbook: xw.Book,
    metadata: FileModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"skipped regression extraction for {source_file}: missing 'Regression Model' sheet")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"skipped regression extraction for {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1
    header_left, header_row = read_anchor_header_row(sheet, anchor_row, anchor_col)

    col_map = {
        "num_quarters_used": anchor_col - 10,
        "x_col": anchor_col - 11,
        "y_col": anchor_col - 7,
        "forecast_value": anchor_col - 2,
        "actual_value": anchor_col - 1,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }

    col_map["num_quarters_used"] = find_header_column(
        header_left,
        header_row,
        ("quarter", "used"),
        col_map["num_quarters_used"],
    )
    col_map["forecast_value"] = find_header_column(
        header_left,
        header_row,
        ("fcst", "w/o", "sa"),
        col_map["forecast_value"],
    )
    col_map["actual_value"] = find_header_column(
        header_left,
        header_row,
        ("actual",),
        col_map["actual_value"],
    )
    col_map["forecast_max"] = find_header_column(
        header_left,
        header_row,
        ("max",),
        col_map["forecast_max"],
    )
    col_map["forecast_min"] = find_header_column(
        header_left,
        header_row,
        ("min",),
        col_map["forecast_min"],
    )
    col_map = {key: clamp_column(value) for key, value in col_map.items()}

    intercept_col = max(col_map.values()) + 3
    slope_col = intercept_col + 1
    intercept_range = sheet.range((data_start_row, intercept_col), (data_end_row, intercept_col))
    slope_range = sheet.range((data_start_row, slope_col), (data_end_row, slope_col))

    for row_idx in range(N_QUARTERS):
        row_num = data_start_row + row_idx
        range_start = max(data_start_row, row_num - row_idx)
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{range_start}C{col_map["y_col"]}:R{row_num}C{col_map["y_col"]},'
            f'R{range_start}C{col_map["x_col"]}:R{row_num}C{col_map["x_col"]}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{range_start}C{col_map["y_col"]}:R{row_num}C{col_map["y_col"]},'
            f'R{range_start}C{col_map["x_col"]}:R{row_num}C{col_map["x_col"]}),"")'
        )
        sheet.cells(row_num, intercept_col).formula2 = intercept_formula
        sheet.cells(row_num, slope_col).formula2 = slope_formula

    workbook.app.calculate()
    intercept_values = intercept_range.value
    slope_values = slope_range.value
    intercept_range.value = None
    slope_range.value = None

    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]

    block_left = min(col_map.values())
    block_right = max(col_map.values())
    row_block = to_2d(sheet.range((data_start_row, block_left), (data_end_row, block_right)).value)

    def row_cell(row_idx: int, col_num: int) -> Any:
        return row_block[row_idx][col_num - block_left]

    rows: list[dict[str, Any]] = []
    previous_key: tuple[Any, ...] | None = None
    for idx in range(min(N_QUARTERS, len(row_block))):
        num_quarters = as_int(row_cell(idx, col_map["num_quarters_used"])) or (idx + 1)
        x_value = as_float(row_cell(idx, col_map["x_col"]))
        forecast_value = as_float(row_cell(idx, col_map["forecast_value"]))
        actual_value = as_float(row_cell(idx, col_map["actual_value"]))
        forecast_max = as_float(row_cell(idx, col_map["forecast_max"]))
        forecast_min = as_float(row_cell(idx, col_map["forecast_min"]))
        intercept = as_float(intercept_values[idx] if idx < len(intercept_values) else None)
        slope = as_float(slope_values[idx] if idx < len(slope_values) else None)

        if forecast_value is None and intercept is not None and slope is not None and x_value is not None:
            forecast_value = intercept + (slope * x_value)

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        current_key = (
            num_quarters,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )
        if previous_key is not None and current_key == previous_key:
            continue
        previous_key = current_key

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
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


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)


def write_sheet(sheet: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])

    for col_idx, header in enumerate(columns, start=1):
        header_cell = sheet.cell(row=1, column=col_idx)
        header_cell.font = Font(bold=True)
        max_len = len(header)
        for row_idx in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = max(12, min(max_len + 2, 44))

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{sheet.max_row}"


def is_xlsx_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".xlsx"


def run() -> None:
    src_dir = Path(input_dir).expanduser().resolve()
    dst_dir = Path(output_dir).expanduser().resolve()

    if not src_dir.exists() or not src_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {src_dir}")

    output_file = build_output_path(src_dir, dst_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    excel_app = xw.App(visible=False, add_book=False)
    excel_app.display_alerts = False
    excel_app.screen_updating = False

    try:
        for file_path in sorted(src_dir.iterdir()):
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue
            if not is_xlsx_file(file_path):
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            metadata = parse_metadata_from_filename(file_path)
            if metadata is None:
                print(f"skipped file: {file_path.name} (filename format mismatch)")
                continue

            try:
                workbook = excel_app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"skipped file: {file_path.name} (open failed: {exc})")
                continue

            try:
                print(f"processed file: {file_path.name}")
                empirical_rows.extend(
                    extract_empirical_candidates(workbook, metadata, file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(workbook, metadata, file_path.name)
                )
                processed_file_count += 1
            finally:
                safe_close_workbook(workbook)
    finally:
        excel_app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_file_count}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
