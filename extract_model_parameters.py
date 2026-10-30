#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --------------------------
# Configure these two paths.
# --------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

N_QUARTERS_EMPIRICAL = 10
N_ROWS_REGRESSION = 10
HEADER_SCAN_LEFT = 30
HEADER_SCAN_RIGHT = 12

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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}
FILE_META_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>Early|Mid|Late)\s*(?P<month>[A-Za-z]+)\s*(?P<year>\d{4})",
    re.IGNORECASE,
)


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_token(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", "", text)
    text = text.replace("_", "").replace("-", "").replace("/", "")
    text = re.sub(r"[^a-z0-9%]", "", text)
    return text


def to_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def to_number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    as_num = to_number(value)
    if as_num is None:
        return None
    try:
        return int(as_num)
    except (TypeError, ValueError):
        return None


def safe_col(col: int) -> int:
    return 1 if col < 1 else col


def read_single_col(sheet: xw.Sheet, col: int, start_row: int, row_count: int) -> list[Any]:
    end_row = start_row + row_count - 1
    values = sheet.range((start_row, safe_col(col)), (end_row, safe_col(col))).value
    matrix = to_matrix(values)
    return [row[0] if row else None for row in matrix]


def find_anchor(sheet: xw.Sheet, needle: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    matrix = to_matrix(used.value)
    if not matrix:
        return None

    needle_norm = normalize_token(needle)
    start_row, start_col = used.row, used.column

    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_token(value) == needle_norm:
                return (start_row + r_idx, start_col + c_idx)
    return None


def build_header_map(
    sheet: xw.Sheet, header_row: int, start_col: int, end_col: int
) -> dict[str, int]:
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(values, list):
        values = [values]

    header_map: dict[str, int] = {}
    for idx, value in enumerate(values):
        key = normalize_token(value)
        if key and key not in header_map:
            header_map[key] = start_col + idx
    return header_map


def find_col_by_patterns(
    header_map: dict[str, int], patterns: list[re.Pattern[str]]
) -> int | None:
    for key, col in header_map.items():
        if any(pat.search(key) for pat in patterns):
            return col
    return None


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem
    match = FILE_META_PATTERN.search(stem)
    if match:
        ticker = match.group("ticker").upper()
        period = match.group("period").capitalize()
        month_text = match.group("month")
        year_text = match.group("year")

        month_key = month_text[:3].lower()
        month_num = MONTH_MAP.get(month_key)
        if month_num is not None:
            month_name = month_key.capitalize()
            day = PERIOD_DAY[period.lower()]
            model_period = f"{period}{month_name}_{year_text}"
            model_date = date(int(year_text), month_num, day).isoformat()
            model = f"{ticker}_{model_period}"
            return FileMeta(
                model=model,
                ticker=ticker,
                model_period=model_period,
                model_date=model_date,
            )

    parts = [part.strip() for part in stem.split("-")]
    ticker_fallback = parts[1].upper() if len(parts) > 1 and parts[1] else "UNKNOWN"
    model_period_fallback = "unknown_period"
    model_fallback = f"{ticker_fallback}_{model_period_fallback}"
    return FileMeta(
        model=model_fallback,
        ticker=ticker_fallback,
        model_period=model_period_fallback,
        model_date="",
    )


def output_path_for_input_folder(input_folder: Path, out_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM.xlsx"
    candidate = out_folder / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_folder / f"{input_folder.name}_PARAM.{index}.xlsx"
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

    try:
        workbook.api.Close(SaveChanges=False)
    except TypeError:
        workbook.api.Close(False)


def find_history_start_row(sheet: xw.Sheet, x_col: int, y_col: int, end_row: int) -> int:
    max_lookback = 120
    start_row = end_row
    row = end_row

    while row >= 1 and end_row - row < max_lookback:
        x_val = sheet.range((row, safe_col(x_col))).value
        y_val = sheet.range((row, safe_col(y_col))).value
        if to_number(x_val) is None or to_number(y_val) is None:
            break
        start_row = row
        row -= 1

    return start_row


def process_empirical_sheet(sheet: xw.Sheet, workbook: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(
        sheet,
        anchor_row,
        safe_col(anchor_col - HEADER_SCAN_LEFT),
        safe_col(anchor_col + HEADER_SCAN_RIGHT),
    )

    num_quarters_col = find_col_by_patterns(
        header_map,
        [re.compile(r"numquarters"), re.compile(r"quartersused"), re.compile(r"nquarters")],
    ) or safe_col(anchor_col - 6)
    last_quarter_col = find_col_by_patterns(
        header_map,
        [re.compile(r"lastquarter"), re.compile(r"quarterused")],
    ) or safe_col(anchor_col - 5)
    forecast_value_col = find_col_by_patterns(
        header_map,
        [
            re.compile(r"estimatedtotalsold"),
            re.compile(r"totfcst"),
            re.compile(r"forecastvalue"),
            re.compile(r"totalforecast"),
        ],
    ) or safe_col(anchor_col - 1)
    reported_sales_col = find_col_by_patterns(
        header_map,
        [re.compile(r"reportedsales"), re.compile(r"actualsales"), re.compile(r"actualvalue")],
    ) or safe_col(anchor_col - 2)
    forecast_min_col = find_col_by_patterns(
        header_map,
        [re.compile(r"^min$"), re.compile(r"forecastmin"), re.compile(r"minforecast")],
    ) or safe_col(anchor_col + 1)
    quarterly_sales_col = find_col_by_patterns(
        header_map,
        [re.compile(r"quarterlysales"), re.compile(r"salesquarterly"), re.compile(r"qtrsls")],
    ) or safe_col(anchor_col - 3)
    growth_rate_col = find_col_by_patterns(
        header_map,
        [re.compile(r"growthrate"), re.compile(r"growth%"), re.compile(r"growthpct")],
    )
    sales_captured_col = find_col_by_patterns(
        header_map,
        [re.compile(r"salescapturedindb"), re.compile(r"capturedindb"), re.compile(r"captured")],
    )
    avg_penetration_source_col = find_col_by_patterns(
        header_map,
        [
            re.compile(r"penetrationpct"),
            re.compile(r"penetration"),
            re.compile(r"avgpenetration"),
            re.compile(r"averagepenetration"),
        ],
    ) or safe_col(anchor_col - 4)

    first_data_row = anchor_row + 1
    row_count = N_QUARTERS_EMPIRICAL

    num_quarters_values = read_single_col(sheet, num_quarters_col, first_data_row, row_count)
    helper_col = sheet.used_range.last_cell.column + 2

    avg_formulas = []
    for idx in range(row_count):
        row_num = first_data_row + idx
        n_quarters = to_int(num_quarters_values[idx]) or (idx + 1)
        start_row = max(first_data_row, row_num - n_quarters + 1)
        avg_formulas.append(
            [
                f'=IFERROR(AVERAGE(R{start_row}C{safe_col(avg_penetration_source_col)}:R{row_num}C{safe_col(avg_penetration_source_col)}),"")'
            ]
        )

    sheet.range((first_data_row, helper_col), (first_data_row + row_count - 1, helper_col)).formula2 = avg_formulas
    workbook.app.calculate()

    last_quarter_values = read_single_col(sheet, last_quarter_col, first_data_row, row_count)
    forecast_values = read_single_col(sheet, forecast_value_col, first_data_row, row_count)
    reported_sales_values = read_single_col(sheet, reported_sales_col, first_data_row, row_count)
    forecast_max_values = read_single_col(sheet, anchor_col, first_data_row, row_count)
    forecast_min_values = read_single_col(sheet, forecast_min_col, first_data_row, row_count)
    quarterly_sales_values = read_single_col(sheet, quarterly_sales_col, first_data_row, row_count)
    avg_pen_values = read_single_col(sheet, helper_col, first_data_row, row_count)
    growth_rate_values = (
        read_single_col(sheet, growth_rate_col, first_data_row, row_count)
        if growth_rate_col
        else [None] * row_count
    )
    sales_captured_values = (
        read_single_col(sheet, sales_captured_col, first_data_row, row_count)
        if sales_captured_col
        else [None] * row_count
    )

    output_rows: list[dict[str, Any]] = []
    for idx in range(row_count):
        forecast_value = to_number(forecast_values[idx])
        forecast_max = to_number(forecast_max_values[idx])
        forecast_min = to_number(forecast_min_values[idx])
        reported_sales = to_number(reported_sales_values[idx])
        quarterly_sales = to_number(quarterly_sales_values[idx])
        avg_penetration_pct = to_number(avg_pen_values[idx])
        growth_rate_pct = to_number(growth_rate_values[idx])
        captured_pct = to_number(sales_captured_values[idx])

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and reported_sales is None
            and avg_penetration_pct is None
        ):
            continue

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        num_quarters = to_int(num_quarters_values[idx]) or (idx + 1)
        output_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_values[idx],
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return output_rows


def process_regression_sheet(sheet: xw.Sheet, workbook: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped regression in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(
        sheet,
        anchor_row,
        safe_col(anchor_col - HEADER_SCAN_LEFT),
        safe_col(anchor_col + HEADER_SCAN_RIGHT),
    )

    num_quarters_col = find_col_by_patterns(
        header_map,
        [re.compile(r"numquarters"), re.compile(r"quartersused"), re.compile(r"nquarters")],
    ) or safe_col(anchor_col - 6)
    forecast_total_col = find_col_by_patterns(
        header_map,
        [
            re.compile(r"totfcstwosa"),
            re.compile(r"totforecastwosa"),
            re.compile(r"forecasttotalwosa"),
            re.compile(r"totfcstwithoutsa"),
        ],
    ) or safe_col(anchor_col - 1)
    forecast_min_col = find_col_by_patterns(
        header_map,
        [re.compile(r"^min$"), re.compile(r"forecastmin"), re.compile(r"minforecast")],
    ) or safe_col(anchor_col + 1)
    actual_value_col = find_col_by_patterns(
        header_map,
        [re.compile(r"actual"), re.compile(r"reportedsales")],
    )

    y_col = safe_col(anchor_col - 7)
    x_col = safe_col(anchor_col - 11)

    first_data_row = anchor_row + 1
    row_count = N_ROWS_REGRESSION
    num_quarters_values = read_single_col(sheet, num_quarters_col, first_data_row, row_count)

    history_end_row = anchor_row - 1
    history_start_row = find_history_start_row(sheet, x_col, y_col, history_end_row)

    helper_col = sheet.used_range.last_cell.column + 3
    intercept_formulas: list[list[str]] = []
    slope_formulas: list[list[str]] = []

    for idx in range(row_count):
        n_quarters = to_int(num_quarters_values[idx]) or (idx + 1)
        range_start = max(history_start_row, history_end_row - n_quarters + 1)

        intercept_formulas.append(
            [
                f'=IFERROR(INTERCEPT(R{range_start}C{y_col}:R{history_end_row}C{y_col},R{range_start}C{x_col}:R{history_end_row}C{x_col}),"")'
            ]
        )
        slope_formulas.append(
            [
                f'=IFERROR(SLOPE(R{range_start}C{y_col}:R{history_end_row}C{y_col},R{range_start}C{x_col}:R{history_end_row}C{x_col}),"")'
            ]
        )

    sheet.range((first_data_row, helper_col), (first_data_row + row_count - 1, helper_col)).formula2 = intercept_formulas
    sheet.range((first_data_row, helper_col + 1), (first_data_row + row_count - 1, helper_col + 1)).formula2 = slope_formulas
    workbook.app.calculate()

    intercept_values = read_single_col(sheet, helper_col, first_data_row, row_count)
    slope_values = read_single_col(sheet, helper_col + 1, first_data_row, row_count)
    forecast_values = read_single_col(sheet, forecast_total_col, first_data_row, row_count)
    forecast_max_values = read_single_col(sheet, anchor_col, first_data_row, row_count)
    forecast_min_values = read_single_col(sheet, forecast_min_col, first_data_row, row_count)
    actual_values = (
        read_single_col(sheet, actual_value_col, first_data_row, row_count)
        if actual_value_col
        else [None] * row_count
    )

    output_rows: list[dict[str, Any]] = []
    last_signature: tuple[Any, ...] | None = None

    for idx in range(row_count):
        num_quarters = to_int(num_quarters_values[idx]) or (idx + 1)
        forecast_value = to_number(forecast_values[idx])
        forecast_max = to_number(forecast_max_values[idx])
        forecast_min = to_number(forecast_min_values[idx])
        intercept = to_number(intercept_values[idx])
        slope = to_number(slope_values[idx])
        actual_value = to_number(actual_values[idx])

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        signature = (num_quarters, forecast_value, forecast_max, forecast_min, intercept, slope)
        if signature == last_signature:
            continue
        last_signature = signature

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        output_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
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

    return output_rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    if workbook.sheetnames and workbook.active.title == "Sheet" and workbook.active.max_row == 1:
        sheet = workbook.active
        sheet.title = sheet_name
    else:
        sheet = workbook.create_sheet(title=sheet_name)

    sheet.append(columns)
    for record in rows:
        sheet.append([record.get(column) for column in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    output_book = Workbook()
    write_sheet(output_book, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_book, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_book.save(output_path)


def run() -> int:
    if not input_dir.exists():
        print(f"Input folder not found: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_path_for_input_folder(input_dir, output_dir)

    source_files: list[Path] = []
    for file_path in sorted(input_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        source_files.append(file_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"Processing: {file_path.name}")
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path)

                try:
                    empirical_sheet = workbook.sheets["Empirical Model"]
                    empirical_rows.extend(
                        process_empirical_sheet(empirical_sheet, workbook, meta, file_path.name)
                    )
                except Exception as exc:
                    print(f"Skipped empirical in {file_path.name}: {exc}")

                try:
                    regression_sheet = workbook.sheets["Regression Model"]
                    regression_rows.extend(
                        process_regression_sheet(regression_sheet, workbook, meta, file_path.name)
                    )
                except Exception as exc:
                    print(f"Skipped regression in {file_path.name}: {exc}")

                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
