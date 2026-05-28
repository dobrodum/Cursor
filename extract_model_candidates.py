#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these directories before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

MONTH_LOOKUP = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[_\s]+", " ", text)
    text = re.sub(r"[^a-z0-9 %./-]+", "", text)
    return text.strip()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def maybe_number(value: Any) -> Any:
    numeric = to_float(value)
    if numeric is None:
        return value
    if abs(numeric - round(numeric)) < 1e-10:
        return int(round(numeric))
    return numeric


def values_equal(left: Any, right: Any, tol: float = 1e-9) -> bool:
    lnum = to_float(left)
    rnum = to_float(right)
    if lnum is not None and rnum is not None:
        return abs(lnum - rnum) <= tol
    return (left or "") == (right or "")


def flatten_vertical(values: Any, expected_len: int) -> List[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flattened = [row[0] if row else None for row in values]
        else:
            flattened = values
    else:
        flattened = [values]

    if len(flattened) < expected_len:
        flattened.extend([None] * (expected_len - len(flattened)))
    return flattened[:expected_len]


def vertical_values(sheet: xw.Sheet, start_row: int, end_row: int, col: Optional[int]) -> List[Any]:
    row_count = max(0, end_row - start_row + 1)
    if row_count == 0:
        return []
    if col is None or col < 1:
        return [None] * row_count
    values = sheet.range((start_row, col), (end_row, col)).value
    return flatten_vertical(values, row_count)


def find_anchor(sheet: xw.Sheet, target_text: str) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if isinstance(values, list):
        if values and isinstance(values[0], list):
            grid = values
        else:
            grid = [values]
    else:
        grid = [[values]]

    target = normalize_text(target_text)
    base_row = used.row
    base_col = used.column

    for r_idx, row in enumerate(grid):
        if not isinstance(row, list):
            row = [row]
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return base_row + r_idx, base_col + c_idx
    return None


def build_header_maps(sheet: xw.Sheet, anchor_row: int, anchor_col: int, span: int = 35) -> List[Dict[str, int]]:
    start_col = max(1, anchor_col - span)
    end_col = anchor_col + span
    maps: List[Dict[str, int]] = []

    for row in (anchor_row - 1, anchor_row, anchor_row + 1):
        if row < 1:
            continue
        row_values = sheet.range((row, start_col), (row, end_col)).value
        if not isinstance(row_values, list):
            row_values = [row_values]

        header_map: Dict[str, int] = {}
        for idx, value in enumerate(row_values):
            text = normalize_text(value)
            if text and text not in header_map:
                header_map[text] = start_col + idx
        maps.append(header_map)

    return maps


def match_column(header_maps: Sequence[Dict[str, int]], keyword_sets: Sequence[Sequence[str]]) -> Optional[int]:
    for keywords in keyword_sets:
        normalized_keywords = [normalize_text(keyword) for keyword in keywords]
        for header_map in header_maps:
            for header, col in header_map.items():
                if all(keyword in header for keyword in normalized_keywords):
                    return col
    return None


def parse_file_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]{1,10})\s*-", stem)
        ticker = ticker_match.group(1).upper() if ticker_match else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )

    if period_match:
        period_bucket = period_match.group(1).capitalize()
        month_token = period_match.group(2).strip()
        year = int(period_match.group(3))
        month_key = month_token[:3].lower()

        if month_key in MONTH_LOOKUP:
            month_number = MONTH_LOOKUP[month_key]
            month_abbrev = calendar.month_abbr[month_number]
            day = PERIOD_DAY[period_bucket.lower()]
            model_period = f"{period_bucket}{month_abbrev}_{year}"
            model_date = date(year, month_number, day).isoformat()
        else:
            model_period = f"{period_bucket}{month_token[:3].title()}_{year}"
            model_date = ""
    else:
        model_period = "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def close_source_workbook(wb: xw.Book) -> None:
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
        # Nothing else to do; workbook may already be closed.
        pass


def choose_output_path(input_folder: Path, destination_folder: Path) -> Path:
    destination_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM.xlsx"
    candidate = destination_folder / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = destination_folder / f"{input_folder.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def empirical_rows_from_sheet(
    sheet: xw.Sheet,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  empirical: skipped, could not find 'max' anchor in sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    data_start = anchor_row + 1
    data_end = data_start + N_QUARTERS - 1

    header_maps = build_header_maps(sheet, anchor_row, anchor_col)

    num_quarters_col = match_column(
        header_maps,
        [
            ("num", "quarter"),
            ("quarters", "used"),
            ("n", "quarter"),
        ],
    ) or (anchor_col - 12)

    last_quarter_col = match_column(
        header_maps,
        [
            ("last", "quarter"),
            ("quarter", "used"),
        ],
    ) or (anchor_col - 11)

    forecast_value_col = match_column(
        header_maps,
        [
            ("estimated", "total", "sold"),
            ("forecast", "total"),
            ("tot", "fcst"),
        ],
    ) or (anchor_col - 1)

    actual_value_col = match_column(
        header_maps,
        [
            ("reported", "sales"),
            ("actual", "sales"),
            ("actual", "value"),
        ],
    ) or (anchor_col - 2)

    forecast_min_col = match_column(
        header_maps,
        [
            ("min",),
            ("forecast", "min"),
        ],
    ) or (anchor_col + 1)

    quarterly_sales_col = match_column(
        header_maps,
        [
            ("quarterly", "sales"),
            ("qtr", "sales"),
        ],
    ) or (anchor_col - 10)

    reported_sales_col = match_column(
        header_maps,
        [
            ("reported", "sales"),
            ("sales", "reported"),
        ],
    ) or (anchor_col - 9)

    growth_rate_col = match_column(
        header_maps,
        [
            ("growth", "rate"),
            ("growth", "%"),
        ],
    ) or (anchor_col - 8)

    sales_captured_col = match_column(
        header_maps,
        [
            ("sales", "captured", "db"),
            ("captured", "db"),
            ("penetration", "%"),
        ],
    ) or (anchor_col - 7)

    helper_col = anchor_col + 35

    formula_rows: List[List[str]] = []
    for row in range(data_start, data_end + 1):
        if sales_captured_col and sales_captured_col > 0:
            formula = f'=IFERROR(AVERAGE(R{data_start}C{sales_captured_col}:R{row}C{sales_captured_col}),"")'
        elif quarterly_sales_col and reported_sales_col:
            formula = (
                f'=IFERROR('
                f'SUM(R{data_start}C{quarterly_sales_col}:R{row}C{quarterly_sales_col})/'
                f'SUM(R{data_start}C{reported_sales_col}:R{row}C{reported_sales_col}),'
                f'""'
                f')'
            )
        else:
            formula = '=""'
        formula_rows.append([formula])

    sheet.range((data_start, helper_col), (data_end, helper_col)).formula2 = formula_rows
    sheet.book.app.calculate()

    avg_penetration_values = vertical_values(sheet, data_start, data_end, helper_col)
    num_quarters_values = vertical_values(sheet, data_start, data_end, num_quarters_col)
    last_quarter_values = vertical_values(sheet, data_start, data_end, last_quarter_col)
    forecast_values = vertical_values(sheet, data_start, data_end, forecast_value_col)
    actual_values = vertical_values(sheet, data_start, data_end, actual_value_col)
    max_values = vertical_values(sheet, data_start, data_end, anchor_col)
    min_values = vertical_values(sheet, data_start, data_end, forecast_min_col)
    quarterly_sales_values = vertical_values(sheet, data_start, data_end, quarterly_sales_col)
    reported_sales_values = vertical_values(sheet, data_start, data_end, reported_sales_col)
    growth_rate_values = vertical_values(sheet, data_start, data_end, growth_rate_col)
    sales_captured_values = vertical_values(sheet, data_start, data_end, sales_captured_col)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_forecast = forecast_values[idx]
        row_actual = actual_values[idx]
        row_max = max_values[idx]
        row_min = min_values[idx]
        row_q_sales = quarterly_sales_values[idx]
        row_reported_sales = reported_sales_values[idx]
        row_growth = growth_rate_values[idx]
        row_capture = sales_captured_values[idx]
        row_avg_pen = avg_penetration_values[idx]

        if all(
            is_blank(value)
            for value in (
                row_forecast,
                row_actual,
                row_max,
                row_min,
                row_q_sales,
                row_reported_sales,
                row_growth,
                row_capture,
                row_avg_pen,
            )
        ):
            continue

        max_num = to_float(row_max)
        min_num = to_float(row_min)
        range_width = None
        if max_num is not None and min_num is not None:
            range_width = max_num - min_num

        num_quarters = maybe_number(num_quarters_values[idx])
        if is_blank(num_quarters):
            num_quarters = idx + 1

        avg_penetration = maybe_number(row_avg_pen)
        if is_blank(avg_penetration):
            avg_penetration = maybe_number(row_capture)

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": maybe_number(last_quarter_values[idx]),
                "forecast_value": maybe_number(row_forecast),
                "actual_value": maybe_number(row_actual),
                "forecast_max": maybe_number(row_max),
                "forecast_min": maybe_number(row_min),
                "range_width": maybe_number(range_width),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": maybe_number(row_q_sales),
                "reported_sales": maybe_number(row_reported_sales),
                "growth_rate_pct": maybe_number(row_growth),
                "sales_captured_in_db_pct": maybe_number(row_capture),
                "source_file": source_file,
            }
        )

    return rows


def regression_rows_from_sheet(
    sheet: xw.Sheet,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  regression: skipped, could not find 'max' anchor in sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    data_start = anchor_row + 1
    data_end = data_start + N_QUARTERS - 1
    header_maps = build_header_maps(sheet, anchor_row, anchor_col)

    num_quarters_col = match_column(
        header_maps,
        [
            ("num", "quarter"),
            ("quarters", "used"),
            ("n", "quarter"),
        ],
    ) or (anchor_col - 12)

    forecast_total_col = match_column(
        header_maps,
        [
            ("tot", "fcst", "w/o", "sa"),
            ("total", "forecast", "without", "sa"),
            ("forecast", "total"),
        ],
    ) or (anchor_col - 1)

    forecast_min_col = match_column(
        header_maps,
        [
            ("min",),
            ("forecast", "min"),
        ],
    ) or (anchor_col + 1)

    helper_intercept_col = anchor_col + 35
    helper_slope_col = anchor_col + 36

    intercept_formulas: List[List[str]] = []
    slope_formulas: List[List[str]] = []
    for row in range(data_start, data_end + 1):
        intercept_formulas.append(
            [
                f'=IFERROR(INTERCEPT(R{data_start}C{y_col}:R{row}C{y_col},'
                f'R{data_start}C{x_col}:R{row}C{x_col}),"")'
            ]
        )
        slope_formulas.append(
            [
                f'=IFERROR(SLOPE(R{data_start}C{y_col}:R{row}C{y_col},'
                f'R{data_start}C{x_col}:R{row}C{x_col}),"")'
            ]
        )

    sheet.range((data_start, helper_intercept_col), (data_end, helper_intercept_col)).formula2 = intercept_formulas
    sheet.range((data_start, helper_slope_col), (data_end, helper_slope_col)).formula2 = slope_formulas
    sheet.book.app.calculate()

    num_quarters_values = vertical_values(sheet, data_start, data_end, num_quarters_col)
    forecast_values = vertical_values(sheet, data_start, data_end, forecast_total_col)
    max_values = vertical_values(sheet, data_start, data_end, anchor_col)
    min_values = vertical_values(sheet, data_start, data_end, forecast_min_col)
    intercept_values = vertical_values(sheet, data_start, data_end, helper_intercept_col)
    slope_values = vertical_values(sheet, data_start, data_end, helper_slope_col)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters = maybe_number(num_quarters_values[idx])
        if is_blank(num_quarters):
            num_quarters = idx + 1

        forecast_value = maybe_number(forecast_values[idx])
        forecast_max = maybe_number(max_values[idx])
        forecast_min = maybe_number(min_values[idx])
        intercept = maybe_number(intercept_values[idx])
        slope = maybe_number(slope_values[idx])

        if all(is_blank(value) for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = None
        if max_num is not None and min_num is not None:
            range_width = max_num - min_num

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": maybe_number(range_width),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows:
            previous = rows[-1]
            duplicate = all(
                values_equal(row[key], previous[key])
                for key in (
                    "num_quarters_used",
                    "forecast_value",
                    "forecast_max",
                    "forecast_min",
                    "intercept",
                    "slope",
                )
            )
            if duplicate:
                continue

        rows.append(row)

    return rows


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    max_width = 48
    min_width = 12
    for col_idx, column in enumerate(columns, start=1):
        values = [column]
        values.extend(
            "" if row.get(column) is None else str(row.get(column))
            for row in rows
        )
        width = max(len(value) for value in values) + 2
        width = min(max_width, max(min_width, width))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def iter_source_files(folder: Path) -> Iterable[Path]:
    if not folder.exists():
        print(f"Input directory does not exist: {folder}")
        return

    for path in sorted(folder.iterdir()):
        if path.is_dir():
            print(f"Skipped {path.name}: directory")
            continue

        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temp file")
            continue

        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue

        if re.search(r"_PARAM(?:\.\d+)?$", path.stem, flags=re.IGNORECASE):
            print(f"Skipped {path.name}: appears to be an output workbook")
            continue

        yield path


def run() -> None:
    output_path = choose_output_path(input_dir, output_dir)
    source_files = list(iter_source_files(input_dir))
    if not source_files:
        print("No source files to process.")
        print(f"Output path (reserved): {output_path}")
        return

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open ({exc})")
                continue

            try:
                meta = parse_file_label(file_path)

                try:
                    empirical_sheet = workbook.sheets["Empirical Model"]
                    empirical_rows.extend(
                        empirical_rows_from_sheet(
                            empirical_sheet,
                            meta=meta,
                            source_file=file_path.name,
                        )
                    )
                except Exception as exc:
                    print(f"  empirical: skipped for {file_path.name} ({exc})")

                try:
                    regression_sheet = workbook.sheets["Regression Model"]
                    regression_rows.extend(
                        regression_rows_from_sheet(
                            regression_sheet,
                            meta=meta,
                            source_file=file_path.name,
                        )
                    )
                except Exception as exc:
                    print(f"  regression: skipped for {file_path.name} ({exc})")

                files_processed += 1
            finally:
                close_source_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
