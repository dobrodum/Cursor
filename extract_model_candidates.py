#!/usr/bin/env python3
"""
Build one PARAM workbook from all source .xlsx files in an input folder.

The script opens each source workbook only once, processes both
"Empirical Model" and "Regression Model" while it is open, and closes
the source workbook without saving.
"""

from __future__ import annotations

import argparse
import calendar
import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Required top-level variables for folder configuration.
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

PERIOD_DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", re.IGNORECASE)

EMPIRICAL_HEADER_KEYWORDS: dict[str, tuple[str, ...]] = {
    "num_quarters_used": ("numquarters", "numqtrs", "quartersused", "nquarters"),
    "last_quarter_used": ("lastquarterused", "lastquarter", "latestquarter", "quarterused"),
    "forecast_value": ("estimatedtotalsold", "forecastvalue", "forecast", "totfcst", "estimatetotal"),
    "actual_value": ("reportedsales", "actualvalue", "actualsales", "actual"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "avg_penetration_pct": ("avgpenetration", "averagepenetration", "penetrationavg", "penetration"),
    "quarterly_sales": ("quarterlysales", "qtrsales", "salesquarterly"),
    "reported_sales": ("reportedsales", "salesreported"),
    "growth_rate_pct": ("growthrate", "growthpct", "growth"),
    "sales_captured_in_db_pct": ("salescapturedindb", "capturedindb", "dbcapture", "dbcaptured"),
}

REGRESSION_HEADER_KEYWORDS: dict[str, tuple[str, ...]] = {
    "num_quarters_used": ("numquarters", "numqtrs", "quartersused", "nquarters"),
    "forecast_value": ("totfcstwosa", "forecasttotalwithoutsa", "forecastvalue", "forecast"),
    "actual_value": ("actualvalue", "actual", "reportedsales"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
}

# Anchor-based fallback offsets when headers are unavailable.
EMPIRICAL_FALLBACK_OFFSETS: dict[str, int | None] = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -1,
    "actual_value": -4,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -6,
    "quarterly_sales": -5,
    "reported_sales": -4,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -2,
}

REGRESSION_FALLBACK_OFFSETS: dict[str, int | None] = {
    "num_quarters_used": -8,
    "forecast_value": -1,
    "actual_value": None,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass(frozen=True)
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def sanitize_value(value: Any) -> Any:
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        if not value or value.startswith("#"):
            return None
        return value
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def to_number(value: Any) -> float | None:
    value = sanitize_value(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("$", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
            try:
                return float(cleaned) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> int | None:
    number = to_number(value)
    if number is None:
        return None
    return int(round(number))


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], tuple):
            return [list(row) if isinstance(row, tuple) else row for row in values]
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def flatten_column(values: Any) -> list[Any]:
    rows = to_2d(values)
    output: list[Any] = []
    for row in rows:
        if not row:
            output.append(None)
        else:
            output.append(row[0])
    return output


def parse_month(month_token: str) -> int:
    token = re.sub(r"[^A-Za-z]", "", month_token).title()
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token, fmt).month
        except ValueError:
            continue
    if len(token) >= 3:
        return datetime.strptime(token[:3], "%b").month
    raise ValueError(f"Unable to parse month token '{month_token}'")


def parse_file_meta(file_path: Path) -> FileMeta:
    name_no_ext = file_path.stem
    parts = [part.strip() for part in name_no_ext.split(" - ")]
    if len(parts) < 2 or not parts[1]:
        raise ValueError("filename does not contain ticker in expected ' - TICKER - ' format")

    ticker = parts[1].upper()
    period_match = PERIOD_PATTERN.search(name_no_ext)
    if not period_match:
        raise ValueError("filename does not contain Early/Mid/Late + month + year token")

    phase = period_match.group(1).title()
    month = parse_month(period_match.group(2))
    year = int(period_match.group(3))
    day = PERIOD_DAY_MAP[phase.lower()]

    month_label = calendar.month_abbr[month]
    model_period = f"{phase}{month_label}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def get_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    input_folder_name = input_path.name
    base_stem = f"{input_folder_name}_PARAM"
    candidate = output_path / f"{base_stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def collect_source_files(input_path: Path) -> list[Path]:
    files: list[Path] = []
    param_prefix = f"{input_path.name}_PARAM"
    for path in sorted(input_path.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped: {path.name} (not an .xlsx file)")
            continue
        if path.stem.startswith(param_prefix):
            print(f"Skipped: {path.name} (existing PARAM output)")
            continue
        files.append(path)
    return files


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    fallbacks = (
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.api.Close(False),
        lambda: workbook.api.Close(0),
    )
    for close_call in fallbacks:
        try:
            close_call()
            return
        except Exception:
            continue


def set_formula2(range_obj: xw.Range, formula_values: list[list[str]]) -> None:
    try:
        range_obj.formula2 = formula_values
    except Exception:
        range_obj.formula = formula_values


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int]:
    used_range = sheet.used_range
    values = to_2d(used_range.value)
    start_row = used_range.row
    start_col = used_range.column

    matches: list[tuple[int, int]] = []
    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                matches.append((start_row + row_idx, start_col + col_idx))

    if not matches:
        raise ValueError(f"missing 'max' anchor on '{sheet.name}'")

    for row, col in matches:
        right_value = sanitize_value(sheet.cells(row, col + 1).value)
        if normalize_text(right_value) == "min":
            return row, col

    return matches[0]


def get_header_map(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> dict[str, int]:
    min_col = max(1, anchor_col - 20)
    max_col = anchor_col + 10
    header_map: dict[str, int] = {}

    for row in (anchor_row - 1, anchor_row):
        if row < 1:
            continue
        values = to_2d(sheet.range((row, min_col), (row, max_col)).value)
        if not values:
            continue
        row_values = values[0]
        for idx, value in enumerate(row_values):
            key = normalize_text(value)
            if key and key not in header_map:
                header_map[key] = min_col + idx
    return header_map


def resolve_columns(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    keywords: dict[str, tuple[str, ...]],
    fallback_offsets: dict[str, int | None],
) -> dict[str, int | None]:
    header_map = get_header_map(sheet, anchor_row, anchor_col)
    resolved: dict[str, int | None] = {}

    for field_name, key_tokens in keywords.items():
        chosen_col = None
        for key, col in header_map.items():
            if any(token in key for token in key_tokens):
                chosen_col = col
                break
        if chosen_col is None:
            offset = fallback_offsets.get(field_name)
            chosen_col = anchor_col + offset if offset is not None else None
        resolved[field_name] = chosen_col
    return resolved


def choose_penetration_history_col(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    empirical_cols: dict[str, int | None],
) -> int:
    candidates: list[int] = []
    preferred = empirical_cols.get("avg_penetration_pct")
    if preferred is not None and preferred > 0:
        candidates.append(preferred)
    for offset in (-6, -7, -8, -9):
        candidate = anchor_col + offset
        if candidate > 0 and candidate not in candidates:
            candidates.append(candidate)

    end_row = anchor_row - 1
    start_row = max(1, end_row - 30)
    if end_row < start_row:
        return candidates[0] if candidates else max(1, anchor_col - 6)

    for col in candidates:
        values = flatten_column(sheet.range((start_row, col), (end_row, col)).value)
        numeric_count = sum(1 for value in values if to_number(value) is not None)
        if numeric_count >= 3:
            return col

    return candidates[0] if candidates else max(1, anchor_col - 6)


def block_value(block: list[list[Any]], row_idx: int, min_col: int, col: int | None) -> Any:
    if col is None:
        return None
    if row_idx < 0 or row_idx >= len(block):
        return None
    row = block[row_idx]
    rel_idx = col - min_col
    if rel_idx < 0 or rel_idx >= len(row):
        return None
    return sanitize_value(row[rel_idx])


def has_empirical_payload(row: dict[str, Any]) -> bool:
    fields = (
        "forecast_value",
        "actual_value",
        "forecast_max",
        "forecast_min",
        "avg_penetration_pct",
        "quarterly_sales",
        "reported_sales",
        "growth_rate_pct",
        "sales_captured_in_db_pct",
    )
    return any(row.get(field) is not None for field in fields)


def has_regression_payload(row: dict[str, Any]) -> bool:
    fields = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    return any(row.get(field) is not None for field in fields)


def same_regression_row(a: dict[str, Any], b: dict[str, Any], tolerance: float = 1e-9) -> bool:
    keys = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    for key in keys:
        av = to_number(a.get(key))
        bv = to_number(b.get(key))
        if av is None and bv is None:
            continue
        if av is None or bv is None:
            return False
        if abs(av - bv) > tolerance:
            return False
    return to_int(a.get("num_quarters_used")) == to_int(b.get("num_quarters_used"))


def process_empirical_sheet(workbook: xw.Book, meta: FileMeta) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped: {meta.source_file} (missing sheet 'Empirical Model')")
        return []

    try:
        anchor_row, anchor_col = find_max_anchor(sheet)
    except ValueError as error:
        print(f"Skipped: {meta.source_file} ({error})")
        return []

    columns = resolve_columns(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        keywords=EMPIRICAL_HEADER_KEYWORDS,
        fallback_offsets=EMPIRICAL_FALLBACK_OFFSETS,
    )

    penetration_col = choose_penetration_history_col(sheet, anchor_row, anchor_col, columns)

    helper_col = max(
        anchor_col + 12,
        max(col for col in columns.values() if col is not None) + 1,
    )
    history_end = anchor_row - 1
    formula_rows: list[list[str]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        history_start = max(1, history_end - n_quarters + 1)
        if history_end < history_start:
            formula_rows.append(["=NA()"])
            continue
        formula_rows.append(
            [
                (
                    f"=IFERROR(AVERAGE(R{history_start}C{penetration_col}:"
                    f"R{history_end}C{penetration_col}),NA())"
                )
            ]
        )

    helper_range = sheet.range(
        (anchor_row + 1, helper_col),
        (anchor_row + N_QUARTERS, helper_col),
    )
    set_formula2(helper_range, formula_rows)
    workbook.app.calculate()

    min_col = min(col for col in columns.values() if col is not None)
    max_col = max(helper_col, max(col for col in columns.values() if col is not None))
    block = to_2d(
        sheet.range(
            (anchor_row + 1, min_col),
            (anchor_row + N_QUARTERS, max_col),
        ).value
    )

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        default_num_quarters = idx + 1
        num_quarters_used = to_int(block_value(block, idx, min_col, columns.get("num_quarters_used")))
        if num_quarters_used is None:
            num_quarters_used = default_num_quarters

        avg_penetration = block_value(block, idx, min_col, columns.get("avg_penetration_pct"))
        if avg_penetration is None:
            avg_penetration = block_value(block, idx, min_col, helper_col)

        forecast_max = to_number(block_value(block, idx, min_col, columns.get("forecast_max")))
        forecast_min = to_number(block_value(block, idx, min_col, columns.get("forecast_min")))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": to_number(avg_penetration),
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": block_value(block, idx, min_col, columns.get("last_quarter_used")),
            "forecast_value": to_number(block_value(block, idx, min_col, columns.get("forecast_value"))),
            "actual_value": to_number(block_value(block, idx, min_col, columns.get("actual_value"))),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": to_number(avg_penetration),
            "quarterly_sales": to_number(
                block_value(block, idx, min_col, columns.get("quarterly_sales"))
            ),
            "reported_sales": to_number(
                block_value(block, idx, min_col, columns.get("reported_sales"))
            ),
            "growth_rate_pct": to_number(
                block_value(block, idx, min_col, columns.get("growth_rate_pct"))
            ),
            "sales_captured_in_db_pct": to_number(
                block_value(block, idx, min_col, columns.get("sales_captured_in_db_pct"))
            ),
            "source_file": meta.source_file,
        }
        if has_empirical_payload(row):
            rows.append(row)
    return rows


def process_regression_sheet(workbook: xw.Book, meta: FileMeta) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped: {meta.source_file} (missing sheet 'Regression Model')")
        return []

    try:
        anchor_row, anchor_col = find_max_anchor(sheet)
    except ValueError as error:
        print(f"Skipped: {meta.source_file} ({error})")
        return []

    columns = resolve_columns(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        keywords=REGRESSION_HEADER_KEYWORDS,
        fallback_offsets=REGRESSION_FALLBACK_OFFSETS,
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    helper_intercept_col = max(
        anchor_col + 12,
        max(col for col in columns.values() if col is not None) + 1,
    )
    helper_slope_col = helper_intercept_col + 1
    history_end = anchor_row - 1

    intercept_formulas: list[list[str]] = []
    slope_formulas: list[list[str]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        history_start = max(1, history_end - n_quarters + 1)
        if history_end < history_start:
            intercept_formulas.append(["=NA()"])
            slope_formulas.append(["=NA()"])
            continue
        intercept_formulas.append(
            [
                (
                    f"=IFERROR(INTERCEPT(R{history_start}C{y_col}:R{history_end}C{y_col},"
                    f"R{history_start}C{x_col}:R{history_end}C{x_col}),NA())"
                )
            ]
        )
        slope_formulas.append(
            [
                (
                    f"=IFERROR(SLOPE(R{history_start}C{y_col}:R{history_end}C{y_col},"
                    f"R{history_start}C{x_col}:R{history_end}C{x_col}),NA())"
                )
            ]
        )

    intercept_range = sheet.range(
        (anchor_row + 1, helper_intercept_col),
        (anchor_row + N_QUARTERS, helper_intercept_col),
    )
    slope_range = sheet.range(
        (anchor_row + 1, helper_slope_col),
        (anchor_row + N_QUARTERS, helper_slope_col),
    )
    set_formula2(intercept_range, intercept_formulas)
    set_formula2(slope_range, slope_formulas)
    workbook.app.calculate()

    min_col = min(
        col
        for col in (
            *[c for c in columns.values() if c is not None],
            helper_intercept_col,
            helper_slope_col,
        )
    )
    max_col = max(
        col
        for col in (
            *[c for c in columns.values() if c is not None],
            helper_intercept_col,
            helper_slope_col,
        )
    )
    block = to_2d(
        sheet.range(
            (anchor_row + 1, min_col),
            (anchor_row + N_QUARTERS, max_col),
        ).value
    )

    next_x_value = to_number(sheet.cells(anchor_row, x_col).value)
    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        default_num_quarters = idx + 1
        num_quarters_used = to_int(block_value(block, idx, min_col, columns.get("num_quarters_used")))
        if num_quarters_used is None:
            num_quarters_used = default_num_quarters

        intercept = to_number(block_value(block, idx, min_col, helper_intercept_col))
        slope = to_number(block_value(block, idx, min_col, helper_slope_col))
        forecast_value = to_number(block_value(block, idx, min_col, columns.get("forecast_value")))
        if forecast_value is None and intercept is not None and slope is not None and next_x_value is not None:
            forecast_value = intercept + (slope * next_x_value)

        forecast_max = to_number(block_value(block, idx, min_col, columns.get("forecast_max")))
        forecast_min = to_number(block_value(block, idx, min_col, columns.get("forecast_min")))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": to_number(block_value(block, idx, min_col, columns.get("actual_value"))),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": meta.source_file,
        }

        if has_regression_payload(row):
            rows.append(row)

    if len(rows) >= 2 and same_regression_row(rows[-1], rows[-2]):
        rows.pop()
    return rows


def format_sheet(sheet, columns: list[str]) -> None:
    header_font = Font(bold=True)
    for cell in sheet[1]:
        cell.font = header_font

    sheet.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    last_row = max(2, sheet.max_row)
    sheet.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)


def write_sheet(
    workbook: Workbook,
    title: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    sheet = workbook.create_sheet(title=title)
    sheet.append(columns)
    for row in rows:
        sheet.append([sanitize_value(row.get(column)) for column in columns])
    format_sheet(sheet, columns)


def write_output(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def run(input_path: Path, output_path: Path) -> None:
    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_file = get_output_path(input_path, output_path)
    source_files = collect_source_files(input_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    try:
        try:
            app.display_alerts = False
        except Exception:
            pass
        try:
            app.screen_updating = False
        except Exception:
            pass

        for source_file in source_files:
            try:
                meta = parse_file_meta(source_file)
            except Exception as error:
                print(f"Skipped: {source_file.name} ({error})")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(source_file), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, meta))
                regression_rows.extend(process_regression_sheet(workbook, meta))
                files_processed += 1
                print(f"Processed: {source_file.name}")
            except Exception as error:
                print(f"Skipped: {source_file.name} (workbook error: {error})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output(output_file, empirical_rows, regression_rows)
    print(f"Output path: {output_file}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract empirical and regression model candidates into one PARAM workbook."
    )
    parser.add_argument("--input-dir", default=str(input_dir), help="Folder containing source .xlsx files")
    parser.add_argument("--output-dir", default=str(output_dir), help="Folder for PARAM output workbook")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    run(
        input_path=Path(args.input_dir).expanduser().resolve(),
        output_path=Path(args.output_dir).expanduser().resolve(),
    )
