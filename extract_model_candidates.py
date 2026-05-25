#!/usr/bin/env python3
"""Extract empirical and regression candidate parameters from Excel model files."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure paths here.
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}

EMPIRICAL_HEADER_ALIASES = {
    "num_quarters_used": ("num quarters", "quarters used", "n quarters"),
    "last_quarter_used": ("last quarter",),
    "forecast_value": (
        "estimated total sold",
        "est total sold",
        "forecast value",
        "tot fcst",
        "total sold",
    ),
    "actual_value": ("reported sales", "actual sales", "actual value", "actual"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "avg_penetration_pct": ("avg penetration", "average penetration"),
    "quarterly_sales": ("quarterly sales", "quarter sales", "q sales"),
    "reported_sales": ("reported sales",),
    "growth_rate_pct": ("growth rate", "growth pct", "growth %"),
    "sales_captured_in_db_pct": (
        "sales captured in db",
        "captured in db",
        "db penetration",
        "penetration",
    ),
}

REGRESSION_HEADER_ALIASES = {
    "num_quarters_used": ("num quarters", "quarters used", "n quarters"),
    "forecast_value": (
        "tot fcst w/o sa",
        "tot fcst without sa",
        "forecast value",
        "total forecast",
    ),
    "actual_value": ("actual", "actual value", "reported sales"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
}

# Offset fallbacks (relative to the "max" anchor column) used when headers are absent.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "forecast_value": -4,
    "actual_value": -3,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -6,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -2,
    "sales_captured_in_db_pct": -5,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "forecast_value": -4,
    "actual_value": -3,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s_]+", " ", text)
    text = re.sub(r"[^a-z0-9 %/()-]+", "", text)
    return text.strip()


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("%", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def subtract_numbers(lhs: Any, rhs: Any) -> Optional[float]:
    left = to_number(lhs)
    right = to_number(rhs)
    if left is None or right is None:
        return None
    return left - right


def parse_file_metadata(file_path: Path) -> Optional[FileMetadata]:
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1]
    period_segment = parts[2]
    period_match = re.search(
        r"(Early|Mid|Late)([A-Za-z]+)(\d{4})",
        period_segment,
        flags=re.IGNORECASE,
    )
    if not period_match:
        return None

    period_prefix = period_match.group(1).title()
    month_text = period_match.group(2).title()
    year_text = period_match.group(3)

    period_key = period_prefix.lower()
    day_value = DAY_BY_PERIOD.get(period_key)
    if day_value is None:
        return None

    month_abbrev = month_text[:3]
    try:
        month_value = date.fromisoformat(f"{year_text}-01-01").replace(
            month=month_name_to_number(month_abbrev)
        ).month
    except ValueError:
        return None

    model_period = f"{period_prefix}{month_abbrev}_{year_text}"
    model_date = date(int(year_text), month_value, day_value).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def month_name_to_number(month_text: str) -> int:
    month_lookup = {
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
    key = month_text.lower()[:3]
    if key not in month_lookup:
        raise ValueError(f"Unknown month abbreviation: {month_text}")
    return month_lookup[key]


def resolve_output_path(input_folder: Path, destination_folder: Path) -> Path:
    destination_folder.mkdir(parents=True, exist_ok=True)
    stem = f"{input_folder.name}_PARAM"
    candidate = destination_folder / f"{stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = destination_folder / f"{stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def iter_input_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue
        yield path


def get_sheet_by_name(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor_cell(sheet: xw.Sheet, target: str = "max") -> Optional[xw.Range]:
    used_range = sheet.used_range
    values = as_2d(used_range.value)
    if not values:
        return None

    target_text = normalize_text(target)
    first_row = used_range.row
    first_col = used_range.column

    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target_text:
                return sheet.cells(first_row + row_idx, first_col + col_idx)
    return None


def collect_header_candidates(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    scan_width: int = 45,
) -> Dict[str, int]:
    min_col = max(1, anchor_col - scan_width)
    max_col = anchor_col + scan_width
    header_rows = [row for row in (anchor_row - 1, anchor_row, anchor_row + 1) if row >= 1]
    headers: Dict[str, int] = {}

    for row in header_rows:
        row_values = sheet.range((row, min_col), (row, max_col)).value
        if row_values is None:
            continue
        if not isinstance(row_values, list):
            row_values = [row_values]
        for idx, value in enumerate(row_values):
            key = normalize_text(value)
            if key and key not in headers:
                headers[key] = min_col + idx
    return headers


def map_columns(
    header_candidates: Dict[str, int],
    aliases: Dict[str, Tuple[str, ...]],
) -> Dict[str, Optional[int]]:
    mapped: Dict[str, Optional[int]] = {}
    for field_name, alias_values in aliases.items():
        mapped[field_name] = None
        for alias in alias_values:
            normalized_alias = normalize_text(alias)
            for header_text, col in header_candidates.items():
                if normalized_alias and normalized_alias in header_text:
                    mapped[field_name] = col
                    break
            if mapped[field_name] is not None:
                break
    return mapped


def value_from_row(
    sheet: xw.Sheet,
    row_number: int,
    anchor_col: int,
    col_map: Dict[str, Optional[int]],
    field_name: str,
    fallback_offsets: Dict[str, int],
) -> Any:
    col = col_map.get(field_name)
    if col is None:
        col = anchor_col + fallback_offsets[field_name]
    if col < 1:
        return None
    return sheet.cells(row_number, col).value


def set_formula2(target_cell: xw.Range, formula: str) -> None:
    try:
        target_cell.formula2 = formula
    except Exception:
        target_cell.formula = formula


def close_workbook_safely(workbook: xw.Book) -> None:
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


def extract_empirical_rows(sheet: xw.Sheet, metadata: FileMetadata) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    anchor = find_anchor_cell(sheet, target="max")
    if anchor is None:
        print("  - Empirical Model: skipped, 'max' anchor not found")
        return rows

    anchor_row = anchor.row
    anchor_col = anchor.column
    output_start_row = anchor_row + 1
    source_end_row = anchor_row - 1
    used_last_col = sheet.used_range.last_cell.column

    header_candidates = collect_header_candidates(sheet, anchor_row, anchor_col)
    col_map = map_columns(header_candidates, EMPIRICAL_HEADER_ALIASES)
    scratch_col = used_last_col + 6

    penetration_col = (
        col_map.get("sales_captured_in_db_pct")
        or col_map.get("avg_penetration_pct")
        or (anchor_col + EMPIRICAL_FALLBACK_OFFSETS["sales_captured_in_db_pct"])
    )

    if penetration_col >= 1 and source_end_row >= 1:
        for idx in range(N_QUARTERS):
            quarter_count = idx + 1
            start_row = source_end_row - quarter_count + 1
            if start_row < 1:
                start_row = 1
            formula_cell = sheet.cells(output_start_row + idx, scratch_col)
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:'
                f"R{source_end_row}C{penetration_col}),\"\")"
            )
            # Temporary write-only formulas to trigger workbook calculations.
            set_formula2(formula_cell, formula)
        sheet.book.app.calculate()

    for idx in range(N_QUARTERS):
        quarter_count = idx + 1
        row_number = output_start_row + idx

        avg_penetration = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "avg_penetration_pct",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        if (avg_penetration is None or avg_penetration == "") and penetration_col >= 1:
            avg_penetration = sheet.cells(row_number, scratch_col).value

        num_quarters_used = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "num_quarters_used",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = quarter_count

        forecast_value = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "forecast_value",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        actual_value = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "actual_value",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        forecast_max = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "forecast_max",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        forecast_min = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "forecast_min",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        range_width = subtract_numbers(forecast_max, forecast_min)

        last_quarter_used = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "last_quarter_used",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        quarterly_sales = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "quarterly_sales",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        reported_sales = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "reported_sales",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        growth_rate_pct = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "growth_rate_pct",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        sales_captured_in_db_pct = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "sales_captured_in_db_pct",
            EMPIRICAL_FALLBACK_OFFSETS,
        )

        key_values = (
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            avg_penetration,
            quarterly_sales,
            reported_sales,
        )
        if all(value in (None, "") for value in key_values):
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": metadata.source_file,
            }
        )

    return rows


def extract_regression_rows(sheet: xw.Sheet, metadata: FileMetadata) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    anchor = find_anchor_cell(sheet, target="max")
    if anchor is None:
        print("  - Regression Model: skipped, 'max' anchor not found")
        return rows

    anchor_row = anchor.row
    anchor_col = anchor.column
    output_start_row = anchor_row + 1
    source_end_row = anchor_row - 1
    used_last_col = sheet.used_range.last_cell.column

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    header_candidates = collect_header_candidates(sheet, anchor_row, anchor_col)
    col_map = map_columns(header_candidates, REGRESSION_HEADER_ALIASES)

    scratch_intercept_col = used_last_col + 6
    scratch_slope_col = used_last_col + 7
    scratch_forecast_col = used_last_col + 8

    if y_col >= 1 and x_col >= 1 and source_end_row >= 1:
        for idx in range(N_QUARTERS):
            quarter_count = idx + 1
            output_row = output_start_row + idx
            start_row = source_end_row - quarter_count + 1
            if start_row < 1:
                start_row = 1

            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{source_end_row}C{y_col},'
                f"R{start_row}C{x_col}:R{source_end_row}C{x_col}),\"\")"
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{source_end_row}C{y_col},'
                f"R{start_row}C{x_col}:R{source_end_row}C{x_col}),\"\")"
            )
            forecast_formula = (
                f'=IFERROR(RC[-2]+RC[-1]*R{source_end_row}C{x_col},"")'
            )

            # Temporary write-only formulas to trigger workbook calculations.
            set_formula2(sheet.cells(output_row, scratch_intercept_col), intercept_formula)
            set_formula2(sheet.cells(output_row, scratch_slope_col), slope_formula)
            set_formula2(sheet.cells(output_row, scratch_forecast_col), forecast_formula)

        sheet.book.app.calculate()

    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx in range(N_QUARTERS):
        quarter_count = idx + 1
        row_number = output_start_row + idx

        num_quarters_used = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "num_quarters_used",
            REGRESSION_FALLBACK_OFFSETS,
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = quarter_count

        forecast_value = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "forecast_value",
            REGRESSION_FALLBACK_OFFSETS,
        )
        if forecast_value in (None, ""):
            forecast_value = sheet.cells(row_number, scratch_forecast_col).value

        actual_value = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "actual_value",
            REGRESSION_FALLBACK_OFFSETS,
        )

        forecast_max = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "forecast_max",
            REGRESSION_FALLBACK_OFFSETS,
        )
        forecast_min = value_from_row(
            sheet,
            row_number,
            anchor_col,
            col_map,
            "forecast_min",
            REGRESSION_FALLBACK_OFFSETS,
        )
        range_width = subtract_numbers(forecast_max, forecast_min)

        intercept_value = sheet.cells(row_number, scratch_intercept_col).value
        slope_value = sheet.cells(row_number, scratch_slope_col).value

        key_values = (
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept_value,
            slope_value,
        )
        if all(value in (None, "") for value in key_values):
            continue

        signature = (
            to_number(num_quarters_used),
            to_number(forecast_value),
            to_number(forecast_max),
            to_number(forecast_min),
            to_number(intercept_value),
            to_number(slope_value),
        )
        # Guard against duplicated terminal row from model templates.
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": metadata.source_file,
            }
        )

    return rows


def format_output_sheet(sheet: Any, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    sheet.append(columns)
    for col_idx in range(1, len(columns) + 1):
        sheet.cell(row=1, column=col_idx).font = Font(bold=True)

    for row_data in rows:
        sheet.append([row_data.get(column) for column in columns])

    sheet.freeze_panes = "A2"
    if rows:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(rows) + 1}"
    else:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"

    for col_idx, header in enumerate(columns, start=1):
        values = [header]
        for row_data in rows:
            value = row_data.get(header)
            values.append("" if value is None else str(value))
        max_width = max(len(value) for value in values)
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_width + 2), 48)


def write_output_workbook(
    destination_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    format_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    format_output_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(destination_path)


def main() -> None:
    if not input_dir.exists():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    output_path = resolve_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_input_files(input_dir):
            metadata = parse_file_metadata(file_path)
            if metadata is None:
                print(f"Skipped {file_path.name}: filename did not match expected pattern")
                continue

            print(f"Processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)

                empirical_sheet = get_sheet_by_name(workbook, "Empirical Model")
                regression_sheet = get_sheet_by_name(workbook, "Regression Model")

                if empirical_sheet is None:
                    print("  - Empirical Model sheet not found")
                else:
                    empirical_rows.extend(extract_empirical_rows(empirical_sheet, metadata))

                if regression_sheet is None:
                    print("  - Regression Model sheet not found")
                else:
                    regression_rows.extend(extract_regression_rows(regression_sheet, metadata))

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
