from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure paths here.
input_dir = Path("./input")
output_dir = Path("./output")

EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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

N_QUARTERS = 10

MONTH_MAP = {
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

MODEL_PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]+)\s*[_-]?(\d{4})", re.IGNORECASE)


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def log(message: str) -> None:
    print(message, flush=True)


def normalize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        if stripped.upper() in {"#N/A", "#VALUE!", "#DIV/0!", "#REF!", "#NUM!", "#NAME?", "#NULL!"}:
            return None
        return stripped
    return value


def as_number(value: Any) -> float | None:
    value = normalize_value(value)
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = as_number(max_value)
    min_num = as_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_month_token(token: str) -> int:
    key = token.strip().lower()
    if key in MONTH_MAP:
        return MONTH_MAP[key]
    key3 = key[:3]
    if key3 in MONTH_MAP:
        return MONTH_MAP[key3]
    raise ValueError(f"Unsupported month token: {token}")


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = parts[1].upper() if len(parts) >= 2 else ""
    period_source = parts[2] if len(parts) >= 3 else stem

    match = MODEL_PERIOD_PATTERN.search(period_source)
    if not match:
        match = MODEL_PERIOD_PATTERN.search(stem)

    if not match:
        model_period = ""
        model_date = ""
    else:
        phase = match.group(1).title()
        month_token = match.group(2)
        year = int(match.group(3))
        month_num = parse_month_token(month_token)
        month_abbrev = datetime(year, month_num, 1).strftime("%b")
        model_period = f"{phase}{month_abbrev}_{year}"
        day = {"early": 5, "mid": 15, "late": 25}[phase.lower()]
        model_date = date(year, month_num, day).isoformat()

    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    else:
        model = stem.replace(" ", "_")

    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_sheet_case_insensitive(book: xw.Book, target_name: str) -> xw.Sheet | None:
    target = target_name.strip().lower()
    for sheet in book.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []

    if isinstance(values, tuple):
        values = list(values)

    if not isinstance(values, list):
        return [[values]]

    if not values:
        return []

    if isinstance(values[0], tuple):
        values = [list(row) for row in values]

    if not isinstance(values[0], list):
        return [values]

    return values


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column

    best_anchor: tuple[int, int] | None = None
    best_score = -1

    for row_idx, row_values in enumerate(values):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for col_idx, cell_value in enumerate(row_values):
            if not isinstance(cell_value, str) or cell_value.strip().lower() != "max":
                continue

            score = 0
            if col_idx + 1 < len(row_values):
                right_value = row_values[col_idx + 1]
                if isinstance(right_value, str) and right_value.strip().lower() == "min":
                    score += 2

            if row_idx + 1 < len(values):
                below_row = values[row_idx + 1]
                if isinstance(below_row, list) and col_idx < len(below_row):
                    if as_number(below_row[col_idx]) is not None:
                        score += 1

            if score > best_score:
                best_score = score
                best_anchor = (start_row + row_idx, start_col + col_idx)

    return best_anchor


def set_formula2_r1c1(target_range: xw.Range, formula_r1c1: str) -> None:
    try:
        target_range.formula2 = formula_r1c1
    except Exception:
        try:
            target_range.api.Formula2R1C1 = formula_r1c1
        except Exception:
            target_range.api.FormulaR1C1 = formula_r1c1


def close_workbook_without_saving(workbook: xw.Book) -> None:
    attempts: Iterable[Any] = (
        lambda: workbook.close(save=False),
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.api.Close(False),
        lambda: workbook.close(),
    )

    for close_func in attempts:
        try:
            close_func()
            return
        except Exception:
            continue


def safe_get_row(values_2d: list[list[Any]], idx: int) -> list[Any]:
    if idx < 0 or idx >= len(values_2d):
        return []
    row = values_2d[idx]
    if isinstance(row, list):
        return row
    return [row]


def extract_empirical_rows(sheet: xw.Sheet, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        log(f"  - skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    table_left_col = anchor_col - 9
    table_right_col = anchor_col + 1
    helper_col = anchor_col + 12
    penetration_history_col = anchor_col - 6
    history_end_row = anchor_row - 1

    if table_left_col < 1 or penetration_history_col < 1:
        log(f"  - skipped empirical extraction for {source_file}: anchor offsets out of bounds")
        return []

    formulas_written = False
    if history_end_row > 0:
        for quarter_count in range(1, N_QUARTERS + 1):
            target_row = data_start_row + quarter_count - 1
            history_start_row = max(1, history_end_row - quarter_count + 1)
            formula = (
                f"=AVERAGE(R{history_start_row}C{penetration_history_col}:"
                f"R{history_end_row}C{penetration_history_col})"
            )
            set_formula2_r1c1(sheet.range((target_row, helper_col)), formula)
            formulas_written = True

    if formulas_written:
        sheet.book.app.calculate()

    table_values = ensure_2d(
        sheet.range((data_start_row, table_left_col), (data_start_row + N_QUARTERS - 1, table_right_col)).value
    )
    helper_values = ensure_2d(
        sheet.range((data_start_row, helper_col), (data_start_row + N_QUARTERS - 1, helper_col)).value
    )

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_vals = safe_get_row(table_values, idx)
        helper_row = safe_get_row(helper_values, idx)
        if len(row_vals) < 11:
            continue

        num_quarters_used = normalize_value(row_vals[0]) or (idx + 1)
        last_quarter_used = normalize_value(row_vals[1])
        avg_penetration_pct = normalize_value(helper_row[0] if helper_row else None)
        if avg_penetration_pct is None:
            avg_penetration_pct = normalize_value(row_vals[3])

        sales_captured_in_db_pct = normalize_value(row_vals[4])
        growth_rate_pct = normalize_value(row_vals[5])
        quarterly_sales = normalize_value(row_vals[6])
        reported_sales = normalize_value(row_vals[7])
        forecast_value = normalize_value(row_vals[8])
        forecast_max = normalize_value(row_vals[9])
        forecast_min = normalize_value(row_vals[10])
        actual_value = reported_sales
        width = range_width(forecast_max, forecast_min)

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def numbers_close(a: Any, b: Any, tol: float = 1e-9) -> bool:
    an = as_number(a)
    bn = as_number(b)
    if an is None or bn is None:
        return normalize_value(a) == normalize_value(b)
    return abs(an - bn) <= tol


def duplicate_regression_row(previous_row: dict[str, Any], current_row: dict[str, Any]) -> bool:
    keys = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    return all(numbers_close(previous_row.get(key), current_row.get(key)) for key in keys)


def extract_regression_rows(sheet: xw.Sheet, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        log(f"  - skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    table_left_col = anchor_col - 9
    table_right_col = anchor_col + 1
    intercept_helper_col = anchor_col + 12
    slope_helper_col = anchor_col + 13
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = anchor_row - 1

    if min(table_left_col, x_col, y_col) < 1:
        log(f"  - skipped regression extraction for {source_file}: anchor offsets out of bounds")
        return []

    formulas_written = False
    if history_end_row > 1:
        for quarter_count in range(1, N_QUARTERS + 1):
            target_row = data_start_row + quarter_count - 1
            if quarter_count < 2:
                set_formula2_r1c1(sheet.range((target_row, intercept_helper_col)), "=NA()")
                set_formula2_r1c1(sheet.range((target_row, slope_helper_col)), "=NA()")
                formulas_written = True
                continue

            history_start_row = max(1, history_end_row - quarter_count + 1)
            intercept_formula = (
                f"=INTERCEPT(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
                f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
                f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
            )
            set_formula2_r1c1(sheet.range((target_row, intercept_helper_col)), intercept_formula)
            set_formula2_r1c1(sheet.range((target_row, slope_helper_col)), slope_formula)
            formulas_written = True

    if formulas_written:
        sheet.book.app.calculate()

    table_values = ensure_2d(
        sheet.range((data_start_row, table_left_col), (data_start_row + N_QUARTERS - 1, table_right_col)).value
    )
    helper_values = ensure_2d(
        sheet.range(
            (data_start_row, intercept_helper_col),
            (data_start_row + N_QUARTERS - 1, slope_helper_col),
        ).value
    )

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_vals = safe_get_row(table_values, idx)
        helper_row = safe_get_row(helper_values, idx)
        if len(row_vals) < 11:
            continue

        num_quarters_used = normalize_value(row_vals[0]) or (idx + 1)
        actual_candidate = normalize_value(row_vals[7])
        forecast_value = normalize_value(row_vals[8])  # TOT FCST w/o SA
        forecast_max = normalize_value(row_vals[9])
        forecast_min = normalize_value(row_vals[10])
        intercept = normalize_value(helper_row[0] if len(helper_row) > 0 else None)
        slope = normalize_value(helper_row[1] if len(helper_row) > 1 else None)
        width = range_width(forecast_max, forecast_min)

        if all(value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_candidate,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and duplicate_regression_row(rows[-1], row):
            continue
        rows.append(row)

    return rows


def write_sheet(worksheet: Any, headers: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row.get(header) for header in headers])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx in range(1, worksheet.max_column + 1):
        column_letter = get_column_letter(col_idx)
        max_len = len(str(worksheet.cell(row=1, column=col_idx).value or ""))
        for row_idx in range(2, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        worksheet.column_dimensions[column_letter].width = min(max(max_len + 2, 12), 42)


def output_workbook_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    path = output_folder / f"{base_name}.xlsx"
    if not path.exists():
        return path

    suffix = 1
    while True:
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def collect_source_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            log(f"Skipped file: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            log(f"Skipped file: {file_path.name} (not .xlsx)")
            continue
        files.append(file_path)
    return files


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
    if not input_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    source_files = collect_source_files(input_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

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
            log(f"Processing file: {file_path.name}")
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)

                empirical_sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
                if empirical_sheet is None:
                    log(f"  - skipped empirical extraction for {file_path.name}: sheet not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = get_sheet_case_insensitive(workbook, "Regression Model")
                if regression_sheet is None:
                    log(f"  - skipped regression extraction for {file_path.name}: sheet not found")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_count += 1
                log(f"Processed file: {file_path.name}")
            except Exception as exc:
                log(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_without_saving(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_path = output_workbook_path(input_dir, output_dir)
    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    empirical_ws = output_wb.create_sheet("empirical_candidates")
    regression_ws = output_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)
    output_wb.save(output_path)

    log(f"Output path: {output_path}")
    log(f"Number of files processed: {processed_count}")
    log(f"Number of empirical rows: {len(empirical_rows)}")
    log(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
