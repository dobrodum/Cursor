#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet

# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

MODEL_FILE_RE = re.compile(
    r"^.*?-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<timing>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})",
    re.IGNORECASE,
)

TIMING_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_LOOKUP = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct")
    text = re.sub(r"[_/\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def parse_model_metadata(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    match = MODEL_FILE_RE.search(stem)
    if not match:
        raise ValueError("filename does not match expected model naming convention")

    ticker = match.group("ticker").upper()
    timing = match.group("timing").title()
    month_token = match.group("month").lower()
    year = int(match.group("year"))
    month = MONTH_LOOKUP.get(month_token)
    if month is None:
        raise ValueError(f"unsupported month token '{match.group('month')}' in filename")

    day = TIMING_DAY[timing.lower()]
    month_abbr = date(year, month, 1).strftime("%b")
    model_period = f"{timing}{month_abbr}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if isinstance(values, list):
        if not values:
            return []
        first = values[0]
        if isinstance(first, tuple):
            return [list(row) for row in values]
        if isinstance(first, list):
            return values
        return [values]
    return [[values]]


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None
    first_row = used.row
    first_col = used.column

    anchor_text = anchor_text.strip().lower()
    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == anchor_text:
                return first_row + r_idx, first_col + c_idx
    return None


def build_header_map(sheet: xw.Sheet, header_row: int) -> tuple[dict[int, str], int, int]:
    used = sheet.used_range
    first_col = used.column
    last_col = used.column + used.columns.count - 1
    row_values = sheet.range((header_row, first_col), (header_row, last_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]

    headers: dict[int, str] = {}
    for idx, value in enumerate(row_values):
        normalized = normalize_text(value)
        if normalized:
            headers[first_col + idx] = normalized
    return headers, first_col, last_col


def find_column(
    headers: dict[int, str], keywords: list[str], anchor_col: int | None = None
) -> int | None:
    keyword_set = [normalize_text(key) for key in keywords]
    candidates: list[int] = []
    for col, header in headers.items():
        if any(keyword in header for keyword in keyword_set):
            candidates.append(col)

    if not candidates:
        return None
    if anchor_col is None:
        return candidates[0]
    return min(candidates, key=lambda col: abs(col - anchor_col))


def safe_cell_value(sheet: xw.Sheet, row: int, col: int | None) -> Any:
    if col is None or row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def calculate_once(wb: xw.Book, formulas_written: bool) -> None:
    if formulas_written:
        wb.app.calculate()


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Saved = True
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def detect_numeric_block(sheet: xw.Sheet, col: int, max_row: int) -> tuple[int, int] | None:
    row = max_row
    while row >= 1 and to_float(safe_cell_value(sheet, row, col)) is None:
        row -= 1
    if row < 1:
        return None
    end_row = row
    while row >= 1 and to_float(safe_cell_value(sheet, row, col)) is not None:
        row -= 1
    start_row = row + 1
    return start_row, end_row


def is_all_blank(values: list[Any]) -> bool:
    return all(value in (None, "") for value in values)


def calculate_range_width(max_value: Any, min_value: Any) -> float | None:
    max_float = to_float(max_value)
    min_float = to_float(min_value)
    if max_float is None or min_float is None:
        return None
    return max_float - min_float


def extract_empirical_candidates(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"skipped empirical sheet in {source_file}: sheet not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped empirical sheet in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    headers, _, _ = build_header_map(sheet, anchor_row)

    max_col = anchor_col
    min_col = find_column(headers, ["min"], anchor_col) or (anchor_col + 1)
    num_quarters_col = find_column(
        headers, ["num quarters used", "num quarters", "quarters used", "n quarters"], anchor_col
    )
    last_quarter_col = find_column(headers, ["last quarter used", "last quarter"], anchor_col)
    avg_penetration_col = find_column(
        headers, ["avg penetration pct", "average penetration", "avg penetration"], anchor_col
    )
    forecast_col = find_column(
        headers,
        [
            "estimated total sold",
            "estimate total sold",
            "forecast value",
            "forecast",
            "tot fcst",
        ],
        anchor_col,
    )
    actual_col = find_column(headers, ["reported sales", "actual value", "actual"], anchor_col)
    quarterly_sales_col = find_column(headers, ["quarterly sales"], anchor_col)
    reported_sales_col = find_column(headers, ["reported sales"], anchor_col) or actual_col
    growth_rate_col = find_column(headers, ["growth rate pct", "growth rate"], anchor_col)
    sales_captured_col = find_column(
        headers, ["sales captured in db pct", "captured in db", "sales captured"], anchor_col
    )

    formulas_written = False
    if avg_penetration_col is not None:
        for n_quarters in range(1, 11):
            row = anchor_row + n_quarters
            avg_cell = sheet.cells(row, avg_penetration_col)
            try:
                existing_formula = avg_cell.formula
            except Exception:
                existing_formula = None

            if isinstance(existing_formula, str) and existing_formula.startswith("="):
                set_formula2(avg_cell, existing_formula)
                formulas_written = True
                continue

            left_span = min(n_quarters, max(avg_penetration_col - 1, 1))
            if left_span > 0:
                set_formula2(avg_cell, f"=AVERAGE(RC[-{left_span}]:RC[-1])")
                formulas_written = True

    calculate_once(wb, formulas_written)

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, 11):
        row = anchor_row + n_quarters
        row_snapshot = [
            safe_cell_value(sheet, row, num_quarters_col),
            safe_cell_value(sheet, row, last_quarter_col),
            safe_cell_value(sheet, row, avg_penetration_col),
            safe_cell_value(sheet, row, forecast_col),
            safe_cell_value(sheet, row, actual_col),
            safe_cell_value(sheet, row, max_col),
            safe_cell_value(sheet, row, min_col),
        ]
        if is_all_blank(row_snapshot):
            continue

        avg_penetration_pct = safe_cell_value(sheet, row, avg_penetration_col)
        forecast_max = safe_cell_value(sheet, row, max_col)
        forecast_min = safe_cell_value(sheet, row, min_col)
        reported_sales_value = safe_cell_value(sheet, row, reported_sales_col)

        num_quarters_used = safe_cell_value(sheet, row, num_quarters_col)
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

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
                "last_quarter_used": safe_cell_value(sheet, row, last_quarter_col),
                "forecast_value": safe_cell_value(sheet, row, forecast_col),
                "actual_value": safe_cell_value(sheet, row, actual_col),
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calculate_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": safe_cell_value(sheet, row, quarterly_sales_col),
                "reported_sales": reported_sales_value,
                "growth_rate_pct": safe_cell_value(sheet, row, growth_rate_col),
                "sales_captured_in_db_pct": safe_cell_value(sheet, row, sales_captured_col),
                "source_file": source_file,
            }
        )

    return rows


def rounded_signature(*values: Any) -> tuple[Any, ...]:
    out: list[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is None:
            out.append(value)
        else:
            out.append(round(numeric, 10))
    return tuple(out)


def extract_regression_candidates(
    wb: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"skipped regression sheet in {source_file}: sheet not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped regression sheet in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    headers, _, _ = build_header_map(sheet, anchor_row)

    y_col = max(anchor_col - 7, 1)
    x_col = max(anchor_col - 11, 1)

    max_col = anchor_col
    min_col = find_column(headers, ["min"], anchor_col) or (anchor_col + 1)
    num_quarters_col = find_column(
        headers, ["num quarters used", "num quarters", "quarters used", "n quarters"], anchor_col
    )
    intercept_col = find_column(headers, ["intercept"], anchor_col)
    slope_col = find_column(headers, ["slope"], anchor_col)
    forecast_col = find_column(
        headers,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast total without sa", "tot fcst"],
        anchor_col,
    )
    actual_col = find_column(headers, ["actual value", "actual", "reported sales"], anchor_col)

    block = detect_numeric_block(sheet, y_col, anchor_row - 1)
    formulas_written = False
    if block and intercept_col is not None and slope_col is not None:
        data_start, data_end = block
        for n_quarters in range(1, 11):
            result_row = anchor_row + n_quarters
            start_row = max(data_start, data_end - n_quarters + 1)
            used_points = data_end - start_row + 1

            y_ref = f"R{start_row}C{y_col}:R{data_end}C{y_col}"
            x_ref = f"R{start_row}C{x_col}:R{data_end}C{x_col}"
            set_formula2(sheet.cells(result_row, intercept_col), f"=INTERCEPT({y_ref},{x_ref})")
            set_formula2(sheet.cells(result_row, slope_col), f"=SLOPE({y_ref},{x_ref})")
            formulas_written = True

            if num_quarters_col is not None:
                sheet.cells(result_row, num_quarters_col).value = used_points

    calculate_once(wb, formulas_written)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for n_quarters in range(1, 11):
        row = anchor_row + n_quarters
        num_quarters_used = safe_cell_value(sheet, row, num_quarters_col)
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        intercept_value = safe_cell_value(sheet, row, intercept_col)
        slope_value = safe_cell_value(sheet, row, slope_col)
        forecast_value = safe_cell_value(sheet, row, forecast_col)
        forecast_max = safe_cell_value(sheet, row, max_col)
        forecast_min = safe_cell_value(sheet, row, min_col)

        if is_all_blank(
            [num_quarters_used, intercept_value, slope_value, forecast_value, forecast_max, forecast_min]
        ):
            continue

        signature = rounded_signature(
            num_quarters_used,
            intercept_value,
            slope_value,
            forecast_value,
            forecast_max,
            forecast_min,
        )
        if signature == previous_signature:
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
                "actual_value": safe_cell_value(sheet, row, actual_col),
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calculate_range_width(forecast_max, forecast_min),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def choose_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_sheet(ws: Worksheet, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(
            max(max_len + 2, 12), 48
        )


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_input_files(path: Path) -> list[Path]:
    files = sorted(path.iterdir(), key=lambda p: p.name.lower())
    selected: list[Path] = []
    for file_path in files:
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped: {file_path.name} (reason: temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {file_path.name} (reason: not an .xlsx file)")
            continue
        selected.append(file_path)
    return selected


def main() -> None:
    source_input_dir = Path(input_dir).expanduser().resolve()
    target_output_dir = Path(output_dir).expanduser().resolve()

    if not source_input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {source_input_dir}")
    if not source_input_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a folder: {source_input_dir}")

    target_output_dir.mkdir(parents=True, exist_ok=True)
    output_path = choose_output_path(source_input_dir, target_output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    original_calculation = None
    try:
        original_calculation = app.calculation
        app.calculation = "manual"
    except Exception:
        original_calculation = None

    try:
        for file_path in iter_input_files(source_input_dir):
            if source_input_dir == target_output_dir and file_path.name == output_path.name:
                print(f"skipped: {file_path.name} (reason: output workbook)")
                continue

            try:
                metadata = parse_model_metadata(file_path.name)
            except ValueError as exc:
                print(f"skipped: {file_path.name} (reason: {exc})")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical = extract_empirical_candidates(workbook, metadata, file_path.name)
                regression = extract_regression_candidates(workbook, metadata, file_path.name)
                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_count += 1
                print(
                    f"processed: {file_path.name} (empirical_rows={len(empirical)}, "
                    f"regression_rows={len(regression)})"
                )
            except Exception as exc:
                print(f"skipped: {file_path.name} (reason: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        try:
            if original_calculation is not None:
                app.calculation = original_calculation
        except Exception:
            pass
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_count}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
