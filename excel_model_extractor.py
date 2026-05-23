from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configure these two folders before running.
# ---------------------------------------------------------------------------
input_dir = "./input"
output_dir = "./output"


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
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}


def to_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def as_number(value: Any) -> float | int | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("%", "").strip()
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def as_int_or_none(value: Any) -> int | None:
    num = as_number(value)
    if num is None:
        return None
    return int(round(float(num)))


def empty_to_none(value: Any) -> Any:
    if value == "":
        return None
    return value


def safe_formula2(target_range: Any, formula: str) -> None:
    try:
        target_range.formula2 = formula
    except Exception:
        target_range.formula = formula


def close_workbook_safely(wb: Any) -> None:
    # Never save source workbooks.
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
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        pass


def get_sheet_by_name(wb: Any, target_name: str) -> Any | None:
    needle = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == needle:
            return sheet
    return None


def find_max_anchor(sheet: Any) -> tuple[int, int] | None:
    used = sheet.used_range
    top_row = used.row
    left_col = used.column
    values = to_2d(used.value)

    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return top_row + r_idx, left_col + c_idx
    return None


def parse_filename_metadata(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 and parts[1] else "UNKNOWN"

    token_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )

    if token_match:
        phase_raw, month_raw, year_raw = token_match.groups()
        phase = phase_raw.capitalize()
        month = month_raw.capitalize()
        year = int(year_raw)
    else:
        phase = "Mid"
        month = "Jan"
        year = 1900

    model_period = f"{phase}{month}_{year}"
    model_date = date(year, MONTH_MAP[month], DAY_MAP[phase]).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    serial = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{serial}.xlsx"
        serial += 1
    return candidate


def range_width(max_value: Any, min_value: Any) -> float | int | None:
    max_num = as_number(max_value)
    min_num = as_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def extract_empirical_candidates(
    wb: Any,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Empirical Model")
    if sheet is None:
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    data_start_row = anchor_row + 1
    data_end_row = min(sheet.used_range.last_cell.row, data_start_row + n_quarters - 1)
    if data_end_row < data_start_row:
        return []

    helper_col = anchor_col + 20
    sales_capture_col = anchor_col - 6

    # Write all formulas once, then calculate once for speed.
    for row in range(data_start_row, data_end_row + 1):
        window = row - data_start_row + 1
        if window == 1:
            start_ref = f"RC{sales_capture_col}"
        else:
            start_ref = f"R[-{window - 1}]C{sales_capture_col}"
        formula = f'=IFERROR(AVERAGE({start_ref}:RC{sales_capture_col}),"")'
        safe_formula2(sheet.range((row, helper_col)), formula)
    wb.app.calculate()

    min_col = anchor_col - 9
    max_col = anchor_col + 1
    raw_matrix = to_2d(sheet.range((data_start_row, min_col), (data_end_row, max_col)).value)
    avg_penetrations = to_2d(
        sheet.range((data_start_row, helper_col), (data_end_row, helper_col)).value
    )

    def rel(row_data: list[Any], col_offset: int) -> Any:
        index = (anchor_col + col_offset) - min_col
        if index < 0 or index >= len(row_data):
            return None
        return empty_to_none(row_data[index])

    rows: list[dict[str, Any]] = []
    for idx, row_data in enumerate(raw_matrix):
        avg_pen_val = (
            empty_to_none(avg_penetrations[idx][0])
            if idx < len(avg_penetrations) and avg_penetrations[idx]
            else None
        )
        default_num_quarters = idx + 1
        num_quarters_used = as_int_or_none(rel(row_data, -5)) or default_num_quarters

        forecast_max = rel(row_data, 0)
        forecast_min = rel(row_data, 1)
        record = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen_val,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": rel(row_data, -4),
            "forecast_value": rel(row_data, -2),
            "actual_value": rel(row_data, -1),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_pen_val,
            "quarterly_sales": rel(row_data, -9),
            "reported_sales": rel(row_data, -8),
            "growth_rate_pct": rel(row_data, -7),
            "sales_captured_in_db_pct": rel(row_data, -6),
            "source_file": source_file,
        }
        rows.append(record)

    # Remove temporary formulas (workbook closes without save anyway).
    sheet.range((data_start_row, helper_col), (data_end_row, helper_col)).clear_contents()
    return rows


def row_signature(row: dict[str, Any], keys: list[str]) -> tuple[Any, ...]:
    normalized: list[Any] = []
    for key in keys:
        value = row.get(key)
        if isinstance(value, (int, float)):
            normalized.append(round(float(value), 10))
        else:
            normalized.append(value)
    return tuple(normalized)


def extract_regression_candidates(
    wb: Any,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Regression Model")
    if sheet is None:
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    n_quarters = 10
    data_start_row = anchor_row + 1
    data_end_row = min(sheet.used_range.last_cell.row, data_start_row + n_quarters - 1)
    if data_end_row < data_start_row:
        return []

    intercept_col = anchor_col + 20
    slope_col = anchor_col + 21

    # Write regression formulas once and calculate once.
    for row in range(data_start_row, data_end_row + 1):
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        safe_formula2(sheet.range((row, intercept_col)), intercept_formula)
        safe_formula2(sheet.range((row, slope_col)), slope_formula)
    wb.app.calculate()

    min_col = anchor_col - 11
    max_col = anchor_col + 1
    raw_matrix = to_2d(sheet.range((data_start_row, min_col), (data_end_row, max_col)).value)
    intercept_vals = to_2d(
        sheet.range((data_start_row, intercept_col), (data_end_row, intercept_col)).value
    )
    slope_vals = to_2d(sheet.range((data_start_row, slope_col), (data_end_row, slope_col)).value)

    def rel(row_data: list[Any], col_offset: int) -> Any:
        index = (anchor_col + col_offset) - min_col
        if index < 0 or index >= len(row_data):
            return None
        return empty_to_none(row_data[index])

    rows: list[dict[str, Any]] = []
    compare_keys = [
        "num_quarters_used",
        "intercept",
        "slope",
        "forecast_value",
        "forecast_max",
        "forecast_min",
    ]

    for idx, row_data in enumerate(raw_matrix):
        row_num = idx + 1
        num_quarters_used = as_int_or_none(rel(row_data, -10)) or row_num
        intercept_value = (
            empty_to_none(intercept_vals[idx][0]) if idx < len(intercept_vals) else None
        )
        slope_value = empty_to_none(slope_vals[idx][0]) if idx < len(slope_vals) else None
        forecast_max = rel(row_data, 0)
        forecast_min = rel(row_data, 1)

        record = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": rel(row_data, -1),
            "actual_value": rel(row_data, -2),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width(forecast_max, forecast_min),
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        if rows and row_signature(rows[-1], compare_keys) == row_signature(record, compare_keys):
            continue
        rows.append(record)

    # Remove temporary formulas (workbook closes without save anyway).
    sheet.range((data_start_row, intercept_col), (data_end_row, slope_col)).clear_contents()
    return rows


def write_sheet(
    ws: Any,
    columns: list[str],
    rows: list[dict[str, Any]],
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
        for row in rows:
            value = row.get(column_name)
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def should_skip_file(file_path: Path, input_folder_name: str) -> str | None:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary Excel file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    output_matcher = re.compile(
        rf"^{re.escape(input_folder_name)}_PARAM(?:\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )
    if output_matcher.match(file_path.name):
        return "generated output workbook"
    return None


def main() -> None:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not in_path.exists():
        raise FileNotFoundError(f"Input directory does not exist: {in_path}")
    if not in_path.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {in_path}")

    processed_files = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(in_path.iterdir(), key=lambda p: p.name.lower()):
            skip_reason = should_skip_file(file_path, in_path.name)
            if skip_reason is not None:
                print(f"Skipped file: {file_path.name} ({skip_reason})")
                continue

            print(f"Processing file: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path)
                empirical_rows.extend(
                    extract_empirical_candidates(wb, metadata, file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(wb, metadata, file_path.name)
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    output_file = unique_output_path(in_path, out_path)
    final_wb = Workbook()
    empirical_ws = final_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = final_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    final_wb.save(output_file)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
