from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running the script.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}
MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def is_non_empty(value: Any) -> bool:
    return value not in (None, "")


def as_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def calculate_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = as_number(max_value)
    min_num = as_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        max_len = max(len(row) for row in values)
        return [row + [None] * (max_len - len(row)) for row in values]
    return [values]


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) > 1 else ""
    period_source = parts[2] if len(parts) > 2 else ""
    period_token = period_source.split("_", 1)[0].strip()

    model_period = period_token.replace(" ", "_") if period_token else ""
    model_date = ""

    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", period_token, re.IGNORECASE)
    if period_match:
        phase = period_match.group(1).title()
        month_text = period_match.group(2)[:3].title()
        year = int(period_match.group(3))
        month = MONTHS.get(month_text.lower())
        day = DAY_BY_PHASE[phase.lower()]
        if month:
            model_period = f"{phase}{month_text}_{year}"
            model_date = date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else stem.replace(" ", "_")
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def get_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    stem = f"{input_path.name}_PARAM"
    first_choice = output_path / f"{stem}.xlsx"
    if not first_choice.exists():
        return first_choice

    suffix = 1
    while True:
        candidate = output_path / f"{stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        workbook.api.Close(False)


def get_sheet_snapshot(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int, int, int]:
    used_range = sheet.used_range
    matrix = to_2d(used_range.value)
    start_row = used_range.row
    start_col = used_range.column
    if not matrix:
        return matrix, start_row, start_col, start_row, start_col

    last_row = start_row + len(matrix) - 1
    last_col = start_col + len(matrix[0]) - 1
    return matrix, start_row, start_col, last_row, last_col


def find_max_anchor(matrix: Sequence[Sequence[Any]], start_row: int, start_col: int) -> Optional[Tuple[int, int]]:
    best: Optional[Tuple[int, int, int]] = None
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) != "max":
                continue
            right_value = normalize_text(row_values[c_idx + 1]) if c_idx + 1 < len(row_values) else ""
            score = 1 if right_value in {"min", "minimum"} else 0
            anchor_row = start_row + r_idx
            anchor_col = start_col + c_idx
            candidate = (score, anchor_row, anchor_col)
            if best is None or candidate > best:
                best = candidate
    if best is None:
        return None
    return best[1], best[2]


def find_columns_near_anchor(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    anchor_col: int,
    aliases: Dict[str, Sequence[str]],
    row_window: int = 4,
    col_window: int = 28,
) -> Dict[str, int]:
    if not matrix:
        return {}

    row_count = len(matrix)
    col_count = len(matrix[0])
    anchor_r = anchor_row - start_row
    anchor_c = anchor_col - start_col
    found: Dict[str, int] = {}

    row_start = max(0, anchor_r - row_window)
    row_end = min(row_count - 1, anchor_r + row_window)
    col_start = max(0, anchor_c - col_window)
    col_end = min(col_count - 1, anchor_c + col_window)

    for key, key_aliases in aliases.items():
        best: Optional[Tuple[int, int]] = None
        for r_idx in range(row_start, row_end + 1):
            for c_idx in range(col_start, col_end + 1):
                text = normalize_text(matrix[r_idx][c_idx]).replace("_", " ")
                if not text:
                    continue
                if not any(alias in text for alias in key_aliases):
                    continue
                distance = (abs(r_idx - anchor_r) * 3) + abs(c_idx - anchor_c)
                candidate = (distance, start_col + c_idx)
                if best is None or candidate < best:
                    best = candidate
        if best:
            found[key] = best[1]
    return found


def infer_scenario_rows(sheet: xw.Sheet, anchor_row: int, anchor_col: int, n_quarters: int) -> List[int]:
    down_rows = [anchor_row + i for i in range(1, n_quarters + 1)]
    up_rows = [anchor_row - i for i in range(1, n_quarters + 1)]

    down_score = sum(1 for row in down_rows if as_number(sheet.cells(row, anchor_col).value) is not None)
    up_score = sum(1 for row in up_rows if as_number(sheet.cells(row, anchor_col).value) is not None)

    if up_score > down_score:
        return sorted(up_rows)
    return down_rows


def comparable_value(value: Any) -> Any:
    number = as_number(value)
    if number is not None:
        return round(number, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def extract_empirical_rows(workbook: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if EMPIRICAL_SHEET_NAME not in [sheet.name for sheet in workbook.sheets]:
        print(f"Skipped {source_file}: missing '{EMPIRICAL_SHEET_NAME}' sheet.")
        return []

    sheet = workbook.sheets[EMPIRICAL_SHEET_NAME]
    matrix, start_row, start_col, _, last_col = get_sheet_snapshot(sheet)
    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped {source_file}: no 'max' anchor found in '{EMPIRICAL_SHEET_NAME}'.")
        return []

    anchor_row, anchor_col = anchor
    scenario_rows = infer_scenario_rows(sheet, anchor_row, anchor_col, N_QUARTERS)

    aliases = {
        "num_quarters_used": ("num quarters", "quarters used", "n quarters", "qtrs"),
        "last_quarter_used": ("last quarter", "last qtr"),
        "forecast_value": ("estimated total sold", "tot fcst", "forecast", "total sold"),
        "actual_value": ("reported sales", "actual sales", "actual"),
        "avg_penetration_pct": ("avg penetration", "penetration %"),
        "quarterly_sales": ("quarterly sales", "qtr sales"),
        "reported_sales": ("reported sales", "sales reported"),
        "growth_rate_pct": ("growth rate", "growth %"),
        "sales_captured_in_db_pct": ("captured in db", "sales captured"),
    }
    column_map = find_columns_near_anchor(matrix, start_row, start_col, anchor_row, anchor_col, aliases)

    num_quarters_col = column_map.get("num_quarters_used")
    last_quarter_col = column_map.get("last_quarter_used")
    forecast_col = column_map.get("forecast_value", anchor_col - 1 if anchor_col > 1 else anchor_col)
    actual_col = column_map.get("actual_value")
    avg_pen_col = column_map.get("avg_penetration_pct")
    quarterly_sales_col = column_map.get("quarterly_sales")
    reported_sales_col = column_map.get("reported_sales")
    growth_col = column_map.get("growth_rate_pct")
    captured_col = column_map.get("sales_captured_in_db_pct")

    helper_col = last_col + 2
    computed_avg: Dict[int, Any] = {}
    if quarterly_sales_col and reported_sales_col:
        for row in scenario_rows:
            formula = (
                f'=IFERROR('
                f'R{row}C{quarterly_sales_col}/R{row}C{reported_sales_col},'
                f'""'
                f')'
            )
            sheet.cells(row, helper_col).formula2 = formula
        workbook.app.calculate()
        for row in scenario_rows:
            computed_avg[row] = sheet.cells(row, helper_col).value

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(scenario_rows, start=1):
        num_quarters_value = sheet.cells(row, num_quarters_col).value if num_quarters_col else idx
        num_quarters_num = as_number(num_quarters_value)
        num_quarters_used = int(num_quarters_num) if num_quarters_num is not None else idx

        last_quarter_used = sheet.cells(row, last_quarter_col).value if last_quarter_col else ""
        forecast_value = sheet.cells(row, forecast_col).value if forecast_col else None
        actual_value = sheet.cells(row, actual_col).value if actual_col else None
        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, anchor_col + 1).value

        avg_penetration_pct = (
            computed_avg.get(row)
            if row in computed_avg and is_non_empty(computed_avg.get(row))
            else (sheet.cells(row, avg_pen_col).value if avg_pen_col else None)
        )

        quarterly_sales = sheet.cells(row, quarterly_sales_col).value if quarterly_sales_col else None
        reported_sales = sheet.cells(row, reported_sales_col).value if reported_sales_col else actual_value
        growth_rate_pct = sheet.cells(row, growth_col).value if growth_col else None
        sales_captured_in_db_pct = sheet.cells(row, captured_col).value if captured_col else None

        if not any(
            is_non_empty(value)
            for value in (forecast_value, actual_value, forecast_max, forecast_min, avg_penetration_pct)
        ):
            continue

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
                "range_width": calculate_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def infer_data_end_row(sheet: xw.Sheet, anchor_row: int, y_col: int, lookback: int = 80) -> int:
    for row in range(anchor_row - 1, max(1, anchor_row - lookback), -1):
        if as_number(sheet.cells(row, y_col).value) is not None:
            return row
    return max(1, anchor_row - 1)


def extract_regression_rows(workbook: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if REGRESSION_SHEET_NAME not in [sheet.name for sheet in workbook.sheets]:
        print(f"Skipped {source_file}: missing '{REGRESSION_SHEET_NAME}' sheet.")
        return []

    sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    matrix, start_row, start_col, _, last_col = get_sheet_snapshot(sheet)
    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped {source_file}: no 'max' anchor found in '{REGRESSION_SHEET_NAME}'.")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    if x_col < 1 or y_col < 1:
        print(f"Skipped {source_file}: regression anchor offsets are out of bounds.")
        return []

    scenario_rows = infer_scenario_rows(sheet, anchor_row, anchor_col, N_QUARTERS)
    data_end_row = infer_data_end_row(sheet, anchor_row, y_col)

    aliases = {
        "num_quarters_used": ("num quarters", "quarters used", "n quarters", "qtrs"),
        "forecast_value": ("tot fcst w/o sa", "tot fcst wo sa", "total fcst", "forecast"),
        "actual_value": ("actual", "reported sales"),
    }
    column_map = find_columns_near_anchor(matrix, start_row, start_col, anchor_row, anchor_col, aliases)

    num_quarters_col = column_map.get("num_quarters_used")
    forecast_col = column_map.get("forecast_value", anchor_col - 1 if anchor_col > 1 else anchor_col)
    actual_col = column_map.get("actual_value")

    intercept_col = last_col + 2
    slope_col = last_col + 3

    for idx, row in enumerate(scenario_rows, start=1):
        start_data_row = data_end_row - idx + 1
        if start_data_row < 1:
            start_data_row = 1
        intercept_formula = (
            f'=IFERROR(INTERCEPT('
            f'R{start_data_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{start_data_row}C{x_col}:R{data_end_row}C{x_col}),'
            f'""'
            f')'
        )
        slope_formula = (
            f'=IFERROR(SLOPE('
            f'R{start_data_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{start_data_row}C{x_col}:R{data_end_row}C{x_col}),'
            f'""'
            f')'
        )
        sheet.cells(row, intercept_col).formula2 = intercept_formula
        sheet.cells(row, slope_col).formula2 = slope_formula

    workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx, row in enumerate(scenario_rows, start=1):
        num_quarters_value = sheet.cells(row, num_quarters_col).value if num_quarters_col else idx
        num_quarters_num = as_number(num_quarters_value)
        num_quarters_used = int(num_quarters_num) if num_quarters_num is not None else idx

        forecast_value = sheet.cells(row, forecast_col).value if forecast_col else None
        actual_value = sheet.cells(row, actual_col).value if actual_col else ""
        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, anchor_col + 1).value
        intercept_value = sheet.cells(row, intercept_col).value
        slope_value = sheet.cells(row, slope_col).value

        if not any(
            is_non_empty(value)
            for value in (forecast_value, forecast_max, forecast_min, intercept_value, slope_value)
        ):
            continue

        signature = (
            comparable_value(forecast_value),
            comparable_value(forecast_max),
            comparable_value(forecast_min),
            comparable_value(intercept_value),
            comparable_value(slope_value),
        )

        # Avoid duplicate trailing row if the final calculated candidate repeats the prior one.
        if idx == len(scenario_rows) and previous_signature == signature:
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
                "range_width": calculate_range_width(forecast_max, forecast_min),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    worksheet = workbook.create_sheet(sheet_name)
    worksheet.append(list(headers))

    for row in rows:
        worksheet.append([row.get(header, "") for header in headers])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_idx in range(1, worksheet.max_column + 1):
        max_length = 0
        for row_idx in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=column_idx).value
            if value is None:
                length = 0
            else:
                length = len(str(value))
            if length > max_length:
                max_length = length
        width = min(max(max_length + 2, 12), 48)
        worksheet.column_dimensions[get_column_letter(column_idx)].width = width


def save_output(output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_path = get_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file.")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook.")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file.")
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                metadata = parse_file_label(file_path.name)
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)

        save_output(output_path, empirical_rows, regression_rows)

        print(f"Output path: {output_path}")
        print(f"Files processed: {processed_files}")
        print(f"Empirical rows: {len(empirical_rows)}")
        print(f"Regression rows: {len(regression_rows)}")
    finally:
        if app is not None:
            app.quit()


if __name__ == "__main__":
    main()
