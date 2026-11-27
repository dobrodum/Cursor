from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Set these before running.
input_dir = "/path/to/input"
output_dir = "/path/to/output"


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

MAX_QUARTERS = 10
SKIP_OUTPUT_PATTERN = re.compile(r"_PARAM(\.\d+)?\.xlsx$", re.IGNORECASE)

FILENAME_PATTERN = re.compile(
    r"^\s*.+?\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
    r"(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})_Send\s*$",
    re.IGNORECASE,
)

MONTH_TO_NUMBER = {
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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def as_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
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


def to_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_1d(values: Any) -> List[Any]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def safe_subtract(left: Optional[float], right: Optional[float]) -> Optional[float]:
    if left is None or right is None:
        return None
    return left - right


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_safely(wb: Any) -> None:
    if wb is None:
        return
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
        wb.close()
    except Exception:
        pass


def parse_file_metadata(file_path: Path) -> FileMeta:
    stem = file_path.stem
    match = FILENAME_PATTERN.match(stem)
    if not match:
        parts = [part.strip() for part in stem.split("-")]
        ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else stem.upper()
        model_period = "unknown_period"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    ticker = match.group("ticker").upper()
    period_token = match.group("period")

    phase_match = re.match(r"(Early|Mid|Late)", period_token, re.IGNORECASE)
    if not phase_match:
        model_period = period_token
        model_date = ""
    else:
        raw_phase = phase_match.group(1)
        phase = raw_phase[0].upper() + raw_phase[1:].lower()
        month_abbrev = period_token[len(raw_phase) : len(raw_phase) + 3].title()
        year = period_token[-4:]
        model_period = f"{phase}{month_abbrev}_{year}"

        month_number = MONTH_TO_NUMBER.get(month_abbrev.lower())
        day = PHASE_TO_DAY.get(phase.lower())
        if month_number is None or day is None:
            model_date = ""
        else:
            model_date = f"{year}-{month_number:02d}-{day:02d}"

    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_root: Path, output_root: Path) -> Path:
    base_name = f"{input_root.name}_PARAM.xlsx"
    output_path = output_root / base_name
    if not output_path.exists():
        return output_path

    index = 1
    while True:
        candidate = output_root / f"{input_root.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def get_used_bounds(sheet: Any) -> Tuple[int, int, int, int]:
    used = sheet.used_range
    first_row = int(used.row)
    first_col = int(used.column)
    last_row = first_row + int(used.rows.count) - 1
    last_col = first_col + int(used.columns.count) - 1
    return first_row, last_row, first_col, last_col


def find_max_anchor(
    sheet: Any,
    first_row: int,
    last_row: int,
    first_col: int,
    last_col: int,
) -> Optional[Tuple[int, int]]:
    data = sheet.range((first_row, first_col), (last_row, last_col)).value
    matrix = to_2d(data)
    candidates: List[Tuple[int, int]] = []

    for r_off, row in enumerate(matrix):
        for c_off, value in enumerate(row):
            if normalize_text(value) == "max":
                candidates.append((first_row + r_off, first_col + c_off))

    if not candidates:
        return None

    for row, col in candidates:
        right = normalize_text(sheet.cells(row, col + 1).value) if col + 1 <= last_col else ""
        left = normalize_text(sheet.cells(row, col - 1).value) if col - 1 >= first_col else ""
        if right == "min" or left == "min":
            return row, col
    return candidates[0]


def collect_numeric_pairs(
    sheet: Any,
    start_row: int,
    end_row: int,
    x_col: int,
    y_col: int,
) -> Tuple[List[int], Dict[int, float], Dict[int, float]]:
    x_values = to_1d(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_values = to_1d(sheet.range((start_row, y_col), (end_row, y_col)).value)

    rows: List[int] = []
    x_map: Dict[int, float] = {}
    y_map: Dict[int, float] = {}

    for idx, (x_value, y_value) in enumerate(zip(x_values, y_values)):
        abs_row = start_row + idx
        x_num = as_float(x_value)
        y_num = as_float(y_value)
        if x_num is None or y_num is None:
            continue
        rows.append(abs_row)
        x_map[abs_row] = x_num
        y_map[abs_row] = y_num

    return rows, x_map, y_map


def map_column_values(
    sheet: Any,
    start_row: int,
    end_row: int,
    col: int,
) -> Dict[int, Any]:
    values = to_1d(sheet.range((start_row, col), (end_row, col)).value)
    return {start_row + idx: value for idx, value in enumerate(values)}


def compute_growth_rate(values: Sequence[float]) -> Optional[float]:
    if len(values) < 2:
        return None
    first = values[0]
    last = values[-1]
    if first == 0:
        return None
    return (last - first) / abs(first)


def read_table_value(sheet: Any, row: int, col: int) -> Optional[float]:
    if row <= 0 or col <= 0:
        return None
    return as_float(sheet.cells(row, col).value)


def process_empirical_sheet(sheet: Any, meta: FileMeta, source_file: str) -> List[Dict[str, Any]]:
    rows_out: List[Dict[str, Any]] = []
    first_row, last_row, first_col, last_col = get_used_bounds(sheet)
    anchor = find_max_anchor(sheet, first_row, last_row, first_col, last_col)
    if anchor is None:
        print(f"  - empirical_candidates skipped: 'max' anchor not found in {sheet.name}")
        return rows_out

    anchor_row, anchor_col = anchor
    # Existing layout convention: penetration and sales columns sit left of the max anchor.
    penetration_col = anchor_col - 11
    sales_col = anchor_col - 7
    quarter_col = penetration_col - 1
    if penetration_col <= 0 or sales_col <= 0:
        print(f"  - empirical_candidates skipped: invalid anchor offsets in {sheet.name}")
        return rows_out

    data_start_row = anchor_row + 1
    if data_start_row > last_row:
        print(f"  - empirical_candidates skipped: no data rows under anchor in {sheet.name}")
        return rows_out

    data_rows, penetration_by_row, sales_by_row = collect_numeric_pairs(
        sheet=sheet,
        start_row=data_start_row,
        end_row=last_row,
        x_col=penetration_col,
        y_col=sales_col,
    )
    if not data_rows:
        print(f"  - empirical_candidates skipped: no numeric quarter data in {sheet.name}")
        return rows_out

    quarter_map = map_column_values(sheet, data_start_row, last_row, quarter_col) if quarter_col > 0 else {}

    max_n = min(MAX_QUARTERS, len(data_rows))
    helper_row = anchor_row
    helper_col = max(last_col + 2, anchor_col + 6)

    for n_quarters in range(1, max_n + 1):
        window_rows = data_rows[-n_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        set_formula2(sheet.cells(helper_row, helper_col), avg_formula)
        sheet.book.app.calculate()
        avg_penetration_pct = as_float(sheet.cells(helper_row, helper_col).value)
        if avg_penetration_pct is None:
            continue

        penetration_window = [penetration_by_row[row] for row in window_rows]
        sales_window = [sales_by_row[row] for row in window_rows]

        quarterly_sales = sum(sales_window)
        reported_sales = sales_window[-1]
        growth_rate_pct = compute_growth_rate(sales_window)
        sales_captured_in_db_pct = avg_penetration_pct
        last_quarter_used = quarter_map.get(end_row, end_row)

        table_row = anchor_row + n_quarters
        table_forecast = read_table_value(sheet, table_row, anchor_col - 1)
        table_actual = read_table_value(sheet, table_row, anchor_col - 2)
        table_max = read_table_value(sheet, table_row, anchor_col)
        table_min = read_table_value(sheet, table_row, anchor_col + 1)

        forecast_value = table_forecast
        if forecast_value is None and avg_penetration_pct != 0:
            forecast_value = reported_sales / avg_penetration_pct

        actual_value = table_actual if table_actual is not None else reported_sales

        min_penetration = min(penetration_window) if penetration_window else None
        max_penetration = max(penetration_window) if penetration_window else None
        fallback_max = (
            (reported_sales / min_penetration)
            if min_penetration is not None and min_penetration != 0
            else None
        )
        fallback_min = (
            (reported_sales / max_penetration)
            if max_penetration is not None and max_penetration != 0
            else None
        )

        forecast_max = table_max if table_max is not None else fallback_max
        forecast_min = table_min if table_min is not None else fallback_min
        range_width = safe_subtract(forecast_max, forecast_min)

        rows_out.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows_out


def process_regression_sheet(sheet: Any, meta: FileMeta, source_file: str) -> List[Dict[str, Any]]:
    rows_out: List[Dict[str, Any]] = []
    first_row, last_row, first_col, last_col = get_used_bounds(sheet)
    anchor = find_max_anchor(sheet, first_row, last_row, first_col, last_col)
    if anchor is None:
        print(f"  - regression_candidates skipped: 'max' anchor not found in {sheet.name}")
        return rows_out

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col <= 0 or x_col <= 0:
        print(f"  - regression_candidates skipped: invalid anchor offsets in {sheet.name}")
        return rows_out

    data_start_row = anchor_row + 1
    if data_start_row > last_row:
        print(f"  - regression_candidates skipped: no data rows under anchor in {sheet.name}")
        return rows_out

    data_rows, x_by_row, y_by_row = collect_numeric_pairs(
        sheet=sheet,
        start_row=data_start_row,
        end_row=last_row,
        x_col=x_col,
        y_col=y_col,
    )
    if not data_rows:
        print(f"  - regression_candidates skipped: no numeric quarter data in {sheet.name}")
        return rows_out

    max_n = min(MAX_QUARTERS, len(data_rows))
    helper_row = anchor_row
    helper_intercept_col = max(last_col + 2, anchor_col + 8)
    helper_slope_col = helper_intercept_col + 1

    previous_signature: Optional[Tuple[Any, ...]] = None
    for n_quarters in range(1, max_n + 1):
        window_rows = data_rows[-n_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2(sheet.cells(helper_row, helper_intercept_col), intercept_formula)
        set_formula2(sheet.cells(helper_row, helper_slope_col), slope_formula)
        sheet.book.app.calculate()

        intercept = as_float(sheet.cells(helper_row, helper_intercept_col).value)
        slope = as_float(sheet.cells(helper_row, helper_slope_col).value)
        latest_x = x_by_row[end_row]

        table_row = anchor_row + n_quarters
        table_forecast = read_table_value(sheet, table_row, anchor_col - 1)
        table_actual = read_table_value(sheet, table_row, anchor_col - 2)
        table_max = read_table_value(sheet, table_row, anchor_col)
        table_min = read_table_value(sheet, table_row, anchor_col + 1)

        fallback_forecast = None
        if intercept is not None and slope is not None:
            fallback_forecast = intercept + slope * latest_x

        forecast_value = table_forecast if table_forecast is not None else fallback_forecast
        y_window = [y_by_row[row] for row in window_rows]
        forecast_max = table_max if table_max is not None else max(y_window)
        forecast_min = table_min if table_min is not None else min(y_window)
        range_width = safe_subtract(forecast_max, forecast_min)
        actual_value = table_actual if table_actual is not None else ""

        current_signature = (
            n_quarters,
            round(forecast_value, 8) if forecast_value is not None else None,
            round(intercept, 8) if intercept is not None else None,
            round(slope, 8) if slope is not None else None,
            round(forecast_max, 8) if forecast_max is not None else None,
            round(forecast_min, 8) if forecast_min is not None else None,
        )

        # Guard against the known duplicate trailing row pattern in regression output.
        if previous_signature is not None and current_signature[1:] == previous_signature[1:]:
            continue

        rows_out.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
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
        previous_signature = current_signature

    return rows_out


def write_rows_to_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(list(columns))

    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 52)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    write_rows_to_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def should_skip_file(file_path: Path) -> Optional[str]:
    if file_path.name.startswith("~"):
        return "temporary workbook"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if SKIP_OUTPUT_PATTERN.search(file_path.name):
        return "previous PARAM output"
    return None


def main() -> None:
    input_root = Path(input_dir).expanduser().resolve()
    output_root = Path(output_dir).expanduser().resolve()

    if not input_root.exists() or not input_root.is_dir():
        raise ValueError(f"input_dir does not exist or is not a folder: {input_root}")
    output_root.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(input_root, output_root)
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

        for file_path in sorted(input_root.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue

            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            print(f"Processing {file_path.name}")
            wb = None
            try:
                metadata = parse_file_metadata(file_path)
                wb = app.books.open(str(file_path), update_links=False)
                processed_files += 1

                try:
                    empirical_sheet = wb.sheets["Empirical Model"]
                    empirical_rows.extend(
                        process_empirical_sheet(empirical_sheet, metadata, file_path.name)
                    )
                except Exception as exc:
                    print(f"  - empirical_candidates skipped for {file_path.name}: {exc}")

                try:
                    regression_sheet = wb.sheets["Regression Model"]
                    regression_rows.extend(
                        process_regression_sheet(regression_sheet, metadata, file_path.name)
                    )
                except Exception as exc:
                    print(f"  - regression_candidates skipped for {file_path.name}: {exc}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to process ({exc})")
            finally:
                close_workbook_safely(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
