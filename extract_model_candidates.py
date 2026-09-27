from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------- User inputs -----------------------------
input_dir = Path("input")
output_dir = Path("output")
# ----------------------------------------------------------------------

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
PERIOD_REGEX = re.compile(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", re.IGNORECASE)
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
PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value in (None, "") or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def calc_range_width(forecast_max: Any, forecast_min: Any) -> Optional[float]:
    max_val = to_float(forecast_max)
    min_val = to_float(forecast_min)
    if max_val is None or min_val is None:
        return None
    return max_val - min_val


def parse_month(token: str) -> Optional[int]:
    return MONTH_LOOKUP.get(token.strip().lower())


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""
    period_source = parts[2] if len(parts) >= 3 else stem

    model_period = ""
    model_date = ""

    match = PERIOD_REGEX.search(period_source)
    if match:
        period_tag = match.group(1).title()
        month_token = match.group(2)
        year = int(match.group(3))
        month = parse_month(month_token)
        if month:
            month_abbrev = date(year, month, 1).strftime("%b")
            model_period = f"{period_tag}{month_abbrev}_{year}"
            day = PERIOD_DAY[period_tag]
            model_date = date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or stem
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{in_dir.resolve().name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base}.{suffix}.xlsx"
        suffix += 1
    return candidate


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def get_used_bounds(sheet: xw.Sheet) -> Tuple[int, int, int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    last_row = used.last_cell.row
    last_col = used.last_cell.column
    return start_row, start_col, last_row, last_col


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return None

    start_row = used.row
    start_col = used.column
    target = normalize_text(anchor_text)

    for row_idx, row_values in enumerate(matrix):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                return start_row + row_idx, start_col + col_idx
    return None


def build_header_index(
    sheet: xw.Sheet, candidate_rows: Sequence[int], start_col: int, end_col: int
) -> Dict[str, int]:
    index: Dict[str, int] = {}
    for row in candidate_rows:
        if row < 1:
            continue
        row_values = sheet.range((row, start_col), (row, end_col)).value
        if not isinstance(row_values, list):
            row_values = [row_values]
        if row_values and isinstance(row_values[0], list):
            row_values = row_values[0]
        for offset, value in enumerate(row_values):
            key = normalize_text(value)
            if key and key not in index:
                index[key] = start_col + offset
    return index


def find_col(index: Dict[str, int], keywords: Sequence[str]) -> Optional[int]:
    normalized_keywords = [normalize_text(keyword) for keyword in keywords if keyword]

    for keyword in normalized_keywords:
        if keyword in index:
            return index[keyword]

    for header, col in index.items():
        for keyword in normalized_keywords:
            if keyword and keyword in header:
                return col
    return None


def read_col_values(sheet: xw.Sheet, col: Optional[int], start_row: int, n_rows: int) -> List[Any]:
    if col is None:
        return [None] * n_rows

    values = sheet.range((start_row, col), (start_row + n_rows - 1, col)).value
    if n_rows == 1:
        return [values]

    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flat = [row[0] for row in values]
        else:
            flat = list(values)
    else:
        flat = [values]

    if len(flat) < n_rows:
        flat.extend([None] * (n_rows - len(flat)))
    return flat[:n_rows]


def set_formula2(target: xw.Range, formula: str) -> None:
    try:
        target.formula2 = formula
    except Exception:
        target.formula = formula


def close_workbook_without_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for close_call in (
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ):
        try:
            close_call()
            return
        except Exception:
            continue


def process_empirical_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if "Empirical Model" not in {sheet.name for sheet in wb.sheets}:
        print(f"  skipped empirical extraction ({source_file}): missing 'Empirical Model' sheet")
        return rows

    sheet = wb.sheets["Empirical Model"]
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  skipped empirical extraction ({source_file}): could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    start_row, start_col, _, end_col = get_used_bounds(sheet)
    header_index = build_header_index(sheet, [anchor_row, anchor_row - 1], start_col, end_col)

    num_quarters_col = find_col(header_index, ["num quarters used", "quarters used", "n quarters"])
    last_quarter_col = find_col(header_index, ["last quarter used", "last quarter", "last qtr"])
    forecast_value_col = find_col(
        header_index, ["estimated total sold", "forecast value", "tot fcst", "total forecast"]
    )
    reported_sales_col = find_col(header_index, ["reported sales", "actual sales", "actual value"])
    quarterly_sales_col = find_col(header_index, ["quarterly sales", "quarter sales"])
    growth_rate_col = find_col(header_index, ["growth rate pct", "growth rate"])
    sales_captured_col = find_col(
        header_index, ["sales captured in db pct", "sales captured in db", "captured in db"]
    )
    penetration_source_col = find_col(
        header_index, ["penetration pct", "avg penetration pct", "penetration", "sales captured in db pct"]
    )

    forecast_max_col = anchor_col
    forecast_min_col = find_col(header_index, ["min"])
    if forecast_min_col is None and (anchor_col + 1) <= end_col:
        forecast_min_col = anchor_col + 1

    data_start_row = anchor_row + 1
    n_rows = N_QUARTERS
    temp_col = end_col + 3
    avg_penetration_vals = [None] * n_rows

    if penetration_source_col is not None:
        for idx in range(n_rows):
            row_num = data_start_row + idx
            n_quarters = idx + 1
            window_start = max(data_start_row, row_num - n_quarters + 1)
            formula = (
                f"=AVERAGE(R{window_start}C{penetration_source_col}:"
                f"R{row_num}C{penetration_source_col})"
            )
            set_formula2(sheet.range((row_num, temp_col)), formula)

        wb.app.calculate()
        avg_penetration_vals = read_col_values(sheet, temp_col, data_start_row, n_rows)
        sheet.range((data_start_row, temp_col), (data_start_row + n_rows - 1, temp_col)).clear_contents()

    num_quarters_vals = read_col_values(sheet, num_quarters_col, data_start_row, n_rows)
    last_quarter_vals = read_col_values(sheet, last_quarter_col, data_start_row, n_rows)
    forecast_vals = read_col_values(sheet, forecast_value_col, data_start_row, n_rows)
    reported_sales_vals = read_col_values(sheet, reported_sales_col, data_start_row, n_rows)
    forecast_max_vals = read_col_values(sheet, forecast_max_col, data_start_row, n_rows)
    forecast_min_vals = read_col_values(sheet, forecast_min_col, data_start_row, n_rows)
    quarterly_sales_vals = read_col_values(sheet, quarterly_sales_col, data_start_row, n_rows)
    growth_rate_vals = read_col_values(sheet, growth_rate_col, data_start_row, n_rows)
    sales_captured_vals = read_col_values(sheet, sales_captured_col, data_start_row, n_rows)

    for idx in range(n_rows):
        num_quarters = num_quarters_vals[idx]
        if num_quarters in (None, ""):
            num_quarters = idx + 1

        avg_penetration = avg_penetration_vals[idx]
        forecast_max = forecast_max_vals[idx]
        forecast_min = forecast_min_vals[idx]
        reported_sales = reported_sales_vals[idx]
        forecast_value = forecast_vals[idx]

        row_data: Dict[str, Any] = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters,
            "last_quarter_used": last_quarter_vals[idx],
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales_vals[idx],
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_vals[idx],
            "sales_captured_in_db_pct": sales_captured_vals[idx],
            "source_file": source_file,
        }

        has_signal = any(
            row_data[key] not in (None, "")
            for key in (
                "forecast_value",
                "actual_value",
                "forecast_max",
                "forecast_min",
                "avg_penetration_pct",
                "quarterly_sales",
            )
        )
        if has_signal:
            rows.append(row_data)

    return rows


def rows_are_equivalent(current: Dict[str, Any], previous: Dict[str, Any], fields: Sequence[str]) -> bool:
    for field in fields:
        cur = current.get(field)
        prev = previous.get(field)
        cur_float = to_float(cur)
        prev_float = to_float(prev)
        if cur_float is not None and prev_float is not None:
            if abs(cur_float - prev_float) > 1e-9:
                return False
            continue
        if cur != prev:
            return False
    return True


def process_regression_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if "Regression Model" not in {sheet.name for sheet in wb.sheets}:
        print(f"  skipped regression extraction ({source_file}): missing 'Regression Model' sheet")
        return rows

    sheet = wb.sheets["Regression Model"]
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  skipped regression extraction ({source_file}): could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    start_row, start_col, _, end_col = get_used_bounds(sheet)
    header_index = build_header_index(sheet, [anchor_row, anchor_row - 1], start_col, end_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = find_col(header_index, ["num quarters used", "quarters used", "n quarters"])
    forecast_value_col = find_col(
        header_index,
        [
            "tot fcst w o sa",
            "tot fcst w/o sa",
            "tot fcst wo sa",
            "forecast total without sa",
            "forecast value",
        ],
    )
    actual_value_col = find_col(header_index, ["actual value", "actual sales", "reported sales"])

    forecast_max_col = anchor_col
    forecast_min_col = find_col(header_index, ["min"])
    if forecast_min_col is None and (anchor_col + 1) <= end_col:
        forecast_min_col = anchor_col + 1

    data_start_row = anchor_row + 1
    n_rows = N_QUARTERS
    temp_intercept_col = end_col + 3
    temp_slope_col = end_col + 4

    for idx in range(n_rows):
        row_num = data_start_row + idx
        n_quarters = idx + 1
        window_start = max(data_start_row, row_num - n_quarters + 1)
        intercept_formula = (
            f"=INTERCEPT(R{window_start}C{y_col}:R{row_num}C{y_col},"
            f"R{window_start}C{x_col}:R{row_num}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{window_start}C{y_col}:R{row_num}C{y_col},"
            f"R{window_start}C{x_col}:R{row_num}C{x_col})"
        )
        set_formula2(sheet.range((row_num, temp_intercept_col)), intercept_formula)
        set_formula2(sheet.range((row_num, temp_slope_col)), slope_formula)

    wb.app.calculate()
    intercept_vals = read_col_values(sheet, temp_intercept_col, data_start_row, n_rows)
    slope_vals = read_col_values(sheet, temp_slope_col, data_start_row, n_rows)
    sheet.range((data_start_row, temp_intercept_col), (data_start_row + n_rows - 1, temp_slope_col)).clear_contents()

    num_quarters_vals = read_col_values(sheet, num_quarters_col, data_start_row, n_rows)
    forecast_vals = read_col_values(sheet, forecast_value_col, data_start_row, n_rows)
    actual_vals = read_col_values(sheet, actual_value_col, data_start_row, n_rows)
    forecast_max_vals = read_col_values(sheet, forecast_max_col, data_start_row, n_rows)
    forecast_min_vals = read_col_values(sheet, forecast_min_col, data_start_row, n_rows)

    for idx in range(n_rows):
        num_quarters = num_quarters_vals[idx]
        if num_quarters in (None, ""):
            num_quarters = idx + 1

        forecast_max = forecast_max_vals[idx]
        forecast_min = forecast_min_vals[idx]
        row_data: Dict[str, Any] = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_vals[idx],
            "actual_value": actual_vals[idx],
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "intercept": intercept_vals[idx],
            "slope": slope_vals[idx],
            "source_file": source_file,
        }

        has_signal = any(
            row_data[key] not in (None, "")
            for key in ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
        )
        if not has_signal:
            continue

        if idx == n_rows - 1 and rows:
            if rows_are_equivalent(
                row_data,
                rows[-1],
                ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope"),
            ):
                continue

        rows.append(row_data)

    return rows


def write_sheet(ws: Any, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, ws.max_row)}"

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for cell in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=idx, max_col=idx):
            value = cell[0].value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = max(12, min(max_len + 2, 48))


def write_output_workbook(
    output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    wb.save(output_path)


def list_input_files(in_dir: Path) -> List[Path]:
    if not in_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {in_dir}")
    return sorted(path for path in in_dir.iterdir())


def run() -> None:
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_file_count = 0

    files = list_input_files(input_dir)
    output_path = build_output_path(input_dir, output_dir)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        app.screen_updating = False
    except Exception:
        pass

    try:
        for file_path in files:
            if not file_path.is_file():
                print(f"skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not an .xlsx file)")
                continue

            print(f"processing: {file_path.name}")
            label = parse_file_label(file_path)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, label, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, label, file_path.name))
                processed_file_count += 1
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"files processed: {processed_file_count}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
