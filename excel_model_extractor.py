from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("./input")
output_dir = Path("./output")

N_QUARTERS = 10
EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"

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

DAY_BY_WINDOW = {"early": 5, "mid": 15, "late": 25}
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


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def as_scalar_list(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            try:
                return float(text[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    num = to_float(value)
    if num is None:
        return None
    return int(round(num))


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_month(month_token: str) -> int:
    key = month_token.strip().lower()
    if key in MONTH_LOOKUP:
        return MONTH_LOOKUP[key]
    short_key = key[:3]
    if short_key in MONTH_LOOKUP:
        return MONTH_LOOKUP[short_key]
    raise ValueError(f"Unrecognized month token: {month_token}")


def parse_filename_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = "UNKNOWN"
    period_token = ""

    if len(parts) >= 3:
        ticker = re.sub(r"\s+", "", parts[1]).upper() or "UNKNOWN"
        period_token = parts[2]
    else:
        match = re.search(
            r"([A-Za-z0-9]+)\s*-\s*((?:Early|Mid|Late)[A-Za-z]{3,9}\d{4})",
            stem,
            flags=re.IGNORECASE,
        )
        if match:
            ticker = match.group(1).upper()
            period_token = match.group(2)

    period_token = re.sub(r"_?send$", "", period_token, flags=re.IGNORECASE)
    period_token = re.sub(r"[^A-Za-z0-9]", "", period_token)

    model_period = period_token or "Unknown_Period"
    model_date = ""

    period_match = re.match(
        r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if period_match:
        window_raw, month_raw, year_raw = period_match.groups()
        window = window_raw.title()
        year_int = int(year_raw)
        month_int = parse_month(month_raw)
        month_abbrev = datetime(year_int, month_int, 1).strftime("%b")
        model_period = f"{window}{month_abbrev}_{year_int}"
        model_day = DAY_BY_WINDOW[window.lower()]
        model_date = date(year_int, month_int, model_day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_folder: Path, out_folder: Path) -> Path:
    base = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_folder / f"{base}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def safe_close_workbook(wb: Any) -> None:
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
        wb.close()
    except Exception:
        pass


def set_formula2_r1c1(cell: Any, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.FormulaR1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def get_sheet_by_name(wb: Any, sheet_name: str) -> Optional[Any]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def get_used_bounds(sheet: Any) -> Tuple[int, int, int, int]:
    try:
        used = sheet.api.UsedRange
        first_row = int(used.Row)
        first_col = int(used.Column)
        last_row = first_row + int(used.Rows.Count) - 1
        last_col = first_col + int(used.Columns.Count) - 1
        return first_row, first_col, last_row, last_col
    except Exception:
        used = sheet.used_range
        first_row = int(used.row)
        first_col = int(used.column)
        values = ensure_2d(used.value)
        row_count = len(values)
        col_count = len(values[0]) if row_count else 1
        last_row = first_row + max(row_count - 1, 0)
        last_col = first_col + max(col_count - 1, 0)
        return first_row, first_col, last_row, last_col


def find_anchor_max(sheet: Any) -> Optional[Tuple[int, int]]:
    try:
        found = sheet.api.Cells.Find(What="max", LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    first_row, first_col, last_row, last_col = get_used_bounds(sheet)
    grid = ensure_2d(sheet.range((first_row, first_col), (last_row, last_col)).value)
    for r_offset, row in enumerate(grid):
        for c_offset, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return first_row + r_offset, first_col + c_offset
    return None


def build_header_index(
    sheet: Any, header_row: int, first_col: int, last_col: int
) -> List[Tuple[str, int]]:
    row_values = sheet.range((header_row, first_col), (header_row, last_col)).value
    if row_values is None:
        return []
    if not isinstance(row_values, list):
        row_values = [row_values]

    headers: List[Tuple[str, int]] = []
    for idx, value in enumerate(row_values):
        normalized = normalize_header(value)
        if normalized:
            headers.append((normalized, first_col + idx))
    return headers


def phrase_in_header(header: str, phrase: str) -> bool:
    header_tokens = header.split()
    phrase_tokens = phrase.split()
    if not phrase_tokens:
        return False
    if len(phrase_tokens) == 1:
        return phrase_tokens[0] in header_tokens

    span = len(phrase_tokens)
    for i in range(len(header_tokens) - span + 1):
        if header_tokens[i : i + span] == phrase_tokens:
            return True
    return False


def find_col_by_keywords(
    headers: Sequence[Tuple[str, int]], phrases: Sequence[str]
) -> Optional[int]:
    normalized_phrases = [normalize_header(phrase) for phrase in phrases if phrase]
    for phrase in normalized_phrases:
        for header, col in headers:
            if phrase_in_header(header, phrase):
                return col
    return None


def read_table_value(
    table_grid: List[List[Any]],
    absolute_row: int,
    absolute_col: Optional[int],
    table_start_row: int,
    first_col: int,
) -> Any:
    if absolute_col is None:
        return None
    row_idx = absolute_row - table_start_row
    col_idx = absolute_col - first_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(table_grid):
        return None
    row = table_grid[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


def find_last_numeric_row(
    sheet: Any,
    column: int,
    first_row: int,
    last_row: int,
) -> Optional[int]:
    if last_row < first_row:
        return None
    values = as_scalar_list(sheet.range((first_row, column), (last_row, column)).value)
    for idx in range(len(values) - 1, -1, -1):
        if to_float(values[idx]) is not None:
            return first_row + idx
    return None


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def comparable_value(value: Any) -> Any:
    num = to_float(value)
    if num is not None:
        return round(num, 10)
    if value is None:
        return None
    return str(value).strip()


def extract_empirical_rows(
    wb: Any,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_by_name(wb, EMPIRICAL_SHEET)
    if sheet is None:
        return rows

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor
    first_row, first_col, last_row, last_col = get_used_bounds(sheet)
    headers = build_header_index(sheet, anchor_row, first_col, last_col)

    num_quarters_col = find_col_by_keywords(
        headers, ["num quarters used", "quarters used", "num quarters", "n quarters"]
    )
    last_quarter_col = find_col_by_keywords(
        headers, ["last quarter used", "last quarter", "quarter used"]
    )
    forecast_col = find_col_by_keywords(
        headers,
        [
            "estimated total sold",
            "est total sold",
            "total forecast",
            "tot fcst",
            "forecast",
        ],
    )
    reported_col = find_col_by_keywords(
        headers, ["reported sales", "actual sales", "reported", "actual"]
    )
    forecast_min_col = find_col_by_keywords(headers, ["forecast min", "min", "minimum"])
    avg_penetration_col = find_col_by_keywords(
        headers, ["avg penetration pct", "average penetration", "avg penetration", "penetration"]
    )
    quarterly_sales_col = find_col_by_keywords(
        headers, ["quarterly sales", "quarter sales", "qtr sales"]
    )
    growth_rate_col = find_col_by_keywords(headers, ["growth rate pct", "growth rate", "growth"])
    captured_col = find_col_by_keywords(
        headers,
        [
            "sales captured in db pct",
            "sales captured in db",
            "captured in db",
            "captured",
        ],
    )

    forecast_max_col = anchor_col
    if forecast_min_col is None:
        forecast_min_col = anchor_col + 1

    table_start_row = anchor_row + 1
    table_end_row = min(last_row, table_start_row + N_QUARTERS - 1)
    if table_end_row < table_start_row:
        return rows

    table_grid = ensure_2d(
        sheet.range((table_start_row, first_col), (table_end_row, last_col)).value
    )

    scratch_avg_col = last_col + 2
    formula_rows: List[int] = []
    pending: List[Dict[str, Any]] = []

    for i in range(N_QUARTERS):
        row = table_start_row + i
        if row > table_end_row:
            break

        num_quarters_raw = read_table_value(table_grid, row, num_quarters_col, table_start_row, first_col)
        num_quarters_used = to_int(num_quarters_raw) or (i + 1)
        last_quarter_used = read_table_value(table_grid, row, last_quarter_col, table_start_row, first_col)
        forecast_value = read_table_value(table_grid, row, forecast_col, table_start_row, first_col)
        forecast_max = read_table_value(table_grid, row, forecast_max_col, table_start_row, first_col)
        forecast_min = read_table_value(table_grid, row, forecast_min_col, table_start_row, first_col)
        quarterly_sales = read_table_value(table_grid, row, quarterly_sales_col, table_start_row, first_col)
        reported_sales = read_table_value(table_grid, row, reported_col, table_start_row, first_col)
        growth_rate_pct = read_table_value(table_grid, row, growth_rate_col, table_start_row, first_col)
        captured_pct = read_table_value(table_grid, row, captured_col, table_start_row, first_col)
        avg_penetration_pct = read_table_value(
            table_grid, row, avg_penetration_col, table_start_row, first_col
        )

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and quarterly_sales is None
            and reported_sales is None
            and avg_penetration_pct is None
            and last_quarter_used is None
        ):
            continue

        rolling_start = max(table_start_row, row - num_quarters_used + 1)
        avg_formula: Optional[str] = None
        if avg_penetration_col is not None:
            avg_formula = (
                f'=IFERROR(AVERAGE(R{rolling_start}C{avg_penetration_col}:R{row}C{avg_penetration_col}),"")'
            )
        elif quarterly_sales_col is not None and reported_col is not None:
            avg_formula = (
                f'=IFERROR(SUM(R{rolling_start}C{quarterly_sales_col}:R{row}C{quarterly_sales_col})/'
                f'SUM(R{rolling_start}C{reported_col}:R{row}C{reported_col}),"")'
            )

        if avg_formula:
            set_formula2_r1c1(sheet.cells(row, scratch_avg_col), avg_formula)
            formula_rows.append(row)

        pending.append(
            {
                "row": row,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": captured_pct,
            }
        )

    if formula_rows:
        wb.app.calculate()

    formula_row_set = set(formula_rows)
    for record in pending:
        row = record["row"]
        avg_penetration_pct = record["avg_penetration_pct"]
        if row in formula_row_set:
            avg_penetration_pct = sheet.cells(row, scratch_avg_col).value

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": record["num_quarters_used"],
                "last_quarter_used": record["last_quarter_used"],
                "forecast_value": record["forecast_value"],
                "actual_value": record["reported_sales"],
                "forecast_max": record["forecast_max"],
                "forecast_min": record["forecast_min"],
                "range_width": calc_range_width(record["forecast_max"], record["forecast_min"]),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": record["quarterly_sales"],
                "reported_sales": record["reported_sales"],
                "growth_rate_pct": record["growth_rate_pct"],
                "sales_captured_in_db_pct": record["sales_captured_in_db_pct"],
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: Any,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_by_name(wb, REGRESSION_SHEET)
    if sheet is None:
        return rows

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor
    first_row, first_col, last_row, last_col = get_used_bounds(sheet)
    headers = build_header_index(sheet, anchor_row, first_col, last_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = find_col_by_keywords(
        headers, ["num quarters used", "quarters used", "num quarters", "n quarters"]
    )
    forecast_col = find_col_by_keywords(
        headers,
        [
            "tot fcst w o sa",
            "tot fcst wout sa",
            "tot fcst without sa",
            "total forecast without sa",
            "forecast without sa",
            "forecast",
        ],
    )
    actual_col = find_col_by_keywords(headers, ["actual sales", "reported sales", "actual", "reported"])
    forecast_min_col = find_col_by_keywords(headers, ["forecast min", "min", "minimum"])

    forecast_max_col = anchor_col
    if forecast_min_col is None:
        forecast_min_col = anchor_col + 1

    table_start_row = anchor_row + 1
    table_end_row = min(last_row, table_start_row + N_QUARTERS - 1)
    if table_end_row < table_start_row:
        return rows

    data_end_row = find_last_numeric_row(sheet, y_col, first_row, anchor_row - 1)
    if data_end_row is None:
        data_end_row = anchor_row - 1

    table_grid = ensure_2d(
        sheet.range((table_start_row, first_col), (table_end_row, last_col)).value
    )

    scratch_intercept_col = last_col + 2
    scratch_slope_col = last_col + 3
    formula_rows: List[int] = []

    for i in range(N_QUARTERS):
        row = table_start_row + i
        if row > table_end_row:
            break

        num_quarters_raw = read_table_value(table_grid, row, num_quarters_col, table_start_row, first_col)
        num_quarters_used = to_int(num_quarters_raw) or (i + 1)
        if num_quarters_used < 2:
            num_quarters_used = 2

        data_start_row = data_end_row - num_quarters_used + 1
        if data_start_row < first_row:
            data_start_row = first_row

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        set_formula2_r1c1(sheet.cells(row, scratch_intercept_col), intercept_formula)
        set_formula2_r1c1(sheet.cells(row, scratch_slope_col), slope_formula)
        formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    previous_key: Optional[Tuple[Any, ...]] = None
    for i in range(N_QUARTERS):
        row = table_start_row + i
        if row > table_end_row:
            break

        num_quarters_raw = read_table_value(table_grid, row, num_quarters_col, table_start_row, first_col)
        num_quarters_used = to_int(num_quarters_raw) or (i + 1)
        forecast_value = read_table_value(table_grid, row, forecast_col, table_start_row, first_col)
        actual_value = read_table_value(table_grid, row, actual_col, table_start_row, first_col)
        forecast_max = read_table_value(table_grid, row, forecast_max_col, table_start_row, first_col)
        forecast_min = read_table_value(table_grid, row, forecast_min_col, table_start_row, first_col)
        intercept = sheet.cells(row, scratch_intercept_col).value
        slope = sheet.cells(row, scratch_slope_col).value

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept in (None, "")
            and slope in (None, "")
        ):
            continue

        current_key = (
            num_quarters_used,
            comparable_value(intercept),
            comparable_value(slope),
            comparable_value(forecast_value),
            comparable_value(forecast_max),
            comparable_value(forecast_min),
        )
        if previous_key is not None and current_key == previous_key:
            continue
        previous_key = current_key

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def iter_source_workbooks(source_dir: Path) -> Iterable[Path]:
    param_output_pattern = re.compile(
        rf"^{re.escape(source_dir.name)}_PARAM(?:\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )
    for path in sorted(source_dir.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue
        if param_output_pattern.match(path.name):
            print(f"Skipped {path.name}: prior PARAM output file")
            continue
        yield path


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    records: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)

    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)

    for row_idx, record in enumerate(records, start=2):
        for col_idx, col_name in enumerate(headers, start=1):
            ws.cell(row=row_idx, column=col_idx, value=record.get(col_name))

    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(headers))
    last_row = max(1, len(records) + 1)
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for col_idx, col_name in enumerate(headers, start=1):
        max_len = len(col_name)
        for record in records:
            value = record.get(col_name)
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir.resolve(), output_dir.resolve())

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in iter_source_workbooks(input_dir):
            print(f"Processing {file_path.name}")
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: open failed ({exc})")
                continue

            try:
                meta = parse_filename_metadata(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing failed ({exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
