from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw

try:
    import openpyxl
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None
    Font = None
    get_column_letter = None


# ---------- User-configurable paths ----------
input_dir = "/path/to/input"
output_dir = "/path/to/output"


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


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def parse_model_metadata(filename: str) -> Tuple[Optional[Dict[str, str]], Optional[str]]:
    stem = Path(filename).stem
    parts = [part.strip() for part in stem.split(" - ")]

    if len(parts) < 3:
        return None, "filename format does not match expected pattern"

    ticker = parts[1].strip().upper()
    period_token = parts[2].split("_", 1)[0].strip()

    match = re.search(r"(?i)(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token)
    if not match:
        return None, "period segment does not match Early/Mid/LateMonthYear pattern"

    period_bucket = match.group(1).title()
    month_token = match.group(2)
    year = int(match.group(3))

    month_value = None
    for fmt in ("%b", "%B"):
        try:
            month_value = datetime.strptime(month_token, fmt).month
            break
        except ValueError:
            continue
    if month_value is None:
        try:
            month_value = datetime.strptime(month_token[:3], "%b").month
        except ValueError:
            return None, f"unable to parse month token '{month_token}'"

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_map[period_bucket]
    model_date = date(year, month_value, day).isoformat()
    month_short = datetime(year, month_value, 1).strftime("%b")
    model_period = f"{period_bucket}{month_short}_{year}"
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }, None


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def to_2d(values: Any, rows_hint: Optional[int] = None, cols_hint: Optional[int] = None) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        if rows_hint and rows_hint > 1 and (cols_hint == 1 or cols_hint is None):
            return [[item] for item in values]
        return [values]
    return [[values]]


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any, default: Optional[int] = None) -> Optional[int]:
    number = to_number(value)
    if number is None:
        return default
    return int(round(number))


def find_anchor(sheet: xw.main.Sheet, label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    used_rows = used.rows.count
    used_cols = used.columns.count
    values = to_2d(used.value, used_rows, used_cols)
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == label:
                return used.row + r_idx, used.column + c_idx
    return None


def get_header_map(
    sheet: xw.main.Sheet,
    header_row: int,
    anchor_col: int,
    window: int = 30,
) -> Dict[str, int]:
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window
    row_values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if row_values is None:
        row_values = []
    if not isinstance(row_values, list):
        row_values = [row_values]

    header_map: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        key = normalize_header(value)
        if key:
            header_map[key] = start_col + idx

    # Ensure the anchor column is always available as "max".
    header_map.setdefault("max", anchor_col)
    return header_map


def resolve_col(
    header_map: Dict[str, int],
    aliases: Sequence[str],
    fallback_col: Optional[int] = None,
) -> Optional[int]:
    normalized_aliases = [normalize_header(alias) for alias in aliases]
    for alias in normalized_aliases:
        for header_key, col in header_map.items():
            if alias == header_key or alias in header_key:
                return col
    if fallback_col is not None and fallback_col >= 1:
        return fallback_col
    return None


def read_row_block(
    sheet: xw.main.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> List[List[Any]]:
    if start_col < 1:
        start_col = 1
    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    expected_rows = end_row - start_row + 1
    expected_cols = end_col - start_col + 1
    rows = to_2d(values, expected_rows, expected_cols)

    while len(rows) < expected_rows:
        rows.append([None] * expected_cols)
    for idx, row in enumerate(rows):
        if not isinstance(row, list):
            rows[idx] = [row] + [None] * (expected_cols - 1)
        elif len(row) < expected_cols:
            rows[idx] = row + [None] * (expected_cols - len(row))
    return rows


def block_value(block: List[List[Any]], row_idx: int, col: Optional[int], min_col: int) -> Any:
    if col is None:
        return None
    col_idx = col - min_col
    if row_idx < 0 or row_idx >= len(block):
        return None
    row = block[row_idx]
    if col_idx < 0 or col_idx >= len(row):
        return None
    return row[col_idx]


def set_formula2_r1c1(cell: xw.main.Range, formula_r1c1: str) -> None:
    # .formula2 first (as requested), then API fallback for engines that require Formula2R1C1.
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        cell.api.FormulaR1C1 = formula_r1c1


def close_source_workbook(wb: Optional[xw.main.Book]) -> None:
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
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    # Last-resort close attempt (display alerts are disabled on the app).
    try:
        wb.close()
    except Exception:
        pass


def number_or_blank(value: Any) -> Any:
    number = to_number(value)
    if number is None:
        return ""
    if float(number).is_integer():
        return int(number)
    return number


def process_empirical_sheet(
    wb: xw.main.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped Empirical Model in {source_file}: sheet not found")
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"Skipped Empirical Model in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = get_header_map(sheet, anchor_row, anchor_col)

    num_quarters_col = resolve_col(
        header_map,
        ["num_quarters_used", "num quarters used", "n quarters", "quarters used"],
        anchor_col - 10,
    )
    last_quarter_col = resolve_col(
        header_map,
        ["last_quarter_used", "last quarter used", "last quarter", "last qtr"],
        anchor_col - 9,
    )
    forecast_value_col = resolve_col(
        header_map,
        ["estimated total sold", "est total sold", "forecast value", "total forecast"],
        anchor_col - 8,
    )
    forecast_max_col = resolve_col(header_map, ["max"], anchor_col)
    forecast_min_col = resolve_col(header_map, ["min"], anchor_col + 1)
    quarterly_sales_col = resolve_col(
        header_map,
        ["quarterly_sales", "quarterly sales", "quarter sales", "q sales"],
        anchor_col - 6,
    )
    reported_sales_col = resolve_col(
        header_map,
        ["reported_sales", "reported sales", "actual sales"],
        anchor_col - 5,
    )
    avg_penetration_col = resolve_col(
        header_map,
        ["avg_penetration_pct", "average penetration", "avg penetration", "penetration pct"],
        anchor_col - 4,
    )
    growth_rate_col = resolve_col(
        header_map,
        ["growth_rate_pct", "growth rate", "growth pct"],
        anchor_col - 3,
    )
    captured_db_col = resolve_col(
        header_map,
        ["sales_captured_in_db_pct", "sales captured in db", "captured in db"],
        anchor_col - 2,
    )

    data_start_row = anchor_row + 1
    data_end_row = anchor_row + 10  # n_quarters = 10
    read_columns = [
        col
        for col in [
            num_quarters_col,
            last_quarter_col,
            forecast_value_col,
            forecast_max_col,
            forecast_min_col,
            quarterly_sales_col,
            reported_sales_col,
            avg_penetration_col,
            growth_rate_col,
            captured_db_col,
        ]
        if col is not None and col >= 1
    ]
    if not read_columns:
        print(f"Skipped Empirical Model in {source_file}: no readable columns resolved")
        return []

    min_col = min(read_columns)
    max_col = max(read_columns)
    block = read_row_block(sheet, data_start_row, data_end_row, min_col, max_col)

    helper_col = max(sheet.used_range.last_cell.column, max_col) + 2
    helper_rows: List[int] = []

    for row_idx in range(10):
        absolute_row = data_start_row + row_idx
        n_val = block_value(block, row_idx, num_quarters_col, min_col)
        n_quarters = max(1, min(10, to_int(n_val, row_idx + 1) or (row_idx + 1)))
        avg_val = block_value(block, row_idx, avg_penetration_col, min_col)

        # Keep formula writes temporary; workbook is never saved.
        if (
            to_number(avg_val) is None
            and quarterly_sales_col is not None
            and reported_sales_col is not None
            and quarterly_sales_col >= 1
            and reported_sales_col >= 1
        ):
            start_row = max(data_start_row, absolute_row - n_quarters + 1)
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{quarterly_sales_col}:R{absolute_row}C{quarterly_sales_col}'
                f"/R{start_row}C{reported_sales_col}:R{absolute_row}C{reported_sales_col}),\"\")"
            )
            set_formula2_r1c1(sheet.cells(absolute_row, helper_col), formula)
            helper_rows.append(absolute_row)

    helper_avg_values: Dict[int, Any] = {}
    if helper_rows:
        wb.app.calculate()
        first_helper_row = min(helper_rows)
        last_helper_row = max(helper_rows)
        helper_data = sheet.range((first_helper_row, helper_col), (last_helper_row, helper_col)).value
        helper_block = to_2d(helper_data, last_helper_row - first_helper_row + 1, 1)
        for offset, abs_row in enumerate(range(first_helper_row, last_helper_row + 1)):
            value = helper_block[offset][0] if offset < len(helper_block) else None
            helper_avg_values[abs_row] = value

    rows: List[Dict[str, Any]] = []
    for row_idx in range(10):
        absolute_row = data_start_row + row_idx
        num_quarters_used = block_value(block, row_idx, num_quarters_col, min_col)
        last_quarter_used = block_value(block, row_idx, last_quarter_col, min_col)
        forecast_value = block_value(block, row_idx, forecast_value_col, min_col)
        actual_value = block_value(block, row_idx, reported_sales_col, min_col)
        forecast_max = block_value(block, row_idx, forecast_max_col, min_col)
        forecast_min = block_value(block, row_idx, forecast_min_col, min_col)
        quarterly_sales = block_value(block, row_idx, quarterly_sales_col, min_col)
        reported_sales = block_value(block, row_idx, reported_sales_col, min_col)
        growth_rate_pct = block_value(block, row_idx, growth_rate_col, min_col)
        sales_captured_in_db_pct = block_value(block, row_idx, captured_db_col, min_col)
        avg_penetration_pct = block_value(block, row_idx, avg_penetration_col, min_col)

        if to_number(avg_penetration_pct) is None and absolute_row in helper_avg_values:
            avg_penetration_pct = helper_avg_values[absolute_row]

        # Skip empty candidate lines.
        if all(
            value in (None, "")
            for value in [forecast_value, forecast_max, forecast_min, quarterly_sales, reported_sales]
        ):
            continue

        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = max_num - min_num if max_num is not None and min_num is not None else ""

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": number_or_blank(avg_penetration_pct),
                "num_quarters_used": number_or_blank(num_quarters_used),
                "last_quarter_used": last_quarter_used if last_quarter_used is not None else "",
                "forecast_value": number_or_blank(forecast_value),
                "actual_value": number_or_blank(actual_value),
                "forecast_max": number_or_blank(forecast_max),
                "forecast_min": number_or_blank(forecast_min),
                "range_width": number_or_blank(range_width),
                "avg_penetration_pct": number_or_blank(avg_penetration_pct),
                "quarterly_sales": number_or_blank(quarterly_sales),
                "reported_sales": number_or_blank(reported_sales),
                "growth_rate_pct": number_or_blank(growth_rate_pct),
                "sales_captured_in_db_pct": number_or_blank(sales_captured_in_db_pct),
                "source_file": source_file,
            }
        )

    return rows


def values_match(left: Any, right: Any, tol: float = 1e-9) -> bool:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is not None and right_num is not None:
        return abs(left_num - right_num) <= tol
    left_blank = left in (None, "")
    right_blank = right in (None, "")
    if left_blank and right_blank:
        return True
    return str(left).strip() == str(right).strip()


def is_duplicate_regression_row(current: Dict[str, Any], previous: Dict[str, Any]) -> bool:
    keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    return all(values_match(current.get(key), previous.get(key)) for key in keys)


def process_regression_sheet(
    wb: xw.main.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped Regression Model in {source_file}: sheet not found")
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"Skipped Regression Model in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = get_header_map(sheet, anchor_row, anchor_col)

    num_quarters_col = resolve_col(
        header_map,
        ["num_quarters_used", "num quarters used", "n quarters", "quarters used"],
        anchor_col - 10,
    )
    forecast_value_col = resolve_col(
        header_map,
        ["tot fcst w/o sa", "tot fcst without sa", "total forecast without sa", "forecast value"],
        anchor_col - 8,
    )
    actual_value_col = resolve_col(
        header_map,
        ["actual value", "actual sales", "reported sales"],
        None,
    )
    forecast_max_col = resolve_col(header_map, ["max"], anchor_col)
    forecast_min_col = resolve_col(header_map, ["min"], anchor_col + 1)

    data_start_row = anchor_row + 1
    data_end_row = anchor_row + 10  # n_quarters = 10
    read_columns = [
        col
        for col in [num_quarters_col, forecast_value_col, actual_value_col, forecast_max_col, forecast_min_col]
        if col is not None and col >= 1
    ]
    if not read_columns:
        print(f"Skipped Regression Model in {source_file}: no readable columns resolved")
        return []

    min_col = min(read_columns)
    max_col = max(read_columns)
    block = read_row_block(sheet, data_start_row, data_end_row, min_col, max_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"Skipped Regression Model in {source_file}: derived x/y columns are out of bounds")
        return []
    helper_start_col = max(sheet.used_range.last_cell.column, max_col) + 2
    intercept_col = helper_start_col
    slope_col = helper_start_col + 1

    for row_idx in range(10):
        absolute_row = data_start_row + row_idx
        n_val = block_value(block, row_idx, num_quarters_col, min_col)
        n_quarters = max(2, min(10, to_int(n_val, row_idx + 1) or (row_idx + 1)))
        start_hist_row = max(1, anchor_row - n_quarters)
        end_hist_row = anchor_row - 1

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},'
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col}),\"\")"
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},'
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col}),\"\")"
        )
        set_formula2_r1c1(sheet.cells(absolute_row, intercept_col), intercept_formula)
        set_formula2_r1c1(sheet.cells(absolute_row, slope_col), slope_formula)

    wb.app.calculate()

    intercept_values = to_2d(
        sheet.range((data_start_row, intercept_col), (data_end_row, intercept_col)).value,
        data_end_row - data_start_row + 1,
        1,
    )
    slope_values = to_2d(
        sheet.range((data_start_row, slope_col), (data_end_row, slope_col)).value,
        data_end_row - data_start_row + 1,
        1,
    )

    rows: List[Dict[str, Any]] = []
    for row_idx in range(10):
        num_quarters_used = block_value(block, row_idx, num_quarters_col, min_col)
        forecast_value = block_value(block, row_idx, forecast_value_col, min_col)
        actual_value = block_value(block, row_idx, actual_value_col, min_col)
        forecast_max = block_value(block, row_idx, forecast_max_col, min_col)
        forecast_min = block_value(block, row_idx, forecast_min_col, min_col)
        intercept = intercept_values[row_idx][0] if row_idx < len(intercept_values) else None
        slope = slope_values[row_idx][0] if row_idx < len(slope_values) else None

        # Skip empty candidate lines.
        if all(
            value in (None, "")
            for value in [num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope]
        ):
            continue

        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = max_num - min_num if max_num is not None and min_num is not None else ""

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": number_or_blank(num_quarters_used),
            "num_quarters_used": number_or_blank(num_quarters_used),
            "forecast_value": number_or_blank(forecast_value),
            "actual_value": number_or_blank(actual_value),
            "forecast_max": number_or_blank(forecast_max),
            "forecast_min": number_or_blank(forecast_min),
            "range_width": number_or_blank(range_width),
            "intercept": number_or_blank(intercept),
            "slope": number_or_blank(slope),
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(row, rows[-1]):
            continue
        rows.append(row)

    return rows


def apply_sheet_formatting_openpyxl(ws: Any, headers: Sequence[str]) -> None:
    if Font is None or get_column_letter is None:
        return

    for col_idx, header in enumerate(headers, start=1):
        header_cell = ws.cell(row=1, column=col_idx)
        if header_cell.value in (None, ""):
            header_cell.value = header
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, ws.max_row)}"

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        col_letter = get_column_letter(col_idx)
        for cell in ws[col_letter]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 40)


def apply_sheet_formatting_xlwings(
    sheet: xw.main.Sheet,
    headers: Sequence[str],
    data_rows: List[List[Any]],
) -> None:
    row_count = max(1, len(data_rows) + 1)
    col_count = len(headers)
    used_range = sheet.range((1, 1), (row_count, col_count))

    sheet.range((1, 1), (1, col_count)).api.Font.Bold = True
    used_range.api.AutoFilter()

    sheet.activate()
    sheet.range("A2").select()
    sheet.book.app.api.ActiveWindow.FreezePanes = True

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_values in data_rows:
            value = row_values[col_idx - 1]
            if value not in (None, ""):
                max_len = max(max_len, len(str(value)))
        sheet.range((1, col_idx), (row_count, col_idx)).column_width = min(max(max_len + 2, 12), 40)


def write_output_workbook_xlwings(
    app: xw.main.App,
    output_file: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = app.books.add()
    try:
        empirical_sheet = out_wb.sheets[0]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet = out_wb.sheets.add("regression_candidates", after=empirical_sheet)

        empirical_data = [[row.get(header, "") for header in EMPIRICAL_HEADERS] for row in empirical_rows]
        empirical_sheet.range((1, 1)).value = [EMPIRICAL_HEADERS] + empirical_data
        apply_sheet_formatting_xlwings(empirical_sheet, EMPIRICAL_HEADERS, empirical_data)

        regression_data = [[row.get(header, "") for header in REGRESSION_HEADERS] for row in regression_rows]
        regression_sheet.range((1, 1)).value = [REGRESSION_HEADERS] + regression_data
        apply_sheet_formatting_xlwings(regression_sheet, REGRESSION_HEADERS, regression_data)

        out_wb.save(str(output_file))
    finally:
        out_wb.close()


def write_output_workbook(
    output_file: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
    app: xw.main.App,
) -> None:
    if openpyxl is None:
        write_output_workbook_xlwings(app, output_file, empirical_rows, regression_rows)
        return

    out_wb = openpyxl.Workbook()
    default_ws = out_wb.active
    out_wb.remove(default_ws)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_ws.append([row.get(header, "") for header in EMPIRICAL_HEADERS])
    apply_sheet_formatting_openpyxl(empirical_ws, EMPIRICAL_HEADERS)

    regression_ws = out_wb.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_ws.append([row.get(header, "") for header in REGRESSION_HEADERS])
    apply_sheet_formatting_openpyxl(regression_ws, REGRESSION_HEADERS)

    out_wb.save(output_file)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = unique_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary Excel file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not .xlsx)")
                continue

            output_pattern = rf"^{re.escape(input_path.name)}_PARAM(?:\.\d+)?\.xlsx$"
            if re.match(output_pattern, file_path.name):
                print(f"Skipped: {file_path.name} (existing output workbook)")
                continue

            metadata, reason = parse_model_metadata(file_path.name)
            if metadata is None:
                print(f"Skipped: {file_path.name} ({reason})")
                continue

            print(f"Processing: {file_path.name}")
            wb: Optional[xw.main.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                processed_files += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (open/process error: {exc})")
            finally:
                close_source_workbook(wb)
        write_output_workbook(output_file, empirical_rows, regression_rows, app)
    finally:
        app.quit()

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
