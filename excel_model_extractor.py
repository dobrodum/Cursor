#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path(r"/path/to/input")
output_dir = Path(r"/path/to/output")

N_QUARTERS = 10
PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def to_1d(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] for row in values]
        return values
    return [values]


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    is_percent = "%" in text
    cleaned = text.replace(",", "").replace("%", "").strip()
    if cleaned in {"", "-", "--"}:
        return None

    try:
        numeric = float(cleaned)
    except ValueError:
        return None
    return numeric / 100.0 if is_percent else numeric


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def parse_model_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) > 1 and parts[1] else ""
    if not ticker:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()
    period_token_source = parts[2] if len(parts) > 2 else stem
    period_token = period_token_source.split("_")[0]

    match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if match:
        period_label = match.group(1).title()
        month_abbr = match.group(2).title()
        year = int(match.group(3))
        model_period = f"{period_label}{month_abbr}_{year}"
        try:
            month_number = datetime.strptime(month_abbr, "%b").month
            day = PERIOD_DAY_MAP[period_label.lower()]
            model_date = date(year, month_number, day).isoformat()
        except ValueError:
            model_date = ""

    model = "_".join(part for part in (ticker, model_period) if part)
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(input_path: Path, out_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def close_workbook_no_save(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        workbook.close()


def read_cell(sheet: xw.Sheet, row: Optional[int], col: Optional[int]) -> Any:
    if row is None or col is None:
        return None
    return sheet.range((row, col)).value


def find_anchor_cell(sheet: xw.Sheet, anchor_value: str = "max") -> Optional[Tuple[int, int]]:
    try:
        found = sheet.api.Cells.Find(What=anchor_value, LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    used = sheet.used_range
    values_2d = to_2d(used.value)
    anchor_text = anchor_value.strip().lower()
    for row_idx, row_vals in enumerate(values_2d):
        for col_idx, value in enumerate(row_vals):
            if str(value).strip().lower() == anchor_text:
                return used.row + row_idx, used.column + col_idx
    return None


def header_map_from_anchor_row(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int, window: int = 30
) -> Dict[int, str]:
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window
    values = sheet.range((anchor_row, start_col), (anchor_row, end_col)).value
    header_values = to_1d(values)

    headers: Dict[int, str] = {}
    for idx, value in enumerate(header_values):
        headers[start_col + idx] = normalize_text(value)
    return headers


def find_column(
    headers: Dict[int, str],
    include_groups: Sequence[Sequence[str]],
    exclude_terms: Sequence[str] = (),
) -> Optional[int]:
    for include_terms in include_groups:
        for col, header_text in headers.items():
            if not header_text:
                continue
            if all(term in header_text for term in include_terms) and not any(
                term in header_text for term in exclude_terms
            ):
                return col
    return None


def candidate_rows_from_num_quarters(
    sheet: xw.Sheet, anchor_row: int, num_quarters_col: Optional[int], max_rows: int
) -> List[int]:
    if num_quarters_col is None:
        return [anchor_row + idx for idx in range(1, max_rows + 1)]

    probe_end = anchor_row + 200
    values = sheet.range((anchor_row + 1, num_quarters_col), (probe_end, num_quarters_col)).value
    value_list = to_1d(values)

    rows: List[int] = []
    blanks_in_a_row = 0
    for idx, value in enumerate(value_list, start=anchor_row + 1):
        if to_int(value) is None:
            if rows:
                blanks_in_a_row += 1
                if blanks_in_a_row >= 3:
                    break
            continue

        blanks_in_a_row = 0
        rows.append(idx)
        if len(rows) >= max_rows:
            break

    return rows if rows else [anchor_row + idx for idx in range(1, max_rows + 1)]


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula
    except Exception:
        cell.api.FormulaR1C1 = formula


def find_penetration_series(sheet: xw.Sheet) -> Tuple[Optional[int], List[int]]:
    used = sheet.used_range
    values_2d = to_2d(used.value)

    label_row = None
    label_col = None
    for row_idx, row_values in enumerate(values_2d):
        for col_idx, value in enumerate(row_values):
            text = normalize_text(value)
            if "penetration" in text and "avg" not in text:
                label_row = used.row + row_idx
                label_col = used.column + col_idx
                break
        if label_col is not None:
            break

    if label_col is None or label_row is None:
        return None, []

    data_start_row = label_row + 1
    data_end_row = data_start_row + 400
    values = sheet.range((data_start_row, label_col), (data_end_row, label_col)).value
    value_list = to_1d(values)

    rows: List[int] = []
    blanks_in_a_row = 0
    for idx, value in enumerate(value_list, start=data_start_row):
        if to_float(value) is None:
            if rows:
                blanks_in_a_row += 1
                if blanks_in_a_row >= 2:
                    break
            continue

        blanks_in_a_row = 0
        rows.append(idx)

    return label_col, rows


def compute_avg_penetration_by_n(
    workbook: xw.Book,
    sheet: xw.Sheet,
    penetration_col: Optional[int],
    penetration_rows: List[int],
    n_quarters: int,
    helper_row: int,
    helper_col: int,
) -> Dict[int, Optional[float]]:
    if penetration_col is None or not penetration_rows:
        return {}

    max_n = min(n_quarters, len(penetration_rows))
    newest_row = penetration_rows[-1]
    for n in range(1, max_n + 1):
        start_row = penetration_rows[-n]
        formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{newest_row}C{penetration_col})"
        )
        helper_cell = sheet.range((helper_row + n - 1, helper_col))
        set_formula2_r1c1(helper_cell, formula)

    workbook.app.calculate()

    result_range = sheet.range((helper_row, helper_col), (helper_row + max_n - 1, helper_col))
    values = to_1d(result_range.value)
    return {n: to_float(value) for n, value in enumerate(values, start=1)}


def build_row_index_by_num_quarters(
    sheet: xw.Sheet, rows: List[int], num_quarters_col: Optional[int]
) -> Dict[int, int]:
    if num_quarters_col is None:
        return {idx: row for idx, row in enumerate(rows, start=1)}

    mapping: Dict[int, int] = {}
    for fallback_num, row in enumerate(rows, start=1):
        num = to_int(read_cell(sheet, row, num_quarters_col))
        if num is None:
            num = fallback_num
        if num not in mapping:
            mapping[num] = row
    return mapping


def collect_xy_rows(
    sheet: xw.Sheet, x_col: int, y_col: int, anchor_row: int
) -> List[int]:
    if x_col < 1 or y_col < 1:
        return []

    used = sheet.used_range
    start_row = used.row
    end_row = min(used.last_cell.row, anchor_row - 1)
    if end_row < start_row:
        end_row = used.last_cell.row

    x_values = to_1d(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_values = to_1d(sheet.range((start_row, y_col), (end_row, y_col)).value)

    rows: List[int] = []
    for idx, (x_val, y_val) in enumerate(zip(x_values, y_values), start=start_row):
        if to_float(x_val) is not None and to_float(y_val) is not None:
            rows.append(idx)
    return rows


def compute_regression_coefficients(
    workbook: xw.Book,
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    xy_rows: List[int],
    n_quarters: int,
    helper_row: int,
    helper_col: int,
) -> Dict[int, Tuple[Optional[float], Optional[float]]]:
    if not xy_rows:
        return {}

    max_n = min(n_quarters, len(xy_rows))
    newest_row = xy_rows[-1]

    for n in range(1, max_n + 1):
        start_row = xy_rows[-n]
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{newest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{newest_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{newest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{newest_row}C{x_col})"
        )
        intercept_cell = sheet.range((helper_row + n - 1, helper_col))
        slope_cell = sheet.range((helper_row + n - 1, helper_col + 1))
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)

    workbook.app.calculate()

    value_range = sheet.range(
        (helper_row, helper_col), (helper_row + max_n - 1, helper_col + 1)
    ).value
    values_2d = to_2d(value_range)
    output: Dict[int, Tuple[Optional[float], Optional[float]]] = {}
    for n, pair in enumerate(values_2d, start=1):
        intercept = to_float(pair[0]) if len(pair) > 0 else None
        slope = to_float(pair[1]) if len(pair) > 1 else None
        output[n] = (intercept, slope)
    return output


def extract_empirical_candidates(
    workbook: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"  Skipped empirical extraction for {source_file}: missing 'Empirical Model'.")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  Skipped empirical extraction for {source_file}: could not find 'max' anchor.")
        return rows

    anchor_row, anchor_col = anchor
    headers = header_map_from_anchor_row(sheet, anchor_row, anchor_col)

    num_quarters_col = find_column(
        headers,
        include_groups=[
            ("num", "quarter"),
            ("num", "qtr"),
            ("quarters", "used"),
        ],
    )
    last_quarter_col = find_column(headers, include_groups=[("last", "quarter")])
    forecast_col = find_column(
        headers,
        include_groups=[
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("forecast",),
        ],
        exclude_terms=("max", "min"),
    )
    actual_col = find_column(
        headers,
        include_groups=[
            ("reported", "sales"),
            ("actual", "sales"),
            ("actual",),
        ],
    )
    avg_pen_col = find_column(
        headers, include_groups=[("avg", "penetration"), ("average", "penetration")]
    )
    quarterly_sales_col = find_column(headers, include_groups=[("quarterly", "sales")])
    growth_col = find_column(headers, include_groups=[("growth", "rate"), ("growth",)])
    sales_capture_col = find_column(
        headers,
        include_groups=[
            ("captured", "db"),
            ("captured", "database"),
            ("sales", "captured"),
        ],
    )

    # Anchor-relative fallback positions for known model templates.
    num_quarters_col = num_quarters_col or max(1, anchor_col - 8)
    last_quarter_col = last_quarter_col or max(1, anchor_col - 7)
    forecast_col = forecast_col or max(1, anchor_col - 2)
    actual_col = actual_col or max(1, anchor_col - 3)
    avg_pen_col = avg_pen_col or max(1, anchor_col - 5)
    quarterly_sales_col = quarterly_sales_col or max(1, anchor_col - 4)

    max_col = anchor_col
    min_col = find_column(headers, include_groups=[("min",)]) or (anchor_col + 1)
    candidate_rows = candidate_rows_from_num_quarters(
        sheet, anchor_row, num_quarters_col, max_rows=N_QUARTERS
    )
    row_index_by_n = build_row_index_by_num_quarters(sheet, candidate_rows, num_quarters_col)

    used = sheet.used_range
    helper_row = used.last_cell.row + 5
    helper_col = used.last_cell.column + 2
    penetration_col, penetration_rows = find_penetration_series(sheet)
    avg_pen_from_formula = compute_avg_penetration_by_n(
        workbook,
        sheet,
        penetration_col,
        penetration_rows,
        n_quarters=N_QUARTERS,
        helper_row=helper_row,
        helper_col=helper_col,
    )

    for n in range(1, N_QUARTERS + 1):
        row_num = row_index_by_n.get(n)
        if row_num is None:
            row_num = anchor_row + n

        num_quarters_used = to_int(read_cell(sheet, row_num, num_quarters_col)) or n
        forecast_value = to_float(read_cell(sheet, row_num, forecast_col))
        actual_value = to_float(read_cell(sheet, row_num, actual_col))
        forecast_max = to_float(read_cell(sheet, row_num, max_col))
        forecast_min = to_float(read_cell(sheet, row_num, min_col))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        avg_penetration_pct = avg_pen_from_formula.get(n)
        if avg_penetration_pct is None:
            avg_penetration_pct = to_float(read_cell(sheet, row_num, avg_pen_col))

        quarterly_sales = to_float(read_cell(sheet, row_num, quarterly_sales_col))
        growth_rate_pct = to_float(read_cell(sheet, row_num, growth_col))
        sales_captured_in_db_pct = to_float(read_cell(sheet, row_num, sales_capture_col))
        last_quarter_used = read_cell(sheet, row_num, last_quarter_col)
        reported_sales = actual_value

        if (
            avg_penetration_pct is None
            and forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and quarterly_sales is None
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
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    workbook: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"  Skipped regression extraction for {source_file}: missing 'Regression Model'.")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  Skipped regression extraction for {source_file}: could not find 'max' anchor.")
        return rows

    anchor_row, anchor_col = anchor
    headers = header_map_from_anchor_row(sheet, anchor_row, anchor_col)

    num_quarters_col = find_column(
        headers,
        include_groups=[
            ("num", "quarter"),
            ("num", "qtr"),
            ("quarters", "used"),
        ],
    )
    forecast_col = find_column(
        headers,
        include_groups=[
            ("tot", "fcst", "sa"),
            ("tot", "forecast", "sa"),
            ("forecast", "without", "sa"),
            ("forecast", "w_o", "sa"),
        ],
        exclude_terms=("max", "min"),
    )
    intercept_col = find_column(headers, include_groups=[("intercept",)])
    slope_col = find_column(headers, include_groups=[("slope",)])

    # Anchor-relative fallback positions for known model templates.
    num_quarters_col = num_quarters_col or max(1, anchor_col - 4)
    intercept_col = intercept_col or max(1, anchor_col - 3)
    slope_col = slope_col or max(1, anchor_col - 2)
    forecast_col = forecast_col or max(1, anchor_col - 1)

    max_col = anchor_col
    min_col = find_column(headers, include_groups=[("min",)]) or (anchor_col + 1)

    candidate_rows = candidate_rows_from_num_quarters(
        sheet, anchor_row, num_quarters_col, max_rows=N_QUARTERS
    )
    row_index_by_n = build_row_index_by_num_quarters(sheet, candidate_rows, num_quarters_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    xy_rows = collect_xy_rows(sheet, x_col, y_col, anchor_row)

    used = sheet.used_range
    helper_row = used.last_cell.row + 5
    helper_col = used.last_cell.column + 2
    coeffs_by_n = compute_regression_coefficients(
        workbook,
        sheet,
        x_col=x_col,
        y_col=y_col,
        xy_rows=xy_rows,
        n_quarters=N_QUARTERS,
        helper_row=helper_row,
        helper_col=helper_col,
    )

    previous_signature: Optional[Tuple[Any, ...]] = None
    for n in range(1, N_QUARTERS + 1):
        row_num = row_index_by_n.get(n)
        if row_num is None and n <= len(candidate_rows):
            row_num = candidate_rows[n - 1]
        if row_num is None:
            row_num = anchor_row + n

        num_quarters_used = to_int(read_cell(sheet, row_num, num_quarters_col)) or n

        intercept, slope = coeffs_by_n.get(n, (None, None))
        if intercept is None:
            intercept = to_float(read_cell(sheet, row_num, intercept_col))
        if slope is None:
            slope = to_float(read_cell(sheet, row_num, slope_col))

        forecast_value = to_float(read_cell(sheet, row_num, forecast_col))
        if forecast_value is None and intercept is not None and slope is not None and xy_rows:
            last_x = to_float(read_cell(sheet, xy_rows[-1], x_col))
            if last_x is not None:
                forecast_value = intercept + slope * (last_x + 1.0)

        forecast_max = to_float(read_cell(sheet, row_num, max_col))
        forecast_min = to_float(read_cell(sheet, row_num, min_col))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if (
            intercept is None
            and slope is None
            and forecast_value is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        signature = (
            num_quarters_used,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
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
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    worksheet: Any, columns: List[str], rows: List[Dict[str, Any]], max_width: int = 40
) -> None:
    worksheet.append(columns)
    for item in rows:
        worksheet.append([item.get(column) for column in columns])

    bold_font = Font(bold=True)
    for col_idx in range(1, len(columns) + 1):
        worksheet.cell(row=1, column=col_idx).font = bold_font

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        width = len(column_name) + 2
        for row in rows:
            value = row.get(column_name)
            if value is None:
                continue
            width = max(width, len(str(value)) + 2)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(width, 12), max_width)


def write_output_workbook(
    output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def should_skip_file(file_path: Path, output_prefix: str) -> Tuple[bool, str]:
    if not file_path.is_file():
        return True, "not a regular file"
    if file_path.name.startswith("~"):
        return True, "temporary Excel file"
    if file_path.suffix.lower() != ".xlsx":
        return True, "not an .xlsx file"
    if file_path.name.startswith(output_prefix):
        return True, "generated PARAM output file"
    return False, ""


def main() -> None:
    input_path = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input folder does not exist: {input_path}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_prefix = f"{input_path.name}_PARAM"

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.api.AskToUpdateLinks = False
    except Exception:
        pass

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            skip, reason = should_skip_file(file_path, output_prefix)
            if skip:
                print(f"Skipped: {file_path.name} ({reason})")
                continue

            print(f"Processing: {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path.name)
                empirical_rows.extend(
                    extract_empirical_candidates(workbook, metadata, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(workbook, metadata, source_file=file_path.name)
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_no_save(workbook)
    finally:
        app.quit()

    output_path = next_output_path(input_path, out_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
