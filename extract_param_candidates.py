from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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

MONTHS = {
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def flatten_column(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if stripped.endswith("%"):
            stripped = stripped[:-1]
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def normalize_number(value: Any) -> Any:
    num = to_float(value)
    if num is None:
        return value
    if num.is_integer():
        return int(num)
    return num


def values_close(a: Any, b: Any, tol: float = 1e-9) -> bool:
    af = to_float(a)
    bf = to_float(b)
    if af is not None and bf is not None:
        return abs(af - bf) <= tol
    return a == b


def is_blank(value: Any) -> bool:
    return value is None or value == ""


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)([A-Za-z]+)(\d{4}))",
        stem,
        flags=re.IGNORECASE,
    )
    if not match:
        ticker = stem.split("_")[0].upper()
        return {
            "ticker": ticker,
            "model_period": "",
            "model_date": "",
            "model": ticker,
        }

    ticker = match.group("ticker").upper()
    period_token = match.group("period")
    period_match = re.match(
        r"(?P<label>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )
    if not period_match:
        return {
            "ticker": ticker,
            "model_period": "",
            "model_date": "",
            "model": ticker,
        }

    label = period_match.group("label").title()
    month_text = period_match.group("month")[:3].title()
    year = period_match.group("year")
    month_num = MONTHS.get(month_text.lower())
    day = PERIOD_DAY.get(label.lower())

    model_period = f"{label}{month_text}_{year}"
    model_date = ""
    if month_num and day:
        model_date = date(int(year), month_num, day).isoformat()

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}" if model_period else ticker,
    }


def get_next_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def close_workbook_safe(wb: xw.Book) -> None:
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
    except Exception:
        try:
            wb.api.Close(False)
        except Exception:
            pass


def find_max_anchor(values: List[List[Any]], start_row: int, start_col: int) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    for row_idx, row in enumerate(values):
        for col_idx, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            score = 0
            left = max(0, col_idx - 5)
            right = min(len(row), col_idx + 6)
            nearby = {normalize_text(v) for v in row[left:right]}
            if "min" in nearby:
                score += 2
            if "forecast" in nearby:
                score += 1
            candidates.append((score, start_row + row_idx, start_col + col_idx))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def build_header_cells(values: List[List[Any]], start_row: int, start_col: int, header_row: int) -> List[Tuple[str, int]]:
    rel_row = header_row - start_row
    if rel_row < 0 or rel_row >= len(values):
        return []
    header_cells: List[Tuple[str, int]] = []
    for col_idx, value in enumerate(values[rel_row]):
        key = normalize_text(value)
        if key:
            header_cells.append((key, start_col + col_idx))
    return header_cells


def find_column(header_cells: Sequence[Tuple[str, int]], keyword_sets: Sequence[Sequence[str]]) -> Optional[int]:
    for keywords in keyword_sets:
        for key, col in header_cells:
            if all(keyword in key for keyword in keywords):
                return col
    return None


def get_cell_value(
    sheet: xw.Sheet,
    values: List[List[Any]],
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    row: int,
    col: int,
) -> Any:
    if start_row <= row <= end_row and start_col <= col <= end_col:
        return values[row - start_row][col - start_col]
    return sheet.range((row, col)).value


def process_empirical_sheet(wb: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return []

    start_row = used.row
    start_col = used.column
    end_row = start_row + len(values) - 1
    end_col = start_col + len(values[0]) - 1

    anchor = find_max_anchor(values, start_row, start_col)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    header_cells = build_header_cells(values, start_row, start_col, anchor_row)
    if not header_cells and anchor_row - 1 >= start_row:
        header_cells = build_header_cells(values, start_row, start_col, anchor_row - 1)

    min_col = find_column(header_cells, [("min",)]) or (anchor_col + 1)
    num_quarters_col = find_column(
        header_cells,
        [
            ("num", "quarters", "used"),
            ("num", "quarter"),
            ("quarters", "used"),
            ("quarter",),
        ],
    )
    last_quarter_col = find_column(header_cells, [("last", "quarter", "used"), ("last", "quarter")])
    avg_pen_col = find_column(header_cells, [("avg", "penetration"), ("average", "penetration")])
    quarterly_sales_col = find_column(header_cells, [("quarterly", "sales"), ("quarter", "sales")])
    reported_sales_col = find_column(header_cells, [("reported", "sales"), ("actual", "sales")])
    growth_rate_col = find_column(header_cells, [("growth", "rate"), ("growth",)])
    captured_col = find_column(
        header_cells,
        [
            ("sales", "captured", "db"),
            ("captured", "db"),
            ("captured",),
        ],
    )
    forecast_value_col = find_column(
        header_cells,
        [
            ("estimated", "total", "sold"),
            ("estimate", "total", "sold"),
            ("forecast", "value"),
            ("tot", "fcst"),
            ("forecast",),
        ],
    )

    data_start_row = anchor_row + 1
    helper_col = end_col + 2
    avg_from_formula: Dict[int, Any] = {}

    if captured_col is not None:
        for offset in range(10):
            row = data_start_row + offset
            n_quarters = offset + 1
            if n_quarters == 1:
                start_ref = f"RC{captured_col}"
            else:
                start_ref = f"R[-{n_quarters - 1}]C{captured_col}"
            formula = f'=IFERROR(AVERAGE({start_ref}:RC{captured_col}),"")'
            sheet.range((row, helper_col)).formula2 = formula

        wb.app.calculate()
        formula_values = flatten_column(
            sheet.range((data_start_row, helper_col), (data_start_row + 9, helper_col)).value
        )
        for idx, value in enumerate(formula_values):
            avg_from_formula[data_start_row + idx] = value

    rows: List[Dict[str, Any]] = []
    for offset in range(10):
        row = data_start_row + offset

        num_quarters = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, num_quarters_col)
            if num_quarters_col is not None
            else offset + 1
        )
        last_quarter = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, last_quarter_col)
            if last_quarter_col is not None
            else None
        )
        forecast_value = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, forecast_value_col)
            if forecast_value_col is not None
            else None
        )
        reported_sales = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, reported_sales_col)
            if reported_sales_col is not None
            else None
        )
        forecast_max = get_cell_value(
            sheet, values, start_row, start_col, end_row, end_col, row, anchor_col
        )
        forecast_min = get_cell_value(
            sheet, values, start_row, start_col, end_row, end_col, row, min_col
        )
        quarterly_sales = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, quarterly_sales_col)
            if quarterly_sales_col is not None
            else None
        )
        growth_rate = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, growth_rate_col)
            if growth_rate_col is not None
            else None
        )
        captured_pct = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, captured_col)
            if captured_col is not None
            else None
        )
        avg_penetration = avg_from_formula.get(row)
        if avg_penetration in ("", None) and avg_pen_col is not None:
            avg_penetration = get_cell_value(
                sheet, values, start_row, start_col, end_row, end_col, row, avg_pen_col
            )

        row_is_empty = all(
            is_blank(value)
            for value in (
                forecast_value,
                reported_sales,
                forecast_max,
                forecast_min,
                quarterly_sales,
                growth_rate,
                captured_pct,
                avg_penetration,
            )
        )
        if row_is_empty:
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": normalize_number(avg_penetration),
                "num_quarters_used": normalize_number(num_quarters),
                "last_quarter_used": last_quarter,
                "forecast_value": normalize_number(forecast_value),
                "actual_value": normalize_number(reported_sales),
                "forecast_max": normalize_number(forecast_max),
                "forecast_min": normalize_number(forecast_min),
                "range_width": normalize_number(range_width),
                "avg_penetration_pct": normalize_number(avg_penetration),
                "quarterly_sales": normalize_number(quarterly_sales),
                "reported_sales": normalize_number(reported_sales),
                "growth_rate_pct": normalize_number(growth_rate),
                "sales_captured_in_db_pct": normalize_number(captured_pct),
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(wb: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return []

    start_row = used.row
    start_col = used.column
    end_row = start_row + len(values) - 1
    end_col = start_col + len(values[0]) - 1

    anchor = find_max_anchor(values, start_row, start_col)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        return []

    header_cells = build_header_cells(values, start_row, start_col, anchor_row)
    if not header_cells and anchor_row - 1 >= start_row:
        header_cells = build_header_cells(values, start_row, start_col, anchor_row - 1)

    min_col = find_column(header_cells, [("min",)]) or (anchor_col + 1)
    num_quarters_col = find_column(
        header_cells,
        [
            ("num", "quarters", "used"),
            ("num", "quarter"),
            ("quarters", "used"),
            ("quarter",),
        ],
    )
    forecast_col = find_column(
        header_cells,
        [
            ("tot", "fcst", "w_o", "sa"),
            ("tot", "fcst", "wo", "sa"),
            ("forecast", "without", "sa"),
            ("forecast", "w_o", "sa"),
            ("tot", "fcst"),
        ],
    )
    actual_col = find_column(header_cells, [("actual",), ("reported", "sales")])

    pairs: List[Tuple[int, float, float]] = []
    for row in range(start_row, end_row + 1):
        x_val = get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, x_col)
        y_val = get_cell_value(sheet, values, start_row, start_col, end_row, end_col, row, y_col)
        x_num = to_float(x_val)
        y_num = to_float(y_val)
        if x_num is None or y_num is None:
            continue
        pairs.append((row, x_num, y_num))

    if len(pairs) < 2:
        return []

    above_anchor = [item for item in pairs if item[0] < anchor_row]
    data_pairs = above_anchor if len(above_anchor) >= 2 else pairs

    max_n = min(10, len(data_pairs))
    n_values = list(range(2, max_n + 1))
    if not n_values:
        return []

    helper_intercept_col = end_col + 2
    helper_slope_col = end_col + 3
    helper_start_row = anchor_row + 1

    plan: List[Tuple[int, int, int, int]] = []
    for idx, n_quarters in enumerate(n_values):
        window = data_pairs[-n_quarters:]
        start_r = window[0][0]
        end_r = window[-1][0]
        formula_row = helper_start_row + idx

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_r}C{y_col}:R{end_r}C{y_col},'
            f'R{start_r}C{x_col}:R{end_r}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_r}C{y_col}:R{end_r}C{y_col},'
            f'R{start_r}C{x_col}:R{end_r}C{x_col}),"")'
        )
        sheet.range((formula_row, helper_intercept_col)).formula2 = intercept_formula
        sheet.range((formula_row, helper_slope_col)).formula2 = slope_formula
        plan.append((n_quarters, formula_row, start_r, end_r))

    wb.app.calculate()

    intercept_vals = flatten_column(
        sheet.range(
            (helper_start_row, helper_intercept_col),
            (helper_start_row + len(plan) - 1, helper_intercept_col),
        ).value
    )
    slope_vals = flatten_column(
        sheet.range(
            (helper_start_row, helper_slope_col),
            (helper_start_row + len(plan) - 1, helper_slope_col),
        ).value
    )

    latest_x = data_pairs[-1][1]
    rows: List[Dict[str, Any]] = []
    previous_values: Optional[Tuple[Any, Any, Any, Any, Any]] = None

    for idx, (n_quarters, _, _, _) in enumerate(plan):
        table_row = anchor_row + n_quarters
        intercept = intercept_vals[idx] if idx < len(intercept_vals) else None
        slope = slope_vals[idx] if idx < len(slope_vals) else None

        forecast_value = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, table_row, forecast_col)
            if forecast_col is not None
            else None
        )
        if is_blank(forecast_value):
            intercept_num = to_float(intercept)
            slope_num = to_float(slope)
            forecast_value = (
                (intercept_num + slope_num * latest_x)
                if intercept_num is not None and slope_num is not None
                else None
            )

        actual_value = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, table_row, actual_col)
            if actual_col is not None
            else None
        )
        forecast_max = get_cell_value(
            sheet, values, start_row, start_col, end_row, end_col, table_row, anchor_col
        )
        forecast_min = get_cell_value(
            sheet, values, start_row, start_col, end_row, end_col, table_row, min_col
        )
        num_quarters_used = (
            get_cell_value(sheet, values, start_row, start_col, end_row, end_col, table_row, num_quarters_col)
            if num_quarters_col is not None
            else n_quarters
        )

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        current_values = (
            normalize_number(intercept),
            normalize_number(slope),
            normalize_number(forecast_value),
            normalize_number(forecast_max),
            normalize_number(forecast_min),
        )
        if previous_values is not None and all(
            values_close(current_values[pos], previous_values[pos]) for pos in range(len(current_values))
        ):
            break
        previous_values = current_values

        row_is_empty = all(is_blank(item) for item in current_values)
        if row_is_empty:
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": normalize_number(num_quarters_used),
                "num_quarters_used": normalize_number(num_quarters_used),
                "forecast_value": normalize_number(forecast_value),
                "actual_value": normalize_number(actual_value),
                "forecast_max": normalize_number(forecast_max),
                "forecast_min": normalize_number(forecast_min),
                "range_width": normalize_number(range_width),
                "intercept": normalize_number(intercept),
                "slope": normalize_number(slope),
                "source_file": source_file,
            }
        )

    return rows


def set_column_widths(sheet, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = row.get(header)
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        sheet.column_dimensions[get_column_letter(idx)].width = min(60, max(12, max_len + 2))


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(column) for column in EMPIRICAL_COLUMNS])

    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(column) for column in REGRESSION_COLUMNS])

    for sheet, headers, rows in (
        (empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows),
        (regression_sheet, REGRESSION_COLUMNS, regression_rows),
    ):
        for header_cell in sheet[1]:
            header_cell.font = Font(bold=True)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        set_column_widths(sheet, headers, rows)

    workbook.save(output_file)


def should_skip_file(file_path: Path, output_prefix: str) -> Optional[str]:
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if file_path.stem.startswith(output_prefix):
        return "existing PARAM output file"
    return None


def process_workbooks(input_path: Path, output_path: Path) -> None:
    output_prefix = f"{input_path.name}_PARAM"
    output_file = get_next_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue

            skip_reason = should_skip_file(file_path, output_prefix)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            print(f"Processing {file_path.name}")
            metadata = parse_file_label(file_path.name)

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to process ({exc})")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


def main() -> None:
    input_path = Path(input_dir).expanduser()
    output_path = Path(output_dir).expanduser()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a directory: {input_path}")

    process_workbooks(input_path, output_path)


if __name__ == "__main__":
    main()
