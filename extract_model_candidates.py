#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Input/output locations
input_dir = Path("./input")
output_dir = Path("./output")

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

MONTH_MAP = {
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

PERIOD_DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

FILE_LABEL_PATTERN = re.compile(
    r"-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if not stripped:
            return None
        try:
            if stripped.endswith("%"):
                return float(stripped[:-1]) / 100.0
            return float(stripped)
        except ValueError:
            return None
    return None


def numeric_difference(left: Any, right: Any) -> Any:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def rounded_signature(value: Any) -> Optional[float]:
    as_float = to_float(value)
    if as_float is None:
        return None
    return round(as_float, 12)


def parse_file_labels(file_name: str) -> Optional[FileLabels]:
    stem = Path(file_name).stem
    match = FILE_LABEL_PATTERN.search(stem)
    if not match:
        return None

    ticker = match.group(1).upper()
    period_word = match.group(2).title()
    month_word = match.group(3)
    year = int(match.group(4))

    month_key = month_word[:3].lower()
    month_number = MONTH_MAP.get(month_key)
    if month_number is None:
        return None

    day = PERIOD_DAY_MAP[period_word.lower()]
    model_period = f"{period_word}{month_word[:3].title()}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_output_path(input_path: Path, target_output_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def find_anchor(
    values: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    label: str,
) -> Optional[Tuple[int, int]]:
    wanted = label.strip().lower()
    for row_index, row_values in enumerate(values):
        for col_index, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == wanted:
                return start_row + row_index, start_col + col_index
    return None


def matrix_cell(
    values: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int,
) -> Any:
    rel_row = row - start_row
    rel_col = col - start_col
    if rel_row < 0 or rel_col < 0:
        return None
    if rel_row >= len(values):
        return None
    row_values = values[rel_row]
    if rel_col >= len(row_values):
        return None
    return row_values[rel_col]


def read_cell(
    sheet: xw.Sheet,
    values: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int,
) -> Any:
    from_matrix = matrix_cell(values, start_row, start_col, row, col)
    if from_matrix is not None:
        return from_matrix
    return sheet.cells(row, col).value


def numeric_rows_single_col(
    values: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    col: int,
    row_upper_exclusive: int,
) -> List[int]:
    rows: List[int] = []
    rel_col = col - start_col
    if rel_col < 0:
        return rows
    for rel_row, row_values in enumerate(values):
        abs_row = start_row + rel_row
        if abs_row >= row_upper_exclusive:
            break
        if rel_col >= len(row_values):
            continue
        if to_float(row_values[rel_col]) is not None:
            rows.append(abs_row)
    return rows


def numeric_rows_two_cols(
    values: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    col_a: int,
    col_b: int,
    row_upper_exclusive: int,
) -> List[int]:
    rows: List[int] = []
    rel_a = col_a - start_col
    rel_b = col_b - start_col
    if rel_a < 0 or rel_b < 0:
        return rows
    for rel_row, row_values in enumerate(values):
        abs_row = start_row + rel_row
        if abs_row >= row_upper_exclusive:
            break
        if rel_a >= len(row_values) or rel_b >= len(row_values):
            continue
        if to_float(row_values[rel_a]) is None or to_float(row_values[rel_b]) is None:
            continue
        rows.append(abs_row)
    return rows


def safe_close_source_workbook(wb: xw.Book) -> None:
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


def process_empirical_sheet(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"SKIPPED: {source_file} (missing sheet: Empirical Model)")
        return []

    used_range = sheet.used_range
    values = normalize_2d(used_range.value)
    if not values:
        print(f"SKIPPED: {source_file} (Empirical Model is empty)")
        return []

    start_row = used_range.row
    start_col = used_range.column
    anchor = find_anchor(values, start_row, start_col, "max")
    if anchor is None:
        print(f"SKIPPED: {source_file} (Empirical Model missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor

    # Anchor-based offsets from the "max" cell.
    quarter_col = anchor_col - 12
    quarterly_sales_col = anchor_col - 11
    reported_sales_col = anchor_col - 10
    growth_rate_col = anchor_col - 9
    sales_captured_col = anchor_col - 8
    penetration_col = anchor_col - 7

    history_rows = numeric_rows_single_col(
        values=values,
        start_row=start_row,
        start_col=start_col,
        col=penetration_col,
        row_upper_exclusive=anchor_row,
    )
    if not history_rows:
        print(f"SKIPPED: {source_file} (no empirical history rows found)")
        return []

    history_rows = history_rows[-N_QUARTERS:]
    data_end_row = history_rows[-1]

    helper_avg_cell = sheet.cells(anchor_row, anchor_col + 4)

    output_rows: List[Dict[str, Any]] = []
    for n_used in range(1, len(history_rows) + 1):
        data_start_row = history_rows[-n_used]
        helper_avg_cell.formula2 = (
            f"=AVERAGE(R{data_start_row}C{penetration_col}:"
            f"R{data_end_row}C{penetration_col})"
        )
        wb.app.calculate()
        avg_penetration_pct = helper_avg_cell.value
        avg_penetration_num = to_float(avg_penetration_pct)

        forecast_value = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row - 1,
            anchor_col + 1,
        )
        reported_sales = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row - 2,
            anchor_col + 1,
        )
        if reported_sales is None:
            reported_sales = read_cell(
                sheet,
                values,
                start_row,
                start_col,
                data_end_row,
                reported_sales_col,
            )

        if to_float(forecast_value) is None:
            reported_sales_num = to_float(reported_sales)
            if (
                reported_sales_num is not None
                and avg_penetration_num is not None
                and avg_penetration_num != 0
            ):
                forecast_value = reported_sales_num / avg_penetration_num

        forecast_max = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row,
            anchor_col + 1,
        )
        forecast_min = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row + 1,
            anchor_col + 1,
        )

        output_rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_used,
                "last_quarter_used": read_cell(
                    sheet,
                    values,
                    start_row,
                    start_col,
                    data_end_row,
                    quarter_col,
                ),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_difference(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": read_cell(
                    sheet,
                    values,
                    start_row,
                    start_col,
                    data_end_row,
                    quarterly_sales_col,
                ),
                "reported_sales": reported_sales,
                "growth_rate_pct": read_cell(
                    sheet,
                    values,
                    start_row,
                    start_col,
                    data_end_row,
                    growth_rate_col,
                ),
                "sales_captured_in_db_pct": read_cell(
                    sheet,
                    values,
                    start_row,
                    start_col,
                    data_end_row,
                    sales_captured_col,
                ),
                "source_file": source_file,
            }
        )

    return output_rows


def process_regression_sheet(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"SKIPPED: {source_file} (missing sheet: Regression Model)")
        return []

    used_range = sheet.used_range
    values = normalize_2d(used_range.value)
    if not values:
        print(f"SKIPPED: {source_file} (Regression Model is empty)")
        return []

    start_row = used_range.row
    start_col = used_range.column
    anchor = find_anchor(values, start_row, start_col, "max")
    if anchor is None:
        print(f"SKIPPED: {source_file} (Regression Model missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    history_rows = numeric_rows_two_cols(
        values=values,
        start_row=start_row,
        start_col=start_col,
        col_a=x_col,
        col_b=y_col,
        row_upper_exclusive=anchor_row,
    )
    if len(history_rows) < 2:
        print(f"SKIPPED: {source_file} (not enough regression history rows)")
        return []

    history_rows = history_rows[-N_QUARTERS:]
    data_end_row = history_rows[-1]

    helper_intercept_cell = sheet.cells(anchor_row, anchor_col + 4)
    helper_slope_cell = sheet.cells(anchor_row + 1, anchor_col + 4)

    output_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for n_used in range(2, len(history_rows) + 1):
        data_start_row = history_rows[-n_used]

        helper_intercept_cell.formula2 = (
            f"=INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        helper_slope_cell.formula2 = (
            f"=SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        wb.app.calculate()

        intercept = helper_intercept_cell.value
        slope = helper_slope_cell.value
        forecast_total_without_sa = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row - 1,
            anchor_col + 1,
        )
        forecast_max = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row,
            anchor_col + 1,
        )
        forecast_min = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row + 1,
            anchor_col + 1,
        )
        actual_value = read_cell(
            sheet,
            values,
            start_row,
            start_col,
            anchor_row - 2,
            anchor_col + 1,
        )
        if actual_value is None:
            actual_value = ""

        signature = (
            rounded_signature(intercept),
            rounded_signature(slope),
            rounded_signature(forecast_total_without_sa),
            rounded_signature(forecast_max),
            rounded_signature(forecast_min),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        output_rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_used,
                "num_quarters_used": n_used,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_difference(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return output_rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"

    widths = [len(col) for col in columns]
    for row in rows:
        for idx, column in enumerate(columns):
            value = row.get(column, "")
            as_text = "" if value is None else str(value)
            widths[idx] = min(60, max(widths[idx], len(as_text)))

    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width + 2


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def iter_source_files(input_path: Path) -> List[Path]:
    valid_files: List[Path] = []
    prior_output_pattern = re.compile(
        rf"^{re.escape(input_path.name)}_PARAM(\.\d+)?\.xlsx$",
        re.IGNORECASE,
    )

    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"SKIPPED: {file_path.name} (temporary Excel file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"SKIPPED: {file_path.name} (not an .xlsx file)")
            continue
        if prior_output_pattern.match(file_path.name):
            print(f"SKIPPED: {file_path.name} (prior output workbook)")
            continue
        valid_files.append(file_path)
    return valid_files


def main() -> None:
    resolved_input_dir = input_dir.expanduser().resolve()
    resolved_output_dir = output_dir.expanduser().resolve()
    resolved_output_dir.mkdir(parents=True, exist_ok=True)

    if not resolved_input_dir.exists() or not resolved_input_dir.is_dir():
        raise SystemExit(f"Input directory does not exist: {resolved_input_dir}")

    output_path = get_output_path(resolved_input_dir, resolved_output_dir)
    source_files = iter_source_files(resolved_input_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files:
            labels = parse_file_labels(file_path.name)
            if labels is None:
                print(
                    "SKIPPED: "
                    f"{file_path.name} (filename does not match ticker/period pattern)"
                )
                continue

            print(f"PROCESSING: {file_path.name}")
            source_wb: Optional[xw.Book] = None
            try:
                source_wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    process_empirical_sheet(source_wb, labels, file_path.name)
                )
                regression_rows.extend(
                    process_regression_sheet(source_wb, labels, file_path.name)
                )
                processed_files += 1
                print(f"PROCESSED: {file_path.name}")
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (processing error: {exc})")
            finally:
                if source_wb is not None:
                    safe_close_source_workbook(source_wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"OUTPUT: {output_path}")
    print(f"FILES PROCESSED: {processed_files}")
    print(f"EMPIRICAL ROWS: {len(empirical_rows)}")
    print(f"REGRESSION ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
