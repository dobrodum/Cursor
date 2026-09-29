#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---- Required input/output variables ----
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

N_QUARTERS = 10


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = text.replace("/", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def safe_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def numeric_difference(a: Any, b: Any) -> float | None:
    a_num = safe_number(a)
    b_num = safe_number(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def first_not_blank(*values: Any) -> Any:
    for value in values:
        if value not in (None, ""):
            return value
    return None


def close_workbook_safely(wb: xw.Book) -> None:
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
        wb.api.Close(False)


def parse_model_metadata(file_name: str) -> dict[str, Any]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 else ""
    period_token = parts[2] if len(parts) > 2 else ""

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if match:
        phase = match.group(1).title()
        month_token = match.group(2)[:3].title()
        year = int(match.group(3))
        month_lookup = {
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
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        month_num = month_lookup.get(month_token)
        day_num = day_lookup[phase]
        if month_num:
            model_period = f"{phase}{month_token}_{year}"
            model_date = date(year, month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    first_path = out_dir / f"{base_name}.xlsx"
    if not first_path.exists():
        return first_path

    index = 1
    while True:
        candidate = out_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def get_sheet_case_insensitive(wb: xw.Book, target_name: str) -> xw.Sheet | None:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def scan_sheet_once(
    ws: xw.Sheet, anchor_text: str = "max"
) -> tuple[list[list[Any]], int, int, int, int, int, int]:
    used = ws.used_range
    matrix = ensure_2d(used.value)
    base_row = used.row
    base_col = used.column
    row_count = len(matrix)
    col_count = max((len(row) for row in matrix), default=0)

    target = normalize_text(anchor_text)
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return (
                    matrix,
                    base_row,
                    base_col,
                    row_count,
                    col_count,
                    base_row + r_idx,
                    base_col + c_idx,
                )

    raise ValueError(f'Anchor "{anchor_text}" not found in sheet "{ws.name}"')


def build_header_map(
    matrix: list[list[Any]], base_row: int, base_col: int, header_row: int
) -> dict[str, int]:
    row_idx = header_row - base_row
    if row_idx < 0 or row_idx >= len(matrix):
        return {}

    mapping: dict[str, int] = {}
    for c_idx, value in enumerate(matrix[row_idx]):
        header = normalize_text(value)
        if header and header not in mapping:
            mapping[header] = base_col + c_idx
    return mapping


def find_col(
    header_map: dict[str, int], alternatives: list[tuple[str, ...]], fallback: int | None
) -> int | None:
    for keywords in alternatives:
        for header, col in header_map.items():
            if all(keyword in header for keyword in keywords):
                return col
    return fallback


def get_cell(ws: xw.Sheet, row: int, col: int | None) -> Any:
    if col is None or row < 1 or col < 1:
        return None
    return ws.cells(row, col).value


def extract_empirical_rows(
    wb: xw.Book,
    ws: xw.Sheet,
    metadata: dict[str, Any],
    source_file: str,
) -> list[dict[str, Any]]:
    (
        matrix,
        base_row,
        base_col,
        _row_count,
        col_count,
        anchor_row,
        anchor_col,
    ) = scan_sheet_once(ws, anchor_text="max")

    header_map = build_header_map(matrix, base_row, base_col, anchor_row)
    used_end_col = base_col + max(col_count, 1) - 1

    num_quarters_col = find_col(
        header_map,
        [("num", "quarter"), ("quarters", "used")],
        anchor_col - 8,
    )
    last_quarter_col = find_col(
        header_map,
        [("last", "quarter"), ("most", "recent", "quarter")],
        anchor_col - 7,
    )
    forecast_col = find_col(
        header_map,
        [("estimated", "total", "sold"), ("forecast", "value"), ("forecast",)],
        anchor_col - 3,
    )
    actual_col = find_col(
        header_map,
        [("reported", "sales"), ("actual", "value"), ("actual",)],
        anchor_col - 2,
    )
    max_col = find_col(header_map, [("max",)], anchor_col)
    min_col = find_col(header_map, [("min",)], anchor_col + 1)
    quarterly_sales_col = find_col(
        header_map,
        [("quarterly", "sales"), ("sales", "quarterly")],
        anchor_col - 6,
    )
    growth_rate_col = find_col(
        header_map,
        [("growth", "rate"), ("growth", "pct")],
        anchor_col - 5,
    )
    captured_col = find_col(
        header_map,
        [("sales", "captured", "db"), ("captured", "db"), ("captured", "pct")],
        anchor_col - 4,
    )
    avg_pen_col = find_col(
        header_map,
        [("avg", "penetration"), ("average", "penetration"), ("penetration", "pct")],
        anchor_col - 9,
    )
    penetration_history_col = find_col(
        header_map,
        [("penetration",), ("pct", "penetration")],
        avg_pen_col,
    )

    data_start_row = anchor_row + 1
    temp_calc_col = used_end_col + 2
    temp_calc_row = anchor_row
    calc_cell = ws.cells(temp_calc_row, temp_calc_col)

    history_available = max(1, anchor_row - base_row)
    rows: list[dict[str, Any]] = []

    for n in range(1, N_QUARTERS + 1):
        current_row = data_start_row + (n - 1)
        if num_quarters_col is not None:
            ws.cells(current_row, num_quarters_col).value = n

        avg_penetration = None
        if penetration_history_col is not None:
            sample_size = min(n, history_available)
            start_row = anchor_row - sample_size
            end_row = anchor_row - 1

            row_offset_start = start_row - temp_calc_row
            row_offset_end = end_row - temp_calc_row
            col_offset = penetration_history_col - temp_calc_col
            calc_cell.formula2 = (
                f'=IFERROR(AVERAGE(R[{row_offset_start}]C[{col_offset}]'
                f':R[{row_offset_end}]C[{col_offset}]),"")'
            )
            wb.app.calculate()
            avg_penetration = calc_cell.value

        avg_penetration = first_not_blank(avg_penetration, get_cell(ws, current_row, avg_pen_col))
        quarterly_sales = get_cell(ws, current_row, quarterly_sales_col)
        reported_sales = get_cell(ws, current_row, actual_col)
        forecast_value = get_cell(ws, current_row, forecast_col)
        forecast_max = get_cell(ws, current_row, max_col)
        forecast_min = get_cell(ws, current_row, min_col)

        if forecast_value in (None, ""):
            sales_num = safe_number(quarterly_sales)
            pen_num = safe_number(avg_penetration)
            if sales_num is not None and pen_num not in (None, 0):
                if pen_num > 1:
                    pen_num = pen_num / 100.0
                if pen_num not in (None, 0):
                    forecast_value = sales_num / pen_num

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": first_not_blank(get_cell(ws, current_row, num_quarters_col), n),
            "last_quarter_used": get_cell(ws, current_row, last_quarter_col),
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": numeric_difference(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": get_cell(ws, current_row, growth_rate_col),
            "sales_captured_in_db_pct": get_cell(ws, current_row, captured_col),
            "source_file": source_file,
        }
        rows.append(row)

    calc_cell.value = None
    return rows


def sameish(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    a_num = safe_number(a)
    b_num = safe_number(b)
    if a_num is not None and b_num is not None:
        return abs(a_num - b_num) <= tolerance
    return a == b


def is_duplicate_regression_row(prev: dict[str, Any], curr: dict[str, Any]) -> bool:
    keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    return all(sameish(prev.get(key), curr.get(key)) for key in keys)


def extract_regression_rows(
    wb: xw.Book,
    ws: xw.Sheet,
    metadata: dict[str, Any],
    source_file: str,
) -> list[dict[str, Any]]:
    (
        matrix,
        base_row,
        base_col,
        _row_count,
        col_count,
        anchor_row,
        anchor_col,
    ) = scan_sheet_once(ws, anchor_text="max")

    header_map = build_header_map(matrix, base_row, base_col, anchor_row)
    used_end_col = base_col + max(col_count, 1) - 1

    # Required by spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = find_col(
        header_map,
        [("num", "quarter"), ("quarters", "used")],
        anchor_col - 8,
    )
    forecast_col = find_col(
        header_map,
        [
            ("tot", "fcst", "w", "o", "sa"),
            ("tot", "fcst", "without", "sa"),
            ("forecast", "without", "sa"),
            ("forecast",),
        ],
        anchor_col - 3,
    )
    actual_col = find_col(
        header_map,
        [("actual",), ("reported", "sales")],
        None,
    )
    max_col = find_col(header_map, [("max",)], anchor_col)
    min_col = find_col(header_map, [("min",)], anchor_col + 1)

    data_start_row = anchor_row + 1
    temp_intercept_col = used_end_col + 2
    temp_slope_col = used_end_col + 3
    intercept_cell = ws.cells(anchor_row, temp_intercept_col)
    slope_cell = ws.cells(anchor_row, temp_slope_col)

    history_available = max(2, anchor_row - base_row)
    rows: list[dict[str, Any]] = []

    for n in range(1, N_QUARTERS + 1):
        current_row = data_start_row + (n - 1)
        if num_quarters_col is not None:
            ws.cells(current_row, num_quarters_col).value = n

        sample_size = max(2, min(n, history_available))
        start_row = anchor_row - sample_size
        end_row = anchor_row - 1

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT('
            f'R[{start_row - anchor_row}]C[{y_col - temp_intercept_col}]'
            f':R[{end_row - anchor_row}]C[{y_col - temp_intercept_col}],'
            f'R[{start_row - anchor_row}]C[{x_col - temp_intercept_col}]'
            f':R[{end_row - anchor_row}]C[{x_col - temp_intercept_col}]'
            f'),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE('
            f'R[{start_row - anchor_row}]C[{y_col - temp_slope_col}]'
            f':R[{end_row - anchor_row}]C[{y_col - temp_slope_col}],'
            f'R[{start_row - anchor_row}]C[{x_col - temp_slope_col}]'
            f':R[{end_row - anchor_row}]C[{x_col - temp_slope_col}]'
            f'),"")'
        )
        wb.app.calculate()

        forecast_max = get_cell(ws, current_row, max_col)
        forecast_min = get_cell(ws, current_row, min_col)

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": first_not_blank(get_cell(ws, current_row, num_quarters_col), n),
            "num_quarters_used": first_not_blank(get_cell(ws, current_row, num_quarters_col), n),
            "forecast_value": get_cell(ws, current_row, forecast_col),
            "actual_value": get_cell(ws, current_row, actual_col),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": numeric_difference(forecast_max, forecast_min),
            "intercept": intercept_cell.value,
            "slope": slope_cell.value,
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(rows[-1], row):
            continue
        rows.append(row)

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_sheet(
    ws,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col}{max(1, ws.max_row)}"

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for r in range(2, ws.max_row + 1):
            value = ws.cell(row=r, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir not found: {input_dir}")

    all_entries = sorted(input_dir.iterdir())
    files_to_process: list[Path] = []
    for file_path in all_entries:
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped file: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped file: {file_path.name} (not .xlsx)")
            continue
        files_to_process.append(file_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in files_to_process:
            print(f"processing file: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path.name)

                empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
                regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")

                if empirical_sheet is None and regression_sheet is None:
                    print(
                        f"skipped file: {file_path.name} "
                        f"(missing Empirical Model and Regression Model sheets)"
                    )
                    continue

                if empirical_sheet is None:
                    print(f"skipped Empirical Model in {file_path.name} (sheet missing)")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(wb, empirical_sheet, metadata, file_path.name)
                    )

                if regression_sheet is None:
                    print(f"skipped Regression Model in {file_path.name} (sheet missing)")
                else:
                    regression_rows.extend(
                        extract_regression_rows(wb, regression_sheet, metadata, file_path.name)
                    )

                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    output_path = next_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
