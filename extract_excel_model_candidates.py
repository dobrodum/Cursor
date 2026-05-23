#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Optional, Sequence

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency guard
    raise SystemExit(
        "Missing dependency 'xlwings'. Install it before running this script."
    ) from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover - runtime dependency guard
    raise SystemExit(
        "Missing dependency 'openpyxl'. Install it before running this script."
    ) from exc


# User-configurable paths.
input_dir = Path("input")
output_dir = Path("output")

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
    "sept": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

MONTH_ABBR = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}

PART_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def normalize_header(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
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


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def safe_subtract(left: Any, right: Any) -> Any:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def coerce_num_quarters(raw_value: Any, fallback: int) -> int:
    parsed = to_int(raw_value)
    if parsed is None:
        return fallback
    if parsed < 1 or parsed > N_QUARTERS:
        return fallback
    return parsed


def normalize_matrix(value: Any) -> list[list[Any]]:
    if value is None:
        return []
    if isinstance(value, tuple):
        value = list(value)
    if not isinstance(value, list):
        return [[value]]
    if not value:
        return []

    if isinstance(value[0], tuple):
        value = [list(row) if isinstance(row, tuple) else row for row in value]
    if isinstance(value[0], list):
        return value
    return [value]


def matrix_cell(
    matrix: list[list[Any]],
    matrix_start_row: int,
    matrix_start_col: int,
    row: int,
    col: Optional[int],
) -> Any:
    if col is None:
        return None

    row_index = row - matrix_start_row
    col_index = col - matrix_start_col
    if row_index < 0 or col_index < 0:
        return None
    if row_index >= len(matrix):
        return None
    row_values = matrix[row_index]
    if col_index >= len(row_values):
        return None
    return row_values[col_index]


def first_non_blank(*values: Any) -> Any:
    for value in values:
        if not is_blank(value):
            return value
    return ""


def has_data_signal(values: Sequence[Any]) -> bool:
    return any(not is_blank(value) for value in values)


def compact_signature(values: Sequence[Any]) -> tuple[Any, ...]:
    signature: list[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is None:
            signature.append(normalize_text(value))
        else:
            signature.append(round(numeric, 10))
    return tuple(signature)


def build_header_cells(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
    header_row: int,
) -> list[tuple[str, int]]:
    row_index = header_row - start_row
    if row_index < 0 or row_index >= len(matrix):
        return []
    row_values = matrix[row_index]
    headers: list[tuple[str, int]] = []
    for col_offset, value in enumerate(row_values):
        normalized = normalize_header(value)
        if normalized:
            headers.append((normalized, start_col + col_offset))
    return headers


def find_column(
    headers: list[tuple[str, int]],
    required_tokens: Sequence[str],
    forbidden_tokens: Sequence[str] = (),
) -> Optional[int]:
    required = [token.lower() for token in required_tokens]
    forbidden = [token.lower() for token in forbidden_tokens]
    for normalized_text, col in headers:
        if all(token in normalized_text for token in required) and all(
            token not in normalized_text for token in forbidden
        ):
            return col
    return None


def parse_file_label(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) >= 2 and parts[1].strip() else "UNKNOWN"
    period_token = parts[2] if len(parts) >= 3 else stem
    period_token = re.split(r"[_\s-]*send\b", period_token, flags=re.IGNORECASE)[0].strip()

    model_period = re.sub(r"[\s-]+", "_", period_token).strip("_") if period_token else ""
    model_date = ""

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )
    if match:
        phase = match.group(1).capitalize()
        month_token = match.group(2)
        year = int(match.group(3))

        month_num = MONTH_MAP.get(month_token.lower()) or MONTH_MAP.get(month_token[:3].lower())
        if month_num is not None:
            model_period = f"{phase}{MONTH_ABBR[month_num]}_{year}"
            model_date = date(year, month_num, PART_DAY_MAP[phase.lower()]).isoformat()

    if not model_period:
        model_period = "unknown_period"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def should_skip_file(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temporary file"
    if path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.search(r"_PARAM(?:\.\d+)?$", path.stem, flags=re.IGNORECASE):
        return "looks like generated output"
    return None


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback for Excel builds that don't expose formula2.
        cell.formula = formula


def safe_close_source_workbook(wb: Any) -> None:
    close_attempts = (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.close(),
    )
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception:
            continue


def safe_quit_excel_app(app: Any) -> None:
    try:
        app.quit()
    except Exception:
        pass


def get_sheet_case_insensitive(wb: Any, target_name: str) -> Optional[Any]:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def sheet_used_range_matrix(ws: Any) -> tuple[list[list[Any]], int, int, int]:
    used = ws.used_range
    start_row = used.row
    start_col = used.column
    matrix = normalize_matrix(used.options(ndim=2).value)
    end_row = start_row + len(matrix) - 1 if matrix else (start_row - 1)
    return matrix, start_row, start_col, end_row


def find_anchor(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
    text: str = "max",
) -> Optional[tuple[int, int]]:
    target = normalize_text(text)
    for row_index, row_values in enumerate(matrix):
        for col_index, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                return start_row + row_index, start_col + col_index
    return None


def process_empirical_sheet(
    ws: Any,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    matrix, start_row, start_col, end_row = sheet_used_range_matrix(ws)
    anchor = find_anchor(matrix, start_row, start_col, text="max")
    if anchor is None:
        print("  - Empirical Model: skipped (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    headers = build_header_cells(matrix, start_row, start_col, anchor_row)

    num_quarters_col = (
        find_column(headers, ["num", "quarter"])
        or find_column(headers, ["quarters", "used"])
        or (anchor_col - 11)
    )
    last_quarter_col = find_column(headers, ["last", "quarter"]) or (anchor_col - 10)
    quarterly_sales_col = (
        find_column(headers, ["quarterly", "sales"], ["reported"])
        or find_column(headers, ["quarter", "sales"], ["reported"])
        or (anchor_col - 8)
    )
    growth_rate_col = find_column(headers, ["growth"]) or (anchor_col - 6)
    sales_captured_col = (
        find_column(headers, ["captured", "db"])
        or find_column(headers, ["captured"])
        or (anchor_col - 5)
    )
    forecast_total_col = (
        find_column(headers, ["estimated", "total"])
        or find_column(headers, ["tot", "fcst"])
        or find_column(headers, ["forecast"])
        or (anchor_col - 3)
    )
    reported_sales_col = (
        find_column(headers, ["reported", "sales"])
        or find_column(headers, ["actual", "sales"])
        or (anchor_col - 2)
    )
    forecast_max_col = anchor_col
    forecast_min_col = find_column(headers, ["min"]) or (anchor_col + 1)

    helper_avg_col = anchor_col + 20
    helper_forecast_col = anchor_col + 21

    rows_to_read: list[tuple[int, int]] = []
    for offset in range(N_QUARTERS):
        row = data_start_row + offset
        if row > end_row:
            break

        probe_values = [
            matrix_cell(matrix, start_row, start_col, row, forecast_max_col),
            matrix_cell(matrix, start_row, start_col, row, quarterly_sales_col),
            matrix_cell(matrix, start_row, start_col, row, reported_sales_col),
            matrix_cell(matrix, start_row, start_col, row, sales_captured_col),
        ]
        if all(is_blank(value) for value in probe_values):
            if rows_to_read:
                break
            continue

        avg_cell = ws.cells(row, helper_avg_col)
        forecast_cell = ws.cells(row, helper_forecast_col)

        set_formula2(
            avg_cell,
            (
                f'=IFERROR(AVERAGE(R{data_start_row}C{sales_captured_col}:'
                f'R{row}C{sales_captured_col}),"")'
            ),
        )
        set_formula2(
            forecast_cell,
            f'=IFERROR(R{row}C{quarterly_sales_col}/R{row}C{helper_avg_col},"")',
        )
        rows_to_read.append((row, offset + 1))

    if rows_to_read:
        ws.book.app.calculate()

    extracted_rows: list[dict[str, Any]] = []
    for row, fallback_num_quarters in rows_to_read:
        num_quarters_used = coerce_num_quarters(
            matrix_cell(matrix, start_row, start_col, row, num_quarters_col),
            fallback_num_quarters,
        )

        avg_penetration_pct = ws.cells(row, helper_avg_col).value
        calc_forecast_value = ws.cells(row, helper_forecast_col).value
        forecast_value = first_non_blank(
            ws.cells(row, forecast_total_col).value,
            matrix_cell(matrix, start_row, start_col, row, forecast_total_col),
            calc_forecast_value,
        )
        forecast_max = first_non_blank(
            ws.cells(row, forecast_max_col).value,
            matrix_cell(matrix, start_row, start_col, row, forecast_max_col),
        )
        forecast_min = first_non_blank(
            ws.cells(row, forecast_min_col).value,
            matrix_cell(matrix, start_row, start_col, row, forecast_min_col),
        )
        quarterly_sales = first_non_blank(
            ws.cells(row, quarterly_sales_col).value,
            matrix_cell(matrix, start_row, start_col, row, quarterly_sales_col),
        )
        reported_sales = first_non_blank(
            ws.cells(row, reported_sales_col).value,
            matrix_cell(matrix, start_row, start_col, row, reported_sales_col),
        )
        growth_rate_pct = first_non_blank(
            ws.cells(row, growth_rate_col).value,
            matrix_cell(matrix, start_row, start_col, row, growth_rate_col),
        )
        sales_captured_in_db_pct = first_non_blank(
            ws.cells(row, sales_captured_col).value,
            matrix_cell(matrix, start_row, start_col, row, sales_captured_col),
        )
        last_quarter_used = first_non_blank(
            ws.cells(row, last_quarter_col).value,
            matrix_cell(matrix, start_row, start_col, row, last_quarter_col),
        )

        if not has_data_signal(
            [
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            ]
        ):
            continue

        extracted_rows.append(
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
                "actual_value": first_non_blank(reported_sales, ""),
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    for row, _ in rows_to_read:
        ws.cells(row, helper_avg_col).value = None
        ws.cells(row, helper_forecast_col).value = None

    return extracted_rows


def process_regression_sheet(
    ws: Any,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    matrix, start_row, start_col, end_row = sheet_used_range_matrix(ws)
    anchor = find_anchor(matrix, start_row, start_col, text="max")
    if anchor is None:
        print("  - Regression Model: skipped (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    headers = build_header_cells(matrix, start_row, start_col, anchor_row)

    # Required by spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = (
        find_column(headers, ["num", "quarter"])
        or find_column(headers, ["quarters", "used"])
        or (anchor_col - 11)
    )
    forecast_total_col = (
        find_column(headers, ["tot", "fcst", "w", "o", "sa"])
        or find_column(headers, ["tot", "fcst", "without", "sa"])
        or find_column(headers, ["forecast"])
        or (anchor_col - 3)
    )
    actual_value_col = (
        find_column(headers, ["actual", "value"])
        or find_column(headers, ["actual", "sales"])
        or find_column(headers, ["reported", "sales"])
    )
    forecast_max_col = anchor_col
    forecast_min_col = find_column(headers, ["min"]) or (anchor_col + 1)
    helper_intercept_col = anchor_col + 20
    helper_slope_col = anchor_col + 21

    rows_to_read: list[tuple[int, int]] = []
    for offset in range(N_QUARTERS):
        row = data_start_row + offset
        if row > end_row:
            break

        probe_values = [
            matrix_cell(matrix, start_row, start_col, row, y_col),
            matrix_cell(matrix, start_row, start_col, row, x_col),
            matrix_cell(matrix, start_row, start_col, row, forecast_total_col),
            matrix_cell(matrix, start_row, start_col, row, forecast_max_col),
        ]
        if all(is_blank(value) for value in probe_values):
            if rows_to_read:
                break
            continue

        intercept_cell = ws.cells(row, helper_intercept_col)
        slope_cell = ws.cells(row, helper_slope_col)

        set_formula2(
            intercept_cell,
            (
                f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
            ),
        )
        set_formula2(
            slope_cell,
            (
                f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
            ),
        )
        rows_to_read.append((row, offset + 1))

    if rows_to_read:
        ws.book.app.calculate()

    extracted_rows: list[dict[str, Any]] = []
    previous_signature: Optional[tuple[Any, ...]] = None
    final_row_index = len(rows_to_read) - 1

    for idx, (row, fallback_num_quarters) in enumerate(rows_to_read):
        num_quarters_used = coerce_num_quarters(
            matrix_cell(matrix, start_row, start_col, row, num_quarters_col),
            max(2, fallback_num_quarters),
        )
        intercept = to_float(ws.cells(row, helper_intercept_col).value)
        slope = to_float(ws.cells(row, helper_slope_col).value)
        forecast_value = first_non_blank(
            ws.cells(row, forecast_total_col).value,
            matrix_cell(matrix, start_row, start_col, row, forecast_total_col),
        )
        forecast_max = first_non_blank(
            ws.cells(row, forecast_max_col).value,
            matrix_cell(matrix, start_row, start_col, row, forecast_max_col),
        )
        forecast_min = first_non_blank(
            ws.cells(row, forecast_min_col).value,
            matrix_cell(matrix, start_row, start_col, row, forecast_min_col),
        )
        actual_value = first_non_blank(
            ws.cells(row, actual_value_col).value if actual_value_col is not None else "",
            matrix_cell(matrix, start_row, start_col, row, actual_value_col),
        )
        if is_blank(actual_value):
            actual_value = ""

        if not has_data_signal(
            [num_quarters_used, intercept, slope, forecast_value, forecast_max, forecast_min]
        ):
            continue

        signature = compact_signature(
            [num_quarters_used, intercept, slope, forecast_value, forecast_max, forecast_min]
        )
        if idx == final_row_index and previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        extracted_rows.append(
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
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    for row, _ in rows_to_read:
        ws.cells(row, helper_intercept_col).value = None
        ws.cells(row, helper_slope_col).value = None

    return extracted_rows


def process_workbook_once(
    app: Any,
    file_path: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    metadata = parse_file_label(file_path.name)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    wb = app.books.open(str(file_path), update_links=False)
    try:
        empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
        if empirical_sheet is None:
            print("  - Empirical Model: skipped (sheet not found)")
        else:
            empirical_rows = process_empirical_sheet(
                empirical_sheet, metadata=metadata, source_file=file_path.name
            )

        regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")
        if regression_sheet is None:
            print("  - Regression Model: skipped (sheet not found)")
        else:
            regression_rows = process_regression_sheet(
                regression_sheet, metadata=metadata, source_file=file_path.name
            )
    finally:
        safe_close_source_workbook(wb)

    return empirical_rows, regression_rows


def apply_sheet_formatting(ws: Any, headers: Sequence[str]) -> None:
    if not headers:
        return

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    bold_font = Font(bold=True)
    for index in range(1, len(headers) + 1):
        ws.cell(row=1, column=index).font = bold_font

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_sheet_rows(ws: Any, headers: Sequence[str], rows: Sequence[dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header, "") for header in headers])
    apply_sheet_formatting(ws, headers)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[dict[str, Any]],
    regression_rows: Sequence[dict[str, Any]],
) -> None:
    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    empirical_ws = output_wb.create_sheet("empirical_candidates")
    regression_ws = output_wb.create_sheet("regression_candidates")

    write_sheet_rows(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet_rows(regression_ws, REGRESSION_COLUMNS, regression_rows)

    output_wb.save(output_path)


def main() -> None:
    in_folder = Path(input_dir).expanduser().resolve()
    out_folder = Path(output_dir).expanduser().resolve()
    if not in_folder.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_folder}")
    if not in_folder.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {in_folder}")

    output_path = build_output_path(in_folder, out_folder)
    source_files = sorted(in_folder.iterdir())

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
    except Exception:
        pass
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass

    processed_files = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    try:
        for file_path in source_files:
            skip_reason = should_skip_file(file_path)
            if skip_reason is not None:
                print(f"skipped {file_path.name}: {skip_reason}")
                continue

            print(f"processing {file_path.name}")
            try:
                file_empirical, file_regression = process_workbook_once(app, file_path)
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
                continue

            empirical_rows.extend(file_empirical)
            regression_rows.extend(file_regression)
            processed_files += 1
            print(f"processed {file_path.name}")
    finally:
        safe_quit_excel_app(app)

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
