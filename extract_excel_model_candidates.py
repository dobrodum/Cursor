#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Missing dependency 'xlwings'. Install it before running this script.") from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Missing dependency 'openpyxl'. Install it before running this script.") from exc


# User-configurable paths.
input_dir = Path("input")
output_dir = Path("output")

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
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    return isinstance(value, str) and not value.strip()


def to_float(value: Any) -> float | None:
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


def to_int(value: Any) -> int | None:
    number = to_float(value)
    if number is None:
        return None
    return int(number)


def first_non_blank(*values: Any) -> Any:
    for value in values:
        if not is_blank(value):
            return value
    return ""


def safe_subtract(left: Any, right: Any) -> Any:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def has_data_signal(values: list[Any]) -> bool:
    return any(not is_blank(value) for value in values)


def compact_signature(values: list[Any]) -> tuple[Any, ...]:
    signature: list[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is None:
            signature.append(normalize_text(value))
        else:
            signature.append(round(numeric, 10))
    return tuple(signature)


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
    col: int,
) -> Any:
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


def matrix_end_col(matrix: list[list[Any]], start_col: int) -> int:
    if not matrix:
        return start_col
    widest = max((len(row) for row in matrix), default=1)
    return start_col + max(widest - 1, 0)


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def parse_month(month_token: str) -> int:
    token = month_token.strip()
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token.title(), fmt).month
        except ValueError:
            continue
    return datetime.strptime(token[:3].title(), "%b").month


def parse_model_metadata(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"
    period_token = parts[2] if len(parts) >= 3 else stem
    period_token = re.sub(r"(?i)[_\s-]*send.*$", "", period_token).strip()

    model_period = re.sub(r"[\s-]+", "_", period_token).strip("_") if period_token else "unknown_period"
    model_date = ""

    match = re.search(
        r"(?i)(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})",
        period_token,
    )
    if match:
        phase = match.group(1).title()
        month = parse_month(match.group(2))
        year = int(match.group(3))
        model_period = f"{phase}{date(year, month, 1).strftime('%b')}_{year}"
        model_date = date(year, month, PERIOD_DAY_MAP[phase.lower()]).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    output_path = out_dir / f"{base_name}.xlsx"
    if not output_path.exists():
        return output_path

    suffix = 1
    while True:
        output_path = out_dir / f"{base_name}.{suffix}.xlsx"
        if not output_path.exists():
            return output_path
        suffix += 1


def should_skip_file(path: Path) -> str | None:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temp file"
    if path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    if re.search(r"(?i)_PARAM(?:\.\d+)?$", path.stem):
        return "looks like output artifact"
    return None


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


def get_sheet_case_insensitive(wb: Any, target_name: str) -> Any | None:
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
    end_row = start_row + len(matrix) - 1 if matrix else start_row - 1
    return matrix, start_row, start_col, end_row


def find_anchor(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
    anchor_text: str = "max",
) -> tuple[int, int] | None:
    needle = normalize_text(anchor_text)
    for row_index, row_values in enumerate(matrix):
        for col_index, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == needle:
                return start_row + row_index, start_col + col_index
    return None


def coerce_num_quarters(raw_value: Any, fallback: int) -> int:
    parsed = to_int(raw_value)
    if parsed is None:
        return fallback
    if parsed < 1 or parsed > N_QUARTERS:
        return fallback
    return parsed


def process_empirical_sheet(
    ws: Any,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    matrix, start_row, start_col, end_row = sheet_used_range_matrix(ws)
    anchor = find_anchor(matrix, start_row, start_col, anchor_text="max")
    if anchor is None:
        print("  skipped Empirical Model: no 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1

    num_quarters_col = anchor_col - 11
    last_quarter_col = anchor_col - 10
    quarterly_sales_col = anchor_col - 8
    growth_rate_col = anchor_col - 6
    sales_captured_col = anchor_col - 5
    forecast_total_col = anchor_col - 3
    reported_sales_col = anchor_col - 2
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1

    helper_base_col = max(matrix_end_col(matrix, start_col), anchor_col) + 4
    helper_avg_col = helper_base_col
    helper_forecast_col = helper_base_col + 1

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

        # Existing empirical logic: rolling average penetration and derived forecast.
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
    anchor = find_anchor(matrix, start_row, start_col, anchor_text="max")
    if anchor is None:
        print("  skipped Regression Model: no 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    num_quarters_col = anchor_col - 11
    forecast_total_col = anchor_col - 3
    actual_value_col = anchor_col - 2
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1

    helper_base_col = max(matrix_end_col(matrix, start_col), anchor_col) + 8
    helper_intercept_col = helper_base_col
    helper_slope_col = helper_base_col + 1

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
    previous_signature: tuple[Any, ...] | None = None
    final_row_index = len(rows_to_read) - 1

    for idx, (row, fallback_num_quarters) in enumerate(rows_to_read):
        num_quarters_used = coerce_num_quarters(
            matrix_cell(matrix, start_row, start_col, row, num_quarters_col),
            fallback_num_quarters,
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
            ws.cells(row, actual_value_col).value,
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


def write_sheet(ws: Any, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(column, "") for column in headers])

    bold = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, ws.max_row)}"

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 52)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_sheet = wb.create_sheet("empirical_candidates")
    regression_sheet = wb.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> int:
    input_path = Path(input_dir).expanduser().resolve()
    output_path_root = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        print(f"input directory does not exist or is not a directory: {input_path}")
        return 1

    output_path = build_output_path(input_path, output_path_root)

    source_files: list[Path] = []
    for path in sorted(input_path.iterdir()):
        reason = should_skip_file(path)
        if reason is not None:
            print(f"skipped {path.name}: {reason}")
            continue
        source_files.append(path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

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
    try:
        app.api.EnableEvents = False
    except Exception:
        pass

    try:
        for file_path in source_files:
            print(f"processing {file_path.name}")
            wb = None
            try:
                metadata = parse_model_metadata(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)

                empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
                if empirical_sheet is None:
                    print("  skipped Empirical Model: sheet not found")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")
                if regression_sheet is None:
                    print("  skipped Regression Model: sheet not found")
                else:
                    regression_rows.extend(
                        process_regression_sheet(
                            regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
                print(f"processed {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        safe_quit_excel_app(app)

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
