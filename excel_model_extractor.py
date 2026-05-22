#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable locations
# ----------------------------
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
MONTH_ABBR_TO_NUM = {
    month.lower(): month_idx
    for month_idx, month in enumerate(calendar.month_abbr)
    if month
}


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetContext:
    ws: Any
    anchor_row: int
    anchor_col: int
    header_cols: Dict[str, int]


def normalize_text(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    normalized = re.sub(r"[^a-z0-9]+", " ", value.strip().lower())
    return re.sub(r"\s+", " ", normalized).strip()


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


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("%", "").strip()
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def value_signature(value: Any) -> Any:
    numeric = to_float(value)
    if numeric is not None:
        return round(numeric, 10)
    if value is None:
        return None
    return str(value).strip()


def safe_read(ws: Any, row: int, col: int) -> Any:
    if row <= 0 or col <= 0:
        return None
    return ws.cells(row, col).value


def clamp_col(col: int) -> int:
    return max(1, col)


def compute_range_width(max_value: Any, min_value: Any) -> Any:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return ""
    return max_num - min_num


def parse_file_labels(file_name: str) -> Optional[FileLabels]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    period_chunk = stem
    if len(parts) >= 3:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        period_chunk = parts[2]
    else:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_chunk = re.sub(
        r"[_\-\s]*(send|final|draft|rev\d+|v\d+).*$",
        "",
        period_chunk,
        flags=re.IGNORECASE,
    )
    period_match = re.search(
        r"(early|mid|late)\s*([A-Za-z]{3,9})\s*[_-]?\s*(\d{4})",
        period_chunk,
        flags=re.IGNORECASE,
    )
    if not ticker or not period_match:
        return None

    period_type = period_match.group(1).lower()
    month_token = period_match.group(2)[:3].lower()
    year = period_match.group(3)
    month_num = MONTH_ABBR_TO_NUM.get(month_token)
    if month_num is None:
        return None

    period_prefix = period_match.group(1).title()
    month_prefix = period_match.group(2)[:3].title()
    model_period = f"{period_prefix}{month_prefix}_{year}"
    model_date = date(int(year), month_num, DAY_BY_PERIOD[period_type]).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.resolve().name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def close_source_workbook(wb: Any) -> None:
    try:
        wb.close(save=False)
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

    try:
        wb.close()
    except Exception:
        pass


def unmerge_sheet_cells(ws: Any) -> bool:
    try:
        used = ws.used_range
    except Exception:
        return False

    try:
        merge_state = used.merge_cells
        if merge_state is False:
            return False
    except Exception:
        pass

    try:
        used.unmerge()
        return True
    except Exception:
        pass

    try:
        ws.api.UsedRange.UnMerge()
        return True
    except Exception:
        return False


def find_anchor(used_start_row: int, used_start_col: int, matrix: List[List[Any]]) -> Optional[Tuple[int, int]]:
    for row_idx, row_values in enumerate(matrix):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return used_start_row + row_idx, used_start_col + col_idx
    return None


def build_header_lookup(
    matrix: List[List[Any]],
    used_start_row: int,
    used_start_col: int,
    anchor_row: int,
) -> Dict[str, int]:
    header_cols: Dict[str, int] = {}
    min_row = max(used_start_row, anchor_row - 2)
    max_row = anchor_row + 2

    for row_number in range(min_row, max_row + 1):
        matrix_row_idx = row_number - used_start_row
        if matrix_row_idx < 0 or matrix_row_idx >= len(matrix):
            continue
        row_values = matrix[matrix_row_idx]
        for col_idx, cell_value in enumerate(row_values):
            key = normalize_text(cell_value)
            if not key:
                continue
            abs_col = used_start_col + col_idx
            header_cols.setdefault(key, abs_col)

    return header_cols


def find_sheet_context(wb: Any, sheet_name: str) -> Optional[SheetContext]:
    if sheet_name not in {sheet.name for sheet in wb.sheets}:
        return None

    ws = wb.sheets[sheet_name]
    used = ws.used_range
    matrix = ensure_2d(used.value)
    if not matrix:
        return None

    anchor = find_anchor(used.row, used.column, matrix)
    if anchor is None:
        return None

    anchor_row, anchor_col = anchor
    header_cols = build_header_lookup(matrix, used.row, used.column, anchor_row)
    return SheetContext(
        ws=ws,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        header_cols=header_cols,
    )


def resolve_column(
    anchor_col: int,
    header_cols: Dict[str, int],
    synonyms: Sequence[str],
    default_offset: int,
) -> int:
    for header_key, abs_col in header_cols.items():
        if any(syn in header_key for syn in synonyms):
            return abs_col
    return anchor_col + default_offset


def extract_empirical_rows(
    ctx: SheetContext,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    ws = ctx.ws
    anchor_row = ctx.anchor_row
    anchor_col = ctx.anchor_col
    headers = ctx.header_cols

    num_quarters_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("num quarters", "quarters used", "n quarters"), -4
        )
    )
    last_quarter_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("last quarter", "quarter used"), -3
        )
    )
    forecast_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("estimated total sold", "forecast", "tot fcst"), -2
        )
    )
    actual_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("reported sales", "actual"), -1
        )
    )
    reported_sales_col = clamp_col(
        resolve_column(
            anchor_col,
            headers,
            ("reported sales",),
            actual_col - anchor_col,
        )
    )
    min_col = clamp_col(resolve_column(anchor_col, headers, ("min",), 1))
    avg_pen_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("avg penetration", "average penetration"), -5
        )
    )
    quarterly_sales_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("quarterly sales",), -7
        )
    )
    growth_rate_col = clamp_col(
        resolve_column(anchor_col, headers, ("growth rate",), -6)
    )
    sales_captured_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("sales captured", "captured in db"), -8
        )
    )
    penetration_source_col = clamp_col(
        resolve_column(
            anchor_col,
            headers,
            ("penetration",),
            sales_captured_col - anchor_col,
        )
    )

    formula_written = False
    for n_quarters in range(1, N_QUARTERS + 1):
        result_row = anchor_row + n_quarters
        if anchor_row <= 1:
            break
        start_hist_row = max(1, anchor_row - n_quarters)
        end_hist_row = anchor_row - 1
        avg_cell = ws.cells(result_row, avg_pen_col)
        if is_blank(avg_cell.value) and start_hist_row <= end_hist_row:
            avg_cell.formula2 = (
                f'=IFERROR(AVERAGE(R{start_hist_row}C{penetration_source_col}'
                f":R{end_hist_row}C{penetration_source_col}),\"\")"
            )
            formula_written = True

    if formula_written:
        ws.book.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        result_row = anchor_row + n_quarters
        num_quarters_used = safe_read(ws, result_row, num_quarters_col)
        if to_int(num_quarters_used) is None:
            num_quarters_used = n_quarters

        forecast_value = safe_read(ws, result_row, forecast_col)
        actual_value = safe_read(ws, result_row, actual_col)
        forecast_max = safe_read(ws, result_row, anchor_col)
        forecast_min = safe_read(ws, result_row, min_col)
        avg_penetration_pct = safe_read(ws, result_row, avg_pen_col)
        quarterly_sales = safe_read(ws, result_row, quarterly_sales_col)
        reported_sales = safe_read(ws, result_row, reported_sales_col)
        growth_rate_pct = safe_read(ws, result_row, growth_rate_col)
        sales_captured_in_db_pct = safe_read(ws, result_row, sales_captured_col)
        last_quarter_used = safe_read(ws, result_row, last_quarter_col)

        row_is_blank = all(
            is_blank(value)
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                sales_captured_in_db_pct,
            )
        )
        if row_is_blank:
            if rows:
                break
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    ctx: SheetContext,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    ws = ctx.ws
    anchor_row = ctx.anchor_row
    anchor_col = ctx.anchor_col
    headers = ctx.header_cols

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("num quarters", "quarters used", "n quarters"), -4
        )
    )
    forecast_col = clamp_col(
        resolve_column(
        anchor_col,
        headers,
        ("tot fcst w/o sa", "total fcst w/o sa", "forecast", "tot fcst"),
        -2,
        )
    )
    actual_col = clamp_col(
        resolve_column(
        anchor_col, headers, ("actual", "reported sales"), -1
        )
    )
    min_col = clamp_col(resolve_column(anchor_col, headers, ("min",), 1))

    temp_intercept_col = anchor_col + 24
    temp_slope_col = anchor_col + 25

    formula_written = False
    for n_quarters in range(1, N_QUARTERS + 1):
        target_row = anchor_row + n_quarters
        num_quarters_value = safe_read(ws, target_row, num_quarters_col)
        num_quarters_used = to_int(num_quarters_value) or n_quarters
        end_row = anchor_row - 1
        start_row = max(1, end_row - num_quarters_used + 1)
        if start_row > end_row:
            continue

        ws.cells(target_row, temp_intercept_col).formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
        )
        ws.cells(target_row, temp_slope_col).formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
        )
        formula_written = True

    if formula_written:
        ws.book.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None
    for n_quarters in range(1, N_QUARTERS + 1):
        result_row = anchor_row + n_quarters
        num_quarters_value = safe_read(ws, result_row, num_quarters_col)
        num_quarters_used = to_int(num_quarters_value) or n_quarters

        intercept_value = safe_read(ws, result_row, temp_intercept_col)
        slope_value = safe_read(ws, result_row, temp_slope_col)
        forecast_value = safe_read(ws, result_row, forecast_col)
        actual_value = safe_read(ws, result_row, actual_col)
        forecast_max = safe_read(ws, result_row, anchor_col)
        forecast_min = safe_read(ws, result_row, min_col)

        if is_blank(forecast_value):
            x_for_forecast = safe_read(ws, anchor_row, x_col)
            intercept_num = to_float(intercept_value)
            slope_num = to_float(slope_value)
            x_num = to_float(x_for_forecast)
            if intercept_num is not None and slope_num is not None and x_num is not None:
                forecast_value = intercept_num + (slope_num * x_num)

        row_is_blank = all(
            is_blank(value)
            for value in (
                intercept_value,
                slope_value,
                forecast_value,
                forecast_max,
                forecast_min,
            )
        )
        if row_is_blank:
            if rows:
                break
            continue

        signature = (
            value_signature(intercept_value),
            value_signature(slope_value),
            value_signature(forecast_value),
            value_signature(forecast_max),
            value_signature(forecast_min),
        )
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if not is_blank(actual_value) else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    ws: Any,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(col_name, "") for col_name in headers])

    bold_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold_font

    ws.freeze_panes = "A2"
    max_row = max(1, len(rows) + 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max_row}"

    for col_idx, col_name in enumerate(headers, start=1):
        values = [col_name]
        for row in rows:
            value = row.get(col_name, "")
            values.append("" if value is None else str(value))
        max_length = max(len(str(value)) for value in values)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def should_skip_file(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temp file"
    if path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    return None


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    output_path = build_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_dir.iterdir()):
            skip_reason = should_skip_file(file_path)
            if skip_reason is not None:
                print(f"Skipped file: {file_path.name} ({skip_reason})")
                continue

            labels = parse_file_labels(file_path.name)
            if labels is None:
                print(f"Skipped file: {file_path.name} (could not parse ticker/period)")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                sheet_names = {sheet.name for sheet in wb.sheets}
                if "Empirical Model" in sheet_names and unmerge_sheet_cells(
                    wb.sheets["Empirical Model"]
                ):
                    print(f"Unmerged cells: {file_path.name} [Empirical Model]")
                if "Regression Model" in sheet_names and unmerge_sheet_cells(
                    wb.sheets["Regression Model"]
                ):
                    print(f"Unmerged cells: {file_path.name} [Regression Model]")

                empirical_ctx = find_sheet_context(wb, "Empirical Model")
                regression_ctx = find_sheet_context(wb, "Regression Model")
                if empirical_ctx is None and regression_ctx is None:
                    print(
                        f"Skipped file: {file_path.name} "
                        "(missing 'Empirical Model' and 'Regression Model' sheets)"
                    )
                    continue

                if empirical_ctx is not None:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_ctx, labels, file_path.name)
                    )
                if regression_ctx is not None:
                    regression_rows.extend(
                        extract_regression_rows(regression_ctx, labels, file_path.name)
                    )

                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
