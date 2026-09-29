#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Configure these paths
# -----------------------------------------------------------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10

EMPIRICAL_OUTPUT_COLUMNS = [
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

REGRESSION_OUTPUT_COLUMNS = [
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

DAY_BY_WINDOW = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).replace(",", "").strip()
        if not text:
            return None
        return float(text)
    except (TypeError, ValueError):
        return None


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_number(max_value)
    min_num = to_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def resolve_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def parse_model_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_token = parts[2] if len(parts) >= 3 else ""
    period_token = period_token.split("_")[0].strip()

    model_period = ""
    model_date = ""

    match = re.match(
        r"^(early|mid|late)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if match:
        window = match.group(1).lower()
        month_abbr = match.group(2).lower()
        year = int(match.group(3))

        window_label = window.capitalize()
        month_label = month_abbr.capitalize()
        model_period = f"{window_label}{month_label}_{year}"
        model_date = date(year, MONTHS[month_abbr], DAY_BY_WINDOW[window]).isoformat()
    else:
        cleaned = re.sub(r"[^A-Za-z0-9]+", "_", period_token).strip("_")
        model_period = cleaned

    model = f"{ticker}_{model_period}".strip("_")

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def read_used_range(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    values = used.options(ndim=2).value
    if values is None:
        values = [[]]
    return start_row, start_col, values


def matrix_get(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int,
) -> Any:
    r_idx = row - start_row
    c_idx = col - start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(matrix):
        return None
    row_values = matrix[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def find_anchor_cell(
    matrix: Sequence[Sequence[Any]], start_row: int, start_col: int, anchor_text: str = "max"
) -> Optional[Tuple[int, int]]:
    target = normalize_text(anchor_text)
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) == target:
                return start_row + r_idx, start_col + c_idx
    return None


def build_header_lookup(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    header_row: int,
) -> Dict[int, str]:
    r_idx = header_row - start_row
    if r_idx < 0 or r_idx >= len(matrix):
        return {}
    header_values = matrix[r_idx]
    lookup: Dict[int, str] = {}
    for c_idx, value in enumerate(header_values):
        text = normalize_text(value)
        if text:
            lookup[start_col + c_idx] = text
    return lookup


def find_col_with_aliases(
    header_lookup: Dict[int, str],
    aliases: Iterable[str],
    anchor_col: int,
    fallback_offset: Optional[int] = None,
) -> Optional[int]:
    normalized_aliases = [alias.lower() for alias in aliases]
    matches: List[int] = []
    for col, text in header_lookup.items():
        if any(alias in text for alias in normalized_aliases):
            matches.append(col)
    if matches:
        return min(matches, key=lambda col: abs(col - anchor_col))
    if fallback_offset is not None:
        return anchor_col + fallback_offset
    return None


def row_has_any_data(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    cols: Iterable[Optional[int]],
) -> bool:
    for col in cols:
        if col is None:
            continue
        value = matrix_get(matrix, start_row, start_col, row, col)
        if value not in (None, ""):
            return True
    return False


def process_empirical_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    start_row, start_col, matrix = read_used_range(sheet)
    anchor = find_anchor_cell(matrix, start_row, start_col, anchor_text="max")
    if anchor is None:
        print(f"  skipped empirical extraction (missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    header_lookup = build_header_lookup(matrix, start_row, start_col, anchor_row)

    cols = {
        "num_quarters_used": find_col_with_aliases(
            header_lookup,
            aliases=("num quarter", "quarters used", "qtrs used", "n quarter"),
            anchor_col=anchor_col,
            fallback_offset=-4,
        ),
        "last_quarter_used": find_col_with_aliases(
            header_lookup,
            aliases=("last quarter", "last qtr"),
            anchor_col=anchor_col,
            fallback_offset=-3,
        ),
        "forecast_value": find_col_with_aliases(
            header_lookup,
            aliases=("estimated total sold", "est total sold", "forecast"),
            anchor_col=anchor_col,
            fallback_offset=-2,
        ),
        "actual_value": find_col_with_aliases(
            header_lookup,
            aliases=("reported sales", "actual"),
            anchor_col=anchor_col,
            fallback_offset=-1,
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_col_with_aliases(
            header_lookup,
            aliases=("min",),
            anchor_col=anchor_col,
            fallback_offset=1,
        ),
        "avg_penetration_pct": find_col_with_aliases(
            header_lookup,
            aliases=("avg penetration", "average penetration"),
            anchor_col=anchor_col,
            fallback_offset=2,
        ),
        "quarterly_sales": find_col_with_aliases(
            header_lookup,
            aliases=("quarterly sales", "quarter sales"),
            anchor_col=anchor_col,
            fallback_offset=3,
        ),
        "growth_rate_pct": find_col_with_aliases(
            header_lookup,
            aliases=("growth rate",),
            anchor_col=anchor_col,
            fallback_offset=4,
        ),
        "sales_captured_in_db_pct": find_col_with_aliases(
            header_lookup,
            aliases=("sales captured", "captured in db", "db pct"),
            anchor_col=anchor_col,
            fallback_offset=5,
        ),
    }

    helper_col = max(header_lookup) + 2 if header_lookup else (anchor_col + 8)
    helper_row = anchor_row
    helper_cell = sheet.range((helper_row, helper_col))
    historical_pen_col = (
        cols["sales_captured_in_db_pct"]
        if cols["sales_captured_in_db_pct"] is not None
        else cols["avg_penetration_pct"]
    )

    rows: List[Dict[str, Any]] = []
    blank_streak = 0
    trailing_end_row = anchor_row - 1

    try:
        for i in range(N_QUARTERS):
            data_row = anchor_row + 1 + i
            if not row_has_any_data(
                matrix,
                start_row,
                start_col,
                data_row,
                (
                    cols["num_quarters_used"],
                    cols["forecast_value"],
                    cols["forecast_max"],
                    cols["forecast_min"],
                    cols["actual_value"],
                    cols["avg_penetration_pct"],
                ),
            ):
                blank_streak += 1
                if blank_streak >= 2:
                    break
                continue

            blank_streak = 0
            num_quarters = matrix_get(
                matrix, start_row, start_col, data_row, cols["num_quarters_used"]
            )
            if num_quarters in (None, ""):
                num_quarters = i + 1

            avg_pen_calc = None
            num_q_for_formula = int(to_number(num_quarters) or (i + 1))
            if historical_pen_col is not None and trailing_end_row >= start_row:
                start_hist = max(start_row, trailing_end_row - num_q_for_formula + 1)
                helper_cell.formula2 = (
                    f'=IFERROR(AVERAGE(R{start_hist}C{historical_pen_col}:'
                    f'R{trailing_end_row}C{historical_pen_col}),"")'
                )
                workbook.app.calculate()
                avg_pen_calc = helper_cell.value

            forecast_max = matrix_get(matrix, start_row, start_col, data_row, cols["forecast_max"])
            forecast_min = matrix_get(matrix, start_row, start_col, data_row, cols["forecast_min"])
            forecast_value = matrix_get(matrix, start_row, start_col, data_row, cols["forecast_value"])
            reported_sales = matrix_get(matrix, start_row, start_col, data_row, cols["actual_value"])
            avg_pen = matrix_get(
                matrix, start_row, start_col, data_row, cols["avg_penetration_pct"]
            )
            if avg_pen in (None, ""):
                avg_pen = avg_pen_calc

            row = {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": num_quarters,
                "last_quarter_used": matrix_get(
                    matrix, start_row, start_col, data_row, cols["last_quarter_used"]
                ),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": matrix_get(
                    matrix, start_row, start_col, data_row, cols["quarterly_sales"]
                ),
                "reported_sales": reported_sales,
                "growth_rate_pct": matrix_get(
                    matrix, start_row, start_col, data_row, cols["growth_rate_pct"]
                ),
                "sales_captured_in_db_pct": matrix_get(
                    matrix, start_row, start_col, data_row, cols["sales_captured_in_db_pct"]
                ),
                "source_file": source_file,
            }
            rows.append(row)
    finally:
        helper_cell.clear_contents()

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    start_row, start_col, matrix = read_used_range(sheet)
    anchor = find_anchor_cell(matrix, start_row, start_col, anchor_text="max")
    if anchor is None:
        print("  skipped regression extraction (missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    header_lookup = build_header_lookup(matrix, start_row, start_col, anchor_row)

    cols = {
        "num_quarters_used": find_col_with_aliases(
            header_lookup,
            aliases=("num quarter", "quarters used", "qtrs used", "n quarter"),
            anchor_col=anchor_col,
            fallback_offset=-2,
        ),
        "forecast_value": find_col_with_aliases(
            header_lookup,
            aliases=("tot fcst w/o sa", "forecast", "fcst"),
            anchor_col=anchor_col,
            fallback_offset=-1,
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_col_with_aliases(
            header_lookup,
            aliases=("min",),
            anchor_col=anchor_col,
            fallback_offset=1,
        ),
        "actual_value": find_col_with_aliases(
            header_lookup,
            aliases=("actual", "reported"),
            anchor_col=anchor_col,
            fallback_offset=None,
        ),
    }

    # Required anchor-based source columns for regression formulas.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = anchor_row - 1

    helper_base_col = max(header_lookup) + 2 if header_lookup else (anchor_col + 6)
    intercept_cell = sheet.range((anchor_row, helper_base_col))
    slope_cell = sheet.range((anchor_row, helper_base_col + 1))

    rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Any, ...]] = None

    try:
        for i in range(N_QUARTERS):
            n_quarters = i + 1
            start_hist_row = history_end_row - n_quarters + 1
            if start_hist_row < start_row:
                break

            intercept_cell.formula2 = (
                f'=IFERROR(INTERCEPT(R{start_hist_row}C{y_col}:R{history_end_row}C{y_col},'
                f'R{start_hist_row}C{x_col}:R{history_end_row}C{x_col}),"")'
            )
            slope_cell.formula2 = (
                f'=IFERROR(SLOPE(R{start_hist_row}C{y_col}:R{history_end_row}C{y_col},'
                f'R{start_hist_row}C{x_col}:R{history_end_row}C{x_col}),"")'
            )
            workbook.app.calculate()

            intercept = intercept_cell.value
            slope = slope_cell.value

            data_row = anchor_row + 1 + i
            file_num_quarters = matrix_get(
                matrix, start_row, start_col, data_row, cols["num_quarters_used"]
            )
            if file_num_quarters not in (None, ""):
                n_quarters = file_num_quarters

            forecast_value = matrix_get(
                matrix, start_row, start_col, data_row, cols["forecast_value"]
            )
            if forecast_value in (None, ""):
                x_next = matrix_get(matrix, start_row, start_col, history_end_row + 1, x_col)
                intercept_num = to_number(intercept)
                slope_num = to_number(slope)
                x_next_num = to_number(x_next)
                if (
                    intercept_num is not None
                    and slope_num is not None
                    and x_next_num is not None
                ):
                    forecast_value = intercept_num + (slope_num * x_next_num)

            forecast_max = matrix_get(matrix, start_row, start_col, data_row, cols["forecast_max"])
            forecast_min = matrix_get(matrix, start_row, start_col, data_row, cols["forecast_min"])
            actual_value = matrix_get(
                matrix, start_row, start_col, data_row, cols["actual_value"]
            )
            if actual_value in (None,):
                actual_value = ""

            dedupe_key = (
                n_quarters,
                to_number(forecast_value),
                to_number(forecast_max),
                to_number(forecast_min),
                to_number(intercept),
                to_number(slope),
            )
            if dedupe_key == previous_key:
                continue
            previous_key = dedupe_key

            rows.append(
                {
                    "model": metadata["model"],
                    "ticker": metadata["ticker"],
                    "model_period": metadata["model_period"],
                    "model_date": metadata["model_date"],
                    "method": "regression",
                    "parameter_name": "num_quarters_used",
                    "parameter_value": n_quarters,
                    "num_quarters_used": n_quarters,
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
    finally:
        intercept_cell.clear_contents()
        slope_cell.clear_contents()

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    if sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = workbook.create_sheet(sheet_name)

    ws.append(list(columns))
    for col_index in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_index).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_index, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_index).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max(12, max_len + 2), 45)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb_out = Workbook()

    # Replace default sheet with empirical sheet.
    default_ws = wb_out.active
    wb_out.remove(default_ws)

    write_sheet(wb_out, "empirical_candidates", EMPIRICAL_OUTPUT_COLUMNS, empirical_rows)
    write_sheet(wb_out, "regression_candidates", REGRESSION_OUTPUT_COLUMNS, regression_rows)

    wb_out.save(output_path)


def should_skip_file(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    if path.name.startswith("~"):
        return "temporary file"
    return None


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = resolve_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    files = sorted(input_dir.iterdir())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files:
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"SKIPPED: {file_path.name} ({skip_reason})")
                continue

            workbook: Optional[xw.Book] = None
            try:
                print(f"PROCESSING: {file_path.name}")
                workbook = app.books.open(str(file_path), update_links=False)

                metadata = parse_model_metadata(file_path.name)

                try:
                    empirical_sheet = workbook.sheets[EMPIRICAL_SHEET_NAME]
                except Exception:
                    empirical_sheet = None

                try:
                    regression_sheet = workbook.sheets[REGRESSION_SHEET_NAME]
                except Exception:
                    regression_sheet = None

                if empirical_sheet is None and regression_sheet is None:
                    print(
                        "SKIPPED: "
                        f"{file_path.name} (missing '{EMPIRICAL_SHEET_NAME}' and "
                        f"'{REGRESSION_SHEET_NAME}')"
                    )
                    continue

                if empirical_sheet is not None:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            workbook=workbook,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"  skipped empirical extraction (missing '{EMPIRICAL_SHEET_NAME}')")

                if regression_sheet is not None:
                    regression_rows.extend(
                        process_regression_sheet(
                            workbook=workbook,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"  skipped regression extraction (missing '{REGRESSION_SHEET_NAME}')")

                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {processed_files}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
