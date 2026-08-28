from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ======================
# User-configurable paths
# ======================
input_dir = "/path/to/input"
output_dir = "/path/to/output"

N_QUARTERS = 10


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

PERIOD_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_float(value: Any) -> Optional[float]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        txt = value.strip().replace(",", "")
        if not txt:
            return None
        if txt.endswith("%"):
            txt = txt[:-1]
            try:
                return float(txt) / 100.0
            except ValueError:
                return None
        try:
            return float(txt)
        except ValueError:
            return None
    return None


def numeric_difference(left: Any, right: Any) -> Optional[float]:
    l_num = to_float(left)
    r_num = to_float(right)
    if l_num is None or r_num is None:
        return None
    return l_num - r_num


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def safe_set_formula2(cell: Any, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_book(wb: Any) -> None:
    # Primary path requested by requirements.
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Fallback with explicit COM close that avoids save.
    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    # Last fallback; display alerts are disabled on app.
    try:
        wb.close()
    except Exception:
        pass


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if not match:
        raise ValueError(
            "filename does not match expected pattern (e.g., '... - AORT - MidJan2026_...')"
        )

    ticker = match.group("ticker").upper()
    period_token = match.group("period").title()
    month_token = match.group("month")
    year_token = match.group("year")

    month_key = month_token[:3].lower()
    if month_key not in MONTH_MAP:
        raise ValueError(f"unrecognized month token '{month_token}'")

    month_number = MONTH_MAP[month_key]
    day_number = PERIOD_DAY_MAP[period_token]
    month_label = month_key.title()

    model_period = f"{period_token}{month_label}_{year_token}"
    model_date = date(int(year_token), month_number, day_number).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.name
    base_name = f"{folder_name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def find_anchor_cell(sheet: Any, anchor_label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    anchor_norm = normalize_text(anchor_label)
    start_row = used.row
    start_col = used.column
    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == anchor_norm:
                return start_row + r_idx, start_col + c_idx
    return None


def get_header_cells(
    sheet: Any, header_row: int, anchor_col: int, span: int = 35
) -> List[Tuple[int, str, str]]:
    start_col = max(1, anchor_col - span)
    end_col = anchor_col + span
    row_values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]

    result: List[Tuple[int, str, str]] = []
    for idx, value in enumerate(row_values):
        abs_col = start_col + idx
        raw = "" if value is None else str(value)
        result.append((abs_col, normalize_text(value), raw))
    return result


def pick_column_near_anchor(
    header_cells: Sequence[Tuple[int, str, str]],
    anchor_col: int,
    keyword_patterns: Sequence[str],
    default_col: Optional[int] = None,
) -> Optional[int]:
    matches: List[int] = []
    for col, norm, _ in header_cells:
        if any(pattern in norm for pattern in keyword_patterns):
            matches.append(col)
    if not matches:
        return default_col
    return min(matches, key=lambda col: abs(col - anchor_col))


def read_block(
    sheet: Any, start_row: int, end_row: int, start_col: int, end_col: int
) -> List[List[Any]]:
    values = sheet.range((start_row, start_col), (end_row, end_col)).options(ndim=2).value
    return ensure_2d(values)


def block_value(
    block: Sequence[Sequence[Any]], row_idx: int, abs_col: Optional[int], min_col: int
) -> Any:
    if abs_col is None:
        return None
    col_idx = abs_col - min_col
    if row_idx < 0 or row_idx >= len(block):
        return None
    row = block[row_idx]
    if col_idx < 0 or col_idx >= len(row):
        return None
    return row[col_idx]


def collect_numeric_rows_single_col(sheet: Any, col: int, max_row: int) -> List[int]:
    if max_row < 1:
        return []
    start_row = max(1, sheet.used_range.row)
    if max_row < start_row:
        return []
    values = sheet.range((start_row, col), (max_row, col)).value
    if not isinstance(values, list):
        values = [values]
    rows: List[int] = []
    for idx, value in enumerate(values):
        if to_float(value) is not None:
            rows.append(start_row + idx)
    return rows


def collect_numeric_rows_pair(sheet: Any, x_col: int, y_col: int, max_row: int) -> List[int]:
    if max_row < 1:
        return []
    start_row = max(1, sheet.used_range.row)
    if max_row < start_row:
        return []

    x_values = sheet.range((start_row, x_col), (max_row, x_col)).value
    y_values = sheet.range((start_row, y_col), (max_row, y_col)).value
    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]

    paired_rows: List[int] = []
    for idx, (x_value, y_value) in enumerate(zip(x_values, y_values)):
        if to_float(x_value) is not None and to_float(y_value) is not None:
            paired_rows.append(start_row + idx)
    return paired_rows


def build_empirical_rows(
    wb: Any, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        return []
    sheet = wb.sheets["Empirical Model"]

    anchor = find_anchor_cell(sheet, "max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor
    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1

    header_cells = get_header_cells(sheet, anchor_row, anchor_col)
    col_map = {
        "num_quarters_used": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["numquartersused", "quartersused", "nquarters", "numqtrs", "numquarters"],
        ),
        "last_quarter_used": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["lastquarterused", "lastquarter", "latestquarter"],
        ),
        "avg_penetration_pct": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["avgpenetrationpct", "avgpenetration", "averagepenetration", "penetrationpct"],
        ),
        "forecast_value": pick_column_near_anchor(
            header_cells,
            anchor_col,
            [
                "estimatedtotalsold",
                "esttotalsold",
                "forecastvalue",
                "forecasttotalsold",
                "totfcstwosa",
            ],
        ),
        "reported_sales": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["reportedsales", "actualsales", "salesreported"],
        ),
        "quarterly_sales": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["quarterlysales", "salesquarter", "quartersales"],
        ),
        "growth_rate_pct": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["growthratepct", "growthrate", "growthpct"],
        ),
        "sales_captured_in_db_pct": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["salescapturedindbpct", "capturedindb", "dbcapturepct", "salescapturedpct"],
        ),
        "forecast_min": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["forecastmin", "min"],
            default_col=anchor_col + 1,
        ),
    }

    penetration_source_col = pick_column_near_anchor(
        header_cells,
        anchor_col,
        ["salescapturedindbpct", "penetrationpct", "penetration"],
    )
    formulas_written = False
    avg_pen_col = col_map["avg_penetration_pct"]
    if avg_pen_col is not None and penetration_source_col is not None:
        hist_rows = collect_numeric_rows_single_col(sheet, penetration_source_col, anchor_row - 1)
        if hist_rows:
            for idx in range(N_QUARTERS):
                n_use = min(idx + 1, len(hist_rows))
                first_hist_row = hist_rows[-n_use]
                last_hist_row = hist_rows[-1]
                formula = (
                    f'=IFERROR(AVERAGE(R{first_hist_row}C{penetration_source_col}:'
                    f'R{last_hist_row}C{penetration_source_col}),"")'
                )
                target_cell = sheet.range((first_data_row + idx, avg_pen_col))
                safe_set_formula2(target_cell, formula)
                formulas_written = True

    if formulas_written:
        wb.app.calculate()

    tracked_cols = [anchor_col, col_map["forecast_min"]]
    tracked_cols.extend(col for col in col_map.values() if col is not None)
    min_col = min(tracked_cols)
    max_col = max(tracked_cols)
    block = read_block(sheet, first_data_row, last_data_row, min_col, max_col)

    results: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        forecast_max = block_value(block, idx, anchor_col, min_col)
        forecast_min = block_value(block, idx, col_map["forecast_min"], min_col)
        forecast_value = block_value(block, idx, col_map["forecast_value"], min_col)
        avg_penetration = block_value(block, idx, col_map["avg_penetration_pct"], min_col)

        if is_blank(forecast_max) and is_blank(forecast_min) and is_blank(forecast_value) and is_blank(
            avg_penetration
        ):
            continue

        raw_num_quarters = block_value(block, idx, col_map["num_quarters_used"], min_col)
        num_quarters_used = raw_num_quarters if not is_blank(raw_num_quarters) else idx + 1
        reported_sales = block_value(block, idx, col_map["reported_sales"], min_col)

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": block_value(block, idx, col_map["last_quarter_used"], min_col),
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": numeric_difference(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": block_value(block, idx, col_map["quarterly_sales"], min_col),
            "reported_sales": reported_sales,
            "growth_rate_pct": block_value(block, idx, col_map["growth_rate_pct"], min_col),
            "sales_captured_in_db_pct": block_value(
                block, idx, col_map["sales_captured_in_db_pct"], min_col
            ),
            "source_file": source_file,
        }
        results.append(row)

    return results


def value_signature(value: Any) -> Any:
    number_value = to_float(value)
    if number_value is not None:
        return round(number_value, 10)
    if is_blank(value):
        return None
    return str(value).strip()


def build_regression_rows(
    wb: Any, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        return []
    sheet = wb.sheets["Regression Model"]

    anchor = find_anchor_cell(sheet, "max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor
    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_cells = get_header_cells(sheet, anchor_row, anchor_col)
    col_map = {
        "num_quarters_used": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["numquartersused", "quartersused", "nquarters", "numqtrs", "numquarters"],
        ),
        "forecast_value": pick_column_near_anchor(
            header_cells,
            anchor_col,
            [
                "totfcstwosa",
                "totfcstwithoutsa",
                "forecasttotalwithoutsa",
                "forecastvalue",
                "totforecastwosa",
            ],
        ),
        "actual_value": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["actualvalue", "actualsales", "reportedsales"],
        ),
        "forecast_min": pick_column_near_anchor(
            header_cells,
            anchor_col,
            ["forecastmin", "min"],
            default_col=anchor_col + 1,
        ),
    }

    paired_rows = collect_numeric_rows_pair(sheet, x_col, y_col, anchor_row - 1)
    used_last_col = sheet.used_range.last_cell.column
    scratch_intercept_col = max(used_last_col + 2, anchor_col + 2)
    scratch_slope_col = scratch_intercept_col + 1

    formulas_written = False
    for idx in range(N_QUARTERS):
        n_use = min(idx + 1, len(paired_rows))
        target_row = first_data_row + idx

        if n_use >= 2:
            first_hist_row = paired_rows[-n_use]
            last_hist_row = paired_rows[-1]
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{first_hist_row}C{y_col}:R{last_hist_row}C{y_col},'
                f'R{first_hist_row}C{x_col}:R{last_hist_row}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{first_hist_row}C{y_col}:R{last_hist_row}C{y_col},'
                f'R{first_hist_row}C{x_col}:R{last_hist_row}C{x_col}),"")'
            )
        else:
            intercept_formula = '=""'
            slope_formula = '=""'

        safe_set_formula2(sheet.range((target_row, scratch_intercept_col)), intercept_formula)
        safe_set_formula2(sheet.range((target_row, scratch_slope_col)), slope_formula)
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    tracked_cols = [
        anchor_col,
        col_map["forecast_min"],
        col_map["num_quarters_used"],
        col_map["forecast_value"],
        col_map["actual_value"],
        scratch_intercept_col,
        scratch_slope_col,
    ]
    tracked_cols = [col for col in tracked_cols if col is not None]
    min_col = min(tracked_cols)
    max_col = max(tracked_cols)
    block = read_block(sheet, first_data_row, last_data_row, min_col, max_col)

    results: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx in range(N_QUARTERS):
        forecast_max = block_value(block, idx, anchor_col, min_col)
        forecast_min = block_value(block, idx, col_map["forecast_min"], min_col)
        forecast_value = block_value(block, idx, col_map["forecast_value"], min_col)
        intercept = block_value(block, idx, scratch_intercept_col, min_col)
        slope = block_value(block, idx, scratch_slope_col, min_col)

        if (
            is_blank(forecast_max)
            and is_blank(forecast_min)
            and is_blank(forecast_value)
            and is_blank(intercept)
            and is_blank(slope)
        ):
            continue

        row_signature = (
            value_signature(intercept),
            value_signature(slope),
            value_signature(forecast_value),
            value_signature(forecast_max),
            value_signature(forecast_min),
        )
        if previous_signature is not None and row_signature == previous_signature:
            continue
        previous_signature = row_signature

        raw_num_quarters = block_value(block, idx, col_map["num_quarters_used"], min_col)
        num_quarters_used = raw_num_quarters if not is_blank(raw_num_quarters) else idx + 1

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": block_value(block, idx, col_map["actual_value"], min_col),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": numeric_difference(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        results.append(row)

    return results


def write_sheet(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(col, "") for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = min(value_len, 60)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 64)


def write_output_workbook(
    output_path: Path, empirical_rows: Sequence[Dict[str, Any]], regression_rows: Sequence[Dict[str, Any]]
) -> None:
    wb_out = Workbook()
    ws_emp = wb_out.active
    ws_emp.title = "empirical_candidates"
    ws_reg = wb_out.create_sheet("regression_candidates")

    write_sheet(ws_emp, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(ws_reg, REGRESSION_HEADERS, regression_rows)
    wb_out.save(output_path)


def should_skip_file(file_path: Path, input_folder_name: str) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary lock file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.match(
        rf"^{re.escape(input_folder_name)}_PARAM(\.\d+)?\.xlsx$",
        file_path.name,
        flags=re.IGNORECASE,
    ):
        return "looks like previously generated output"
    return None


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")
    if not input_path.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {input_path}")

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0
    input_folder_name = input_path.name

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            skip_reason = should_skip_file(file_path, input_folder_name)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            try:
                metadata = parse_file_metadata(file_path)
            except Exception as exc:
                print(f"Skipped {file_path.name}: filename parse error ({exc})")
                continue

            print(f"Processing {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(build_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(build_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: workbook processing error ({exc})")
            finally:
                if wb is not None:
                    safe_close_book(wb)

        final_output = build_output_path(input_path, output_path)
        write_output_workbook(final_output, empirical_rows, regression_rows)

        print(f"Output path: {final_output}")
        print(f"Number of files processed: {processed_files}")
        print(f"Number of empirical rows: {len(empirical_rows)}")
        print(f"Number of regression rows: {len(regression_rows)}")
    finally:
        app.quit()


if __name__ == "__main__":
    main()
