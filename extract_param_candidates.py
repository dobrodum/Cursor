from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -------------------------
# User-configurable paths
# -------------------------
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


# Offsets from the "max" anchor when header labels are not available.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "avg_penetration_pct": -7,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "forecast_value": -3,  # estimated total sold
    "sales_captured_in_db_pct": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}


REGRESSION_OFFSETS = {
    "num_quarters_used": -8,
    "forecast_value": -3,  # TOT FCST w/o SA
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


N_QUARTERS = 10


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def safe_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        if stripped == "" or stripped.startswith("#"):
            return None
        return stripped
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    return value


def to_number(value: Any) -> Optional[float]:
    value = safe_value(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


def subtraction(left: Any, right: Any) -> Optional[float]:
    l_num = to_number(left)
    r_num = to_number(right)
    if l_num is None or r_num is None:
        return None
    return l_num - r_num


def get_unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.resolve().name
    base_name = f"{folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def month_to_number(month_token: str) -> int:
    token = month_token.strip().title()
    if len(token) >= 3:
        token = token[:3]
    return dt.datetime.strptime(token, "%b").month


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    ticker = "UNKNOWN"
    model_period = "Unknown_0000"
    model_date = ""

    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"

    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([0-9]{4})", stem, re.IGNORECASE)
    if period_match:
        period_name, month_name, year = period_match.groups()
        period_name = period_name.title()
        month_short = month_name.strip().title()[:3]
        model_period = f"{period_name}{month_short}_{year}"
        day_by_period = {"early": 5, "mid": 15, "late": 25}
        try:
            month_num = month_to_number(month_short)
            day_num = day_by_period[period_name.lower()]
            model_date = dt.date(int(year), month_num, day_num).isoformat()
        except ValueError:
            model_date = ""

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def close_workbook_no_save(wb: xw.Book) -> None:
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
        pass


def get_sheet_if_exists(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    lookup = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == lookup:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    try:
        found = sheet.api.Cells.Find(What=anchor_text, LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if isinstance(values, (list, tuple)) and values and not isinstance(values[0], (list, tuple)):
        values = [values]
    elif not isinstance(values, (list, tuple)):
        values = [[values]]

    start_row = used.row
    start_col = used.column

    for r_idx, row_values in enumerate(values):
        for c_idx, raw_value in enumerate(row_values):
            if isinstance(raw_value, str) and raw_value.strip().lower() == anchor_text.lower():
                return start_row + r_idx, start_col + c_idx
    return None


def build_header_map(sheet: xw.Sheet, header_row: int, min_col: int, max_col: int) -> Dict[str, int]:
    if max_col < min_col:
        return {}
    row_values = sheet.range((header_row, min_col), (header_row, max_col)).value
    if not isinstance(row_values, (list, tuple)):
        row_values = [row_values]

    header_map: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        normalized = normalize_header(value)
        if normalized:
            header_map[normalized] = min_col + idx
    return header_map


def find_column(
    header_map: Dict[str, int],
    keys: Iterable[str],
    default_col: int,
) -> int:
    normalized_keys = [normalize_header(k) for k in keys]

    for key in normalized_keys:
        if key in header_map:
            return header_map[key]

    for header, col in header_map.items():
        for key in normalized_keys:
            if key and key in header:
                return col

    return default_col


def range_to_grid(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if values and not isinstance(values[0], (list, tuple)):
        return [list(values)]
    return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> List[List[Any]]:
    raw = sheet.range((start_row, start_col), (end_row, end_col)).value
    return range_to_grid(raw)


def block_value(block: List[List[Any]], row_idx: int, col: int, block_start_col: int) -> Any:
    col_idx = col - block_start_col
    if row_idx < 0 or row_idx >= len(block):
        return None
    row = block[row_idx]
    if col_idx < 0 or col_idx >= len(row):
        return None
    return safe_value(row[col_idx])


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def extract_empirical_rows(
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  Skipped empirical extraction in {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    min_col = used.column
    max_col = min_col + used.columns.count - 1
    header_row = anchor_row
    data_start_row = header_row + 1
    data_end_row = data_start_row + (N_QUARTERS - 1)

    header_map = build_header_map(sheet, header_row, min_col, max_col)

    num_quarters_col = find_column(
        header_map,
        ["num_quarters_used", "num_qtrs_used", "num_quarters", "num_qtrs"],
        anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"],
    )
    last_quarter_col = find_column(
        header_map,
        ["last_quarter_used", "last_quarter", "last_qtr_used", "last_qtr"],
        anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"],
    )
    avg_penetration_col = find_column(
        header_map,
        ["avg_penetration_pct", "avg_penetration", "average_penetration"],
        anchor_col + EMPIRICAL_OFFSETS["avg_penetration_pct"],
    )
    quarterly_sales_col = find_column(
        header_map,
        ["quarterly_sales", "qtr_sales", "quarter_sales"],
        anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"],
    )
    reported_sales_col = find_column(
        header_map,
        ["reported_sales", "actual_sales", "actual_value"],
        anchor_col + EMPIRICAL_OFFSETS["reported_sales"],
    )
    growth_rate_col = find_column(
        header_map,
        ["growth_rate_pct", "growth_rate"],
        anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"],
    )
    forecast_value_col = find_column(
        header_map,
        ["estimated_total_sold", "forecast_value", "est_total_sold", "tot_fcst"],
        anchor_col + EMPIRICAL_OFFSETS["forecast_value"],
    )
    captured_in_db_col = find_column(
        header_map,
        ["sales_captured_in_db_pct", "captured_in_db_pct", "db_capture_pct", "penetration_pct"],
        anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
    )
    actual_value_col = find_column(
        header_map,
        ["actual_value", "reported_sales", "actual_sales"],
        anchor_col + EMPIRICAL_OFFSETS["actual_value"],
    )
    forecast_max_col = anchor_col + EMPIRICAL_OFFSETS["forecast_max"]
    forecast_min_col = find_column(
        header_map,
        ["forecast_min", "min"],
        anchor_col + EMPIRICAL_OFFSETS["forecast_min"],
    )

    # Use R1C1 Formula2 once per candidate row, then calculate once.
    if captured_in_db_col > 0 and avg_penetration_col > 0:
        for n_quarters in range(1, N_QUARTERS + 1):
            row_num = data_start_row + (n_quarters - 1)
            start_row = max(data_start_row, row_num - n_quarters + 1)
            formula = f"=AVERAGE(R{start_row}C{captured_in_db_col}:R{row_num}C{captured_in_db_col})"
            set_formula2_r1c1(sheet.range((row_num, avg_penetration_col)), formula)
        sheet.book.app.calculate()

    needed_cols = [
        num_quarters_col,
        last_quarter_col,
        avg_penetration_col,
        quarterly_sales_col,
        reported_sales_col,
        growth_rate_col,
        forecast_value_col,
        captured_in_db_col,
        actual_value_col,
        forecast_max_col,
        forecast_min_col,
    ]
    block_start_col = min(needed_cols)
    block_end_col = max(needed_cols)
    data_block = read_block(sheet, data_start_row, data_end_row, block_start_col, block_end_col)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters_used = block_value(data_block, idx, num_quarters_col, block_start_col)
        last_quarter_used = block_value(data_block, idx, last_quarter_col, block_start_col)
        avg_penetration_pct = block_value(data_block, idx, avg_penetration_col, block_start_col)
        quarterly_sales = block_value(data_block, idx, quarterly_sales_col, block_start_col)
        reported_sales = block_value(data_block, idx, reported_sales_col, block_start_col)
        growth_rate_pct = block_value(data_block, idx, growth_rate_col, block_start_col)
        forecast_value = block_value(data_block, idx, forecast_value_col, block_start_col)
        sales_captured_in_db_pct = block_value(data_block, idx, captured_in_db_col, block_start_col)
        actual_value = block_value(data_block, idx, actual_value_col, block_start_col)
        forecast_max = block_value(data_block, idx, forecast_max_col, block_start_col)
        forecast_min = block_value(data_block, idx, forecast_min_col, block_start_col)

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        if num_quarters_used is None:
            num_quarters_used = idx + 1

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
                "range_width": subtraction(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def normalize_fingerprint(values: Iterable[Any]) -> Tuple[Any, ...]:
    normalized: List[Any] = []
    for value in values:
        number = to_number(value)
        if number is not None:
            normalized.append(round(number, 10))
        else:
            normalized.append(safe_value(value))
    return tuple(normalized)


def extract_regression_rows(
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  Skipped regression extraction in {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    min_col = used.column
    max_col = min_col + used.columns.count - 1
    header_row = anchor_row
    data_start_row = header_row + 1
    data_end_row = data_start_row + (N_QUARTERS - 1)

    header_map = build_header_map(sheet, header_row, min_col, max_col)

    num_quarters_col = find_column(
        header_map,
        ["num_quarters_used", "num_qtrs_used", "num_quarters", "num_qtrs"],
        anchor_col + REGRESSION_OFFSETS["num_quarters_used"],
    )
    forecast_value_col = find_column(
        header_map,
        ["tot_fcst_w_o_sa", "tot_fcst_wo_sa", "tot_fcst_without_sa", "forecast_value"],
        anchor_col + REGRESSION_OFFSETS["forecast_value"],
    )
    actual_value_col = find_column(
        header_map,
        ["actual_value", "actual_sales", "reported_sales"],
        anchor_col + REGRESSION_OFFSETS["actual_value"],
    )
    forecast_max_col = anchor_col + REGRESSION_OFFSETS["forecast_max"]
    forecast_min_col = find_column(
        header_map,
        ["forecast_min", "min"],
        anchor_col + REGRESSION_OFFSETS["forecast_min"],
    )
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    helper_intercept_col = max(max_col + 2, anchor_col + 2)
    helper_slope_col = helper_intercept_col + 1
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        start_row = data_start_row
        end_row = row_num
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2_r1c1(sheet.range((row_num, helper_intercept_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((row_num, helper_slope_col)), slope_formula)
    sheet.book.app.calculate()

    needed_cols = [
        num_quarters_col,
        forecast_value_col,
        actual_value_col,
        forecast_max_col,
        forecast_min_col,
        helper_intercept_col,
        helper_slope_col,
    ]
    block_start_col = min(needed_cols)
    block_end_col = max(needed_cols)
    data_block = read_block(sheet, data_start_row, data_end_row, block_start_col, block_end_col)

    rows: List[Dict[str, Any]] = []
    previous_fingerprint: Optional[Tuple[Any, ...]] = None
    for idx in range(N_QUARTERS):
        num_quarters_used = block_value(data_block, idx, num_quarters_col, block_start_col)
        forecast_value = block_value(data_block, idx, forecast_value_col, block_start_col)
        actual_value = block_value(data_block, idx, actual_value_col, block_start_col)
        forecast_max = block_value(data_block, idx, forecast_max_col, block_start_col)
        forecast_min = block_value(data_block, idx, forecast_min_col, block_start_col)
        intercept = block_value(data_block, idx, helper_intercept_col, block_start_col)
        slope = block_value(data_block, idx, helper_slope_col, block_start_col)

        if all(value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        if num_quarters_used is None:
            num_quarters_used = idx + 1

        fingerprint = normalize_fingerprint(
            [num_quarters_used, intercept, slope, forecast_value, forecast_max, forecast_min]
        )
        if previous_fingerprint is not None and fingerprint == previous_fingerprint:
            continue
        previous_fingerprint = fingerprint

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
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtraction(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def set_column_widths(ws, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = row.get(header)
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 60)


def write_sheet(ws, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    set_column_widths(ws, headers, rows)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def should_skip_file(file_path: Path, in_dir: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary excel file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if file_path.name.startswith(f"{in_dir.resolve().name}_PARAM"):
        return "already an output workbook"
    return None


def main() -> None:
    in_dir = Path(input_dir).resolve()
    out_dir = Path(output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_dir}")
    if not in_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a folder: {in_dir}")

    source_files = sorted(in_dir.iterdir(), key=lambda p: p.name.lower())
    output_path = get_unique_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False
    old_calc_mode = app.calculation
    app.calculation = "manual"

    try:
        for file_path in source_files:
            skip_reason = should_skip_file(file_path, in_dir)
            if skip_reason is not None:
                print(f"SKIPPED: {file_path.name} ({skip_reason})")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                files_processed += 1
                print(f"PROCESSED: {file_path.name}")

                metadata = parse_file_metadata(file_path.name)

                empirical_sheet = get_sheet_if_exists(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"  Skipped empirical extraction in {file_path.name}: sheet not found.")
                else:
                    empirical_rows.extend(extract_empirical_rows(empirical_sheet, metadata, file_path.name))

                regression_sheet = get_sheet_if_exists(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"  Skipped regression extraction in {file_path.name}: sheet not found.")
                else:
                    regression_rows.extend(extract_regression_rows(regression_sheet, metadata, file_path.name))
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (read error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_no_save(wb)
    finally:
        try:
            app.calculation = old_calc_mode
        except Exception:
            pass
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {files_processed}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
