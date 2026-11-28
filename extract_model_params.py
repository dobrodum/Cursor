from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# --------- User-configurable paths ---------
input_dir = Path("./input")
output_dir = Path("./output")
# -------------------------------------------


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"
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


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_number_or_none(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(",", "")
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def month_to_int(mon_abbrev: str) -> int:
    return datetime.strptime(mon_abbrev.title(), "%b").month


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"

    period_token = ""
    if len(parts) >= 3 and parts[2]:
        period_token = parts[2].split("_")[0].strip()

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    if period_match:
        cadence = period_match.group(1).title()
        month_abbrev = period_match.group(2).title()
        year = int(period_match.group(3))
        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        model_period = f"{cadence}{month_abbrev}_{year}"
        model_date = date(year, month_to_int(month_abbrev), day_map[cadence]).isoformat()
    else:
        sanitized = re.sub(r"[^A-Za-z0-9]+", "_", period_token).strip("_")
        model_period = sanitized or "UnknownPeriod"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def output_path_for_run(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        numbered = output_folder / f"{base}.{idx}.xlsx"
        if not numbered.exists():
            return numbered
        idx += 1


def safe_close_workbook(wb: Any) -> None:
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

    # Last-resort COM close fallback.
    try:
        wb.api.Close(SaveChanges=False)
    except TypeError:
        wb.api.Close(False)


def get_sheet_or_none(wb: Any, sheet_name: str) -> Optional[Any]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def find_anchor_cell(sheet: Any, token: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = used.value
    if matrix is None:
        return None

    if isinstance(matrix, list):
        if matrix and not isinstance(matrix[0], list):
            matrix_2d = [matrix]
        else:
            matrix_2d = matrix
    else:
        matrix_2d = [[matrix]]

    token_norm = normalize_label(token)
    base_row = used.row
    base_col = used.column

    for r_idx, row in enumerate(matrix_2d):
        if not isinstance(row, list):
            row = [row]
        for c_idx, cell_val in enumerate(row):
            if normalize_label(cell_val) == token_norm:
                return base_row + r_idx, base_col + c_idx
    return None


def build_header_map(sheet: Any, row: int, start_col: int, end_col: int) -> Dict[str, int]:
    start_col = max(1, start_col)
    values = sheet.range((row, start_col), (row, end_col)).value

    if not isinstance(values, list):
        values = [values]

    header_map: Dict[str, int] = {}
    for offset, raw in enumerate(values):
        norm = normalize_label(raw)
        if norm:
            header_map[norm] = start_col + offset
    return header_map


def find_header_col(header_map: Dict[str, int], aliases: Iterable[str]) -> Optional[int]:
    for alias in aliases:
        alias_norm = normalize_label(alias)
        if alias_norm in header_map:
            return header_map[alias_norm]

    for alias in aliases:
        alias_norm = normalize_label(alias)
        for key, col in header_map.items():
            if alias_norm and alias_norm in key:
                return col
    return None


def find_last_numeric_pair_row(
    sheet: Any,
    top_row: int,
    bottom_row: int,
    col_a: int,
    col_b: int,
) -> Optional[int]:
    if top_row > bottom_row:
        return None
    for row in range(bottom_row, top_row - 1, -1):
        a = sheet.cells(row, col_a).value
        b = sheet.cells(row, col_b).value
        if is_number(a) and is_number(b):
            return row
    return None


def read_cell(sheet: Any, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1 or row < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def calc_range_width(max_val: Any, min_val: Any) -> Optional[float]:
    max_num = to_number_or_none(max_val)
    min_num = to_number_or_none(min_val)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def extract_empirical_rows(
    wb: Any,
    metadata: FileMetadata,
    source_file_name: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_or_none(wb, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    headers = build_header_map(sheet, anchor_row, anchor_col - 25, anchor_col + 10)

    max_col = anchor_col
    min_col = find_header_col(headers, ["min"]) or (anchor_col + 1)
    num_quarters_col = find_header_col(headers, ["num quarters used", "num quarters", "n quarters"])
    last_quarter_col = find_header_col(headers, ["last quarter used", "last quarter"])
    reported_sales_col = find_header_col(headers, ["reported sales", "actual sales", "actual"])
    growth_col = find_header_col(headers, ["growth rate", "growth"])
    sales_captured_col = find_header_col(headers, ["sales captured in db", "sales captured", "captured in db"])
    quarterly_sales_col = find_header_col(headers, ["quarterly sales", "sales", "tot fcst w/o sa"]) or (anchor_col - 7)
    penetration_col = find_header_col(headers, ["penetration", "captured in db"]) or (anchor_col - 11)

    data_end_row = find_last_numeric_pair_row(
        sheet=sheet,
        top_row=1,
        bottom_row=anchor_row - 1,
        col_a=penetration_col,
        col_b=quarterly_sales_col,
    )
    if data_end_row is None:
        return []

    helper_col = anchor_col + 20
    formula_rows: List[Tuple[int, int]] = []

    for n in range(1, N_QUARTERS + 1):
        start_row = data_end_row - n + 1
        if start_row < 1:
            break
        helper_row = anchor_row + n

        # Temporary formulas (not saved) to force Excel-native calculations quickly.
        avg_pen_formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{data_end_row}C{penetration_col})"
        )
        forecast_formula = f'=IFERROR(R{data_end_row}C{quarterly_sales_col}/RC[-1],"")'
        max_formula = (
            f'=IFERROR(R{data_end_row}C{quarterly_sales_col}/MAX('
            f"R{start_row}C{penetration_col}:R{data_end_row}C{penetration_col}"
            '),"")'
        )
        min_formula = (
            f'=IFERROR(R{data_end_row}C{quarterly_sales_col}/MIN('
            f"R{start_row}C{penetration_col}:R{data_end_row}C{penetration_col}"
            '),"")'
        )

        sheet.cells(helper_row, helper_col).formula2 = avg_pen_formula
        sheet.cells(helper_row, helper_col + 1).formula2 = forecast_formula
        sheet.cells(helper_row, helper_col + 2).formula2 = max_formula
        sheet.cells(helper_row, helper_col + 3).formula2 = min_formula
        formula_rows.append((n, helper_row))

    if formula_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n, helper_row in formula_rows:
        candidate_row = anchor_row + n
        avg_pen = read_cell(sheet, helper_row, helper_col)
        forecast_value = read_cell(sheet, helper_row, helper_col + 1)
        fallback_max = read_cell(sheet, helper_row, helper_col + 2)
        fallback_min = read_cell(sheet, helper_row, helper_col + 3)

        forecast_max = read_cell(sheet, candidate_row, max_col)
        if to_number_or_none(forecast_max) is None:
            forecast_max = fallback_max

        forecast_min = read_cell(sheet, candidate_row, min_col)
        if to_number_or_none(forecast_min) is None:
            forecast_min = fallback_min

        quarterly_sales = read_cell(sheet, data_end_row, quarterly_sales_col)
        reported_sales = read_cell(sheet, candidate_row, reported_sales_col)
        actual_value = reported_sales
        growth_rate_pct = read_cell(sheet, candidate_row, growth_col)
        sales_captured_in_db_pct = read_cell(sheet, candidate_row, sales_captured_col)
        if sales_captured_in_db_pct is None:
            sales_captured_in_db_pct = read_cell(sheet, data_end_row, penetration_col)

        num_quarters_used = read_cell(sheet, candidate_row, num_quarters_col)
        if to_number_or_none(num_quarters_used) is None:
            num_quarters_used = n

        last_quarter_used = read_cell(sheet, candidate_row, last_quarter_col)
        if last_quarter_used is None:
            last_quarter_used = read_cell(sheet, data_end_row, max(1, penetration_col - 1))

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file_name,
        }
        rows.append(row)

    return rows


def extract_regression_rows(
    wb: Any,
    metadata: FileMetadata,
    source_file_name: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_or_none(wb, REGRESSION_SHEET_NAME)
    if sheet is None:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    headers = build_header_map(sheet, anchor_row, anchor_col - 25, anchor_col + 10)
    max_col = anchor_col
    min_col = find_header_col(headers, ["min"]) or (anchor_col + 1)
    num_quarters_col = find_header_col(headers, ["num quarters used", "num quarters", "n quarters"])
    forecast_col = find_header_col(headers, ["tot fcst w/o sa", "forecast w/o sa", "forecast"])
    intercept_col = find_header_col(headers, ["intercept"])
    slope_col = find_header_col(headers, ["slope"])
    actual_col = find_header_col(headers, ["actual", "reported sales", "actual sales"])

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_end_row = find_last_numeric_pair_row(
        sheet=sheet,
        top_row=1,
        bottom_row=anchor_row - 1,
        col_a=x_col,
        col_b=y_col,
    )
    if data_end_row is None:
        return []

    forecast_x_row = anchor_row if is_number(read_cell(sheet, anchor_row, x_col)) else data_end_row

    helper_col = anchor_col + 30
    formula_rows: List[Tuple[int, int]] = []
    for n in range(1, N_QUARTERS + 1):
        start_row = data_end_row - n + 1
        if start_row < 1:
            break

        helper_row = anchor_row + n
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        forecast_formula = f"=RC[-2]+(RC[-1]*R{forecast_x_row}C{x_col})"

        sheet.cells(helper_row, helper_col).formula2 = intercept_formula
        sheet.cells(helper_row, helper_col + 1).formula2 = slope_formula
        sheet.cells(helper_row, helper_col + 2).formula2 = forecast_formula
        formula_rows.append((n, helper_row))

    if formula_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for n, helper_row in formula_rows:
        candidate_row = anchor_row + n

        num_quarters_used = read_cell(sheet, candidate_row, num_quarters_col)
        if to_number_or_none(num_quarters_used) is None:
            num_quarters_used = n

        intercept = read_cell(sheet, candidate_row, intercept_col)
        if to_number_or_none(intercept) is None:
            intercept = read_cell(sheet, helper_row, helper_col)

        slope = read_cell(sheet, candidate_row, slope_col)
        if to_number_or_none(slope) is None:
            slope = read_cell(sheet, helper_row, helper_col + 1)

        forecast_value = read_cell(sheet, candidate_row, forecast_col)
        if to_number_or_none(forecast_value) is None:
            forecast_value = read_cell(sheet, helper_row, helper_col + 2)

        forecast_max = read_cell(sheet, candidate_row, max_col)
        forecast_min = read_cell(sheet, candidate_row, min_col)
        actual_value = read_cell(sheet, candidate_row, actual_col)
        range_width = calc_range_width(forecast_max, forecast_min)

        signature = (
            round(to_number_or_none(num_quarters_used) or 0.0, 10),
            round(to_number_or_none(intercept) or 0.0, 10),
            round(to_number_or_none(slope) or 0.0, 10),
            round(to_number_or_none(forecast_value) or 0.0, 10),
            round(to_number_or_none(forecast_max) or 0.0, 10),
            round(to_number_or_none(forecast_min) or 0.0, 10),
        )
        if prev_signature == signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file_name,
            }
        )

    return rows


def write_sheet(ws: Any, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h) for h in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None:
                continue
            val_len = len(str(val))
            if val_len > max_len:
                max_len = val_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = EMPIRICAL_OUTPUT_SHEET
    ws_regression = wb.create_sheet(REGRESSION_OUTPUT_SHEET)

    write_sheet(ws_empirical, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_HEADERS, regression_rows)
    wb.save(output_path)


def should_skip(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temporary file"
    if path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    return None


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    out_path = output_path_for_run(input_folder=input_dir, output_folder=output_dir)
    source_files = sorted(input_dir.iterdir())

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        try:
            app.calculation = "manual"
        except Exception:
            # Not all Excel backends expose this cleanly.
            pass

        for file_path in source_files:
            reason = should_skip(file_path)
            if reason:
                print(f"Skipped: {file_path.name} ({reason})")
                continue

            print(f"Processing: {file_path.name}")
            metadata = parse_file_metadata(file_path)
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file_name=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file_name=file_path.name,
                    )
                )
                processed_count += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)

        write_output_workbook(
            output_path=out_path,
            empirical_rows=empirical_rows,
            regression_rows=regression_rows,
        )
    finally:
        app.quit()

    print(f"Output path: {out_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
