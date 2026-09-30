from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Font
import xlwings as xw


# ========= USER CONFIG =========
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")
# ===============================


N_QUARTERS = 10
SCAN_COLS_LEFT = 24
SCAN_COLS_RIGHT = 24


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
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def to_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, tuple):
        values = list(values)
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], (list, tuple)):
            return [list(row) if isinstance(row, tuple) else row for row in values]
        return [values]
    return [[values]]


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def coerce_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        return float(value)
    except Exception:
        return None


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    a_num = coerce_float(a)
    b_num = coerce_float(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def parse_month_token(month_token: str) -> Tuple[str, int]:
    token = month_token.strip().lower()
    month_map = {
        "jan": ("Jan", 1),
        "january": ("Jan", 1),
        "feb": ("Feb", 2),
        "february": ("Feb", 2),
        "mar": ("Mar", 3),
        "march": ("Mar", 3),
        "apr": ("Apr", 4),
        "april": ("Apr", 4),
        "may": ("May", 5),
        "jun": ("Jun", 6),
        "june": ("Jun", 6),
        "jul": ("Jul", 7),
        "july": ("Jul", 7),
        "aug": ("Aug", 8),
        "august": ("Aug", 8),
        "sep": ("Sep", 9),
        "sept": ("Sep", 9),
        "september": ("Sep", 9),
        "oct": ("Oct", 10),
        "october": ("Oct", 10),
        "nov": ("Nov", 11),
        "november": ("Nov", 11),
        "dec": ("Dec", 12),
        "december": ("Dec", 12),
    }
    if token in month_map:
        return month_map[token]
    for key, value in month_map.items():
        if token.startswith(key):
            return value
    raise ValueError(f"Unsupported month token: {month_token}")


def parse_filename_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    # Example: MedMiner_Model - AORT - MidJan2026_Send
    pattern = re.compile(
        r".*?-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        pieces = [p.strip() for p in stem.split(" - ")]
        ticker = pieces[1].upper() if len(pieces) >= 2 else ""
        return FileLabel(
            model=ticker or stem,
            ticker=ticker,
            model_period="",
            model_date="",
        )

    ticker = match.group(1).upper()
    period_bucket = match.group(2).title()
    month_token = match.group(3)
    year = int(match.group(4))
    month_label, month_num = parse_month_token(month_token)

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_map[period_bucket]
    model_period = f"{period_bucket}{month_label}_{year}"
    model = f"{ticker}_{model_period}"
    model_date = date(year, month_num, day).isoformat()

    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    folder_name = in_dir.name
    base = out_dir / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    index = 1
    while True:
        candidate = out_dir / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    used_values = to_2d(used.value)
    if not used_values:
        return None

    base_row = used.row
    base_col = used.column
    for r_idx, row_values in enumerate(used_values):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return base_row + r_idx, base_col + c_idx
    return None


def build_header_map(
    sheet: xw.Sheet, header_row: int, min_col: int, max_col: int
) -> Dict[str, int]:
    row_values = sheet.range((header_row, min_col), (header_row, max_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]
    header_map: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        key = normalize_text(value)
        if key and key not in header_map:
            header_map[key] = min_col + idx
    return header_map


def find_col(
    header_map: Dict[str, int], patterns: Sequence[str], fallback: Optional[int] = None
) -> Optional[int]:
    keys = list(header_map.keys())
    for pattern in patterns:
        target = normalize_text(pattern)
        for key in keys:
            if target and target in key:
                return header_map[key]
    return fallback


def read_block(
    sheet: xw.Sheet, row_start: int, row_end: int, col_start: int, col_end: int
) -> List[List[Any]]:
    values = sheet.range((row_start, col_start), (row_end, col_end)).value
    return to_2d(values)


def block_get(
    block: List[List[Any]], row: int, col: int, row_start: int, col_start: int
) -> Any:
    r_idx = row - row_start
    c_idx = col - col_start
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(block):
        return None
    if c_idx >= len(block[r_idx]):
        return None
    return block[r_idx][c_idx]


def safely_close_workbook(wb: xw.Book) -> None:
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
            wb.close()
        except Exception:
            pass


def process_empirical_sheet(
    wb: xw.Book, sheet: xw.Sheet, label: FileLabel, source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if not anchor:
        print("  - Empirical Model: skipped (no 'max' anchor found)")
        return []

    anchor_row, anchor_col = anchor
    min_col = max(1, anchor_col - SCAN_COLS_LEFT)
    max_col = anchor_col + SCAN_COLS_RIGHT
    header_map = build_header_map(sheet, anchor_row, min_col, max_col)

    forecast_max_col = anchor_col
    forecast_min_col = find_col(header_map, ["min"], fallback=anchor_col + 1)
    forecast_value_col = find_col(
        header_map,
        ["estimatedtotalsold", "forecastvalue", "estimate", "totalforecast", "totfcst"],
        fallback=anchor_col - 1,
    )
    reported_sales_col = find_col(
        header_map,
        ["reportedsales", "actualsales", "reported", "actualvalue", "actual"],
        fallback=anchor_col - 2,
    )
    quarterly_sales_col = find_col(
        header_map,
        ["quarterlysales", "qtrsales", "quartersales"],
        fallback=anchor_col - 3,
    )
    growth_rate_col = find_col(
        header_map,
        ["growthrate", "growthpct"],
        fallback=anchor_col - 4,
    )
    captured_col = find_col(
        header_map,
        ["capturedindb", "salescapturedindb", "dbcaptured", "capturedpct"],
        fallback=anchor_col - 5,
    )
    avg_penetration_source_col = find_col(
        header_map,
        ["penetrationpct", "avgpenetration", "penetration"],
        fallback=anchor_col - 6,
    )
    num_quarters_col = find_col(
        header_map,
        ["numquartersused", "numquarters", "quartersused", "nquarters"],
        fallback=None,
    )
    last_quarter_col = find_col(
        header_map,
        ["lastquarterused", "lastquarter", "lastqtr"],
        fallback=None,
    )

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1
    block = read_block(sheet, data_start_row, data_end_row, min_col, max_col)

    avg_penetration_values: List[Any] = [None] * N_QUARTERS
    if avg_penetration_source_col:
        scratch_col = max_col + 2
        scratch_range = sheet.range((data_start_row, scratch_col), (data_end_row, scratch_col))
        formulas = [
            [
                f"=AVERAGE(R{data_start_row}C{avg_penetration_source_col}:R{data_start_row + i}C{avg_penetration_source_col})"
            ]
            for i in range(N_QUARTERS)
        ]
        scratch_range.formula2 = formulas
        wb.app.calculate()
        raw_avg_values = to_2d(scratch_range.value)
        for i in range(min(N_QUARTERS, len(raw_avg_values))):
            avg_penetration_values[i] = raw_avg_values[i][0] if raw_avg_values[i] else None

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row_num = data_start_row + i
        num_quarters_used = (
            block_get(block, row_num, num_quarters_col, data_start_row, min_col)
            if num_quarters_col
            else i + 1
        )
        if is_blank(num_quarters_used):
            num_quarters_used = i + 1

        last_quarter_used = (
            block_get(block, row_num, last_quarter_col, data_start_row, min_col)
            if last_quarter_col
            else None
        )

        forecast_value = block_get(block, row_num, forecast_value_col, data_start_row, min_col)
        reported_sales = block_get(block, row_num, reported_sales_col, data_start_row, min_col)
        forecast_max = block_get(block, row_num, forecast_max_col, data_start_row, min_col)
        forecast_min = block_get(block, row_num, forecast_min_col, data_start_row, min_col)
        quarterly_sales = block_get(block, row_num, quarterly_sales_col, data_start_row, min_col)
        growth_rate_pct = block_get(block, row_num, growth_rate_col, data_start_row, min_col)
        captured_pct = block_get(block, row_num, captured_col, data_start_row, min_col)

        avg_penetration_pct = avg_penetration_values[i]
        if is_blank(avg_penetration_pct):
            avg_penetration_pct = block_get(
                block, row_num, avg_penetration_source_col, data_start_row, min_col
            )

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                reported_sales,
                quarterly_sales,
                avg_penetration_pct,
            )
        ):
            continue

        range_width = safe_subtract(forecast_max, forecast_min)
        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def rows_equal_signature(left_row: Sequence[Any], right_row: Sequence[Any]) -> bool:
    if len(left_row) != len(right_row):
        return False
    for left, right in zip(left_row, right_row):
        left_num = coerce_float(left)
        right_num = coerce_float(right)
        if left_num is not None and right_num is not None:
            if abs(left_num - right_num) > 1e-9:
                return False
        else:
            if left != right:
                return False
    return True


def find_last_numeric_row(sheet: xw.Sheet, col: int, max_row: int) -> Optional[int]:
    if max_row < 1:
        return None
    values = sheet.range((1, col), (max_row, col)).value
    if not isinstance(values, list):
        values = [values]
    for idx in range(len(values) - 1, -1, -1):
        if coerce_float(values[idx]) is not None:
            return idx + 1
    return None


def process_regression_sheet(
    wb: xw.Book, sheet: xw.Sheet, label: FileLabel, source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if not anchor:
        print("  - Regression Model: skipped (no 'max' anchor found)")
        return []

    anchor_row, anchor_col = anchor
    min_col = max(1, anchor_col - SCAN_COLS_LEFT)
    max_col = anchor_col + SCAN_COLS_RIGHT
    header_map = build_header_map(sheet, anchor_row, min_col, max_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    forecast_max_col = anchor_col
    forecast_min_col = find_col(header_map, ["min"], fallback=anchor_col + 1)
    forecast_col = find_col(
        header_map,
        ["totfcstwosa", "totalforecastwithoutsa", "forecastvalue", "totfcst", "forecast"],
        fallback=anchor_col - 1,
    )
    actual_col = find_col(
        header_map,
        ["actualvalue", "actualsales", "actual", "reported"],
        fallback=None,
    )
    num_quarters_col = find_col(
        header_map,
        ["numquartersused", "numquarters", "quartersused", "nquarters"],
        fallback=None,
    )

    history_end_row = find_last_numeric_row(sheet, y_col, anchor_row - 1)
    x_last_value = sheet.range((history_end_row, x_col)).value if history_end_row else None
    x_last_numeric = coerce_float(x_last_value)

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1
    block = read_block(sheet, data_start_row, data_end_row, min_col, max_col)

    scratch_col = max_col + 2
    intercept_cells = [sheet.range((data_start_row + i, scratch_col)) for i in range(N_QUARTERS)]
    slope_cells = [sheet.range((data_start_row + i, scratch_col + 1)) for i in range(N_QUARTERS)]

    wrote_regression_formulas = False
    for i in range(N_QUARTERS):
        n_quarters = i + 1
        if history_end_row is None:
            continue
        start_row = history_end_row - n_quarters + 1
        if start_row < 1:
            continue
        intercept_cells[i].formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        slope_cells[i].formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        wrote_regression_formulas = True

    if wrote_regression_formulas:
        wb.app.calculate()

    intercept_values = [cell.value for cell in intercept_cells]
    slope_values = [cell.value for cell in slope_cells]

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(N_QUARTERS):
        n_quarters = i + 1
        row_num = data_start_row + i
        intercept_value = intercept_values[i]
        slope_value = slope_values[i]

        forecast_value = block_get(block, row_num, forecast_col, data_start_row, min_col)
        if is_blank(forecast_value):
            intercept_num = coerce_float(intercept_value)
            slope_num = coerce_float(slope_value)
            if intercept_num is not None and slope_num is not None and x_last_numeric is not None:
                forecast_value = intercept_num + (slope_num * (x_last_numeric + 1))

        forecast_max = block_get(block, row_num, forecast_max_col, data_start_row, min_col)
        forecast_min = block_get(block, row_num, forecast_min_col, data_start_row, min_col)
        actual_value = (
            block_get(block, row_num, actual_col, data_start_row, min_col) if actual_col else None
        )
        num_quarters_used = (
            block_get(block, row_num, num_quarters_col, data_start_row, min_col)
            if num_quarters_col
            else n_quarters
        )
        if is_blank(num_quarters_used):
            num_quarters_used = n_quarters

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        ):
            continue

        signature = (
            forecast_value,
            forecast_max,
            forecast_min,
            intercept_value,
            slope_value,
        )
        if previous_signature and rows_equal_signature(signature, previous_signature):
            # Prevent duplicate final row when the last two parameterizations are identical.
            continue

        range_width = safe_subtract(forecast_max, forecast_min)
        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )
        previous_signature = signature

    return rows


def write_rows_to_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    headers: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_cells in ws.columns:
        header = col_cells[0].value or ""
        max_len = len(str(header))
        for cell in col_cells[1:]:
            if cell.value is None:
                continue
            value_len = len(str(cell.value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[col_cells[0].column_letter].width = min(60, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_rows_to_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_rows_to_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)
    wb.save(output_path)


def list_input_files(folder: Path) -> List[Path]:
    return sorted([path for path in folder.iterdir() if path.is_file()], key=lambda p: p.name.lower())


def run() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    files = list_input_files(input_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in files:
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not .xlsx")
                continue

            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_filename_label(file_path)
                sheet_names = {sheet.name for sheet in wb.sheets}

                if "Empirical Model" in sheet_names:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            wb=wb,
                            sheet=wb.sheets["Empirical Model"],
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print("  - Empirical Model: skipped (sheet missing)")

                if "Regression Model" in sheet_names:
                    regression_rows.extend(
                        process_regression_sheet(
                            wb=wb,
                            sheet=wb.sheets["Regression Model"],
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print("  - Regression Model: skipped (sheet missing)")

                processed_files += 1
            except Exception as exc:
                print(f"  - Failed {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safely_close_workbook(wb)
    finally:
        app.quit()

    output_path = unique_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
