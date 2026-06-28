from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# User-editable paths
input_dir = Path("input")
output_dir = Path("output")


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
MONTH_TO_NUMBER = {
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


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetScan:
    start_row: int
    start_col: int
    last_row: int
    last_col: int
    values: List[List[Any]]
    text_cells: List[Tuple[int, int, str]]


def normalize_used_range(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    normalized = [list(row) if isinstance(row, list) else [row] for row in values]
    max_cols = max((len(row) for row in normalized), default=0)
    for row in normalized:
        if len(row) < max_cols:
            row.extend([None] * (max_cols - len(row)))
    return normalized


def scan_sheet_once(sheet: xw.Sheet) -> SheetScan:
    used = sheet.used_range
    values = normalize_used_range(used.value)
    row_count = len(values)
    col_count = max((len(row) for row in values), default=0)
    start_row = used.row
    start_col = used.column
    last_row = start_row + max(row_count - 1, 0)
    last_col = start_col + max(col_count - 1, 0)

    text_cells: List[Tuple[int, int, str]] = []
    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip():
                text_cells.append((start_row + r_idx, start_col + c_idx, value.strip().lower()))

    return SheetScan(
        start_row=start_row,
        start_col=start_col,
        last_row=last_row,
        last_col=last_col,
        values=values,
        text_cells=text_cells,
    )


def scan_value(scan: SheetScan, row: int, col: int) -> Any:
    r_idx = row - scan.start_row
    c_idx = col - scan.start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(scan.values):
        return None
    if c_idx >= len(scan.values[r_idx]):
        return None
    return scan.values[r_idx][c_idx]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def rounded_signature(values: Sequence[Optional[float]], digits: int = 10) -> Tuple[Optional[float], ...]:
    result: List[Optional[float]] = []
    for value in values:
        if value is None:
            result.append(None)
        else:
            result.append(round(value, digits))
    return tuple(result)


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    period_token = ""
    if len(parts) >= 3:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        period_token = parts[2]
    else:
        fallback = re.search(
            r"-\s*([A-Za-z0-9]+)\s*-\s*((?:Early|Mid|Late)[A-Za-z]{3,9}\d{4})",
            stem,
            flags=re.IGNORECASE,
        )
        if fallback:
            ticker = fallback.group(1).upper()
            period_token = fallback.group(2)

    if not ticker:
        raise ValueError("ticker not found in filename")

    period_token = re.sub(r"[_\-\s]*send$", "", period_token, flags=re.IGNORECASE).strip()
    match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_\-]?\s*(\d{4})", period_token, flags=re.IGNORECASE)
    if not match:
        raise ValueError("model period not found in filename")

    period_text = match.group(1).title()
    month_text = match.group(2).title()
    year = int(match.group(3))
    month_key = month_text[:3].lower()
    if month_key not in MONTH_TO_NUMBER:
        raise ValueError(f"unsupported month token: {month_text}")

    day = DAY_BY_PERIOD[period_text.lower()]
    month_num = MONTH_TO_NUMBER[month_key]
    model_period = f"{period_text}{month_text[:3]}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def choose_output_path(input_path: Path, target_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    counter = 1
    while True:
        candidate = target_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


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
        # Last-resort fallback: close via app context if direct close fails.
        wb.app.api.DisplayAlerts = False
        wb.api.Close(SaveChanges=False)


def find_anchor(scan: SheetScan, anchor_label: str = "max") -> Optional[Tuple[int, int]]:
    candidates = [(r, c) for r, c, text in scan.text_cells if text == anchor_label]
    if not candidates:
        return None

    def score(pos: Tuple[int, int]) -> Tuple[int, int]:
        row, col = pos
        right_value = scan_value(scan, row, col + 1)
        is_numeric = 0 if to_float(right_value) is not None else 1
        return (is_numeric, abs(scan.last_col - col))

    return sorted(candidates, key=score)[0]


def find_label_cell(
    scan: SheetScan,
    keyword_groups: Iterable[Iterable[str]],
    anchor_col: int,
) -> Optional[Tuple[int, int]]:
    matches: List[Tuple[int, int, int]] = []
    for row, col, text in scan.text_cells:
        for keyword_group in keyword_groups:
            lowered = [word.lower() for word in keyword_group]
            if all(word in text for word in lowered):
                matches.append((abs(col - anchor_col), row, col))
                break
    if not matches:
        return None
    matches.sort()
    _, row, col = matches[0]
    return row, col


def find_column(
    scan: SheetScan,
    keyword_groups: Iterable[Iterable[str]],
    anchor_col: int,
    fallback_col: int,
) -> int:
    label_cell = find_label_cell(scan, keyword_groups, anchor_col)
    if label_cell is None:
        return fallback_col
    _, col = label_cell
    return col


def numeric_rows_for_column(scan: SheetScan, column: int) -> List[Tuple[int, float]]:
    rows: List[Tuple[int, float]] = []
    for row in range(scan.start_row, scan.last_row + 1):
        numeric_value = to_float(scan_value(scan, row, column))
        if numeric_value is not None:
            rows.append((row, numeric_value))
    return rows


def get_sheet_if_exists(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sheet in wb.sheets:
        if sheet.name == sheet_name:
            return sheet
    return None


def get_live_value(sheet: xw.Sheet, row: int, col: int, offset: int = 1) -> Any:
    return sheet.cells(row, col + offset).value


def process_empirical_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    scan = scan_sheet_once(sheet)
    anchor = find_anchor(scan, "max")
    if anchor is None:
        print(f"Skipping empirical extraction for {source_file}: max anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    penetration_col = find_column(scan, [("penetration",)], anchor_col, anchor_col - 11)
    quarter_col = find_column(scan, [("quarter",)], anchor_col, penetration_col - 1)
    quarterly_sales_col = find_column(
        scan,
        [("quarterly", "sales"), ("db", "sales"), ("database", "sales")],
        anchor_col,
        anchor_col - 7,
    )
    reported_sales_col = find_column(scan, [("reported", "sales"), ("actual", "sales")], anchor_col, anchor_col - 6)
    growth_rate_col = find_column(scan, [("growth", "rate"), ("growth",)], anchor_col, anchor_col - 5)
    sales_captured_col = find_column(
        scan,
        [("captured", "db"), ("sales", "captured"), ("penetration", "pct")],
        anchor_col,
        penetration_col,
    )

    penetration_rows = numeric_rows_for_column(scan, penetration_col)
    if not penetration_rows:
        print(f"Skipping empirical extraction for {source_file}: no numeric penetration data")
        return rows

    est_total_label = find_label_cell(
        scan,
        [
            ("estimated", "total", "sold"),
            ("tot", "fcst", "w/o", "sa"),
            ("tot", "forecast"),
        ],
        anchor_col,
    )
    reported_label = find_label_cell(scan, [("reported", "sales"), ("actual", "sales")], anchor_col)

    scratch_col = scan.last_col + 2
    scratch_row = anchor_row
    avg_formula_cell = sheet.cells(scratch_row, scratch_col)

    max_iterations = 10
    available_quarters = len(penetration_rows)
    for num_quarters in range(1, max_iterations + 1):
        if num_quarters > available_quarters:
            break

        selected_rows = penetration_rows[-num_quarters:]
        start_row = selected_rows[0][0]
        end_row = selected_rows[-1][0]

        avg_formula_cell.formula2 = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        wb.app.calculate()

        avg_penetration_pct = to_float(avg_formula_cell.value)
        forecast_max = to_float(sheet.cells(anchor_row, anchor_col + 1).value)
        forecast_min = to_float(sheet.cells(anchor_row + 1, anchor_col + 1).value)
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        quarterly_sales = to_float(sheet.cells(end_row, quarterly_sales_col).value)
        reported_sales = to_float(sheet.cells(end_row, reported_sales_col).value)
        if reported_sales is None and reported_label is not None:
            reported_sales = to_float(get_live_value(sheet, reported_label[0], reported_label[1], offset=1))

        growth_rate_pct = to_float(sheet.cells(end_row, growth_rate_col).value)
        sales_captured_in_db_pct = to_float(sheet.cells(end_row, sales_captured_col).value)
        if sales_captured_in_db_pct is None:
            sales_captured_in_db_pct = avg_penetration_pct

        forecast_value = None
        if est_total_label is not None:
            forecast_value = to_float(get_live_value(sheet, est_total_label[0], est_total_label[1], offset=1))

        if forecast_value is None and avg_penetration_pct not in (None, 0) and quarterly_sales is not None:
            penetration_ratio = avg_penetration_pct / 100 if avg_penetration_pct > 1 else avg_penetration_pct
            if penetration_ratio not in (None, 0):
                forecast_value = quarterly_sales / penetration_ratio

        last_quarter_used = sheet.cells(end_row, quarter_col).value
        actual_value = reported_sales

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    avg_formula_cell.value = None
    return rows


def process_regression_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    scan = scan_sheet_once(sheet)
    anchor = find_anchor(scan, "max")
    if anchor is None:
        print(f"Skipping regression extraction for {source_file}: max anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    numeric_pairs: List[Tuple[int, float, float]] = []
    for row in range(scan.start_row, scan.last_row + 1):
        y_value = to_float(scan_value(scan, row, y_col))
        x_value = to_float(scan_value(scan, row, x_col))
        if y_value is not None and x_value is not None:
            numeric_pairs.append((row, x_value, y_value))

    if len(numeric_pairs) < 2:
        print(f"Skipping regression extraction for {source_file}: insufficient x/y data")
        return rows

    actual_label = find_label_cell(scan, [("actual", "sales"), ("reported", "sales")], anchor_col)
    fcst_label = find_label_cell(
        scan,
        [
            ("tot", "fcst", "w/o", "sa"),
            ("total", "forecast", "without", "sa"),
            ("tot", "forecast"),
        ],
        anchor_col,
    )

    scratch_col = scan.last_col + 2
    scratch_row = anchor_row
    intercept_cell = sheet.cells(scratch_row, scratch_col)
    slope_cell = sheet.cells(scratch_row, scratch_col + 1)
    forecast_cell = sheet.cells(scratch_row, scratch_col + 2)

    previous_signature: Optional[Tuple[Optional[float], ...]] = None
    max_iterations = min(10, len(numeric_pairs))
    for num_quarters in range(2, max_iterations + 1):
        selected_rows = numeric_pairs[-num_quarters:]
        start_row = selected_rows[0][0]
        end_row = selected_rows[-1][0]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        forecast_cell.formula2 = f"=R{end_row}C{x_col}*R{scratch_row}C{scratch_col + 1}+R{scratch_row}C{scratch_col}"
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_total_without_sa = None
        if fcst_label is not None:
            forecast_total_without_sa = to_float(get_live_value(sheet, fcst_label[0], fcst_label[1], offset=1))
        if forecast_total_without_sa is None:
            forecast_total_without_sa = to_float(forecast_cell.value)

        forecast_max = to_float(sheet.cells(anchor_row, anchor_col + 1).value)
        forecast_min = to_float(sheet.cells(anchor_row + 1, anchor_col + 1).value)
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        actual_value = None
        if actual_label is not None:
            actual_value = to_float(get_live_value(sheet, actual_label[0], actual_label[1], offset=1))

        current_signature = rounded_signature(
            [intercept, slope, forecast_total_without_sa, forecast_max, forecast_min, range_width]
        )
        if previous_signature is not None and current_signature == previous_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    intercept_cell.value = None
    slope_cell.value = None
    forecast_cell.value = None
    return rows


def auto_width(worksheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row in rows:
            value = row.get(col_name)
            as_text = "" if value is None else str(value)
            if len(as_text) > max_len:
                max_len = len(as_text)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 50)


def write_candidate_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_row = max(1, len(rows) + 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{last_row}"
    auto_width(ws, columns, rows)


def create_output_workbook(
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
    out_path: Path,
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    write_candidate_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_candidate_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(out_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = choose_output_path(input_dir, output_dir)

    source_files: List[Path] = []
    for entry in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
        if not entry.is_file():
            continue
        if entry.name.startswith("~"):
            print(f"Skipping {entry.name}: temporary file")
            continue
        if entry.suffix.lower() != ".xlsx":
            print(f"Skipping {entry.name}: not an .xlsx file")
            continue
        source_files.append(entry)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"Processing: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                metadata = parse_file_metadata(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)

                empirical_sheet = get_sheet_if_exists(wb, "Empirical Model")
                regression_sheet = get_sheet_if_exists(wb, "Regression Model")

                if empirical_sheet is None and regression_sheet is None:
                    print(f"Skipping {file_path.name}: missing both target sheets")
                    continue

                if empirical_sheet is not None:
                    empirical_rows.extend(process_empirical_sheet(wb, empirical_sheet, metadata, file_path.name))
                else:
                    print(f"Skipping Empirical Model in {file_path.name}: sheet not found")

                if regression_sheet is not None:
                    regression_rows.extend(process_regression_sheet(wb, regression_sheet, metadata, file_path.name))
                else:
                    print(f"Skipping Regression Model in {file_path.name}: sheet not found")

                processed_files += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)

    finally:
        app.quit()

    create_output_workbook(empirical_rows, regression_rows, out_path)

    print(f"Output path: {out_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
