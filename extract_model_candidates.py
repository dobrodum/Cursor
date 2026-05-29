#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
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

DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetScan:
    top_row: int
    left_col: int
    values: List[List[Any]]
    labels: Dict[str, List[Tuple[int, int]]]


def normalize_matrix(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if not isinstance(value, list):
        return [[value]]
    if not value:
        return []
    if not isinstance(value[0], list):
        return [value]
    matrix: List[List[Any]] = []
    for row in value:
        if isinstance(row, list):
            matrix.append(row)
        else:
            matrix.append([row])
    return matrix


def normalize_label(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            if math.isnan(value):  # type: ignore[arg-type]
                return None
        except TypeError:
            pass
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").strip()
        if not cleaned:
            return None
        pct = cleaned.endswith("%")
        if pct:
            cleaned = cleaned[:-1]
        try:
            parsed = float(cleaned)
        except ValueError:
            return None
        return parsed / 100.0 if pct else parsed
    return None


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def fmt_signature(values: Sequence[Optional[float]]) -> Tuple[Optional[float], ...]:
    result: List[Optional[float]] = []
    for value in values:
        if value is None:
            result.append(None)
        else:
            result.append(round(value, 10))
    return tuple(result)


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("expected '<prefix> - <ticker> - <period>' naming pattern")

    ticker = parts[-2].strip().upper()
    period_token = re.sub(r"([_\s-]*send)$", "", parts[-1], flags=re.IGNORECASE).strip()
    period_match = re.fullmatch(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", period_token, flags=re.IGNORECASE)
    if not period_match:
        raise ValueError("period token must look like EarlyJan2026 / MidJan2026 / LateJan2026")

    bucket = period_match.group(1).title()
    month_short = period_match.group(2).title()
    year = int(period_match.group(3))
    month = MONTH_MAP.get(month_short.lower())
    if month is None:
        raise ValueError(f"unsupported month abbreviation '{month_short}'")

    day = DAY_MAP[bucket.lower()]
    model_period = f"{bucket}{month_short}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def scan_sheet(sheet: xw.Sheet) -> SheetScan:
    used_range = sheet.used_range
    values = normalize_matrix(used_range.value)
    labels: Dict[str, List[Tuple[int, int]]] = {}
    for row_idx, row_vals in enumerate(values):
        for col_idx, cell_val in enumerate(row_vals):
            if isinstance(cell_val, str):
                key = normalize_label(cell_val)
                if not key:
                    continue
                abs_row = used_range.row + row_idx
                abs_col = used_range.column + col_idx
                labels.setdefault(key, []).append((abs_row, abs_col))
    return SheetScan(
        top_row=used_range.row,
        left_col=used_range.column,
        values=values,
        labels=labels,
    )


def matrix_value(scan: SheetScan, row: int, col: int) -> Any:
    row_idx = row - scan.top_row
    col_idx = col - scan.left_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(scan.values):
        return None
    row_values = scan.values[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def get_value(scan: SheetScan, sheet: xw.Sheet, row: int, col: int) -> Any:
    from_matrix = matrix_value(scan, row, col)
    if from_matrix is not None:
        return from_matrix
    return sheet.cells(row, col).value


def find_label_cell(scan: SheetScan, options: Iterable[str]) -> Optional[Tuple[int, int]]:
    for label in options:
        key = normalize_label(label)
        if key in scan.labels and scan.labels[key]:
            return scan.labels[key][0]
    return None


def find_label_value_right(scan: SheetScan, sheet: xw.Sheet, options: Iterable[str]) -> Any:
    for label in options:
        key = normalize_label(label)
        for row, col in scan.labels.get(key, []):
            value = get_value(scan, sheet, row, col + 1)
            if value not in (None, ""):
                return value
    return None


def find_anchor_max(scan: SheetScan) -> Optional[Tuple[int, int]]:
    max_cells = scan.labels.get("max", [])
    if not max_cells:
        return None
    min_set = set(scan.labels.get("min", []))
    for row, col in max_cells:
        if (row + 1, col) in min_set or (row - 1, col) in min_set:
            return row, col
    return max_cells[0]


def safe_close_workbook(workbook: Optional[xw.Book]) -> None:
    if workbook is None:
        return
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)  # type: ignore[attr-defined]
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{in_dir.name}_PARAM.{suffix}.xlsx"
        suffix += 1
    return candidate


def collect_source_files(folder: Path) -> List[Path]:
    if not folder.exists():
        print(f"Skipped: {folder} (input directory does not exist)")
        return []
    if not folder.is_dir():
        print(f"Skipped: {folder} (input path is not a directory)")
        return []

    accepted: List[Path] = []
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            print(f"Skipped: {path.name} (not a file)")
            continue
        if path.name.startswith("~"):
            print(f"Skipped: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped: {path.name} (not .xlsx)")
            continue
        accepted.append(path)
    return accepted


def extract_empirical_rows(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        return []

    scan = scan_sheet(sheet)
    anchor = find_anchor_max(scan)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    n_quarters = 10
    end_col = max(1, anchor_col - 2)
    quarter_row = max(1, anchor_row - 7)
    penetration_row = max(1, anchor_row - 6)
    quarterly_sales_row = max(1, anchor_row - 5)
    reported_sales_row = max(1, anchor_row - 4)

    temp_col = anchor_col + 24
    avg_cell = sheet.cells(anchor_row, temp_col)
    forecast_cell = sheet.cells(anchor_row + 1, temp_col)
    max_cell = sheet.cells(anchor_row + 2, temp_col)
    min_cell = sheet.cells(anchor_row + 3, temp_col)

    anchor_forecast_max = to_float(get_value(scan, sheet, anchor_row, anchor_col + 1))
    anchor_forecast_min = to_float(get_value(scan, sheet, anchor_row + 1, anchor_col + 1))

    estimated_total_from_label = to_float(
        find_label_value_right(
            scan,
            sheet,
            options=["estimated total sold", "est total sold", "tot fcst", "total forecast"],
        )
    )
    reported_sales_from_label = to_float(
        find_label_value_right(
            scan,
            sheet,
            options=["reported sales", "actual sales"],
        )
    )

    rows: List[Dict[str, Any]] = []
    for used_quarters in range(1, n_quarters + 1):
        start_col = end_col - used_quarters + 1
        if start_col < 1:
            break

        # Temporary worksheet formulas to keep logic in Excel and avoid column-letter conversion.
        avg_cell.formula2 = (
            f"=AVERAGE(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
        )
        forecast_cell.formula2 = (
            f"=R{quarterly_sales_row}C{end_col}*R{avg_cell.row}C{avg_cell.column}"
        )
        max_cell.formula2 = (
            f"=MAX(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
            f"*R{quarterly_sales_row}C{end_col}"
        )
        min_cell.formula2 = (
            f"=MIN(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
            f"*R{quarterly_sales_row}C{end_col}"
        )
        workbook.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        forecast_value = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)

        if forecast_value is None:
            forecast_value = estimated_total_from_label
        if forecast_max is None:
            forecast_max = anchor_forecast_max
        if forecast_min is None:
            forecast_min = anchor_forecast_min

        quarterly_sales = to_float(get_value(scan, sheet, quarterly_sales_row, end_col))
        reported_sales = to_float(get_value(scan, sheet, reported_sales_row, end_col))
        if reported_sales is None:
            reported_sales = reported_sales_from_label
        first_quarter_sales = to_float(get_value(scan, sheet, quarterly_sales_row, start_col))

        last_quarter_used = get_value(scan, sheet, quarter_row, end_col)
        if last_quarter_used in (None, ""):
            last_quarter_used = f"C{end_col}"

        growth_ratio = safe_div(quarterly_sales, first_quarter_sales)
        growth_rate_pct = None if growth_ratio is None else (growth_ratio - 1.0)
        sales_captured_in_db_pct = safe_div(reported_sales, quarterly_sales)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": used_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        return []

    scan = scan_sheet(sheet)
    anchor = find_anchor_max(scan)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        return []

    data_points: List[Tuple[int, float, float]] = []
    row_start = scan.top_row
    row_end = anchor_row - 1
    for row in range(row_start, row_end + 1):
        x_val = to_float(get_value(scan, sheet, row, x_col))
        y_val = to_float(get_value(scan, sheet, row, y_col))
        if x_val is None or y_val is None:
            continue
        data_points.append((row, x_val, y_val))

    if len(data_points) < 2:
        return []

    temp_col = anchor_col + 24
    intercept_cell = sheet.cells(anchor_row, temp_col)
    slope_cell = sheet.cells(anchor_row + 1, temp_col)
    forecast_cell = sheet.cells(anchor_row + 2, temp_col)
    max_cell = sheet.cells(anchor_row + 3, temp_col)
    min_cell = sheet.cells(anchor_row + 4, temp_col)

    forecast_total_without_sa_label = to_float(
        find_label_value_right(scan, sheet, options=["tot fcst w/o sa", "tot fcst without sa"])
    )
    anchor_max = to_float(get_value(scan, sheet, anchor_row, anchor_col + 1))
    anchor_min = to_float(get_value(scan, sheet, anchor_row + 1, anchor_col + 1))

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Optional[float], ...]] = None
    max_window = min(10, len(data_points))
    for used_quarters in range(2, max_window + 1):
        window = data_points[-used_quarters:]
        first_row = window[0][0]
        last_row = window[-1][0]
        next_x = window[-1][1] + 1.0

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        forecast_cell.formula2 = (
            f"=R{intercept_cell.row}C{intercept_cell.column}"
            f"+R{slope_cell.row}C{slope_cell.column}*{next_x}"
        )
        max_cell.formula2 = f"=MAX(R{first_row}C{y_col}:R{last_row}C{y_col})"
        min_cell.formula2 = f"=MIN(R{first_row}C{y_col}:R{last_row}C{y_col})"
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)

        if forecast_value is None:
            forecast_value = forecast_total_without_sa_label
        if forecast_max is None:
            forecast_max = anchor_max
        if forecast_min is None:
            forecast_min = anchor_min

        signature = fmt_signature(
            [intercept, slope, forecast_value, forecast_max, forecast_min]
        )
        if signature == prev_signature:
            # Keep only unique trailing rows when final calc duplicates previous output.
            continue
        prev_signature = signature

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": used_quarters,
                "num_quarters_used": used_quarters,
                "forecast_value": forecast_value,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_table_sheet(worksheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    worksheet.append(list(columns))
    for col_idx in range(1, len(columns) + 1):
        worksheet.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        worksheet.append([row.get(col) for col in columns])

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{worksheet.max_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for cell in worksheet[get_column_letter(col_idx)]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 45)


def write_output_workbook(
    path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_table_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_table_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(path)


def main() -> None:
    source_files = collect_source_files(input_dir)
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in source_files:
            try:
                metadata = parse_file_metadata(file_path.name)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (filename parse error: {exc})")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_empirical = extract_empirical_rows(workbook, metadata, file_path.name)
                file_regression = extract_regression_rows(workbook, metadata, file_path.name)

                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                processed_files += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
