from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# -------- User-configurable paths --------
input_dir = Path("./input")
output_dir = Path("./output")

# -------- Extraction constants --------
N_QUARTERS = 10

EMPIRICAL_SOURCE_SHEET = "Empirical Model"
REGRESSION_SOURCE_SHEET = "Regression Model"

EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"

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

# Offsets are relative to the "max" anchor cell.
# They are intentionally centralized for quick adjustment if workbook templates shift.
EMPIRICAL_OFFSETS = {
    "row_start": 1,
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "avg_penetration_source_col": -9,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "forecast_value": -2,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_helper_col": 2,
}

REGRESSION_OFFSETS = {
    "row_start": 1,
    "num_quarters_used": -12,
    "forecast_value": -2,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept_helper_col": 3,
    "slope_helper_col": 4,
}


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def _parse_model_period(period_token: str) -> Tuple[str, str]:
    """
    Converts strings like MidJan2026 -> (MidJan_2026, 2026-01-15).
    Day convention:
      Early -> 05, Mid -> 15, Late -> 25
    """
    match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, flags=re.IGNORECASE)
    if not match:
        return "", ""

    period_prefix = match.group(1).title()
    month_token = match.group(2)[:3].title()
    year = int(match.group(3))
    day_map = {"Early": 5, "Mid": 15, "Late": 25}

    try:
        month = datetime.strptime(month_token, "%b").month
    except ValueError:
        return "", ""

    model_period = f"{period_prefix}{month_token}_{year}"
    model_date = f"{year:04d}-{month:02d}-{day_map[period_prefix]:02d}"
    return model_period, model_date


def parse_file_metadata(file_path: Path) -> FileMetadata:
    """
    Expected filename pattern example:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = parts[1].upper() if len(parts) >= 2 else ""
    period_chunk = parts[2] if len(parts) >= 3 else ""
    period_token = period_chunk.split("_")[0].strip()
    model_period, model_date = _parse_model_period(period_token)

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or stem
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def should_skip_file(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    return None


def safe_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def subtract_values(left: Any, right: Any) -> Optional[float]:
    left_number = safe_number(left)
    right_number = safe_number(right)
    if left_number is None or right_number is None:
        return None
    return left_number - right_number


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    return sheet.cells(row, col).value


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    # Fast path: Excel-native Find.
    try:
        found = sheet.api.Cells.Find(What=anchor_text, LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    # Fallback: one-pass scan of used range.
    used = sheet.used_range
    values = used.options(ndim=2).value
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    anchor_text_lower = anchor_text.lower()

    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == anchor_text_lower:
                return start_row + row_offset, start_col + col_offset
    return None


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    if not formula.startswith("="):
        formula = f"={formula}"

    # Prefer .formula2 as requested.
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass

    # Safe fallback paths for engine differences.
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass

    cell.api.FormulaR1C1 = formula


def close_source_workbook(wb: xw.Book) -> None:
    # Preferred close call.
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Signature fallback for environments where keyword args are unsupported.
    try:
        wb.close(False)
        return
    except Exception:
        pass

    # COM-level fallbacks.
    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    wb.api.Close(False)


def _coerce_quarters(raw_value: Any, fallback: int) -> int:
    value = safe_number(raw_value)
    if value is None:
        return fallback
    quarter_count = int(round(value))
    return max(1, quarter_count)


def extract_empirical_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SOURCE_SHEET]
    except Exception:
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        return []

    anchor_row, anchor_col = anchor
    helper_col = anchor_col + EMPIRICAL_OFFSETS["avg_penetration_helper_col"]
    prepared: List[Tuple[int, int, xw.Range]] = []

    for idx in range(N_QUARTERS):
        row_num = anchor_row + EMPIRICAL_OFFSETS["row_start"] + idx
        raw_quarters = get_cell_value(sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"])
        quarters_used = _coerce_quarters(raw_quarters, fallback=idx + 1)

        hist_start = max(1, anchor_row - quarters_used)
        hist_end = max(1, anchor_row - 1)
        penetration_col = anchor_col + EMPIRICAL_OFFSETS["avg_penetration_source_col"]

        helper_cell = sheet.cells(row_num, helper_col)
        formula = f"=AVERAGE(R{hist_start}C{penetration_col}:R{hist_end}C{penetration_col})"
        set_formula2_r1c1(helper_cell, formula)
        prepared.append((row_num, quarters_used, helper_cell))

    if prepared:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for row_num, quarters_used, helper_cell in prepared:
        avg_penetration_pct = helper_cell.value
        forecast_value = get_cell_value(sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["forecast_value"])
        reported_sales = get_cell_value(sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["reported_sales"])
        forecast_max = get_cell_value(sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["forecast_max"])
        forecast_min = get_cell_value(sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["forecast_min"])

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": get_cell_value(
                sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"]
            )
            or quarters_used,
            "last_quarter_used": get_cell_value(
                sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"]
            ),
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_values(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": get_cell_value(sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]),
            "reported_sales": reported_sales,
            "growth_rate_pct": get_cell_value(sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]),
            "sales_captured_in_db_pct": get_cell_value(
                sheet, row_num, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
            ),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def extract_regression_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SOURCE_SHEET]
    except Exception:
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    prepared: List[Tuple[int, int, xw.Range, xw.Range]] = []
    for idx in range(N_QUARTERS):
        row_num = anchor_row + REGRESSION_OFFSETS["row_start"] + idx
        raw_quarters = get_cell_value(sheet, row_num, anchor_col + REGRESSION_OFFSETS["num_quarters_used"])
        quarters_used = _coerce_quarters(raw_quarters, fallback=idx + 1)

        hist_start = max(1, anchor_row - quarters_used)
        hist_end = max(1, anchor_row - 1)

        intercept_cell = sheet.cells(row_num, anchor_col + REGRESSION_OFFSETS["intercept_helper_col"])
        slope_cell = sheet.cells(row_num, anchor_col + REGRESSION_OFFSETS["slope_helper_col"])

        intercept_formula = (
            f"=INTERCEPT(R{hist_start}C{y_col}:R{hist_end}C{y_col},R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{hist_start}C{y_col}:R{hist_end}C{y_col},R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        prepared.append((row_num, quarters_used, intercept_cell, slope_cell))

    if prepared:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for row_num, quarters_used, intercept_cell, slope_cell in prepared:
        forecast_value = get_cell_value(sheet, row_num, anchor_col + REGRESSION_OFFSETS["forecast_value"])
        forecast_max = get_cell_value(sheet, row_num, anchor_col + REGRESSION_OFFSETS["forecast_max"])
        forecast_min = get_cell_value(sheet, row_num, anchor_col + REGRESSION_OFFSETS["forecast_min"])
        intercept_value = intercept_cell.value
        slope_value = slope_cell.value
        num_quarters_used = (
            get_cell_value(sheet, row_num, anchor_col + REGRESSION_OFFSETS["num_quarters_used"])
            or quarters_used
        )

        signature = (
            safe_number(num_quarters_used),
            safe_number(intercept_value),
            safe_number(slope_value),
            safe_number(forecast_value),
            safe_number(forecast_max),
            safe_number(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            # Prevent duplicate final row if the template repeats the same computed values.
            continue
        previous_signature = signature

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
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_values(forecast_max, forecast_min),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def write_rows(ws: Worksheet, columns: Iterable[str], rows: List[Dict[str, Any]]) -> None:
    columns = list(columns)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(ws.max_row, 1)}"

    for index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_index in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_index, column=index).value
            if cell_value is None:
                continue
            max_length = max(max_length, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(index)].width = min(max(max_length + 2, 12), 42)


def save_output_workbook(
    target_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = EMPIRICAL_OUTPUT_SHEET
    regression_ws = workbook.create_sheet(REGRESSION_OUTPUT_SHEET)

    write_rows(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(target_path)


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    target_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {source_dir}")
    if not source_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a folder: {source_dir}")
    target_dir.mkdir(parents=True, exist_ok=True)

    output_path = get_output_path(source_dir, target_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        # Keep a single hidden Excel app alive for full run.
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(source_dir.iterdir(), key=lambda path: path.name.lower()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            wb: Optional[xw.Book] = None
            try:
                metadata = parse_file_metadata(file_path)
                wb = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))

                files_processed += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    save_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
