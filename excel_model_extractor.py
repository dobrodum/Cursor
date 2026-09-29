from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from typing import Any

import pandas as pd
import xlwings as xw
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------
# User-configurable locations
# ---------------------------
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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


# Offsets relative to the "max" anchor column on each model tab.
# These are intentionally centralized for quick adjustment if layout shifts.
EMPIRICAL_OFFSETS = {
    "first_data_row": 1,
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "forecast_value": -4,  # estimated total sold
    "actual_value": -3,  # reported sales
    "forecast_max": 0,
    "forecast_min_primary": 1,
    "forecast_min_fallback": -1,
    "avg_pen_formula_col_offset": 20,  # temp formula area
}

REGRESSION_OFFSETS = {
    "first_data_row": 1,
    "num_quarters_used": -10,
    "forecast_total_without_sa": -4,  # TOT FCST w/o SA
    "actual_value": -3,  # optional if present
    "forecast_max": 0,
    "forecast_min_primary": 1,
    "forecast_min_fallback": -1,
    "intercept_formula_col_offset": 20,  # temp formula area
    "slope_formula_col_offset": 21,  # temp formula area
}


PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12]\d{3})", re.IGNORECASE)
TICKER_PATTERN = re.compile(r"\s-\s([A-Za-z0-9._-]+)\s-\s")

DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> float | None:
    if is_blank(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    as_float = to_float(value)
    if as_float is None:
        return None
    try:
        return int(round(as_float))
    except (TypeError, ValueError):
        return None


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def sheet_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    target = normalize_text(anchor_text)

    for r_idx, row in enumerate(values):
        for c_idx, raw in enumerate(row):
            if normalize_text(raw) == target:
                return start_row + r_idx, start_col + c_idx
    return None


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def safe_close_workbook(wb: xw.Book) -> None:
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


def parse_month_number(month_token: str) -> int | None:
    clean = re.sub(r"[^A-Za-z]", "", month_token or "")
    if not clean:
        return None

    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(clean, fmt).month
        except ValueError:
            continue

    short = clean[:3]
    try:
        return datetime.strptime(short, "%b").month
    except ValueError:
        return None


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    ticker = ""

    ticker_match = TICKER_PATTERN.search(stem)
    if ticker_match:
        ticker = ticker_match.group(1).strip().upper()
    else:
        pieces = [piece.strip() for piece in stem.split("-")]
        if len(pieces) >= 2:
            ticker = pieces[1].upper()

    period_match = PERIOD_PATTERN.search(stem)
    if not period_match:
        model_period = "unknown_period"
        model_date = ""
        model = f"{ticker}_{model_period}" if ticker else model_period
        return FileMetadata(model=model, ticker=ticker or "unknown", model_period=model_period, model_date=model_date)

    period_label = period_match.group(1).capitalize()
    month_token = period_match.group(2)
    year = int(period_match.group(3))

    month_number = parse_month_number(month_token)
    if month_number is None:
        model_period = f"{period_label}{month_token}_{year}"
        model_date = ""
        model = f"{ticker}_{model_period}" if ticker else model_period
        return FileMetadata(model=model, ticker=ticker or "unknown", model_period=model_period, model_date=model_date)

    month_abbr = datetime(year=year, month=month_number, day=1).strftime("%b")
    model_period = f"{period_label}{month_abbr}_{year}"
    day = DAY_MAP[period_label.lower()]
    model_date = datetime(year=year, month=month_number, day=day).strftime("%Y-%m-%d")
    model = f"{ticker}_{model_period}" if ticker else model_period

    return FileMetadata(model=model, ticker=ticker or "unknown", model_period=model_period, model_date=model_date)


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1

    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def keep_numeric_signature(*values: Any) -> tuple[Any, ...]:
    sig: list[Any] = []
    for value in values:
        num = to_float(value)
        if num is None:
            sig.append(value if not is_blank(value) else None)
        else:
            sig.append(round(num, 10))
    return tuple(sig)


def extract_empirical_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + EMPIRICAL_OFFSETS["first_data_row"]

    formula_cells: list[tuple[int, int]] = []
    provisional_rows: list[dict[str, Any]] = []

    for idx in range(N_QUARTERS):
        row = start_row + idx

        num_quarters_used = to_int(sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"]))
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = idx + 1

        last_quarter_used = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"])
        quarterly_sales = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"])
        reported_sales = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["reported_sales"])
        growth_rate_pct = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"])
        sales_captured_in_db_pct = sheet_cell_value(
            sheet, row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
        )
        forecast_value = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_value"])
        actual_value = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["actual_value"])
        forecast_max = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_max"])

        forecast_min = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_min_primary"])
        if is_blank(forecast_min):
            forecast_min = sheet_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_min_fallback"])

        # Keep only meaningful candidate rows.
        if is_blank(forecast_value) and is_blank(forecast_max) and is_blank(forecast_min) and is_blank(reported_sales):
            continue

        history_end = row
        history_start = max(start_row, row - num_quarters_used + 1)
        penetration_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
        avg_pen_formula_col = anchor_col + EMPIRICAL_OFFSETS["avg_pen_formula_col_offset"]
        avg_pen_formula_cell = sheet.cells(row, avg_pen_formula_col)

        avg_pen_formula = (
            f'=IFERROR(AVERAGE(R{history_start}C{penetration_col}:R{history_end}C{penetration_col}),"")'
        )
        set_formula2_r1c1(avg_pen_formula_cell, avg_pen_formula)
        formula_cells.append((row, avg_pen_formula_col))

        provisional_rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": None,  # filled after calculate()
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,  # estimated total sold
                "actual_value": reported_sales if is_blank(actual_value) else actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": None,  # computed below
                "avg_penetration_pct": None,  # filled after calculate()
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    if formula_cells:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for row_data, (row, col) in zip(provisional_rows, formula_cells):
        avg_penetration_pct = sheet_cell_value(sheet, row, col)
        row_data["avg_penetration_pct"] = avg_penetration_pct
        row_data["parameter_value"] = avg_penetration_pct

        max_value = to_float(row_data["forecast_max"])
        min_value = to_float(row_data["forecast_min"])
        row_data["range_width"] = (max_value - min_value) if max_value is not None and min_value is not None else None
        rows.append(row_data)

    return rows


def extract_regression_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + REGRESSION_OFFSETS["first_data_row"]

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    formula_cells: list[tuple[int, int, int]] = []
    provisional_rows: list[dict[str, Any]] = []

    for idx in range(N_QUARTERS):
        row = start_row + idx

        num_quarters_used = to_int(sheet_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["num_quarters_used"]))
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = idx + 1

        history_end = row
        history_start = max(start_row, row - num_quarters_used + 1)

        intercept_col = anchor_col + REGRESSION_OFFSETS["intercept_formula_col_offset"]
        slope_col = anchor_col + REGRESSION_OFFSETS["slope_formula_col_offset"]
        intercept_cell = sheet.cells(row, intercept_col)
        slope_cell = sheet.cells(row, slope_col)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{history_start}C{y_col}:R{history_end}C{y_col},'
            f'R{history_start}C{x_col}:R{history_end}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{history_start}C{y_col}:R{history_end}C{y_col},'
            f'R{history_start}C{x_col}:R{history_end}C{x_col}),"")'
        )
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)

        forecast_value = sheet_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_total_without_sa"])
        actual_value = sheet_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["actual_value"])
        forecast_max = sheet_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_max"])
        forecast_min = sheet_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_min_primary"])
        if is_blank(forecast_min):
            forecast_min = sheet_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_min_fallback"])

        if is_blank(forecast_value) and is_blank(forecast_max) and is_blank(forecast_min):
            continue

        formula_cells.append((row, intercept_col, slope_col))
        provisional_rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,  # TOT FCST w/o SA
                "actual_value": actual_value if not is_blank(actual_value) else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": None,  # computed below
                "intercept": None,  # filled after calculate()
                "slope": None,  # filled after calculate()
                "source_file": source_file,
            }
        )

    if formula_cells:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for row_data, (row, intercept_col, slope_col) in zip(provisional_rows, formula_cells):
        intercept_value = sheet_cell_value(sheet, row, intercept_col)
        slope_value = sheet_cell_value(sheet, row, slope_col)
        row_data["intercept"] = intercept_value
        row_data["slope"] = slope_value

        max_value = to_float(row_data["forecast_max"])
        min_value = to_float(row_data["forecast_min"])
        row_data["range_width"] = (max_value - min_value) if max_value is not None and min_value is not None else None

        signature = keep_numeric_signature(
            row_data["forecast_value"],
            row_data["forecast_max"],
            row_data["forecast_min"],
            row_data["intercept"],
            row_data["slope"],
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature
        rows.append(row_data)

    return rows


def format_worksheet(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    ws = writer.book[sheet_name]

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"

    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for idx, column_name in enumerate(dataframe.columns, start=1):
        max_len = len(column_name)
        for value in dataframe[column_name]:
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    empirical_df = pd.DataFrame(empirical_rows, columns=EMPIRICAL_COLUMNS)
    regression_df = pd.DataFrame(regression_rows, columns=REGRESSION_COLUMNS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        empirical_df.to_excel(writer, sheet_name="empirical_candidates", index=False)
        regression_df.to_excel(writer, sheet_name="regression_candidates", index=False)
        format_worksheet(writer, "empirical_candidates", empirical_df)
        format_worksheet(writer, "regression_candidates", regression_df)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input directory does not exist or is not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    output_file = unique_output_path(input_path, output_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    excel_candidates = sorted(input_path.iterdir(), key=lambda p: p.name.lower())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in excel_candidates:
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary workbook)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            wb: xw.Book | None = None
            try:
                print(f"Processing file: {file_path.name}")
                wb = app.books.open(str(file_path), update_links=False)

                metadata = parse_file_metadata(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
