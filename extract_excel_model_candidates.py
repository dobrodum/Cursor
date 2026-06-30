from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# User-configurable folders
# =========================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


# =======================
# Workbook/Sheet settings
# =======================
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10


# =====================================
# Anchor-based offsets (relative to max)
# =====================================
#
# These offsets mirror the expected model layout and keep all positional logic in
# one place. If your workbook layout changes, update only these constants.

# Empirical rows are expected to start directly below the "max" anchor row.
EMPIRICAL_ROW_START_OFFSET = 1
EMP_COL_OFFSETS = {
    "avg_penetration_pct": -9,
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "forecast_value": -4,  # estimated total sold
    "actual_value": -3,  # reported sales
    "growth_rate_pct": 2,
    "sales_captured_in_db_pct": 3,
    "forecast_max": 0,
    "forecast_min": 1,
}

# Avg penetration formula source settings for empirical rows.
# Formula is written as:
#   =AVERAGE(R[source_row_offset]C[left_col_offset]:R[source_row_offset]C[right_col_offset])
EMP_AVG_SOURCE_ROW_OFFSET = -2
EMP_AVG_SOURCE_RIGHTMOST_COL_OFFSET = -10


# Regression rows are expected to start directly below the "max" anchor row.
REGRESSION_ROW_START_OFFSET = 1
REG_COL_OFFSETS = {
    "num_quarters_used": -2,
    "actual_value": -3,  # optional; may be blank
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": 2,
    "slope": 3,
}


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


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


MONTH_LOOKUP = {
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

PERIOD_DAY_LOOKUP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

FILENAME_REGEX = re.compile(
    r"^\s*.*?-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})",
    re.IGNORECASE,
)


def parse_file_metadata(file_path: Path) -> FileMetadata:
    match = FILENAME_REGEX.search(file_path.stem)
    if not match:
        ticker = file_path.stem
        return FileMetadata(model=ticker, ticker=ticker, model_period="", model_date="")

    ticker = match.group(1).upper()
    period_label = match.group(2).title()
    month_token = match.group(3)[:3].lower()
    year = int(match.group(4))

    month = MONTH_LOOKUP.get(month_token)
    if month is None:
        return FileMetadata(
            model=ticker,
            ticker=ticker,
            model_period=f"{period_label}{match.group(3)}_{year}",
            model_date="",
        )

    day = PERIOD_DAY_LOOKUP[period_label.lower()]
    model_period = f"{period_label}{match.group(3)[:3].title()}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def make_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    try:
        num = float(value)
    except (TypeError, ValueError):
        return None
    if num != num:  # NaN guard
        return None
    return num


def clean_value(value: Any) -> Any:
    if isinstance(value, str):
        return value.strip()
    return value


def compute_range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = to_number(max_value)
    min_num = to_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


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
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        print("Warning: failed to close workbook cleanly without saving.")


def find_anchor_max(sheet: xw.Sheet) -> tuple[int, int]:
    try:
        found = sheet.api.Cells.Find(What="max", LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    used = sheet.used_range
    values = used.options(ndim=2).value or []
    start_row, start_col = used.row, used.column

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return start_row + r_idx, start_col + c_idx

    raise ValueError(f"Could not find 'max' anchor in sheet '{sheet.name}'.")


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    return clean_value(sheet.cells(row, col).value)


def process_empirical_sheet(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"  Skipped empirical: missing sheet '{EMPIRICAL_SHEET_NAME}'.")
        return []

    anchor_row, anchor_col = find_anchor_max(sheet)
    rows: list[dict[str, Any]] = []

    # Write all average penetration formulas first, then calculate once.
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        row_num = anchor_row + EMPIRICAL_ROW_START_OFFSET + idx
        avg_col = anchor_col + EMP_COL_OFFSETS["avg_penetration_pct"]
        avg_cell = sheet.cells(row_num, avg_col)
        right_col_offset = EMP_AVG_SOURCE_RIGHTMOST_COL_OFFSET
        left_col_offset = right_col_offset - (n_quarters - 1)
        avg_cell.formula2 = (
            f"=AVERAGE(R[{EMP_AVG_SOURCE_ROW_OFFSET}]C[{left_col_offset}]"
            f":R[{EMP_AVG_SOURCE_ROW_OFFSET}]C[{right_col_offset}])"
        )

    wb.app.calculate()

    for idx in range(N_QUARTERS):
        row_num = anchor_row + EMPIRICAL_ROW_START_OFFSET + idx
        n_quarters = idx + 1

        forecast_max = read_cell(sheet, row_num, anchor_col + EMP_COL_OFFSETS["forecast_max"])
        forecast_min = read_cell(sheet, row_num, anchor_col + EMP_COL_OFFSETS["forecast_min"])
        avg_penetration_pct = read_cell(
            sheet, row_num, anchor_col + EMP_COL_OFFSETS["avg_penetration_pct"]
        )

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": read_cell(
                sheet, row_num, anchor_col + EMP_COL_OFFSETS["num_quarters_used"]
            )
            or n_quarters,
            "last_quarter_used": read_cell(
                sheet, row_num, anchor_col + EMP_COL_OFFSETS["last_quarter_used"]
            ),
            "forecast_value": read_cell(
                sheet, row_num, anchor_col + EMP_COL_OFFSETS["forecast_value"]
            ),
            "actual_value": read_cell(sheet, row_num, anchor_col + EMP_COL_OFFSETS["actual_value"]),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": read_cell(
                sheet, row_num, anchor_col + EMP_COL_OFFSETS["quarterly_sales"]
            ),
            "reported_sales": read_cell(sheet, row_num, anchor_col + EMP_COL_OFFSETS["reported_sales"]),
            "growth_rate_pct": read_cell(sheet, row_num, anchor_col + EMP_COL_OFFSETS["growth_rate_pct"]),
            "sales_captured_in_db_pct": read_cell(
                sheet, row_num, anchor_col + EMP_COL_OFFSETS["sales_captured_in_db_pct"]
            ),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def values_match(a: Any, b: Any, tol: float = 1e-10) -> bool:
    a_num = to_number(a)
    b_num = to_number(b)
    if a_num is not None and b_num is not None:
        return abs(a_num - b_num) <= tol
    return clean_value(a) == clean_value(b)


def process_regression_sheet(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"  Skipped regression: missing sheet '{REGRESSION_SHEET_NAME}'.")
        return []

    anchor_row, anchor_col = find_anchor_max(sheet)
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    max_rows = min(N_QUARTERS, max(anchor_row - 1, 0))
    rows: list[dict[str, Any]] = []

    # Write all INTERCEPT/SLOPE formulas first, then calculate once.
    for idx in range(max_rows):
        row_num = anchor_row + REGRESSION_ROW_START_OFFSET + idx
        n_quarters = idx + 1
        start_row = anchor_row - n_quarters
        end_row = anchor_row - 1

        intercept_cell = sheet.cells(row_num, anchor_col + REG_COL_OFFSETS["intercept"])
        slope_cell = sheet.cells(row_num, anchor_col + REG_COL_OFFSETS["slope"])

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

    wb.app.calculate()

    for idx in range(max_rows):
        row_num = anchor_row + REGRESSION_ROW_START_OFFSET + idx
        n_quarters = idx + 1

        forecast_value = read_cell(sheet, row_num, anchor_col + REG_COL_OFFSETS["forecast_value"])
        forecast_max = read_cell(sheet, row_num, anchor_col + REG_COL_OFFSETS["forecast_max"])
        forecast_min = read_cell(sheet, row_num, anchor_col + REG_COL_OFFSETS["forecast_min"])
        intercept = read_cell(sheet, row_num, anchor_col + REG_COL_OFFSETS["intercept"])
        slope = read_cell(sheet, row_num, anchor_col + REG_COL_OFFSETS["slope"])
        num_quarters_used = read_cell(sheet, row_num, anchor_col + REG_COL_OFFSETS["num_quarters_used"])
        if not num_quarters_used:
            num_quarters_used = n_quarters

        next_row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": read_cell(sheet, row_num, anchor_col + REG_COL_OFFSETS["actual_value"])
            or "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows:
            prev = rows[-1]
            duplicate = (
                values_match(next_row["forecast_value"], prev["forecast_value"])
                and values_match(next_row["forecast_max"], prev["forecast_max"])
                and values_match(next_row["forecast_min"], prev["forecast_min"])
                and values_match(next_row["intercept"], prev["intercept"])
                and values_match(next_row["slope"], prev["slope"])
            )
            if duplicate:
                continue

        rows.append(next_row)

    return rows


def write_output_workbook(
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
    output_path: Path,
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    write_sheet(
        wb=wb,
        sheet_name="empirical_candidates",
        columns=EMPIRICAL_COLUMNS,
        rows=empirical_rows,
    )
    write_sheet(
        wb=wb,
        sheet_name="regression_candidates",
        columns=REGRESSION_COLUMNS,
        rows=regression_rows,
    )
    wb.save(output_path)


def write_sheet(wb: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(columns)
    for record in rows:
        ws.append([record.get(column, "") for column in columns])

    header_font = Font(bold=True)
    for col_idx, column_name in enumerate(columns, start=1):
        header_cell = ws.cell(row=1, column=col_idx)
        header_cell.font = header_font

        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            max_len = max(max_len, len(str(cell_value)) if cell_value is not None else 0)

        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 45)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"


def iter_input_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for item in sorted(folder.iterdir()):
        if not item.is_file():
            print(f"Skipped: {item.name} (not a file)")
            continue
        if item.name.startswith("~"):
            print(f"Skipped: {item.name} (temporary file)")
            continue
        if item.suffix.lower() != ".xlsx":
            print(f"Skipped: {item.name} (not .xlsx)")
            continue
        files.append(item)
    return files


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    output_path = make_output_path(input_dir, output_dir)
    source_files = iter_input_files(input_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processed: {file_path.name}")
            wb = app.books.open(str(file_path), update_links=False)
            try:
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                processed_file_count += 1
            except Exception as exc:
                print(f"  Skipped: {file_path.name} (processing error: {exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(empirical_rows, regression_rows, output_path)
    print(f"Output: {output_path}")
    print(f"Files processed: {processed_file_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
