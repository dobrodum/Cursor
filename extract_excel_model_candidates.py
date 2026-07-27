from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these folders before running.
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


# Fallback offsets from the "max" anchor cell when matching headers are not found.
EMPIRICAL_FALLBACK_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
    "actual_value": -2,
    "reported_sales": -2,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -4,
    "quarterly_sales": -5,
    "last_quarter_used": -6,
    "num_quarters_used": -7,
    "penetration_input": -8,
}

REGRESSION_FALLBACK_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,  # TOT FCST w/o SA
    "actual_value": -2,
    "num_quarters_used": -7,
}


EMPIRICAL_HEADER_ALIASES = {
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "forecast_value": ("estimated total sold", "forecast total", "tot fcst", "forecast"),
    "actual_value": ("actual sales", "actual", "reported sales"),
    "reported_sales": ("reported sales", "actual sales"),
    "growth_rate_pct": ("growth rate", "growth pct", "growth %"),
    "sales_captured_in_db_pct": ("sales captured in db", "captured in db", "db capture"),
    "quarterly_sales": ("quarterly sales", "qtr sales"),
    "last_quarter_used": ("last quarter used", "last quarter", "last qtr"),
    "num_quarters_used": ("num quarters used", "n quarters", "quarters used", "qtrs used"),
    "penetration_input": ("penetration", "pen %", "avg penetration"),
}

REGRESSION_HEADER_ALIASES = {
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "forecast_value": ("tot fcst w/o sa", "forecast total without sa", "tot fcst", "forecast"),
    "actual_value": ("actual sales", "actual", "reported sales"),
    "num_quarters_used": ("num quarters used", "n quarters", "quarters used", "qtrs used"),
}


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

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class ModelInfo:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text.startswith("#"):
        return None
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def to_positive_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    rounded = int(round(numeric))
    if rounded <= 0:
        return None
    return rounded


def safe_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def as_matrix(value: Any, expected_rows: Optional[int] = None, expected_cols: Optional[int] = None) -> List[List[Any]]:
    if isinstance(value, tuple):
        value = list(value)

    if isinstance(value, list):
        if not value:
            matrix: List[List[Any]] = []
        elif isinstance(value[0], (list, tuple)):
            matrix = [list(row) if isinstance(row, tuple) else row for row in value]
        else:
            matrix = [value]
    else:
        matrix = [[value]]

    if expected_rows is not None:
        while len(matrix) < expected_rows:
            matrix.append([])
        matrix = matrix[:expected_rows]

    if expected_cols is not None:
        for idx, row in enumerate(matrix):
            row_copy = list(row)
            if len(row_copy) < expected_cols:
                row_copy.extend([None] * (expected_cols - len(row_copy)))
            matrix[idx] = row_copy[:expected_cols]

    return matrix


def as_column_list(value: Any, size: int) -> List[Any]:
    matrix = as_matrix(value, expected_rows=size)
    column: List[Any] = []
    for row in matrix:
        if isinstance(row, Sequence) and row:
            column.append(row[0])
        else:
            column.append(None)
    return column


def pick_sheet(workbook: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target_norm = normalize_text(target_name)
    for sheet in workbook.sheets:
        if normalize_text(sheet.name) == target_norm:
            return sheet
    return None


def parse_file_label(file_path: Path) -> ModelInfo:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = parts[-2].upper() if len(parts) >= 2 else "UNKNOWN"
    period_raw = parts[-1] if parts else stem
    period_token = period_raw.split("_")[0]

    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$", period_token, flags=re.IGNORECASE)
    if not match:
        model_period = "UnknownPeriod"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return ModelInfo(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    phase = match.group(1)
    month_abbrev = match.group(2)
    year = int(match.group(3))

    month_num = MONTH_MAP.get(month_abbrev.lower())
    day = PHASE_DAY.get(phase.lower())
    if month_num is None or day is None:
        model_period = "UnknownPeriod"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return ModelInfo(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    model_period = f"{phase}{month_abbrev.title()}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelInfo(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Tuple[int, int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    last_col = used.last_cell.column
    values = as_matrix(used.value)
    target = normalize_text(anchor_text)

    for row_idx, row_values in enumerate(values):
        row_seq = list(row_values) if isinstance(row_values, Sequence) else [row_values]
        for col_idx, cell_value in enumerate(row_seq):
            normalized_cell = normalize_text(cell_value)
            if normalized_cell == target or target in normalized_cell.split(" "):
                return start_row + row_idx, start_col + col_idx, last_col

    raise ValueError(f'Could not find "{anchor_text}" anchor in sheet "{sheet.name}".')


def resolve_offsets(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    aliases: Dict[str, Tuple[str, ...]],
    fallback_offsets: Dict[str, int],
    last_col: int,
) -> Dict[str, int]:
    offsets = dict(fallback_offsets)
    header_row_values = as_matrix(sheet.range((anchor_row, 1), (anchor_row, last_col)).value, expected_rows=1)[0]
    normalized_headers = [normalize_text(cell) for cell in header_row_values]

    for field_name, field_aliases in aliases.items():
        chosen_col: Optional[int] = None
        for col_idx, header in enumerate(normalized_headers, start=1):
            if not header:
                continue
            if any(alias in header for alias in field_aliases):
                chosen_col = col_idx
                break
        if chosen_col is not None:
            offsets[field_name] = chosen_col - anchor_col

    return offsets


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass

    try:
        cell.api.Formula2 = formula
        return
    except Exception:
        pass

    cell.formula = formula


def extract_from_block(
    block: List[List[Any]],
    block_col_start: int,
    row_index: int,
    anchor_col: int,
    offset: int,
) -> Any:
    abs_col = anchor_col + offset
    relative_col = abs_col - block_col_start
    if relative_col < 0:
        return None
    if row_index < 0 or row_index >= len(block):
        return None
    row = block[row_index]
    if relative_col >= len(row):
        return None
    return row[relative_col]


def close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    workbook.api.Close(False)


def process_empirical_sheet(workbook: xw.Book, info: ModelInfo, source_file: str) -> List[Dict[str, Any]]:
    sheet = pick_sheet(workbook, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        print(f"SKIPPED {source_file}: missing sheet '{EMPIRICAL_SHEET_NAME}'")
        return []

    try:
        anchor_row, anchor_col, last_col = find_anchor(sheet, "max")
    except ValueError as error:
        print(f"SKIPPED {source_file}: {error}")
        return []

    offsets = resolve_offsets(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        aliases=EMPIRICAL_HEADER_ALIASES,
        fallback_offsets=EMPIRICAL_FALLBACK_OFFSETS,
        last_col=last_col,
    )

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1
    helper_col = last_col + 2

    num_quarters_col = max(1, anchor_col + offsets["num_quarters_used"])
    num_quarters_values = as_column_list(
        sheet.range((first_data_row, num_quarters_col), (last_data_row, num_quarters_col)).value,
        size=N_QUARTERS,
    )

    penetration_col = anchor_col + offsets["penetration_input"]
    formula_rows: List[int] = []
    if penetration_col >= 1:
        for idx in range(N_QUARTERS):
            row = first_data_row + idx
            num_quarters_used = to_positive_int(num_quarters_values[idx]) or (idx + 1)
            start_row = max(1, row - num_quarters_used + 1)
            avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{row}C{penetration_col})"
            set_formula2(sheet.cells(row, helper_col), avg_formula)
            formula_rows.append(row)

    if formula_rows:
        workbook.app.calculate()

    avg_penetration_values = (
        as_column_list(
            sheet.range((first_data_row, helper_col), (last_data_row, helper_col)).value,
            size=N_QUARTERS,
        )
        if formula_rows
        else [None] * N_QUARTERS
    )

    offsets_to_read = {
        offsets["num_quarters_used"],
        offsets["last_quarter_used"],
        offsets["forecast_value"],
        offsets["actual_value"],
        offsets["forecast_max"],
        offsets["forecast_min"],
        offsets["quarterly_sales"],
        offsets["reported_sales"],
        offsets["growth_rate_pct"],
        offsets["sales_captured_in_db_pct"],
    }
    min_col = max(1, anchor_col + min(offsets_to_read))
    max_col = max(1, anchor_col + max(offsets_to_read))
    block = as_matrix(
        sheet.range((first_data_row, min_col), (last_data_row, max_col)).value,
        expected_rows=N_QUARTERS,
        expected_cols=max_col - min_col + 1,
    )

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters_used = (
            to_positive_int(
                extract_from_block(block, min_col, idx, anchor_col, offsets["num_quarters_used"])
            )
            or (idx + 1)
        )
        last_quarter_used = extract_from_block(block, min_col, idx, anchor_col, offsets["last_quarter_used"])
        forecast_value = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["forecast_value"]))
        actual_value = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["actual_value"]))
        forecast_max = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["forecast_max"]))
        forecast_min = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["forecast_min"]))
        avg_penetration_pct = to_float(avg_penetration_values[idx])
        quarterly_sales = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["quarterly_sales"]))
        reported_sales = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["reported_sales"]))
        growth_rate_pct = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["growth_rate_pct"]))
        sales_captured_in_db_pct = to_float(
            extract_from_block(block, min_col, idx, anchor_col, offsets["sales_captured_in_db_pct"])
        )

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": info.model,
                "ticker": info.ticker,
                "model_period": info.model_period,
                "model_date": info.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": "" if last_quarter_used is None else str(last_quarter_used),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    if formula_rows:
        sheet.range((first_data_row, helper_col), (last_data_row, helper_col)).value = None

    return rows


def process_regression_sheet(workbook: xw.Book, info: ModelInfo, source_file: str) -> List[Dict[str, Any]]:
    sheet = pick_sheet(workbook, REGRESSION_SHEET_NAME)
    if sheet is None:
        print(f"SKIPPED {source_file}: missing sheet '{REGRESSION_SHEET_NAME}'")
        return []

    try:
        anchor_row, anchor_col, last_col = find_anchor(sheet, "max")
    except ValueError as error:
        print(f"SKIPPED {source_file}: {error}")
        return []

    offsets = resolve_offsets(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        aliases=REGRESSION_HEADER_ALIASES,
        fallback_offsets=REGRESSION_FALLBACK_OFFSETS,
        last_col=last_col,
    )

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1
    helper_intercept_col = last_col + 2
    helper_slope_col = last_col + 3

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print(f"SKIPPED {source_file}: invalid regression anchor offsets (x_col={x_col}, y_col={y_col})")
        return []

    num_quarters_col = max(1, anchor_col + offsets["num_quarters_used"])
    num_quarters_values = as_column_list(
        sheet.range((first_data_row, num_quarters_col), (last_data_row, num_quarters_col)).value,
        size=N_QUARTERS,
    )

    formula_rows: List[int] = []
    for idx in range(N_QUARTERS):
        row = first_data_row + idx
        num_quarters_used = to_positive_int(num_quarters_values[idx]) or (idx + 1)
        start_row = max(1, row - num_quarters_used + 1)

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},"
            f"R{start_row}C{x_col}:R{row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{row}C{y_col},"
            f"R{start_row}C{x_col}:R{row}C{x_col})"
        )

        set_formula2(sheet.cells(row, helper_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row, helper_slope_col), slope_formula)
        formula_rows.append(row)

    if formula_rows:
        workbook.app.calculate()

    intercept_values = as_column_list(
        sheet.range((first_data_row, helper_intercept_col), (last_data_row, helper_intercept_col)).value,
        size=N_QUARTERS,
    )
    slope_values = as_column_list(
        sheet.range((first_data_row, helper_slope_col), (last_data_row, helper_slope_col)).value,
        size=N_QUARTERS,
    )

    offsets_to_read = {
        offsets["num_quarters_used"],
        offsets["forecast_value"],
        offsets["actual_value"],
        offsets["forecast_max"],
        offsets["forecast_min"],
    }
    min_col = max(1, anchor_col + min(offsets_to_read))
    max_col = max(1, anchor_col + max(offsets_to_read))
    block = as_matrix(
        sheet.range((first_data_row, min_col), (last_data_row, max_col)).value,
        expected_rows=N_QUARTERS,
        expected_cols=max_col - min_col + 1,
    )

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(N_QUARTERS):
        num_quarters_used = (
            to_positive_int(
                extract_from_block(block, min_col, idx, anchor_col, offsets["num_quarters_used"])
            )
            or (idx + 1)
        )
        forecast_value = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["forecast_value"]))
        actual_value_raw = extract_from_block(block, min_col, idx, anchor_col, offsets["actual_value"])
        actual_value = to_float(actual_value_raw)
        forecast_max = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["forecast_max"]))
        forecast_min = to_float(extract_from_block(block, min_col, idx, anchor_col, offsets["forecast_min"]))
        intercept = to_float(intercept_values[idx])
        slope = to_float(slope_values[idx])

        if all(
            value is None
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        signature = (
            num_quarters_used,
            None if intercept is None else round(intercept, 10),
            None if slope is None else round(slope, 10),
            None if forecast_value is None else round(forecast_value, 10),
            None if forecast_max is None else round(forecast_max, 10),
            None if forecast_min is None else round(forecast_min, 10),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": info.model,
                "ticker": info.ticker,
                "model_period": info.model_period,
                "model_date": info.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    sheet.range((first_data_row, helper_intercept_col), (last_data_row, helper_slope_col)).value = None
    return rows


def write_sheet(work_sheet, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    for col_idx, col_name in enumerate(columns, start=1):
        header_cell = work_sheet.cell(row=1, column=col_idx, value=col_name)
        header_cell.font = Font(bold=True)

    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, col_name in enumerate(columns, start=1):
            value = row_data.get(col_name)
            work_sheet.cell(row=row_idx, column=col_idx, value=value)

    work_sheet.freeze_panes = "A2"

    last_col_letter = get_column_letter(len(columns))
    last_row = max(1, len(rows) + 1)
    work_sheet.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, len(rows) + 2):
            value = work_sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        work_sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 45)


def write_output_workbook(
    output_file: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)


def run() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a folder: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_path, output_path)
    generated_prefix = f"{input_path.name}_param"

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"SKIPPED {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"SKIPPED {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED {file_path.name}: not an .xlsx file")
                continue
            if file_path.stem.lower().startswith(generated_prefix):
                print(f"SKIPPED {file_path.name}: generated output workbook")
                continue

            print(f"PROCESSING {file_path.name}")

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                info = parse_file_label(file_path)
                empirical_rows.extend(process_empirical_sheet(workbook, info, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, info, file_path.name))
                processed_files += 1
                print(f"PROCESSED {file_path.name}")
            except Exception as error:
                print(f"SKIPPED {file_path.name}: processing error ({error})")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            app.kill()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"OUTPUT_FILE {output_file}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    run()
