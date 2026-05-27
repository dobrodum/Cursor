from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "quarterly_sales": -7,
    "reported_sales": -6,
    "sales_captured_in_db_pct": -5,
    "growth_rate_pct": -4,
    "avg_penetration_pct": -3,
    "forecast_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -9,
    "actual_value": -4,
    "intercept": -3,
    "slope": -2,
    "forecast_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}

EMPIRICAL_ALIASES = {
    "num_quarters_used": ("num quarters", "n quarters", "quarters used", "# quarters"),
    "last_quarter_used": ("last quarter used", "latest quarter", "quarter used"),
    "quarterly_sales": ("quarterly sales", "qtr sales"),
    "reported_sales": ("reported sales", "actual sales"),
    "sales_captured_in_db_pct": ("sales captured in db", "captured in db", "db capture"),
    "growth_rate_pct": ("growth rate", "growth %"),
    "avg_penetration_pct": ("avg penetration", "average penetration"),
    "forecast_value": ("estimated total sold", "forecast", "tot fcst", "total forecast"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
}

REGRESSION_ALIASES = {
    "num_quarters_used": ("num quarters", "n quarters", "quarters used", "# quarters"),
    "actual_value": ("actual", "reported sales", "actual sales"),
    "intercept": ("intercept",),
    "slope": ("slope",),
    "forecast_value": ("tot fcst w/o sa", "tot fcst without sa", "forecast", "total forecast"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
}

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}
FILE_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<phase>Early|Mid|Late)"
    r"(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
    re.IGNORECASE,
)
MONTH_INDEX = {
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


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalized_label(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def normalize_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values  # already 2D
    return [values]


def numeric(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1]
        try:
            return float(text)
        except ValueError:
            return None
    return None


def range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = numeric(max_value)
    min_num = numeric(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_model_metadata(file_name: str) -> Optional[ModelMetadata]:
    match = FILE_PATTERN.search(file_name)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    phase_raw = match.group("phase").lower()
    month_raw = match.group("month").lower()[:3]
    year = int(match.group("year"))

    if month_raw not in MONTH_INDEX or phase_raw not in DAY_BY_PHASE:
        return None

    phase_title = phase_raw.capitalize()
    month_title = month_raw.capitalize()
    model_period = f"{phase_title}{month_title}_{year}"
    model_date = date(year, MONTH_INDEX[month_raw], DAY_BY_PHASE[phase_raw]).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, target_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM.xlsx"
    candidate = target_folder / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = target_folder / f"{input_folder.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_workbook(workbook: xw.Book) -> None:
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
    except Exception:
        # Last fallback: close without any explicit save flag.
        workbook.close()


def find_anchor_cell(matrix: Sequence[Sequence[Any]]) -> Optional[Tuple[int, int]]:
    for row_idx, row in enumerate(matrix):
        for col_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return row_idx, col_idx
    return None


def resolve_offsets(
    matrix: Sequence[Sequence[Any]],
    anchor_row_idx: int,
    anchor_col_idx: int,
    aliases: Dict[str, Tuple[str, ...]],
    fallback_offsets: Dict[str, int],
) -> Dict[str, int]:
    offsets = dict(fallback_offsets)
    if not matrix:
        return offsets

    header_row = matrix[anchor_row_idx] if anchor_row_idx < len(matrix) else []
    normalized_headers = [normalized_label(cell) for cell in header_row]

    for field_name, terms in aliases.items():
        for col_idx, label in enumerate(normalized_headers):
            if not label:
                continue
            if any(term in label for term in terms):
                offsets[field_name] = col_idx - anchor_col_idx
                break

    return offsets


def get_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if col < 1 or row < 1:
        return None
    return sheet.range((row, col)).value


def set_formula2(sheet: xw.Sheet, row: int, col: int, formula: str) -> bool:
    if col < 1 or row < 1:
        return False
    sheet.range((row, col)).formula2 = formula
    return True


def extract_empirical_candidates(
    workbook: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    used_range = sheet.used_range
    matrix = normalize_matrix(used_range.value)
    anchor = find_anchor_cell(matrix)
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row_idx, anchor_col_idx = anchor
    anchor_row = used_range.row + anchor_row_idx
    anchor_col = used_range.column + anchor_col_idx
    offsets = resolve_offsets(matrix, anchor_row_idx, anchor_col_idx, EMPIRICAL_ALIASES, EMPIRICAL_FALLBACK_OFFSETS)
    start_row = anchor_row + 1

    wrote_formulas = False
    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters = get_cell(sheet, row, anchor_col + offsets["num_quarters_used"])
        n = int(numeric(num_quarters) or (i + 1))
        avg_col = anchor_col + offsets["avg_penetration_pct"]
        source_col = anchor_col + offsets["sales_captured_in_db_pct"]
        rel_col = source_col - avg_col
        if n > 0 and avg_col >= 1 and source_col >= 1:
            formula = f'=IFERROR(AVERAGE(R[-{n}]C[{rel_col}]:R[-1]C[{rel_col}]),"")'
            wrote_formulas = set_formula2(sheet, row, avg_col, formula) or wrote_formulas

    if wrote_formulas:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters_used = get_cell(sheet, row, anchor_col + offsets["num_quarters_used"])
        last_quarter_used = get_cell(sheet, row, anchor_col + offsets["last_quarter_used"])
        forecast_value = get_cell(sheet, row, anchor_col + offsets["forecast_value"])
        actual_value = get_cell(sheet, row, anchor_col + offsets["reported_sales"])
        forecast_max = get_cell(sheet, row, anchor_col + offsets["forecast_max"])
        forecast_min = get_cell(sheet, row, anchor_col + offsets["forecast_min"])
        avg_penetration = get_cell(sheet, row, anchor_col + offsets["avg_penetration_pct"])
        quarterly_sales = get_cell(sheet, row, anchor_col + offsets["quarterly_sales"])
        reported_sales = get_cell(sheet, row, anchor_col + offsets["reported_sales"])
        growth_rate = get_cell(sheet, row, anchor_col + offsets["growth_rate_pct"])
        captured_in_db = get_cell(sheet, row, anchor_col + offsets["sales_captured_in_db_pct"])

        is_empty = all(
            value in (None, "")
            for value in (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration,
            )
        )
        if is_empty:
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_in_db,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    workbook: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    used_range = sheet.used_range
    matrix = normalize_matrix(used_range.value)
    anchor = find_anchor_cell(matrix)
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row_idx, anchor_col_idx = anchor
    anchor_row = used_range.row + anchor_row_idx
    anchor_col = used_range.column + anchor_col_idx
    offsets = resolve_offsets(matrix, anchor_row_idx, anchor_col_idx, REGRESSION_ALIASES, REGRESSION_FALLBACK_OFFSETS)
    start_row = anchor_row + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    wrote_formulas = False
    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters = get_cell(sheet, row, anchor_col + offsets["num_quarters_used"])
        n = int(numeric(num_quarters) or (i + 1))
        data_start_row = max(1, row - n)
        data_end_row = row - 1
        intercept_col = anchor_col + offsets["intercept"]
        slope_col = anchor_col + offsets["slope"]

        if n > 0 and data_end_row >= data_start_row:
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),\"\")"
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),\"\")"
            )
            wrote_formulas = set_formula2(sheet, row, intercept_col, intercept_formula) or wrote_formulas
            wrote_formulas = set_formula2(sheet, row, slope_col, slope_formula) or wrote_formulas

    if wrote_formulas:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters_used = get_cell(sheet, row, anchor_col + offsets["num_quarters_used"])
        intercept = get_cell(sheet, row, anchor_col + offsets["intercept"])
        slope = get_cell(sheet, row, anchor_col + offsets["slope"])
        forecast_value = get_cell(sheet, row, anchor_col + offsets["forecast_value"])
        actual_value = get_cell(sheet, row, anchor_col + offsets["actual_value"])
        forecast_max = get_cell(sheet, row, anchor_col + offsets["forecast_max"])
        forecast_min = get_cell(sheet, row, anchor_col + offsets["forecast_min"])

        is_empty = all(
            value in (None, "")
            for value in (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        )
        if is_empty:
            continue

        signature = (
            round(numeric(num_quarters_used) or 0.0, 8),
            round(numeric(forecast_value) or 0.0, 8),
            round(numeric(forecast_max) or 0.0, 8),
            round(numeric(forecast_min) or 0.0, 8),
            round(numeric(intercept) or 0.0, 8),
            round(numeric(slope) or 0.0, 8),
        )
        if i == N_QUARTERS - 1 and previous_signature == signature:
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
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def apply_sheet_formatting(sheet) -> None:
    sheet.freeze_panes = "A2"
    if sheet.max_row >= 1 and sheet.max_column >= 1:
        sheet.auto_filter.ref = sheet.dimensions

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for col_idx in range(1, sheet.max_column + 1):
        max_len = 0
        for row_idx in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = min(max(max_len + 2, 12), 42)
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: Iterable[Dict[str, Any]],
    regression_rows: Iterable[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(column) for column in EMPIRICAL_COLUMNS])
    apply_sheet_formatting(empirical_sheet)

    regression_sheet = workbook.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(column) for column in REGRESSION_COLUMNS])
    apply_sheet_formatting(regression_sheet)

    workbook.save(output_path)


def run() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a directory: {in_dir}")

    output_path = next_output_path(in_dir, out_dir)
    output_name_pattern = re.compile(
        rf"^{re.escape(in_dir.name)}_PARAM(?:\.\d+)?\.xlsx$",
        re.IGNORECASE,
    )

    files_processed = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel lock file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            if output_name_pattern.match(file_path.name):
                print(f"Skipped {file_path.name}: generated output workbook")
                continue

            metadata = parse_model_metadata(file_path.name)
            if metadata is None:
                print(f"Skipped {file_path.name}: filename does not match expected model pattern")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_empirical = extract_empirical_candidates(workbook, metadata, file_path.name)
                file_regression = extract_regression_candidates(workbook, metadata, file_path.name)
                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                files_processed += 1
                print(
                    f"Processed {file_path.name}: "
                    f"{len(file_empirical)} empirical rows, {len(file_regression)} regression rows"
                )
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
