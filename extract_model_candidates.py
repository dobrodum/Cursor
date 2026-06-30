from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these two paths before running.
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

N_QUARTERS = 10

DAY_BY_PHASE = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_TO_NUMBER = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

FILE_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<phase>Early|Mid|Late)\s*(?P<month>[A-Za-z]{3})\s*(?P<year>\d{4})",
    flags=re.IGNORECASE,
)

# Fallback offsets are relative to the "max" anchor column.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -11,
    "last_quarter_used": -10,
    "forecast_value": -6,
    "actual_value": -5,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -7,
    "quarterly_sales": -9,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -10,
    "forecast_value": -6,
    "actual_value": -5,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": -3,
    "slope": -2,
}


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def normalize_header(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def parse_file_metadata(file_path: Path) -> FileMetadata | None:
    match = FILE_PATTERN.search(file_path.stem)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    phase = match.group("phase").title()
    month = match.group("month").title()
    year = int(match.group("year"))

    month_number = MONTH_TO_NUMBER.get(month)
    day = DAY_BY_PHASE.get(phase)
    if month_number is None or day is None:
        return None

    model_period = f"{phase}{month}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def safe_close_workbook(wb: Any) -> None:
    if wb is None:
        return

    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def get_sheet_by_name(wb: Any, sheet_name: str) -> Any | None:
    target = normalize_text(sheet_name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == target:
            return sheet
    return None


def matrixify(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor(sheet: Any, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = matrixify(used.value)
    target = normalize_text(anchor_text)
    for row_idx, row in enumerate(values):
        for col_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return used.row + row_idx, used.column + col_idx
    return None


def build_header_map(sheet: Any, header_row: int, anchor_col: int, span: int = 24) -> dict[str, int]:
    start_col = max(1, anchor_col - span)
    end_col = anchor_col + span
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    row_values = values if isinstance(values, list) else [values]

    headers: dict[str, int] = {}
    for idx, value in enumerate(row_values):
        key = normalize_header(value)
        if key and key not in headers:
            headers[key] = start_col + idx
    return headers


def resolve_column(
    header_map: dict[str, int],
    aliases: list[str],
    anchor_col: int,
    fallback_offset: int | None,
) -> int | None:
    for alias in aliases:
        if alias in header_map:
            return header_map[alias]
    if fallback_offset is None:
        return None
    return anchor_col + fallback_offset


def get_value(sheet: Any, row: int, col: int | None) -> Any:
    if col is None or col < 1:
        return None
    return sheet.cells(row, col).value


def to_int(value: Any, default: int) -> int:
    try:
        number = int(float(value))
    except (TypeError, ValueError):
        return default
    return number


def to_float(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def relative_ref(offset: int) -> str:
    if offset == 0:
        return "RC"
    return f"RC[{offset}]"


def write_formula2(sheet: Any, row: int, col: int | None, formula: str) -> bool:
    if col is None or col < 1:
        return False
    try:
        sheet.cells(row, col).formula2 = formula
        return True
    except Exception:
        return False


def has_any_data(*values: Any) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return True
    return False


def calculate_range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def extract_empirical_rows(wb: Any, meta: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical extraction for {source_file}: sheet 'Empirical Model' not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)

    col_map = {
        "num_quarters_used": resolve_column(
            header_map,
            ["num_quarters_used", "quarters_used", "n_quarters", "num_qtrs"],
            anchor_col,
            EMPIRICAL_OFFSETS["num_quarters_used"],
        ),
        "last_quarter_used": resolve_column(
            header_map,
            ["last_quarter_used", "last_quarter", "last_qtr"],
            anchor_col,
            EMPIRICAL_OFFSETS["last_quarter_used"],
        ),
        "forecast_value": resolve_column(
            header_map,
            ["estimated_total_sold", "forecast_value", "tot_fcst_wo_sa", "tot_fcst_w_o_sa", "tot_fcst"],
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_value"],
        ),
        "actual_value": resolve_column(
            header_map,
            ["actual_value", "reported_sales", "actual_sales"],
            anchor_col,
            EMPIRICAL_OFFSETS["actual_value"],
        ),
        "forecast_max": resolve_column(
            header_map,
            ["max", "forecast_max"],
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_max"],
        ),
        "forecast_min": resolve_column(
            header_map,
            ["min", "forecast_min"],
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_min"],
        ),
        "avg_penetration_pct": resolve_column(
            header_map,
            ["avg_penetration_pct", "average_penetration_pct", "avg_penetration"],
            anchor_col,
            EMPIRICAL_OFFSETS["avg_penetration_pct"],
        ),
        "quarterly_sales": resolve_column(
            header_map,
            ["quarterly_sales", "qtr_sales", "quarter_sales"],
            anchor_col,
            EMPIRICAL_OFFSETS["quarterly_sales"],
        ),
        "reported_sales": resolve_column(
            header_map,
            ["reported_sales", "actual_sales"],
            anchor_col,
            EMPIRICAL_OFFSETS["reported_sales"],
        ),
        "growth_rate_pct": resolve_column(
            header_map,
            ["growth_rate_pct", "growth_pct", "growth_rate"],
            anchor_col,
            EMPIRICAL_OFFSETS["growth_rate_pct"],
        ),
        "sales_captured_in_db_pct": resolve_column(
            header_map,
            ["sales_captured_in_db_pct", "sales_captured_db_pct", "sales_captured_in_db"],
            anchor_col,
            EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
        ),
    }

    # Use temporary formula writes for avg penetration and recalc once.
    formula_write_count = 0
    avg_col = col_map["avg_penetration_pct"]
    penetration_start_col = anchor_col - 11
    for index in range(N_QUARTERS):
        row = anchor_row + 1 + index
        num_quarters = to_int(get_value(sheet, row, col_map["num_quarters_used"]), index + 1)
        num_quarters = max(1, min(num_quarters, N_QUARTERS))

        if avg_col is None:
            continue
        if penetration_start_col < 1:
            continue

        start_rel = penetration_start_col - avg_col
        end_rel = start_rel + num_quarters - 1
        formula = f'=IFERROR(AVERAGE({relative_ref(start_rel)}:{relative_ref(end_rel)}),"")'
        if write_formula2(sheet, row, avg_col, formula):
            formula_write_count += 1

    if formula_write_count:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for index in range(N_QUARTERS):
        row = anchor_row + 1 + index
        num_quarters_used = get_value(sheet, row, col_map["num_quarters_used"])
        if num_quarters_used in (None, ""):
            num_quarters_used = index + 1

        last_quarter_used = get_value(sheet, row, col_map["last_quarter_used"])
        forecast_value = get_value(sheet, row, col_map["forecast_value"])
        actual_value = get_value(sheet, row, col_map["actual_value"])
        forecast_max = get_value(sheet, row, col_map["forecast_max"])
        forecast_min = get_value(sheet, row, col_map["forecast_min"])
        avg_penetration_pct = get_value(sheet, row, col_map["avg_penetration_pct"])
        quarterly_sales = get_value(sheet, row, col_map["quarterly_sales"])
        reported_sales = get_value(sheet, row, col_map["reported_sales"])
        growth_rate_pct = get_value(sheet, row, col_map["growth_rate_pct"])
        sales_captured_in_db_pct = get_value(sheet, row, col_map["sales_captured_in_db_pct"])

        if not has_any_data(
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            avg_penetration_pct,
            quarterly_sales,
            reported_sales,
            growth_rate_pct,
            sales_captured_in_db_pct,
        ):
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calculate_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def build_regression_formula(function_name: str, target_col: int, y_col: int, x_col: int, width: int) -> str:
    y_start_rel = y_col - target_col
    y_end_rel = y_start_rel + width - 1
    x_start_rel = x_col - target_col
    x_end_rel = x_start_rel + width - 1
    y_range = f"{relative_ref(y_start_rel)}:{relative_ref(y_end_rel)}"
    x_range = f"{relative_ref(x_start_rel)}:{relative_ref(x_end_rel)}"
    return f'=IFERROR({function_name}({y_range},{x_range}),"")'


def extract_regression_rows(wb: Any, meta: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped regression extraction for {source_file}: sheet 'Regression Model' not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    header_map = build_header_map(sheet, anchor_row, anchor_col)

    col_map = {
        "num_quarters_used": resolve_column(
            header_map,
            ["num_quarters_used", "quarters_used", "n_quarters", "num_qtrs"],
            anchor_col,
            REGRESSION_OFFSETS["num_quarters_used"],
        ),
        "forecast_value": resolve_column(
            header_map,
            ["tot_fcst_wo_sa", "tot_fcst_w_o_sa", "forecast_value", "tot_fcst_without_sa", "tot_fcst"],
            anchor_col,
            REGRESSION_OFFSETS["forecast_value"],
        ),
        "actual_value": resolve_column(
            header_map,
            ["actual_value", "reported_sales", "actual_sales"],
            anchor_col,
            REGRESSION_OFFSETS["actual_value"],
        ),
        "forecast_max": resolve_column(
            header_map,
            ["max", "forecast_max"],
            anchor_col,
            REGRESSION_OFFSETS["forecast_max"],
        ),
        "forecast_min": resolve_column(
            header_map,
            ["min", "forecast_min"],
            anchor_col,
            REGRESSION_OFFSETS["forecast_min"],
        ),
        "intercept": resolve_column(
            header_map,
            ["intercept"],
            anchor_col,
            REGRESSION_OFFSETS["intercept"],
        ),
        "slope": resolve_column(
            header_map,
            ["slope"],
            anchor_col,
            REGRESSION_OFFSETS["slope"],
        ),
    }

    formula_write_count = 0
    for index in range(N_QUARTERS):
        row = anchor_row + 1 + index
        num_quarters = to_int(get_value(sheet, row, col_map["num_quarters_used"]), index + 1)
        num_quarters = max(2, min(num_quarters, N_QUARTERS))

        intercept_col = col_map["intercept"]
        slope_col = col_map["slope"]
        if intercept_col is not None:
            intercept_formula = build_regression_formula("INTERCEPT", intercept_col, y_col, x_col, num_quarters)
            if write_formula2(sheet, row, intercept_col, intercept_formula):
                formula_write_count += 1
        if slope_col is not None:
            slope_formula = build_regression_formula("SLOPE", slope_col, y_col, x_col, num_quarters)
            if write_formula2(sheet, row, slope_col, slope_formula):
                formula_write_count += 1

    if formula_write_count:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for index in range(N_QUARTERS):
        row = anchor_row + 1 + index
        num_quarters_used = get_value(sheet, row, col_map["num_quarters_used"])
        if num_quarters_used in (None, ""):
            num_quarters_used = index + 1

        intercept = get_value(sheet, row, col_map["intercept"])
        slope = get_value(sheet, row, col_map["slope"])
        forecast_value = get_value(sheet, row, col_map["forecast_value"])
        actual_value = get_value(sheet, row, col_map["actual_value"])
        forecast_max = get_value(sheet, row, col_map["forecast_max"])
        forecast_min = get_value(sheet, row, col_map["forecast_min"])

        if not has_any_data(intercept, slope, forecast_value, forecast_max, forecast_min):
            continue

        signature = (
            num_quarters_used,
            intercept,
            slope,
            forecast_value,
            forecast_max,
            forecast_min,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calculate_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        width = max(12, len(column_name) + 2)
        for excel_row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
            value = excel_row[0].value
            if value is None:
                continue
            width = max(width, min(60, len(str(value)) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(60, width)


def write_output_workbook(path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    wb = Workbook()
    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(path)


def process_workbooks(input_folder: Path, output_folder: Path) -> Path:
    if not input_folder.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_folder}")
    if not input_folder.is_dir():
        raise NotADirectoryError(f"Input path is not a folder: {input_folder}")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_folder.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            meta = parse_file_metadata(file_path)
            if meta is None:
                print(f"Skipped {file_path.name}: filename format not recognized")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                files_processed += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: workbook processing error: {exc}")
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    output_path = next_output_path(input_folder, output_folder)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")
    return output_path


if __name__ == "__main__":
    process_workbooks(Path(input_dir), Path(output_dir))
