from __future__ import annotations

import re
import traceback
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

# -----------------------------
# Constants
# -----------------------------
N_QUARTERS = 10
PHASE_DAY = {"Early": 5, "Mid": 15, "Late": 25}

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def safe_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_int(value: Any) -> int | None:
    num = safe_float(value)
    if num is None:
        return None
    return int(round(num))


def has_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str) and not value.strip():
        return False
    return True


def subtract_numbers(left: Any, right: Any) -> float | None:
    left_num = safe_float(left)
    right_num = safe_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def signature_value(value: Any) -> Any:
    num = safe_float(value)
    if num is not None:
        return round(num, 10)
    if value is None:
        return None
    return str(value).strip()


def parse_file_label(filename: str) -> dict[str, str]:
    stem = Path(filename).stem
    compact = stem.replace("_", " ")
    pattern = re.compile(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(compact)

    if not match:
        ticker = "UNKNOWN"
        model_period = "UNKNOWN"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return {
            "model": model,
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
        }

    ticker = match.group("ticker").upper()
    phase = match.group("phase").title()
    month_abbr = match.group("month").title()
    year = int(match.group("year"))

    try:
        month_num = datetime.strptime(f"{month_abbr} 1 {year}", "%b %d %Y").month
        day = PHASE_DAY[phase]
        model_date = date(year, month_num, day).isoformat()
    except ValueError:
        month_num = None
        model_date = ""

    model_period = f"{phase}{month_abbr}_{year}"
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def get_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    start_row = used.row
    start_col = used.column

    if not isinstance(values, list):
        matrix = [[values]]
    elif values and not isinstance(values[0], list):
        matrix = [values]
    else:
        matrix = values

    target = normalize_text(anchor_text)
    for row_index, row_values in enumerate(matrix):
        for col_index, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                return start_row + row_index, start_col + col_index
    return None


def build_header_map(
    sheet: xw.Sheet,
    header_row: int,
    anchor_col: int,
    search_width: int = 35,
) -> dict[str, int]:
    left_col = max(1, anchor_col - search_width)
    right_col = anchor_col + search_width
    values = sheet.range((header_row, left_col), (header_row, right_col)).value

    if values is None:
        return {}
    if isinstance(values, list):
        row_values = values
    else:
        row_values = [values]

    mapping: dict[str, int] = {}
    for idx, value in enumerate(row_values):
        key = normalize_text(value)
        if key:
            mapping.setdefault(key, left_col + idx)
    return mapping


def find_column(header_map: dict[str, int], aliases: list[str]) -> int | None:
    normalized_aliases = [normalize_text(alias) for alias in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for key, col in header_map.items():
        for alias in normalized_aliases:
            if alias and alias in key:
                return col
    return None


def get_cell_value(sheet: xw.Sheet, row: int, col: int | None) -> Any:
    if col is None or row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def safe_close_workbook(wb: xw.Book) -> None:
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


def collect_empirical_rows(
    wb: xw.Book,
    file_meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        print(f"  - skipped Empirical Model in {source_file}: sheet missing")
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  - skipped Empirical Model in {source_file}: 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    used_last_col = sheet.used_range.last_cell.column

    min_col = find_column(header_map, ["min"]) or (anchor_col + 1)
    forecast_col = find_column(
        header_map,
        [
            "estimated total sold",
            "total forecast",
            "tot fcst",
            "forecast value",
        ],
    ) or (anchor_col - 1)
    reported_col = find_column(
        header_map,
        ["reported sales", "actual value", "actual", "reported"],
    ) or (anchor_col - 2)
    quarterly_sales_col = find_column(
        header_map,
        ["quarterly sales", "quarter sales", "sales"],
    ) or (anchor_col - 3)
    growth_rate_col = find_column(
        header_map,
        ["growth rate pct", "growth rate", "growth"],
    ) or (anchor_col - 4)
    sales_captured_col = find_column(
        header_map,
        ["sales captured in db pct", "sales captured", "captured in db"],
    ) or (anchor_col - 5)
    last_quarter_col = find_column(
        header_map,
        ["last quarter used", "last quarter"],
    ) or (anchor_col - 6)
    num_quarters_col = find_column(
        header_map,
        ["num quarters used", "number of quarters used", "quarters used"],
    ) or (anchor_col - 7)

    penetration_cols = [
        col
        for header, col in header_map.items()
        if "penetration" in header and "avg" not in header
    ]
    if penetration_cols:
        penetration_start_col = min(penetration_cols)
        penetration_end_col = max(penetration_cols)
    else:
        penetration_start_col = max(1, anchor_col - N_QUARTERS)
        penetration_end_col = max(1, anchor_col - 1)

    temp_avg_col = max(used_last_col + 2, anchor_col + 2)
    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    start_offset = penetration_start_col - temp_avg_col
    end_offset = penetration_end_col - temp_avg_col
    if end_offset < start_offset:
        start_offset, end_offset = end_offset, start_offset

    for row in range(data_start_row, data_end_row + 1):
        formula = f'=IFERROR(AVERAGE(RC[{start_offset}]:RC[{end_offset}]),"")'
        sheet.cells(row, temp_avg_col).formula2 = formula

    wb.app.calculate()

    for idx, row in enumerate(range(data_start_row, data_end_row + 1), start=1):
        forecast_max = get_cell_value(sheet, row, anchor_col)
        forecast_min = get_cell_value(sheet, row, min_col)
        forecast_value = get_cell_value(sheet, row, forecast_col)
        reported_sales = get_cell_value(sheet, row, reported_col)
        quarterly_sales = get_cell_value(sheet, row, quarterly_sales_col)
        growth_rate = get_cell_value(sheet, row, growth_rate_col)
        sales_captured = get_cell_value(sheet, row, sales_captured_col)
        last_quarter_used = get_cell_value(sheet, row, last_quarter_col)
        num_quarters_used = get_cell_value(sheet, row, num_quarters_col)
        avg_penetration = get_cell_value(sheet, row, temp_avg_col)

        if not any(
            has_value(v)
            for v in (
                forecast_max,
                forecast_min,
                forecast_value,
                reported_sales,
                quarterly_sales,
                avg_penetration,
            )
        ):
            continue

        row_data = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": safe_int(num_quarters_used) or idx,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_numbers(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": source_file,
        }
        rows.append(row_data)

    try:
        sheet.range((data_start_row, temp_avg_col), (data_end_row, temp_avg_col)).clear_contents()
    except Exception:
        pass

    return rows


def collect_regression_rows(
    wb: xw.Book,
    file_meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        print(f"  - skipped Regression Model in {source_file}: sheet missing")
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  - skipped Regression Model in {source_file}: 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    used_last_col = sheet.used_range.last_cell.column

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print(
            f"  - skipped Regression Model in {source_file}: "
            "anchor position does not allow x/y column offsets"
        )
        return rows

    min_col = find_column(header_map, ["min"]) or (anchor_col + 1)
    forecast_col = find_column(
        header_map,
        ["tot fcst w/o sa", "total forecast without sa", "tot fcst wo sa", "forecast value"],
    ) or (anchor_col - 1)
    num_quarters_col = find_column(
        header_map,
        ["num quarters used", "number of quarters used", "quarters used"],
    ) or (anchor_col - 2)
    actual_col = find_column(
        header_map,
        ["actual value", "actual", "reported sales"],
    )

    temp_intercept_col = max(used_last_col + 2, anchor_col + 2)
    temp_slope_col = temp_intercept_col + 1

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    for idx, row in enumerate(range(data_start_row, data_end_row + 1), start=1):
        q_used = safe_int(get_cell_value(sheet, row, num_quarters_col)) or idx
        q_used = max(2, min(N_QUARTERS, q_used))
        start_row = max(data_start_row, row - q_used + 1)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        sheet.cells(row, temp_intercept_col).formula2 = intercept_formula
        sheet.cells(row, temp_slope_col).formula2 = slope_formula

    wb.app.calculate()

    previous_signature: tuple[Any, ...] | None = None
    for idx, row in enumerate(range(data_start_row, data_end_row + 1), start=1):
        num_quarters_used = safe_int(get_cell_value(sheet, row, num_quarters_col)) or idx
        forecast_value = get_cell_value(sheet, row, forecast_col)
        actual_value = get_cell_value(sheet, row, actual_col)
        forecast_max = get_cell_value(sheet, row, anchor_col)
        forecast_min = get_cell_value(sheet, row, min_col)
        intercept = get_cell_value(sheet, row, temp_intercept_col)
        slope = get_cell_value(sheet, row, temp_slope_col)

        if not any(
            has_value(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        ):
            continue

        row_data = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_numbers(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        signature = (
            row_data["num_quarters_used"],
            signature_value(row_data["intercept"]),
            signature_value(row_data["slope"]),
            signature_value(row_data["forecast_value"]),
            signature_value(row_data["forecast_max"]),
            signature_value(row_data["forecast_min"]),
        )

        if previous_signature is not None and signature == previous_signature:
            continue

        previous_signature = signature
        rows.append(row_data)

    try:
        sheet.range((data_start_row, temp_intercept_col), (data_end_row, temp_slope_col)).clear_contents()
    except Exception:
        pass

    return rows


def write_table(sheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, header in enumerate(headers, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=row_data.get(header))

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def get_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"

    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def create_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_table(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)
    write_table(regression_sheet, REGRESSION_HEADERS, regression_rows)

    workbook.save(output_file)


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_path}")

    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue
            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_meta = parse_file_label(file_path.name)
                empirical_rows.extend(collect_empirical_rows(wb, file_meta, file_path.name))
                regression_rows.extend(collect_regression_rows(wb, file_meta, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
                traceback.print_exc()
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_file = get_output_path(input_path, output_path)
    create_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
