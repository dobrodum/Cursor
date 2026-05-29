from __future__ import annotations

import calendar
import datetime as dt
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these two paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def coerce_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def to_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def subtract_if_numeric(left: Any, right: Any) -> Any:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def month_name_to_number(month_token: str) -> Optional[int]:
    token = month_token.strip().lower()
    if not token:
        return None

    month_lookup: Dict[str, int] = {}
    for index in range(1, 13):
        month_lookup[calendar.month_name[index].lower()] = index
        month_lookup[calendar.month_abbr[index].lower()] = index

    if token in month_lookup:
        return month_lookup[token]
    return month_lookup.get(token[:3])


def parse_filename_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) > 1 else "UNKNOWN"
    ticker = re.sub(r"\s+", "", ticker)

    period_source = parts[2] if len(parts) > 2 else stem
    period_source = period_source.split("_")[0]

    match = re.search(r"(?i)(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_source)
    model_period = period_source
    model_date = ""

    if match:
        window = match.group(1).title()
        month_token = match.group(2)
        year = int(match.group(3))

        month_number = month_name_to_number(month_token)
        if month_number:
            month_abbr = calendar.month_abbr[month_number]
            day = {"Early": 5, "Mid": 15, "Late": 25}[window]
            model_period = f"{window}{month_abbr}_{year}"
            model_date = dt.date(year, month_number, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def output_path_with_increment(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    suffix = 1

    while candidate.exists():
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1

    return candidate


def locate_anchor_and_headers(
    sheet: xw.Sheet, anchor_text: str = "max"
) -> Optional[Tuple[int, int, Dict[str, int], int]]:
    used = sheet.used_range
    values = coerce_2d(used.value)
    base_row = used.row
    base_col = used.column

    if not values:
        return None

    anchor_row: Optional[int] = None
    anchor_col: Optional[int] = None
    target = normalize_text(anchor_text)

    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                anchor_row = base_row + row_idx
                anchor_col = base_col + col_idx
                break
        if anchor_row is not None and anchor_col is not None:
            break

    if anchor_row is None or anchor_col is None:
        return None

    header_map: Dict[str, int] = {}
    header_row_index = anchor_row - base_row
    header_values = values[header_row_index]
    for col_offset, header_cell in enumerate(header_values):
        key = normalize_text(header_cell)
        if key and key not in header_map:
            header_map[key] = base_col + col_offset

    return anchor_row, anchor_col, header_map, used.last_cell.column


def resolve_column(
    header_map: Dict[str, int], token_sets: Sequence[Sequence[str]], fallback_col: int
) -> int:
    for header_text, column in header_map.items():
        for tokens in token_sets:
            if all(token in header_text for token in tokens):
                return column
    return fallback_col


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.api.Formula2 = formula


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
    except Exception:
        pass


def row_has_any_data(values: Iterable[Any]) -> bool:
    return any(value not in (None, "") for value in values)


def clean_signature_value(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 10)
    return value


def extract_empirical_rows(
    workbook: xw.Book, sheet: xw.Sheet, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    located = locate_anchor_and_headers(sheet, anchor_text="max")
    if not located:
        print(f"Skipped empirical extraction in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col, header_map, last_used_col = located
    row_start = anchor_row + 1

    num_quarters_col = resolve_column(
        header_map,
        token_sets=[
            ("num", "quarter"),
            ("quarters", "used"),
            ("#", "qtr"),
        ],
        fallback_col=anchor_col - 8,
    )
    last_quarter_col = resolve_column(
        header_map,
        token_sets=[("last", "quarter")],
        fallback_col=anchor_col - 9,
    )
    estimated_total_col = resolve_column(
        header_map,
        token_sets=[
            ("estimated", "total", "sold"),
            ("est", "total", "sold"),
            ("total", "sold"),
        ],
        fallback_col=anchor_col - 2,
    )
    reported_sales_col = resolve_column(
        header_map,
        token_sets=[("reported", "sales"), ("actual", "sales")],
        fallback_col=anchor_col - 1,
    )
    min_col = resolve_column(
        header_map,
        token_sets=[("min",)],
        fallback_col=anchor_col + 1,
    )
    quarterly_sales_col = resolve_column(
        header_map,
        token_sets=[("quarterly", "sales")],
        fallback_col=anchor_col - 6,
    )
    growth_rate_col = resolve_column(
        header_map,
        token_sets=[("growth",), ("growth", "rate")],
        fallback_col=anchor_col - 4,
    )
    sales_captured_col = resolve_column(
        header_map,
        token_sets=[
            ("captured", "db"),
            ("sales", "captured", "db"),
            ("captured", "database"),
        ],
        fallback_col=anchor_col - 3,
    )

    avg_pen_temp_col = last_used_col + 2
    formulas_written = False

    for offset in range(N_QUARTERS):
        row = row_start + offset
        formula = (
            f'=IFERROR(AVERAGE(R{row_start}C{sales_captured_col}:R{row}C{sales_captured_col}),"")'
        )
        set_formula2(sheet.cells(row, avg_pen_temp_col), formula)
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for offset in range(N_QUARTERS):
        row = row_start + offset

        num_quarters_used = sheet.cells(row, num_quarters_col).value
        if num_quarters_used in (None, ""):
            num_quarters_used = offset + 1

        last_quarter_used = sheet.cells(row, last_quarter_col).value
        forecast_value = sheet.cells(row, estimated_total_col).value
        actual_value = sheet.cells(row, reported_sales_col).value
        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, min_col).value
        avg_penetration_pct = sheet.cells(row, avg_pen_temp_col).value
        quarterly_sales = sheet.cells(row, quarterly_sales_col).value
        reported_sales = sheet.cells(row, reported_sales_col).value
        growth_rate_pct = sheet.cells(row, growth_rate_col).value
        sales_captured_in_db_pct = sheet.cells(row, sales_captured_col).value

        if not row_has_any_data(
            [
                num_quarters_used,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            ]
        ):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book, sheet: xw.Sheet, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    located = locate_anchor_and_headers(sheet, anchor_text="max")
    if not located:
        print(f"Skipped regression extraction in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col, header_map, last_used_col = located
    row_start = anchor_row + 1

    num_quarters_col = resolve_column(
        header_map,
        token_sets=[
            ("num", "quarter"),
            ("quarters", "used"),
            ("#", "qtr"),
        ],
        fallback_col=anchor_col - 8,
    )
    forecast_total_col = resolve_column(
        header_map,
        token_sets=[
            ("tot", "fcst", "w/o", "sa"),
            ("tot", "fcst", "without", "sa"),
            ("forecast", "without", "sa"),
        ],
        fallback_col=anchor_col - 1,
    )
    actual_value_col = resolve_column(
        header_map,
        token_sets=[("actual",), ("reported", "sales")],
        fallback_col=anchor_col - 2,
    )
    min_col = resolve_column(
        header_map,
        token_sets=[("min",)],
        fallback_col=anchor_col + 1,
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    intercept_temp_col = last_used_col + 2
    slope_temp_col = last_used_col + 3
    formulas_written = False

    for offset in range(N_QUARTERS):
        row = row_start + offset
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{row_start}C{y_col}:R{row}C{y_col},'
            f'R{row_start}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{row_start}C{y_col}:R{row}C{y_col},'
            f'R{row_start}C{x_col}:R{row}C{x_col}),"")'
        )
        set_formula2(sheet.cells(row, intercept_temp_col), intercept_formula)
        set_formula2(sheet.cells(row, slope_temp_col), slope_formula)
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for offset in range(N_QUARTERS):
        row = row_start + offset
        num_quarters_used = sheet.cells(row, num_quarters_col).value
        if num_quarters_used in (None, ""):
            num_quarters_used = offset + 1

        forecast_value = sheet.cells(row, forecast_total_col).value
        actual_value = sheet.cells(row, actual_value_col).value
        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, min_col).value
        intercept = sheet.cells(row, intercept_temp_col).value
        slope = sheet.cells(row, slope_temp_col).value

        if not row_has_any_data(
            [num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope]
        ):
            continue

        signature = tuple(
            clean_signature_value(value)
            for value in (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        )
        if offset == N_QUARTERS - 1 and previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    worksheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]
) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(column, "") for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = (
        f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
    )

    for column_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_index in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_index, column=column_index).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(12, max_length + 2), 50
        )


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    output_workbook = Workbook()
    empirical_ws = output_workbook.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = output_workbook.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    output_workbook.save(output_path)


def iter_candidate_files(input_folder: Path) -> Iterable[Path]:
    for file_path in sorted(input_folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_path_with_increment(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_candidate_files(input_dir):
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path.name)

                try:
                    empirical_sheet = workbook.sheets[EMPIRICAL_SHEET_NAME]
                    empirical_rows.extend(
                        extract_empirical_rows(
                            workbook=workbook,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )
                except Exception as exc:
                    print(
                        f"Skipped empirical extraction in {file_path.name}: "
                        f"{EMPIRICAL_SHEET_NAME} issue ({exc})"
                    )

                try:
                    regression_sheet = workbook.sheets[REGRESSION_SHEET_NAME]
                    regression_rows.extend(
                        extract_regression_rows(
                            workbook=workbook,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )
                except Exception as exc:
                    print(
                        f"Skipped regression extraction in {file_path.name}: "
                        f"{REGRESSION_SHEET_NAME} issue ({exc})"
                    )

                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
