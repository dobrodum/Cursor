from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw


# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"
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


DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
MONTH_ABBR_TO_NUM = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        cleaned = str(value).replace(",", "").strip()
        return float(cleaned)
    except Exception:
        return None


def maybe_blank(value: Any) -> Any:
    return "" if value is None else value


def safe_subtract(left: Any, right: Any) -> Any:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def approx_equal(left: Any, right: Any, tolerance: float = 1e-9) -> bool:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is not None and right_num is not None:
        return abs(left_num - right_num) <= tolerance
    return maybe_blank(left) == maybe_blank(right)


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    pieces = [piece.strip() for piece in stem.split(" - ")]
    ticker = pieces[1].strip() if len(pieces) >= 2 else ""
    period_token = pieces[2].strip() if len(pieces) >= 3 else ""
    period_token = period_token.split("_")[0]

    model_period = ""
    model_date = ""
    period_match = re.match(
        r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", period_token, flags=re.IGNORECASE
    )
    if period_match:
        period_name = period_match.group(1).title()
        month_chunk = period_match.group(2)
        month_abbr = month_chunk[:3].title()
        year = int(period_match.group(3))
        month_number = MONTH_ABBR_TO_NUM.get(month_abbr.lower())
        day = DAY_BY_PERIOD[period_name.lower()]
        if month_number is not None:
            model_period = f"{period_name}{month_abbr}_{year}"
            model_date = date(year, month_number, day).isoformat()

    if not model_period:
        model_period = period_token or "unknown_period"
    if not model_date:
        model_date = ""

    model = f"{ticker}_{model_period}" if ticker else model_period
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def generate_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for fallback in (
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.api.Close(False),
    ):
        try:
            fallback()
            return
        except Exception:
            continue


def list_source_files(input_path: Path) -> Iterable[Path]:
    if not input_path.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")
    if not input_path.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {input_path}")

    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            print(f"Skipped: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def get_sheet(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sheet in workbook.sheets:
        if sheet.name == sheet_name:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Tuple[int, int]:
    used = sheet.used_range
    values = used.value
    if values is None:
        raise ValueError(f"Sheet '{sheet.name}' is empty.")

    if not isinstance(values, list):
        matrix = [[values]]
    elif values and not isinstance(values[0], list):
        matrix = [values]
    else:
        matrix = values

    target = normalize_text(anchor_text)
    for row_idx, row in enumerate(matrix):
        for col_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return used.row + row_idx, used.column + col_idx

    raise ValueError(f"Could not find anchor '{anchor_text}' on sheet '{sheet.name}'.")


def get_header_map(
    sheet: xw.Sheet, header_row: int, start_col: int, end_col: int
) -> Dict[str, int]:
    if end_col < start_col:
        return {}
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if values is None:
        return {}
    if not isinstance(values, list):
        values = [values]

    header_map: Dict[str, int] = {}
    for idx, value in enumerate(values):
        normalized = normalize_text(value)
        if normalized and normalized not in header_map:
            header_map[normalized] = start_col + idx
    return header_map


def resolve_column(
    header_map: Dict[str, int],
    aliases: Sequence[str],
    fallback: int,
) -> int:
    normalized_aliases = [normalize_text(alias) for alias in aliases]
    for key, col in header_map.items():
        for alias in normalized_aliases:
            if alias in key:
                return col
    return max(1, fallback)


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    return sheet.range((row, max(1, col))).value


def build_empirical_rows(
    workbook: xw.Book, sheet: xw.Sheet, labels: FileLabel, source_file: str
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_anchor_cell(sheet, "max")
    header_map = get_header_map(
        sheet,
        header_row=anchor_row,
        start_col=max(1, anchor_col - 25),
        end_col=anchor_col + 25,
    )

    num_quarters_col = resolve_column(
        header_map,
        aliases=["num_quarters_used", "numquartersused", "quartersused", "nquarters"],
        fallback=anchor_col - 9,
    )
    last_quarter_col = resolve_column(
        header_map,
        aliases=["last_quarter_used", "lastquarter", "lastqtr"],
        fallback=anchor_col - 8,
    )
    avg_penetration_col = resolve_column(
        header_map,
        aliases=["avg_penetration_pct", "avgpenetration", "penetrationpct"],
        fallback=anchor_col - 3,
    )
    forecast_value_col = resolve_column(
        header_map,
        aliases=["estimatedtotalsold", "forecastvalue", "estimatedsold", "totalsold"],
        fallback=anchor_col - 1,
    )
    actual_value_col = resolve_column(
        header_map,
        aliases=["actualvalue", "reportedsales", "actualsales", "actual"],
        fallback=anchor_col - 2,
    )
    forecast_min_col = resolve_column(header_map, aliases=["min"], fallback=anchor_col + 1)
    quarterly_sales_col = resolve_column(
        header_map,
        aliases=["quarterlysales", "qtrsales"],
        fallback=anchor_col - 7,
    )
    reported_sales_col = resolve_column(
        header_map,
        aliases=["reportedsales", "reported"],
        fallback=actual_value_col,
    )
    growth_rate_col = resolve_column(
        header_map,
        aliases=["growthratepct", "growthrate", "growthpct"],
        fallback=anchor_col - 5,
    )
    captured_in_db_col = resolve_column(
        header_map,
        aliases=["salescapturedindbpct", "capturedindbpct", "captured"],
        fallback=anchor_col - 4,
    )

    # Temporary helper formulas for average penetration percentages using R1C1 + formula2.
    helper_col = max(
        sheet.used_range.last_cell.column,
        anchor_col,
        forecast_min_col,
        captured_in_db_col,
    ) + 2
    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1
    helper_cell_range = sheet.range((first_data_row, helper_col), (last_data_row, helper_col))

    formulas: List[List[str]] = []
    for i in range(N_QUARTERS):
        row = first_data_row + i
        n_quarters = i + 1
        source_start_row = max(first_data_row, row - n_quarters + 1)
        formula = (
            f'=IFERROR(AVERAGE(R{source_start_row}C{avg_penetration_col}:'
            f'R{row}C{avg_penetration_col}),"")'
        )
        formulas.append([formula])
    helper_cell_range.formula2 = formulas
    workbook.app.calculate()
    avg_values = helper_cell_range.value
    if not isinstance(avg_values, list):
        avg_values = [avg_values]
    if avg_values and isinstance(avg_values[0], list):
        avg_values = [row[0] for row in avg_values]

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row = first_data_row + i
        num_quarters = read_cell(sheet, row, num_quarters_col)
        if num_quarters in (None, ""):
            num_quarters = i + 1

        avg_penetration_pct = avg_values[i] if i < len(avg_values) else read_cell(
            sheet, row, avg_penetration_col
        )
        forecast_value = read_cell(sheet, row, forecast_value_col)
        actual_value = read_cell(sheet, row, actual_value_col)
        forecast_max = read_cell(sheet, row, anchor_col)
        forecast_min = read_cell(sheet, row, forecast_min_col)

        row_data = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": maybe_blank(avg_penetration_pct),
            "num_quarters_used": maybe_blank(num_quarters),
            "last_quarter_used": maybe_blank(read_cell(sheet, row, last_quarter_col)),
            "forecast_value": maybe_blank(forecast_value),
            "actual_value": maybe_blank(actual_value),
            "forecast_max": maybe_blank(forecast_max),
            "forecast_min": maybe_blank(forecast_min),
            "range_width": maybe_blank(safe_subtract(forecast_max, forecast_min)),
            "avg_penetration_pct": maybe_blank(avg_penetration_pct),
            "quarterly_sales": maybe_blank(read_cell(sheet, row, quarterly_sales_col)),
            "reported_sales": maybe_blank(read_cell(sheet, row, reported_sales_col)),
            "growth_rate_pct": maybe_blank(read_cell(sheet, row, growth_rate_col)),
            "sales_captured_in_db_pct": maybe_blank(read_cell(sheet, row, captured_in_db_col)),
            "source_file": source_file,
        }
        rows.append(row_data)

    return rows


def build_regression_rows(
    workbook: xw.Book, sheet: xw.Sheet, labels: FileLabel, source_file: str
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_anchor_cell(sheet, "max")
    header_map = get_header_map(
        sheet,
        header_row=anchor_row,
        start_col=max(1, anchor_col - 25),
        end_col=anchor_col + 25,
    )

    # Required anchor-relative setup from existing logic.
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    num_quarters_col = resolve_column(
        header_map,
        aliases=["num_quarters_used", "numquartersused", "quartersused", "nquarters"],
        fallback=anchor_col - 9,
    )
    forecast_total_col = resolve_column(
        header_map,
        aliases=["totfcstwosa", "forecasttotalwithoutsa", "totforecastwithoutsa"],
        fallback=anchor_col - 1,
    )
    actual_col = resolve_column(
        header_map,
        aliases=["actualvalue", "reportedsales", "actualsales", "actual"],
        fallback=anchor_col - 2,
    )
    forecast_min_col = resolve_column(header_map, aliases=["min"], fallback=anchor_col + 1)

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1
    helper_start_col = max(sheet.used_range.last_cell.column, forecast_min_col) + 2
    intercept_col = helper_start_col
    slope_col = helper_start_col + 1

    intercept_formulas: List[List[str]] = []
    slope_formulas: List[List[str]] = []
    for i in range(N_QUARTERS):
        row = first_data_row + i
        n_quarters = i + 1
        source_start_row = max(first_data_row, row - n_quarters + 1)
        intercept_formulas.append(
            [
                f'=IFERROR(INTERCEPT(R{source_start_row}C{y_col}:R{row}C{y_col},'
                f'R{source_start_row}C{x_col}:R{row}C{x_col}),"")'
            ]
        )
        slope_formulas.append(
            [
                f'=IFERROR(SLOPE(R{source_start_row}C{y_col}:R{row}C{y_col},'
                f'R{source_start_row}C{x_col}:R{row}C{x_col}),"")'
            ]
        )

    intercept_range = sheet.range((first_data_row, intercept_col), (last_data_row, intercept_col))
    slope_range = sheet.range((first_data_row, slope_col), (last_data_row, slope_col))
    intercept_range.formula2 = intercept_formulas
    slope_range.formula2 = slope_formulas
    workbook.app.calculate()

    intercept_values = intercept_range.value
    slope_values = slope_range.value
    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]
    if intercept_values and isinstance(intercept_values[0], list):
        intercept_values = [row[0] for row in intercept_values]
    if slope_values and isinstance(slope_values[0], list):
        slope_values = [row[0] for row in slope_values]

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row = first_data_row + i
        num_quarters = read_cell(sheet, row, num_quarters_col)
        if num_quarters in (None, ""):
            num_quarters = i + 1

        forecast_value = read_cell(sheet, row, forecast_total_col)
        actual_value = read_cell(sheet, row, actual_col)
        forecast_max = read_cell(sheet, row, anchor_col)
        forecast_min = read_cell(sheet, row, forecast_min_col)
        intercept = intercept_values[i] if i < len(intercept_values) else ""
        slope = slope_values[i] if i < len(slope_values) else ""

        row_data = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": maybe_blank(num_quarters),
            "num_quarters_used": maybe_blank(num_quarters),
            "forecast_value": maybe_blank(forecast_value),
            "actual_value": maybe_blank(actual_value),
            "forecast_max": maybe_blank(forecast_max),
            "forecast_min": maybe_blank(forecast_min),
            "range_width": maybe_blank(safe_subtract(forecast_max, forecast_min)),
            "intercept": maybe_blank(intercept),
            "slope": maybe_blank(slope),
            "source_file": source_file,
        }

        if i == N_QUARTERS - 1 and rows:
            prev = rows[-1]
            duplicate_last_row = (
                approx_equal(prev.get("forecast_value"), row_data.get("forecast_value"))
                and approx_equal(prev.get("forecast_max"), row_data.get("forecast_max"))
                and approx_equal(prev.get("forecast_min"), row_data.get("forecast_min"))
                and approx_equal(prev.get("intercept"), row_data.get("intercept"))
                and approx_equal(prev.get("slope"), row_data.get("slope"))
            )
            if duplicate_last_row:
                continue

        rows.append(row_data)

    return rows


def compute_column_widths(columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> List[float]:
    widths: List[float] = []
    sample_size = min(len(rows), 1000)
    for column in columns:
        max_len = len(column)
        for i in range(sample_size):
            value = rows[i].get(column, "")
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        widths.append(float(min(40, max_len + 2)))
    return widths


def write_sheet(
    sheet: xw.Sheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]], app: xw.App
) -> None:
    grid = [list(columns)]
    for row in rows:
        grid.append([row.get(col, "") for col in columns])
    sheet.range("A1").value = grid

    row_count = len(grid)
    col_count = len(columns)

    header_range = sheet.range((1, 1), (1, col_count))
    header_range.api.Font.Bold = True

    data_range = sheet.range((1, 1), (max(1, row_count), col_count))
    data_range.api.AutoFilter()

    widths = compute_column_widths(columns, rows)
    for i, width in enumerate(widths, start=1):
        sheet.range((1, i), (1, i)).column_width = width

    # Freeze top row (best effort in hidden mode).
    try:
        sheet.activate()
        app.api.ActiveWindow.SplitRow = 1
        app.api.ActiveWindow.SplitColumn = 0
        app.api.ActiveWindow.FreezePanes = True
    except Exception:
        pass


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    output_book = app.books.add()
    try:
        empirical_sheet = output_book.sheets[0]
        empirical_sheet.name = EMPIRICAL_OUTPUT_SHEET

        if len(output_book.sheets) > 1:
            regression_sheet = output_book.sheets[1]
        else:
            regression_sheet = output_book.sheets.add(after=empirical_sheet)
        regression_sheet.name = REGRESSION_OUTPUT_SHEET

        while len(output_book.sheets) > 2:
            output_book.sheets[-1].delete()

        write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows, app)
        write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows, app)
        output_book.save(str(output_path))
    finally:
        output_book.close()


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_file = generate_output_path(input_path, output_path)

    print(f"Input directory: {input_path}")
    print(f"Output directory: {output_path}")

    files = list(list_source_files(input_path))
    if not files:
        print("No eligible .xlsx files found.")
        return

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in files:
            print(f"Processed file: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_label(file_path.name)

                empirical_sheet = get_sheet(wb, EMPIRICAL_SHEET_NAME)
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction: {file_path.name} (sheet missing)")
                else:
                    empirical_rows.extend(
                        build_empirical_rows(wb, empirical_sheet, labels, file_path.name)
                    )

                regression_sheet = get_sheet(wb, REGRESSION_SHEET_NAME)
                if regression_sheet is None:
                    print(f"Skipped regression extraction: {file_path.name} (sheet missing)")
                else:
                    regression_rows.extend(
                        build_regression_rows(wb, regression_sheet, labels, file_path.name)
                    )

                processed_count += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)

        write_output_workbook(app, output_file, empirical_rows, regression_rows)
    finally:
        app.quit()

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
