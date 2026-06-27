from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Inputs
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")

# Core extraction settings
N_QUARTERS = 10

# Anchor-based empirical offsets (relative to the located "max" cell)
EMPIRICAL_OFFSETS = {
    "quarter_col_delta": -12,
    "quarterly_sales_col_delta": -11,
    "reported_sales_col_delta": -10,
    "growth_rate_col_delta": -9,
    "sales_captured_col_delta": -8,
    "penetration_col_delta": -7,
    "forecast_max_value_col_delta": 1,
    "forecast_min_row_delta": 1,
    "forecast_min_value_col_delta": 1,
    "temp_formula_row_delta": 6,
    "temp_formula_col_delta": 4,
}

# Anchor-based regression offsets (relative to the located "max" cell)
REGRESSION_OFFSETS = {
    "forecast_max_value_col_delta": 1,
    "forecast_min_row_delta": 1,
    "forecast_min_value_col_delta": 1,
    "forecast_total_row_delta": -2,
    "forecast_total_col_delta": 1,
    "actual_value_row_delta": -3,
    "actual_value_col_delta": 1,
    "temp_intercept_row_delta": 6,
    "temp_intercept_col_delta": 4,
    "temp_slope_row_delta": 7,
    "temp_slope_col_delta": 4,
}

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

PERIOD_RE = re.compile(
    r"(?P<timing>Early|Mid|Late)(?P<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(?P<year>\d{4})",
    flags=re.IGNORECASE,
)


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        if text.endswith("%"):
            return float(text[:-1]) / 100.0
        return float(text)
    except ValueError:
        return None


def safe_subtract(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left - right


def safe_divide(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def coalesce(*values: Any) -> Any:
    for value in values:
        if value not in (None, ""):
            return value
    return None


def parse_file_metadata(file_path: Path) -> FileMetadata | None:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].upper()
    period_token = parts[2].split("_")[0]
    match = PERIOD_RE.search(period_token)
    if not match:
        return None

    timing = match.group("timing").title()
    month_abbr = match.group("month").title()
    year = int(match.group("year"))
    month = list(calendar.month_abbr).index(month_abbr)

    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_lookup[timing]

    model_period = f"{timing}{month_abbr}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"

    candidate = output_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def find_anchor(sheet: xw.Sheet, text: str = "max") -> tuple[int, int] | None:
    # xlValues (-4163), xlWhole (1), xlByRows (1), xlNext (1)
    anchor = sheet.api.Cells.Find(
        What=text,
        LookIn=-4163,
        LookAt=1,
        SearchOrder=1,
        SearchDirection=1,
        MatchCase=False,
    )
    if anchor is None:
        # Fallback to partial match if the label includes extra text.
        anchor = sheet.api.Cells.Find(
            What=text,
            LookIn=-4163,
            LookAt=2,
            SearchOrder=1,
            SearchDirection=1,
            MatchCase=False,
        )
    if anchor is None:
        return None
    return int(anchor.Row), int(anchor.Column)


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def set_formula2(cell: xw.Range, formula: str) -> None:
    cell.formula2 = formula


def process_empirical_sheet(
    sheet: xw.Sheet,
    metadata: FileMetadata,
    source_file: str,
    app: xw.App,
) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    quarter_col = anchor_col + EMPIRICAL_OFFSETS["quarter_col_delta"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col_delta"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales_col_delta"]
    growth_rate_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_col_delta"]
    sales_captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_col_delta"]
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["penetration_col_delta"]

    forecast_max = to_float(
        get_cell_value(
            sheet,
            anchor_row,
            anchor_col + EMPIRICAL_OFFSETS["forecast_max_value_col_delta"],
        )
    )
    forecast_min = to_float(
        get_cell_value(
            sheet,
            anchor_row + EMPIRICAL_OFFSETS["forecast_min_row_delta"],
            anchor_col + EMPIRICAL_OFFSETS["forecast_min_value_col_delta"],
        )
    )

    end_row = anchor_row - 1
    temp_formula_cell = sheet.cells(
        anchor_row + EMPIRICAL_OFFSETS["temp_formula_row_delta"],
        anchor_col + EMPIRICAL_OFFSETS["temp_formula_col_delta"],
    )

    rows: list[dict[str, Any]] = []
    for num_quarters in range(1, N_QUARTERS + 1):
        start_row = end_row - num_quarters + 1
        if start_row < 1:
            continue

        avg_formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        )
        set_formula2(temp_formula_cell, avg_formula)
        app.calculate()
        avg_penetration_pct = to_float(temp_formula_cell.value)

        last_quarter_used = get_cell_value(sheet, end_row, quarter_col)
        quarterly_sales = to_float(get_cell_value(sheet, end_row, quarterly_sales_col))
        reported_sales = to_float(get_cell_value(sheet, end_row, reported_sales_col))
        growth_rate_pct = to_float(get_cell_value(sheet, end_row, growth_rate_col))
        sales_captured_pct = to_float(get_cell_value(sheet, end_row, sales_captured_col))

        forecast_value = safe_divide(quarterly_sales, avg_penetration_pct)
        actual_value = reported_sales
        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }
        rows.append(row)

    temp_formula_cell.clear_contents()
    return rows


def process_regression_sheet(
    sheet: xw.Sheet,
    metadata: FileMetadata,
    source_file: str,
    app: xw.App,
) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    end_row = anchor_row - 1

    forecast_max = to_float(
        get_cell_value(
            sheet,
            anchor_row,
            anchor_col + REGRESSION_OFFSETS["forecast_max_value_col_delta"],
        )
    )
    forecast_min = to_float(
        get_cell_value(
            sheet,
            anchor_row + REGRESSION_OFFSETS["forecast_min_row_delta"],
            anchor_col + REGRESSION_OFFSETS["forecast_min_value_col_delta"],
        )
    )

    forecast_total_without_sa = to_float(
        get_cell_value(
            sheet,
            anchor_row + REGRESSION_OFFSETS["forecast_total_row_delta"],
            anchor_col + REGRESSION_OFFSETS["forecast_total_col_delta"],
        )
    )
    actual_value = to_float(
        get_cell_value(
            sheet,
            anchor_row + REGRESSION_OFFSETS["actual_value_row_delta"],
            anchor_col + REGRESSION_OFFSETS["actual_value_col_delta"],
        )
    )

    intercept_cell = sheet.cells(
        anchor_row + REGRESSION_OFFSETS["temp_intercept_row_delta"],
        anchor_col + REGRESSION_OFFSETS["temp_intercept_col_delta"],
    )
    slope_cell = sheet.cells(
        anchor_row + REGRESSION_OFFSETS["temp_slope_row_delta"],
        anchor_col + REGRESSION_OFFSETS["temp_slope_col_delta"],
    )

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for num_quarters in range(2, N_QUARTERS + 1):
        start_row = end_row - num_quarters + 1
        if start_row < 1:
            continue

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        latest_x = to_float(get_cell_value(sheet, end_row, x_col))
        calculated_forecast = (
            None
            if intercept is None or slope is None or latest_x is None
            else intercept + (slope * latest_x)
        )
        forecast_value = coalesce(forecast_total_without_sa, calculated_forecast)

        signature = (
            round(intercept, 8) if intercept is not None else None,
            round(slope, 8) if slope is not None else None,
            round(forecast_value, 8) if forecast_value is not None else None,
            round(forecast_max, 8) if forecast_max is not None else None,
            round(forecast_min, 8) if forecast_min is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    intercept_cell.clear_contents()
    slope_cell.clear_contents()
    return rows


def ensure_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def close_source_workbook(wb: xw.Book) -> None:
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


def autofit_columns(ws, headers: Iterable[str], rows: list[dict[str, Any]]) -> None:
    for idx, header in enumerate(headers, start=1):
        width = max(len(header), 12)
        for row in rows:
            value = row.get(header)
            if value is None:
                continue
            width = max(width, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(width + 2, 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    empirical_ws.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_ws.append([row.get(header) for header in EMPIRICAL_HEADERS])

    regression_ws.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_ws.append([row.get(header) for header in REGRESSION_HEADERS])

    for ws, headers, rows in (
        (empirical_ws, EMPIRICAL_HEADERS, empirical_rows),
        (regression_ws, REGRESSION_HEADERS, regression_rows),
    ):
        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=1, column=col_idx).font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autofit_columns(ws, headers, rows)

    wb.save(output_path)


def iter_source_files(input_folder: Path) -> Iterable[Path]:
    for file_path in sorted(input_folder.iterdir()):
        if not file_path.is_file():
            continue
        yield file_path


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            # Calculation mode can vary across Excel backends.
            pass

        for file_path in iter_source_files(input_dir):
            filename = file_path.name
            if filename.startswith("~"):
                print(f"Skipped file: {filename} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {filename} (not .xlsx)")
                continue

            metadata = parse_file_metadata(file_path)
            if metadata is None:
                print(f"Skipped file: {filename} (filename does not match expected pattern)")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_sheet = ensure_sheet(wb, "Empirical Model")
                regression_sheet = ensure_sheet(wb, "Regression Model")

                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {filename}: missing 'Empirical Model'")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            empirical_sheet,
                            metadata,
                            filename,
                            app,
                        )
                    )

                if regression_sheet is None:
                    print(f"Skipped regression extraction for {filename}: missing 'Regression Model'")
                else:
                    regression_rows.extend(
                        process_regression_sheet(
                            regression_sheet,
                            metadata,
                            filename,
                            app,
                        )
                    )

                processed_count += 1
                print(f"Processed file: {filename}")
            except Exception as exc:
                print(f"Skipped file: {filename} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
