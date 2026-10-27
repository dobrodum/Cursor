from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING
from typing import Any

try:
    import xlwings as xw
except ImportError:  # pragma: no cover - runtime dependency check
    xw = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover - runtime dependency check
    Workbook = None
    Font = None
    get_column_letter = None

if TYPE_CHECKING:
    import xlwings as xw_types


# Configure paths here.
input_dir = "/workspace/input"
output_dir = "/workspace/output"


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


@dataclass
class SheetSnapshot:
    top_row: int
    left_col: int
    values: list[list[Any]]

    @property
    def nrows(self) -> int:
        return len(self.values)

    @property
    def ncols(self) -> int:
        if not self.values:
            return 0
        return max(len(row) for row in self.values)

    def get(self, row: int, col: int) -> Any:
        row_index = row - self.top_row
        col_index = col - self.left_col
        if row_index < 0 or col_index < 0:
            return None
        if row_index >= self.nrows:
            return None
        row_values = self.values[row_index]
        if col_index >= len(row_values):
            return None
        return row_values[col_index]

    def find_text(self, expected: str) -> tuple[int, int] | None:
        expected_clean = expected.strip().lower()
        for row_offset, row in enumerate(self.values):
            for col_offset, value in enumerate(row):
                if isinstance(value, str) and value.strip().lower() == expected_clean:
                    return self.top_row + row_offset, self.left_col + col_offset
        return None


def normalize_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return [[None]]
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def snapshot_sheet(sheet: "xw_types.Sheet") -> SheetSnapshot:
    used_range = sheet.used_range
    matrix = normalize_matrix(used_range.value)
    return SheetSnapshot(top_row=used_range.row, left_col=used_range.column, values=matrix)


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        is_percent = text.endswith("%")
        if is_percent:
            text = text[:-1].strip()
        text = text.replace(",", "")
        try:
            numeric = float(text)
        except ValueError:
            return None
        if is_percent:
            numeric /= 100.0
        return numeric
    return None


def normalize_output_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def set_formula2(cell: "xw_types.Range", formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_workbook(wb: "xw_types.Book") -> None:
    close_errors: list[Exception] = []
    for close_attempt in (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
    ):
        try:
            close_attempt()
            return
        except Exception as exc:
            close_errors.append(exc)
    raise RuntimeError(f"Unable to safely close workbook: {close_errors[-1]}")


def month_number(month_text: str) -> int | None:
    month_lookup = {
        "jan": 1,
        "january": 1,
        "feb": 2,
        "february": 2,
        "mar": 3,
        "march": 3,
        "apr": 4,
        "april": 4,
        "may": 5,
        "jun": 6,
        "june": 6,
        "jul": 7,
        "july": 7,
        "aug": 8,
        "august": 8,
        "sep": 9,
        "sept": 9,
        "september": 9,
        "oct": 10,
        "october": 10,
        "nov": 11,
        "november": 11,
        "dec": 12,
        "december": 12,
    }
    key = month_text.strip().lower()
    if key in month_lookup:
        return month_lookup[key]
    if len(key) >= 3 and key[:3] in month_lookup:
        return month_lookup[key[:3]]
    return None


def parse_model_metadata(file_name: str) -> dict[str, Any]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"\s+", "", parts[1].upper())

    period_source = parts[2] if len(parts) >= 3 else stem
    period_core = period_source.split("_")[0]

    period_match = re.search(
        r"(?i)(early|mid|late)\s*([A-Za-z]+)\s*(\d{4})",
        period_core,
    )

    model_period = period_core
    model_date = ""

    if period_match:
        phase = period_match.group(1).title()
        month_text = period_match.group(2)
        year = int(period_match.group(3))
        month = month_number(month_text)
        if month is not None:
            month_abbrev = date(year, month, 1).strftime("%b")
            model_period = f"{phase}{month_abbrev}_{year}"
            day_by_phase = {"Early": 5, "Mid": 15, "Late": 25}
            model_date = date(year, month, day_by_phase[phase]).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def get_sheet_case_insensitive(wb: "xw_types.Book", name: str) -> "xw_types.Sheet | None":
    expected = name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == expected:
            return sheet
    return None


def extract_empirical_rows(
    sheet: "xw_types.Sheet",
    wb: "xw_types.Book",
    metadata: dict[str, Any],
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = snapshot.find_text("max")
    if anchor is None:
        print(f"skipped empirical in {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    data_end_row = anchor_row - 1
    if data_end_row < 1:
        print(f"skipped empirical in {source_file}: not enough rows above anchor")
        return []

    # Anchor-driven offsets for this model layout.
    quarter_col = anchor_col - 12
    penetration_col = anchor_col - 8
    quarterly_sales_col = anchor_col - 7
    reported_sales_col = anchor_col - 6
    growth_col = anchor_col - 5
    sales_captured_col = anchor_col - 4

    temp_avg_col = anchor_col + 2
    temp_growth_col = anchor_col + 3
    rows: list[dict[str, Any]] = []

    try:
        for quarter_count in range(1, n_quarters + 1):
            start_row = data_end_row - quarter_count + 1
            if start_row < 1:
                break

            avg_cell = sheet.cells(anchor_row + quarter_count, temp_avg_col)
            growth_cell = sheet.cells(anchor_row + quarter_count, temp_growth_col)

            set_formula2(
                avg_cell,
                f"=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{data_end_row}C{penetration_col}),\"\")",
            )
            if quarter_count > 1:
                set_formula2(
                    growth_cell,
                    (
                        "=IFERROR((R{end_row}C{sales_col}/R{start_row}C{sales_col})^(1/{step})-1,\"\")".format(
                            end_row=data_end_row,
                            sales_col=quarterly_sales_col,
                            start_row=start_row,
                            step=quarter_count - 1,
                        )
                    ),
                )
            else:
                growth_cell.value = None

            wb.app.calculate()

            avg_penetration = to_float(avg_cell.value)
            growth_rate = to_float(growth_cell.value)
            quarterly_sales = to_float(snapshot.get(data_end_row, quarterly_sales_col))
            reported_sales = to_float(snapshot.get(data_end_row, reported_sales_col))

            if reported_sales is None:
                reported_sales = quarterly_sales

            forecast_value = None
            if avg_penetration not in (None, 0) and quarterly_sales is not None:
                forecast_value = quarterly_sales / avg_penetration

            forecast_max = to_float(sheet.cells(anchor_row + quarter_count, anchor_col).value)
            forecast_min = to_float(sheet.cells(anchor_row + quarter_count, anchor_col + 1).value)
            if forecast_max is None and forecast_value is not None:
                forecast_max = forecast_value * 1.05
            if forecast_min is None and forecast_value is not None:
                forecast_min = forecast_value * 0.95

            if (
                avg_penetration is None
                and forecast_value is None
                and forecast_max is None
                and forecast_min is None
            ):
                continue

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            sales_captured_pct = to_float(snapshot.get(data_end_row, sales_captured_col))
            if (
                sales_captured_pct is None
                and quarterly_sales is not None
                and reported_sales not in (None, 0)
            ):
                sales_captured_pct = quarterly_sales / reported_sales

            if growth_rate is None:
                growth_rate = to_float(snapshot.get(data_end_row, growth_col))

            row = {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": quarter_count,
                "last_quarter_used": snapshot.get(data_end_row, quarter_col),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
            rows.append(row)
    finally:
        try:
            sheet.range(
                (anchor_row + 1, temp_avg_col),
                (anchor_row + n_quarters + 1, temp_growth_col),
            ).clear_contents()
        except Exception:
            pass

    return rows


def rounded_signature(*values: Any) -> tuple[Any, ...]:
    normalized: list[Any] = []
    for value in values:
        number = to_float(value)
        if number is None:
            normalized.append(None)
        else:
            normalized.append(round(number, 10))
    return tuple(normalized)


def extract_regression_rows(
    sheet: "xw_types.Sheet",
    wb: "xw_types.Book",
    metadata: dict[str, Any],
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = snapshot.find_text("max")
    if anchor is None:
        print(f"skipped regression in {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    data_end_row = anchor_row - 1
    if data_end_row < 1:
        print(f"skipped regression in {source_file}: not enough rows above anchor")
        return []

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    temp_intercept_col = anchor_col + 2
    temp_slope_col = anchor_col + 3
    temp_forecast_col = anchor_col + 4

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    try:
        for quarter_count in range(1, 11):
            start_row = data_end_row - quarter_count + 1
            if start_row < 1:
                break

            intercept_cell = sheet.cells(anchor_row + quarter_count, temp_intercept_col)
            slope_cell = sheet.cells(anchor_row + quarter_count, temp_slope_col)
            forecast_cell = sheet.cells(anchor_row + quarter_count, temp_forecast_col)

            set_formula2(
                intercept_cell,
                (
                    "=IFERROR(INTERCEPT(R{start}C{y}:R{end}C{y},R{start}C{x}:R{end}C{x}),\"\")".format(
                        start=start_row,
                        end=data_end_row,
                        y=y_col,
                        x=x_col,
                    )
                ),
            )
            set_formula2(
                slope_cell,
                (
                    "=IFERROR(SLOPE(R{start}C{y}:R{end}C{y},R{start}C{x}:R{end}C{x}),\"\")".format(
                        start=start_row,
                        end=data_end_row,
                        y=y_col,
                        x=x_col,
                    )
                ),
            )
            set_formula2(
                forecast_cell,
                (
                    "=IFERROR(R{row}C{int_col}+R{row}C{slope_col}*(R{end}C{x_col}+1),\"\")".format(
                        row=anchor_row + quarter_count,
                        int_col=temp_intercept_col,
                        slope_col=temp_slope_col,
                        end=data_end_row,
                        x_col=x_col,
                    )
                ),
            )

            wb.app.calculate()

            intercept = to_float(intercept_cell.value)
            slope = to_float(slope_cell.value)
            forecast_total_without_sa = to_float(forecast_cell.value)
            forecast_max = to_float(sheet.cells(anchor_row + quarter_count, anchor_col).value)
            forecast_min = to_float(sheet.cells(anchor_row + quarter_count, anchor_col + 1).value)

            if forecast_max is None and forecast_total_without_sa is not None:
                forecast_max = forecast_total_without_sa * 1.05
            if forecast_min is None and forecast_total_without_sa is not None:
                forecast_min = forecast_total_without_sa * 0.95

            signature = rounded_signature(
                intercept,
                slope,
                forecast_total_without_sa,
                forecast_max,
                forecast_min,
            )

            # Prevent duplicate final rows when the model stops changing.
            if signature == previous_signature:
                continue
            previous_signature = signature

            if (
                intercept is None
                and slope is None
                and forecast_total_without_sa is None
                and forecast_max is None
                and forecast_min is None
            ):
                continue

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            row = {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": quarter_count,
                "num_quarters_used": quarter_count,
                "forecast_value": forecast_total_without_sa,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
            rows.append(row)
    finally:
        try:
            sheet.range(
                (anchor_row + 1, temp_intercept_col),
                (anchor_row + 12, temp_forecast_col),
            ).clear_contents()
        except Exception:
            pass

    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)

    for row in rows:
        ws.append([normalize_output_value(row.get(col)) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    max_row_for_width = min(ws.max_row, 2000)
    for col_index, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_index in range(1, max_row_for_width + 1):
            cell_value = ws.cell(row=row_index, column=col_index).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max(12, max_len + 2), 44)


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    write_sheet(output_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    output_wb.save(output_file)


def main() -> None:
    if xw is None or Workbook is None or Font is None or get_column_letter is None:
        raise ModuleNotFoundError(
            "xlwings and openpyxl are required. Install dependencies first: pip install xlwings openpyxl"
        )

    source_dir = Path(input_dir).expanduser()
    target_dir = Path(output_dir).expanduser()

    if not source_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {source_dir}")

    target_dir.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(source_dir, target_dir)

    source_files = sorted(source_dir.iterdir(), key=lambda path: path.name.lower())

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    previous_calculation = None
    try:
        previous_calculation = app.api.Calculation
        app.api.Calculation = -4135  # xlCalculationManual
    except Exception:
        previous_calculation = None

    try:
        for file_path in source_files:
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"processing file: {file_path.name}")
            metadata = parse_model_metadata(file_path.name)
            wb: xw.Book | None = None

            try:
                wb = app.books.open(str(file_path), update_links=False)

                empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"skipped empirical in {file_path.name}: sheet not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            empirical_sheet,
                            wb,
                            metadata,
                            file_path.name,
                        )
                    )

                regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"skipped regression in {file_path.name}: sheet not found")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            regression_sheet,
                            wb,
                            metadata,
                            file_path.name,
                        )
                    )

                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    try:
                        safe_close_workbook(wb)
                    except Exception as close_exc:
                        print(f"warning: close failed for {file_path.name}: {close_exc}")
    finally:
        try:
            if previous_calculation is not None:
                app.api.Calculation = previous_calculation
        except Exception:
            pass
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
