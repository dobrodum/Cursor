from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running the script.
input_dir = Path("./input")
output_dir = Path("./output")

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

MONTHS = {
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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}

REGEX_PERIOD = re.compile(
    r"\b(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?(\d{4})\b", re.IGNORECASE
)


@dataclass
class SheetSnapshot:
    sheet: xw.main.Sheet
    start_row: int
    start_col: int
    matrix: list[list[Any]]
    text_index: dict[str, list[tuple[int, int]]]

    @property
    def row_count(self) -> int:
        return len(self.matrix)

    @property
    def col_count(self) -> int:
        if not self.matrix:
            return 0
        return len(self.matrix[0])

    def value(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if 0 <= row_idx < self.row_count and 0 <= col_idx < self.col_count:
            return self.matrix[row_idx][col_idx]
        return self.sheet.cells(row, col).value


def normalize_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, tuple):
        values = [list(item) if isinstance(item, tuple) else item for item in values]
        first = values[0]
    if isinstance(first, list):
        matrix: list[list[Any]] = []
        for row in values:
            if isinstance(row, tuple):
                matrix.append(list(row))
            elif isinstance(row, list):
                matrix.append(row)
            else:
                matrix.append([row])
        return matrix
    return [list(values)]


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip().lower()
    return str(value).strip().lower()


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        is_pct = stripped.endswith("%")
        cleaned = stripped.replace(",", "").replace("%", "")
        try:
            number = float(cleaned)
        except ValueError:
            return None
        return number / 100.0 if is_pct else number
    return None


def normalize_label(label: str) -> str:
    return re.sub(r"\s+", " ", label.strip().lower())


def build_text_index(
    matrix: list[list[Any]], start_row: int, start_col: int
) -> dict[str, list[tuple[int, int]]]:
    index: dict[str, list[tuple[int, int]]] = {}
    for row_offset, row_values in enumerate(matrix):
        for col_offset, value in enumerate(row_values):
            if not isinstance(value, str):
                continue
            normalized = normalize_label(value)
            if not normalized:
                continue
            abs_row = start_row + row_offset
            abs_col = start_col + col_offset
            index.setdefault(normalized, []).append((abs_row, abs_col))
    return index


def take_sheet_snapshot(sheet: xw.main.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    matrix = normalize_matrix(used.value)
    text_index = build_text_index(matrix, used.row, used.column)
    return SheetSnapshot(
        sheet=sheet,
        start_row=used.row,
        start_col=used.column,
        matrix=matrix,
        text_index=text_index,
    )


def choose_max_anchor(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    max_cells = snapshot.text_index.get("max", [])
    min_cells = snapshot.text_index.get("min", [])
    if not max_cells:
        return None
    if not min_cells:
        return sorted(max_cells)[0]

    def distance_to_nearest_min(max_cell: tuple[int, int]) -> int:
        max_row, max_col = max_cell
        return min(
            abs(max_row - min_row) + abs(max_col - min_col)
            for min_row, min_col in min_cells
        )

    return min(max_cells, key=distance_to_nearest_min)


def find_label_cell_near_anchor(
    snapshot: SheetSnapshot,
    labels: list[str],
    anchor_row: int,
    anchor_col: int,
) -> tuple[int, int] | None:
    candidates: list[tuple[int, int]] = []
    for label in labels:
        normalized = normalize_label(label)
        candidates.extend(snapshot.text_index.get(normalized, []))
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda pos: abs(pos[0] - anchor_row) + abs(pos[1] - anchor_col),
    )


def first_numeric_to_right(
    snapshot: SheetSnapshot,
    row: int,
    col: int,
    max_steps: int = 6,
) -> float | None:
    for step in range(1, max_steps + 1):
        value = to_float(snapshot.value(row, col + step))
        if value is not None:
            return value
    return None


def lookup_metric(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    labels: list[str],
    fallback_offset: tuple[int, int] | None = None,
) -> float | None:
    label_cell = find_label_cell_near_anchor(snapshot, labels, anchor_row, anchor_col)
    if label_cell is not None:
        value = first_numeric_to_right(snapshot, label_cell[0], label_cell[1])
        if value is not None:
            return value
    if fallback_offset is not None:
        row = anchor_row + fallback_offset[0]
        col = anchor_col + fallback_offset[1]
        return to_float(snapshot.value(row, col))
    return None


def scalar_signature(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value, 8)


def parse_file_labels(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = parts[1].upper() if len(parts) >= 2 else "UNKNOWN"

    model_period = "Unknown"
    model_date = ""
    match = REGEX_PERIOD.search(stem)
    if match:
        phase_raw, month_raw, year = match.groups()
        phase = phase_raw.capitalize()
        month_abbrev = month_raw[:3].capitalize()
        month_num = MONTHS.get(month_abbrev.lower())
        day = PHASE_TO_DAY.get(phase.lower())
        if month_num and day:
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = f"{year}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if model_period != "Unknown" else ticker
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def choose_output_path(input_folder: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    default_path = target_dir / f"{base_name}.xlsx"
    if not default_path.exists():
        return default_path
    counter = 1
    while True:
        candidate = target_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def safe_close_workbook(workbook: xw.main.Book) -> None:
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
        pass


def get_sheet_if_exists(workbook: xw.main.Book, name: str) -> xw.main.Sheet | None:
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == name.lower():
            return sheet
    return None


def infer_last_quarter_label(
    snapshot: SheetSnapshot, row: int, x_col: int, y_col: int
) -> str:
    candidate_cols = [
        x_col - 2,
        x_col - 1,
        y_col - 2,
        y_col - 1,
        snapshot.start_col,
    ]
    for col in candidate_cols:
        value = snapshot.value(row, col)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def scale_from_base(
    forecast_value: float | None,
    base_forecast: float | None,
    base_bound: float | None,
) -> float | None:
    if base_bound is None:
        return None
    if forecast_value is None:
        return base_bound
    if base_forecast in (None, 0):
        return base_bound
    return base_bound * (forecast_value / base_forecast)


def process_empirical_sheet(
    workbook: xw.main.Book,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, "Empirical Model")
    if sheet is None:
        print(f"Skipping empirical extraction for {source_file}: sheet not found")
        return []

    snapshot = take_sheet_snapshot(sheet)
    anchor = choose_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipping empirical extraction for {source_file}: max anchor not found")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_end_row = anchor_row - 1
    max_available = data_end_row - snapshot.start_row + 1
    quarters_to_use = min(N_QUARTERS, max_available)
    if quarters_to_use <= 0:
        return []

    base_max = lookup_metric(snapshot, anchor_row, anchor_col, ["max"], (0, 1))
    base_min = lookup_metric(snapshot, anchor_row, anchor_col, ["min"], (1, 1))
    base_forecast = lookup_metric(
        snapshot,
        anchor_row,
        anchor_col,
        [
            "estimated total sold",
            "est total sold",
            "forecast",
            "forecast value",
        ],
        (-1, 1),
    )

    temp_col = max(anchor_col + 12, snapshot.start_col + snapshot.col_count + 2)
    temp_row = anchor_row
    formula_cell = sheet.cells(temp_row, temp_col)

    rows: list[dict[str, Any]] = []
    for num_quarters in range(1, quarters_to_use + 1):
        start_row = data_end_row - num_quarters + 1
        sum_x_formula = f"R{start_row}C{x_col}:R{data_end_row}C{x_col}"
        sum_y_formula = f"R{start_row}C{y_col}:R{data_end_row}C{y_col}"
        formula_cell.formula2 = f"=IFERROR(SUM({sum_x_formula})/SUM({sum_y_formula}),0)"
        workbook.app.calculate()
        avg_penetration_pct = to_float(formula_cell.value)

        latest_quarterly_sales = to_float(snapshot.value(data_end_row, x_col))
        latest_reported_sales = to_float(snapshot.value(data_end_row, y_col))
        prev_reported_sales = (
            to_float(snapshot.value(data_end_row - 1, y_col))
            if data_end_row - 1 >= start_row
            else None
        )

        growth_rate_pct = None
        if (
            latest_reported_sales is not None
            and prev_reported_sales not in (None, 0)
            and prev_reported_sales is not None
        ):
            growth_rate_pct = (latest_reported_sales - prev_reported_sales) / prev_reported_sales

        sales_captured_pct = None
        if (
            latest_quarterly_sales is not None
            and latest_reported_sales not in (None, 0)
            and latest_reported_sales is not None
        ):
            sales_captured_pct = latest_quarterly_sales / latest_reported_sales

        forecast_value = None
        if latest_quarterly_sales is not None and avg_penetration_pct not in (None, 0):
            forecast_value = latest_quarterly_sales / avg_penetration_pct
        elif base_forecast is not None:
            forecast_value = base_forecast

        forecast_max = scale_from_base(forecast_value, base_forecast, base_max)
        forecast_min = scale_from_base(forecast_value, base_forecast, base_min)
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters,
            "last_quarter_used": infer_last_quarter_label(snapshot, start_row, x_col, y_col),
            "forecast_value": forecast_value,
            "actual_value": latest_reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": latest_quarterly_sales,
            "reported_sales": latest_reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def process_regression_sheet(
    workbook: xw.main.Book,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, "Regression Model")
    if sheet is None:
        print(f"Skipping regression extraction for {source_file}: sheet not found")
        return []

    snapshot = take_sheet_snapshot(sheet)
    anchor = choose_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipping regression extraction for {source_file}: max anchor not found")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_end_row = anchor_row - 1
    max_available = data_end_row - snapshot.start_row + 1
    quarters_to_use = min(N_QUARTERS, max_available)
    if quarters_to_use < 2:
        return []

    base_max = lookup_metric(snapshot, anchor_row, anchor_col, ["max"], (0, 1))
    base_min = lookup_metric(snapshot, anchor_row, anchor_col, ["min"], (1, 1))
    base_forecast = lookup_metric(
        snapshot,
        anchor_row,
        anchor_col,
        [
            "tot fcst w/o sa",
            "tot fcst without sa",
            "forecast total without sa",
            "forecast",
        ],
        (-1, 1),
    )

    temp_col = max(anchor_col + 12, snapshot.start_col + snapshot.col_count + 2)
    temp_row = anchor_row
    intercept_cell = sheet.cells(temp_row, temp_col)
    slope_cell = sheet.cells(temp_row, temp_col + 1)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    latest_x = to_float(snapshot.value(data_end_row, x_col))
    latest_actual = to_float(snapshot.value(data_end_row, y_col))

    for num_quarters in range(2, quarters_to_use + 1):
        start_row = data_end_row - num_quarters + 1
        y_formula = f"R{start_row}C{y_col}:R{data_end_row}C{y_col}"
        x_formula = f"R{start_row}C{x_col}:R{data_end_row}C{x_col}"
        intercept_cell.formula2 = f"=IFERROR(INTERCEPT({y_formula},{x_formula}),\"\")"
        slope_cell.formula2 = f"=IFERROR(SLOPE({y_formula},{x_formula}),\"\")"
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        forecast_value = None
        if intercept is not None and slope is not None and latest_x is not None:
            forecast_value = intercept + slope * latest_x
        elif base_forecast is not None:
            forecast_value = base_forecast

        forecast_max = scale_from_base(forecast_value, base_forecast, base_max)
        forecast_min = scale_from_base(forecast_value, base_forecast, base_min)
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            scalar_signature(intercept),
            scalar_signature(slope),
            scalar_signature(forecast_value),
            scalar_signature(forecast_max),
            scalar_signature(forecast_min),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": latest_actual,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def autosize_columns(worksheet: Any) -> None:
    for idx, column_cells in enumerate(worksheet.columns, start=1):
        max_length = 0
        for cell in column_cells:
            if cell.value is None:
                continue
            max_length = max(max_length, len(str(cell.value)))
        adjusted = min(max(max_length + 2, 12), 48)
        worksheet.column_dimensions[get_column_letter(idx)].width = adjusted


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    sheet_specs = [
        ("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows),
        ("regression_candidates", REGRESSION_COLUMNS, regression_rows),
    ]

    for sheet_name, columns, rows in sheet_specs:
        worksheet = workbook.create_sheet(sheet_name)
        worksheet.append(columns)
        for row in rows:
            worksheet.append([row.get(column) for column in columns])
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        autosize_columns(worksheet)

    workbook.save(output_path)


def gather_source_files(folder: Path) -> list[Path]:
    files_to_process: list[Path] = []
    if not folder.exists():
        print(f"Input directory does not exist: {folder}")
        return files_to_process

    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipping {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipping {path.name}: not an .xlsx file")
            continue
        files_to_process.append(path)
    return files_to_process


def main() -> None:
    source_files = gather_source_files(input_dir)
    output_path = choose_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    if not source_files:
        write_output_workbook(output_path, empirical_rows, regression_rows)
        print(f"Output path: {output_path}")
        print("Number of files processed: 0")
        print("Number of empirical rows: 0")
        print("Number of regression rows: 0")
        return

    app: xw.main.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"Processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                parsed_labels = parse_file_labels(file_path)
                empirical_rows.extend(
                    process_empirical_sheet(workbook, parsed_labels, file_path.name)
                )
                regression_rows.extend(
                    process_regression_sheet(workbook, parsed_labels, file_path.name)
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
