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


# User-configurable paths.
input_dir = Path("./input")
output_dir = Path("./output")

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

EMPIRICAL_OUTPUT_COLUMNS = [
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

REGRESSION_OUTPUT_COLUMNS = [
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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
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
PERIOD_PATTERN = re.compile(
    r"\b(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})\b",
    re.IGNORECASE,
)


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetGrid:
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: list[list[Any]]
    labels: dict[str, list[tuple[int, int]]]


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    collapsed = re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
    return collapsed


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and value is not None


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def round_or_none(value: Any, precision: int = 10) -> float | None:
    parsed = to_float(value)
    if parsed is None:
        return None
    return round(parsed, precision)


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_match = PERIOD_PATTERN.search(stem)
    if not period_match:
        raise ValueError("filename does not include Early/Mid/Late + month + year period token")

    period_name, month_name, year_token = period_match.groups()
    period_name = period_name.title()
    month_name = month_name.title()
    year = int(year_token)

    month_number = MONTH_MAP[month_name.lower()]
    day = PERIOD_DAY_MAP[period_name.lower()]
    model_date = date(year, month_number, day).isoformat()
    model_period = f"{period_name}{month_name}_{year}"
    model = f"{ticker}_{model_period}"

    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_next_output_path(source_input_dir: Path, destination_dir: Path) -> Path:
    folder_name = source_input_dir.name
    candidate = destination_dir / f"{folder_name}_PARAM.xlsx"
    if not candidate.exists():
        return candidate

    version = 1
    while True:
        candidate = destination_dir / f"{folder_name}_PARAM.{version}.xlsx"
        if not candidate.exists():
            return candidate
        version += 1


def read_sheet_grid(sheet: xw.main.Sheet) -> SheetGrid:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    values = used.options(ndim=2).value
    rows = [list(row) for row in values] if values else [[None]]

    height = len(rows)
    width = max(len(row) for row in rows)
    normalized_rows: list[list[Any]] = []
    for row in rows:
        if len(row) < width:
            normalized_rows.append(row + [None] * (width - len(row)))
        else:
            normalized_rows.append(row)

    end_row = start_row + height - 1
    end_col = start_col + width - 1

    labels: dict[str, list[tuple[int, int]]] = {}
    for row_offset, row in enumerate(normalized_rows):
        abs_row = start_row + row_offset
        for col_offset, raw_value in enumerate(row):
            normalized = normalize_label(raw_value)
            if normalized:
                abs_col = start_col + col_offset
                labels.setdefault(normalized, []).append((abs_row, abs_col))

    return SheetGrid(
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        values=normalized_rows,
        labels=labels,
    )


def grid_value(grid: SheetGrid, row: int, col: int) -> Any:
    if row < grid.start_row or row > grid.end_row:
        return None
    if col < grid.start_col or col > grid.end_col:
        return None
    return grid.values[row - grid.start_row][col - grid.start_col]


def cell_value(sheet: xw.main.Sheet, grid: SheetGrid, row: int, col: int) -> Any:
    in_grid = grid_value(grid, row, col)
    if in_grid is not None:
        return in_grid
    return sheet.cells(row, col).value


def find_max_anchor(grid: SheetGrid) -> tuple[int, int]:
    candidates = grid.labels.get("max", [])
    if not candidates:
        for label, positions in grid.labels.items():
            if label.startswith("max "):
                candidates.extend(positions)
    if not candidates:
        raise ValueError("could not find 'max' anchor cell")

    def candidate_score(position: tuple[int, int]) -> int:
        row, col = position
        right_label = normalize_label(grid_value(grid, row, col + 1))
        return 1 if right_label == "min" else 0

    candidates.sort(key=candidate_score, reverse=True)
    return candidates[0]


def collect_header_map(grid: SheetGrid, rows: list[int]) -> dict[str, list[int]]:
    headers: dict[str, list[int]] = {}
    for row in rows:
        if row < grid.start_row or row > grid.end_row:
            continue
        for col in range(grid.start_col, grid.end_col + 1):
            raw = grid_value(grid, row, col)
            normalized = normalize_label(raw)
            if normalized:
                headers.setdefault(normalized, []).append(col)
    return headers


def find_col_by_header(
    header_map: dict[str, list[int]],
    phrases: list[str],
    default_col: int,
) -> int:
    for label, cols in header_map.items():
        for phrase in phrases:
            if phrase in label:
                return cols[0]
    return default_col


def get_sheet_by_name(workbook: xw.main.Book, name: str) -> xw.main.Sheet:
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == name.lower():
            return sheet
    raise KeyError(f"sheet '{name}' not found")


def collect_candidate_rows(
    grid: SheetGrid,
    anchor_row: int,
    max_col: int,
    min_col: int,
    limit: int = 10,
) -> list[int]:
    rows: list[int] = []
    blank_streak = 0
    scan_end_row = min(grid.end_row, anchor_row + 150)
    for row in range(anchor_row + 1, scan_end_row + 1):
        max_value = grid_value(grid, row, max_col)
        min_value = grid_value(grid, row, min_col)
        if is_number(max_value) or is_number(min_value):
            rows.append(row)
            blank_streak = 0
            if len(rows) >= limit:
                break
            continue

        if rows:
            blank_streak += 1
            if blank_streak >= 3:
                break
    return rows


def find_numeric_rows(grid: SheetGrid, column: int, stop_row_exclusive: int) -> list[int]:
    numeric_rows: list[int] = []
    for row in range(grid.start_row, stop_row_exclusive):
        value = grid_value(grid, row, column)
        if is_number(value):
            numeric_rows.append(row)
    return numeric_rows


def find_numeric_pair_rows(
    grid: SheetGrid,
    x_col: int,
    y_col: int,
    stop_row_exclusive: int,
) -> list[int]:
    rows: list[int] = []
    for row in range(grid.start_row, stop_row_exclusive):
        x_value = grid_value(grid, row, x_col)
        y_value = grid_value(grid, row, y_col)
        if is_number(x_value) and is_number(y_value):
            rows.append(row)
    return rows


def safe_close_source_workbook(workbook: xw.main.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass


def build_empirical_rows(
    workbook: xw.main.Book,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, EMPIRICAL_SHEET_NAME)
    grid = read_sheet_grid(sheet)
    anchor_row, anchor_col = find_max_anchor(grid)

    header_map = collect_header_map(grid, [anchor_row - 1, anchor_row])
    max_col = anchor_col
    min_col = find_col_by_header(header_map, ["min"], anchor_col + 1)
    forecast_col = find_col_by_header(
        header_map,
        ["estimated total sold", "forecast value", "tot fcst", "forecast"],
        anchor_col - 1,
    )
    reported_col = find_col_by_header(
        header_map,
        ["reported sales", "actual sales", "reported"],
        anchor_col - 4,
    )
    quarterly_sales_col = find_col_by_header(
        header_map,
        ["quarterly sales", "qtr sales", "quarter sales"],
        anchor_col - 5,
    )
    avg_pen_col = find_col_by_header(
        header_map,
        ["avg penetration", "average penetration", "avg pen"],
        anchor_col - 6,
    )
    penetration_input_col = anchor_col - 9
    for label, cols in header_map.items():
        if "penetration" in label and "avg" not in label and "average" not in label:
            penetration_input_col = cols[0]
            break
    if penetration_input_col == anchor_col - 9:
        penetration_input_col = find_col_by_header(
            header_map,
            ["penetration pct", "pen pct", "penetration"],
            anchor_col - 9,
        )
    num_quarters_col = find_col_by_header(
        header_map,
        ["num quarters", "quarters used", "n quarters"],
        anchor_col - 8,
    )
    last_quarter_col = find_col_by_header(
        header_map,
        ["last quarter", "last qtr", "quarter used"],
        anchor_col - 7,
    )
    growth_col = find_col_by_header(
        header_map,
        ["growth rate", "growth pct", "growth"],
        anchor_col - 3,
    )
    sales_captured_col = find_col_by_header(
        header_map,
        ["sales captured in db", "captured in db", "sales captured"],
        anchor_col - 2,
    )

    candidate_rows = collect_candidate_rows(
        grid=grid,
        anchor_row=anchor_row,
        max_col=max_col,
        min_col=min_col,
        limit=10,
    )
    numeric_pen_rows = find_numeric_rows(
        grid=grid,
        column=penetration_input_col,
        stop_row_exclusive=anchor_row,
    )

    scratch_row = grid.end_row + 3
    scratch_col = max(grid.end_col + 1, anchor_col + 3)
    scratch_cell = sheet.cells(scratch_row, scratch_col)

    rows: list[dict[str, Any]] = []
    max_iterations = 10
    max_quarters_with_data = len(numeric_pen_rows)
    for idx in range(max_iterations):
        n_quarters = idx + 1
        if max_quarters_with_data < n_quarters and idx >= len(candidate_rows):
            break

        if idx < len(candidate_rows):
            source_row = candidate_rows[idx]
        else:
            source_row = anchor_row + idx + 1

        avg_pen_calc = None
        if max_quarters_with_data >= n_quarters:
            start_row = numeric_pen_rows[-n_quarters]
            end_row = numeric_pen_rows[-1]
            scratch_cell.formula2 = (
                f"=AVERAGE(R{start_row}C{penetration_input_col}:R{end_row}C{penetration_input_col})"
            )
            workbook.app.calculate()
            avg_pen_calc = scratch_cell.value

        avg_pen_existing = cell_value(sheet, grid, source_row, avg_pen_col)
        avg_penetration = (
            avg_pen_existing if is_number(avg_pen_existing) else avg_pen_calc
        )

        forecast_max = cell_value(sheet, grid, source_row, max_col)
        forecast_min = cell_value(sheet, grid, source_row, min_col)
        forecast_value = cell_value(sheet, grid, source_row, forecast_col)
        reported_sales = cell_value(sheet, grid, source_row, reported_col)
        quarterly_sales = cell_value(sheet, grid, source_row, quarterly_sales_col)
        growth_rate = cell_value(sheet, grid, source_row, growth_col)
        sales_captured = cell_value(sheet, grid, source_row, sales_captured_col)
        num_quarters_used = cell_value(sheet, grid, source_row, num_quarters_col)
        last_quarter_used = cell_value(sheet, grid, source_row, last_quarter_col)

        num_quarters_used_value = (
            int(num_quarters_used) if is_number(num_quarters_used) else n_quarters
        )
        forecast_max_value = to_float(forecast_max)
        forecast_min_value = to_float(forecast_min)
        range_width = None
        if forecast_max_value is not None and forecast_min_value is not None:
            range_width = forecast_max_value - forecast_min_value

        row = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used_value,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": source_file,
        }

        if all(
            row[key] in (None, "")
            for key in (
                "forecast_value",
                "forecast_max",
                "forecast_min",
                "parameter_value",
                "actual_value",
            )
        ):
            continue

        rows.append(row)

    try:
        scratch_cell.value = None
    except Exception:
        pass
    return rows


def build_regression_rows(
    workbook: xw.main.Book,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, REGRESSION_SHEET_NAME)
    grid = read_sheet_grid(sheet)
    anchor_row, anchor_col = find_max_anchor(grid)

    # Required by the extraction specification.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_map = collect_header_map(grid, [anchor_row - 1, anchor_row])
    max_col = anchor_col
    min_col = find_col_by_header(header_map, ["min"], anchor_col + 1)
    forecast_col = find_col_by_header(
        header_map,
        ["tot fcst w o sa", "tot fcst wo sa", "tot fcst without sa", "tot fcst", "forecast"],
        anchor_col - 1,
    )
    num_quarters_col = find_col_by_header(
        header_map,
        ["num quarters", "quarters used", "n quarters"],
        anchor_col - 8,
    )
    actual_col = find_col_by_header(
        header_map,
        ["actual", "reported sales", "actual value"],
        anchor_col - 4,
    )

    candidate_rows = collect_candidate_rows(
        grid=grid,
        anchor_row=anchor_row,
        max_col=max_col,
        min_col=min_col,
        limit=10,
    )
    numeric_pair_rows = find_numeric_pair_rows(
        grid=grid,
        x_col=x_col,
        y_col=y_col,
        stop_row_exclusive=anchor_row,
    )

    if len(numeric_pair_rows) < 2:
        return []

    scratch_row = grid.end_row + 3
    intercept_col = max(grid.end_col + 1, anchor_col + 3)
    slope_col = intercept_col + 1
    intercept_cell = sheet.cells(scratch_row, intercept_col)
    slope_cell = sheet.cells(scratch_row, slope_col)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    max_iterations = min(10, len(numeric_pair_rows))
    for n_quarters in range(2, max_iterations + 1):
        start_row = numeric_pair_rows[-n_quarters]
        end_row = numeric_pair_rows[-1]
        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        workbook.app.calculate()

        intercept_value = intercept_cell.value
        slope_value = slope_cell.value

        source_row_idx = n_quarters - 1
        if source_row_idx < len(candidate_rows):
            source_row = candidate_rows[source_row_idx]
        else:
            source_row = anchor_row + source_row_idx + 1

        num_quarters_raw = cell_value(sheet, grid, source_row, num_quarters_col)
        num_quarters_used = (
            int(num_quarters_raw) if is_number(num_quarters_raw) else n_quarters
        )
        forecast_value = cell_value(sheet, grid, source_row, forecast_col)
        actual_value = cell_value(sheet, grid, source_row, actual_col)
        forecast_max = cell_value(sheet, grid, source_row, max_col)
        forecast_min = cell_value(sheet, grid, source_row, min_col)

        max_float = to_float(forecast_max)
        min_float = to_float(forecast_min)
        range_width = None
        if max_float is not None and min_float is not None:
            range_width = max_float - min_float

        signature = (
            round_or_none(intercept_value),
            round_or_none(slope_value),
            round_or_none(forecast_value),
            round_or_none(forecast_max),
            round_or_none(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        if all(
            row[key] in (None, "")
            for key in ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
        ):
            continue
        rows.append(row)

    try:
        intercept_cell.value = None
        slope_cell.value = None
    except Exception:
        pass
    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    sheet = workbook.create_sheet(sheet_name)
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    last_row = max(sheet.max_row, 1)
    sheet.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=idx).value
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_len:
                max_len = value_len
        sheet.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 50)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    if workbook.active is not None:
        workbook.remove(workbook.active)

    write_sheet(
        workbook=workbook,
        sheet_name="empirical_candidates",
        columns=EMPIRICAL_OUTPUT_COLUMNS,
        rows=empirical_rows,
    )
    write_sheet(
        workbook=workbook,
        sheet_name="regression_candidates",
        columns=REGRESSION_OUTPUT_COLUMNS,
        rows=regression_rows,
    )
    workbook.save(output_path)


def discover_input_files(source_dir: Path) -> list[Path]:
    files: list[Path] = []
    output_file_regex = re.compile(rf"^{re.escape(source_dir.name)}_PARAM(?:\.\d+)?\.xlsx$", re.IGNORECASE)

    for path in sorted(source_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            print(f"Skipped {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue
        if output_file_regex.match(path.name):
            print(f"Skipped {path.name}: appears to be a generated PARAM output file")
            continue
        files.append(path)
    return files


def main() -> None:
    source_input_dir = input_dir.expanduser().resolve()
    destination_dir = output_dir.expanduser().resolve()

    if not source_input_dir.exists() or not source_input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {source_input_dir}")

    destination_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_next_output_path(source_input_dir, destination_dir)

    files = discover_input_files(source_input_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.main.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in files:
            print(f"Processing: {file_path.name}")
            try:
                file_label = parse_file_label(file_path)
            except Exception as exc:
                print(f"Skipped {file_path.name}: filename parse failed ({exc})")
                continue

            workbook: xw.main.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: workbook open failed ({exc})")
                continue

            try:
                empirical_rows.extend(
                    build_empirical_rows(
                        workbook=workbook,
                        file_label=file_label,
                        source_file=file_path.name,
                    )
                )
            except Exception as exc:
                print(f"Skipped empirical extraction for {file_path.name}: {exc}")

            try:
                regression_rows.extend(
                    build_regression_rows(
                        workbook=workbook,
                        file_label=file_label,
                        source_file=file_path.name,
                    )
                )
            except Exception as exc:
                print(f"Skipped regression extraction for {file_path.name}: {exc}")

            processed_files += 1
            safe_close_source_workbook(workbook)

    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

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
