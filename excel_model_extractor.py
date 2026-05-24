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


# Configure these two paths before running.
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


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

# Fallback offsets if labels cannot be discovered near the "max" anchor.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "avg_penetration_pct": -5,
    "quarterly_sales": -6,
    "reported_sales": -2,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetGrid:
    values: List[List[Any]]
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or row > self.end_row:
            return None
        if col < self.start_col or col > self.end_col:
            return None
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if r_idx < 0 or r_idx >= len(self.values):
            return None
        row_values = self.values[r_idx]
        if c_idx < 0 or c_idx >= len(row_values):
            return None
        return row_values[c_idx]


def normalize_2d(values: Any) -> List[List[Any]]:
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
        first = list(first)
    if isinstance(first, list):
        matrix: List[List[Any]] = []
        max_len = 0
        for row in values:
            if isinstance(row, tuple):
                row = list(row)
            if not isinstance(row, list):
                row = [row]
            matrix.append(row)
            if len(row) > max_len:
                max_len = len(row)
        if max_len:
            for row in matrix:
                if len(row) < max_len:
                    row.extend([None] * (max_len - len(row)))
        return matrix
    return [list(values)]


def to_list(values: Any) -> List[Any]:
    matrix = normalize_2d(values)
    if not matrix:
        return []
    if len(matrix) == 1:
        return matrix[0]
    if len(matrix[0]) == 1:
        return [row[0] for row in matrix]
    return [item for row in matrix for item in row]


def read_sheet_grid(sheet: xw.Sheet) -> SheetGrid:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    start_row = used.row
    start_col = used.column
    row_count = len(matrix)
    col_count = len(matrix[0]) if matrix else 0
    end_row = start_row + row_count - 1 if row_count else start_row
    end_col = start_col + col_count - 1 if col_count else start_col
    return SheetGrid(
        values=matrix,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def find_anchor_cell(grid: SheetGrid, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    anchor_text = anchor_text.strip().lower()
    for r_idx, row in enumerate(grid.values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == anchor_text:
                return grid.start_row + r_idx, grid.start_col + c_idx
    return None


def choose_col(current: Optional[int], candidate: int, anchor_col: int) -> int:
    if current is None:
        return candidate
    if abs(candidate - anchor_col) < abs(current - anchor_col):
        return candidate
    return current


def discover_columns(
    grid: SheetGrid,
    anchor_row: int,
    anchor_col: int,
    header_hints: Dict[str, Sequence[str]],
    fallback_offsets: Dict[str, int],
    header_window: int = 3,
) -> Dict[str, int]:
    columns: Dict[str, int] = {}
    row_min = max(grid.start_row, anchor_row - header_window)
    row_max = min(grid.end_row, anchor_row + 1)
    for row in range(row_min, row_max + 1):
        for col in range(grid.start_col, grid.end_col + 1):
            normalized = normalize_text(grid.get(row, col))
            if not normalized:
                continue
            for field, hints in header_hints.items():
                if any(hint in normalized for hint in hints):
                    current = columns.get(field)
                    columns[field] = choose_col(current, col, anchor_col)

    for field, offset in fallback_offsets.items():
        if field not in columns:
            columns[field] = anchor_col + offset

    columns["forecast_max"] = anchor_col
    return columns


def maybe_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
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


def numeric_or_original(value: Any) -> Any:
    number = maybe_float(value)
    return number if number is not None else value


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = maybe_float(left)
    right_num = maybe_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def set_r1c1_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    cell.formula = formula_r1c1


def close_workbook_no_save(wb: xw.Book) -> None:
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


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    pieces = [piece.strip() for piece in stem.split("-") if piece.strip()]
    ticker = "UNKNOWN"
    period_token = ""
    if len(pieces) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", pieces[1]).upper() or "UNKNOWN"
    if len(pieces) >= 3:
        period_token = pieces[2].split("_")[0].strip()
    period_match = re.search(
        r"(?i)(early|mid|late)\s*([A-Za-z]{3})\s*(\d{4})",
        period_token,
    )
    if period_match:
        period_part = period_match.group(1).title()
        month_part = period_match.group(2).title()
        year_part = period_match.group(3)
        month_num = _month_to_num(month_part)
        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        day = day_map.get(period_part, 15)
        model_period = f"{period_part}{month_part}_{year_part}"
        model_date = date(int(year_part), month_num, day).isoformat()
    else:
        model_period = period_token or "UnknownPeriod"
        model_date = ""
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def _month_to_num(month_abbrev: str) -> int:
    month_map = {
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
    return month_map.get(month_abbrev.title(), 1)


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    grid = read_sheet_grid(sheet)
    anchor = find_anchor_cell(grid, "max")
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    header_hints = {
        "num_quarters_used": ("num quarter", "quarters used", "# quarter"),
        "last_quarter_used": ("last quarter",),
        "avg_penetration_pct": ("avg penetration", "average penetration"),
        "penetration_input_pct": ("penetration",),
        "quarterly_sales": ("quarterly sales", "qtr sales"),
        "reported_sales": ("reported sales",),
        "forecast_value": ("estimated total sold", "forecast total", "tot fcst"),
        "actual_value": ("actual", "reported"),
        "forecast_min": ("min",),
        "growth_rate_pct": ("growth rate", "growth"),
        "sales_captured_in_db_pct": ("captured in db", "sales captured", "db pct"),
    }
    column_map = discover_columns(
        grid=grid,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        header_hints=header_hints,
        fallback_offsets=EMPIRICAL_FALLBACK_OFFSETS,
    )

    data_start_row = anchor_row + 1
    helper_col = grid.end_col + 1
    penetration_source_col = column_map.get("penetration_input_pct") or column_map.get("avg_penetration_pct")
    wrote_formula = False
    if penetration_source_col:
        for idx in range(N_QUARTERS):
            row = data_start_row + idx
            formula = (
                f'=IFERROR(AVERAGE(R{data_start_row}C{penetration_source_col}:'
                f"RC{penetration_source_col}),\"\")"
            )
            set_r1c1_formula2(sheet.cells(row, helper_col), formula)
            wrote_formula = True

    helper_values: List[Any] = [None] * N_QUARTERS
    if wrote_formula:
        wb.app.calculate()
        helper_range = sheet.range(
            (data_start_row, helper_col),
            (data_start_row + N_QUARTERS - 1, helper_col),
        ).value
        helper_values = to_list(helper_range)
        if len(helper_values) < N_QUARTERS:
            helper_values.extend([None] * (N_QUARTERS - len(helper_values)))

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        num_quarters_used = grid.get(row, column_map["num_quarters_used"])
        last_quarter_used = grid.get(row, column_map["last_quarter_used"])
        forecast_value = grid.get(row, column_map["forecast_value"])
        forecast_max = grid.get(row, column_map["forecast_max"])
        forecast_min = grid.get(row, column_map["forecast_min"])
        quarterly_sales = grid.get(row, column_map["quarterly_sales"])
        reported_sales = grid.get(row, column_map["reported_sales"])
        growth_rate_pct = grid.get(row, column_map["growth_rate_pct"])
        sales_captured_in_db_pct = grid.get(row, column_map["sales_captured_in_db_pct"])
        avg_penetration_pct = grid.get(row, column_map["avg_penetration_pct"])
        if avg_penetration_pct in (None, "") and idx < len(helper_values):
            avg_penetration_pct = helper_values[idx]
        if num_quarters_used in (None, ""):
            num_quarters_used = idx + 1

        if all(
            value in (None, "")
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        actual_value = grid.get(row, column_map["actual_value"])
        if actual_value in (None, ""):
            actual_value = reported_sales

        row_payload = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": numeric_or_original(avg_penetration_pct),
            "num_quarters_used": numeric_or_original(num_quarters_used),
            "last_quarter_used": last_quarter_used,
            "forecast_value": numeric_or_original(forecast_value),
            "actual_value": numeric_or_original(actual_value),
            "forecast_max": numeric_or_original(forecast_max),
            "forecast_min": numeric_or_original(forecast_min),
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": numeric_or_original(avg_penetration_pct),
            "quarterly_sales": numeric_or_original(quarterly_sales),
            "reported_sales": numeric_or_original(reported_sales),
            "growth_rate_pct": numeric_or_original(growth_rate_pct),
            "sales_captured_in_db_pct": numeric_or_original(sales_captured_in_db_pct),
            "source_file": source_file,
        }
        rows.append(row_payload)

    return rows


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    grid = read_sheet_grid(sheet)
    anchor = find_anchor_cell(grid, "max")
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    header_hints = {
        "num_quarters_used": ("num quarter", "quarters used", "# quarter"),
        "forecast_value": ("tot fcst w o sa", "tot fcst", "forecast total"),
        "actual_value": ("actual", "reported"),
        "forecast_min": ("min",),
    }
    column_map = discover_columns(
        grid=grid,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        header_hints=header_hints,
        fallback_offsets=REGRESSION_FALLBACK_OFFSETS,
    )

    data_start_row = anchor_row + 1
    intercept_col = grid.end_col + 1
    slope_col = intercept_col + 1

    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        formula_intercept = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f"R{data_start_row}C{x_col}:R{row}C{x_col}),\"\")"
        )
        formula_slope = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f"R{data_start_row}C{x_col}:R{row}C{x_col}),\"\")"
        )
        set_r1c1_formula2(sheet.cells(row, intercept_col), formula_intercept)
        set_r1c1_formula2(sheet.cells(row, slope_col), formula_slope)

    wb.app.calculate()
    intercept_values = to_list(
        sheet.range(
            (data_start_row, intercept_col),
            (data_start_row + N_QUARTERS - 1, intercept_col),
        ).value
    )
    slope_values = to_list(
        sheet.range(
            (data_start_row, slope_col),
            (data_start_row + N_QUARTERS - 1, slope_col),
        ).value
    )
    if len(intercept_values) < N_QUARTERS:
        intercept_values.extend([None] * (N_QUARTERS - len(intercept_values)))
    if len(slope_values) < N_QUARTERS:
        slope_values.extend([None] * (N_QUARTERS - len(slope_values)))

    rows: List[Dict[str, Any]] = []
    previous_final_signature: Optional[Tuple[Any, ...]] = None
    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        num_quarters_used = grid.get(row, column_map["num_quarters_used"])
        if num_quarters_used in (None, ""):
            num_quarters_used = idx + 1
        forecast_value = grid.get(row, column_map["forecast_value"])
        actual_value = grid.get(row, column_map["actual_value"])
        forecast_max = grid.get(row, column_map["forecast_max"])
        forecast_min = grid.get(row, column_map["forecast_min"])
        intercept_value = intercept_values[idx] if idx < len(intercept_values) else None
        slope_value = slope_values[idx] if idx < len(slope_values) else None

        if all(
            value in (None, "")
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        ):
            continue

        signature = (
            maybe_float(intercept_value),
            maybe_float(slope_value),
            maybe_float(forecast_value),
            maybe_float(forecast_max),
            maybe_float(forecast_min),
        )
        if idx == N_QUARTERS - 1 and previous_final_signature == signature:
            continue
        previous_final_signature = signature

        row_payload = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": numeric_or_original(num_quarters_used),
            "num_quarters_used": numeric_or_original(num_quarters_used),
            "forecast_value": numeric_or_original(forecast_value),
            "actual_value": numeric_or_original(actual_value),
            "forecast_max": numeric_or_original(forecast_max),
            "forecast_min": numeric_or_original(forecast_min),
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": numeric_or_original(intercept_value),
            "slope": numeric_or_original(slope_value),
            "source_file": source_file,
        }
        rows.append(row_payload)

    return rows


def iter_candidate_files(folder: Path) -> Iterable[Path]:
    if not folder.exists():
        print(f"Skipped input scan: folder does not exist -> {folder}")
        return []
    paths: List[Path] = []
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        paths.append(file_path)
    return paths


def choose_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def format_output_sheet(ws, headers: Sequence[str]) -> None:
    for index, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=index, max_col=index):
            value = row[0].value
            if value is None:
                continue
            cell_len = len(str(value))
            if cell_len > max_len:
                max_len = cell_len
        ws.column_dimensions[get_column_letter(index)].width = min(max(max_len + 2, 12), 45)

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_ws.append([row.get(column, "") for column in EMPIRICAL_COLUMNS])
    format_output_sheet(empirical_ws, EMPIRICAL_COLUMNS)

    regression_ws = workbook.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_ws.append([row.get(column, "") for column in REGRESSION_COLUMNS])
    format_output_sheet(regression_ws, REGRESSION_COLUMNS)

    workbook.save(output_path)


def run() -> None:
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    files = list(iter_candidate_files(input_dir))
    output_path = choose_output_path(input_dir, output_dir)

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in files:
            print(f"Processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                files_processed += 1
                label = parse_file_label(file_path)

                empirical_sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical in {file_path.name}: sheet 'Empirical Model' not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=workbook,
                            sheet=empirical_sheet,
                            label=label,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = get_sheet_case_insensitive(workbook, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression in {file_path.name}: sheet 'Regression Model' not found")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=workbook,
                            sheet=regression_sheet,
                            label=label,
                            source_file=file_path.name,
                        )
                    )
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_workbook_no_save(workbook)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
