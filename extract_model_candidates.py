from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ========= USER INPUTS =========
input_dir = Path("input")
output_dir = Path("output")
# ===============================

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
MAX_QUARTERS = 10

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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}
FILE_PATTERN = re.compile(
    r"Model\s*-\s*(?P<ticker>[^-]+?)\s*-\s*(?P<period>(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4}))_Send$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


@dataclass(frozen=True)
class SheetGrid:
    start_row: int
    start_col: int
    values: list[list[Any]]

    @classmethod
    def from_sheet(cls, sheet: xw.Sheet) -> "SheetGrid":
        used = sheet.used_range
        raw = used.value
        values = ensure_2d(raw)
        return cls(start_row=used.row, start_col=used.column, values=values)

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        row_values = self.values[row_idx]
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]

    @property
    def max_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def max_col(self) -> int:
        width = max((len(row) for row in self.values), default=0)
        return self.start_col + width - 1


def ensure_2d(raw: Any) -> list[list[Any]]:
    if raw is None:
        return []
    if isinstance(raw, (list, tuple)):
        raw_list = list(raw)
        if not raw_list:
            return []
        if isinstance(raw_list[0], (list, tuple)):
            return [list(row) for row in raw_list]
        return [raw_list]
    return [[raw]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def to_float(value: Any) -> float | None:
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
            try:
                return float(text[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def coalesce(*values: Any) -> Any:
    for value in values:
        if value is not None and value != "":
            return value
    return None


def safe_divide(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def parse_model_metadata(file_path: Path) -> ModelMetadata:
    match = FILE_PATTERN.search(file_path.stem)
    if not match:
        raise ValueError("filename does not match expected pattern")

    ticker = match.group("ticker").strip().upper().replace(" ", "")
    bucket = match.group("bucket").title()
    month_token = match.group("month")
    year = int(match.group("year"))

    month_abbrev = month_token[:3].title()
    month_number = datetime.strptime(month_abbrev, "%b").month
    day = DAY_MAP[bucket.lower()]

    model_period = f"{bucket}{month_abbrev}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return ModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def build_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def safe_close_workbook(workbook: xw.Book | None) -> None:
    if workbook is None:
        return
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


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback for older Excel engines that do not expose formula2.
        cell.formula = formula


def find_anchor_cell(grid: SheetGrid, needle: str = "max") -> tuple[int, int]:
    target = needle.strip().lower()
    for r_idx, row_values in enumerate(grid.values):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) == target:
                return grid.start_row + r_idx, grid.start_col + c_idx
    raise ValueError(f"anchor '{needle}' not found")


def collect_numeric_rows(grid: SheetGrid, max_row: int, x_col: int, y_col: int) -> list[int]:
    numeric_rows: list[int] = []
    start = max(grid.start_row, 1)
    for row in range(start, max_row):
        x_val = to_float(grid.get(row, x_col))
        y_val = to_float(grid.get(row, y_col))
        if x_val is not None and y_val is not None:
            numeric_rows.append(row)
    return numeric_rows


def find_penetration_column(grid: SheetGrid, anchor_row: int, anchor_col: int, y_col: int) -> int:
    # Try nearby headers first.
    header_candidates = {"penetration", "avg penetration", "sales captured", "captured in db"}
    for row in range(max(grid.start_row, anchor_row - 2), min(grid.max_row, anchor_row + 2) + 1):
        for col in range(max(grid.start_col, anchor_col - 20), min(grid.max_col, anchor_col + 2) + 1):
            text = normalize_text(grid.get(row, col))
            if any(token in text for token in header_candidates):
                return col

    # Fallback: choose a nearby column with many values in [0, 1.2].
    best_col = max(grid.start_col, y_col - 1)
    best_score = -1
    scan_min = max(grid.start_col, anchor_col - 20)
    scan_max = min(grid.max_col, anchor_col - 1)
    for col in range(scan_min, scan_max + 1):
        score = 0
        for row in range(grid.start_row, anchor_row):
            val = to_float(grid.get(row, col))
            if val is not None and 0 <= val <= 1.2:
                score += 1
        if score > best_score:
            best_score = score
            best_col = col
    return best_col


def maybe_table_value(
    grid: SheetGrid, base_row: int, base_col: int, row_offset: int, col_offset: int
) -> Any:
    return grid.get(base_row + row_offset, base_col + col_offset)


def process_empirical_sheet(wb: xw.Book, metadata: ModelMetadata) -> list[dict[str, Any]]:
    if EMPIRICAL_SHEET_NAME not in {sheet.name for sheet in wb.sheets}:
        return []

    sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    grid = SheetGrid.from_sheet(sheet)
    if not grid.values:
        return []

    anchor_row, anchor_col = find_anchor_cell(grid, "max")
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    penetration_col = find_penetration_column(grid, anchor_row, anchor_col, y_col)
    numeric_rows = collect_numeric_rows(grid, anchor_row, x_col, y_col)
    if not numeric_rows:
        return []

    quarter_rows = [r for r in numeric_rows if to_float(grid.get(r, penetration_col)) is not None]
    if not quarter_rows:
        return []

    n_limit = min(MAX_QUARTERS, len(quarter_rows))
    helper_row = anchor_row + 2
    helper_col = anchor_col + 6

    # Write all average formulas first, then calculate once.
    for idx in range(n_limit):
        n_quarters = idx + 1
        selected = quarter_rows[-n_quarters:]
        start_row = selected[0]
        end_row = selected[-1]
        avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        set_formula2(sheet.cells(helper_row + idx, helper_col), avg_formula)

    wb.app.calculate()
    avg_values = ensure_2d(
        sheet.range((helper_row, helper_col), (helper_row + n_limit - 1, helper_col)).value
    )

    rows: list[dict[str, Any]] = []
    for idx in range(n_limit):
        n_quarters = idx + 1
        selected = quarter_rows[-n_quarters:]
        end_row = selected[-1]
        prev_row = selected[-2] if len(selected) > 1 else None

        avg_penetration = to_float(avg_values[idx][0] if avg_values[idx] else None)
        quarterly_sales = to_float(grid.get(end_row, y_col))
        previous_sales = to_float(grid.get(prev_row, y_col)) if prev_row else None
        growth_rate = (
            (quarterly_sales - previous_sales) / previous_sales
            if quarterly_sales is not None and previous_sales not in (None, 0)
            else None
        )
        estimated_total_sold = safe_divide(quarterly_sales, avg_penetration)

        table_row_offset = n_quarters
        forecast_max = to_float(maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, 0))
        forecast_min = to_float(maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, 1))
        reported_sales = to_float(
            coalesce(
                maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, -2),
                grid.get(end_row, y_col + 1),
                quarterly_sales,
            )
        )
        forecast_value = to_float(
            coalesce(
                maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, -1),
                estimated_total_sold,
            )
        )
        if forecast_max is None:
            forecast_max = forecast_value
        if forecast_min is None:
            forecast_min = forecast_value
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        last_quarter_used = coalesce(
            maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, -5),
            grid.get(end_row, x_col),
        )
        num_quarters_used = to_float(
            coalesce(
                maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, -6),
                n_quarters,
            )
        )
        sales_captured_pct = to_float(
            coalesce(
                maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, -7),
                avg_penetration,
            )
        )

        row = {
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
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": metadata.source_file,
        }
        rows.append(row)

    return rows


def rows_equal_for_regression(a: dict[str, Any], b: dict[str, Any]) -> bool:
    keys = ["num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for key in keys:
        va = a.get(key)
        vb = b.get(key)
        if va is None and vb is None:
            continue
        if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
            if abs(float(va) - float(vb)) > 1e-9:
                return False
        elif va != vb:
            return False
    return True


def process_regression_sheet(wb: xw.Book, metadata: ModelMetadata) -> list[dict[str, Any]]:
    if REGRESSION_SHEET_NAME not in {sheet.name for sheet in wb.sheets}:
        return []

    sheet = wb.sheets[REGRESSION_SHEET_NAME]
    grid = SheetGrid.from_sheet(sheet)
    if not grid.values:
        return []

    anchor_row, anchor_col = find_anchor_cell(grid, "max")
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    numeric_rows = collect_numeric_rows(grid, anchor_row, x_col, y_col)
    if len(numeric_rows) < 2:
        return []

    n_limit = min(MAX_QUARTERS, len(numeric_rows))
    n_values = [n for n in range(1, n_limit + 1)]
    if not n_values:
        return []

    helper_row = anchor_row + 2
    intercept_col = anchor_col + 6
    slope_col = anchor_col + 7
    forecast_col = anchor_col + 8

    for idx, n_quarters in enumerate(n_values):
        selected = numeric_rows[-n_quarters:]
        start_row = selected[0]
        end_row = selected[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2(sheet.cells(helper_row + idx, intercept_col), intercept_formula)
        set_formula2(sheet.cells(helper_row + idx, slope_col), slope_formula)

        next_x = to_float(grid.get(end_row + 1, x_col))
        if next_x is None:
            end_x = to_float(grid.get(end_row, x_col))
            next_x = (end_x + 1) if end_x is not None else float(n_quarters + 1)
        set_formula2(
            sheet.cells(helper_row + idx, forecast_col),
            f"=R{helper_row + idx}C{intercept_col}+R{helper_row + idx}C{slope_col}*{next_x}",
        )

    wb.app.calculate()
    helper_values = ensure_2d(
        sheet.range(
            (helper_row, intercept_col),
            (helper_row + len(n_values) - 1, forecast_col),
        ).value
    )

    rows: list[dict[str, Any]] = []
    for idx, n_quarters in enumerate(n_values):
        helper_row_vals = helper_values[idx] if idx < len(helper_values) else []
        intercept = to_float(helper_row_vals[0] if len(helper_row_vals) > 0 else None)
        slope = to_float(helper_row_vals[1] if len(helper_row_vals) > 1 else None)
        forecast_calc = to_float(helper_row_vals[2] if len(helper_row_vals) > 2 else None)

        table_row_offset = n_quarters
        num_quarters_used = to_float(
            coalesce(
                maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, -5),
                n_quarters,
            )
        )
        forecast_value = to_float(
            coalesce(
                maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, -1),
                forecast_calc,
            )
        )
        forecast_max = to_float(maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, 0))
        forecast_min = to_float(maybe_table_value(grid, anchor_row, anchor_col, table_row_offset, 1))
        if forecast_max is None:
            forecast_max = forecast_value
        if forecast_min is None:
            forecast_min = forecast_value
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": metadata.source_file,
        }

        if rows and rows_equal_for_regression(rows[-1], row):
            continue
        rows.append(row)

    return rows


def write_sheet(ws, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_row = max(ws.max_row, 1)
    last_col_letter = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for col_idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    wb = Workbook()
    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def iter_source_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary workbook")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input directory does not exist: {input_dir}")

    output_path = build_output_path(input_dir, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = "manual"
        try:
            app.api.EnableEvents = False
        except Exception:
            pass

        for file_path in iter_source_files(input_dir):
            try:
                metadata = parse_model_metadata(file_path)
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
                continue

            workbook: xw.Book | None = None
            try:
                print(f"Processing {file_path.name}")
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, metadata))
                regression_rows.extend(process_regression_sheet(workbook, metadata))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                safe_close_workbook(workbook)

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
