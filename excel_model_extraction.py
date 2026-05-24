from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
MAX_ANCHOR_TEXT = "max"
MAX_N_QUARTERS = 10

# Anchor-based column offsets from the "max" anchor cell.
X_COL_OFFSET = -11
Y_COL_OFFSET = -7
QUARTER_LABEL_COL_OFFSET_FROM_X = -1

# Temporary calculation area starts to the right of anchor.
EMPIRICAL_TEMP_COL_OFFSET = 2
REGRESSION_TEMP_COL_OFFSET = 2

# Optional pull offsets around the anchor (row_offset, col_offset).
REGRESSION_FORECAST_TOTAL_OFFSET = (-1, 1)
REGRESSION_FORECAST_MAX_OFFSET = (0, 1)
REGRESSION_FORECAST_MIN_OFFSET = (1, 1)

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

PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)[ _-]*([A-Za-z]{3,9})[ _-]*(\d{4})", re.IGNORECASE)
TICKER_FALLBACK_PATTERN = re.compile(r"-\s*([A-Za-z0-9]{1,12})\s*-")
MONTH_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass(frozen=True)
class ModelFileMeta:
    ticker: str
    model_period: str
    model_date: str
    model: str


def to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def as_list(value: Any) -> list[Any]:
    if isinstance(value, list):
        return value
    return [value]


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def safe_close_workbook(workbook: xw.Book) -> None:
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


def set_formula2(target_cell: xw.Range, formula: str) -> None:
    try:
        target_cell.formula2 = formula
        return
    except Exception:
        pass

    try:
        target_cell.api.Formula2 = formula
        return
    except Exception:
        pass

    # Last fallback for older Excel engines.
    target_cell.formula = formula


def parse_month_token(month_token: str) -> tuple[int, str]:
    normalized = re.sub(r"[^A-Za-z]", "", month_token).lower()
    for month_num in range(1, 13):
        month_full = calendar.month_name[month_num].lower()
        month_abbr = calendar.month_abbr[month_num].lower()
        if normalized.startswith(month_full) or normalized.startswith(month_abbr):
            return month_num, calendar.month_abbr[month_num]
    raise ValueError(f"Unknown month token: {month_token}")


def parse_file_metadata(file_name: str) -> ModelFileMeta:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()
    if not ticker:
        ticker_match = TICKER_FALLBACK_PATTERN.search(stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()
    if not ticker:
        ticker = "UNKNOWN"

    model_period = "UNKNOWN_PERIOD"
    model_date = ""
    period_match = PERIOD_PATTERN.search(stem)
    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month_num, month_abbr = parse_month_token(month_token)
        day = MONTH_DAY_MAP[phase]
        model_period = f"{phase}{month_abbr}_{year}"
        model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return ModelFileMeta(ticker=ticker, model_period=model_period, model_date=model_date, model=model)


def get_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def get_sheet_by_name(book: xw.Book, target_name: str) -> xw.Sheet | None:
    lowered_target = target_name.strip().lower()
    for sheet in book.sheets:
        if sheet.name.strip().lower() == lowered_target:
            return sheet
    return None


def find_anchor(sheet: xw.Sheet, anchor_text: str = MAX_ANCHOR_TEXT) -> tuple[int, int]:
    used_range = sheet.used_range
    values = ensure_2d(used_range.value)
    base_row = used_range.row
    base_col = used_range.column

    lowered_anchor = anchor_text.strip().lower()
    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == lowered_anchor:
                return base_row + row_offset, base_col + col_offset

    raise ValueError(f'Anchor "{anchor_text}" not found in sheet "{sheet.name}"')


def collect_history_rows(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    anchor_row: int,
    quarter_col: int | None = None,
) -> list[dict[str, Any]]:
    if anchor_row <= 1 or x_col <= 0 or y_col <= 0:
        return []

    last_data_row = anchor_row - 1
    x_values = as_list(sheet.range((1, x_col), (last_data_row, x_col)).options(ndim=1).value)
    y_values = as_list(sheet.range((1, y_col), (last_data_row, y_col)).options(ndim=1).value)
    q_values: list[Any] = []

    if quarter_col and quarter_col > 0:
        q_values = as_list(sheet.range((1, quarter_col), (last_data_row, quarter_col)).options(ndim=1).value)

    rows: list[dict[str, Any]] = []
    for idx in range(last_data_row):
        x_val = to_float(x_values[idx] if idx < len(x_values) else None)
        y_val = to_float(y_values[idx] if idx < len(y_values) else None)
        if x_val is None or y_val is None or y_val == 0:
            continue

        quarter_label = q_values[idx] if idx < len(q_values) else None
        rows.append(
            {
                "row": idx + 1,
                "x": x_val,
                "y": y_val,
                "quarter": quarter_label,
            }
        )

    return rows


def format_quarter_label(raw_value: Any, row_num: int) -> str:
    if raw_value not in (None, ""):
        return str(raw_value).strip()
    return f"row_{row_num}"


def numeric_equal(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    if a in (None, "") and b in (None, ""):
        return True

    a_float = to_float(a)
    b_float = to_float(b)
    if a_float is not None and b_float is not None:
        return abs(a_float - b_float) <= tolerance
    return a == b


def is_duplicate_regression_row(previous: dict[str, Any], current: dict[str, Any]) -> bool:
    comparison_keys = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    return all(numeric_equal(previous.get(key), current.get(key)) for key in comparison_keys)


def extract_empirical_candidates(
    sheet: xw.Sheet,
    metadata: ModelFileMeta,
    source_file: str,
    workbook: xw.Book,
) -> list[dict[str, Any]]:
    anchor_row, anchor_col = find_anchor(sheet, MAX_ANCHOR_TEXT)
    x_col = anchor_col + X_COL_OFFSET
    y_col = anchor_col + Y_COL_OFFSET
    quarter_col = x_col + QUARTER_LABEL_COL_OFFSET_FROM_X

    history = collect_history_rows(sheet, x_col, y_col, anchor_row, quarter_col)
    if not history:
        return []

    latest = history[-1]
    latest_row = latest["row"]
    latest_quarterly_sales = latest["x"]
    latest_reported_sales = latest["y"]
    latest_quarter_label = format_quarter_label(latest.get("quarter"), latest_row)

    previous_reported_sales = history[-2]["y"] if len(history) >= 2 else None
    growth_rate_pct = None
    if previous_reported_sales not in (None, 0):
        growth_rate_pct = (latest_reported_sales - previous_reported_sales) / previous_reported_sales

    sales_captured_pct = None
    if latest_reported_sales not in (None, 0):
        sales_captured_pct = latest_quarterly_sales / latest_reported_sales

    max_n = min(MAX_N_QUARTERS, len(history))
    temp_col = anchor_col + EMPIRICAL_TEMP_COL_OFFSET

    calc_specs: list[dict[str, Any]] = []
    for n_quarters in range(1, max_n + 1):
        start_row = history[-n_quarters]["row"]
        calc_row = anchor_row + n_quarters

        avg_formula = (
            f"=AVERAGE(R{start_row}C{x_col}:R{latest_row}C{x_col}/"
            f"R{start_row}C{y_col}:R{latest_row}C{y_col})"
        )
        min_formula = (
            f"=MIN(R{start_row}C{x_col}:R{latest_row}C{x_col}/"
            f"R{start_row}C{y_col}:R{latest_row}C{y_col})"
        )
        max_formula = (
            f"=MAX(R{start_row}C{x_col}:R{latest_row}C{x_col}/"
            f"R{start_row}C{y_col}:R{latest_row}C{y_col})"
        )

        set_formula2(sheet.range((calc_row, temp_col)), avg_formula)
        set_formula2(sheet.range((calc_row, temp_col + 1)), min_formula)
        set_formula2(sheet.range((calc_row, temp_col + 2)), max_formula)
        calc_specs.append({"n_quarters": n_quarters, "calc_row": calc_row})

    if calc_specs:
        workbook.app.calculate()

    rows: list[dict[str, Any]] = []
    for spec in calc_specs:
        calc_row = spec["calc_row"]
        n_quarters = spec["n_quarters"]

        avg_penetration_pct = to_float(sheet.range((calc_row, temp_col)).value)
        min_penetration_pct = to_float(sheet.range((calc_row, temp_col + 1)).value)
        max_penetration_pct = to_float(sheet.range((calc_row, temp_col + 2)).value)

        forecast_value = None
        forecast_max = None
        forecast_min = None

        if avg_penetration_pct not in (None, 0):
            forecast_value = latest_quarterly_sales / avg_penetration_pct
        if min_penetration_pct not in (None, 0):
            forecast_max = latest_quarterly_sales / min_penetration_pct
        if max_penetration_pct not in (None, 0):
            forecast_min = latest_quarterly_sales / max_penetration_pct

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": latest_quarter_label,
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
        )

    return rows


def extract_regression_candidates(
    sheet: xw.Sheet,
    metadata: ModelFileMeta,
    source_file: str,
    workbook: xw.Book,
) -> list[dict[str, Any]]:
    anchor_row, anchor_col = find_anchor(sheet, MAX_ANCHOR_TEXT)
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    history = collect_history_rows(sheet, x_col, y_col, anchor_row)
    if len(history) < 2:
        return []

    latest = history[-1]
    latest_row = latest["row"]
    latest_x_value = latest["x"]
    max_n = min(MAX_N_QUARTERS, len(history))
    temp_col = anchor_col + REGRESSION_TEMP_COL_OFFSET

    calc_specs: list[dict[str, Any]] = []
    for n_quarters in range(2, max_n + 1):
        start_row = history[-n_quarters]["row"]
        calc_row = anchor_row + n_quarters

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )
        forecast_formula = f"=RC[-2]+RC[-1]*{latest_x_value:.15g}"

        set_formula2(sheet.range((calc_row, temp_col)), intercept_formula)
        set_formula2(sheet.range((calc_row, temp_col + 1)), slope_formula)
        set_formula2(sheet.range((calc_row, temp_col + 2)), forecast_formula)

        calc_specs.append({"n_quarters": n_quarters, "calc_row": calc_row})

    if calc_specs:
        workbook.app.calculate()

    base_forecast_total = to_float(
        sheet.range(
            (
                anchor_row + REGRESSION_FORECAST_TOTAL_OFFSET[0],
                anchor_col + REGRESSION_FORECAST_TOTAL_OFFSET[1],
            )
        ).value
    )
    base_forecast_max = to_float(
        sheet.range(
            (
                anchor_row + REGRESSION_FORECAST_MAX_OFFSET[0],
                anchor_col + REGRESSION_FORECAST_MAX_OFFSET[1],
            )
        ).value
    )
    base_forecast_min = to_float(
        sheet.range(
            (
                anchor_row + REGRESSION_FORECAST_MIN_OFFSET[0],
                anchor_col + REGRESSION_FORECAST_MIN_OFFSET[1],
            )
        ).value
    )

    rows: list[dict[str, Any]] = []
    for spec in calc_specs:
        calc_row = spec["calc_row"]
        n_quarters = spec["n_quarters"]

        intercept_value = to_float(sheet.range((calc_row, temp_col)).value)
        slope_value = to_float(sheet.range((calc_row, temp_col + 1)).value)
        forecast_value = to_float(sheet.range((calc_row, temp_col + 2)).value)

        # If workbook has a dedicated TOT FCST w/o SA anchor value, prefer it for the
        # widest (max_n) window where existing models typically place final output.
        if n_quarters == max_n and base_forecast_total is not None:
            forecast_value = base_forecast_total

        forecast_max = base_forecast_max
        forecast_min = base_forecast_min
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(rows[-1], row):
            continue
        rows.append(row)

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row in rows:
            value = row.get(column, "")
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 42)


def write_output_workbook(
    destination_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    workbook.save(destination_path)


def run() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {in_dir}")
    if not in_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a folder: {in_dir}")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
            if not path.is_file():
                print(f"SKIPPED: {path.name} (not a file)")
                continue
            if path.name.startswith("~"):
                print(f"SKIPPED: {path.name} (temp file)")
                continue
            if path.suffix.lower() != ".xlsx":
                print(f"SKIPPED: {path.name} (not .xlsx)")
                continue

            print(f"Processing: {path.name}")
            metadata = parse_file_metadata(path.name)

            source_book: xw.Book | None = None
            try:
                source_book = app.books.open(str(path), update_links=False)

                empirical_sheet = get_sheet_by_name(source_book, EMPIRICAL_SHEET_NAME)
                if empirical_sheet is None:
                    print(f"SKIPPED: {path.name} ({EMPIRICAL_SHEET_NAME} not found)")
                else:
                    empirical_rows.extend(
                        extract_empirical_candidates(
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=path.name,
                            workbook=source_book,
                        )
                    )

                regression_sheet = get_sheet_by_name(source_book, REGRESSION_SHEET_NAME)
                if regression_sheet is None:
                    print(f"SKIPPED: {path.name} ({REGRESSION_SHEET_NAME} not found)")
                else:
                    regression_rows.extend(
                        extract_regression_candidates(
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=path.name,
                            workbook=source_book,
                        )
                    )

                files_processed += 1
            except Exception as exc:
                print(f"SKIPPED: {path.name} (error: {exc})")
            finally:
                if source_book is not None:
                    safe_close_workbook(source_book)
    finally:
        app.quit()

    output_path = get_output_path(in_dir=in_dir, out_dir=out_dir)
    write_output_workbook(
        destination_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
