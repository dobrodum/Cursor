from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------
# User-configurable paths
# ---------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


@dataclass(frozen=True)
class SheetScan:
    origin_row: int
    origin_col: int
    values: list[list[Any]]


def coerce_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = "UNKNOWN"
    if len(parts) >= 2:
        maybe_ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if maybe_ticker:
            ticker = maybe_ticker
    else:
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )
    if period_match is None:
        period_match = re.search(
            r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})",
            stem,
            flags=re.IGNORECASE,
        )

    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)[:3].title()
        year = period_match.group(3)
        month_num = MONTH_MAP.get(month_token.lower(), 1)
        day_num = DAY_MAP.get(phase.lower(), 15)
        model_period = f"{phase}{month_token}_{year}"
        model_date = date(int(year), month_num, day_num).isoformat()
    else:
        model_period = "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def get_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_stem = f"{input_folder.name}_PARAM"
    primary = output_folder / f"{base_stem}.xlsx"
    if not primary.exists():
        return primary

    suffix = 1
    while True:
        candidate = output_folder / f"{base_stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def scan_sheet(sheet: xw.Sheet) -> SheetScan:
    used = sheet.used_range
    return SheetScan(
        origin_row=used.row,
        origin_col=used.column,
        values=coerce_2d(used.value),
    )


def find_anchor(scan: SheetScan, target: str = "max") -> tuple[int, int] | None:
    expected = target.strip().lower()
    for r_idx, row in enumerate(scan.values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == expected:
                return scan.origin_row + r_idx, scan.origin_col + c_idx
    return None


def get_row_value_map(scan: SheetScan, abs_row: int) -> dict[int, Any]:
    row_idx = abs_row - scan.origin_row
    if row_idx < 0 or row_idx >= len(scan.values):
        return {}
    row_values = scan.values[row_idx]
    return {scan.origin_col + idx: value for idx, value in enumerate(row_values)}


def find_header_col(
    scan: SheetScan,
    anchor_row: int,
    anchor_col: int,
    keywords: Iterable[str],
    default_col: int,
) -> int:
    row_map = get_row_value_map(scan, anchor_row)
    lowered_keywords = [k.lower() for k in keywords]
    best_col: int | None = None
    best_distance = float("inf")

    for abs_col, value in row_map.items():
        text = normalize_text(value)
        if not text:
            continue
        if any(k in text for k in lowered_keywords):
            distance = abs(abs_col - anchor_col)
            if distance < best_distance:
                best_distance = distance
                best_col = abs_col

    if best_col is not None:
        return best_col
    return max(1, default_col)


def get_block_value(block: list[list[Any]], row_offset: int, abs_col: int, min_col: int) -> Any:
    if row_offset < 0 or row_offset >= len(block):
        return None
    row = block[row_offset]
    col_offset = abs_col - min_col
    if col_offset < 0 or col_offset >= len(row):
        return None
    return row[col_offset]


def set_formula2(target_range: xw.Range, formula: str) -> None:
    try:
        target_range.formula2 = formula
    except Exception:
        target_range.formula = formula


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


def extract_empirical_rows(wb: xw.Book, meta: FileMeta) -> list[dict[str, Any]]:
    if "Empirical Model" not in [sht.name for sht in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    scan = scan_sheet(sheet)
    anchor = find_anchor(scan, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_row = anchor_row

    forecast_max_col = anchor_col
    forecast_min_col = find_header_col(scan, header_row, anchor_col, ["min"], anchor_col + 1)
    forecast_value_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["estimated total sold", "est total sold", "forecast", "total sold"],
        anchor_col - 2,
    )
    actual_value_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["reported sales", "actual value", "actual"],
        anchor_col - 1,
    )
    num_quarters_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["num quarters", "quarters used"],
        anchor_col - 6,
    )
    last_quarter_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["last quarter", "quarter"],
        anchor_col - 5,
    )
    quarterly_sales_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["quarterly sales", "qtr sales", "quarter sales"],
        anchor_col - 11,
    )
    reported_sales_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["reported sales", "reported"],
        anchor_col - 10,
    )
    growth_rate_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["growth rate", "growth"],
        anchor_col - 9,
    )
    captured_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["captured in db", "captured", "sales captured"],
        anchor_col - 8,
    )
    avg_pen_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["avg penetration", "average penetration", "penetration"],
        anchor_col - 7,
    )

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1

    cols = [
        num_quarters_col,
        last_quarter_col,
        forecast_value_col,
        actual_value_col,
        forecast_max_col,
        forecast_min_col,
        quarterly_sales_col,
        reported_sales_col,
        growth_rate_col,
        captured_col,
        avg_pen_col,
    ]
    min_col = min(cols)
    max_col = max(cols)
    block = coerce_2d(sheet.range((first_data_row, min_col), (last_data_row, max_col)).value)

    helper_cell = sheet.range((anchor_row, max_col + 3))

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        abs_row = first_data_row + idx

        formula = (
            f'=IFERROR(AVERAGE(R{first_data_row}C{reported_sales_col}:R{abs_row}C{reported_sales_col})'
            f'/AVERAGE(R{first_data_row}C{quarterly_sales_col}:R{abs_row}C{quarterly_sales_col}),"")'
        )
        set_formula2(helper_cell, formula)
        wb.app.calculate()
        avg_pen_formula = to_number(helper_cell.value)

        avg_pen_cell = to_number(get_block_value(block, idx, avg_pen_col, min_col))
        avg_pen = avg_pen_formula if avg_pen_formula is not None else avg_pen_cell

        num_quarters_used = to_number(get_block_value(block, idx, num_quarters_col, min_col))
        if num_quarters_used is None:
            num_quarters_used = float(idx + 1)

        last_quarter_used = get_block_value(block, idx, last_quarter_col, min_col)
        forecast_value = to_number(get_block_value(block, idx, forecast_value_col, min_col))
        actual_value = to_number(get_block_value(block, idx, actual_value_col, min_col))
        forecast_max = to_number(get_block_value(block, idx, forecast_max_col, min_col))
        forecast_min = to_number(get_block_value(block, idx, forecast_min_col, min_col))
        quarterly_sales = to_number(get_block_value(block, idx, quarterly_sales_col, min_col))
        reported_sales = to_number(get_block_value(block, idx, reported_sales_col, min_col))
        growth_rate = to_number(get_block_value(block, idx, growth_rate_col, min_col))
        captured = to_number(get_block_value(block, idx, captured_col, min_col))

        if all(
            value is None
            for value in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
                avg_pen,
            ]
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured,
                "source_file": meta.source_file,
            }
        )

    return rows


def _regression_rows_match(left: dict[str, Any], right: dict[str, Any]) -> bool:
    keys = ["num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for key in keys:
        l_val = left.get(key)
        r_val = right.get(key)
        if l_val is None and r_val is None:
            continue
        if l_val is None or r_val is None:
            return False
        if abs(float(l_val) - float(r_val)) > 1e-9:
            return False
    return True


def extract_regression_rows(wb: xw.Book, meta: FileMeta) -> list[dict[str, Any]]:
    if "Regression Model" not in [sht.name for sht in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    scan = scan_sheet(sheet)
    anchor = find_anchor(scan, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_row = anchor_row

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    forecast_max_col = anchor_col
    forecast_min_col = find_header_col(scan, header_row, anchor_col, ["min"], anchor_col + 1)
    forecast_value_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["tot fcst w/o sa", "total forecast w/o sa", "forecast w/o sa", "tot fcst"],
        anchor_col - 2,
    )
    actual_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["actual", "reported sales"],
        anchor_col - 1,
    )
    num_quarters_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["num quarters", "quarters used"],
        anchor_col - 6,
    )

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1

    cols = [
        x_col,
        y_col,
        num_quarters_col,
        forecast_value_col,
        actual_col,
        forecast_max_col,
        forecast_min_col,
    ]
    min_col = min(cols)
    max_col = max(cols)
    block = coerce_2d(sheet.range((first_data_row, min_col), (last_data_row, max_col)).value)

    intercept_cell = sheet.range((anchor_row, max_col + 3))
    slope_cell = sheet.range((anchor_row, max_col + 4))

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        abs_row = first_data_row + idx
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{first_data_row}C{y_col}:R{abs_row}C{y_col},'
            f'R{first_data_row}C{x_col}:R{abs_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{first_data_row}C{y_col}:R{abs_row}C{y_col},'
            f'R{first_data_row}C{x_col}:R{abs_row}C{x_col}),"")'
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_number(intercept_cell.value)
        slope = to_number(slope_cell.value)
        num_quarters_used = to_number(get_block_value(block, idx, num_quarters_col, min_col))
        if num_quarters_used is None:
            num_quarters_used = float(idx + 1)

        forecast_value = to_number(get_block_value(block, idx, forecast_value_col, min_col))
        if forecast_value is None and intercept is not None and slope is not None:
            x_val = to_number(get_block_value(block, idx, x_col, min_col))
            if x_val is not None:
                forecast_value = intercept + slope * x_val

        actual_value = get_block_value(block, idx, actual_col, min_col)
        if actual_value == "":
            actual_value = None
        actual_value_num = to_number(actual_value)
        if actual_value_num is not None:
            actual_value = actual_value_num

        forecast_max = to_number(get_block_value(block, idx, forecast_max_col, min_col))
        forecast_min = to_number(get_block_value(block, idx, forecast_min_col, min_col))

        if all(value is None for value in [forecast_value, forecast_max, forecast_min, intercept, slope]):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": meta.source_file,
            }
        )

    if len(rows) >= 2 and _regression_rows_match(rows[-1], rows[-2]):
        rows.pop()

    return rows


def format_sheet(ws, columns: list[str]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=idx, max_col=idx):
            value = row[0].value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row_data in empirical_rows:
        empirical_ws.append([row_data.get(col) for col in EMPIRICAL_COLUMNS])
    format_sheet(empirical_ws, EMPIRICAL_COLUMNS)

    regression_ws = out_wb.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_COLUMNS)
    for row_data in regression_rows:
        regression_ws.append([row_data.get(col) for col in REGRESSION_COLUMNS])
    format_sheet(regression_ws, REGRESSION_COLUMNS)

    out_wb.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        yield path


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_dir}")

    output_path = get_output_path(input_dir, output_dir)
    output_prefix = f"{input_dir.name}_PARAM"

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in iter_input_files(input_dir):
            lower_name = file_path.name.lower()
            if file_path.name.startswith("~"):
                print(f"SKIPPED {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED {file_path.name}: not an .xlsx file")
                continue
            if lower_name.startswith(output_prefix.lower()):
                print(f"SKIPPED {file_path.name}: output artifact pattern")
                continue

            print(f"PROCESSING {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path)
                file_empirical_rows = extract_empirical_rows(wb, meta)
                file_regression_rows = extract_regression_rows(wb, meta)

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
