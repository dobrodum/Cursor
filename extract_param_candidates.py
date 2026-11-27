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


# ---------- user-configurable paths ----------
input_dir = Path("/path/to/input_folder")
output_dir = Path("/path/to/output_folder")


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
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).replace(",", "").strip()
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


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = "UNKNOWN"
    if len(parts) >= 2:
        candidate = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if candidate:
            ticker = candidate
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

    return FileMeta(
        model=f"{ticker}_{model_period}",
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
    return SheetScan(origin_row=used.row, origin_col=used.column, values=coerce_2d(used.value))


def find_anchor(scan: SheetScan, target: str = "max") -> tuple[int, int] | None:
    target_text = target.strip().lower()
    for r_idx, row in enumerate(scan.values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target_text:
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
    header_row: int,
    anchor_col: int,
    keywords: Iterable[str],
    default_col: int,
) -> int:
    row_values = get_row_value_map(scan, header_row)
    lowered_keywords = [keyword.lower() for keyword in keywords]

    best_col = None
    best_distance = float("inf")
    for abs_col, value in row_values.items():
        text = normalize_text(value)
        if not text:
            continue
        if any(keyword in text for keyword in lowered_keywords):
            distance = abs(abs_col - anchor_col)
            if distance < best_distance:
                best_distance = distance
                best_col = abs_col

    return best_col if best_col is not None else max(1, default_col)


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


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def extract_empirical_rows(workbook: xw.Book, meta: FileMeta) -> list[dict[str, Any]]:
    sheet = workbook.sheets["Empirical Model"]
    scan = scan_sheet(sheet)
    anchor = find_anchor(scan, "max")
    if anchor is None:
        raise ValueError("max anchor not found in Empirical Model")

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

    helper_col = max_col + 3
    for idx in range(N_QUARTERS):
        abs_row = first_data_row + idx
        formula = (
            f'=IFERROR(AVERAGE(R{first_data_row}C{reported_sales_col}:R{abs_row}C{reported_sales_col})'
            f'/AVERAGE(R{first_data_row}C{quarterly_sales_col}:R{abs_row}C{quarterly_sales_col}),"")'
        )
        set_formula2(sheet.range((first_data_row + idx, helper_col)), formula)

    workbook.app.calculate()
    avg_pen_values = sheet.range((first_data_row, helper_col), (last_data_row, helper_col)).value
    avg_pen_list = avg_pen_values if isinstance(avg_pen_values, list) else [avg_pen_values]

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        avg_pen_formula = to_number(avg_pen_list[idx]) if idx < len(avg_pen_list) else None
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

    sheet.range((first_data_row, helper_col), (last_data_row, helper_col)).clear_contents()
    return rows


def regression_rows_match(left: dict[str, Any], right: dict[str, Any]) -> bool:
    compare_keys = ["num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for key in compare_keys:
        left_value = left.get(key)
        right_value = right.get(key)
        if left_value is None and right_value is None:
            continue
        if left_value is None or right_value is None:
            return False
        if abs(float(left_value) - float(right_value)) > 1e-9:
            return False
    return True


def extract_regression_rows(workbook: xw.Book, meta: FileMeta) -> list[dict[str, Any]]:
    sheet = workbook.sheets["Regression Model"]
    scan = scan_sheet(sheet)
    anchor = find_anchor(scan, "max")
    if anchor is None:
        raise ValueError("max anchor not found in Regression Model")

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
    actual_col = find_header_col(scan, header_row, anchor_col, ["actual", "reported sales"], anchor_col - 1)
    num_quarters_col = find_header_col(
        scan,
        header_row,
        anchor_col,
        ["num quarters", "quarters used"],
        anchor_col - 6,
    )

    first_data_row = anchor_row + 1
    last_data_row = first_data_row + N_QUARTERS - 1

    cols = [x_col, y_col, num_quarters_col, forecast_value_col, actual_col, forecast_max_col, forecast_min_col]
    min_col = min(cols)
    max_col = max(cols)
    block = coerce_2d(sheet.range((first_data_row, min_col), (last_data_row, max_col)).value)

    intercept_col = max_col + 3
    slope_col = max_col + 4
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
        set_formula2(sheet.range((first_data_row + idx, intercept_col)), intercept_formula)
        set_formula2(sheet.range((first_data_row + idx, slope_col)), slope_formula)

    workbook.app.calculate()
    intercept_values = sheet.range((first_data_row, intercept_col), (last_data_row, intercept_col)).value
    slope_values = sheet.range((first_data_row, slope_col), (last_data_row, slope_col)).value
    intercept_list = intercept_values if isinstance(intercept_values, list) else [intercept_values]
    slope_list = slope_values if isinstance(slope_values, list) else [slope_values]

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        intercept = to_number(intercept_list[idx]) if idx < len(intercept_list) else None
        slope = to_number(slope_list[idx]) if idx < len(slope_list) else None

        num_quarters_used = to_number(get_block_value(block, idx, num_quarters_col, min_col))
        if num_quarters_used is None:
            num_quarters_used = float(idx + 1)

        forecast_value = to_number(get_block_value(block, idx, forecast_value_col, min_col))
        if forecast_value is None and intercept is not None and slope is not None:
            x_value = to_number(get_block_value(block, idx, x_col, min_col))
            if x_value is not None:
                forecast_value = intercept + slope * x_value

        actual_value = get_block_value(block, idx, actual_col, min_col)
        if actual_value == "":
            actual_value = None
        actual_numeric = to_number(actual_value)
        if actual_numeric is not None:
            actual_value = actual_numeric

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
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": meta.source_file,
            }
        )

    if len(rows) >= 2 and regression_rows_match(rows[-1], rows[-2]):
        rows.pop()

    sheet.range((first_data_row, intercept_col), (last_data_row, slope_col)).clear_contents()
    return rows


def format_sheet(worksheet, columns: list[str]) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, worksheet.max_row)}"
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for idx, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=idx, max_col=idx):
            value = row[0].value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 2, 60)


def write_output_workbook(output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    output_book = Workbook()
    output_book.remove(output_book.active)

    empirical_sheet = output_book.create_sheet("empirical_candidates")
    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(col) for col in EMPIRICAL_COLUMNS])
    format_sheet(empirical_sheet, EMPIRICAL_COLUMNS)

    regression_sheet = output_book.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(col) for col in REGRESSION_COLUMNS])
    format_sheet(regression_sheet, REGRESSION_COLUMNS)

    output_book.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir()):
        if path.is_file():
            yield path


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_dir}")

    output_path = get_output_path(input_dir, output_dir)
    output_prefix = f"{input_dir.name}_PARAM".lower()

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
            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not an .xlsx file")
                continue
            if file_path.name.lower().startswith(output_prefix):
                print(f"skipped {file_path.name}: output artifact pattern")
                continue

            print(f"processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path)

                workbook_sheet_names = [sheet.name for sheet in workbook.sheets]
                file_empirical_rows: list[dict[str, Any]] = []
                file_regression_rows: list[dict[str, Any]] = []

                if "Empirical Model" in workbook_sheet_names:
                    try:
                        file_empirical_rows = extract_empirical_rows(workbook, meta)
                    except Exception as exc:
                        print(f"skipped {file_path.name}: empirical extraction failed ({exc})")
                else:
                    print(f"skipped {file_path.name}: missing sheet 'Empirical Model'")

                if "Regression Model" in workbook_sheet_names:
                    try:
                        file_regression_rows = extract_regression_rows(workbook, meta)
                    except Exception as exc:
                        print(f"skipped {file_path.name}: regression extraction failed ({exc})")
                else:
                    print(f"skipped {file_path.name}: missing sheet 'Regression Model'")

                if file_empirical_rows or file_regression_rows:
                    empirical_rows.extend(file_empirical_rows)
                    regression_rows.extend(file_regression_rows)
                    processed_files += 1
                else:
                    print(f"skipped {file_path.name}: no extractable rows")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
