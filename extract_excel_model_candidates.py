from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# User-configurable paths
# =========================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


@dataclass
class SheetScan:
    anchor: tuple[int, int] | None
    labels: dict[str, list[tuple[int, int]]]


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

PERIOD_DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


def normalize_label(value: str) -> str:
    text = re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
    return re.sub(r"\s+", " ", text)


def to_number(value: Any) -> float | None:
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
        is_pct = text.endswith("%")
        if is_pct:
            text = text[:-1]
        try:
            number = float(text)
        except ValueError:
            return None
        return number / 100 if is_pct else number
    return None


def parse_filename_metadata(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    ticker = "UNKNOWN"
    model_period = "Unknown_0000"
    model_date = ""

    file_regex = re.compile(
        r"""
        ^.*?
        \s*-\s*
        (?P<ticker>[A-Za-z0-9]+)
        \s*-\s*
        (?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})
        (?:_Send)?$
        """,
        re.IGNORECASE | re.VERBOSE,
    )
    match = file_regex.match(stem)

    if match:
        ticker = match.group("ticker").upper()
        period_raw = match.group("period")
        period_match = re.match(
            r"(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
            period_raw,
            flags=re.IGNORECASE,
        )
        if period_match:
            bucket = period_match.group("bucket")
            month_token = period_match.group("month")
            year_token = period_match.group("year")
            month_num = MONTH_MAP.get(month_token.lower())
            if month_num:
                model_period = f"{bucket}{month_token}_{year_token}"
                day_num = PERIOD_DAY_MAP[bucket.lower()]
                model_date = date(int(year_token), month_num, day_num).isoformat()
    else:
        parts = [p.strip() for p in stem.split("-")]
        if len(parts) >= 3:
            ticker_guess = parts[1].strip()
            if ticker_guess:
                ticker = ticker_guess.upper()

    model = f"{ticker}_{model_period}"
    return ModelMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def find_unique_output_path(source_input_dir: Path, source_output_dir: Path) -> Path:
    input_folder_name = source_input_dir.resolve().name
    base_name = f"{input_folder_name}_PARAM"
    candidate = source_output_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = source_output_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def scan_sheet_once(sheet: xw.Sheet) -> SheetScan:
    used = sheet.used_range
    values = used.options(ndim=2).value
    if values is None:
        return SheetScan(anchor=None, labels={})

    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    labels: dict[str, list[tuple[int, int]]] = {}
    anchor: tuple[int, int] | None = None

    base_row = used.row
    base_col = used.column
    for r_idx, row_values in enumerate(values):
        for c_idx, raw_value in enumerate(row_values):
            if not isinstance(raw_value, str):
                continue
            key = normalize_label(raw_value)
            if not key:
                continue
            row = base_row + r_idx
            col = base_col + c_idx
            labels.setdefault(key, []).append((row, col))
            if key == "max" and anchor is None:
                anchor = (row, col)

    return SheetScan(anchor=anchor, labels=labels)


def first_label_position(scan: SheetScan, patterns: list[str]) -> tuple[int, int] | None:
    norm_patterns = [normalize_label(p) for p in patterns]
    for key, positions in scan.labels.items():
        if any(pattern in key for pattern in norm_patterns):
            return positions[0]
    return None


def value_cell_from_label(
    sheet: xw.Sheet, scan: SheetScan, label_patterns: list[str], right_probe: int = 5
) -> tuple[int, int] | None:
    label_pos = first_label_position(scan, label_patterns)
    if not label_pos:
        return None

    row, col = label_pos
    right_values = sheet.range((row, col + 1), (row, col + right_probe)).value
    if not isinstance(right_values, list):
        right_values = [right_values]

    for offset, value in enumerate(right_values, start=1):
        if value not in (None, ""):
            return (row, col + offset)
    return (row, col + 1)


def fallback_cell(anchor: tuple[int, int] | None, row_offset: int, col_offset: int) -> tuple[int, int] | None:
    if not anchor:
        return None
    return (anchor[0] + row_offset, anchor[1] + col_offset)


def resolve_cell(
    sheet: xw.Sheet,
    scan: SheetScan,
    labels: list[str],
    fallback_offset: tuple[int, int] | None,
) -> tuple[int, int] | None:
    from_label = value_cell_from_label(sheet, scan, labels)
    if from_label:
        return from_label
    if fallback_offset is None:
        return None
    return fallback_cell(scan.anchor, fallback_offset[0], fallback_offset[1])


def read_cell(sheet: xw.Sheet, cell: tuple[int, int] | None) -> Any:
    if not cell:
        return None
    return sheet.range(cell).value


def write_cell(sheet: xw.Sheet, cell: tuple[int, int] | None, value: Any) -> None:
    if not cell:
        return
    sheet.range(cell).value = value


def collect_numeric_rows(
    sheet: xw.Sheet,
    col: int,
    start_row: int,
    end_row: int,
) -> list[tuple[int, float]]:
    if start_row > end_row or col <= 0:
        return []

    values = sheet.range((start_row, col), (end_row, col)).value
    if not isinstance(values, list):
        values = [values]

    rows: list[tuple[int, float]] = []
    for idx, raw in enumerate(values):
        numeric = to_number(raw)
        if numeric is None:
            continue
        rows.append((start_row + idx, numeric))
    return rows


def collect_regression_pairs(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    start_row: int,
    end_row: int,
) -> list[tuple[int, float, float]]:
    if start_row > end_row or x_col <= 0 or y_col <= 0:
        return []

    x_values = sheet.range((start_row, x_col), (end_row, x_col)).value
    y_values = sheet.range((start_row, y_col), (end_row, y_col)).value

    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]

    pairs: list[tuple[int, float, float]] = []
    for idx, (raw_x, raw_y) in enumerate(zip(x_values, y_values)):
        x = to_number(raw_x)
        y = to_number(raw_y)
        if x is None or y is None:
            continue
        pairs.append((start_row + idx, x, y))
    return pairs


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
    except Exception as exc:
        print(f"warning: unable to close workbook cleanly ({exc})")


def get_sheet_case_insensitive(wb: xw.Book, target_name: str) -> xw.Sheet | None:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def extract_empirical_rows(sheet: xw.Sheet, meta: ModelMeta) -> list[dict[str, Any]]:
    scan = scan_sheet_once(sheet)
    if not scan.anchor:
        print(f"  empirical: skipped (no 'max' anchor found in sheet '{sheet.name}')")
        return []

    anchor_row, anchor_col = scan.anchor
    n_quarters = 10
    y_col = anchor_col - 7
    penetration_col = anchor_col - 11
    quarter_col = anchor_col - 12

    num_quarters_input_cell = resolve_cell(
        sheet,
        scan,
        ["num quarters used", "quarters used", "num quarters"],
        fallback_offset=(-6, 1),
    )
    last_quarter_cell = resolve_cell(
        sheet,
        scan,
        ["last quarter used", "last quarter"],
        fallback_offset=(-5, 1),
    )
    forecast_value_cell = resolve_cell(
        sheet,
        scan,
        ["estimated total sold", "forecast value", "total forecast"],
        fallback_offset=(-4, 1),
    )
    actual_value_cell = resolve_cell(
        sheet,
        scan,
        ["reported sales", "actual value", "actual sales"],
        fallback_offset=(-3, 1),
    )
    forecast_max_cell = resolve_cell(sheet, scan, ["max"], fallback_offset=(0, 1))
    forecast_min_cell = resolve_cell(sheet, scan, ["min"], fallback_offset=(1, 1))
    growth_rate_cell = resolve_cell(
        sheet,
        scan,
        ["growth rate", "growth %", "growth pct"],
        fallback_offset=(-2, 1),
    )
    sales_captured_cell = resolve_cell(
        sheet,
        scan,
        ["sales captured in db", "captured in db", "captured pct"],
        fallback_offset=(-1, 1),
    )

    data_start = max(1, anchor_row - 300)
    data_end = max(1, anchor_row - 1)
    penetration_rows = collect_numeric_rows(sheet, penetration_col, data_start, data_end)
    y_rows = {row: value for row, value in collect_numeric_rows(sheet, y_col, data_start, data_end)}
    quarter_rows = {row: sheet.range((row, quarter_col)).value for row, _ in penetration_rows}

    if not penetration_rows:
        print(f"  empirical: skipped (no numeric penetration data in col {penetration_col})")
        return []

    avg_pen_temp_cell = (anchor_row + 30, anchor_col + 6)
    rows: list[dict[str, Any]] = []
    max_n = min(n_quarters, len(penetration_rows))
    for num_quarters_used in range(1, max_n + 1):
        write_cell(sheet, num_quarters_input_cell, num_quarters_used)
        selected = penetration_rows[-num_quarters_used:]
        first_row = selected[0][0]
        last_row = selected[-1][0]

        avg_formula = f"=AVERAGE(R{first_row}C{penetration_col}:R{last_row}C{penetration_col})"
        sheet.range(avg_pen_temp_cell).formula2 = avg_formula
        sheet.book.app.calculate()

        avg_pen = read_cell(sheet, avg_pen_temp_cell)
        forecast_value = read_cell(sheet, forecast_value_cell)
        actual_value = read_cell(sheet, actual_value_cell)
        forecast_max = read_cell(sheet, forecast_max_cell)
        forecast_min = read_cell(sheet, forecast_min_cell)
        growth_rate = read_cell(sheet, growth_rate_cell)
        captured_pct = read_cell(sheet, sales_captured_cell)
        quarterly_sales = y_rows.get(last_row)
        reported_sales = actual_value
        last_quarter_used = quarter_rows.get(last_row)
        if last_quarter_used in (None, ""):
            last_quarter_used = read_cell(sheet, last_quarter_cell)

        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = max_num - min_num if max_num is not None and min_num is not None else None

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
                "sales_captured_in_db_pct": captured_pct,
                "source_file": meta.source_file,
            }
        )

    # Clear temporary formula cell before closing workbook (no save anyway).
    sheet.range(avg_pen_temp_cell).value = None
    return rows


def extract_regression_rows(sheet: xw.Sheet, meta: ModelMeta) -> list[dict[str, Any]]:
    scan = scan_sheet_once(sheet)
    if not scan.anchor:
        print(f"  regression: skipped (no 'max' anchor found in sheet '{sheet.name}')")
        return []

    anchor_row, anchor_col = scan.anchor
    n_quarters = 10
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_input_cell = resolve_cell(
        sheet,
        scan,
        ["num quarters used", "quarters used", "num quarters"],
        fallback_offset=(-6, 1),
    )
    forecast_total_cell = resolve_cell(
        sheet,
        scan,
        ["tot fcst w o sa", "tot fcst wo sa", "total forecast without sa", "tot fcst"],
        fallback_offset=(-4, 1),
    )
    actual_value_cell = resolve_cell(
        sheet,
        scan,
        ["actual value", "reported sales", "actual sales"],
        fallback_offset=None,
    )
    forecast_max_cell = resolve_cell(sheet, scan, ["max"], fallback_offset=(0, 1))
    forecast_min_cell = resolve_cell(sheet, scan, ["min"], fallback_offset=(1, 1))

    data_start = max(1, anchor_row - 300)
    data_end = max(1, anchor_row - 1)
    pairs = collect_regression_pairs(sheet, x_col, y_col, data_start, data_end)
    if not pairs:
        print(f"  regression: skipped (no numeric x/y data in cols x={x_col}, y={y_col})")
        return []

    intercept_temp_cell = (anchor_row + 32, anchor_col + 6)
    slope_temp_cell = (anchor_row + 33, anchor_col + 6)

    rows: list[dict[str, Any]] = []
    previous_key: tuple[Any, ...] | None = None
    max_n = min(n_quarters, len(pairs))
    for num_quarters_used in range(1, max_n + 1):
        write_cell(sheet, num_quarters_input_cell, num_quarters_used)
        selected = pairs[-num_quarters_used:]
        first_row = selected[0][0]
        last_row = selected[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        sheet.range(intercept_temp_cell).formula2 = intercept_formula
        sheet.range(slope_temp_cell).formula2 = slope_formula
        sheet.book.app.calculate()

        intercept = read_cell(sheet, intercept_temp_cell)
        slope = read_cell(sheet, slope_temp_cell)
        forecast_value = read_cell(sheet, forecast_total_cell)
        max_value = read_cell(sheet, forecast_max_cell)
        min_value = read_cell(sheet, forecast_min_cell)
        actual_value = read_cell(sheet, actual_value_cell)

        if forecast_value in (None, ""):
            x_anchor = to_number(sheet.range((anchor_row, x_col)).value)
            if x_anchor is None:
                x_anchor = selected[-1][1]
            intercept_num = to_number(intercept)
            slope_num = to_number(slope)
            if intercept_num is not None and slope_num is not None:
                forecast_value = intercept_num + slope_num * x_anchor

        max_num = to_number(max_value)
        min_num = to_number(min_value)
        range_width = max_num - min_num if max_num is not None and min_num is not None else None

        dedupe_key = (
            num_quarters_used,
            round(to_number(intercept) or 0.0, 10),
            round(to_number(slope) or 0.0, 10),
            round(to_number(forecast_value) or 0.0, 10),
            round(to_number(max_value) or 0.0, 10),
            round(to_number(min_value) or 0.0, 10),
        )
        if previous_key is not None and dedupe_key == previous_key:
            continue
        previous_key = dedupe_key

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
                "forecast_max": max_value,
                "forecast_min": min_value,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": meta.source_file,
            }
        )

    sheet.range(intercept_temp_cell).value = None
    sheet.range(slope_temp_cell).value = None
    return rows


def apply_worksheet_formatting(ws: Any, headers: list[str]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    empirical_ws = wb.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_ws.append([row.get(header) for header in EMPIRICAL_HEADERS])
    apply_worksheet_formatting(empirical_ws, EMPIRICAL_HEADERS)

    regression_ws = wb.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_ws.append([row.get(header) for header in REGRESSION_HEADERS])
    apply_worksheet_formatting(regression_ws, REGRESSION_HEADERS)

    wb.save(output_path)


def process_all_workbooks() -> None:
    source_input_dir = Path(input_dir)
    source_output_dir = Path(output_dir)
    if not source_input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {source_input_dir}")
    source_output_dir.mkdir(parents=True, exist_ok=True)

    output_path = find_unique_output_path(source_input_dir, source_output_dir)
    output_stem_prefix = f"{source_input_dir.resolve().name}_PARAM"

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(source_input_dir.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temp file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue
            if file_path.stem.startswith(output_stem_prefix):
                print(f"skipped file: {file_path.name} (output artifact)")
                continue

            print(f"processed file: {file_path.name}")
            meta = parse_filename_metadata(file_path)
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"  skipped model extraction in {file_path.name}: unable to open workbook ({exc})")
                continue
            try:
                empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
                regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")

                if empirical_sheet is None:
                    print(f"  empirical: skipped (sheet 'Empirical Model' not found)")
                else:
                    empirical_rows.extend(extract_empirical_rows(empirical_sheet, meta))

                if regression_sheet is None:
                    print(f"  regression: skipped (sheet 'Regression Model' not found)")
                else:
                    regression_rows.extend(extract_regression_rows(regression_sheet, meta))

                files_processed += 1
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_all_workbooks()
