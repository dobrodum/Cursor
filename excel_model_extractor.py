from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import re
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw


# Configure these paths for your environment.
input_dir = "./input"
output_dir = "./output"


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
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class UsedRangeData:
    values: list[list[Any]]
    start_row: int
    start_col: int
    row_count: int
    col_count: int


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            try:
                return float(cleaned[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def parse_file_meta(file_path: Path) -> FileMeta | None:
    name = file_path.stem
    # Example: MedMiner_Model - AORT - MidJan2026_Send
    match = re.search(
        r"model\s*-\s*([A-Za-z0-9]+)\s*-\s*((?:Early|Mid|Late)(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{4})",
        name,
        re.IGNORECASE,
    )
    if not match:
        return None

    ticker = match.group(1).upper()
    period_token = match.group(2)

    period_match = re.match(
        r"^(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})$",
        period_token,
        re.IGNORECASE,
    )
    if not period_match:
        return None

    phase = period_match.group(1).capitalize()
    month_abbrev = period_match.group(2).capitalize()
    year = int(period_match.group(3))

    day_by_phase = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_by_phase[phase]
    month_num = {
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
    }[month_abbrev]

    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def output_path_for_folder(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    folder_name = input_path.name
    base_name = f"{folder_name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_path / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def list_source_files(input_path: Path) -> tuple[list[Path], list[tuple[Path, str]]]:
    files: list[Path] = []
    skipped: list[tuple[Path, str]] = []

    if not input_path.exists():
        return files, [(input_path, "input directory does not exist")]

    for item in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
        if not item.is_file():
            continue
        if item.name.startswith("~"):
            skipped.append((item, "temporary file"))
            continue
        if item.suffix.lower() != ".xlsx":
            skipped.append((item, "not an .xlsx file"))
            continue
        files.append(item)

    return files, skipped


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
    except TypeError:
        try:
            workbook.api.Close(False)
        except Exception:
            try:
                workbook.close()
            except Exception:
                pass
    except Exception:
        try:
            workbook.api.Close(False)
        except Exception:
            pass


def get_used_range_data(sheet: xw.Sheet) -> UsedRangeData:
    used = sheet.used_range
    values = ensure_2d(used.value)
    row_count = len(values)
    col_count = max((len(row) for row in values), default=0)
    # Normalize row lengths for predictable indexing.
    normalized_values = [row + [None] * (col_count - len(row)) for row in values]
    return UsedRangeData(
        values=normalized_values,
        start_row=used.row,
        start_col=used.column,
        row_count=row_count,
        col_count=col_count,
    )


def matrix_value(used_data: UsedRangeData, abs_row: int, abs_col: int) -> Any:
    r = abs_row - used_data.start_row
    c = abs_col - used_data.start_col
    if r < 0 or c < 0:
        return None
    if r >= used_data.row_count or c >= used_data.col_count:
        return None
    return used_data.values[r][c]


def find_max_anchor(used_data: UsedRangeData) -> tuple[int, int] | None:
    candidates: list[tuple[int, int, int]] = []
    for r_index, row in enumerate(used_data.values):
        for c_index, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            abs_row = used_data.start_row + r_index
            abs_col = used_data.start_col + c_index
            neighbor_min = normalize_text(matrix_value(used_data, abs_row, abs_col + 1)) == "min"
            score = 0
            if neighbor_min:
                score += 10
            below_value = matrix_value(used_data, abs_row + 1, abs_col)
            if to_number(below_value) is not None:
                score += 5
            candidates.append((score, abs_row, abs_col))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def detect_columns(
    used_data: UsedRangeData,
    anchor_row: int,
    anchor_col: int,
    keyword_map: dict[str, list[str]],
    row_window: Iterable[int] = (-2, -1, 0, 1, 2),
) -> dict[str, int]:
    detected: dict[str, int] = {}
    candidate_rows = [anchor_row + delta for delta in row_window]

    for key, keywords in keyword_map.items():
        best: tuple[int, int] | None = None
        for row in candidate_rows:
            for col in range(used_data.start_col, used_data.start_col + used_data.col_count):
                text = normalize_text(matrix_value(used_data, row, col))
                if not text:
                    continue
                if any(keyword in text for keyword in keywords):
                    distance = abs(col - anchor_col) + (abs(row - anchor_row) * 2)
                    if best is None or distance < best[0]:
                        best = (distance, col)
        if best is not None:
            detected[key] = best[1]

    # Anchor-based defaults for speed and resilience.
    if "forecast_max_col" not in detected:
        detected["forecast_max_col"] = anchor_col
    if "forecast_min_col" not in detected:
        detected["forecast_min_col"] = anchor_col + 1

    return detected


def find_fallback_col(
    detected: dict[str, int],
    key: str,
    anchor_col: int,
    fallback_offset: int,
) -> int:
    return detected.get(key, anchor_col + fallback_offset)


def make_read_block(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    row_span_before: int = 12,
    row_span_after: int = 12,
    col_span_left: int = 20,
    col_span_right: int = 20,
) -> tuple[int, int, list[list[Any]]]:
    row_start = max(1, anchor_row - row_span_before)
    row_end = max(row_start, anchor_row + row_span_after)
    col_start = max(1, anchor_col - col_span_left)
    col_end = max(col_start, anchor_col + col_span_right)
    values = ensure_2d(sheet.range((row_start, col_start), (row_end, col_end)).value)
    width = max((len(r) for r in values), default=0)
    padded = [r + [None] * (width - len(r)) for r in values]
    return row_start, col_start, padded


def block_value(
    row_start: int,
    col_start: int,
    block: list[list[Any]],
    abs_row: int,
    abs_col: int,
) -> Any:
    r = abs_row - row_start
    c = abs_col - col_start
    if r < 0 or c < 0:
        return None
    if r >= len(block):
        return None
    if c >= len(block[r]):
        return None
    return block[r][c]


def process_empirical_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    meta: FileMeta,
    source_name: str,
) -> list[dict[str, Any]]:
    used_data = get_used_range_data(sheet)
    anchor = find_max_anchor(used_data)
    if anchor is None:
        print(f"Skipped empirical extraction for {source_name}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    column_keywords = {
        "forecast_value_col": ["estimated total sold", "est total sold", "tot fcst", "forecast total"],
        "actual_value_col": ["reported sales", "actual sales", "actual value", "actual"],
        "quarterly_sales_col": ["quarterly sales", "qtr sales"],
        "reported_sales_col": ["reported sales"],
        "growth_rate_col": ["growth rate", "growth %"],
        "sales_captured_col": ["captured in db", "sales captured"],
        "last_quarter_col": ["last quarter", "quarter used", "quarter", "qtr"],
        "penetration_col": ["penetration"],
        "num_quarters_col": ["num quarters", "quarters used", "n quarters"],
    }
    detected = detect_columns(used_data, anchor_row, anchor_col, column_keywords)

    forecast_value_col = find_fallback_col(detected, "forecast_value_col", anchor_col, -2)
    actual_value_col = find_fallback_col(detected, "actual_value_col", anchor_col, -3)
    forecast_max_col = find_fallback_col(detected, "forecast_max_col", anchor_col, 0)
    forecast_min_col = find_fallback_col(detected, "forecast_min_col", anchor_col, 1)
    quarterly_sales_col = find_fallback_col(detected, "quarterly_sales_col", anchor_col, -11)
    reported_sales_col = find_fallback_col(detected, "reported_sales_col", anchor_col, -7)
    growth_rate_col = find_fallback_col(detected, "growth_rate_col", anchor_col, -6)
    sales_captured_col = find_fallback_col(detected, "sales_captured_col", anchor_col, -5)
    last_quarter_col = find_fallback_col(detected, "last_quarter_col", anchor_col, -12)
    penetration_col = find_fallback_col(detected, "penetration_col", anchor_col, -10)
    num_quarters_col = find_fallback_col(detected, "num_quarters_col", anchor_col, -1)

    row_start, col_start, block = make_read_block(sheet, anchor_row, anchor_col)

    n_quarters = 10
    scratch_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 25)
    scratch_start = anchor_row + 1
    scratch_end = anchor_row + n_quarters
    formulas: list[list[str | None]] = []

    for n in range(1, n_quarters + 1):
        hist_start = anchor_row - n
        hist_end = anchor_row - 1
        if hist_start < 1 or penetration_col < 1:
            formulas.append([None])
            continue
        formulas.append([f"=AVERAGE(R{hist_start}C{penetration_col}:R{hist_end}C{penetration_col})"])

    sheet.range((scratch_start, scratch_col), (scratch_end, scratch_col)).formula2 = formulas
    workbook.app.calculate()
    avg_values = ensure_2d(sheet.range((scratch_start, scratch_col), (scratch_end, scratch_col)).value)
    sheet.range((scratch_start, scratch_col), (scratch_end, scratch_col)).clear_contents()

    rows: list[dict[str, Any]] = []
    for n in range(1, n_quarters + 1):
        data_row = anchor_row + n
        base_hist_row = anchor_row - n
        if base_hist_row < 1:
            continue

        inferred_num = to_number(block_value(row_start, col_start, block, data_row, num_quarters_col))
        num_quarters_used = int(inferred_num) if inferred_num is not None else n

        avg_penetration = avg_values[n - 1][0] if n - 1 < len(avg_values) else None
        forecast_value = block_value(row_start, col_start, block, data_row, forecast_value_col)
        actual_value = block_value(row_start, col_start, block, data_row, actual_value_col)
        if actual_value in (None, ""):
            actual_value = block_value(row_start, col_start, block, anchor_row - 1, reported_sales_col)

        forecast_max = block_value(row_start, col_start, block, data_row, forecast_max_col)
        forecast_min = block_value(row_start, col_start, block, data_row, forecast_min_col)
        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": block_value(row_start, col_start, block, base_hist_row, last_quarter_col),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": block_value(row_start, col_start, block, base_hist_row, quarterly_sales_col),
            "reported_sales": block_value(row_start, col_start, block, base_hist_row, reported_sales_col),
            "growth_rate_pct": block_value(row_start, col_start, block, base_hist_row, growth_rate_col),
            "sales_captured_in_db_pct": block_value(row_start, col_start, block, base_hist_row, sales_captured_col),
            "source_file": source_name,
        }
        rows.append(row)

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    meta: FileMeta,
    source_name: str,
) -> list[dict[str, Any]]:
    used_data = get_used_range_data(sheet)
    anchor = find_max_anchor(used_data)
    if anchor is None:
        print(f"Skipped regression extraction for {source_name}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    column_keywords = {
        "num_quarters_col": ["num quarters", "quarters used", "n quarters"],
        "forecast_value_col": ["tot fcst w/o sa", "tot fcst without sa", "tot fcst", "forecast total"],
        "actual_value_col": ["actual sales", "actual value", "actual"],
        "intercept_col": ["intercept"],
        "slope_col": ["slope"],
    }
    detected = detect_columns(used_data, anchor_row, anchor_col, column_keywords)

    num_quarters_col = find_fallback_col(detected, "num_quarters_col", anchor_col, -1)
    forecast_value_col = find_fallback_col(detected, "forecast_value_col", anchor_col, -2)
    actual_value_col = find_fallback_col(detected, "actual_value_col", anchor_col, -3)
    forecast_max_col = find_fallback_col(detected, "forecast_max_col", anchor_col, 0)
    forecast_min_col = find_fallback_col(detected, "forecast_min_col", anchor_col, 1)

    row_start, col_start, block = make_read_block(sheet, anchor_row, anchor_col)

    n_quarters = 10
    scratch_col_intercept = max(sheet.used_range.last_cell.column + 2, anchor_col + 25)
    scratch_col_slope = scratch_col_intercept + 1
    scratch_start = anchor_row + 1
    scratch_end = anchor_row + n_quarters

    intercept_formulas: list[list[str | None]] = []
    slope_formulas: list[list[str | None]] = []
    for n in range(1, n_quarters + 1):
        hist_start = anchor_row - n
        hist_end = anchor_row - 1
        if hist_start < 1 or x_col < 1 or y_col < 1:
            intercept_formulas.append([None])
            slope_formulas.append([None])
            continue

        y_range = f"R{hist_start}C{y_col}:R{hist_end}C{y_col}"
        x_range = f"R{hist_start}C{x_col}:R{hist_end}C{x_col}"
        intercept_formulas.append([f"=INTERCEPT({y_range},{x_range})"])
        slope_formulas.append([f"=SLOPE({y_range},{x_range})"])

    sheet.range((scratch_start, scratch_col_intercept), (scratch_end, scratch_col_intercept)).formula2 = intercept_formulas
    sheet.range((scratch_start, scratch_col_slope), (scratch_end, scratch_col_slope)).formula2 = slope_formulas
    workbook.app.calculate()
    intercept_values = ensure_2d(
        sheet.range((scratch_start, scratch_col_intercept), (scratch_end, scratch_col_intercept)).value
    )
    slope_values = ensure_2d(sheet.range((scratch_start, scratch_col_slope), (scratch_end, scratch_col_slope)).value)
    sheet.range((scratch_start, scratch_col_intercept), (scratch_end, scratch_col_slope)).clear_contents()

    rows: list[dict[str, Any]] = []
    for n in range(1, n_quarters + 1):
        data_row = anchor_row + n

        inferred_num = to_number(block_value(row_start, col_start, block, data_row, num_quarters_col))
        num_quarters_used = int(inferred_num) if inferred_num is not None else n
        forecast_value = block_value(row_start, col_start, block, data_row, forecast_value_col)
        actual_value = block_value(row_start, col_start, block, data_row, actual_value_col)
        forecast_max = block_value(row_start, col_start, block, data_row, forecast_max_col)
        forecast_min = block_value(row_start, col_start, block, data_row, forecast_min_col)
        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        intercept_value = intercept_values[n - 1][0] if n - 1 < len(intercept_values) else None
        slope_value = slope_values[n - 1][0] if n - 1 < len(slope_values) else None

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value not in (None, "") else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_name,
        }

        if rows:
            prev = rows[-1]
            duplicate_signature = (
                row["forecast_value"],
                row["forecast_max"],
                row["forecast_min"],
                row["intercept"],
                row["slope"],
            )
            prev_signature = (
                prev["forecast_value"],
                prev["forecast_max"],
                prev["forecast_min"],
                prev["intercept"],
                prev["slope"],
            )
            if duplicate_signature == prev_signature:
                continue

        rows.append(row)

    return rows


def rows_to_matrix(rows: list[dict[str, Any]], headers: list[str]) -> list[list[Any]]:
    return [[row.get(header) for header in headers] for row in rows]


def write_output_sheet(sheet: xw.Sheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    matrix = rows_to_matrix(rows, headers)
    data = [headers] + matrix
    end_row = len(data)
    end_col = len(headers)
    sheet.range((1, 1), (end_row, end_col)).value = data

    header_range = sheet.range((1, 1), (1, end_col))
    header_range.api.Font.Bold = True

    if end_row > 1:
        table_range = sheet.range((1, 1), (end_row, end_col))
        table_range.api.AutoFilter()

    sheet.activate()
    window = sheet.api.Application.ActiveWindow
    window.SplitRow = 1
    window.SplitColumn = 0
    window.FreezePanes = True

    # Reasonable fixed widths keep this fast and readable.
    width_overrides = {
        "model": 22,
        "ticker": 10,
        "model_period": 16,
        "model_date": 14,
        "method": 12,
        "parameter_name": 22,
        "source_file": 36,
    }
    default_width = 14
    for index, header in enumerate(headers, start=1):
        width = width_overrides.get(header, default_width)
        sheet.range((1, index)).column_width = width


def build_output_workbook(
    app: xw.App,
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    output_wb = app.books.add()
    try:
        while len(output_wb.sheets) < 2:
            output_wb.sheets.add()
        while len(output_wb.sheets) > 2:
            output_wb.sheets[-1].delete()

        empirical_sheet = output_wb.sheets[0]
        regression_sheet = output_wb.sheets[1]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet.name = "regression_candidates"

        write_output_sheet(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)
        write_output_sheet(regression_sheet, REGRESSION_HEADERS, regression_rows)

        output_wb.save(str(output_file))
    finally:
        safe_close_workbook(output_wb)


def main() -> None:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    source_files, skipped_files = list_source_files(in_path)
    for file_path, reason in skipped_files:
        print(f"Skipped file: {file_path.name} ({reason})")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files:
            meta = parse_file_meta(file_path)
            if meta is None:
                print(f"Skipped file: {file_path.name} (filename pattern not recognized)")
                continue

            print(f"Processing file: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                sheet_names = {sheet.name: sheet for sheet in workbook.sheets}

                empirical_sheet = sheet_names.get("Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name}: sheet 'Empirical Model' not found")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(workbook, empirical_sheet, meta, file_path.name)
                    )

                regression_sheet = sheet_names.get("Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression extraction for {file_path.name}: sheet 'Regression Model' not found")
                else:
                    regression_rows.extend(
                        process_regression_sheet(workbook, regression_sheet, meta, file_path.name)
                    )

                processed_count += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (open/process error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)

        output_file = output_path_for_folder(in_path, out_path)
        build_output_workbook(app, output_file, empirical_rows, regression_rows)

        print(f"Output path: {output_file}")
        print(f"Number of files processed: {processed_count}")
        print(f"Number of empirical rows: {len(empirical_rows)}")
        print(f"Number of regression rows: {len(regression_rows)}")
    finally:
        try:
            app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
