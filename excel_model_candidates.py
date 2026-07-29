from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User inputs
# ---------------------------------------------------------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


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


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_month(month_text: str) -> int | None:
    clean = month_text.strip()
    lower_map = {
        name.lower(): idx
        for idx, name in enumerate(calendar.month_name)
        if idx and name
    }
    lower_map.update(
        {
            abbr.lower(): idx
            for idx, abbr in enumerate(calendar.month_abbr)
            if idx and abbr
        }
    )

    if clean.lower() in lower_map:
        return lower_map[clean.lower()]

    short = clean[:3].lower()
    return lower_map.get(short)


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0].strip()

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )
    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )

    if not period_match:
        model_period = "unknown_period"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileLabel(
            model=model,
            ticker=ticker,
            model_period=model_period,
            model_date=model_date,
        )

    phase = period_match.group(1).title()
    month_name = period_match.group(2)
    year = int(period_match.group(3))
    month = parse_month(month_name)

    if month is None:
        model_period = "unknown_period"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileLabel(
            model=model,
            ticker=ticker,
            model_period=model_period,
            model_date=model_date,
        )

    month_abbr = calendar.month_abbr[month]
    model_period = f"{phase}{month_abbr}_{year}"
    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_map[phase]
    model_date = date(year, month, model_day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1].strip()
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def number_or_original(value: Any) -> Any:
    numeric = to_float(value)
    return numeric if numeric is not None else value


def numeric_diff(left: Any, right: Any) -> float | None:
    left_n = to_float(left)
    right_n = to_float(right)
    if left_n is None or right_n is None:
        return None
    return left_n - right_n


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []

    if isinstance(values, (list, tuple)):
        if not values:
            return []
        if isinstance(values[0], (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]

    return [[values]]


def safe_close_without_save(wb: Any) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
    except Exception:
        pass

    wb.close()


def set_formula2(target: Any, formula_value: Any) -> None:
    try:
        target.formula2 = formula_value
    except Exception:
        target.formula = formula_value


def find_anchor_cell(sheet: Any, anchor_text: str = "max") -> tuple[int, int] | None:
    used_range = sheet.used_range
    values = normalize_2d(used_range.value)
    if not values:
        return None

    base_row = used_range.row
    base_col = used_range.column
    needle = anchor_text.strip().lower()

    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == needle:
                return base_row + row_idx, base_col + col_idx
    return None


def get_sheet_if_exists(wb: Any, sheet_name: str) -> Any | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM.xlsx"
    output_path = output_folder / base_name
    if not output_path.exists():
        return output_path

    suffix_index = 1
    while True:
        candidate = output_folder / f"{input_folder.name}_PARAM.{suffix_index}.xlsx"
        if not candidate.exists():
            return candidate
        suffix_index += 1


def as_row_values(values_2d: list[list[Any]], row_idx: int) -> list[Any]:
    if row_idx < 0 or row_idx >= len(values_2d):
        return []
    return values_2d[row_idx]


def process_empirical_sheet(
    sheet: Any,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if anchor is None:
        print(f"Skipped empirical for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    row_start = anchor_row + 1
    row_end = row_start + n_quarters - 1

    offsets = {
        "num_quarters_used": -8,
        "last_quarter_used": -7,
        "avg_penetration_pct": -6,
        "quarterly_sales": -5,
        "reported_sales": -4,
        "growth_rate_pct": -3,
        "sales_captured_in_db_pct": -2,
        "forecast_value": -1,
        "forecast_max": 0,
        "forecast_min": 1,
    }

    num_col = anchor_col + offsets["num_quarters_used"]
    avg_pen_col = anchor_col + offsets["avg_penetration_pct"]
    penetration_source_col = anchor_col + offsets["sales_captured_in_db_pct"]
    history_end_row = anchor_row - 1

    if history_end_row >= 1:
        num_quarter_values = [[quarter_idx] for quarter_idx in range(1, n_quarters + 1)]
        sheet.range((row_start, num_col), (row_end, num_col)).value = num_quarter_values

        avg_formulas: list[list[str]] = []
        for n_used in range(1, n_quarters + 1):
            history_start_row = max(1, history_end_row - n_used + 1)
            avg_formula = (
                f"=AVERAGE(R{history_start_row}C{penetration_source_col}:"
                f"R{history_end_row}C{penetration_source_col})"
            )
            avg_formulas.append([avg_formula])
        set_formula2(sheet.range((row_start, avg_pen_col), (row_end, avg_pen_col)), avg_formulas)
        sheet.book.app.calculate()

    min_offset = min(offsets.values())
    max_offset = max(offsets.values())
    read_start_col = anchor_col + min_offset
    read_end_col = anchor_col + max_offset

    value_grid = normalize_2d(
        sheet.range((row_start, read_start_col), (row_end, read_end_col)).value
    )
    if not value_grid:
        return []

    offset_base = -min_offset
    result_rows: list[dict[str, Any]] = []
    for row_idx in range(len(value_grid)):
        row_values = as_row_values(value_grid, row_idx)
        if not row_values:
            continue

        def col_value(field_name: str) -> Any:
            col_idx = offsets[field_name] + offset_base
            return row_values[col_idx] if 0 <= col_idx < len(row_values) else None

        num_quarters_used = col_value("num_quarters_used")
        last_quarter_used = col_value("last_quarter_used")
        avg_penetration_pct = col_value("avg_penetration_pct")
        quarterly_sales = col_value("quarterly_sales")
        reported_sales = col_value("reported_sales")
        growth_rate_pct = col_value("growth_rate_pct")
        sales_captured_pct = col_value("sales_captured_in_db_pct")
        forecast_value = col_value("forecast_value")
        forecast_max = col_value("forecast_max")
        forecast_min = col_value("forecast_min")
        range_width = numeric_diff(forecast_max, forecast_min)

        if all(
            value is None
            for value in (
                num_quarters_used,
                avg_penetration_pct,
                forecast_value,
                forecast_max,
                forecast_min,
            )
        ):
            continue

        result_rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": number_or_original(avg_penetration_pct),
                "num_quarters_used": number_or_original(num_quarters_used),
                "last_quarter_used": last_quarter_used,
                "forecast_value": number_or_original(forecast_value),
                "actual_value": number_or_original(reported_sales),
                "forecast_max": number_or_original(forecast_max),
                "forecast_min": number_or_original(forecast_min),
                "range_width": range_width,
                "avg_penetration_pct": number_or_original(avg_penetration_pct),
                "quarterly_sales": number_or_original(quarterly_sales),
                "reported_sales": number_or_original(reported_sales),
                "growth_rate_pct": number_or_original(growth_rate_pct),
                "sales_captured_in_db_pct": number_or_original(sales_captured_pct),
                "source_file": source_file,
            }
        )

    return result_rows


def signatures_match(previous: Iterable[Any], current: Iterable[Any], tol: float = 1e-10) -> bool:
    for prev_val, curr_val in zip(previous, current):
        prev_num = to_float(prev_val)
        curr_num = to_float(curr_val)
        if prev_num is not None and curr_num is not None:
            if abs(prev_num - curr_num) > tol:
                return False
            continue
        if prev_val != curr_val:
            return False
    return True


def process_regression_sheet(
    sheet: Any,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if anchor is None:
        print(f"Skipped regression for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    row_start = anchor_row + 1
    row_end = row_start + n_quarters - 1

    offsets = {
        "num_quarters_used": -4,
        "intercept": -3,
        "slope": -2,
        "forecast_value": -1,
        "forecast_max": 0,
        "forecast_min": 1,
    }

    num_col = anchor_col + offsets["num_quarters_used"]
    intercept_col = anchor_col + offsets["intercept"]
    slope_col = anchor_col + offsets["slope"]

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = anchor_row - 1

    if history_end_row >= 1:
        num_quarter_values = [[quarter_idx] for quarter_idx in range(1, n_quarters + 1)]
        sheet.range((row_start, num_col), (row_end, num_col)).value = num_quarter_values

        intercept_formulas: list[list[str]] = []
        slope_formulas: list[list[str]] = []
        for n_used in range(1, n_quarters + 1):
            history_start_row = max(1, history_end_row - n_used + 1)
            intercept_formula = (
                f"=INTERCEPT(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
                f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
                f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
            )
            intercept_formulas.append([intercept_formula])
            slope_formulas.append([slope_formula])

        set_formula2(
            sheet.range((row_start, intercept_col), (row_end, intercept_col)),
            intercept_formulas,
        )
        set_formula2(
            sheet.range((row_start, slope_col), (row_end, slope_col)),
            slope_formulas,
        )
        sheet.book.app.calculate()

    min_offset = min(offsets.values())
    max_offset = max(offsets.values())
    read_start_col = anchor_col + min_offset
    read_end_col = anchor_col + max_offset
    value_grid = normalize_2d(
        sheet.range((row_start, read_start_col), (row_end, read_end_col)).value
    )
    if not value_grid:
        return []

    offset_base = -min_offset
    result_rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for row_values in value_grid:
        def col_value(field_name: str) -> Any:
            col_idx = offsets[field_name] + offset_base
            return row_values[col_idx] if 0 <= col_idx < len(row_values) else None

        num_quarters_used = col_value("num_quarters_used")
        intercept = col_value("intercept")
        slope = col_value("slope")
        forecast_value = col_value("forecast_value")
        forecast_max = col_value("forecast_max")
        forecast_min = col_value("forecast_min")
        range_width = numeric_diff(forecast_max, forecast_min)

        if all(
            value is None
            for value in (
                num_quarters_used,
                intercept,
                slope,
                forecast_value,
                forecast_max,
                forecast_min,
            )
        ):
            continue

        current_signature = (
            number_or_original(num_quarters_used),
            number_or_original(intercept),
            number_or_original(slope),
            number_or_original(forecast_value),
            number_or_original(forecast_max),
            number_or_original(forecast_min),
        )
        if previous_signature is not None and signatures_match(previous_signature, current_signature):
            continue

        result_rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": number_or_original(num_quarters_used),
                "num_quarters_used": number_or_original(num_quarters_used),
                "forecast_value": number_or_original(forecast_value),
                "actual_value": None,
                "forecast_max": number_or_original(forecast_max),
                "forecast_min": number_or_original(forecast_min),
                "range_width": range_width,
                "intercept": number_or_original(intercept),
                "slope": number_or_original(slope),
                "source_file": source_file,
            }
        )
        previous_signature = current_signature

    return result_rows


def set_sheet_formatting(ws: Any) -> None:
    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            text_value = str(cell_value)
            if len(text_value) > max_len:
                max_len = len(text_value)
        width = min(max(12, max_len + 2), 48)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_ws.append([row.get(header) for header in EMPIRICAL_HEADERS])
    set_sheet_formatting(empirical_ws)

    regression_ws = workbook.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_ws.append([row.get(header) for header in REGRESSION_HEADERS])
    set_sheet_formatting(regression_ws)

    workbook.save(output_path)


def iter_candidate_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            print(f"Skipped: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped: {file_path.name} (not an .xlsx file)")
            continue
        if re.search(r"_PARAM(?:\.\d+)?\.xlsx$", file_path.name, flags=re.IGNORECASE):
            print(f"Skipped: {file_path.name} (output workbook pattern)")
            continue
        yield file_path


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir not found: {in_dir}")
    if not in_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a folder: {in_dir}")

    output_path = build_output_path(in_dir, out_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_candidate_files(in_dir):
            print(f"Processing: {file_path.name}")
            file_label = parse_file_label(file_path.name)
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                processed_files += 1

                empirical_sheet = get_sheet_if_exists(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical for {file_path.name}: sheet missing")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(empirical_sheet, file_label, file_path.name)
                    )

                regression_sheet = get_sheet_if_exists(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression for {file_path.name}: sheet missing")
                else:
                    regression_rows.extend(
                        process_regression_sheet(regression_sheet, file_label, file_path.name)
                    )
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_without_save(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
