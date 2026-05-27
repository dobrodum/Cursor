#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "xlwings is required. Install with: pip install xlwings openpyxl"
    ) from exc


# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("./input")
output_dir = Path("./output")


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
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

PHASE_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_ABBR = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def normalize_header(value: Any) -> str:
    base = normalize_text(value)
    base = base.replace("%", " pct ")
    return re.sub(r"[^a-z0-9]+", " ", base).strip()


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("%", "")
    try:
        return float(text)
    except ValueError:
        return None


def values_to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    filename_match = re.match(
        r"^.+?\s*-\s*([A-Za-z0-9]+)\s*-\s*([A-Za-z]+\d{4})(?:_.+)?$",
        stem,
    )
    if not filename_match:
        raise ValueError("filename does not match expected pattern")

    ticker, period_token = filename_match.groups()
    token_match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token)
    if not token_match:
        raise ValueError(f"cannot parse period token '{period_token}'")

    phase, month_token, year_text = token_match.groups()
    month_num = MONTH_MAP.get(month_token.lower()) or MONTH_MAP.get(
        month_token[:3].lower()
    )
    if month_num is None:
        raise ValueError(f"unknown month '{month_token}'")

    year = int(year_text)
    day = PHASE_DAY_MAP[phase]
    model_period = f"{phase}{MONTH_ABBR[month_num]}_{year}"
    model_date = f"{year:04d}-{month_num:02d}-{day:02d}"
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def output_path_for_run(input_folder: Path, out_dir: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    suffix_idx = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{suffix_idx}.xlsx"
        suffix_idx += 1
    return candidate


def safe_close_workbook(wb: xw.Book) -> None:
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
    except TypeError:
        pass
    except Exception:
        pass

    wb.close()


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int]:
    used = sheet.used_range
    used_values = values_to_2d(used.value)
    base_row = used.row
    base_col = used.column

    for r_idx, row_values in enumerate(used_values):
        for c_idx, raw_value in enumerate(row_values):
            if normalize_text(raw_value) == "max":
                return base_row + r_idx, base_col + c_idx

    raise ValueError("could not find 'max' anchor")


def collect_header_offsets(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int
) -> dict[str, int]:
    used = sheet.used_range
    start_col = used.column
    end_col = start_col + used.columns.count - 1
    header_values = sheet.range((anchor_row, start_col), (anchor_row, end_col)).value
    if not isinstance(header_values, list):
        header_values = [header_values]

    offsets: dict[str, int] = {}
    for idx, value in enumerate(header_values):
        col = start_col + idx
        normalized = normalize_header(value)
        if normalized:
            offsets[normalized] = col - anchor_col
    return offsets


def find_offset_by_keywords(
    offsets: dict[str, int], include_all: tuple[str, ...], include_any: tuple[str, ...] = ()
) -> int | None:
    for header, offset in offsets.items():
        if all(token in header for token in include_all):
            if include_any and not any(token in header for token in include_any):
                continue
            return offset
    return None


def offset_or_default(value: int | None, default: int) -> int:
    return default if value is None else value


def get_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    return sheet.cells(row, col).value


def numeric_row_indices(sheet: xw.Sheet, col: int, end_row: int) -> list[int]:
    used = sheet.used_range
    start_row = used.row
    if end_row < start_row:
        return []

    col_values = sheet.range((start_row, col), (end_row, col)).value
    if not isinstance(col_values, list):
        col_values = [col_values]
    rows: list[int] = []
    for idx, value in enumerate(col_values):
        if to_float(value) is not None:
            rows.append(start_row + idx)
    return rows


def paired_numeric_rows(sheet: xw.Sheet, x_col: int, y_col: int, end_row: int) -> list[int]:
    used = sheet.used_range
    start_row = used.row
    if end_row < start_row:
        return []

    x_values = sheet.range((start_row, x_col), (end_row, x_col)).value
    y_values = sheet.range((start_row, y_col), (end_row, y_col)).value
    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]

    rows: list[int] = []
    for idx, (x_val, y_val) in enumerate(zip(x_values, y_values)):
        if to_float(x_val) is not None and to_float(y_val) is not None:
            rows.append(start_row + idx)
    return rows


def empirical_offsets(anchor_col: int, header_offsets: dict[str, int]) -> dict[str, int]:
    defaults = {
        "num_quarters_used": -9,
        "last_quarter_used": -8,
        "avg_penetration_pct": -6,
        "forecast_value": -4,
        "actual_value": -3,
        "forecast_min": 1,
        "quarterly_sales": -11,
        "reported_sales": -3,
        "growth_rate_pct": -5,
        "sales_captured_in_db_pct": -2,
    }

    resolved: dict[str, int] = {}
    resolved["num_quarters_used"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("num",), ("quarter", "qtr")),
        defaults["num_quarters_used"],
    )
    resolved["last_quarter_used"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("last", "quarter")),
        defaults["last_quarter_used"],
    )
    resolved["avg_penetration_pct"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("avg", "penetration")),
        defaults["avg_penetration_pct"],
    )
    resolved["forecast_value"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("estimated", "total", "sold"))
        or find_offset_by_keywords(header_offsets, ("tot", "fcst"), ("without", "w o", "w/o")),
        defaults["forecast_value"],
    )
    resolved["actual_value"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("reported", "sales")),
        defaults["actual_value"],
    )
    resolved["forecast_max"] = 0
    resolved["forecast_min"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("min",)),
        defaults["forecast_min"],
    )
    resolved["quarterly_sales"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("quarterly", "sales"))
        or find_offset_by_keywords(header_offsets, ("qtr", "sales")),
        defaults["quarterly_sales"],
    )
    resolved["reported_sales"] = resolved["actual_value"]
    resolved["growth_rate_pct"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("growth", "rate")),
        defaults["growth_rate_pct"],
    )
    resolved["sales_captured_in_db_pct"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("sales", "captured"), ("db",)),
        defaults["sales_captured_in_db_pct"],
    )
    resolved["penetration_source_col"] = anchor_col - 11
    return resolved


def regression_offsets(header_offsets: dict[str, int]) -> dict[str, int]:
    defaults = {
        "num_quarters_used": -9,
        "forecast_value": -4,
        "forecast_min": 1,
        "actual_value": -3,
    }

    resolved: dict[str, int] = {}
    resolved["num_quarters_used"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("num",), ("quarter", "qtr")),
        defaults["num_quarters_used"],
    )
    resolved["forecast_value"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("tot", "fcst"), ("without", "w o", "w/o", "sa")),
        defaults["forecast_value"],
    )
    resolved["forecast_max"] = 0
    resolved["forecast_min"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("min",)),
        defaults["forecast_min"],
    )
    resolved["actual_value"] = offset_or_default(
        find_offset_by_keywords(header_offsets, ("actual", "value"))
        or find_offset_by_keywords(header_offsets, ("reported", "sales")),
        defaults["actual_value"],
    )
    intercept_offset = find_offset_by_keywords(header_offsets, ("intercept",))
    slope_offset = find_offset_by_keywords(header_offsets, ("slope",))
    if intercept_offset is not None:
        resolved["intercept"] = intercept_offset
    if slope_offset is not None:
        resolved["slope"] = slope_offset
    return resolved


def process_empirical_sheet(
    wb: xw.Book, meta: FileLabel, source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)
    headers = collect_header_offsets(sheet, anchor_row, anchor_col)
    offsets = empirical_offsets(anchor_col, headers)

    avg_col = anchor_col + offsets["avg_penetration_pct"]
    source_col = offsets["penetration_source_col"]
    history_rows = numeric_row_indices(sheet, source_col, anchor_row - 1)

    formula_rows: list[int] = []
    if history_rows:
        for n in range(1, N_QUARTERS + 1):
            if len(history_rows) < n:
                continue
            row = anchor_row + n
            start_row = history_rows[-n]
            end_row = history_rows[-1]
            formula = f"=AVERAGE(R{start_row}C{source_col}:R{end_row}C{source_col})"
            sheet.cells(row, avg_col).formula2 = formula
            formula_rows.append(row)
        if formula_rows:
            wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        num_quarters = get_cell(sheet, row, anchor_col + offsets["num_quarters_used"]) or n
        last_quarter = get_cell(sheet, row, anchor_col + offsets["last_quarter_used"])
        avg_penetration = get_cell(sheet, row, avg_col)
        forecast_value = get_cell(sheet, row, anchor_col + offsets["forecast_value"])
        actual_value = get_cell(sheet, row, anchor_col + offsets["actual_value"])
        forecast_max = get_cell(sheet, row, anchor_col + offsets["forecast_max"])
        forecast_min = get_cell(sheet, row, anchor_col + offsets["forecast_min"])
        quarterly_sales = get_cell(sheet, row, anchor_col + offsets["quarterly_sales"])
        reported_sales = get_cell(sheet, row, anchor_col + offsets["reported_sales"])
        growth_rate = get_cell(sheet, row, anchor_col + offsets["growth_rate_pct"])
        sales_captured = get_cell(
            sheet, row, anchor_col + offsets["sales_captured_in_db_pct"]
        )

        if all(
            value in (None, "")
            for value in (
                avg_penetration,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
            )
        ):
            continue

        fmax_num = to_float(forecast_max)
        fmin_num = to_float(forecast_min)
        range_width = (fmax_num - fmin_num) if fmax_num is not None and fmin_num is not None else None

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
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
        )
    return rows


def rows_match_for_duplicate(a: dict[str, Any], b: dict[str, Any]) -> bool:
    comparable_keys = (
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    )
    for key in comparable_keys:
        va = a.get(key)
        vb = b.get(key)
        fa = to_float(va)
        fb = to_float(vb)
        if fa is not None or fb is not None:
            if fa is None or fb is None:
                return False
            if round(fa, 10) != round(fb, 10):
                return False
        else:
            if (va or "") != (vb or ""):
                return False
    return True


def process_regression_sheet(
    wb: xw.Book, meta: FileLabel, source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)
    headers = collect_header_offsets(sheet, anchor_row, anchor_col)
    offsets = regression_offsets(headers)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_rows = paired_numeric_rows(sheet, x_col=x_col, y_col=y_col, end_row=anchor_row - 1)

    used = sheet.used_range
    temp_col_base = used.column + used.columns.count + 2
    intercept_col = anchor_col + offsets.get("intercept", temp_col_base - anchor_col)
    slope_col = anchor_col + offsets.get("slope", (temp_col_base + 1) - anchor_col)

    formula_rows: list[int] = []
    if history_rows:
        for n in range(1, N_QUARTERS + 1):
            if len(history_rows) < n:
                continue
            row = anchor_row + n
            start_row = history_rows[-n]
            end_row = history_rows[-1]
            intercept_formula = (
                f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            )
            sheet.cells(row, intercept_col).formula2 = intercept_formula
            sheet.cells(row, slope_col).formula2 = slope_formula
            formula_rows.append(row)
        if formula_rows:
            wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        num_quarters = get_cell(sheet, row, anchor_col + offsets["num_quarters_used"]) or n
        forecast_value = get_cell(sheet, row, anchor_col + offsets["forecast_value"])
        actual_value = get_cell(sheet, row, anchor_col + offsets["actual_value"]) or ""
        forecast_max = get_cell(sheet, row, anchor_col + offsets["forecast_max"])
        forecast_min = get_cell(sheet, row, anchor_col + offsets["forecast_min"])
        intercept_value = get_cell(sheet, row, intercept_col)
        slope_value = get_cell(sheet, row, slope_col)

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

        fmax_num = to_float(forecast_max)
        fmin_num = to_float(forecast_min)
        range_width = (fmax_num - fmin_num) if fmax_num is not None and fmin_num is not None else None

        row_dict = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        if n == N_QUARTERS and rows and rows_match_for_duplicate(row_dict, rows[-1]):
            continue
        rows.append(row_dict)

    return rows


def write_sheet(ws, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row_data in rows:
        ws.append([(row_data.get(col) if row_data.get(col) is not None else "") for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def main() -> int:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input directory not found: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_path_for_run(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue
            if file_path.name == output_path.name and file_path.parent.resolve() == output_path.parent.resolve():
                print(f"skipped: {file_path.name} (output target)")
                continue

            try:
                label = parse_file_label(file_path)
            except Exception as parse_error:
                print(f"skipped: {file_path.name} ({parse_error})")
                continue

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    process_empirical_sheet(wb=wb, meta=label, source_file=file_path.name)
                )
                regression_rows.extend(
                    process_regression_sheet(wb=wb, meta=label, source_file=file_path.name)
                )
                files_processed += 1
                print(f"processed: {file_path.name}")
            except Exception as workbook_error:
                print(f"skipped: {file_path.name} ({workbook_error})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {files_processed}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
