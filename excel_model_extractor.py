from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("input")
output_dir = Path("output")


N_QUARTERS = 10

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

TIMING_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
PERIOD_RE = re.compile(
    r"(?i)\b(early|mid|late)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(20\d{2})\b"
)


@dataclass(frozen=True)
class FileModelInfo:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    try:
        parsed = float(str(value).replace(",", "").strip())
        if math.isnan(parsed) or math.isinf(parsed):
            return None
        return parsed
    except Exception:
        return None


def clean_scalar(value: Any) -> Any:
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        if text.startswith("#"):  # Excel error strings like #DIV/0!
            return None
        return text
    return value


def as_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def parse_model_info(file_name: str) -> FileModelInfo:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()

    period_match = PERIOD_RE.search(stem)
    if period_match:
        timing = period_match.group(1).lower()
        month_txt = period_match.group(2).title()
        year_txt = period_match.group(3)
        month_num = MONTH_MAP[month_txt.lower()]
        day_num = TIMING_DAY_MAP[timing]
        model_period = f"{timing.title()}{month_txt}_{year_txt}"
        model_date = date(int(year_txt), month_num, day_num).isoformat()
    else:
        model_period = "UnknownPeriod"
        model_date = ""

    if not ticker:
        ticker_match = re.search(r"\b([A-Z]{1,8})\b", stem)
        ticker = ticker_match.group(1) if ticker_match else "UNKNOWN"

    model = f"{ticker}_{model_period}"
    return FileModelInfo(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for closer in (
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ):
        try:
            closer()
            return
        except Exception:
            continue


def set_formula2(target: xw.Range, formula_r1c1: str) -> None:
    try:
        target.formula2 = formula_r1c1
    except Exception:
        # Fallback for older Excel/engines that don't expose formula2.
        target.formula = formula_r1c1


def get_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def find_anchor(sheet: xw.Sheet, anchor_label: str = "max") -> tuple[int, int] | None:
    found = None
    try:
        found = sheet.api.Cells.Find(
            What=anchor_label,
            After=sheet.api.Cells(1, 1),
            LookAt=1,  # xlWhole
            SearchOrder=1,  # xlByRows
            SearchDirection=1,  # xlNext
            MatchCase=False,
        )
    except Exception:
        pass

    if found is not None:
        return int(found.Row), int(found.Column)

    used = sheet.used_range
    used_values = as_2d(used.value)
    top = used.row
    left = used.column
    target = normalize_text(anchor_label)
    for r_idx, row_vals in enumerate(used_values):
        for c_idx, val in enumerate(row_vals):
            if normalize_text(val) == target:
                return top + r_idx, left + c_idx
    return None


def build_text_index(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    row_radius: int = 60,
    col_radius: int = 22,
) -> dict[str, tuple[int, int]]:
    top = max(1, anchor_row - row_radius)
    bottom = max(top, anchor_row + row_radius)
    left = max(1, anchor_col - col_radius)
    right = max(left, anchor_col + col_radius)

    values = as_2d(sheet.range((top, left), (bottom, right)).value)
    index: dict[str, tuple[int, int]] = {}
    for r_idx, row_vals in enumerate(values):
        for c_idx, cell_value in enumerate(row_vals):
            if isinstance(cell_value, str):
                key = normalize_text(cell_value)
                if key and key not in index:
                    index[key] = (top + r_idx, left + c_idx)
    return index


def find_label(index: dict[str, tuple[int, int]], aliases: Iterable[str]) -> tuple[int, int] | None:
    normalized_aliases = [normalize_text(alias) for alias in aliases]
    for alias in normalized_aliases:
        if alias in index:
            return index[alias]
    for key, cell in index.items():
        for alias in normalized_aliases:
            if alias and alias in key:
                return cell
    return None


def find_numeric_value_cell(sheet: xw.Sheet, row: int, col: int) -> tuple[int, int]:
    candidate_offsets = [1, 2, -1, 3]
    for offset in candidate_offsets:
        if col + offset < 1:
            continue
        value = clean_scalar(sheet.range((row, col + offset)).value)
        if not isinstance(value, str):
            return row, col + offset
    return row, col + 1


def resolve_value_cell(
    sheet: xw.Sheet,
    text_index: dict[str, tuple[int, int]],
    aliases: Iterable[str],
    fallback: tuple[int, int],
) -> tuple[int, int]:
    label_cell = find_label(text_index, aliases)
    if label_cell:
        return find_numeric_value_cell(sheet, label_cell[0], label_cell[1])
    return fallback


def get_value(sheet: xw.Sheet, cell: tuple[int, int] | None) -> Any:
    if cell is None:
        return None
    row, col = cell
    if row < 1 or col < 1:
        return None
    return clean_scalar(sheet.range((row, col)).value)


def build_empirical_cells(
    sheet: xw.Sheet,
    text_index: dict[str, tuple[int, int]],
    anchor_row: int,
    anchor_col: int,
) -> dict[str, tuple[int, int]]:
    return {
        "num_quarters_input": resolve_value_cell(
            sheet,
            text_index,
            ["num quarters used", "quarters used", "num quarters", "n quarters"],
            (anchor_row - 6, anchor_col + 1),
        ),
        "last_quarter_used": resolve_value_cell(
            sheet,
            text_index,
            ["last quarter used", "last quarter"],
            (anchor_row - 5, anchor_col + 1),
        ),
        "forecast_value": resolve_value_cell(
            sheet,
            text_index,
            ["estimated total sold", "total sold estimate", "tot fcst", "forecast total"],
            (anchor_row - 2, anchor_col + 1),
        ),
        "actual_value": resolve_value_cell(
            sheet,
            text_index,
            ["reported sales", "actual sales"],
            (anchor_row + 2, anchor_col + 1),
        ),
        "forecast_max": (anchor_row, anchor_col + 1),
        "forecast_min": (anchor_row + 1, anchor_col + 1),
        "avg_penetration_value": resolve_value_cell(
            sheet,
            text_index,
            ["avg penetration", "average penetration", "avg pen"],
            (anchor_row + 3, anchor_col + 1),
        ),
        "quarterly_sales": resolve_value_cell(
            sheet,
            text_index,
            ["quarterly sales", "q sales", "quarter sales"],
            (anchor_row + 4, anchor_col + 1),
        ),
        "reported_sales": resolve_value_cell(
            sheet,
            text_index,
            ["reported sales", "sales reported"],
            (anchor_row + 5, anchor_col + 1),
        ),
        "growth_rate_pct": resolve_value_cell(
            sheet,
            text_index,
            ["growth rate", "growth rate pct", "growth %"],
            (anchor_row + 6, anchor_col + 1),
        ),
        "sales_captured_in_db_pct": resolve_value_cell(
            sheet,
            text_index,
            ["sales captured in db", "captured in db", "captured %"],
            (anchor_row + 7, anchor_col + 1),
        ),
    }


def build_regression_cells(
    sheet: xw.Sheet,
    text_index: dict[str, tuple[int, int]],
    anchor_row: int,
    anchor_col: int,
) -> dict[str, tuple[int, int]]:
    return {
        "num_quarters_input": resolve_value_cell(
            sheet,
            text_index,
            ["num quarters used", "quarters used", "num quarters", "n quarters"],
            (anchor_row - 6, anchor_col + 1),
        ),
        "forecast_value": resolve_value_cell(
            sheet,
            text_index,
            ["tot fcst w/o sa", "tot fcst wo sa", "total forecast without sa", "tot fcst"],
            (anchor_row - 2, anchor_col + 1),
        ),
        "actual_value": resolve_value_cell(
            sheet,
            text_index,
            ["actual sales", "reported sales"],
            (anchor_row + 2, anchor_col + 1),
        ),
        "forecast_max": (anchor_row, anchor_col + 1),
        "forecast_min": (anchor_row + 1, anchor_col + 1),
    }


def find_penetration_history_row(
    text_index: dict[str, tuple[int, int]],
    default_row: int,
) -> int:
    for alias in ("penetration", "sales captured in db", "captured in db"):
        cell = find_label(text_index, [alias])
        if cell:
            return cell[0]
    return default_row


def find_last_numeric_row(sheet: xw.Sheet, col: int, max_row: int, scan_window: int = 300) -> int | None:
    top = max(1, max_row - scan_window)
    column_values = as_2d(sheet.range((top, col), (max_row, col)).value)
    flat = [row[0] for row in column_values]
    for idx in range(len(flat) - 1, -1, -1):
        if to_float(flat[idx]) is not None:
            return top + idx
    return None


def rounded_for_compare(value: Any) -> Any:
    num = to_float(value)
    if num is not None:
        return round(num, 10)
    return clean_scalar(value)


def extract_empirical_rows(
    wb: xw.Book,
    model_info: FileModelInfo,
    source_file_name: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        print(f"SKIPPED {source_file_name}: missing sheet 'Empirical Model'")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIPPED {source_file_name}: no 'max' anchor in 'Empirical Model'")
        return []

    anchor_row, anchor_col = anchor
    text_index = build_text_index(sheet, anchor_row, anchor_col)
    cells = build_empirical_cells(sheet, text_index, anchor_row, anchor_col)
    penetration_row = find_penetration_history_row(text_index, cells["sales_captured_in_db_pct"][0])

    temp_formula_cell = sheet.range((anchor_row + 35, anchor_col + 8))
    rows: list[dict[str, Any]] = []

    for n_quarters in range(1, N_QUARTERS + 1):
        sheet.range(cells["num_quarters_input"]).value = n_quarters

        history_end_col = max(1, anchor_col - 1)
        history_start_col = max(1, history_end_col - n_quarters + 1)
        avg_formula = (
            f'=IFERROR(AVERAGE(R{penetration_row}C{history_start_col}:'
            f'R{penetration_row}C{history_end_col}),"")'
        )
        set_formula2(temp_formula_cell, avg_formula)

        wb.app.calculate()

        forecast_value = to_float(get_value(sheet, cells["forecast_value"]))
        actual_value = to_float(get_value(sheet, cells["actual_value"]))
        forecast_max = to_float(get_value(sheet, cells["forecast_max"]))
        forecast_min = to_float(get_value(sheet, cells["forecast_min"]))
        avg_penetration = to_float(temp_formula_cell.value)
        if avg_penetration is None:
            avg_penetration = to_float(get_value(sheet, cells["avg_penetration_value"]))

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                actual_value,
            )
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": model_info.model,
                "ticker": model_info.ticker,
                "model_period": model_info.model_period,
                "model_date": model_info.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": get_value(sheet, cells["last_quarter_used"]),
                "forecast_value": forecast_value,  # estimated total sold
                "actual_value": actual_value,  # reported sales
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": to_float(get_value(sheet, cells["quarterly_sales"])),
                "reported_sales": to_float(get_value(sheet, cells["reported_sales"])),
                "growth_rate_pct": to_float(get_value(sheet, cells["growth_rate_pct"])),
                "sales_captured_in_db_pct": to_float(get_value(sheet, cells["sales_captured_in_db_pct"])),
                "source_file": source_file_name,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    model_info: FileModelInfo,
    source_file_name: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        print(f"SKIPPED {source_file_name}: missing sheet 'Regression Model'")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIPPED {source_file_name}: no 'max' anchor in 'Regression Model'")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print(f"SKIPPED {source_file_name}: invalid regression anchor offsets")
        return []

    text_index = build_text_index(sheet, anchor_row, anchor_col)
    cells = build_regression_cells(sheet, text_index, anchor_row, anchor_col)

    y_end_row = find_last_numeric_row(sheet, y_col, anchor_row - 1)
    x_end_row = find_last_numeric_row(sheet, x_col, anchor_row - 1)
    if y_end_row is None or x_end_row is None:
        data_end_row = None
    else:
        data_end_row = min(y_end_row, x_end_row)

    if data_end_row is None:
        print(f"SKIPPED {source_file_name}: no regression data found")
        return []

    intercept_cell = sheet.range((anchor_row + 35, anchor_col + 8))
    slope_cell = sheet.range((anchor_row + 36, anchor_col + 8))

    rows: list[dict[str, Any]] = []
    prev_signature: tuple[Any, ...] | None = None

    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = data_end_row - n_quarters + 1
        if start_row < 1:
            break

        sheet.range(cells["num_quarters_input"]).value = n_quarters

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)

        wb.app.calculate()

        forecast_value = to_float(get_value(sheet, cells["forecast_value"]))
        actual_value = to_float(get_value(sheet, cells["actual_value"]))
        forecast_max = to_float(get_value(sheet, cells["forecast_max"]))
        forecast_min = to_float(get_value(sheet, cells["forecast_min"]))
        intercept_val = to_float(intercept_cell.value)
        slope_val = to_float(slope_cell.value)

        if all(
            value is None
            for value in (forecast_value, intercept_val, slope_val, forecast_max, forecast_min)
        ):
            continue

        signature = (
            rounded_for_compare(forecast_value),
            rounded_for_compare(forecast_max),
            rounded_for_compare(forecast_min),
            rounded_for_compare(intercept_val),
            rounded_for_compare(slope_val),
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": model_info.model,
                "ticker": model_info.ticker,
                "model_period": model_info.model_period,
                "model_date": model_info.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,  # TOT FCST w/o SA
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_val,
                "slope": slope_val,
                "source_file": source_file_name,
            }
        )

    return rows


def write_sheet(ws, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    max_row = max(ws.max_row, 1)
    max_col = max(ws.max_column, 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb_out = Workbook()
    ws_emp = wb_out.active
    ws_emp.title = "empirical_candidates"
    write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = wb_out.create_sheet("regression_candidates")
    write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb_out.save(output_path)


def collect_source_files(source_dir: Path) -> list[Path]:
    if not source_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {source_dir}")
    if not source_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {source_dir}")

    sources: list[Path] = []
    for entry in sorted(source_dir.iterdir()):
        if not entry.is_file():
            print(f"SKIPPED {entry.name}: not a file")
            continue
        if entry.name.startswith("~"):
            print(f"SKIPPED {entry.name}: temp file")
            continue
        if entry.suffix.lower() != ".xlsx":
            print(f"SKIPPED {entry.name}: not an .xlsx file")
            continue
        sources.append(entry)
    return sources


def process_all_files(source_files: list[Path]) -> tuple[int, list[dict[str, Any]], list[dict[str, Any]]]:
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"PROCESSING {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                model_info = parse_model_info(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, model_info, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, model_info, file_path.name))
                processed_count += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    return processed_count, empirical_rows, regression_rows


def main() -> None:
    source_files = collect_source_files(input_dir)
    output_path = get_output_path(input_dir, output_dir)

    processed_files, empirical_rows, regression_rows = process_all_files(source_files)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT {output_path}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    main()
