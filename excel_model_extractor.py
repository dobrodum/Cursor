#!/usr/bin/env python3
from __future__ import annotations

import calendar
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

# -------- User-configurable paths --------
input_dir = Path("input")
output_dir = Path("output")


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
PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def log(message: str) -> None:
    print(message, flush=True)


def as_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    try:
        parsed = float(str(value).strip().replace(",", ""))
        if math.isnan(parsed):
            return None
        return parsed
    except (TypeError, ValueError):
        return None


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"[^a-z0-9%/ ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


@dataclass
class SheetCache:
    sheet: xw.main.Sheet
    start_row: int
    start_col: int
    values: list[list[Any]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        if not self.values:
            return 0
        return max(len(row) for row in self.values)

    @property
    def end_row(self) -> int:
        return self.start_row + self.row_count - 1

    @property
    def end_col(self) -> int:
        return self.start_col + self.col_count - 1

    @classmethod
    def build(cls, sheet: xw.main.Sheet) -> "SheetCache":
        used = sheet.used_range
        values = to_2d(used.value)
        return cls(sheet=sheet, start_row=used.row, start_col=used.column, values=values)

    def in_bounds(self, row: int, col: int) -> bool:
        return self.start_row <= row <= self.end_row and self.start_col <= col <= self.end_col

    def value_at(self, row: int, col: int) -> Any:
        if not self.in_bounds(row, col):
            return None
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        row_values = self.values[r_idx]
        if c_idx >= len(row_values):
            return None
        return row_values[c_idx]


def parse_filename_metadata(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem

    ticker = ""
    model_period = ""
    model_date = ""

    main_match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3,9}\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if main_match:
        ticker = main_match.group("ticker").upper()
        period_token = main_match.group("period")
    else:
        parts = [part.strip() for part in stem.split("-")]
        if len(parts) >= 3:
            ticker = parts[1].upper()
            period_token = re.sub(r"[^A-Za-z0-9]", "", parts[2])
        else:
            period_token = ""

    token_match = re.match(
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4})",
        period_token or "",
        flags=re.IGNORECASE,
    )
    if token_match:
        phase = token_match.group("phase").capitalize()
        month_text = token_match.group("month")
        year = int(token_match.group("year"))

        month_num = month_to_number(month_text)
        if month_num is not None:
            month_abbrev = calendar.month_abbr[month_num]
            model_period = f"{phase}{month_abbrev}_{year}"
            day = PERIOD_DAY_MAP[phase.lower()]
            model_date = date(year, month_num, day).isoformat()

    if not ticker:
        ticker = "UNKNOWN"
    if not model_period:
        model_period = "UNKNOWN_PERIOD"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def month_to_number(month_text: str) -> int | None:
    month_text = month_text.strip()
    if not month_text:
        return None

    for index in range(1, 13):
        if month_text.lower() in {
            calendar.month_abbr[index].lower(),
            calendar.month_name[index].lower(),
        }:
            return index

    # Support slightly noisy labels by matching first 3 chars.
    short = month_text[:3].lower()
    for index in range(1, 13):
        if short == calendar.month_abbr[index].lower():
            return index
    return None


def get_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    input_folder_name = input_path.resolve().name

    base_file = output_path / f"{input_folder_name}_PARAM.xlsx"
    if not base_file.exists():
        return base_file

    index = 1
    while True:
        candidate = output_path / f"{input_folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_workbook(wb: xw.main.Book) -> None:
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


def sheet_by_name(wb: xw.main.Book, sheet_name: str) -> xw.main.Sheet | None:
    for sh in wb.sheets:
        if sh.name == sheet_name:
            return sh
    return None


def find_max_anchor(cache: SheetCache) -> tuple[int, int] | None:
    candidates: list[tuple[int, int]] = []
    for r_idx, row_values in enumerate(cache.values):
        for c_idx, raw in enumerate(row_values):
            if normalize_label(raw) == "max":
                row = cache.start_row + r_idx
                col = cache.start_col + c_idx
                candidates.append((row, col))

    if not candidates:
        return None

    for row, col in candidates:
        if normalize_label(cache.value_at(row, col + 1)) == "min":
            return row, col

    return candidates[0]


def build_header_map(cache: SheetCache, anchor_row: int) -> dict[str, int]:
    header_map: dict[str, int] = {}
    rows_to_scan = [anchor_row - 1, anchor_row, anchor_row + 1]
    for row in rows_to_scan:
        if row < cache.start_row or row > cache.end_row:
            continue
        for col in range(cache.start_col, cache.end_col + 1):
            label = normalize_label(cache.value_at(row, col))
            if label and label not in header_map:
                header_map[label] = col
    return header_map


def find_column(header_map: dict[str, int], keywords: Iterable[str], fallback: int | None = None) -> int | None:
    for keyword in keywords:
        needle = normalize_label(keyword)
        for label, col in header_map.items():
            if needle in label:
                return col
    return fallback


def set_formula2_r1c1(cell: xw.main.Range, formula_r1c1: str) -> None:
    # Prefer .formula2 assignment first, then fall back to COM interfaces that
    # explicitly support R1C1 formula injection.
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

    cell.api.FormulaR1C1 = formula_r1c1


def extract_empirical_rows(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    cache = SheetCache.build(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(cache, anchor_row)

    row_start = anchor_row + 1
    max_col = anchor_col
    min_col = find_column(header_map, [" min "], fallback=anchor_col + 1) or (anchor_col + 1)
    num_q_col = find_column(header_map, ["num quarters used", "num quarters", "quarters used"], fallback=anchor_col - 5)
    last_q_col = find_column(header_map, ["last quarter used", "last quarter"], fallback=anchor_col - 4)
    avg_pen_col = find_column(header_map, ["avg penetration", "average penetration"], fallback=anchor_col - 3)
    forecast_col = find_column(
        header_map,
        ["estimated total sold", "est total sold", "forecast value", "tot fcst w/o sa", "tot fcst"],
        fallback=anchor_col - 2,
    )
    reported_col = find_column(header_map, ["reported sales", "actual sales", "actual value"], fallback=anchor_col - 1)
    quarterly_sales_col = find_column(header_map, ["quarterly sales"], fallback=anchor_col - 6)
    growth_col = find_column(header_map, ["growth rate", "growth %"], fallback=anchor_col + 2)
    captured_col = find_column(header_map, ["sales captured in db", "captured in db", "db pct"], fallback=anchor_col + 3)

    penetration_source_col = find_column(
        header_map,
        ["penetration", "pen %", "penetration pct"],
        fallback=avg_pen_col,
    )

    scratch_col = max(cache.end_col + 6, anchor_col + 12)
    scratch_start_row = max(cache.end_row + 2, row_start + N_QUARTERS + 2)

    formula_cells: list[xw.main.Range] = []
    source_data_start = row_start
    formulas_written = 0

    for n in range(1, N_QUARTERS + 1):
        if penetration_source_col is None:
            break
        source_end = source_data_start + n - 1
        formula = (
            f'=IFERROR(AVERAGE(R{source_data_start}C{penetration_source_col}:'
            f"R{source_end}C{penetration_source_col}),\"\")"
        )
        cell = sheet.range((scratch_start_row + n - 1, scratch_col))
        set_formula2_r1c1(cell, formula)
        formula_cells.append(cell)
        formulas_written += 1

    if formulas_written:
        wb.app.calculate()

    averaged_penetrations: list[float | None] = []
    for cell in formula_cells:
        averaged_penetrations.append(as_number(cell.value))

    rows: list[dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row = row_start + n - 1

        num_quarters_used = as_number(cache.value_at(row, num_q_col)) if num_q_col is not None else None
        if num_quarters_used is None:
            num_quarters_used = float(n)

        last_quarter_used = cache.value_at(row, last_q_col) if last_q_col is not None else None
        forecast_value = as_number(cache.value_at(row, forecast_col)) if forecast_col is not None else None
        reported_sales = as_number(cache.value_at(row, reported_col)) if reported_col is not None else None
        forecast_max = as_number(cache.value_at(row, max_col))
        forecast_min = as_number(cache.value_at(row, min_col))
        quarterly_sales = as_number(cache.value_at(row, quarterly_sales_col)) if quarterly_sales_col is not None else None
        growth_rate = as_number(cache.value_at(row, growth_col)) if growth_col is not None else None
        captured_pct = as_number(cache.value_at(row, captured_col)) if captured_col is not None else None

        avg_pen_sheet = as_number(cache.value_at(row, avg_pen_col)) if avg_pen_col is not None else None
        avg_pen_calc = averaged_penetrations[n - 1] if n - 1 < len(averaged_penetrations) else None
        avg_penetration = avg_pen_calc if avg_pen_calc is not None else avg_pen_sheet

        # Skip clearly empty rows while still supporting sparse layouts.
        key_fields = [forecast_value, forecast_max, forecast_min, avg_penetration, reported_sales]
        if all(value is None for value in key_fields):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
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
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    # Clear scratch formulas after reading calculated values.
    for cell in formula_cells:
        cell.value = None

    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    cache = SheetCache.build(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(cache, anchor_row)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    row_start = anchor_row + 1
    max_col = anchor_col
    min_col = find_column(header_map, [" min "], fallback=anchor_col + 1) or (anchor_col + 1)
    num_q_col = find_column(header_map, ["num quarters used", "num quarters", "quarters used"], fallback=anchor_col - 5)
    forecast_col = find_column(
        header_map,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast total without sa", "tot fcst"],
        fallback=anchor_col - 2,
    )
    actual_col = find_column(header_map, ["actual value", "actual sales", "reported sales"], fallback=None)

    xy_rows: list[tuple[int, float, float]] = []
    blank_streak = 0
    for row in range(row_start, cache.end_row + 1):
        x_val = as_number(cache.value_at(row, x_col))
        y_val = as_number(cache.value_at(row, y_col))
        if x_val is None or y_val is None:
            blank_streak += 1
            if blank_streak >= 8 and row > row_start + 20:
                break
            continue
        blank_streak = 0
        xy_rows.append((row, x_val, y_val))

    if len(xy_rows) < 2:
        return []

    max_n = min(N_QUARTERS, len(xy_rows))
    next_x = xy_rows[-1][1] + 1.0

    scratch_col = max(cache.end_col + 6, anchor_col + 12)
    scratch_start_row = max(cache.end_row + 2, row_start + max_n + 2)

    calc_blocks: list[dict[str, Any]] = []
    for n in range(2, max_n + 1):
        start_row = xy_rows[-n][0]
        end_row = xy_rows[-1][0]

        out_row = scratch_start_row + n
        intercept_cell = sheet.range((out_row, scratch_col))
        slope_cell = sheet.range((out_row, scratch_col + 1))
        forecast_cell = sheet.range((out_row, scratch_col + 2))

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
        )
        forecast_formula = f'=IFERROR(RC[-2]+RC[-1]*{next_x},"")'

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)

        calc_blocks.append(
            {
                "n": n,
                "intercept_cell": intercept_cell,
                "slope_cell": slope_cell,
                "forecast_cell": forecast_cell,
            }
        )

    if calc_blocks:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for block in calc_blocks:
        n = block["n"]
        row = row_start + n - 1

        num_quarters_used = as_number(cache.value_at(row, num_q_col)) if num_q_col is not None else None
        if num_quarters_used is None:
            num_quarters_used = float(n)

        intercept = as_number(block["intercept_cell"].value)
        slope = as_number(block["slope_cell"].value)

        forecast_from_sheet = as_number(cache.value_at(row, forecast_col)) if forecast_col is not None else None
        forecast_calc = as_number(block["forecast_cell"].value)
        forecast_value = forecast_from_sheet if forecast_from_sheet is not None else forecast_calc

        forecast_max = as_number(cache.value_at(row, max_col))
        forecast_min = as_number(cache.value_at(row, min_col))
        actual_value = cache.value_at(row, actual_col) if actual_col is not None else None

        if actual_value == "":
            actual_value = None

        if all(v is None for v in [intercept, slope, forecast_value, forecast_max, forecast_min]):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            round(float(num_quarters_used), 6) if num_quarters_used is not None else None,
            round(intercept, 6) if intercept is not None else None,
            round(slope, 6) if slope is not None else None,
            round(forecast_value, 6) if forecast_value is not None else None,
            round(forecast_max, 6) if forecast_max is not None else None,
            round(forecast_min, 6) if forecast_min is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
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
                "source_file": source_file,
            }
        )

    for block in calc_blocks:
        block["intercept_cell"].value = None
        block["slope_cell"].value = None
        block["forecast_cell"].value = None

    return rows


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def autosize_columns(ws: Any, min_width: int = 12, max_width: int = 42) -> None:
    for col_idx, _ in enumerate(ws[1], start=1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[col_letter]:
            value = cell.value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_dir}")

    output_path = get_output_path(in_dir, out_dir)
    generated_prefix = f"{in_dir.resolve().name}_param"

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
        all_files = sorted(in_dir.iterdir(), key=lambda path: path.name.lower())
        for file_path in all_files:
            if not file_path.is_file():
                continue
            if file_path.suffix.lower() != ".xlsx":
                log(f"Skipped {file_path.name}: not an .xlsx file")
                continue
            if file_path.name.startswith("~"):
                log(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.stem.lower().startswith(generated_prefix):
                log(f"Skipped {file_path.name}: generated output workbook")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path.name)

                empirical_sheet = sheet_by_name(wb, "Empirical Model")
                if empirical_sheet is not None:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = sheet_by_name(wb, "Regression Model")
                if regression_sheet is not None:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
                log(f"Processed {file_path.name}")
            except Exception as exc:
                log(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    log(f"Output workbook: {output_path}")
    log(f"Files processed: {processed_files}")
    log(f"Empirical rows: {len(empirical_rows)}")
    log(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
