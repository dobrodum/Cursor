#!/usr/bin/env python3
"""
Extract empirical and regression parameter candidates from Excel model workbooks.

Requirements implemented:
- Process every .xlsx file in input_dir once (skip temp "~" files).
- Open each workbook one time, process both model sheets, and close safely without saving.
- Write one output workbook with:
    - empirical_candidates
    - regression_candidates
- Use R1C1 `.formula2` for temporary calculation formulas.
- Add output formatting (bold header, frozen top row, filters, column widths).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font
import xlwings as xw


# -------------------------------
# User-configurable paths
# -------------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


# -------------------------------
# Constants
# -------------------------------
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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

N_QUARTERS = 10


# -------------------------------
# Helpers
# -------------------------------
def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        txt = value.strip().replace(",", "")
        if not txt:
            return None
        try:
            return float(txt)
        except ValueError:
            return None
    return None


def as_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def normalize_key(value: Any) -> str:
    if value is None:
        return ""
    txt = str(value).strip().lower()
    txt = re.sub(r"[^a-z0-9]+", "", txt)
    return txt


def close_workbook_safely(wb: Any) -> None:
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
    except Exception as exc:
        print(f"WARN: workbook close fallback failed: {exc}")


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base
    if not candidate.exists():
        return candidate

    i = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem

    # Expected pattern example:
    # MedMiner_Model - AORT - MidJan2026_Send.xlsx
    ticker = "UNKNOWN"
    ticker_match = re.search(r"-\s*([A-Za-z0-9_.-]+)\s*-\s*", stem)
    if ticker_match:
        ticker = ticker_match.group(1).strip()

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12][0-9]{3})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "UnknownPeriod"
    model_date = ""
    if period_match:
        period_part = period_match.group(1).title()  # Early/Mid/Late
        month_text = period_match.group(2)[:3].lower()
        year_text = period_match.group(3)
        month_num = MONTH_MAP.get(month_text)
        day_num = PERIOD_DAY_MAP.get(period_part.lower())

        if month_num and day_num:
            parsed_date = date(int(year_text), month_num, day_num)
            model_date = parsed_date.isoformat()
            model_period = f"{period_part}{month_text.title()}_{year_text}"

    model = f"{ticker}_{model_period}"
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


class SheetCache:
    def __init__(self, sheet: Any) -> None:
        self.sheet = sheet
        used = sheet.used_range
        self.start_row = used.row
        self.start_col = used.column
        self.values = as_2d(used.value)
        self.row_count = len(self.values)
        self.col_count = max((len(r) for r in self.values), default=0)
        self.end_row = self.start_row + self.row_count - 1
        self.end_col = self.start_col + self.col_count - 1

    def get(self, abs_row: int, abs_col: int) -> Any:
        r = abs_row - self.start_row
        c = abs_col - self.start_col
        if r < 0 or c < 0 or r >= self.row_count:
            return None
        row = self.values[r]
        if c >= len(row):
            return None
        return row[c]

    def row_values(self, abs_row: int) -> list[Any]:
        r = abs_row - self.start_row
        if r < 0 or r >= self.row_count:
            return []
        return self.values[r]

    def find_anchor_max(self) -> tuple[int, int] | None:
        # Find all "max" candidates, then pick one that likely belongs to a min/max table.
        candidates: list[tuple[int, int]] = []
        for r_idx, row in enumerate(self.values):
            for c_idx, cell in enumerate(row):
                if normalize_key(cell) == "max":
                    candidates.append(
                        (self.start_row + r_idx, self.start_col + c_idx)
                    )

        if not candidates:
            return None
        if len(candidates) == 1:
            return candidates[0]

        best = candidates[0]
        best_score = -1
        for row_num, col_num in candidates:
            score = 0
            # Score rows where nearby headers contain useful terms.
            row_vals = [normalize_key(v) for v in self.row_values(row_num)]
            row_txt = " ".join(row_vals)
            if "min" in row_vals:
                score += 3
            if "forecast" in row_txt or "fcst" in row_txt:
                score += 2
            if "quarter" in row_txt:
                score += 1
            if score > best_score:
                best_score = score
                best = (row_num, col_num)
        return best

    def header_map(self, header_row: int) -> dict[str, int]:
        out: dict[str, int] = {}
        for col in range(self.start_col, self.end_col + 1):
            key = normalize_key(self.get(header_row, col))
            if key and key not in out:
                out[key] = col
        return out


def find_col(header_map: dict[str, int], synonyms: list[str]) -> int | None:
    normalized = {normalize_key(k): v for k, v in header_map.items()}
    for synonym in synonyms:
        key = normalize_key(synonym)
        if key in normalized:
            return normalized[key]

    # Fuzzy fallback (contains-based).
    for synonym in synonyms:
        key = normalize_key(synonym)
        for existing, col in normalized.items():
            if key and key in existing:
                return col
    return None


def compute_range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def process_empirical_sheet(
    wb: Any,
    source_file: str,
    labels: FileLabels,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        print(f"SKIP SHEET: {source_file} -> missing 'Empirical Model'")
        return rows

    sheet = wb.sheets["Empirical Model"]
    cache = SheetCache(sheet)
    anchor = cache.find_anchor_max()
    if not anchor:
        print(f"SKIP SHEET: {source_file} -> Empirical 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    headers = cache.header_map(anchor_row)

    col_num_q = find_col(
        headers, ["num_quarters_used", "num quarters used", "n quarters", "quarters"]
    )
    col_last_q = find_col(headers, ["last_quarter_used", "last quarter used"])
    col_forecast = find_col(
        headers,
        [
            "estimated total sold",
            "forecast",
            "forecast value",
            "total forecast",
            "tot fcst",
        ],
    )
    col_actual = find_col(headers, ["actual", "actual value", "reported sales"])
    col_min = find_col(headers, ["min", "forecast min"])
    col_quarterly_sales = find_col(headers, ["quarterly sales"])
    col_reported_sales = find_col(headers, ["reported sales"])
    col_growth = find_col(headers, ["growth rate pct", "growth rate"])
    col_capture = find_col(headers, ["sales captured in db pct", "captured in db"])
    col_avg_pen = find_col(headers, ["avg penetration pct", "avg penetration"])

    # If no explicit min column found, use max+1 as a typical layout fallback.
    if col_min is None:
        col_min = anchor_col + 1

    data_start = anchor_row + 1
    scratch_col = cache.end_col + 2
    scratch_start = cache.end_row + 5
    formula_slots: list[tuple[int, int, int]] = []  # (row_index, scratch_row, scratch_col)

    for i in range(N_QUARTERS):
        row_num = data_start + i
        num_q = cache.get(row_num, col_num_q) if col_num_q else (i + 1)
        if is_blank(num_q):
            num_q = i + 1

        last_q = cache.get(row_num, col_last_q) if col_last_q else None
        forecast_value = cache.get(row_num, col_forecast) if col_forecast else None
        actual_value = cache.get(row_num, col_actual) if col_actual else None
        forecast_max = cache.get(row_num, anchor_col)
        forecast_min = cache.get(row_num, col_min)
        quarterly_sales = (
            cache.get(row_num, col_quarterly_sales) if col_quarterly_sales else None
        )
        reported_sales = (
            cache.get(row_num, col_reported_sales) if col_reported_sales else None
        )
        growth_rate = cache.get(row_num, col_growth) if col_growth else None
        capture_pct = cache.get(row_num, col_capture) if col_capture else None
        avg_penetration = cache.get(row_num, col_avg_pen) if col_avg_pen else None

        # Stop early if row appears fully empty in primary fields.
        if all(
            is_blank(v)
            for v in (forecast_value, actual_value, forecast_max, forecast_min, num_q)
        ):
            continue

        out_row: dict[str, Any] = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_q,
            "last_quarter_used": last_q,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": capture_pct,
            "source_file": source_file,
        }

        # If avg penetration is absent, compute via temporary R1C1 formula.
        if is_blank(avg_penetration) and col_quarterly_sales and col_reported_sales:
            n_used = int(to_float(num_q) or (i + 1))
            end_row = row_num
            start_row = max(data_start, end_row - n_used + 1)
            scratch_row = scratch_start + i
            formula = (
                f'=IFERROR(AVERAGE('
                f'R{start_row}C{col_quarterly_sales}:R{end_row}C{col_quarterly_sales}'
                f'/R{start_row}C{col_reported_sales}:R{end_row}C{col_reported_sales}'
                f'),"")'
            )
            sheet.range((scratch_row, scratch_col)).formula2 = formula
            formula_slots.append((len(rows), scratch_row, scratch_col))

        rows.append(out_row)

    if formula_slots:
        wb.app.calculate()
        for row_idx, s_row, s_col in formula_slots:
            calc_value = sheet.range((s_row, s_col)).value
            rows[row_idx]["avg_penetration_pct"] = calc_value
            rows[row_idx]["parameter_value"] = calc_value

    return rows


def process_regression_sheet(
    wb: Any,
    source_file: str,
    labels: FileLabels,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if "Regression Model" not in [s.name for s in wb.sheets]:
        print(f"SKIP SHEET: {source_file} -> missing 'Regression Model'")
        return rows

    sheet = wb.sheets["Regression Model"]
    cache = SheetCache(sheet)
    anchor = cache.find_anchor_max()
    if not anchor:
        print(f"SKIP SHEET: {source_file} -> Regression 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    headers = cache.header_map(anchor_row)

    # Required by prompt:
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    col_num_q_table = find_col(
        headers, ["num_quarters_used", "num quarters used", "n quarters", "quarters"]
    )
    col_fcst_table = find_col(
        headers,
        ["tot fcst w/o sa", "total forecast without sa", "forecast", "forecast value"],
    )
    col_min = find_col(headers, ["min", "forecast min"])
    if col_min is None:
        col_min = anchor_col + 1

    data_end_row = anchor_row - 1
    if data_end_row <= cache.start_row:
        print(f"SKIP SHEET: {source_file} -> insufficient rows above regression anchor")
        return rows

    scratch_col_intercept = cache.end_col + 2
    scratch_col_slope = cache.end_col + 3
    scratch_col_forecast = cache.end_col + 4
    scratch_start = cache.end_row + 5

    formulas_written = False
    calc_rows: list[dict[str, Any]] = []

    for i in range(N_QUARTERS):
        n_used = i + 1
        start_row = data_end_row - n_used + 1
        if start_row < cache.start_row:
            continue

        scratch_row = scratch_start + i
        intercept_formula = (
            f"=INTERCEPT("
            f"R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE("
            f"R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        # Forecast X uses anchor row X value as the next target point.
        forecast_formula = (
            f"=FORECAST.LINEAR("
            f"R{anchor_row}C{x_col},"
            f"R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )

        sheet.range((scratch_row, scratch_col_intercept)).formula2 = intercept_formula
        sheet.range((scratch_row, scratch_col_slope)).formula2 = slope_formula
        sheet.range((scratch_row, scratch_col_forecast)).formula2 = forecast_formula
        formulas_written = True

        table_row = anchor_row + i + 1
        table_num_q = cache.get(table_row, col_num_q_table) if col_num_q_table else None
        table_forecast = cache.get(table_row, col_fcst_table) if col_fcst_table else None
        table_max = cache.get(table_row, anchor_col)
        table_min = cache.get(table_row, col_min)

        calc_rows.append(
            {
                "_scratch_row": scratch_row,
                "num_quarters_used": table_num_q if not is_blank(table_num_q) else n_used,
                "table_forecast": table_forecast,
                "forecast_max": table_max,
                "forecast_min": table_min,
            }
        )

    if formulas_written:
        wb.app.calculate()

    for calc in calc_rows:
        s_row = calc["_scratch_row"]
        intercept = sheet.range((s_row, scratch_col_intercept)).value
        slope = sheet.range((s_row, scratch_col_slope)).value
        forecast_calc = sheet.range((s_row, scratch_col_forecast)).value
        forecast_value = (
            calc["table_forecast"]
            if not is_blank(calc["table_forecast"])
            else forecast_calc
        )
        forecast_max = calc["forecast_max"]
        forecast_min = calc["forecast_min"]

        out_row: dict[str, Any] = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": calc["num_quarters_used"],
            "num_quarters_used": calc["num_quarters_used"],
            "forecast_value": forecast_value,
            "actual_value": "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        # Prevent duplicate final row.
        if rows:
            prev = rows[-1]
            same = (
                prev.get("num_quarters_used") == out_row.get("num_quarters_used")
                and prev.get("forecast_value") == out_row.get("forecast_value")
                and prev.get("forecast_max") == out_row.get("forecast_max")
                and prev.get("forecast_min") == out_row.get("forecast_min")
                and prev.get("intercept") == out_row.get("intercept")
                and prev.get("slope") == out_row.get("slope")
            )
            if same:
                continue

        rows.append(out_row)

    return rows


def apply_sheet_formatting(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    if ws.max_row < 1 or ws.max_column < 1:
        return
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Set reasonable column widths based on content.
    for col_cells in ws.columns:
        max_len = 0
        for cell in col_cells:
            val = cell.value
            if val is None:
                continue
            val_len = len(str(val))
            if val_len > max_len:
                max_len = val_len
        adjusted = min(max(max_len + 2, 12), 48)
        ws.column_dimensions[col_cells[0].column_letter].width = adjusted


def write_output_workbook(
    path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = openpyxl.Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    ws_emp = wb.create_sheet("empirical_candidates")
    ws_emp.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        ws_emp.append([row.get(col, "") for col in EMPIRICAL_HEADERS])
    apply_sheet_formatting(ws_emp)

    ws_reg = wb.create_sheet("regression_candidates")
    ws_reg.append(REGRESSION_HEADERS)
    for row in regression_rows:
        ws_reg.append([row.get(col, "") for col in REGRESSION_HEADERS])
    apply_sheet_formatting(ws_reg)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def should_process_file(path: Path) -> tuple[bool, str]:
    if not path.is_file():
        return False, "not a file"
    if path.name.startswith("~"):
        return False, "temp file (~)"
    if path.suffix.lower() != ".xlsx":
        return False, "not .xlsx"
    return True, ""


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    source_files = sorted(input_dir.iterdir())
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in source_files:
            process_ok, reason = should_process_file(file_path)
            if not process_ok:
                print(f"SKIP: {file_path.name} ({reason})")
                continue

            print(f"PROCESS: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                files_processed += 1
                labels = parse_file_labels(file_path.name)

                empirical_rows.extend(
                    process_empirical_sheet(
                        wb=wb,
                        source_file=file_path.name,
                        labels=labels,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        wb=wb,
                        source_file=file_path.name,
                        labels=labels,
                    )
                )
            except Exception as exc:
                print(f"ERROR: {file_path.name} -> {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    output_path = unique_output_path(input_dir, output_dir)
    write_output_workbook(
        path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {files_processed}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
