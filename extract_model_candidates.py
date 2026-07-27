#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:
    raise SystemExit("xlwings is required to run this script.") from exc


# --------- configurable paths ---------
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


@dataclass(frozen=True)
class ModelLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetSnapshot:
    start_row: int
    start_col: int
    values: List[List[Any]]

    @property
    def n_rows(self) -> int:
        return len(self.values)

    @property
    def n_cols(self) -> int:
        return max((len(r) for r in self.values), default=0)


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if values and isinstance(values[0], (list, tuple)):
        return [list(row) for row in values]
    return [list(values)]


def build_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    return SheetSnapshot(start_row=used.row, start_col=used.column, values=matrix)


def snapshot_get(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    if row < snapshot.start_row or col < snapshot.start_col:
        return None
    r_idx = row - snapshot.start_row
    c_idx = col - snapshot.start_col
    if r_idx >= snapshot.n_rows or r_idx < 0:
        return None
    row_vals = snapshot.values[r_idx]
    if c_idx >= len(row_vals) or c_idx < 0:
        return None
    return row_vals[c_idx]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s_/\\\-]+", " ", text)
    return re.sub(r"[^a-z0-9 %]+", "", text)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def safe_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def find_anchor(snapshot: SheetSnapshot, target: str = "max") -> Optional[Tuple[int, int]]:
    target_norm = normalize_text(target)
    for r_idx, row_vals in enumerate(snapshot.values):
        abs_row = snapshot.start_row + r_idx
        for c_idx, raw_value in enumerate(row_vals):
            if normalize_text(raw_value) == target_norm:
                abs_col = snapshot.start_col + c_idx
                return abs_row, abs_col
    return None


def find_keyword_columns(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    patterns: Dict[str, Sequence[str]],
) -> Dict[str, int]:
    row_min = max(snapshot.start_row, anchor_row - 2)
    row_max = min(snapshot.start_row + snapshot.n_rows - 1, anchor_row + 2)
    col_min = max(1, anchor_col - 40)
    col_max = anchor_col + 40
    found: Dict[str, int] = {}

    for key, regex_list in patterns.items():
        best: Optional[Tuple[int, int]] = None
        for row in range(row_min, row_max + 1):
            for col in range(col_min, col_max + 1):
                label = normalize_text(snapshot_get(snapshot, row, col))
                if not label:
                    continue
                if any(re.search(rx, label) for rx in regex_list):
                    distance = abs(col - anchor_col)
                    if best is None or distance < best[0]:
                        best = (distance, col)
        if best:
            found[key] = best[1]
    return found


def default_empirical_columns(anchor_col: int) -> Dict[str, int]:
    return {
        "num_quarters_used": anchor_col - 6,
        "last_quarter_used": anchor_col - 5,
        "forecast_value": anchor_col - 2,
        "actual_value": anchor_col - 1,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
        "avg_penetration_pct": anchor_col - 3,
        "quarterly_sales": anchor_col - 8,
        "reported_sales": anchor_col - 1,
        "growth_rate_pct": anchor_col - 7,
        "sales_captured_in_db_pct": anchor_col - 4,
    }


def default_regression_columns(anchor_col: int) -> Dict[str, int]:
    return {
        "num_quarters_used": anchor_col - 5,
        "forecast_value": anchor_col - 2,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }


def set_formula2_r1c1(rng: xw.Range, formula_r1c1: str) -> None:
    try:
        # Keep formula2 usage explicit, then force R1C1 through the API when available.
        rng.formula2 = formula_r1c1
        rng.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    try:
        rng.formula2 = formula_r1c1
        return
    except Exception:
        pass
    rng.formula = formula_r1c1


def safe_close_without_save(wb: xw.Book) -> None:
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
    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        wb.api.Close(False)


def parse_model_label(file_name: str) -> ModelLabel:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""
    period_part = parts[2] if len(parts) >= 3 else ""
    period_part = period_part.split("_")[0]

    match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", period_part, flags=re.IGNORECASE)
    if not match:
        model_period = period_part or "unknown_period"
        model_date = ""
    else:
        phase_raw, month_raw, year_str = match.groups()
        phase = phase_raw.capitalize()
        month_three = month_raw[:3].title()
        year = int(year_str)
        month_num = datetime.strptime(month_three, "%b").month
        day = {"Early": 5, "Mid": 15, "Late": 25}[phase]
        model_period = f"{phase}{month_three}_{year}"
        model_date = f"{year:04d}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if ticker and model_period else stem
    return ModelLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def calc_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def live_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def best_last_numeric_row(snapshot: SheetSnapshot, col: int, before_row: int) -> Optional[int]:
    last_row: Optional[int] = None
    for row in range(snapshot.start_row, before_row + 1):
        value = snapshot_get(snapshot, row, col)
        if to_float(value) is not None:
            last_row = row
    return last_row


def build_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    snapshot: SheetSnapshot,
    label: ModelLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(snapshot, "max")
    if not anchor:
        print(f"Skipped empirical rows for {source_file}: could not find 'max' anchor in Empirical Model.")
        return []

    anchor_row, anchor_col = anchor
    pattern_cols = find_keyword_columns(
        snapshot,
        anchor_row,
        anchor_col,
        {
            "num_quarters_used": [r"num.*quarter"],
            "last_quarter_used": [r"last.*quarter"],
            "forecast_value": [r"estimated.*total.*sold", r"tot.*fcst", r"forecast"],
            "actual_value": [r"reported.*sales", r"actual"],
            "forecast_min": [r"^min$"],
            "avg_penetration_pct": [r"avg.*penetration"],
            "quarterly_sales": [r"quarterly.*sales"],
            "reported_sales": [r"reported.*sales"],
            "growth_rate_pct": [r"growth.*rate"],
            "sales_captured_in_db_pct": [r"sales.*captured.*db"],
        },
    )
    cols = default_empirical_columns(anchor_col)
    cols.update(pattern_cols)

    penetration_series_col = max(1, anchor_col - 11)
    data_end_row = best_last_numeric_row(snapshot, penetration_series_col, anchor_row - 1)

    # Write all temporary avg-penetration formulas first, then calculate once.
    for n in range(1, N_QUARTERS + 1):
        candidate_row = anchor_row + n
        avg_col = cols.get("avg_penetration_pct", anchor_col + 25)
        if avg_col < 1:
            avg_col = anchor_col + 25
        if data_end_row is None:
            formula = '=IFERROR(AVERAGE(RC[-20]:RC[-1]),"")'
        else:
            start_row = max(1, data_end_row - n + 1)
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{penetration_series_col}:'
                f'R{data_end_row}C{penetration_series_col}),"")'
            )
        set_formula2_r1c1(sheet.range((candidate_row, avg_col)), formula)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n
        num_quarters_used = to_int(live_cell_value(sheet, row_idx, cols["num_quarters_used"])) or n
        last_quarter_used = safe_text(live_cell_value(sheet, row_idx, cols["last_quarter_used"]))
        forecast_value = to_float(live_cell_value(sheet, row_idx, cols["forecast_value"]))
        actual_value = to_float(live_cell_value(sheet, row_idx, cols["actual_value"]))
        forecast_max = to_float(live_cell_value(sheet, row_idx, cols["forecast_max"]))
        forecast_min = to_float(live_cell_value(sheet, row_idx, cols["forecast_min"]))
        avg_penetration_pct = to_float(live_cell_value(sheet, row_idx, cols["avg_penetration_pct"]))
        quarterly_sales = to_float(live_cell_value(sheet, row_idx, cols["quarterly_sales"]))
        reported_sales = to_float(live_cell_value(sheet, row_idx, cols["reported_sales"]))
        growth_rate_pct = to_float(live_cell_value(sheet, row_idx, cols["growth_rate_pct"]))
        sales_captured_in_db_pct = to_float(live_cell_value(sheet, row_idx, cols["sales_captured_in_db_pct"]))

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales if reported_sales is not None else actual_value,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )
    return rows


def regression_data_block_rows(
    snapshot: SheetSnapshot,
    anchor_row: int,
    x_col: int,
    y_col: int,
) -> List[int]:
    numeric_rows: List[int] = []
    for row in range(snapshot.start_row, anchor_row):
        x_val = to_float(snapshot_get(snapshot, row, x_col))
        y_val = to_float(snapshot_get(snapshot, row, y_col))
        if x_val is not None and y_val is not None:
            numeric_rows.append(row)
    if not numeric_rows:
        return []
    # Keep only the contiguous tail block that ends nearest the anchor.
    tail: List[int] = [numeric_rows[-1]]
    for idx in range(len(numeric_rows) - 2, -1, -1):
        if numeric_rows[idx] == tail[0] - 1:
            tail.insert(0, numeric_rows[idx])
        else:
            break
    return tail


def build_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    snapshot: SheetSnapshot,
    label: ModelLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(snapshot, "max")
    if not anchor:
        print(f"Skipped regression rows for {source_file}: could not find 'max' anchor in Regression Model.")
        return []

    anchor_row, anchor_col = anchor
    x_col = max(1, anchor_col - 11)
    y_col = max(1, anchor_col - 7)

    pattern_cols = find_keyword_columns(
        snapshot,
        anchor_row,
        anchor_col,
        {
            "num_quarters_used": [r"num.*quarter"],
            "forecast_value": [r"tot.*fcst.*w ?o.*sa", r"total.*forecast.*without.*sa", r"forecast"],
            "actual_value": [r"actual", r"reported.*sales"],
            "forecast_min": [r"^min$"],
        },
    )
    cols = default_regression_columns(anchor_col)
    cols.update(pattern_cols)

    data_rows = regression_data_block_rows(snapshot, anchor_row, x_col, y_col)
    if data_rows:
        data_end_row = data_rows[-1]
    else:
        data_end_row = max(1, anchor_row - 1)

    intercept_col = anchor_col + 20
    slope_col = anchor_col + 21

    for n in range(1, N_QUARTERS + 1):
        candidate_row = anchor_row + n
        if data_rows:
            start_index = max(0, len(data_rows) - n)
            start_row = data_rows[start_index]
        else:
            start_row = max(1, data_end_row - n + 1)

        x_range = f"R{start_row}C{x_col}:R{data_end_row}C{x_col}"
        y_range = f"R{start_row}C{y_col}:R{data_end_row}C{y_col}"
        set_formula2_r1c1(
            sheet.range((candidate_row, intercept_col)),
            f'=IFERROR(INTERCEPT({y_range},{x_range}),"")',
        )
        set_formula2_r1c1(
            sheet.range((candidate_row, slope_col)),
            f'=IFERROR(SLOPE({y_range},{x_range}),"")',
        )

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n
        num_quarters_used = to_int(live_cell_value(sheet, row_idx, cols["num_quarters_used"])) or n
        forecast_value = to_float(live_cell_value(sheet, row_idx, cols["forecast_value"]))
        actual_col = cols.get("actual_value")
        actual_value = to_float(live_cell_value(sheet, row_idx, actual_col)) if actual_col else None
        forecast_max = to_float(live_cell_value(sheet, row_idx, cols["forecast_max"]))
        forecast_min = to_float(live_cell_value(sheet, row_idx, cols["forecast_min"]))
        intercept = to_float(live_cell_value(sheet, row_idx, intercept_col))
        slope = to_float(live_cell_value(sheet, row_idx, slope_col))

        if all(
            value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        row_data = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        # Prevent duplicate final row.
        if rows and n == N_QUARTERS:
            previous = rows[-1]
            keys = ("num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
            if all(row_data.get(k) == previous.get(k) for k in keys):
                continue

        rows.append(row_data)

    return rows


def is_generated_output_file(file_name: str) -> bool:
    return bool(re.search(r"_PARAM(?:\.\d+)?\.xlsx$", file_name, flags=re.IGNORECASE))


def iter_input_files(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            print(f"Skipped: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped: {file_path.name} (temporary workbook)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped: {file_path.name} (not .xlsx)")
            continue
        if is_generated_output_file(file_path.name):
            print(f"Skipped: {file_path.name} (generated output workbook)")
            continue
        yield file_path


def next_output_path(input_path: Path, out_path: Path) -> Path:
    folder_name = input_path.name
    base = out_path / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base
    index = 1
    while True:
        candidate = out_path / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def apply_sheet_format(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    bold = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold

    for col_idx, header in enumerate(headers, start=1):
        values = [header]
        for row in rows:
            value = row.get(header)
            if value is not None:
                values.append(str(value))
        width = max(12, min(48, max(len(v) for v in values) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb_out = Workbook()
    ws_emp = wb_out.active
    ws_emp.title = "empirical_candidates"
    ws_emp.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_emp.append([row.get(col) for col in EMPIRICAL_COLUMNS])
    apply_sheet_format(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = wb_out.create_sheet("regression_candidates")
    ws_reg.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_reg.append([row.get(col) for col in REGRESSION_COLUMNS])
    apply_sheet_format(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb_out.save(output_path)


def main() -> None:
    in_path = Path(input_dir)
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    if not in_path.exists() or not in_path.is_dir():
        raise SystemExit(f"Input directory does not exist or is not a folder: {in_path}")

    output_path = next_output_path(in_path, out_path)
    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_input_files(in_path):
            print(f"Processing: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (open failed: {exc})")
                continue

            try:
                label = parse_model_label(file_path.name)

                if "Empirical Model" in [s.name for s in wb.sheets]:
                    emp_sheet = wb.sheets["Empirical Model"]
                    emp_snapshot = build_snapshot(emp_sheet)
                    empirical_rows.extend(
                        build_empirical_rows(wb, emp_sheet, emp_snapshot, label, file_path.name)
                    )
                else:
                    print(f"Skipped empirical rows for {file_path.name}: missing sheet 'Empirical Model'.")

                if "Regression Model" in [s.name for s in wb.sheets]:
                    reg_sheet = wb.sheets["Regression Model"]
                    reg_snapshot = build_snapshot(reg_sheet)
                    regression_rows.extend(
                        build_regression_rows(wb, reg_sheet, reg_snapshot, label, file_path.name)
                    )
                else:
                    print(f"Skipped regression rows for {file_path.name}: missing sheet 'Regression Model'.")

                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing failed: {exc})")
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
