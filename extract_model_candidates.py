#!/usr/bin/env python3
"""
Extract empirical and regression candidates from a folder of Excel models.

The script opens each source workbook once with xlwings, processes both
"Empirical Model" and "Regression Model" sheets, then writes one output
workbook containing:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ----------------------------
# User inputs
# ----------------------------
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


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    values: List[List[Any]]
    start_row: int
    start_col: int
    label_map: Dict[str, List[Tuple[int, int]]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        if not self.values:
            return 0
        return len(self.values[0])


FILE_META_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
    r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
    re.IGNORECASE,
)

MONTH_TO_NUM = {
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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def safe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        pct_text = text[:-1].strip()
        try:
            return float(pct_text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def compute_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def parse_model_meta(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    match = FILE_META_PATTERN.search(stem)

    if not match:
        parts = [p.strip() for p in stem.split("-")]
        ticker = parts[1].upper() if len(parts) >= 2 else "UNKNOWN"
        model_period = "Unknown_0000"
        return ModelMeta(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    ticker = match.group("ticker").upper()
    phase = match.group("phase").title()
    month_token = match.group("month")[:3].lower()
    month_num = MONTH_TO_NUM.get(month_token)
    year = int(match.group("year"))

    if month_num is None:
        model_period = f"{phase}{match.group('month')}_{year}"
        model_date = ""
    else:
        month_abbr = match.group("month")[:3].title()
        day = PHASE_TO_DAY[phase.lower()]
        model_period = f"{phase}{month_abbr}_{year}"
        model_date = date(year, month_num, day).isoformat()

    return ModelMeta(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def make_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    raw_values = used.options(ndim=2).value

    if raw_values is None:
        values: List[List[Any]] = []
    elif isinstance(raw_values, list):
        if raw_values and not isinstance(raw_values[0], list):
            values = [raw_values]
        else:
            values = raw_values
    else:
        values = [[raw_values]]

    if values:
        width = max(len(row) for row in values)
        for row in values:
            if len(row) < width:
                row.extend([None] * (width - len(row)))

    label_map: Dict[str, List[Tuple[int, int]]] = {}
    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            label = normalize_text(cell_value)
            if label:
                abs_row = used.row + r_idx
                abs_col = used.column + c_idx
                label_map.setdefault(label, []).append((abs_row, abs_col))

    return SheetSnapshot(
        values=values,
        start_row=used.row,
        start_col=used.column,
        label_map=label_map,
    )


def snapshot_get(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    r_idx = row - snapshot.start_row
    c_idx = col - snapshot.start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= snapshot.row_count or c_idx >= snapshot.col_count:
        return None
    return snapshot.values[r_idx][c_idx]


def find_anchor_max(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    if "max" in snapshot.label_map:
        return snapshot.label_map["max"][0]

    candidates: List[Tuple[int, int]] = []
    for label, positions in snapshot.label_map.items():
        if " max" in f" {label}" or label.endswith("max") or label.startswith("max "):
            candidates.extend(positions)
    return candidates[0] if candidates else None


def nearest_positions(
    positions: Iterable[Tuple[int, int]], anchor: Optional[Tuple[int, int]]
) -> List[Tuple[int, int]]:
    position_list = list(positions)
    if anchor is None:
        return position_list
    return sorted(position_list, key=lambda p: abs(p[0] - anchor[0]) + abs(p[1] - anchor[1]))


def find_label_positions(snapshot: SheetSnapshot, tokens: Sequence[str]) -> List[Tuple[int, int]]:
    matches: List[Tuple[int, int]] = []
    for label, positions in snapshot.label_map.items():
        if any(token in label for token in tokens):
            matches.extend(positions)
    return matches


def find_value_cell_near_label(
    snapshot: SheetSnapshot,
    tokens: Sequence[str],
    anchor: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    label_positions = nearest_positions(find_label_positions(snapshot, tokens), anchor)
    neighbor_offsets = ((0, 1), (0, 2), (1, 0), (-1, 0), (1, 1), (-1, 1), (0, -1))

    for row, col in label_positions:
        for d_row, d_col in neighbor_offsets:
            value_row = row + d_row
            value_col = col + d_col
            value = snapshot_get(snapshot, value_row, value_col)
            if value not in (None, ""):
                return (value_row, value_col)
    return None


def find_value_by_label(
    snapshot: SheetSnapshot,
    tokens: Sequence[str],
    anchor: Optional[Tuple[int, int]] = None,
    numeric_only: bool = False,
) -> Any:
    value_cell = find_value_cell_near_label(snapshot, tokens, anchor=anchor)
    if value_cell is None:
        return None

    value = snapshot_get(snapshot, value_cell[0], value_cell[1])
    if not numeric_only:
        return value
    return safe_float(value)


def first_numeric(*values: Any) -> Optional[float]:
    for value in values:
        parsed = safe_float(value)
        if parsed is not None:
            return parsed
    return None


def guess_penetration_col(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> int:
    history_end = anchor_row - 1
    history_start = max(1, history_end - 24)
    candidate_cols = [c for c in range(max(1, anchor_col - 12), anchor_col) if c > 0]

    best_col = max(1, anchor_col - 1)
    best_score = -1

    for col in candidate_cols:
        score = 0
        for row in range(history_start, history_end + 1):
            value = safe_float(snapshot_get(snapshot, row, col))
            if value is None:
                continue
            if 0 <= value <= 1.5:
                score += 2
            elif 1.5 < value <= 100:
                score += 1
        if score > best_score:
            best_col = col
            best_score = score
    return best_col


def guess_quarter_col(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> Optional[int]:
    history_end = anchor_row - 1
    history_start = max(1, history_end - 24)
    candidate_cols = [c for c in range(max(1, anchor_col - 15), anchor_col)]
    quarter_regex = re.compile(r"(q[1-4]|20\d{2}|fy)", re.IGNORECASE)

    best_col: Optional[int] = None
    best_score = 0
    for col in candidate_cols:
        score = 0
        for row in range(history_start, history_end + 1):
            value = snapshot_get(snapshot, row, col)
            if value is None:
                continue
            text = str(value).strip()
            if quarter_regex.search(text):
                score += 1
        if score > best_score:
            best_col = col
            best_score = score
    return best_col


def set_formula2(target: xw.Range, formula_r1c1: str) -> None:
    try:
        target.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        target.api.Formula2 = formula_r1c1
        return
    except Exception:
        pass
    target.formula = formula_r1c1


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(SaveChanges=False)
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

    workbook.close()


def determine_forecast_bounds(
    snapshot: SheetSnapshot, anchor_row: int, anchor_col: int
) -> Tuple[Optional[float], Optional[float]]:
    # Anchor-based offsets from the "max" cell.
    max_candidate = snapshot_get(snapshot, anchor_row, anchor_col + 1)
    min_candidate = snapshot_get(snapshot, anchor_row + 1, anchor_col + 1)

    max_value = safe_float(max_candidate)
    min_value = safe_float(min_candidate)

    if max_value is None:
        max_value = find_value_by_label(snapshot, ["forecast max", "max"], (anchor_row, anchor_col), True)
    if min_value is None:
        min_value = find_value_by_label(snapshot, ["forecast min", "min"], (anchor_row, anchor_col), True)

    return max_value, min_value


def process_empirical_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    model_meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = make_snapshot(sheet)
    if snapshot.row_count == 0:
        return []

    anchor = find_anchor_max(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    forecast_max, forecast_min = determine_forecast_bounds(snapshot, anchor_row, anchor_col)

    quarterly_sales = find_value_by_label(
        snapshot, ["quarterly sales", "qtr sales", "sales this quarter"], anchor, numeric_only=True
    )
    reported_sales = find_value_by_label(snapshot, ["reported sales", "actual sales", "reported"], anchor, True)
    growth_rate_pct = find_value_by_label(snapshot, ["growth rate", "growth %"], anchor, True)
    sales_captured_in_db_pct = find_value_by_label(
        snapshot, ["sales captured in db", "captured in db", "captured"], anchor, True
    )

    estimated_total_cell = find_value_cell_near_label(
        snapshot, ["estimated total sold", "est total sold", "forecast value"], anchor
    )

    avg_pen_cell = find_value_cell_near_label(snapshot, ["avg penetration", "average penetration"], anchor)
    if avg_pen_cell is None:
        avg_pen_cell = (anchor_row + 2, anchor_col + 4)

    penetration_col = guess_penetration_col(snapshot, anchor_row, anchor_col)
    quarter_col = guess_quarter_col(snapshot, anchor_row, anchor_col)

    history_end_row = anchor_row - 1
    rows: List[Dict[str, Any]] = []

    for num_quarters_used in range(1, N_QUARTERS + 1):
        history_start_row = history_end_row - num_quarters_used + 1
        if history_start_row < 1:
            break

        avg_formula = (
            f"=AVERAGE(R{history_start_row}C{penetration_col}:"
            f"R{history_end_row}C{penetration_col})"
        )
        set_formula2(sheet.range(avg_pen_cell), avg_formula)
        workbook.app.calculate()

        avg_penetration_pct = safe_float(sheet.range(avg_pen_cell).value)
        if avg_penetration_pct is None:
            continue

        forecast_value = None
        if estimated_total_cell is not None:
            forecast_value = safe_float(sheet.range(estimated_total_cell).value)
        if forecast_value is None and quarterly_sales not in (None, 0) and avg_penetration_pct not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration_pct

        last_quarter_used = find_value_by_label(snapshot, ["last quarter used"], anchor, numeric_only=False)
        if last_quarter_used in (None, "") and quarter_col is not None:
            last_quarter_used = snapshot_get(snapshot, history_end_row, quarter_col)

        row = {
            "model": model_meta.model,
            "ticker": model_meta.ticker,
            "model_period": model_meta.model_period,
            "model_date": model_meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    model_meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = make_snapshot(sheet)
    if snapshot.row_count == 0:
        return []

    anchor = find_anchor_max(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    forecast_max, forecast_min = determine_forecast_bounds(snapshot, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    forecast_cell = find_value_cell_near_label(
        snapshot,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast total without sa", "tot fcst"],
        anchor,
    )
    actual_value = find_value_by_label(snapshot, ["actual value", "actual"], anchor, numeric_only=True)

    intercept_cell = (anchor_row + 2, anchor_col + 4)
    slope_cell = (anchor_row + 2, anchor_col + 5)

    history_end_row = anchor_row - 1
    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for num_quarters_used in range(2, N_QUARTERS + 1):
        history_start_row = history_end_row - num_quarters_used + 1
        if history_start_row < 1:
            break

        intercept_formula = (
            f"=INTERCEPT(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{history_start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{history_start_row}C{x_col}:R{history_end_row}C{x_col})"
        )

        set_formula2(sheet.range(intercept_cell), intercept_formula)
        set_formula2(sheet.range(slope_cell), slope_formula)
        workbook.app.calculate()

        intercept = safe_float(sheet.range(intercept_cell).value)
        slope = safe_float(sheet.range(slope_cell).value)
        if intercept is None or slope is None:
            continue

        forecast_total_without_sa = None
        if forecast_cell is not None:
            forecast_total_without_sa = safe_float(sheet.range(forecast_cell).value)

        if forecast_total_without_sa is None:
            x_next = first_numeric(
                snapshot_get(snapshot, history_end_row + 1, x_col),
                (
                    first_numeric(snapshot_get(snapshot, history_end_row, x_col))
                    if snapshot_get(snapshot, history_end_row, x_col) is not None
                    else None
                ),
            )
            if x_next is not None and snapshot_get(snapshot, history_end_row + 1, x_col) is None:
                x_next = x_next + 1
            if x_next is not None:
                forecast_total_without_sa = intercept + slope * x_next

        signature = (
            round(intercept, 10),
            round(slope, 10),
            round(forecast_total_without_sa, 10) if forecast_total_without_sa is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        # Guard against duplicate trailing rows.
        if previous_signature == signature:
            continue
        previous_signature = signature

        row = {
            "model": model_meta.model,
            "ticker": model_meta.ticker,
            "model_period": model_meta.model_period,
            "model_date": model_meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_total_without_sa,
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def iter_input_files(folder: Path) -> List[Path]:
    files: List[Path] = []
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            print(f"skipped file: {path.name} (not a file)")
            continue
        if path.name.startswith("~"):
            print(f"skipped file: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"skipped file: {path.name} (not .xlsx)")
            continue
        files.append(path)
    return files


def build_output_path(out_dir: Path, input_folder_name: str) -> Path:
    base = f"{input_folder_name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = out_dir / f"{base}.{idx}.xlsx"
        idx += 1
    return candidate


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(sheet_name)
    ws.append(list(columns))

    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir not found: {input_dir.resolve()}")

    output_dir.mkdir(parents=True, exist_ok=True)
    source_files = iter_input_files(input_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            workbook: Optional[xw.Book] = None
            try:
                # Open source workbook in read/process mode only.
                workbook = app.books.open(str(file_path), update_links=False)
                model_meta = parse_model_meta(file_path)

                try:
                    empirical_sheet = workbook.sheets["Empirical Model"]
                    empirical_rows.extend(
                        process_empirical_sheet(workbook, empirical_sheet, model_meta, file_path.name)
                    )
                except Exception as exc:
                    print(f"skipped empirical sheet in {file_path.name}: {exc}")

                try:
                    regression_sheet = workbook.sheets["Regression Model"]
                    regression_rows.extend(
                        process_regression_sheet(workbook, regression_sheet, model_meta, file_path.name)
                    )
                except Exception as exc:
                    print(f"skipped regression sheet in {file_path.name}: {exc}")

                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    # Never save source workbooks.
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    output_path = build_output_path(output_dir, input_dir.name)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
