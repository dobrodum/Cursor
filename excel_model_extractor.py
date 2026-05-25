#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Required top-level configuration variables.
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


@dataclass
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


@dataclass
class SheetSnapshot:
    values: List[List[Any]]
    start_row: int
    start_col: int
    end_row: int
    end_col: int


def to_float(value: Any) -> Optional[float]:
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
        try:
            return float(text)
        except ValueError:
            return None
    return None


def normalize_key(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def r1c1_abs(row: int, col: int) -> str:
    return f"R{row}C{col}"


def safe_close_workbook(wb: Any) -> None:
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
        pass


def set_formula2_r1c1(cell: Any, formula_r1c1: str) -> None:
    # Prefer formula2 as requested; fallback to COM property for compatibility.
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
    cell.formula = formula_r1c1


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, tuple):
        return [list(row) for row in values]
    if isinstance(first, list):
        return values
    return [values]


def snapshot_sheet(sheet: Any) -> SheetSnapshot:
    used = sheet.used_range
    values = ensure_2d(used.value)
    start_row = int(used.row)
    start_col = int(used.column)
    n_rows = len(values)
    n_cols = len(values[0]) if values else 0
    end_row = start_row + n_rows - 1
    end_col = start_col + n_cols - 1
    return SheetSnapshot(
        values=values,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )


def get_snapshot_cell(snapshot: SheetSnapshot, abs_row: int, abs_col: int) -> Any:
    row_idx = abs_row - snapshot.start_row
    col_idx = abs_col - snapshot.start_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(snapshot.values):
        return None
    row = snapshot.values[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


def get_rows_with_numeric(
    snapshot: SheetSnapshot,
    required_cols: Sequence[int],
    max_row: Optional[int] = None,
) -> List[int]:
    rows: List[int] = []
    upper_row = snapshot.end_row if max_row is None else min(snapshot.end_row, max_row)
    for abs_row in range(snapshot.start_row, upper_row + 1):
        values = [to_float(get_snapshot_cell(snapshot, abs_row, col)) for col in required_cols]
        if all(v is not None for v in values):
            rows.append(abs_row)
    return rows


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int]] = []
    for r_idx, row in enumerate(snapshot.values):
        for c_idx, value in enumerate(row):
            if normalize_key(value) == "max":
                abs_row = snapshot.start_row + r_idx
                abs_col = snapshot.start_col + c_idx
                candidates.append((abs_row, abs_col))
    if not candidates:
        return None

    # Prefer an anchor where "min" exists nearby in the same row.
    for abs_row, abs_col in candidates:
        for delta in range(1, 5):
            near_value = get_snapshot_cell(snapshot, abs_row, abs_col + delta)
            if normalize_key(near_value) == "min":
                return (abs_row, abs_col)
    return candidates[0]


def build_header_maps(snapshot: SheetSnapshot, anchor_row: int) -> Dict[str, int]:
    header_rows = [anchor_row, anchor_row - 1]
    header_map: Dict[str, int] = {}
    for header_row in header_rows:
        if header_row < snapshot.start_row or header_row > snapshot.end_row:
            continue
        for abs_col in range(snapshot.start_col, snapshot.end_col + 1):
            key = normalize_key(get_snapshot_cell(snapshot, header_row, abs_col))
            if key:
                header_map.setdefault(key, abs_col)
    return header_map


def find_col_by_aliases(
    header_map: Dict[str, int],
    aliases: Iterable[str],
    fallback: Optional[int] = None,
) -> Optional[int]:
    normalized_aliases = [normalize_key(alias) for alias in aliases]
    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]
    for alias in normalized_aliases:
        for header, col in header_map.items():
            if alias and alias in header:
                return col
    return fallback


def month_number(mon_text: str) -> int:
    short = mon_text[:3].title()
    lookup = {calendar.month_abbr[idx]: idx for idx in range(1, 13)}
    if short not in lookup:
        raise ValueError(f"Unknown month abbreviation: {mon_text}")
    return lookup[short]


def parse_metadata(file_path: Path) -> ModelMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[-2] if len(parts) >= 2 else "UNKNOWN"
    period_token = parts[-1].split("_")[0].strip() if parts else "UNKNOWN"

    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$", period_token, re.IGNORECASE)
    if match:
        phase = match.group(1).title()
        mon = match.group(2).title()
        year = int(match.group(3))
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        month_idx = month_number(mon)
        model_period = f"{phase}{mon}_{year}"
        model_date = date(year, month_idx, day_lookup[phase]).isoformat()
    else:
        model_period = period_token
        model_date = ""

    model = f"{ticker}_{model_period}" if ticker and model_period else stem
    return ModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def eq_or_blank(left: Any, right: Any, tol: float = 1e-9) -> bool:
    lf = to_float(left)
    rf = to_float(right)
    if lf is None and rf is None:
        return True
    if lf is not None and rf is not None:
        return abs(lf - rf) <= tol
    return str(left) == str(right)


def next_output_path(input_root: Path, output_root: Path) -> Path:
    output_root.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_root.name}_PARAM"
    candidate = output_root / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_root / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def build_empirical_rows(wb: Any, sheet: Any, metadata: ModelMetadata) -> List[Dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"  - Empirical Model: no 'max' anchor found in {metadata.source_file}")
        return []

    anchor_row, anchor_col = anchor
    headers = build_header_maps(snapshot, anchor_row)

    max_col = anchor_col
    min_col = find_col_by_aliases(headers, ["min"], fallback=anchor_col + 1) or anchor_col + 1
    penetration_col = find_col_by_aliases(
        headers,
        ["avg_penetration_pct", "penetration_pct", "penetration", "pen_pct"],
        fallback=anchor_col - 8,
    )
    quarter_col = find_col_by_aliases(headers, ["quarter", "last_quarter"], fallback=anchor_col - 10)
    quarterly_sales_col = find_col_by_aliases(
        headers,
        ["quarterly_sales", "qtr_sales", "sales_qtr"],
        fallback=anchor_col - 6,
    )
    forecast_col = find_col_by_aliases(
        headers,
        ["estimated_total_sold", "forecast_value", "tot_fcst", "total_forecast"],
        fallback=anchor_col - 1,
    )
    reported_sales_col = find_col_by_aliases(
        headers,
        ["reported_sales", "actual_sales", "reported"],
        fallback=anchor_col - 5,
    )
    growth_col = find_col_by_aliases(headers, ["growth_rate_pct", "growth_pct", "growth"], fallback=anchor_col - 4)
    captured_col = find_col_by_aliases(
        headers,
        ["sales_captured_in_db_pct", "captured_in_db_pct", "captured_pct"],
        fallback=anchor_col - 3,
    )

    if penetration_col is None or quarterly_sales_col is None:
        print(f"  - Empirical Model: required columns missing in {metadata.source_file}")
        return []

    historical_rows = get_rows_with_numeric(
        snapshot,
        required_cols=[penetration_col, quarterly_sales_col],
        max_row=anchor_row - 1,
    )
    if not historical_rows:
        print(f"  - Empirical Model: no historical numeric rows in {metadata.source_file}")
        return []

    n_rows = min(N_QUARTERS, len(historical_rows))
    helper_col = max(anchor_col + 8, snapshot.end_col + 1)
    helper_start_row = anchor_row + 1

    windows: List[Tuple[int, int, int]] = []
    for idx in range(n_rows):
        n_q = idx + 1
        start_idx = len(historical_rows) - n_q
        start_row = historical_rows[start_idx]
        end_row = historical_rows[-1]
        windows.append((n_q, start_row, end_row))
        avg_cell = sheet.range((helper_start_row + idx, helper_col))
        avg_formula = f"=AVERAGE({r1c1_abs(start_row, penetration_col)}:{r1c1_abs(end_row, penetration_col)})"
        set_formula2_r1c1(avg_cell, avg_formula)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx, (n_q, start_row, end_row) in enumerate(windows):
        avg_penetration = to_float(sheet.range((helper_start_row + idx, helper_col)).value)
        quarterly_sales = to_float(get_snapshot_cell(snapshot, end_row, quarterly_sales_col))
        reported_sales = to_float(get_snapshot_cell(snapshot, end_row, reported_sales_col))
        growth_rate = to_float(get_snapshot_cell(snapshot, end_row, growth_col))
        captured_pct = to_float(get_snapshot_cell(snapshot, end_row, captured_col))
        last_quarter_used = get_snapshot_cell(snapshot, end_row, quarter_col)

        candidate_row = anchor_row + n_q
        forecast_value = to_float(get_snapshot_cell(snapshot, candidate_row, forecast_col))
        if forecast_value is None and avg_penetration is not None and quarterly_sales is not None:
            forecast_value = avg_penetration * quarterly_sales

        # Prefer workbook-provided candidate max/min rows if present.
        forecast_max = to_float(get_snapshot_cell(snapshot, candidate_row, max_col))
        forecast_min = to_float(get_snapshot_cell(snapshot, candidate_row, min_col))

        # Fallback to range from selected penetration window.
        if forecast_max is None or forecast_min is None:
            selected_penetrations = [
                to_float(get_snapshot_cell(snapshot, row_idx, penetration_col))
                for row_idx in historical_rows[-n_q:]
            ]
            selected_penetrations = [v for v in selected_penetrations if v is not None]
            if quarterly_sales is not None and selected_penetrations:
                forecast_max = max(selected_penetrations) * quarterly_sales
                forecast_min = min(selected_penetrations) * quarterly_sales

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_q,
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
                "source_file": metadata.source_file,
            }
        )

    return rows


def build_regression_rows(wb: Any, sheet: Any, metadata: ModelMetadata) -> List[Dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"  - Regression Model: no 'max' anchor found in {metadata.source_file}")
        return []

    anchor_row, anchor_col = anchor
    headers = build_header_maps(snapshot, anchor_row)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    max_col = anchor_col
    min_col = find_col_by_aliases(headers, ["min"], fallback=anchor_col + 1) or anchor_col + 1
    actual_col = find_col_by_aliases(headers, ["actual_value", "actual", "reported_sales"], fallback=None)
    forecast_col = find_col_by_aliases(
        headers,
        ["tot_fcst_w_o_sa", "tot_fcst_wo_sa", "tot_fcst_without_sa", "forecast_total_without_sa"],
        fallback=anchor_col - 1,
    )

    historical_rows = get_rows_with_numeric(snapshot, required_cols=[x_col, y_col], max_row=anchor_row - 1)
    if not historical_rows:
        print(f"  - Regression Model: no historical x/y rows in {metadata.source_file}")
        return []

    n_rows = min(N_QUARTERS, len(historical_rows))
    helper_col_intercept = max(anchor_col + 8, snapshot.end_col + 1)
    helper_col_slope = helper_col_intercept + 1
    helper_start_row = anchor_row + 1

    windows: List[Tuple[int, int, int]] = []
    for idx in range(n_rows):
        n_q = idx + 1
        start_idx = len(historical_rows) - n_q
        start_row = historical_rows[start_idx]
        end_row = historical_rows[-1]
        windows.append((n_q, start_row, end_row))

        intercept_formula = (
            f"=INTERCEPT({r1c1_abs(start_row, y_col)}:{r1c1_abs(end_row, y_col)},"
            f"{r1c1_abs(start_row, x_col)}:{r1c1_abs(end_row, x_col)})"
        )
        slope_formula = (
            f"=SLOPE({r1c1_abs(start_row, y_col)}:{r1c1_abs(end_row, y_col)},"
            f"{r1c1_abs(start_row, x_col)}:{r1c1_abs(end_row, x_col)})"
        )
        set_formula2_r1c1(sheet.range((helper_start_row + idx, helper_col_intercept)), intercept_formula)
        set_formula2_r1c1(sheet.range((helper_start_row + idx, helper_col_slope)), slope_formula)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx, (n_q, _start_row, end_row) in enumerate(windows):
        intercept = to_float(sheet.range((helper_start_row + idx, helper_col_intercept)).value)
        slope = to_float(sheet.range((helper_start_row + idx, helper_col_slope)).value)

        last_x = to_float(get_snapshot_cell(snapshot, end_row, x_col))
        candidate_row = anchor_row + n_q
        forecast_total_without_sa = to_float(get_snapshot_cell(snapshot, candidate_row, forecast_col))
        if forecast_total_without_sa is None and intercept is not None and slope is not None and last_x is not None:
            forecast_total_without_sa = intercept + slope * (last_x + 1.0)

        forecast_max = to_float(get_snapshot_cell(snapshot, candidate_row, max_col))
        forecast_min = to_float(get_snapshot_cell(snapshot, candidate_row, min_col))

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        actual_value = get_snapshot_cell(snapshot, candidate_row, actual_col) if actual_col else None

        row_data = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_q,
            "num_quarters_used": n_q,
            "forecast_value": forecast_total_without_sa,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": metadata.source_file,
        }

        # Prevent duplicated final row.
        if rows:
            prev = rows[-1]
            duplicate = (
                eq_or_blank(prev["forecast_value"], row_data["forecast_value"])
                and eq_or_blank(prev["forecast_max"], row_data["forecast_max"])
                and eq_or_blank(prev["forecast_min"], row_data["forecast_min"])
                and eq_or_blank(prev["intercept"], row_data["intercept"])
                and eq_or_blank(prev["slope"], row_data["slope"])
            )
            if duplicate:
                continue

        rows.append(row_data)

    return rows


def apply_sheet_formatting(sheet: Any, headers: Sequence[str]) -> None:
    for col_idx, name in enumerate(headers, start=1):
        sheet.cell(row=1, column=col_idx, value=name)
        sheet.cell(row=1, column=col_idx).font = Font(bold=True)

    sheet.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(headers))
    sheet.auto_filter.ref = f"A1:{last_col_letter}{max(1, sheet.max_row)}"

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    empirical_ws.append(EMPIRICAL_COLUMNS)
    regression_ws.append(REGRESSION_COLUMNS)

    for row_data in empirical_rows:
        empirical_ws.append([row_data.get(col) for col in EMPIRICAL_COLUMNS])
    for row_data in regression_rows:
        regression_ws.append([row_data.get(col) for col in REGRESSION_COLUMNS])

    apply_sheet_formatting(empirical_ws, EMPIRICAL_COLUMNS)
    apply_sheet_formatting(regression_ws, REGRESSION_COLUMNS)
    wb.save(output_path)


def iter_source_files(folder: Path) -> Tuple[List[Path], List[Tuple[str, str]]]:
    files: List[Path] = []
    skipped: List[Tuple[str, str]] = []

    if not folder.exists():
        skipped.append((str(folder), "input directory does not exist"))
        return files, skipped

    for item in sorted(folder.iterdir()):
        if not item.is_file():
            continue
        if item.name.startswith("~"):
            skipped.append((item.name, "temporary file"))
            continue
        if item.suffix.lower() != ".xlsx":
            skipped.append((item.name, "non-xlsx extension"))
            continue
        files.append(item)

    return files, skipped


def main() -> None:
    source_files, skipped = iter_source_files(input_dir)
    for filename, reason in skipped:
        print(f"Skipped: {filename} ({reason})")

    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_metadata(file_path)
                sheet_names = {s.name for s in wb.sheets}

                if "Empirical Model" in sheet_names:
                    empirical_sheet = wb.sheets["Empirical Model"]
                    empirical_rows.extend(build_empirical_rows(wb, empirical_sheet, metadata))
                else:
                    print(f"  - Missing sheet 'Empirical Model' in {file_path.name}")

                if "Regression Model" in sheet_names:
                    regression_sheet = wb.sheets["Regression Model"]
                    regression_rows.extend(build_regression_rows(wb, regression_sheet, metadata))
                else:
                    print(f"  - Missing sheet 'Regression Model' in {file_path.name}")

                processed_count += 1
            except Exception as exc:
                print(f"  - Failed to process {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
