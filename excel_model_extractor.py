#!/usr/bin/env python3
"""
Extract empirical/regression candidate rows from Excel model workbooks.

This script opens each source workbook once, processes both model sheets while
open, and writes one combined output workbook with two tabs:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ==============================
# User-configurable paths
# ==============================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


N_QUARTERS = 10

EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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


DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}

# Fallback offsets are relative to the `max` anchor column.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "forecast_value": -6,
    "actual_value": -5,
    "avg_penetration_pct": -4,
    "quarterly_sales": -12,
    "reported_sales": -11,
    "growth_rate_pct": -8,
    "sales_captured_in_db_pct": -7,
}

EMPIRICAL_ALIASES = {
    "num_quarters_used": [
        "num_quarters_used",
        "num quarters used",
        "n quarters",
        "# quarters used",
        "quarters used",
    ],
    "last_quarter_used": [
        "last_quarter_used",
        "last quarter used",
        "last quarter",
    ],
    "forecast_value": [
        "estimated total sold",
        "forecast value",
        "forecast",
        "tot fcst",
    ],
    "actual_value": [
        "actual value",
        "actual",
        "reported sales",
    ],
    "avg_penetration_pct": [
        "avg_penetration_pct",
        "avg penetration pct",
        "avg penetration %",
        "average penetration",
        "avg penetration",
    ],
    "quarterly_sales": [
        "quarterly_sales",
        "quarterly sales",
        "quarter sales",
    ],
    "reported_sales": [
        "reported_sales",
        "reported sales",
    ],
    "growth_rate_pct": [
        "growth_rate_pct",
        "growth rate pct",
        "growth rate %",
    ],
    "sales_captured_in_db_pct": [
        "sales_captured_in_db_pct",
        "sales captured in db pct",
        "sales captured in db %",
        "captured in db",
    ],
}

REGRESSION_ALIASES = {
    "num_quarters_used": [
        "num_quarters_used",
        "num quarters used",
        "n quarters",
        "quarters used",
    ],
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "forecast total without sa",
        "forecast value",
        "forecast",
    ],
    "actual_value": [
        "actual value",
        "actual",
        "reported sales",
    ],
}

PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12][0-9]{3})", re.IGNORECASE)


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
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


def to_int(value: Any) -> Optional[int]:
    number = to_number(value)
    if number is None:
        return None
    return int(round(number))


def to_text(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text if text else None


def signature_value(value: Any) -> Any:
    number = to_number(value)
    if number is None:
        return value
    return round(number, 10)


def set_formula2(cell: xw.main.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Older Excel builds may not support formula2; fallback keeps script usable.
        cell.formula = formula_r1c1


def close_source_workbook(wb: xw.main.Book) -> None:
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
        try:
            wb.close()
        except Exception:
            pass


def parse_model_meta(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) > 1 and parts[1] else "UNKNOWN"

    match = PERIOD_RE.search(stem)
    if not match:
        model_period = "Unknown_0000"
        model_date = ""
    else:
        phase_raw, month_raw, year_raw = match.groups()
        phase = phase_raw.capitalize()
        month_abbrev = month_raw[:3].title()
        try:
            month_number = datetime.strptime(month_abbrev, "%b").month
            day = DAY_BY_PHASE[phase.lower()]
            year = int(year_raw)
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = date(year, month_number, day).isoformat()
        except ValueError:
            model_period = "Unknown_0000"
            model_date = ""

    model = f"{ticker}_{model_period}"
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def find_max_anchor(ws: xw.main.Sheet) -> Optional[Tuple[int, int]]:
    used = ws.used_range
    values = to_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return start_row + row_offset, start_col + col_offset
    return None


def build_label_map(
    ws: xw.main.Sheet,
    anchor_row: int,
    start_col: int,
    end_col: int,
) -> Dict[str, int]:
    label_map: Dict[str, int] = {}
    for row in (anchor_row, anchor_row - 1):
        if row < 1:
            continue
        row_values = ws.range((row, start_col), (row, end_col)).value
        if not isinstance(row_values, list):
            row_values = [row_values]
        for idx, value in enumerate(row_values):
            normalized = normalize_label(value)
            if normalized and normalized not in label_map:
                label_map[normalized] = start_col + idx
    return label_map


def resolve_col(label_map: Dict[str, int], aliases: Sequence[str], fallback_col: int) -> int:
    alias_norms = [normalize_label(alias) for alias in aliases]
    for alias in alias_norms:
        if alias in label_map:
            return label_map[alias]
    for alias in alias_norms:
        for label, col in label_map.items():
            if alias and alias in label:
                return col
    return fallback_col


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def empirical_columns(ws: xw.main.Sheet, anchor_row: int, anchor_col: int) -> Dict[str, int]:
    scan_start = max(1, anchor_col - 30)
    scan_end = min(16384, anchor_col + 10)
    labels = build_label_map(ws, anchor_row, scan_start, scan_end)

    cols = {
        "num_quarters_used": resolve_col(
            labels,
            EMPIRICAL_ALIASES["num_quarters_used"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["num_quarters_used"],
        ),
        "last_quarter_used": resolve_col(
            labels,
            EMPIRICAL_ALIASES["last_quarter_used"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["last_quarter_used"],
        ),
        "forecast_value": resolve_col(
            labels,
            EMPIRICAL_ALIASES["forecast_value"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["forecast_value"],
        ),
        "actual_value": resolve_col(
            labels,
            EMPIRICAL_ALIASES["actual_value"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["actual_value"],
        ),
        "avg_penetration_pct": resolve_col(
            labels,
            EMPIRICAL_ALIASES["avg_penetration_pct"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["avg_penetration_pct"],
        ),
        "quarterly_sales": resolve_col(
            labels,
            EMPIRICAL_ALIASES["quarterly_sales"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["quarterly_sales"],
        ),
        "reported_sales": resolve_col(
            labels,
            EMPIRICAL_ALIASES["reported_sales"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["reported_sales"],
        ),
        "growth_rate_pct": resolve_col(
            labels,
            EMPIRICAL_ALIASES["growth_rate_pct"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["growth_rate_pct"],
        ),
        "sales_captured_in_db_pct": resolve_col(
            labels,
            EMPIRICAL_ALIASES["sales_captured_in_db_pct"],
            anchor_col + EMPIRICAL_FALLBACK_OFFSETS["sales_captured_in_db_pct"],
        ),
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }
    return cols


def regression_columns(
    ws: xw.main.Sheet,
    anchor_row: int,
    anchor_col: int,
    y_col: int,
) -> Dict[str, int]:
    scan_start = max(1, anchor_col - 30)
    scan_end = min(16384, anchor_col + 10)
    labels = build_label_map(ws, anchor_row, scan_start, scan_end)

    cols = {
        "num_quarters_used": resolve_col(
            labels,
            REGRESSION_ALIASES["num_quarters_used"],
            anchor_col - 10,
        ),
        "forecast_value": resolve_col(
            labels,
            REGRESSION_ALIASES["forecast_value"],
            y_col,
        ),
        "actual_value": resolve_col(
            labels,
            REGRESSION_ALIASES["actual_value"],
            anchor_col - 6,
        ),
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }
    return cols


def extract_empirical_rows(
    wb: xw.main.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        ws = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped sheet in {source_file}: Empirical Model not found")
        return []

    anchor = find_max_anchor(ws)
    if not anchor:
        print(f"Skipped sheet in {source_file}: Empirical Model max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    cols = empirical_columns(ws, anchor_row, anchor_col)
    scratch_col = min(16384, anchor_col + 25)
    rows: List[Dict[str, Any]] = []

    for idx in range(1, N_QUARTERS + 1):
        row_num = anchor_row + idx

        num_quarters_used = to_int(ws.cells(row_num, cols["num_quarters_used"]).value) or idx
        last_quarter_used = to_text(ws.cells(row_num, cols["last_quarter_used"]).value)

        forecast_max = to_number(ws.cells(row_num, cols["forecast_max"]).value)
        forecast_min = to_number(ws.cells(row_num, cols["forecast_min"]).value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        quarterly_sales = to_number(ws.cells(row_num, cols["quarterly_sales"]).value)
        reported_sales = to_number(ws.cells(row_num, cols["reported_sales"]).value)
        growth_rate_pct = to_number(ws.cells(row_num, cols["growth_rate_pct"]).value)
        sales_captured_in_db_pct = to_number(ws.cells(row_num, cols["sales_captured_in_db_pct"]).value)

        avg_penetration_pct = to_number(ws.cells(row_num, cols["avg_penetration_pct"]).value)
        if avg_penetration_pct is None and sales_captured_in_db_pct is not None:
            start_row = max(anchor_row + 1, row_num - num_quarters_used + 1)
            avg_formula_cell = ws.cells(row_num, scratch_col)
            set_formula2(
                avg_formula_cell,
                (
                    f'=IFERROR(AVERAGE('
                    f'R{start_row}C{cols["sales_captured_in_db_pct"]}:'
                    f'R{row_num}C{cols["sales_captured_in_db_pct"]}),"")'
                ),
            )
            wb.app.calculate()
            avg_penetration_pct = to_number(avg_formula_cell.value)
            avg_formula_cell.value = None

        forecast_value = to_number(ws.cells(row_num, cols["forecast_value"]).value)
        actual_value = to_number(ws.cells(row_num, cols["actual_value"]).value)
        if actual_value is None:
            actual_value = reported_sales

        if forecast_value is None and avg_penetration_pct not in (None, 0) and quarterly_sales is not None:
            forecast_value = quarterly_sales / avg_penetration_pct

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
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        ws = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped sheet in {source_file}: Regression Model not found")
        return []

    anchor = find_max_anchor(ws)
    if not anchor:
        print(f"Skipped sheet in {source_file}: Regression Model max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    cols = regression_columns(ws, anchor_row, anchor_col, y_col)
    intercept_col = min(16382, anchor_col + 25)
    slope_col = intercept_col + 1

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(1, N_QUARTERS + 1):
        row_num = anchor_row + idx
        num_quarters_used = to_int(ws.cells(row_num, cols["num_quarters_used"]).value) or idx

        intercept: Optional[float] = None
        slope: Optional[float] = None

        end_row = row_num - 1
        start_row = max(anchor_row + 1, end_row - num_quarters_used + 1)
        if start_row <= end_row and x_col > 0 and y_col > 0:
            has_x = to_number(ws.cells(start_row, x_col).value) is not None and to_number(
                ws.cells(end_row, x_col).value
            ) is not None
            has_y = to_number(ws.cells(start_row, y_col).value) is not None and to_number(
                ws.cells(end_row, y_col).value
            ) is not None

            if has_x and has_y:
                intercept_cell = ws.cells(row_num, intercept_col)
                slope_cell = ws.cells(row_num, slope_col)

                set_formula2(
                    intercept_cell,
                    (
                        f'=IFERROR(INTERCEPT('
                        f'R{start_row}C{y_col}:R{end_row}C{y_col},'
                        f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
                    ),
                )
                set_formula2(
                    slope_cell,
                    (
                        f'=IFERROR(SLOPE('
                        f'R{start_row}C{y_col}:R{end_row}C{y_col},'
                        f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
                    ),
                )
                wb.app.calculate()
                intercept = to_number(intercept_cell.value)
                slope = to_number(slope_cell.value)
                intercept_cell.value = None
                slope_cell.value = None

        forecast_value = to_number(ws.cells(row_num, cols["forecast_value"]).value)
        actual_value = to_number(ws.cells(row_num, cols["actual_value"]).value)
        forecast_max = to_number(ws.cells(row_num, cols["forecast_max"]).value)
        forecast_min = to_number(ws.cells(row_num, cols["forecast_min"]).value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if forecast_value is None and intercept is not None and slope is not None:
            x_val = to_number(ws.cells(row_num, x_col).value)
            if x_val is not None:
                forecast_value = intercept + (slope * x_val)

        signature = (
            num_quarters_used,
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        if all(
            value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

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
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws: Any, columns: Sequence[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col}{max(ws.max_row, 1)}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_width = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_width = max(max_width, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_width + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = wb.create_sheet("regression_candidates")
    write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def should_skip_file(file_path: Path, input_folder_name: str) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary workbook"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.match(rf"^{re.escape(input_folder_name)}_PARAM(\.\d+)?\.xlsx$", file_path.name):
        return "existing output workbook"
    return None


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir()):
            skip_reason = should_skip_file(file_path, in_dir.name)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            print(f"Processing {file_path.name}")
            workbook = None
            try:
                meta = parse_model_meta(file_path)
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, meta, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
