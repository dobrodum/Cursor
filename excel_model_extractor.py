#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks.

This script processes all .xlsx files in input_dir, opens each source workbook
exactly once, extracts both model sheets while the workbook is open, and writes
one consolidated output workbook with two sheets:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency
    raise SystemExit("Missing dependency: xlwings is required to run this script.") from exc


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


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

WINDOW_TO_DAY = {
    "Early": 5,
    "Mid": 15,
    "Late": 25,
}

MONTH_TO_NUMBER = {
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

# Anchor-based offsets from the "max" anchor cell.
# These are used first for speed; label-based lookups are fallbacks.
EMPIRICAL_MAX_VALUE_OFFSETS = ((1, 0), (0, 1), (1, 1))
EMPIRICAL_MIN_VALUE_OFFSETS = ((1, -1), (0, -1), (1, -2))
EMPIRICAL_AVG_PEN_OFFSETS = ((2, -3), (1, -3), (2, -2))
EMPIRICAL_EST_TOTAL_OFFSETS = ((3, -2), (3, -1), (2, -2))
EMPIRICAL_REPORTED_SALES_OFFSETS = ((4, -2), (4, -1), (3, -1))
EMPIRICAL_QUARTERLY_SALES_OFFSETS = ((5, -2), (5, -1), (4, -1))
EMPIRICAL_GROWTH_RATE_OFFSETS = ((6, -2), (6, -1), (5, -1))
EMPIRICAL_CAPTURED_OFFSETS = ((7, -2), (7, -1), (6, -1))

REGRESSION_MAX_VALUE_OFFSETS = ((1, 0), (0, 1), (1, 1))
REGRESSION_MIN_VALUE_OFFSETS = ((1, -1), (0, -1), (1, -2))
REGRESSION_FORECAST_OFFSETS = ((3, -2), (2, -2), (3, -1))
REGRESSION_INTERCEPT_OFFSETS = ((2, 2), (3, 2), (2, 3))
REGRESSION_SLOPE_OFFSETS = ((3, 2), (4, 2), (3, 3))


@dataclass
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class NumericSeries:
    cells: list[Any]
    orientation: str  # "horizontal" or "vertical"


@dataclass
class EmpiricalContext:
    max_anchor: Any
    min_anchor: Any
    forecast_max_cell: Any
    forecast_min_cell: Any
    avg_penetration_cell: Any
    estimated_total_sold_cell: Any
    reported_sales_cell: Any
    quarterly_sales_cell: Any
    growth_rate_cell: Any
    sales_captured_cell: Any
    penetration_series: NumericSeries | None


@dataclass
class RegressionContext:
    max_anchor: Any
    min_anchor: Any
    forecast_max_cell: Any
    forecast_min_cell: Any
    forecast_wo_sa_cell: Any
    actual_value_cell: Any
    intercept_cell: Any
    slope_cell: Any
    num_quarters_cell: Any
    x_col: int
    y_col: int
    xy_rows: list[int]


def log(message: str) -> None:
    print(message, flush=True)


def normalize_text(value: Any) -> str:
    return str(value).strip().lower() if value is not None else ""


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text or text.startswith("#"):
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except ValueError:
        return None


def clean_cell_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return value


def number_or_blank(value: Any) -> float | str:
    numeric = to_float(value)
    return numeric if numeric is not None else ""


def safe_subtract(left: Any, right: Any) -> float | str:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


def normalize_signature_value(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 10)
    return value


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def gather_input_files(folder: Path) -> tuple[list[Path], list[tuple[Path, str]]]:
    process_files: list[Path] = []
    skipped: list[tuple[Path, str]] = []

    if not folder.exists():
        raise FileNotFoundError(f"Input directory does not exist: {folder}")

    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            skipped.append((file_path, "temp file"))
            continue
        if file_path.suffix.lower() != ".xlsx":
            skipped.append((file_path, "not .xlsx"))
            continue
        process_files.append(file_path)
    return process_files, skipped


def parse_model_metadata(file_path: Path) -> ModelMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].strip().upper() if len(parts) >= 2 else ""

    period_raw = parts[2] if len(parts) >= 3 else stem
    period_raw = re.sub(r"(?i)[ _-]*send.*$", "", period_raw).strip(" _-")

    match = re.search(
        r"(?i)\b(early|mid|late)[ _-]*([a-z]{3,9})[ _-]*(\d{2,4})\b",
        period_raw,
    )

    model_period = ""
    model_date = ""
    if match:
        window = match.group(1).title()
        month_key = match.group(2).lower()
        month_number = MONTH_TO_NUMBER.get(month_key)
        if month_number is None:
            month_number = MONTH_TO_NUMBER.get(month_key[:3])

        year_token = match.group(3)
        year = int(year_token)
        if year < 100:
            year += 2000

        if month_number is not None:
            month_abbrev = date(year, month_number, 1).strftime("%b")
            model_period = f"{window}{month_abbrev}_{year}"
            model_date = date(year, month_number, WINDOW_TO_DAY[window]).isoformat()

    if not ticker:
        ticker_match = re.search(r" - ([A-Za-z0-9]+) - ", stem)
        ticker = ticker_match.group(1).upper() if ticker_match else "UNKNOWN"

    if model_period:
        model = f"{ticker}_{model_period}"
    else:
        sanitized = re.sub(r"\s+", "_", stem).strip("_")
        model = f"{ticker}_{sanitized}"

    return ModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_sheet_case_insensitive(workbook: Any, sheet_name: str) -> Any | None:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def safe_close_workbook(workbook: Any) -> None:
    """Close source workbook without saving with safe fallbacks."""
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        workbook.api.Close(False)


def set_formula2_r1c1(cell: Any, formula_r1c1: str) -> None:
    if not formula_r1c1.startswith("="):
        formula_r1c1 = "=" + formula_r1c1

    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.FormulaR1C1 = formula_r1c1
        return
    except Exception:
        cell.formula = formula_r1c1


def find_cell(sheet: Any, text: str, *, exact: bool = False) -> Any | None:
    look_at_options = (1,) if exact else (1, 2)
    for look_at in look_at_options:
        try:
            found = sheet.api.Cells.Find(
                What=text,
                After=sheet.api.Cells(1, 1),
                LookIn=-4163,  # xlValues
                LookAt=look_at,  # xlWhole / xlPart
                SearchOrder=1,  # xlByRows
                SearchDirection=1,  # xlNext
                MatchCase=False,
            )
        except Exception:
            found = None

        if found is not None:
            return sheet.range((found.Row, found.Column))
    return None


def find_label_cell(sheet: Any, labels: Iterable[str], *, exact: bool = False) -> Any | None:
    for label in labels:
        found = find_cell(sheet, label, exact=exact)
        if found is not None:
            return found
    return None


def pick_neighbor_value_cell(label_cell: Any) -> Any:
    right = label_cell.offset(0, 1)
    below = label_cell.offset(1, 0)

    if to_float(right.value) is not None:
        return right
    if to_float(below.value) is not None:
        return below

    right_text = normalize_text(right.value)
    below_text = normalize_text(below.value)
    if right_text and not below_text:
        return right
    if below_text and not right_text:
        return below
    return right


def find_value_cell(sheet: Any, labels: Iterable[str], *, exact: bool = False) -> Any | None:
    label_cell = find_label_cell(sheet, labels, exact=exact)
    if label_cell is None:
        return None
    return pick_neighbor_value_cell(label_cell)


def first_numeric_cell(cells: Sequence[Any]) -> Any | None:
    first = None
    for cell in cells:
        if cell is None:
            continue
        if first is None:
            first = cell
        if to_float(cell.value) is not None:
            return cell
    return first


def cells_from_offsets(anchor_cell: Any, offsets: Sequence[tuple[int, int]]) -> list[Any]:
    return [anchor_cell.offset(row_off, col_off) for row_off, col_off in offsets]


def value_cell_from_anchor(anchor_cell: Any, offsets: Sequence[tuple[int, int]]) -> Any | None:
    return first_numeric_cell(cells_from_offsets(anchor_cell, offsets))


def scan_numeric_series(start_cell: Any, row_step: int, col_step: int) -> list[Any]:
    cells: list[Any] = []
    blank_run = 0
    max_scan = 64
    for step in range(1, max_scan + 1):
        current = start_cell.offset(step * row_step, step * col_step)
        if to_float(current.value) is not None:
            cells.append(current)
            blank_run = 0
        else:
            blank_run += 1
            if cells and blank_run >= 2:
                break
    return cells


def extract_numeric_series(label_cell: Any) -> NumericSeries | None:
    options: list[NumericSeries] = []
    scans = [
        ("horizontal", scan_numeric_series(label_cell, 0, 1)),
        ("vertical", scan_numeric_series(label_cell, 1, 0)),
        ("horizontal", scan_numeric_series(label_cell, 0, -1)),
        ("vertical", scan_numeric_series(label_cell, -1, 0)),
    ]
    for orientation, cells in scans:
        if cells:
            ordered = cells
            if orientation == "horizontal" and len(cells) > 1 and cells[0].column > cells[-1].column:
                ordered = list(reversed(cells))
            if orientation == "vertical" and len(cells) > 1 and cells[0].row > cells[-1].row:
                ordered = list(reversed(cells))
            options.append(NumericSeries(cells=ordered, orientation=orientation))

    if not options:
        return None
    return max(options, key=lambda series: len(series.cells))


def infer_last_quarter_used(series: NumericSeries) -> Any:
    if not series.cells:
        return ""
    last_cell = series.cells[-1]
    if series.orientation == "horizontal":
        candidates = [last_cell.offset(-1, 0), last_cell.offset(1, 0)]
    else:
        candidates = [last_cell.offset(0, -1), last_cell.offset(0, 1)]

    for candidate in candidates:
        value = clean_cell_value(candidate.value)
        if value != "":
            return value
    return clean_cell_value(last_cell.value)


def collect_xy_rows(sheet: Any, anchor_row: int, x_col: int, y_col: int) -> list[int]:
    rows: list[int] = []
    blank_run = 0
    for row_idx in range(anchor_row - 1, 0, -1):
        x_value = to_float(sheet.range((row_idx, x_col)).value)
        y_value = to_float(sheet.range((row_idx, y_col)).value)
        if x_value is not None and y_value is not None:
            rows.append(row_idx)
            blank_run = 0
        else:
            blank_run += 1
            if rows and blank_run >= 3:
                break
    rows.reverse()
    return rows


def build_empirical_context(sheet: Any) -> EmpiricalContext:
    max_anchor = find_cell(sheet, "max", exact=True)
    if max_anchor is None:
        raise ValueError("Empirical Model missing 'max' anchor")

    min_anchor = find_cell(sheet, "min", exact=True)

    forecast_max_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_MAX_VALUE_OFFSETS)
    if forecast_max_cell is None:
        forecast_max_cell = find_value_cell(sheet, ["max"], exact=True)

    forecast_min_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_MIN_VALUE_OFFSETS)
    if forecast_min_cell is None and min_anchor is not None:
        forecast_min_cell = first_numeric_cell(
            [min_anchor.offset(1, 0), min_anchor.offset(0, 1), min_anchor.offset(1, 1)]
        )
    if forecast_min_cell is None:
        forecast_min_cell = find_value_cell(sheet, ["min"], exact=True)

    avg_penetration_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_AVG_PEN_OFFSETS)
    if avg_penetration_cell is None:
        avg_penetration_cell = find_value_cell(
            sheet,
            ["avg penetration", "average penetration", "avg penetration pct"],
        )

    estimated_total_sold_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_EST_TOTAL_OFFSETS)
    if estimated_total_sold_cell is None:
        estimated_total_sold_cell = find_value_cell(
            sheet,
            ["estimated total sold", "est total sold", "forecast total sold"],
        )

    reported_sales_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_REPORTED_SALES_OFFSETS)
    if reported_sales_cell is None:
        reported_sales_cell = find_value_cell(sheet, ["reported sales", "actual sales"])

    quarterly_sales_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_QUARTERLY_SALES_OFFSETS)
    if quarterly_sales_cell is None:
        quarterly_sales_cell = find_value_cell(sheet, ["quarterly sales"])

    growth_rate_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_GROWTH_RATE_OFFSETS)
    if growth_rate_cell is None:
        growth_rate_cell = find_value_cell(sheet, ["growth rate", "growth rate %"])

    sales_captured_cell = value_cell_from_anchor(max_anchor, EMPIRICAL_CAPTURED_OFFSETS)
    if sales_captured_cell is None:
        sales_captured_cell = find_value_cell(
            sheet,
            ["sales captured in db", "sales captured", "captured in db"],
        )

    penetration_label = find_label_cell(
        sheet,
        ["penetration", "penetration %", "quarter penetration", "pen %"],
    )
    penetration_series = extract_numeric_series(penetration_label) if penetration_label else None

    return EmpiricalContext(
        max_anchor=max_anchor,
        min_anchor=min_anchor,
        forecast_max_cell=forecast_max_cell,
        forecast_min_cell=forecast_min_cell,
        avg_penetration_cell=avg_penetration_cell,
        estimated_total_sold_cell=estimated_total_sold_cell,
        reported_sales_cell=reported_sales_cell,
        quarterly_sales_cell=quarterly_sales_cell,
        growth_rate_cell=growth_rate_cell,
        sales_captured_cell=sales_captured_cell,
        penetration_series=penetration_series,
    )


def build_regression_context(sheet: Any) -> RegressionContext:
    max_anchor = find_cell(sheet, "max", exact=True)
    if max_anchor is None:
        raise ValueError("Regression Model missing 'max' anchor")

    min_anchor = find_cell(sheet, "min", exact=True)
    x_col = max_anchor.column - 11
    y_col = max_anchor.column - 7
    xy_rows = collect_xy_rows(sheet, max_anchor.row, x_col, y_col)

    forecast_max_cell = value_cell_from_anchor(max_anchor, REGRESSION_MAX_VALUE_OFFSETS)
    if forecast_max_cell is None:
        forecast_max_cell = find_value_cell(sheet, ["max"], exact=True)

    forecast_min_cell = value_cell_from_anchor(max_anchor, REGRESSION_MIN_VALUE_OFFSETS)
    if forecast_min_cell is None and min_anchor is not None:
        forecast_min_cell = first_numeric_cell(
            [min_anchor.offset(1, 0), min_anchor.offset(0, 1), min_anchor.offset(1, 1)]
        )
    if forecast_min_cell is None:
        forecast_min_cell = find_value_cell(sheet, ["min"], exact=True)

    forecast_wo_sa_cell = value_cell_from_anchor(max_anchor, REGRESSION_FORECAST_OFFSETS)
    if forecast_wo_sa_cell is None:
        forecast_wo_sa_cell = find_value_cell(
            sheet,
            ["tot fcst w/o sa", "tot fcst without sa", "total fcst w/o sa"],
        )

    intercept_cell = value_cell_from_anchor(max_anchor, REGRESSION_INTERCEPT_OFFSETS)
    if intercept_cell is None:
        intercept_cell = find_value_cell(sheet, ["intercept"])
    if intercept_cell is None:
        intercept_cell = max_anchor.offset(2, 4)

    slope_cell = value_cell_from_anchor(max_anchor, REGRESSION_SLOPE_OFFSETS)
    if slope_cell is None:
        slope_cell = find_value_cell(sheet, ["slope"])
    if slope_cell is None:
        slope_cell = max_anchor.offset(3, 4)

    actual_value_cell = find_value_cell(sheet, ["actual sales", "reported sales"])
    num_quarters_cell = find_value_cell(sheet, ["num quarters used", "quarters used"])

    return RegressionContext(
        max_anchor=max_anchor,
        min_anchor=min_anchor,
        forecast_max_cell=forecast_max_cell,
        forecast_min_cell=forecast_min_cell,
        forecast_wo_sa_cell=forecast_wo_sa_cell,
        actual_value_cell=actual_value_cell,
        intercept_cell=intercept_cell,
        slope_cell=slope_cell,
        num_quarters_cell=num_quarters_cell,
        x_col=x_col,
        y_col=y_col,
        xy_rows=xy_rows,
    )


def extract_empirical_rows(
    workbook: Any,
    sheet: Any,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    context = build_empirical_context(sheet)
    if context.avg_penetration_cell is None:
        raise ValueError("Empirical Model missing avg penetration target cell")
    if context.penetration_series is None or not context.penetration_series.cells:
        raise ValueError("Empirical Model missing penetration history series")

    rows: list[dict[str, Any]] = []
    num_candidates = min(N_QUARTERS, len(context.penetration_series.cells))
    last_quarter_used = infer_last_quarter_used(context.penetration_series)

    for num_quarters in range(1, num_candidates + 1):
        start_cell = context.penetration_series.cells[-num_quarters]
        end_cell = context.penetration_series.cells[-1]
        avg_formula = (
            f"=AVERAGE(R{start_cell.row}C{start_cell.column}:R{end_cell.row}C{end_cell.column})"
        )
        set_formula2_r1c1(context.avg_penetration_cell, avg_formula)
        workbook.app.calculate()

        avg_penetration = number_or_blank(context.avg_penetration_cell.value)
        forecast_value = number_or_blank(
            context.estimated_total_sold_cell.value if context.estimated_total_sold_cell else None
        )
        actual_value = number_or_blank(context.reported_sales_cell.value if context.reported_sales_cell else None)
        forecast_max = number_or_blank(context.forecast_max_cell.value if context.forecast_max_cell else None)
        forecast_min = number_or_blank(context.forecast_min_cell.value if context.forecast_min_cell else None)
        quarterly_sales = number_or_blank(context.quarterly_sales_cell.value if context.quarterly_sales_cell else None)
        reported_sales = number_or_blank(context.reported_sales_cell.value if context.reported_sales_cell else None)
        growth_rate = number_or_blank(context.growth_rate_cell.value if context.growth_rate_cell else None)
        sales_captured = number_or_blank(context.sales_captured_cell.value if context.sales_captured_cell else None)
        range_width = safe_subtract(forecast_max, forecast_min)

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
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


def extract_regression_rows(
    workbook: Any,
    sheet: Any,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    context = build_regression_context(sheet)
    if not context.xy_rows:
        raise ValueError("Regression Model missing usable x/y data rows")

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    max_candidates = min(N_QUARTERS, len(context.xy_rows))

    for num_quarters in range(1, max_candidates + 1):
        start_row = context.xy_rows[-num_quarters]
        end_row = context.xy_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{context.y_col}:R{end_row}C{context.y_col},"
            f"R{start_row}C{context.x_col}:R{end_row}C{context.x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{context.y_col}:R{end_row}C{context.y_col},"
            f"R{start_row}C{context.x_col}:R{end_row}C{context.x_col})"
        )

        set_formula2_r1c1(context.intercept_cell, intercept_formula)
        set_formula2_r1c1(context.slope_cell, slope_formula)
        if context.num_quarters_cell is not None:
            context.num_quarters_cell.value = num_quarters
        workbook.app.calculate()

        intercept = number_or_blank(context.intercept_cell.value)
        slope = number_or_blank(context.slope_cell.value)
        forecast_value = number_or_blank(context.forecast_wo_sa_cell.value if context.forecast_wo_sa_cell else None)

        # Fallback forecast if explicit workbook cell is unavailable.
        if forecast_value == "" and intercept != "" and slope != "":
            last_x = to_float(sheet.range((end_row, context.x_col)).value)
            if last_x is not None:
                forecast_value = intercept + (slope * (last_x + 1))

        actual_value = number_or_blank(context.actual_value_cell.value if context.actual_value_cell else None)
        forecast_max = number_or_blank(context.forecast_max_cell.value if context.forecast_max_cell else None)
        forecast_min = number_or_blank(context.forecast_min_cell.value if context.forecast_min_cell else None)
        range_width = safe_subtract(forecast_max, forecast_min)

        signature = (
            normalize_signature_value(num_quarters),
            normalize_signature_value(intercept),
            normalize_signature_value(slope),
            normalize_signature_value(forecast_value),
            normalize_signature_value(forecast_max),
            normalize_signature_value(forecast_min),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
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


def autofit_columns(ws: Any, headers: Sequence[str]) -> None:
    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 48)


def write_output_sheet(ws: Any, headers: Sequence[str], rows: Sequence[dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(column, "") for column in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"

    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{max(2, ws.max_row)}"
    autofit_columns(ws, headers)


def write_output_workbook(
    destination: Path,
    empirical_rows: Sequence[dict[str, Any]],
    regression_rows: Sequence[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    write_output_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    workbook.save(destination)


def run_extraction() -> None:
    files_to_process, skipped_files = gather_input_files(input_dir)
    for skipped_file, reason in skipped_files:
        log(f"Skipped: {skipped_file.name} ({reason})")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    if files_to_process:
        app = None
        try:
            app = xw.App(visible=False, add_book=False)
            app.visible = False
            app.display_alerts = False
            app.screen_updating = False

            for file_path in files_to_process:
                log(f"Processing: {file_path.name}")
                workbook = None
                try:
                    workbook = app.books.open(str(file_path), update_links=False)
                    metadata = parse_model_metadata(file_path)

                    empirical_sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
                    if empirical_sheet is None:
                        log(f"Skipped: {file_path.name} (missing sheet: Empirical Model)")
                    else:
                        empirical_rows.extend(
                            extract_empirical_rows(workbook, empirical_sheet, metadata, file_path.name)
                        )

                    regression_sheet = get_sheet_case_insensitive(workbook, "Regression Model")
                    if regression_sheet is None:
                        log(f"Skipped: {file_path.name} (missing sheet: Regression Model)")
                    else:
                        regression_rows.extend(
                            extract_regression_rows(workbook, regression_sheet, metadata, file_path.name)
                        )

                    processed_count += 1
                    log(f"Processed: {file_path.name}")
                except Exception as exc:
                    log(f"Skipped: {file_path.name} (processing error: {exc})")
                finally:
                    if workbook is not None:
                        safe_close_workbook(workbook)
        finally:
            if app is not None:
                app.quit()

    output_path = build_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    log(f"Output path: {output_path}")
    log(f"Number of files processed: {processed_count}")
    log(f"Number of empirical rows: {len(empirical_rows)}")
    log(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run_extraction()
