#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "xlwings is required to run this script. Install it with: pip install xlwings"
    ) from exc

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----- Inputs -----
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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
    "sept": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}

FILENAME_PATTERN = re.compile(
    r".*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
    r"(?P<period>(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]{3,4})(?P<year>\d{4}))"
    r"(?:[_\-\s\.].*)?$",
    flags=re.IGNORECASE,
)

RowDict = Dict[str, Any]


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def to_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def number_or_blank(value: Any) -> Any:
    num = to_float(value)
    return num if num is not None else ""


def parse_model_meta(file_name: str) -> Optional[ModelMeta]:
    """
    Example:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    -> ticker=AORT, model_period=MidJan_2026, model_date=2026-01-15
    """
    match = FILENAME_PATTERN.match(file_name)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    bucket = match.group("bucket").title()
    month_token = match.group("month").title()
    year = int(match.group("year"))
    month_number = MONTH_MAP.get(month_token.lower())
    if month_number is None:
        return None

    day = DAY_BY_PERIOD[bucket.lower()]
    model_period = f"{bucket}{month_token}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return ModelMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def build_output_path(input_folder: Path, target_dir: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    wanted = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == wanted:
            return sheet
    return None


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = used.options(ndim=2).value
    if not values:
        return None

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return used.row + r_idx, used.column + c_idx
    return None


def build_header_offset_map(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    row_window: int = 2,
    col_window: int = 28,
) -> Dict[str, int]:
    start_row = max(1, anchor_row - row_window)
    end_row = anchor_row + row_window
    start_col = max(1, anchor_col - col_window)
    end_col = anchor_col + col_window

    values = sheet.range((start_row, start_col), (end_row, end_col)).options(ndim=2).value
    offsets: Dict[str, int] = {}
    for row in values:
        for c_idx, value in enumerate(row):
            key = normalize_text(value)
            if not key:
                continue
            abs_col = start_col + c_idx
            offsets.setdefault(key, abs_col - anchor_col)
    return offsets


def find_offset(
    header_offsets: Dict[str, int],
    aliases: Sequence[str],
    default: Optional[int] = None,
) -> Optional[int]:
    for alias in aliases:
        wanted = normalize_text(alias)
        for header, offset in header_offsets.items():
            if wanted in header:
                return offset
    return default


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    """
    Prefer Formula2R1C1 when available; fallback to formula2/formula.
    """
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def close_workbook_without_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
    except TypeError:
        try:
            wb.api.Close(SaveChanges=False)
        except Exception:
            wb.close()


def cell_from_block(
    block: List[List[Any]],
    block_top_row: int,
    block_left_col: int,
    abs_row: int,
    abs_col: int,
) -> Any:
    r_idx = abs_row - block_top_row
    c_idx = abs_col - block_left_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(block):
        return None
    row = block[r_idx]
    if c_idx >= len(row):
        return None
    return row[c_idx]


def value_from_anchor_offset(
    block: List[List[Any]],
    block_top_row: int,
    block_left_col: int,
    abs_row: int,
    anchor_col: int,
    offset: Optional[int],
) -> Any:
    if offset is None:
        return None
    return cell_from_block(
        block=block,
        block_top_row=block_top_row,
        block_left_col=block_left_col,
        abs_row=abs_row,
        abs_col=anchor_col + offset,
    )


def build_history_rows(sheet: xw.Sheet, value_col: int, anchor_row: int) -> List[int]:
    start_row = max(1, anchor_row - 300)
    end_row = anchor_row - 1
    if end_row < start_row:
        return []

    values = sheet.range((start_row, value_col), (end_row, value_col)).options(ndim=2).value
    rows: List[int] = []
    for idx, row in enumerate(values):
        if to_float(row[0]) is not None:
            rows.append(start_row + idx)
    return rows


def extract_empirical_candidates(
    wb: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[RowDict]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print("  - skipped empirical: missing sheet 'Empirical Model'")
        return []

    anchor = find_max_anchor(sheet)
    if not anchor:
        print("  - skipped empirical: no 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    n_quarters = 10
    header_offsets = build_header_offset_map(sheet, anchor_row, anchor_col)

    # Anchor-relative offsets; defaults are from the current template family.
    min_offset = find_offset(header_offsets, ["min", "minimum"], default=1)
    forecast_offset = find_offset(
        header_offsets,
        ["estimated total sold", "estimate total sold", "tot fcst", "forecast"],
        default=-2,
    )
    actual_offset = find_offset(header_offsets, ["reported sales", "actual"], default=-1)
    num_quarters_offset = find_offset(
        header_offsets,
        ["num quarters used", "num quarters", "n quarters", "quarters used"],
        default=None,
    )
    last_quarter_offset = find_offset(
        header_offsets,
        ["last quarter used", "last quarter"],
        default=None,
    )
    penetration_offset = find_offset(
        header_offsets,
        ["avg penetration", "penetration pct", "penetration"],
        default=-8,
    )
    quarterly_sales_offset = find_offset(
        header_offsets,
        ["quarterly sales", "qtr sales"],
        default=None,
    )
    reported_sales_offset = find_offset(
        header_offsets,
        ["reported sales"],
        default=actual_offset,
    )
    growth_offset = find_offset(
        header_offsets,
        ["growth rate", "growth %", "growth pct"],
        default=None,
    )
    captured_offset = find_offset(
        header_offsets,
        ["sales captured in db", "captured in db", "capture in db"],
        default=None,
    )

    penetration_col = anchor_col + (penetration_offset if penetration_offset is not None else -8)
    history_rows = build_history_rows(sheet, penetration_col, anchor_row)

    calc_col = anchor_col + 35
    calc_row_start = anchor_row + 220
    formulas_written = 0
    for idx in range(n_quarters):
        n_used = idx + 1
        if len(history_rows) < n_used:
            continue
        start_hist_row = history_rows[-n_used]
        end_hist_row = history_rows[-1]
        target_cell = sheet.range((calc_row_start + idx, calc_col))
        set_formula2_r1c1(
            target_cell,
            f"=AVERAGE(R{start_hist_row}C{penetration_col}:R{end_hist_row}C{penetration_col})",
        )
        formulas_written += 1

    avg_penetration_values: List[Any] = [""] * n_quarters
    if formulas_written:
        wb.app.calculate()
        calc_values = sheet.range(
            (calc_row_start, calc_col),
            (calc_row_start + n_quarters - 1, calc_col),
        ).options(ndim=2).value
        for idx, row in enumerate(calc_values):
            avg_penetration_values[idx] = row[0]

    block_top_row = anchor_row + 1
    block_bottom_row = anchor_row + n_quarters
    block_left_col = max(1, anchor_col - 28)
    block_right_col = anchor_col + 28
    block = sheet.range(
        (block_top_row, block_left_col),
        (block_bottom_row, block_right_col),
    ).options(ndim=2).value

    rows: List[RowDict] = []
    for idx in range(n_quarters):
        current_row = anchor_row + idx + 1
        default_n_quarters = idx + 1

        num_quarters_used = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=num_quarters_offset,
        )
        if to_float(num_quarters_used) is None:
            num_quarters_used = default_n_quarters
        num_quarters_used = int(float(num_quarters_used))

        forecast_value = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=forecast_offset,
        )
        actual_value = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=actual_offset,
        )
        forecast_max = cell_from_block(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            abs_col=anchor_col,
        )
        forecast_min = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=min_offset,
        )

        if all(v in (None, "") for v in [forecast_value, actual_value, forecast_max, forecast_min]):
            continue

        forecast_max_num = to_float(forecast_max)
        forecast_min_num = to_float(forecast_min)
        range_width: Any = ""
        if forecast_max_num is not None and forecast_min_num is not None:
            range_width = forecast_max_num - forecast_min_num

        avg_penetration_pct = avg_penetration_values[idx]
        if avg_penetration_pct in (None, ""):
            avg_penetration_pct = value_from_anchor_offset(
                block=block,
                block_top_row=block_top_row,
                block_left_col=block_left_col,
                abs_row=current_row,
                anchor_col=anchor_col,
                offset=penetration_offset,
            )

        row_data: RowDict = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": number_or_blank(avg_penetration_pct),
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": value_from_anchor_offset(
                block=block,
                block_top_row=block_top_row,
                block_left_col=block_left_col,
                abs_row=current_row,
                anchor_col=anchor_col,
                offset=last_quarter_offset,
            ),
            "forecast_value": number_or_blank(forecast_value),
            "actual_value": number_or_blank(actual_value),
            "forecast_max": number_or_blank(forecast_max),
            "forecast_min": number_or_blank(forecast_min),
            "range_width": number_or_blank(range_width),
            "avg_penetration_pct": number_or_blank(avg_penetration_pct),
            "quarterly_sales": number_or_blank(
                value_from_anchor_offset(
                    block=block,
                    block_top_row=block_top_row,
                    block_left_col=block_left_col,
                    abs_row=current_row,
                    anchor_col=anchor_col,
                    offset=quarterly_sales_offset,
                )
            ),
            "reported_sales": number_or_blank(
                value_from_anchor_offset(
                    block=block,
                    block_top_row=block_top_row,
                    block_left_col=block_left_col,
                    abs_row=current_row,
                    anchor_col=anchor_col,
                    offset=reported_sales_offset,
                )
            ),
            "growth_rate_pct": number_or_blank(
                value_from_anchor_offset(
                    block=block,
                    block_top_row=block_top_row,
                    block_left_col=block_left_col,
                    abs_row=current_row,
                    anchor_col=anchor_col,
                    offset=growth_offset,
                )
            ),
            "sales_captured_in_db_pct": number_or_blank(
                value_from_anchor_offset(
                    block=block,
                    block_top_row=block_top_row,
                    block_left_col=block_left_col,
                    abs_row=current_row,
                    anchor_col=anchor_col,
                    offset=captured_offset,
                )
            ),
            "source_file": source_file,
        }
        rows.append(row_data)

    return rows


def extract_regression_candidates(
    wb: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[RowDict]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print("  - skipped regression: missing sheet 'Regression Model'")
        return []

    anchor = find_max_anchor(sheet)
    if not anchor:
        print("  - skipped regression: no 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    # Required offsets from spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    n_quarters = 10

    header_offsets = build_header_offset_map(sheet, anchor_row, anchor_col)
    num_quarters_offset = find_offset(
        header_offsets,
        ["num quarters used", "num quarters", "n quarters", "quarters used"],
        default=None,
    )
    forecast_offset = find_offset(
        header_offsets,
        ["tot fcst w/o sa", "tot fcst wo sa", "forecast total without sa", "tot fcst"],
        default=-2,
    )
    max_offset = find_offset(header_offsets, ["max"], default=0)
    min_offset = find_offset(header_offsets, ["min", "minimum"], default=1)
    actual_offset = find_offset(
        header_offsets,
        ["actual", "reported sales", "total sold actual"],
        default=None,
    )

    history_start = max(1, anchor_row - 300)
    history_end = anchor_row - 1
    x_values = sheet.range((history_start, x_col), (history_end, x_col)).options(ndim=2).value
    y_values = sheet.range((history_start, y_col), (history_end, y_col)).options(ndim=2).value
    valid_history_rows: List[int] = []
    for idx in range(len(x_values)):
        x_val = to_float(x_values[idx][0])
        y_val = to_float(y_values[idx][0])
        if x_val is not None and y_val is not None:
            valid_history_rows.append(history_start + idx)

    calc_row_start = anchor_row + 260
    calc_intercept_col = anchor_col + 35
    calc_slope_col = anchor_col + 36
    stats_by_n: Dict[int, Tuple[Any, Any]] = {}
    n_values = [n for n in range(2, n_quarters + 1) if len(valid_history_rows) >= n]

    for idx, n_used in enumerate(n_values):
        start_row = valid_history_rows[-n_used]
        end_row = valid_history_rows[-1]
        intercept_cell = sheet.range((calc_row_start + idx, calc_intercept_col))
        slope_cell = sheet.range((calc_row_start + idx, calc_slope_col))
        set_formula2_r1c1(
            intercept_cell,
            (
                f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            ),
        )
        set_formula2_r1c1(
            slope_cell,
            (
                f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            ),
        )

    if n_values:
        wb.app.calculate()
        intercept_values = sheet.range(
            (calc_row_start, calc_intercept_col),
            (calc_row_start + len(n_values) - 1, calc_intercept_col),
        ).options(ndim=2).value
        slope_values = sheet.range(
            (calc_row_start, calc_slope_col),
            (calc_row_start + len(n_values) - 1, calc_slope_col),
        ).options(ndim=2).value
        for idx, n_used in enumerate(n_values):
            stats_by_n[n_used] = (intercept_values[idx][0], slope_values[idx][0])

    block_top_row = anchor_row + 1
    block_bottom_row = anchor_row + n_quarters
    block_left_col = max(1, anchor_col - 28)
    block_right_col = anchor_col + 28
    block = sheet.range(
        (block_top_row, block_left_col),
        (block_bottom_row, block_right_col),
    ).options(ndim=2).value

    rows: List[RowDict] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(n_quarters):
        current_row = anchor_row + idx + 1
        default_n_quarters = idx + 1

        n_candidate = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=num_quarters_offset,
        )
        if to_float(n_candidate) is None:
            n_candidate = default_n_quarters
        num_quarters_used = int(float(n_candidate))

        forecast_value = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=forecast_offset,
        )
        forecast_max = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=max_offset,
        )
        forecast_min = value_from_anchor_offset(
            block=block,
            block_top_row=block_top_row,
            block_left_col=block_left_col,
            abs_row=current_row,
            anchor_col=anchor_col,
            offset=min_offset,
        )
        actual_value: Any = ""
        if actual_offset is not None:
            actual_value = value_from_anchor_offset(
                block=block,
                block_top_row=block_top_row,
                block_left_col=block_left_col,
                abs_row=current_row,
                anchor_col=anchor_col,
                offset=actual_offset,
            )

        if all(v in (None, "") for v in [forecast_value, forecast_max, forecast_min]):
            continue

        intercept, slope = stats_by_n.get(num_quarters_used, ("", ""))
        forecast_max_num = to_float(forecast_max)
        forecast_min_num = to_float(forecast_min)
        range_width: Any = ""
        if forecast_max_num is not None and forecast_min_num is not None:
            range_width = forecast_max_num - forecast_min_num

        signature = (
            num_quarters_used,
            number_or_blank(forecast_value),
            number_or_blank(forecast_max),
            number_or_blank(forecast_min),
            number_or_blank(intercept),
            number_or_blank(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            # Avoid duplicate final rows emitted by some model sheets.
            continue
        previous_signature = signature

        row_data: RowDict = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": number_or_blank(forecast_value),
            "actual_value": number_or_blank(actual_value),
            "forecast_max": number_or_blank(forecast_max),
            "forecast_min": number_or_blank(forecast_min),
            "range_width": number_or_blank(range_width),
            "intercept": number_or_blank(intercept),
            "slope": number_or_blank(slope),
            "source_file": source_file,
        }
        rows.append(row_data)

    return rows


def write_sheet(ws: Any, rows: List[RowDict], columns: Sequence[str]) -> None:
    ws.append(list(columns))
    for row_data in rows:
        ws.append([row_data.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for value_tuple in ws.iter_rows(
            min_row=2,
            max_row=ws.max_row,
            min_col=col_idx,
            max_col=col_idx,
            values_only=True,
        ):
            value = value_tuple[0]
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    out_path: Path,
    empirical_rows: List[RowDict],
    regression_rows: List[RowDict],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_empirical = wb.create_sheet("empirical_candidates")
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, empirical_rows, EMPIRICAL_COLUMNS)
    write_sheet(ws_regression, regression_rows, REGRESSION_COLUMNS)
    wb.save(out_path)


def main() -> int:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input directory does not exist or is not a directory: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = build_output_path(input_dir, output_dir)

    empirical_rows: List[RowDict] = []
    regression_rows: List[RowDict] = []
    processed_files = 0

    excel_app = xw.App(visible=False, add_book=False)
    excel_app.display_alerts = False
    excel_app.screen_updating = False
    try:
        excel_app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"SKIPPED {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"SKIPPED {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED {file_path.name}: not .xlsx")
                continue

            model_meta = parse_model_meta(file_path.name)
            if model_meta is None:
                print(f"SKIPPED {file_path.name}: filename does not match expected pattern")
                continue

            print(f"PROCESSING {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = excel_app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_candidates(
                        wb=wb,
                        meta=model_meta,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_candidates(
                        wb=wb,
                        meta=model_meta,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: processing error -> {exc}")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        excel_app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT {output_path}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
