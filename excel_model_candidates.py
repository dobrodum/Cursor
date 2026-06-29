#!/usr/bin/env python3
"""Build empirical and regression candidate sheets from model workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = "./input"
output_dir = "./output"


EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"
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


# Fallback layout offsets relative to anchor_col ("max").
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "avg_penetration_pct": -6,
    "quarterly_sales": -5,
    "reported_sales": -4,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -2,
    "forecast_value": -1,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": -4,  # reported sales fallback
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -2,
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": -3,
}


EMPIRICAL_LABEL_KEYWORDS = {
    "num_quarters_used": ["num quarter", "quarters used", "n quarter"],
    "last_quarter_used": ["last quarter"],
    "avg_penetration_pct": ["avg penetration", "average penetration"],
    "quarterly_sales": ["quarterly sales"],
    "reported_sales": ["reported sales"],
    "growth_rate_pct": ["growth rate"],
    "sales_captured_in_db_pct": ["captured in db", "sales captured"],
    "forecast_value": ["estimated total sold", "estimate total sold", "forecast"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "actual_value": ["actual", "reported sales"],
}

REGRESSION_LABEL_KEYWORDS = {
    "num_quarters_used": ["num quarter", "quarters used", "n quarter"],
    "forecast_value": ["tot fcst w/o sa", "total forecast", "forecast"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "actual_value": ["actual"],
}


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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


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


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def subtract_if_numeric(a: Any, b: Any) -> Optional[float]:
    fa = to_float(a)
    fb = to_float(b)
    if fa is None or fb is None:
        return None
    return fa - fb


def is_effectively_empty(*values: Any) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return False
    return True


def parse_file_labels(file_path: Path) -> FileLabels:
    """
    Parse naming style similar to:
    MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1].strip() if len(parts) >= 2 else "UNKNOWN"
    period_segment = parts[2].strip() if len(parts) >= 3 else ""
    period_token = period_segment.split("_", 1)[0]

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    if not match:
        # Fallback if filename does not strictly follow expected pattern.
        model_period = period_token if period_token else "UnknownPeriod"
        return FileLabels(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    phase = match.group(1).title()  # Early/Mid/Late
    month_text = match.group(2)
    year = int(match.group(3))

    month_key = month_text[:3].lower()
    month_num = MONTH_MAP.get(month_key)
    if month_num is None:
        model_period = f"{phase}{month_text}_{year}"
        return FileLabels(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    day = DAY_MAP[phase.lower()]
    model_date = date(year, month_num, day).isoformat()
    model_period = f"{phase}{month_text[:3].title()}_{year}"
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def source_files(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary Excel file")
            continue
        yield file_path


def get_sheet_safe(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        # Continue to fallback close attempts.
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    top_row = used.row
    left_col = used.column
    wanted = anchor_text.strip().lower()

    for r_index, row in enumerate(values):
        for c_index, value in enumerate(row):
            if normalize_text(value) == wanted:
                return top_row + r_index, left_col + c_index
    return None


def detect_offsets_from_labels(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    label_keywords: Dict[str, List[str]],
    row_radius: int = 3,
    col_radius: int = 24,
) -> Dict[str, int]:
    start_row = max(1, anchor_row - row_radius)
    end_row = anchor_row + row_radius
    start_col = max(1, anchor_col - col_radius)
    end_col = anchor_col + col_radius

    values = to_2d(sheet.range((start_row, start_col), (end_row, end_col)).value)
    offsets: Dict[str, int] = {}

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            text = normalize_text(value)
            if not text:
                continue
            absolute_col = start_col + c_idx

            for field, needles in label_keywords.items():
                if field in offsets:
                    continue
                if any(needle in text for needle in needles):
                    offsets[field] = absolute_col - anchor_col
                    break
    return offsets


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if col < 1 or row < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def empirical_average_formula(num_quarters: int) -> str:
    """
    Average over a supporting penetration column immediately to the left.
    Written in R1C1 style with formula2 for speed and reliability.
    """
    if num_quarters <= 1:
        return '=IFERROR(RC[-1],"")'
    return f'=IFERROR(AVERAGE(R[-{num_quarters - 1}]C[-1]:RC[-1]),"")'


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped empirical extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor

    offsets = EMPIRICAL_FALLBACK_OFFSETS.copy()
    offsets.update(
        detect_offsets_from_labels(
            sheet=sheet,
            anchor_row=anchor_row,
            anchor_col=anchor_col,
            label_keywords=EMPIRICAL_LABEL_KEYWORDS,
        )
    )

    avg_col = anchor_col + offsets["avg_penetration_pct"]
    formulas_written = False
    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        if avg_col < 1:
            break
        try:
            sheet.cells(row, avg_col).formula2 = empirical_average_formula(n)
            formulas_written = True
        except Exception:
            # Continue extraction even if formula write fails in this workbook.
            pass

    if formulas_written:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    empty_streak = 0

    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        row_values = {
            "num_quarters_used": read_cell(sheet, row, anchor_col + offsets["num_quarters_used"]),
            "last_quarter_used": read_cell(sheet, row, anchor_col + offsets["last_quarter_used"]),
            "forecast_value": read_cell(sheet, row, anchor_col + offsets["forecast_value"]),
            "actual_value": read_cell(sheet, row, anchor_col + offsets["actual_value"]),
            "forecast_max": read_cell(sheet, row, anchor_col + offsets["forecast_max"]),
            "forecast_min": read_cell(sheet, row, anchor_col + offsets["forecast_min"]),
            "avg_penetration_pct": read_cell(
                sheet,
                row,
                anchor_col + offsets["avg_penetration_pct"],
            ),
            "quarterly_sales": read_cell(sheet, row, anchor_col + offsets["quarterly_sales"]),
            "reported_sales": read_cell(sheet, row, anchor_col + offsets["reported_sales"]),
            "growth_rate_pct": read_cell(sheet, row, anchor_col + offsets["growth_rate_pct"]),
            "sales_captured_in_db_pct": read_cell(
                sheet,
                row,
                anchor_col + offsets["sales_captured_in_db_pct"],
            ),
        }

        if is_effectively_empty(
            row_values["forecast_value"],
            row_values["forecast_max"],
            row_values["forecast_min"],
            row_values["avg_penetration_pct"],
        ):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        num_quarters_used = row_values["num_quarters_used"]
        if is_effectively_empty(num_quarters_used):
            num_quarters_used = n

        forecast_max = row_values["forecast_max"]
        forecast_min = row_values["forecast_min"]
        range_width = subtract_if_numeric(forecast_max, forecast_min)

        out = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": row_values["avg_penetration_pct"],
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": row_values["last_quarter_used"],
            "forecast_value": row_values["forecast_value"],  # estimated total sold
            "actual_value": row_values["actual_value"],  # reported sales
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": row_values["avg_penetration_pct"],
            "quarterly_sales": row_values["quarterly_sales"],
            "reported_sales": row_values["reported_sales"],
            "growth_rate_pct": row_values["growth_rate_pct"],
            "sales_captured_in_db_pct": row_values["sales_captured_in_db_pct"],
            "source_file": source_file,
        }
        rows.append(out)

    return rows


def build_intercept_formula(
    y_col: int,
    x_col: int,
    start_row: int,
    end_row: int,
) -> str:
    return (
        f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
        f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
    )


def build_slope_formula(
    y_col: int,
    x_col: int,
    start_row: int,
    end_row: int,
) -> str:
    return (
        f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
        f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
    )


def dedupe_key(*values: Any) -> Tuple[Any, ...]:
    def normalize(value: Any) -> Any:
        number = to_float(value)
        if number is not None:
            return round(number, 12)
        if isinstance(value, str):
            return value.strip()
        return value

    return tuple(normalize(v) for v in values)


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, anchor_text="max")
    if not anchor:
        print(f"Skipped regression extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor

    # Required anchor-driven columns from the existing logic.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    offsets = REGRESSION_FALLBACK_OFFSETS.copy()
    offsets.update(
        detect_offsets_from_labels(
            sheet=sheet,
            anchor_row=anchor_row,
            anchor_col=anchor_col,
            label_keywords=REGRESSION_LABEL_KEYWORDS,
        )
    )

    # Write regression formulas into temporary columns to avoid touching core model cells.
    temp_col_start = max(sheet.used_range.last_cell.column + 2, anchor_col + 4)
    intercept_col = temp_col_start
    slope_col = temp_col_start + 1

    formulas_written = False
    data_end_row = anchor_row - 1
    for n in range(1, N_QUARTERS + 1):
        target_row = anchor_row + n
        data_start_row = max(1, data_end_row - n + 1)
        if data_end_row <= data_start_row:
            continue

        try:
            sheet.cells(target_row, intercept_col).formula2 = build_intercept_formula(
                y_col=y_col,
                x_col=x_col,
                start_row=data_start_row,
                end_row=data_end_row,
            )
            sheet.cells(target_row, slope_col).formula2 = build_slope_formula(
                y_col=y_col,
                x_col=x_col,
                start_row=data_start_row,
                end_row=data_end_row,
            )
            formulas_written = True
        except Exception:
            pass

    if formulas_written:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Any, ...]] = None
    empty_streak = 0

    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        num_quarters_used = read_cell(sheet, row, anchor_col + offsets["num_quarters_used"])
        forecast_value = read_cell(sheet, row, anchor_col + offsets["forecast_value"])
        forecast_max = read_cell(sheet, row, anchor_col + offsets["forecast_max"])
        forecast_min = read_cell(sheet, row, anchor_col + offsets["forecast_min"])
        actual_value = read_cell(sheet, row, anchor_col + offsets["actual_value"])
        intercept = read_cell(sheet, row, intercept_col)
        slope = read_cell(sheet, row, slope_col)

        if is_effectively_empty(
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        ):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        if is_effectively_empty(num_quarters_used):
            num_quarters_used = n

        current_key = dedupe_key(
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )

        # Prevent duplicate final row when the model spills one repeated candidate.
        if n == N_QUARTERS and previous_key is not None and current_key == previous_key:
            continue

        previous_key = current_key
        range_width = subtract_if_numeric(forecast_max, forecast_min)

        out = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,  # TOT FCST w/o SA
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(out)

    return rows


def format_output_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        letter = get_column_letter(idx)
        for cell in ws[letter]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_len + 2, 60)


def write_output_workbook(
    target_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_emp = wb.create_sheet("empirical_candidates")
    ws_reg = wb.create_sheet("regression_candidates")

    format_output_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)
    format_output_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb.save(target_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_path=input_path, output_path=output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files(input_path):
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                labels = parse_file_labels(file_path)
                empirical_sheet = get_sheet_safe(wb, EMPIRICAL_MODEL_SHEET)
                regression_sheet = get_sheet_safe(wb, REGRESSION_MODEL_SHEET)

                if empirical_sheet is None and regression_sheet is None:
                    print(
                        f"Skipped {file_path.name}: missing both '{EMPIRICAL_MODEL_SHEET}' "
                        f"and '{REGRESSION_MODEL_SHEET}' sheets"
                    )
                    continue

                if empirical_sheet is None:
                    print(f"Skipped empirical in {file_path.name}: missing '{EMPIRICAL_MODEL_SHEET}'")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                if regression_sheet is None:
                    print(
                        f"Skipped regression in {file_path.name}: missing '{REGRESSION_MODEL_SHEET}'"
                    )
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(
        target_path=output_file,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
