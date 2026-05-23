#!/usr/bin/env python3
"""Extract empirical/regression parameter candidates from Excel model files.

This script processes all `.xlsx` files in `input_dir` and writes one output
workbook in `output_dir` containing:
  - empirical_candidates
  - regression_candidates

Runtime is optimized by:
  - using one hidden Excel app for the full run
  - opening each source workbook exactly once
  - processing both model sheets while that workbook is open
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


# Required top-level directories.
input_dir = Path("input")
output_dir = Path("output")


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


EMPIRICAL_FALLBACK_OFFSETS = {
    # Offsets are relative to the "max" anchor column.
    "num_quarters_used": -7,
    "last_quarter_used": -8,
    "forecast_value": -3,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -5,
    "quarterly_sales": -4,
    "reported_sales": -2,
    "growth_rate_pct": -9,
    "sales_captured_in_db_pct": -10,
}

REGRESSION_FALLBACK_OFFSETS = {
    # Offsets are relative to the "max" anchor column.
    "num_quarters_used": -8,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s/_\-]+", " ", text)
    return text


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
    try:
        return float(text)
    except ValueError:
        return None


def has_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str) and not value.strip():
        return False
    return True


def safe_close_workbook(wb: xw.Book) -> None:
    """Close source workbook without saving; fallback across xlwings variants."""
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        # Last-resort swallow to avoid masking upstream processing.
        pass


def find_unique_output_path(input_folder: Path, out_dir: Path) -> Path:
    folder_name = input_folder.resolve().name
    base_name = f"{folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def parse_file_metadata(file_name: str) -> FileMetadata:
    """Parse ticker/model_period/model_date/model from source filename."""
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    period_token = ""
    if len(parts) >= 2:
        ticker = parts[1].upper()
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_token,
        re.IGNORECASE,
    )

    day_by_window = {"early": 5, "mid": 15, "late": 25}
    month_by_token = {
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

    model_period = ""
    model_date = ""

    if match:
        window = match.group(1).capitalize()
        month_token = match.group(2)[:3].lower()
        year = int(match.group(3))
        month_num = month_by_token.get(month_token)
        if month_num:
            day = day_by_window[window.lower()]
            model_period = f"{window}{match.group(2)[:3].capitalize()}_{year}"
            model_date = date(year, month_num, day).isoformat()

    if not model_period:
        model_period = period_token.replace(" ", "_")
    if not model_date:
        model_date = ""
    if not ticker and parts:
        ticker = parts[0].upper()

    model = f"{ticker}_{model_period}".strip("_")
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def values_to_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_max_anchor(sheet: xw.Sheet) -> Tuple[int, int]:
    used = sheet.used_range
    matrix = values_to_matrix(used.value)
    start_row = used.row
    start_col = used.column

    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return start_row + r_idx, start_col + c_idx

    raise ValueError(f'Could not find "max" anchor on sheet "{sheet.name}"')


def discover_offsets(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    aliases_by_key: Dict[str, Sequence[str]],
    *,
    row_window: int = 3,
    col_window: int = 35,
) -> Dict[str, int]:
    row_start = max(1, anchor_row - row_window)
    row_end = anchor_row + row_window
    col_start = max(1, anchor_col - col_window)
    col_end = anchor_col + col_window

    region = sheet.range((row_start, col_start), (row_end, col_end))
    matrix = values_to_matrix(region.value)

    candidates: List[Tuple[str, int, int]] = []
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            normalized = normalize_text(value)
            if normalized:
                abs_row = row_start + r_idx
                abs_col = col_start + c_idx
                candidates.append((normalized, abs_row, abs_col))

    offsets: Dict[str, int] = {}
    for key, aliases in aliases_by_key.items():
        best_match: Optional[Tuple[int, int]] = None
        for text, abs_row, abs_col in candidates:
            if any(alias in text for alias in aliases):
                # Weight row distance slightly to prioritize same-row headers.
                score = abs(abs_col - anchor_col) + (2 * abs(abs_row - anchor_row))
                if best_match is None or score < best_match[0]:
                    best_match = (score, abs_col)
        if best_match:
            offsets[key] = best_match[1] - anchor_col
    return offsets


def get_cell_value_by_offset(
    sheet: xw.Sheet,
    row: int,
    anchor_col: int,
    offsets: Dict[str, int],
    key: str,
    fallback_offsets: Dict[str, int],
) -> Any:
    offset = offsets.get(key, fallback_offsets.get(key))
    if offset is None:
        return None
    return sheet.cells(row, anchor_col + offset).value


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    """Prefer `.formula2` while keeping compatibility with COM variants."""
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


def format_sheet(ws, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        max_width = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            text = "" if value is None else str(value)
            max_width = max(max_width, len(text))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_width + 2, 11), 48)


def find_numeric_history_rows(
    sheet: xw.Sheet,
    end_row: int,
    x_col: int,
    y_col: int,
    *,
    max_rows: int = 80,
) -> List[int]:
    rows: List[int] = []
    row = end_row
    while row >= 1 and len(rows) < max_rows:
        x_val = to_float(sheet.cells(row, x_col).value)
        y_val = to_float(sheet.cells(row, y_col).value)
        if x_val is None or y_val is None:
            if rows:
                break
            row -= 1
            continue
        rows.append(row)
        row -= 1

    rows.reverse()
    return rows


def process_empirical_model(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction in {source_file}: missing sheet 'Empirical Model'")
        return []

    try:
        anchor_row, anchor_col = find_max_anchor(sheet)
    except Exception as exc:
        print(f"Skipped empirical extraction in {source_file}: {exc}")
        return []

    empirical_aliases = {
        "num_quarters_used": ("num quarters", "n quarters", "quarters used"),
        "last_quarter_used": ("last quarter", "last qtr"),
        "forecast_value": ("estimated total sold", "tot fcst", "forecast"),
        "actual_value": ("actual", "reported sales", "actual sales"),
        "forecast_max": ("max",),
        "forecast_min": ("min",),
        "avg_penetration_pct": ("avg penetration", "average penetration", "penetration"),
        "quarterly_sales": ("quarterly sales", "qtr sales", "quarter sales"),
        "reported_sales": ("reported sales", "reported"),
        "growth_rate_pct": ("growth rate", "growth"),
        "sales_captured_in_db_pct": ("captured in db", "sales captured", "db pct"),
    }
    offsets = discover_offsets(sheet, anchor_row, anchor_col, empirical_aliases)

    # Build temporary average-penetration formulas in scratch columns.
    # This keeps extraction fast and avoids letter-based formula generation.
    helper_penetration_col = anchor_col + 40
    helper_avg_col = anchor_col + 41

    # Historical columns for ratio calculation.
    quarterly_offset = offsets.get("quarterly_sales", EMPIRICAL_FALLBACK_OFFSETS["quarterly_sales"])
    reported_offset = offsets.get("reported_sales", EMPIRICAL_FALLBACK_OFFSETS["reported_sales"])
    quarterly_col = anchor_col + quarterly_offset
    reported_col = anchor_col + reported_offset

    history_rows = find_numeric_history_rows(
        sheet,
        end_row=anchor_row - 1,
        x_col=quarterly_col,
        y_col=reported_col,
    )

    formula_updates = 0
    for row in history_rows:
        penetration_cell = sheet.cells(row, helper_penetration_col)
        ratio_formula = (
            f"=IFERROR(R{row}C{reported_col}/R{row}C{quarterly_col},NA())"
        )
        set_formula2_r1c1(penetration_cell, ratio_formula)
        formula_updates += 1

    n_quarters = 10
    for i in range(1, n_quarters + 1):
        candidate_row = anchor_row + i
        if history_rows:
            end_row = history_rows[-1]
            start_row = max(history_rows[0], end_row - i + 1)
            avg_formula = (
                f"=IFERROR(AVERAGE(R{start_row}C{helper_penetration_col}:"
                f"R{end_row}C{helper_penetration_col}),NA())"
            )
            set_formula2_r1c1(sheet.cells(candidate_row, helper_avg_col), avg_formula)
            formula_updates += 1

    if formula_updates:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(1, n_quarters + 1):
        row_num = anchor_row + i
        num_quarters_raw = get_cell_value_by_offset(
            sheet,
            row_num,
            anchor_col,
            offsets,
            "num_quarters_used",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        num_quarters_used = int(to_float(num_quarters_raw) or i)

        last_quarter_used = get_cell_value_by_offset(
            sheet,
            row_num,
            anchor_col,
            offsets,
            "last_quarter_used",
            EMPIRICAL_FALLBACK_OFFSETS,
        )
        forecast_value = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "forecast_value",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        actual_value = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "actual_value",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        forecast_max = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "forecast_max",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        forecast_min = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "forecast_min",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        avg_penetration_pct = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "avg_penetration_pct",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        if avg_penetration_pct is None:
            avg_penetration_pct = to_float(sheet.cells(row_num, helper_avg_col).value)

        quarterly_sales = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "quarterly_sales",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        reported_sales = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "reported_sales",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        growth_rate_pct = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "growth_rate_pct",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )
        sales_captured_in_db_pct = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "sales_captured_in_db_pct",
                EMPIRICAL_FALLBACK_OFFSETS,
            )
        )

        if forecast_value is None and avg_penetration_pct is not None and quarterly_sales is not None:
            forecast_value = avg_penetration_pct * quarterly_sales

        if actual_value is None:
            actual_value = reported_sales

        if not any(
            has_value(value)
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
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
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


def process_regression_model(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction in {source_file}: missing sheet 'Regression Model'")
        return []

    try:
        anchor_row, anchor_col = find_max_anchor(sheet)
    except Exception as exc:
        print(f"Skipped regression extraction in {source_file}: {exc}")
        return []

    # Required by prompt.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_rows = find_numeric_history_rows(
        sheet,
        end_row=anchor_row - 1,
        x_col=x_col,
        y_col=y_col,
    )
    if not history_rows:
        print(f"Skipped regression extraction in {source_file}: no numeric history for x/y")
        return []

    regression_aliases = {
        "num_quarters_used": ("num quarters", "n quarters", "quarters used"),
        "forecast_value": ("tot fcst w/o sa", "tot fcst wo sa", "forecast"),
        "actual_value": ("actual", "actual sales", "reported sales"),
        "forecast_max": ("max",),
        "forecast_min": ("min",),
    }
    offsets = discover_offsets(sheet, anchor_row, anchor_col, regression_aliases)

    helper_intercept_col = anchor_col + 45
    helper_slope_col = anchor_col + 46
    helper_forecast_col = anchor_col + 47

    forecast_x = to_float(sheet.cells(anchor_row, x_col).value)
    if forecast_x is None:
        forecast_x = to_float(sheet.cells(history_rows[-1], x_col).value) or 0.0

    candidate_count = len(history_rows)
    formula_updates = 0
    for i in range(1, candidate_count + 1):
        candidate_row = anchor_row + i
        start_row = history_rows[-i]
        end_row = history_rows[-1]

        intercept_formula = (
            f"=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),NA())"
        )
        slope_formula = (
            f"=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),NA())"
        )
        forecast_formula = (
            f"=IFERROR(R{candidate_row}C{helper_intercept_col}+"
            f"R{candidate_row}C{helper_slope_col}*{forecast_x},NA())"
        )

        set_formula2_r1c1(sheet.cells(candidate_row, helper_intercept_col), intercept_formula)
        set_formula2_r1c1(sheet.cells(candidate_row, helper_slope_col), slope_formula)
        set_formula2_r1c1(sheet.cells(candidate_row, helper_forecast_col), forecast_formula)
        formula_updates += 3

    if formula_updates:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(1, candidate_count + 1):
        row_num = anchor_row + i
        num_quarters_raw = get_cell_value_by_offset(
            sheet,
            row_num,
            anchor_col,
            offsets,
            "num_quarters_used",
            REGRESSION_FALLBACK_OFFSETS,
        )
        num_quarters_used = int(to_float(num_quarters_raw) or i)

        intercept = to_float(sheet.cells(row_num, helper_intercept_col).value)
        slope = to_float(sheet.cells(row_num, helper_slope_col).value)
        forecast_total_without_sa = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "forecast_value",
                REGRESSION_FALLBACK_OFFSETS,
            )
        )
        if forecast_total_without_sa is None:
            forecast_total_without_sa = to_float(sheet.cells(row_num, helper_forecast_col).value)

        actual_value = get_cell_value_by_offset(
            sheet,
            row_num,
            anchor_col,
            offsets,
            "actual_value",
            REGRESSION_FALLBACK_OFFSETS,
        )
        if not has_value(actual_value):
            actual_value = ""

        forecast_max = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "forecast_max",
                REGRESSION_FALLBACK_OFFSETS,
            )
        )
        forecast_min = to_float(
            get_cell_value_by_offset(
                sheet,
                row_num,
                anchor_col,
                offsets,
                "forecast_min",
                REGRESSION_FALLBACK_OFFSETS,
            )
        )
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            num_quarters_used,
            round(intercept or 0.0, 8),
            round(slope or 0.0, 8),
            round(forecast_total_without_sa or 0.0, 8),
            round(forecast_max or 0.0, 8),
            round(forecast_min or 0.0, 8),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        if not any(
            has_value(value)
            for value in (
                intercept,
                slope,
                forecast_total_without_sa,
                forecast_max,
                forecast_min,
            )
        ):
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_total_without_sa,
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


def list_input_files(source_dir: Path) -> Iterable[Path]:
    for path in sorted(source_dir.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary Excel lock file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue
        yield path


def write_output_workbook(
    path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    format_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    format_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    workbook.save(path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir.resolve()}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = find_unique_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in list_input_files(input_dir):
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                metadata = parse_file_metadata(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(
                    process_empirical_model(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_model(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path.resolve()}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
