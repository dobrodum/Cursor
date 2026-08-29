#!/usr/bin/env python3
"""Extract empirical/regression parameter candidates from model workbooks.

This script opens each source workbook once, processes both model sheets while
the workbook is open, and writes one consolidated output workbook.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"

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


def normalize_key(value: Any) -> str:
    """Normalize labels for robust matching."""
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def to_2d(values: Any) -> List[List[Any]]:
    """Coerce xlwings returned value into a 2D list."""
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values  # already 2D
    return [values]  # 1D row


def to_float(value: Any) -> Optional[float]:
    """Best-effort numeric conversion."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("%", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any, default: int) -> int:
    """Convert to int with default fallback."""
    converted = to_float(value)
    if converted is None:
        return default
    return int(round(converted))


def safe_subtract(left: Optional[float], right: Optional[float]) -> Optional[float]:
    """Subtract nullable numerics."""
    if left is None or right is None:
        return None
    return left - right


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    """Return sheet cell value safely."""
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def find_anchor_cell(sheet: xw.Sheet, target_text: str = "max") -> Optional[Tuple[int, int]]:
    """Find anchor cell coordinates by scanning used range once."""
    used = sheet.used_range
    values_2d = to_2d(used.value)
    if not values_2d:
        return None

    target = normalize_key(target_text)
    start_row = int(used.row)
    start_col = int(used.column)

    for row_idx, row_values in enumerate(values_2d):
        for col_idx, value in enumerate(row_values):
            if normalize_key(value) == target:
                return start_row + row_idx, start_col + col_idx
    return None


def build_header_offsets(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    used_start_col: int,
    used_col_count: int,
) -> Dict[str, int]:
    """Read one header row and map normalized labels to anchor-relative offsets."""
    end_col = used_start_col + used_col_count - 1
    header_values = sheet.range((anchor_row, used_start_col), (anchor_row, end_col)).value
    if not isinstance(header_values, list):
        header_values = [header_values]

    offsets: Dict[str, int] = {}
    for idx, value in enumerate(header_values):
        key = normalize_key(value)
        if key and key not in offsets:
            col = used_start_col + idx
            offsets[key] = col - anchor_col
    return offsets


def resolve_offset(
    offsets: Dict[str, int],
    aliases: Iterable[str],
    default: int,
) -> int:
    """Pick a column offset by alias list, else use default."""
    for alias in aliases:
        key = normalize_key(alias)
        if key in offsets:
            return offsets[key]
    return default


def resolve_optional_offset(
    offsets: Dict[str, int],
    aliases: Iterable[str],
) -> Optional[int]:
    """Pick a column offset by alias list, else None."""
    for alias in aliases:
        key = normalize_key(alias)
        if key in offsets:
            return offsets[key]
    return None


def numeric_signature(*values: Optional[float]) -> Tuple[Optional[float], ...]:
    """Rounded signature to avoid duplicate floating-point rows."""
    out: List[Optional[float]] = []
    for value in values:
        out.append(None if value is None else round(value, 10))
    return tuple(out)


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    """Parse ticker/model period/date from source filename."""
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) > 1 and parts[1] else "UNKNOWN"
    period_raw = parts[2] if len(parts) > 2 else ""
    period_raw = re.sub(r"_?send.*$", "", period_raw, flags=re.IGNORECASE).strip()

    period_match = re.search(
        r"(?P<window>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
        period_raw,
        flags=re.IGNORECASE,
    )

    model_period = period_raw or "UNKNOWN"
    model_date = ""

    if period_match:
        window = period_match.group("window").title()
        month_text = period_match.group("month")[:3].title()
        year = int(period_match.group("year"))

        month_lookup = {
            "Jan": 1,
            "Feb": 2,
            "Mar": 3,
            "Apr": 4,
            "May": 5,
            "Jun": 6,
            "Jul": 7,
            "Aug": 8,
            "Sep": 9,
            "Oct": 10,
            "Nov": 11,
            "Dec": 12,
        }
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}

        month_num = month_lookup.get(month_text)
        if month_num:
            model_period = f"{window}{month_text}_{year}"
            model_date = date(year, month_num, day_lookup[window]).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_root: Path, output_root: Path) -> Path:
    """Build unique output path with required naming convention."""
    base_name = f"{input_root.name}_PARAM"
    base_path = output_root / f"{base_name}.xlsx"
    if not base_path.exists():
        return base_path

    idx = 1
    while True:
        candidate = output_root / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def get_sheet_by_name(workbook: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    """Case-insensitive sheet lookup."""
    lowered_target = target_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == lowered_target:
            return sheet
    return None


def safe_close_source_workbook(workbook: xw.Book) -> None:
    """Close workbook without saving. Fallback if close(save=False) is unsupported."""
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
    except Exception:
        pass


def process_empirical_sheet(
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    """Extract empirical candidate rows using anchor-based offsets."""
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    used_start_col = int(used.column)
    used_col_count = int(used.columns.count)
    header_offsets = build_header_offsets(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        used_start_col=used_start_col,
        used_col_count=used_col_count,
    )

    offsets = {
        "num_quarters_used": resolve_offset(
            header_offsets,
            ["num_quarters_used", "num quarters used", "quarters used", "num qtrs"],
            default=-8,
        ),
        "last_quarter_used": resolve_offset(
            header_offsets,
            ["last_quarter_used", "last quarter used", "last qtr used", "last quarter"],
            default=-7,
        ),
        "forecast_value": resolve_offset(
            header_offsets,
            ["estimated total sold", "est total sold", "forecast_value", "forecast value"],
            default=-2,
        ),
        "actual_value": resolve_offset(
            header_offsets,
            ["reported_sales", "reported sales", "actual_sales", "actual value"],
            default=-1,
        ),
        "forecast_max": 0,
        "forecast_min": resolve_offset(
            header_offsets,
            ["forecast_min", "min"],
            default=1,
        ),
        "avg_penetration_pct": resolve_offset(
            header_offsets,
            ["avg_penetration_pct", "avg penetration pct", "avg penetration"],
            default=-4,
        ),
        "quarterly_sales": resolve_offset(
            header_offsets,
            ["quarterly_sales", "quarterly sales", "qtr sales"],
            default=-6,
        ),
        "reported_sales": resolve_offset(
            header_offsets,
            ["reported_sales", "reported sales", "actual_sales", "actual value"],
            default=-1,
        ),
        "growth_rate_pct": resolve_offset(
            header_offsets,
            ["growth_rate_pct", "growth rate pct", "growth rate"],
            default=-3,
        ),
        "sales_captured_in_db_pct": resolve_offset(
            header_offsets,
            [
                "sales_captured_in_db_pct",
                "sales captured in db pct",
                "sales captured in db",
            ],
            default=-5,
        ),
    }

    helper_avg_col = used_start_col + used_col_count + 2
    helper_forecast_col = helper_avg_col + 1
    available_hist_rows = anchor_row - 1

    rows: List[Dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n

        num_quarters_used = to_int(
            get_cell_value(sheet, row_idx, anchor_col + offsets["num_quarters_used"]),
            default=n,
        )
        num_quarters_used = max(1, min(num_quarters_used, max(1, available_hist_rows)))

        hist_start_row = anchor_row - num_quarters_used
        hist_end_row = anchor_row - 1

        quarterly_sales_col = anchor_col + offsets["quarterly_sales"]
        reported_sales_col = anchor_col + offsets["reported_sales"]

        avg_penetration_pct: Optional[float] = None
        forecast_value: Optional[float] = None

        if quarterly_sales_col > 0 and reported_sales_col > 0 and hist_start_row >= 1:
            avg_cell = sheet.cells(row_idx, helper_avg_col)
            forecast_cell = sheet.cells(row_idx, helper_forecast_col)

            # R1C1 + formula2 avoids any column-letter conversion.
            avg_cell.formula2 = (
                f"=IFERROR("
                f"SUM(R{hist_start_row}C{quarterly_sales_col}:R{hist_end_row}C{quarterly_sales_col})/"
                f"SUM(R{hist_start_row}C{reported_sales_col}:R{hist_end_row}C{reported_sales_col})"
                f",\"\")"
            )
            forecast_cell.formula2 = (
                f"=IFERROR(R{row_idx}C{quarterly_sales_col}/R{row_idx}C{helper_avg_col},\"\")"
            )

            sheet.book.app.calculate()
            avg_penetration_pct = to_float(avg_cell.value)
            forecast_value = to_float(forecast_cell.value)

        if avg_penetration_pct is None:
            avg_penetration_pct = to_float(
                get_cell_value(sheet, row_idx, anchor_col + offsets["avg_penetration_pct"])
            )
        if forecast_value is None:
            forecast_value = to_float(
                get_cell_value(sheet, row_idx, anchor_col + offsets["forecast_value"])
            )

        reported_sales = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["reported_sales"])
        )
        forecast_max = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["forecast_max"])
        )
        forecast_min = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["forecast_min"])
        )
        range_width = safe_subtract(forecast_max, forecast_min)

        last_quarter_used = get_cell_value(
            sheet, row_idx, anchor_col + offsets["last_quarter_used"]
        )
        quarterly_sales = to_float(get_cell_value(sheet, row_idx, quarterly_sales_col))
        growth_rate_pct = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["growth_rate_pct"])
        )
        sales_captured_in_db_pct = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["sales_captured_in_db_pct"])
        )

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
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


def process_regression_sheet(
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    """Extract regression candidate rows using anchor-based offsets."""
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    used_start_col = int(used.column)
    used_col_count = int(used.columns.count)
    header_offsets = build_header_offsets(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        used_start_col=used_start_col,
        used_col_count=used_col_count,
    )

    offsets = {
        "num_quarters_used": resolve_offset(
            header_offsets,
            ["num_quarters_used", "num quarters used", "quarters used", "num qtrs"],
            default=-8,
        ),
        "forecast_value": resolve_offset(
            header_offsets,
            [
                "tot fcst w/o sa",
                "tot fcst wo sa",
                "total forecast without sa",
                "forecast total without sa",
                "tot fcst without sa",
            ],
            default=-1,
        ),
        "forecast_max": 0,
        "forecast_min": resolve_offset(header_offsets, ["forecast_min", "min"], default=1),
        "actual_value": resolve_optional_offset(
            header_offsets, ["actual_value", "actual sales", "reported sales", "actual"]
        ),
    }

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    helper_intercept_col = used_start_col + used_col_count + 2
    helper_slope_col = helper_intercept_col + 1

    available_hist_rows = anchor_row - 1
    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for n in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n

        num_quarters_used = to_int(
            get_cell_value(sheet, row_idx, anchor_col + offsets["num_quarters_used"]),
            default=n,
        )
        num_quarters_used = max(1, min(num_quarters_used, max(1, available_hist_rows)))

        hist_start_row = anchor_row - num_quarters_used
        hist_end_row = anchor_row - 1

        intercept: Optional[float] = None
        slope: Optional[float] = None

        if x_col > 0 and y_col > 0 and hist_start_row >= 1:
            intercept_cell = sheet.cells(row_idx, helper_intercept_col)
            slope_cell = sheet.cells(row_idx, helper_slope_col)

            intercept_cell.formula2 = (
                f"=IFERROR(INTERCEPT("
                f"R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
                f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col}"
                f"),\"\")"
            )
            slope_cell.formula2 = (
                f"=IFERROR(SLOPE("
                f"R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
                f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col}"
                f"),\"\")"
            )

            sheet.book.app.calculate()
            intercept = to_float(intercept_cell.value)
            slope = to_float(slope_cell.value)

        forecast_value = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["forecast_value"])
        )
        if forecast_value is None and intercept is not None and slope is not None:
            x_value_for_row = to_float(get_cell_value(sheet, row_idx, x_col))
            if x_value_for_row is not None:
                forecast_value = intercept + (slope * x_value_for_row)

        forecast_max = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["forecast_max"])
        )
        forecast_min = to_float(
            get_cell_value(sheet, row_idx, anchor_col + offsets["forecast_min"])
        )
        range_width = safe_subtract(forecast_max, forecast_min)

        actual_value = ""
        actual_offset = offsets["actual_value"]
        if actual_offset is not None:
            actual_value = get_cell_value(sheet, row_idx, anchor_col + actual_offset)

        signature = numeric_signature(
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        if all(
            value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

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

    return rows


def write_output_sheet(
    worksheet,
    columns: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    """Write rows to worksheet with required formatting."""
    worksheet.append(columns)
    for row in rows:
        worksheet.append([row.get(column, "") for column in columns])

    for header_cell in worksheet[1]:
        header_cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 44)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    """Write both candidate sheets to one workbook."""
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_output_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def main() -> None:
    in_root = Path(input_dir).expanduser().resolve()
    out_root = Path(output_dir).expanduser().resolve()

    if not in_root.exists() or not in_root.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {in_root}")
    out_root.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(in_root, out_root)
    generated_prefix = f"{in_root.name}_PARAM"

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(in_root.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file.")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file.")
                continue
            if file_path.stem.startswith(generated_prefix):
                print(f"Skipped {file_path.name}: previously generated PARAM output.")
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)

                empirical_sheet = get_sheet_by_name(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name}: sheet not found.")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(empirical_sheet, metadata, file_path.name)
                    )

                regression_sheet = get_sheet_by_name(workbook, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression extraction for {file_path.name}: sheet not found.")
                else:
                    regression_rows.extend(
                        process_regression_sheet(regression_sheet, metadata, file_path.name)
                    )

                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed during processing ({exc}).")
            finally:
                if workbook is not None:
                    safe_close_source_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
