#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from Excel workbooks.

This script scans input_dir for .xlsx files, skips temp files, and writes one
combined output workbook containing:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
input_dir = "/path/to/input"
output_dir = "/path/to/output"


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


# Anchor-based fallback offsets from the "max" header cell.
EMP_DEFAULT_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -3,  # estimated total sold
    "actual_value": -2,  # reported sales
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -5,
    "reported_sales": -2,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -1,
    "penetration": -9,
}

REG_DEFAULT_OFFSETS = {
    "num_quarters_used": -8,
    "forecast_value": -3,  # TOT FCST w/o SA
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


# Temp formula cells (still anchor-based) used only while source workbook is open.
EMP_FORMULA_COL_OFFSET = 20
REG_INTERCEPT_COL_OFFSET = 20
REG_SLOPE_COL_OFFSET = 21


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


MONTHS = {
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

PERIOD_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

FILENAME_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})",
    re.IGNORECASE,
)
PERIOD_PATTERN = re.compile(
    r"^(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})$",
    re.IGNORECASE,
)


def parse_file_labels(file_name: str) -> FileLabels:
    """Parse ticker/model period/date from file name."""
    stem = Path(file_name).stem
    match = FILENAME_PATTERN.search(stem)
    if not match:
        raise ValueError(
            "expected '<prefix> - <TICKER> - <Early|Mid|Late><Mon><YYYY>_...'"
        )

    ticker = match.group("ticker").upper()
    period_token = match.group("period")
    period_match = PERIOD_PATTERN.match(period_token)
    if not period_match:
        raise ValueError(f"invalid period token: {period_token}")

    phase_raw = period_match.group("phase")
    month_raw = period_match.group("month")
    year = period_match.group("year")

    month_key = month_raw.lower()
    if month_key not in MONTHS:
        raise ValueError(f"invalid month token: {month_raw}")

    phase_key = phase_raw.lower()
    day = PERIOD_DAY[phase_key]
    model_date = date(int(year), MONTHS[month_key], day).isoformat()

    phase = phase_raw[0].upper() + phase_raw[1:].lower()
    month = month_raw[0].upper() + month_raw[1:].lower()
    model_period = f"{phase}{month}_{year}"
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def choose_output_path(output_folder: Path, input_folder_name: str) -> Path:
    """Create {input_folder}_PARAM.xlsx, then .1/.2/etc when needed."""
    base = f"{input_folder_name}_PARAM"
    candidate = output_folder / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_folder / f"{base}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def maybe_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def maybe_int(value: Any) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def pct(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return (numerator / denominator) * 100.0


def subtract(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left - right


def normalize_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def safe_close_workbook(wb: Any) -> None:
    """Close source workbook with no save, with safe fallbacks."""
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
    except Exception:
        pass


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and isinstance(values[0], list):
        return values
    return [values]


def get_sheet_by_name(workbook: Any, target_name: str) -> Any | None:
    target = target_name.strip().lower()
    for sheet in workbook.sheets:
        if str(sheet.name).strip().lower() == target:
            return sheet
    return None


def find_anchor(sheet: Any, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    grid = normalize_2d(used.value)
    if not grid:
        return None

    start_row = used.row
    start_col = used.column
    needle = anchor_text.strip().lower()

    for r_idx, row_values in enumerate(grid):
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == needle:
                return start_row + r_idx, start_col + c_idx
    return None


def row_header_offsets(sheet: Any, anchor_row: int, anchor_col: int) -> list[tuple[str, int]]:
    """Scan one row and return (normalized_header, offset_from_anchor_col)."""
    last_col = sheet.used_range.last_cell.column
    row_values = sheet.range((anchor_row, 1), (anchor_row, last_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]

    offsets: list[tuple[str, int]] = []
    for col_idx, value in enumerate(row_values, start=1):
        if isinstance(value, str) and value.strip():
            offsets.append((normalize_header(value), col_idx - anchor_col))
    return offsets


def pick_offset(
    header_offsets: list[tuple[str, int]],
    aliases: Iterable[str],
    default_offset: int,
) -> int:
    """Pick nearest matching header offset; fallback to fixed offset."""
    aliases_norm = [normalize_header(a) for a in aliases]
    best_offset = None
    best_distance = None

    for header_text, offset in header_offsets:
        if any(alias in header_text for alias in aliases_norm):
            distance = abs(offset)
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_offset = offset

    if best_offset is None:
        return default_offset
    return best_offset


def empirical_offsets(header_offsets: list[tuple[str, int]]) -> dict[str, int]:
    return {
        "num_quarters_used": pick_offset(
            header_offsets,
            ("num quarters used", "num quarters", "quarters used", "n quarters"),
            EMP_DEFAULT_OFFSETS["num_quarters_used"],
        ),
        "last_quarter_used": pick_offset(
            header_offsets,
            ("last quarter used", "last quarter"),
            EMP_DEFAULT_OFFSETS["last_quarter_used"],
        ),
        "forecast_value": pick_offset(
            header_offsets,
            ("estimated total sold", "tot fcst", "forecast"),
            EMP_DEFAULT_OFFSETS["forecast_value"],
        ),
        "actual_value": pick_offset(
            header_offsets,
            ("reported sales", "actual sales", "actual"),
            EMP_DEFAULT_OFFSETS["actual_value"],
        ),
        "forecast_max": EMP_DEFAULT_OFFSETS["forecast_max"],
        "forecast_min": pick_offset(
            header_offsets,
            ("min",),
            EMP_DEFAULT_OFFSETS["forecast_min"],
        ),
        "quarterly_sales": pick_offset(
            header_offsets,
            ("quarterly sales", "quarter sales"),
            EMP_DEFAULT_OFFSETS["quarterly_sales"],
        ),
        "reported_sales": pick_offset(
            header_offsets,
            ("reported sales",),
            EMP_DEFAULT_OFFSETS["reported_sales"],
        ),
        "growth_rate_pct": pick_offset(
            header_offsets,
            ("growth rate pct", "growth rate"),
            EMP_DEFAULT_OFFSETS["growth_rate_pct"],
        ),
        "sales_captured_in_db_pct": pick_offset(
            header_offsets,
            ("sales captured in db pct", "sales captured in db"),
            EMP_DEFAULT_OFFSETS["sales_captured_in_db_pct"],
        ),
        "penetration": pick_offset(
            header_offsets,
            ("penetration pct", "penetration"),
            EMP_DEFAULT_OFFSETS["penetration"],
        ),
    }


def regression_offsets(header_offsets: list[tuple[str, int]]) -> dict[str, int]:
    return {
        "num_quarters_used": pick_offset(
            header_offsets,
            ("num quarters used", "num quarters", "quarters used", "n quarters"),
            REG_DEFAULT_OFFSETS["num_quarters_used"],
        ),
        "forecast_value": pick_offset(
            header_offsets,
            ("tot fcst w/o sa", "tot fcst without sa", "tot fcst", "forecast"),
            REG_DEFAULT_OFFSETS["forecast_value"],
        ),
        "actual_value": pick_offset(
            header_offsets,
            ("actual sales", "reported sales", "actual"),
            REG_DEFAULT_OFFSETS["actual_value"],
        ),
        "forecast_max": REG_DEFAULT_OFFSETS["forecast_max"],
        "forecast_min": pick_offset(
            header_offsets,
            ("min",),
            REG_DEFAULT_OFFSETS["forecast_min"],
        ),
    }


def extract_empirical_rows(sheet: Any, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, anchor_text="max")
    if anchor is None:
        print(f"  Skipped empirical extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    offsets = empirical_offsets(row_header_offsets(sheet, anchor_row, anchor_col))
    first_candidate_row = anchor_row + 1

    formula_cells: list[tuple[int, int]] = []
    row_info: list[tuple[int, int, int]] = []
    penetration_col = anchor_col + offsets["penetration"]
    formula_col = anchor_col + EMP_FORMULA_COL_OFFSET
    history_end_row = anchor_row - 1
    history_start_floor = max(1, history_end_row - N_QUARTERS + 1)

    for idx in range(N_QUARTERS):
        row = first_candidate_row + idx
        num_quarters = maybe_int(
            sheet.range((row, anchor_col + offsets["num_quarters_used"])).value
        )
        if num_quarters is None or num_quarters <= 0:
            num_quarters = idx + 1

        start_row = max(history_start_floor, history_end_row - num_quarters + 1)
        if history_end_row >= start_row:
            start_rel = start_row - row
            end_rel = history_end_row - row
            col_rel = penetration_col - formula_col
            formula = (
                f"=AVERAGE(R[{start_rel}]C[{col_rel}]:R[{end_rel}]C[{col_rel}])"
            )
            sheet.range((row, formula_col)).formula2 = formula
            formula_cells.append((row, formula_col))
        row_info.append((row, num_quarters, formula_col))

    if formula_cells:
        sheet.book.app.calculate()

    rows: list[dict[str, Any]] = []
    for row, num_quarters, avg_pen_col in row_info:
        forecast_value = maybe_float(
            sheet.range((row, anchor_col + offsets["forecast_value"])).value
        )
        actual_value = maybe_float(
            sheet.range((row, anchor_col + offsets["actual_value"])).value
        )
        forecast_max = maybe_float(
            sheet.range((row, anchor_col + offsets["forecast_max"])).value
        )
        forecast_min = maybe_float(
            sheet.range((row, anchor_col + offsets["forecast_min"])).value
        )
        avg_penetration_pct = maybe_float(sheet.range((row, avg_pen_col)).value)

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

        last_quarter_used = sheet.range(
            (row, anchor_col + offsets["last_quarter_used"])
        ).value
        quarterly_sales = maybe_float(
            sheet.range((row, anchor_col + offsets["quarterly_sales"])).value
        )
        reported_sales = maybe_float(
            sheet.range((row, anchor_col + offsets["reported_sales"])).value
        )
        growth_rate_pct = maybe_float(
            sheet.range((row, anchor_col + offsets["growth_rate_pct"])).value
        )
        if growth_rate_pct is None:
            if quarterly_sales not in (None, 0) and reported_sales is not None:
                growth_rate_pct = ((reported_sales / quarterly_sales) - 1.0) * 100.0

        sales_captured_in_db_pct = maybe_float(
            sheet.range((row, anchor_col + offsets["sales_captured_in_db_pct"])).value
        )
        if sales_captured_in_db_pct is None:
            sales_captured_in_db_pct = pct(quarterly_sales, reported_sales)

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(sheet: Any, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, anchor_text="max")
    if anchor is None:
        print(f"  Skipped regression extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    offsets = regression_offsets(row_header_offsets(sheet, anchor_row, anchor_col))
    first_candidate_row = anchor_row + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = anchor_row - 1
    history_start_floor = max(1, history_end_row - N_QUARTERS + 1)

    intercept_col = anchor_col + REG_INTERCEPT_COL_OFFSET
    slope_col = anchor_col + REG_SLOPE_COL_OFFSET

    row_info: list[tuple[int, int]] = []
    for idx in range(N_QUARTERS):
        row = first_candidate_row + idx
        num_quarters = maybe_int(
            sheet.range((row, anchor_col + offsets["num_quarters_used"])).value
        )
        if num_quarters is None or num_quarters <= 0:
            num_quarters = idx + 1

        start_row = max(history_start_floor, history_end_row - num_quarters + 1)
        if history_end_row >= start_row:
            start_rel = start_row - row
            end_rel = history_end_row - row
            y_rel_i = y_col - intercept_col
            x_rel_i = x_col - intercept_col
            y_rel_s = y_col - slope_col
            x_rel_s = x_col - slope_col

            sheet.range((row, intercept_col)).formula2 = (
                f"=INTERCEPT(R[{start_rel}]C[{y_rel_i}]:R[{end_rel}]C[{y_rel_i}],"
                f"R[{start_rel}]C[{x_rel_i}]:R[{end_rel}]C[{x_rel_i}])"
            )
            sheet.range((row, slope_col)).formula2 = (
                f"=SLOPE(R[{start_rel}]C[{y_rel_s}]:R[{end_rel}]C[{y_rel_s}],"
                f"R[{start_rel}]C[{x_rel_s}]:R[{end_rel}]C[{x_rel_s}])"
            )

        row_info.append((row, num_quarters))

    if row_info:
        sheet.book.app.calculate()

    rows: list[dict[str, Any]] = []
    prev_signature: tuple[Any, ...] | None = None

    for row, num_quarters in row_info:
        forecast_value = maybe_float(
            sheet.range((row, anchor_col + offsets["forecast_value"])).value
        )
        actual_value = maybe_float(
            sheet.range((row, anchor_col + offsets["actual_value"])).value
        )
        forecast_max = maybe_float(
            sheet.range((row, anchor_col + offsets["forecast_max"])).value
        )
        forecast_min = maybe_float(
            sheet.range((row, anchor_col + offsets["forecast_min"])).value
        )
        intercept_value = maybe_float(sheet.range((row, intercept_col)).value)
        slope_value = maybe_float(sheet.range((row, slope_col)).value)

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        ):
            continue

        signature = (
            num_quarters,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept_value,
            slope_value,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract(forecast_max, forecast_min),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    worksheet: Any,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    for col_idx, header in enumerate(columns, start=1):
        cell = worksheet.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    for row in rows:
        worksheet.append([row.get(col_name) for col_name in columns])

    worksheet.freeze_panes = "A2"
    max_row = max(worksheet.max_row, 1)
    max_col = len(columns)
    worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(
            60, max(12, max_len + 2)
        )


def collect_source_files(input_folder: Path) -> list[Path]:
    files: list[Path] = []
    for file_path in sorted(input_folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        files.append(file_path)
    return files


def extract_all(source_files: list[Path]) -> tuple[int, list[dict[str, Any]], list[dict[str, Any]]]:
    try:
        import xlwings as xw  # Imported lazily to keep no-file dry runs simple.
    except ImportError as exc:
        raise RuntimeError(
            "xlwings is required for workbook extraction. Install xlwings and run on a machine with Excel."
        ) from exc

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    try:
        for file_path in source_files:
            try:
                labels = parse_file_labels(file_path.name)
            except ValueError as exc:
                print(f"Skipped {file_path.name}: filename parse failed ({exc})")
                continue

            print(f"Processed file: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_sheet = get_sheet_by_name(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"  Empirical Model not found in {file_path.name}")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_sheet, labels, file_path.name)
                    )

                regression_sheet = get_sheet_by_name(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"  Regression Model not found in {file_path.name}")
                else:
                    regression_rows.extend(
                        extract_regression_rows(regression_sheet, labels, file_path.name)
                    )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: extraction error ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    return files_processed, empirical_rows, regression_rows


def write_output(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def main() -> None:
    input_folder = Path(input_dir).expanduser()
    output_folder = Path(output_dir).expanduser()

    if not input_folder.exists() or not input_folder.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_folder}")

    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = choose_output_path(output_folder, input_folder.name)

    source_files = collect_source_files(input_folder)
    if source_files:
        files_processed, empirical_rows, regression_rows = extract_all(source_files)
    else:
        files_processed, empirical_rows, regression_rows = 0, [], []

    write_output(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
