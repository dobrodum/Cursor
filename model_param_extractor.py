#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "forecast_value": -1,
    "actual_value": 2,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -6,
    "quarterly_sales": -4,
    "reported_sales": -3,
    "growth_rate_pct": -2,
    "sales_captured_in_db_pct": -5,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -1,
    "actual_value": 2,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": -5,
    "slope": -4,
}

EMPIRICAL_HEADER_PATTERNS = {
    "num_quarters_used": [
        "num quarters used",
        "num quarters",
        "quarters used",
        "n quarters",
        "# quarters",
    ],
    "last_quarter_used": ["last quarter used", "last quarter", "ending quarter"],
    "forecast_value": [
        "estimated total sold",
        "total sold estimate",
        "tot fcst",
        "total forecast",
        "forecast value",
    ],
    "actual_value": ["reported sales", "actual sales", "actual value", "actual"],
    "forecast_max": ["max", "forecast max"],
    "forecast_min": ["min", "forecast min"],
    "avg_penetration_pct": [
        "avg penetration",
        "average penetration",
        "avg penetration pct",
    ],
    "quarterly_sales": ["quarterly sales", "qtr sales"],
    "reported_sales": ["reported sales", "sales reported"],
    "growth_rate_pct": ["growth rate pct", "growth rate", "growth pct"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db",
    ],
}

REGRESSION_HEADER_PATTERNS = {
    "num_quarters_used": [
        "num quarters used",
        "num quarters",
        "quarters used",
        "n quarters",
        "# quarters",
    ],
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst without sa",
        "total forecast without sa",
        "forecast value",
        "total forecast",
    ],
    "actual_value": ["actual value", "actual sales", "reported sales"],
    "forecast_max": ["max", "forecast max"],
    "forecast_min": ["min", "forecast min"],
    "intercept": ["intercept"],
    "slope": ["slope"],
}

DAY_BY_PERIOD_PREFIX = {"early": 5, "mid": 15, "late": 25}
MONTH_LOOKUP = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def clean_value(value: Any) -> Any:
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return value


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if not stripped:
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def compute_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_float = to_float(max_value)
    min_float = to_float(min_value)
    if max_float is None or min_float is None:
        return None
    return max_float - min_float


def quantize(value: Any) -> Any:
    value_float = to_float(value)
    if value_float is None:
        return clean_value(value)
    return round(value_float, 8)


def as_2d(values: Any, rows: Optional[int] = None, cols: Optional[int] = None) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        if rows is not None and cols is not None:
            if rows > 1 and cols == 1:
                return [[value] for value in values]
            if rows == 1 and cols > 1:
                return [values]
        return [values]
    return [[values]]


def set_formula2(cell: Any, formula: str) -> bool:
    try:
        cell.formula2 = formula
        return True
    except Exception:
        return False


def safe_close_workbook(wb: Any) -> None:
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


def parse_model_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) > 1 else ""
    period_section = parts[2] if len(parts) > 2 else ""
    period_token = period_section.split("_")[0].strip()

    model_period = period_token
    model_date = ""

    period_match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token, flags=re.IGNORECASE)
    if period_match:
        prefix = period_match.group(1).title()
        month_token = period_match.group(2).lower()
        year = int(period_match.group(3))
        month_number = MONTH_LOOKUP.get(month_token)
        day = DAY_BY_PERIOD_PREFIX.get(prefix.lower())
        if month_number and day:
            month_abbr = date(year, month_number, 1).strftime("%b")
            model_period = f"{prefix}{month_abbr}_{year}"
            model_date = date(year, month_number, day).isoformat()

    if not ticker:
        ticker = parts[0] if parts else stem
    model = f"{ticker}_{model_period}" if model_period else ticker

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def should_skip_file(path: Path) -> Tuple[bool, str]:
    if not path.is_file():
        return True, "not a file"
    if path.name.startswith("~"):
        return True, "temporary file"
    if path.suffix.lower() != ".xlsx":
        return True, "not an .xlsx file"
    if re.search(r"_param(\.\d+)?\.xlsx$", path.name, flags=re.IGNORECASE):
        return True, "output-like file"
    return False, ""


def find_max_anchor(sheet: Any) -> Optional[Tuple[int, int]]:
    used_range = sheet.used_range
    values = as_2d(
        used_range.value,
        rows=used_range.rows.count,
        cols=used_range.columns.count,
    )
    start_row = used_range.row
    start_col = used_range.column

    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return start_row + row_offset, start_col + col_offset
    return None


def build_header_map(sheet: Any, header_row: int, anchor_col: int, span: int = 35) -> Dict[str, int]:
    start_col = max(1, anchor_col - span)
    end_col = anchor_col + span
    row_values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]

    header_map: Dict[str, int] = {}
    for index, raw_value in enumerate(row_values):
        normalized = normalize_text(raw_value)
        if normalized and normalized not in header_map:
            header_map[normalized] = start_col + index
    return header_map


def resolve_column(
    key: str,
    header_map: Dict[str, int],
    patterns: Dict[str, Sequence[str]],
    anchor_col: int,
    fallback_offsets: Dict[str, int],
) -> Optional[int]:
    for pattern in patterns.get(key, []):
        normalized_pattern = normalize_text(pattern)
        if not normalized_pattern:
            continue
        for normalized_header, column in header_map.items():
            if normalized_pattern == normalized_header or normalized_pattern in normalized_header:
                return column

    if key in fallback_offsets:
        fallback_col = anchor_col + fallback_offsets[key]
        if fallback_col >= 1:
            return fallback_col
    return None


def read_block(sheet: Any, row_start: int, row_end: int, col_start: int, col_end: int) -> List[List[Any]]:
    values = sheet.range((row_start, col_start), (row_end, col_end)).value
    return as_2d(
        values,
        rows=(row_end - row_start + 1),
        cols=(col_end - col_start + 1),
    )


def block_value(
    block: List[List[Any]],
    row: int,
    col: Optional[int],
    row_start: int,
    col_start: int,
) -> Any:
    if col is None:
        return None
    row_index = row - row_start
    col_index = col - col_start
    if row_index < 0 or col_index < 0:
        return None
    if row_index >= len(block):
        return None
    if col_index >= len(block[row_index]):
        return None
    return clean_value(block[row_index][col_index])


def has_row_content(values: Sequence[Any]) -> bool:
    for value in values:
        cleaned = clean_value(value)
        if cleaned is None:
            continue
        if isinstance(cleaned, str) and not cleaned:
            continue
        return True
    return False


def extract_empirical_rows(wb: Any, source_file: str, metadata: Dict[str, str]) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  - Empirical Model sheet missing in {source_file}")
        return []

    anchor = find_max_anchor(sheet)
    if not anchor:
        print(f"  - max anchor not found in Empirical Model for {source_file}")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)

    column_map = {
        key: resolve_column(
            key=key,
            header_map=header_map,
            patterns=EMPIRICAL_HEADER_PATTERNS,
            anchor_col=anchor_col,
            fallback_offsets=EMPIRICAL_FALLBACK_OFFSETS,
        )
        for key in EMPIRICAL_FALLBACK_OFFSETS
    }

    row_start = anchor_row + 1
    row_end = row_start + N_QUARTERS - 1

    formulas_updated = False
    avg_pen_col = column_map.get("avg_penetration_pct")
    if avg_pen_col:
        for index in range(N_QUARTERS):
            row = row_start + index
            avg_cell = sheet.cells(row, avg_pen_col)
            # R1C1 formula keeps work local to each row and avoids column-letter conversion.
            avg_formula = f'=IFERROR(AVERAGE(RC[-{index + 1}]:RC[-1]),"")'
            if set_formula2(avg_cell, avg_formula):
                formulas_updated = True
    if formulas_updated:
        wb.app.calculate()

    useful_cols = [col for col in column_map.values() if col is not None]
    if not useful_cols:
        return []
    col_start = min(useful_cols)
    col_end = max(useful_cols)
    block = read_block(sheet, row_start, row_end, col_start, col_end)

    rows: List[Dict[str, Any]] = []
    for index in range(N_QUARTERS):
        row_number = row_start + index
        row_values = {
            key: block_value(block, row_number, col, row_start, col_start)
            for key, col in column_map.items()
        }

        content_probe = [
            row_values.get("forecast_value"),
            row_values.get("forecast_max"),
            row_values.get("forecast_min"),
            row_values.get("avg_penetration_pct"),
            row_values.get("reported_sales"),
        ]
        if not has_row_content(content_probe):
            continue

        num_quarters = row_values.get("num_quarters_used")
        if num_quarters is None:
            num_quarters = index + 1

        actual_value = row_values.get("actual_value")
        if actual_value is None:
            actual_value = row_values.get("reported_sales")

        forecast_max = row_values.get("forecast_max")
        forecast_min = row_values.get("forecast_min")
        avg_penetration = row_values.get("avg_penetration_pct")

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": row_values.get("last_quarter_used"),
                "forecast_value": row_values.get("forecast_value"),
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": row_values.get("quarterly_sales"),
                "reported_sales": row_values.get("reported_sales"),
                "growth_rate_pct": row_values.get("growth_rate_pct"),
                "sales_captured_in_db_pct": row_values.get("sales_captured_in_db_pct"),
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(wb: Any, source_file: str, metadata: Dict[str, str]) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  - Regression Model sheet missing in {source_file}")
        return []

    anchor = find_max_anchor(sheet)
    if not anchor:
        print(f"  - max anchor not found in Regression Model for {source_file}")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_map = build_header_map(sheet, anchor_row, anchor_col)
    column_map = {
        key: resolve_column(
            key=key,
            header_map=header_map,
            patterns=REGRESSION_HEADER_PATTERNS,
            anchor_col=anchor_col,
            fallback_offsets=REGRESSION_FALLBACK_OFFSETS,
        )
        for key in REGRESSION_FALLBACK_OFFSETS
    }

    row_start = anchor_row + 1
    row_end = row_start + N_QUARTERS - 1

    formulas_updated = False
    intercept_col = column_map.get("intercept")
    slope_col = column_map.get("slope")

    if intercept_col and x_col >= 1 and y_col >= 1:
        for index in range(N_QUARTERS):
            row = row_start + index
            n_rows = index + 1
            intercept_cell = sheet.cells(row, intercept_col)
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R[-{n_rows - 1}]C[{y_col - intercept_col}]:RC[{y_col - intercept_col}],'
                f'R[-{n_rows - 1}]C[{x_col - intercept_col}]:RC[{x_col - intercept_col}]),"")'
            )
            if set_formula2(intercept_cell, intercept_formula):
                formulas_updated = True

    if slope_col and x_col >= 1 and y_col >= 1:
        for index in range(N_QUARTERS):
            row = row_start + index
            n_rows = index + 1
            slope_cell = sheet.cells(row, slope_col)
            slope_formula = (
                f'=IFERROR(SLOPE(R[-{n_rows - 1}]C[{y_col - slope_col}]:RC[{y_col - slope_col}],'
                f'R[-{n_rows - 1}]C[{x_col - slope_col}]:RC[{x_col - slope_col}]),"")'
            )
            if set_formula2(slope_cell, slope_formula):
                formulas_updated = True

    if formulas_updated:
        wb.app.calculate()

    useful_cols = [col for col in column_map.values() if col is not None]
    if not useful_cols:
        return []
    col_start = min(useful_cols)
    col_end = max(useful_cols)
    block = read_block(sheet, row_start, row_end, col_start, col_end)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for index in range(N_QUARTERS):
        row_number = row_start + index
        row_values = {
            key: block_value(block, row_number, col, row_start, col_start)
            for key, col in column_map.items()
        }

        content_probe = [
            row_values.get("forecast_value"),
            row_values.get("forecast_max"),
            row_values.get("forecast_min"),
            row_values.get("intercept"),
            row_values.get("slope"),
        ]
        if not has_row_content(content_probe):
            continue

        num_quarters = row_values.get("num_quarters_used")
        if num_quarters is None:
            num_quarters = index + 1

        signature = (
            quantize(num_quarters),
            quantize(row_values.get("forecast_value")),
            quantize(row_values.get("forecast_max")),
            quantize(row_values.get("forecast_min")),
            quantize(row_values.get("intercept")),
            quantize(row_values.get("slope")),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        forecast_max = row_values.get("forecast_max")
        forecast_min = row_values.get("forecast_min")
        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": row_values.get("forecast_value"),
                "actual_value": row_values.get("actual_value"),
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "intercept": row_values.get("intercept"),
                "slope": row_values.get("slope"),
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(worksheet: Any, columns: Sequence[str], rows: List[Dict[str, Any]]) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for cell in worksheet.iter_cols(
            min_col=column_index,
            max_col=column_index,
            min_row=2,
            max_row=worksheet.max_row,
            values_only=True,
        ):
            for value in cell:
                if value is None:
                    continue
                max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(max_length + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_sheet = wb.create_sheet("empirical_candidates")
    regression_sheet = wb.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")

    files_processed = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(in_dir.iterdir(), key=lambda path: path.name.lower()):
            skip, reason = should_skip_file(file_path)
            if skip:
                print(f"SKIPPED {file_path.name}: {reason}")
                continue

            print(f"PROCESSING {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path.name)

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        source_file=file_path.name,
                        metadata=metadata,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        source_file=file_path.name,
                        metadata=metadata,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: failed to process ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    output_path = next_output_path(in_dir=in_dir, out_dir=out_dir)
    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"OUTPUT: {output_path}")
    print(f"FILES PROCESSED: {files_processed}")
    print(f"EMPIRICAL ROWS: {len(empirical_rows)}")
    print(f"REGRESSION ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
