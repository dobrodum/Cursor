from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ------------------------------
# Configure paths here
# ------------------------------
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


EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -7,
    "last_quarter_used": -6,
    "forecast_value": -2,
    "actual_value": -1,
    "avg_penetration_pct": -5,
    "quarterly_sales": -4,
    "reported_sales": -1,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -8,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -3,
    "forecast_value": -2,  # TOT FCST w/o SA
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": 2,
    "slope": 3,
}


EMPIRICAL_HEADER_ALIASES = {
    "num_quarters_used": ["num quarters", "quarters used", "n quarters", "n_qtrs"],
    "last_quarter_used": ["last quarter", "last qtr"],
    "forecast_value": [
        "estimated total sold",
        "est total sold",
        "forecast value",
        "tot fcst",
    ],
    "actual_value": ["reported sales", "actual sales", "actual value"],
    "avg_penetration_pct": ["avg penetration", "average penetration"],
    "quarterly_sales": ["quarterly sales", "qtr sales"],
    "reported_sales": ["reported sales", "actual sales"],
    "growth_rate_pct": ["growth rate", "growth %"],
    "sales_captured_in_db_pct": ["sales captured", "captured in db", "db capture"],
    "forecast_min": ["min"],
}

REGRESSION_HEADER_ALIASES = {
    "num_quarters_used": ["num quarters", "quarters used", "n quarters", "n_qtrs"],
    "forecast_value": ["tot fcst w/o sa", "tot fcst", "forecast without sa"],
    "actual_value": ["actual", "reported sales", "actual sales"],
    "forecast_min": ["min"],
    "intercept": ["intercept"],
    "slope": ["slope"],
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def to_2d(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def as_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except (ValueError, TypeError):
        return None


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback for environments where formula2 is unavailable.
        cell.formula = formula_r1c1


def safe_close_workbook(wb: xw.Book) -> None:
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
    except Exception as close_error:
        print(f"SKIPPED: {wb.name} | close failed: {close_error}")


def find_anchor_max(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = to_2d(used.options(ndim=2).value)
    if not values:
        return None

    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == "max":
                return used.row + row_idx, used.column + col_idx
    return None


def build_header_map(sheet: xw.Sheet, header_row: int) -> Dict[int, str]:
    used = sheet.used_range
    first_col = used.column
    last_col = used.last_cell.column
    header_values = sheet.range((header_row, first_col), (header_row, last_col)).value

    if not isinstance(header_values, list):
        header_values = [header_values]

    header_map: Dict[int, str] = {}
    for idx, value in enumerate(header_values):
        header_map[first_col + idx] = normalize_text(value)
    return header_map


def resolve_column(
    header_map: Dict[int, str],
    aliases: Sequence[str],
    anchor_col: int,
    default_offset: Optional[int] = None,
) -> Optional[int]:
    matches: List[int] = []
    for col, header_text in header_map.items():
        if any(alias in header_text for alias in aliases):
            matches.append(col)

    if matches:
        return min(matches, key=lambda col: abs(col - anchor_col))
    if default_offset is None:
        return None
    resolved = anchor_col + default_offset
    return resolved if resolved > 0 else None


def parse_filename_metadata(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    # Example: MedMiner_Model - AORT - MidJan2026_Send.xlsx
    pattern = re.compile(
        r"-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        raise ValueError("filename does not match expected pattern")

    ticker = match.group(1).upper()
    period_type = match.group(2).lower()
    month_token = match.group(3)[:3].lower()
    year = int(match.group(4))

    if month_token not in MONTH_MAP:
        raise ValueError(f"unrecognized month token: {match.group(3)}")

    month_num = MONTH_MAP[month_token]
    day = PERIOD_DAY[period_type]
    model_date = date(year, month_num, day).isoformat()
    month_label = match.group(3)[:3].title()
    period_label = f"{match.group(2).title()}{month_label}_{year}"
    model = f"{ticker}_{period_label}"

    return ModelMeta(
        model=model,
        ticker=ticker,
        model_period=period_label,
        model_date=model_date,
        source_file=file_path.name,
    )


def compute_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = as_number(max_value)
    min_num = as_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def should_skip_empty_row(values: Sequence[Any]) -> bool:
    for value in values:
        if value not in (None, ""):
            return False
    return True


def extract_empirical_candidates(wb: xw.Book, meta: ModelMeta) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"SKIPPED: {meta.source_file} | missing sheet: Empirical Model")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"SKIPPED: {meta.source_file} | no 'max' anchor in Empirical Model")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row)

    resolved_cols: Dict[str, Optional[int]] = {
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(
            header_map,
            EMPIRICAL_HEADER_ALIASES["forecast_min"],
            anchor_col,
            EMPIRICAL_DEFAULT_OFFSETS["forecast_min"],
        ),
    }

    for key, aliases in EMPIRICAL_HEADER_ALIASES.items():
        if key == "forecast_min":
            continue
        resolved_cols[key] = resolve_column(
            header_map,
            aliases,
            anchor_col,
            EMPIRICAL_DEFAULT_OFFSETS.get(key),
        )

    start_row = anchor_row + 1
    avg_col = resolved_cols.get("avg_penetration_pct")
    formula_updates = False

    if avg_col and avg_col > 1:
        for i in range(N_QUARTERS):
            row = start_row + i
            avg_span = min(i + 1, avg_col - 1)
            if avg_span < 1:
                continue

            left_start_col = avg_col - avg_span
            support_values = sheet.range((row, left_start_col), (row, avg_col - 1)).value
            support_values_1d = (
                support_values
                if isinstance(support_values, list)
                else [support_values]
            )
            if all(value in (None, "") for value in support_values_1d):
                continue

            # R1C1 formula via formula2 per requirement.
            set_formula2(sheet.cells(row, avg_col), f"=AVERAGE(RC[-{avg_span}]:RC[-1])")
            formula_updates = True

    if formula_updates:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters_used = read_cell(sheet, row, resolved_cols.get("num_quarters_used") or -1)
        if num_quarters_used in (None, ""):
            num_quarters_used = i + 1

        last_quarter_used = read_cell(sheet, row, resolved_cols.get("last_quarter_used") or -1)
        forecast_value = read_cell(sheet, row, resolved_cols.get("forecast_value") or -1)
        actual_value = read_cell(sheet, row, resolved_cols.get("actual_value") or -1)
        forecast_max = read_cell(sheet, row, resolved_cols["forecast_max"] or -1)
        forecast_min = read_cell(sheet, row, resolved_cols.get("forecast_min") or -1)
        avg_penetration_pct = read_cell(sheet, row, resolved_cols.get("avg_penetration_pct") or -1)
        quarterly_sales = read_cell(sheet, row, resolved_cols.get("quarterly_sales") or -1)
        reported_sales = read_cell(sheet, row, resolved_cols.get("reported_sales") or -1)
        growth_rate_pct = read_cell(sheet, row, resolved_cols.get("growth_rate_pct") or -1)
        sales_captured_in_db_pct = read_cell(
            sheet,
            row,
            resolved_cols.get("sales_captured_in_db_pct") or -1,
        )

        if should_skip_empty_row(
            [forecast_value, forecast_max, forecast_min, avg_penetration_pct, quarterly_sales]
        ):
            continue

        row_data: Dict[str, Any] = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,  # estimated total sold
            "actual_value": actual_value,  # reported sales
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales if reported_sales not in (None, "") else actual_value,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": meta.source_file,
        }
        rows.append(row_data)

    return rows


def rows_equal_for_dedup(current_row: Dict[str, Any], previous_row: Dict[str, Any]) -> bool:
    compare_keys = ["num_quarters_used", "intercept", "slope", "forecast_value", "forecast_max", "forecast_min"]

    for key in compare_keys:
        current_value = current_row.get(key)
        previous_value = previous_row.get(key)
        current_num = as_number(current_value)
        previous_num = as_number(previous_value)

        if current_num is not None and previous_num is not None:
            if abs(current_num - previous_num) > 1e-9:
                return False
            continue

        if current_value != previous_value:
            return False

    return True


def extract_regression_candidates(wb: xw.Book, meta: ModelMeta) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"SKIPPED: {meta.source_file} | missing sheet: Regression Model")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"SKIPPED: {meta.source_file} | no 'max' anchor in Regression Model")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    header_map = build_header_map(sheet, anchor_row)

    resolved_cols: Dict[str, Optional[int]] = {
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(
            header_map,
            REGRESSION_HEADER_ALIASES["forecast_min"],
            anchor_col,
            REGRESSION_DEFAULT_OFFSETS["forecast_min"],
        ),
    }

    for key, aliases in REGRESSION_HEADER_ALIASES.items():
        if key == "forecast_min":
            continue
        resolved_cols[key] = resolve_column(
            header_map,
            aliases,
            anchor_col,
            REGRESSION_DEFAULT_OFFSETS.get(key),
        )

    start_row = anchor_row + 1
    data_end_row = anchor_row - 1
    formula_updates = False

    intercept_col = resolved_cols.get("intercept")
    slope_col = resolved_cols.get("slope")

    if intercept_col and slope_col and y_col > 0 and x_col > 0 and data_end_row > 0:
        for i in range(N_QUARTERS):
            n_used = i + 1
            data_start_row = max(1, data_end_row - n_used + 1)
            if data_start_row > data_end_row:
                continue

            target_row = start_row + i
            intercept_formula = (
                f"=INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
                f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
                f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
            )

            # R1C1 formulas via formula2 per requirement.
            set_formula2(sheet.cells(target_row, intercept_col), intercept_formula)
            set_formula2(sheet.cells(target_row, slope_col), slope_formula)
            formula_updates = True

    if formula_updates:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters_used = read_cell(sheet, row, resolved_cols.get("num_quarters_used") or -1)
        if num_quarters_used in (None, ""):
            num_quarters_used = i + 1

        intercept = read_cell(sheet, row, intercept_col or -1)
        slope = read_cell(sheet, row, slope_col or -1)
        forecast_value = read_cell(sheet, row, resolved_cols.get("forecast_value") or -1)
        forecast_max = read_cell(sheet, row, resolved_cols["forecast_max"] or -1)
        forecast_min = read_cell(sheet, row, resolved_cols.get("forecast_min") or -1)
        actual_value = read_cell(sheet, row, resolved_cols.get("actual_value") or -1)

        if forecast_value in (None, ""):
            intercept_num = as_number(intercept)
            slope_num = as_number(slope)
            x_latest = as_number(read_cell(sheet, data_end_row, x_col))
            if intercept_num is not None and slope_num is not None and x_latest is not None:
                forecast_value = intercept_num + slope_num * x_latest

        if should_skip_empty_row([intercept, slope, forecast_value, forecast_max, forecast_min]):
            continue

        row_data: Dict[str, Any] = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,  # TOT FCST w/o SA
            "actual_value": actual_value if actual_value not in (None, "") else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": meta.source_file,
        }

        if rows and rows_equal_for_dedup(row_data, rows[-1]):
            continue

        rows.append(row_data)

    return rows


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    primary = out_dir / f"{base_name}.xlsx"
    if not primary.exists():
        return primary

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def safe_display(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    headers: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(col, "") for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=idx).value
            max_len = max(max_len, len(safe_display(cell_value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 40)


def write_output_workbook(
    out_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(out_path)


def main() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir not found or not a directory: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = unique_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"SKIPPED: {file_path.name} | temp file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED: {file_path.name} | not .xlsx")
                continue

            try:
                meta = parse_filename_metadata(file_path)
            except ValueError as parse_error:
                print(f"SKIPPED: {file_path.name} | {parse_error}")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_empirical_rows = extract_empirical_candidates(wb, meta)
                file_regression_rows = extract_regression_candidates(wb, meta)
                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                files_processed += 1
                print(f"PROCESSED: {file_path.name}")
            except Exception as workbook_error:
                print(f"SKIPPED: {file_path.name} | workbook processing error: {workbook_error}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(out_path, empirical_rows, regression_rows)

    print(f"OUTPUT: {out_path}")
    print(f"FILES_PROCESSED: {files_processed}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
