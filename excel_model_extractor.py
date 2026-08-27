from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import xlwings as xw
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable paths
# ----------------------------
input_dir = Path("./input")
output_dir = Path("./output")


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
MODEL_PHASE_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_TO_NUM = {
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


def to_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
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


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def build_output_path(src_input_dir: Path, src_output_dir: Path) -> Path:
    folder_name = src_input_dir.resolve().name
    base_name = f"{folder_name}_PARAM"
    candidate = src_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = src_output_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_file_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem

    ticker = ""
    period_token = ""

    # Primary pattern:
    # MedMiner_Model - AORT - MidJan2026_Send.xlsx
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) >= 3:
        ticker = parts[-2]
        period_token = parts[-1]
    else:
        fallback = re.search(
            r"(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})",
            stem,
            flags=re.IGNORECASE,
        )
        if fallback:
            ticker = fallback.group("ticker").strip()
            period_token = fallback.group("period").strip()

    period_token = period_token.replace("_Send", "").strip()
    period_match = re.search(
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    model_period = period_token or "unknown_period"
    model_date = ""

    if period_match:
        phase = period_match.group("phase").capitalize()
        month_abbrev = period_match.group("month").capitalize()
        year = int(period_match.group("year"))
        model_period = f"{phase}{month_abbrev}_{year}"

        month_num = MONTH_TO_NUM.get(month_abbrev.lower())
        day = MODEL_PHASE_DAY.get(phase.lower())
        if month_num and day:
            model_date = date(year, month_num, day).isoformat()

    ticker = ticker or "UNKNOWN"
    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def iter_source_files(src_input_dir: Path) -> Iterable[Path]:
    for path in sorted(src_input_dir.iterdir()):
        if not path.is_file():
            print(f"Skipped {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not .xlsx")
            continue
        if re.search(r"_PARAM(\.\d+)?\.xlsx$", path.name, flags=re.IGNORECASE):
            print(f"Skipped {path.name}: output workbook pattern")
            continue
        yield path


def close_workbook_without_save(wb: xw.Book) -> None:
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
        # Last-resort fallback: attempt plain close.
        try:
            wb.close()
        except Exception:
            pass


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Compatibility fallback if formula2 is unavailable.
        cell.formula = formula


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    target = anchor_text.strip().lower()

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return start_row + r_idx, start_col + c_idx
    return None


def calc_bounds(
    forecast_value: Optional[float],
    raw_max_value: Optional[float],
    raw_min_value: Optional[float],
) -> Tuple[Optional[float], Optional[float]]:
    if forecast_value is None:
        return raw_max_value, raw_min_value

    # If the workbook stores multipliers (e.g. 1.1 / 0.9), apply them.
    # Otherwise, treat nearby values as already absolute bounds.
    if raw_max_value is not None and 0 < raw_max_value < 10:
        forecast_max = forecast_value * raw_max_value
    elif raw_max_value is not None:
        forecast_max = raw_max_value
    else:
        forecast_max = forecast_value * 1.05

    if raw_min_value is not None and 0 < raw_min_value < 10:
        forecast_min = forecast_value * raw_min_value
    elif raw_min_value is not None:
        forecast_min = raw_min_value
    else:
        forecast_min = forecast_value * 0.95

    if forecast_min is not None and forecast_max is not None and forecast_min > forecast_max:
        forecast_min, forecast_max = forecast_max, forecast_min

    return forecast_max, forecast_min


def rounded_signature(*values: Any) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is None:
            signature.append(None)
        else:
            signature.append(round(numeric, 8))
    return tuple(signature)


def extract_empirical_rows(
    wb: xw.Book,
    source_file: str,
    file_meta: Dict[str, str],
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical for {source_file}: sheet 'Empirical Model' missing")
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"Skipped empirical for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    last_data_row = anchor_row - 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = x_col - 1
    scratch_col = anchor_col + 12

    raw_max = to_float(sheet.cells(anchor_row, anchor_col + 1).value)
    raw_min = to_float(sheet.cells(anchor_row + 1, anchor_col + 1).value)

    rows: List[Dict[str, Any]] = []

    for num_quarters in range(1, N_QUARTERS + 1):
        first_row = max(1, last_data_row - num_quarters + 1)
        if first_row > last_data_row:
            continue

        avg_pen_cell = sheet.cells(anchor_row, scratch_col)
        avg_pen_formula = (
            f"=AVERAGE(R{first_row}C{y_col}:R{last_data_row}C{y_col}"
            f"/R{first_row}C{x_col}:R{last_data_row}C{x_col})"
        )
        set_formula2(avg_pen_cell, avg_pen_formula)
        wb.app.calculate()

        avg_penetration_pct = to_float(avg_pen_cell.value)
        quarterly_sales = to_float(sheet.cells(last_data_row, y_col).value)
        reported_sales = to_float(sheet.cells(last_data_row, x_col).value)
        previous_reported_sales = (
            to_float(sheet.cells(last_data_row - 1, x_col).value)
            if last_data_row - 1 >= first_row
            else None
        )
        growth_rate_pct = safe_div(
            None if reported_sales is None or previous_reported_sales is None else reported_sales - previous_reported_sales,
            previous_reported_sales,
        )
        sales_captured_in_db_pct = safe_div(quarterly_sales, reported_sales)
        forecast_value = safe_div(quarterly_sales, avg_penetration_pct)
        actual_value = reported_sales
        forecast_max, forecast_min = calc_bounds(forecast_value, raw_max, raw_min)

        rows.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": sheet.cells(last_data_row, quarter_col).value,
                "forecast_value": forecast_value,  # estimated total sold
                "actual_value": actual_value,  # reported sales
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width(forecast_max, forecast_min),
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
    wb: xw.Book,
    source_file: str,
    file_meta: Dict[str, str],
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression for {source_file}: sheet 'Regression Model' missing")
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"Skipped regression for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    last_data_row = anchor_row - 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    scratch_col = anchor_col + 12

    raw_max = to_float(sheet.cells(anchor_row, anchor_col + 1).value)
    raw_min = to_float(sheet.cells(anchor_row + 1, anchor_col + 1).value)

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for num_quarters in range(1, N_QUARTERS + 1):
        first_row = max(1, last_data_row - num_quarters + 1)
        if first_row > last_data_row:
            continue

        intercept_cell = sheet.cells(anchor_row, scratch_col)
        slope_cell = sheet.cells(anchor_row + 1, scratch_col)
        forecast_cell = sheet.cells(anchor_row + 2, scratch_col)

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_data_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_data_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_data_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_data_row}C{x_col})"
        )
        forecast_formula = (
            f"=R{anchor_row}C{scratch_col}"
            f"+R{anchor_row + 1}C{scratch_col}*R{last_data_row}C{x_col}"
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        set_formula2(forecast_cell, forecast_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = to_float(forecast_cell.value)  # TOT FCST w/o SA

        forecast_max, forecast_min = calc_bounds(forecast_value, raw_max, raw_min)
        signature = rounded_signature(
            forecast_value,
            intercept,
            slope,
            forecast_max,
            forecast_min,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,  # TOT FCST w/o SA
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def format_output_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    ws = writer.book[sheet_name]
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(df.columns, start=1):
        values = [str(v) for v in df[col_name].dropna().head(300)]
        max_len = max([len(col_name)] + [len(v) for v in values]) if values else len(col_name)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 50)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    empirical_df = pd.DataFrame(empirical_rows, columns=EMPIRICAL_COLUMNS)
    regression_df = pd.DataFrame(regression_rows, columns=REGRESSION_COLUMNS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        empirical_df.to_excel(writer, sheet_name="empirical_candidates", index=False)
        regression_df.to_excel(writer, sheet_name="regression_candidates", index=False)
        format_output_sheet(writer, "empirical_candidates", empirical_df)
        format_output_sheet(writer, "regression_candidates", regression_df)


def main() -> None:
    src_input_dir = Path(input_dir).expanduser().resolve()
    src_output_dir = Path(output_dir).expanduser().resolve()

    if not src_input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {src_input_dir}")

    src_output_dir.mkdir(parents=True, exist_ok=True)

    source_files = list(iter_source_files(src_input_dir))
    output_path = build_output_path(src_input_dir, src_output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_label(file_path)

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        source_file=file_path.name,
                        file_meta=metadata,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        source_file=file_path.name,
                        file_meta=metadata,
                    )
                )
                processed_count += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
