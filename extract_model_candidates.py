from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# User-configurable inputs
# =========================
input_dir = "./input"
output_dir = "./output"


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


DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if value != value:  # NaN
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.startswith("#"):  # Excel error literals, e.g. #N/A
            return None
        if text.endswith("%"):
            try:
                return float(text[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_number(value)
    if numeric is None:
        return None
    return int(round(numeric))


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
    except TypeError:
        wb.api.Close(False)
    except Exception:
        # Last safety fallback: close without arguments.
        try:
            wb.close()
        except Exception:
            pass


def get_sheet_case_insensitive(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == normalize_text(anchor_text):
                return used.row + row_idx, used.column + col_idx
    return None


def build_nearby_header_map(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    row_offsets: Iterable[int] = (-1, 0, 1),
    col_window: int = 24,
) -> Dict[str, int]:
    header_map: Dict[str, int] = {}
    left_col = max(1, anchor_col - col_window)
    right_col = anchor_col + col_window

    for row_offset in row_offsets:
        row_num = anchor_row + row_offset
        if row_num < 1:
            continue
        row_values = sheet.range((row_num, left_col), (row_num, right_col)).value
        if row_values is None:
            continue
        if not isinstance(row_values, list):
            row_values = [row_values]
        for i, raw_value in enumerate(row_values):
            normalized = normalize_text(raw_value)
            if normalized and normalized not in header_map:
                header_map[normalized] = left_col + i

    return header_map


def find_column(header_map: Dict[str, int], patterns: Iterable[str]) -> Optional[int]:
    normalized_patterns = [normalize_text(pattern) for pattern in patterns]
    for pattern in normalized_patterns:
        for header, col in header_map.items():
            if pattern and pattern in header:
                return col
    return None


def read_cell(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1 or row < 1:
        return None
    return sheet.range((row, col)).value


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback for environments that require explicit Formula2R1C1 API assignment.
        cell.api.Formula2R1C1 = formula_r1c1


def parse_model_metadata(filename: str) -> Dict[str, str]:
    stem = Path(filename).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    period_token = ""
    if len(parts) >= 3:
        ticker = parts[1].strip().upper()
        period_token = parts[2].split("_")[0].strip()
    else:
        ticker_match = re.search(r"\b([A-Z]{2,6})\b", stem)
        ticker = ticker_match.group(1) if ticker_match else ""
        period_match = re.search(r"(Early|Mid|Late)[A-Za-z]{3,9}\d{4}", stem, flags=re.IGNORECASE)
        period_token = period_match.group(0) if period_match else ""

    model_period = period_token
    model_date = ""

    token_match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", period_token, flags=re.IGNORECASE)
    if token_match:
        period_word_raw, month_raw, year_raw = token_match.groups()
        period_word = period_word_raw.title()
        month_num = None
        for fmt in ("%b", "%B"):
            try:
                month_num = datetime.strptime(month_raw[:3] if fmt == "%b" else month_raw, fmt).month
                break
            except ValueError:
                continue
        if month_num is None:
            try:
                month_num = datetime.strptime(month_raw[:3], "%b").month
            except ValueError:
                month_num = None

        if month_num is not None:
            month_abbrev = datetime(2000, month_num, 1).strftime("%b")
            year_int = int(year_raw)
            day = DAY_BY_PERIOD[period_word.lower()]
            model_period = f"{period_word}{month_abbrev}_{year_int}"
            model_date = date(year_int, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or model_period

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def resolve_output_path(input_path: Path, output_path: Path) -> Path:
    input_folder_name = input_path.resolve().name
    base_stem = f"{input_folder_name}_PARAM"
    candidate = output_path / f"{base_stem}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{base_stem}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def is_effectively_empty(values: Iterable[Any]) -> bool:
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in values)


def process_empirical_sheet(
    wb: xw.Book, sheet: xw.Sheet, file_path: Path, metadata: Dict[str, str]
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped {file_path.name} empirical rows: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_nearby_header_map(sheet, anchor_row, anchor_col)
    used_last_col = sheet.used_range.last_cell.column

    min_col = find_column(header_map, ["min"]) or (anchor_col + 1)
    forecast_value_col = find_column(
        header_map,
        ["estimated total sold", "est total sold", "forecast value", "forecast total"],
    ) or (anchor_col - 1)
    actual_value_col = find_column(header_map, ["reported sales", "actual sales", "actual"]) or None
    num_quarters_col = find_column(header_map, ["num quarters used", "num quarters"]) or (anchor_col - 8)
    last_quarter_col = find_column(header_map, ["last quarter used", "last quarter"]) or (anchor_col - 9)
    avg_penetration_col = find_column(header_map, ["avg penetration", "average penetration"]) or None
    quarterly_sales_col = find_column(header_map, ["quarterly sales"]) or None
    reported_sales_col = find_column(header_map, ["reported sales"]) or actual_value_col
    growth_rate_col = find_column(header_map, ["growth rate"]) or None
    captured_sales_col = find_column(
        header_map,
        ["sales captured in db", "captured in db", "sales captured"],
    ) or None

    avg_penetration_source_col = captured_sales_col or avg_penetration_col
    avg_pen_temp_col = used_last_col + 2

    row_payloads: List[Dict[str, Any]] = []
    formula_written = False

    for n_quarters in range(1, 11):
        row_num = anchor_row + n_quarters
        num_quarters_used = to_int(read_cell(sheet, row_num, num_quarters_col)) or n_quarters

        payload: Dict[str, Any] = {
            "row_num": row_num,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": read_cell(sheet, row_num, last_quarter_col),
            "forecast_value": to_number(read_cell(sheet, row_num, forecast_value_col)),
            "actual_value": to_number(read_cell(sheet, row_num, actual_value_col)),
            "forecast_max": to_number(read_cell(sheet, row_num, anchor_col)),
            "forecast_min": to_number(read_cell(sheet, row_num, min_col)),
            "quarterly_sales": to_number(read_cell(sheet, row_num, quarterly_sales_col)),
            "reported_sales": to_number(read_cell(sheet, row_num, reported_sales_col)),
            "growth_rate_pct": to_number(read_cell(sheet, row_num, growth_rate_col)),
            "sales_captured_in_db_pct": to_number(read_cell(sheet, row_num, captured_sales_col)),
            "avg_penetration_pct": to_number(read_cell(sheet, row_num, avg_penetration_col)),
            "avg_pen_formula_cell": None,
        }

        if avg_penetration_source_col is not None and num_quarters_used > 0:
            start_row = max(anchor_row + 1, row_num - num_quarters_used + 1)
            temp_cell = sheet.range((row_num, avg_pen_temp_col))
            formula = f"=AVERAGE(R{start_row}C{avg_penetration_source_col}:R{row_num}C{avg_penetration_source_col})"
            set_formula2_r1c1(temp_cell, formula)
            payload["avg_pen_formula_cell"] = temp_cell
            formula_written = True

        row_payloads.append(payload)

    if formula_written:
        wb.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    for payload in row_payloads:
        avg_penetration_pct = payload["avg_penetration_pct"]
        if payload["avg_pen_formula_cell"] is not None:
            avg_penetration_pct = to_number(payload["avg_pen_formula_cell"].value)

        row_check = [
            payload["forecast_value"],
            payload["actual_value"],
            payload["forecast_max"],
            payload["forecast_min"],
            payload["quarterly_sales"],
            payload["reported_sales"],
            payload["growth_rate_pct"],
            payload["sales_captured_in_db_pct"],
            avg_penetration_pct,
        ]
        if is_effectively_empty(row_check):
            continue

        forecast_max = payload["forecast_max"]
        forecast_min = payload["forecast_min"]
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        extracted_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": payload["num_quarters_used"],
                "last_quarter_used": payload["last_quarter_used"],
                "forecast_value": payload["forecast_value"],
                "actual_value": payload["actual_value"],
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": payload["quarterly_sales"],
                "reported_sales": payload["reported_sales"],
                "growth_rate_pct": payload["growth_rate_pct"],
                "sales_captured_in_db_pct": payload["sales_captured_in_db_pct"],
                "source_file": file_path.name,
            }
        )

    return extracted_rows


def process_regression_sheet(
    wb: xw.Book, sheet: xw.Sheet, file_path: Path, metadata: Dict[str, str]
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped {file_path.name} regression rows: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"Skipped {file_path.name} regression rows: invalid x/y offsets from anchor")
        return []

    header_map = build_nearby_header_map(sheet, anchor_row, anchor_col)
    used_last_row = sheet.used_range.last_cell.row
    used_last_col = sheet.used_range.last_cell.column

    min_col = find_column(header_map, ["min"]) or (anchor_col + 1)
    num_quarters_col = find_column(header_map, ["num quarters used", "num quarters"]) or (anchor_col - 8)
    forecast_wo_sa_col = find_column(
        header_map,
        ["tot fcst w o sa", "tot fcst wo sa", "total fcst without sa", "forecast w o sa"],
    ) or None
    actual_value_col = find_column(header_map, ["actual value", "actual sales", "actual"]) or None

    x_values = sheet.range((1, x_col), (used_last_row, x_col)).value
    y_values = sheet.range((1, y_col), (used_last_row, y_col)).value
    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]

    paired_rows: List[int] = []
    for row_num, (x_val, y_val) in enumerate(zip(x_values, y_values), start=1):
        if to_number(x_val) is not None and to_number(y_val) is not None:
            paired_rows.append(row_num)

    if len(paired_rows) < 2:
        print(f"Skipped {file_path.name} regression rows: not enough numeric x/y data")
        return []

    temp_intercept_col = used_last_col + 2
    temp_slope_col = used_last_col + 3
    temp_forecast_col = used_last_col + 4

    calc_plan: List[Dict[str, Any]] = []
    max_windows = min(10, len(paired_rows))
    for n_quarters in range(2, max_windows + 1):
        window_rows = paired_rows[-n_quarters:]
        start_row = window_rows[0]
        end_row = window_rows[-1]
        output_row = anchor_row + (n_quarters - 1)

        intercept_cell = sheet.range((output_row, temp_intercept_col))
        slope_cell = sheet.range((output_row, temp_slope_col))
        forecast_cell = sheet.range((output_row, temp_forecast_col))

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        forecast_formula = f"=R{output_row}C{temp_intercept_col}+R{output_row}C{temp_slope_col}*R{end_row}C{x_col}"

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)

        calc_plan.append(
            {
                "output_row": output_row,
                "n_quarters": n_quarters,
                "intercept_cell": intercept_cell,
                "slope_cell": slope_cell,
                "forecast_cell": forecast_cell,
            }
        )

    if calc_plan:
        wb.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for plan in calc_plan:
        output_row = plan["output_row"]
        n_quarters = to_int(read_cell(sheet, output_row, num_quarters_col)) or int(plan["n_quarters"])

        intercept = to_number(plan["intercept_cell"].value)
        slope = to_number(plan["slope_cell"].value)

        forecast_value = (
            to_number(read_cell(sheet, output_row, forecast_wo_sa_col))
            if forecast_wo_sa_col is not None
            else to_number(plan["forecast_cell"].value)
        )
        actual_value = to_number(read_cell(sheet, output_row, actual_value_col))
        forecast_max = to_number(read_cell(sheet, output_row, anchor_col))
        forecast_min = to_number(read_cell(sheet, output_row, min_col))

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            n_quarters,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row_check = [forecast_value, forecast_max, forecast_min, intercept, slope]
        if is_effectively_empty(row_check):
            continue

        extracted_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": file_path.name,
            }
        )

    return extracted_rows


def write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row_data in rows:
        ws.append([row_data.get(col) for col in columns])

    bold_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(1, ws.max_row)}"

    for col_idx in range(1, ws.max_column + 1):
        max_len = len(str(ws.cell(row=1, column=col_idx).value or ""))
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 50)


def write_output_workbook(
    output_file: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory not found or not a folder: {input_path}")

    output_file = resolve_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing file: {file_path.name}")
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open ({exc})")
                continue

            try:
                metadata = parse_model_metadata(file_path.name)

                empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical for {file_path.name}: sheet not found")
                else:
                    empirical_rows.extend(process_empirical_sheet(wb, empirical_sheet, file_path, metadata))

                regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression for {file_path.name}: sheet not found")
                else:
                    regression_rows.extend(process_regression_sheet(wb, regression_sheet, file_path, metadata))

                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
