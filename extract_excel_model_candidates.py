from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ----------------------------
# Configure folders here
# ----------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"


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
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "quarterly_sales": -6,
    "growth_rate_pct": -5,
    "avg_penetration_pct": -4,
    "forecast_value": -3,  # estimated total sold
    "actual_value": -2,  # reported sales
    "reported_sales": -2,
    "sales_captured_in_db_pct": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "penetration_series_col": -4,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "forecast_value": -3,  # TOT FCST w/o SA
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

MONTH_MAP = {
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


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        if isinstance(values[0], (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "").replace("%", "")
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def set_formula2(target_range: xw.Range, formula_r1c1: str) -> None:
    try:
        target_range.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        target_range.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        target_range.api.FormulaR1C1 = formula_r1c1
        return
    except Exception:
        # Last-resort fallback for older Excel APIs.
        target_range.formula = formula_r1c1


@dataclass
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


class SheetSnapshot:
    def __init__(self, sheet: xw.Sheet):
        self.sheet = sheet
        used = sheet.used_range
        self.first_row = used.row
        self.first_col = used.column
        self.values = to_2d(used.value)
        self.n_rows = len(self.values)
        self.n_cols = len(self.values[0]) if self.values else 0
        self.last_row = self.first_row + self.n_rows - 1 if self.n_rows else self.first_row
        self.last_col = self.first_col + self.n_cols - 1 if self.n_cols else self.first_col

    def get(self, abs_row: int, abs_col: int) -> Any:
        if self.n_rows == 0 or self.n_cols == 0:
            return None
        row_idx = abs_row - self.first_row
        col_idx = abs_col - self.first_col
        if row_idx < 0 or col_idx < 0 or row_idx >= self.n_rows or col_idx >= self.n_cols:
            return None
        return self.values[row_idx][col_idx]

    def find_max_anchor(self) -> Optional[Tuple[int, int]]:
        for r in range(self.first_row, self.last_row + 1):
            for c in range(self.first_col, self.last_col + 1):
                value = normalize_text(self.get(r, c))
                if value == "max":
                    return r, c
        return None

    def identify_headers_near_anchor(self, anchor_row: int, anchor_col: int) -> Dict[str, int]:
        header_map: Dict[str, int] = {"forecast_max": anchor_col}
        row_candidates = [anchor_row - 1, anchor_row, anchor_row + 1]
        start_col = max(self.first_col, anchor_col - 24)
        end_col = min(self.last_col, anchor_col + 12)

        for r in row_candidates:
            if r < self.first_row or r > self.last_row:
                continue
            for c in range(start_col, end_col + 1):
                key = header_key(normalize_text(self.get(r, c)))
                if key and key not in header_map:
                    header_map[key] = c
        return header_map

    def numeric_rows(self, col: int, max_row: Optional[int] = None) -> List[int]:
        end_row = min(self.last_row, max_row) if max_row is not None else self.last_row
        rows: List[int] = []
        for r in range(self.first_row, end_row + 1):
            if to_float(self.get(r, col)) is not None:
                rows.append(r)
        return rows

    def paired_numeric_rows(self, x_col: int, y_col: int, max_row: Optional[int] = None) -> List[int]:
        end_row = min(self.last_row, max_row) if max_row is not None else self.last_row
        rows: List[int] = []
        for r in range(self.first_row, end_row + 1):
            if to_float(self.get(r, x_col)) is not None and to_float(self.get(r, y_col)) is not None:
                rows.append(r)
        return rows


def header_key(text: str) -> Optional[str]:
    if not text:
        return None
    if text == "max":
        return "forecast_max"
    if text == "min":
        return "forecast_min"
    if "num" in text and "quarter" in text:
        return "num_quarters_used"
    if "last quarter" in text:
        return "last_quarter_used"
    if "estimated total sold" in text:
        return "forecast_value"
    if "tot fcst" in text and ("w/o sa" in text or "wo sa" in text):
        return "forecast_value"
    if "forecast" in text and "total" in text and "sa" in text:
        return "forecast_value"
    if "reported sales" in text:
        return "reported_sales"
    if text == "actual" or "actual sales" in text:
        return "actual_value"
    if "quarterly sales" in text:
        return "quarterly_sales"
    if "growth rate" in text:
        return "growth_rate_pct"
    if "captured" in text and "db" in text:
        return "sales_captured_in_db_pct"
    if "avg penetration" in text or "average penetration" in text:
        return "avg_penetration_pct"
    if text == "penetration" or text == "penetration %":
        return "penetration_series_col"
    if text == "intercept":
        return "intercept"
    if text == "slope":
        return "slope"
    return None


def parse_model_metadata(file_name: str) -> ModelMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    period_token = ""
    if len(parts) >= 3:
        ticker = parts[1].strip().upper() or "UNKNOWN"
        period_token = parts[2].split("_")[0].strip()
    else:
        # Loose fallback parse if file is not in expected "A - TICKER - PERIOD" format.
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1)
        period_token = stem.split("_")[0].split()[-1]

    period_match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$", period_token, flags=re.IGNORECASE)
    if period_match:
        phase = period_match.group(1).capitalize()
        month_abbr = period_match.group(2).title()
        year = int(period_match.group(3))
        day = DAY_MAP[phase.lower()]
        month_num = MONTH_MAP.get(month_abbr, 1)
        model_period = f"{phase}{month_abbr}_{year}"
        model_date = date(year, month_num, day).isoformat()
    else:
        model_period = period_token if period_token else "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def safe_close_source_workbook(wb: xw.Book) -> None:
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


def resolve_col(header_map: Dict[str, int], key: str, anchor_col: int, fallback_offsets: Dict[str, int]) -> int:
    if key in header_map:
        return header_map[key]
    return anchor_col + fallback_offsets[key]


def r1c1_range(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    return f"R{start_row}C{start_col}:R{end_row}C{end_col}"


def process_empirical_sheet(wb: xw.Book, metadata: ModelMetadata, source_file: str) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        print(f"skipped empirical in {source_file}: missing sheet 'Empirical Model'")
        return []

    sheet = wb.sheets["Empirical Model"]
    snap = SheetSnapshot(sheet)
    anchor = snap.find_max_anchor()
    if anchor is None:
        print(f"skipped empirical in {source_file}: could not find 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    header_map = snap.identify_headers_near_anchor(anchor_row, anchor_col)

    num_col = resolve_col(header_map, "num_quarters_used", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    last_q_col = resolve_col(header_map, "last_quarter_used", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    forecast_col = resolve_col(header_map, "forecast_value", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    actual_col = header_map.get(
        "actual_value",
        header_map.get("reported_sales", anchor_col + EMPIRICAL_FALLBACK_OFFSETS["actual_value"]),
    )
    max_col = resolve_col(header_map, "forecast_max", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    min_col = resolve_col(header_map, "forecast_min", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    avg_pen_col = resolve_col(header_map, "avg_penetration_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    q_sales_col = resolve_col(header_map, "quarterly_sales", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    reported_sales_col = header_map.get("reported_sales", actual_col)
    growth_col = resolve_col(header_map, "growth_rate_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    captured_col = resolve_col(header_map, "sales_captured_in_db_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    penetration_series_col = header_map.get("penetration_series_col", anchor_col + EMPIRICAL_FALLBACK_OFFSETS["penetration_series_col"])

    # Use R1C1 formula2 for avg penetration in a scratch column.
    scratch_avg_col = snap.last_col + 2
    penetration_rows = snap.numeric_rows(penetration_series_col, max_row=anchor_row - 1)
    if not penetration_rows:
        penetration_rows = snap.numeric_rows(penetration_series_col)

    formula_rows: Dict[int, Tuple[int, int]] = {}
    if penetration_rows:
        max_n = min(N_QUARTERS, len(penetration_rows))
        for i in range(1, max_n + 1):
            start_row = penetration_rows[-i]
            end_row = penetration_rows[-1]
            target_row = anchor_row + i
            avg_formula = f'=IFERROR(AVERAGE({r1c1_range(start_row, penetration_series_col, end_row, penetration_series_col)}),"")'
            set_formula2(sheet.range((target_row, scratch_avg_col)), avg_formula)
            formula_rows[i] = (target_row, scratch_avg_col)
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + i
        num_quarters_used = to_int(sheet.range((row_idx, num_col)).value) or i
        last_quarter_used = sheet.range((row_idx, last_q_col)).value
        forecast_value = to_float(sheet.range((row_idx, forecast_col)).value)
        actual_value = to_float(sheet.range((row_idx, actual_col)).value)
        forecast_max = to_float(sheet.range((row_idx, max_col)).value)
        forecast_min = to_float(sheet.range((row_idx, min_col)).value)
        quarterly_sales = to_float(sheet.range((row_idx, q_sales_col)).value)
        reported_sales = to_float(sheet.range((row_idx, reported_sales_col)).value)
        growth_rate_pct = to_float(sheet.range((row_idx, growth_col)).value)
        sales_captured_pct = to_float(sheet.range((row_idx, captured_col)).value)

        avg_penetration_pct = None
        if i in formula_rows:
            formula_row, formula_col = formula_rows[i]
            avg_penetration_pct = to_float(sheet.range((formula_row, formula_col)).value)
        if avg_penetration_pct is None:
            avg_penetration_pct = to_float(sheet.range((row_idx, avg_pen_col)).value)

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

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

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
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(wb: xw.Book, metadata: ModelMetadata, source_file: str) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        print(f"skipped regression in {source_file}: missing sheet 'Regression Model'")
        return []

    sheet = wb.sheets["Regression Model"]
    snap = SheetSnapshot(sheet)
    anchor = snap.find_max_anchor()
    if anchor is None:
        print(f"skipped regression in {source_file}: could not find 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    header_map = snap.identify_headers_near_anchor(anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    history_rows = snap.paired_numeric_rows(x_col, y_col, max_row=anchor_row - 1)
    if not history_rows:
        history_rows = snap.paired_numeric_rows(x_col, y_col)

    scratch_intercept_col = snap.last_col + 2
    scratch_slope_col = snap.last_col + 3

    max_n = min(N_QUARTERS, len(history_rows))
    for i in range(1, max_n + 1):
        sample_rows = history_rows[-i:]
        start_row = sample_rows[0]
        end_row = sample_rows[-1]
        target_row = anchor_row + i

        intercept_formula = (
            f'=IFERROR(INTERCEPT({r1c1_range(start_row, y_col, end_row, y_col)},'
            f"{r1c1_range(start_row, x_col, end_row, x_col)}),\"\")"
        )
        slope_formula = (
            f'=IFERROR(SLOPE({r1c1_range(start_row, y_col, end_row, y_col)},'
            f"{r1c1_range(start_row, x_col, end_row, x_col)}),\"\")"
        )
        set_formula2(sheet.range((target_row, scratch_intercept_col)), intercept_formula)
        set_formula2(sheet.range((target_row, scratch_slope_col)), slope_formula)

    if max_n > 0:
        wb.app.calculate()

    num_col = resolve_col(header_map, "num_quarters_used", anchor_col, REGRESSION_FALLBACK_OFFSETS)
    forecast_col = resolve_col(header_map, "forecast_value", anchor_col, REGRESSION_FALLBACK_OFFSETS)
    actual_col = header_map.get("actual_value")
    max_col = resolve_col(header_map, "forecast_max", anchor_col, REGRESSION_FALLBACK_OFFSETS)
    min_col = resolve_col(header_map, "forecast_min", anchor_col, REGRESSION_FALLBACK_OFFSETS)

    latest_x = to_float(sheet.range((history_rows[-1], x_col)).value) if history_rows else None

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for i in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + i
        num_quarters_used = to_int(sheet.range((row_idx, num_col)).value) or i

        intercept_value = to_float(sheet.range((row_idx, scratch_intercept_col)).value)
        slope_value = to_float(sheet.range((row_idx, scratch_slope_col)).value)

        forecast_value = to_float(sheet.range((row_idx, forecast_col)).value)
        if forecast_value is None and intercept_value is not None and slope_value is not None and latest_x is not None:
            forecast_value = intercept_value + slope_value * (latest_x + 1)

        actual_value = to_float(sheet.range((row_idx, actual_col)).value) if actual_col is not None else None
        forecast_max = to_float(sheet.range((row_idx, max_col)).value)
        forecast_min = to_float(sheet.range((row_idx, min_col)).value)

        if all(value is None for value in (forecast_value, forecast_max, forecast_min, intercept_value, slope_value)):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept_value, 10) if intercept_value is not None else None,
            round(slope_value, 10) if slope_value is not None else None,
        )

        # Some model files duplicate the last candidate row; skip that final duplicate.
        if i == N_QUARTERS and previous_signature is not None and signature == previous_signature:
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
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def next_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.name
    base_name = f"{folder_name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    i = 1
    while True:
        candidate = output_path / f"{base_name}.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def apply_sheet_formatting(ws, headers: List[str], row_count: int) -> None:
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(row_count, 1)}"

    for col_idx in range(1, len(headers) + 1):
        max_len = len(headers[col_idx - 1])
        for row_idx in range(2, row_count + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            length = len(str(value)) if value is not None else 0
            if length > max_len:
                max_len = length
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 50)


def write_output_workbook(
    output_file: Path,
    empirical_rows: Iterable[Dict[str, Any]],
    regression_rows: Iterable[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    ws_emp = workbook.create_sheet("empirical_candidates")
    ws_reg = workbook.create_sheet("regression_candidates")

    emp_rows_list = list(empirical_rows)
    reg_rows_list = list(regression_rows)

    for row_idx, row_data in enumerate(emp_rows_list, start=2):
        for col_idx, col_name in enumerate(EMPIRICAL_COLUMNS, start=1):
            ws_emp.cell(row=row_idx, column=col_idx, value=row_data.get(col_name))

    for row_idx, row_data in enumerate(reg_rows_list, start=2):
        for col_idx, col_name in enumerate(REGRESSION_COLUMNS, start=1):
            ws_reg.cell(row=row_idx, column=col_idx, value=row_data.get(col_name))

    apply_sheet_formatting(ws_emp, EMPIRICAL_COLUMNS, len(emp_rows_list) + 1)
    apply_sheet_formatting(ws_reg, REGRESSION_COLUMNS, len(reg_rows_list) + 1)

    workbook.save(output_file)


def iter_input_files(input_path: Path) -> List[Path]:
    files: List[Path] = []
    for path in sorted(input_path.iterdir()):
        if path.is_dir():
            print(f"skipped file: {path.name} (reason: directory)")
            continue
        if path.name.startswith("~"):
            print(f"skipped file: {path.name} (reason: temp file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"skipped file: {path.name} (reason: not .xlsx)")
            continue
        files.append(path)
    return files


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    source_files = iter_input_files(input_path)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"processing file: {file_path.name}")
            metadata = parse_model_metadata(file_path.name)
            wb: Optional[xw.Book] = None

            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                files_processed += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (reason: {exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    output_file = next_output_path(input_path, output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
