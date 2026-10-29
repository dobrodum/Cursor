from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

try:
    import xlwings as xw
except ImportError as exc:
    raise SystemExit("Missing dependency: xlwings. Install with `pip install xlwings`.") from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit("Missing dependency: openpyxl. Install with `pip install openpyxl`.") from exc


# --------- user-configurable paths ----------
input_dir = "./input"
output_dir = "./output"
# -------------------------------------------


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

PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


class SheetSnapshot:
    def __init__(self, first_row: int, first_col: int, matrix: list[list[Any]]):
        self.first_row = first_row
        self.first_col = first_col
        self.matrix = matrix
        self.row_count = len(matrix)
        self.col_count = len(matrix[0]) if matrix else 0
        self.last_row = self.first_row + self.row_count - 1 if self.row_count else self.first_row
        self.last_col = self.first_col + self.col_count - 1 if self.col_count else self.first_col
        self.label_positions: dict[str, list[tuple[int, int]]] = {}
        self._build_label_index()

    def _build_label_index(self) -> None:
        for r_idx, row_vals in enumerate(self.matrix):
            for c_idx, value in enumerate(row_vals):
                key = normalize_label(value)
                if not key:
                    continue
                row = self.first_row + r_idx
                col = self.first_col + c_idx
                self.label_positions.setdefault(key, []).append((row, col))

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.first_row
        c_idx = col - self.first_col
        if r_idx < 0 or c_idx < 0:
            return None
        if r_idx >= self.row_count or c_idx >= self.col_count:
            return None
        return self.matrix[r_idx][c_idx]


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if len(values) == 0:
        return []
    first = values[0]
    if isinstance(first, (list, tuple)):
        return [list(row) for row in values]
    return [list(values)]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    if not isinstance(value, str):
        return ""
    return re.sub(r"\s+", " ", value.strip().lower())


def as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    pct = text.endswith("%")
    text = text.replace(",", "").replace("%", "")
    try:
        numeric = float(text)
    except ValueError:
        return None
    return numeric / 100.0 if pct else numeric


def as_int(value: Any) -> int | None:
    numeric = as_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def as_clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text


def normalize_vector(values: Any, expected_len: int) -> list[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            output = [row[0] if row else None for row in values]
        else:
            output = values
    else:
        output = [values]
    if len(output) < expected_len:
        output.extend([None] * (expected_len - len(output)))
    return output[:expected_len]


def parse_model_metadata(file_name: str) -> ModelMetadata:
    stem = Path(file_name).stem
    parts = [chunk.strip() for chunk in stem.split(" - ") if chunk.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].replace(" ", "").upper()
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "UNKNOWN"
    model_date = ""
    if period_match:
        period_bucket = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month_key = month_token.lower()
        month_num = MONTH_LOOKUP.get(month_key) or MONTH_LOOKUP.get(month_key[:3])
        if month_num:
            month_abbr = calendar.month_abbr[month_num]
            model_period = f"{period_bucket}{month_abbr}_{year}"
            day = PERIOD_DAY.get(period_bucket, 15)
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period != "UNKNOWN" else ticker
    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    version = 1
    while candidate.exists():
        candidate = output_path / f"{input_path.name}_PARAM.{version}.xlsx"
        version += 1
    return candidate


def capture_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    return SheetSnapshot(used.row, used.column, matrix)


def find_anchor_cell(snapshot: SheetSnapshot, target: str = "max") -> tuple[int, int] | None:
    desired = target.lower().strip()
    for r_idx, row_vals in enumerate(snapshot.matrix):
        for c_idx, value in enumerate(row_vals):
            if isinstance(value, str) and value.strip().lower() == desired:
                return snapshot.first_row + r_idx, snapshot.first_col + c_idx
    return None


def find_nearest_label(
    snapshot: SheetSnapshot,
    labels: list[str],
    anchor_row: int,
    anchor_col: int,
) -> tuple[int, int] | None:
    norm_labels = [normalize_label(label) for label in labels]
    best: tuple[int, int] | None = None
    best_score: tuple[int, int] | None = None

    for key, positions in snapshot.label_positions.items():
        for label in norm_labels:
            if not label:
                continue
            if key == label or key.startswith(label) or label in key:
                for row, col in positions:
                    score = (abs(row - anchor_row), abs(col - anchor_col))
                    if best_score is None or score < best_score:
                        best_score = score
                        best = (row, col)
    return best


def trailing_numeric_rows(snapshot: SheetSnapshot, col: int, end_row: int, max_scan: int = 300) -> list[int]:
    start_row = max(snapshot.first_row, end_row - max_scan + 1)
    rows = []
    for row in range(start_row, end_row + 1):
        if as_float(snapshot.get(row, col)) is not None:
            rows.append(row)
    if not rows:
        return []

    # Keep only the trailing contiguous block.
    contiguous = [rows[-1]]
    for row in reversed(rows[:-1]):
        if contiguous[0] - row == 1:
            contiguous.insert(0, row)
        else:
            break
    return contiguous


def trailing_xy_rows(
    snapshot: SheetSnapshot,
    x_col: int,
    y_col: int,
    end_row: int,
    max_scan: int = 300,
) -> list[int]:
    start_row = max(snapshot.first_row, end_row - max_scan + 1)
    rows = []
    for row in range(start_row, end_row + 1):
        x_val = as_float(snapshot.get(row, x_col))
        y_val = as_float(snapshot.get(row, y_col))
        if x_val is not None and y_val is not None:
            rows.append(row)
    if not rows:
        return []

    contiguous = [rows[-1]]
    for row in reversed(rows[:-1]):
        if contiguous[0] - row == 1:
            contiguous.insert(0, row)
        else:
            break
    return contiguous


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Compatibility fallback for older Excel builds.
        cell.formula = formula_r1c1


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    fallback_calls = [
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ]
    for call in fallback_calls:
        try:
            call()
            return
        except Exception:
            continue


def extract_empirical_candidates(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = capture_sheet_snapshot(sheet)
    anchor = find_anchor_cell(snapshot, "max")
    if not anchor:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    min_label = find_nearest_label(snapshot, ["min"], anchor_row, anchor_col)
    min_col = min_label[1] if min_label and min_label[0] == anchor_row else anchor_col + 1

    cols = {
        "num_quarters_col": anchor_col - 12,
        "last_quarter_col": anchor_col - 10,
        "forecast_col": anchor_col - 9,
        "actual_col": anchor_col - 8,
        "quarterly_sales_col": anchor_col - 6,
        "reported_sales_col": anchor_col - 8,
        "growth_rate_col": anchor_col - 4,
        "captured_col": anchor_col - 3,
        "penetration_col": anchor_col - 5,
        "max_col": anchor_col,
        "min_col": min_col,
    }

    label_map = {
        "num_quarters_col": ["num_quarters_used", "num quarters used", "quarters used"],
        "last_quarter_col": ["last_quarter_used", "last quarter used", "last quarter"],
        "forecast_col": ["estimated total sold", "forecast value", "tot fcst"],
        "actual_col": ["actual value", "actual sales", "actual"],
        "quarterly_sales_col": ["quarterly sales", "sales in db"],
        "reported_sales_col": ["reported sales", "total sales", "reported"],
        "growth_rate_col": ["growth rate", "growth"],
        "captured_col": ["sales captured in db", "captured in db", "capture"],
        "penetration_col": ["avg penetration", "penetration"],
    }

    start_row = anchor_row + 1
    num_quarter_label = find_nearest_label(snapshot, label_map["num_quarters_col"], anchor_row, anchor_col)
    if num_quarter_label:
        start_row = num_quarter_label[0] + 1
        cols["num_quarters_col"] = num_quarter_label[1]

    for key, labels in label_map.items():
        if key == "num_quarters_col":
            continue
        label_pos = find_nearest_label(snapshot, labels, anchor_row, anchor_col)
        if label_pos:
            cols[key] = label_pos[1]

    n_quarters = 10
    history_rows = trailing_numeric_rows(snapshot, cols["penetration_col"], start_row - 1, max_scan=240)

    temp_col = max(snapshot.last_col + 2, anchor_col + 15)
    temp_start_row = max(snapshot.last_row + 2, anchor_row + 2)
    temp_range = sheet.range((temp_start_row, temp_col), (temp_start_row + n_quarters - 1, temp_col))

    for idx in range(n_quarters):
        count = idx + 1
        if len(history_rows) >= count:
            start = history_rows[-count]
            end = history_rows[-1]
            formula = f'=IFERROR(AVERAGE(R{start}C{cols["penetration_col"]}:R{end}C{cols["penetration_col"]}),"")'
        else:
            formula = '=""'
        set_formula2(sheet.range((temp_start_row + idx, temp_col)), formula)

    wb.app.calculate()
    avg_penetration_values = normalize_vector(temp_range.value, n_quarters)
    temp_range.value = None

    rows: list[dict[str, Any]] = []
    for idx in range(n_quarters):
        data_row = start_row + idx
        num_quarters_used = as_int(snapshot.get(data_row, cols["num_quarters_col"])) or (idx + 1)
        last_quarter_used = as_clean_text(snapshot.get(data_row, cols["last_quarter_col"]))

        avg_penetration_pct = as_float(avg_penetration_values[idx])
        if avg_penetration_pct is None:
            avg_penetration_pct = as_float(snapshot.get(data_row, cols["penetration_col"]))

        quarterly_sales = as_float(snapshot.get(data_row, cols["quarterly_sales_col"]))
        reported_sales = as_float(snapshot.get(data_row, cols["reported_sales_col"]))
        growth_rate_pct = as_float(snapshot.get(data_row, cols["growth_rate_col"]))
        sales_captured_in_db_pct = as_float(snapshot.get(data_row, cols["captured_col"]))

        forecast_value = as_float(snapshot.get(data_row, cols["forecast_col"]))
        if forecast_value is None and avg_penetration_pct not in (None, 0):
            numerator = reported_sales if reported_sales is not None else quarterly_sales
            if numerator is not None:
                forecast_value = numerator / avg_penetration_pct

        forecast_max = as_float(snapshot.get(data_row, cols["max_col"]))
        forecast_min = as_float(snapshot.get(data_row, cols["min_col"]))
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration_pct is None
            and quarterly_sales is None
            and reported_sales is None
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row = {
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
        rows.append(row)

    return rows


def _rows_equivalent_for_duplicate_guard(a: dict[str, Any], b: dict[str, Any]) -> bool:
    numeric_keys = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    for key in numeric_keys:
        a_val = as_float(a.get(key))
        b_val = as_float(b.get(key))
        if a_val is None and b_val is None:
            continue
        if a_val is None or b_val is None:
            return False
        if abs(a_val - b_val) > 1e-9:
            return False
    return as_int(a.get("num_quarters_used")) == as_int(b.get("num_quarters_used"))


def extract_regression_candidates(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = capture_sheet_snapshot(sheet)
    anchor = find_anchor_cell(snapshot, "max")
    if not anchor:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    min_label = find_nearest_label(snapshot, ["min"], anchor_row, anchor_col)
    min_col = min_label[1] if min_label and min_label[0] == anchor_row else anchor_col + 1

    forecast_total_col = anchor_col - 1
    forecast_label = find_nearest_label(
        snapshot,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast total without sa", "tot fcst"],
        anchor_row,
        anchor_col,
    )
    if forecast_label:
        forecast_total_col = forecast_label[1]

    actual_col = None
    actual_label = find_nearest_label(snapshot, ["actual value", "actual sales", "actual"], anchor_row, anchor_col)
    if actual_label:
        actual_col = actual_label[1]

    xy_rows = trailing_xy_rows(snapshot, x_col, y_col, anchor_row - 1, max_scan=360)
    if len(xy_rows) < 2:
        print(f"Skipped regression extraction for {source_file}: not enough x/y data.")
        return []

    n_quarters = min(10, len(xy_rows))
    temp_start_row = max(snapshot.last_row + 2, anchor_row + 2)
    temp_intercept_col = max(snapshot.last_col + 2, anchor_col + 15)
    temp_slope_col = temp_intercept_col + 1

    for idx in range(n_quarters):
        count = idx + 1
        sample_rows = xy_rows[-count:]
        first_row = sample_rows[0]
        last_row = sample_rows[-1]
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},'
            f'R{first_row}C{x_col}:R{last_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},'
            f'R{first_row}C{x_col}:R{last_row}C{x_col}),"")'
        )
        set_formula2(sheet.range((temp_start_row + idx, temp_intercept_col)), intercept_formula)
        set_formula2(sheet.range((temp_start_row + idx, temp_slope_col)), slope_formula)

    wb.app.calculate()
    intercept_values = normalize_vector(
        sheet.range((temp_start_row, temp_intercept_col), (temp_start_row + n_quarters - 1, temp_intercept_col)).value,
        n_quarters,
    )
    slope_values = normalize_vector(
        sheet.range((temp_start_row, temp_slope_col), (temp_start_row + n_quarters - 1, temp_slope_col)).value,
        n_quarters,
    )

    sheet.range((temp_start_row, temp_intercept_col), (temp_start_row + n_quarters - 1, temp_slope_col)).value = None

    forecast_x = as_float(snapshot.get(anchor_row, x_col))
    if forecast_x is None:
        forecast_x = as_float(snapshot.get(xy_rows[-1], x_col))
        if forecast_x is not None and len(xy_rows) >= 2:
            prev_x = as_float(snapshot.get(xy_rows[-2], x_col))
            if prev_x is not None:
                forecast_x += forecast_x - prev_x

    rows: list[dict[str, Any]] = []
    for idx in range(n_quarters):
        num_quarters_used = idx + 1
        intercept = as_float(intercept_values[idx])
        slope = as_float(slope_values[idx])
        if intercept is None or slope is None:
            continue

        data_row = anchor_row + idx + 1
        forecast_value = as_float(snapshot.get(data_row, forecast_total_col))
        if forecast_value is None and forecast_x is not None:
            forecast_value = intercept + slope * forecast_x

        forecast_max = as_float(snapshot.get(data_row, anchor_col))
        forecast_min = as_float(snapshot.get(data_row, min_col))
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        actual_value = as_float(snapshot.get(data_row, actual_col)) if actual_col else None
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row = {
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
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and _rows_equivalent_for_duplicate_guard(row, rows[-1]):
            continue
        rows.append(row)

    return rows


def write_sheet(
    ws,
    headers: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col_letter}{max(ws.max_row, 1)}"

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for (value,) in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=idx, max_col=idx, values_only=True):
            if value is None:
                continue
            text = value.isoformat() if isinstance(value, date) else str(value)
            max_len = max(max_len, len(text))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def save_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    wb.save(output_file)


def process_workbooks(input_path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook.")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file.")
                continue

            print(f"Processing {file_path.name}")

            try:
                source_wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: open failed ({exc}).")
                continue

            try:
                metadata = parse_model_metadata(file_path.name)

                try:
                    empirical_sheet = source_wb.sheets["Empirical Model"]
                except Exception:
                    empirical_sheet = None
                    print(f"Skipped empirical extraction for {file_path.name}: sheet not found.")

                if empirical_sheet is not None:
                    empirical_rows.extend(
                        extract_empirical_candidates(
                            wb=source_wb,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                try:
                    regression_sheet = source_wb.sheets["Regression Model"]
                except Exception:
                    regression_sheet = None
                    print(f"Skipped regression extraction for {file_path.name}: sheet not found.")

                if regression_sheet is not None:
                    regression_rows.extend(
                        extract_regression_candidates(
                            wb=source_wb,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc}).")
            finally:
                close_workbook_safely(source_wb)
    finally:
        app.quit()

    return empirical_rows, regression_rows, processed_files


def main() -> int:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        print(f"Input directory not found: {input_path}")
        return 1

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_path, output_path)

    empirical_rows, regression_rows, processed_files = process_workbooks(input_path)
    save_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output written to: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
