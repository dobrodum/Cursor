from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import math
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ========= User-configurable paths =========
input_dir = "/workspace/input"
output_dir = "/workspace/output"
# ==========================================


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


@dataclass
class FileMetadata:
    ticker: str
    model_period: str
    model_date: str
    model: str


@dataclass
class SheetScan:
    start_row: int
    start_col: int
    values: List[List[Any]]
    labels: Dict[str, Tuple[int, int]]
    max_anchor: Optional[Tuple[int, int]]

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        width = max((len(row) for row in self.values), default=0)
        return self.start_col + width - 1

    def get_value(self, row: int, col: int) -> Any:
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if r_idx < 0 or c_idx < 0:
            return None
        if r_idx >= len(self.values):
            return None
        row_vals = self.values[r_idx]
        if c_idx >= len(row_vals):
            return None
        return row_vals[c_idx]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        f = float(value)
        if math.isnan(f):
            return None
        return f
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            try:
                return float(cleaned[:-1]) / 100.0
            except ValueError:
                return None
        try:
            f = float(cleaned)
            if math.isnan(f):
                return None
            return f
        except ValueError:
            return None
    return None


def safe_div(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def safe_sub(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return a - b


def parse_month_token(token: str) -> Optional[int]:
    month_lookup = {
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
    return month_lookup.get(token[:3].lower())


def parse_filename_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem

    ticker = "UNKNOWN"
    ticker_match = re.search(r"-\s*([A-Za-z0-9]{1,12})\s*-", stem)
    if ticker_match:
        ticker = ticker_match.group(1).upper()
    else:
        token_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        if token_match:
            ticker = token_match.group(0).upper()

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12]\d{3})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "UnknownPeriod"
    model_date = ""
    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month_num = parse_month_token(month_token)
        if month_num:
            month_abbrev = date(year, month_num, 1).strftime("%b")
            day = {"Early": 5, "Mid": 15, "Late": 25}[phase]
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return FileMetadata(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def find_sheet_case_insensitive(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def scan_sheet(sheet: xw.Sheet) -> SheetScan:
    used = sheet.used_range
    values = as_2d(used.value)
    start_row = used.row
    start_col = used.column

    labels: Dict[str, Tuple[int, int]] = {}
    max_anchor: Optional[Tuple[int, int]] = None
    wanted_labels = [
        "min",
        "max",
        "penetration",
        "reported sales",
        "quarterly sales",
        "growth rate",
        "sales captured in db",
        "tot fcst w/o sa",
        "actual",
    ]

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if not isinstance(cell_value, str):
                continue
            normalized = normalize_text(cell_value)
            absolute_row = start_row + r_idx
            absolute_col = start_col + c_idx
            if normalized == "max" and max_anchor is None:
                max_anchor = (absolute_row, absolute_col)
            for label in wanted_labels:
                if label in normalized and label not in labels:
                    labels[label] = (absolute_row, absolute_col)

    return SheetScan(
        start_row=start_row,
        start_col=start_col,
        values=values,
        labels=labels,
        max_anchor=max_anchor,
    )


def first_numeric_near(
    scan: SheetScan,
    row: int,
    col: int,
    offsets: Sequence[Tuple[int, int]],
) -> Optional[float]:
    for row_offset, col_offset in offsets:
        value = scan.get_value(row + row_offset, col + col_offset)
        number = to_float(value)
        if number is not None:
            return number
    return None


def extract_max_min(scan: SheetScan, anchor: Tuple[int, int]) -> Tuple[Optional[float], Optional[float]]:
    anchor_row, anchor_col = anchor
    max_value = first_numeric_near(
        scan,
        anchor_row,
        anchor_col,
        offsets=[(1, 0), (0, 1), (1, 1), (0, 2), (2, 0), (-1, 0), (0, -1)],
    )

    min_pos = scan.labels.get("min")
    if min_pos:
        min_value = first_numeric_near(
            scan,
            min_pos[0],
            min_pos[1],
            offsets=[(1, 0), (0, 1), (1, 1), (0, 2), (2, 0), (-1, 0), (0, -1)],
        )
    else:
        min_value = first_numeric_near(
            scan,
            anchor_row,
            anchor_col,
            offsets=[(2, 0), (0, 2), (2, 1), (1, 2), (3, 0), (0, 3)],
        )
    return max_value, min_value


def extract_xy_pairs(
    scan: SheetScan,
    x_col: int,
    y_col: int,
    data_end_row: int,
) -> List[Tuple[int, float, float]]:
    pairs: List[Tuple[int, float, float]] = []
    for row in range(scan.start_row, data_end_row + 1):
        x_val = to_float(scan.get_value(row, x_col))
        y_val = to_float(scan.get_value(row, y_col))
        if x_val is None or y_val is None:
            continue
        pairs.append((row, x_val, y_val))
    return pairs


def get_last_quarter_label(scan: SheetScan, row: int, x_col: int) -> str:
    for col in [x_col - 3, x_col - 2, x_col - 1, 1]:
        value = scan.get_value(row, col)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return f"row_{row}"


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
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

    cell.api.FormulaR1C1 = formula_r1c1


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
    except Exception:
        pass


def process_empirical_sheet(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        return []

    scan = scan_sheet(sheet)
    if scan.max_anchor is None:
        return []

    anchor_row, anchor_col = scan.max_anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    pairs = extract_xy_pairs(scan, x_col=x_col, y_col=y_col, data_end_row=anchor_row - 1)
    if not pairs:
        return []

    max_quarters = min(10, len(pairs))
    latest_row, latest_x, latest_y = pairs[-1]
    previous_y = pairs[-2][2] if len(pairs) >= 2 else None
    growth_rate = safe_div(latest_y, previous_y)
    if growth_rate is not None:
        growth_rate -= 1.0

    sales_captured_pct = safe_div(latest_x, latest_y)
    max_value, min_value = extract_max_min(scan, (anchor_row, anchor_col))
    last_quarter_used = get_last_quarter_label(scan, latest_row, x_col=x_col)

    penetration_col = None
    for label, (row, col) in scan.labels.items():
        if "penetration" in label and row < anchor_row:
            penetration_col = col
            break

    helper_col = anchor_col + 20
    helper_start_row = anchor_row + 2

    for i in range(max_quarters):
        n = i + 1
        start_row = pairs[-n][0]
        target_row = helper_start_row + i

        sheet.cells(target_row, helper_col).value = n
        avg_pen_cell = sheet.cells(target_row, helper_col + 1)
        forecast_cell = sheet.cells(target_row, helper_col + 2)

        if penetration_col is not None:
            avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{latest_row}C{penetration_col})"
        else:
            avg_formula = (
                f"=AVERAGE(R{start_row}C{x_col}:R{latest_row}C{x_col}/"
                f"R{start_row}C{y_col}:R{latest_row}C{y_col})"
            )

        set_formula2_r1c1(avg_pen_cell, avg_formula)
        set_formula2_r1c1(
            forecast_cell,
            f"=IF(RC[-1]=0,NA(),{latest_x}/RC[-1])",
        )

    wb.app.calculate()

    result_block = as_2d(
        sheet.range(
            (helper_start_row, helper_col + 1),
            (helper_start_row + max_quarters - 1, helper_col + 2),
        ).value
    )

    rows: List[Dict[str, Any]] = []
    for i in range(max_quarters):
        n = i + 1
        avg_penetration = to_float(result_block[i][0]) if i < len(result_block) and len(result_block[i]) > 0 else None
        forecast_total = to_float(result_block[i][1]) if i < len(result_block) and len(result_block[i]) > 1 else None

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_total,
            "actual_value": latest_y,
            "forecast_max": max_value,
            "forecast_min": min_value,
            "range_width": safe_sub(max_value, min_value),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": latest_x,
            "reported_sales": latest_y,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def process_regression_sheet(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        return []

    scan = scan_sheet(sheet)
    if scan.max_anchor is None:
        return []

    anchor_row, anchor_col = scan.max_anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    pairs = extract_xy_pairs(scan, x_col=x_col, y_col=y_col, data_end_row=anchor_row - 1)
    if len(pairs) < 2:
        return []

    max_quarters = min(10, len(pairs))
    latest_row, latest_x, _ = pairs[-1]
    max_value, min_value = extract_max_min(scan, (anchor_row, anchor_col))

    helper_col = anchor_col + 20
    helper_start_row = anchor_row + 2

    for i in range(max_quarters):
        n = i + 1
        start_row = pairs[-n][0]
        target_row = helper_start_row + i

        sheet.cells(target_row, helper_col).value = n
        intercept_cell = sheet.cells(target_row, helper_col + 1)
        slope_cell = sheet.cells(target_row, helper_col + 2)
        forecast_cell = sheet.cells(target_row, helper_col + 3)

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_row}C{x_col}:R{latest_row}C{x_col})"
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, f"=RC[-2]+RC[-1]*{latest_x}")

    wb.app.calculate()

    result_block = as_2d(
        sheet.range(
            (helper_start_row, helper_col),
            (helper_start_row + max_quarters - 1, helper_col + 3),
        ).value
    )

    actual_value = None
    actual_label_pos = scan.labels.get("actual")
    if actual_label_pos:
        actual_value = first_numeric_near(
            scan,
            actual_label_pos[0],
            actual_label_pos[1],
            offsets=[(1, 0), (0, 1), (1, 1), (0, 2), (2, 0)],
        )

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Optional[float], Optional[float], Optional[float], Optional[float], Optional[float]]] = None
    for i in range(max_quarters):
        if i >= len(result_block):
            continue
        current = result_block[i]
        if len(current) < 4:
            continue

        n = int(to_float(current[0]) or (i + 1))
        intercept_val = to_float(current[1])
        slope_val = to_float(current[2])
        forecast_wo_sa = to_float(current[3])

        signature = (
            round(intercept_val, 10) if intercept_val is not None else None,
            round(slope_val, 10) if slope_val is not None else None,
            round(forecast_wo_sa, 10) if forecast_wo_sa is not None else None,
            round(max_value, 10) if max_value is not None else None,
            round(min_value, 10) if min_value is not None else None,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n,
            "num_quarters_used": n,
            "forecast_value": forecast_wo_sa,
            "actual_value": actual_value,
            "forecast_max": max_value,
            "forecast_min": min_value,
            "range_width": safe_sub(max_value, min_value),
            "intercept": intercept_val,
            "slope": slope_val,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_rows_to_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_index in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_index).font = Font(bold=True)

        max_len = len(columns[col_index - 1])
        for row_index in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_index, column=col_index).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max(max_len + 2, 12), 44)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise NotADirectoryError(f"Input directory does not exist: {input_path}")

    files_to_process: List[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"SKIP: {file_path.name} (temporary Excel file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"SKIP: {file_path.name} (not .xlsx)")
            continue
        files_to_process.append(file_path)

    if not files_to_process:
        print(f"No .xlsx files found in: {input_path}")
        return

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in files_to_process:
            print(f"PROCESS: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path)

                empirical_rows.extend(
                    process_empirical_sheet(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        wb=wb,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                processed_count += 1
            except Exception as exc:
                print(f"SKIP: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    final_output = build_output_path(input_folder=input_path, output_folder=output_path)
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    write_rows_to_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    workbook.save(final_output)

    print(f"OUTPUT: {final_output}")
    print(f"FILES_PROCESSED: {processed_count}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()
