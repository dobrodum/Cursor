#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# Configure inputs/outputs here
# -----------------------------
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

PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", re.IGNORECASE)
PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = text.replace("_", " ")
    text = re.sub(r"[^a-z0-9% ]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def is_number(value: Any) -> bool:
    try:
        if value is None:
            return False
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def to_number(value: Any) -> Optional[float]:
    return float(value) if is_number(value) else None


def subtract_if_numeric(a: Any, b: Any) -> Optional[float]:
    a_num = to_number(a)
    b_num = to_number(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def value_signature(value: Any) -> Any:
    if is_number(value):
        return round(float(value), 10)
    return value


def normalize_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


@dataclass
class SheetSnapshot:
    sheet: xw.Sheet
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: List[List[Any]]

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or row > self.end_row:
            return None
        if col < self.start_col or col > self.end_col:
            return None
        return self.values[row - self.start_row][col - self.start_col]


def build_snapshot(sheet: xw.Sheet) -> Optional[SheetSnapshot]:
    used = sheet.used_range
    matrix = normalize_matrix(used.value)
    if not matrix:
        return None
    row_count = len(matrix)
    col_count = max((len(r) for r in matrix), default=0)
    if col_count == 0:
        return None
    # Normalize row widths for safe indexing.
    for row in matrix:
        if len(row) < col_count:
            row.extend([None] * (col_count - len(row)))
    start_row = used.row
    start_col = used.column
    end_row = start_row + row_count - 1
    end_col = start_col + col_count - 1
    return SheetSnapshot(sheet=sheet, start_row=start_row, start_col=start_col, end_row=end_row, end_col=end_col, values=matrix)


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int]] = []
    for r in range(snapshot.start_row, snapshot.end_row + 1):
        for c in range(snapshot.start_col, snapshot.end_col + 1):
            if normalize_text(snapshot.get(r, c)) == "max":
                candidates.append((r, c))
    if not candidates:
        return None

    # Prefer an anchor that has "min" adjacent in the same row.
    for r, c in candidates:
        if normalize_text(snapshot.get(r, c + 1)) == "min" or normalize_text(snapshot.get(r, c - 1)) == "min":
            return (r, c)
    return candidates[0]


def build_header_map(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int, window: int = 24) -> Dict[int, str]:
    header_map: Dict[int, str] = {}
    for row in (anchor_row - 1, anchor_row):
        if row < snapshot.start_row or row > snapshot.end_row:
            continue
        for col in range(max(snapshot.start_col, anchor_col - window), min(snapshot.end_col, anchor_col + window) + 1):
            text = normalize_text(snapshot.get(row, col))
            if text and col not in header_map:
                header_map[col] = text
    return header_map


def find_col(header_map: Dict[int, str], patterns: Sequence[str], default: Optional[int] = None) -> Optional[int]:
    for col, text in header_map.items():
        for pattern in patterns:
            if re.search(pattern, text):
                return col
    return default


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    # Formula2R1C1 is fastest and avoids column letter conversion.
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass
    cell.formula = formula


def close_workbook_safely(wb: xw.Book) -> None:
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


def parse_month_number(token: str) -> Optional[int]:
    cleaned = re.sub(r"[^A-Za-z]", "", token).title()
    if not cleaned:
        return None
    short = cleaned[:3]
    for month_num in range(1, 13):
        if short == calendar.month_abbr[month_num]:
            return month_num
    return None


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 else ""
    period_raw = parts[2] if len(parts) > 2 else ""
    period_token = period_raw.split("_")[0].strip()

    model_period = period_token
    model_date = ""

    match = PERIOD_RE.search(period_token)
    if match:
        period_tag = match.group(1).title()
        month_token = match.group(2)
        year_num = int(match.group(3))
        month_num = parse_month_number(month_token)
        if month_num is not None:
            month_abbr = calendar.month_abbr[month_num]
            model_period = f"{period_tag}{month_abbr}_{year_num}"
            model_date = date(year_num, month_num, PERIOD_DAY[period_tag]).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or stem)
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_dir / f"{base}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def get_sheet(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[name]
    except Exception:
        return None


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = build_snapshot(sheet)
    if snapshot is None:
        return []

    anchor = find_max_anchor(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row, anchor_col)

    num_quarters_col = find_col(header_map, [r"\bnum\b.*\bquarter", r"\bquarters?\s*used\b", r"\bn\s*quarters?\b"], default=anchor_col - 6)
    last_quarter_col = find_col(header_map, [r"\blast\b.*\bquarter", r"\blast\b.*\bqtr"], default=anchor_col - 5)
    forecast_value_col = find_col(header_map, [r"\bestimated total sold\b", r"\bforecast value\b", r"\btot\b.*\bfcst\b"], default=anchor_col - 1)
    actual_value_col = find_col(header_map, [r"\breported sales\b", r"\bactual\b"], default=anchor_col - 2)
    forecast_max_col = anchor_col
    forecast_min_col = find_col(header_map, [r"^min$"], default=anchor_col + 1)
    avg_pen_col = find_col(header_map, [r"\bavg\b.*\bpenetration", r"\bpenetration\b.*%"])
    quarterly_sales_col = find_col(header_map, [r"\bquarterly sales\b", r"\bqtr\b.*\bsales\b"], default=anchor_col - 10)
    reported_sales_col = find_col(header_map, [r"\breported sales\b", r"\bsales reported\b"], default=anchor_col - 9)
    growth_rate_col = find_col(header_map, [r"\bgrowth\b.*\brate\b", r"\bgrowth\b.*%"], default=anchor_col - 8)
    sales_captured_col = find_col(header_map, [r"\bcaptured\b.*\bdb\b", r"\bsales captured\b"], default=anchor_col - 7)

    output_rows: List[Dict[str, Any]] = []
    formula_rows: List[Tuple[int, int]] = []
    scratch_col = max(snapshot.end_col + 2, anchor_col + 15)
    empty_streak = 0

    for i in range(N_QUARTERS):
        row = anchor_row + 1 + i
        num_quarters_used = snapshot.get(row, num_quarters_col) if num_quarters_col else None
        last_quarter_used = snapshot.get(row, last_quarter_col) if last_quarter_col else None
        forecast_value = snapshot.get(row, forecast_value_col) if forecast_value_col else None
        actual_value = snapshot.get(row, actual_value_col) if actual_value_col else None
        forecast_max = snapshot.get(row, forecast_max_col)
        forecast_min = snapshot.get(row, forecast_min_col) if forecast_min_col else None
        quarterly_sales = snapshot.get(row, quarterly_sales_col) if quarterly_sales_col else None
        reported_sales = snapshot.get(row, reported_sales_col) if reported_sales_col else None
        growth_rate_pct = snapshot.get(row, growth_rate_col) if growth_rate_col else None
        sales_captured_in_db_pct = snapshot.get(row, sales_captured_col) if sales_captured_col else None
        avg_penetration_pct = snapshot.get(row, avg_pen_col) if avg_pen_col else None

        if all(
            value in (None, "")
            for value in (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
            )
        ):
            empty_streak += 1
            if empty_streak >= 3:
                break
            continue
        empty_streak = 0

        row_data: Dict[str, Any] = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used if num_quarters_used not in (None, "") else i + 1,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_if_numeric(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        output_rows.append(row_data)

        # If avg penetration isn't directly populated, compute it with a temporary R1C1 formula.
        if not is_number(avg_penetration_pct) and quarterly_sales_col is not None and reported_sales_col is not None:
            reported_offset = reported_sales_col - scratch_col
            quarterly_offset = quarterly_sales_col - scratch_col
            formula = f'=IFERROR(RC[{reported_offset}]/RC[{quarterly_offset}], "")'
            set_formula2_r1c1(sheet.range((row, scratch_col)), formula)
            formula_rows.append((len(output_rows) - 1, row))

    if formula_rows:
        wb.app.calculate()
        for output_index, row in formula_rows:
            avg_value = sheet.range((row, scratch_col)).value
            output_rows[output_index]["avg_penetration_pct"] = avg_value
            output_rows[output_index]["parameter_value"] = avg_value

    return output_rows


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = build_snapshot(sheet)
    if snapshot is None:
        return []

    anchor = find_max_anchor(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = find_col(header_map, [r"\bnum\b.*\bquarter", r"\bquarters?\s*used\b"], default=anchor_col - 10)
    forecast_value_col = find_col(header_map, [r"\btot\b.*\bfcst\b.*\bw/o\b.*\bsa\b", r"\bforecast\b.*\bwithout\b.*\bsa\b"], default=anchor_col - 2)
    actual_value_col = find_col(header_map, [r"\bactual\b", r"\breported\b.*\bsales\b"])
    forecast_max_col = anchor_col
    forecast_min_col = find_col(header_map, [r"^min$"], default=anchor_col + 1)

    data_rows: List[int] = []
    for row in range(snapshot.start_row, anchor_row):
        x_val = snapshot.get(row, x_col)
        y_val = snapshot.get(row, y_col)
        if is_number(x_val) and is_number(y_val):
            data_rows.append(row)

    if len(data_rows) < 2:
        return []

    n_rows = min(N_QUARTERS, len(data_rows))
    scratch_start_col = max(snapshot.end_col + 3, anchor_col + 15)

    for i in range(n_rows):
        n_quarters = i + 1
        recent_rows = data_rows[-n_quarters:]
        first_row = recent_rows[0]
        last_row = recent_rows[-1]
        target_row = anchor_row + 1 + i

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},'
            f'R{first_row}C{x_col}:R{last_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},'
            f'R{first_row}C{x_col}:R{last_row}C{x_col}),"")'
        )
        forecast_formula = f'=IFERROR(RC[-2]+RC[-1]*R{last_row}C{x_col},"")'

        set_formula2_r1c1(sheet.range((target_row, scratch_start_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((target_row, scratch_start_col + 1)), slope_formula)
        set_formula2_r1c1(sheet.range((target_row, scratch_start_col + 2)), forecast_formula)

    wb.app.calculate()

    output_rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for i in range(n_rows):
        n_quarters = i + 1
        row = anchor_row + 1 + i
        intercept_val = sheet.range((row, scratch_start_col)).value
        slope_val = sheet.range((row, scratch_start_col + 1)).value
        fallback_forecast_val = sheet.range((row, scratch_start_col + 2)).value

        num_quarters_used = snapshot.get(row, num_quarters_col) if num_quarters_col else None
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        forecast_value = snapshot.get(row, forecast_value_col) if forecast_value_col else None
        if forecast_value in (None, ""):
            forecast_value = fallback_forecast_val

        actual_value = snapshot.get(row, actual_value_col) if actual_value_col else ""
        forecast_max = snapshot.get(row, forecast_max_col)
        forecast_min = snapshot.get(row, forecast_min_col) if forecast_min_col else None
        if forecast_max in (None, "") and is_number(forecast_value):
            forecast_max = forecast_value
        if forecast_min in (None, "") and is_number(forecast_value):
            forecast_min = forecast_value

        row_data: Dict[str, Any] = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_if_numeric(forecast_max, forecast_min),
            "intercept": intercept_val,
            "slope": slope_val,
            "source_file": source_file,
        }

        signature = (
            value_signature(row_data["num_quarters_used"]),
            value_signature(row_data["forecast_value"]),
            value_signature(row_data["forecast_max"]),
            value_signature(row_data["forecast_min"]),
            value_signature(row_data["intercept"]),
            value_signature(row_data["slope"]),
        )
        # Avoid duplicate final row caused by repeated formula output.
        if prev_signature is not None and signature == prev_signature:
            continue
        output_rows.append(row_data)
        prev_signature = signature

    return output_rows


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(ws.max_row, 1)}"

    for col_idx, name in enumerate(columns, start=1):
        max_len = len(name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 45)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    empirical_sheet = wb.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = wb.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not .xlsx")
                continue

            print(f"Processing {file_path.name}")
            labels = parse_file_labels(file_path)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                processed_files += 1

                empirical_sheet = get_sheet(wb, "Empirical Model")
                regression_sheet = get_sheet(wb, "Regression Model")

                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name}: sheet 'Empirical Model' missing")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                if regression_sheet is None:
                    print(f"Skipped regression extraction for {file_path.name}: sheet 'Regression Model' missing")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
