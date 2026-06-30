from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure paths here.
input_dir = Path("input")
output_dir = Path("output")

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

PHASE_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}


@dataclass
class SheetSnapshot:
    values: list[list[Any]]
    start_row: int
    start_col: int

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        if not self.values:
            return self.start_col
        return self.start_col + len(self.values[0]) - 1

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or row > self.end_row:
            return None
        if col < self.start_col or col > self.end_col:
            return None
        return self.values[row - self.start_row][col - self.start_col]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def coerce_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def numeric_delta(a: Any, b: Any) -> float | None:
    left = coerce_float(a)
    right = coerce_float(b)
    if left is None or right is None:
        return None
    return left - right


def load_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = used.value
    if values is None:
        matrix = [[]]
    elif isinstance(values, list):
        if not values:
            matrix = [[]]
        elif isinstance(values[0], list):
            matrix = [list(row) for row in values]
        else:
            matrix = [list(values)]
    else:
        matrix = [[values]]

    if not matrix:
        matrix = [[]]

    width = max((len(r) for r in matrix), default=0)
    if width == 0:
        width = 1
    padded = [row + [None] * (width - len(row)) for row in matrix]
    return SheetSnapshot(values=padded, start_row=used.row, start_col=used.column)


def build_text_index(snapshot: SheetSnapshot) -> tuple[dict[str, list[tuple[int, int]]], list[tuple[str, int, int]]]:
    exact: dict[str, list[tuple[int, int]]] = {}
    entries: list[tuple[str, int, int]] = []
    for r_idx, row in enumerate(snapshot.values):
        row_num = snapshot.start_row + r_idx
        for c_idx, cell_value in enumerate(row):
            col_num = snapshot.start_col + c_idx
            token = normalize_text(cell_value)
            if not token:
                continue
            exact.setdefault(token, []).append((row_num, col_num))
            entries.append((token, row_num, col_num))
    return exact, entries


def find_anchor_max(snapshot: SheetSnapshot, entries: list[tuple[str, int, int]]) -> tuple[int, int] | None:
    for token, row, col in entries:
        if token == "max":
            return row, col
    for token, row, col in entries:
        if token.startswith("max"):
            return row, col
    return None


def find_label_cell(
    text_index: tuple[dict[str, list[tuple[int, int]]], list[tuple[str, int, int]]],
    aliases: list[str],
) -> tuple[int, int] | None:
    exact, entries = text_index
    normalized_aliases = [normalize_text(a) for a in aliases if normalize_text(a)]
    for alias in normalized_aliases:
        if alias in exact:
            return exact[alias][0]
    for alias in normalized_aliases:
        for token, row, col in entries:
            if alias in token:
                return row, col
    return None


def resolve_value_cell(
    snapshot: SheetSnapshot,
    text_index: tuple[dict[str, list[tuple[int, int]]], list[tuple[str, int, int]]],
    aliases: list[str],
    fallback: tuple[int, int] | None = None,
) -> tuple[int, int] | None:
    label_cell = find_label_cell(text_index, aliases)
    if not label_cell:
        return fallback
    label_row, label_col = label_cell
    right = (label_row, label_col + 1)
    below = (label_row + 1, label_col)
    right_value = snapshot.get(*right)
    below_value = snapshot.get(*below)
    if right_value is not None and right_value != "":
        return right
    if below_value is not None and below_value != "":
        return below
    return right


def safe_get_cell_value(sheet: xw.Sheet, coord: tuple[int, int] | None) -> Any:
    if coord is None:
        return None
    row, col = coord
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def safe_set_value(sheet: xw.Sheet, coord: tuple[int, int] | None, value: Any) -> bool:
    if coord is None:
        return False
    row, col = coord
    try:
        sheet.range((row, col)).value = value
        return True
    except Exception:
        return False


def safe_set_formula2(sheet: xw.Sheet, coord: tuple[int, int] | None, formula_r1c1: str) -> bool:
    if coord is None:
        return False
    row, col = coord
    rng = sheet.range((row, col))
    try:
        rng.api.Formula2R1C1 = formula_r1c1
        return True
    except Exception:
        pass
    try:
        rng.formula2 = formula_r1c1
        return True
    except Exception:
        try:
            rng.formula = formula_r1c1
            return True
        except Exception:
            return False


def safe_clear_cell(sheet: xw.Sheet, coord: tuple[int, int] | None) -> None:
    if coord is None:
        return
    row, col = coord
    try:
        sheet.range((row, col)).value = None
    except Exception:
        pass


def parse_file_labels(file_path: Path) -> dict[str, str]:
    stem = file_path.stem

    ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
    ticker = ticker_match.group(1).upper() if ticker_match else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*([0-9]{4})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "UNKNOWN"
    model_date = ""
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.capitalize()
        month = month_raw.capitalize()
        year = int(year_raw)
        month_num = MONTHS.get(month.lower())
        day = PHASE_DAY.get(phase.lower())
        if month_num and day:
            model_period = f"{phase}{month}_{year}"
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period != "UNKNOWN" else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def safe_close_workbook(wb: xw.Book) -> None:
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
        pass


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.resolve().name
    base = out_dir / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base
    idx = 1
    while True:
        candidate = out_dir / f"{folder_name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def extract_empirical_candidates(wb: xw.Book, file_labels: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  - skipped empirical extraction in {source_file}: missing 'Empirical Model'")
        return []

    snapshot = load_snapshot(sheet)
    text_index = build_text_index(snapshot)
    anchor = find_anchor_max(snapshot, text_index[1])
    if anchor is None:
        print(f"  - skipped empirical extraction in {source_file}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor

    num_quarters_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["num quarters used", "num qtrs used", "number of quarters used", "n quarters"],
        fallback=(anchor_row - 1, anchor_col - 2),
    )
    avg_pen_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["avg penetration pct", "average penetration pct", "avg penetration"],
        fallback=(anchor_row - 2, anchor_col - 2),
    )
    estimated_total_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["estimated total sold", "forecast value", "total forecast", "forecast"],
        fallback=(anchor_row - 3, anchor_col + 1),
    )
    reported_sales_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["reported sales", "actual value", "actual sales"],
        fallback=(anchor_row - 2, anchor_col + 1),
    )
    forecast_max_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["max"],
        fallback=(anchor_row, anchor_col + 1),
    )
    forecast_min_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["min"],
        fallback=(anchor_row + 1, anchor_col + 1),
    )
    quarterly_sales_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["quarterly sales"],
        fallback=(anchor_row - 4, anchor_col + 1),
    )
    growth_rate_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["growth rate pct", "growth rate"],
        fallback=(anchor_row - 5, anchor_col + 1),
    )
    sales_captured_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["sales captured in db pct", "sales captured"],
        fallback=(anchor_row - 6, anchor_col + 1),
    )
    last_quarter_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["last quarter used", "latest quarter used"],
        fallback=(anchor_row - 1, anchor_col - 1),
    )

    # Anchor-based history offsets for avg penetration formula.
    penetration_history_col = anchor_col - 11
    history_end_row = max(snapshot.start_row, anchor_row - 1)
    temp_formula_cell = (anchor_row, anchor_col + 20)

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        history_start_row = max(snapshot.start_row, history_end_row - n_quarters + 1)
        avg_formula = (
            f"=AVERAGE(R{history_start_row}C{penetration_history_col}:"
            f"R{history_end_row}C{penetration_history_col})"
        )

        changed = False
        if safe_set_value(sheet, num_quarters_cell, n_quarters):
            changed = True

        formula_target = avg_pen_cell if avg_pen_cell is not None else temp_formula_cell
        if safe_set_formula2(sheet, formula_target, avg_formula):
            changed = True

        if changed:
            wb.app.calculate()

        avg_pen = safe_get_cell_value(sheet, avg_pen_cell or temp_formula_cell)
        forecast_max = safe_get_cell_value(sheet, forecast_max_cell)
        forecast_min = safe_get_cell_value(sheet, forecast_min_cell)
        forecast_value = safe_get_cell_value(sheet, estimated_total_cell)
        actual_value = safe_get_cell_value(sheet, reported_sales_cell)
        quarterly_sales = safe_get_cell_value(sheet, quarterly_sales_cell)
        reported_sales = safe_get_cell_value(sheet, reported_sales_cell)
        growth_rate = safe_get_cell_value(sheet, growth_rate_cell)
        sales_captured = safe_get_cell_value(sheet, sales_captured_cell)
        last_quarter_used = safe_get_cell_value(sheet, last_quarter_cell)

        rows.append(
            {
                "model": file_labels["model"],
                "ticker": file_labels["ticker"],
                "model_period": file_labels["model_period"],
                "model_date": file_labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_delta(forecast_max, forecast_min),
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    safe_clear_cell(sheet, temp_formula_cell)
    return rows


def extract_regression_candidates(wb: xw.Book, file_labels: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  - skipped regression extraction in {source_file}: missing 'Regression Model'")
        return []

    snapshot = load_snapshot(sheet)
    text_index = build_text_index(snapshot)
    anchor = find_anchor_max(snapshot, text_index[1])
    if anchor is None:
        print(f"  - skipped regression extraction in {source_file}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["num quarters used", "num qtrs used", "number of quarters used", "n quarters"],
        fallback=(anchor_row - 1, anchor_col - 2),
    )
    forecast_total_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["tot fcst w/o sa", "tot fcst without sa", "forecast total without sa"],
        fallback=(anchor_row - 2, anchor_col + 1),
    )
    actual_value_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["actual value", "reported sales", "actual sales"],
        fallback=None,
    )
    forecast_max_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["max"],
        fallback=(anchor_row, anchor_col + 1),
    )
    forecast_min_cell = resolve_value_cell(
        snapshot,
        text_index,
        aliases=["min"],
        fallback=(anchor_row + 1, anchor_col + 1),
    )

    intercept_formula_cell = (anchor_row, anchor_col + 20)
    slope_formula_cell = (anchor_row, anchor_col + 21)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters in range(1, N_QUARTERS + 1):
        changed = False
        if safe_set_value(sheet, num_quarters_cell, n_quarters):
            changed = True

        y_end_row = max(snapshot.start_row, anchor_row - 1)
        y_start_row = max(snapshot.start_row, y_end_row - n_quarters + 1)

        intercept_formula = (
            f"=INTERCEPT(R{y_start_row}C{y_col}:R{y_end_row}C{y_col},"
            f"R{y_start_row}C{x_col}:R{y_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{y_start_row}C{y_col}:R{y_end_row}C{y_col},"
            f"R{y_start_row}C{x_col}:R{y_end_row}C{x_col})"
        )

        if safe_set_formula2(sheet, intercept_formula_cell, intercept_formula):
            changed = True
        if safe_set_formula2(sheet, slope_formula_cell, slope_formula):
            changed = True

        if changed:
            wb.app.calculate()

        intercept = safe_get_cell_value(sheet, intercept_formula_cell)
        slope = safe_get_cell_value(sheet, slope_formula_cell)
        forecast_value = safe_get_cell_value(sheet, forecast_total_cell)
        actual_value = safe_get_cell_value(sheet, actual_value_cell)
        forecast_max = safe_get_cell_value(sheet, forecast_max_cell)
        forecast_min = safe_get_cell_value(sheet, forecast_min_cell)

        signature = (
            round(coerce_float(intercept) or 0.0, 10),
            round(coerce_float(slope) or 0.0, 10),
            round(coerce_float(forecast_value) or 0.0, 10),
            round(coerce_float(forecast_max) or 0.0, 10),
            round(coerce_float(forecast_min) or 0.0, 10),
        )
        if previous_signature is not None and signature == previous_signature:
            continue

        rows.append(
            {
                "model": file_labels["model"],
                "ticker": file_labels["ticker"],
                "model_period": file_labels["model_period"],
                "model_date": file_labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_delta(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )
        previous_signature = signature

    safe_clear_cell(sheet, intercept_formula_cell)
    safe_clear_cell(sheet, slope_formula_cell)
    return rows


def write_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(48, max(12, max_len + 2))


def write_output_workbook(output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    wb = Workbook()

    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        print(f"Input directory does not exist: {in_dir}")
        return

    files = sorted(in_dir.iterdir())
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    try:
        app = xw.App(visible=False, add_book=False)
    except Exception as exc:
        print(f"Unable to start hidden Excel app: {exc}")
        return
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files:
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not an .xlsx file)")
                continue

            print(f"processed: {file_path.name}")
            file_labels = parse_file_labels(file_path)

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(wb, file_labels, file_path.name))
                regression_rows.extend(extract_regression_candidates(wb, file_labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"  - skipped file during processing: {file_path.name} ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    output_path = next_output_path(in_dir, out_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
