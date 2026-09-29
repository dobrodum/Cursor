from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------- User-configurable paths ----------
input_dir = Path("input")
output_dir = Path("output")


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
DAY_BY_PERIOD_PREFIX = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetSnapshot:
    top_row: int
    left_col: int
    values: list[list[Any]]

    @property
    def bottom_row(self) -> int:
        return self.top_row + len(self.values) - 1

    @property
    def right_col(self) -> int:
        if not self.values or not self.values[0]:
            return self.left_col
        return self.left_col + len(self.values[0]) - 1


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def parse_file_labels(file_name: str) -> FileLabels:
    base = Path(file_name).stem
    file_pattern = re.compile(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3,9}\d{4})",
        re.IGNORECASE,
    )
    match = file_pattern.search(base)
    if not match:
        raise ValueError("filename does not match expected ticker/period pattern")

    ticker = match.group("ticker").upper()
    raw_period = match.group("period")
    period_pattern = re.compile(
        r"^(?P<prefix>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})$",
        re.IGNORECASE,
    )
    period_match = period_pattern.match(raw_period)
    if not period_match:
        raise ValueError("cannot parse model period token")

    prefix = period_match.group("prefix").title()
    month_text = period_match.group("month")
    year = int(period_match.group("year"))
    month_abbrev = month_text[:3].title()
    try:
        month_num = datetime.strptime(month_abbrev, "%b").month
    except ValueError as exc:
        raise ValueError(f"unknown month token '{month_text}'") from exc

    day = DAY_BY_PERIOD_PREFIX[prefix.lower()]
    model_period = f"{prefix}{month_abbrev}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_output_path(src_input_dir: Path, dst_output_dir: Path) -> Path:
    dst_output_dir.mkdir(parents=True, exist_ok=True)
    folder_name = src_input_dir.name or "input"
    base_name = f"{folder_name}_PARAM"
    output_path = dst_output_dir / f"{base_name}.xlsx"
    if not output_path.exists():
        return output_path

    idx = 1
    while True:
        candidate = dst_output_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def make_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    raw_values = used.value

    if raw_values is None:
        values: list[list[Any]] = []
    elif isinstance(raw_values, list):
        if raw_values and isinstance(raw_values[0], list):
            values = raw_values
        else:
            values = [raw_values]
    else:
        values = [[raw_values]]

    return SheetSnapshot(top_row=used.row, left_col=used.column, values=values)


def find_text_cell(snapshot: SheetSnapshot, needle: str) -> tuple[int, int] | None:
    normalized_target = normalize_text(needle)
    for r_offset, row in enumerate(snapshot.values):
        for c_offset, cell_value in enumerate(row):
            if normalize_text(cell_value) == normalized_target:
                return snapshot.top_row + r_offset, snapshot.left_col + c_offset
    return None


def find_header_col(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    aliases: Iterable[str],
    row_window: int = 2,
) -> int | None:
    alias_tokens = [normalize_text(alias) for alias in aliases]
    best_match: tuple[int, int] | None = None

    for r_offset, row in enumerate(snapshot.values):
        row_num = snapshot.top_row + r_offset
        if abs(row_num - anchor_row) > row_window:
            continue
        for c_offset, cell_value in enumerate(row):
            text = normalize_text(cell_value)
            if not text:
                continue
            if any(alias == text or alias in text for alias in alias_tokens):
                col_num = snapshot.left_col + c_offset
                score = abs(row_num - anchor_row) * 100 + abs(col_num - anchor_col)
                if best_match is None or score < best_match[0]:
                    best_match = (score, col_num)

    return None if best_match is None else best_match[1]


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def formula2_r1c1(target: xw.Range, r1c1_formula: str) -> None:
    try:
        target.formula2 = r1c1_formula
    except Exception:
        try:
            target.api.Formula2R1C1 = r1c1_formula
        except Exception:
            target.api.FormulaR1C1 = r1c1_formula


def numeric_rows_for_col(sheet: xw.Sheet, col: int, row_start: int, row_end: int) -> list[tuple[int, float]]:
    rows: list[tuple[int, float]] = []
    for row in range(max(1, row_start), row_end + 1):
        val = to_float(sheet.cells(row, col).value)
        if val is not None:
            rows.append((row, val))
    return rows


def close_source_workbook(wb: xw.Book) -> None:
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
        wb.api.Close(False)
    except Exception:
        # Last fallback: close without arguments (still never save explicitly).
        wb.close()


def build_empirical_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    snapshot = make_snapshot(sheet)
    anchor = find_text_cell(snapshot, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    cols = {
        "num_quarters_used": find_header_col(snapshot, anchor_row, anchor_col, ["num_quarters_used", "num quarters", "quarters used", "# quarters"]) or (anchor_col - 5),
        "last_quarter_used": find_header_col(snapshot, anchor_row, anchor_col, ["last_quarter_used", "last quarter used", "last quarter"]) or (anchor_col - 4),
        "forecast_value": find_header_col(snapshot, anchor_row, anchor_col, ["estimated total sold", "forecast", "forecast value", "tot fcst"]) or (anchor_col - 2),
        "actual_value": find_header_col(snapshot, anchor_row, anchor_col, ["reported sales", "actual sales", "actual"]) or (anchor_col - 1),
        "forecast_max": find_header_col(snapshot, anchor_row, anchor_col, ["max", "forecast max"]) or anchor_col,
        "forecast_min": find_header_col(snapshot, anchor_row, anchor_col, ["min", "forecast min"]) or (anchor_col + 1),
        "avg_penetration_pct": find_header_col(snapshot, anchor_row, anchor_col, ["avg penetration", "average penetration", "avg_penetration_pct"]) or (anchor_col - 3),
        "quarterly_sales": find_header_col(snapshot, anchor_row, anchor_col, ["quarterly sales", "sales per quarter"]) or (anchor_col - 6),
        "reported_sales": find_header_col(snapshot, anchor_row, anchor_col, ["reported sales", "actual sales"]) or (anchor_col - 1),
        "growth_rate_pct": find_header_col(snapshot, anchor_row, anchor_col, ["growth rate", "growth %", "growth_rate_pct"]) or (anchor_col - 7),
        "sales_captured_in_db_pct": find_header_col(snapshot, anchor_row, anchor_col, ["sales captured in db", "sales captured", "captured in db %"]) or (anchor_col - 8),
    }

    penetration_history_col = cols["sales_captured_in_db_pct"]
    penetration_history = numeric_rows_for_col(sheet, penetration_history_col, max(1, anchor_row - 400), anchor_row - 1)

    write_formula = len(penetration_history) > 0
    if write_formula:
        for idx in range(N_QUARTERS):
            target_row = anchor_row + 1 + idx
            n = idx + 1
            sheet.cells(target_row, cols["num_quarters_used"]).value = n

            if len(penetration_history) >= n:
                start_row = penetration_history[-n][0]
            else:
                start_row = penetration_history[0][0]
            end_row = penetration_history[-1][0]

            avg_cell = sheet.cells(target_row, cols["avg_penetration_pct"])
            formula2_r1c1(
                avg_cell,
                f'=IFERROR(AVERAGE(R{start_row}C{penetration_history_col}:R{end_row}C{penetration_history_col}),"")',
            )

        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = anchor_row + 1 + idx

        num_quarters = sheet.cells(row_num, cols["num_quarters_used"]).value
        last_quarter = sheet.cells(row_num, cols["last_quarter_used"]).value
        forecast_value = sheet.cells(row_num, cols["forecast_value"]).value
        actual_value = sheet.cells(row_num, cols["actual_value"]).value
        forecast_max = sheet.cells(row_num, cols["forecast_max"]).value
        forecast_min = sheet.cells(row_num, cols["forecast_min"]).value
        avg_pen = sheet.cells(row_num, cols["avg_penetration_pct"]).value
        quarterly_sales = sheet.cells(row_num, cols["quarterly_sales"]).value
        reported_sales = sheet.cells(row_num, cols["reported_sales"]).value
        growth_rate = sheet.cells(row_num, cols["growth_rate_pct"]).value
        sales_captured = sheet.cells(row_num, cols["sales_captured_in_db_pct"]).value

        meaningful_values = [forecast_value, forecast_max, forecast_min, avg_pen, num_quarters]
        if all(v in (None, "") for v in meaningful_values):
            continue

        max_val = to_float(forecast_max)
        min_val = to_float(forecast_min)
        range_width = (max_val - min_val) if (max_val is not None and min_val is not None) else None

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def build_regression_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    snapshot = make_snapshot(sheet)
    anchor = find_text_cell(snapshot, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    cols = {
        "num_quarters_used": find_header_col(snapshot, anchor_row, anchor_col, ["num_quarters_used", "num quarters", "quarters used", "# quarters"]) or (anchor_col - 4),
        "forecast_value": find_header_col(snapshot, anchor_row, anchor_col, ["tot fcst w/o sa", "forecast total without sa", "tot fcst wo sa", "forecast"]) or (anchor_col - 2),
        "actual_value": find_header_col(snapshot, anchor_row, anchor_col, ["actual value", "actual sales", "reported sales"]),
        "forecast_max": find_header_col(snapshot, anchor_row, anchor_col, ["max", "forecast max"]) or anchor_col,
        "forecast_min": find_header_col(snapshot, anchor_row, anchor_col, ["min", "forecast min"]) or (anchor_col + 1),
        "intercept": find_header_col(snapshot, anchor_row, anchor_col, ["intercept"]) or (anchor_col + 2),
        "slope": find_header_col(snapshot, anchor_row, anchor_col, ["slope"]) or (anchor_col + 3),
    }

    xy_rows: list[tuple[int, float, float]] = []
    for row in range(max(1, anchor_row - 400), anchor_row):
        x_val = to_float(sheet.cells(row, x_col).value)
        y_val = to_float(sheet.cells(row, y_col).value)
        if x_val is None or y_val is None:
            continue
        xy_rows.append((row, x_val, y_val))

    if not xy_rows:
        return []

    last_x = xy_rows[-1][1]
    prev_signature: tuple[Any, ...] | None = None
    rows: list[dict[str, Any]] = []

    for idx in range(N_QUARTERS):
        row_num = anchor_row + 1 + idx
        n = idx + 1
        sample = xy_rows[-n:] if len(xy_rows) >= n else xy_rows
        sample_start = sample[0][0]
        sample_end = sample[-1][0]

        sheet.cells(row_num, cols["num_quarters_used"]).value = n
        formula2_r1c1(
            sheet.cells(row_num, cols["intercept"]),
            f'=IFERROR(INTERCEPT(R{sample_start}C{y_col}:R{sample_end}C{y_col},R{sample_start}C{x_col}:R{sample_end}C{x_col}),"")',
        )
        formula2_r1c1(
            sheet.cells(row_num, cols["slope"]),
            f'=IFERROR(SLOPE(R{sample_start}C{y_col}:R{sample_end}C{y_col},R{sample_start}C{x_col}:R{sample_end}C{x_col}),"")',
        )

    wb.app.calculate()

    for idx in range(N_QUARTERS):
        row_num = anchor_row + 1 + idx
        num_quarters = sheet.cells(row_num, cols["num_quarters_used"]).value
        forecast_value = sheet.cells(row_num, cols["forecast_value"]).value
        actual_value = sheet.cells(row_num, cols["actual_value"]).value if cols["actual_value"] else None
        forecast_max = sheet.cells(row_num, cols["forecast_max"]).value
        forecast_min = sheet.cells(row_num, cols["forecast_min"]).value
        intercept = sheet.cells(row_num, cols["intercept"]).value
        slope = sheet.cells(row_num, cols["slope"]).value

        if forecast_value in (None, ""):
            i_val = to_float(intercept)
            s_val = to_float(slope)
            if i_val is not None and s_val is not None:
                forecast_value = i_val + (s_val * last_x)

        max_val = to_float(forecast_max)
        min_val = to_float(forecast_min)
        range_width = (max_val - min_val) if (max_val is not None and min_val is not None) else None

        signature = (
            round(to_float(forecast_value) or 0.0, 10),
            round(to_float(forecast_max) or 0.0, 10),
            round(to_float(forecast_min) or 0.0, 10),
            round(to_float(intercept) or 0.0, 10),
            round(to_float(slope) or 0.0, 10),
        )
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        meaningful_values = [num_quarters, forecast_value, forecast_max, forecast_min, intercept, slope]
        if all(v in (None, "") for v in meaningful_values):
            continue

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
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def format_sheet(ws, headers: list[str]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_emp = wb.create_sheet("empirical_candidates")
    ws_emp.append(EMPIRICAL_COLUMNS)
    for header_cell in ws_emp[1]:
        header_cell.font = Font(bold=True)
    for row in empirical_rows:
        ws_emp.append([row.get(col) for col in EMPIRICAL_COLUMNS])
    format_sheet(ws_emp, EMPIRICAL_COLUMNS)

    ws_reg = wb.create_sheet("regression_candidates")
    ws_reg.append(REGRESSION_COLUMNS)
    for header_cell in ws_reg[1]:
        header_cell.font = Font(bold=True)
    for row in regression_rows:
        ws_reg.append([row.get(col) for col in REGRESSION_COLUMNS])
    format_sheet(ws_reg, REGRESSION_COLUMNS)

    wb.save(output_path)


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"skipped input_dir: '{input_dir}' does not exist or is not a directory")
        return

    output_path = get_output_path(input_dir, output_dir)
    entries = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for item in entries:
            if not item.is_file():
                print(f"skipped file: {item.name} (not a file)")
                continue
            if item.name.startswith("~"):
                print(f"skipped file: {item.name} (temporary file)")
                continue
            if item.suffix.lower() != ".xlsx":
                print(f"skipped file: {item.name} (not .xlsx)")
                continue

            try:
                labels = parse_file_labels(item.name)
            except ValueError as exc:
                print(f"skipped file: {item.name} ({exc})")
                continue

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(item), update_links=False)
                print(f"processed file: {item.name}")
                empirical_rows.extend(build_empirical_rows(wb, labels, item.name))
                regression_rows.extend(build_regression_rows(wb, labels, item.name))
                files_processed += 1
            except Exception as exc:
                print(f"skipped file: {item.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)

    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"output path: {output_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
