#!/usr/bin/env python3
"""Extract empirical and regression model candidates from source workbooks."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
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

PERIOD_RE = re.compile(
    r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})",
    re.IGNORECASE,
)
MONTH_NUM = {
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
PHASE_DAY = {"Early": 5, "Mid": 15, "Late": 25}
PARAM_OUTPUT_RE = re.compile(r"_PARAM(?:\.\d+)?\.xlsx$", re.IGNORECASE)


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class LabelCell:
    text: str
    row: int
    col: int


def normalize_text(value: Any) -> str:
    """Lowercase, trim, and collapse spaces."""
    if value is None:
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip().lower()
    return text


def to_number(value: Any) -> float | None:
    """Parse a numeric value from Excel cell content."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(int(value))
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        pct = raw.endswith("%")
        cleaned = raw.replace(",", "").replace("$", "").replace("%", "")
        try:
            number = float(cleaned)
        except ValueError:
            return None
        if pct:
            return number / 100.0
        return number
    return None


def number_or_blank(value: Any) -> Any:
    """Return parsed numeric value or blank."""
    parsed = to_number(value)
    return parsed if parsed is not None else None


def choose_output_path(src_dir: Path, dst_dir: Path) -> Path:
    """Build unique output path: folder_PARAM.xlsx, then .1/.2..."""
    dst_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{src_dir.name}_PARAM"
    candidate = dst_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = dst_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def parse_file_labels(file_name: str) -> FileLabels:
    """Parse ticker/model period/date from file name."""
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    if not ticker:
        raise ValueError("ticker could not be parsed")

    period_match = PERIOD_RE.search(stem)
    if not period_match:
        raise ValueError("model period token (Early/Mid/Late + month + year) missing")

    phase = period_match.group(1).title()
    month_token = period_match.group(2).title()
    year = int(period_match.group(3))
    if month_token not in MONTH_NUM:
        raise ValueError(f"unrecognized month token: {month_token}")

    model_period = f"{phase}{month_token}_{year}"
    model_date = date(year, MONTH_NUM[month_token], PHASE_DAY[phase]).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def collect_source_files(src_dir: Path) -> list[Path]:
    """Collect valid source .xlsx files and log skipped entries."""
    files: list[Path] = []
    if not src_dir.exists():
        print(f"Skipped folder: input_dir does not exist -> {src_dir}")
        return files

    for path in sorted(src_dir.iterdir()):
        if not path.is_file():
            print(f"Skipped {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not .xlsx")
            continue
        if PARAM_OUTPUT_RE.search(path.name):
            print(f"Skipped {path.name}: prior PARAM output workbook")
            continue
        files.append(path)
    return files


def scan_sheet(sheet: xw.Sheet) -> tuple[tuple[int, int] | None, list[LabelCell]]:
    """Scan used range once to find 'max' anchor and labels."""
    used = sheet.used_range
    values = used.value
    if values is None:
        return None, []

    rows: list[list[Any]]
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            rows = values
        else:
            rows = [values]
    else:
        rows = [[values]]

    labels: list[LabelCell] = []
    anchor: tuple[int, int] | None = None
    start_row = used.row
    start_col = used.column

    for row_idx, row_values in enumerate(rows):
        current_row = row_values if isinstance(row_values, list) else [row_values]
        for col_idx, value in enumerate(current_row):
            text = normalize_text(value)
            if not text:
                continue
            row = start_row + row_idx
            col = start_col + col_idx
            labels.append(LabelCell(text=text, row=row, col=col))
            if anchor is None and text == "max":
                anchor = (row, col)
    return anchor, labels


def find_label(labels: Iterable[LabelCell], phrase_sets: list[tuple[str, ...]]) -> LabelCell | None:
    """Return first label that contains all terms in a candidate phrase set."""
    for phrase_set in phrase_sets:
        for label in labels:
            if all(term in label.text for term in phrase_set):
                return label
    return None


def cell_if_valid(sheet: xw.Sheet, row: int, col: int) -> xw.Range | None:
    """Return cell if row/col are valid Excel coordinates."""
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col)


def get_value_from_anchor(sheet: xw.Sheet, anchor: tuple[int, int], row_off: int, col_off: int) -> Any:
    """Read value using anchor offsets."""
    row, col = anchor
    cell = cell_if_valid(sheet, row + row_off, col + col_off)
    if cell is None:
        return None
    return cell.value


def get_value_near_label(
    sheet: xw.Sheet,
    labels: list[LabelCell],
    phrase_sets: list[tuple[str, ...]],
    row_off: int = 0,
    col_off: int = 1,
) -> Any:
    """Read value from offset near the first matching label."""
    hit = find_label(labels, phrase_sets)
    if hit is None:
        return None
    cell = cell_if_valid(sheet, hit.row + row_off, hit.col + col_off)
    if cell is None:
        return None
    return cell.value


def get_row_from_label(
    labels: list[LabelCell],
    phrase_sets: list[tuple[str, ...]],
    default_row: int,
) -> int:
    """Resolve a row based on label text; fallback to default row."""
    hit = find_label(labels, phrase_sets)
    return hit.row if hit else default_row


def get_cell_from_label(
    sheet: xw.Sheet,
    labels: list[LabelCell],
    phrase_sets: list[tuple[str, ...]],
    row_off: int = 0,
    col_off: int = 1,
) -> xw.Range | None:
    """Resolve an xlwings cell from a label and offset."""
    hit = find_label(labels, phrase_sets)
    if hit is None:
        return None
    return cell_if_valid(sheet, hit.row + row_off, hit.col + col_off)


def latest_numeric_in_row(sheet: xw.Sheet, row: int, start_col: int, end_col: int) -> float | None:
    """Get latest (rightmost) numeric value in a row segment."""
    if start_col < 1 or end_col < start_col:
        return None
    values = sheet.range((row, start_col), (row, end_col)).value
    if not isinstance(values, list):
        values = [values]
    for value in reversed(values):
        parsed = to_number(value)
        if parsed is not None:
            return parsed
    return None


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    """Write formula using Formula2 first, fallback to Formula."""
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def make_trailing_average_formula(row: int, start_col: int, end_col: int, n_quarters: int) -> str:
    """R1C1 formula that averages the trailing n populated values."""
    return (
        f"=LET(rng,R{row}C{start_col}:R{row}C{end_col},"
        f"n,{n_quarters},"
        "cnt,COUNTA(rng),"
        "AVERAGE(INDEX(rng,1,MAX(1,cnt-n+1)):INDEX(rng,1,cnt)))"
    )


def close_source_workbook(wb: xw.Book) -> None:
    """Close workbook without saving with compatibility fallbacks."""
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
        try:
            wb.close()
        except Exception:
            pass


def process_empirical_sheet(
    wb: xw.Book,
    labels_meta: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    """Extract empirical candidates from one workbook."""
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction ({source_file}): 'Empirical Model' sheet not found")
        return []

    anchor, labels = scan_sheet(sheet)
    if anchor is None:
        print(f"Skipped empirical extraction ({source_file}): 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    history_start_col = max(1, anchor_col - 11)
    history_end_col = max(history_start_col, anchor_col - 2)
    n_quarters_max = 10

    num_quarters_cell = get_cell_from_label(
        sheet,
        labels,
        phrase_sets=[("num", "quarters", "used"), ("quarters", "used")],
    )
    avg_pen_cell = get_cell_from_label(
        sheet,
        labels,
        phrase_sets=[("avg", "penetration"), ("average", "penetration")],
    )
    if avg_pen_cell is None:
        avg_pen_cell = cell_if_valid(sheet, anchor_row, anchor_col + 3)
        if avg_pen_cell is None:
            return []

    penetration_row = get_row_from_label(
        labels,
        phrase_sets=[
            ("sales", "captured", "db"),
            ("penetration",),
        ],
        default_row=max(1, anchor_row - 4),
    )
    quarterly_sales_row = get_row_from_label(
        labels,
        phrase_sets=[("quarterly", "sales"), ("sales", "in", "db")],
        default_row=max(1, anchor_row - 3),
    )
    reported_sales_row = get_row_from_label(
        labels,
        phrase_sets=[("reported", "sales"), ("actual", "sales")],
        default_row=max(1, anchor_row - 2),
    )
    growth_rate_row = get_row_from_label(
        labels,
        phrase_sets=[("growth", "rate")],
        default_row=max(1, anchor_row - 1),
    )
    quarter_label_row = max(1, penetration_row - 1)

    rows: list[dict[str, Any]] = []

    for n_quarters in range(1, n_quarters_max + 1):
        if num_quarters_cell is not None:
            num_quarters_cell.value = n_quarters

        avg_formula = make_trailing_average_formula(
            row=penetration_row,
            start_col=history_start_col,
            end_col=history_end_col,
            n_quarters=n_quarters,
        )
        set_formula2(avg_pen_cell, avg_formula)
        wb.app.calculate()

        avg_penetration_pct = to_number(avg_pen_cell.value)
        sales_captured_in_db_pct = latest_numeric_in_row(
            sheet, penetration_row, history_start_col, history_end_col
        )
        quarterly_sales = latest_numeric_in_row(
            sheet, quarterly_sales_row, history_start_col, history_end_col
        )
        reported_sales = latest_numeric_in_row(
            sheet, reported_sales_row, history_start_col, history_end_col
        )
        growth_rate_pct = latest_numeric_in_row(
            sheet, growth_rate_row, history_start_col, history_end_col
        )

        estimated_total_sold = to_number(
            get_value_near_label(
                sheet,
                labels,
                phrase_sets=[("estimated", "total", "sold"), ("total", "sold")],
            )
        )
        if estimated_total_sold is None and quarterly_sales is not None and avg_penetration_pct:
            penetration = avg_penetration_pct
            if penetration > 1:
                penetration = penetration / 100.0
            if penetration:
                estimated_total_sold = quarterly_sales / penetration

        forecast_max = to_number(get_value_from_anchor(sheet, anchor, 0, 1))
        if forecast_max is None:
            forecast_max = to_number(
                get_value_near_label(sheet, labels, phrase_sets=[("max",)], row_off=0, col_off=1)
            )

        forecast_min = to_number(get_value_from_anchor(sheet, anchor, 1, 1))
        if forecast_min is None:
            forecast_min = to_number(
                get_value_near_label(sheet, labels, phrase_sets=[("min",)], row_off=0, col_off=1)
            )

        last_quarter_used = sheet.cells(quarter_label_row, history_end_col).value
        if last_quarter_used in ("", None):
            last_quarter_used = get_value_near_label(
                sheet,
                labels,
                phrase_sets=[("last", "quarter", "used")],
            )

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row = {
            "model": labels_meta.model,
            "ticker": labels_meta.ticker,
            "model_period": labels_meta.model_period,
            "model_date": labels_meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": n_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": estimated_total_sold,
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


def regression_row_signature(row: dict[str, Any]) -> tuple[Any, ...]:
    """Build duplicate-detection signature for regression rows."""
    def rounded(value: Any) -> Any:
        if isinstance(value, (int, float)):
            return round(float(value), 8)
        return value

    return (
        rounded(row.get("num_quarters_used")),
        rounded(row.get("forecast_value")),
        rounded(row.get("forecast_max")),
        rounded(row.get("forecast_min")),
        rounded(row.get("intercept")),
        rounded(row.get("slope")),
    )


def process_regression_sheet(
    wb: xw.Book,
    labels_meta: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    """Extract regression candidates from one workbook."""
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction ({source_file}): 'Regression Model' sheet not found")
        return []

    anchor, labels = scan_sheet(sheet)
    if anchor is None:
        print(f"Skipped regression extraction ({source_file}): 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)
    history_end_col = max(x_col, anchor_col - 2)
    n_quarters_max = min(10, history_end_col - x_col + 1)

    num_quarters_cell = get_cell_from_label(
        sheet,
        labels,
        phrase_sets=[("num", "quarters", "used"), ("quarters", "used")],
    )
    intercept_cell = get_cell_from_label(
        sheet,
        labels,
        phrase_sets=[("intercept",)],
    )
    slope_cell = get_cell_from_label(
        sheet,
        labels,
        phrase_sets=[("slope",)],
    )
    if intercept_cell is None:
        intercept_cell = cell_if_valid(sheet, anchor_row, anchor_col + 3)
    if slope_cell is None:
        slope_cell = cell_if_valid(sheet, anchor_row, anchor_col + 4)
    if intercept_cell is None or slope_cell is None:
        return []

    y_row = get_row_from_label(
        labels,
        phrase_sets=[("tot", "fcst", "w/o", "sa"), ("tot", "fcst")],
        default_row=max(1, anchor_row - 2),
    )
    x_row = get_row_from_label(
        labels,
        phrase_sets=[("quarter",), ("qtr",)],
        default_row=max(1, y_row - 1),
    )

    rows: list[dict[str, Any]] = []
    prior_signature: tuple[Any, ...] | None = None

    for n_quarters in range(1, n_quarters_max + 1):
        if num_quarters_cell is not None:
            num_quarters_cell.value = n_quarters

        start_col = history_end_col - n_quarters + 1
        intercept_formula = (
            f"=INTERCEPT(R{y_row}C{start_col}:R{y_row}C{history_end_col},"
            f"R{x_row}C{start_col}:R{x_row}C{history_end_col})"
        )
        slope_formula = (
            f"=SLOPE(R{y_row}C{start_col}:R{y_row}C{history_end_col},"
            f"R{x_row}C{start_col}:R{x_row}C{history_end_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_number(intercept_cell.value)
        slope = to_number(slope_cell.value)

        forecast_total_without_sa = latest_numeric_in_row(
            sheet, y_row, start_col, history_end_col
        )
        if forecast_total_without_sa is None:
            forecast_total_without_sa = to_number(
                get_value_near_label(
                    sheet,
                    labels,
                    phrase_sets=[("tot", "fcst", "w/o", "sa"), ("tot", "fcst")],
                )
            )

        forecast_max = to_number(get_value_from_anchor(sheet, anchor, 0, 1))
        if forecast_max is None:
            forecast_max = to_number(
                get_value_near_label(sheet, labels, phrase_sets=[("max",)], row_off=0, col_off=1)
            )

        forecast_min = to_number(get_value_from_anchor(sheet, anchor, 1, 1))
        if forecast_min is None:
            forecast_min = to_number(
                get_value_near_label(sheet, labels, phrase_sets=[("min",)], row_off=0, col_off=1)
            )

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row = {
            "model": labels_meta.model,
            "ticker": labels_meta.ticker,
            "model_period": labels_meta.model_period,
            "model_date": labels_meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_total_without_sa,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        signature = regression_row_signature(row)
        if prior_signature is not None and signature == prior_signature:
            continue
        rows.append(row)
        prior_signature = signature

    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    """Write one output sheet with formatting."""
    ws = wb.create_sheet(title=sheet_name)
    ws.append(columns)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    max_lengths = [len(column) for column in columns]
    for row in rows:
        ordered = [row.get(column) for column in columns]
        ws.append(ordered)
        for idx, value in enumerate(ordered):
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_lengths[idx]:
                max_lengths[idx] = value_len

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, width in enumerate(max_lengths, start=1):
        column_letter = get_column_letter(idx)
        ws.column_dimensions[column_letter].width = min(max(12, width + 2), 45)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    """Create output workbook with both required sheets."""
    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def configure_app(app: xw.App) -> None:
    """Set app to low-noise, faster settings."""
    app.visible = False
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass


def main() -> None:
    src_dir = Path(input_dir).expanduser().resolve()
    dst_dir = Path(output_dir).expanduser().resolve()

    source_files = collect_source_files(src_dir)
    output_path = choose_output_path(src_dir, dst_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    if not source_files:
        write_output_workbook(output_path, empirical_rows, regression_rows)
        print(f"Output path: {output_path}")
        print("Number of files processed: 0")
        print("Number of empirical rows: 0")
        print("Number of regression rows: 0")
        return

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        configure_app(app)

        for file_path in source_files:
            try:
                labels_meta = parse_file_labels(file_path.name)
            except ValueError as exc:
                print(f"Skipped {file_path.name}: {exc}")
                continue

            wb: xw.Book | None = None
            try:
                print(f"Processed file: {file_path.name}")
                wb = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(
                    process_empirical_sheet(
                        wb=wb,
                        labels_meta=labels_meta,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        wb=wb,
                        labels_meta=labels_meta,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed while processing ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
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
