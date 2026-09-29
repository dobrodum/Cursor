#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these two paths before running the script.
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

MONTH_LOOKUP = {
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

DAY_LOOKUP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

Cell = Tuple[int, int]

FILE_PATTERN = re.compile(
    r"^.+?\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<phase>Early|Mid|Late)"
    r"(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
    flags=re.IGNORECASE,
)


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetScan:
    start_row: int
    start_col: int
    values: List[List[Any]]
    labels: Dict[str, Cell]
    anchor_max: Optional[Cell]


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_label(value: str) -> str:
    cleaned = value.replace("\n", " ").replace("\r", " ").strip().lower()
    return " ".join(cleaned.split())


def parse_model_meta(file_name: str) -> Optional[ModelMeta]:
    stem = Path(file_name).stem
    match = FILE_PATTERN.search(stem)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    phase_raw = match.group("phase").lower()
    month_raw = match.group("month").lower()[:3]
    year = int(match.group("year"))

    month_num = MONTH_LOOKUP.get(month_raw)
    day_num = DAY_LOOKUP.get(phase_raw)
    if month_num is None or day_num is None:
        return None

    phase = phase_raw.capitalize()
    month_short = datetime(year, month_num, 1).strftime("%b")
    model_period = f"{phase}{month_short}_{year}"
    model_date = f"{year:04d}-{month_num:02d}-{day_num:02d}"
    model = f"{ticker}_{model_period}"
    return ModelMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def safe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        pct = raw.endswith("%")
        raw = raw.replace(",", "").rstrip("%")
        try:
            number = float(raw)
        except ValueError:
            return None
        return number / 100.0 if pct else number
    return None


def safe_int(value: Any, default: int) -> int:
    as_float = safe_float(value)
    if as_float is None:
        return default
    return int(round(as_float))


def safe_value(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def scan_sheet(sheet: xw.Sheet) -> SheetScan:
    used = sheet.used_range
    values = normalize_2d(used.value)
    labels: Dict[str, Cell] = {}
    anchor_max: Optional[Cell] = None

    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if not isinstance(cell_value, str):
                continue
            label = normalize_label(cell_value)
            if not label:
                continue
            abs_row = used.row + r_idx
            abs_col = used.column + c_idx
            labels.setdefault(label, (abs_row, abs_col))
            if anchor_max is None and label == "max":
                anchor_max = (abs_row, abs_col)

    return SheetScan(
        start_row=used.row,
        start_col=used.column,
        values=values,
        labels=labels,
        anchor_max=anchor_max,
    )


def scan_value(scan: SheetScan, row: int, col: int) -> Any:
    r_idx = row - scan.start_row
    c_idx = col - scan.start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(scan.values):
        return None
    row_values = scan.values[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def find_label_cell(scan: SheetScan, aliases: Sequence[str]) -> Optional[Cell]:
    for alias in aliases:
        position = scan.labels.get(normalize_label(alias))
        if position is not None:
            return position
    return None


def find_value_cell_from_label(scan: SheetScan, aliases: Sequence[str]) -> Optional[Cell]:
    label_cell = find_label_cell(scan, aliases)
    if label_cell is None:
        return None

    row, col = label_cell
    candidates = [(row, col + 1), (row + 1, col), (row, col - 1), (row - 1, col)]
    for cand_row, cand_col in candidates:
        candidate_value = scan_value(scan, cand_row, cand_col)
        if candidate_value not in (None, ""):
            return (cand_row, cand_col)

    return (row, col + 1)


def build_offsets(
    anchor: Cell,
    scan: SheetScan,
    defaults: Dict[str, Cell],
    aliases: Dict[str, Sequence[str]],
) -> Dict[str, Cell]:
    offsets = dict(defaults)
    anchor_row, anchor_col = anchor
    for key, alias_list in aliases.items():
        value_cell = find_value_cell_from_label(scan, alias_list)
        if value_cell is None:
            continue
        value_row, value_col = value_cell
        offsets[key] = (value_row - anchor_row, value_col - anchor_col)
    return offsets


def cell_from_offset(anchor: Cell, offset: Cell) -> Cell:
    return (anchor[0] + offset[0], anchor[1] + offset[1])


def get_offset_value(sheet: xw.Sheet, anchor: Cell, offsets: Dict[str, Cell], key: str) -> Any:
    if key not in offsets:
        return None
    row, col = cell_from_offset(anchor, offsets[key])
    return sheet.range((row, col)).value


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_no_save(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    api = getattr(workbook, "api", None)
    if api is None:
        return

    try:
        api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        api.Close(False)
    except Exception:
        pass


def output_path_for_run(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def process_empirical_model(
    workbook: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical for {source_file}: missing sheet 'Empirical Model'")
        return rows

    scan = scan_sheet(sheet)
    if scan.anchor_max is None:
        print(f"Skipped empirical for {source_file}: could not find 'max' anchor")
        return rows

    anchor = scan.anchor_max
    anchor_row, anchor_col = anchor

    default_offsets = {
        "num_quarters_input": (-3, 1),
        "last_quarter_used": (-4, 1),
        "forecast_value": (-1, 1),
        "actual_value": (2, 1),
        "forecast_max": (0, 1),
        "forecast_min": (1, 1),
        "avg_penetration_pct": (-2, 1),
        "quarterly_sales": (3, 1),
        "reported_sales": (2, 1),
        "growth_rate_pct": (4, 1),
        "sales_captured_in_db_pct": (5, 1),
    }

    aliases = {
        "num_quarters_input": (
            "num quarters used",
            "number of quarters used",
            "quarters used",
        ),
        "last_quarter_used": (
            "last quarter used",
            "last qtr used",
        ),
        "forecast_value": (
            "estimated total sold",
            "forecast total",
            "forecast value",
        ),
        "actual_value": (
            "reported sales",
            "actual sales",
            "actual value",
        ),
        "forecast_max": ("max",),
        "forecast_min": ("min",),
        "avg_penetration_pct": (
            "avg penetration %",
            "avg penetration pct",
            "average penetration",
        ),
        "quarterly_sales": (
            "quarterly sales",
            "quarterly sale",
        ),
        "reported_sales": (
            "reported sales",
            "actual sales",
        ),
        "growth_rate_pct": (
            "growth rate %",
            "growth rate pct",
            "growth rate",
        ),
        "sales_captured_in_db_pct": (
            "sales captured in db %",
            "sales captured in db pct",
            "sales captured in db",
        ),
    }

    offsets = build_offsets(anchor, scan, default_offsets, aliases)

    num_quarters_cell = sheet.range(cell_from_offset(anchor, offsets["num_quarters_input"]))
    avg_pen_cell = sheet.range(cell_from_offset(anchor, offsets["avg_penetration_pct"]))

    penetration_label_cell = find_label_cell(
        scan,
        (
            "sales captured in db %",
            "sales captured in db pct",
            "quarterly penetration",
            "penetration",
        ),
    )
    penetration_row = penetration_label_cell[0] if penetration_label_cell else anchor_row - 1

    for n_quarters in range(1, N_QUARTERS + 1):
        num_quarters_cell.value = n_quarters
        start_col = max(1, anchor_col - n_quarters)
        end_col = max(1, anchor_col - 1)
        avg_formula = f"=AVERAGE(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
        set_formula2_r1c1(avg_pen_cell, avg_formula)

        workbook.app.calculate()

        num_used = safe_int(num_quarters_cell.value, n_quarters)
        last_quarter_used = get_offset_value(sheet, anchor, offsets, "last_quarter_used")
        forecast_value = safe_float(get_offset_value(sheet, anchor, offsets, "forecast_value"))
        actual_value = safe_float(get_offset_value(sheet, anchor, offsets, "actual_value"))
        forecast_max = safe_float(get_offset_value(sheet, anchor, offsets, "forecast_max"))
        forecast_min = safe_float(get_offset_value(sheet, anchor, offsets, "forecast_min"))
        avg_penetration_pct = safe_float(avg_pen_cell.value)
        quarterly_sales = safe_float(get_offset_value(sheet, anchor, offsets, "quarterly_sales"))
        reported_sales = safe_float(get_offset_value(sheet, anchor, offsets, "reported_sales"))
        growth_rate_pct = safe_float(get_offset_value(sheet, anchor, offsets, "growth_rate_pct"))
        sales_captured_in_db_pct = safe_float(
            get_offset_value(sheet, anchor, offsets, "sales_captured_in_db_pct")
        )

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_used,
                "last_quarter_used": safe_value(last_quarter_used),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
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
        )

    return rows


def almost_equal(left: Any, right: Any, eps: float = 1e-9) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return abs(float(left) - float(right)) <= eps
    return left == right


def is_duplicate_regression_row(current: Dict[str, Any], previous: Dict[str, Any]) -> bool:
    keys = (
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    )
    return all(almost_equal(current.get(key), previous.get(key)) for key in keys)


def process_regression_model(
    workbook: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression for {source_file}: missing sheet 'Regression Model'")
        return rows

    scan = scan_sheet(sheet)
    if scan.anchor_max is None:
        print(f"Skipped regression for {source_file}: could not find 'max' anchor")
        return rows

    anchor = scan.anchor_max
    anchor_row, anchor_col = anchor

    default_offsets = {
        "num_quarters_input": (-3, 1),
        "forecast_total_without_sa": (-1, 1),
        "actual_value": (2, 1),
        "forecast_max": (0, 1),
        "forecast_min": (1, 1),
    }

    aliases = {
        "num_quarters_input": (
            "num quarters used",
            "number of quarters used",
            "quarters used",
        ),
        "forecast_total_without_sa": (
            "tot fcst w/o sa",
            "tot fcst wo sa",
            "tot fcst without sa",
            "forecast total w/o sa",
        ),
        "actual_value": (
            "actual value",
            "reported sales",
            "actual sales",
        ),
        "forecast_max": ("max",),
        "forecast_min": ("min",),
    }

    offsets = build_offsets(anchor, scan, default_offsets, aliases)
    num_quarters_cell = sheet.range(cell_from_offset(anchor, offsets["num_quarters_input"]))

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    intercept_cell = sheet.range((anchor_row, anchor_col + 3))
    slope_cell = sheet.range((anchor_row + 1, anchor_col + 3))

    for n_quarters in range(1, N_QUARTERS + 1):
        num_quarters_cell.value = n_quarters

        end_row = anchor_row - 1
        start_row = max(1, end_row - n_quarters + 1)

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)

        workbook.app.calculate()

        num_used = safe_int(num_quarters_cell.value, n_quarters)
        intercept = safe_float(intercept_cell.value)
        slope = safe_float(slope_cell.value)
        forecast_value = safe_float(
            get_offset_value(sheet, anchor, offsets, "forecast_total_without_sa")
        )
        forecast_max = safe_float(get_offset_value(sheet, anchor, offsets, "forecast_max"))
        forecast_min = safe_float(get_offset_value(sheet, anchor, offsets, "forecast_min"))
        actual_value = safe_float(get_offset_value(sheet, anchor, offsets, "actual_value"))
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_used,
            "num_quarters_used": num_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(row, rows[-1]):
            continue
        rows.append(row)

    return rows


def write_sheet(worksheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([safe_value(row.get(column)) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    max_row = max(worksheet.max_row, 1)
    max_col = len(columns)
    worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    for col_idx, name in enumerate(columns, start=1):
        longest = len(name)
        for row_idx in range(2, worksheet.max_row + 1):
            cell_val = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_val is None:
                continue
            longest = max(longest, len(str(cell_val)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 2, 12), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def iter_input_files(source_dir: Path) -> Iterable[Path]:
    for path in sorted(source_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        yield path


def main() -> None:
    source_dir = input_dir.expanduser().resolve()
    destination_dir = output_dir.expanduser().resolve()

    if not source_dir.exists() or not source_dir.is_dir():
        raise SystemExit(f"Input directory does not exist: {source_dir}")

    destination_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_path_for_run(source_dir, destination_dir)

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

        for file_path in iter_input_files(source_dir):
            file_name = file_path.name

            if file_name.startswith("~"):
                print(f"Skipped {file_name}: temporary Excel file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_name}: not an .xlsx file")
                continue

            if file_path.resolve() == output_path.resolve():
                print(f"Skipped {file_name}: output workbook for this run")
                continue

            meta = parse_model_meta(file_name)
            if meta is None:
                print(f"Skipped {file_name}: filename did not match expected pattern")
                continue

            print(f"Processed {file_name}")

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_model(workbook, meta, file_name))
                regression_rows.extend(process_regression_model(workbook, meta, file_name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_name}: workbook processing error ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_no_save(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
