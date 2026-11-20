"""Streaming CSV / XLSX / XLS readers that preserve original cell values."""

from __future__ import annotations

import codecs
import csv
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from .columns import RESULT_COLUMNS, detect_columns

csv.field_size_limit(min(2**31 - 1, 50 * 1024 * 1024))


def safe_columns(headers: Iterable) -> tuple[list[str], list[dict]]:
    """Keep names where possible; make blank/duplicate/reserved names explicit."""
    used, result, renames = set(RESULT_COLUMNS), [], []
    headers = list(headers)
    # Drop trailing blank headers (common in Excel exports).
    while headers and (headers[-1] is None or not str(headers[-1]).strip()):
        headers.pop()
    for index, raw in enumerate(headers):
        original = "" if raw is None else str(raw).strip()
        base = original if original else f"Column_{index + 1}"
        candidate = f"{base}_Original" if base in RESULT_COLUMNS else base
        stem, suffix = candidate, 2
        while candidate in used:
            candidate = f"{stem}_{suffix}"
            suffix += 1
        used.add(candidate)
        result.append(candidate)
        if candidate != original:
            renames.append({"position": index + 1, "original": original, "renamed": candidate})
    if not result:
        raise ValueError("The file has no column headers. The first row must contain column names.")
    return result, renames


def _csv_format(path: Path) -> tuple[str, str]:
    """Validate UTF-8 incrementally, then fall back to Windows-1252 / Latin-1."""
    decoder = codecs.getincrementaldecoder("utf-8-sig")("strict")
    encoding = "utf-8-sig"
    try:
        with path.open("rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                decoder.decode(block)
            decoder.decode(b"", final=True)
    except UnicodeDecodeError:
        encoding = "cp1252"
        try:
            with path.open("rb") as handle:
                for block in iter(lambda: handle.read(1024 * 1024), b""):
                    block.decode("cp1252")
        except UnicodeDecodeError:
            encoding = "latin-1"
    with path.open("r", encoding=encoding, newline="") as handle:
        sample = handle.read(65536)
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        first = sample.splitlines()[0] if sample else ""
        delimiter = max(",;\t|", key=first.count) if first else ","
        if not first.count(delimiter):
            delimiter = ","
    return encoding, delimiter


def _is_blank(row: Iterable) -> bool:
    return all(value is None or (isinstance(value, str) and not value.strip()) for value in row)


@contextmanager
def _xlsx_reader(path):
    with Path(path).open("rb") as handle:
        workbook = load_workbook(handle, read_only=True, data_only=True)
        try:
            yield workbook
        finally:
            workbook.close()


def raw_rows(path: Path, info: dict):
    """Yield data rows (header skipped), padded to the header width."""
    width = len(info["columns"])
    extension = info["extension"]
    if extension == ".csv":
        with Path(path).open("r", encoding=info["encoding"], newline="") as handle:
            reader = csv.reader(handle, delimiter=info["delimiter"])
            next(reader, None)
            for source_row, row in enumerate(reader, 2):
                if _is_blank(row):
                    continue
                if len(row) > width:
                    extra = row[width:]
                    if not _is_blank(extra):
                        raise ValueError(
                            f"CSV row {source_row} has more cells than the header row. "
                            "Check the delimiter or add a header for every column."
                        )
                    row = row[:width]
                yield row + [""] * (width - len(row))
    elif extension == ".xlsx":
        with _xlsx_reader(path) as workbook:
            sheet = workbook[info["sheet_name"]]
            sheet.reset_dimensions()
            rows = sheet.iter_rows(values_only=True)
            next(rows, None)
            for row in rows:
                row = list(row)
                if _is_blank(row):
                    continue
                if len(row) > width:
                    if not _is_blank(row[width:]):
                        raise ValueError(
                            "An Excel data row extends beyond the header. Add names to all data columns."
                        )
                    row = row[:width]
                yield row + [None] * (width - len(row))
    else:
        import xlrd

        workbook = xlrd.open_workbook(str(path), on_demand=True)
        try:
            sheet = workbook.sheet_by_name(info["sheet_name"])
            for row_index in range(1, sheet.nrows):
                row = []
                for column in range(min(sheet.ncols, width)):
                    cell = sheet.cell(row_index, column)
                    value = cell.value
                    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        value = None
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        value = xlrd.xldate.xldate_as_datetime(value, workbook.datemode)
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        value = bool(value)
                    elif cell.ctype == xlrd.XL_CELL_NUMBER and float(value).is_integer():
                        value = int(value)
                    row.append(value)
                if not _is_blank(row):
                    yield row + [None] * (width - len(row))
        finally:
            workbook.release_resources()


def inspect_input(path, filename: str | None = None, sheet_name: str | None = None) -> dict:
    """Read headers/sheets and count rows without loading the whole file."""
    path = Path(path)
    extension = Path(filename or path.name).suffix.lower()
    if extension not in {".csv", ".xlsx", ".xls"}:
        raise ValueError("Choose a CSV, XLSX or XLS file.")
    info = {"extension": extension, "sheets": [], "sheet_name": None, "encoding": None, "delimiter": None}
    if extension == ".csv":
        info["encoding"], info["delimiter"] = _csv_format(path)
        with path.open("r", encoding=info["encoding"], newline="") as handle:
            headers = next(csv.reader(handle, delimiter=info["delimiter"]), [])
    elif extension == ".xlsx":
        with _xlsx_reader(path) as workbook:
            info["sheets"] = workbook.sheetnames
            selected = sheet_name or workbook.sheetnames[0]
            if selected not in workbook.sheetnames:
                raise ValueError("The selected sheet no longer exists.")
            info["sheet_name"] = selected
            sheet = workbook[selected]
            sheet.reset_dimensions()
            headers = next(sheet.iter_rows(values_only=True), ())
    else:
        import xlrd

        workbook = xlrd.open_workbook(str(path), on_demand=True)
        try:
            info["sheets"] = workbook.sheet_names()
            info["sheet_name"] = sheet_name or info["sheets"][0]
            sheet = workbook.sheet_by_name(info["sheet_name"])
            headers = sheet.row_values(0) if sheet.nrows else []
        finally:
            workbook.release_resources()
    info["columns"], info["column_renames"] = safe_columns(headers)
    info["detected"] = detect_columns(info["columns"])
    info["email_index"] = info["detected"].get("email")
    info["row_count"] = sum(1 for _ in raw_rows(path, info))
    return info
