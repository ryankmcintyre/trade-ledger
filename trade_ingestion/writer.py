from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw

from constants import DEDUP_COLUMNS, FIELD_TO_COLUMN, TABLE_NAME
from trade_ingestion.models import CanonicalTrade


def write_trades(workbook_path: Path, sheet_name: str, trades: list[CanonicalTrade]) -> int:
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        table = _find_table(workbook, sheet_name, TABLE_NAME)
        headers = _table_headers(table)
        existing_keys = _existing_dedup_keys(table, headers)

        # Build column position map for writable fields
        header_positions: dict[str, int] = {}
        for col_name in FIELD_TO_COLUMN.values():
            if col_name in headers:
                header_positions[col_name] = headers.index(col_name) + 1

        pending: list[CanonicalTrade] = []
        for trade in trades:
            key = _make_dedup_key(trade, headers)
            if key not in existing_keys:
                pending.append(trade)

        for trade in pending:
            row = table.ListRows.Add()
            for field_name, col_name in FIELD_TO_COLUMN.items():
                if col_name not in header_positions:
                    continue
                value = getattr(trade, field_name, None)
                if value is None:
                    continue
                row.Range.Cells(1, header_positions[col_name]).Value = value

        workbook.save()
        return len(pending)
    finally:
        if not was_open:
            workbook.close()
            app.quit()


def read_existing_lot_ids(workbook_path: Path, sheet_name: str) -> set[str]:
    """DEPRECATED: read existing composite dedup keys as pipe-delimited strings."""
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        table = _find_table(workbook, sheet_name, TABLE_NAME)
        headers = _table_headers(table)
        return _existing_dedup_keys(table, headers)
    finally:
        if not was_open:
            workbook.close()
            app.quit()


def _make_dedup_key(trade: CanonicalTrade, headers: list[str]) -> str:
    """Build a composite dedup key from trade fields matching DEDUP_COLUMNS."""
    parts: list[str] = []
    for col_name in DEDUP_COLUMNS:
        if col_name == "Stock":
            parts.append(str(trade.stock or ""))
        elif col_name == "Open Date":
            parts.append(trade.open_date.isoformat() if trade.open_date else "")
        elif col_name == "B/S":
            parts.append(trade.side or "")
        elif col_name == "C":
            parts.append(f"{trade.quantity:g}")
    return "|".join(parts)


def _existing_dedup_keys(table: Any, headers: list[str]) -> set[str]:
    """Read existing rows and build composite dedup keys."""
    values: set[str] = set()
    data_range = getattr(table, "DataBodyRange", None)
    if data_range is None or data_range.Value in (None, ""):
        return values

    rows = _normalize_table_rows(data_range.Value, len(headers))

    # Find column indices for dedup columns
    col_indices: dict[str, int | None] = {}
    for col_name in DEDUP_COLUMNS:
        col_indices[col_name] = headers.index(col_name) if col_name in headers else None

    for row in rows:
        parts: list[str] = []
        for col_name in DEDUP_COLUMNS:
            idx = col_indices[col_name]
            if idx is None:
                parts.append("")
                continue
            val = row[idx]
            if val is None or val == "":
                parts.append("")
            elif col_name == "Open Date" and isinstance(val, (int, float)):
                # Excel serial date — convert to ISO format for comparison
                d = _excel_serial_to_date(val)
                parts.append(d.isoformat() if d else "")
            elif col_name == "C":
                parts.append(f"{float(val):g}")
            elif col_name == "Open Date" and hasattr(val, "date") and callable(getattr(val, "date", None)):
                parts.append(val.date().isoformat())
            else:
                parts.append(str(val))
        key = "|".join(parts)
        if any(p for p in parts):
            values.add(key)

    return values


def _excel_serial_to_date(serial: float) -> date | None:
    """Convert Excel serial number to Python date."""
    try:
        from datetime import datetime, timedelta
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=int(serial))).date()
    except (ValueError, OverflowError):
        return None


def _open_workbook(workbook_path: Path) -> tuple[Any, Any, bool]:
    resolved_path = str(workbook_path.resolve())
    existing_book = _find_open_book(resolved_path)
    if existing_book is not None:
        return existing_book, existing_book.app, True

    app = xw.App(visible=False, add_book=False)
    workbook = app.books.open(resolved_path)
    return workbook, app, False


def _find_open_book(resolved_path: str) -> Any | None:
    for app in xw.apps:
        for book in app.books:
            fullname = str(Path(book.fullname).resolve())
            if fullname == resolved_path:
                return book
    return None


def _find_table(workbook: Any, sheet_name: str, table_name: str) -> Any:
    sheet = _find_sheet(workbook, sheet_name)
    try:
        return sheet.api.ListObjects(table_name)
    except Exception as exc:
        raise ValueError(
            f"Could not find table {table_name!r} on worksheet {sheet_name!r}"
        ) from exc


def _find_sheet(workbook: Any, sheet_name: str) -> Any:
    available: list[str] = []
    for sheet in workbook.sheets:
        name = str(sheet.name)
        available.append(name)
        if name.casefold() == sheet_name.casefold():
            return sheet
    known = ", ".join(repr(name) for name in available) or "none"
    raise ValueError(
        f"Could not find worksheet {sheet_name!r} in the workbook. Available worksheets: {known}"
    )


def _table_headers(table: Any) -> list[str]:
    header_values = table.HeaderRowRange.Value
    if isinstance(header_values, tuple):
        if header_values and isinstance(header_values[0], tuple):
            return [str(value) for value in header_values[0]]
        return [str(value) for value in header_values]
    if isinstance(header_values, list):
        if header_values and isinstance(header_values[0], list):
            return [str(value) for value in header_values[0]]
        return [str(value) for value in header_values]
    return [str(header_values)]


def _normalize_table_rows(raw_value: Any, width: int) -> list[list[Any]]:
    if isinstance(raw_value, tuple):
        raw_value = [list(item) if isinstance(item, tuple) else item for item in raw_value]
    if isinstance(raw_value, list):
        if raw_value and not isinstance(raw_value[0], list):
            return [list(raw_value)]
        return [list(row) for row in raw_value]
    return [[raw_value] + [None] * (width - 1)]
