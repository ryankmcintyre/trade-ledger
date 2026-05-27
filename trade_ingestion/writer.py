from __future__ import annotations

from pathlib import Path
from typing import Any, Sequence

import xlwings as xw

from trade_ingestion.models import CanonicalTrade

TABLE_NAME = "tbl_trades"
WRITABLE_COLUMNS: tuple[str, ...] = (
    "lot_id",
    "trade_id",
    "underlying",
    "symbol",
    "open_date",
    "exp_date",
    "call_or_put",
    "side",
    "strike",
    "stock_price_open",
    "premium",
    "quantity",
    "fees",
    "exit_price",
    "close_date",
    "account",
)


# TODO: The requirements mention reading existing trade_id values, but the matcher and schema use
# TODO: lot_id as the primary dedup key. The writer therefore prefers lot_id when available and
# TODO: falls back to trade_id only to avoid duplicate appends in workbooks with legacy layouts.
def write_trades(workbook_path: Path, trades: list[CanonicalTrade]) -> int:
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        table = _find_table(workbook, TABLE_NAME)
        headers = _table_headers(table)
        existing_keys = _existing_dedup_values(table, headers)
        writable_headers = [header for header in headers if header in WRITABLE_COLUMNS]
        header_positions = {header: headers.index(header) + 1 for header in writable_headers}
        pending = [trade for trade in trades if trade.lot_id not in existing_keys and trade.trade_id not in existing_keys]

        for trade in pending:
            row = table.ListRows.Add()
            for header in writable_headers:
                row.Range.Cells(1, header_positions[header]).Value = getattr(trade, header)

        workbook.save()
        return len(pending)
    finally:
        if not was_open:
            workbook.close()
            app.quit()


def read_existing_lot_ids(workbook_path: Path) -> set[str]:
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        table = _find_table(workbook, TABLE_NAME)
        headers = _table_headers(table)
        if "lot_id" not in headers:
            return set()

        lot_index = headers.index("lot_id")
        data_range = getattr(table, "DataBodyRange", None)
        if data_range is None or data_range.Value in (None, ""):
            return set()

        rows = _normalize_table_rows(data_range.Value, len(headers))
        return {str(row[lot_index]) for row in rows if row[lot_index] not in (None, "")}
    finally:
        if not was_open:
            workbook.close()
            app.quit()


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


def _find_table(workbook: Any, table_name: str) -> Any:
    for sheet in workbook.sheets:
        try:
            return sheet.api.ListObjects(table_name)
        except Exception:
            continue
    raise ValueError(f"Could not find table {table_name!r}")


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


def _existing_dedup_values(table: Any, headers: Sequence[str]) -> set[str]:
    values: set[str] = set()
    data_range = getattr(table, "DataBodyRange", None)
    if data_range is None or data_range.Value in (None, ""):
        return values

    rows = _normalize_table_rows(data_range.Value, len(headers))
    lot_index = headers.index("lot_id") if "lot_id" in headers else None
    trade_index = headers.index("trade_id") if "trade_id" in headers else None
    for row in rows:
        if lot_index is not None and row[lot_index] not in (None, ""):
            values.add(str(row[lot_index]))
        if trade_index is not None and row[trade_index] not in (None, ""):
            values.add(str(row[trade_index]))
    return values


def _normalize_table_rows(raw_value: Any, width: int) -> list[list[Any]]:
    if isinstance(raw_value, tuple):
        raw_value = [list(item) if isinstance(item, tuple) else item for item in raw_value]
    if isinstance(raw_value, list):
        if raw_value and not isinstance(raw_value[0], list):
            return [list(raw_value)]
        return [list(row) for row in raw_value]
    return [[raw_value] + [None] * (width - 1)]
