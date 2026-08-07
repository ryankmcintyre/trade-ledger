from __future__ import annotations

import time
from datetime import date
from pathlib import Path
from typing import Any, Callable, TypeVar

import xlwings as xw

from constants import DEDUP_COLUMNS, FIELD_TO_COLUMN, TABLE_NAME
from trade_ingestion.models import CanonicalTrade

try:  # pragma: no cover - pywin32 is only importable on Windows
    import pythoncom
    import pywintypes
except ImportError:  # pragma: no cover - non-Windows/test environments
    pythoncom = None
    pywintypes = None

# RPC_E_CALL_REJECTED: Excel's COM server was busy (e.g. still finishing an
# open/save, showing a dialog, or recalculating) and rejected the call.
# RPC_E_SERVERCALL_RETRYLATER: the server explicitly asked the caller to retry.
# Both are transient — retrying (after pumping the message queue) resolves them.
_RETRYABLE_HRESULTS = frozenset({-2147418111, -2147417846})

# TODO: retry attempt count / backoff are conservative defaults; the issue
# does not specify exact timing requirements, so these are tunable if real
# workbooks need longer waits (e.g. very large tables/pivot recalculation).
_COM_RETRY_ATTEMPTS = 5
_COM_RETRY_BASE_DELAY = 0.2

T = TypeVar("T")


class ComRetryExhaustedError(RuntimeError):
    """Raised when a COM call keeps failing with a transient/busy error
    even after all retry attempts have been exhausted."""


def _is_retryable_com_error(exc: BaseException) -> bool:
    """Return True if `exc` looks like a transient COM busy/rejected-call error.

    Excel occasionally rejects COM calls (RPC_E_CALL_REJECTED /
    RPC_E_SERVERCALL_RETRYLATER) while it is busy. When that happens mid
    enumeration (e.g. iterating `workbook.sheets`), the generated win32com
    wrapper's `__iter__` swallows the real `com_error` and raises a secondary,
    misleading `TypeError: This object does not support enumeration` instead.
    We treat both shapes as retryable.
    """
    if pywintypes is not None and isinstance(exc, pywintypes.com_error):
        hresult = exc.args[0] if exc.args else None
        return hresult in _RETRYABLE_HRESULTS
    if isinstance(exc, TypeError) and "does not support enumeration" in str(exc):
        return True
    return False


def _call_with_com_retry(
    func: Callable[[], T],
    *,
    attempts: int = _COM_RETRY_ATTEMPTS,
    base_delay: float = _COM_RETRY_BASE_DELAY,
) -> T:
    """Call `func`, retrying on transient COM busy/rejected-call errors.

    Between attempts we pump Excel's pending Windows message queue (this is
    what actually clears RPC_E_CALL_REJECTED) and back off briefly before
    retrying. If every attempt fails, raise a clear ComRetryExhaustedError
    instead of letting the confusing raw COM/TypeError propagate.
    """
    last_error: BaseException | None = None
    for attempt in range(1, attempts + 1):
        try:
            return func()
        except Exception as exc:  # noqa: BLE001 - re-raised below if not retryable
            if not _is_retryable_com_error(exc):
                raise
            last_error = exc
            if attempt < attempts:
                if pythoncom is not None:
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:  # noqa: BLE001 - pumping is best-effort
                        # If pumping fails (e.g. COM not initialized on this
                        # thread) we still want to keep retrying rather than
                        # abort the whole retry loop.
                        pass
                time.sleep(base_delay * attempt)
    raise ComRetryExhaustedError(
        "Excel COM server was busy/unresponsive and rejected the call after "
        f"{attempts} retries. The workbook may be showing a dialog, "
        "recalculating, or otherwise blocked. Original error: "
        f"{last_error!r}"
    ) from last_error


def write_trades(workbook_path: Path, sheet_name: str, trades: list[CanonicalTrade]) -> int:
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        sheet = _find_sheet(workbook, sheet_name)
        table = _find_table(sheet, TABLE_NAME)
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

        insertion_position = _last_populated_row_position(table, len(headers)) + 1
        for trade in pending:
            row = _call_with_com_retry(
                lambda insertion_position=insertion_position: table.ListRows.Add(
                    Position=insertion_position
                )
            )
            base_row = _call_with_com_retry(lambda: row.Range.Row)
            base_column = _call_with_com_retry(lambda: row.Range.Column)
            for field_name, col_name in FIELD_TO_COLUMN.items():
                if col_name not in header_positions:
                    continue
                value = getattr(trade, field_name, None)
                if value is None:
                    continue
                cell_index = header_positions[col_name]

                def _write_cell(
                    base_row: int = base_row,
                    base_column: int = base_column,
                    cell_index: int = cell_index,
                    value: Any = value,
                ) -> None:
                    sheet.range((base_row, base_column + cell_index - 1)).value = value

                _call_with_com_retry(_write_cell)
            insertion_position += 1

        _call_with_com_retry(workbook.save)
        return len(pending)
    finally:
        if not was_open:
            workbook.close()
            app.quit()


def read_existing_lot_ids(workbook_path: Path, sheet_name: str) -> set[str]:
    """DEPRECATED: read existing composite dedup keys as pipe-delimited strings."""
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        sheet = _find_sheet(workbook, sheet_name)
        table = _find_table(sheet, TABLE_NAME)
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
    existing_book = _call_with_com_retry(lambda: _find_open_book(resolved_path))
    if existing_book is not None:
        return existing_book, existing_book.app, True

    app = xw.App(visible=False, add_book=False)
    # Reduce the chance Excel blocks on a dialog (e.g. file-format prompts,
    # "keep changes") which is a common trigger for RPC_E_CALL_REJECTED.
    # Only touch these settings on an app we created ourselves — never on
    # the user's own already-open Excel session (see was_open branch above).
    try:
        app.display_alerts = False
        app.screen_updating = False
    except Exception:  # noqa: BLE001 - best-effort, not critical to success
        pass
    workbook = _call_with_com_retry(lambda: app.books.open(resolved_path))
    return workbook, app, False


def _find_open_book(resolved_path: str) -> Any | None:
    for app in xw.apps:
        for book in app.books:
            fullname = str(Path(book.fullname).resolve())
            if fullname == resolved_path:
                return book
    return None


def _find_table(sheet: Any, table_name: str) -> Any:
    try:
        return _call_with_com_retry(lambda: sheet.api.ListObjects(table_name))
    except ComRetryExhaustedError:
        # Exhausted retries on a transient COM error — surface that clearly
        # rather than masking it as "table not found".
        raise
    except Exception as exc:
        raise ValueError(
            f"Could not find table {table_name!r} on worksheet {sheet.name!r}"
        ) from exc


def _find_sheet(workbook: Any, sheet_name: str) -> Any:
    def _search() -> Any:
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

    return _call_with_com_retry(_search)


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


def _last_populated_row_position(table: Any, headers: list[str]) -> int:
    """Return the one-based position of the table's last non-empty data row.

    A row is considered populated if any of the user-entered dedup columns has a value.
    This avoids treating formula-driven columns as data when trailing rows are otherwise blank.
    """
    data_range = getattr(table, "DataBodyRange", None)
    if data_range is None:
        return 0

    value = data_range.Value
    if value in (None, ""):
        return 0

    rows = _normalize_table_rows(value, len(headers))
    indices = [headers.index(col) for col in DEDUP_COLUMNS if col in headers]

    for position in range(len(rows), 0, -1):
        row = rows[position - 1]
        if indices:
            if any((row[idx] if idx < len(row) else None) not in (None, "") for idx in indices):
                return position
        elif any(cell not in (None, "") for cell in row):
            return position
    return 0
