from __future__ import annotations

import time
from dataclasses import dataclass, field, replace
from datetime import date
from pathlib import Path
from typing import Any, Callable, TypeVar

import xlwings as xw

from constants import (
    DEDUP_COLUMNS,
    FIELD_TO_COLUMN,
    LINKED_DATA_TYPE_CULTURE,
    LOT_ID_COLUMN,
    STOCK_FIELD_NAME,
    STOCK_SYMBOL_COLUMN,
    STOCKS_SERVICE_ID,
    TABLE_NAME,
)
from trade_ingestion.models import CanonicalTrade, make_trade_id
from trade_ingestion.retry import resolve_with_retry

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

# Tolerance for floating-point quantity comparisons when reconciling closes
# against existing open rows (mirrors matcher.MATCH_EPSILON).
MATCH_EPSILON = 1e-9


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


@dataclass(slots=True)
class ConversionFailure:
    """Details about a failed stock-conversion retry for a single trade."""

    trade: CanonicalTrade
    input_ticker: str
    attempted_ticker: str | None
    error: str


@dataclass(slots=True)
class WriteResult:
    """Outcome of a write_trades run."""

    rows_written: int
    # Underlying tickers whose Column A cell could not be converted to the Stocks
    # linked data type; those rows keep the plain Column A text and their
    # formula-driven "Stock Symbol"/"Current Stock Price" columns will not resolve.
    # NOTE: the underlying ticker is reported rather than CanonicalTrade.stock,
    # because the latter is a Column A *display* value that may have been remapped
    # via UNDERLYING_DISPLAY_MAP (e.g. "SPXW" -> "S&P 500 INDEX").
    failed_conversions: list[str]
    conversion_failures: list[ConversionFailure] = field(default_factory=list)


def write_trades(
    workbook_path: Path,
    table_name: str,
    trades: list[CanonicalTrade],
    ticker_prompt: Callable[[str, CanonicalTrade], str | None] | None = None,
) -> int:
    """Append trades to the named Excel table and return the number of rows written."""
    return write_trades_detailed(workbook_path, table_name, trades, ticker_prompt=ticker_prompt).rows_written


def write_trades_detailed(
    workbook_path: Path,
    table_name: str,
    trades: list[CanonicalTrade],
    ticker_prompt: Callable[[str, CanonicalTrade], str | None] | None = None,
) -> WriteResult:
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        table, sheet = _find_table_and_sheet(workbook, table_name)
        headers = _table_headers(table)
        existing_keys = _existing_dedup_keys(table, headers)

        # Build column position map for writable fields
        header_positions: dict[str, int] = {}
        for col_name in FIELD_TO_COLUMN.values():
            if col_name in headers:
                header_positions[col_name] = headers.index(col_name) + 1

        pending: list[CanonicalTrade] = []
        updated_rows = 0
        for trade in trades:
            match = _find_existing_row_to_update(table, headers, trade)
            if match is not None:
                row_index, row = match
                remaining_quantity = _remaining_open_quantity(row, headers, trade.quantity)
                if remaining_quantity is not None and remaining_quantity > MATCH_EPSILON:
                    # Partial close: the matched row's open quantity is larger than what
                    # this close accounts for (e.g. two separate sell tickets closing a
                    # single larger open lot). The closed portion is written as its own
                    # new row carrying the original open-side data, so the composite
                    # dedup key (when there is no Lot ID column) must be derived only
                    # after merging those open-side fields in.
                    split_trade = _merge_open_fields_from_row(trade, row, headers)

                    row_quantity = _row_quantity(row, headers)
                    row_fees = _row_fees(row, headers)
                    open_fee_share = _allocate_fee(row_fees, row_quantity, remaining_quantity)
                    close_fee_share = _allocate_fee(row_fees, row_quantity, trade.quantity)
                    if close_fee_share is not None:
                        split_trade = replace(split_trade, fees=(split_trade.fees or 0.0) + close_fee_share)

                    key = _make_dedup_key(split_trade, headers)
                    if key in existing_keys:
                        # Already reconciled in a previous run/import — re-importing the
                        # same closing ticket must not shrink the (already-shrunk) open
                        # row a second time.
                        continue

                    # Shrink the existing row down to the quantity still open, carrying
                    # forward its proportional share of the row's original opening fees.
                    _reduce_existing_row_quantity(
                        sheet,
                        table,
                        headers,
                        header_positions,
                        row_index,
                        remaining_quantity,
                        remaining_fees=open_fee_share,
                    )
                    pending.append(split_trade)
                    existing_keys.add(key)
                else:
                    exact_trade = trade
                    if trade.open_date is None:
                        existing_open_fees = _row_fees(row, headers)
                        if existing_open_fees is not None:
                            exact_trade = replace(
                                trade, fees=existing_open_fees + (trade.fees or 0.0)
                            )
                    # Register this close's key (with the row's open-side fields
                    # merged in, matching the split path) so re-importing the same
                    # ticket after the row is fully closed — and thus no longer an
                    # open candidate — is recognized as already reconciled.
                    existing_keys.add(
                        _make_dedup_key(_merge_open_fields_from_row(trade, row, headers), headers)
                    )
                    _update_existing_trade_row(
                        sheet, table, headers, header_positions, row_index, exact_trade
                    )
                    updated_rows += 1
                continue

            key = _make_dedup_key(trade, headers)
            if key not in existing_keys:
                pending.append(trade)
                existing_keys.add(key)

        stock_column = header_positions.get(FIELD_TO_COLUMN["stock"])
        symbol_column = headers.index(STOCK_SYMBOL_COLUMN) + 1 if STOCK_SYMBOL_COLUMN in headers else None
        failed_conversions: list[str] = []
        conversion_failures: list[ConversionFailure] = []
        rows = _normalize_table_rows(getattr(table, "DataBodyRange", None).Value, len(headers))

        for trade in pending:
            insertion_position = _determine_insertion_position(rows, headers, trade)

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

            if stock_column is not None and trade.stock:
                converted, conversion_failure = _convert_stock_cell_with_recovery(
                    sheet,
                    row_number=base_row,
                    stock_cell_column=base_column + stock_column - 1,
                    symbol_cell_column=(
                        base_column + symbol_column - 1 if symbol_column is not None else None
                    ),
                    trade=trade,
                    ticker_prompt=ticker_prompt,
                )
                if not converted:
                    failed_conversions.append(trade.underlying or trade.stock)
                    if conversion_failure is not None:
                        conversion_failures.append(conversion_failure)

            row_values = [None] * len(headers)
            stock_index = header_positions.get(FIELD_TO_COLUMN["stock"])
            if stock_index is not None and trade.stock:
                row_values[stock_index - 1] = trade.stock
            open_date_index = header_positions.get(FIELD_TO_COLUMN["open_date"])
            if open_date_index is not None and trade.open_date is not None:
                row_values[open_date_index - 1] = trade.open_date
            rows.insert(insertion_position - 1, row_values)

        _call_with_com_retry(workbook.save)
        return WriteResult(
            rows_written=len(pending) + updated_rows,
            failed_conversions=failed_conversions,
            conversion_failures=conversion_failures,
        )
    finally:
        if not was_open:
            workbook.close()
            app.quit()


# NOTE: Column A ("Stock") is not a text column in the workbook — it must hold an
# NOTE: Excel Stocks *linked data type* entity, because "Stock Symbol" and
# NOTE: "Current Stock Price" are driven by _FV(A, ...) formulas which only accept a
# NOTE: rich value. Writing a plain ticker string leaves those columns unresolved.
# NOTE: Converting requires Microsoft 365 with the Stocks data type available and an
# NOTE: internet connection; otherwise the conversion is a no-op and we fall back to text.
def _convert_stock_cell_with_recovery(
    sheet: Any,
    *,
    row_number: int,
    stock_cell_column: int,
    symbol_cell_column: int | None,
    trade: CanonicalTrade,
    ticker_prompt: Callable[[str, CanonicalTrade], str | None] | None = None,
) -> tuple[bool, ConversionFailure | None]:
    input_ticker = trade.stock or trade.underlying or ""
    if not input_ticker:
        return False, None

    def try_convert(ticker: str) -> tuple[bool | None, str | None]:
        # Retrying with a replacement ticker means the cell must hold that
        # replacement's text before Excel can convert it — a no-op on the
        # first attempt, since the row-writing loop already wrote input_ticker.
        _set_plain_ticker(sheet, row_number, stock_cell_column, ticker)
        converted, error_message = _convert_stock_cell(
            sheet,
            row_number=row_number,
            stock_cell_column=stock_cell_column,
            symbol_cell_column=symbol_cell_column,
            ticker=ticker,
        )
        return (True if converted else None), error_message

    context_label = trade.trade_id or trade.lot_id or trade.symbol or "trade"
    resolved_prompt = (lambda value, _label, trade=trade: ticker_prompt(value, trade)) if ticker_prompt else None
    result, failure = resolve_with_retry(input_ticker, context_label, try_convert, resolved_prompt)
    if result:
        return True, None
    if failure is None:
        return False, None
    return False, ConversionFailure(
        trade=trade,
        input_ticker=failure.input_value,
        attempted_ticker=failure.attempted_value,
        error=failure.error,
    )


def _convert_stock_cell(
    sheet: Any,
    *,
    row_number: int,
    stock_cell_column: int,
    symbol_cell_column: int | None,
    ticker: str,
) -> tuple[bool, str | None]:
    """Convert one Column A cell to the Stocks data type. Return True on success.

    TODO: Excel offers no programmatic way to disambiguate an ambiguous ticker match,
    so we verify the result via the resolved "Stock Symbol" column and conservatively
    restore the plain ticker text when the entity did not resolve.
    """
    # TODO: cells are converted one at a time rather than as a single batched range.
    # This costs one COM round-trip per row but is what makes per-row verification and
    # fallback possible; revisit if ingest volume makes it slow.
    stock_cell = sheet.range((row_number, stock_cell_column))
    try:
        _call_with_com_retry(
            lambda: stock_cell.api.ConvertToLinkedDataType(
                ServiceID=STOCKS_SERVICE_ID,
                LanguageCulture=LINKED_DATA_TYPE_CULTURE,
            )
        )
    except Exception as exc:  # noqa: BLE001 - conversion is best-effort; ingest must not fail
        _restore_plain_ticker(stock_cell, ticker)
        return False, str(exc)

    if symbol_cell_column is None:
        # Without the verification column we cannot confirm the entity resolved, so
        # undo the conversion by rewriting plain text. This keeps the workbook state
        # consistent with the failure we report, and keeps Column A readable for dedup.
        _restore_plain_ticker(stock_cell, ticker)
        return False, "Workbook does not include a verification column for the stock symbol"

    resolved = _call_with_com_retry(lambda: sheet.range((row_number, symbol_cell_column)).value)
    if _is_resolved_symbol(resolved):
        return True, None

    _restore_plain_ticker(stock_cell, ticker)
    return False, "The stock symbol could not be resolved"


def _set_plain_ticker(sheet: Any, row_number: int, stock_cell_column: int, ticker: str) -> None:
    """Set the plain ticker text for a stock cell without attempting conversion."""
    stock_cell = sheet.range((row_number, stock_cell_column))
    try:
        _call_with_com_retry(lambda: setattr(stock_cell, "value", ticker))
    except Exception:  # noqa: BLE001 - best effort for prompt-retry flow
        pass


def _restore_plain_ticker(stock_cell: Any, ticker: str) -> None:
    """Put the plain ticker back so a failed row still carries usable data."""
    try:
        _call_with_com_retry(lambda: setattr(stock_cell, "value", ticker))
    except Exception:  # noqa: BLE001 - already on the failure path
        pass


def _is_resolved_symbol(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    # Excel surfaces unresolved linked data as #FIELD!/#VALUE!/#N/A error text.
    return not text.startswith("#")


def read_existing_lot_ids(workbook_path: Path, table_name: str) -> set[str]:
    """DEPRECATED: read existing composite dedup keys as pipe-delimited strings."""
    workbook, app, was_open = _open_workbook(workbook_path)
    try:
        table, _sheet = _find_table_and_sheet(workbook, table_name)
        headers = _table_headers(table)
        return _existing_dedup_keys(table, headers)
    finally:
        if not was_open:
            workbook.close()
            app.quit()


_LOT_ID_KEY_PREFIX = "lotid:"


def _make_dedup_key(trade: CanonicalTrade, headers: list[str]) -> str:
    """Build a dedup key for `trade`.

    When the workbook has a Lot ID column, `CanonicalTrade.lot_id` — the primary
    dedup discriminator — is used directly so that legitimate distinct lots (e.g.
    two same-day trades on the same contract/side/quantity) are never collapsed.
    Otherwise falls back to a composite key built from DEDUP_COLUMNS, for
    workbooks/rows written before the Lot ID column existed. Close events
    (trades carrying a Close Date/Exit Price — including a partial-close split
    merged with its open row's fields) additionally include close-side
    discriminators, so that two distinct close tickets against the same open
    lot (e.g. equal quantities closed on different dates) never collide.
    """
    if LOT_ID_COLUMN in headers and trade.lot_id:
        return f"{_LOT_ID_KEY_PREFIX}{trade.lot_id}"
    if trade.close_date is not None or trade.exit_price is not None:
        return _make_close_composite_dedup_key(trade)
    return _make_composite_dedup_key(trade)


def _make_composite_dedup_key(trade: CanonicalTrade) -> str:
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
        elif col_name == "Strike Price":
            parts.append(f"{trade.strike:g}" if trade.strike is not None else "")
    return "|".join(parts)


def _make_close_composite_dedup_key(trade: CanonicalTrade) -> str:
    """Build a composite dedup key for a close event.

    Extends `_make_composite_dedup_key` with close-side discriminators (Close
    Date, Exit Price) so the key can be reconstructed identically both from the
    incoming close ticket (merged with the open row's fields, before that row
    is mutated) and from the already-persisted closed/split row on a later run
    (see `_existing_dedup_keys`), instead of relying solely on open-side fields
    that two distinct close tickets against the same open lot can share.
    """
    base = _make_composite_dedup_key(trade)
    close_date = trade.close_date.isoformat() if trade.close_date else ""
    exit_price = f"{trade.exit_price:g}" if trade.exit_price is not None else ""
    return f"{base}|{close_date}|{exit_price}"


def _find_existing_row_to_update(
    table: Any, headers: list[str], trade: CanonicalTrade
) -> tuple[int, list[Any]] | None:
    """Return the (1-based table row index, row values) for a matching open row.

    Close-only trades can be reconciled to an existing open position in the table
    even when the incoming trade has no Open Date value of its own; in that case we
    fall back to matching on the row's stock/side/quantity/account values. A row
    whose open quantity is greater than the incoming close's quantity (e.g. two
    separate closing tickets against one larger open lot) is also considered a
    match — the caller is responsible for splitting the partial close.
    """
    if trade.close_date is None and trade.exit_price is None:
        return None

    data_range = getattr(table, "DataBodyRange", None)
    rows = _normalize_table_rows(data_range.Value, len(headers))
    if not rows:
        return None

    candidates: list[tuple[int, list[Any]]] = []
    col_indices = {col_name: headers.index(col_name) if col_name in headers else None for col_name in DEDUP_COLUMNS}
    symbol_index = headers.index(STOCK_SYMBOL_COLUMN) if STOCK_SYMBOL_COLUMN in headers else None
    account_index = headers.index("Account") if "Account" in headers else None

    for row_index, row in enumerate(rows):
        if _row_has_close_values(row, headers):
            continue

        if not _row_matches_trade(row, headers, trade, col_indices, symbol_index, account_index):
            continue

        candidates.append((row_index + 1, row))

    if not candidates:
        return None

    # `_row_matches_trade` never compares Exp Date/Call or Put, so distinct
    # option contracts (e.g. different expiries, or a put vs. a call) can both
    # reach this point sharing the same underlying/side/strike/quantity
    # profile. Narrow to rows whose option identifiers agree with the incoming
    # trade *before* considering exact-quantity/uniqueness, so a same-quantity
    # row for the wrong contract is never preferred over the correct contract
    # (which may hold a different, larger quantity).
    option_candidates = _filter_candidates_by_option_fields(candidates, headers, trade)

    # An exact-quantity match is unambiguous even when a larger, partially-open
    # row also satisfies the relaxed `row_quantity >= close_quantity` comparison
    # in `_row_matches_trade` (e.g. rows open for 50 and 100 shares both "match"
    # a 50-share close). Prefer the exact match so it closes outright instead of
    # being folded into a partial-close split.
    exact_candidates = [
        (idx, row) for idx, row in option_candidates if _is_exact_quantity_match(row, headers, trade.quantity)
    ]
    if exact_candidates:
        if len(exact_candidates) > 1:
            context = trade.trade_id or trade.lot_id or trade.symbol or "trade"
            raise ValueError(
                f"Multiple existing rows matched close trade '{context}'; cannot reconcile automatically"
            )
        return exact_candidates[0]

    if len(option_candidates) == 1:
        return option_candidates[0]

    if trade.exp_date is not None or trade.call_or_put:
        context = trade.trade_id or trade.lot_id or trade.symbol or "trade"
        raise ValueError(
            f"Multiple existing rows matched close trade '{context}'; cannot reconcile automatically"
        )

    # TODO: Several partially-open rows qualify for this close and none matches
    # exactly. Default to FIFO (earliest Open Date first, ties broken by table
    # row order) for review — the issue does not specify a required order when
    # more than one open lot could absorb the same close.
    return _fifo_earliest_candidate(option_candidates, headers)


def _filter_candidates_by_option_fields(
    candidates: list[tuple[int, list[Any]]], headers: list[str], trade: CanonicalTrade
) -> list[tuple[int, list[Any]]]:
    """Narrow `candidates` to rows whose Exp Date/Call or Put agree with `trade`.

    `_row_matches_trade` never compares these fields, so equity candidates are
    returned unchanged; only applies when the incoming trade actually carries
    option identifiers.
    """
    if trade.exp_date is None and not trade.call_or_put:
        return candidates

    exp_date_index = headers.index("Exp Date") if "Exp Date" in headers else None
    call_put_index = headers.index("Call or Put") if "Call or Put" in headers else None

    filtered: list[tuple[int, list[Any]]] = []
    for idx, row in candidates:
        if trade.exp_date is not None and exp_date_index is not None and exp_date_index < len(row):
            row_value = row[exp_date_index]
            if row_value not in (None, ""):
                if isinstance(row_value, (int, float)):
                    row_exp_date = _excel_serial_to_date(float(row_value))
                elif hasattr(row_value, "date") and callable(getattr(row_value, "date", None)):
                    row_exp_date = row_value.date()
                elif isinstance(row_value, date):
                    row_exp_date = row_value
                else:
                    row_exp_date = None
                if row_exp_date is not None and row_exp_date != trade.exp_date:
                    continue

        if trade.call_or_put and call_put_index is not None and call_put_index < len(row):
            row_value = row[call_put_index]
            if row_value not in (None, "") and str(row_value).strip() != trade.call_or_put.strip():
                continue

        filtered.append((idx, row))
    return filtered


def _row_quantity(row: list[Any], headers: list[str]) -> float | None:
    """Return the row's "C" (quantity) value, or None if unavailable/unreadable."""
    quantity_index = headers.index("C") if "C" in headers else None
    if quantity_index is None or quantity_index >= len(row):
        return None
    row_quantity_raw = row[quantity_index]
    if row_quantity_raw in (None, ""):
        return None
    try:
        return float(row_quantity_raw)
    except (TypeError, ValueError):
        return None


def _row_fees(row: list[Any], headers: list[str]) -> float | None:
    """Return the row's "Fees" value, or None if unavailable/unreadable."""
    fees_index = headers.index("Fees") if "Fees" in headers else None
    if fees_index is None or fees_index >= len(row):
        return None
    value = row[fees_index]
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _allocate_fee(total_fee: float | None, quantity_pool: float | None, quantity_slice: float) -> float | None:
    """Allocate a proportional share of `total_fee` for `quantity_slice` out of `quantity_pool`.

    Mirrors matcher._allocate_fee's proportional-by-quantity allocation so that
    splitting an existing open row's fees follows the same convention as the
    matcher's own partial-lot fee allocation.
    """
    if total_fee is None or quantity_pool is None:
        return None
    if quantity_pool <= MATCH_EPSILON:
        return None
    return total_fee * (quantity_slice / quantity_pool)


def _is_exact_quantity_match(row: list[Any], headers: list[str], close_quantity: float) -> bool:
    row_quantity = _row_quantity(row, headers)
    if row_quantity is None:
        return False
    return abs(row_quantity - close_quantity) <= MATCH_EPSILON


def _fifo_earliest_candidate(
    candidates: list[tuple[int, list[Any]]], headers: list[str]
) -> tuple[int, list[Any]]:
    """Return the candidate with the earliest Open Date (ties broken by row order)."""
    open_date_index = headers.index("Open Date") if "Open Date" in headers else None

    def _sort_key(candidate: tuple[int, list[Any]]) -> date:
        _, row = candidate
        if open_date_index is None or open_date_index >= len(row):
            return date.max
        value = row[open_date_index]
        if isinstance(value, (int, float)):
            parsed = _excel_serial_to_date(float(value))
        elif hasattr(value, "date") and callable(getattr(value, "date", None)):
            parsed = value.date()
        elif isinstance(value, date):
            parsed = value
        else:
            parsed = None
        return parsed if parsed is not None else date.max

    return min(candidates, key=_sort_key)


def _remaining_open_quantity(row: list[Any], headers: list[str], close_quantity: float) -> float | None:
    """Return the open quantity left on `row` after applying `close_quantity`.

    Returns None when the row's quantity is unavailable/unreadable (in which case
    the caller should treat the match as a full close, matching prior behavior),
    or when the close accounts for the entire open quantity (nothing remains open).
    """
    row_quantity = _row_quantity(row, headers)
    if row_quantity is None:
        return None

    remaining = row_quantity - close_quantity
    if remaining <= MATCH_EPSILON:
        return None
    return remaining


def _row_has_close_values(row: list[Any], headers: list[str]) -> bool:
    close_index = headers.index("Close Date") if "Close Date" in headers else None
    exit_index = headers.index("Exit Price") if "Exit Price" in headers else None
    if close_index is not None and close_index < len(row) and row[close_index] not in (None, ""):
        return True
    if exit_index is not None and exit_index < len(row) and row[exit_index] not in (None, ""):
        return True
    return False


def _row_matches_trade(
    row: list[Any],
    headers: list[str],
    trade: CanonicalTrade,
    col_indices: dict[str, int | None],
    symbol_index: int | None,
    account_index: int | None,
) -> bool:
    stock_value = _row_ticker(row, col_indices.get("Stock"), symbol_index)
    trade_stock_value = str(trade.stock or trade.underlying or "").strip()
    # An empty/unused table row (e.g. a trailing blank row left for future entries) has no
    # Stock value of its own, so it must never be treated as a match: without this check every
    # blank row would trivially satisfy the remaining comparisons below and be reported as an
    # ambiguous duplicate alongside the one real open row.
    if not stock_value:
        return False
    if trade_stock_value and stock_value != trade_stock_value:
        return False

    side_index = col_indices.get("B/S")
    if side_index is not None and side_index < len(row):
        row_side = row[side_index]
        if row_side not in (None, "") and str(row_side).strip() != str(trade.side or "").strip():
            return False

    quantity_index = col_indices.get("C")
    if quantity_index is not None and quantity_index < len(row):
        row_quantity = row[quantity_index]
        if row_quantity not in (None, ""):
            try:
                row_quantity_value = float(row_quantity)
            except (TypeError, ValueError):
                return False
            # A row can reconcile a close whose quantity is less than or equal to
            # the row's open quantity — e.g. two separate closing tickets against
            # one larger open lot. A close quantity greater than what's open on
            # the row can never be satisfied by it.
            if row_quantity_value - float(trade.quantity) < -MATCH_EPSILON:
                return False

    if account_index is not None and account_index < len(row):
        row_account = row[account_index]
        if row_account not in (None, "") and str(row_account).strip() != str(trade.account or "").strip():
            return False

    lot_id_index = headers.index(LOT_ID_COLUMN) if LOT_ID_COLUMN in headers else None
    if trade.open_date is not None and trade.lot_id and lot_id_index is not None and lot_id_index < len(row):
        row_lot_id = row[lot_id_index]
        if row_lot_id not in (None, "") and str(row_lot_id).strip() != trade.lot_id.strip():
            return False

    strike_index = col_indices.get("Strike Price")
    if strike_index is not None and strike_index < len(row) and trade.strike is not None:
        row_strike = row[strike_index]
        if row_strike not in (None, ""):
            try:
                row_strike_value = float(row_strike)
            except (TypeError, ValueError):
                return False
            if abs(row_strike_value - float(trade.strike)) > 1e-9:
                return False

    if trade.open_date is not None:
        open_date_index = headers.index("Open Date") if "Open Date" in headers else None
        if open_date_index is not None and open_date_index < len(row):
            row_open_date_value = row[open_date_index]
            if row_open_date_value in (None, ""):
                return False
            if isinstance(row_open_date_value, (int, float)):
                row_open_date = _excel_serial_to_date(float(row_open_date_value))
            elif hasattr(row_open_date_value, "date") and callable(getattr(row_open_date_value, "date", None)):
                row_open_date = row_open_date_value.date()
            else:
                row_open_date = None
            if row_open_date is None:
                return False
            if row_open_date != trade.open_date:
                return False

    return True


def _reduce_existing_row_quantity(
    sheet: Any,
    table: Any,
    headers: list[str],
    header_positions: dict[str, int],
    row_index: int,
    remaining_quantity: float,
    remaining_fees: float | None = None,
) -> None:
    """Shrink an existing open row's quantity (and, if provided, its Fees) to
    what's still open.

    Used when a close only accounts for part of an existing open row's
    quantity (e.g. two separate closing tickets against one larger open
    lot). The row is left open — no close-side columns are touched.
    `remaining_fees` is the row's own opening fees reallocated to the portion
    that stays open (see `_allocate_fee`); when None the Fees cell is left
    untouched.
    """
    quantity_column = header_positions.get("C")
    fees_column = header_positions.get("Fees")
    if quantity_column is None and fees_column is None:
        return

    list_row = _call_with_com_retry(lambda: table.ListRows(row_index))
    base_row = _call_with_com_retry(lambda: list_row.Range.Row)
    base_column = _call_with_com_retry(lambda: list_row.Range.Column)

    if quantity_column is not None:

        def _write_quantity() -> None:
            sheet.range((base_row, base_column + quantity_column - 1)).value = remaining_quantity

        _call_with_com_retry(_write_quantity)

    if fees_column is not None and remaining_fees is not None:

        def _write_fees() -> None:
            sheet.range((base_row, base_column + fees_column - 1)).value = remaining_fees

        _call_with_com_retry(_write_fees)


def _merge_open_fields_from_row(trade: CanonicalTrade, row: list[Any], headers: list[str]) -> CanonicalTrade:
    """Return a copy of `trade` enriched with open-side data from `row`.

    Used when a close-only trade (no matching open in the current import) is
    reconciled against part of an existing open row's quantity: the trade is
    written as its own new row, so it needs the original open-side fields
    (open date, strike, premium, etc.) carried over from the row it's
    partially closing, not just the incoming close's own fields.
    """
    date_fields = {"open_date", "exp_date"}
    float_fields = {"strike", "stock_price_open", "premium"}
    field_columns: tuple[tuple[str, str], ...] = (
        ("open_date", "Open Date"),
        ("exp_date", "Exp Date"),
        ("call_or_put", "Call or Put"),
        ("side", "B/S"),
        ("strike", "Strike Price"),
        ("stock_price_open", "Stock Price DOC"),
        ("premium", "Premium"),
        ("account", "Account"),
    )

    updates: dict[str, Any] = {}
    for field_name, col_name in field_columns:
        if col_name not in headers:
            continue
        col_index = headers.index(col_name)
        if col_index >= len(row):
            continue
        value = row[col_index]
        if value in (None, ""):
            continue
        if field_name in date_fields:
            if isinstance(value, (int, float)):
                value = _excel_serial_to_date(float(value))
            elif hasattr(value, "date") and callable(getattr(value, "date", None)):
                value = value.date()
            if value is None:
                continue
        elif field_name in float_fields:
            try:
                value = float(value)
            except (TypeError, ValueError):
                continue
        updates[field_name] = value

    if not updates:
        return trade

    merged = replace(trade, **updates)
    return replace(merged, trade_id=make_trade_id(merged))


def _update_existing_trade_row(
    sheet: Any,
    table: Any,
    headers: list[str],
    header_positions: dict[str, int],
    row_index: int,
    trade: CanonicalTrade,
) -> None:
    # NOTE: table.ListRows is a COM collection wrapped by xlwings'
    # NOTE: COMRetryObjectWrapper, which does not implement __getitem__, so
    # NOTE: bracket subscription (table.ListRows[row_index]) raises
    # NOTE: "'COMRetryObjectWrapper' object is not subscriptable". COM
    # NOTE: collections are indexed via call syntax instead.
    list_row = _call_with_com_retry(lambda: table.ListRows(row_index))
    base_row = _call_with_com_retry(lambda: list_row.Range.Row)
    base_column = _call_with_com_retry(lambda: list_row.Range.Column)

    for field_name, col_name in FIELD_TO_COLUMN.items():
        if col_name not in header_positions:
            continue
        if field_name == STOCK_FIELD_NAME:
            # NOTE: preserve the existing Excel Stocks linked-data cell instead of
            # NOTE: overwriting it with plain text during in-place close reconciliation.
            continue
        if field_name == "lot_id":
            # NOTE: preserve the row's original Lot ID during in-place close
            # NOTE: reconciliation — the incoming close trade may carry a
            # NOTE: different lot_id (its own close event's, for an orphan close
            # NOTE: matched to a previously-written open row) and must not
            # NOTE: overwrite the identity of the lot being closed.
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


def _existing_dedup_keys(table: Any, headers: list[str]) -> set[str]:
    """Read existing rows and build dedup keys.

    Rows with a Lot ID value are keyed by that value directly (the primary dedup
    discriminator); rows without one (e.g. written before the Lot ID column
    existed) fall back to the composite key built from DEDUP_COLUMNS. Rows that
    carry close info (Close Date/Exit Price) additionally get a close-side key
    with those discriminators appended, matching `_make_close_composite_dedup_key`,
    so a re-imported close ticket for an already-persisted closed/split row is
    recognized even though the row itself no longer qualifies as an open candidate.
    """
    values: set[str] = set()
    data_range = getattr(table, "DataBodyRange", None)
    if data_range is None or data_range.Value in (None, ""):
        return values

    rows = _normalize_table_rows(data_range.Value, len(headers))

    lot_id_index = headers.index(LOT_ID_COLUMN) if LOT_ID_COLUMN in headers else None
    close_date_index = headers.index("Close Date") if "Close Date" in headers else None
    exit_price_index = headers.index("Exit Price") if "Exit Price" in headers else None

    # Find column indices for dedup columns
    col_indices: dict[str, int | None] = {}
    for col_name in DEDUP_COLUMNS:
        col_indices[col_name] = headers.index(col_name) if col_name in headers else None
    # NOTE: once Column A holds a Stocks entity, reading it returns the entity's display
    # NOTE: name (e.g. "GraniteShares ETF Trust") rather than the ticker we wrote. The
    # NOTE: resolved "Stock Symbol" column is therefore the reliable dedup source.
    symbol_index = headers.index(STOCK_SYMBOL_COLUMN) if STOCK_SYMBOL_COLUMN in headers else None

    for row in rows:
        if lot_id_index is not None and lot_id_index < len(row) and row[lot_id_index] not in (None, ""):
            values.add(f"{_LOT_ID_KEY_PREFIX}{row[lot_id_index]}")
            continue

        parts: list[str] = []
        for col_name in DEDUP_COLUMNS:
            idx = col_indices[col_name]
            if col_name == "Stock":
                parts.append(_row_ticker(row, idx, symbol_index))
                continue
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
            elif col_name in ("C", "Strike Price"):
                parts.append(f"{float(val):g}")
            elif col_name == "Open Date" and hasattr(val, "date") and callable(getattr(val, "date", None)):
                parts.append(val.date().isoformat())
            else:
                parts.append(str(val))
        key = "|".join(parts)
        if not any(p for p in parts):
            continue

        close_date_val = (
            row[close_date_index] if close_date_index is not None and close_date_index < len(row) else None
        )
        exit_price_val = (
            row[exit_price_index] if exit_price_index is not None and exit_price_index < len(row) else None
        )
        if close_date_val not in (None, "") or exit_price_val not in (None, ""):
            close_date_str = _row_date_isoformat(close_date_val)
            exit_price_str = ""
            if exit_price_val not in (None, ""):
                try:
                    exit_price_str = f"{float(exit_price_val):g}"
                except (TypeError, ValueError):
                    exit_price_str = str(exit_price_val)
            values.add(f"{key}|{close_date_str}|{exit_price_str}")

            # A re-imported close ticket that never matches an open row (e.g.
            # because the row it originally reconciled is now fully closed and
            # therefore excluded as an open candidate) builds its key from its
            # own fields only, without the row's Open Date merged in. Also
            # index the same close discriminators keyed off a blank Open Date
            # so that lookup still finds this row.
            open_date_col_index = col_indices.get("Open Date")
            if open_date_col_index is not None:
                open_date_part_index = DEDUP_COLUMNS.index("Open Date")
                if parts[open_date_part_index]:
                    blank_open_parts = list(parts)
                    blank_open_parts[open_date_part_index] = ""
                    blank_open_key = "|".join(blank_open_parts)
                    values.add(f"{blank_open_key}|{close_date_str}|{exit_price_str}")
        else:
            values.add(key)

    return values


def _row_date_isoformat(value: Any) -> str:
    """Return the ISO date string for a raw cell `value`, or "" if unavailable."""
    if value in (None, ""):
        return ""
    if isinstance(value, (int, float)):
        parsed = _excel_serial_to_date(float(value))
        return parsed.isoformat() if parsed else ""
    if hasattr(value, "date") and callable(getattr(value, "date", None)):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def _row_ticker(row: list[Any], stock_index: int | None, symbol_index: int | None) -> str:
    """Return the ticker for a row, preferring the resolved "Stock Symbol" column.

    Column A may hold a Stocks entity whose value reads back as a company/fund name,
    so the formula-resolved symbol is used when available and falls back to Column A.
    """
    if symbol_index is not None and symbol_index < len(row):
        symbol_value = row[symbol_index]
        if _is_resolved_symbol(symbol_value):
            return str(symbol_value).strip()
    if stock_index is not None and stock_index < len(row):
        stock_value = row[stock_index]
        if stock_value not in (None, ""):
            return str(stock_value)
    return ""


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


def _find_table_and_sheet(workbook: Any, table_name: str) -> tuple[Any, Any]:
    """Find the named Excel table and its owning sheet in a single workbook scan.

    Looks up `ListObjects(table_name)` on each sheet and returns the table
    together with its owning sheet as soon as a match is found, so callers
    no longer need a second COM lookup (via a since-removed
    `_find_sheet_for_table`) to resolve the sheet.

    A sheet without the requested table can raise anything from a plain
    KeyError/ValueError (as raised by non-Excel/COM lookups) to a genuine
    `pywintypes.com_error` (e.g. "Subscript out of range") from real Excel.
    That `com_error` shape is detected explicitly by type rather than by
    matching its message text, which is not guaranteed to contain phrases
    like "not found" or "does not exist".
    """

    def _search() -> tuple[Any, Any]:
        available_sheets: list[str] = []
        for sheet in workbook.sheets:
            available_sheets.append(str(sheet.name))
            try:
                table = _call_with_com_retry(lambda sheet=sheet: sheet.api.ListObjects(table_name))
                return table, sheet
            except KeyError:
                continue
            except ValueError:
                continue
            except ComRetryExhaustedError:
                # Exhausted retries on a transient COM error — surface that clearly
                # rather than masking it as "table not found".
                raise
            except Exception as exc:
                if pywintypes is not None and isinstance(exc, pywintypes.com_error):
                    if _is_retryable_com_error(exc):
                        # _call_with_com_retry already retries transient COM
                        # errors internally, so reaching here with a still-
                        # retryable error means retries were somehow bypassed
                        # (e.g. future refactor) — surface it rather than
                        # silently treating a busy Excel as "table missing".
                        raise
                    # A missing keyed ListObjects lookup on this sheet; keep
                    # searching the rest of the workbook.
                    continue
                raise

        known = ", ".join(repr(name) for name in available_sheets) or "none"
        raise ValueError(f"Could not find table {table_name!r} in the workbook. Available sheets: {known}")

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
    if raw_value in (None, ""):
        return []
    if isinstance(raw_value, tuple):
        raw_value = [list(item) if isinstance(item, tuple) else item for item in raw_value]
    if isinstance(raw_value, list):
        if raw_value and not isinstance(raw_value[0], list):
            return [list(raw_value)]
        return [list(row) for row in raw_value]
    return [[raw_value] + [None] * (width - 1)]


def _row_is_populated(row: list[Any], headers: list[str]) -> bool:
    """Return True when `row` contains a user-entered value in a dedup column."""
    indices = [headers.index(col) for col in DEDUP_COLUMNS if col in headers]
    if indices:
        return any((row[idx] if idx < len(row) else None) not in (None, "") for idx in indices)
    return any(cell not in (None, "") for cell in row)


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
    return _last_populated_row_position_from_rows(rows, headers)


def _last_populated_row_position_from_rows(rows: list[list[Any]], headers: list[str]) -> int:
    """Return the one-based position of the last populated row in `rows`."""
    for position in range(len(rows), 0, -1):
        if _row_is_populated(rows[position - 1], headers):
            return position
    return 0


def _determine_insertion_position(rows: list[list[Any]], headers: list[str], trade: CanonicalTrade) -> int:
    """Return the 1-based position where `trade` should be inserted in `rows`."""
    if not rows:
        return 1

    last_populated_position = _last_populated_row_position_from_rows(rows, headers)
    if last_populated_position == 0:
        return 1

    ticker = str(trade.underlying or trade.stock or "").strip()
    stock_index = headers.index(FIELD_TO_COLUMN["stock"]) if FIELD_TO_COLUMN["stock"] in headers else None
    symbol_index = headers.index(STOCK_SYMBOL_COLUMN) if STOCK_SYMBOL_COLUMN in headers else None
    open_date_index = headers.index(FIELD_TO_COLUMN["open_date"]) if FIELD_TO_COLUMN["open_date"] in headers else None

    if trade.open_date is not None and ticker:
        same_group_position: int | None = None
        for row_position, row in enumerate(rows, start=1):
            if not _row_is_populated(row, headers):
                continue
            if _row_ticker(row, stock_index, symbol_index).strip() != ticker:
                continue
            row_open_date = _row_open_date(row, headers, open_date_index)
            if row_open_date == trade.open_date:
                same_group_position = row_position
        if same_group_position is not None:
            return same_group_position + 1

    if trade.open_date is not None and ticker:
        ticker_last_position: int | None = None
        ticker_first_later_position: int | None = None
        for row_position, row in enumerate(rows, start=1):
            if not _row_is_populated(row, headers):
                continue
            if _row_ticker(row, stock_index, symbol_index).strip() != ticker:
                continue
            row_open_date = _row_open_date(row, headers, open_date_index)
            if row_open_date is None:
                continue
            if row_open_date > trade.open_date:
                ticker_first_later_position = (
                    row_position
                    if ticker_first_later_position is None
                    else min(ticker_first_later_position, row_position)
                )
            if row_open_date <= trade.open_date:
                ticker_last_position = row_position
        if ticker_first_later_position is not None:
            return ticker_first_later_position
        if ticker_last_position is not None:
            return ticker_last_position + 1

    if trade.open_date is not None:
        for row_position, row in enumerate(rows, start=1):
            if not _row_is_populated(row, headers):
                continue
            row_open_date = _row_open_date(row, headers, open_date_index)
            if row_open_date is None:
                continue
            if row_open_date > trade.open_date:
                return row_position
        return last_populated_position + 1

    if ticker:
        ticker_last_position = None
        for row_position, row in enumerate(rows, start=1):
            if not _row_is_populated(row, headers):
                continue
            if _row_ticker(row, stock_index, symbol_index).strip() == ticker:
                ticker_last_position = row_position
        if ticker_last_position is not None:
            return ticker_last_position + 1

    return last_populated_position + 1


def _row_open_date(row: list[Any], headers: list[str], open_date_index: int | None = None) -> date | None:
    """Return the parsed Open Date from a table row, if one is present."""
    if open_date_index is None:
        open_date_index = headers.index(FIELD_TO_COLUMN["open_date"]) if FIELD_TO_COLUMN["open_date"] in headers else None
    if open_date_index is None or open_date_index >= len(row):
        return None

    raw_open_date = row[open_date_index]
    if raw_open_date in (None, ""):
        return None
    if isinstance(raw_open_date, (int, float)):
        return _excel_serial_to_date(float(raw_open_date))
    if hasattr(raw_open_date, "date") and callable(getattr(raw_open_date, "date", None)):
        return raw_open_date.date()
    if isinstance(raw_open_date, date):
        return raw_open_date
    return None


def _build_group_last_row_positions(
    table: Any, headers: list[str]
) -> tuple[dict[str, int], dict[tuple[str, date], int]]:
    """Return the last row position for each ticker, and each (ticker, Open Date) pair.

    Used to insert new rows for the same stock symbol/open date immediately after
    the last existing row of that group, so related trades stay grouped in
    consecutive rows (see `write_trades_detailed`). `ticker_last_row` supports
    grouping closing-only trades (no Open Date) by ticker alone.
    """
    data_range = getattr(table, "DataBodyRange", None)
    ticker_last_row: dict[str, int] = {}
    exact_group_last_row: dict[tuple[str, date], int] = {}
    if data_range is None:
        return ticker_last_row, exact_group_last_row

    rows = _normalize_table_rows(data_range.Value, len(headers))
    if not rows:
        return ticker_last_row, exact_group_last_row

    stock_index = headers.index(FIELD_TO_COLUMN["stock"]) if FIELD_TO_COLUMN["stock"] in headers else None
    symbol_index = headers.index(STOCK_SYMBOL_COLUMN) if STOCK_SYMBOL_COLUMN in headers else None

    for row_index, row in enumerate(rows, start=1):
        ticker = _row_ticker(row, stock_index, symbol_index).strip()
        if not ticker:
            continue
        ticker_last_row[ticker] = row_index

        row_open_date = _row_open_date(row, headers)
        if row_open_date is not None:
            exact_group_last_row[(ticker, row_open_date)] = row_index

    return ticker_last_row, exact_group_last_row
