from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import pytest

from constants import STOCK_SYMBOL_COLUMN, STOCKS_SERVICE_ID, TABLE_NAME
import trade_ingestion.writer as writer
from trade_ingestion.models import CanonicalTrade

try:  # pragma: no cover - pywin32 is only importable on Windows
    import pywintypes
except ImportError:  # pragma: no cover - non-Windows/no-pywin32 environments

    class _FakeComError(Exception):
        """Stand-in for pywintypes.com_error when pywin32 isn't installed."""

    class _FakePywintypes:
        com_error = _FakeComError

    pywintypes = _FakePywintypes()  # type: ignore[assignment]

SHEET_NAME = "Trades"


class FakeCell:
    def __init__(self, row: "FakeRowRange", column_index: int) -> None:
        self._row = row
        self._column_index = column_index

    @property
    def Value(self) -> Any:
        return self._row.values.get(self._column_index)

    @Value.setter
    def Value(self, value: Any) -> None:
        self._row.values[self._column_index] = value


class FakeCellApi:
    """Stand-in for the COM Range behind an xlwings cell.

    Simulates Excel's Stocks linked data type conversion: on success the workbook's
    _FV formula in "Stock Symbol" resolves to the ticker; on failure it stays blank
    (or an error string), which is what the writer verifies against.
    """

    def __init__(self, sheet: "FakeSheet", row: "FakeRowRange", column_index: int) -> None:
        self._sheet = sheet
        self._row = row
        self._column_index = column_index

    def ConvertToLinkedDataType(self, ServiceID: int, LanguageCulture: str) -> None:
        self._sheet.conversion_calls.append((self._row.Row, self._column_index, ServiceID, LanguageCulture))
        if self._sheet.conversion_error is not None:
            raise self._sheet.conversion_error
        ticker = self._row.values.get(self._column_index)
        resolved = self._sheet.stock_resolver(str(ticker) if ticker is not None else "")
        symbol_index = self._sheet.symbol_column_index()
        if symbol_index is not None:
            self._row.values[symbol_index] = resolved


class FakeXlwingsCell:
    def __init__(self, row: "FakeRowRange", column_index: int, sheet: "FakeSheet | None" = None) -> None:
        self._row = row
        self._column_index = column_index
        self._sheet = sheet

    @property
    def api(self) -> FakeCellApi:
        assert self._sheet is not None
        return FakeCellApi(self._sheet, self._row, self._column_index)

    @property
    def value(self) -> Any:
        return self._row.values.get(self._column_index)

    @value.setter
    def value(self, value: Any) -> None:
        self._row.values[self._column_index] = value


class FakeRowRange:
    def __init__(self, table: "FakeTable", row_number: int) -> None:
        self.table = table
        self.Row = row_number
        self.Column = 1
        self.values: dict[int, Any] = {}

    def Cells(self, _row_index: int, column_index: int) -> FakeCell:
        return FakeCell(self, column_index)


class FakeListRow:
    def __init__(self, table: "FakeTable", position: int) -> None:
        rows = table.DataBodyRange.Value
        if rows is None:
            rows = []
            table.DataBodyRange.Value = rows
        if position < 1 or position > len(rows) + 1:
            raise ValueError(f"Invalid ListRow position: {position}")

        row_number = position + 1
        self.Range = FakeRowRange(table, row_number)
        rows.insert(position - 1, self.Range.values)
        table.added_rows.append(self.Range.values)
        for existing_row_number in sorted(table.rows_by_number, reverse=True):
            if existing_row_number >= row_number:
                existing_row = table.rows_by_number.pop(existing_row_number)
                existing_row.Row = existing_row_number + 1
                table.rows_by_number[existing_row_number + 1] = existing_row
        table.rows_by_number[row_number] = self.Range


class FakeListRows:
    def __init__(self, table: "FakeTable") -> None:
        self.table = table

    @property
    def Count(self) -> int:
        return len(self.table.DataBodyRange.Value or [])

    def Add(self, Position: int | None = None) -> FakeListRow:
        position = Position if Position is not None else self.Count + 1
        return FakeListRow(self.table, position)


class FakeRange:
    def __init__(self, value: Any) -> None:
        self.Value = value


class FakeTable:
    def __init__(self, headers: list[str], rows: list[list[Any]]) -> None:
        self.HeaderRowRange = FakeRange([headers])
        self.DataBodyRange = FakeRange(rows if rows else None)
        self.ListRows = FakeListRows(self)
        self.added_rows: list[dict[int, Any]] = []
        self.rows_by_number: dict[int, FakeRowRange] = {}


class FakeSheetApi:
    def __init__(self, table: FakeTable) -> None:
        self._table = table

    def ListObjects(self, name: str) -> FakeTable:
        if name != TABLE_NAME:
            raise KeyError(name)
        return self._table


class FakeSheet:
    def __init__(self, table: FakeTable, name: str = SHEET_NAME) -> None:
        self.api = FakeSheetApi(table)
        self.name = name
        self._table = table
        self.conversion_calls: list[tuple[int, int, int, str]] = []
        self.conversion_error: Exception | None = None
        # Default: every ticker resolves to itself, mirroring a working Stocks lookup.
        self.stock_resolver: Any = lambda ticker: ticker

    def symbol_column_index(self) -> int | None:
        headers = self._table.HeaderRowRange.Value[0]
        if STOCK_SYMBOL_COLUMN not in headers:
            return None
        return headers.index(STOCK_SYMBOL_COLUMN) + 1

    def range(self, coords: tuple[int, int]) -> FakeXlwingsCell:
        row_number, column_number = coords
        row = self._table.rows_by_number[row_number]
        relative_column = column_number - row.Column + 1
        return FakeXlwingsCell(row, relative_column, self)


class FakeBooks(list):
    def open(self, fullname: str) -> "FakeBook":
        raise AssertionError(f"Unexpected open() for already-open workbook: {fullname}")


class FakeBook:
    def __init__(self, fullname: str, table: FakeTable, app: "FakeApp", sheet_name: str = SHEET_NAME) -> None:
        self.fullname = fullname
        self.sheets = [FakeSheet(table, sheet_name)]
        self.app = app
        self.saved = False
        self.closed = False

    def save(self) -> None:
        self.saved = True

    def close(self) -> None:
        self.closed = True


class FakeApp:
    def __init__(self, books: list[FakeBook]) -> None:
        self.books = FakeBooks(books)
        self.quit_called = False

    def quit(self) -> None:
        self.quit_called = True


class FakeXw:
    def __init__(self, app: FakeApp) -> None:
        self.apps = [app]

    def App(self, *args: Any, **kwargs: Any) -> Any:
        raise AssertionError(f"Unexpected App() call for already-open workbook: {args} {kwargs}")


def _trade(
    stock: str = "SPY",
    open_date: date | None = date(2024, 1, 2),
    quantity: float = 1.0,
    side: str = "B",
    underlying: str | None = None,
) -> CanonicalTrade:
    return CanonicalTrade(
        lot_id="lot-1",
        trade_id="",
        underlying=underlying if underlying is not None else stock,
        symbol="SPY 240119C00450000",
        open_date=open_date,
        exp_date=date(2024, 1, 19),
        call_or_put="C",
        side=side,
        strike=450.0,
        stock_price_open=470.0,
        premium=2.0,
        quantity=quantity,
        fees=None,
        exit_price=3.0,
        close_date=date(2024, 1, 5),
        account="Fidelity",
        stock=stock,
    )


def test_write_trades_uses_column_mapping(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Stock Symbol", "Open Date", "Exp Date", "Call or Put", "B/S",
               "Stock Price DOC", "DTE", "Current Stock Price", "Break Even Price",
               "Strike Price", "Premium", "C", "Collateral", "(Put) Margin Cash Reserve",
               "(Call) Cost Basis/Share", "Fees", "Exit Price", "Close Date",
               "Profit/Loss", "Days Held", "Return on Capital",
               "Annualized ROR for Options", "Margin Annualized ROR",
               "Status", "Account", "Source"]
    table = FakeTable(headers, [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    trade = _trade()
    written = writer.write_trades(workbook_path, SHEET_NAME, [trade])

    assert written == 1
    assert book.saved is True
    assert len(table.added_rows) == 1

    row = table.added_rows[0]
    # Column positions (1-based): Stock=1, Open Date=3, Exp Date=4, C/P=5, B/S=6,
    # Strike Price=11, Premium=12, C=13, Exit Price=18, Close Date=19, Account=26
    assert row[1] == "SPY"  # Stock (Col A)
    assert row[3] == date(2024, 1, 2)  # Open Date
    assert row[4] == date(2024, 1, 19)  # Exp Date
    assert row[5] == "C"  # Call or Put
    assert row[6] == "B"  # B/S
    assert row[11] == 450.0  # Strike Price
    assert row[12] == 2.0  # Premium
    assert row[13] == 1.0  # C (quantity)
    assert row[18] == 3.0  # Exit Price
    assert row[19] == date(2024, 1, 5)  # Close Date
    assert row[26] == "Fidelity"  # Account
    # Formula columns should NOT be written
    assert 8 not in row  # DTE
    assert 9 not in row  # Current Stock Price
    assert 10 not in row  # Break Even Price


def test_write_trades_inserts_below_last_populated_row(
    monkeypatch: Any, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C"]
    existing_row = ["AAPL", date(2024, 1, 1), "B", 1.0]
    blank_row = [None, None, None, None]
    table = FakeTable(headers, [existing_row, blank_row.copy(), blank_row.copy()])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    written = writer.write_trades(workbook_path, SHEET_NAME, [_trade()])

    assert written == 1
    assert table.DataBodyRange.Value[0] == existing_row
    assert table.DataBodyRange.Value[1].get(1) == "SPY"
    assert table.DataBodyRange.Value[2:] == [blank_row, blank_row]


def test_write_trades_ignores_formula_only_trailing_rows(
    monkeypatch: Any, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C", "Status"]
    existing_row = ["AAPL", date(2024, 1, 1), "B", 1.0, "Open"]
    formula_only_row = [None, None, None, None, ""]
    table = FakeTable(headers, [existing_row, formula_only_row.copy()])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    written = writer.write_trades(workbook_path, SHEET_NAME, [_trade()])

    assert written == 1
    assert table.DataBodyRange.Value[0] == existing_row
    assert table.DataBodyRange.Value[1].get(1) == "SPY"
    assert table.DataBodyRange.Value[2] == formula_only_row


def test_write_trades_preserves_order_above_trailing_blank_rows(
    monkeypatch: Any, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C"]
    existing_row = ["AAPL", date(2024, 1, 1), "B", 1.0]
    blank_row = [None, None, None, None]
    table = FakeTable(headers, [existing_row, blank_row.copy()])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    trades = [_trade(stock="SPY"), _trade(stock="QQQ", quantity=2.0)]
    written = writer.write_trades(workbook_path, SHEET_NAME, trades)

    assert written == 2
    assert table.DataBodyRange.Value[0] == existing_row
    assert [row.get(1) for row in table.DataBodyRange.Value[1:3]] == ["SPY", "QQQ"]
    assert table.DataBodyRange.Value[3] == blank_row


def test_write_trades_inserts_at_start_of_blank_table(
    monkeypatch: Any, tmp_path: Path
) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C"]
    blank_row = [None, None, None, None]
    table = FakeTable(headers, [blank_row.copy(), blank_row.copy()])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    written = writer.write_trades(workbook_path, SHEET_NAME, [_trade()])

    assert written == 1
    assert table.DataBodyRange.Value[0][1] == "SPY"
    assert table.DataBodyRange.Value[1:] == [blank_row, blank_row]


def test_write_trades_dedup_by_composite_key(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Stock Symbol", "Open Date", "Exp Date", "Call or Put", "B/S",
               "Stock Price DOC", "DTE", "Current Stock Price", "Break Even Price",
               "Strike Price", "Premium", "C", "Collateral", "(Put) Margin Cash Reserve",
               "(Call) Cost Basis/Share", "Fees", "Exit Price", "Close Date",
               "Profit/Loss", "Days Held", "Return on Capital",
               "Annualized ROR for Options", "Margin Annualized ROR",
               "Status", "Account", "Source"]
    # Existing row: Stock=SPY, Open Date=2024-01-02 (serial 45293), B/S=B, C=1
    existing_rows = [["SPY", "SPY", 45293, None, "C", "B", None, None, None, None,
                      450.0, 2.0, 1.0, None, None, None, None, 3.0, 45296,
                      None, None, None, None, None, None, "Fidelity", None]]
    table = FakeTable(headers, existing_rows)
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    # Try to write the same trade — should be deduped
    trade = _trade()
    written = writer.write_trades(workbook_path, SHEET_NAME, [trade])

    assert written == 0
    assert len(table.added_rows) == 0


def test_write_trades_skips_none_values(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Stock Symbol", "Open Date", "Exp Date", "Call or Put", "B/S",
               "Stock Price DOC", "DTE", "Current Stock Price", "Break Even Price",
               "Strike Price", "Premium", "C", "Collateral", "(Put) Margin Cash Reserve",
               "(Call) Cost Basis/Share", "Fees", "Exit Price", "Close Date",
               "Profit/Loss", "Days Held", "Return on Capital",
               "Annualized ROR for Options", "Margin Annualized ROR",
               "Status", "Account", "Source"]
    table = FakeTable(headers, [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    # Orphan close — no open_date, no premium
    trade = CanonicalTrade(
        lot_id="orphan",
        trade_id="",
        underlying="SOLS",
        symbol="SOLS",
        open_date=None,
        exp_date=None,
        call_or_put=None,
        side="C",
        strike=None,
        stock_price_open=None,
        premium=None,
        quantity=1.0,
        fees=None,
        exit_price=86.3,
        close_date=date(2026, 5, 27),
        account="Fidelity",
        stock="SOLS",
    )
    written = writer.write_trades(workbook_path, SHEET_NAME, [trade])

    assert written == 1
    row = table.added_rows[0]
    assert row[1] == "SOLS"  # Stock
    assert 3 not in row  # Open Date — None, not written
    assert 12 not in row  # Premium — None, not written
    assert row[13] == 1.0  # C (quantity)
    assert row[18] == 86.3  # Exit Price
    assert row[19] == date(2026, 5, 27)  # Close Date


def test_write_trades_rejects_unknown_sheet(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    table = FakeTable(["Stock", "Open Date", "B/S", "C"], [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    with pytest.raises(ValueError, match="Could not find worksheet 'Missing'"):
        writer.write_trades(workbook_path, "Missing", [_trade()])

    assert len(table.added_rows) == 0


def test_write_trades_uses_only_the_named_sheet(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C"]
    other_table = FakeTable(headers, [])
    target_table = FakeTable(headers, [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), other_table, app, sheet_name="Other")
    book.sheets.append(FakeSheet(target_table, SHEET_NAME))
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    written = writer.write_trades(workbook_path, SHEET_NAME, [_trade()])

    assert written == 1
    assert len(target_table.added_rows) == 1
    assert len(other_table.added_rows) == 0


def test_write_trades_rejects_missing_table_on_sheet(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    class EmptySheetApi:
        def ListObjects(self, name: str) -> Any:
            raise KeyError(name)

    table = FakeTable(["Stock", "Open Date", "B/S", "C"], [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    book.sheets[0].api = EmptySheetApi()
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    with pytest.raises(ValueError, match=f"Could not find table '{TABLE_NAME}' on worksheet '{SHEET_NAME}'"):
        writer.write_trades(workbook_path, SHEET_NAME, [_trade()])

    assert len(table.added_rows) == 0


def _make_transient_com_error() -> Exception:
    """Build an exception shaped like the real RPC_E_CALL_REJECTED failure."""
    return pywintypes.com_error(-2147418111, "Call was rejected by callee.", None, None)


class FlakySheets:
    """Iterable that fails with a transient COM error the first N times."""

    def __init__(self, sheets: list[FakeSheet], fail_times: int) -> None:
        self._sheets = sheets
        self._remaining_failures = fail_times

    def __iter__(self) -> Any:
        if self._remaining_failures > 0:
            self._remaining_failures -= 1
            raise _make_transient_com_error()
        return iter(self._sheets)


class FlakyEnumerationSheets:
    """Iterable that raises the secondary TypeError symptom the first N times."""

    def __init__(self, sheets: list[FakeSheet], fail_times: int) -> None:
        self._sheets = sheets
        self._remaining_failures = fail_times

    def __iter__(self) -> Any:
        if self._remaining_failures > 0:
            self._remaining_failures -= 1
            raise TypeError("This object does not support enumeration")
        return iter(self._sheets)


def test_write_trades_retries_transient_com_error_on_sheet_lookup(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C"]
    table = FakeTable(headers, [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    # Fail twice with a transient RPC_E_CALL_REJECTED-style error, then succeed.
    book.sheets = FlakySheets([FakeSheet(table, SHEET_NAME)], fail_times=2)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))
    monkeypatch.setattr(writer, "pywintypes", pywintypes)
    monkeypatch.setattr(writer.time, "sleep", lambda _seconds: None)

    written = writer.write_trades(workbook_path, SHEET_NAME, [_trade()])

    assert written == 1
    assert len(table.added_rows) == 1


def test_write_trades_retries_enumeration_type_error_on_sheet_lookup(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C"]
    table = FakeTable(headers, [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    # Fail with the secondary "does not support enumeration" TypeError, then succeed.
    book.sheets = FlakyEnumerationSheets([FakeSheet(table, SHEET_NAME)], fail_times=1)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))
    monkeypatch.setattr(writer.time, "sleep", lambda _seconds: None)

    written = writer.write_trades(workbook_path, SHEET_NAME, [_trade()])

    assert written == 1
    assert len(table.added_rows) == 1


def test_write_trades_raises_clear_error_when_com_never_recovers(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["Stock", "Open Date", "B/S", "C"]
    table = FakeTable(headers, [])
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    # Always fails — retries should be exhausted and a clear error raised.
    book.sheets = FlakySheets([FakeSheet(table, SHEET_NAME)], fail_times=999)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))
    monkeypatch.setattr(writer, "pywintypes", pywintypes)
    monkeypatch.setattr(writer.time, "sleep", lambda _seconds: None)

    with pytest.raises(writer.ComRetryExhaustedError, match="Excel COM server was busy/unresponsive"):
        writer.write_trades(workbook_path, SHEET_NAME, [_trade()])



FULL_HEADERS = ["Stock", "Stock Symbol", "Open Date", "Exp Date", "Call or Put", "B/S",
                "Stock Price DOC", "DTE", "Current Stock Price", "Break Even Price",
                "Strike Price", "Premium", "C", "Collateral", "(Put) Margin Cash Reserve",
                "(Call) Cost Basis/Share", "Fees", "Exit Price", "Close Date",
                "Profit/Loss", "Days Held", "Return on Capital",
                "Annualized ROR for Options", "Margin Annualized ROR",
                "Status", "Account", "Source"]


def _build_workbook(tmp_path: Path, headers: list[str], rows: list[list[Any]]) -> tuple[Path, FakeTable, FakeBook, FakeApp]:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")
    table = FakeTable(headers, rows)
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)
    return workbook_path, table, book, app


def test_write_trades_converts_stock_cell_to_stocks_data_type(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path, table, book, app = _build_workbook(tmp_path, FULL_HEADERS, [])
    monkeypatch.setattr(writer, "xw", FakeXw(app))

    result = writer.write_trades_detailed(workbook_path, SHEET_NAME, [_trade(stock="NVDL")])

    assert result.rows_written == 1
    assert result.failed_conversions == []

    sheet = book.sheets[0]
    assert len(sheet.conversion_calls) == 1
    _row_number, column_index, service_id, culture = sheet.conversion_calls[0]
    assert column_index == 1  # Column A ("Stock")
    assert service_id == STOCKS_SERVICE_ID
    assert culture == "en-US"

    row = table.added_rows[0]
    assert row[1] == "NVDL"
    # The workbook's _FV formula resolves the ticker once the entity exists.
    assert row[FULL_HEADERS.index(STOCK_SYMBOL_COLUMN) + 1] == "NVDL"


def test_write_trades_reports_unresolved_conversion_and_keeps_plain_text(
    monkeypatch: Any, tmp_path: Path
) -> None:
    workbook_path, table, book, app = _build_workbook(tmp_path, FULL_HEADERS, [])
    monkeypatch.setattr(writer, "xw", FakeXw(app))
    # Excel accepts the call but cannot resolve the entity (offline / unknown ticker).
    book.sheets[0].stock_resolver = lambda ticker: "#FIELD!"

    result = writer.write_trades_detailed(workbook_path, SHEET_NAME, [_trade(stock="SPXW")])

    assert result.rows_written == 1
    assert result.failed_conversions == ["SPXW"]
    assert table.added_rows[0][1] == "SPXW"


def test_write_trades_survives_conversion_com_error(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path, table, book, app = _build_workbook(tmp_path, FULL_HEADERS, [])
    monkeypatch.setattr(writer, "xw", FakeXw(app))
    # Older Excel builds raise when the Stocks service is unavailable.
    book.sheets[0].conversion_error = RuntimeError("Stocks data type unavailable")

    result = writer.write_trades_detailed(workbook_path, SHEET_NAME, [_trade(stock="NVDL")])

    assert result.rows_written == 1
    assert result.failed_conversions == ["NVDL"]
    assert table.added_rows[0][1] == "NVDL"
    assert book.saved is True


def test_write_trades_dedups_using_resolved_stock_symbol_column(
    monkeypatch: Any, tmp_path: Path
) -> None:
    # Column A reads back as the Stocks entity display name, not the ticker.
    existing_rows = [["SPDR S&P 500 ETF Trust", "SPY", 45293, None, "C", "B", None, None,
                      None, None, 450.0, 2.0, 1.0, None, None, None, None, 3.0, 45296,
                      None, None, None, None, None, None, "Fidelity", None]]
    workbook_path, table, _book, app = _build_workbook(tmp_path, FULL_HEADERS, existing_rows)
    monkeypatch.setattr(writer, "xw", FakeXw(app))

    written = writer.write_trades(workbook_path, SHEET_NAME, [_trade(stock="SPY")])

    assert written == 0
    assert len(table.added_rows) == 0


def test_write_trades_second_run_writes_no_rows(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path, table, book, app = _build_workbook(tmp_path, FULL_HEADERS, [])
    monkeypatch.setattr(writer, "xw", FakeXw(app))

    trades = [_trade(stock="SPY")]
    assert writer.write_trades(workbook_path, SHEET_NAME, trades) == 1

    # Rebuild the workbook state as Excel would persist it: Column A now shows the
    # Stocks entity display name and Column B holds the resolved ticker.
    written_row = table.added_rows[0]
    persisted = [written_row.get(index + 1) for index in range(len(FULL_HEADERS))]
    persisted[0] = "SPDR S&P 500 ETF Trust"

    workbook_path2, table2, _book2, app2 = _build_workbook(tmp_path, FULL_HEADERS, [persisted])
    monkeypatch.setattr(writer, "xw", FakeXw(app2))

    assert writer.write_trades(workbook_path2, SHEET_NAME, trades) == 0
    assert len(table2.added_rows) == 0


def test_write_trades_reports_underlying_ticker_for_remapped_display_value(
    monkeypatch: Any, tmp_path: Path
) -> None:
    workbook_path, _table, book, app = _build_workbook(tmp_path, FULL_HEADERS, [])
    monkeypatch.setattr(writer, "xw", FakeXw(app))
    book.sheets[0].stock_resolver = lambda ticker: "#FIELD!"

    # Column A holds a remapped display value (UNDERLYING_DISPLAY_MAP), but the
    # warning should name the ticker the user actually imported.
    trade = _trade(stock="S&P 500 INDEX", underlying="SPXW")
    result = writer.write_trades_detailed(workbook_path, SHEET_NAME, [trade])

    assert result.failed_conversions == ["SPXW"]


def test_write_trades_restores_plain_text_without_verification_column(
    monkeypatch: Any, tmp_path: Path
) -> None:
    # A workbook without the "Stock Symbol" column offers no way to verify the
    # entity resolved, so the cell must be left as plain text rather than a
    # half-converted entity that later reads back as a display name.
    headers = ["Stock", "Open Date", "B/S", "C"]
    workbook_path, table, book, app = _build_workbook(tmp_path, headers, [])
    monkeypatch.setattr(writer, "xw", FakeXw(app))

    result = writer.write_trades_detailed(workbook_path, SHEET_NAME, [_trade(stock="SPY")])

    assert result.rows_written == 1
    assert result.failed_conversions == ["SPY"]
    assert len(book.sheets[0].conversion_calls) == 1
    assert table.added_rows[0][1] == "SPY"
