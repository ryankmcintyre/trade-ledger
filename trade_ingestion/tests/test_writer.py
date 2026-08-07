from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import pytest

from constants import TABLE_NAME
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


class FakeXlwingsCell:
    def __init__(self, row: "FakeRowRange", column_index: int) -> None:
        self._row = row
        self._column_index = column_index

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
    def __init__(self, table: "FakeTable") -> None:
        row_number = len(table.added_rows) + 2
        self.Range = FakeRowRange(table, row_number)
        table.added_rows.append(self.Range.values)
        table.rows_by_number[row_number] = self.Range


class FakeListRows:
    def __init__(self, table: "FakeTable") -> None:
        self.table = table

    def Add(self) -> FakeListRow:
        return FakeListRow(self.table)


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

    def range(self, coords: tuple[int, int]) -> FakeXlwingsCell:
        row_number, column_number = coords
        row = self._table.rows_by_number[row_number]
        relative_column = column_number - row.Column + 1
        return FakeXlwingsCell(row, relative_column)


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


def _trade(stock: str = "SPY", open_date: date | None = date(2024, 1, 2), quantity: float = 1.0, side: str = "B") -> CanonicalTrade:
    return CanonicalTrade(
        lot_id="lot-1",
        trade_id="",
        underlying="SPY",
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
