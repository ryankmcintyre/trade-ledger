from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

from constants import TABLE_NAME
import trade_ingestion.writer as writer
from trade_ingestion.models import CanonicalTrade


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


class FakeRowRange:
    def __init__(self, table: "FakeTable") -> None:
        self.table = table
        self.values: dict[int, Any] = {}

    def Cells(self, _row_index: int, column_index: int) -> FakeCell:
        return FakeCell(self, column_index)


class FakeListRow:
    def __init__(self, table: "FakeTable") -> None:
        self.Range = FakeRowRange(table)
        table.added_rows.append(self.Range.values)


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


class FakeSheetApi:
    def __init__(self, table: FakeTable) -> None:
        self._table = table

    def ListObjects(self, name: str) -> FakeTable:
        if name != TABLE_NAME:
            raise KeyError(name)
        return self._table


class FakeSheet:
    def __init__(self, table: FakeTable) -> None:
        self.api = FakeSheetApi(table)


class FakeBooks(list):
    def open(self, fullname: str) -> "FakeBook":
        raise AssertionError(f"Unexpected open() for already-open workbook: {fullname}")


class FakeBook:
    def __init__(self, fullname: str, table: FakeTable, app: "FakeApp") -> None:
        self.fullname = fullname
        self.sheets = [FakeSheet(table)]
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
    written = writer.write_trades(workbook_path, [trade])

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
    written = writer.write_trades(workbook_path, [trade])

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
    written = writer.write_trades(workbook_path, [trade])

    assert written == 1
    row = table.added_rows[0]
    assert row[1] == "SOLS"  # Stock
    assert 3 not in row  # Open Date — None, not written
    assert 12 not in row  # Premium — None, not written
    assert row[13] == 1.0  # C (quantity)
    assert row[18] == 86.3  # Exit Price
    assert row[19] == date(2026, 5, 27)  # Close Date
