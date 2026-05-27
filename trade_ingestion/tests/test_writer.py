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


def _trade(lot_id: str, trade_id: str) -> CanonicalTrade:
    return CanonicalTrade(
        lot_id=lot_id,
        trade_id=trade_id,
        underlying="SPY",
        symbol="SPY 240119C00450000",
        open_date=date(2024, 1, 2),
        exp_date=date(2024, 1, 19),
        call_or_put="C",
        side="B",
        strike=450.0,
        stock_price_open=470.0,
        premium=2.0,
        quantity=1.0,
        fees=0.3,
        exit_price=3.0,
        close_date=date(2024, 1, 5),
        account="Fidelity:IRA-1",
    )


def test_write_trades_writes_only_raw_columns_and_skips_duplicates(monkeypatch: Any, tmp_path: Path) -> None:
    workbook_path = tmp_path / "ledger.xlsx"
    workbook_path.write_text("placeholder", encoding="utf-8")

    headers = ["lot_id", "trade_id", "underlying", "symbol", "formula_col", "open_date", "account"]
    existing_rows = [["existing-lot", "existing-trade", "SPY", "SPY 240119C00450000", "=SUM(A1:A1)", date(2024, 1, 1), "Fidelity:IRA-1"]]
    table = FakeTable(headers, existing_rows)
    app = FakeApp([])
    book = FakeBook(str(workbook_path.resolve()), table, app)
    app.books.append(book)

    monkeypatch.setattr(writer, "xw", FakeXw(app))

    written = writer.write_trades(
        workbook_path,
        [
            _trade("existing-lot", "new-trade"),
            _trade("new-lot", "existing-trade"),
            _trade("fresh-lot", "fresh-trade"),
        ],
    )

    assert written == 1
    assert book.saved is True
    assert book.closed is False
    assert app.quit_called is False
    assert len(table.added_rows) == 1

    row = table.added_rows[0]
    assert row[1] == "fresh-lot"
    assert row[2] == "fresh-trade"
    assert row[3] == "SPY"
    assert row[4] == "SPY 240119C00450000"
    assert 5 not in row
    assert row[6] == date(2024, 1, 2)
    assert row[7] == "Fidelity:IRA-1"
