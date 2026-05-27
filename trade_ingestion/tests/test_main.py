from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

import main
from trade_ingestion.models import CanonicalTrade, RawEvent


def test_run_pipeline_uses_broker_adapter_and_writer(monkeypatch: Any, tmp_path: Path) -> None:
    csv_path = tmp_path / "input.csv"
    workbook_path = tmp_path / "ledger.xlsx"
    csv_path.write_text("example", encoding="utf-8")
    workbook_path.write_text("placeholder", encoding="utf-8")

    captured: dict[str, Any] = {}
    events = [
        RawEvent(
            lot_id="lot-1",
            broker="Fidelity",
            account="Fidelity:IRA",
            underlying="AAPL",
            symbol="AAPL",
            trade_date=__import__("datetime").date(2024, 1, 2),
            exp_date=None,
            call_or_put=None,
            side="C",
            strike=None,
            stock_price=180.0,
            premium=180.0,
            quantity=1.0,
            fees=0.0,
            effect="OPEN",
        )
    ]
    trades = [
        CanonicalTrade(
            lot_id="lot-1",
            trade_id="",
            underlying="AAPL",
            symbol="AAPL",
            open_date=__import__("datetime").date(2024, 1, 2),
            exp_date=None,
            call_or_put=None,
            side="C",
            strike=None,
            stock_price_open=180.0,
            premium=180.0,
            quantity=1.0,
            fees=0.0,
            exit_price=None,
            close_date=None,
            account="Fidelity:IRA",
        )
    ]

    def fake_adapter(content: str) -> list[RawEvent]:
        captured["content"] = content
        return events

    def fake_read_existing_lot_ids(path: Path) -> set[str]:
        captured["dedup_path"] = path
        return {"existing-lot"}

    def fake_match_trades(input_events: list[RawEvent], existing_lot_ids: set[str]) -> list[CanonicalTrade]:
        captured["events"] = input_events
        captured["existing_lot_ids"] = existing_lot_ids
        return trades

    def fake_write_trades(path: Path, input_trades: list[CanonicalTrade]) -> int:
        captured["write_path"] = path
        captured["trades"] = input_trades
        return len(input_trades)

    monkeypatch.setitem(main.ADAPTERS, "fake", fake_adapter)
    monkeypatch.setattr(main, "read_existing_lot_ids", fake_read_existing_lot_ids)
    monkeypatch.setattr(main, "match_trades", fake_match_trades)
    monkeypatch.setattr(main, "write_trades", fake_write_trades)

    written = main.run_pipeline(broker="fake", csv_path=csv_path, workbook_path=workbook_path)

    assert written == 1
    assert captured["content"] == "example"
    assert captured["dedup_path"] == workbook_path
    assert captured["events"] == events
    assert captured["existing_lot_ids"] == {"existing-lot"}
    assert captured["write_path"] == workbook_path
    assert captured["trades"] == trades


def test_run_pipeline_rejects_unsupported_broker(tmp_path: Path) -> None:
    csv_path = tmp_path / "input.csv"
    workbook_path = tmp_path / "ledger.xlsx"
    csv_path.write_text("example", encoding="utf-8")
    workbook_path.write_text("placeholder", encoding="utf-8")

    with pytest.raises(ValueError, match="Unsupported broker"):
        main.run_pipeline(broker="unknown", csv_path=csv_path, workbook_path=workbook_path)


def test_main_parses_cli_arguments(monkeypatch: Any, capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    csv_path = tmp_path / "input.csv"
    workbook_path = tmp_path / "ledger.xlsx"

    def fake_run_pipeline(*, broker: str, csv_path: Path, workbook_path: Path) -> int:
        assert broker == "fidelity"
        assert csv_path == tmp_path / "input.csv"
        assert workbook_path == tmp_path / "ledger.xlsx"
        return 3

    monkeypatch.setattr(main, "run_pipeline", fake_run_pipeline)

    exit_code = main.main(["fidelity", str(csv_path), "--workbook", str(workbook_path)])

    assert exit_code == 0
    assert capsys.readouterr().out.strip() == f"Wrote 3 trade rows to {workbook_path}"
