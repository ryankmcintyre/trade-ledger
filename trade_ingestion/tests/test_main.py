from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

import main
from trade_ingestion.matcher import MatchResult
from trade_ingestion.models import CanonicalTrade, RawEvent
from trade_ingestion.writer import WriteResult


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
            account="Fidelity",
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
            fees=None,
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
            fees=None,
            exit_price=None,
            close_date=None,
            account="Fidelity",
            stock="AAPL",
        )
    ]

    def fake_adapter(content: str) -> list[RawEvent]:
        captured["content"] = content
        return events

    def fake_match_trades(input_events: list[RawEvent], existing_lot_ids: set[str] | None = None) -> MatchResult:
        captured["events"] = input_events
        return MatchResult(trades=trades, skipped_duplicates=0, open_positions=1)

    def fake_write_trades(path: Path, sheet_name: str, input_trades: list[CanonicalTrade]) -> WriteResult:
        captured["write_path"] = path
        captured["sheet_name"] = sheet_name
        captured["trades"] = input_trades
        return WriteResult(rows_written=len(input_trades), failed_conversions=[])

    monkeypatch.setitem(main.ADAPTERS, "fake", fake_adapter)
    monkeypatch.setattr(main, "match_trades_with_summary", fake_match_trades)
    monkeypatch.setattr(main, "write_trades_detailed", fake_write_trades)

    result = main.run_pipeline(
        broker="fake", csv_path=csv_path, workbook_path=workbook_path, sheet_name="Trades"
    )

    assert result == main.PipelineResult(rows_ingested=1, rows_skipped=0, open_positions=1)
    assert captured["content"] == "example"
    assert captured["events"] == events
    assert captured["write_path"] == workbook_path
    assert captured["sheet_name"] == "Trades"
    assert captured["trades"] == trades


def test_run_pipeline_rejects_unsupported_broker(tmp_path: Path) -> None:
    csv_path = tmp_path / "input.csv"
    workbook_path = tmp_path / "ledger.xlsx"
    csv_path.write_text("example", encoding="utf-8")
    workbook_path.write_text("placeholder", encoding="utf-8")

    with pytest.raises(ValueError, match="Unsupported broker"):
        main.run_pipeline(
            broker="unknown", csv_path=csv_path, workbook_path=workbook_path, sheet_name="Trades"
        )


def test_main_parses_cli_arguments(monkeypatch: Any, capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    csv_path = tmp_path / "input.csv"
    workbook_path = tmp_path / "ledger.xlsx"

    def fake_run_pipeline(
        *, broker: str, csv_path: Path, workbook_path: Path, sheet_name: str
    ) -> main.PipelineResult:
        assert broker == "fidelity"
        assert csv_path == tmp_path / "input.csv"
        assert workbook_path == tmp_path / "ledger.xlsx"
        assert sheet_name == "Trades"
        return main.PipelineResult(rows_ingested=3, rows_skipped=2, open_positions=1)

    monkeypatch.setattr(main, "run_pipeline", fake_run_pipeline)

    exit_code = main.main(
        ["fidelity", str(csv_path), "--workbook", str(workbook_path), "--sheet", "Trades"]
    )

    assert exit_code == 0
    assert capsys.readouterr().out.strip() == (
        f"Ingested 3 trade rows to {workbook_path} [Trades]; "
        "skipped 2 duplicate rows; "
        "left 1 open position unmatched"
    )


def test_main_requires_sheet_argument(tmp_path: Path) -> None:
    csv_path = tmp_path / "input.csv"
    workbook_path = tmp_path / "ledger.xlsx"

    with pytest.raises(SystemExit):
        main.main(["fidelity", str(csv_path), "--workbook", str(workbook_path)])
