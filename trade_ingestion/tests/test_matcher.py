from datetime import date

import pytest

from trade_ingestion.matcher import match_trades, match_trades_with_summary
from trade_ingestion.models import RawEvent


def _event(
    *,
    lot_id: str,
    trade_date: date,
    effect: str,
    quantity: float,
    premium: float,
    fees: float,
) -> RawEvent:
    return RawEvent(
        lot_id=lot_id,
        broker="Fidelity",
        account="Fidelity:IRA-1",
        underlying="SPY",
        symbol="SPY 240119C00450000",
        trade_date=trade_date,
        exp_date=date(2024, 1, 19),
        call_or_put="C",
        side="B",
        strike=450.0,
        stock_price=470.0,
        premium=premium,
        quantity=quantity,
        fees=fees,
        effect=effect,
    )


def test_match_trades_full_close() -> None:
    trades = match_trades(
        [
            _event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=1.0, premium=2.0, fees=0.1),
            _event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.2),
        ],
        existing_lot_ids=set(),
    )

    assert len(trades) == 1
    trade = trades[0]
    assert trade.lot_id == "open-1"
    assert trade.quantity == 1.0
    assert trade.exit_price == 3.0
    assert trade.close_date == date(2024, 1, 3)
    assert trade.fees == pytest.approx(0.3)


def test_match_trades_partial_close_produces_matched_and_open_rows() -> None:
    trades = match_trades(
        [
            _event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=2.0, premium=2.0, fees=0.2),
            _event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.1),
        ],
        existing_lot_ids=set(),
    )

    assert len(trades) == 2
    matched = next(trade for trade in trades if trade.close_date is not None)
    remaining = next(trade for trade in trades if trade.close_date is None)

    assert matched.lot_id == "open-1"
    assert matched.quantity == 1.0
    assert matched.fees == pytest.approx(0.2)
    assert remaining.lot_id.startswith("open-1:open-")
    assert remaining.quantity == 1.0
    assert remaining.fees == pytest.approx(0.1)
    assert remaining.exit_price is None


def test_match_trades_returns_unmatched_open_position() -> None:
    trades = match_trades(
        [_event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=1.0, premium=2.0, fees=0.1)],
        existing_lot_ids=set(),
    )

    assert len(trades) == 1
    trade = trades[0]
    assert trade.close_date is None
    assert trade.exit_price is None
    assert trade.quantity == 1.0


def test_match_trades_skips_duplicates_by_lot_id() -> None:
    trades = match_trades(
        [
            _event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=1.0, premium=2.0, fees=0.1),
            _event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.2),
        ],
        existing_lot_ids={"open-1"},
    )

    assert trades == []


def test_match_trades_warns_and_skips_orphaned_close() -> None:
    with pytest.warns(UserWarning, match="could not be matched to an open lot and was skipped"):
        trades = match_trades(
            [_event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.2)],
            existing_lot_ids=set(),
        )

    assert trades == []


def test_match_trades_with_summary_reports_duplicates_and_open_positions() -> None:
    result = match_trades_with_summary(
        [
            _event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=2.0, premium=2.0, fees=0.2),
            _event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.1),
        ],
        existing_lot_ids={"open-1"},
    )

    assert len(result.trades) == 1
    assert result.trades[0].lot_id == "open-1:open-1"
    assert result.trades[0].close_date is None
    assert result.skipped_duplicates == 1
    assert result.open_positions == 1
