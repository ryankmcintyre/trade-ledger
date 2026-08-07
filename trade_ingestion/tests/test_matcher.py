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
    fees: float | None = None,
) -> RawEvent:
    return RawEvent(
        lot_id=lot_id,
        broker="Fidelity",
        account="Fidelity",
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
    )

    assert len(trades) == 1
    trade = trades[0]
    assert trade.lot_id == "open-1"
    assert trade.quantity == 1.0
    assert trade.exit_price == 3.0
    assert trade.close_date == date(2024, 1, 3)
    assert trade.fees == pytest.approx(0.3)
    assert trade.stock == "SPY"
    assert trade.status == "Closed"


def test_match_trades_partial_close_produces_matched_and_open_rows() -> None:
    trades = match_trades(
        [
            _event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=2.0, premium=2.0, fees=0.2),
            _event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.1),
        ],
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
    assert matched.status == "Closed"
    assert remaining.status == "Open"


def test_match_trades_returns_unmatched_open_position() -> None:
    trades = match_trades(
        [_event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=1.0, premium=2.0, fees=0.1)],
    )

    assert len(trades) == 1
    trade = trades[0]
    assert trade.close_date is None
    assert trade.exit_price is None
    assert trade.quantity == 1.0
    assert trade.status == "Open"


@pytest.mark.parametrize(
    ("effect", "expected_status"),
    [("CLOSE", "Closed"), ("OPEN", "Open"), ("ASSIGNED", "Assigned"), ("EXERCISED", "Exercised"), ("EXPIRED", "Expired")],
)
def test_match_trades_derives_status_from_effect(effect: str, expected_status: str) -> None:
    trades = match_trades([_event(lot_id="close-1", trade_date=date(2024, 1, 3), effect=effect, quantity=1.0, premium=3.0, fees=0.2)])

    assert len(trades) == 1
    assert trades[0].status == expected_status


def test_match_trades_orphan_close_creates_close_only_row() -> None:
    """Orphan closes (no matching open) should produce a close-only row."""
    trades = match_trades(
        [_event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.2)],
    )

    assert len(trades) == 1
    trade = trades[0]
    assert trade.open_date is None
    assert trade.premium is None
    assert trade.exit_price == 3.0
    assert trade.close_date == date(2024, 1, 3)
    assert trade.quantity == 1.0
    assert trade.stock == "SPY"


def test_pre_aggregation_merges_same_day_events() -> None:
    """Same-day, same-symbol, same-side, same-effect events are merged with weighted avg."""
    trades = match_trades(
        [
            _event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=4.0, premium=3.36),
            _event(lot_id="open-2", trade_date=date(2024, 1, 2), effect="OPEN", quantity=2.0, premium=3.15),
        ],
    )

    assert len(trades) == 1
    trade = trades[0]
    assert trade.quantity == 6.0
    # Weighted avg: (4*3.36 + 2*3.15) / 6 = 3.29
    assert trade.premium == pytest.approx(3.29, abs=0.01)
    assert trade.close_date is None


def test_match_trades_with_summary_reports_open_positions() -> None:
    result = match_trades_with_summary(
        [
            _event(lot_id="open-1", trade_date=date(2024, 1, 2), effect="OPEN", quantity=2.0, premium=2.0, fees=0.2),
            _event(lot_id="close-1", trade_date=date(2024, 1, 3), effect="CLOSE", quantity=1.0, premium=3.0, fees=0.1),
        ],
    )

    assert len(result.trades) == 2
    assert result.open_positions == 1
