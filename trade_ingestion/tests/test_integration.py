from collections import Counter
from datetime import date
from pathlib import Path

import pytest

from trade_ingestion.adapters.fidelity import parse_fidelity_csv_detailed
from trade_ingestion.matcher import match_trades_with_summary
from trade_ingestion.models import CanonicalTrade

FIXTURE_DIR = Path(__file__).parent / "fixtures"
FIDELITY_HISTORY_FIXTURE = FIXTURE_DIR / "History_for_Account_Zxxxxxxx.csv"


@pytest.mark.integration
@pytest.mark.parametrize("cover_quantity", [0, 4, 10, 15])
def test_fidelity_short_stock_uses_signed_common_stock_quantity(cover_quantity: int) -> None:
    content = "Date,Activity Description,Symbol,Quantity,Price,Amount\n"
    if cover_quantity:
        content += f'Oct-29-2026,YOU BOUGHT SHORT COVER,AMZN,{cover_quantity},229.08,"-2,290.80"\n'
    content += 'Oct-29-2026,YOU SOLD SHORT SALE,AMZN,-10,231.3,"2,312.95"\n'

    events = parse_fidelity_csv_detailed(content).events
    result = match_trades_with_summary(events)

    assert all(event.side == "S" and event.quantity > 0 for event in events)
    assert all(trade.side == "C" and trade.quantity < 0 for trade in result.trades)
    assert sum(trade.quantity for trade in result.trades) == pytest.approx(-max(10, cover_quantity) / 100)
    assert result.open_positions == int(cover_quantity < 10)
    for trade in result.trades:
        assert trade.call_or_put is None
        assert trade.exp_date is None
        if trade.open_date is not None:
            assert trade.open_date == date(2026, 10, 29)
            assert trade.premium == 231.3
        if trade.close_date is not None:
            assert trade.exit_price == 229.08
            assert trade.status == "Closed"
            if trade.premium is not None:
                assert (trade.exit_price - trade.premium) * trade.quantity * 100 > 0
        else:
            assert trade.exit_price is None
            assert trade.status == "Open"


@pytest.mark.integration
def test_fidelity_short_cover_does_not_close_long_stock() -> None:
    content = """Date,Activity Description,Symbol,Quantity,Price
Oct-29-2026,YOU BOUGHT,AMZN,10,230
Oct-29-2026,YOU SOLD SHORT SALE,AMZN,-10,231.3
Oct-29-2026,YOU BOUGHT SHORT COVER,AMZN,10,229.08
"""
    trades = match_trades_with_summary(parse_fidelity_csv_detailed(content).events).trades

    assert len(trades) == 2
    closed = next(trade for trade in trades if trade.close_date is not None)
    remaining = next(trade for trade in trades if trade.close_date is None)
    assert closed.side == remaining.side == "C"
    assert closed.quantity == -0.1
    assert closed.premium == 231.3
    assert remaining.quantity == 0.1
    assert remaining.premium == 230


def _row(trade: CanonicalTrade) -> tuple[object, ...]:
    return (
        trade.stock,
        trade.open_date,
        trade.exp_date,
        trade.call_or_put,
        trade.side,
        trade.strike,
        trade.premium,
        trade.quantity,
        trade.fees,
        trade.exit_price,
        trade.close_date,
        trade.account,
    )


@pytest.mark.integration
def test_fidelity_fixture_matches_expected_canonical_trades() -> None:
    """Fixture retains BOM and Fidelity disclaimer rows to exercise CSV parsing boundaries."""
    content = FIDELITY_HISTORY_FIXTURE.read_text(encoding="utf-8-sig")

    events = parse_fidelity_csv_detailed(content).events
    result = match_trades_with_summary(events)

    expected = [
        (
            "P",
            date(2026, 5, 27),
            None,
            None,
            "C",
            None,
            76.0,
            0.5,
            None,
            79.322,
            date(2026, 5, 27),
            "Fidelity",
        ),
        ("P", date(2026, 5, 27), None, None, "C", None, 76.0, 0.5, None, None, None, "Fidelity"),
        ("SOLS", None, None, None, "C", None, None, 1.0, None, 86.3, date(2026, 5, 27), "Fidelity"),
        (
            "NOK",
            date(2026, 5, 27),
            date(2026, 10, 16),
            "Call",
            "B",
            16.0,
            3.29,
            6.0,
            None,
            None,
            None,
            "Fidelity",
        ),
        (
            "S&P 500 INDEX",
            date(2026, 5, 27),
            date(2026, 6, 18),
            "Put",
            "S",
            7410.0,
            56.47,
            2.0,
            None,
            None,
            None,
            "Fidelity",
        ),
        (
            "S&P 500 INDEX",
            date(2026, 5, 27),
            date(2026, 6, 18),
            "Put",
            "B",
            7400.0,
            54.07,
            2.0,
            None,
            None,
            None,
            "Fidelity",
        ),
    ]

    assert len(result.trades) == 6
    assert result.skipped_duplicates == 0
    assert result.open_positions == 4
    assert Counter(_row(trade) for trade in result.trades) == Counter(expected)

    nok_trade = next(trade for trade in result.trades if trade.underlying == "NOK")
    assert nok_trade.stock == "NOK"
    assert nok_trade.underlying == "NOK"
    assert nok_trade.symbol == "NOK 261016C00016000"

    spxw_trades = [trade for trade in result.trades if trade.underlying == "SPXW"]
    assert len(spxw_trades) == 2
    assert {trade.stock for trade in spxw_trades} == {"S&P 500 INDEX"}
    assert {trade.underlying for trade in spxw_trades} == {"SPXW"}
    assert {trade.symbol for trade in spxw_trades} == {
        "SPXW 260618P07400000",
        "SPXW 260618P07410000",
    }
