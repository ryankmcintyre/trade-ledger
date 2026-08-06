from datetime import date

import pytest

from trade_ingestion.adapters.robinhood import parse_robinhood_csv
from trade_ingestion.models import make_fallback_lot_id


def test_parse_robinhood_csv_parses_equities_and_options_and_skips_non_trades() -> None:
    content = """Metadata,Value
Generated,2024-01-01

Date,Description,Symbol,Quantity,Price,Commission,Fees,Account,Transaction ID,Security Type
2024-01-02,Buy,AAPL,100,180.50,1.00,0.50,IRA-1,EQ-BUY,Equity
2024-01-03,Sell,AAPL,100,181.75,1.00,0.50,IRA-1,EQ-SELL,Equity
2024-01-04,Buy to Open,SPY 2024-01-19 450 Call,1,2.15,0.65,0.05,IRA-1,OPT-BUY,Option
2024-01-05,Sell to Close,SPY 2024-01-19 450 Call,1,3.10,0.65,0.05,IRA-1,OPT-SELL,Option
2024-01-06,Dividend,AAPL,0,0,0,0,IRA-1,DIV-1,Cash
"""

    events = parse_robinhood_csv(content)

    assert len(events) == 4
    equity_buy, equity_sell, option_buy, option_sell = events

    assert equity_buy.effect == "OPEN"
    assert equity_buy.side == "C"
    assert equity_buy.quantity == pytest.approx(1.0)
    assert equity_buy.fees == pytest.approx(1.5)

    assert equity_sell.effect == "CLOSE"
    assert equity_sell.side == "C"

    assert option_buy.symbol == "SPY 240119C00450000"
    assert option_buy.underlying == "SPY"
    assert option_buy.exp_date == date(2024, 1, 19)
    assert option_buy.call_or_put == "C"
    assert option_buy.strike == pytest.approx(450.0)
    assert option_buy.effect == "OPEN"
    assert option_buy.side == "B"

    assert option_sell.effect == "CLOSE"
    assert option_sell.side == "S"


def test_parse_robinhood_csv_uses_hash_fallback_when_transaction_id_missing() -> None:
    content = """Date,Description,Symbol,Quantity,Price,Commission,Fees,Account
2024-01-02,Buy,AAPL,100,180.50,1.00,0.50,IRA-1
"""

    events = parse_robinhood_csv(content)

    assert len(events) == 1
    event = events[0]
    expected_lot_id = make_fallback_lot_id(
        trade_date=event.trade_date,
        symbol=event.symbol,
        quantity=event.quantity,
        premium=event.premium,
    )
    assert event.lot_id == expected_lot_id
    assert len(event.lot_id) == 12
