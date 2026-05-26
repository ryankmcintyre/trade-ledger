from datetime import date

from trade_ingestion.adapters.fidelity import parse_fidelity_csv


def test_parse_fidelity_csv_skips_metadata_and_non_trade_rows() -> None:
    content = """Metadata,Value
Generated,2024-01-01
Trade Date,Action,Symbol,Quantity,Price,Commission,Fees,Account,Transaction ID,Security Type,Underlying Price
2024-01-02,Buy,AAPL,100,180.50,1.00,0.50,IRA-1,EQ-BUY,Equity,180.50
2024-01-03,Sell,AAPL,100,181.75,1.00,0.50,IRA-1,EQ-SELL,Equity,181.75
2024-01-04,Buy to Open,SPY 01/19/2024 450 C,1,2.15,0.65,0.05,IRA-1,OPT-BUY,Option,470.00
2024-01-05,Sell to Close,SPY 01/19/2024 450 C,1,3.10,0.65,0.05,IRA-1,OPT-SELL,Option,472.00
2024-01-06,Dividend,AAPL,0,0,0,0,IRA-1,DIV-1,Cash,0
"""

    events = parse_fidelity_csv(content)

    assert len(events) == 4

    equity_buy, equity_sell, option_buy, option_sell = events

    assert equity_buy.lot_id == "EQ-BUY"
    assert equity_buy.account == "Fidelity:IRA-1"
    assert equity_buy.side == "C"
    assert equity_buy.effect == "OPEN"
    assert equity_buy.quantity == 1.0
    assert equity_buy.fees == 1.5
    assert equity_buy.trade_date == date(2024, 1, 2)

    assert equity_sell.side == "C"
    assert equity_sell.effect == "CLOSE"
    assert equity_sell.quantity == 1.0

    assert option_buy.symbol == "SPY 240119C00450000"
    assert option_buy.underlying == "SPY"
    assert option_buy.exp_date == date(2024, 1, 19)
    assert option_buy.call_or_put == "C"
    assert option_buy.strike == 450.0
    assert option_buy.side == "B"
    assert option_buy.effect == "OPEN"

    assert option_sell.symbol == "SPY 240119C00450000"
    assert option_sell.side == "B"
    assert option_sell.effect == "CLOSE"
