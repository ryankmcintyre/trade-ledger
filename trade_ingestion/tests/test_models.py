from datetime import date

from trade_ingestion.models import CanonicalTrade, make_trade_id


def test_make_trade_id_for_option_trade() -> None:
    trade = CanonicalTrade(
        lot_id="lot-1",
        trade_id="",
        underlying="SPY",
        symbol="SPY 230915C00450000",
        open_date=date(2023, 9, 1),
        exp_date=date(2023, 9, 15),
        call_or_put="C",
        side="B",
        strike=450.0,
        stock_price_open=449.5,
        premium=2.15,
        quantity=1.0,
        fees=1.3,
        exit_price=3.0,
        close_date=date(2023, 9, 8),
        account="Fidelity:IRA",
    )

    assert make_trade_id(trade) == "2023-09-01|SPY 230915C00450000|B|1"
    assert trade.trade_id == "2023-09-01|SPY 230915C00450000|B|1"


def test_make_trade_id_for_equity_trade() -> None:
    trade = CanonicalTrade(
        lot_id="lot-2",
        trade_id="",
        underlying="AAPL",
        symbol="AAPL",
        open_date=date(2023, 8, 1),
        exp_date=None,
        call_or_put=None,
        side="C",
        strike=None,
        stock_price_open=190.5,
        premium=190.5,
        quantity=1.0,
        fees=0.65,
        exit_price=193.0,
        close_date=date(2023, 8, 2),
        account="Fidelity:IRA",
    )

    assert make_trade_id(trade) == "2023-08-01|AAPL|C|1"


def test_make_trade_id_with_optional_fields_missing() -> None:
    trade = CanonicalTrade(
        lot_id="lot-3",
        trade_id="",
        underlying="MSFT",
        symbol="MSFT",
        open_date=date(2023, 7, 1),
        exp_date=None,
        call_or_put=None,
        side="C",
        strike=None,
        stock_price_open=None,
        premium=None,
        quantity=2.5,
        fees=0.0,
        exit_price=None,
        close_date=None,
        account="Fidelity:TAXABLE",
    )

    assert make_trade_id(trade) == "2023-07-01|MSFT|C|2.5"
