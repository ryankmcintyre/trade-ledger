from datetime import date

import pytest

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


def test_parse_fidelity_csv_real_world_column_names_and_verbose_actions() -> None:
    """Exercises the actual column headers and action strings from a real Fidelity export."""
    content = (
        "\n"
        "Run Date,Action,Symbol,Description,Type,Price ($),Quantity,Commission ($),Fees ($),"
        "Accrued Interest ($),Amount ($),Cash Balance ($),Settlement Date\n"
        '05/27/2026,"YOU BOUGHT CLOUDFLARE INC CL A COM (NET) (Margin)",NET,'
        '"CLOUDFLARE INC CL A COM",Margin,205,30,,,,-6150,,05/28/2026\n'
        '05/27/2026,"YOU SOLD SOLSTICE ADVANCED MATLS INC COM SHS (SOLS) (Margin)",SOLS,'
        '"SOLSTICE ADVANCED MATLS INC COM SHS",Margin,86.3,-100,,0.19,,8629.81,,05/28/2026\n'
        '05/27/2026,"YOU BOUGHT OPENING TRANSACTION PUT (SPXW) NEW S & P 500 INDEX JUN 18 26 '
        '$7400 (100 SHS) (Margin)"," -SPXW260618P7400","PUT (SPXW)...",Margin,54.07,2,,0.05,,'
        "-10814.05,,05/28/2026\n"
        '05/27/2026,"YOU SOLD OPENING TRANSACTION PUT (SPXW) NEW S & P 500 INDEX JUN 18 26 '
        '$7410 (100 SHS) (Margin)"," -SPXW260618P7410","PUT (SPXW)...",Margin,56.47,-2,,0.05,,'
        "11293.95,,05/28/2026\n"
        '05/22/2026,"YOU SOLD SHORT SALE DIREXION SHARES ETF TRUST DAILY S&P ... (SPXL) (Short)",'
        'SPXL,"DIREXION SHARES ETF",Short,273.11,-10,,0.06,,2731.04,,05/26/2026\n'
        '05/20/2026,"YOU BOUGHT SHORT COVER POET TECHNOLOGIES INC COM NPV (POET) (Short)",POET,'
        '"POET TECHNOLOGIES INC",Short,14.55,50,,,,-727.5,,05/21/2026\n'
        '05/27/2026,"SHORT VS MARGIN MARK TO MARKET (Margin)",,"No Description",Margin,,0.000,,,,'
        "-105.96,,\n"
    )

    events = parse_fidelity_csv(content)

    # Non-trade row (MARK TO MARKET) must be excluded
    assert len(events) == 6

    equity_buy, equity_sell, opt_buy_open, opt_sell_open, short_sale, short_cover = events

    # Long equity buy → OPEN, side C
    assert equity_buy.underlying == "NET"
    assert equity_buy.effect == "OPEN"
    assert equity_buy.side == "C"
    assert equity_buy.premium == 205.0
    assert equity_buy.quantity == pytest.approx(30 / 100)
    assert equity_buy.fees == 0.0

    # Long equity sell → CLOSE, side C, fees from Fees ($) column
    assert equity_sell.underlying == "SOLS"
    assert equity_sell.effect == "CLOSE"
    assert equity_sell.side == "C"
    assert equity_sell.fees == pytest.approx(0.19)

    # Option buy to open — compact symbol normalised to OCC format
    assert opt_buy_open.symbol == "SPXW 260618P07400000"
    assert opt_buy_open.underlying == "SPXW"
    assert opt_buy_open.exp_date == date(2026, 6, 18)
    assert opt_buy_open.call_or_put == "P"
    assert opt_buy_open.strike == pytest.approx(7400.0)
    assert opt_buy_open.effect == "OPEN"
    assert opt_buy_open.side == "B"
    assert opt_buy_open.premium == pytest.approx(54.07)
    assert opt_buy_open.quantity == 2.0
    assert opt_buy_open.fees == pytest.approx(0.05)

    # Option sell to open (short option)
    assert opt_sell_open.symbol == "SPXW 260618P07410000"
    assert opt_sell_open.effect == "OPEN"
    assert opt_sell_open.side == "S"

    # Short equity sale → OPEN, side S
    assert short_sale.underlying == "SPXL"
    assert short_sale.effect == "OPEN"
    assert short_sale.side == "S"

    # Short equity cover → CLOSE, side S
    assert short_cover.underlying == "POET"
    assert short_cover.effect == "CLOSE"
    assert short_cover.side == "S"


def test_parse_fidelity_csv_compact_option_symbol_small_strike() -> None:
    """Compact symbols with small strikes (e.g. NOK $16 call) must zero-pad to 8 digits."""
    content = (
        "\n"
        "Run Date,Action,Symbol,Description,Type,Price ($),Quantity,Commission ($),Fees ($),"
        "Accrued Interest ($),Amount ($),Cash Balance ($),Settlement Date\n"
        '05/26/2026,"YOU SOLD CLOSING TRANSACTION CALL (NOK) NOKIA OYJ ADR EACH DEC 18 26 '
        '$24 (100 SHS) (Cash)"," -NOK261218C24","CALL (NOK)...",Cash,1.9,-2,,0.05,,379.95,,'
        "05/27/2026\n"
    )

    events = parse_fidelity_csv(content)

    assert len(events) == 1
    event = events[0]
    assert event.symbol == "NOK 261218C00024000"
    assert event.underlying == "NOK"
    assert event.strike == pytest.approx(24.0)
    assert event.exp_date == date(2026, 12, 18)
    assert event.call_or_put == "C"
    assert event.effect == "CLOSE"
    assert event.side == "B"

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
