FIDELITY_BROKER_NAME = "Fidelity"

TABLE_NAME = "tbl_trades"

# Mapping from CanonicalTrade field names to workbook column headers.
FIELD_TO_COLUMN: dict[str, str] = {
    "stock": "Stock",
    "open_date": "Open Date",
    "exp_date": "Exp Date",
    "call_or_put": "Call or Put",
    "side": "B/S",
    "stock_price_open": "Stock Price DOC",
    "strike": "Strike Price",
    "premium": "Premium",
    "quantity": "C",
    "fees": "Fees",
    "exit_price": "Exit Price",
    "close_date": "Close Date",
    "account": "Account",
    "status": "Status",
}

# Columns used to form the composite dedup key when reading existing rows.
DEDUP_COLUMNS: tuple[str, ...] = ("Stock", "Open Date", "B/S", "C")

# Formula-driven column holding the ticker resolved from the Column A Stocks entity.
STOCK_SYMBOL_COLUMN = "Stock Symbol"

# Excel linked data type service identifier for Stocks, and the culture used to
# resolve the entity. Column A must hold a Stocks entity (not text) because the
# workbook drives "Stock Symbol" and "Current Stock Price" from _FV(A, ...).
STOCKS_SERVICE_ID = 268435456
LINKED_DATA_TYPE_CULTURE = "en-US"

# Maps option root symbols to the display name written to Column A (Stock).
# Regular equity tickers are written as-is.
UNDERLYING_DISPLAY_MAP: dict[str, str] = {
    "SPXW": "S&P 500 INDEX",
    "SPX": "S&P 500 INDEX",
}
