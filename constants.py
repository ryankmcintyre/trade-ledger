FIDELITY_BROKER_NAME = "Fidelity"

TABLE_NAME = "tbl_trades"

STOCK_FIELD_NAME = "stock"

# Optional workbook column holding CanonicalTrade.lot_id, the primary dedup
# discriminator. Not every workbook has this column; when present, the writer
# prefers it over the composite DEDUP_COLUMNS key (see writer._make_dedup_key /
# writer._existing_dedup_keys) because the composite key is not guaranteed to be
# unique — e.g. two legitimate lots opened same-day on the same contract with the
# same side/quantity would otherwise collide and one would be dropped as a
# duplicate.
LOT_ID_COLUMN = "Lot ID"

# Mapping from CanonicalTrade field names to workbook column headers.
FIELD_TO_COLUMN: dict[str, str] = {
    "lot_id": LOT_ID_COLUMN,
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

# Columns used to form the composite dedup key when reading existing rows that
# have no Lot ID value (e.g. rows written before LOT_ID_COLUMN existed).
# NOTE: "Strike Price" is required here because index options (e.g. SPXW/SPX) all
# NOTE: resolve to the same "Stock" display name via UNDERLYING_DISPLAY_MAP, so without
# NOTE: the strike, distinct same-day/same-side/same-quantity trades on different
# NOTE: strikes would collide on the same composite key and be dropped as duplicates.
DEDUP_COLUMNS: tuple[str, ...] = ("Stock", "Open Date", "B/S", "C", "Strike Price")

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

# Canonical option-type values written to the workbook, and the broker-supplied
# aliases (upper-cased) that normalize to them.
CALL_OPTION_TYPE = "Call"
PUT_OPTION_TYPE = "Put"
OPTION_TYPE_ALIASES: dict[str, str] = {
    "C": CALL_OPTION_TYPE,
    "CALL": CALL_OPTION_TYPE,
    "P": PUT_OPTION_TYPE,
    "PUT": PUT_OPTION_TYPE,
}
