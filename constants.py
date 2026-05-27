FIDELITY_BROKER_NAME = "Fidelity"

TABLE_NAME = "tbl_trades"
WRITABLE_COLUMNS: tuple[str, ...] = (
    "lot_id",
    "trade_id",
    "underlying",
    "symbol",
    "open_date",
    "exp_date",
    "call_or_put",
    "side",
    "strike",
    "stock_price_open",
    "premium",
    "quantity",
    "fees",
    "exit_price",
    "close_date",
    "account",
)
