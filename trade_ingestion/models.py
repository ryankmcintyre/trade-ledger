from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from hashlib import sha256


@dataclass(slots=True)
class CanonicalTrade:
    lot_id: str
    trade_id: str
    underlying: str
    symbol: str
    open_date: date
    exp_date: date | None
    call_or_put: str | None
    side: str
    strike: float | None
    stock_price_open: float | None
    premium: float | None
    quantity: float
    fees: float
    exit_price: float | None
    close_date: date | None
    account: str

    def __post_init__(self) -> None:
        if not self.trade_id:
            self.trade_id = make_trade_id(self)


@dataclass(slots=True)
class RawEvent:
    lot_id: str
    broker: str
    account: str
    underlying: str
    symbol: str
    trade_date: date
    exp_date: date | None
    call_or_put: str | None
    side: str
    strike: float | None
    stock_price: float | None
    premium: float | None
    quantity: float
    fees: float
    effect: str


@dataclass(slots=True)
class OpenLot:
    event: RawEvent
    remaining_quantity: float
    remaining_fees: float
    split_index: int = field(default=0)


def make_trade_id(trade: CanonicalTrade) -> str:
    quantity = f"{trade.quantity:g}"
    return f"{trade.open_date.isoformat()}|{trade.symbol}|{trade.side}|{quantity}"


def make_fallback_lot_id(
    *,
    trade_date: date,
    symbol: str,
    quantity: float,
    premium: float | None,
) -> str:
    premium_value = "" if premium is None else f"{premium:.8f}"
    payload = f"{trade_date.isoformat()}|{symbol}|{quantity:.8f}|{premium_value}"
    return sha256(payload.encode("utf-8")).hexdigest()[:12]
