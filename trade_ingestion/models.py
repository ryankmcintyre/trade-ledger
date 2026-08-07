from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from hashlib import sha256
from typing import Generic, TypeVar

T = TypeVar("T")


@dataclass(slots=True)
class CanonicalTrade:
    lot_id: str
    trade_id: str
    underlying: str
    symbol: str
    open_date: date | None
    exp_date: date | None
    call_or_put: str | None
    side: str | None
    strike: float | None
    stock_price_open: float | None
    premium: float | None
    quantity: float
    fees: float | None
    exit_price: float | None
    close_date: date | None
    account: str
    stock: str  # Display value for Column A (ticker or display name from UNDERLYING_DISPLAY_MAP)
    status: str | None = None

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
    side: str | None
    strike: float | None
    stock_price: float | None
    premium: float | None
    quantity: float
    fees: float | None
    effect: str


@dataclass(slots=True)
class OpenLot:
    event: RawEvent
    remaining_quantity: float
    remaining_fees: float | None
    split_index: int = field(default=0)


@dataclass(slots=True)
class ResolutionFailure(Generic[T]):
    """Details about a resolution that could not be completed, even after a
    single prompt-and-retry attempt (shared by the writer's stock-ticker
    conversion recovery and the Fidelity adapter's option-symbol recovery)."""

    context_label: str
    input_value: str
    attempted_value: str | None
    error: str


@dataclass(slots=True)
class FidelityParseResult:
    """Outcome of a parse_fidelity_csv_detailed run."""

    events: list[RawEvent]
    # Rows whose option symbol could not be parsed, even after a single
    # prompt-and-retry attempt with an operator-supplied replacement ticker.
    # These rows are skipped rather than aborting the whole import.
    symbol_failures: list[ResolutionFailure] = field(default_factory=list)


def make_trade_id(trade: CanonicalTrade) -> str:
    quantity = f"{trade.quantity:g}"
    date_str = trade.open_date.isoformat() if trade.open_date else "no-open"
    side = trade.side or ""
    return f"{date_str}|{trade.symbol}|{side}|{quantity}"


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
