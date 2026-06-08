from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import date
from typing import Deque

from constants import UNDERLYING_DISPLAY_MAP
from trade_ingestion.models import CanonicalTrade, OpenLot, RawEvent

MATCH_EPSILON = 1e-9


@dataclass(slots=True)
class MatchResult:
    trades: list[CanonicalTrade]
    skipped_duplicates: int
    open_positions: int


def match_trades(events: list[RawEvent], existing_lot_ids: set[str] | None = None) -> list[CanonicalTrade]:
    return match_trades_with_summary(events, existing_lot_ids).trades


def match_trades_with_summary(events: list[RawEvent], existing_lot_ids: set[str] | None = None) -> MatchResult:
    aggregated = _pre_aggregate(events)
    open_lots: dict[tuple[str, str, str], Deque[OpenLot]] = defaultdict(deque)
    results: list[CanonicalTrade] = []
    open_positions = 0

    for event in sorted(aggregated, key=_event_sort_key):
        key = (event.account, event.symbol, event.side)
        if event.effect == "OPEN":
            open_lots[key].append(
                OpenLot(
                    event=event,
                    remaining_quantity=event.quantity,
                    remaining_fees=event.fees,
                )
            )
            continue

        remaining_close_quantity = event.quantity
        close_fee_rate = (event.fees or 0.0) / event.quantity if event.quantity else 0.0
        lots = open_lots[key]
        while remaining_close_quantity > MATCH_EPSILON and lots:
            lot = lots[0]
            matched_quantity = min(lot.remaining_quantity, remaining_close_quantity)
            open_fee_share = _allocate_fee(lot.remaining_fees, lot.remaining_quantity, matched_quantity)
            close_fee_share = close_fee_rate * matched_quantity
            total_fee = (open_fee_share or 0.0) + close_fee_share
            trade = _make_trade(
                lot=lot,
                quantity=matched_quantity,
                fees=total_fee if total_fee > 0.0 else None,
                exit_price=event.premium,
                close_date=event.trade_date,
                split_suffix=None,
            )
            results.append(trade)

            lot.remaining_quantity -= matched_quantity
            lot.remaining_fees = (lot.remaining_fees or 0.0) - (open_fee_share or 0.0)
            remaining_close_quantity -= matched_quantity

            if lot.remaining_quantity <= MATCH_EPSILON:
                lots.popleft()

        # Orphan close: no matching open lot — write as close-only row
        if remaining_close_quantity > MATCH_EPSILON:
            close_fees = close_fee_rate * remaining_close_quantity
            trade = _make_orphan_close(event, remaining_close_quantity, close_fees if close_fees > 0.0 else None)
            results.append(trade)

    for lots in open_lots.values():
        for lot in lots:
            if lot.remaining_quantity <= MATCH_EPSILON:
                continue
            open_positions += 1
            lot.split_index += 1
            trade = _make_trade(
                lot=lot,
                quantity=lot.remaining_quantity,
                fees=lot.remaining_fees if (lot.remaining_fees or 0.0) > 0.0 else None,
                exit_price=None,
                close_date=None,
                split_suffix=f"open-{lot.split_index}" if lot.remaining_quantity != lot.event.quantity else None,
            )
            results.append(trade)

    return MatchResult(
        trades=results,
        skipped_duplicates=0,
        open_positions=open_positions,
    )


def _pre_aggregate(events: list[RawEvent]) -> list[RawEvent]:
    """Merge same-day, same-symbol, same-side, same-effect events with weighted-average pricing."""
    groups: dict[tuple[str, str, str, str, date], list[RawEvent]] = defaultdict(list)
    for event in events:
        key = (event.account, event.symbol, event.side, event.effect, event.trade_date)
        groups[key].append(event)

    aggregated: list[RawEvent] = []
    for group_events in groups.values():
        if len(group_events) == 1:
            aggregated.append(group_events[0])
            continue

        total_qty = sum(e.quantity for e in group_events)
        # Weighted average premium
        premiums_with_qty = [(e.premium or 0.0, e.quantity) for e in group_events]
        total_premium_value = sum(p * q for p, q in premiums_with_qty)
        avg_premium = total_premium_value / total_qty if total_qty > 0 else None
        # Round to avoid floating point noise
        if avg_premium is not None:
            avg_premium = round(avg_premium, 10)

        # Sum fees (only non-None)
        fee_values = [e.fees for e in group_events if e.fees is not None]
        total_fees = sum(fee_values) if fee_values else None

        base = group_events[0]
        merged_lot_id = "|".join(e.lot_id for e in group_events)
        aggregated.append(RawEvent(
            lot_id=merged_lot_id,
            broker=base.broker,
            account=base.account,
            underlying=base.underlying,
            symbol=base.symbol,
            trade_date=base.trade_date,
            exp_date=base.exp_date,
            call_or_put=base.call_or_put,
            side=base.side,
            strike=base.strike,
            stock_price=base.stock_price,
            premium=avg_premium,
            quantity=total_qty,
            fees=total_fees,
            effect=base.effect,
        ))
    return aggregated


def _event_sort_key(event: RawEvent) -> tuple[object, int, str]:
    return (event.trade_date, 0 if event.effect == "OPEN" else 1, event.lot_id)


def _allocate_fee(total_fee: float | None, quantity_pool: float, quantity_slice: float) -> float | None:
    if total_fee is None:
        return None
    if quantity_pool <= MATCH_EPSILON:
        return None
    return total_fee * (quantity_slice / quantity_pool)


def _resolve_stock(underlying: str) -> str:
    """Resolve underlying ticker to Column A display value."""
    return UNDERLYING_DISPLAY_MAP.get(underlying, underlying)


def _make_trade(
    *,
    lot: OpenLot,
    quantity: float,
    fees: float | None,
    exit_price: float | None,
    close_date: date | None,
    split_suffix: str | None,
) -> CanonicalTrade:
    lot_id = lot.event.lot_id if split_suffix is None else f"{lot.event.lot_id}:{split_suffix}"
    return CanonicalTrade(
        lot_id=lot_id,
        trade_id="",
        underlying=lot.event.underlying,
        symbol=lot.event.symbol,
        open_date=lot.event.trade_date,
        exp_date=lot.event.exp_date,
        call_or_put=lot.event.call_or_put,
        side=lot.event.side,
        strike=lot.event.strike,
        stock_price_open=lot.event.stock_price,
        premium=lot.event.premium,
        quantity=quantity,
        fees=fees,
        exit_price=exit_price,
        close_date=close_date,
        account=lot.event.account,
        stock=_resolve_stock(lot.event.underlying),
    )


def _make_orphan_close(event: RawEvent, quantity: float, fees: float | None) -> CanonicalTrade:
    """Create a close-only row for sells without a matching open in the import data."""
    return CanonicalTrade(
        lot_id=event.lot_id,
        trade_id="",
        underlying=event.underlying,
        symbol=event.symbol,
        open_date=None,
        exp_date=event.exp_date,
        call_or_put=event.call_or_put,
        side=event.side,
        strike=event.strike,
        stock_price_open=None,
        premium=None,
        quantity=quantity,
        fees=fees,
        exit_price=event.premium,
        close_date=event.trade_date,
        account=event.account,
        stock=_resolve_stock(event.underlying),
    )
