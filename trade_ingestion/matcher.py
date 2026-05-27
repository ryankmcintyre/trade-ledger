from __future__ import annotations

import warnings
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import date
from typing import Deque

from trade_ingestion.models import CanonicalTrade, OpenLot, RawEvent

MATCH_EPSILON = 1e-9


@dataclass(slots=True)
class MatchResult:
    trades: list[CanonicalTrade]
    skipped_duplicates: int
    open_positions: int


# TODO: Broker exports in scope do not provide an explicit open-lot reference on close rows.
# TODO: Until one is available, closes are matched FIFO within (account, symbol, side), which is
# TODO: the most conservative assumption that preserves order and avoids cross-account matching.
def match_trades(events: list[RawEvent], existing_lot_ids: set[str]) -> list[CanonicalTrade]:
    return match_trades_with_summary(events, existing_lot_ids).trades


def match_trades_with_summary(events: list[RawEvent], existing_lot_ids: set[str]) -> MatchResult:
    open_lots: dict[tuple[str, str, str], Deque[OpenLot]] = defaultdict(deque)
    results: list[CanonicalTrade] = []
    skipped_duplicates = 0
    open_positions = 0

    for event in sorted(events, key=_event_sort_key):
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
        close_fee_rate = event.fees / event.quantity if event.quantity else 0.0
        lots = open_lots[key]
        while remaining_close_quantity > MATCH_EPSILON and lots:
            lot = lots[0]
            matched_quantity = min(lot.remaining_quantity, remaining_close_quantity)
            open_fee_share = _allocate_fee(lot.remaining_fees, lot.remaining_quantity, matched_quantity)
            close_fee_share = close_fee_rate * matched_quantity
            trade = _make_trade(
                lot=lot,
                quantity=matched_quantity,
                fees=open_fee_share + close_fee_share,
                exit_price=event.premium,
                close_date=event.trade_date,
                split_suffix=None,
            )
            if trade.lot_id not in existing_lot_ids:
                results.append(trade)
            else:
                skipped_duplicates += 1

            lot.remaining_quantity -= matched_quantity
            lot.remaining_fees -= open_fee_share
            remaining_close_quantity -= matched_quantity

            if lot.remaining_quantity <= MATCH_EPSILON:
                lots.popleft()

        if remaining_close_quantity > MATCH_EPSILON:
            # TODO: Date-range-limited broker exports can contain close rows whose opening lots are
            # outside the imported window. Until there is a durable way to persist or recover
            # those opens, ingestion skips the orphaned close after warning instead of aborting.
            warnings.warn(
                "Close event could not be matched to an open lot and was skipped: "
                f"lot_id={event.lot_id} trade_date={event.trade_date.isoformat()} "
                f"account={event.account} symbol={event.symbol} side={event.side} quantity={event.quantity} "
                f"remaining_quantity={remaining_close_quantity}",
                stacklevel=2,
            )

    for lots in open_lots.values():
        for lot in lots:
            if lot.remaining_quantity <= MATCH_EPSILON:
                continue
            open_positions += 1
            lot.split_index += 1
            trade = _make_trade(
                lot=lot,
                quantity=lot.remaining_quantity,
                fees=lot.remaining_fees,
                exit_price=None,
                close_date=None,
                split_suffix=f"open-{lot.split_index}" if lot.remaining_quantity != lot.event.quantity else None,
            )
            if trade.lot_id not in existing_lot_ids:
                results.append(trade)
            else:
                skipped_duplicates += 1

    return MatchResult(
        trades=results,
        skipped_duplicates=skipped_duplicates,
        open_positions=open_positions,
    )


def _event_sort_key(event: RawEvent) -> tuple[object, int, str]:
    return (event.trade_date, 0 if event.effect == "OPEN" else 1, event.lot_id)


def _allocate_fee(total_fee: float, quantity_pool: float, quantity_slice: float) -> float:
    if quantity_pool <= MATCH_EPSILON:
        return 0.0
    return total_fee * (quantity_slice / quantity_pool)


def _make_trade(
    *,
    lot: OpenLot,
    quantity: float,
    fees: float,
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
    )
