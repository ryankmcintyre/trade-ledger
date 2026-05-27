from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime
from typing import Iterable

from constants import FIDELITY_BROKER_NAME
from trade_ingestion.models import RawEvent, make_fallback_lot_id
HEADER_REQUIREMENTS = {"Action", "Symbol"}
DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y")
OPTION_SYMBOL_RE = re.compile(
    r"^(?P<underlying>[A-Z.]+)\s+(?P<exp>\d{2}/\d{2}/\d{4})\s+(?P<strike>\d+(?:\.\d+)?)\s+(?P<cp>[CP])$"
)
OCC_SYMBOL_RE = re.compile(
    r"^(?P<underlying>[A-Z.]+)\s(?P<exp>\d{6})(?P<cp>[CP])(?P<strike>\d{8})$"
)
FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "trade_date": ("Run Date", "Trade Date", "Date", "Settlement Date"),
    "action": ("Action", "Transaction Type", "Type"),
    "symbol": ("Symbol", "Description"),
    "quantity": ("Quantity", "Qty"),
    "price": ("Price", "Net Amount Per Share", "Amount"),
    "commission": ("Commission",),
    "fees": ("Fees", "Fee", "Reg Fee", "Additional Fees"),
    "account": ("Account", "Account Number"),
    "transaction_id": ("Transaction ID", "Reference Number", "Trade ID"),
    "security_type": ("Security Type", "Type Detail"),
    "underlying_price": ("Underlying Price", "Underlying Last Price"),
}


class FidelityParseError(ValueError):
    pass


# NOTE: Fidelity transaction exports are expected to include a transaction identifier.
# NOTE: When that identifier is missing, this adapter falls back to a deterministic hash
# NOTE: of trade_date + symbol + quantity + premium, per the pipeline requirements.
def parse_fidelity_csv(content: str) -> list[RawEvent]:
    rows = list(csv.reader(io.StringIO(content)))
    header_index = _find_header_index(rows)
    if header_index is None:
        raise FidelityParseError("Could not locate Fidelity header row")

    reader = csv.DictReader(io.StringIO(content), fieldnames=rows[header_index])
    for _ in range(header_index + 1):
        next(reader, None)

    events: list[RawEvent] = []
    for row in reader:
        event = _parse_row(row)
        if event is not None:
            events.append(event)
    return events


def _find_header_index(rows: list[list[str]]) -> int | None:
    for index, row in enumerate(rows):
        normalized = {cell.strip() for cell in row if cell.strip()}
        if HEADER_REQUIREMENTS.issubset(normalized):
            return index
    return None


def _parse_row(row: dict[str, str | None]) -> RawEvent | None:
    action = (_get_value(row, "action") or "").strip()
    mapping = _map_action(action)
    if mapping is None:
        return None

    symbol_text = (_get_value(row, "symbol") or "").strip()
    if not symbol_text:
        return None

    trade_date = _parse_date(_required_value(row, "trade_date"))
    quantity_raw = abs(_parse_float(_required_value(row, "quantity")))
    price = _parse_optional_float(_get_value(row, "price"))
    account_value = (_get_value(row, "account") or "UNKNOWN").strip()
    security_type = (_get_value(row, "security_type") or "").strip().lower()

    if mapping["instrument"] == "option" or "option" in security_type:
        parsed_symbol = _normalize_option_symbol(symbol_text)
        quantity = quantity_raw
        premium = price
    else:
        parsed_symbol = {
            "underlying": symbol_text,
            "symbol": symbol_text,
            "exp_date": None,
            "call_or_put": None,
            "strike": None,
        }
        quantity = quantity_raw / 100.0
        premium = price

    fees = _sum_fees(row)
    lot_id = (_get_value(row, "transaction_id") or "").strip() or make_fallback_lot_id(
        trade_date=trade_date,
        symbol=parsed_symbol["symbol"],
        quantity=quantity,
        premium=premium,
    )

    return RawEvent(
        lot_id=lot_id,
        broker=FIDELITY_BROKER_NAME,
        account=f"{FIDELITY_BROKER_NAME}:{account_value}",
        underlying=parsed_symbol["underlying"],
        symbol=parsed_symbol["symbol"],
        trade_date=trade_date,
        exp_date=parsed_symbol["exp_date"],
        call_or_put=parsed_symbol["call_or_put"],
        side=mapping["side"],
        strike=parsed_symbol["strike"],
        stock_price=_parse_optional_float(_get_value(row, "underlying_price")),
        premium=premium,
        quantity=quantity,
        fees=fees,
        effect=mapping["effect"],
    )


def _get_value(row: dict[str, str | None], alias_key: str) -> str | None:
    for field_name in FIELD_ALIASES[alias_key]:
        if field_name in row and row[field_name] not in (None, ""):
            return row[field_name]
    return None


def _required_value(row: dict[str, str | None], alias_key: str) -> str:
    value = _get_value(row, alias_key)
    if value in (None, ""):
        raise FidelityParseError(f"Missing required Fidelity field: {alias_key}")
    return value


def _map_action(action: str) -> dict[str, str] | None:
    normalized = action.strip().lower()
    mapping: dict[str, dict[str, str]] = {
        "buy": {"effect": "OPEN", "side": "C", "instrument": "equity"},
        "sell": {"effect": "CLOSE", "side": "C", "instrument": "equity"},
        "buy to open": {"effect": "OPEN", "side": "B", "instrument": "option"},
        "sell to close": {"effect": "CLOSE", "side": "B", "instrument": "option"},
        "sell to open": {"effect": "OPEN", "side": "S", "instrument": "option"},
        "buy to close": {"effect": "CLOSE", "side": "S", "instrument": "option"},
    }
    # TODO: Fidelity can export short equity transactions, but the canonical schema only
    # TODO: specifies side='C' for equities. This adapter conservatively treats plain buy/sell
    # TODO: as long equity open/close until a short-equity requirement is specified.
    return mapping.get(normalized)


def _normalize_option_symbol(symbol: str) -> dict[str, object]:
    occ_match = OCC_SYMBOL_RE.match(symbol)
    if occ_match:
        exp_date = datetime.strptime(occ_match.group("exp"), "%y%m%d").date()
        strike = int(occ_match.group("strike")) / 1000.0
        return {
            "underlying": occ_match.group("underlying"),
            "symbol": symbol,
            "exp_date": exp_date,
            "call_or_put": occ_match.group("cp"),
            "strike": strike,
        }

    match = OPTION_SYMBOL_RE.match(symbol)
    if not match:
        raise FidelityParseError(f"Unsupported Fidelity option symbol format: {symbol}")

    exp_date = datetime.strptime(match.group("exp"), "%m/%d/%Y").date()
    strike_value = float(match.group("strike"))
    occ_symbol = (
        f"{match.group('underlying')} {exp_date.strftime('%y%m%d')}"
        f"{match.group('cp')}{int(round(strike_value * 1000)):08d}"
    )
    return {
        "underlying": match.group("underlying"),
        "symbol": occ_symbol,
        "exp_date": exp_date,
        "call_or_put": match.group("cp"),
        "strike": strike_value,
    }


def _sum_fees(row: dict[str, str | None]) -> float:
    return sum(
        _parse_optional_float(row.get(field)) or 0.0
        for alias in ("commission", "fees")
        for field in FIELD_ALIASES[alias]
        if field in row
    )


def _parse_date(value: str) -> date:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            continue
    raise FidelityParseError(f"Unsupported Fidelity date format: {value}")


def _parse_float(value: str) -> float:
    return float(value.replace("$", "").replace(",", "").strip())


def _parse_optional_float(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    return _parse_float(value)
