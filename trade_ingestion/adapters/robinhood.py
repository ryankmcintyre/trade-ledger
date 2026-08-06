from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime
from typing import Any

from constants import ROBINHOOD_BROKER_NAME
from trade_ingestion.models import RawEvent, make_fallback_lot_id

DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d", "%b %d, %Y", "%b %d %Y")
OCC_SYMBOL_RE = re.compile(r"^(?P<underlying>[A-Z.]+)\s(?P<exp>\d{6})(?P<cp>[CP])(?P<strike>\d{8})$")
OPTION_TEXT_RE = re.compile(
    r"^(?P<underlying>[A-Z.]+)\s+(?P<exp>\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})\s+(?P<strike>\d+(?:\.\d+)?)\s*(?P<cp>CALL|PUT|C|P)$",
    re.IGNORECASE,
)
FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "trade_date": ("Date", "Trade Date", "Transaction Date", "Execution Date", "Settlement Date"),
    "action": ("Description", "Transaction Type", "Type", "Action", "Side"),
    "symbol": ("Symbol", "Underlying", "Security", "Description"),
    "quantity": ("Quantity", "Qty", "Shares", "Quantity (Shares)"),
    "price": ("Price", "Price ($)", "Average Price", "Cost Basis"),
    "commission": ("Commission", "Commission ($)", "Commission Fee"),
    "fees": ("Fees", "Fees ($)", "Fee", "Fee ($)", "Net Fees"),
    "account": ("Account", "Account Name"),
    "transaction_id": ("Transaction ID", "Order ID", "ID", "Reference"),
    "security_type": ("Security Type", "Instrument", "Type", "Option Type"),
    "underlying_price": ("Underlying Price", "Underlying Last Price"),
}


class RobinhoodParseError(ValueError):
    pass


# NOTE: Robinhood exports may contain separate equity and option files, so the adapter normalizes both.
# NOTE: For rows without a broker transaction identifier, the adapter falls back to a deterministic hash.
def parse_robinhood_csv(content: str) -> list[RawEvent]:
    rows = list(csv.reader(io.StringIO(content)))
    header_index = _find_header_index(rows)
    if header_index is None:
        raise RobinhoodParseError("Could not locate Robinhood header row")

    empty_before_header = sum(1 for row in rows[:header_index] if not row)
    reader = csv.DictReader(io.StringIO(content), fieldnames=rows[header_index])
    for _ in range(header_index - empty_before_header + 1):
        next(reader, None)

    events: list[RawEvent] = []
    for row in reader:
        event = _parse_row(row)
        if event is not None:
            events.append(event)
    return events


def _find_header_index(rows: list[list[str]]) -> int | None:
    for index, row in enumerate(rows):
        normalized = {cell.strip().lower() for cell in row if cell.strip()}
        if normalized.intersection({"date", "description", "quantity", "price", "symbol"}) and normalized.intersection(
            {"description", "type", "transaction type"}
        ):
            return index
    return None


def _parse_row(row: dict[str, str | None]) -> RawEvent | None:
    action = (_get_value(row, "action") or "").strip()
    if not action:
        return None

    classification = _classify_action(action)
    if classification is None:
        # NOTE: Dividend, assignment, expiration, transfer, and other non-trade rows are skipped.
        return None

    symbol_text = (_get_value(row, "symbol") or "").strip()
    if not symbol_text:
        return None

    trade_date = _parse_date(_required_value(row, "trade_date"))
    quantity_raw = _parse_optional_float(_get_value(row, "quantity"))
    if quantity_raw is None:
        return None

    quantity_signed = quantity_raw
    quantity = abs(quantity_signed)
    price = _parse_optional_float(_get_value(row, "price"))
    fees = _sum_fees(row)

    security_type = (_get_value(row, "security_type") or "").strip().lower()
    is_option = classification["instrument"] == "option" or "option" in security_type or "call" in security_type or "put" in security_type
    if is_option:
        parsed_symbol = _normalize_option_symbol(symbol_text)
        normalized_quantity = quantity
    else:
        parsed_symbol = {
            "underlying": _normalize_equity_symbol(symbol_text),
            "symbol": _normalize_equity_symbol(symbol_text),
            "exp_date": None,
            "call_or_put": None,
            "strike": None,
        }
        normalized_quantity = quantity / 100.0 if quantity > 0 else quantity / 100.0

    lot_id = (_get_value(row, "transaction_id") or "").strip() or make_fallback_lot_id(
        trade_date=trade_date,
        symbol=parsed_symbol["symbol"],
        quantity=normalized_quantity,
        premium=price,
    )

    return RawEvent(
        lot_id=lot_id,
        broker=ROBINHOOD_BROKER_NAME,
        account=(_get_value(row, "account") or ROBINHOOD_BROKER_NAME).strip() or ROBINHOOD_BROKER_NAME,
        underlying=parsed_symbol["underlying"],
        symbol=parsed_symbol["symbol"],
        trade_date=trade_date,
        exp_date=parsed_symbol["exp_date"],
        call_or_put=parsed_symbol["call_or_put"],
        side=classification["side"],
        strike=parsed_symbol["strike"],
        stock_price=_parse_optional_float(_get_value(row, "underlying_price")),
        premium=price,
        quantity=normalized_quantity,
        fees=fees,
        effect=classification["effect"],
    )


def _get_value(row: dict[str, str | None], alias_key: str) -> str | None:
    for field_name in FIELD_ALIASES[alias_key]:
        if field_name in row and row[field_name] not in (None, ""):
            return row[field_name]
    return None


def _required_value(row: dict[str, str | None], alias_key: str) -> str:
    value = _get_value(row, alias_key)
    if value in (None, ""):
        raise RobinhoodParseError(f"Missing required Robinhood field: {alias_key}")
    return value


def _classify_action(action: str) -> dict[str, str] | None:
    normalized = action.strip().lower()
    if any(token in normalized for token in ("dividend", "interest", "deposit", "withdraw", "transfer", "assignment", "expiration", "expire", "cash", "market to market")):
        return None

    if any(token in normalized for token in ("buy to open", "buy open")):
        return {"effect": "OPEN", "side": "B", "instrument": "option"}
    if any(token in normalized for token in ("sell to open", "sell open", "short", "short sale")):
        return {"effect": "OPEN", "side": "S", "instrument": "option"}
    if any(token in normalized for token in ("buy to close", "buy close")):
        return {"effect": "CLOSE", "side": "S", "instrument": "option"}
    if any(token in normalized for token in ("sell to close", "sell close")):
        return {"effect": "CLOSE", "side": "B", "instrument": "option"}

    if any(token in normalized for token in ("buy", "bought", "purchase")):
        if "cover" in normalized:
            return {"effect": "CLOSE", "side": "S", "instrument": "equity"}
        return {"effect": "OPEN", "side": "C", "instrument": "equity"}
    if any(token in normalized for token in ("sell", "sold")):
        if "cover" in normalized:
            return {"effect": "OPEN", "side": "S", "instrument": "equity"}
        return {"effect": "CLOSE", "side": "C", "instrument": "equity"}

    return None


def _normalize_option_symbol(symbol: str) -> dict[str, Any]:
    symbol_text = symbol.strip()
    occ_match = OCC_SYMBOL_RE.match(symbol_text)
    if occ_match:
        exp_date = datetime.strptime(occ_match.group("exp"), "%y%m%d").date()
        strike = int(occ_match.group("strike")) / 1000.0
        return {
            "underlying": occ_match.group("underlying"),
            "symbol": symbol_text,
            "exp_date": exp_date,
            "call_or_put": occ_match.group("cp"),
            "strike": strike,
        }

    option_text_match = OPTION_TEXT_RE.match(symbol_text)
    if option_text_match:
        exp_text = option_text_match.group("exp")
        exp_date = _parse_exp_date(exp_text)
        strike_value = float(option_text_match.group("strike"))
        cp = option_text_match.group("cp").upper()
        if cp in {"CALL", "C"}:
            call_or_put = "C"
        else:
            call_or_put = "P"
        occ_symbol = (
            f"{option_text_match.group('underlying')} {exp_date.strftime('%y%m%d')}"
            f"{call_or_put}{int(round(strike_value * 1000)):08d}"
        )
        return {
            "underlying": option_text_match.group("underlying"),
            "symbol": occ_symbol,
            "exp_date": exp_date,
            "call_or_put": call_or_put,
            "strike": strike_value,
        }

    # TODO: Robinhood option symbols vary by export; use the raw symbol when a format is unsupported.
    return {
        "underlying": symbol_text.split()[0] if symbol_text.split() else symbol_text,
        "symbol": symbol_text,
        "exp_date": None,
        "call_or_put": None,
        "strike": None,
    }


def _normalize_equity_symbol(symbol: str) -> str:
    return symbol.split()[0].strip() if symbol.split() else symbol.strip()


def _sum_fees(row: dict[str, str | None]) -> float | None:
    total = 0.0
    seen = False
    for field in tuple(FIELD_ALIASES["commission"]) + tuple(FIELD_ALIASES["fees"]):
        if field in row and row[field] not in (None, ""):
            total += _parse_optional_float(row[field]) or 0.0
            seen = True
    return total if seen and total > 0.0 else None


def _parse_date(value: str) -> date:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            continue
    raise RobinhoodParseError(f"Unsupported Robinhood date format: {value}")


def _parse_exp_date(value: str) -> date:
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%y%m%d"):
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            continue
    raise RobinhoodParseError(f"Unsupported Robinhood option date format: {value}")


def _parse_optional_float(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    return float(str(value).replace("$", "").replace(",", "").strip())
