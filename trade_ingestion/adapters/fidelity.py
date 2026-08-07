from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Callable

from constants import FIDELITY_BROKER_NAME
from trade_ingestion.models import RawEvent, make_fallback_lot_id
from trade_ingestion.retry import ResolutionFailure, resolve_with_retry

DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%b-%d-%Y")
OPTION_SYMBOL_RE = re.compile(
    r"^(?P<underlying>[A-Z.]+)\s+(?P<exp>\d{2}/\d{2}/\d{4})\s+(?P<strike>\d+(?:\.\d+)?)\s+(?P<cp>[CP])$"
)
OCC_SYMBOL_RE = re.compile(
    r"^(?P<underlying>[A-Z.]+)\s(?P<exp>\d{6})(?P<cp>[CP])(?P<strike>\d{8})$"
)
# NOTE: Real Fidelity CSV exports use a compact symbol format with an optional leading dash and no
# zero-padded strike: e.g. -SPXW260618P7400, -NOK261016C16, or NVDL260807P26. This is distinct from full OCC format.
FIDELITY_COMPACT_RE = re.compile(
    r"^(?:-)?(?P<underlying>[A-Z.]+)(?P<exp>\d{6})(?P<cp>[CP])(?P<strike>\d+(?:\.\d+)?)$"
)
# NOTE: Underlying symbols are normally letters/dots only ([A-Z.]+ above), but brokers
# occasionally rename a ticker to include a digit after a corporate action (e.g. a stock
# split renames HON to HON2 in the broker's own system while the public ticker is still
# HON). None of the strict regexes above can match that. This fallback isolates the
# trailing date+call/put+strike suffix regardless of what precedes it, so a malformed
# prefix can still be spliced out and swapped for an operator-supplied replacement ticker.
COMPACT_SUFFIX_RE = re.compile(
    r"^(?P<dash>-?)(?P<prefix>.*?)(?P<exp>\d{6})(?P<cp>[CP])(?P<strike>\d+(?:\.\d+)?)$"
)
FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "trade_date": ("Run Date", "Trade Date", "Date", "Settlement Date"),
    "action": ("Action", "Activity Description", "Transaction Type", "Type"),
    "symbol": ("Symbol", "Description"),
    "quantity": ("Quantity", "Qty"),
    "price": ("Price ($)", "Price", "Net Amount Per Share", "Amount"),
    "commission": ("Commission ($)", "Commission"),
    "fees": ("Fees ($)", "Fees", "Fee", "Reg Fee", "Additional Fees"),
    "account": ("Account", "Account Number"),
    "transaction_id": ("Transaction ID", "Reference Number", "Trade ID"),
    "security_type": ("Security Type", "Type Detail"),
    "underlying_price": ("Underlying Price", "Underlying Last Price"),
}
# Reuse FIELD_ALIASES to keep header detection aligned with parsing aliases.
HEADER_ACTION_ALIASES = set(FIELD_ALIASES["action"])
HEADER_SYMBOL_ALIASES = set(FIELD_ALIASES["symbol"])


class FidelityParseError(ValueError):
    pass


@dataclass(slots=True)
class FidelityParseResult:
    """Outcome of a parse_fidelity_csv_detailed run."""

    events: list[RawEvent]
    # Rows whose option symbol could not be parsed, even after a single
    # prompt-and-retry attempt with an operator-supplied replacement ticker.
    # These rows are skipped rather than aborting the whole import.
    symbol_failures: list[ResolutionFailure] = field(default_factory=list)


# NOTE: Fidelity transaction exports are expected to include a transaction identifier.
# NOTE: When that identifier is missing, this adapter falls back to a deterministic hash
# NOTE: of trade_date + symbol + quantity + premium, per the pipeline requirements.
def parse_fidelity_csv(
    content: str,
    symbol_prompt: Callable[[str, str], str | None] | None = None,
) -> list[RawEvent]:
    """Parse a Fidelity export into RawEvents. Thin wrapper over
    parse_fidelity_csv_detailed for callers that don't need failure details."""
    return parse_fidelity_csv_detailed(content, symbol_prompt).events


def parse_fidelity_csv_detailed(
    content: str,
    symbol_prompt: Callable[[str, str], str | None] | None = None,
) -> FidelityParseResult:
    rows = list(csv.reader(io.StringIO(content)))
    header_index = _find_header_index(rows)
    if header_index is None:
        raise FidelityParseError("Could not locate Fidelity header row")

    # csv.DictReader silently skips empty rows (row == []) while csv.reader counts them.
    # Subtract any empty rows that precede the header so the skip loop doesn't over-advance
    # and accidentally consume the first real data row.
    empty_before_header = sum(1 for row in rows[:header_index] if not row)
    reader = csv.DictReader(io.StringIO(content), fieldnames=rows[header_index])
    for _ in range(header_index - empty_before_header + 1):
        next(reader, None)

    events: list[RawEvent] = []
    symbol_failures: list[ResolutionFailure] = []
    for row in reader:
        event, failure = _parse_row(row, symbol_prompt)
        if event is not None:
            events.append(event)
        if failure is not None:
            symbol_failures.append(failure)
    return FidelityParseResult(events=events, symbol_failures=symbol_failures)


def _find_header_index(rows: list[list[str]]) -> int | None:
    for index, row in enumerate(rows):
        normalized = {cell.strip() for cell in row if cell.strip()}
        if normalized.intersection(HEADER_ACTION_ALIASES) and normalized.intersection(
            HEADER_SYMBOL_ALIASES
        ):
            return index
    return None


def _parse_row(
    row: dict[str, str | None],
    symbol_prompt: Callable[[str, str], str | None] | None = None,
) -> tuple[RawEvent | None, ResolutionFailure | None]:
    action = (_get_value(row, "action") or "").strip()
    mapping = _map_action(action)
    if mapping is None:
        return None, None

    symbol_text = (_get_value(row, "symbol") or "").strip()
    if not symbol_text:
        return None, None

    trade_date = _parse_date(_required_value(row, "trade_date"))
    quantity_raw = abs(_parse_float(_required_value(row, "quantity")))
    price = _parse_optional_float(_get_value(row, "price"))
    security_type = (_get_value(row, "security_type") or "").strip().lower()

    if mapping["instrument"] == "option" or "option" in security_type:
        parsed_symbol, failure = _resolve_option_symbol(symbol_text, trade_date, action, symbol_prompt)
        if parsed_symbol is None:
            return None, failure
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
        account=FIDELITY_BROKER_NAME,
        underlying=parsed_symbol["underlying"],
        symbol=parsed_symbol["symbol"],
        trade_date=trade_date,
        exp_date=parsed_symbol["exp_date"],
        call_or_put=parsed_symbol["call_or_put"],
        side=mapping.get("side"),
        strike=parsed_symbol["strike"],
        stock_price=_parse_optional_float(_get_value(row, "underlying_price")),
        premium=premium,
        quantity=quantity,
        fees=fees,
        effect=mapping["effect"],
    ), None


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

    # Short-form exact matches — kept for alternate CSV layouts and test fixtures.
    _EXACT: dict[str, dict[str, str]] = {
        "buy": {"effect": "OPEN", "side": "C", "instrument": "equity"},
        "sell": {"effect": "CLOSE", "side": "C", "instrument": "equity"},
        "buy to open": {"effect": "OPEN", "side": "B", "instrument": "option"},
        "sell to close": {"effect": "CLOSE", "side": "B", "instrument": "option"},
        "sell to open": {"effect": "OPEN", "side": "S", "instrument": "option"},
        "buy to close": {"effect": "CLOSE", "side": "S", "instrument": "option"},
    }
    if normalized in _EXACT:
        return _EXACT[normalized]

    # NOTE: Fidelity emits assignment, exercise, and expiration rows as distinct lifecycle events;
    # NOTE: route them through the matcher with explicit effects so the import can write statuses.
    if "assigned" in normalized:
        return {"effect": "ASSIGNED", "instrument": "option"}
    if "exercised" in normalized:
        return {"effect": "EXERCISED", "instrument": "option"}
    if "expired" in normalized:
        return {"effect": "EXPIRED", "instrument": "option"}

    # Verbose Fidelity descriptions: "YOU BOUGHT/SOLD [OPENING/CLOSING TRANSACTION] ..."
    bought = "you bought" in normalized
    sold = "you sold" in normalized

    if "opening transaction" in normalized:
        if bought:
            return {"effect": "OPEN", "side": "B", "instrument": "option"}
        if sold:
            return {"effect": "OPEN", "side": "S", "instrument": "option"}

    if "closing transaction" in normalized:
        if bought:
            return {"effect": "CLOSE", "side": "S", "instrument": "option"}
        if sold:
            return {"effect": "CLOSE", "side": "B", "instrument": "option"}

    # TODO: Short equity transactions use side='S' as a placeholder; the canonical schema does
    # TODO: not yet explicitly specify how short equity open/close positions should be represented.
    if "short sale" in normalized and sold:
        return {"effect": "OPEN", "side": "S", "instrument": "equity"}
    if "short cover" in normalized and bought:
        return {"effect": "CLOSE", "side": "S", "instrument": "equity"}

    if bought:
        return {"effect": "OPEN", "side": "C", "instrument": "equity"}
    if sold:
        return {"effect": "CLOSE", "side": "C", "instrument": "equity"}

    return None


def _resolve_option_symbol(
    symbol_text: str,
    trade_date: date,
    action: str,
    symbol_prompt: Callable[[str, str], str | None] | None,
) -> tuple[dict[str, object] | None, ResolutionFailure | None]:
    """Normalize an option symbol, recovering from an unparseable prefix (e.g. a
    broker-renamed ticker like "HON2") by prompting once for a replacement ticker
    and retrying. On failure the caller skips the row rather than aborting the
    whole import."""
    context_label = f"{trade_date.isoformat()} {action} {symbol_text}".strip()

    def try_resolve(candidate: str) -> tuple[dict[str, object] | None, str | None]:
        try:
            return _normalize_option_symbol(candidate), None
        except (FidelityParseError, ValueError) as exc:
            return None, str(exc)

    def prompt(_input_symbol: str, _context_label: str) -> str | None:
        replacement_ticker = symbol_prompt(symbol_text, context_label)  # type: ignore[misc]
        if not replacement_ticker:
            return None
        return _reconstruct_compact_symbol(symbol_text, replacement_ticker)

    return resolve_with_retry(
        symbol_text,
        context_label,
        try_resolve,
        prompt if symbol_prompt is not None else None,
    )


def _reconstruct_compact_symbol(symbol: str, replacement_ticker: str) -> str | None:
    """Swap the prefix of a compact option symbol for a replacement ticker,
    preserving the trailing date+call/put+strike suffix. Returns None if the
    suffix can't be isolated at all (e.g. the legacy space-delimited format),
    in which case there is nothing sensible to reconstruct.
    """
    match = COMPACT_SUFFIX_RE.match(symbol)
    if not match:
        return None
    return (
        f"{match.group('dash')}{replacement_ticker.strip()}"
        f"{match.group('exp')}{match.group('cp')}{match.group('strike')}"
    )


def _normalize_option_symbol(symbol: str) -> dict[str, object]:
    # Fidelity compact format with leading dash: -UNDERLYING[YYMMDD][C/P][STRIKE]
    # e.g. -SPXW260618P7400, -NOK261016C16
    compact_match = FIDELITY_COMPACT_RE.match(symbol)
    if compact_match:
        exp_date = datetime.strptime(compact_match.group("exp"), "%y%m%d").date()
        strike_value = float(compact_match.group("strike"))
        occ_symbol = (
            f"{compact_match.group('underlying')} {compact_match.group('exp')}"
            f"{compact_match.group('cp')}{int(round(strike_value * 1000)):08d}"
        )
        return {
            "underlying": compact_match.group("underlying"),
            "symbol": occ_symbol,
            "exp_date": exp_date,
            "call_or_put": compact_match.group("cp"),
            "strike": strike_value,
        }

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


def _sum_fees(row: dict[str, str | None]) -> float | None:
    """Only include Commission ($) in fees — Fees ($) column is excluded per requirements."""
    total = 0.0
    has_commission_value = False
    for field in FIELD_ALIASES["commission"]:
        if field not in row:
            continue
        value = row.get(field)
        if value in (None, ""):
            continue
        has_commission_value = True
        total += _parse_fee_value(value)

    if not has_commission_value:
        return None
    return total if total > 0.0 else 0.0


def _parse_date(value: str) -> date:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            continue
    raise FidelityParseError(f"Unsupported Fidelity date format: {value}")


def _parse_float(value: str) -> float:
    return float(value.replace("$", "").replace(",", "").strip())


def _parse_fee_value(value: str | None) -> float:
    if value in (None, ""):
        return 0.0

    stripped = (value or "").strip()
    if stripped in {"--", "-"}:
        return 0.0

    try:
        return _parse_float(value)
    except ValueError as exc:
        raise ValueError(f"Invalid Fidelity commission value: {value}") from exc


def _parse_optional_float(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    return _parse_float(value)
