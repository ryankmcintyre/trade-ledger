from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Sequence

from trade_ingestion.adapters import parse_fidelity_csv_detailed
from trade_ingestion.matcher import match_trades_with_summary
from trade_ingestion.models import CanonicalTrade
from trade_ingestion.adapters.fidelity import FidelityParseResult
from trade_ingestion.retry import ResolutionFailure
from trade_ingestion.writer import ConversionFailure, write_trades_detailed

Adapter = Callable[[str, "Callable[[str, str], str | None] | None"], FidelityParseResult]
ADAPTERS: dict[str, Adapter] = {
    "fidelity": parse_fidelity_csv_detailed,
}


@dataclass(slots=True)
class PipelineResult:
    rows_ingested: int
    rows_skipped: int
    open_positions: int
    failed_conversions: list[str] = field(default_factory=list)
    conversion_failures: list[ConversionFailure] = field(default_factory=list)
    symbol_failures: list[ResolutionFailure] = field(default_factory=list)


def run_pipeline(
    *,
    broker: str,
    csv_path: Path,
    workbook_path: Path,
    sheet_name: str,
    ticker_prompt: Callable[[str, CanonicalTrade], str | None] | None = None,
    symbol_prompt: Callable[[str, str], str | None] | None = None,
) -> PipelineResult:
    adapter = ADAPTERS.get(broker.strip().lower())
    if adapter is None:
        supported = ", ".join(sorted(ADAPTERS))
        raise ValueError(f"Unsupported broker {broker!r}. Supported brokers: {supported}")

    csv_content = csv_path.read_text(encoding="utf-8-sig")
    parse_result = adapter(csv_content, symbol_prompt)
    match_result = match_trades_with_summary(parse_result.events)
    write_result = write_trades_detailed(
        workbook_path,
        sheet_name,
        match_result.trades,
        ticker_prompt=ticker_prompt,
    )
    written = write_result.rows_written
    # The writer deduplicates against existing rows using composite keys.
    skipped = len(match_result.trades) - written
    return PipelineResult(
        rows_ingested=written,
        rows_skipped=skipped,
        open_positions=match_result.open_positions,
        failed_conversions=write_result.failed_conversions,
        conversion_failures=write_result.conversion_failures,
        symbol_failures=parse_result.symbol_failures,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Ingest broker trade CSVs into the trade ledger workbook.")
    parser.add_argument("broker", help="Broker adapter name, for example: fidelity")
    parser.add_argument("csv_path", type=Path, help="Path to the broker export CSV file")
    parser.add_argument(
        "--workbook",
        type=Path,
        required=True,
        help="Path to the Excel workbook containing tbl_trades",
    )
    parser.add_argument(
        "--sheet",
        required=True,
        help="Name of the worksheet inside the workbook that contains the tbl_trades table",
    )
    parser.add_argument(
        "--no-prompt",
        action="store_true",
        help="Skip interactive ticker retry prompts and report conversion failures instead",
    )
    return parser


def _stdin_is_interactive() -> bool:
    try:
        return bool(sys.stdin.isatty())
    except (AttributeError, OSError, ValueError):
        return False


def _prompt_for_replacement_value(input_value: str, record_label: str, *, what: str) -> str | None:
    prompt_text = (
        f"Could not resolve {what} {input_value!r} for record {record_label!r}. "
        "Enter an updated stock ticker to retry (leave blank to skip): "
    )
    try:
        response = input(prompt_text).strip()
    except EOFError:
        return None
    return response or None


def _prompt_for_replacement_ticker(input_ticker: str, trade: CanonicalTrade) -> str | None:
    record_label = trade.trade_id or trade.lot_id or trade.symbol or "trade"
    return _prompt_for_replacement_value(input_ticker, record_label, what="stock ticker")


def _prompt_for_replacement_symbol(input_symbol: str, record_label: str) -> str | None:
    return _prompt_for_replacement_value(input_symbol, record_label, what="option symbol")


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    ticker_prompt: Callable[[str, CanonicalTrade], str | None] | None = None
    symbol_prompt: Callable[[str, str], str | None] | None = None
    if not args.no_prompt and _stdin_is_interactive():
        ticker_prompt = _prompt_for_replacement_ticker
        symbol_prompt = _prompt_for_replacement_symbol

    result = run_pipeline(
        broker=args.broker,
        csv_path=args.csv_path,
        workbook_path=args.workbook,
        sheet_name=args.sheet,
        ticker_prompt=ticker_prompt,
        symbol_prompt=symbol_prompt,
    )
    open_label = "open position" if result.open_positions == 1 else "open positions"
    print(
        f"Ingested {result.rows_ingested} trade rows to {args.workbook} [{args.sheet}]; "
        f"skipped {result.rows_skipped} duplicate rows; "
        f"left {result.open_positions} {open_label} unmatched"
    )
    if result.symbol_failures:
        row_label = "row" if len(result.symbol_failures) == 1 else "rows"
        print(
            f"Warning: {len(result.symbol_failures)} {row_label} could not be parsed and were "
            "skipped due to an unsupported option symbol format."
        )
    for symbol_failure in result.symbol_failures:
        attempted_symbol = symbol_failure.attempted_value or symbol_failure.input_value
        print(
            f"Symbol parse failed for record {symbol_failure.context_label!r} "
            f"(input symbol {symbol_failure.input_value!r}; attempted symbol {attempted_symbol!r}): "
            f"{symbol_failure.error}"
        )
    if result.failed_conversions:
        tickers = ", ".join(sorted(set(result.failed_conversions)))
        row_label = "row" if len(result.failed_conversions) == 1 else "rows"
        print(
            f"Warning: {len(result.failed_conversions)} {row_label} could not be converted to the "
            f"Stocks data type ({tickers}); 'Stock Symbol' and 'Current Stock Price' "
            "will not resolve for them."
        )
    for failure in result.conversion_failures:
        trade_label = failure.trade.trade_id or failure.trade.lot_id or failure.trade.symbol or "trade"
        attempted_ticker = failure.attempted_ticker or failure.input_ticker
        print(
            f"Conversion failed for record {trade_label!r} "
            f"(input ticker {failure.input_ticker!r}; attempted ticker {attempted_ticker!r}): "
            f"{failure.error}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
