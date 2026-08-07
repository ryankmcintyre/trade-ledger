from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Sequence

from trade_ingestion.adapters import parse_fidelity_csv
from trade_ingestion.matcher import match_trades_with_summary
from trade_ingestion.models import RawEvent
from trade_ingestion.writer import write_trades_detailed

Adapter = Callable[[str], list[RawEvent]]
ADAPTERS: dict[str, Adapter] = {
    "fidelity": parse_fidelity_csv,
}


@dataclass(slots=True)
class PipelineResult:
    rows_ingested: int
    rows_skipped: int
    open_positions: int
    failed_conversions: list[str] = field(default_factory=list)


def run_pipeline(*, broker: str, csv_path: Path, workbook_path: Path, sheet_name: str) -> PipelineResult:
    adapter = ADAPTERS.get(broker.strip().lower())
    if adapter is None:
        supported = ", ".join(sorted(ADAPTERS))
        raise ValueError(f"Unsupported broker {broker!r}. Supported brokers: {supported}")

    csv_content = csv_path.read_text(encoding="utf-8-sig")
    events = adapter(csv_content)
    match_result = match_trades_with_summary(events)
    write_result = write_trades_detailed(workbook_path, sheet_name, match_result.trades)
    written = write_result.rows_written
    # The writer deduplicates against existing rows using composite keys.
    skipped = len(match_result.trades) - written
    return PipelineResult(
        rows_ingested=written,
        rows_skipped=skipped,
        open_positions=match_result.open_positions,
        failed_conversions=write_result.failed_conversions,
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
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    result = run_pipeline(
        broker=args.broker,
        csv_path=args.csv_path,
        workbook_path=args.workbook,
        sheet_name=args.sheet,
    )
    open_label = "open position" if result.open_positions == 1 else "open positions"
    print(
        f"Ingested {result.rows_ingested} trade rows to {args.workbook} [{args.sheet}]; "
        f"skipped {result.rows_skipped} duplicate rows; "
        f"left {result.open_positions} {open_label} unmatched"
    )
    if result.failed_conversions:
        tickers = ", ".join(sorted(set(result.failed_conversions)))
        row_label = "row" if len(result.failed_conversions) == 1 else "rows"
        print(
            f"Warning: {len(result.failed_conversions)} {row_label} could not be converted to the "
            f"Stocks data type ({tickers}); 'Stock Symbol' and 'Current Stock Price' "
            "will not resolve for them."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
