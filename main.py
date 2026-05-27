from __future__ import annotations

import argparse
from pathlib import Path
from typing import Callable, Sequence

from trade_ingestion.adapters import parse_fidelity_csv
from trade_ingestion.matcher import match_trades
from trade_ingestion.models import RawEvent
from trade_ingestion.writer import read_existing_lot_ids, write_trades

Adapter = Callable[[str], list[RawEvent]]
ADAPTERS: dict[str, Adapter] = {
    "fidelity": parse_fidelity_csv,
}


def run_pipeline(*, broker: str, csv_path: Path, workbook_path: Path) -> int:
    adapter = ADAPTERS.get(broker.strip().lower())
    if adapter is None:
        supported = ", ".join(sorted(ADAPTERS))
        raise ValueError(f"Unsupported broker {broker!r}. Supported brokers: {supported}")

    csv_content = csv_path.read_text(encoding="utf-8-sig")
    events = adapter(csv_content)
    existing_lot_ids = read_existing_lot_ids(workbook_path)
    trades = match_trades(events, existing_lot_ids)
    return write_trades(workbook_path, trades)


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
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    written = run_pipeline(broker=args.broker, csv_path=args.csv_path, workbook_path=args.workbook)
    print(f"Wrote {written} trade rows to {args.workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
