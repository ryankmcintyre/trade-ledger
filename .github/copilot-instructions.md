# Copilot Instructions — trade-ledger

## Project Purpose

This project ingests broker trade export CSVs from multiple brokers (currently Fidelity and Robinhood),
normalizes them into a canonical schema, matches open and close events into single trade rows, and
writes the results into an existing Excel workbook. The workbook contains a formal Excel Table
(`tbl_trades`) that feeds pivot table reports on a separate sheet. The script writes only raw input
fields — formula-driven columns are owned by Excel and must never be overwritten.

---

## Architecture Overview

```
broker CSV(s)
     │
     ▼
adapters/          # one module per broker — parses CSV into list[RawEvent]
     │
     ▼
matcher.py         # matches RawEvents into list[CanonicalTrade] (open + close → one row)
     │
     ▼
writer.py          # appends CanonicalTrade rows to tbl_trades via xlwings
```

Entry point: `main.py` — accepts broker name and CSV path as CLI arguments, runs the full pipeline.

---

## Core Data Models

All models live in `models.py`. Do not define data models anywhere else.

### `RawEvent`
Represents a single broker event (one buy or one sell) as emitted by the broker CSV.
Not written to Excel — intermediate representation only.

### `CanonicalTrade`
Represents one complete trade: open and close in a single row, matching the structure of `tbl_trades`.

```python
@dataclass
class CanonicalTrade:
    lot_id: str               # PRIMARY dedup key — broker transaction ID if available,
                              # else SHA-256 hash of open_date+symbol+quantity+premium (12 chars)
    trade_id: str             # human-readable composite: open_date+symbol+side+quantity (display only)
    underlying: str           # ticker of the underlying (e.g. 'SPY')
    symbol: str               # raw symbol — OCC format for options (e.g. 'SPY 230915C00450000')
    open_date: date
    exp_date: date | None     # options only
    call_or_put: str | None   # 'C', 'P', or None for equities
    side: str                 # 'B' or 'S' for options or 'C' for equities
    strike: float | None      # options only
    stock_price_open: float | None
    premium: float | None     # per-contract for options, per-share for equities
    quantity: float
    fees: float               # total fees for both legs combined
    exit_price: float | None  # None for open positions
    close_date: date | None   # None for open positions
    account: str              # broker name + account identifier
```

**One row per lot.** Adding to a position on the same day produces multiple `CanonicalTrade` rows,
each with their own `lot_id`. Never merge or average same-day adds.

---

## Module Responsibilities

### `models.py`
- `RawEvent` and `CanonicalTrade` dataclasses
- `make_trade_id(trade: CanonicalTrade) -> str` — builds the human-readable composite ID
- `make_lot_id_hash(open_date, symbol, quantity, premium) -> str` — SHA-256 fallback, 12-char truncation
- No business logic beyond these helpers

### `adapters/fidelity.py`
- Accepts a CSV file path, returns `list[RawEvent]`
- Skips metadata rows at the top of the file — locate the true header row dynamically, do not hardcode a row number
- Normalizes: dates → `datetime.date`, side → `'B'`/`'S'`/`C`, fees by summing all commission and fee columns, symbols to OCC format where applicable
- Uses broker-assigned transaction ID as `lot_id` where available
- Add a `# NOTE:` comment documenting whether native transaction IDs are used or the hash fallback

### `adapters/robinhood.py`
- Same interface and contract as the Fidelity adapter
- Robinhood exports equities and options separately — the adapter must handle both file types
- Filter out non-trade rows (assignments, expirations, dividends) and document each filtered type with a `# NOTE:` comment
- Add a `# NOTE:` comment documenting lot_id strategy

### `matcher.py`
- Accepts `list[RawEvent]` and `set[str]` of existing `lot_id` values already in the workbook
- Matches open and close events into `CanonicalTrade` objects:
  - Match criteria: same `symbol`, `open_date <= close_date`, quantity alignment
  - Default cost basis method: FIFO — document this as a `# TODO:` for review
- Unmatched opens → `CanonicalTrade` with `exit_price=None`, `close_date=None`
- Partial closes → split into two `CanonicalTrade` rows: one matched (closed), one remaining open
- Skips any trade whose `lot_id` is already in the existing set (dedup)
- Returns `list[CanonicalTrade]`

### `writer.py`
- Uses `xlwings` exclusively — no other Excel library
- Opens the workbook, locates `tbl_trades` by table name
- Reads existing `lot_id` column to build the dedup set passed to the matcher
- Appends new rows in correct column order
- Never writes to formula-driven columns: `current_stock_price`, `break_even_price`, `dte`,
  `profit_loss`, `days_held`, `return_on_capital`, `status`
- Does not close the workbook if it was already open when the script started — check xlwings app state
- Entry point: `write_trades(workbook_path: Path, trades: list[CanonicalTrade]) -> int`
  returns count of rows written

### `main.py`
- CLI entry point using `argparse`
- Arguments: `--broker [fidelity|robinhood]`, `--csv <path>`, `--workbook <path>`
- Runs the full pipeline: adapter → matcher → writer
- Prints a summary on completion: rows ingested, rows skipped (dedup), open positions left unmatched

---

## Coding Standards

- **Type hints everywhere** — all function signatures, all dataclass fields, no untyped code
- **No magic numbers or hardcoded strings** — column names, table names, and broker identifiers
  go in a `constants.py` file at the project root
- **Explicit over clever** — prefer readable conditionals over compact expressions when matching logic
  is involved; this code will be debugged against real broker data
- **Fail loudly on schema mismatches** — if a broker CSV is missing an expected column, raise a
  descriptive `ValueError` identifying the column and the broker, do not silently skip
- **`# TODO:` for assumptions** — any behavior not explicitly specified above should be implemented
  as the most conservative choice and flagged with a `# TODO:` comment explaining the assumption
- **`# NOTE:` for broker-specific decisions** — document format quirks, lot_id strategy, and
  filtered event types in the relevant adapter with `# NOTE:` comments

---

## Testing Standards

- All tests use `pytest`
- No tests depend on real broker files — use inline CSV fixtures as strings
- No tests hit the real filesystem for Excel — use a temp workbook fixture or mock xlwings
- Each module has a corresponding test file under `tests/`
- Required test cases per module are specified in the Copilot scaffold prompt — cover all of them
  before adding extras

---

## Project Structure

```
trade-ledger/
├── .github/
│   └── copilot-instructions.md   # this file
├── trade_ingestion/
│   ├── constants.py
│   ├── models.py
│   ├── matcher.py
│   ├── writer.py
│   ├── adapters/
│   │   ├── __init__.py
│   │   ├── fidelity.py
│   │   └── robinhood.py
│   └── tests/
│       ├── test_models.py
│       ├── test_fidelity.py
│       ├── test_robinhood.py
│       ├── test_matcher.py
│       └── test_writer.py
├── main.py
├── requirements.txt
└── README.md
```

---

## Dependencies

- `xlwings` — Excel interaction (writer only)
- `pandas` — CSV parsing in adapters (optional but preferred for header-skipping logic)
- `pytest` — testing
- Python 3.11+

---

## What This Project Does Not Do

- Does not pull live data from broker APIs — file-based ingestion only
- Does not modify formula-driven columns in the workbook
- Does not support Google Sheets or any non-Excel target
- Does not implement the AI feedback layer — that is a separate pipeline
