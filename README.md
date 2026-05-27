# trade-ledger

Ingests broker trade export CSVs, normalises them into a canonical schema, matches open and close
events into single trade rows, and appends the results to an Excel workbook containing a table
named `tbl_trades`.

---

## Requirements

- Python 3.11+
- [xlwings](https://www.xlwings.org/) (Excel must be installed on the machine)
- pandas (optional, used internally for CSV parsing)

Install dependencies:

```bash
pip install xlwings pandas
```

---

## Usage

```
python main.py <broker> <csv_path> --workbook <workbook_path>
```

| Argument | Description |
|---|---|
| `broker` | Broker adapter name (see [Supported brokers](#supported-brokers)) |
| `csv_path` | Path to the broker-exported CSV file |
| `--workbook` | Path to the Excel workbook containing `tbl_trades` |

**Example:**

```bash
python main.py fidelity ~/Downloads/History.csv --workbook ~/trades/ledger.xlsx
```

**Output:**

```
Ingested 12 trade rows to /Users/ryan/trades/ledger.xlsx; skipped 3 duplicate rows; left 2 open positions unmatched
```

- **Ingested** — new rows written to `tbl_trades`
- **Skipped** — rows already present in the workbook (deduplicated by `lot_id`)
- **Open positions** — trades with no matching close event in the imported file (written as open rows with blank `exit_price` and `close_date`)

---

## Supported brokers

### Fidelity (`fidelity`)

Parses Fidelity transaction history CSV exports. The adapter handles:

- **Equities** — `Buy` and `Sell` actions
- **Options** — `Buy to Open`, `Sell to Close`, `Sell to Open`, `Buy to Close` actions
- Metadata rows at the top of the file are skipped automatically; no hardcoded row offset required
- Option symbols are normalised to OCC format (e.g. `SPY 240119C00450000`)
- Date formats accepted: `YYYY-MM-DD`, `MM/DD/YYYY`, `MM/DD/YY`
- Fees are summed across all commission and fee columns

To export from Fidelity: **Accounts & Trade → Activity & Orders → History** → select a date range → **Download**.

---

## How it works

1. The adapter parses the CSV into a list of raw trade events (one buy or sell per row).
2. The matcher pairs open and close events FIFO within each `(account, symbol, side)` group into complete trade rows. Partial closes produce two rows: one matched, one remaining open.
3. The writer reads existing `lot_id` values from `tbl_trades` to skip duplicates, then appends new rows.
4. Formula-driven columns (`current_stock_price`, `break_even_price`, `dte`, `profit_loss`, `days_held`, `return_on_capital`, `status`) are never written — they remain owned by Excel.

---

## Running the tests

```bash
python -m pytest trade_ingestion/tests -q
```