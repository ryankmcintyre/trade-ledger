# trade-ledger

Ingests broker trade export CSVs, normalises them into a canonical schema, matches open and close
events into single trade rows, and appends the results to an Excel workbook containing a table
named `tbl_trades`.

---

## Download

Windows users can download a pre-built executable from the [latest release](https://github.com/ryankmcintyre/trade-ledger/releases/latest).

From the release's **Assets** section, download `trade_ledger.exe` and place it in a folder of
your choice (for example `C:\Tools\trade_ledger\`).

Excel is required on the machine because `xlwings` uses Excel via COM automation.

**Usage:**

```
trade_ledger.exe <broker> <csv_path> --workbook <workbook_path>
```

**Example:**

```
trade_ledger.exe fidelity "C:\Downloads\History.csv" --workbook "C:\trades\ledger.xlsx"
```

---

## Requirements

If you're running from source as a developer:

- Python 3.11+
- [xlwings](https://www.xlwings.org/) (Excel must be installed on the machine)

Install dependencies:

```bash
pip install -r requirements.txt
```

Install development and test tooling:

```bash
pip install -r requirements.txt -r requirements-dev.txt
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
- **Skipped** — rows already present in the workbook (deduplicated by composite key)
- **Open positions** — trades with no matching close event in the imported file (written as open rows with blank `exit_price` and `close_date`)

---

## Supported brokers

### Fidelity (`fidelity`)

Parses Fidelity transaction history CSV exports. The adapter handles:

- **Equities** — buy and sell actions (both short-form like `Buy` and the verbose descriptions Fidelity uses in real exports, e.g. `YOU BOUGHT CLOUDFLARE INC CL A COM (NET) (Margin)`)
- **Options** — open and close transactions for calls and puts (e.g. `YOU BOUGHT OPENING TRANSACTION PUT (SPXW)...`)
- Metadata rows and footer disclaimer text at the top/bottom of the file are skipped automatically
- Option symbols are normalised to OCC format from Fidelity's compact notation (e.g. `-SPXW260618P7400` → `SPXW 260618P07400000`)
- Date formats accepted: `YYYY-MM-DD`, `MM/DD/YYYY`, `MM/DD/YY`
- Fees include only the `Commission ($)` column (the `Fees ($)` column is excluded)

To export from Fidelity: **Accounts & Trade → Activity & Orders → History** → select a date range → **Download**.

---

## How it works

1. The adapter parses the CSV into a list of raw trade events (one buy or sell per row).
2. Same-day events for the same symbol/side/effect are pre-aggregated (weighted-average pricing, summed quantities).
3. The matcher pairs open and close events FIFO within each `(account, symbol, side)` group into complete trade rows. Partial closes produce two rows: one matched, one remaining open. Sells without a matching open produce close-only rows.
4. The writer reads existing rows from `tbl_trades` to skip duplicates (composite key: Stock + Open Date + B/S + Quantity), then appends new rows.
5. Formula-driven columns are never written — they remain owned by Excel.

---

## Running the tests

```bash
python -m pytest trade_ingestion/tests -q
```