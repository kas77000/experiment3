# TCA Client Report

A **standalone** Transaction Cost Analysis report generator. It reads one
`orders.csv` plus a few optional auxiliary CSVs and writes a client-ready report
in three formats:

| Output | What it is |
|--------|-----------|
| `<client>_TCA.xlsx` | Multi-sheet workbook: summary, the two order-type tables, dark analysis, venues, algo summary, and an embedded-chart sheet. |
| `<client>_TCA.pdf`  | Paginated slide-style report: cover, executive summary, tables, and every chart. |
| `png/*.png`         | Each chart as a standalone image for decks / emails. |

No parquet, no KDB, no Bloomberg — just the CSVs. One file: `tca_report.py`.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python tca_report.py --orders orders.csv --aux-dir ./aux \
    --outdir ./report --client "DYMON" --period "Q2 2026"
```

| Flag | Meaning |
|------|---------|
| `--orders` | path to `orders.csv` (required) |
| `--aux-dir` | folder holding the optional auxiliary CSVs (below); omit to skip them |
| `--outdir` | output folder (default `./report`) |
| `--client` | client name shown on the cover / filenames |
| `--period` | period label; auto-derived from the `Date` column if omitted |
| `--cost-positive` | set this **only** if your export uses `+ = cost`. See sign convention. |

### Try it with synthetic data first

```bash
python make_sample_data.py --outdir ./sample_data --n 900
python tca_report.py --orders sample_data/orders.csv --aux-dir sample_data \
    --outdir report --client "DYMON"
```

The sample generator writes `orders.csv` and all four aux CSVs with the exact
column names below, drawn so higher `%DARK` goes with better slippage (so the
dark story is visible). It is synthetic — replace with the real files and the
same command runs unchanged.

## Inputs

### `orders.csv` (required)

The friendly per-order export. Column headers are matched **loosely** — case is
ignored and any `(...)` note is stripped, so `"$Mln (x1000000)"` resolves to
`$Mln` and `"Start(HK)"` to `Start`. Recognised columns (aliases in brackets):

`client`, `Trader`, `Date`, `Sym`, `Side`, `Start(HK)`, `End(HK)`,
`aggrTgtId`, `strike`, `Strategy`, `Cap`, `sector`, `auctionOnly`,
`arrivalTime`, `marketLimit`, `#Shares` (×1000), `$Mln` (×1e6),
`%DV` (= %ADV), `ePR`, `FR`, `Vol`, `Sprd`, `IS`, `Open`, `Close`,
`ePvwap`, `ePvwap/Sprd`, `Dur`, `Pvwap`, `%DARK`.

Only three are strictly required: `marketLimit`, `IS`, `Pvwap`. Everything else
degrades gracefully if missing.

- **Executed notional (the weight for every weighted average)** = `$Mln` × 1e6.
  If `$Mln` is absent it falls back to `#Shares` × 1000 × `strike`.
- `FR` and `%DARK` are auto-normalised: a 0–1 fraction is scaled to 0–100%.

### Auxiliary CSVs (all optional, in `--aux-dir`)

Headers are matched loosely, same as above. Missing files are skipped silently.

**`dark_routed.csv`** — *Routed Volume by Dark Venue* pie:

```
Venue,Routed %
SIGMA-X,22.4
UBS-ATS,18.1
...
```

**`dark_executed.csv`** — *Executed Volume by Venue* pie:

```
Venue,Executed %
SIGMA-X,19.7
UBS-ATS,20.3
...
```

**`dark_venue_summary.csv`** — the *Dark Venue Summary* table:

```
Venue Category,Shares,Exec Value,% Total Shares,% Total Exec Value
Continuous Dark,42000000,3200000000,38.1,35.4
Conditional (Block),...
```

**`algo_summary.csv`** — *Performance Summary By Algorithm*:

```
Strategy,Orders,Exec Shares,Exec Value,%Weight,Period Part,%ADV,Spread Bps,Benchmark,Impact Bps,Wgt Impact Bps,Wgt Impact Cps
VWAP,312,...,Order PVWAP,...
```

## What the report contains

**Reproduced from the earlier deliverable** (notional-weighted, `+ good / − cost`):

1. **Total Slippage Breakdown by Order Type** — all strategies (Market / Limit / All).
2. **Slippage Breakdown by Order Type per Strategy**.

Both carry: `# Orders`, `Arrival / Open / Close / PVWAP Slippage (bps)`, `%ADV`,
`Volatility`, `Spread (bps)`, `Fill Rate (%)`.

**Added — the DARK execution case** (the four requested angles):

- **A. Lower slippage when dark %** — arrival & PVWAP slippage by `%DARK` bucket
  (`<5% … 50%+`) with an indicative recoverable-cost estimate.
- **B. Spread capture** — spread-normalised slippage (`ePvwap/Sprd`) rising with dark.
- **C. Dark benefit by order size** — arrival slippage, low vs high dark, per `%ADV`
  bucket, showing the benefit concentrates on the largest / hardest orders (also
  the reduced-signalling angle).
- **D. Venue reach** — routed vs executed volume by venue (gap = capture headroom)
  and the dark venue summary.
- Plus the **Performance Summary by Algorithm** chart when `algo_summary.csv` is present.

## Sign convention (important)

Slippage columns are used **as-is**, matching the existing pipeline:
**`+` = outperformance (good), `−` = cost (bad)**, in bps. Every chart and table
reads that way, and the dark "benefit" is a positive delta.

If your `orders.csv` uses the opposite convention (`+` = cost), pass
`--cost-positive` and the loader flips the sign of every slippage column
(`IS`, `Open`, `Close`, `Pvwap`, `ePvwap`, `ePvwap/Sprd`) on the way in, so the
report still reads `+ = good` everywhere. **Check one known order before sending
to the client** — this is the single most important thing to get right.

## Notes / assumptions to confirm against real data

- The **$ savings** figure is an *indicative upper bound*: the notional-weighted
  arrival-slippage gap between the high-dark (≥30%) and low-dark (<15%) bands,
  applied to the notional still executing low-dark. It assumes low-dark flow can
  reach high-dark performance — real capacity is finite. Label it as indicative
  in any client conversation.
- Dark and size buckets are fixed (`DARK_BINS` / `SIZE_BINS` near the top of
  `tca_report.py`) — adjust to taste.
- Charts embedded in the XLSX are the same PNGs written to `png/`.
