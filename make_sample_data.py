#!/usr/bin/env python3
"""
Synthetic sample data for tca_report.py
=======================================

Writes an ``orders.csv`` plus the four auxiliary CSVs (dark routed / executed
splits, dark venue summary, algo summary) using the EXACT column names the real
client export uses, so you can run ``tca_report.py`` end-to-end before the real
files arrive. The numbers are drawn so that **higher %DARK goes with better
slippage** -- this is synthetic and only there to exercise the persuasion story;
swap in the real files and the same script runs unchanged.

    python make_sample_data.py --outdir ./sample_data --n 900

Nothing here is real market data.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

STRATEGIES = ["VWAP", "TWAP", "IS", "CLOSE", "DARK_SEEK", "POV"]
SECTORS = ["Financials", "Tech", "Industrials", "Materials", "Health", "Energy"]
CAPS = ["Large", "Mid", "Small"]
AUCTION = ["ContinuousOnly", "Mixed", "CloseOnly", "OpenOnly"]
ARRIVAL = ["Pre-Open", "Day", "Last30Mins", "First30Mins"]
MKT_LIMIT = ["Market", "Limit"]
SIDES = ["Buy", "Sell"]
DARK_VENUES = ["BLKX (Dark)", "SIGMA-X", "UBS-ATS", "MS-Pool", "Instinct-X",
               "CS-Crossfinder", "Liquidnet"]


def _rng(seed: int) -> np.random.Generator:
    return np.random.default_rng(seed)


def make_orders(n: int, seed: int) -> pd.DataFrame:
    rng = _rng(seed)
    dates = pd.bdate_range("2026-04-01", "2026-06-30")
    date = rng.choice(dates, n)

    strategy = rng.choice(STRATEGIES, n, p=[0.34, 0.14, 0.12, 0.10, 0.16, 0.14])
    market_limit = rng.choice(MKT_LIMIT, n, p=[0.58, 0.42])
    side = rng.choice(SIDES, n)
    auction_only = rng.choice(AUCTION, n, p=[0.6, 0.2, 0.15, 0.05])

    # %adv: order participation as a fraction of ADV -- the REAL %ADV. Right-skewed.
    # (The legacy %DV column below is kept for back-compat but is NOT the %ADV.)
    pct_adv = np.round(rng.lognormal(mean=0.6, sigma=0.9, size=n), 2).clip(0.05, 45)
    # %DV: legacy order-value column, retained but no longer used as %ADV.
    pct_dv = np.round(rng.lognormal(mean=0.7, sigma=0.9, size=n), 2).clip(0.05, 60)
    vol = np.round(rng.normal(28, 9, n).clip(8, 70), 1)          # annualised vol %
    sprd = np.round(rng.lognormal(mean=1.6, sigma=0.5, size=n), 1).clip(1.5, 60)

    # %DARK: DARK_SEEK strategy leans high; market orders slightly higher dark
    base_dark = rng.beta(1.6, 3.2, n) * 100
    base_dark += np.where(strategy == "DARK_SEEK", 22, 0)
    base_dark += np.where(market_limit == "Market", 4, -3)
    dark = np.round(base_dark.clip(0, 95), 1)

    # ---- slippage model (convention: + = good / - = cost, in bps) -------------
    # cost grows with spread, size and vol; DARK participation RECOVERS cost.
    noise = rng.normal(0, 6, n)
    dark_benefit = dark / 100.0                     # 0..~1
    # arrival slippage (IS): dark helps most on the hardest (wide-spread/large) orders
    is_bps = (
        2.0
        - 0.35 * sprd
        - 0.20 * pct_adv
        - 0.05 * (vol - 28)
        + dark_benefit * (6 + 0.30 * sprd + 0.25 * pct_adv)   # <- dark recaptures
        + noise
    )
    # pvwap slippage: tighter, also improved by dark (spread capture)
    pvwap_bps = (
        1.0 - 0.10 * sprd + dark_benefit * (3.5 + 0.12 * sprd)
        + rng.normal(0, 4, n)
    )
    open_bps = rng.normal(-1.5, 8, n) + dark_benefit * 3
    close_bps = rng.normal(-2.0, 9, n) + dark_benefit * 3.5

    epvwap = pvwap_bps + rng.normal(0, 1.5, n)
    epvwap_sprd = np.round(epvwap / sprd, 3)

    fr = np.round(rng.beta(9, 1.3, n).clip(0.4, 1.0), 3)        # fill rate 0..1
    epr = np.round(rng.uniform(3, 25, n), 1)                    # participation %
    dur = np.round(rng.lognormal(mean=3.2, sigma=0.7, size=n), 1).clip(2, 390)

    shares = np.round(rng.lognormal(mean=3.2, sigma=1.0, size=n), 1).clip(1, 4000)  # x1000
    price = rng.uniform(5, 300, n)
    mln = np.round(shares * 1000 * price / 1e6, 3)             # $Mln

    # ---- participation split: %POST + %TAKE + %OPEN + %CLOSE + %DARK == 100 ------
    # Whatever isn't executed dark is shared across lit-passive (%POST),
    # lit-aggressive (%TAKE), the opening auction (%OPEN) and the closing auction
    # (%CLOSE). Dirichlet weights are tilted by strategy / auction preference.
    #   columns of alpha: [post, take, open, close]
    alpha = np.column_stack([
        np.full(n, 3.0),                                       # post (passive lit)
        np.full(n, 2.0),                                       # take (aggressive lit)
        np.full(n, 0.6),                                       # open auction
        np.full(n, 1.2),                                       # close auction
    ])
    alpha[:, 0] += np.where(market_limit == "Limit", 1.5, 0.0)     # limit -> more passive
    alpha[:, 1] += np.where(strategy == "IS", 1.5, 0.0)            # IS -> more aggressive
    alpha[:, 2] += np.where(auction_only == "OpenOnly", 4.0, 0.0)  # open auction
    alpha[:, 2] += np.where(auction_only == "Mixed", 0.8, 0.0)
    alpha[:, 3] += np.where(strategy == "CLOSE", 6.0, 0.0)         # close-strat -> close cross
    alpha[:, 3] += np.where(auction_only == "CloseOnly", 4.0, 0.0)
    alpha[:, 3] += np.where(auction_only == "Mixed", 0.8, 0.0)

    frac = rng.gamma(alpha)                                   # (n, 4) gamma draws
    frac = frac / frac.sum(axis=1, keepdims=True)             # each row sums to 1
    remaining = 100.0 - dark                                  # non-dark share of the order
    pct_post = np.round(frac[:, 0] * remaining, 1)
    pct_take = np.round(frac[:, 1] * remaining, 1)
    pct_open = np.round(frac[:, 2] * remaining, 1)
    pct_close = np.round(frac[:, 3] * remaining, 1)
    pct_auction = np.round(pct_open + pct_close, 1)           # %AUCTION = %OPEN + %CLOSE

    def tod(h_lo, h_hi):
        h = rng.integers(h_lo, h_hi, n)
        m = rng.integers(0, 60, n)
        s = rng.integers(0, 60, n)
        ms = rng.integers(0, 1000, n)
        return [f"{H:02d}:{M:02d}:{S:02d}.{MS:03d}" for H, M, S, MS in zip(h, m, s, ms)]

    df = pd.DataFrame({
        "client": "BLAROC.HK.A_DSA",
        "Trader": rng.choice(["A_CHAN", "B_LEE", "C_WONG", "D_NG"], n),
        "Date": pd.to_datetime(date).strftime("%Y-%m-%d"),
        "Sym": [f"{c}{n_:04d} HK" for c, n_ in zip(rng.choice(list("0"), n), rng.integers(1, 3000, n))],
        "Side": side,
        "Start(HK)": tod(9, 12),
        "End(HK)": tod(13, 16),
        "aggrTgtId": rng.integers(10_000_000, 99_999_999, n),
        "strike": np.round(price, 2),
        "Strategy": strategy,
        "Cap": rng.choice(CAPS, n, p=[0.55, 0.30, 0.15]),
        "sector": rng.choice(SECTORS, n),
        "auctionOnly": auction_only,
        "arrivalTime": rng.choice(ARRIVAL, n, p=[0.15, 0.6, 0.15, 0.10]),
        "marketLimit": market_limit,
        "#Shares (x1000)": shares,
        "$Mln (x1000000)": mln,
        "%DV": pct_dv,
        "%adv": pct_adv,
        "%POST": pct_post,
        "%TAKE": pct_take,
        "%OPEN": pct_open,
        "%CLOSE": pct_close,
        "%AUCTION": pct_auction,
        "ePR": epr,
        "FR": fr,
        "Vol": vol,
        "Sprd": sprd,
        "IS": np.round(is_bps, 2),
        "Open": np.round(open_bps, 2),
        "Close": np.round(close_bps, 2),
        "ePvwap": np.round(epvwap, 2),
        "ePvwap/Sprd": epvwap_sprd,
        "Dur": dur,
        "Pvwap": np.round(pvwap_bps, 2),
        "%DARK": dark,
    })
    return df


def make_dark_routed(seed: int) -> pd.DataFrame:
    rng = _rng(seed + 1)
    w = rng.uniform(0.5, 3, len(DARK_VENUES))
    pct = np.round(w / w.sum() * 100, 1)
    return pd.DataFrame({"Venue": DARK_VENUES, "Routed %": pct})


def make_dark_executed(seed: int) -> pd.DataFrame:
    rng = _rng(seed + 2)
    w = rng.uniform(0.5, 3, len(DARK_VENUES))
    pct = np.round(w / w.sum() * 100, 1)
    return pd.DataFrame({"Venue": DARK_VENUES, "Executed %": pct})


def make_venue_summary(seed: int) -> pd.DataFrame:
    """Dark Venue Summary table (Venue Category level)."""
    rng = _rng(seed + 3)
    cats = ["Continuous Dark", "Conditional (Block)", "Periodic Auction",
            "Close Cross", "Lit (reference)"]
    shares = rng.integers(2_000, 60_000, len(cats)) * 1000
    val = shares * rng.uniform(8, 120, len(cats))
    tot_sh, tot_val = shares.sum(), val.sum()
    return pd.DataFrame({
        "Venue Category": cats,
        "Shares": shares,
        "Exec Value": np.round(val, 0),
        "% Total Shares": np.round(shares / tot_sh * 100, 1),
        "% Total Exec Value": np.round(val / tot_val * 100, 1),
    })


def make_algo_summary(orders: pd.DataFrame) -> pd.DataFrame:
    """Performance Summary By Algorithm, aggregated from the orders (so it ties out)."""
    orders = orders.copy()
    orders["_notional"] = orders["$Mln (x1000000)"] * 1e6
    orders["_shares"] = orders["#Shares (x1000)"] * 1000
    tot = orders["_notional"].sum()
    rows = []
    for strat, g in orders.groupby("Strategy"):
        w = g["_notional"]
        wavg = lambda c: np.average(g[c], weights=w) if w.sum() else np.nan  # noqa: E731
        rows.append({
            "Strategy": strat,
            "Orders": len(g),
            "Exec Shares": int(g["_shares"].sum()),
            "Exec Value": round(float(w.sum()), 0),
            "%Weight": round(float(w.sum() / tot * 100), 1),
            "Period Part": round(float(g["ePR"].mean()), 1),
            "%ADV": round(float(wavg("%adv")), 1),
            "Spread Bps": round(float(g["Sprd"].mean()), 1),
            "Benchmark": "Order PVWAP",
            "Impact Bps": round(float(wavg("Pvwap")), 1),
            "Wgt Impact Bps": round(float(wavg("Pvwap")), 1),
            "Wgt Impact Cps": round(float(wavg("Pvwap")) / 100 * g["strike"].mean(), 3),
        })
    return pd.DataFrame(rows).sort_values("Exec Value", ascending=False)


def main() -> None:
    p = argparse.ArgumentParser(description="Generate synthetic TCA sample data")
    p.add_argument("--outdir", default="./sample_data")
    p.add_argument("--n", type=int, default=900, help="number of orders")
    p.add_argument("--seed", type=int, default=7)
    args = p.parse_args()

    out = Path(args.outdir)
    out.mkdir(parents=True, exist_ok=True)

    orders = make_orders(args.n, args.seed)
    orders.to_csv(out / "orders.csv", index=False)
    make_dark_routed(args.seed).to_csv(out / "dark_routed.csv", index=False)
    make_dark_executed(args.seed).to_csv(out / "dark_executed.csv", index=False)
    make_venue_summary(args.seed).to_csv(out / "dark_venue_summary.csv", index=False)
    make_algo_summary(orders).to_csv(out / "algo_summary.csv", index=False)

    print(f"Wrote sample data to {out.resolve()}:")
    for f in ["orders.csv", "dark_routed.csv", "dark_executed.csv",
              "dark_venue_summary.csv", "algo_summary.csv"]:
        print(f"  - {f}")


if __name__ == "__main__":
    main()
