# Bhav Analyzer — Technique Deep Dive

## Overview

Bhav Analyzer is built around one core idea: **tracking where institutional money is moving in the Nifty options market using Open Interest (OI) data from NSE's daily bhav copy.**

It doesn't use price action, indicators, or charts. Instead, it follows the footprint of large participants through OI changes — specifically the **monetary value** of those changes, not just contract counts.

---

## The Central Technique: OI Value Analysis

### Why Value, Not Just Contracts?

Raw OI (number of contracts) can be misleading because a contract at strike 18000 is worth far more than one at strike 10000. The app converts OI into **rupee value** to make comparisons meaningful:

```
EOD Value       = Open Interest × Close Price
OI Change Value = Change in Open Interest × Close Price
```

This gives a crore-denominated view of how much money is sitting in each strike, and how much was added or removed on that day.

---

## The ATM/ITM Split — The Core Signal

The most important logic in the app is splitting contracts into **All strikes** vs **ITM (In-The-Money) strikes**.

### ITM Definition Used Here

| Option Type | ITM Condition         | Interpretation                        |
|-------------|-----------------------|---------------------------------------|
| CE (Call)   | Strike Price < Spot   | Calls that are already profitable     |
| PE (Put)    | Strike Price > Spot   | Puts that are already profitable      |

### Why Does This Matter?

OTM (Out-of-the-Money) options are cheap and heavily traded by retail participants for directional bets. **ITM options, however, are expensive and mostly held or written by institutional players** — hedgers, market makers, and large funds.

By isolating ITM OI separately, the app answers:
> "Of all the money in the options market today, how much is sitting in contracts that are already in profit — and is that money increasing or decreasing?"

---

## What Each Metric Tells You

### CE Side (Calls)

| Metric          | Formula                          | Signal                                                   |
|-----------------|----------------------------------|----------------------------------------------------------|
| `CE_EOD`        | Sum of all Call EOD values       | Total call OI in crores — size of the call side          |
| `CE_CHANGE`     | Sum of all Call OI change values | Net money added/removed in calls today                   |
| `CE_%CHANGE`    | CE_CHANGE / CE_EOD × 100         | % shift in call OI — velocity of change                 |
| `ITM_CE_EOD`    | Sum of ITM call EOD values       | Money locked in deep/ITM calls                           |
| `ITM_CE_%EOD`   | ITM_CE_EOD / CE_EOD × 100        | ITM share of total call OI — how "heavy" the call side is|
| `ITM_CE_CHANGE` | Sum of ITM call OI change values | Money added/removed specifically in ITM calls            |
| `ITM_CE_%CHANGE`| ITM_CE_CHANGE / ITM_CE_EOD × 100 | Rate of change within ITM calls                          |

### PE Side (Puts)

Exact mirror of the CE side, but for puts with the ITM condition reversed (Strike > Spot).

---

## Reading the Signals

### Bullish Signs
- `CE_CHANGE` is **negative** (call writers booking profit / shorts closing) → resistance weakening
- `PE_CHANGE` is **positive** (put writers adding positions) → support being built
- `ITM_PE_%EOD` rising → more money defending lower levels

### Bearish Signs
- `CE_CHANGE` is **positive** (new call writing above spot) → resistance being built
- `PE_CHANGE` is **negative** (put longs closing / put writers exiting) → support weakening
- `ITM_CE_%EOD` rising → more money hedging or shorting via ITM calls

### Neutral / Range-Bound
- Both `CE_CHANGE` and `PE_CHANGE` strongly positive → both sides building → market pinned in a range

---

## Expiry Segmentation

The app tracks **4 expiry buckets** separately:

| Sheet    | Instrument  | Expiry Type | Why Separate?                                              |
|----------|-------------|-------------|-------------------------------------------------------------|
| Nifty-W  | NIFTY       | Weekly      | Short-term sentiment, intraday/swing traders                |
| Nifty-M  | NIFTY       | Monthly     | Positional view, institutional hedging                      |
| Bank-M   | BANKNIFTY   | Monthly     | Banking sector specific OI flow                             |
| Fin-M    | FINNIFTY    | Monthly     | Financial sector specific OI flow                           |

Weekly and monthly expiries often tell different stories. A weekly expiry may show aggressive call writing (short-term resistance) while the monthly expiry shows put writing (longer-term bullishness). Comparing both gives a fuller picture.

---

## Historical Accumulation

Each time the app runs for a new trading day, the new row is **appended** to the existing master Excel file. Over time this builds a day-by-day time series that lets you:

- Spot trends in OI buildup over days/weeks
- See if ITM OI is steadily increasing (sustained directional pressure)
- Compare how OI behaved before large moves vs quiet days

---

## Data Source

The input is the **NSE Bhav Copy** — the official end-of-day data file published by NSE every trading day. It contains closing prices and OI for every F&O contract traded that day. The filename typically contains the date in `YYYYMMDD` format which the app uses to auto-name the output file.

---

## Limitations

- This is **end-of-day analysis only** — intraday OI shifts are not captured
- The app does not account for **rollover activity** near expiry (OI may drop due to rolls, not market direction)
- No price trend context is built in — OI signals should be read alongside price action for best results
- ITM/OTM classification is based on the **spot value entered manually** by the user, so accuracy depends on entering the correct spot

---

## Summary

```
NSE Bhav Copy (CSV)
        ↓
Filter: Index Options (IDO) → NIFTY / BANKNIFTY / FINNIFTY
        ↓
Filter by selected expiry date → CE and PE separately
        ↓
Compute: EOD Value = OI × Close Price
         OI Change Value = ΔOI × Close Price
        ↓
Split: All strikes vs ITM strikes (relative to entered spot)
        ↓
Aggregate into crore values + % metrics
        ↓
Append to master Excel → historical time series
```

The technique is essentially **options flow analysis via OI monetization** — converting raw contract data into rupee terms and isolating institutional-grade ITM activity from the broader market noise.
