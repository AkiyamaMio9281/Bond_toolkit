# Project 2 — Interest Rate / Bond Cash‑Flow Toolkit (v2)

This project is a **deliverable-style** bond / interest-rate toolkit:

- **Excel = front-end** (inputs + viewing tables)
- **Python = engine** (cashflows, pricing, curve scenarios, key-rate duration, exports)

## Features

### 1) Cashflows / PV / Price
- Generates a cashflow table (time, accrual, coupon, cashflow, DF, PV)
- Prices off **nominal rates j^(m)** (compounded *m* times/year)

### 2) Duration / Convexity
- **Flat-yield metrics** (match the original Excel template):
  - Price, Macaulay duration, Modified duration, discrete convexity
- **Curve analytics**:
  - Effective duration / convexity for a **parallel curve shift**
  - **Key-rate durations** (bump one curve pillar at a time)

### 3) Scenario testing (curve-aware)
- Parallel shifts: ±50bp, ±100bp
- Simple steepener / flattener:
  - Steepener: ≤2y +100bp, ≥10y +30bp (linear between)
  - Flattener: ≤2y +30bp, ≥10y +100bp (linear between)
- Optional compounding-frequency sensitivity: keep nominal j but change *m*

### 4) Stub period support (non-integer cashflow timing)
If maturity **T · p** is not an integer, the engine can create a **last stub** period:
- cashflow times: 1/p, 2/p, …, floor(Tp)/p, T
- coupon amount uses accrual fraction **dt = t_k − t_{k−1}**

(Excel template still assumes an integer schedule; stub is a Python-only extension.)

---

## Repository layout

- `tool/bond_toolkit.xlsx` — original Excel template
- `tool/bond_toolkit_v2.xlsx` — template with an optional `Curve` sheet (pillars + zero rates)
- `python/engine/` — Python pricing engine + Excel IO helpers
- `python/scripts/` — CLI tools (pricer + scenario runner)
- `python/bond_check.py` — flat-yield cross-check against the original template (reads cached values)

---

## Quick start

### A) Price + export cashflows
From `python/scripts/`:

```bash
python run_pricer.py --xlsx ../../tool/bond_toolkit_v2.xlsx --outdir ../../output --use-curve --export-xlsx
```

This writes:
- `output/cashflows.csv`
- `output/summary.json`
- `output/bond_pricer_output.xlsx` (optional)

If you want the results injected into a copy of the template:

```bash
python run_pricer.py --xlsx ../../tool/bond_toolkit_v2.xlsx --outdir ../../output --use-curve --inject-template
```

### B) Run curve scenarios + key-rate durations
```bash
python run_scenarios.py --xlsx ../../tool/bond_toolkit_v2.xlsx --outdir ../../output --use-curve --export-xlsx
```

Outputs:
- `output/scenarios.csv`
- `output/key_rate_durations.csv`
- `output/scenario_output.xlsx` (optional)

### C) Verify Excel vs Python (flat-yield)
```bash
python bond_check.py --xlsx ../tool/bond_toolkit.xlsx
```

> openpyxl reads cached formula values — if you changed the workbook, open Excel and Save once before running the check.

---

## Curve sheet format (optional)

Create a sheet named **`Curve`** with:

| A (Tenor years) | B (Zero rate j^(m)) |
|---:|---:|
| 0.5 | 0.0600 |
| 1.0 | 0.0600 |
| 2.0 | 0.0600 |
| … | … |

Interpolation is piecewise-linear in the nominal rate.

---

## Notes for portfolio write-up
- **Excel = UI**, **Python = engine**
- Curve scenarios and key-rate duration demonstrate rate-risk tooling beyond parallel shocks
- Stub-period support shows you can handle “non-tidy” real-world cashflow schedules
