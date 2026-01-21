#!/usr/bin/env python3
"""
Scenario analysis for Project 2 (v2):
- Parallel shifts
- Simple steepener/flattener shocks (tenor-dependent)
- Optional compounding-frequency sensitivity (alt m)

Also exports key-rate durations (KRD).

Examples:
  python run_scenarios.py --xlsx ../tool/bond_toolkit_v2.xlsx --outdir ../output --use-curve --export-xlsx
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from engine.bond_engine import (
    cashflows, build_flat_curve, price_from_curve, price,
    effective_duration_convexity, key_rate_durations,
    apply_tenor_dependent_shock, steepener_shock_bp,
    YieldCurve,
)
from engine.excel_io import read_inputs_xlsx, read_curve_xlsx, write_output_workbook


def make_flattener_fn(short_end=2.0, long_start=10.0, short_bp=30.0, long_bp=100.0):
    def f(t: float) -> float:
        if t <= short_end:
            return short_bp
        if t >= long_start:
            return long_bp
        w = (t - short_end) / (long_start - short_end)
        return short_bp + w * (long_bp - short_bp)
    return f


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--outdir", default="../output")
    ap.add_argument("--use-curve", action="store_true")
    ap.add_argument("--stub", action="store_true", help="Enable stub period support (default off to mirror template)")
    ap.add_argument("--export-xlsx", action="store_true", help="Export a scenario workbook")
    ap.add_argument("--krd-bp", type=float, default=1.0, help="Key-rate duration bump in bp (default 1bp)")
    args = ap.parse_args()

    xlsx_path = Path(args.xlsx).resolve()
    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    bond, y = read_inputs_xlsx(str(xlsx_path))
    curve = read_curve_xlsx(str(xlsx_path)) if args.use_curve else None
    if curve is None:
        curve = build_flat_curve(y)

    cf = cashflows(bond, stub=args.stub)
    base_cf = price_from_curve(cf, curve)
    P0 = price(base_cf)

    # use effective duration/convexity for parallel-shift approximations
    eff = effective_duration_convexity(bond, curve, bump_bp=1.0, stub=args.stub)
    D = eff["eff_duration"]
    C = eff["eff_convexity"]

    scenarios = []

    def add_scenario(name: str, curve_s: YieldCurve, dy_bp: float | None = None, note: str = ""):
        nonlocal scenarios
        P = price(price_from_curve(cf, curve_s))
        pct = (P / P0 - 1.0) * 100.0
        row = {"scenario": name, "price": P, "%chg_vs_base": pct, "note": note}
        if dy_bp is not None:
            dy = dy_bp / 10000.0
            est_dur = P0 * (1.0 - D * dy)
            est_dc = P0 * (1.0 - D * dy + 0.5 * C * dy * dy)
            row.update({
                "dy_bp": dy_bp,
                "est_price_dur": est_dur,
                "est_price_durconv": est_dc,
                "err_dur": P - est_dur,
                "err_durconv": P - est_dc,
            })
        scenarios.append(row)

    # Base
    add_scenario("Base", curve, dy_bp=0.0)

    # Parallel shifts
    for bp in [50, 100, -50, -100]:
        add_scenario(f"Parallel {bp:+}bp", curve.bump_parallel(bp), dy_bp=float(bp))

    # Steepener / flattener
    add_scenario("Steepener (≤2y +100bp, ≥10y +30bp)", apply_tenor_dependent_shock(curve, steepener_shock_bp), note="tenor-dependent shock")
    add_scenario("Flattener (≤2y +30bp, ≥10y +100bp)", apply_tenor_dependent_shock(curve, make_flattener_fn()), note="tenor-dependent shock")

    # Compounding sensitivity: keep the same nominal j values but change m
    for alt_m in [4, 1]:
        curve_m = YieldCurve(curve.pillars, curve.rates, alt_m)
        add_scenario(f"Same nominal j, m={alt_m}", curve_m, note="changes compounding frequency only")

    scen_df = pd.DataFrame(scenarios)
    scen_csv = outdir / "scenarios.csv"
    scen_df.to_csv(scen_csv, index=False)

    # Key-rate durations
    krd = key_rate_durations(bond, curve, bump_bp=args.krd_bp, stub=args.stub)
    krd_csv = outdir / "key_rate_durations.csv"
    krd.to_csv(krd_csv, index=False)

    # bundle summary as json for convenience
    summary = {
        "base_price": P0,
        "effective_duration_1bp": float(D),
        "effective_convexity_1bp": float(C),
        "stub_enabled": bool(args.stub),
        "curve_pillars": list(curve.pillars),
        "curve_m": int(curve.m),
        "krd_bump_bp": float(args.krd_bp),
    }
    (outdir / "scenario_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")

    if args.export_xlsx:
        out_xlsx = outdir / "scenario_output.xlsx"
        # re-use write_output_workbook: put scenarios + key-rate, cashflows omitted here (still in run_pricer)
        write_output_workbook(str(out_xlsx), summary=summary, cashflows=base_cf, scenarios=scen_df, key_rate=krd)

    print("Done.")
    print(f"- Scenarios: {scen_csv}")
    print(f"- Key-rate durations: {krd_csv}")
    if args.export_xlsx:
        print(f"- Output workbook: {outdir / 'scenario_output.xlsx'}")


if __name__ == "__main__":
    main()
