#!/usr/bin/env python3
"""
Run the Python pricer engine from Excel inputs and export results.

Examples:
  python run_pricer.py --xlsx ../tool/bond_toolkit_v2.xlsx --outdir ../output --use-curve --export-xlsx
  python run_pricer.py --xlsx ../tool/bond_toolkit_v2.xlsx --outdir ../output --inject-template

Outputs:
  - cashflows.csv (includes df and pv columns)
  - summary.json
  - optional: bond_pricer_output.xlsx
  - optional: bond_toolkit_v2_with_python_outputs.xlsx  (if --inject-template)
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

import pandas as pd

# Make imports work when running as a script
import sys
sys.path.append(str(Path(__file__).resolve().parents[1]))

from engine.bond_engine import (
    BondSpec, FlatYield, YieldCurve,
    cashflows, build_flat_curve,
    price_from_curve, price, macaulay_duration,
    modified_duration_flat, convexity_flat,
    effective_duration_convexity,
)
from engine.excel_io import read_inputs_xlsx, read_curve_xlsx, write_output_workbook, inject_python_outputs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, help="Path to Excel template/workbook")
    ap.add_argument("--outdir", default="../output", help="Output folder")
    ap.add_argument("--use-curve", action="store_true", help="Use Curve sheet if present (else fall back to flat)")
    ap.add_argument("--stub", action="store_true", help="Enable stub period support (default off to mirror template)")
    ap.add_argument("--export-xlsx", action="store_true", help="Export a standalone output workbook")
    ap.add_argument("--inject-template", action="store_true", help="Write outputs into a copy of the template (Python Outputs sheet)")
    args = ap.parse_args()

    xlsx_path = Path(args.xlsx).resolve()
    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    bond, y = read_inputs_xlsx(str(xlsx_path))
    curve = None
    if args.use_curve:
        curve = read_curve_xlsx(str(xlsx_path))
    if curve is None:
        curve = build_flat_curve(y)

    cf = cashflows(bond, stub=args.stub)
    cf_pv = price_from_curve(cf, curve)
    P = price(cf_pv)
    mac = macaulay_duration(cf_pv)

    # flat-yield metrics are only "canonical" for a flat curve (or when you want to report them anyway)
    # We compute them using the nominal rate at t = maturity (for flat this equals y.j).
    flat = FlatYield(j=curve.zero_rate(bond.maturity), m=curve.m)
    mod = modified_duration_flat(mac, flat)
    conv = convexity_flat(cf_pv, flat)

    eff = effective_duration_convexity(bond, curve, bump_bp=1.0, stub=args.stub)

    summary = {
        "fv": bond.fv,
        "coupon_rate": bond.coupon_rate,
        "maturity": bond.maturity,
        "coupon_freq": bond.coupon_freq,
        "curve_m": curve.m,
        "price": P,
        "macaulay_duration": mac,
        "modified_duration": mod,
        "convexity": conv,
        "effective_duration_1bp": eff["eff_duration"],
        "effective_convexity_1bp": eff["eff_convexity"],
        "stub_enabled": bool(args.stub),
        "n_cashflows": int(len(cf_pv)),
    }

    # exports
    cf_csv = outdir / "cashflows.csv"
    cf_pv.to_csv(cf_csv, index=False)

    summary_json = outdir / "summary.json"
    summary_json.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    if args.export_xlsx:
        out_xlsx = outdir / "bond_pricer_output.xlsx"
        write_output_workbook(str(out_xlsx), summary=summary, cashflows=cf_pv)

    if args.inject_template:
        out_xlsx2 = outdir / f"{xlsx_path.stem}_with_python_outputs.xlsx"
        inject_python_outputs(
            xlsx_in=str(xlsx_path),
            xlsx_out=str(out_xlsx2),
            summary=summary,
            cashflows=cf_pv,
        )

    print("Done.")
    print(f"- Cashflows: {cf_csv}")
    print(f"- Summary  : {summary_json}")
    if args.export_xlsx:
        print(f"- Output workbook: {outdir / 'bond_pricer_output.xlsx'}")
    if args.inject_template:
        print(f"- Template copy with outputs: {outdir / f'{xlsx_path.stem}_with_python_outputs.xlsx'}")


if __name__ == "__main__":
    main()
