#!/usr/bin/env python3
"""
Bond pricer engine for Project 2 (v2).

Core features:
- Cash-flow generation with optional stub period (non-integer T * p).
- Pricing off a nominal zero-rate curve j^(m) with piecewise-linear interpolation.
- Flat-yield metrics (Price, Macaulay, Modified, discrete Convexity) consistent with the original Excel template.
- Curve analytics: effective duration/convexity (parallel curve shift) + key-rate durations.

Conventions:
- Rates are nominal j convertible m times per year (j^(m)).
- Discount factor: DF(t) = (1 + j(t)/m)^(-m t)
- Time t is in years.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence, Tuple, Dict
import math
import pandas as pd


@dataclass(frozen=True)
class BondSpec:
    fv: float
    coupon_rate: float   # annual coupon rate c (as decimal)
    maturity: float      # years, T
    coupon_freq: int     # payments per year, p


@dataclass(frozen=True)
class FlatYield:
    j: float             # nominal rate j^(m), decimal
    m: int               # compounding times/year


@dataclass(frozen=True)
class YieldCurve:
    """
    Nominal zero-rate curve j^(m) at pillars (tenors in years).
    Interpolation: piecewise-linear in the nominal rate.
    Extrapolation: flat beyond the end pillars.
    """
    pillars: Tuple[float, ...]
    rates: Tuple[float, ...]
    m: int

    def __post_init__(self):
        if len(self.pillars) != len(self.rates):
            raise ValueError("pillars and rates must have the same length")
        if len(self.pillars) < 1:
            raise ValueError("Need at least one pillar")
        if any(t <= 0 for t in self.pillars):
            raise ValueError("All pillars must be > 0 (years)")
        if any(self.pillars[i] >= self.pillars[i+1] for i in range(len(self.pillars)-1)):
            raise ValueError("pillars must be strictly increasing")
        if self.m <= 0:
            raise ValueError("m must be positive")

    def zero_rate(self, t: float) -> float:
        """Nominal zero rate j^(m) at time t (years)."""
        if t <= self.pillars[0]:
            return self.rates[0]
        if t >= self.pillars[-1]:
            return self.rates[-1]

        # find segment
        # linear interpolation between (t_i, r_i) and (t_{i+1}, r_{i+1})
        for i in range(len(self.pillars)-1):
            t0, t1 = self.pillars[i], self.pillars[i+1]
            if t0 <= t <= t1:
                r0, r1 = self.rates[i], self.rates[i+1]
                w = (t - t0) / (t1 - t0)
                return r0 + w * (r1 - r0)
        # should never hit
        return self.rates[-1]

    def df(self, t: float) -> float:
        r = self.zero_rate(t)
        base = 1.0 + r / self.m
        if base <= 0:
            raise ValueError(f"Invalid rate/compounding: 1 + r/m <= 0 at t={t}, r={r}, m={self.m}")
        return base ** (-self.m * t)

    def bump_parallel(self, bp: float) -> "YieldCurve":
        """Parallel shift by bp basis points."""
        d = bp / 10000.0
        return YieldCurve(self.pillars, tuple(r + d for r in self.rates), self.m)

    def bump_key(self, idx: int, bp: float) -> "YieldCurve":
        """Bump a single pillar by bp basis points."""
        if not (0 <= idx < len(self.rates)):
            raise IndexError("idx out of range")
        d = bp / 10000.0
        new_rates = list(self.rates)
        new_rates[idx] += d
        return YieldCurve(self.pillars, tuple(new_rates), self.m)


def cashflows(bond: BondSpec, stub: bool = True) -> pd.DataFrame:
    """
    Generate cash-flow schedule from t=0 to maturity.

    If stub=True and T*p is not an integer, we use a *last* stub period:
      times: 1/p, 2/p, ..., n/p, T where n=floor(T*p)
      accrual: dt_k = t_k - t_{k-1}
      coupon_k = FV * c * dt_k
    """
    fv, c, T, p = bond.fv, bond.coupon_rate, bond.maturity, bond.coupon_freq
    if fv <= 0 or T <= 0:
        raise ValueError("fv and maturity must be positive")
    if p <= 0:
        raise ValueError("coupon_freq p must be positive integer")
    if c < 0:
        raise ValueError("coupon_rate cannot be negative")

    n_exact = T * p
    n_full = int(math.floor(n_exact + 1e-12))
    times: List[float] = []
    for k in range(1, n_full + 1):
        times.append(k / p)

    if stub:
        if abs(n_exact - n_full) > 1e-10:
            # add final maturity time if it's not already exactly included
            if times and abs(times[-1] - T) < 1e-12:
                pass
            else:
                times.append(T)
        else:
            # exact schedule: last time already equals T
            if not times or abs(times[-1] - T) > 1e-12:
                times.append(T)
    else:
        # force integer number of periods
        T_used = n_full / p
        times = [k / p for k in range(1, n_full + 1)]
        if abs(times[-1] - T_used) > 1e-12:
            times.append(T_used)

    rows = []
    t_prev = 0.0
    for i, t in enumerate(times, start=1):
        dt = t - t_prev
        if dt <= 0:
            raise ValueError("Non-increasing time grid generated")
        coupon = fv * c * dt
        cf = coupon
        if abs(t - T) < 1e-10:
            cf += fv
        rows.append({"k": i, "t": float(t), "accrual": float(dt), "cashflow": float(cf), "coupon": float(coupon)})
        t_prev = t

    return pd.DataFrame(rows)


def price_from_curve(cf: pd.DataFrame, curve: YieldCurve) -> pd.DataFrame:
    """Add DF and PV columns, return the enriched CF table."""
    df = cf.copy()
    df["df"] = df["t"].map(curve.df)
    df["pv"] = df["cashflow"] * df["df"]
    return df


def price(cf_pv: pd.DataFrame) -> float:
    return float(cf_pv["pv"].sum())


def macaulay_duration(cf_pv: pd.DataFrame) -> float:
    P = price(cf_pv)
    if P == 0:
        return float("nan")
    return float((cf_pv["t"] * cf_pv["pv"]).sum() / P)


def modified_duration_flat(mac_dur: float, y: FlatYield) -> float:
    # For nominal j^(m): modified duration = MacDur / (1 + j/m)
    return float(mac_dur / (1.0 + y.j / y.m))


def convexity_flat(cf_pv: pd.DataFrame, y: FlatYield) -> float:
    """
    Discrete convexity consistent with the original Excel template for nominal j^(m).

    Formula used:
      Conv = (1/P) * Σ PV_k * t_k * (t_k + 1/m) / (1 + j/m)^2
    """
    P = price(cf_pv)
    if P == 0:
        return float("nan")
    a = 1.0 + y.j / y.m
    return float(((cf_pv["pv"] * cf_pv["t"] * (cf_pv["t"] + 1.0 / y.m)).sum() / P) / (a * a))


def effective_duration_convexity(bond: BondSpec, curve: YieldCurve, bump_bp: float = 1.0, stub: bool = True) -> Dict[str, float]:
    """
    Effective duration and convexity for a *parallel* curve shift by +/- bump_bp.
    bump_bp default 1bp for stability; you can pass 50/100bp for scenario alignment.
    """
    cf = cashflows(bond, stub=stub)
    base = price(price_from_curve(cf, curve))
    up = price(price_from_curve(cf, curve.bump_parallel(+bump_bp)))
    dn = price(price_from_curve(cf, curve.bump_parallel(-bump_bp)))
    dy = bump_bp / 10000.0

    eff_dur = (dn - up) / (2.0 * base * dy)
    eff_conv = (up + dn - 2.0 * base) / (base * dy * dy)
    return {"price": base, "eff_duration": float(eff_dur), "eff_convexity": float(eff_conv)}


def key_rate_durations(bond: BondSpec, curve: YieldCurve, bump_bp: float = 1.0, stub: bool = True) -> pd.DataFrame:
    """
    Key-rate durations (KRD) by bumping one curve pillar at a time.
    Returns a dataframe with columns: pillar, krd
    """
    cf = cashflows(bond, stub=stub)
    base = price(price_from_curve(cf, curve))
    dy = bump_bp / 10000.0
    rows = []
    for i, ten in enumerate(curve.pillars):
        bumped = curve.bump_key(i, bump_bp)
        pb = price(price_from_curve(cf, bumped))
        krd = -(pb - base) / (base * dy)
        rows.append({"pillar": float(ten), "krd": float(krd)})
    return pd.DataFrame(rows)


def build_flat_curve(y: FlatYield, pillars: Sequence[float] = (0.5, 1, 2, 5, 10, 20, 30)) -> YieldCurve:
    return YieldCurve(tuple(float(t) for t in pillars), tuple(float(y.j) for _ in pillars), y.m)


def steepener_shock_bp(t: float, short_end: float = 2.0, long_start: float = 10.0, short_bp: float = 100.0, long_bp: float = 30.0) -> float:
    """
    Piecewise-linear bump (bp) used for a simple curve steepener:
      t <= short_end      : +short_bp
      t >= long_start     : +long_bp
      short_end..long_start: linear between
    """
    if t <= short_end:
        return short_bp
    if t >= long_start:
        return long_bp
    w = (t - short_end) / (long_start - short_end)
    return short_bp + w * (long_bp - short_bp)


def apply_tenor_dependent_shock(curve: YieldCurve, shock_fn) -> YieldCurve:
    new_rates = []
    for ten, r in zip(curve.pillars, curve.rates):
        new_rates.append(r + shock_fn(float(ten)) / 10000.0)
    return YieldCurve(curve.pillars, tuple(new_rates), curve.m)
