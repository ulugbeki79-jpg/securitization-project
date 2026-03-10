"""Tranche-level analytics: WAL, yield, credit enhancement."""

from __future__ import annotations

import numpy as np
import pandas as pd


def weighted_average_life(tranche_cf: pd.DataFrame) -> float:
    """Compute Weighted Average Life (WAL) in years.

    WAL = Σ (t_i × Principal_i) / Σ Principal_i

    where t_i is the time in years of each principal payment.
    """
    principal = tranche_cf["principal_paid"]
    total_principal = principal.sum()
    if total_principal == 0:
        return 0.0
    time_years = tranche_cf["month"] / 12.0
    return round(float((time_years * principal).sum() / total_principal), 2)


def tranche_yield(
    tranche_cf: pd.DataFrame,
    initial_balance: float,
    price: float = 1.0,
) -> float:
    """Estimate the bond-equivalent yield (IRR) for a tranche.

    Parameters
    ----------
    tranche_cf : pd.DataFrame
        Month-by-month tranche cash flows (needs ``total_cash_flow``).
    initial_balance : float
        Original face amount of the tranche (UZS).
    price : float
        Purchase price as a fraction of par (1.0 = par).

    Returns
    -------
    float
        Annualised yield as a percentage (e.g. 18.5).
        Returns 0.0 if the solver fails.
    """
    invest = -initial_balance * price
    cfs = tranche_cf["total_cash_flow"].values.astype(float)
    cash_flows = np.concatenate([[invest], cfs])

    # np.irr was removed; use np.polynomial approach
    try:
        monthly_irr = _solve_irr(cash_flows)
        annual_yield = ((1.0 + monthly_irr) ** 12 - 1.0) * 100.0
        return round(annual_yield, 4)
    except Exception:
        return 0.0


def _solve_irr(cash_flows: np.ndarray, tol: float = 1e-10, max_iter: int = 500) -> float:
    """Newton-Raphson IRR solver for monthly cash flows."""
    # Initial guess from simple return
    total_in = -cash_flows[0]
    total_out = cash_flows[1:].sum()
    n = len(cash_flows) - 1
    if total_in <= 0 or n == 0:
        return 0.0
    guess = (total_out / total_in) ** (1.0 / n) - 1.0
    guess = max(min(guess, 0.5), -0.5)  # clamp

    rate = guess
    for _ in range(max_iter):
        npv = 0.0
        d_npv = 0.0
        for t, cf in enumerate(cash_flows):
            discount = (1.0 + rate) ** t
            npv += cf / discount
            if t > 0:
                d_npv -= t * cf / ((1.0 + rate) ** (t + 1))
        if abs(d_npv) < 1e-20:
            break
        step = npv / d_npv
        rate -= step
        if abs(step) < tol:
            break
    return rate


def credit_enhancement_pct(sub_balance: float, total_balance: float) -> float:
    """Credit enhancement for the senior tranche as a percentage.

    CE% = subordinated balance / total collateral balance × 100
    """
    if total_balance == 0:
        return 0.0
    return round(sub_balance / total_balance * 100.0, 2)


def summary_metrics(
    tranche_cf: pd.DataFrame,
    initial_balance: float,
    label: str,
) -> dict:
    """Compute all analytics for one tranche and return as a dict."""
    wal = weighted_average_life(tranche_cf)
    yld = tranche_yield(tranche_cf, initial_balance)
    total_interest = tranche_cf["interest_paid"].sum()
    total_principal = tranche_cf["principal_paid"].sum()
    total_loss = tranche_cf["loss_allocated"].sum()
    return {
        "tranche": label,
        "initial_balance_uzs": initial_balance,
        "wal_years": wal,
        "yield_pct": yld,
        "total_interest_uzs": round(total_interest, 2),
        "total_principal_uzs": round(total_principal, 2),
        "total_loss_uzs": round(total_loss, 2),
        "final_balance_uzs": round(tranche_cf["ending_balance"].iloc[-1], 2),
    }
