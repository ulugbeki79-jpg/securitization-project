"""Tranche definitions and waterfall distribution engine.

Implements a simple sequential-pay (senior / subordinated) MBS structure:

    Available Funds
        │
        ├─► Servicing Fee          (already deducted in cashflows.py)
        ├─► Class A Interest       (senior)
        ├─► Class A Principal      (senior)
        ├─► Class B Interest       (subordinated)
        ├─► Class B Principal      (subordinated)
        └─► Residual / Excess spread

Losses are allocated bottom-up: Class B absorbs first, then Class A.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd


@dataclass
class TrancheSpec:
    """Specification for a single tranche."""

    name: str
    balance: float          # initial principal (UZS)
    coupon: float           # annual coupon rate (decimal, e.g. 0.16)
    is_senior: bool = True  # determines loss allocation priority


@dataclass
class TrancheState:
    """Mutable state tracked month-over-month for a tranche."""

    name: str
    beginning_balance: float = 0.0
    interest_due: float = 0.0
    interest_paid: float = 0.0
    interest_shortfall: float = 0.0
    principal_paid: float = 0.0
    loss_allocated: float = 0.0
    ending_balance: float = 0.0


def run_waterfall(
    pool_cf: pd.DataFrame,
    tranche_a: TrancheSpec,
    tranche_b: TrancheSpec,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Distribute collateral cash flows through the senior/sub waterfall.

    Parameters
    ----------
    pool_cf : pd.DataFrame
        Output of :func:`cashflows.project_pool_cashflows`.
    tranche_a : TrancheSpec
        Senior tranche (Class A).
    tranche_b : TrancheSpec
        Subordinated tranche (Class B).

    Returns
    -------
    (class_a_df, class_b_df)
        DataFrames with month-by-month tranche-level cash flows.
    """
    a_bal = tranche_a.balance
    b_bal = tranche_b.balance
    a_shortfall = 0.0
    b_shortfall = 0.0

    a_rows: list[dict] = []
    b_rows: list[dict] = []

    for _, row in pool_cf.iterrows():
        month = int(row["month"])
        available_interest = row["net_interest"] + row["recovery"]
        available_principal = row["total_principal"]
        period_loss = row["loss"]

        # ------------------------------------------------------------------
        # 1. Loss allocation — bottom-up (Class B first, then A)
        # ------------------------------------------------------------------
        b_loss = min(period_loss, b_bal)
        remaining_loss = period_loss - b_loss
        a_loss = min(remaining_loss, a_bal)

        b_bal -= b_loss
        a_bal -= a_loss

        # ------------------------------------------------------------------
        # 2. Interest waterfall
        # ------------------------------------------------------------------
        a_int_due = a_bal * tranche_a.coupon / 12.0 + a_shortfall
        a_int_paid = min(a_int_due, available_interest)
        a_shortfall = a_int_due - a_int_paid
        available_interest -= a_int_paid

        b_int_due = b_bal * tranche_b.coupon / 12.0 + b_shortfall
        b_int_paid = min(b_int_due, available_interest)
        b_shortfall = b_int_due - b_int_paid
        available_interest -= b_int_paid

        # ------------------------------------------------------------------
        # 3. Principal waterfall — sequential pay (A before B)
        # ------------------------------------------------------------------
        a_prin_paid = min(available_principal, a_bal)
        available_principal -= a_prin_paid
        a_bal -= a_prin_paid

        b_prin_paid = min(available_principal, b_bal)
        available_principal -= b_prin_paid
        b_bal -= b_prin_paid

        excess_spread = available_interest + available_principal

        # ------------------------------------------------------------------
        # Record
        # ------------------------------------------------------------------
        a_rows.append(
            {
                "month": month,
                "beginning_balance": round(a_bal + a_prin_paid + a_loss, 2),
                "interest_due": round(a_int_due, 2),
                "interest_paid": round(a_int_paid, 2),
                "interest_shortfall": round(a_shortfall, 2),
                "principal_paid": round(a_prin_paid, 2),
                "loss_allocated": round(a_loss, 2),
                "ending_balance": round(a_bal, 2),
                "total_cash_flow": round(a_int_paid + a_prin_paid, 2),
            }
        )
        b_rows.append(
            {
                "month": month,
                "beginning_balance": round(b_bal + b_prin_paid + b_loss, 2),
                "interest_due": round(b_int_due, 2),
                "interest_paid": round(b_int_paid, 2),
                "interest_shortfall": round(b_shortfall, 2),
                "principal_paid": round(b_prin_paid, 2),
                "loss_allocated": round(b_loss, 2),
                "ending_balance": round(b_bal, 2),
                "total_cash_flow": round(b_int_paid + b_prin_paid, 2),
                "excess_spread": round(excess_spread, 2),
            }
        )

    return pd.DataFrame(a_rows), pd.DataFrame(b_rows)
