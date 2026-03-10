"""Project monthly loan pool cash flows under CPR/CDR/servicing assumptions.

All monetary amounts are in UZS.

Terminology
-----------
- CPR  Conditional Prepayment Rate (annualised voluntary prepayment speed)
- CDR  Conditional Default Rate (annualised default rate)
- SMM  Single Monthly Mortality = 1 - (1 - CPR)^(1/12)
- MDR  Monthly Default Rate     = 1 - (1 - CDR)^(1/12)
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class PoolAssumptions:
    """Assumptions that drive the cash flow projection."""

    cpr: float              # annualised CPR, e.g. 0.08 = 8 %
    cdr: float              # annualised CDR, e.g. 0.02 = 2 %
    recovery_rate: float    # % of defaulted principal recovered, e.g. 0.40
    recovery_lag_months: int  # months until recovery proceeds arrive
    servicing_fee_pct: float  # annual servicing fee as % of outstanding balance


def _smm(cpr: float) -> float:
    """Convert annualised CPR to Single Monthly Mortality."""
    return 1.0 - (1.0 - cpr) ** (1.0 / 12.0)


def _mdr(cdr: float) -> float:
    """Convert annualised CDR to Monthly Default Rate."""
    return 1.0 - (1.0 - cdr) ** (1.0 / 12.0)


def project_pool_cashflows(
    beginning_balance: float,
    wac: float,
    wam: int,
    assumptions: PoolAssumptions,
) -> pd.DataFrame:
    """Generate a month-by-month cash flow schedule for the collateral pool.

    Parameters
    ----------
    beginning_balance : float
        Total outstanding principal at time-0 (UZS).
    wac : float
        Weighted average coupon as a decimal (e.g. 0.19 for 19 %).
    wam : int
        Weighted average remaining maturity in months.
    assumptions : PoolAssumptions
        Prepayment / default / servicing assumptions.

    Returns
    -------
    pd.DataFrame
        One row per month with columns described below.
    """
    monthly_rate = wac / 12.0
    smm = _smm(assumptions.cpr)
    mdr = _mdr(assumptions.cdr)
    monthly_svc = assumptions.servicing_fee_pct / 12.0

    rows: list[dict] = []
    balance = beginning_balance
    # Queue of (month_available, amount) for delayed recoveries
    recovery_queue: list[tuple[int, float]] = []

    for month in range(1, wam + 1):
        if balance <= 0.01:
            break

        # Scheduled payment (level-pay amortisation)
        remaining_months = wam - month + 1
        if monthly_rate > 0:
            pmt = balance * monthly_rate / (1.0 - (1.0 + monthly_rate) ** -remaining_months)
        else:
            pmt = balance / remaining_months

        scheduled_interest = balance * monthly_rate
        scheduled_principal = pmt - scheduled_interest

        # Defaults (applied to balance after scheduled principal)
        default_amount = (balance - scheduled_principal) * mdr

        # Prepayments (applied to surviving balance after defaults)
        surviving = balance - scheduled_principal - default_amount
        prepayment = surviving * smm

        # Total principal received this month
        total_principal = scheduled_principal + prepayment

        # Losses
        loss = default_amount * (1.0 - assumptions.recovery_rate)

        # Queue recovery proceeds
        recovered = default_amount * assumptions.recovery_rate
        if recovered > 0:
            recovery_queue.append((month + assumptions.recovery_lag_months, recovered))

        # Collect any recoveries arriving this month
        recovery_received = sum(amt for m, amt in recovery_queue if m == month)

        # Servicing fee (deducted from interest)
        servicing_fee = balance * monthly_svc
        net_interest = scheduled_interest - servicing_fee

        # Ending balance
        ending_balance = balance - total_principal - default_amount

        rows.append(
            {
                "month": month,
                "beginning_balance": round(balance, 2),
                "scheduled_pmt": round(pmt, 2),
                "scheduled_principal": round(scheduled_principal, 2),
                "scheduled_interest": round(scheduled_interest, 2),
                "prepayment": round(prepayment, 2),
                "default_amount": round(default_amount, 2),
                "loss": round(loss, 2),
                "recovery": round(recovery_received, 2),
                "servicing_fee": round(servicing_fee, 2),
                "net_interest": round(net_interest, 2),
                "total_principal": round(total_principal, 2),
                "ending_balance": round(max(ending_balance, 0), 2),
            }
        )

        balance = max(ending_balance, 0)

    return pd.DataFrame(rows)
