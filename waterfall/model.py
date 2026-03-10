"""MBS cash flow waterfall model — main entry point.

Usage:
    python -m waterfall.model sample_data/loan_pool.xlsx
    python -m waterfall.model sample_data/loan_pool.xlsx --cpr 0.10 --cdr 0.03
    python -m waterfall.model sample_data/loan_pool.xlsx -o my_output/

This reads the loan pool, sizes tranches, projects cash flows, runs the
waterfall, computes analytics, and writes a multi-sheet Excel workbook.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from waterfall.cashflows import PoolAssumptions, project_pool_cashflows
from waterfall.tranches import TrancheSpec, run_waterfall
from waterfall.analytics import (
    credit_enhancement_pct,
    summary_metrics,
    weighted_average_life,
)
from waterfall.excel_report import write_waterfall_excel


# ---------------------------------------------------------------------------
# Pool helpers (reuse pool_analysis metrics if available)
# ---------------------------------------------------------------------------

def _pool_summary(df: pd.DataFrame) -> dict:
    """Compute quick pool stats needed for the waterfall."""
    total = df["principal_uzs"].sum()
    w = df["principal_uzs"] / total
    return {
        "total_balance": total,
        "wac": (w * df["coupon_pct"]).sum() / 100.0,      # as decimal
        "wam": int(round((w * df["remaining_term_months"]).sum())),
        "loan_count": len(df),
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run(
    excel_path: str,
    out_dir: str = "output",
    senior_pct: float = 0.80,
    senior_coupon: float = 0.16,
    sub_coupon: float = 0.20,
    cpr: float = 0.08,
    cdr: float = 0.02,
    recovery_rate: float = 0.40,
    recovery_lag: int = 6,
    servicing_fee: float = 0.005,
) -> None:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    # 1. Read pool
    print(f"Reading loan pool from {excel_path} ...")
    df = pd.read_excel(excel_path, sheet_name=0)
    pool = _pool_summary(df)
    print(f"  {pool['loan_count']} loans, total balance {pool['total_balance']:,.0f} UZS")
    print(f"  WAC = {pool['wac'] * 100:.2f}%, WAM = {pool['wam']} months")

    # 2. Size tranches
    total = pool["total_balance"]
    a_balance = round(total * senior_pct)
    b_balance = round(total * (1 - senior_pct))

    tranche_a = TrancheSpec(name="Class A (Senior)", balance=a_balance, coupon=senior_coupon, is_senior=True)
    tranche_b = TrancheSpec(name="Class B (Sub)", balance=b_balance, coupon=sub_coupon, is_senior=False)

    ce = credit_enhancement_pct(b_balance, total)
    print(f"\n  Class A: {a_balance:,.0f} UZS @ {senior_coupon * 100:.1f}%")
    print(f"  Class B: {b_balance:,.0f} UZS @ {sub_coupon * 100:.1f}%")
    print(f"  Credit Enhancement: {ce:.2f}%")

    # 3. Project collateral cash flows
    assumptions = PoolAssumptions(
        cpr=cpr,
        cdr=cdr,
        recovery_rate=recovery_rate,
        recovery_lag_months=recovery_lag,
        servicing_fee_pct=servicing_fee,
    )

    print(f"\nProjecting cash flows (CPR={cpr*100:.1f}%, CDR={cdr*100:.1f}%) ...")
    pool_cf = project_pool_cashflows(
        beginning_balance=total,
        wac=pool["wac"],
        wam=pool["wam"],
        assumptions=assumptions,
    )
    print(f"  Generated {len(pool_cf)} monthly periods")

    # 4. Run waterfall
    print("Running waterfall distribution ...")
    a_cf, b_cf = run_waterfall(pool_cf, tranche_a, tranche_b)

    # 5. Analytics
    a_metrics = summary_metrics(a_cf, a_balance, "Class A (Senior)")
    b_metrics = summary_metrics(b_cf, b_balance, "Class B (Sub)")

    print(f"\n{'Tranche':<20} {'WAL (yr)':>10} {'Yield (%)':>10} {'Loss (UZS)':>18}")
    print("-" * 60)
    for m in [a_metrics, b_metrics]:
        print(f"{m['tranche']:<20} {m['wal_years']:>10.2f} {m['yield_pct']:>10.2f} {m['total_loss_uzs']:>18,.0f}")

    # 6. Write Excel
    assumptions_text = [
        f"Pool balance: {total:,.0f} UZS ({pool['loan_count']} loans)",
        f"WAC: {pool['wac'] * 100:.2f}%  |  WAM: {pool['wam']} months",
        f"Senior %: {senior_pct * 100:.0f}%  |  Sub %: {(1 - senior_pct) * 100:.0f}%",
        f"Class A coupon: {senior_coupon * 100:.1f}%  |  Class B coupon: {sub_coupon * 100:.1f}%",
        f"CPR: {cpr * 100:.1f}%  |  CDR: {cdr * 100:.1f}%",
        f"Recovery rate: {recovery_rate * 100:.0f}%  |  Recovery lag: {recovery_lag} months",
        f"Servicing fee: {servicing_fee * 100:.2f}%",
        f"Credit enhancement (CE%): {ce:.2f}%",
    ]

    excel_path_out = out / "waterfall_output.xlsx"
    print(f"\nWriting Excel report to {excel_path_out} ...")
    write_waterfall_excel(pool_cf, a_cf, b_cf, [a_metrics, b_metrics], assumptions_text, excel_path_out)
    print("Done.")


def main():
    p = argparse.ArgumentParser(description="MBS Cash Flow Waterfall Model")
    p.add_argument("excel_file", help="Loan pool Excel file")
    p.add_argument("-o", "--output-dir", default="output")
    p.add_argument("--senior-pct", type=float, default=0.80,
                   help="Senior tranche as %% of pool (default: 0.80)")
    p.add_argument("--senior-coupon", type=float, default=0.16,
                   help="Class A annual coupon (default: 0.16)")
    p.add_argument("--sub-coupon", type=float, default=0.20,
                   help="Class B annual coupon (default: 0.20)")
    p.add_argument("--cpr", type=float, default=0.08,
                   help="Conditional Prepayment Rate (default: 0.08)")
    p.add_argument("--cdr", type=float, default=0.02,
                   help="Conditional Default Rate (default: 0.02)")
    p.add_argument("--recovery-rate", type=float, default=0.40,
                   help="Recovery rate on defaults (default: 0.40)")
    p.add_argument("--recovery-lag", type=int, default=6,
                   help="Recovery lag in months (default: 6)")
    p.add_argument("--servicing-fee", type=float, default=0.005,
                   help="Annual servicing fee %% (default: 0.005)")
    args = p.parse_args()

    run(
        args.excel_file,
        out_dir=args.output_dir,
        senior_pct=args.senior_pct,
        senior_coupon=args.senior_coupon,
        sub_coupon=args.sub_coupon,
        cpr=args.cpr,
        cdr=args.cdr,
        recovery_rate=args.recovery_rate,
        recovery_lag=args.recovery_lag,
        servicing_fee=args.servicing_fee,
    )


if __name__ == "__main__":
    main()
