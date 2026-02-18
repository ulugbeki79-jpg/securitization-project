"""Generate a sample mortgage loan pool Excel file for testing.

Run this once to create sample_data/loan_pool.xlsx.
"""

import random
from datetime import date, timedelta

import pandas as pd

REGIONS = [
    "Toshkent shahri",
    "Toshkent viloyati",
    "Samarqand",
    "Buxoro",
    "Farg'ona",
    "Andijon",
    "Namangan",
    "Qashqadaryo",
    "Surxondaryo",
    "Xorazm",
    "Navoiy",
    "Jizzax",
    "Sirdaryo",
    "Qoraqalpog'iston",
]

random.seed(42)


def generate_loans(n: int = 200) -> pd.DataFrame:
    rows = []
    for i in range(1, n + 1):
        principal = random.randint(50_000_000, 800_000_000)  # UZS
        ltv = round(random.uniform(40, 95), 1)
        property_value = round(principal / (ltv / 100))
        coupon = round(random.uniform(14.0, 24.0), 2)  # annual %
        original_term = random.choice([60, 120, 180, 240])  # months
        elapsed = random.randint(0, min(original_term - 1, 60))
        remaining_term = original_term - elapsed
        origination_date = date.today() - timedelta(days=elapsed * 30)
        days_past_due = random.choices(
            [0, 0, 0, 0, 0, 0, 0, 15, 35, 60, 90],
            weights=[50, 15, 10, 5, 5, 3, 2, 4, 3, 2, 1],
        )[0]
        region = random.choice(REGIONS)

        rows.append(
            {
                "loan_id": f"UZ-MTG-{i:05d}",
                "region": region,
                "principal_uzs": principal,
                "property_value_uzs": property_value,
                "ltv_pct": ltv,
                "coupon_pct": coupon,
                "original_term_months": original_term,
                "remaining_term_months": remaining_term,
                "origination_date": origination_date,
                "days_past_due": days_past_due,
                "borrower_type": random.choice(["individual", "individual", "individual", "entrepreneur"]),
            }
        )

    return pd.DataFrame(rows)


if __name__ == "__main__":
    df = generate_loans()
    out = "sample_data/loan_pool.xlsx"
    df.to_excel(out, index=False, sheet_name="Loan Pool")
    print(f"Generated {len(df)} loans -> {out}")
