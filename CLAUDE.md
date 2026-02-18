# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MBS (Mortgage-Backed Securities) modeling framework for the Mortgage Refinancing Company of Uzbekistan. This is a greenfield project building tooling for Uzbekistan's emerging securitization practice, combining local regulatory requirements with international structured finance best practices.

**Currency**: UZS (Uzbek som) — all monetary values, cash flow projections, and financial calculations use UZS unless explicitly stated otherwise.

**Bilingual requirement**: All documentation, reports, UI labels, and user-facing strings must support both Uzbek/Russian and English.

## Domain Context

The securitization pipeline follows this workflow:

1. **Loan Pool Selection** — Screening and selecting mortgage loans from the company's portfolio based on eligibility criteria
2. **Due Diligence** — Validating loan-level data quality, legal documentation, and borrower information
3. **SPV Structuring** — Establishing Special Purpose Vehicles to isolate securitized assets from the originator's balance sheet
4. **Credit Enhancement Analysis** — Sizing overcollateralization, subordination, reserve accounts, and other credit support mechanisms
5. **Cash Flow Waterfall Modeling** — Defining priority-of-payment rules and distributing projected cash flows across tranches (senior, mezzanine, equity)
6. **Investor Reporting** — Generating pool performance summaries, compliance documentation, and investor-facing reports

## Regulatory Context

- Uzbekistan's securities regulatory framework is actively evolving; the project must accommodate regulatory changes
- Compliance logic should be modular and configurable rather than hardcoded
- International references: follow conventions from established MBS markets (e.g., SIFMA/ISDA standards) where local rules have not yet been defined

## Planned Modules

- **Asset Pool Management** — Mortgage loan pool definition, eligibility screening, pool stratification
- **Tranche Structuring** — Senior/mezzanine/equity tranching with configurable attachment/detachment points
- **Cash Flow Waterfall** — Priority-of-payment engine with configurable waterfall rules
- **Credit Enhancement** — Overcollateralization, subordination, reserve account modeling
- **Risk Analytics** — Stress testing, default/prepayment modeling (CPR/CDR), scenario analysis
- **Reporting** — Bilingual investor reports, pool tape generation, regulatory filings

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Generate sample loan pool (200 loans)
python sample_data/generate_sample_pool.py

# Run loan pool analysis (outputs report + charts to output/)
python -m pool_analysis.analyze sample_data/loan_pool.xlsx

# Specify custom output directory
python -m pool_analysis.analyze path/to/pool.xlsx -o my_output/
```
