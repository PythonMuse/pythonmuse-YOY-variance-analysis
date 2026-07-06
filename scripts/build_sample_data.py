"""
build_sample_data.py - CodeCritters Inc. sample trial balance generator

Regenerates the two synthetic trial balance CSVs used by
skills/trial-balance-comparison/scripts/generate_visuals.py. Not part of the
analysis skill itself (see CLAUDE.md, Data Locations). Run this only if the
sample data needs to be rebuilt from scratch - the checked-in CSVs in data/
are the read-only source of truth for the demo.

Usage:
    python3 scripts/build_sample_data.py
"""

import csv
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
os.makedirs(DATA_DIR, exist_ok=True)

HEADER = ["Account Code", "Account", "Type", "Jan", "Feb", "Mar", "Apr"]

# (Account Code, Account, Type, prior-year Jan-Apr, current-year Jan-Apr)
ACCOUNTS = [
    ("1000", "Checking",                    "Asset",     [7600, 11270, 18855, 27430],   [42870, 47175, 41325, 41385]),
    ("1010", "Savings",                     "Asset",     [20000, 20000, 20000, 20000],  [20000, 20000, 20000, 20000]),
    ("1200", "Accounts Receivable",         "Asset",     [14000, 14500, 15000, 15200],  [15400, 15700, 16100, 24800]),
    ("1300", "Inventory - Raw Materials",   "Asset",     [8000, 8200, 8400, 8600],      [8700, 8850, 9200, 9500]),
    ("1310", "Inventory - Finished Goods",  "Asset",     [11000, 11200, 11400, 11600],  [11700, 11900, 12100, 12300]),
    ("1400", "Prepaid Insurance",           "Asset",     [2400, 2200, 2000, 1800],      [1600, 1400, 1200, 1000]),
    ("1500", "Equipment",                   "Asset",     [45000, 45000, 45000, 45000],  [45000, 45000, 46800, 46800]),
    ("1510", "Accumulated Depreciation",    "Asset",     [9800, 10600, 11400, 12200],   [13000, 13800, 14600, 15400]),
    ("2000", "Accounts Payable",            "Liability", [7000, 7200, 7400, 7600],      [7700, 7900, 8100, 8300]),
    ("2100", "Credit Card Payable",         "Liability", [3200, 3300, 3100, 3400],      [3500, 3600, 3450, 4700]),
    ("2200", "Sales Tax Payable",           "Liability", [1800, 1850, 1900, 1950],      [1980, 2020, 2060, 2100]),
    ("2300", "Payroll Liabilities",         "Liability", [2200, 2200, 2300, 2300],      [2350, 2350, 2400, 2400]),
    ("2400", "Loan Payable - Equipment",    "Liability", [18000, 17600, 17200, 16800],  [16400, 16000, 15600, 15200]),
    ("3000", "Common Stock",                "Equity",    [10000, 10000, 10000, 10000],  [10000, 10000, 10000, 10000]),
    ("3100", "Retained Earnings",           "Equity",    [58000, 63620, 73355, 83380],  [92540, 98755, 97115, 106485]),
    ("3200", "Owner Draws",                 "Equity",    [2000, 4000, 6000, 8000],      [2200, 4400, 6600, 8800]),
    ("4000", "Sales - PyPal",               "Revenue",   [15000, 15500, 16000, 16200],  [16000, 16300, 17000, 21000]),
    ("4010", "Sales - ByteBot",             "Revenue",   [18000, 18500, 19000, 19200],  [19000, 19500, 19800, 20200]),
    ("4020", "Sales - TensorTurtle",        "Revenue",   [9000, 9200, 9500, 9700],      [9700, 9900, 10100, 10400]),
    ("4900", "Shipping Income",             "Revenue",   [1200, 1250, 1300, 1320],      [1320, 1350, 1400, 1450]),
    ("5000", "COGS - Materials",            "COGS",      [9000, 9200, 9400, 9600],      [9500, 9700, 13500, 14200]),
    ("5010", "COGS - Labor",                "COGS",      [6000, 6100, 6200, 6300],      [6300, 6400, 6500, 6600]),
    ("6000", "Rent Expense",                "Expense",   [3500, 3500, 3500, 3500],      [3600, 3600, 3600, 3600]),
    ("6010", "Utilities",                   "Expense",   [450, 480, 420, 400],          [460, 490, 430, 410]),
    ("6020", "Payroll Expense",             "Expense",   [12000, 12000, 12200, 12200],  [12500, 12500, 12500, 12500]),
    ("6030", "Marketing & Advertising",     "Expense",   [1500, 1500, 1600, 1600],      [1600, 1600, 9800, 3200]),
    ("6040", "Software Subscriptions",      "Expense",   [380, 380, 395, 395],          [400, 400, 410, 410]),
    ("6050", "Insurance Expense",           "Expense",   [600, 600, 600, 600],          [620, 620, 620, 620]),
    ("6060", "Professional Fees",           "Expense",   [500, 3800, 450, 500],         [550, 4200, 500, 550]),
    ("6070", "Depreciation Expense",        "Expense",   [800, 800, 800, 800],          [820, 820, 820, 820]),
    ("6080", "Bank Fees & Merchant Charges", "Expense",  [320, 330, 340, 345],          [350, 355, 360, 600]),
    ("6090", "Office Supplies",             "Expense",   [150, 140, 160, 155],          [160, 150, 900, 170]),
]


def write_csv(path, year_index):
    with open(path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(HEADER)
        for code, name, acct_type, prior, current in ACCOUNTS:
            values = prior if year_index == 0 else current
            writer.writerow([code, name, acct_type] + [float(v) for v in values])


# Contra accounts carry a natural balance opposite their parent category and
# must be SUBTRACTED, not summed in, when reconciling (Accumulated Depreciation
# is a contra-asset; Owner Draws is a contra-equity account).
CONTRA_ASSETS = {"Accumulated Depreciation"}
CONTRA_EQUITY = {"Owner Draws"}


def reconciles(values_by_type, months):
    """Assets should equal Liabilities + Equity for every month, net of contras."""
    ok = True
    for i, m in enumerate(months):
        assets = sum(v[i] for _, n, t, v in values_by_type if t == "Asset" and n not in CONTRA_ASSETS)
        assets -= sum(v[i] for _, n, t, v in values_by_type if n in CONTRA_ASSETS)
        equity = sum(v[i] for _, n, t, v in values_by_type if t == "Equity" and n not in CONTRA_EQUITY)
        equity -= sum(v[i] for _, n, t, v in values_by_type if n in CONTRA_EQUITY)
        liab = sum(v[i] for _, n, t, v in values_by_type if t == "Liability")
        if round(assets - (liab + equity), 2) != 0:
            print(f"[WARN] {m}: Assets {assets} != Liabilities + Equity {liab + equity}")
            ok = False
    return ok


if __name__ == "__main__":
    months = ["Jan", "Feb", "Mar", "Apr"]

    prior_rows = [(c, n, t, p) for c, n, t, p, cur in ACCOUNTS]
    curr_rows = [(c, n, t, cur) for c, n, t, p, cur in ACCOUNTS]

    print("Checking prior-year (2025) reconciliation...")
    ok_2025 = reconciles(prior_rows, months)
    print("Checking current-year (2026) reconciliation...")
    ok_2026 = reconciles(curr_rows, months)

    if not (ok_2025 and ok_2026):
        raise SystemExit("Trial balance does not reconcile - fix source data before writing CSVs.")

    write_csv(os.path.join(DATA_DIR, "trial_balance_2025.csv"), year_index=0)
    write_csv(os.path.join(DATA_DIR, "trial_balance_2026.csv"), year_index=1)
    print("\n[OK] Both trial balances reconcile (Assets = Liabilities + Equity, every month).")
    print("Wrote data/trial_balance_2025.csv and data/trial_balance_2026.csv")
