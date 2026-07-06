# SKILL: trial-balance-comparison

**Skill directory:** `skills/trial-balance-comparison/`
**Script:** `skills/trial-balance-comparison/scripts/generate_visuals.py`

---

## Purpose

Compare two trial balance periods (Year-over-Year and Month-over-Month), flag
accounts that cross materiality thresholds on **both** a dollar and a percent
basis, and produce reproducible charts and an Excel workpaper. Encodes
CLAUDE.md Rules 1, 2, 3, 5, and 8.

---

## When to use this skill

Use this skill any time a reviewer asks Varia to "run the comparison" or
"regenerate the report" for the CodeCritters Inc. trial balances.

---

## Inputs

| File | Location | Notes |
|------|----------|-------|
| Prior-year trial balance | `data/trial_balance_2025.csv` | Read-only source. Never modify (Rule 9). |
| Current-year trial balance | `data/trial_balance_2026.csv` | Read-only source. Never modify (Rule 9). |

## Outputs

| File | Location | Notes |
|------|----------|-------|
| Net income trend chart | `outputs/visuals/01_net_income_trend.png` | Regenerated each run |
| YoY variance by account chart | `outputs/visuals/02_yoy_variance_by_account.png` | Regenerated each run |
| Flagged accounts trend chart | `outputs/visuals/03_flagged_accounts_trend.png` | Regenerated each run |
| Excel workpaper | `outputs/excel/CodeCritters_TB_Comparison_Report.xlsx` | Every variance/total cell is an Excel formula (Rule 8), not a pasted value |

---

## Steps

1. Confirm both source CSVs exist under `data/`.
2. Load each CSV and verify Assets = Liabilities + Equity for every month, net of contra accounts (Rule 5). Stop and report the imbalance if it doesn't reconcile — do not proceed.
3. Classify every account as balance-sheet (point-in-time) or P&L (period activity) before comparing (Rule 1). Never sum a balance-sheet account's monthly columns.
4. Run the Year-over-Year pass: same month, current year vs. prior year.
5. Run the Month-over-Month pass: sequential months within the current year only.
6. Keep the two passes separate — never blend them into one number (Rule 2).
7. Flag an account only if it clears **both** the dollar threshold and the percent threshold for that comparison type (Rule 3). Thresholds live in the script's `CONFIG` block, not in a one-off chat answer (Rule 4).
8. Check every MoM flag against the prior year for the same seasonal pattern; if it recurs, say so explicitly rather than treating it as a surprise (Rule 6).
9. Run the script:
   ```bash
   python3 skills/trial-balance-comparison/scripts/generate_visuals.py
   ```
10. Confirm the three charts and the Excel workpaper were written to `outputs/`.
11. Open the workbook and recalculate — confirm zero formula errors before calling anything final (Rule 8).

---

## Validation checks (after running)

- [ ] Both trial balances reconcile (Assets = Liabilities + Equity) for every month
- [ ] Every flagged account clears both the dollar bar and the percent bar for its comparison type
- [ ] Every recurring seasonal MoM flag is labeled as such, not presented as a fresh finding
- [ ] Every cell in `YoY_Variance`, `MoM_Variance`, and `Flagged_Summary` is a formula referencing `TB_2025`/`TB_2026`, not a static value
- [ ] Workbook recalculates with zero formula errors

---

## Notes

- A flagged account is a question for a human, never a verdict — see `agents/varia_variance_agent.md` and the root `CLAUDE.md` for how to report findings.
- If asked to change thresholds, add accounts, or extend the period range, update this script — do not produce a one-off manual calculation it can't reproduce next time.
