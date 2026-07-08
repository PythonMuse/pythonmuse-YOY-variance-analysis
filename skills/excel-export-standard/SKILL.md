# SKILL: excel-export-standard

**Skill directory:** `skills/excel-export-standard/`
**Script:** `skills/excel-export-standard/scripts/build_formatting_standard_demo.py`

---

## Purpose

Define — and demonstrate — the formatting standard every PythonMuse Excel
deliverable follows, so a reviewer can tell a compliant workbook from a
sloppy one on sight, and so every future skill that writes `.xlsx` output
(starting with `skills/trial-balance-comparison`) uses the same look and the
same audit guarantee: **every calculated cell is a formula, never a pasted
value.**

This skill doesn't produce a client deliverable itself. It produces a small
teaching workbook (`outputs/excel/formatting_standard_demo.xlsx`) that shows
the standard side by side with the anti-pattern it exists to prevent —
useful for onboarding, code review, and the "what does audit-ready actually
look like" moment in a demo.

---

## When to use this skill

- Before building a new skill that exports an Excel workbook — read this
  first, then match its constants (fill colors, fonts, number formats).
- During review of an existing workbook output — regenerate the demo file
  and compare styling side by side if something looks off.
- In a demo, to make Rule 8 ("Excel formulas, not hardcoded results") a
  visible, clickable thing instead of a sentence in CLAUDE.md.

---

## Inputs

None. This skill's script is self-contained — it builds illustrative sample
data, it doesn't read from `data/`.

## Outputs

| File | Location | Notes |
|------|----------|-------|
| Formatting standard demo workbook | `outputs/excel/formatting_standard_demo.xlsx` | Regenerated each run |

---

## The standard

| Element | Rule | Value |
|---|---|---|
| Title cell | Bold, 14pt, Deep Navy text | `Font(bold=True, size=14, color="002030")` |
| Header row fill | Solid Panel Navy | `PatternFill("solid", fgColor="0A3040")` |
| Header row font | White, bold | `Font(color="FFFFFF", bold=True)` |
| Flag / highlight fill | Solid Gold Amber | `PatternFill("solid", fgColor="F8D038")` |
| Currency cells | Thousands separator, 2 decimals | `number_format = "#,##0.00"` |
| Percent cells (detail rows) | 1 decimal | `number_format = "0.0%"` |
| Percent cells (threshold/summary rows) | Whole number | `number_format = "0%"` |
| Header row | Frozen so it stays visible while scrolling | `ws.freeze_panes = "A2"` (or first data row) |
| Column widths | Sized to fit the widest expected value, not left at default | `ws.column_dimensions[col].width = n` |
| **Calculated cells** | **Excel formula referencing the raw data cells — never a Python-computed value pasted in as a static number** | e.g. `='TB_2026'!D3-'TB_2025'!D3` |

These constants match what `skills/trial-balance-comparison/scripts/generate_visuals.py`
already uses (`DEEP_NAVY`, `PANEL_NAVY`, `GOLD_AMBER` in its `CONFIG` block) —
this skill documents that as the standard, it doesn't invent a new one.

---

## The rule that matters most: formulas, not values

A cell like a variance, a percentage, or a total must be written so that a
reviewer can:
1. Click the cell.
2. See a formula in the formula bar, referencing other cells in the
   workbook — not a bare number.
3. Change a source number and watch the calculated cell update.

Writing `ws.cell(row, col, value=computed_variance)` — where `computed_variance`
was calculated in Python — fails this test even if the number is correct.
The workbook stops being auditable the moment a reviewer can't tell whether
a number is a live calculation or a snapshot someone pasted in. This is
CLAUDE.md Rule 8, and it's the one CLAUDE.md says to recalculate and confirm
zero formula errors on before calling anything final.

---

## Steps

1. Run the script:
   ```bash
   python3 skills/excel-export-standard/scripts/build_formatting_standard_demo.py
   ```
2. Open `outputs/excel/formatting_standard_demo.xlsx`.
3. On the **Formula vs. Hardcoded** tab, click the "Correct" column's variance
   cells and confirm the formula bar shows a formula referencing the raw
   data columns on the same sheet. Click the "Wrong" column's cells and
   confirm they show a bare number — that's the anti-pattern, colored red
   and labeled, not something to copy.
4. Change one of the raw input numbers on that tab and recalculate (F9 in
   Excel, or reopen in a viewer that recalculates on load). Confirm the
   "Correct" column updates and the "Wrong" column does not — that gap is
   the entire point of Rule 8.
5. On the **Formatting Standard** tab, confirm the header row is frozen,
   styled per the table above, and the currency/percent columns show the
   correct number formats.

---

## Validation checks (after running)

- [ ] Header rows use the exact fill/font from the standard table, not an
      ad hoc color
- [ ] Every "calculated" cell in the Formula vs. Hardcoded tab that's meant
      to be live is an actual formula (formula bar starts with `=`)
- [ ] Currency columns show 2 decimals and a thousands separator; percent
      columns show `%`
- [ ] Header row is frozen on every data-bearing tab
- [ ] Workbook recalculates with zero formula errors

---

## Notes

- This skill uses `openpyxl` — already a project dependency (`requirements.txt`).
- If you add a new Excel-exporting skill later, import the color/format
  constants from this skill's script rather than re-typing hex codes —
  drift between skills is how "the standard" quietly stops meaning anything.
- This is a teaching/demo artifact, not a client deliverable — it never
  reads real trial balance data and is not subject to CLAUDE.md's
  source-data rules (Rule 9, Rule 11) because it has no source data to
  protect.
