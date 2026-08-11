# Handoff — DebtYield-Reviewer

**Repo:** `bencha5e/dy-reviewer`
**Branch:** `claude/debt-yield-audit-tool-02alfq` (up to date with `origin`, working tree clean)
**Latest commit:** `9c03ec6` — "Replace synthetic definition stubs with the real loan-agreement extracts"
**Tests:** 212 passing (`python -m pytest tests/ -q`)
**PR:** #1 was opened earlier in this project and has since been **merged**. The current branch head (`9c03ec6`) is **not yet in a PR** — nothing has been opened for the work since the merge. If picking this back up, either open a new PR from this branch or ask the user how they want to proceed.

## What this tool is

A Python CLI (`python audit_dy.py`) that audits quarterly debt-yield (DY) test Excel models against the loan agreements that govern them. One workbook per loan. Verifies `DY = NCF / UPB` is calculated correctly — not whether the loan clears its covenant (that's explicitly out of scope, removed early in the project per user instruction).

Entry point: `dy_audit/cli.py` → `dy_audit/audit.py::audit_loan()` runs the full check suite per loan. Output: a findings `.xlsx` per loan (`dy_audit/report.py`) plus a run-level log (`dy_audit/runlog.py`). `--move` (default on) drains reviewed files from the input queue into per-loan output folders.

Full behavioral description (checks, severities, design rationale) is in `README.md` — read that first for anything not covered here.

## Session arc (chronological)

1. **Built the tool from `DebtYield_Audit_Tool_Spec.md`** against 4 loans (Strada, Campus, Hialeah, Ares55thAve) in 7 stages. Reproduced every spec §6 finding with zero false positives, plus found 4 defects the spec missed. PR #1 opened and merged.
2. **User reported a missed defect on a 5th file (Lydian)**: GPR overstated $572,580/yr because the tool's rent-roll trace stopped at a `SUM()`-wrapped summary block instead of reaching tenant rows. Fixed by deep-tracing through summary blocks; added `CHK_RENT_STATUS`, `CHK_GPR_TREND`, `CHK_REVENUE_DOUBLE_COUNT`. Lydian added as 5th acceptance fixture. Commit `f532f19`.
3. **User asked for a full independent rent-roll rebuild** (not just tracing the model's own formula chain) on 3 more files (Memorial Hills, Quincy & Hollingsworth, a corrected Hialeah variant). Built `dy_audit/rebuild.py`: locates the tenant table from its own headers (never trusting the model's formulas), classifies every row, rebuilds annualized base rent, reconciles against the OSAR GPR line with itemized differences when it doesn't tie. Added a **Rent Roll Rebuild tab** to every findings workbook. New checks: `CHK_RENT_EDITS` (hand-edited cells in exported columns), `CHK_RR_TOTAL_TIE` (export's own totals vs. summed rows), `CHK_GPR_SUPPLEMENTARY` (non-base-rent codes inside the GPR line). Found real defects in all 3: Memorial's GPR includes pet rent + gas (+$54,414/yr, also double-counted in Other Income) and a year-stale rent roll; Quincy has 9 cells backfilled with prior-quarter rent (+$554,700/yr) plus an inert vacancy formula; corrected Hialeah ties cleanly, confirming the original repo Hialeah's single-cell `*0.95` haircut is real (−$22,525). Commit `a3ac55b`.
4. **User supplied the real loan-agreement definitions** for those 3 files (previously synthetic stubs). Swapped them in; extended `dy_audit/definitions.py`'s regex parser for Quincy's clause-list drafting style (5 new patterns — vacancy/credit-loss factor phrasing, "per residential unit", "% of effective gross income", "delinquent beyond (N) days", "concessions offered based on the trailing"). Two findings correctly became more cautious with real text: Memorial's other income runs on T10 vs. the agreement's stated T3, Quincy's runs on T6 vs. the agreement's stated T12 — both now `MANUAL_REVIEW` on `CHK_OTHER_INCOME_BASIS`. Commit `9c03ec6`. This is the current HEAD.
5. User ran `/compact` twice (both canceled), switched models twice (opus→sonnet), then asked for this handoff.

## Repository layout

```
dy_audit/
  cli.py            # argparse entry point, --move default True
  discovery.py       # pairs .xlsx <-> *_DYDefinitions.md by parsed loan name
  workbook.py         # dual-open wrapper (data_only=True/False), never recalculates
  osar.py             # OSAR tab selection, Line enum, label-driven row resolution
  formula.py          # HIGHEST-RISK MODULE: ref extraction, pass-through resolver,
                       #   restricted recursive-descent evaluator, dependencies()
  definitions.py      # parses *_DYDefinitions.md -> LoanParams (regex-based)
  params.py           # LoanParams dataclass, Parsed wrapper, CURRENT sentinel
  context.py          # LoanContext dataclass threaded through every check
  recompute.py         # independent GPR/vacancy recompute from the model's own
                       #   formula chain (the ORIGINAL, lighter-weight approach)
  rebuild.py           # independent rent-roll REBUILD from the tenant table's own
                       #   headers (the NEWER, deeper approach added in step 3 above)
  checks/
    blockers.py, high.py, medium.py, low.py, standing.py
  report.py            # findings .xlsx: Summary, Findings, Rent Roll Rebuild, Loan Parameters
  runlog.py            # run-level summary log
  filemove.py          # §0 move/rename logic
tests/
  test_stage0.py, test_blockers.py, test_stage3.py, test_stage4.py, test_output.py
  test_lydian.py       # Lydian acceptance tests (step 2 above)
  test_revenue.py      # Memorial/Quincy/Hialeah-variant acceptance tests (steps 3-4)
  fixtures/revenue/     # the 3 revenue-focused workbooks + their real definitions
  conftest.py           # input_copy fixture, copies to tmp_path — never touches repo files directly in write tests

# Fixture workbooks living at repo root (paired with *_DYDefinitions.md there):
1. Strada DY Test_1Q26_vF vBCS.xlsx
2. Campus at Villa LJ DY Test - 1Q 2026 - vF vBCS.xlsx
3. Hialeah Industrial Park 1Q26 DY Test_vF vBCS.xlsx    <- ORIGINAL, uncorrected (has the *0.95 haircut defect)
4. 55thAve_1Q2026_v2 vBCS.xlsx
5. Lydian DY Test_1Q26_vF Copy.xlsx
```

Note: there are now **two independent revenue-verification paths** — `recompute.py` (traces the model's own GPR formula back to source cells) and `rebuild.py` (ignores the model's formulas entirely and rebuilds from the tenant table's headers). Both run on every loan; `rebuild.py` is the newer and more thorough one and is what produces the Rent Roll Rebuild tab. They occasionally reference each other's facts (`ctx.facts["rent_roll_parse"]` from recompute is used as a fallback in rebuild when header detection fails, e.g. Campus).

## Known non-blocking items / things worth knowing before continuing

- **Two `MANUAL_REVIEW` findings introduced in step 4** are real, unresolved methodology gaps (not bugs): Memorial Hills' other-income window and Quincy's other-income window each disagree with what the model actually computed on. These are correctly flagged as needing human review, not something to "fix" in code — the model would need to change, not the checker.
- **`CHK_EXTERNAL_REFS`** fires with large counts (97, 380) on Memorial and Quincy — confirmed in earlier stages that Campus alone has 130 externalLink parts that are stale residue and must NOT flag; the detection rule (cell-formula `[n]` prefixes only, not zip parts) is already tuned for this. Not re-verified against Memorial/Quincy's specific external-ref content, but the same rule applies uniformly.
- **The `tests/fixtures/revenue/` definitions files are the REAL extracted text** as of commit `9c03ec6` — no longer synthetic stubs. No further swap needed.
- **No PR is currently open** for commits `f532f19`/`a3ac55b`/`9c03ec6`. User has not asked for one this round; ask before opening if picking this back up, per the "don't create PRs unless asked" rule.
- The user's default input folder for real runs is a Windows OneDrive path baked into `dy_audit/cli.py::DEFAULT_INPUT` — this cloud session has never touched it; all testing here is against repo/fixture copies only.

## If continuing this work

Read `README.md` in full first (it's kept current with the tool's actual behavior — checks table, severity ladder, the "nothing is keyed to a cell address" and "workbook is never recalculated" design constraints). Then re-run `python -m pytest tests/ -q` to confirm the 212-test baseline before making changes. The acceptance-test convention throughout this project is strict: **every check is asserted in both directions** (must fire on known defects, must NOT fire on known-clean items) — false positives are treated as failures equally with missed flags.
