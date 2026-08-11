# Debt Yield Audit Tool — Build Spec & Q1 2026 Audit Findings

**Purpose.** This document is written to be handed to Claude Code to build a Python tool that reviews quarterly debt‑yield (DY) test models. It encodes (a) how to translate a loan's NOI definition text into calculation rules, (b) a chronological audit procedure, (c) the explicit error checks the tool must run, (d) the resolved house rules that govern ambiguous cases, and (e) the specific flags found in the Q1 2026 files.

> **Scope of the tool (resolved).** The tool runs a **hybrid**: it **verifies** the model's formulas and links line‑by‑line, and it **independently recomputes** two lines — **Gross Potential Rent (GPR)** and **Vacancy** — from the raw rent roll, comparing back to the model. It does *not* attempt a full independent rebuild of every expense line; ordinary opex is verified as a T12 pass‑through. See §5 for the full set of house rules that were previously open questions.

---

## 0. File I/O, folder structure, and run workflow

**Input folder (one run = one folder):**
```
C:\Users\bstamp\OneDrive - Bellwether\Desktop\Projects\DebtYieldFiles\Input DY Tests\
```
Each quarter, this folder holds the DY-test `.xlsx` files to be reviewed, paired with their loan-definition `.txt`/`.md` files. Files are named so the loan is identifiable from the filename (as in the four examples this spec was built from — e.g. `1__Strada_DY_Test_1Q26_vF_vBCS.xlsx` paired with `1__Strada_DYDefinitions.md`). The tool must **match each Excel file to its corresponding definitions file by loan name parsed out of the filename** — don't assume a fixed pairing order or a fixed count of files; the folder may contain any number of loans on a given run, and the loan roster will change over time (loans get added/paid off).

**Per-loan, per-run processing:**
- The tool discovers all loan pairs present in the input folder and processes **each loan independently**, producing its own audit output — do not consolidate everything into one combined report. One loan's broken formula or unreadable tab should not block the run for any other loan.
- Wrap each loan's processing in its own try/except so a failure on one loan is captured as an error in that loan's output rather than crashing the whole run.

**Output — one folder per loan per run, moved (not copied):**
- Destination root:
  ```
  C:\Users\bstamp\OneDrive - Bellwether\Desktop\Projects\DebtYieldFiles\Input DY Tests\DY Review Output\
  ```
- For each loan processed, create a new subfolder named:
  ```
  <Loan Name> - <YYYY_MM_DD> - v<N>
  ```
  e.g. `Strada - 2026_08_10 - v2`. `<YYYY_MM_DD>` is the run date (today, not the DY-test period). `v<N>` starts at `v1` and auto-increments if a folder with the same loan name and date already exists (i.e., a second run on the same day for the same loan produces `v2`, a third `v3`, etc.) — never overwrite an existing output folder.
- Into that new subfolder, the tool **moves** (not copies) the loan's:
  1. original DY-test `.xlsx` file,
  2. corresponding loan-definitions text file, and
  3. the audit output the tool generated for that loan (see below),
  
  out of the input folder and into the new output subfolder. After a successful run, the input folder should no longer contain the files that were just processed — this keeps the input folder as a queue of "not yet reviewed" files.
- If a loan's processing fails partway through, do **not** move its source files out of the input folder — leave them in place so the loan is retried on the next run, and note the failure in a run-level log instead.

**Audit output format (resolved): `.xlsx` findings workbook, one per loan.** Matches the source material and is easy to scan alongside the model. At minimum it should include: the loan name, run date, reported DY (and covenant pass/fail), every flag from §4 with severity/cell references, the standing UPB-confirmation flag, and the month-count actually used for each annualized line. (Sheet/column layout of this workbook is a detail for the build — a summary sheet up top plus a detail sheet of every check with severity, status, cell reference, and note is a reasonable starting shape.)

**Run-level summary log (resolved): include one.** In addition to each loan's `.xlsx` output, drop a lightweight run-level summary (e.g. `.md` or `.txt`) at the **output root** (`...\DY Review Output\`) for each run, listing: run date/time, which loans succeeded vs. failed (with the failure reason for any that failed and were left in the input queue), and a total flag count by severity across all loans processed that run.

**Deliverable shape (resolved): a CLI script** invoked per-run (e.g. `python audit_dy.py`) that scans the whole input folder in one go, processes every loan pair found, writes each loan's `.xlsx` workbook and moves files into its output subfolder, and writes the run-level summary log at the output root.

---

## 1. The mental model (read this first)

Every workbook has the same skeleton:

- An **OSAR output tab** — named `NEW OSAR`, `(New) Comm OSAR`, `Comm OSAR`, or `OSAR`. The **current DY‑test figures live in column I** (rightmost figure column). **Column H = trailing‑12‑month (T12) actual.** Columns E–G = at‑contribution / prior DY tests (history that rolls).
- Supporting tabs feed column I: **income statement / operating statement (T12)**, **rent roll (RR)**, **AR / aging report**, a **Tax** tab, an **Insurance** tab, and a **Debt Service** tab.

**The single most important rule:**

```
Debt Yield = Net Cash Flow (NCF) / Unpaid Principal Balance (UPB)
```

where **NCF = NOI − Capital Items (the replacement/CapEx reserve)**.

Why NCF and not the model's "NOI" line: each loan's NOI *definition* already subtracts a normalized replacement reserve (e.g., "$250 per unit," "$0.25 per square foot," "$0.10 per square foot"). In the spreadsheets that reserve is placed **below** the NOI line, inside "Capital Items," so the **model's NCF line equals the loan's definitional NOI.** A tool that computed `model_NOI / UPB` would overstate DY on all four loans. **Always drive DY off the NCF line.**

Corollary: the model's on‑tab "NOI" line = definitional NOI *before* the reserve. Don't report that as NOI for the covenant.

---

## 2. Interpreting the loan‑definition text → calculation rules

The definition text file (one per loan) governs how each **revenue** and **expense** line is built for column I. Below is the general pattern, then a per‑loan parameter table. The parser should key off the loan name header and pull the numeric parameters (vacancy floor, reserve rate, delinquency window, occupancy‑expectation windows, credit‑tenant carve‑outs).

### 2.1 Revenue lines

**Gross Potential Rent (GPR) line — the base.**
- Base = **annualized in‑place rent from the current rent roll**: monthly contractual rent of **occupied units/suites whose tenants are current/non‑excluded**, × 12. Because vacant space contributes $0, **physical vacancy is already embedded in this number** — this is the crux of the vacancy logic below.
- Include only tenants that pass the loan's **exclusion filter** (varies by loan; see table): typically excludes tenants > X days delinquent, month‑to‑month, in bankruptcy (unaffirmed), "gone dark" (with credit‑rating/lease‑term carve‑outs), and known/noticed vacates. Add newly‑executed leases not yet in occupancy **only** if occupancy is expected within the loan's window (usually 90 days; 180 for investment‑grade tenants).
- Contractual rent steps over the next 12 months are credited **only for the months they are actually in effect** (do not annualize a step that starts mid‑year as if effective all year).
- Free rent exceeding **1 month per lease year** is deducted from that tenant's revenue.

**Vacancy Loss line — the 5% floor.**
- Loan rule: vacancy factor = **greater of (x) actual vacancy or (y) 5.00%**.
- Because the GPR line is already in‑place (net of actual vacancy), the model only needs to deduct **additional** vacancy when **actual vacancy < 5%** (i.e., occupancy > 95%). When actual vacancy ≥ 5%, the correct vacancy‑loss line is **0** — the actual vacancy is already reflected in in‑place GPR.
- Correct formula shape (occupancy = `occ`, gross‑up to potential then take the shortfall to 5%):
  `vacancy_loss = -(GPR / occ) * (5% - (1 - occ))` when `occ > 95%`, else `0`.
  Equivalent simpler forms exist; **the sign must be negative** (a deduction) and it must be **0 (or blank) when occupancy ≤ 95%.**
- ⚠️ **Sign check required** — see Flag S‑3 (Strada's version adds income in the >95% branch).
- Investment‑grade tenants with ≥ 2 years remaining term may be **excluded from the vacancy factor** in some loans (Campus). Capture this carve‑out.

**Expense Reimbursement / Recovery line (commercial).**
- Annualized in‑place recoveries from occupied tenants (monthly recovery × 12, or the RR's annualized recovery column). Per the definition, recoveries belong to clause (a) revenue and are therefore **subject to the vacancy factor** — verify they aren't parked in "Other Income" to escape it (see Flag A‑7).

**Parking Income / Other Income.**
- Method varies by loan — read the definition:
  - **Strada:** other income = **trailing‑3‑month actual, annualized (T3A)**, **less** concessions on a **trailing‑6‑month annualized (T6A)** basis. Parking = T3A.
  - **Campus / Hialeah / Ares:** other income = **most recent 12‑month (T12)**, adjusted for non‑recurring/extraordinary items. (Ares uses T‑available annualized when < 12 months of history exist.)
- **Do not** apply the vacancy factor to "other income."

### 2.2 Expense lines

- **Most operating expenses = T12 actual** (trailing 12 months), pulled straight from the income‑statement tab into column I.
- **Real Estate Taxes = MAX(T12 actual, actual tax bill/invoice)** — the invoice comes from the **Tax** tab. **Universal hard rule for every loan** — a model that references only one side is a BLOCKER, not a per‑loan option.
- **Property Insurance = MAX(T12 actual, actual premium/invoice)** — the invoice comes from the **Insurance** tab. **Same universal hard rule.** ⚠️ Not implemented on Strada (Flag S‑1); broken on Ares (Flag A‑2). Both must hard‑fail.
- **Management Fee = MAX(actual fee, 3.0% × EGI).** House rule: the 3% base is **always EGI**, regardless of whether the loan text says "Gross Revenues," "gross operating income," or a tab note says "GPR." Do not treat an EGI base as a discrepancy; *do* flag a model that uses anything other than EGI (e.g., 3%×GPR).
- **Replacement / CapEx reserve** = per‑unit or per‑SF rate × count, placed in **Capital Items below NOI**. Rate per loan in the table. This is what makes NCF = definitional NOI.
- **Excluded from opex** by definition (do not let these leak into column I): depreciation/amortization, income taxes, loan/financing costs, capital expenditures, debt service, and any expense paid directly by a tenant. Non‑recurring/extraordinary items are excluded (Hialeah, Ares explicitly).

### 2.3 Per‑loan parameter table

| Parameter | Strada | Campus at Villa La Jolla | Hialeah | Ares 55th Ave |
|---|---|---|---|---|
| Property type | Multifamily | Medical office | Industrial | Industrial |
| Output tab (live) | `NEW OSAR` | `Comm OSAR` | `(New) Comm OSAR` | `Comm OSAR` (built on `Lender Calc` tab) |
| DY cell | `I73` | `I78` | `I78` | `I78` |
| Vacancy floor | 5% | 5% | 5% | 5% |
| Delinquency exclusion window | 60+ days | 45+ days | 60+ days | 60+ days |
| New‑lease occupancy window | (n/a MF) | 90d / 180d IG | 180d | 90d / 180d IG |
| Reserve rate | $250 / unit | $0.25 / sf | $0.25 / sf | $0.10 / sf |
| Other income basis | T3A less T6A concessions | T12 | T12 | T‑avail annualized |
| Mgmt fee | max(actual, 3% EGI) | max(actual, 3% EGI) | max(actual, 3% EGI) | max(actual, 3% EGI) |
| Tax / Insurance | max(T12, invoice) | max(T12, invoice) | max(T12, invoice) | max(T12, invoice) |

*House rule: the 3% mgmt‑fee base is **always EGI** regardless of loan wording. Tax/Insurance MAX(T12, invoice) is **universal** — enforce on every loan.*

### 2.4 Annualization & short operating history

Annualization is literal: **annualize whatever months are available**, scaling by `12 / (number of months used)`. **Assume every month present is a full month of data** — the tool does not attempt to detect or drop "partial" or "stub" months; that judgment is left to the reviewer.

- **Full 12 months present →** standard. T12 lines pass through; T‑N lines (e.g., Strada's T3‑annualized other income, T6‑annualized concessions) use their nominal window × (12/N).
- **Fewer than 12 months present →** annualize on the months you have. Do **not** zero‑fill missing months.
  - A **T12** line with only 11 months = `SUM(11 months) × 12/11`.
  - A **T3‑annualized** line with only **2** months of history = `SUM(2 months) × 12/2`. With **4** months available, still use the nominal **T3** (`× 12/3`) — the window doesn't grow past its definition, it only shrinks when data is short.
- **`CHK_MONTH_COUNT` (verify, don't infer):** for every annualized line, count the months the *model's own formula* actually sums, and compare that count to the number of months present on the source tab for the applicable window (12 for T12, 3 for T3, 6 for T6, etc.). If the model's formula omits a month that's present in the source data (e.g., sums 11 of 12 available months), **flag it** — cite the line, the tab, the month count used vs. available, and let the reviewer inspect why. Do not auto‑correct or reclassify the omitted month.
- The tool should **surface the month count it used** for each annualized line so the reviewer can see when a line was built on short history or is omitting available data.

---

## 3. Step‑by‑step audit procedure (chronological)

Run these in order; each stage assumes the prior one passed.

**Stage 0 — Identify the live output tab and period.**
1. **Tab‑selection rule (resolved):** use the OSAR tab that is **visible**; **ignore any hidden OSAR tabs entirely** (do not read, audit, or flag them). Only in the rare case that **more than one OSAR tab is visible** do you disambiguate by "the tab whose column I is populated and whose `Statement Ending Date` (row 15) = quarter‑end." Read sheet visibility from the workbook (openpyxl `ws.sheet_state == 'visible'`); a template tab of the other property type is typically hidden and should be skipped on that basis.
2. Confirm the **as‑of date** = correct quarter‑end (e.g., 3/31 for 1Q, 6/30 for 2Q, 9/30 for 3Q, 12/31 for 4Q). Check it in row 15/row 3 header and that it flows to dependent date cells.

**Stage 1 — Confirm the correct period financials are pulled.**
3. **Income statement:** confirm the T12 tab is the trailing‑12 **ending at quarter‑end** (look for "Ending Period <Month Year>" and the month column headers). Where multiple period tabs exist (Hialeah has `1Q26 T12`, `2Q T12`, `3Q T12`), confirm column I references the current one.
4. **Rent roll:** confirm the RR is **as of quarter‑end** (header "As of MM/DD/YYYY"). Beware a **report run‑date** later than the as‑of date — that's fine as long as the snapshot is quarter‑end. Where duplicate RR tabs exist (Hialeah `1Q RR` vs `1Q26 RR`), confirm the live one is referenced.
5. **AR / aging:** confirm the delinquency cut used to exclude tenants matches the loan's window (45/60 days) and ties to the same as‑of date.

**Stage 2 — Verify the links from supporting tabs to column I.**
6. For every income and expense line in column I, resolve the formula and confirm it points to the **right cell on the right tab** and that the value is non‑stale (recalc if needed). Watch for:
   - whole‑column `SUM(X:X)` (fragile — picks up stray values),
   - `INDEX/MATCH` / `XLOOKUP` on a label that must exist exactly once on the source tab,
   - links to an **external workbook** (`[1]`, `[2]` prefixes) — these break when the other file is absent and leave only cached values.
7. Confirm each label in column C/B matches the line it's pulling (no row‑offset errors).

**Stage 3 — Recompute the NOI build for column I and tie to definition.**
8. **GPR (independent recompute — mandatory):** rebuild annualized in‑place rent directly from the RR (monthly contractual rent of included tenants × 12) and compare to the model's GPR cell; flag per the §5 threshold. The **delinquency exclusion is driven off the AR aging report** — a tenant past the loan's delinquency window (45/60 days) in the AR aging is excluded. Use the AR aging as the **source of truth** for delinquency; verify the analyst's RR include/exclude flags against it (and against lease status for MTM/dark/BK/known‑vacate).
9. **Vacancy (independent recompute — mandatory):** recompute the vacancy line from occupancy and compare to the model. Verify sign and the >95% / ≤95% branching (§2.1). Confirm floor = 5% and any IG carve‑out.
10. **Other income / parking / recoveries:** verify the basis (T3A/T6A vs T12) matches the loan.
11. **Expenses:** confirm T12 pass‑through for ordinary opex; confirm **Tax = MAX(T12, invoice)** and **Insurance = MAX(T12, invoice)**; confirm **Mgmt = MAX(actual, 3%×base)**; confirm reserve rate.
12. **EGI, Total Opex, NOI** subtotals sum the intended ranges (no dropped or double‑counted rows).

**Stage 4 — Capital items, NCF, and Debt Yield.**
13. Confirm reserve = rate × count (units or SF), and that count ties to the property (watch unit/SF mismatches between RR and OSAR).
14. Confirm **NCF = NOI − Capital Items**.
15. **UPB (resolved):** there is **no external source file for UPB** at this stage. Do **not** try to verify the UPB figure against anything — instead **always emit a standing flag** in the findings that "UPB used in the DY test must be confirmed against internal records" for each loan, and echo the UPB value the model used. (A source file may be wired in later.)
16. **DY of record (resolved):** the **only** DY that matters is the one on the **OSAR output tab** — that is the number that must be correct, and it drives all pass/fail conclusions. If any *other* tab contains a DY calc (e.g., a Debt Service tab), **do not audit it**; only check whether it **ties** to the OSAR DY, and if it doesn't, **raise a flag** naming the tab/cell and both values. Spend no further effort evaluating the off‑tab calc.
17. Compare DY to the covenant threshold and its effective date; note pass/fail and whether the threshold is yet in force.

**Stage 5 — Debt Service tab (feeds DSCR, not DY).**
18. **Rate (resolved):** always use the **actual** effective rate = **MIN(SOFR, Cap) + Spread** — never stress at the cap strike. Confirm Debt Service = `UPB × rate × (365/360)`, and that SOFR, spread, and cap strike inputs are current.

**Stage 6 — Whole‑file hygiene.**
19. Scan every tab for cached error values (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`) and report location + whether it lies on the DY path.
20. Report any hardcoded plugs inside otherwise‑formula cells (constants added into a formula, e.g., `+34000`).

---

## 4. Error checks the tool must enforce (with severity)

**Materiality thresholds (when a recompute/verify difference fires a flag):**
- **NOI, EGI, and Total Operating Expenses:** flag if the recomputed figure differs from the model by **> $1.00** (effectively an exact tie — these subtotals should reconcile to the dollar).
- **Individual revenue line items** (GPR, vacancy, recoveries, parking, other income): flag if the recomputed figure differs from the model by **> 0.1%**.
- These thresholds gate `CHK_GPR_RECOMPUTE`, `CHK_VACANCY_RECOMPUTE`, and the subtotal ties; the existing OSAR‑template ">20% period‑over‑period line variance" note is separate and stays as a review prompt, not a tool failure.

**STANDING FLAG (always emitted, every run):**
- `CHK_UPB_CONFIRM`: there is no UPB source file yet — always emit "confirm UPB against internal records," echoing the UPB value the model used for each loan.

**BLOCKER — wrong DY if inputs shift; must flag loudly:**
- `CHK_TAX_MAX`: RE Taxes cell = `MAX(T12, tax‑tab invoice)`. Fail if it references only one side.
- `CHK_INS_MAX`: Insurance cell = `MAX(T12, insurance‑tab invoice)`, and the invoice term resolves to the **actual dollar premium** (not a factor‑scaled or millions‑formatted display cell). Fail if it pulls T12 only (Strada) or divides/relabels the invoice to ≈0 (Ares).
- `CHK_VACANCY_SIGN`: vacancy‑loss line must be ≤ 0 and must be 0 when occupancy ≤ 95%; when occupancy > 95% it must **reduce** EGI. Recompute with a synthetic >95% occupancy and confirm the sign.
- `CHK_DY_BASIS`: DY numerator = NCF (post‑reserve), not the pre‑reserve NOI line.
- `CHK_DY_CONSISTENCY`: the OSAR DY is the number of record. If a DY appears on any *other* tab, only check that it **ties** to the OSAR DY; if it doesn't, flag the tab/cell and both values — but do **not** audit the off‑tab calc's internals.

**HIGH — methodology / definition adherence:**
- `CHK_PERIOD`: T12 end‑month, RR as‑of, and AR as‑of all equal quarter‑end.
- `CHK_RESERVE_RATE`: reserve rate matches the loan ($250/unit, $0.25/sf, $0.10/sf, …) and count ties to the property.
- `CHK_MGMT_BASE`: mgmt fee = MAX(actual, 3%×**EGI**). Flag any model whose base is not EGI (e.g., 3%×GPR).
- `CHK_OTHER_INCOME_BASIS`: other income basis matches the loan (T3A/T6A vs T12).
- `CHK_EXCLUSIONS`: delinquency exclusions are derived from the **AR aging report** (source of truth) at the loan's window (45/60 days); verify the RR include/exclude flags against the AR aging and against lease status (MTM/dark/BK/known‑vacate). Recoveries must be subject to the vacancy factor where the definition puts them in clause (a).
- `CHK_MONTH_COUNT`: for every annualized line, the model's formula must sum **all** months available on the source tab for its window (12 for T12, 3 for T3, etc.) — no more, no fewer. If the formula omits an available month, flag it with the line, tab, and month count used vs. available. Assume all months present are full months; do not attempt to detect partial/stub months.

**MEDIUM — linkage & robustness:**
- `CHK_LINK_TARGETS`: every column‑I line links to the intended tab/cell; flag whole‑column SUMs, single‑cell GPR anchors, and mismatched labels.
- `CHK_EXTERNAL_REFS`: flag `[n]` external‑workbook links (cached‑value risk).
- `CHK_COLUMN_H_SOURCE`: the T12/reference column links to the actuals tab consistently (flag one‑off links to a different tab — Ares H47).
- `CHK_UNIT_SF_TIE`: unit/SF count is consistent between RR and OSAR (and CapEx uses the intended count).
- `CHK_DUPLICATE_TABS`: when duplicate period/RR/OSAR tabs exist, confirm the live one is referenced and the stale one isn't feeding anything.

**LOW — hygiene:**
- `CHK_ERROR_CELLS`: report all cached error cells; separate "on DY path" from "off path."
- `CHK_HARDCODE_IN_FORMULA`: report numeric constants embedded in formulas (occupancy plugs, ad‑hoc adjustments) with the cell and value.
- `CHK_ANNUALIZATION_FORMULA`: for short‑history annualizers (`SUM(months)/N*12`), ensure the `/N*12` is inside the `IF` so zero rows don't yield `#VALUE!`.

---

## 5. Resolved house rules (previously open questions)

These were the softest spots — where the models diverged from each other or from a literal read of the text. Each now has a settled rule the tool must implement. (Q‑numbers refer to the original open‑question list, kept for traceability.)

1. **Architecture — verify + targeted recompute (Q1).** The tool **verifies** the model's formulas and links, and **independently recomputes GPR and vacancy** from the raw rent roll, comparing back to the model. It does not attempt a full independent rebuild of all expense lines.
2. **RR exclusion source of truth — AR aging (Q10).** Delinquency exclusions are derived from the **AR aging report** at the loan's window (45/60 days). The tool verifies the analyst's RR include/exclude flags against the AR aging (and against lease status for MTM/dark/BK/known‑vacate).
3. **Management‑fee base — always EGI (Q2).** 3% × **EGI** in every loan, regardless of loan wording or tab notes. Flag any model using a different base (e.g., 3%×GPR).
4. **Vacancy — one correct formula (standing).** Deduction (≤ 0), **0 when occupancy ≤ 95%**, and reduces EGI when occupancy > 95%. Test at 100%, 97%, 95%, 92% occupancy. Sign must be negative in the >95% branch (Strada is inverted — Flag S‑3).
5. **Short history / annualization (Q8).** Annualize whatever months exist by `12 / months_used`; a T‑N line shrinks its window only when data is short, never grows past N. **Assume all months present are full months** — the tool does not detect or drop "partial" months. If the model's formula sums fewer months than are actually available on the source tab, **flag the omission** for the reviewer to inspect (`CHK_MONTH_COUNT`). See §2.4.
6. **Debt‑service rate — actual, never stressed (Q3).** MIN(SOFR, cap) + spread. (Affects DSCR reporting, not DY.)
7. **UPB — no source yet; standing flag (Q7).** Do not verify UPB; always emit a "confirm against internal records" flag echoing the model's UPB value.
8. **DY of record — OSAR tab only (Q5).** Only the OSAR DY must be correct. Off‑tab DY calcs are checked for tie‑out and flagged if they disagree, but not audited.
9. **Output tab selection — visible tab (Q4).** Use the visible OSAR tab; ignore hidden ones. If multiple are visible, use the one with a populated column I at quarter‑end.
10. **Tax/Insurance MAX — universal (Q6).** MAX(T12, invoice) on every loan; a missing or broken side is a BLOCKER (Strada S‑1, Ares A‑2).
11. **Materiality (Q9).** NOI/EGI/Total Opex flag at **> $1.00**; individual revenue lines flag at **> 0.1%**.
12. **Contractual rent steps & free rent (open, manual).** Still hard to verify from a static RR snapshot. The tool flags these for **manual review** rather than attempting to verify them — the one item deliberately left to the reviewer.

---

## 6. Q1 2026 audit findings (flags by loan)

Severity: 🔴 blocker · 🟠 high · 🟡 medium · ⚪ low. "Impact this quarter" notes whether the reported DY is currently affected.

### Strada (Multifamily) — DY (I73) = 6.47% · reserve $250/unit
- 🔴 **S‑1 Insurance MAX not applied.** `I37 = INDEX/MATCH(Operating Statement T12)` only; never compares to `Insurance!A1` ($130,449). Taxes (`I36`) correctly does `MAX(H36, Tax!B3)`. *Impact: none this quarter* (T12 $130,999 > invoice $130,449), but wrong if the invoice ever exceeds T12. **Fix:** `I37 = MAX(H37, Insurance!A1)`.
- 🔴 **S‑3 Vacancy‑loss sign inverted (>95% branch).** `I25 = IF(I16>0.95,(I24/I16)*M25,0)`, `M25 = 5%-(1-I16)`. When occupancy > 95% this yields a **positive** number that *raises* EGI. *Impact: none this quarter* (occ 91.95% → I25 = 0), but wrong for any well‑occupied quarter. **Fix:** negate the >95% branch.
- ✅ **S‑6 Mgmt‑fee base** = 3% of EGI — matches the EGI house rule; no action.
- 🟡 **S‑14 Unit count mismatch.** RR total = 497 units (`L540`), OSAR `E12` = 495 (used for CapEx: 495×$250). Occupancy uses 497. Reconcile.
- ⚪ **S‑12 Off‑path errors.** Unused `Comm OSAR` tab `I78/J78 = #DIV/0!`; `Paydown Scenario`, `Variance Analysis`, `Renovation Schedule` carry `#REF!`. None feed `NEW OSAR`.
- ✅ Correct: period (T12 ending 3/31/26, RR 3/31/26), Tax MAX, reserve rate, NCF/UPB basis.

### Campus at Villa La Jolla (Medical office) — DY (I78) = 6.07% · reserve $0.25/sf
- ✅ **Cleanest of the four.** Vacancy blank/0 is **correct** (actual vacancy 31.56% ≫ 5% floor; GPR is in‑place so it's already reflected). Vacancy formula sign is correct. Tax and Insurance both apply `MAX(T12, invoice)`. Reserve $0.25/sf correct. DY = NCF/UPB.
- ✅ **C‑6 Mgmt‑fee base** = 3% of EGI — matches the EGI house rule; no action.
- 🟡 **C‑14 SF mismatch.** RR total 191,544 sf vs OSAR `E12` 191,454 (CapEx uses 191,454; occupancy uses RR total). Reconcile.
- ⚪ **C‑13 Cosmetic.** `K26 = #VALUE!` (variance column only). Confirm `Comm OSAR` (not `OSAR`) is the authoritative tab. Column‑H TTM uses `XLOOKUP` (fine in Excel; note it won't evaluate under LibreOffice/openpyxl recalculation).

### Hialeah (Industrial) — DY (I78) = 2.15% · reserve $0.25/sf
- 🔴 **H‑4 Two conflicting DYs.** Output `I78 = 2.15%` (NCF/UPB, correct) vs `Debt Service!C16 = 2.33%` (NOI/UPB, omits reserve). Decide which is reported and fix the other.
- 🟡 **H‑5 Hardcoded plug in occupancy.** `'1Q26 RR'!F62 = (F54-F41-F7+34000)/F54`. The `+34,000` sf is an undocumented manual adjustment driving occupancy (44.82%) and thus the vacancy branch. Document or source it.
- ⚪ **H‑6 Mgmt‑fee label mismatch (cosmetic).** Formula correctly uses 3%×**EGI** (`I34`, the house rule), but the on‑tab note says "3% of **GPR**." Calculation is correct; **fix the misleading note** so it doesn't invite a wrong "correction" (3%×GPR would be $42.4K vs the correct $60.7K).
- 🟡 **H‑9 Duplicate/stale tabs.** Live `1Q26 RR` vs stale `1Q RR` (has `#REF!`); multiple T12 tabs (`1Q26`/`2Q`/`3Q`). Output correctly references `1Q26 RR` and `1Q26 T12` — but confirm each quarter.
- 🟡 **H‑M External links.** Property‑overview cells pull from `[1]/[2](Old) Comm OSAR` (cosmetic fields only; cached values).
- ⚪ Whole‑column `SUM('1Q26 RR'!V:V)` and `SUM(R:R)*12` for GPR/recoveries — fragile.
- ✅ Correct: T12 "Ending Period March 2026," Tax MAX (invoice $452.7K > T12), Insurance MAX (T12 > invoice), reserve $0.25/sf, DY basis on output tab.

### Ares 55th Ave (Industrial) — DY (I78) = 7.09% · reserve $0.10/sf · 100% occupied
- 🔴 **A‑2 Insurance MAX invoice term broken.** `Lender Calc!E21 = MAX(Actuals!E24, Insurance!C3/Insurance!C4)` where `C3` is a $‑in‑millions display (0.00918) and `C4` = 1,000,000, so the invoice term ≈ 9.2e‑9 instead of ~$9,176. *Impact: none this quarter* (T12 $13,287 wins), but the invoice can never win. **Fix:** reference the raw premium (`Insurance!Y41` = $9,175.82) or `C3*C4`.
- 🟡 **A‑7 Recoveries in "Other Income."** ~$187K of tenant recoveries sit in Other Income, escaping the 5% vacancy factor; the Ares definition puts recoveries in clause (a), which is subject to it (~$9K NOI). Confirm treatment.
- 🟡 **A‑8 Column‑H G&A link.** `H47 = 'Lender Calc'!E33` while the rest of column H pulls from `Actuals (T12)`. Same value here; likely paste error. Reference the actuals tab.
- ⚪ **A‑10 Annualizer throws `#VALUE!`.** `T12!N:N = IF(SUM(B:M)=0,"",SUM(B:M))/N3*12` — the `/N3*12` is outside the `IF`, so all zero/blank rows return `""` then divide → 99 `#VALUE!` cells. None are on matched data rows, so NOI is unaffected. **Fix:** `=IF(SUM(B13:M13)=0,"",SUM(B13:M13)/$N$3*12)`.
- ⚪ **A‑11** `Debt Service!C32 = #DIV/0!` (empty rate‑cap‑renewal scenario). Off path.
- ⚠️ Vacancy floor **correctly bites** here: 100% occupied → `-GPR×5% = -$39,820`. Good worked example for `CHK_VACANCY_SIGN`.
- ✅ Correct: RR "As of 3/31/2026," Tax MAX (raw dollar), reserve $0.10/sf, DY = NCF/UPB, 7‑month annualization ×12/7.

---

## Appendix — column‑I cell map (Q1 2026)

**Strada `NEW OSAR`:** GPR `I24 = 'Rent Roll'!M542*12 + 'Operating Statement'!R54` · Vacancy `I25` (helper `M25=5%-(1-I16)`) · Parking `I29 = SUMIF(OS!A:A,C29,OS!S:S)` (T3A) · Other Income `I30 = 'Operating Statement'!R55` · Taxes `I36 = MAX(H36, Tax!B3)` · Insurance `I37 = INDEX/MATCH(OS T12)` ⚠️ · Mgmt `I40 = MAX(H40, I32*0.03)` · Reserve `I53 = E12*E14` (495×250) · NOI `I51` · NCF `I57` · DY `I73 = I57/E6`.

**Campus `Comm OSAR`:** GPR `I25 = 'Rent Roll'!Y87` · Vacancy `I26 = IF(O20>O19,"",(O19-O20)*-I25)` (O19=5%, O20=actual vac) · Reimb `I29 = 'Rent Roll'!Z87` · Parking/Other `= XLOOKUP(TTM)` · Taxes `I38 = MAX(H38, Tax!D6)` · Insurance `I39 = MAX(H39, Insurance!R9)` · Mgmt `I43 = MAX(H43, 0.03*I34)` · Reserve `I58 = L58*E12` (0.25×191,454) · NOI `I54` · NCF `I62` · DY `I78 = I62/E6`.

**Hialeah `(New) Comm OSAR`:** GPR `I25 = SUM('1Q26 RR'!V:V)` · Vacancy `I26 = -IF(I16>0.95,(I16-0.95)*I25,0)` · Reimb `I29 = SUM('1Q26 RR'!R:R)*12` · Other `I32 = H32` (T12) · Taxes `I38 = MAX(H38, Tax!B7)` · Insurance `I39 = MAX(H39, Insurance!O4)` · Mgmt `I43 = MAX(I34*0.03, H43)` · Reserve `I58 = E12*E14` (298,769×0.25) · NOI `I54` · NCF `I62` · DY `I78 = I62/E6` · (conflicting `Debt Service!C16 = C9/C7`).

**Ares `Comm OSAR` ← `Lender Calc`:** GPR `E5 = SUM(RR!J8)` · Vacancy `E9 = IF(occ<95%, Actuals!E9, -E5*5%)` · Other Income `E12 = INDEX/MATCH(T12 recoveries)` · Taxes `E19 = MAX(Actuals!E22, 'RE Taxes'!T5)` · Insurance `E21 = MAX(Actuals!E24, Insurance!C3/Insurance!C4)` ⚠️ · Mgmt `E28 = MAX(0.03*E15, Actuals!E31)` · Reserve `Comm OSAR!I58 = D14*D12` (0.09996×100,000) · NOI `I54` · NCF `I62` · DY `I78 = I62/D6` (D6 = Debt Service!C4).
