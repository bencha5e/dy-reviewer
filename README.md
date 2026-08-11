# DebtYield-Reviewer

Reviews quarterly debt-yield (DY) test models against the loan agreements that
govern them, one Excel workbook per loan.

```
Debt Yield = Net Cash Flow / Unpaid Principal Balance
```

The question this answers is **whether the workbook is right and the debt yield
is calculated correctly** — not whether the loan clears its covenant. Testing the
reported yield against a threshold happens in a separate workflow, so nothing
here compares one.

NCF, not NOI. Each loan's definition of NOI already nets a replacement reserve,
and these models place that reserve below the NOI line inside Capital Items — so
the model's **NCF line is the definitional NOI**. Driving the yield off the NOI
line overstates it on every loan.

## Running it

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=...                                   # for the revenue review

python audit_dy.py --input-dir "…/Input DY Tests"              # review and drain the queue
python audit_dy.py --input-dir "…/Input DY Tests" --no-move    # review in place, change nothing
python audit_dy.py --input-dir "…" --loan Strada               # one loan
python audit_dy.py --input-dir "…" --no-llm                    # deterministic only, no network
```

One run reviews every loan pair in the folder. Each loan is processed
independently, so a workbook that cannot be opened is recorded as that loan's
failure and leaves every other loan unaffected.

Outputs, under `<input-dir>/DY Review Output/` unless `--output-dir` says
otherwise:

- `<Loan> - <YYYY_MM_DD> - v<N>/` — one folder per loan per run, holding the
  findings workbook and (with `--move`) the source model and definitions file.
  A second run the same day makes `v2`; an existing folder is never written into.
- `DY Audit Run Log - <timestamp>.md` — which loans succeeded, which failed and
  why, and a flag count by severity. Where the revenue review ran, it also
  carries what that cost: input, cache write, cache read and output tokens per
  loan, and the share of prompt tokens served from cache. The first loan of a
  run writes the rules prefix and reads none of it back; the loans after it
  should show most of their prompt coming from cache.

**Moving is on by default**, so the input folder stays a queue of files not yet
reviewed: after a clean run it holds nothing but the output folder. A loan that
*failed* is never moved — its source files stay put so the next run retries it.
Use `--no-move` to review in place and leave the folder exactly as it was, which
is the safe way to preview a run.

## How the work is split

Two engines, one report.

**Python does everything mechanical** — cell errors, date and period alignment,
cross-footing, the expense side, the reserve rate, the management-fee base, link
targets, month counts. This is arithmetic against a known rule, and code does it
more cheaply and more reliably than a model would.

**An LLM reviews the revenue lines** — gross potential rent and the rent roll
behind it, vacancy, other income, and the T12 reconciliation. That work is
judgment rather than arithmetic: whether a Pending-renewal row duplicates a unit
already counted, whether pet rent belongs inside base rent, whether a formula
cell in an exported column is a backfill or a correction. Every one of those took
a bespoke heuristic to catch, and the next workbook shape needed another.

The workbook goes to the model as a real `.xlsx`, uploaded through the Files API
and attached as a `container_upload`, so it lands on the filesystem of the code
execution sandbox. The model reads it with openpyxl the same way this package
does — `data_only=False` for formulas, `data_only=True` for cached values — and
can follow a `SUMIF` into the rent roll rather than reasoning over a flat dump.
The workbook never enters the context window, which is what keeps the cost down.
The rules it works from live in `DY_Revenue_Reviewer_System_Prompt_LLM.md`, sent
as a cached prompt prefix so the second and later loans in a run don't pay for it
again.

**The loan agreement goes with it, verbatim.** The review's first phase builds a
term sheet — delinquency threshold, notice-to-vacate window, the vacancy floor
*and which revenue it attaches to*, the trailing periods — from the clause text
rather than from parsed values, because a regex flattens a clause and the limb it
drops is often the one that matters. The parsed `LoanParams` go in alongside as a
cross-check, and where the two disagree that is its own finding
(`CHK_TERM_SHEET_CONFLICT`): the model's reading of the clause wins, and the
disagreement usually means the parser missed a limb.

**Only defined terms are testable.** Words like *dark*, *bankruptcy*,
*investment grade*, *free rent*, *month-to-month*, *percentage rent* and *rent
steps* appear inside these NOI definitions but are never themselves defined —
no window, no threshold, no screen. The prompt names them and forbids building a
test around any of them, because a workbook cannot depart from a rule its
contract does not state, and a finding of that shape is a false positive. The
mirror case is handled too: an exclusion the workbook applies that the agreement
does *not* require understates revenue, and is recorded as a memo.

Findings come back as one JSON object per loan. The findings themselves merge
into the same list as the deterministic ones, sort by the same key, and render
into the same eight columns — that sheet is unchanged. What the review returns
beyond a finding — the term sheet, the line-by-line rebuild, the reviewer's prose
note, and each finding's rule number, clause, dollar impact and recommendation —
has no column there, so it lands on a new **Revenue Review** sheet.

`--no-llm` falls back to the deterministic revenue checks, which are still in the
codebase and still tested. It needs no API key and no network — use it for
offline runs and for reproducing a prior quarter's output exactly.

**A revenue review that cannot run is never silent.** If the API call fails, the
key is missing, or the rules file is absent, the loan still produces its full
deterministic report plus one `CHK_REVENUE_LLM` finding at HIGH/UNVERIFIABLE
saying the revenue lines are outstanding. It is counted in the severity totals
and coloured distinctly from a pass.

## What it checks

It verifies the model's own formulas and links line by line, and independently
rebuilds two lines — gross potential rent and vacancy — from the rent roll.
Ordinary operating expenses are verified as a T12 pass-through rather than
rebuilt.

Checks marked ✳︎ are the revenue set. With the LLM review on they are answered by
the model; with `--no-llm` the deterministic implementations answer them instead.
Either way they appear under the same check IDs, so the report reads the same.

| Severity | Checks |
|---|---|
| BLOCKER | tax MAX, insurance MAX, ✳︎vacancy sign, DY basis, DY consistency |
| HIGH | ✳︎vacancy floor, ✳︎double-counted vacancy, period tie-out, reserve rate, management-fee base, ✳︎other-income basis, ✳︎exclusions, month count, ✳︎GPR and vacancy recompute, ✳︎rent-status inclusion, ✳︎prior-quarter GPR trend, ✳︎revenue double-count, ✳︎rent-roll rebuild, ✳︎rent-column edits, ✳︎export-total tie, ✳︎supplementary-income-in-GPR |
| MEDIUM | external references, link targets, reference-column source, unit/SF tie, duplicate tabs |
| LOW | cached error cells, hardcoded plugs, short-history annualisers |
| STANDING | UPB confirmation — emitted every run |

Period tie-out, month count and the management-fee base stay with Python even
though they touch revenue. They are a date comparison, a divisor against
months-with-data, and a formula's base reference — arithmetic, not judgment.

**The rent roll is rebuilt from scratch on every run**, under both settings. The
deterministic rebuild keeps running with the LLM review on: it produces the Rent
Roll Rebuild tab and the source-tab resolution later checks depend on, and it
goes to the model as a second opinion to agree or disagree with. What changes is
only who reaches the verdict. The tool finds the
tenant table by its own headers — never by trusting the model's formulas —
identifies the tenant-rent column, decides which tenants belong in base rent,
and writes the whole thing to a **Rent Roll Rebuild** tab in the findings
workbook: every row, its status, its rent, included or excluded and why. If the
rebuilt figure doesn't tie to the OSAR's annualized GPR line, the run does a
second, deeper dive and itemizes the difference — rent on excluded statuses,
supplementary billing codes inside the base-rent line (the GPR line is base
rent only), concession deductions, per-row deltas in a derived rent column,
hand-edited cells, and export totals that no longer match the rows above them.
Three forensic signals flag even when the headline number ties:

- **a formula cell inside an otherwise-literal exported rent column** that
  pulls from another period or workbook — last quarter's rent typed over units
  the export shows producing nothing;
- **the export's own Total rows disagreeing with the tenant rows above them**
  — the sheet was edited after export;
- **the OSAR's note claiming exclusions the formula never applies** ("less
  delinquent tenants & KV's" beside a plain SUM that excludes nobody).

Three of the HIGH checks exist because a workbook once sailed through review
with $572,580/yr of phantom rent (the Lydian fixture, now part of the
acceptance suite):

- **The GPR recompute traces to tenant rows, never to summary blocks.** A rent
  roll that wraps its per-status `SUMIF`s inside a `SUM` used to stop the trace
  at the block — re-summing the model's own subtotals and proving nothing. The
  trace now expands small formula-bearing ranges down to the tenant rows and
  rebuilds rent from occupied-status rows only.
- **Rent-status inclusion.** Rows whose status is Vacant, Applicant, or
  Pending must not contribute rent: an applicant hasn't moved in, and a
  pending-renewal row duplicates a unit already counted as Occupied, so its
  rent lands twice. Rows marked `VACANT` (by status or by tenant name) must
  carry zero rent outright.
- **Prior-quarter trend.** GPR ÷ occupancy — rent per occupied unit — is
  compared against the prior DY test in column G and flags beyond a 5% move.
  Genuine leasing shifts this number slowly; phantom rent moves it instantly
  (Lydian: +11.9% in one quarter).

A fourth, the revenue double-count check, verifies no source cell enters two
revenue lines with the same sign and no concession-labelled deduction is taken
twice — deliberate netting (a parking row added to Parking and subtracted from
Other Income) stays silent.

Findings carry one of four statuses. `MANUAL_REVIEW` and `UNVERIFIABLE` exist so
that a check the tool could not complete never reads as a clean pass — a rent
roll it cannot parse produces no variance flag at all rather than a made-up one.

## Things worth knowing

**Loan parameters come from the loan agreement, not from constants.** The
vacancy floor, reserve rate, management fee, and delinquency window are parsed
out of each `*_DYDefinitions.md`, and every parsed value is reported with the
sentence it came from. Where the build spec's parameter table disagrees with the
loan text, the loan text governs and the conflict is reported rather than
silently resolved.

**Nothing is keyed to a cell address.** The spec's appendix is a Q1 2026
snapshot; columns and tabs drift every quarter. Lines are resolved by label
(which lives in column C on three of the four models and column B on the
fourth), the DY column is derived from the model, and the UPB cell is read off
the debt-yield formula's denominator.

**The workbook is never recalculated.** One model builds its reference column on
`XLOOKUP`, which neither openpyxl nor LibreOffice evaluates, so a recalculation
pass would blank out real figures. Each file is opened twice — once for formulas,
once for the values Excel cached — and audited against the cached values.

**Some defects are invisible in this quarter's numbers.** A vacancy formula with
an inverted sign returns 0 at today's occupancy and looks healthy; it only
misbehaves above the floor. Those checks re-evaluate the model's own formula
under substituted occupancy rather than reading the current result.

## Development

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -q
```

The tests are acceptance tests: they assert the exact findings expected on the
four Q1 2026 models, **in both directions**. A check that fires on a line known
to be correct fails the suite just as a missed defect does. Anything that writes
or moves files runs against copies in `tmp_path`.
