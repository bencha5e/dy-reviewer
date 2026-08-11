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

python audit_dy.py --input-dir "…/Input DY Tests"              # review and drain the queue
python audit_dy.py --input-dir "…/Input DY Tests" --no-move    # review in place, change nothing
python audit_dy.py --input-dir "…" --loan Strada               # one loan
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
  why, and a flag count by severity.

**Moving is on by default**, so the input folder stays a queue of files not yet
reviewed: after a clean run it holds nothing but the output folder. A loan that
*failed* is never moved — its source files stay put so the next run retries it.
Use `--no-move` to review in place and leave the folder exactly as it was, which
is the safe way to preview a run.

## What it checks

A hybrid: it verifies the model's own formulas and links line by line, and
independently rebuilds two lines — gross potential rent and vacancy — from the
rent roll. Ordinary operating expenses are verified as a T12 pass-through rather
than rebuilt.

| Severity | Checks |
|---|---|
| BLOCKER | tax MAX, insurance MAX, vacancy sign, DY basis, DY consistency |
| HIGH | vacancy floor, double-counted vacancy, period tie-out, reserve rate, management-fee base, other-income basis, exclusions, month count, GPR and vacancy recompute, rent-status inclusion, prior-quarter GPR trend, revenue double-count |
| MEDIUM | external references, link targets, reference-column source, unit/SF tie, duplicate tabs |
| LOW | cached error cells, hardcoded plugs, short-history annualisers |
| STANDING | UPB confirmation — emitted every run |

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
