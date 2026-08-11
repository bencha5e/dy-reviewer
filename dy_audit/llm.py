"""Revenue review delegated to an LLM.

Everything mechanical stays in Python. The revenue line items - GPR, the rent
roll it is built from, vacancy, other income, and the T12 reconciliation behind
them - move here, because that work is judgment rather than arithmetic: whether
a Pending-renewal row duplicates a unit already counted, whether pet rent
belongs inside base rent, whether a formula cell in an exported column is a
backfill or a correction. Each of those took a bespoke heuristic to catch.

The workbook is handed over as a real `.xlsx`. Claude's `document` content block
does not accept spreadsheets, so the file goes up through the Files API and is
attached as a `container_upload`, which puts it on the filesystem of the code
execution sandbox. Claude then reads it with openpyxl the same way this package
does - `data_only=False` for formulas, `data_only=True` for cached values - and
can follow a SUMIF into the rent roll instead of reasoning over a flat dump.
The workbook never enters the context window, so a 2.5 MB model costs a few
thousand tokens rather than a few hundred thousand.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .model import Finding, Severity, Status
from .osar import Line

if TYPE_CHECKING:  # pragma: no cover - typing only
    from .context import LoanContext


#: Default model. Opus rather than Sonnet: the failure mode that motivated this
#: rewrite is a trace that stops one hop early, which is exactly where the
#: cheaper model gives way first.
DEFAULT_MODEL = "claude-opus-5"

#: Formula tracing across tabs is the coding/agentic shape that warrants xhigh.
DEFAULT_EFFORT = "xhigh"

MAX_TOKENS = 64_000

#: The server-side code-execution loop pauses every 10 iterations. Each pause
#: costs a round trip, so cap the resumes rather than looping forever.
MAX_CONTINUATIONS = 5

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

#: The revenue rules, as prose for the model. Kept out of this file so it can be
#: edited without touching code, and cached as a stable prompt prefix.
RULES_FILENAME = "DY_Revenue_Reviewer_System_Prompt_LLM.md"


#: Check IDs the LLM owns. Findings the deterministic suite raises under these
#: IDs are discarded when the LLM review is on, so a loan is never reviewed
#: twice for the same defect under the same heading.
#:
#: CHK_PERIOD, CHK_MONTH_COUNT and CHK_MGMT_BASE are deliberately absent. They
#: touch revenue but are pure arithmetic - a date comparison, a divisor against
#: months-with-data, a formula's base reference - and Python does that more
#: cheaply and more reliably than any model.
REVENUE_CHECK_IDS = frozenset(
    {
        "CHK_GPR_RECOMPUTE",
        "CHK_RENT_STATUS",
        "CHK_VACANCY_RECOMPUTE",
        "CHK_RENT_REBUILD",
        "CHK_RENT_EDITS",
        "CHK_RR_TOTAL_TIE",
        "CHK_GPR_SUPPLEMENTARY",
        "CHK_VACANCY_SIGN",
        "CHK_VACANCY_FLOOR",
        "CHK_VACANCY_DOUBLE_COUNT",
        "CHK_OTHER_INCOME_BASIS",
        "CHK_GPR_TREND",
        "CHK_REVENUE_DOUBLE_COUNT",
        "CHK_EXCLUSIONS",
    }
)

#: Raised when the review itself could not be completed.
LLM_CHECK_ID = "CHK_REVENUE_LLM"


# --------------------------------------------------------------------------
# Output schema
# --------------------------------------------------------------------------

#: Mirrors the `Finding` dataclass field for field, so parsing is a constructor
#: call rather than a translation layer. Constraining `check_id` to the IDs the
#: report already knows keeps the Findings sheet's sort and colour coding
#: meaningful; CHK_REVENUE_OTHER is the escape hatch for a defect that fits
#: none of them.
REVENUE_FINDINGS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "findings": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "check_id": {
                        "type": "string",
                        "enum": sorted(REVENUE_CHECK_IDS) + ["CHK_REVENUE_OTHER"],
                    },
                    "severity": {
                        "type": "string",
                        "enum": ["BLOCKER", "HIGH", "MEDIUM", "LOW", "INFO"],
                    },
                    "status": {
                        "type": "string",
                        "enum": ["FLAG", "PASS", "MANUAL_REVIEW", "UNVERIFIABLE"],
                    },
                    "message": {
                        "type": "string",
                        "description": (
                            "What is wrong and what it costs, in one or two sentences. "
                            "Name the dollar impact and the annualisation where there is "
                            "one. Written for a reviewer who has not opened the workbook."
                        ),
                    },
                    "sheet": {"type": ["string", "null"]},
                    "cell": {
                        "type": ["string", "null"],
                        "description": "A1-style coordinate, no sheet prefix.",
                    },
                    "evidence": {
                        "type": ["string", "null"],
                        "description": (
                            "The formula text, cell values, or row counts the finding "
                            "rests on. This is what lets a reviewer check the work."
                        ),
                    },
                    "on_dy_path": {
                        "type": ["boolean", "null"],
                        "description": (
                            "True when the cell feeds the debt-yield calculation, false "
                            "when it does not, null when the distinction does not apply."
                        ),
                    },
                },
                "required": [
                    "check_id",
                    "severity",
                    "status",
                    "message",
                    "sheet",
                    "cell",
                    "evidence",
                    "on_dy_path",
                ],
                "additionalProperties": False,
            },
        }
    },
    "required": ["findings"],
    "additionalProperties": False,
}


# --------------------------------------------------------------------------
# Prompt assembly
# --------------------------------------------------------------------------


def rules_path() -> Path:
    """Location of the revenue rules markdown, repo root by default."""
    override = os.environ.get("DY_REVENUE_RULES")
    if override:
        return Path(override)
    return Path(__file__).resolve().parent.parent / RULES_FILENAME


def load_rules() -> str:
    path = rules_path()
    if not path.is_file():
        raise FileNotFoundError(
            f"Revenue rules not found at {path}. Set DY_REVENUE_RULES to point at "
            f"{RULES_FILENAME}, or run with --no-llm to use the deterministic "
            f"revenue checks instead."
        )
    return path.read_text(encoding="utf-8")


def _params_brief(ctx: LoanContext) -> list[str]:
    """The agreement terms that govern revenue, with the text they came from.

    The quotes matter more than the values. A model told 'the vacancy floor is
    5%' will apply 5%; one shown the clause can notice that the clause says
    something the parser flattened.
    """
    params = ctx.params
    if params is None:
        return ["  (no loan agreement was parsed for this loan)"]

    wanted = [
        ("Vacancy floor", "vacancy_floor"),
        ("Investment-grade carve-out", "ig_vacancy_carveout"),
        ("Management fee %", "mgmt_fee_pct"),
        ("Management fee base, as stated", "mgmt_stated_base"),
        ("Delinquency window", "delinquency"),
        ("Other income window (months)", "other_income_months"),
        ("Concessions window (months)", "concession_months"),
        ("Known-vacate window (days)", "known_vacate_days"),
        ("Rent-step window", "rent_step_window"),
    ]
    lines: list[str] = []
    for label, attribute in wanted:
        parsed = getattr(params, attribute, None)
        if parsed is None:
            continue
        value = getattr(parsed, "value", parsed)
        if value is None:
            continue
        quote = (getattr(parsed, "quote", "") or "").strip()
        lines.append(f"  - {label}: {value}" + (f'  <- "{quote}"' if quote else ""))
    if not lines:
        lines.append("  (nothing governing revenue was parsed out of the agreement)")
    lines.append(
        "  - Management fee base actually enforced: EGI (house rule, overrides the "
        "stated base)"
    )
    return lines


def _rebuild_brief(ctx: LoanContext) -> list[str]:
    """What Python's own rent-roll rebuild found.

    Offered as a second opinion, not an answer. Python locates the tenant table
    from the export's own headers and knows nothing about why a row is there;
    where the two disagree the model should say which is right and why, not
    default to either.
    """
    rebuilds = ctx.facts.get("rebuilds") or []
    if not rebuilds:
        return ["  (Python could not rebuild a rent roll for this loan)"]

    lines: list[str] = []
    for rebuild in rebuilds:
        line = getattr(rebuild, "line", None)
        name = getattr(line, "value", None) or "GPR"
        if not getattr(rebuild, "confident", False):
            notes = "; ".join(getattr(rebuild, "notes", []) or []) or "no detail"
            lines.append(f"  - {name}: table not read confidently ({notes})")
            continue
        included = len(rebuild.included_rows())
        total = len(rebuild.rows)
        rent_column = (rebuild.columns or {}).get("rent")
        basis = "monthly x 12" if rebuild.monthly else "annual"
        lines.append(
            f"  - {name}: rebuilt {rebuild.annualized():,.2f} from column "
            f"{rent_column} ({rebuild.rent_header}, {basis}); {included} of {total} "
            f"rows included; model line reads "
            f"{(rebuild.model_gpr or 0.0):,.2f}"
        )
        for item, amount in getattr(rebuild, "reconciliation", []) or []:
            lines.append(
                f"      * {item}" + (f": {amount:+,.2f}" if amount is not None else "")
            )
    return lines


def _figures_brief(ctx: LoanContext) -> list[str]:
    """The revenue lines as the model currently reports them."""
    wb, tab = ctx.wb, ctx.tab
    wanted = [
        Line.GPR,
        Line.VACANCY,
        Line.BASE_RENT,
        Line.REIMBURSEMENT,
        Line.PERCENTAGE_RENT,
        Line.PARKING,
        Line.OTHER_INCOME,
        Line.EGI,
        Line.OCCUPANCY,
    ]
    lines: list[str] = []
    for line in wanted:
        if not tab.has(line):
            continue
        coord = tab.cell(line)
        if not coord:
            continue
        value = wb.number(tab.sheet, coord)
        formula = wb.formula(tab.sheet, coord)
        rendered = f"{value:,.2f}" if value is not None else "(no cached value)"
        entry = f"  - {line.value}: {tab.sheet}!{coord} = {rendered}"
        if formula:
            entry += f"   formula: {formula}"
        lines.append(entry)
    return lines or ["  (no revenue lines were resolved on the OSAR tab)"]


def build_brief(ctx: LoanContext) -> str:
    """The per-loan half of the prompt. The rules are the stable half."""
    tabs = ctx.facts.get("source_tabs") or {}
    period = ctx.tab.period_end
    period_text = period.strftime("%m/%d/%Y") if period else "not read from the workbook"

    def tab_line(label: str, key: str) -> str:
        name = tabs.get(key)
        return f"  - {label}: {name!r}" if name else f"  - {label}: not found in this workbook"

    sections = [
        f"# Revenue review: {ctx.loan_name}",
        "",
        "The workbook is attached and available in your sandbox. Read it with "
        "openpyxl, twice: `data_only=False` for formula text and `data_only=True` "
        "for the cached values Excel last wrote. Do not recalculate it - several of "
        "these models use functions that blank out on recalculation, and the cached "
        "values are the numbers the analyst actually reported.",
        "",
        "## Where things are",
        f"  - OSAR tab under audit: {ctx.tab.sheet!r}",
        f"  - Debt-yield column on that tab: {ctx.tab.dy_column!r}",
        f"  - Quarter end / statement ending: {period_text}",
        tab_line("Rent roll tab", "rent_roll"),
        tab_line("T12 / operating statement tab", "t12"),
        tab_line("AR aging tab", "ar"),
        "",
        "  Those tab names are what this tool resolved by following the model's own "
        "formulas. Treat them as a starting point: if the GPR line actually draws on "
        "a different sheet, or a second rent roll exists for a commercial schedule, "
        "say so and review the one the model really uses.",
        "",
        "## Revenue lines as the model reports them",
        *_figures_brief(ctx),
        "",
        "## What the loan agreement says",
        *_params_brief(ctx),
        "",
        "## Python's independent rent-roll rebuild",
        "  This is a second opinion, not the answer. It reads the tenant table from "
        "the export's own headers and has no idea why any given row is there. Where "
        "your read disagrees with it, say which is right and why.",
        *_rebuild_brief(ctx),
        "",
        "## What to return",
        "Work through the revenue rules in your instructions against this workbook, "
        "then return findings under the JSON schema you have been given. Emit a "
        "finding for every rule you evaluated, including the ones that passed - a "
        "silent rule is indistinguishable from one that was never run. Where you "
        "could not complete a check, say so with MANUAL_REVIEW or UNVERIFIABLE "
        "rather than passing it.",
    ]
    return "\n".join(sections)


# --------------------------------------------------------------------------
# Response parsing
# --------------------------------------------------------------------------

_FENCE = re.compile(r"```(?:json)?\s*(.*?)```", re.S)


def _iter_text_blocks(content: Any) -> list[str]:
    out: list[str] = []
    for block in content or []:
        if getattr(block, "type", None) == "text":
            text = getattr(block, "text", "")
            if text:
                out.append(text)
    return out


def extract_payload(content: Any) -> dict:
    """Pull the findings object out of a response.

    With `output_config.format` the JSON arrives as a plain text block, but a
    code-execution turn puts server-tool blocks around it, so scan from the end
    rather than assuming position 0. The fenced-code fallback covers the case
    where the schema is refused and the model falls back to prose.
    """
    candidates = _iter_text_blocks(content)
    for text in reversed(candidates):
        stripped = text.strip()
        for attempt in (stripped, *(m.group(1).strip() for m in _FENCE.finditer(text))):
            if not attempt.startswith("{"):
                continue
            try:
                payload = json.loads(attempt)
            except json.JSONDecodeError:
                continue
            if isinstance(payload, dict) and "findings" in payload:
                return payload
    raise ValueError(
        "The revenue review returned no parseable findings object. "
        f"Text blocks received: {len(candidates)}."
    )


def _coerce(value: Any, enum: Any, default: Any) -> Any:
    try:
        return enum(str(value).strip().upper())
    except (ValueError, AttributeError):
        return default


def to_findings(payload: dict) -> list[Finding]:
    """Turn the parsed JSON into Findings the report can already render."""
    findings: list[Finding] = []
    for item in payload.get("findings") or []:
        if not isinstance(item, dict):
            continue
        message = (item.get("message") or "").strip()
        if not message:
            continue
        check_id = (item.get("check_id") or "CHK_REVENUE_OTHER").strip()
        cell = item.get("cell")
        if isinstance(cell, str) and "!" in cell:
            # A model that writes "Sheet!B12" into the cell field would break the
            # Findings sheet's Cell column, which is A1-only.
            cell = cell.rsplit("!", 1)[1]
        findings.append(
            Finding(
                check_id,
                _coerce(item.get("severity"), Severity, Severity.HIGH),
                _coerce(item.get("status"), Status, Status.MANUAL_REVIEW),
                message,
                sheet=item.get("sheet") or None,
                cell=cell or None,
                evidence=(item.get("evidence") or None),
                on_dy_path=item.get("on_dy_path"),
            )
        )
    return findings


def unverifiable(reason: str) -> Finding:
    """The finding raised when the review could not run.

    HIGH and UNVERIFIABLE rather than a silent skip: the house rule is that a
    check the tool could not complete must never read as a clean pass, and
    UNVERIFIABLE has its own fill on the Findings sheet.
    """
    return Finding(
        LLM_CHECK_ID,
        Severity.HIGH,
        Status.UNVERIFIABLE,
        "The revenue line items were not reviewed: the model review did not "
        f"complete ({reason}). Every revenue check on this loan is outstanding - "
        "review GPR, the rent roll behind it, vacancy, and other income by hand.",
        evidence=reason,
        on_dy_path=True,
    )


# --------------------------------------------------------------------------
# Anthropic
# --------------------------------------------------------------------------


def _anthropic_client():
    try:
        import anthropic
    except ImportError as exc:  # pragma: no cover - depends on install state
        raise RuntimeError(
            "The anthropic package is not installed. `pip install -r "
            "requirements.txt`, or run with --no-llm."
        ) from exc
    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        raise RuntimeError(
            "ANTHROPIC_API_KEY is not set. Export it, or run with --no-llm to use "
            "the deterministic revenue checks."
        )
    return anthropic.Anthropic()


def _call_anthropic(
    xlsx_path: Path,
    rules: str,
    brief: str,
    *,
    model: str,
    effort: str,
    client: Any = None,
) -> tuple[dict, dict]:
    """Run the review. Returns (payload, usage)."""
    client = client or _anthropic_client()

    with xlsx_path.open("rb") as handle:
        uploaded = client.beta.files.upload(
            file=(xlsx_path.name, handle, XLSX_MIME),
            betas=["files-api-2025-04-14"],
        )

    try:
        messages: list[dict] = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": brief},
                    {"type": "container_upload", "file_id": uploaded.id},
                ],
            }
        ]
        request = {
            "model": model,
            "max_tokens": MAX_TOKENS,
            # The rules are identical for every loan in a run, so loans 2..n read
            # this prefix from cache instead of paying for it again.
            "system": [
                {"type": "text", "text": rules, "cache_control": {"type": "ephemeral"}}
            ],
            "tools": [{"type": "code_execution_20260521", "name": "code_execution"}],
            "output_config": {
                "effort": effort,
                "format": {"type": "json_schema", "schema": REVENUE_FINDINGS_SCHEMA},
            },
            # Code execution is GA; referencing an uploaded file still needs the
            # Files API beta on the message call as well as on the upload.
            "extra_headers": {"anthropic-beta": "files-api-2025-04-14"},
        }

        usage_total = {"input_tokens": 0, "output_tokens": 0, "cache_read_input_tokens": 0}
        content: Any = []
        for _ in range(MAX_CONTINUATIONS + 1):
            with client.messages.stream(messages=messages, **request) as stream:
                response = stream.get_final_message()

            usage = getattr(response, "usage", None)
            for key in usage_total:
                usage_total[key] += getattr(usage, key, 0) or 0

            content = response.content
            if getattr(response, "stop_reason", None) != "pause_turn":
                break
            # The server-side tool loop hit its iteration cap. Hand the paused
            # turn back and it resumes where it stopped; without this the review
            # is silently truncated on the larger workbooks.
            messages = messages[:1] + [{"role": "assistant", "content": content}]
        else:
            raise RuntimeError(
                f"the code-execution loop still had work after "
                f"{MAX_CONTINUATIONS} resumes"
            )

        return extract_payload(content), usage_total
    finally:
        try:
            client.beta.files.delete(uploaded.id)
        except Exception:  # noqa: BLE001 - cleanup must not mask a real failure
            pass


# --------------------------------------------------------------------------
# OpenAI
# --------------------------------------------------------------------------


def _call_openai(
    xlsx_path: Path,
    rules: str,
    brief: str,
    *,
    model: str,
    effort: str,
    client: Any = None,
) -> tuple[dict, dict]:
    """Structural fallback. Same contract, same Findings out.

    OpenAI's equivalent of container_upload is a code-interpreter container, so
    the file is uploaded with `purpose="assistants"` and attached to a tool
    resource rather than to the message.
    """
    if client is None:
        try:
            from openai import OpenAI
        except ImportError as exc:  # pragma: no cover - optional dependency
            raise RuntimeError(
                "The openai package is not installed; install it or use "
                "--llm-provider anthropic."
            ) from exc
        client = OpenAI()

    with xlsx_path.open("rb") as handle:
        uploaded = client.files.create(file=(xlsx_path.name, handle), purpose="assistants")

    try:
        response = client.responses.create(
            model=model,
            reasoning={"effort": "high" if effort in ("xhigh", "max") else effort},
            instructions=rules,
            input=[{"role": "user", "content": [{"type": "input_text", "text": brief}]}],
            tools=[
                {
                    "type": "code_interpreter",
                    "container": {"type": "auto", "file_ids": [uploaded.id]},
                }
            ],
            text={
                "format": {
                    "type": "json_schema",
                    "name": "revenue_findings",
                    "schema": REVENUE_FINDINGS_SCHEMA,
                    "strict": True,
                }
            },
        )
        payload = json.loads(response.output_text)
        usage = getattr(response, "usage", None)
        return payload, {
            "input_tokens": getattr(usage, "input_tokens", 0) or 0,
            "output_tokens": getattr(usage, "output_tokens", 0) or 0,
            "cache_read_input_tokens": 0,
        }
    finally:
        try:
            client.files.delete(uploaded.id)
        except Exception:  # noqa: BLE001
            pass


PROVIDERS = ("anthropic", "openai")


def _provider(name: str):
    """Resolve the provider at call time.

    Looked up on each call rather than bound into a table at import, so a test
    can substitute the transport without reaching inside this module's state.
    """
    if name == "anthropic":
        return _call_anthropic
    if name == "openai":
        return _call_openai
    return None


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------


def review_revenue_with_llm(
    ctx: LoanContext,
    *,
    provider: str = "anthropic",
    model: str | None = None,
    effort: str = DEFAULT_EFFORT,
    client: Any = None,
) -> list[Finding]:
    """Review this loan's revenue lines and return findings.

    Never raises. A review that cannot run produces one UNVERIFIABLE finding, so
    the deterministic findings still ship and the gap is visible on the report
    rather than reading as a clean revenue section.
    """
    call = _provider(provider)
    if call is None:
        return [unverifiable(f"unknown provider {provider!r}")]

    try:
        rules = load_rules()
        brief = build_brief(ctx)
        payload, usage = call(
            ctx.files.xlsx_path,
            rules,
            brief,
            model=model or DEFAULT_MODEL,
            effort=effort,
            client=client,
        )
    except Exception as exc:  # noqa: BLE001 - a failed review must not fail the loan
        return [unverifiable(f"{type(exc).__name__}: {exc}")]

    ctx.facts["llm_usage"] = usage
    findings = to_findings(payload)
    if not findings:
        return [unverifiable("the review returned no findings")]
    return findings


__all__ = [
    "DEFAULT_EFFORT",
    "DEFAULT_MODEL",
    "LLM_CHECK_ID",
    "REVENUE_CHECK_IDS",
    "REVENUE_FINDINGS_SCHEMA",
    "build_brief",
    "extract_payload",
    "load_rules",
    "review_revenue_with_llm",
    "rules_path",
    "to_findings",
    "unverifiable",
]
