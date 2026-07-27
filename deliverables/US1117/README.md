# SAM OnSite — Guard Onboarding (specification set)

Specification for the SAM OnSite guard onboarding feature: a mandatory first-login voice
training that teaches guards to ask a question, report an incident, and correct a report — using
safe, made-up data. Delivered as two matched documents.

## Deliverables

| File | Length | Audience | Purpose |
|---|---|---|---|
| `SAM_OnSite_Guard_Onboarding_Spec.docx` | 4 pages | Everyone | The brief — understand the feature in one read. |
| `SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx` | ~10 pages | Engineering + QA | Full requirements, acceptance criteria, test matrix, and traceability. |

Builders (Node + `docx`): `build_spec.js` and `build_detailed_spec.js`. Run either with `node <file>`.

The two documents share one visual identity and are kept factually consistent; the detailed spec
carries FR-01–20 → AC-01–15 → UAT-01–21 with a full traceability matrix behind the brief's claims.

## Key product fact (per Vincent Smeyers)

The archive / "training" flag exists and can be used now, but the **filter that excludes that data
from live surfaces is not built yet** — archived data is currently still visible in production.
Both documents state this plainly: the exclusion filter is a required, not-yet-built enhancement,
and the feature **must not go to production until it is built and verified**. Until then it runs
only in DEV/UAT. This is the single dependency that gates go-live.

## Evidence discipline

Every requirement carries a plain-language status: **Confirmed** (agreed behaviour), **Proposed**
(recommended default), or **Check with engineering** (needs a platform confirmation). Nothing reads
as fact unless it is; the added urgent-incident safety path is labelled as added scope, and the
exact app update routes / progress schema are flagged for engineering confirmation.

## How it was verified

1. Structural audit of each `.docx` (real Word numbering, table geometry, margins, no ticket
   references in content) and an ID cross-consistency check on the detailed spec.
2. An unbiased multi-perspective **council** (clarity, completeness/fidelity, safety/data-integrity,
   traceability, honesty/consistency, and document design) whose evidence-backed findings were
   incorporated.
3. A **separate independent verifier** applied strict pass/fail gates to each document and issued
   **PASS** with no blocking defects.

## Guardrails

Describes intended behaviour for internal build and testing only. No production change is made
until Vincent Smeyers has signed off. Universal for all guards (no customer-specific setup); no
existing templates are changed.
