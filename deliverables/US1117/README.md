# SAM OnSite — Guard Onboarding (MVP specification)

A concise, professional **4-page** Word specification for the SAM OnSite guard onboarding
feature (a mandatory first-login voice training that teaches guards to ask a question, report an
incident, and correct a report — using safe, made-up data).

## Files

- `SAM_OnSite_Guard_Onboarding_Spec.docx` — the deliverable.
- `build_spec.js` — reproducible builder (Node + `docx`). Run: `node build_spec.js`.

## What it contains

In-one-minute summary · how it works step by step · a single "what must be true and how we
confirm it" requirements table · the two things that must not go wrong (data isolation &
completion integrity) · handling the unhappy paths · what to confirm with engineering · rollout,
guardrails, and definition of done.

Every requirement is self-contained (no cross-references to chase) and carries a plain-language
status — **Confirmed** (agreed behaviour), **Proposed** (recommended default), or **Check with
engineering** (needs a platform confirmation) — so nothing reads as fact unless it is.

## How it was verified

1. Structural audit of the actual `.docx` (real Word numbering, table geometry, margins, and a
   check that no ticket/user-story references appear in the content).
2. An unbiased multi-perspective **council** — clarity/no-gaps, completeness/fidelity,
   safety/data-integrity, and document-design critics — whose evidence-backed findings were
   incorporated.
3. A **separate independent verifier** applying strict pass/fail gates (≤4 pages · no ticket
   references · effortless clarity · fidelity · honesty · safety · professional quality). It
   issued **PASS** with no blocking defects.

## Design

Georgia serif headings, a restrained slate + teal palette (not default Office blue), eyebrow
section numerals, real Word list definitions, clean tables with repeating headers, and a running
header/footer with page numbers — a polished, human-authored look.

## Guardrails

Describes intended behaviour for internal build and testing only. No production change is made
until Vincent Smeyers has signed off. Universal for all guards (no customer-specific setup); no
existing templates are changed.
