# US 1117 — SAM OnSite Guard Onboarding MVP Specification

Developer-ready Word specification for **User Story 1117 — SAM OnSite: User onboarding scenario**
(mandatory first-login voice training for security guards).

## Files

- `US1117_SAM_OnSite_Guard_Onboarding_MVP_Spec.docx` — the deliverable.
- `build_us1117_spec.js` — reproducible builder (Node + `docx`). Run: `node build_us1117_spec.js`.

## What the document contains

Purpose & confirmed decisions · scope · roles/ownership · end-to-end flow · voice & screen
script · functional requirements (FR-01–FR-20) · progress & completion model · practice-Activity
lifecycle & data isolation · error/recovery · acceptance criteria (AC-01–AC-13) · UAT matrix
(UAT-01–UAT-18) · requirement traceability matrix · rollout/rollback · Definition of Done ·
release-safety gates · engineering-validation list · evidence basis.

Every requirement carries an evidence tag — **[Verified] / [Proposed] / [To validate] /
[Recommended]** — so nothing reads as fact unless it is. `[Verified]` marks a requirement's
*source or decision*, not that the platform already implements it; implementation feasibility is
tracked in §16.

## How it was verified

1. Structural audit of the actual `.docx` (US Letter geometry, real Word numbering — no fake
   bullets, table geometry, repeated headers, cross-reference integrity, FR→AC→UAT coverage).
2. Unbiased multi-perspective **council**: independent Product/UX, Engineering/Data, QA/Traceability,
   Safety/Privacy, and Document-Quality critics. Their evidence-backed findings were incorporated.
3. A **separate final verifier** applied strict pass/fail gates A–I. It issued **PASS** with no
   blocking defects.

## Guardrails

Nothing here authorises a production change. All work proceeds in DEV/UAT and reaches production
only after Vincent Smeyers' approval. No client template pack or protected system template is
modified. No backend endpoint or database field is invented — unknowns are collected in §16.
