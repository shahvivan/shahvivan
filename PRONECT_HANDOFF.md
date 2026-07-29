# PRONECT — COMPLETE HANDOFF

**Prepared:** 27 July 2026
**Purpose:** Continue all Pronect work (currently: User Story 1117, SAM OnSite guard onboarding) in a new chat without losing any context, decisions, corrections, files, or working method.
**Repo:** `shahvivan/shahvivan`, branch `claude/ticket-deliverable-verify-jxv715`
**Status:** Two deliverables built, verified, and pushed. User is learning to operate independently. An email to Vincent Smeyers and Jeroen has been drafted but **not yet sent**.

Read this whole file before doing anything. It is written so a new chat with zero memory can pick up exactly where this one left off.

---

## 1. WHO AND WHAT

**User:** Vivan Shah, at Pronect (a Belgian physical-security SaaS company). Not a technical person — learning as he goes, explicitly said so, and is trying to reach the level of a competent non-technical product owner, not an engineer.

**Product:** SAM OnSite — a voice-first guard assistant. A guard speaks naturally to SAM, and SAM produces structured, time-stamped, geo-tagged incident reports ("Activities").

**People:**
- **Vivan Shah** — ticket owner/requester (the user in this chat).
- **Vincent Smeyers** — Chief Product & Security Officer. Approves every production change. Gave the critical correction in §4 below.
- **Joaquim Mathias** — Technical Lead.
- **Berk** — backend/Supabase specialist.
- **Jordi** — SAM OnSite focus.
- **Sartak** — AI focus.
- **Jeroen** — engineering reviewer (surname not given); will be one of the two recipients of the outstanding email.

**The ticket (User Story 1117 — "SAM OnSite: User onboarding scenario"):**
> As a new security guard, I want to complete a simple, stage-by-stage voice training template upon my first login, so that I can quickly master core features — like asking questions, editing logs, and generating reports — by simply following SAM's voice instructions.

Trigger: fires automatically after first successful login; can be reopened later as a refresher. Original note: investigate whether visual aids (tooltips, highlights) are needed for dark/noisy environments.

**Critical clarification the user gave early on:** "new security guard" does **not** mean only newly-created accounts — it means every guard must get onboarding once, including guards who already existed before this feature shipped.

---

## 2. HOW THE PRIOR INVESTIGATION WAS DONE (before this chat — inherited context)

This work began in an earlier tool ("Codex"), whose full handoff was uploaded as a file at the start of this chat: `/root/.claude/uploads/33332f4f-4aae-5e5f-903b-d2a72bb5aa36/52eeb5d0-US1117_COMPLETE_CHAT_HANDOFF.md`. That file is long (828 lines) and this section distills everything from it that matters. If more raw detail is ever needed, that path may still be readable, but treat it as reference only — its Windows file paths do not exist in this Linux repo environment.

**What was investigated and how:**

- **DEV management console** (read-only, no changes made): inspected clients, templates, and users. Active pack in DEV: `BLECKMANN-2306-v3 (2)`, 76 templates, 1 client. **No onboarding-completion field exists anywhere in the user form** (fields present: Role, Client, Username, Auth Email, PIN Code, Contact Email, Names, Language, Site Assignments, Active toggle). This is why the spec proposes adding a new progress-tracking record rather than assuming one exists.
- **UAT reporting page** (`https://uat.pronect-it.com/platform-event`): "Sam OnSite Activities" has **Edit** and **Delete** actions. The Edit modal exposes: Source, Type, Title, Locations, User, User Email, Geo, Time Zone, Started At, Ended At, Description. Delete opens a confirmation — Vincent later confirmed this **is** soft-delete/archive behaviour.
- **Source documents** (already extracted to text, not this repo): a 109-page "Pronect SAM Four-System Master Source" PDF and a 5-page "Pronect Source Information" PDF. These describe SAM's voice-first behaviour and confirm it produces categorized, time-stamped, geo-tagged reports.
- **The Bleckmann template pack JSON**: 76 templates total. Runtime tools available to SAM include `context_get`, `context_set`, `checklist_advance`, `procedure_advance`, `reclassify_scenario`, `clarification_needed`. **`context_set` is what lets a correction (e.g. "change Dell to MacBook") update structured report data through conversation, even after a scenario is otherwise finished.** No pack tool covers login state, app navigation, onboarding completion, or reporting suppression — those remain application-layer (mobile app + backend) work, not something a template pack can do.
- **No Azure Repos access** — so the exact backend API routes, database field names, and endpoints were never confirmed. The spec deliberately never invents these; they are listed as open engineering questions.
- **No production change was ever made.** All investigation was read-only, in DEV/UAT.

**How the "how did we figure out SAM's behaviour" question breaks down** (asked directly in this chat, worth preserving verbatim as the accurate answer): the **console + UAT reporting page** showed that Activities can be edited/archived and that no onboarding field exists. The **source documents + template pack** showed how SAM itself behaves and how corrections are written via `context_set`. It is not accurate to say the console alone explained SAM's behaviour — it's a combination, and the split matters if asked again.

---

## 3. ALL CONFIRMED PRODUCT DECISIONS (do not reopen without the user's explicit say-so)

1. Every guard gets the onboarding exactly once, including guards who existed before this feature — not just new accounts.
2. It is mandatory; there is no way to skip it to reach the home screen (except the urgent-incident detour, see below — that is not a "skip").
3. "Generate a report" = report a fictional incident by talking to SAM.
4. Applies to all guards; universal, not customer/client-specific configuration.
5. A completed guard can voluntarily reopen it later as a refresher, without losing their completed status.
6. English only for this MVP.
7. Enhanced visual support (tooltips, highlight animations for dark/noisy environments) may be evaluated later; explicitly deferred, not required now.
8. If microphone access is denied, keep prompting; never leave the guard stuck (if it's truly unavailable — broken hardware/device policy — route to a support screen, but the emergency action must still be reachable there).
9. Completion is decided by the system observing the guard actually perform each real action — never a quiz, never a self-reported "I'm done" button.
10. "Editing logs" = editing the guard's own incident report/conversation. Two paths must both work: (a) voice — tell SAM to change the fact; (b) manual — long-press the guard's own sent message, edit it, save. Worked example used throughout: a guard reports "a Dell laptop was stolen," corrects it to "MacBook," and separately edits to add "at reception." The final report must show **both** "MacBook" and "reception," and no duplicate report may ever be created.

**The safety measure the AI added (not an original ticket requirement — always labelled as such):** a narrowly-scoped, always-visible "Report an urgent incident" action available throughout onboarding, usable even without a microphone, opening normal incident reporting and returning the guard to their saved onboarding stage afterward. It never counts as completing or skipping onboarding. Rationale: guards do safety-critical work, so a mandatory training gate must never be able to block a real emergency.

---

## 4. THE SINGLE MOST IMPORTANT CORRECTION — READ THIS CAREFULLY

Partway through this chat, the user shared two screenshots of a Microsoft Teams conversation with **Vincent Smeyers**. The exact exchange:

> **Vivan:** "I need to know can we flag a fictional SAM incident conversation and its generated Activity as `training`, while still allowing the guard to correct it by voice or by long-pressing and editing their message? It must be excluded from operational Activity feeds, reports, alerts, webhooks, analytics and KPIs, then deleted after onboarding while retaining only the user's completion status. If supported, what mechanism is used; if not, where is a backend change required?"
>
> **Vincent:** "Put it in archive mode (soft delete). We need to update the reporting layer to exclude this data. But that flag was created for that purpose of exclusion. So you can use it. Later we can also do a data cleansing in database."
>
> *(later, separately)* **Vincent:** "But again be aware these are still visible in prod. Filtering is an enhancement that needs to be carried out at a later stage."

**What this means, precisely:**
- The archive/soft-delete flag **already exists** on the platform and **can be used now** to mark a report as training data. This part is real and usable today.
- The **filter** that actually makes archived/training data invisible on live/production surfaces — reports, dashboards, alerts, webhooks, analytics, KPIs — **does not exist yet**. Today, archived data is **still visible in production**.
- Building that filter is a **separate, later-stage enhancement**, not something already done.

**Why this mattered so much:** the first draft of the specification (before this correction) said practice data was "excluded from every operational consumer throughout its lifecycle" — stated as settled fact. That was **false as of today** and had to be corrected everywhere it appeared, because handing Vincent or Jeroen a document that claims the isolation already works would be actively misleading and could lead someone to ship this before the filter exists.

**The correction that is now baked into both deliverables, everywhere the topic comes up:**
- The training/archive **flag exists and can be used now**.
- The exclusion **filter does not exist yet** — it is a required, not-yet-built enhancement.
- **Production go-live is blocked** until that filter is built and verified.
- Until then, the feature **runs only in DEV/UAT**.
- Permanent database deletion of practice data is separate, later, and not needed for this MVP — only the guard's completion status must survive.

If continuing this work, **never let any document, email, or statement imply the isolation already works in production.** Always say: flag exists, filter doesn't, production is blocked on the filter, DEV/UAT only for now.

---

## 4b. SECOND STAKEHOLDER CORRECTION — JEROEN'S FEEDBACK (engagement + urgent-incident)

After the first version of both documents was produced, **Jeroen** (engineering reviewer) sent feedback, quoted here verbatim because it drove a full revision:

> "Another very important aspect: engagement and motivation of the guard. The guard has just received a new tool to make life easier, so we must present it as such. The onboarding experience must be fun, like receiving a gift or new toy. 'Congratulations ... I'm your new assistant ... Let's get to know each other ... I can best assist you if you do x ... etc'. It should therefore certainly not be a dry 'now you do this and then you do that' experience. Let's be very engaging."
>
> "A lot of focus in the document is on the 'report urgent incident' flow/scenario, while I don't really think it's so important in practice. When there's a serious incident the guard will just deal with that in a way they've always done so. If it's a new guard and then they're anyway accompanied in the beginning so they don't have to deal with urgent incidents. So while it's not an issue to have this flow, I don't see it as very relevant."
>
> "These are my 2 cents. Happy to discuss or look at the next version."

**What was changed in response (both documents):**

*Engagement:*
- The flow steps were rewritten so SAM introduces itself and speaks warmly, with reference copy quoted inline (e.g. "Hello — I'm SAM, your new assistant. From today I do your paperwork, so you can keep your eyes on the site."), and each step frames a benefit rather than an instruction ("Done. I wrote that up while you talked." / "Fixed. No forms, no rewriting.").
- The brief's summary now leads with "It should feel like unwrapping a good tool, not sitting a test," and **"Feels like a welcome, not a test" is the first row** of the requirements table.
- The detailed doc's script section was renamed **"Tone and script"** and now opens with explicit "HOW IT SHOULD SOUND" principles before the reference copy.
- Engagement was made a real, traceable requirement: **FR-21 → AC-16 → UAT-22** plus a traceability row, with a testable negative clause ("A dry 'do this, then that' sequence does not meet this requirement") and a pass threshold (product read-through + at least four of five test guards calling it welcoming, not a test).

*Urgent-incident de-emphasis:*
- Reduced to exactly one requirement row per document, renamed **"A way out for a real incident (minor)" / FR-17**, explicitly labelled low-priority with Jeroen's own reasoning restated ("In practice guards handle real emergencies as they always have, and new guards are accompanied early on"), plus the instruction "Keep the implementation small."
- Removed from the flow narrative, from the microphone/support-screen requirement, and from the error-recovery and traceability rows where it had been over-weighting the path.
- **Deliberately NOT removed entirely** — Jeroen said "it's not an issue to have this flow," so it retains one FR, one AC, one UAT, and one traceability row. If continuing, do not delete it, and do not let it re-inflate.

**A note on the literal word "Congratulations":** Jeroen's message sketched warmth with an example opening. An early revision transcribed "Congratulations" literally into both the copy and the requirement wording. A reviewer flagged this as a register mismatch (congratulating an adult for being issued mandatory software) and it was reworded — the *requirement* now asks for "a warm, personal greeting that names what it does for the guard and why that helps," leaving the copywriter free to find the line. Treat Jeroen's quoted phrasing as intent, not as locked copy.

**Also fixed in that round:** SAM's spoken line originally promised "two minutes" while the spec's target duration is 3–5 minutes; all time references were harmonised to "a few minutes." And the semantic-equivalence rule (guards need only demonstrate the intended action, not an exact phrase) was moved out of the tone section into FR-07 where QA will find it.

---

## 5. THE MVP PRODUCT FLOW (as specified, in both deliverables — revised twice: once for the permissions/NFC setup sequence, once against real app screenshots)

1. Guard signs in for the first time; SAM opens by itself, before the home screen — new and existing guards alike.
2. SAM introduces itself warmly ("Hello, I'm SAM. From today I do the paperwork with you…"); the screen states the ground rules up front: required, everything is pretend, progress is saved (leave and return resumes at the same step). A step counter shows progress (detailed doc).
3. **Permissions gate:** SAM links to SAM OnSite's App-info page in the phone's settings; the guard opens Permissions, enables everything, returns, and **types** "done" (they cannot speak yet — mic not granted). SAM verifies each permission itself — **including its level** — names anything wrong, and links back. Setup also switches off Android's automatic permission removal for unused apps.
4. **NFC gate:** same pattern — link → enable → return → say/type "done" → SAM confirms NFC (the checkpoint-tag reader) is actually on. (The brief merges steps 3–4 into one step for space; the detailed doc keeps them separate. Same order and behaviour.)
5. Guard presses Talk and asks a practice question ("How do I report an incident?"); the **first successful transcription doubles as the microphone check**. SAM answers. **Nothing live results** (see 5b — a record may exist today but must be suppressed or training-flagged).
6. Guard reports the fictional incident ("A Dell laptop was stolen"). SAM asks for missing details, creates **exactly one** practice report, and **shows it on screen as a card visibly labelled "Training."**
7. With the report in view, guard says "change Dell to MacBook" — same report and message update, no duplicate.
8. Guard long-presses their own message, edits to "A MacBook was stolen at reception," saves. Final report contains **both** "MacBook" and "reception."
9. SAM closes warmly ("You're ready — you can ask, report, and fix reports whenever you need me"); completion is marked only after every step is genuinely observed; a Go-to-home button lands the guard on the normal home screen (Talk / Capture / Type). Refresher can be reopened later without losing completion.

### 5b. REAL APP INTERFACE FACTS (from actual screenshots shared in chat — note: an earlier zip of "onboarding screens" was retracted by the user as mockups; the later individual screenshots are the real app and are authoritative)

- The permissions deep link lands on SAM OnSite's **App-info** page; the guard taps **Permissions** there. Current permission set: **camera, location, microphone, notifications, physical activity** — all "not allowed" on a fresh install.
- **Permissions have levels** ("Allow only while using the app" / "Ask every time" / "Don't allow"; location can be "Allowed all the time"). Verification must check the level, not just on/off — "Ask every time" is not good enough for daily use.
- Android's **"Manage app if unused"** setting auto-removes camera/location/microphone after months of disuse; setup should switch it off (or the app must detect and re-request).
- The app's **Location screen already has** a "Permissions needed" status chip, a "Grant permissions" button, and an "Open system settings" link — so permission-state reading and the settings deep link **partly exist already** (proven for location).
- **Activities have live consequences today**: real transcripts show supervisor alerts being sent, KPI impact, and items created in the **Pronect Action Tracker**. These are named consumers the exclusion filter must cover.
- **A spoken question was recorded as a "Procedure" activity** ("How do I hand out a key?"). So "asking creates no report" is false as an absolute today — the spec now says **nothing live may result**: practice mode either suppresses the record or training-flags it (FR-08).
- An activity is a **report card + transcript** (message-is-the-report confirmed). Conversation UI: mic button, "Listening… speak anytime," pause control, keyboard and camera toggles. Home screen: by-name greeting, **Talk / Capture / Type**, recent activity list.

Key mental model, worth repeating to anyone: **in SAM, a guard's message is what creates and updates their report** — so a correction can happen either by telling SAM (voice) or by editing the original message (manual). Both must work and must converge on the same single report.

---

## 6. THE DELIVERABLES (what exists right now, in the repo)

Location: `deliverables/US1117/` on branch `claude/ticket-deliverable-verify-jxv715`.

| File | What it is |
|---|---|
| `SAM_OnSite_Guard_Onboarding_Spec.docx` | The **brief** — 4 pages, plain language, self-contained (no "see section N" cross-references), no ticket/user-story references anywhere in its content. Everyone-readable. |
| `SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx` | The **detailed companion** — ~10 pages. Full functional requirements (FR-01–20), acceptance criteria (AC-01–15), test matrix (UAT-01–21), and a requirement traceability matrix tying them together. Engineering + QA audience. |
| `build_spec.js` | Node script (uses the `docx` npm package) that generates the brief from scratch. |
| `build_detailed_spec.js` | Node script that generates the detailed companion from scratch. |
| `README.md` | Short internal README describing the pair, the key Vincent correction, and how they were verified. |
| `.gitignore` | Ignores `node_modules/` and `*.log` (the build scripts need `docx` installed locally to run — see §8). |

**Both documents currently state, consistently:** the flag exists, the filter doesn't, production is blocked on the filter, DEV/UAT only until then. Both are ticket-reference-free (no "US1117," no "ticket," no "acceptance criteria [as ticket jargon]," no original-story quoting) per explicit user instruction. Both share one visual design system (see §9).

**Design intent per document:**
- Brief: everyone should be able to read it in one sitting and never have to "connect dots" — every requirement carries its own plain-English proof/confirmation method inline.
- Detailed: the reference document engineering builds from and QA tests from — FR → AC → UAT → traceability, so nothing is unproven or untestable.

---

## 7. THE WORKING METHOD USED (how every revision was actually produced — reuse this exact loop for future work)

This is the process the user explicitly asked to be preserved so it can be repeated without help. It is a **build → verify → fix → re-verify loop**, using independent sub-agents so nothing marks its own homework:

1. **Build.** Write or edit the Node script that generates the `.docx` (see §8 for exact tooling). Never hand-edit a `.docx` binary directly for this project — always regenerate from the script so the source of truth stays in version control as readable code.
2. **Structural audit.** Run a Python script (`audit.py`, built during this session — currently lives only in the ephemeral scratchpad, see §8) against the actual `.docx` XML to check: page size/margins, real Word numbering (no fake Unicode bullets), table geometry, heading counts, and a scan for forbidden ticket-reference terms. This catches mechanical defects fast, before spending agent calls on content review.
3. **Render for visual inspection.** Since this environment has no working Microsoft Word or LibreOffice (see §8 — this was a real, verified dead end, not a shortcut), convert the `.docx` to HTML with the `mammoth` npm package, style it to approximate the real design in CSS, and print to PDF with the pre-installed Chromium via `playwright-core`. Read the resulting PDF to visually check pagination, table rendering, and callouts. This is a **proxy**, not the ground truth — the real `.docx` must eventually be opened in actual Microsoft Word to confirm exact pagination and colour rendering.
4. **Council review.** Spawn multiple independent sub-agents in parallel (via the `Agent` tool), each with a distinct, unbiased lens and **no knowledge of the others' findings** — e.g. clarity/no-gaps, completeness/fidelity, safety/data-integrity, traceability/QA-usefulness, honesty/consistency-with-companion-document, document-design. Each agent reads the actual document content (not a summary) and the authoritative source material, and returns a strict PASS/FAIL verdict with MUST-FIX / SHOULD-FIX findings, each with an exact quote and location.
5. **Apply only evidence-backed fixes.** Read every council report. Apply fixes that are actually correct (verify against source material — don't apply a finding just because an agent said so). Reject or note anything that's a misunderstanding on the agent's part.
6. **Rebuild → re-audit → re-render.** Repeat steps 1–3 after every fix round.
7. **Independent final verifier — separate from the council.** Spawn one more fresh agent, with **no memory of the council agents**, given the strict pass/fail gates explicitly (length, no ticket references, effortless clarity, fidelity to every decision, no unsupported/over-claimed facts, safety, professional design quality — whatever the current bar is). This agent must return an explicit **PASS or FAIL** with a gate-by-gate table. **Do not deliver anything until this verifier returns PASS.** If FAIL, go back to step 1 with its exact findings.
8. **Only after PASS:** commit, push, and deliver the file to the user.

This loop was run **twice** in this chat: once for the original long specification (which surfaced real defects — broken internal cross-references, requirements over-claimed as "Verified" when actually unproven, one orphaned requirement with no test), and once for the redesigned calm-look version of both documents (which passed cleanly). It should be run again for any future revision of these documents, and is a good general pattern for any future Pronect specification work.

---

## 8. THE ENVIRONMENT AND EXACT TOOLING (critical — do not re-attempt what's already ruled out)

This is a fresh Linux container (this is **not** the Windows/Codex machine referenced in the inherited handoff — none of those Windows paths, that Python install, or that Poppler binary exist here). Rebuilding tooling from scratch was necessary and is documented here so it isn't repeated.

**What does NOT work here (verified, don't retry):**
- **LibreOffice (`soffice`)** is installed but its document-conversion pipeline is fundamentally broken in this container — every conversion attempt (docx→pdf, docx→txt, docx→odt, even a trivial one-line docx, even a plain .txt file) fails with `Error: source file could not be loaded`, and running the actual `soffice.bin` binary directly returns exit code 81 (a LibreOffice bootstrap failure), regardless of profile, VCL plugin, or sandbox settings. This is not fixable with flags — it's broken at a lower level. **Do not spend time retrying LibreOffice conversion in this environment.**
- **Microsoft Word** does not exist in this environment at all.
- Because of the above, there is **no way to get a 100%-faithful native render of the final `.docx` in this session.** All visual verification here is a Chromium-based proxy (see below) — genuinely useful for catching structural and layout problems, but the user should open the final files in real Microsoft Word at least once before sending them onward, to confirm exact pagination and color rendering.

**What DOES work, and is the established toolchain — use these:**
- **Node.js v22** is available.
- **`docx` (npm package, v9)** — used to *generate* `.docx` files programmatically from a script. This is how both deliverables are built. Install once per fresh scratchpad: `npm install docx@9` (do not `npm install` inside the repo itself — see below).
- **`mammoth` (npm package)** — converts an existing `.docx` to HTML/Markdown. Used two ways in this project: (a) to extract readable Markdown from the built `.docx` so sub-agents can review actual content instead of a description of it; (b) to convert to HTML for visual proxy rendering.
- **Chromium**, pre-installed at `/opt/pw-browsers/chromium-1194/chrome-linux/chrome`, driven via **`playwright-core`** (npm package) — used to print the mammoth-generated HTML to a PDF for visual inspection. `PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD` is already set; never run `playwright install`.
- **A Python 3 stdlib-only audit script** (`audit.py`) — parses the `.docx` XML directly (unzips it, reads `word/document.xml`, `word/numbering.xml`) to check page geometry, real numbering definitions vs fake Unicode bullets, table structure, and scan body text for forbidden terms (ticket references). No external Python packages needed (avoid `defusedxml`/`python-docx` — not installed and not necessary; plain `zipfile` + `xml.etree.ElementTree` + regex suffice).

**Important practical setup detail:** the `docx` npm package and its `node_modules` were installed in the session's scratchpad directory (a temp path under `/tmp/claude-...`, **not persisted across sessions**), then **symlinked** into the repo's `deliverables/US1117/` folder so the build scripts could `require('docx')` when run from the repo. That symlink is deliberately **not committed** (excluded via `.gitignore`) since `node_modules` shouldn't live in git. **This means: in a fresh chat, before running either build script, you must reinstall `docx` (and `mammoth` + `playwright-core` if you want to re-verify) into some working directory and re-create that symlink, or otherwise ensure `require('docx')` resolves from `deliverables/US1117/`.** The scratchpad path used in this session was under the session-specific temp directory and will not exist in a new session — a new one must be created.

**Reproducible commands (adapt the scratchpad path for the new session):**
```bash
# one-time per fresh session: create a workdir and install deps
mkdir -p /tmp/claude-.../scratchpad/us1117
cd /tmp/claude-.../scratchpad/us1117
npm init -y
npm install docx@9 mammoth playwright-core

# symlink so the repo's build scripts can find docx
ln -sfn /tmp/claude-.../scratchpad/us1117/node_modules /home/user/shahvivan/deliverables/US1117/node_modules

# build both docs
cd /home/user/shahvivan/deliverables/US1117
node build_spec.js
node build_detailed_spec.js

# clean up the symlink before committing (do not commit node_modules)
rm -f /home/user/shahvivan/deliverables/US1117/node_modules
```

The `audit.py` and the mammoth/Chromium render scripts written during this session were saved only in the ephemeral scratchpad and are **not in the repo**. If verification is needed again, they must be rewritten (their logic is fully described in §7 above — page/margin check via `word/pgSz`/`pgMar`, numbering check via `word/numbering.xml` abstractNum count and a regex for literal bullet glyphs at line-start, table `tblGrid`/`tblHeader`/`cantSplit` checks, and a lowercase substring scan of all `<w:t>` text for forbidden ticket terms).

---

## 9. THE VISUAL DESIGN SYSTEM (both documents share this — keep it consistent for any future edits)

The user explicitly rejected an earlier, heavier design (dark near-black table header bars, zebra-striped rows, filled solid-colour status "pills," several competing colours) as "uneasy," "not pleasant," and looking machine-generated. It was redesigned to a **calm, editorial, human-authored** look, confirmed by an independent verifier as achieving that. **Any future edits to these documents should preserve this system, not the old one:**

- **Typography:** Georgia (serif) for section titles and subheadings; Calibri (sans) for body text and table content.
- **Palette (hex, used identically in both `build_spec.js` and `build_detailed_spec.js`):**
  - `INK 3A3A38` — warm charcoal body/heading text (not black).
  - `TEAL 4A6B82` — the single restrained accent colour (dusty blue) — used for section numerals, thin rules, table header text/underline.
  - `TEALD 2E3A44` — deeper blue-grey, used for subheadings inside sections.
  - `GREY 77746E` — warm muted grey for secondary text.
  - `HFILL F4F2EE` / `BOX F7F6F2` — soft warm off-white fills for the cover meta-strip and callout boxes (never a dark fill).
  - `RULE E6E4DF` — hairline colour for thin separators.
  - Status label colours (used as coloured text, never as a filled pill): Confirmed `43724E` (muted green), Proposed `8A6A1F` (muted amber), Check w/ eng. `A05A38` (muted rust).
- **Section headers ("eyebrows"):** a small bold teal numeral ("01", "02"...) followed by a Georgia serif title, with a faint hairline rule beneath. Not heavy borders.
- **Tables:** open style — no cell borders at all except a single accent-coloured rule under the header row, and thin hairline horizontal rules between body rows. No vertical rules, no zebra striping, no dark/filled header background.
- **Status labels:** plain coloured, letter-spaced, small-caps-style text (e.g. "CONFIRMED") — never a filled coloured badge/pill.
- **Callouts:** a soft off-white background box with a thin (not thick) coloured rule on the left edge only.
- **Page setup:** US Letter, margins roughly 0.85"–0.9" (tightened slightly from an initial 1" pass specifically to keep the brief at exactly 4 pages after content additions).
- **Numbering:** real Word `numbering.xml` definitions (`LevelFormat.BULLET` / `LevelFormat.DECIMAL` via the `docx` package) for every list — never literal bullet/dash characters typed into text, and never manually-typed "1." numbering. This was a specific, named defect in the very first version of this document line (inherited from the prior Codex-based attempt) and must never be reintroduced.
- **Header/footer:** running header with document title + subtitle; footer with "Confidential — Pronect internal" and a real `Page X of Y` page-number field.

If asked to adjust the look again, the first place to change values is the palette/constant block near the top of each build script — both files keep this block in the same shape so a change can be mirrored across both documents easily.

---

## 10. CONSTRAINTS THAT MUST NEVER BE VIOLATED (explicit user instructions, collected in one place)

- **No ticket references anywhere in either document's content** — no story ID, no "ticket," no "user story," no iteration number, no quoting the original ticket text. (The word "acceptance criteria" as a plain QA term is fine; it's ticket-system jargon and IDs that are banned.) This was explicitly requested and independently audited for.
- **The brief must stay at 4 pages.** Confirmed via the Chromium-proxy render with visible headroom on page 4; not yet confirmed in native Word.
- **Never claim the practice-data isolation already works in production** — see §4. This is the single most important factual guardrail on this project going forward.
- **Never invent a specific database field name, API endpoint, or app route.** These are explicitly listed as open engineering questions (§11) precisely because they were never confirmed, and the user was firm that nothing should be presented as fact without evidence.
- **Label every claim honestly**: Confirmed (decided fact) / Proposed (a recommended default, not yet agreed) / Check with engineering (an open technical unknown). Never let a Proposed or unverified item read as Confirmed.
- **Nothing goes to production without Vincent Smeyers' explicit sign-off** — stated in both documents' closing lines, and true of the actual working relationship, not just document boilerplate.
- **No production changes have ever been made** by any assistant on this project — everything is read-only investigation (DEV/UAT) plus document generation.
- **Prefer one clean deliverable over unnecessary extra files** — the reason there are two documents (not three, not one) is a deliberate, discussed trade-off: the brief for general readability, the detailed spec specifically because the user decided a QA/engineering companion was worth the exception to "just one file," so that a build/test conversation has complete requirement-to-test traceability.
- **Never take a destructive git action** (force-push, reset --hard, etc.) without explicit permission — none was needed or used in this project; all commits were additive, and each was pushed with `git push -u origin claude/ticket-deliverable-verify-jxv715` after a normal commit.

---

## 11. OPEN QUESTIONS FOR ENGINEERING (unresolved — do not guess these; they are explicitly listed as unknowns in both deliverables and must be answered by Berk/Jeroen, not invented)

1. Does the app already record/know a guard's first successful login anywhere?
2. Can the onboarding exercise be launched directly, on its own, without the app's usual voice-classification or NFC triggers that start other scenarios?
3. Can SAM run in a genuine "practice" conversational mode whose content never enters live/operational reporting?
4. What is the exact technical route the app uses to update an existing report — both when SAM does it via voice correction, and when the guard does a manual long-press edit?
5. Exactly which live consumers (screens, exports, alerts, integrations) must the new exclusion filter cover, and at which exact point in the create/update code path should the training flag be set, so it's atomic with report creation?
6. Is there already a database location to store "this guard has completed onboarding," or must a new table/record be added? (Investigation found no such field visible in the current user-management form — see §2.)

---

## 12. WHAT HAS NOT YET BEEN DONE (the actual outstanding work)

0. **Jeroen has already reviewed a version and sent feedback (see §4b), which has been incorporated.** He closed with "Happy to discuss or look at the next version" — so the natural next step is sending him the revised pair, and the email draft below predates his feedback and should be updated to reference his input before sending.

1. **The email to Vincent Smeyers and Jeroen was drafted earlier in this chat and is NOT sent.** Note it was written *before* Jeroen's feedback arrived, so it needs a line acknowledging his two points were incorporated. The user asked for it to be made "more professional" than an earlier casual draft, and the final professional version (reproduced in full below) is the one to use, unless the user asks for further changes. **Do not assume it was sent — confirm with the user before treating this as done.**

   > **Subject:** SAM OnSite guard onboarding — specification for review
   >
   > Hi Vincent, Jeroen,
   >
   > I've put together the specification for the guard onboarding scenario and would appreciate your review.
   >
   > To prepare it, I went through the requirements and looked at the current behaviour in DEV — the management console (users, templates, and the fact that there's no onboarding-completion field today) and the UAT reporting page, where SAM OnSite Activities can be edited and archived. For how SAM itself handles reports and corrections, I relied on the source documentation and the template pack. From there I made the product decisions and wrote everything up.
   >
   > I've attached two versions:
   >
   > - A short overview (4 pages): what the feature does and how we'll confirm each part works.
   > - A detailed specification: the full requirements, acceptance criteria, test cases, and traceability for build and QA.
   >
   > Vincent, one point I want to flag from your earlier note: the training/archive flag exists, but the filter that excludes that data from the live surfaces isn't built yet. Both documents are explicit that the onboarding stays in DEV/UAT and doesn't go to production until that filter is built and verified, with your sign-off.
   >
   > For transparency: I used AI to help draft and polish the Word documents. The decisions and content are my own.
   >
   > I'd welcome your feedback whenever you have time, and I'm happy to walk through any of it.
   >
   > Best regards,
   > Vivan

2. **Neither `.docx` has been opened in real Microsoft Word yet** — only verified via the Chromium proxy described in §8. Recommended before sending: open both once in Word, confirm the brief is exactly 4 pages and the colours/fonts render as intended, and optionally export a PDF copy that way (the user mentioned to Vincent they'd prepare a PDF for review).
3. **The user is actively learning the technical and product content of these documents themselves**, specifically so they can answer Vincent's and Jeroen's questions without help in a live conversation. A great deal of this chat (not reproduced in full here, but summarized in §13) was spent teaching the underlying concepts from zero, using a repeated analogy: **a "sticker" (the archive/training flag/column, a piece of data on each report) and a "bouncer/filter"** (the not-yet-built check that hides stickered/flagged data from live views) as the throughline for the entire isolation story, layered up to real terms (rows, tables, queries, transactions, webhooks/events, server-side vs client-side, idempotency, concurrency, fail-closed, environments, feature flags). **If continuing this teaching in a new chat, use the same sticker/bouncer/filing-cabinet analogy family — it is what the user has already internalized; switching metaphors would set them back.**
4. **No decision has been made yet on whether/when to actually send the email** — this is the user's call, likely the very next real-world step after this handoff.

---

## 13. THE TEACHING THAT HAPPENED IN THIS CHAT (context for continuity — the user's current knowledge level)

The user asked, escalating in specificity across several turns, to be taught the content of these documents "from zero," because they are not technical and were worried about being asked questions by Vincent or Jeroen that they couldn't answer. This was done in layered passes, each one adding vocabulary on top of concepts already taught, never removing the plain-English base:

1. **Pass 1 — the whole thing as one picture.** "It's a flight simulator for guards" — practice with real stakes removed, and the one universal danger of any simulator: it must never be wired into the real control tower. This is the anchor metaphor for the entire isolation/safety story.
2. **Pass 2 — the seven things that matter**, each in one line: what it is, who/when, what's practised, the safety rule, THE catch (flag exists / filter doesn't), how "done" is decided, the emergency safety valve.
3. **Pass 3 — jargon glossary in plain English** (Activity, archive/soft-delete, "the filter," server-side, DEV/UAT, Confirmed/Proposed/Check-with-engineering) plus a rehearsed Q&A of exactly what to say if Vincent or Jeroen ask specific questions.
4. **Pass 4 (after the user asked specifically for "the catch" explained further) — the sticker-and-bouncer analogy**, introduced and then reused as the fixed metaphor for the rest of the chat: tag = a sticker on a report; filter = a bouncer checking stickers at the door to every live screen; the catch = we can put the sticker on, but nobody's built the bouncer yet.
5. **Pass 5 — the full technical layer**, built as a strict ladder where each new term only uses words already introduced: frontend/backend → database/Supabase/Postgres → table/row/column → query → the tag as a column (the archive flag) → filter as a query condition → view/RLS as "do the filter once, in one shared place, not per-screen" → transaction as "one all-or-nothing step" (why the flag must be set in the same step as report creation) → webhook/event as "an auto-message you can't recall" (why tagging can't happen after creation) → server-side vs client-side (why completion is decided centrally, not trusted from the phone) → completion record + unique constraint + idempotency (why retries/double-taps can't create duplicates) → concurrency (two devices at once) → fail-closed (refuse rather than risk a leak) → environments (DEV/UAT/production) + feature flag (the on/off switch controlling rollout).
6. **Pass 6 (most recent, after the user said even that was too jargon-heavy and asked to start truly "from zero" with no gaps) — the same technical ladder rebuilt with every single term re-explained in full from the ground up**, each with a plain meaning and an everyday comparison, structured so nothing is assumed. This is the most complete and most recent teaching pass and should be treated as the user's actual current baseline of technical understanding if teaching continues.
7. **A "how do I do this myself" pass** — the user asked to be taught the actual *thinking process* for tackling a ticket independently, not just this ticket's content. This was framed around the idea that a non-technical product person's job is like a journalist's (find the right people, ask the right questions, organize the answers — not personally hold the technical answers), broken into six repeatable moves: (1) translate the ticket into one plain sentence + the why, (2) sort everything into "known / my call / their call" buckets, (3) go find answers — ask people for "my call" and "their call" items, but go **look yourself, read-only, in DEV/UAT** for "how does it work today" items, (4) walk the feature as one person's story end-to-end (the happy path), then deliberately break it by asking "what if this fails / what if someone's careless or malicious / what could leak or hurt someone" at every step, (5) turn findings into labelled rules + acceptance checks, being explicit about certainty (Confirmed/Proposed/Check-with-engineering), (6) lay it into the standard document skeleton and then adversarially re-read your own draft looking for holes — this is literally the council-and-verifier loop from §7, just done solo. A concrete "first 20 minutes with any new ticket" checklist was given: minutes 1–5 write the one-sentence translation, minutes 5–15 build a two-column list of "questions for product (me)" vs "questions for engineering," minutes 15–20 go open the relevant part of the system in DEV/UAT and look around.

**If a new chat is asked to continue teaching:** don't restart from scratch or re-explain concepts already covered above — check where the user's understanding currently sits (ask them to explain something back, as was being done at the moment this handoff was requested) and build forward from there, keeping the sticker/bouncer/filing-cabinet analogy family intact.

---

## 14. THE STANDARD SPEC SKELETON (reusable for any future Pronect specification — taught explicitly to the user as a general-purpose template)

This shape was distilled during the teaching passes as the reusable structure behind any well-formed specification document, and both `US1117` deliverables follow it:

> Purpose & confirmed decisions → Scope (in/out) → Roles & ownership → End-to-end flow → Voice/screen script (if applicable) → Functional requirements (each labelled Confirmed/Proposed/Check-with-engineering) → Data/progress model → Safety & integrity guarantees (the "two things that must not go wrong" for this particular feature) → Error & recovery behaviour → Acceptance criteria (one per requirement, Given/When/Then) → Test matrix, including negative/break-it tests → Requirement traceability matrix (ties requirement → criterion → test → evidence) → Rollout & rollback plan → Definition of done → Open engineering questions (explicitly never guessed).

Use this skeleton as the starting outline for any new Pronect ticket that needs a specification written.

---

## 15. EXACT NEXT-ACTION CHECKLIST FOR A NEW CHAT

1. Read this file fully before doing anything else.
2. If asked to modify either `.docx`: re-create the Node toolchain per §8 (fresh scratchpad, `npm install docx@9 mammoth playwright-core`, symlink into `deliverables/US1117/`), edit the relevant `build_*.js`, rebuild, re-verify visually (mammoth + Chromium proxy) and structurally (rewrite the audit logic per §7/§8 if the script itself isn't available), then run the full council-and-independent-verifier loop from §7 before treating any change as final. Never skip straight to "looks fine" without that loop — that discipline is the whole reason these documents are trustworthy.
3. Preserve the design system in §9 exactly unless the user explicitly asks for a different look.
4. Never violate any constraint in §10, especially: no ticket references, never claim the isolation filter already works, never invent unconfirmed technical facts, always label certainty honestly.
5. If asked about the email: it has been drafted (§12) but not sent — confirm with the user rather than assuming either way.
6. If asked to continue teaching the user: build forward from §13's pass 6/pass 7, using the sticker/bouncer analogy family, and check current understanding by asking the user to explain a concept back before adding more.
7. Before delivering any new/changed file to the user, remove any local `node_modules` symlink from the repo folder and confirm nothing extraneous is staged, then commit with a clear message and push to `claude/ticket-deliverable-verify-jxv715` with `git push -u origin claude/ticket-deliverable-verify-jxv715`.
8. Never create a pull request unless the user explicitly asks for one.

---

## 16. QUICK-REFERENCE SUMMARY (if only 60 seconds are available)

SAM OnSite needs a mandatory, one-time voice-guided onboarding for every guard (new and existing) that teaches asking a question, reporting a fictional incident, and correcting it by voice and by manual edit. Two Word documents specify this (brief + detailed, in `deliverables/US1117/`, built from Node scripts using the `docx` package, styled in a calm editorial design, ticket-reference-free, verified via a council-of-agents + independent-verifier loop). The one fact that overrides everything else: **the archive/training flag exists, but the filter that hides that data from production doesn't — so this stays in DEV/UAT and cannot ship until that filter is built and Vincent signs off.** An email to Vincent and Jeroen is drafted but unsent. The user is learning the material themselves and should be taught forward from where they left off, not re-taught from scratch.
