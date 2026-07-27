/*
 * SAM OnSite — Guard Onboarding | Detailed specification (engineering + QA companion)
 * Node + docx.  Run: node build_detailed_spec.js
 * Out: SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx
 *
 * Same visual identity as the 4-page brief (Georgia headings, slate+teal).
 * Ticket-free. References use IDs (FR-/AC-/UAT-) and section NAMES, never
 * "section N" numbers, so there are no breakable pointers. Corrected for the
 * product owner's note that the exclusion filter is not yet built (data is
 * currently visible in production), so this must stay in DEV/UAT until it is.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  Header, Footer, PageNumber, LevelFormat, TabStopType, Tab, VerticalAlign,
} = require("docx");

/* ---------- palette (calm editorial — matches the brief) ---------- */
const INK = "3A3A38", TEAL = "4A6B82", TEALD = "2E3A44", GREY = "77746E";
const HFILL = "F4F2EE", ZEBRA = "F3F6F5", BOX = "F7F6F2", RULE = "E6E4DF";
const SERIF = "Georgia", SANS = "Calibri";
const PAGE_W = 12240, PAGE_H = 15840, MARGIN_TB = 1224, MARGIN_LR = 1296;
const CONTENT_W = PAGE_W - 2 * MARGIN_LR; // 9648

const STATUS = {
  C: { label: "Confirmed",     color: "43724E" },
  P: { label: "Proposed",      color: "8A6A1F" },
  X: { label: "Check w/ eng.", color: "A05A38" },
};

/* ---------- helpers ---------- */
function t(text, o = {}) {
  return new TextRun({ text, font: o.font || SANS, size: o.size || 19, color: o.color || INK,
    bold: !!o.bold, italics: !!o.it, characterSpacing: o.cs });
}
function chip(key) {
  // quiet coloured label (no filled pill)
  const s = STATUS[key];
  return new TextRun({ text: s.label.toUpperCase(), font: SANS, size: 14, bold: true, color: s.color, characterSpacing: 6 });
}
function p(children, o = {}) {
  return new Paragraph({ spacing: { after: o.after == null ? 92 : o.after, line: o.line || 252, lineRule: "auto" },
    alignment: o.align, keepNext: o.keepNext, children: Array.isArray(children) ? children : [t(children, o)] });
}
function eyebrow(num, title) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 220, after: 78 }, keepNext: true,
    border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 4, space: 6 } },
    children: [ new TextRun({ text: num + "   ", font: SANS, size: 19, bold: true, color: TEAL, characterSpacing: 10 }),
      new TextRun({ text: title, font: SERIF, size: 25, color: INK }) ] });
}
function subhead(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 150, after: 56 }, keepNext: true,
    children: [new TextRun({ text, font: SERIF, size: 20, bold: true, color: TEALD })] });
}
const HAIR = { style: BorderStyle.SINGLE, size: 2, color: RULE };
const NONE = { style: BorderStyle.NONE };
const HEADRULE = { style: BorderStyle.SINGLE, size: 8, color: TEAL };
function cell(content, o = {}) {
  const paras = Array.isArray(content) && content[0] instanceof Paragraph ? content
    : [new Paragraph({ spacing: { after: 0, line: 248, lineRule: "auto" }, alignment: o.align,
        children: Array.isArray(content) ? content : [t(content, { size: o.size || 18, bold: o.bold, color: o.color })] })];
  return new TableCell({ width: { size: o.w, type: WidthType.DXA }, verticalAlign: o.va || VerticalAlign.TOP,
    margins: { top: 56, bottom: 56, left: 40, right: 150 }, borders: { top: NONE, bottom: NONE, left: NONE, right: NONE }, children: paras });
}
function hcell(text, w) {
  return new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM,
    margins: { top: 30, bottom: 52, left: 40, right: 150 }, borders: { top: NONE, bottom: HEADRULE, left: NONE, right: NONE },
    children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: text.toUpperCase(), font: SANS, size: 15, bold: true, color: TEAL, characterSpacing: 8 })] })] });
}
function table(colW, headers, rows) {
  const head = new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => hcell(h, colW[i])) });
  const body = rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, ci) => {
    if (c instanceof TableCell) return c;
    if (Array.isArray(c) && (c[0] instanceof TextRun || c[0] instanceof Paragraph)) return cell(c, { w: colW[ci] });
    return cell(String(c == null ? "" : c), { w: colW[ci] });
  }) }));
  return new Table({ columnWidths: colW, width: { size: colW.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: HAIR, insideVertical: NONE }, rows: [head, ...body] });
}
function callout(lines, accent = TEAL, title) {
  const inner = [];
  if (title) inner.push(new Paragraph({ spacing: { after: 46 }, children: [t(title, { bold: true, color: accent })] }));
  lines.forEach((ln, i) => inner.push(new Paragraph({ spacing: { after: i === lines.length - 1 ? 0 : 60, line: 258, lineRule: "auto" },
    children: Array.isArray(ln) ? ln : [t(ln)] })));
  return new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
    borders: { top: NONE, bottom: NONE, right: NONE, left: { style: BorderStyle.SINGLE, size: 12, color: accent }, insideHorizontal: NONE, insideVertical: NONE },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX }, margins: { top: 96, bottom: 96, left: 170, right: 170 }, children: inner })] })] });
}
function bullet(ref, children) { return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 48, line: 250, lineRule: "auto" }, children: Array.isArray(children) ? children : [t(children)] }); }
function numitem(ref, children) { return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 42, line: 248, lineRule: "auto" }, children: Array.isArray(children) ? children : [t(children)] }); }
const gap = (h) => new Paragraph({ spacing: { after: h }, children: [] });
// detail cell = leading quiet label + text
const dcell = (key, text, w, fill) => cell([chip(key), t("  " + text, { size: 18 })], { w });

/* ================= DATA ================= */
const FR = [
  ["FR-01", "C", "Trigger", "Onboarding launches automatically right after a guard’s first successful sign-in, and can be reopened later as a refresher."],
  ["FR-02", "C", "Audience", "Every active guard receives it exactly once — including guards whose accounts existed before this feature."],
  ["FR-03", "C", "Mandatory", "There is no path to the home screen except by completing it (the urgent-incident action is a detour, not a finish)."],
  ["FR-04", "C", "Refresher", "A completed guard can reopen the training; reopening never clears or regresses completion."],
  ["FR-05", "C", "Language", "All onboarding voice and screen copy is English for the MVP."],
  ["FR-06", "C", "Microphone", "The guard must enable the microphone; declining re-prompts with guidance. If it truly cannot be enabled (broken hardware or device policy), the guard reaches a support screen that still shows the urgent-incident action — never trapped."],
  ["FR-07", "P", "Transcription check", "A speaking step is credited only after at least one successful voice transcription; an empty or failed transcription does not advance the step."],
  ["FR-08", "P", "Ask a question", "The guard asks SAM a spoken question and SAM answers. This step runs in a practice context and creates no report."],
  ["FR-09", "P", "One practice report", "Reporting the fictional incident creates exactly one practice report. Retries or repeated messages must not create additional reports."],
  ["FR-10", "C", "Voice correction", "A spoken correction (Dell → MacBook) updates the same report in place, with no duplicate. (The exact way the app updates the report is an engineering confirmation — see Confirm with engineering.)"],
  ["FR-11", "C", "Manual correction", "Long-pressing the guard’s own message allows a manual edit that, when saved, updates the same report; the final report contains both “MacBook” and “reception.” (Whether the app already supports this edit is an engineering confirmation.)"],
  ["FR-12", "C", "Marked at creation", "The training / archive flag is set atomically when the report is created — so a report never exists un-marked. If it cannot be created already marked, it is not created at all."],
  ["FR-13", "C", "Kept out of live surfaces", "The practice report must appear in no live surface (see the Data isolation callout). The platform has the training flag, but the FILTER that excludes marked data is not built yet — data is currently visible in production — so this exclusion filter is a required, not-yet-built enhancement; production is blocked until it is built and verified, and the exercise runs only in DEV/UAT until then."],
  ["FR-14", "P", "Completion integrity", "Completion is recorded server-side and versioned (version 1), issued once per guard, and set only after each required action is observed — never a quiz or an “I’m done” button."],
  ["FR-15", "P", "Fail-closed", "If safe isolation of practice data cannot be assured at any step, the step fails and completion is not marked; the guard stays in the exercise."],
  ["FR-16", "C", "Hand-off", "On confirmed completion the guard reaches the normal home screen (Talk / Capture / Type)."],
  ["FR-17", "P", "Urgent-incident path", "A safety measure we added (not an original requirement): an always-visible “Report an urgent incident” action, usable without a microphone, opens normal reporting and returns the guard to the saved stage. It does not count as finishing or skipping."],
  ["FR-18", "P", "Idempotency / concurrency", "One onboarding record and one active practice-report reference per guard. Retries do not duplicate; the newest valid server-side progress wins with no regression; the completion operation is idempotent."],
  ["FR-19", "P", "Privacy / telemetry", "Telemetry records only stage and error codes and completion metrics. It does not keep microphone-permission state beyond need, does not log raw voice/transcript content beyond existing policy, and never routes fictional incident content into operational analytics."],
  ["FR-20", "C", "Deferred visual support", "Extra visual help (tooltips, highlight animations) for dark or noisy sites is deferred; the need is judged during testing, not built for the MVP. (Scope decision — no separate test.)"],
];
const AC = [
  ["AC-01", "FR-01", "Given an active guard who has not completed onboarding, when they sign in, then it launches automatically before the home screen is reachable."],
  ["AC-02", "FR-02", "Given an existing guard (account predating this feature) who has not completed it, when they sign in, then they receive it exactly once; after completion a later sign-in goes straight to home."],
  ["AC-03", "FR-03", "Given onboarding in progress, when the guard tries to skip, then there is no path to home except completion (or the urgent-incident detour)."],
  ["AC-04", "FR-06", "Given microphone permission denied, when the guard tries to proceed, then SAM re-prompts and speaking steps do not advance; if the mic truly cannot be enabled, the guard reaches a support screen that still shows the urgent-incident action."],
  ["AC-05", "FR-07", "Given a speaking step, when transcription fails or returns empty, then the step is not credited and the guard is asked to try again."],
  ["AC-06", "FR-08 / FR-09", "Given the guard asks a question, then SAM answers and no report is created; given the guard then reports the incident, then exactly one practice report exists (retries add none)."],
  ["AC-07", "FR-10", "Given a practice report naming a Dell, when the guard says to change it to a MacBook, then the same report updates and no second report is created."],
  ["AC-08", "FR-11", "Given the guard long-presses their own message and saves “A MacBook was stolen at reception,” then the final report contains both “MacBook” and “reception.”"],
  ["AC-09", "FR-12 / FR-13", "Given a practice report at any stage, then it is flagged training at creation; once the exclusion filter is built it appears in no live surface; until the filter exists the exercise runs only in DEV/UAT and does not go to production."],
  ["AC-10", "FR-14 / FR-15", "Given every required action observed and isolation assured, when completion runs, then version 1 is recorded server-side; if isolation cannot be assured, completion is withheld."],
  ["AC-11", "FR-04 / FR-16", "Given a completed guard, when they finish, then they reach home; and reopening the training later preserves completion."],
  ["AC-12", "FR-18", "Given duplicate concurrent sessions or retried messages, then no duplicate record or report is created and progress does not regress."],
  ["AC-13", "FR-17", "Given a real urgent incident during onboarding, when the guard uses the urgent action, then normal reporting opens and, on return, the saved stage resumes with completion neither granted nor skipped."],
  ["AC-14", "FR-05", "Given any onboarding screen or prompt, then all voice and screen copy is in English."],
  ["AC-15", "FR-19", "Given a run, when telemetry and analytics are inspected, then only stage/error codes and completion metrics are recorded, with no raw voice/transcript content and no fictional incident content in operational analytics."],
];
const UAT = [
  ["UAT-01", "AC-01", "New guard’s first sign-in triggers onboarding.", "Shown before home; progress record created."],
  ["UAT-02", "AC-02", "Existing (pre-feature) guard is issued it once.", "Runs once; next sign-in after completion goes to home."],
  ["UAT-03", "AC-03", "Attempt to skip onboarding. (negative)", "No skip path to home; only completion or urgent detour."],
  ["UAT-04", "AC-04", "Deny microphone, then try to continue. (negative)", "Re-prompt with guidance; speaking steps blocked."],
  ["UAT-05", "AC-04", "Microphone truly unavailable (device policy). (negative)", "Support screen reached; urgent action present; not trapped."],
  ["UAT-06", "AC-05", "Force a failed / empty transcription. (negative)", "Step not credited; retry prompted."],
  ["UAT-07", "AC-06", "Ask SAM a question.", "SAM answers; no report created."],
  ["UAT-08", "AC-06", "Report the fictional incident.", "Exactly one practice report exists."],
  ["UAT-09", "AC-07", "Voice-correct Dell to MacBook.", "Same report updated; no duplicate."],
  ["UAT-10", "AC-07", "Voice correction misheard, then retry. (negative)", "SAM re-asks; still exactly one report."],
  ["UAT-11", "AC-08", "Manual long-press edit and save.", "Final report contains “MacBook” and “reception.”"],
  ["UAT-12", "AC-08", "Manual save fails, then retry. (negative)", "Error shown; typed text retained; report not half-edited."],
  ["UAT-13", "AC-09", "After the exclusion filter is built, inspect every live surface.", "Absent from all listed consumers; before the filter, run confined to DEV/UAT."],
  ["UAT-14", "AC-09 / AC-10", "Force create-with-flag / isolation failure. (negative)", "Report not created / completion withheld (fail-closed)."],
  ["UAT-15", "AC-10 / AC-11", "Complete all actions with isolation assured.", "Version 1 recorded server-side; guard reaches home."],
  ["UAT-16", "AC-12", "Duplicate concurrent sessions / retried messages. (negative)", "One record, one report; no progress regression."],
  ["UAT-17", "AC-12", "Network drop mid-step, resume on another device. (negative)", "Progress preserved; resumes at saved stage; no loss."],
  ["UAT-18", "AC-11", "Reopen training after completion (refresher).", "Reopens; completion preserved, not cleared."],
  ["UAT-19", "AC-13", "Trigger the urgent-incident path mid-onboarding.", "Normal reporting opens; returns to saved stage; completion neither granted nor skipped."],
  ["UAT-20", "AC-14", "Review all onboarding copy.", "All copy is English."],
  ["UAT-21", "AC-15", "Inspect telemetry and analytics during a run. (negative)", "Only stage/error + completion metrics; no raw content; no fictional data in analytics."],
];
const TRACE = [
  ["Auto-launch + refresher", "FR-01", "AC-01", "UAT-01", "Shown before home; refresher reopens"],
  ["Every guard once, incl. existing", "FR-02", "AC-02", "UAT-02", "Issued once per guard"],
  ["Mandatory", "FR-03", "AC-03", "UAT-03", "No skip path"],
  ["Refresher keeps completion", "FR-04", "AC-11", "UAT-18", "Completion preserved on reopen"],
  ["English MVP", "FR-05", "AC-14", "UAT-20", "All copy English"],
  ["Mic denied never a dead end", "FR-06", "AC-04", "UAT-04/05", "Re-prompt; support screen keeps urgent action"],
  ["Transcription must succeed", "FR-07", "AC-05", "UAT-06", "Failed transcription not credited"],
  ["Ask a question (SAM answers, no report)", "FR-08", "AC-06", "UAT-07", "Answer given; no report"],
  ["One practice report", "FR-09", "AC-06", "UAT-08", "Exactly one report"],
  ["Voice correction, same report", "FR-10", "AC-07", "UAT-09/10", "Same report; no duplicate"],
  ["Manual edit, final MacBook + reception", "FR-11", "AC-08", "UAT-11/12", "Final report has both words"],
  ["Marked training at creation", "FR-12", "AC-09", "UAT-14", "Never un-marked; fail-closed at create"],
  ["Kept out of live surfaces (filter unbuilt)", "FR-13", "AC-09", "UAT-13", "Absent once filter built; DEV/UAT until then"],
  ["Completion observed, server-side, versioned", "FR-14", "AC-10", "UAT-15", "Version 1 server-side"],
  ["Fail-closed isolation", "FR-15", "AC-10", "UAT-14", "Completion withheld on failure"],
  ["Reaches home", "FR-16", "AC-11", "UAT-15", "Home after completion"],
  ["Urgent-incident safety path (added)", "FR-17", "AC-13", "UAT-19", "Reporting opens; stage resumes"],
  ["Idempotency / concurrency", "FR-18", "AC-12", "UAT-16/17", "No duplicates; no regression"],
  ["Privacy / telemetry", "FR-19", "AC-15", "UAT-21", "No raw content; no fictional data in analytics"],
  ["Deferred visual support", "FR-20", "—", "—", "Scope decision; evaluated in testing"],
];

/* ================= CONTENT ================= */
const kids = [];

/* cover */
kids.push(new Paragraph({ spacing: { after: 16 }, children: [t("PRONECT   ·   SAM ONSITE", { size: 16, bold: true, color: TEAL, cs: 40 })] }));
kids.push(new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: "Guard Onboarding", font: SERIF, size: 46, bold: true, color: INK })] }));
kids.push(new Paragraph({ spacing: { after: 140 }, border: { bottom: { color: TEAL, style: BorderStyle.SINGLE, size: 4, space: 10 } },
  children: [new TextRun({ text: "Detailed specification — engineering & QA companion", font: SERIF, size: 21, italics: true, color: GREY })] }));
kids.push(new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [new TableRow({ children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX }, margins: { top: 100, bottom: 100, left: 180, right: 180 }, borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
    children: [new Paragraph({ spacing: { after: 0 }, children: [
      t("Product  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("SAM OnSite       ", { size: 18 }),
      t("Applies to  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Every guard, once       ", { size: 18 }),
      t("Language  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("English       ", { size: 18 }),
      t("Status  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Draft for build       ", { size: 18 }),
      t("Sign-off  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Vincent Smeyers", { size: 18 }),
    ] })] })] })] }));
kids.push(gap(60));
kids.push(callout([
  [t("This companion", { bold: true, color: TEALD }), t(" carries the full functional requirements, acceptance criteria, test matrix, and traceability behind the one-page brief. It reads on its own. Every requirement carries a plain-language status — ")],
  [chip("C"), t("  agreed behaviour (some note an implementation detail to confirm).   "), chip("P"), t("  our recommended default.   "), chip("X"), t("  needs an engineering confirmation.")],
], TEAL));
kids.push(gap(40));

/* 01 Purpose & decisions */
kids.push(eyebrow("01", "Purpose and confirmed decisions"));
kids.push(p([
  t("SAM OnSite is Pronect’s voice-first guard assistant: a guard speaks naturally and SAM captures structured, time-stamped reports. This specifies a "),
  t("mandatory, one-time voice onboarding", { bold: true }),
  t(" that runs at a guard’s first sign-in and teaches the three everyday actions — ask a question, report an incident, correct a report — using a made-up incident so nothing real is affected."),
]));
const decRef = "dec";
[
  "Every guard receives it once, including existing guards; it is mandatory and cannot be skipped.",
  "“Generate a report” means reporting a fictional incident by talking to SAM. Applies to all guards; configuration is universal, not customer-specific.",
  "A completed guard can reopen it later as a refresher without losing completion. English for the MVP.",
  "“Editing logs” means editing the guard’s own report — by voice, and by long-pressing their message and saving. The final report must reflect the correction.",
  "Completion is inferred from the system observing the guard perform each action — not a quiz or a self-confirmation button.",
  "Enhanced visual support (tooltips, highlights) may be evaluated later; it is not required for the MVP.",
].forEach((s) => kids.push(bullet(decRef, s)));

/* 02 Scope */
kids.push(eyebrow("02", "Scope"));
kids.push(subhead("IN SCOPE"));
const inRef = "insc";
[
  "A universal, application-level onboarding that launches at first sign-in and is reopenable as a refresher.",
  "One guided pass through: ask a question, report a fictional incident, correct it by voice and by manual edit.",
  "Server-side, versioned completion tracking that issues version 1 exactly once to every active guard.",
  "Isolation of all fictional practice data from every live surface (dependent on the exclusion filter described later).",
].forEach((s) => kids.push(bullet(inRef, s)));
kids.push(subhead("OUT OF SCOPE"));
const outRef = "outsc";
[
  "Languages other than English; tooltips / highlight animations (deferred for evaluation in testing).",
  "Customer- or site-specific onboarding variants; any change to a live template or protected system template.",
  "Physical deletion / database clean-up of practice data — a separate, later task from the logical archive used here.",
].forEach((s) => kids.push(bullet(outRef, s)));

/* 03 Roles */
kids.push(eyebrow("03", "Roles and ownership"));
kids.push(p("Ownership so no requirement is orphaned across teams. A coordination aid, not a contractual RACI."));
kids.push(table([1900, 2600, 5148], ["Layer", "Owner (indicative)", "Responsibility"], [
  ["Mobile app", "SAM OnSite", "First-sign-in detection, launch/resume, microphone loop, long-press manual edit, home hand-off."],
  ["SAM / AI", "AI", "Practice (non-reporting) conversation; structured-essential capture and correction so a fix updates the same report."],
  ["Backend / data", "Backend / Supabase", "Progress record; training flag set at creation; the exclusion filter; idempotency, concurrency, completion."],
  ["Reporting / integrations", "Backend + Reporting", "Build and verify the filter that excludes training-marked data across every live consumer."],
  ["Product / approval", "Product; Vincent Smeyers", "Decisions and acceptance; Vincent signs off any production step."],
  ["QA", "QA", "Run the test matrix, including the negative and concurrency tests, in DEV/UAT."],
]));

/* 04 Flow */
kids.push(eyebrow("04", "End-to-end flow"));
kids.push(p([t("Good to know: in SAM, a guard’s message is what creates and updates their report — so a report can be fixed by telling SAM or by editing the original message. Target duration is roughly 3–5 minutes ("), chip("P"), t(" a product estimate, not a measured fact; completion is gated on observed actions, not time).")], { after: 60 }));
const flowRef = "flow";
[
  "The guard signs in; if onboarding is not complete, it opens automatically before the home screen (new and existing guards).",
  "A welcome screen explains it is required, uses made-up information, is short, and saves progress.",
  "The guard enables the microphone. If declined, SAM re-prompts; if it truly cannot be enabled, the guard reaches a support screen that still shows the urgent-incident action.",
  "The guard asks SAM a practice question and SAM answers — no report is created.",
  "The guard reports “A Dell laptop was stolen.” SAM asks for anything missing and creates one practice report.",
  "The guard says to change “Dell” to “MacBook”; the same report and message update; no second report appears.",
  "The guard long-presses their message, edits it to “A MacBook was stolen at reception,” and saves; the final report shows both “MacBook” and “reception.”",
  "Completion is marked only after each step is observed for real; the guard reaches the normal home screen. A completed guard may reopen the training later without losing completion.",
].forEach((s) => kids.push(numitem(flowRef, s)));

/* 05 Script */
kids.push(eyebrow("05", "Voice and screen script"));
kids.push(p([chip("P"), t("  Reference copy. Where completion is judged by final system state, the guard need only demonstrate the intended action — semantically-equivalent English is accepted.")], { after: 50 }));
kids.push(table([1500, 4074, 4074], ["Stage", "SAM says", "Guard does"], [
  ["Welcome", "“Welcome to SAM. This quick practice is required and uses made-up information. Your progress is saved.”", "Reads; taps Start."],
  ["Microphone", "“I need your microphone to hear you. Please enable microphone access to continue.”", "Grants access (re-prompted until granted)."],
  ["Ask a question", "“Try asking me something — for example, ‘How do I report an incident?’”", "Asks; SAM answers. No report created."],
  ["Report", "“Now report a practice incident — for example, ‘A Dell laptop was stolen.’ I may ask for a few details.”", "Reports; one practice report is created."],
  ["Voice fix", "“Mistakes happen. Tell me to change Dell to MacBook.”", "“Change the Dell to a MacBook.” Same report updates."],
  ["Manual fix", "“You can also fix it yourself — long-press your message, edit it, and save.”", "Edits message to “…MacBook…at reception.” and saves."],
  ["Finish", "“Great work — here’s your home screen.”", "Continues to home once completion is confirmed."],
]));

/* 06 Functional requirements */
kids.push(eyebrow("06", "Functional requirements"));
kids.push(p("Each requirement is traced to acceptance criteria and tests in the Traceability matrix.", { after: 50 }));
kids.push(table([850, 2100, 6698], ["ID", "Requirement", "Detail"],
  FR.map((r) => [r[0], [t(r[2], { size: 18, bold: true })], [chip(r[1]), t("  " + r[3], { size: 18 })]])));

/* 07 Progress model */
kids.push(eyebrow("07", "Progress and completion model"));
kids.push(p([chip("P"), t("  A recommended record shape, not observed schema — field and key names to be confirmed in build. No completion field was visible in the supplied user form.")], { after: 50 }));
kids.push(table([2600, 2400, 4648], ["Field", "Example", "Purpose"], [
  ["user_id", "—", "The guard the record belongs to."],
  ["onboarding_key", "sam_onsite_guard", "Identifies this onboarding; leaves room for future ones."],
  ["version", "1", "Version issued; each guard gets version 1 once."],
  ["status", "not_started / in_progress / completed", "Server-side state; never trusted from the phone alone."],
  ["current_stage", "e.g. voice_fix", "Resume point after app close or device change."],
  ["started_at / updated_at / completed_at", "timestamps", "Lifecycle timing and idempotency support."],
]));

/* 08 Data isolation */
kids.push(eyebrow("08", "Practice data isolation and completion integrity"));
kids.push(subhead("PRACTICE DATA MUST NEVER REACH LIVE OPERATIONS"));
kids.push(p([
  t("Every practice report is flagged training (archive / soft-delete) the instant it is created — the platform already has this flag. "),
  t("Key dependency:", { bold: true }),
  t(" the FILTER that keeps marked data out of live surfaces is not built yet — today, archived items are still visible in production — so it is a required enhancement. This onboarding therefore runs only in DEV/UAT until that exclusion filter is built and verified, and must not go live before then. When built, it should work hide-by-default at the shared data layer, so every current and future surface excludes training data automatically. Permanent deletion is a separate later clean-up; only the guard’s completion status is kept."),
]));
kids.push(callout([
  "Reports and event pages · dashboards and KPIs · analytics and metrics · exports and scheduled reports · alerts and notifications · webhooks and integrations · the event outbox and any queue feeding the above · read replicas, caches, and search indexes · the SAM conversation / transcript store · the guard’s own activity list and search. A dispatched alert or webhook cannot be recalled, so nothing may be emitted before the training flag is in place.",
], TEAL, "Every consumer the exclusion filter must cover"));
kids.push(subhead("COMPLETION IS SERVER-SIDE, EARNED, AND SAFE UNDER PRESSURE"));
kids.push(p("Completion is recorded only after the system observes each real action; it lives on the server (never trusted from the phone), is versioned (version 1), issued once per guard, resumes safely across app close or device change, and cannot be duplicated by two concurrent sessions. If practice data cannot be safely isolated at any point, completion is withheld and the guard stays in the exercise — protecting live data outranks finishing.", { after: 40 }));

/* 09 Errors */
kids.push(eyebrow("09", "Error and recovery behaviour"));
kids.push(table([3000, 6648], ["Condition", "Behaviour"], [
  ["Microphone declined or unavailable", "Keep prompting with clear guidance; a truly blocked mic reaches a support screen that still shows the urgent-incident action. Speaking steps do not advance until the mic is on."],
  ["Network, speech, or save failure", "Preserve progress; show a retryable error; never silently lose the current step or an edit."],
  ["Voice correction misheard", "SAM re-asks; the guard retries by voice or uses the manual long-press edit."],
  ["Manual save failure", "Show the failure and keep the typed text for retry; the report is never left half-edited."],
  ["Isolation cannot be assured", "Fail closed: withhold completion, keep the guard in the exercise, record only a stage/error code — never the practice content."],
  ["Real urgent incident", "The urgent-incident action opens normal reporting and returns to the saved stage; completion is neither granted nor skipped."],
]));

/* 10 Acceptance criteria */
kids.push(eyebrow("10", "Acceptance criteria"));
kids.push(table([850, 1350, 7448], ["ID", "Traces to", "Criterion"], AC.map((r) => [r[0], r[1], r[2]])));

/* 11 UAT */
kids.push(eyebrow("11", "Test matrix (DEV / UAT)"));
kids.push(p("Positive and negative tests. Negatives cover skip, mic denial, mic unavailable, transcription failure, voice-correction failure, manual-save failure, isolation failure, concurrency, and network/device resume.", { after: 50 }));
kids.push(table([850, 1050, 3900, 3848], ["ID", "Verifies", "Test", "Expected evidence"], UAT.map((r) => [r[0], r[1], r[2], r[3]])));

/* 12 Traceability */
kids.push(eyebrow("12", "Requirement traceability"));
kids.push(p("Each source requirement tied to its functional requirement, acceptance criterion, test, and the evidence QA captures.", { after: 50 }));
kids.push(table([2600, 1150, 1150, 1350, 3398], ["Requirement", "FR", "AC", "Test", "Evidence QA captures"], TRACE.map((r) => r)));

/* 13 Rollout */
kids.push(eyebrow("13", "Rollout and rollback"));
const rollRef = "roll";
[
  [t("Production go-live is "), t("blocked", { bold: true }), t(" until the exclusion filter is built and verified across every consumer above; until then it runs only in DEV/UAT with a dedicated test customer account (tenant) and test guards.")],
  "Release behind a switch so it can be enabled gradually and turned off instantly, without erasing anyone’s completed status.",
  "Universal for all guards — no customer-specific setup — and it changes no existing templates.",
  "Monitor completion-failure and isolation-failure rates; alert on isolation failures. Do not log raw voice/report content in monitoring beyond existing policy.",
  "Nothing reaches production without Vincent Smeyers’ sign-off.",
].forEach((s) => kids.push(bullet(rollRef, s)));

/* 14 DoD */
kids.push(eyebrow("14", "Definition of done"));
const dodRef = "dod";
[
  "All acceptance criteria pass in UAT, including every negative test.",
  "Version 1 issued once to a sample of new and existing guards; the refresher preserves completion.",
  "The exclusion filter is built and the practice report is proven absent from every consumer above; a forced isolation failure withholds completion.",
  "Completion is server-side, versioned, idempotent, and safe under concurrency and device change.",
  "The Confirm-with-engineering items are answered, and Vincent Smeyers has signed off.",
].forEach((s) => kids.push(bullet(dodRef, s)));

/* 15 Engineering validation */
kids.push(eyebrow("15", "Confirm with engineering before building"));
kids.push(p([chip("X"), t("  These depend on platform behaviour not provable from the supplied sources. Settle them once, up front; none is expected to be a blocker.")], { after: 50 }));
const valRef = "val";
[
  "Does the app already record a guard’s first successful sign-in?",
  "Can this exercise be launched on its own, without the usual voice or tap triggers that start other scenarios?",
  "Can SAM run in a practice mode whose conversation and report never enter live reporting?",
  "What is the exact way the app updates an existing report — by voice, and by a saved manual edit?",
  "Which live consumers must the new exclusion filter cover, and where are reports created so the training flag is set at creation? (The flag exists; the filter does not yet.)",
  "Is there already a place to store “this guard has completed onboarding,” or must one be added?",
].forEach((s) => kids.push(bullet(valRef, s)));
kids.push(gap(10));
kids.push(callout([
  "What is proven today: the reporting console exposes edit and soft-delete/archive on SAM OnSite activities, and SAM can write structured-essential corrections through conversation context. The archive flag exists and can be used now. What is not yet built: the filtering that actually excludes archived/training data from live surfaces — it is currently still visible in production and is a later-stage enhancement. No production change has been made.",
], TEALD, "Evidence basis"));

kids.push(new Paragraph({ spacing: { before: 90 }, border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 5, space: 5 } },
  children: [t("Describes intended behaviour for internal build and testing only. No production change is made until Vincent Smeyers has signed off.", { size: 16, it: true, color: GREY })] }));

/* ================= NUMBERING ================= */
const numbering = { config: [] };
[decRef, inRef, outRef, rollRef, dodRef, valRef].forEach((reference) => numbering.config.push({
  reference, levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
    style: { run: { font: SANS, size: 19, color: TEAL }, paragraph: { indent: { left: 360, hanging: 220 } } } }],
}));
[flowRef].forEach((reference) => numbering.config.push({
  reference, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
    style: { run: { font: SANS, size: 19, bold: true, color: TEAL }, paragraph: { indent: { left: 400, hanging: 260 } } } }],
}));

/* ================= DOCUMENT ================= */
const doc = new Document({
  creator: "Pronect", title: "SAM OnSite — Guard Onboarding (detailed specification)", description: "Guard onboarding detailed specification",
  styles: { default: { document: { run: { font: SANS, size: 19, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: SERIF, size: 24, bold: true, color: INK } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: SANS, size: 19, bold: true, color: TEALD } },
    ] },
  numbering,
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN_TB, bottom: MARGIN_TB, left: MARGIN_LR, right: MARGIN_LR } } },
    headers: { default: new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 3, space: 3 } }, spacing: { after: 0 },
      children: [ t("SAM OnSite — Guard Onboarding", { size: 15, color: GREY }), new TextRun({ children: [new Tab()], font: SANS, size: 15 }), t("Detailed specification", { size: 15, color: GREY }) ] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 3, space: 3 } }, spacing: { before: 0 },
      children: [ t("Confidential — Pronect internal", { size: 15, color: GREY }), new TextRun({ children: [new Tab(), "Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: SANS, size: 15, color: GREY }) ] })] }) },
    children: kids,
  }],
});
const OUT = path.join(__dirname, "SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Wrote", OUT, "(" + buf.length + " bytes)"); });
