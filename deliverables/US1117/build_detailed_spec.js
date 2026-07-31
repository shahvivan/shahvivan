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
  ["FR-03", "C", "Mandatory", "There is no path to the home screen except by completing it."],
  ["FR-04", "C", "Refresher", "A completed guard can reopen the training; reopening never clears or regresses completion."],
  ["FR-05", "C", "Language", "All onboarding voice and screen copy is English for the MVP."],
  ["FR-06", "C", "Permissions gate", "SAM explains that the app needs permissions to work at all, then offers a link that opens SAM OnSite’s settings page, where the guard taps Permissions and enables everything; back in onboarding they confirm. The link is duplicated into the onboarding flow rather than sending the guard to the app’s own settings screen — that would hand them a route out of a required exercise. Guidance must name the level needed, since grants can be “only while using the app”, “ask every time”, or for location “all the time”. On the current interface the set is camera, location, microphone, notifications and physical activity, all starting “not allowed” on a fresh install. Setup should also switch off the phone’s automatic permission removal for unused apps (“Manage app if unused”), which would otherwise silently strip camera, location and microphone after months of disuse. Until microphone permission is granted the guard cannot speak, so the typed confirmation must work here. If a permission genuinely cannot be granted (device policy or hardware), the guard reaches a support screen and is never left trapped."],
  ["FR-07", "P", "Transcription check", "A speaking step is credited only after at least one successful voice transcription; an empty or failed transcription does not advance the step. The guard’s first Talk press — the ask-a-question step — doubles as the microphone check: SAM confirms what it heard before crediting it. Steps are judged on the guard demonstrating the intended action and on the resulting report state — semantically-equivalent English is accepted, never an exact phrase."],
  ["FR-08", "P", "Ask a question", "The guard asks SAM a spoken question and SAM answers. Nothing live may result from this step. Note the current app records even procedure questions as activities (e.g. “How do I hand out a key?” became a Procedure activity), so either the practice context suppresses that record or any record it creates carries the training flag from creation, like the practice report."],
  ["FR-09", "P", "One practice report", "Reporting the fictional incident creates exactly one practice report. Retries or repeated messages must not create additional reports."],
  ["FR-10", "C", "Voice correction", "A spoken correction (Dell → MacBook) updates the same report in place, with no duplicate. (The exact way the app updates the report is an engineering confirmation — see Confirm with engineering.)"],
  ["FR-11", "C", "Manual correction", "Long-pressing the guard’s own message allows a manual edit that, when saved, updates the same report; the final report contains both “MacBook” and “reception.” (Whether the app already supports this edit is an engineering confirmation.)"],
  ["FR-12", "C", "Marked at creation", "The training / archive flag is set atomically when the report is created — so a report never exists un-marked. If it cannot be created already marked, it is not created at all. The report card the guard sees carries a visible “Training” label, so it is obviously practice."],
  ["FR-13", "C", "Kept out of live surfaces", "The practice report must appear in no live surface (see the Data isolation callout). The platform has the training flag, but the FILTER that excludes marked data is not built yet — data is currently visible in production — so this exclusion filter is a required, not-yet-built enhancement; production is blocked until it is built and verified, and the exercise runs only in DEV/UAT until then."],
  ["FR-14", "P", "Completion integrity", "Completion is recorded server-side and versioned (version 1), issued once per guard, and set only after each required action is observed — never a quiz or an “I’m done” button."],
  ["FR-15", "P", "Fail-closed", "If safe isolation of practice data cannot be assured at any step, the step fails and completion is not marked; the guard stays in the exercise."],
  ["FR-16", "C", "Hand-off", "On confirmed completion the guard reaches the normal home screen (Talk / Capture / Type)."],
  ["FR-17", "P", "Way out for a real incident", "Low priority, deliberately minimal: a way to leave onboarding, report a real incident through normal reporting, and return to the saved stage without it counting as finishing or skipping. In practice a guard deals with a genuine emergency however they always have, and new guards are accompanied early on — so this is a safeguard, not a focus of the build. Keep the implementation small."],
  ["FR-18", "P", "Idempotency / concurrency", "One onboarding record and one active practice-report reference per guard. Retries do not duplicate; the newest valid server-side progress wins with no regression; the completion operation is idempotent."],
  ["FR-19", "P", "Privacy / telemetry", "Telemetry records only stage and error codes and completion metrics. It does not keep permission or NFC state beyond need, does not log raw voice/transcript content beyond existing policy, and never routes fictional incident content into operational analytics."],
  ["FR-20", "C", "Deferred visual support", "Extra visual help (tooltips, highlight animations) for dark or noisy sites is deferred; the need is judged during testing, not built for the MVP. (Scope decision — no separate test.)"],
  ["FR-21", "C", "Tone and engagement", "The experience must feel like receiving a useful new tool, not sitting a test. SAM opens with a warm, personal greeting that names what it does for the guard and why that helps them, frames each step as a benefit, encourages briefly after each success, uses short everyday language, and closes warmly. A dry “do this, then that” sequence does not meet this requirement. See Tone and script for the principles and reference copy."],
  ["FR-22", "C", "NFC gate", "After permissions, SAM links the guard to the NFC setting — NFC reads the checkpoint tags. The guard switches it on, returns, and confirms. Not every site uses NFC, but having it on costs nothing and avoids a failure later, so it is part of setup for everyone."],
  ["FR-23", "C", "Background tracking gate", "The third and last setup gate: background tracking, which lives in SAM OnSite’s own settings screen. Same shape as the other two — SAM explains why it is needed, links straight to the setting from inside onboarding, the guard switches it on and returns to confirm. After this gate every setting the app needs is active and the practice steps can begin."],
  ["FR-24", "X", "Check the settings rather than asking", "For the MVP the guard confirms each gate themselves. Reading the settings back would be materially better: it removes a step where a guard can confirm without having done it, and converts a confusing failure mid-shift into a clear message during onboarding. The Location screen already surfaces a “Permissions needed” status, so this is at least partly achievable today; how far it extends to NFC and background tracking is for engineering to confirm. Treat as the target state, not MVP scope."],
  ["FR-25", "P", "Tips woven in, never recited", "The practical tips from the train-the-trainer deck are taught inside the practice steps, at the moment each becomes useful — SAM offering one as a natural aside, or the guard meeting the situation the tip is about — rather than listed on a screen or read out in sequence. A slide of tips read aloud does not meet this requirement. Which tips, and where each best fits, to be agreed with product once the slide is shared."],
];
const AC = [
  ["AC-01", "FR-01", "Given an active guard who has not completed onboarding, when they sign in, then it launches automatically before the home screen is reachable."],
  ["AC-02", "FR-02", "Given an existing guard (account predating this feature) who has not completed it, when they sign in, then they receive it exactly once; after completion a later sign-in goes straight to home."],
  ["AC-03", "FR-03", "Given onboarding in progress, when the guard tries to skip, then there is no path to home except completion (or the minor way out for a real incident)."],
  ["AC-04", "FR-06", "Given the permissions gate, when the guard follows the link, enables the permissions and returns, then they confirm by typing and the flow continues; the link opened from inside onboarding and offered no route into the app itself. The typed confirmation works before microphone permission is granted. If a permission cannot be granted at all, the guard reaches a support screen and is not trapped."],
  ["AC-05", "FR-07", "Given a speaking step, when transcription fails or returns empty, then the step is not credited and the guard is asked to try again."],
  ["AC-06", "FR-08 / FR-09", "Given the guard asks a question, then SAM answers and nothing live results — any record the platform logs carries the training flag; given the guard then reports the incident, then exactly one practice report exists (retries add none)."],
  ["AC-07", "FR-10", "Given a practice report naming a Dell, when the guard says to change it to a MacBook, then the same report updates and no second report is created."],
  ["AC-08", "FR-11", "Given the guard long-presses their own message and saves “A MacBook was stolen at reception,” then the final report contains both “MacBook” and “reception.”"],
  ["AC-09", "FR-12 / FR-13", "Given a practice report at any stage, then it is flagged training at creation; once the exclusion filter is built it appears in no live surface; until the filter exists the exercise runs only in DEV/UAT and does not go to production."],
  ["AC-10", "FR-14 / FR-15", "Given every required action observed and isolation assured, when completion runs, then version 1 is recorded server-side; if isolation cannot be assured, completion is withheld."],
  ["AC-11", "FR-04 / FR-16", "Given a completed guard, when they finish, then they reach home; and reopening the training later preserves completion."],
  ["AC-12", "FR-18", "Given duplicate concurrent sessions or retried messages, then no duplicate record or report is created and progress does not regress."],
  ["AC-13", "FR-17", "Given a guard leaves onboarding to report a real incident, then normal reporting opens and, on return, the saved stage resumes with completion neither granted nor skipped."],
  ["AC-14", "FR-05", "Given any onboarding screen or prompt, then all voice and screen copy is in English."],
  ["AC-15", "FR-19", "Given a run, when telemetry and analytics are inspected, then only stage/error codes and completion metrics are recorded, with no raw voice/transcript content and no fictional incident content in operational analytics."],
  ["AC-16", "FR-21", "Given the full run-through, then it opens with a warm personal greeting, each step is framed as a benefit with a brief encouragement after it, and the close is warm — signed off by product on a read-through, and at least four of five test guards describing it as welcoming rather than test-like. If fewer do, the copy is revised and re-run."],
  ["AC-17", "FR-22", "Given the NFC gate, when the guard follows the link, switches NFC on and returns, then they confirm and the flow continues to the third gate."],
  ["AC-18", "FR-23", "Given the background-tracking gate, when the guard follows the link into SAM OnSite’s settings, switches it on and returns, then they confirm and setup is complete — every setting the app needs is now active and the practice steps begin."],
  ["AC-19", "FR-25", "Given a full run-through, then the tips from the train-the-trainer deck appear inside the practice steps at the moment each is useful, and none is presented as a list or read out in sequence."],
];
const UAT = [
  ["UAT-01", "AC-01", "New guard’s first sign-in triggers onboarding.", "Shown before home; progress record created."],
  ["UAT-02", "AC-02", "Existing (pre-feature) guard is issued it once.", "Runs once; next sign-in after completion goes to home."],
  ["UAT-03", "AC-03", "Attempt to skip onboarding. (negative)", "No skip path to home; only completion or the real-incident exit."],
  ["UAT-04", "AC-04", "Complete the permissions gate on a fresh install, following the link and returning.", "Link opens the app’s settings page from within onboarding; no route into the app itself; confirmation advances the flow."],
  ["UAT-05", "AC-04", "A permission cannot be granted at all (device policy). (negative)", "Support screen reached; guard not trapped."],
  ["UAT-06", "AC-05", "Force a failed / empty transcription at the first speaking step. (negative)", "Step not credited; retry prompted."],
  ["UAT-07", "AC-06", "Ask SAM a question.", "SAM answers; nothing live — any logged record carries the training flag."],
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
  ["UAT-19", "AC-13", "Leave onboarding to report a real incident, then return.", "Normal reporting opens; returns to saved stage; completion neither granted nor skipped."],
  ["UAT-20", "AC-14", "Review all onboarding copy.", "All copy is English."],
  ["UAT-21", "AC-15", "Inspect telemetry and analytics during a run. (negative)", "Only stage/error + completion metrics; no raw content; no fictional data in analytics."],
  ["UAT-22", "AC-16", "Full run-through with at least five test guards, reviewing tone.", "Warm personal opening; each step framed as a benefit with brief encouragement; warm close. At least four of five describe it as welcoming, not a test; product signs off the copy."],
  ["UAT-23", "AC-04", "Confirm the permissions step by typing or tapping, before microphone permission is granted.", "Typed/tapped “done” is accepted; no voice needed at this point."],
  ["UAT-24", "AC-17", "NFC gate: follow the link, switch NFC on, return and confirm.", "Link opens the NFC setting; confirmation advances to the background-tracking gate."],
  ["UAT-25", "AC-18", "Background-tracking gate: follow the link into the app’s settings, switch it on, return and confirm.", "Setup ends with permissions, NFC and background tracking all active; practice steps begin."],
  ["UAT-26", "AC-19", "Full run-through, reviewing where the train-the-trainer tips appear.", "Each tip lands inside a practice step at a useful moment; none is read out as a list."],
];
const TRACE = [
  ["Auto-launch + refresher", "FR-01", "AC-01", "UAT-01", "Shown before home; refresher reopens"],
  ["Every guard once, incl. existing", "FR-02", "AC-02", "UAT-02", "Issued once per guard"],
  ["Mandatory", "FR-03", "AC-03", "UAT-03", "No skip path"],
  ["Refresher keeps completion", "FR-04", "AC-11", "UAT-18", "Completion preserved on reopen"],
  ["English MVP", "FR-05", "AC-14", "UAT-20", "All copy English"],
  ["Permissions gate, never a dead end", "FR-06", "AC-04", "UAT-04/05/23", "Link inside the flow; typed confirm works; not trapped"],
  ["NFC gate", "FR-22", "AC-17", "UAT-24", "NFC switched on and confirmed"],
  ["Background-tracking gate", "FR-23", "AC-18", "UAT-25", "All three settings active before practice begins"],
  ["Check settings rather than ask (target)", "FR-24", "—", "—", "Later stage; feasibility to confirm"],
  ["Tips woven in, never recited", "FR-25", "AC-19", "UAT-26", "Tips land in context, not as a list"],
  ["Transcription must succeed", "FR-07", "AC-05", "UAT-06", "Failed transcription not credited"],
  ["Ask a question (nothing live)", "FR-08", "AC-06", "UAT-07", "Answer given; nothing live — any record training-flagged"],
  ["One practice report", "FR-09", "AC-06", "UAT-08", "Exactly one report"],
  ["Voice correction, same report", "FR-10", "AC-07", "UAT-09/10", "Same report; no duplicate"],
  ["Manual edit, final MacBook + reception", "FR-11", "AC-08", "UAT-11/12", "Final report has both words"],
  ["Marked training at creation", "FR-12", "AC-09", "UAT-14", "Never un-marked; fail-closed at create"],
  ["Kept out of live surfaces (filter unbuilt)", "FR-13", "AC-09", "UAT-13", "Absent once filter built; DEV/UAT until then"],
  ["Completion observed, server-side, versioned", "FR-14", "AC-10", "UAT-15", "Version 1 server-side"],
  ["Fail-closed isolation", "FR-15", "AC-10", "UAT-14", "Completion withheld on failure"],
  ["Reaches home", "FR-16", "AC-11", "UAT-15", "Home after completion"],
  ["Way out for a real incident (minor)", "FR-17", "AC-13", "UAT-19", "Reporting opens; stage resumes"],
  ["Idempotency / concurrency", "FR-18", "AC-12", "UAT-16/17", "No duplicates; no regression"],
  ["Privacy / telemetry", "FR-19", "AC-15", "UAT-21", "No raw content; no fictional data in analytics"],
  ["Deferred visual support", "FR-20", "—", "—", "Scope decision; evaluated in testing"],
  ["Feels like a welcome, not a test", "FR-21", "AC-16", "UAT-22", "Warm intro; benefit framing; test guards agree"],
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
  "It must feel like receiving a useful new tool, not sitting a test — SAM introduces itself and the tone stays warm and encouraging throughout.",
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
  "A universal, application-level onboarding that launches at first sign-in and is reopenable as a refresher, replacing the app-onboarding steps currently covered in the onboarding presentation.",
  "Device setup inside the flow: permissions, NFC and background tracking, each reached by a link duplicated into onboarding so the guard is never handed a route out.",
  "One guided pass through: ask a question, report a fictional incident, correct it by voice and by manual edit — delivered in a warm, welcoming tone.",
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
  "SAM greets the guard and introduces itself; the screen also makes clear this is required, everything in it is pretend, it is short, and progress is saved — leaving and returning resumes at the same step. A step counter shows how far along they are.",
  "First setup gate — permissions. SAM explains the app cannot work without them, then links to SAM OnSite’s settings page, where the guard taps Permissions and enables everything at the level asked for; back in onboarding they type “done”. The link is duplicated into the flow, not a jump to the app’s settings screen, so it offers no way out of a required exercise. Setup also switches off the phone’s automatic permission removal for unused apps, so access is not silently lost after a quiet month.",
  "Second gate — NFC, which reads the checkpoint tags: a link to the setting, the guard switches it on, returns and confirms. Not every site uses NFC, but turning it on costs nothing and prevents a failure later.",
  "Third gate — background tracking, in SAM OnSite’s own settings screen: same shape again, link, switch on, return and confirm. With that, every setting the app needs is active and the practice steps can begin.",
  "The guard presses Talk and asks SAM a practice question; the first successful transcription doubles as the microphone check. SAM answers, then reassures them they can ask anything — nothing live results; any record the platform logs carries the training flag.",
  "The guard reports “A Dell laptop was stolen.” SAM asks for anything missing, creates one practice report, and shows it on screen as a card visibly labelled “Training” — seeing the finished report is what makes the two correction steps meaningful.",
  "The guard says to change “Dell” to “MacBook”; the same report and message update; no second report appears.",
  "The guard long-presses their message, edits it to “A MacBook was stolen at reception,” and saves; the final report shows both “MacBook” and “reception.”",
  "SAM closes on a high note — “You’re ready” — and a Go-to-home button lands the guard on the normal home screen (Talk / Capture / Type). Completion is marked only after each step is observed for real. A completed guard may reopen the training later without losing completion.",
].forEach((s) => kids.push(numitem(flowRef, s)));

/* 05 Tone and script */
kids.push(eyebrow("05", "Tone and script"));
kids.push(p([
  t("This is a guard’s first meeting with a tool that is meant to make their job easier, so it should land like unwrapping something good — not like being tested. SAM introduces itself, and every step is framed as a win for the guard rather than an instruction to follow."),
]));
kids.push(subhead("HOW IT SHOULD SOUND"));
const toneRef = "tone";
[
  "Warm and first-person. SAM has a personality: it greets the guard by way of introduction and says plainly what it does for them and why that helps.",
  "Every step sells the benefit, not the mechanic — “I wrote that up while you talked,” not “an Activity record has been created.”",
  "Encourage after each success, briefly. A short “that’s it” or “fixed — no forms, no rewriting” is enough.",
  "Short lines, everyday words, no system jargon. Never a numbered list of instructions read aloud.",
  "Close on a high note and make SAM available: “You’re ready — you can ask, report, and fix reports whenever you need me.”",
].forEach((s) => kids.push(bullet(toneRef, s)));
kids.push(p([chip("P"), t("  Reference copy below — the intended tone, to be polished with product before build. It is not final wording, and no line here is a locked string.")], { after: 50 }));
kids.push(table([1500, 4074, 4074], ["Stage", "SAM says", "Guard does"], [
  ["Welcome", "“Hello, I’m SAM. From today I do the paperwork with you — a few minutes and you’ll know the essentials.”", "Reads; taps Start."],
  ["Practice context", "“Everything here is pretend — nothing you say is filed as a real report, so try things without worrying. I’ll save where you get to.”", "Understands nothing is real."],
  ["Permissions", "“First, let’s switch me on properly — I can’t do much without a few permissions. Tap here, turn them all on, then come back and type done.”", "Follows the link, enables all, returns, types “done”."],
  ["NFC", "“Next, NFC — that’s what reads the checkpoint tags. Tap here, switch it on, and tell me when it’s done.”", "Enables NFC, returns, says or types “done”."],
  ["Background tracking", "“Last one — background tracking, so I keep working when your screen is off. Tap here, switch it on, then come back. That’s everything set up.”", "Enables background tracking, returns, confirms."],
  ["Ask a question", "“Ask me anything you’d ask a colleague — press Talk and try: how do I report an incident?” … “Any time you’re unsure, just ask.”", "Presses Talk and asks — the first transcription doubles as the mic check. SAM answers; nothing live results."],
  ["Report", "“Now tell me about an incident the way you’d tell a colleague. Let’s pretend — a Dell laptop was stolen.” … “Done. I wrote that up while you talked — have a look.”", "Reports; one practice report is created and shown on screen."],
  ["Voice fix", "“Got a detail wrong? Just say so — tell me to change the Dell to a MacBook.” … “Fixed. No forms, no rewriting.”", "“Change the Dell to a MacBook.” Same report updates."],
  ["Manual fix", "“Prefer to type? Press and hold your message and edit it yourself — add where it happened.”", "Edits message to “…MacBook…at reception.” and saves."],
  ["Finish", "“You’re ready — you can ask, report, and fix reports whenever you need me.”", "Taps Go to home once completion is confirmed."],
]));

/* 06 Functional requirements */
kids.push(eyebrow("06", "Functional requirements"));
kids.push(p([
  t("Each requirement is traced to acceptance criteria and tests in the Traceability matrix. Note that "),
  t("FR-21 (tone and engagement)", { bold: true }),
  t(" carries the same weight as any functional rule here — how this feels to the guard is a requirement, not decoration."),
], { after: 50 }));
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
  "Reports and event pages · dashboards and KPIs · analytics and metrics · exports and scheduled reports · alerts and notifications (including supervisor alerts) · webhooks and integrations · the Pronect Action Tracker · the event outbox and any queue feeding the above · read replicas, caches, and search indexes · the SAM conversation / transcript store · the guard’s own activity list and search. Real transcripts show activities send supervisor alerts, move KPIs, and create Action Tracker items today — and a dispatched alert cannot be recalled, so nothing may be emitted before the training flag is in place.",
], TEAL, "Every consumer the exclusion filter must cover"));
kids.push(subhead("COMPLETION IS SERVER-SIDE, EARNED, AND SAFE UNDER PRESSURE"));
kids.push(p("Completion is recorded only after the system observes each real action; it lives on the server (never trusted from the phone), is versioned (version 1), issued once per guard, resumes safely across app close or device change, and cannot be duplicated by two concurrent sessions. If practice data cannot be safely isolated at any point, completion is withheld and the guard stays in the exercise — protecting live data outranks finishing.", { after: 40 }));

/* 09 Errors */
kids.push(eyebrow("09", "Error and recovery behaviour"));
kids.push(table([3000, 6648], ["Condition", "Behaviour"], [
  ["A setting is still off after the guard confirms", "For the MVP the flow takes the confirmation and moves on, so this surfaces later as a feature that does not work — which is the argument for reading the settings back instead of asking (FR-24). Anything that genuinely cannot be enabled reaches a support screen so the guard is not trapped."],
  ["Network, speech, or save failure", "Preserve progress; show a retryable error; never silently lose the current step or an edit."],
  ["Voice correction misheard", "SAM re-asks; the guard retries by voice or uses the manual long-press edit."],
  ["Manual save failure", "Show the failure and keep the typed text for retry; the report is never left half-edited."],
  ["Isolation cannot be assured", "Fail closed: withhold completion, keep the guard in the exercise, record only a stage/error code — never the practice content."],
  ["A real incident comes up", "The guard can leave onboarding, report it through normal reporting, and return to the saved stage; completion is neither granted nor skipped. Minor path — see the requirement note."],
]));

/* 10 Acceptance criteria */
kids.push(eyebrow("10", "Acceptance criteria"));
kids.push(table([850, 1350, 7448], ["ID", "Traces to", "Criterion"], AC.map((r) => [r[0], r[1], r[2]])));

/* 11 UAT */
kids.push(eyebrow("11", "Test matrix (DEV / UAT)"));
kids.push(p("Positive and negative tests. Negatives cover skip, a permission left off, a permission that cannot be granted, NFC off, transcription failure, voice-correction failure, manual-save failure, isolation failure, concurrency, and network/device resume.", { after: 50 }));
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
  "Which live consumers must the new exclusion filter cover — at minimum supervisor alerts, KPIs, and the Pronect Action Tracker, all confirmed live today — and where are reports created so the training flag is set at creation? (The flag exists; the filter does not yet.)",
  "Which permissions and grant levels does the app need (e.g. location “all the time” vs “while using”), can it read each one’s current level to verify it, and can it switch off the phone’s automatic permission removal for unused apps? The app’s Location screen already shows a “Permissions needed” status and an “Open system settings” link, so state-reading and the settings link exist at least for location.",
  "Is there already a place to store “this guard has completed onboarding,” or must one be added?",
].forEach((s) => kids.push(bullet(valRef, s)));
kids.push(gap(10));
kids.push(callout([
  "What is proven today: the reporting console exposes edit and soft-delete/archive on SAM OnSite activities, and SAM can write structured-essential corrections through conversation context. The archive flag exists and can be used now. Real app screens add more: activities have live consequences (supervisor alerts, KPI impact, items in the Pronect Action Tracker — exactly the surfaces the filter must cover); an activity is a report card plus transcript, matching the message-is-the-report model; permission state is readable and a settings link exists at least on the Location screen (“Permissions needed” / “Open system settings”); and a spoken question was recorded as a Procedure activity, so even the ask step may create a record today. What is not yet built: the filtering that actually excludes archived/training data from live surfaces — it is currently still visible in production and is a later-stage enhancement. No production change has been made.",
], TEALD, "Evidence basis"));

kids.push(new Paragraph({ spacing: { before: 90 }, border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 5, space: 5 } },
  children: [t("Describes intended behaviour for internal build and testing only. No production change is made until Vincent Smeyers has signed off.", { size: 16, it: true, color: GREY })] }));

/* ================= NUMBERING ================= */
const numbering = { config: [] };
[decRef, inRef, outRef, rollRef, dodRef, valRef, toneRef].forEach((reference) => numbering.config.push({
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
