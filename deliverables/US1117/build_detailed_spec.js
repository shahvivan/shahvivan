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
  ["FR-02", "C", "Audience", "Every active guard receives it exactly once, including guards whose accounts existed before this feature."],
  ["FR-03", "C", "Mandatory", "There is no path to the home screen except by completing it, other than three defined exits: the real-incident exit (FR-17), the stuck-guard release (FR-28), and leaving a refresher when already complete (FR-04). Nothing else reaches home, and none of the three marks the guard complete."],
  ["FR-04", "C", "Refresher", "A completed guard can reopen the training, and reopening never clears or regresses completion. A refresher starts a fresh practice run with its own practice report; the one-active-report rule is per run, not per lifetime. A guard who is already complete can leave the refresher at any point and return to the home screen."],
  ["FR-05", "C", "Language", "All onboarding voice and screen copy is English for the MVP."],
  ["FR-06", "C", "Permissions gate", "SAM explains that the app needs permissions to work at all, then opens the phone’s settings page for SAM OnSite through a link held inside onboarding, where the guard taps Permissions and enables each one; back in onboarding they confirm. The link opens that one page and returns. It never drops the guard into the app’s ordinary navigation, which would hand them a route out of a required exercise. The on-screen words must name the grant level needed, since Android offers “only while using the app”, “ask every time”, and for location “all the time”; on most Android versions “all the time” is a second, separate step after the first grant, so the copy has to say so. On the current interface the set is camera, location, microphone, notifications and physical activity, all starting “not allowed” on a fresh install. Setup should also switch off the phone’s automatic permission removal for unused apps (“Manage app if unused”), which would otherwise silently strip camera, location and microphone after months of disuse. Until microphone permission is granted the guard cannot speak, so confirmation here is a tap, not a spoken or typed magic word. The tap triggers a read-back (FR-24), not a trust: the gate advances only when the permissions are actually on. If a permission genuinely cannot be granted (device policy or hardware), the guard reaches a support screen and is never left trapped."],
  ["FR-07", "P", "Transcription check", "A speaking step is credited only after at least one successful voice transcription; an empty or failed transcription does not advance the step. The guard’s first Talk press, at the ask-a-question step, doubles as the microphone check: SAM confirms what it heard before crediting it. Steps are judged on the guard demonstrating the intended action and on the resulting report state. Semantically-equivalent English is accepted and an exact phrase is never required, which AC-05 tests by crediting a step on a differently-worded but equivalent utterance."],
  ["FR-08", "P", "Ask a question", "The guard asks SAM a spoken question and SAM answers. Nothing live may result from this step. The current app records even procedure questions as activities (e.g. “How do I hand out a key?” became a Procedure activity). Because the exclusion filter does not exist yet, only suppression is safe: the practice context must not write that activity at all. The answer to the suggested question is written and reviewed in advance, since it is the first thing SAM ever does for this guard and a hollow answer on a fresh site is worse than no question. The prompt therefore names that question rather than inviting an open one, and if the guard asks something else and SAM has no answer, it says so plainly and moves on rather than improvising."],
  ["FR-09", "P", "One practice report", "Reporting the fictional incident creates exactly one practice report. Retries or repeated messages must not create additional reports. SAM asks at most two follow-up questions and “I don’t know” always advances, since the guard is being asked for details of a laptop that does not exist."],
  ["FR-10", "X", "Voice correction", "A spoken correction (whatever the guard called the laptop → MacBook) updates the same report in place, with no duplicate. SAM’s prompt must not assume the guard used the word “Dell”, since they were invited to report in their own words. The train-the-trainer deck says a guard can change scenario within the same conversation when SAM picks the wrong one, and can reopen a conversation to modify or add information, which supports the model but does not prove that a correction updates the report in place without creating a second one. That last point is what engineering must confirm. If a spoken correction is parsed as a new incident, this step teaches the opposite of what is intended. See Confirm with engineering."],
  ["FR-11", "C", "Manual correction", "Long-pressing the guard’s report message opens a manual edit that, when saved, updates the same report, and the final report contains both “MacBook” and “reception”. The train-the-trainer deck confirms the mechanism and the gesture: clicking and holding the text turns it into edit mode so it can be changed, given there as the way to fix spelling errors in names. Two things must still be unambiguous in the build: SAM names the change it wants, so the guard is never asked for a detail nobody gave them, and the target message is identified on screen, since by this point the guard has sent a question, a report, up to two answers and a spoken correction. Describing it in words alone is not enough."],
  ["FR-12", "C", "Marked at creation", "The training / archive flag is set atomically when the report is created, so a report never exists un-marked. If it cannot be created already marked, it is not created at all. The report card the guard sees carries a visible “Training” label, so it is obviously practice."],
  ["FR-13", "C", "Kept out of live surfaces", "The practice report must appear in no live surface (see the Data isolation callout). The platform has the training flag, but the filter that excludes marked data is not built yet. Data is currently visible in production, so this exclusion filter is a required, not-yet-built enhancement; production is blocked until it is built and verified, and the exercise runs only in DEV/UAT until then."],
  ["FR-14", "P", "Completion integrity", "Completion is recorded server-side and versioned (version 1), issued once per guard, and set only after each required action is observed. A quiz or an “I’m done” button does not count."],
  ["FR-15", "P", "Fail-closed", "If safe isolation of practice data cannot be assured at any step, the step fails and completion is not marked; the guard stays in the exercise."],
  ["FR-16", "C", "Hand-off", "On confirmed completion the guard reaches the normal home screen (Talk / Capture / Type)."],
  ["FR-17", "P", "Way out for a real incident", "A way to leave onboarding, report a real incident through normal reporting, and return to the saved stage, with completion neither granted nor skipped. Small to build, but it is the only door out of a flow that cannot otherwise be left, so it is required rather than optional. It is reachable from every step, including the setup gates and the refresher. Crossing it must be unmistakable: the practice banner clears and the guard is told plainly that this one is live and goes to their supervisor. A guard who has just been told nothing counts must not carry that belief through the door."],
  ["FR-18", "P", "Idempotency / concurrency", "One onboarding record per guard, and one active practice-report reference per practice run (a refresher is a new run, per FR-04). Retries do not duplicate; the newest valid server-side progress wins with no regression; the completion operation is idempotent."],
  ["FR-19", "P", "Privacy / telemetry", "Telemetry records only stage and error codes and completion metrics. It does not keep permission or NFC state beyond need, does not log raw voice/transcript content beyond existing policy, and never routes fictional incident content into operational analytics."],
  ["FR-20", "C", "Deferred visual support", "Extra visual help (tooltips, highlight animations) for dark or noisy sites is deferred; the need is judged during testing, not built for the MVP. (Scope decision; no separate test.)"],
  ["FR-21", "C", "Tone and engagement", "The experience must feel like being handed something useful, closer to receiving a good tool than sitting a test. SAM introduces itself as the guard’s new assistant, invites them to get to know each other, says what it does for them and why that helps, frames each step as a benefit, encourages briefly after each success, uses short everyday language, and closes warmly. A dry “do this, then that” sequence does not meet this requirement. See Tone and script for the principles and reference copy."],
  ["FR-22", "C", "NFC gate", "After permissions, SAM links the guard to the NFC setting. NFC reads the checkpoint tags. The guard switches it on, returns, and confirms. Not every site uses NFC, but having it on costs nothing and avoids a failure later, so it is part of setup for everyone."],
  ["FR-23", "C", "Background tracking gate", "The third and last setup gate: background tracking. The first two gates sent the guard to the phone’s settings; this one lives in SAM OnSite’s own in-app settings, and the copy must say so. SAM explains why it is needed and links straight to that one setting from inside onboarding, and the guard switches it on and returns to confirm. The link opens the setting alone and returns, keeping the guard clear of the in-app settings screen and its normal navigation. After this gate every setting the app needs is active and the practice steps can begin."],
  ["FR-24", "X", "Read the settings back", "Onboarding promises a ready phone, so where the app can check a setting it must check it rather than ask. Permissions are readable today: the Location screen already surfaces a “Permissions needed” status, and FR-26 relies on reading microphone state at the first Talk press, so the mechanism exists. The permissions gate therefore reads back rather than asks, and a gate that reads back as off returns the guard to it instead of advancing. Whether NFC and background-tracking state are readable the same way is the open engineering question; until it is answered those two are confirmed by the guard, and the residual risk is recorded in Error and recovery behaviour. If reading them back proves cheap, it should be in the MVP too."],
  ["FR-25", "P", "Tips taught in context", "All eight tips from the General recommendations slide of the train-the-trainer deck are taught inside the practice steps, each at the moment it becomes useful. Two are taught by the guard doing them, not by being told: the long-press correction is the manual-fix step, and full-sentence speech is what the report prompt asks for. The rest arrive as a short aside from SAM at the step where they apply. The mapping is in Where each tip lands. A slide of tips listed on a screen or read out in sequence does not meet this requirement."],
  ["FR-26", "P", "Microphone checked at first use", "The first Talk press checks microphone permission before it checks anything else. If the permission is missing, the guard is returned to the permissions gate with an explanation, not left retrying a button that cannot work. This is the one gate the flow verifies rather than trusts, because every step after it is spoken."],
  ["FR-27", "C", "No NFC hardware", "If the device has no NFC chip, the NFC gate is skipped and the flow continues. A guard is never sent to a support screen over hardware their phone does not have, in an exercise they cannot skip."],
  ["FR-28", "P", "Retry floor and release", "Three rules, all cheap. After three failed attempts at a spoken step, typing is offered. On any step, a guard who is simply stuck can say so and reach the home screen with completion unmarked and the stage saved: the release must not depend on a failure counter, since a guard who cannot find the message to long-press never generates a failed attempt. And if the server cannot be reached at first sign-in, the guard passes through rather than being held at a mandatory screen with no signal. Being un-onboarded is a reporting problem; being locked out of the app mid-shift is a safety one."],
  ["FR-29", "P", "A real incident reported inside practice", "For the MVP, one thing: a practice banner stays visible for the whole exercise rather than appearing only on the finished report card, and it carries the way out to normal reporting (FR-17). This matters because a genuine incident reported during practice is flagged training and hidden from the supervisors it should have reached, by the mechanism meant to protect them. Detecting that what the guard just said is not the practice incident, and asking whether it is real, is the better answer and is listed under What the MVP does not include."],
  ["FR-30", "P", "Go-live gate on the filter", "The MVP is built and tested in DEV/UAT only, so nothing here reaches production during the build. Before the switch is turned on in production, the exclusion filter must be present and verified, and that must be enforced by something other than remembering: activities send supervisor alerts, move KPIs and create Action Tracker items today, and a dispatched alert cannot be recalled. The recommended enforcement is a server-side capability check that refuses to start onboarding when the filter is absent, built at go-live rather than in the MVP."],
  ["FR-31", "P", "Monitoring transparency", "Onboarding states in one plain sentence what location and background tracking are used for and what a supervisor can see. Guards are being asked for all-the-time location and background tracking on a work phone; this is the first question they will ask each other, and in a Belgian workforce it is a works-council matter as well as a tone one. Exact wording to be agreed with product and, where required, with the works council."],
];
const AC = [
  ["AC-01", "FR-01", "Given an active guard who has not completed onboarding, when they sign in, then it launches automatically before the home screen is reachable."],
  ["AC-02", "FR-02", "Given an existing guard (account predating this feature) who has not completed it, when they sign in, then they receive it exactly once; after completion a later sign-in goes straight to home."],
  ["AC-03", "FR-03", "Given onboarding in progress, when the guard tries to skip, then there is no path to home except completion or one of the three defined exits (FR-17, FR-28, FR-04), none of which marks them complete."],
  ["AC-26", "FR-24", "Given the permissions gate, when the guard returns and confirms, then the app reads the permission state back; if any required permission is still off, the guard is returned to the gate with the specific permission named, and the gate does not advance."],
  ["AC-04", "FR-06", "Given the permissions gate, when the guard follows the link, enables the permissions and returns, then a tap triggers a read-back and the flow continues only if every required permission is on; the link opened the phone’s settings page for SAM OnSite from inside onboarding and returned there, offering no route into the app’s ordinary navigation. Confirmation works before microphone permission is granted. If a permission cannot be granted at all, the guard reaches a support screen and is not trapped."],
  ["AC-05", "FR-07", "Given a speaking step, when transcription fails or returns empty, then the step is not credited and the guard is asked to try again; and when the guard says the right thing in different words, then the step is credited."],
  ["AC-06", "FR-08 / FR-09", "Given the guard asks a question, then SAM answers and no activity record of any kind is written for that step, verified in the activity table rather than by inspecting a flag; given the guard then reports the incident, then exactly one practice report exists (retries add none)."],
  ["AC-07", "FR-10", "Given a practice report about the suggested laptop, however the guard worded it, when they say it was a MacBook, then the same report updates and no second report is created."],
  ["AC-08", "FR-11", "Given SAM has named the change and identified the target message on screen, when the guard long-presses their report message and saves it with the location added, then the final report contains both “MacBook” and “reception”."],
  ["AC-09", "FR-12 / FR-13", "Given a practice report at any stage, then it is flagged training at creation and the card shown to the guard carries a visible “Training” label; once the exclusion filter is built it appears in no live surface; until the filter exists the exercise runs only in DEV/UAT and does not go to production."],
  ["AC-10", "FR-14 / FR-15", "Given every required action observed and isolation assured, when completion runs, then version 1 is recorded server-side; if isolation cannot be assured, completion is withheld."],
  ["AC-11", "FR-04 / FR-16", "Given a completed guard, when they finish, then they reach home; and reopening the training later preserves completion."],
  ["AC-12", "FR-18", "Given duplicate concurrent sessions or retried messages, then no duplicate record or report is created and progress does not regress."],
  ["AC-13", "FR-17", "Given a guard leaves onboarding to report a real incident, then the practice banner clears, the guard is told this one is live and reaches their supervisor, normal reporting opens, and on return the saved stage resumes with completion neither granted nor skipped."],
  ["AC-14", "FR-05", "Given any onboarding screen or prompt, then all voice and screen copy is in English."],
  ["AC-15", "FR-19", "Given a run, when telemetry and analytics are inspected, then only stage/error codes and completion metrics are recorded, with no raw voice/transcript content and no fictional incident content in operational analytics."],
  ["AC-16", "FR-21", "Given the full run-through, then it opens with a warm personal greeting, each step is framed as a benefit with a brief encouragement after it, and the close is warm. Signed off by product on a read-through, and at least 80% of a test group of five or more guards describing it as welcoming rather than test-like. If fewer do, the copy is revised and re-run."],
  ["AC-17", "FR-22", "Given the NFC gate, when the guard follows the link, switches NFC on and returns, then they confirm and the flow continues to the third gate."],
  ["AC-18", "FR-23", "Given the background-tracking gate, when the guard follows the link to the in-app setting, switches it on and returns, then they confirm and setup is complete: every setting the app needs is now active and the practice steps begin."],
  ["AC-19", "FR-25", "Given a full run-through, then each of the eight tips from the General recommendations slide appears at the step named in Where each tip lands, none is presented as a list or read out in sequence, and the long-press correction and full-sentence speech are practised rather than described."],
  ["AC-20", "FR-26", "Given a guard who confirmed the permissions gate without granting the microphone, when they press Talk, then they are returned to the permissions gate with an explanation, and are not shown a generic retry."],
  ["AC-21", "FR-27", "Given a device with no NFC hardware, when the flow reaches the NFC gate, then the gate is skipped and the flow continues to background tracking; no support screen is shown."],
  ["AC-22", "FR-28", "Given three failed attempts at a spoken step, then typing is offered; given a guard stuck on any step who asks to leave, then they reach the home screen with completion unmarked and the stage saved; and given no server connectivity at first sign-in, then the guard passes through rather than being held."],
  ["AC-23", "FR-29", "Given a practice run, then a practice banner is visible at every step, including the setup gates, and it offers the way out to normal reporting."],
  ["AC-24", "FR-30", "Given production go-live, then the exclusion filter is present and verified before the switch is enabled, and its absence blocks onboarding from starting rather than relying on release discipline."],
  ["AC-25", "FR-31", "Given the setup gates, then the copy states in plain language what location and background tracking are used for and what a supervisor can see."],
];
const UAT = [
  ["UAT-01", "AC-01", "New guard’s first sign-in triggers onboarding.", "Shown before home; progress record created."],
  ["UAT-02", "AC-02", "Existing (pre-feature) guard is issued it once.", "Runs once; next sign-in after completion goes to home."],
  ["UAT-03", "AC-03", "Attempt to skip onboarding. (negative)", "No skip path to home; only completion or one of the three defined exits, none of which marks completion."],
  ["UAT-04", "AC-04", "Complete the permissions gate on a fresh install, following the link and returning.", "Link opens the phone’s settings page for SAM OnSite and returns; no route into the app’s ordinary navigation; confirmation advances the flow."],
  ["UAT-05", "AC-04", "A permission cannot be granted at all (device policy). (negative)", "Support screen reached; guard not trapped."],
  ["UAT-06", "AC-05", "Force a failed / empty transcription at the first speaking step, then repeat the step in different wording. (negative + positive)", "Step not credited on failure; credited on a differently-worded equivalent."],
  ["UAT-07", "AC-06", "Ask SAM a question, then inspect the activity table.", "SAM answers; no activity row is written for the question step."],
  ["UAT-08", "AC-06", "Report the fictional incident.", "Exactly one practice report exists."],
  ["UAT-09", "AC-07", "Voice-correct Dell to MacBook.", "Same report updated; no duplicate."],
  ["UAT-10", "AC-07", "Voice correction misheard, then retry. (negative)", "SAM re-asks; still exactly one report."],
  ["UAT-11", "AC-08", "Manual long-press edit and save.", "Final report contains “MacBook” and “reception”."],
  ["UAT-12", "AC-08", "Manual save fails, then retry. (negative)", "Error shown; typed text retained; report not half-edited."],
  ["UAT-13", "AC-09", "After the exclusion filter is built, inspect every live surface.", "Absent from all listed consumers; before the filter, run confined to DEV/UAT."],
  ["UAT-14", "AC-09 / AC-10", "Force create-with-flag / isolation failure. (negative)", "Report not created / completion withheld (fail-closed)."],
  ["UAT-15", "AC-10 / AC-11", "Complete all actions with isolation assured.", "Version 1 recorded server-side; guard reaches home."],
  ["UAT-16", "AC-12", "Duplicate concurrent sessions / retried messages within one run. (negative)", "One onboarding record, one report for the run; no progress regression."],
  ["UAT-17", "AC-12", "Network drop mid-step, resume on another device. (negative)", "Progress preserved; resumes at saved stage; no loss."],
  ["UAT-18", "AC-11", "Reopen training after completion (refresher).", "Reopens; completion preserved, not cleared."],
  ["UAT-19", "AC-13", "Leave onboarding to report a real incident, then return.", "Banner cleared and the live warning shown before reporting opens; returns to saved stage; completion neither granted nor skipped."],
  ["UAT-20", "AC-14", "Review all onboarding copy.", "All copy is English."],
  ["UAT-21", "AC-15", "Inspect telemetry and analytics during a run. (negative)", "Only stage/error + completion metrics; no raw content; no fictional data in analytics."],
  ["UAT-22", "AC-16", "Full run-through with a test group of five or more guards, reviewing tone.", "Warm personal opening; each step framed as a benefit with brief encouragement; warm close. At least 80% describe it as welcoming rather than test-like; product signs off the copy."],
  ["UAT-23", "AC-04", "Confirm the permissions step by tapping, before microphone permission is granted.", "The tap is accepted; no voice or typed keyword needed at this point."],
  ["UAT-24", "AC-17", "NFC gate: follow the link, switch NFC on, return and confirm.", "Link opens the NFC setting; confirmation advances to the background-tracking gate."],
  ["UAT-25", "AC-18", "Background-tracking gate: follow the link to the in-app setting, switch it on, return and confirm.", "Setup ends with permissions, NFC and background tracking all active; practice steps begin."],
  ["UAT-26", "AC-19", "Full run-through against the eight-tip mapping.", "All eight covered at the step named; two practised rather than told; none read out as a list."],
  ["UAT-27", "AC-17 / AC-18", "Confirm the NFC or background-tracking gate with the setting still switched off. (negative)", "Accepted for the MVP, since read-back for these two is unconfirmed; the miss surfaces at first use. Recorded as the rationale for extending FR-24."],
  ["UAT-28", "AC-21", "Run the flow on a device with no NFC hardware. (negative)", "NFC gate skipped; flow continues; no support screen."],
  ["UAT-29", "AC-20", "Press Talk having confirmed permissions without granting the microphone. (negative)", "Returned to the permissions gate with an explanation; no endless retry."],
  ["UAT-30", "AC-22", "Fail a spoken step three times; separately, ask to leave from a non-spoken step; separately, first sign-in with no connectivity. (negative)", "Typing offered; the stuck-on-any-step release reaches home with completion unmarked and stage saved; no connectivity does not hold the guard at a mandatory screen."],
  ["UAT-31", "AC-23", "Walk every step, including the setup gates, checking the practice banner.", "Banner visible at every step and offers the way out to normal reporting."],
  ["UAT-32", "AC-24", "At go-live, launch onboarding with the exclusion filter absent. (negative)", "Refuses to start and reports why. Run at go-live, not during MVP build."],
  ["UAT-36", "AC-26", "Confirm the permissions gate with a required permission still switched off. (negative)", "Read-back catches it; the guard is returned to the gate with that permission named; the gate does not advance."],
  ["UAT-34", "AC-09", "Inspect the practice report card on screen.", "Card carries a visible “Training” label at every point it is shown."],
  ["UAT-35", "AC-25", "Read the setup copy as a first-time guard.", "One plain line states what location and background tracking are for and what a supervisor can and cannot see."],
  ["UAT-33", "AC-06", "Answer “I don’t know” to every follow-up question at the report step.", "At most two follow-ups; the step advances; exactly one practice report."],
];
const TRACE = [
  ["Auto-launch + refresher", "FR-01", "AC-01", "UAT-01", "Shown before home; refresher reopens"],
  ["Every guard once, incl. existing", "FR-02", "AC-02", "UAT-02", "Issued once per guard"],
  ["Mandatory", "FR-03", "AC-03", "UAT-03", "No skip path"],
  ["Refresher keeps completion", "FR-04", "AC-11", "UAT-18", "Completion preserved on reopen"],
  ["English MVP", "FR-05", "AC-14", "UAT-20", "All copy English"],
  ["Permissions gate, never a dead end", "FR-06", "AC-04", "UAT-04/05/23/27", "Link inside the flow; tap confirms before the mic is granted; not trapped"],
  ["NFC gate", "FR-22", "AC-17", "UAT-24", "NFC switched on and confirmed"],
  ["Background-tracking gate", "FR-23", "AC-18", "UAT-25", "All three settings active before practice begins"],
  ["Read the settings back", "FR-24", "AC-26", "UAT-36", "Permissions verified, not asked; NFC / background still open"],
  ["Tips taught in context", "FR-25", "AC-19", "UAT-26", "Each tip lands where it is useful"],
  ["Microphone checked at first use", "FR-26", "AC-20", "UAT-29", "Returned to the gate, not left retrying"],
  ["No NFC hardware", "FR-27", "AC-21", "UAT-28", "Gate skipped, not a support screen"],
  ["Retry floor and release", "FR-28", "AC-22", "UAT-30", "Typing offered; never trapped"],
  ["Real incident inside practice", "FR-29", "AC-23", "UAT-31", "Banner at every step; way out offered"],
  ["Go-live gate on the filter", "FR-30", "AC-24", "UAT-32", "Enforced, not remembered (at go-live)"],
  ["Monitoring transparency", "FR-31", "AC-25", "UAT-35", "Copy states what a supervisor sees"],
  ["Visible “Training” label on the card", "FR-12", "AC-09", "UAT-34", "Practice is obvious on screen"],
  ["Bounded clarification", "FR-09", "AC-06", "UAT-33", "Two follow-ups; “I don’t know” advances"],
  ["Transcription must succeed", "FR-07", "AC-05", "UAT-06", "Failed transcription not credited"],
  ["Ask a question (nothing live)", "FR-08", "AC-06", "UAT-07", "Answer given; no activity row written"],
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
  [t("This companion", { bold: true, color: TEALD }), t(" carries the full functional requirements, acceptance criteria, test matrix, and traceability behind the one-page brief. It reads on its own. Every requirement carries a plain-language status: ")],
  [chip("C"), t("  agreed behaviour (some note an implementation detail to confirm).   "), chip("P"), t("  our recommended default.   "), chip("X"), t("  needs an engineering confirmation.")],
], TEAL));
kids.push(gap(40));

/* 01 Purpose & decisions */
kids.push(eyebrow("01", "Purpose and confirmed decisions"));
kids.push(p([
  t("SAM OnSite is Pronect’s voice-first guard assistant: a guard speaks naturally and SAM captures structured, time-stamped reports. This specifies a "),
  t("mandatory, one-time voice onboarding", { bold: true }),
  t(" that runs at a guard’s first sign-in. It has two outputs and both must hold when it ends. The guard can do the three everyday actions (ask a question, report an incident, correct a report), practised on a made-up incident so nothing real is affected. And the phone is ready to work: permissions, NFC and background tracking all on. Getting the device configured is not a preamble to the training, it is half of what onboarding is for. A guard who finishes and then cannot scan a checkpoint has not been onboarded."),
]));
const decRef = "dec";
[
  "It must feel like receiving a useful new tool, not sitting a test. SAM introduces itself and the tone stays warm and encouraging throughout.",
  "Every guard receives it once, including existing guards. It is mandatory and cannot be skipped.",
  "“Generate a report” means reporting a fictional incident by talking to SAM. Applies to all guards, with configuration universal rather than customer-specific.",
  "A completed guard can reopen it later as a refresher without losing completion. English for the MVP.",
  "“Editing logs” means editing the guard’s own report, by voice and by long-pressing their message and saving. The final report must reflect the correction.",
  "Completion is inferred from the system observing the guard perform each action, rather than a quiz or a self-confirmation button.",
  "Onboarding is also the device-setup step. It ends with every setting the app needs switched on, and a guard who cannot get a setting on does not silently pass.",
  "Enhanced visual support (tooltips, highlights) may be evaluated later. It is not required for the MVP.",
].forEach((s) => kids.push(bullet(decRef, s)));

/* 02 Scope */
kids.push(eyebrow("02", "Scope"));
kids.push(subhead("IN SCOPE"));
const inRef = "insc";
[
  "A universal, application-level onboarding that launches at first sign-in and is reopenable as a refresher. It replaces the app-setup slides of the train-the-trainer deck: background tracking opt-in, granting app permissions, and enabling NFC, together with their step-by-step screenshots. It also absorbs the General recommendations tips.",
  "Device setup inside the flow: permissions, NFC and background tracking, each reached by a link duplicated into onboarding so the guard is never handed a route out.",
  "One guided pass through: ask a question, report a fictional incident, correct it by voice and by manual edit, delivered in a warm, welcoming tone.",
  "Server-side, versioned completion tracking that issues version 1 exactly once to every active guard.",
  "Isolation of all fictional practice data from every live surface (dependent on the exclusion filter described later).",
].forEach((s) => kids.push(bullet(inRef, s)));
kids.push(subhead("OUT OF SCOPE"));
const outRef = "outsc";
[
  "Languages other than English; tooltips / highlight animations (deferred for evaluation in testing).",
  "Customer- or site-specific onboarding variants; any change to a live template or protected system template.",
  "Physical deletion / database clean-up of practice data, a separate and later task from the logical archive used here.",
  "Installation and first log-in, which the deck covers on the same slide as the settings. An in-app flow cannot teach a guard to install the app or sign in for the first time, because it only exists once they have. Those two steps stay with the trainer.",
].forEach((s) => kids.push(bullet(outRef, s)));
kids.push(subhead("DELIBERATELY LEFT OUT OF THE MVP"));
kids.push(p("Each of these was considered and is worth doing. None is worth delaying a first build for, and each is written down here so it is a decision rather than an oversight.", { after: 56 }));
const laterRef = "later";
[
  "Detecting that what the guard just reported is not the practice incident, and asking whether it is real. It needs content comparison and a confidence threshold. The MVP covers the same risk more cheaply with a banner at every step and a way out (FR-29, FR-17).",
  "Read-back for NFC and background tracking (part of FR-24). Permissions are read back in the MVP and the microphone is checked at first use. These two wait only because it is not confirmed the app can read their state; if engineering says it can, they belong in the MVP, since a phone that is not ready is an onboarding that did not work.",
  "The server-side capability check that refuses to start onboarding when the exclusion filter is missing (FR-30). Required at go-live, not during a build that never leaves DEV/UAT.",
  "Tooltips and highlight animations for dark or noisy sites (FR-20). Judged during testing on whether they are needed at all.",
].forEach((s) => kids.push(bullet(laterRef, s)));

/* 03 Roles */
kids.push(eyebrow("03", "Roles and ownership"));
kids.push(p("Ownership so no requirement is orphaned across teams. A coordination aid; treat the owners as indicative."));
kids.push(table([1900, 2600, 5148], ["Layer", "Owner (indicative)", "Responsibility"], [
  ["Mobile app", "SAM OnSite", "First-sign-in detection, launch/resume, microphone loop, long-press manual edit, home hand-off."],
  ["SAM / AI", "AI", "Practice (non-reporting) conversation, and capturing the report into its structured fields so a correction updates the same report rather than creating a new one."],
  ["Backend / data", "Backend / Supabase", "Progress record; training flag set at creation; the exclusion filter; idempotency, concurrency, completion."],
  ["Reporting / integrations", "Backend + Reporting", "Build and verify the filter that excludes training-marked data across every live consumer."],
  ["Product / approval", "Product; Vincent Smeyers", "Decisions and acceptance; Vincent signs off any production step."],
  ["QA", "QA", "Run the test matrix, including the negative and concurrency tests, in DEV/UAT."],
]));

/* 04 Flow */
kids.push(eyebrow("04", "End-to-end flow"));
kids.push(p([t("Good to know: in SAM, a guard’s message is what creates and updates their report, so a report can be fixed by telling SAM or by editing the original message. Target duration is roughly 3–5 minutes ("), chip("P"), t(" a product estimate; completion is gated on observed actions rather than elapsed time).")], { after: 60 }));
const flowRef = "flow";
[
  "The guard signs in. If onboarding is not complete, it opens by itself before the home screen, for new and existing guards alike.",
  "SAM greets the guard and introduces itself. The screen also makes clear that this is required, that everything in it is pretend, that it is short, and that progress is saved, so leaving and returning resumes at the same step. A step counter shows how far along they are.",
  "First setup gate: permissions. SAM explains the app cannot work without them, then opens the phone’s settings page for SAM OnSite through a link held inside onboarding, where the guard taps Permissions and enables each one at the level named on screen. Back in onboarding they tap to confirm. The link opens that page and returns, so it offers no way out of a required exercise. Setup also switches off the phone’s automatic permission removal for unused apps, so access is not silently lost after a quiet month.",
  "Second gate: NFC, which reads the checkpoint tags at the gates. Same shape as before. SAM links to the setting, the guard switches it on and comes back to confirm. Not every site uses NFC, but having it on costs nothing and avoids a failure later. A phone with no NFC chip skips this gate entirely.",
  "Third gate: background tracking, in SAM OnSite’s own settings screen. Same shape again: link, switch on, return and confirm. With that, every setting the app needs is active and the practice steps can begin.",
  "The guard presses Talk and asks SAM a practice question. That first successful transcription doubles as the microphone check, and a guard who confirmed the permissions gate without actually granting the microphone is sent back to it here rather than left retrying. SAM answers, then reassures them they can ask anything at all. Nothing live results.",
  "Then the practice report itself. SAM suggests the incident, a laptop that has gone missing, then asks at most two follow-up questions, accepts “I don’t know” for either, creates exactly one report, and puts it on screen as a card visibly labelled “Training”. Seeing the finished report is what makes the two correction steps mean anything.",
  "Correction by voice comes first. The guard asks SAM to change “Dell” to “MacBook”. The same report and the same message update, and no second report appears.",
  "Then the same fix by hand. SAM names the change it wants and points at the message to edit, the guard long-presses their report message, adds that it happened at reception, and saves. The final report shows both “MacBook” and “reception”.",
  "SAM signs off (“That’s everything”) and a Go-to-home button lands the guard on the normal home screen (Talk / Capture / Type). Completion is marked only after each step is observed for real. A completed guard may reopen the training later without losing completion.",
].forEach((s) => kids.push(numitem(flowRef, s)));

/* 05 Tone and script */
kids.push(eyebrow("05", "Tone and script"));
kids.push(p([
  t("This is a guard’s first meeting with a tool meant to make their job easier, so it should land that way. SAM introduces itself, and each step shows the guard something worth having rather than issuing an instruction."),
]));
kids.push(subhead("HOW IT SHOULD SOUND"));
const toneRef = "tone";
[
  "Warm and first-person. SAM introduces itself by name and role, as the guard’s own assistant rather than a system, and says plainly what it does for them.",
  "Say what happened in the guard’s terms. “That’s it written up” beats “an Activity record has been created”.",
  "Acknowledge each success in a few words and move on. “Changed. Same report” is enough.",
  "Short lines, everyday words, no system jargon, and no instructions read aloud as a list.",
  "Sign off plainly and leave the guard knowing what they can now do. “That’s everything. You can report now, and you can put it right when I get it wrong.”",
].forEach((s) => kids.push(bullet(toneRef, s)));
kids.push(p([chip("P"), t("  Reference copy below shows the intended tone. Wording is settled with product before build, and no line here is a locked string.")], { after: 50 }));
kids.push(table([1500, 4074, 4074], ["Stage", "SAM says", "Guard does"], [
  ["Welcome", "“Hello. I’m SAM, your new assistant, and from today you talk and I write it up. Let’s get to know each other, it takes a few minutes.”", "Reads, then taps Start."],
  ["Getting to know each other", "“I’m trained to recognise certain words. If I keep missing something you say, tell your trainer the words you use and we’ll teach me. I’m yours to train.”", "Reads; taps Continue."],
  ["Practice context", "“Nothing here counts. I mark everything we write as training, so it stays out of your real reports. If you have to stop, I’ll remember where you were.”", "Taps Continue."],
  ["What is tracked", "“One thing before we start. Your location and my running in the background are for your reports and your rounds. Your supervisor sees the reports you file and the checkpoints you scan. They do not get a map of your day.”", "Reads; taps Continue."],
  ["Permissions", "“First, permissions. This link opens your phone’s settings for SAM OnSite. Tap Permissions and turn on all five: microphone so I can hear you, location so your reports say where you are, camera for photos, notifications so you get told things, and physical activity so your rounds count. Location asks twice — pick All the time on the second screen. Then, on the same page, turn off Remove permissions if app isn’t used, or your phone quietly switches these off again after a few quiet months. Come back and tap Done.”", "Taps Permissions, enables all five at the named levels, sets location to All the time, turns off the unused-app removal, returns, taps Done."],
  ["NFC", "“Now NFC. On sites with checkpoint tags, that’s what registers your rounds. Switch it on through here and tap Done when you’re back.”", "Enables NFC, returns, taps Done. Skipped if the phone has no NFC chip."],
  ["Background tracking", "“Last one. This one is in my own settings, not your phone’s, so the link looks different. Background tracking is what keeps me working when your screen goes dark. Switch it on, come back, and we’re set up.”", "Enables background tracking in the in-app setting, returns, taps Done."],
  ["Ask a question", "“Press Talk and try this one: how do I report an incident? Ask me anything you like later; for now this is the one I want to show you.”", "Presses Talk and asks. The first transcription doubles as the mic check. SAM answers with the prepared answer; nothing live results."],
  ["Report", "“Now report something to me, the way you’d tell a colleague. Try this one: a Dell laptop’s gone missing.” … “That’s it written up. Have a look.”", "Reports. One practice report is created and shown on screen."],
  ["Voice fix", "“Say the wrong thing and I’ll fix it. Try it now: tell me it was a MacBook, not a Dell.” … “Changed. Same report.”", "Says it was a MacBook. Same report updates."],
  ["Manual fix", "“You can fix it by hand too. Press and hold your report — the long one about the laptop, not the correction you just made — and add that it happened at reception.”", "Long-presses the report message (highlighted on screen), adds “at reception”, saves."],
  ["Finish", "“That’s everything. You can report now, and you can put it right when I get it wrong.”", "Taps Go to home once completion is confirmed."],
]));


/* 05b Where each tip lands */
kids.push(subhead("WHERE EACH TIP LANDS"));
kids.push(p([chip("P"), t("  The eight tips from the General recommendations slide, placed at the step where each is useful. Two are practised rather than told.")], { after: 50 }));
kids.push(table([2500, 2100, 5048], ["Tip (from the deck)", "Where it lands", "How it is taught"], [
  ["Wait for the mic to turn green before speaking", "First Talk press", "SAM says it as the guard reaches for Talk, which is the only moment it means anything."],
  ["SAM handles background noise, but struggles when several people talk at once", "First Talk press", "A short aside straight after the first successful transcription."],
  ["SAM is your new assistant, so get to know each other. It is trained on certain words; tell us the terms you use", "Welcome", "Part of the opening, and the invitation to send terms back is what makes it a two-way tool rather than a fixed one."],
  ["Use full sentences, as you would to a person. “I want to make a theft report”, not “theft report”", "Report step", "Taught by doing: the prompt asks for it the way you would tell a colleague, and a too-short utterance gets a nudge rather than a lecture."],
  ["Take photos whenever you want, even if SAM does not ask", "After the report is shown", "An aside at the moment the guard is looking at their finished report and can see where a photo would have gone."],
  ["Spelling mistakes, in names for instance, are fixed by pressing and holding the text to enter edit mode", "Manual-fix step", "Practised, not described. This tip and that step are the same action."],
  ["You can change scenario mid-conversation if SAM picks the wrong one", "After the voice correction", "An aside once the guard has seen a correction work, so the idea already has a hook."],
  ["You can reopen a conversation later to change or add information", "Close", "Part of the sign-off: what the guard can do after today."],
]));

/* 06 Functional requirements */
kids.push(eyebrow("06", "Functional requirements"));
kids.push(p([
  t("Each requirement is traced to acceptance criteria and tests in the Traceability matrix. Note that "),
  t("FR-21 (tone and engagement)", { bold: true }),
  t(" carries the same weight as any functional rule here. How this feels to the guard is part of the requirement."),
], { after: 50 }));
kids.push(table([850, 2100, 6698], ["ID", "Requirement", "Detail"],
  FR.map((r) => [r[0], [t(r[2], { size: 18, bold: true })], [chip(r[1]), t("  " + r[3], { size: 18 })]])));

/* 07 Progress model */
kids.push(eyebrow("07", "Progress and completion model"));
kids.push(p([chip("P"), t("  A recommended record shape. Field and key names to be confirmed in build. No completion field was visible in the supplied user form.")], { after: 50 }));
kids.push(table([2600, 2400, 4648], ["Field", "Example", "Purpose"], [
  ["user_id", "—", "The guard the record belongs to."],
  ["onboarding_key", "sam_onsite_guard", "Identifies this onboarding; leaves room for future ones."],
  ["version", "1", "Version issued; each guard gets version 1 once."],
  ["status", "not_started / in_progress / completed", "Server-side state; never trusted from the phone alone."],
  ["current_stage", "e.g. voice_fix", "Resume point after app close or device change."],
  ["practice_report_ref", "id of the run’s practice report", "Without it, resuming at voice_fix or manual_fix on another device has no report to correct. Required by the one-active-report rule (FR-18)."],
  ["practice_message_ref", "id of the guard’s report message", "The message the manual-edit step long-presses. Same reason: the resume point is meaningless without it."],
  ["started_at / updated_at / completed_at", "timestamps", "Lifecycle timing and idempotency support."],
]));

/* 08 Data isolation */
kids.push(eyebrow("08", "Practice data isolation and completion integrity"));
kids.push(subhead("PRACTICE DATA MUST NEVER REACH LIVE OPERATIONS"));
kids.push(p([
  t("Every practice report is flagged training (archive / soft-delete) the instant it is created. The platform already has this flag. "),
  t("Key dependency:", { bold: true }),
  t(" the filter that keeps marked data out of live surfaces is not built yet. Today, archived items are still visible in production, so it is a required enhancement. This onboarding therefore runs only in DEV/UAT until that exclusion filter is built and verified, and must not go live before then. When built, it should work hide-by-default at the shared data layer, so every current and future surface excludes training data automatically. Permanent deletion is a separate later clean-up; only the guard’s completion status is kept."),
]));
kids.push(callout([
  "Reports and event pages · dashboards and KPIs · analytics and metrics · exports and scheduled reports · alerts and notifications (including supervisor alerts) · webhooks and integrations · the Pronect Action Tracker · the event outbox and any queue feeding the above · read replicas, caches, and search indexes · the SAM conversation / transcript store · the guard’s own activity list and search. Real transcripts show activities send supervisor alerts, move KPIs, and create Action Tracker items today, and a dispatched alert cannot be recalled, so nothing may be emitted before the training flag is in place.",
], TEAL, "Every consumer the exclusion filter must cover"));
kids.push(subhead("COMPLETION IS SERVER-SIDE, EARNED, AND SAFE UNDER PRESSURE"));
kids.push(p("Completion is recorded only after the system observes each real action; it lives on the server (never trusted from the phone), is versioned (version 1), issued once per guard, resumes safely across app close or device change, and cannot be duplicated by two concurrent sessions. If practice data cannot be safely isolated at any point, completion is withheld and the guard stays where they are. Keeping live data clean outranks getting anyone finished.", { after: 40 }));

/* 09 Errors */
kids.push(eyebrow("09", "Error and recovery behaviour"));
kids.push(table([3000, 6648], ["Condition", "Behaviour"], [
  ["A setting is still off after the guard confirms", "Permissions are read back (FR-24), so a false confirmation is caught at the gate and the guard is sent back with the missing permission named. NFC and background tracking are still taken on trust, because it is not confirmed the app can read their state; a miss on either surfaces later as a feature that does not work. That is the residual risk here, and the reason to extend read-back to all three. Anything that genuinely cannot be enabled reaches a support screen so the guard is not trapped."],
  ["Network, speech, or save failure", "Preserve progress; show a retryable error; never silently lose the current step or an edit."],
  ["Voice correction misheard", "SAM re-asks; the guard retries by voice or uses the manual long-press edit."],
  ["Manual save failure", "Show the failure and keep the typed text for retry; the report is never left half-edited."],
  ["Isolation cannot be assured", "Fail closed: withhold completion, keep the guard in the exercise, record only a stage/error code, and never the practice content."],
  ["A real incident comes up", "The guard can leave onboarding, report it through normal reporting, and return to the saved stage; completion is neither granted nor skipped. Minor path; see the requirement note."],
]));

/* 10 Acceptance criteria */
kids.push(eyebrow("10", "Acceptance criteria"));
kids.push(table([850, 1350, 7448], ["ID", "Traces to", "Criterion"], AC.map((r) => [r[0], r[1], r[2]])));

/* 11 UAT */
kids.push(eyebrow("11", "Test matrix (DEV / UAT)"));
kids.push(p("Positive and negative tests. Negatives cover skip, a permission left off, a permission that cannot be granted, a phone with no NFC, repeated transcription failure, voice-correction failure, manual-save failure, isolation failure, concurrency, and network/device resume. Two rows are go-live rather than build checks: the live-surface sweep and the filter capability check.", { after: 50 }));
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
  "Universal for all guards, with no customer-specific setup, and it changes no existing templates.",
  "Monitor completion-failure and isolation-failure rates; alert on isolation failures. Do not log raw voice/report content in monitoring beyond existing policy.",
  "Nothing reaches production without Vincent Smeyers’ sign-off.",
].forEach((s) => kids.push(bullet(rollRef, s)));

/* 14 DoD */
kids.push(eyebrow("14", "Definition of done"));
const dodRef = "dod";
[
  "All acceptance criteria pass in UAT, including every negative test.",
  "A guard who completes it ends with a working phone: permissions read back as granted, and NFC and background tracking confirmed. Onboarding that leaves a setting off has not done its job.",
  "Version 1 issued once to a sample of new and existing guards; the refresher preserves completion.",
  "The exclusion filter is built and the practice report is proven absent from every consumer above; a forced isolation failure withholds completion.",
  "Completion is server-side, versioned, idempotent, and safe under concurrency and device change.",
  "The Confirm-with-engineering items are answered, and Vincent Smeyers has signed off.",
].forEach((s) => kids.push(bullet(dodRef, s)));

/* 15 Engineering validation */
kids.push(eyebrow("15", "Confirm with engineering before building"));
kids.push(p([chip("X"), t("  These depend on platform behaviour not provable from the supplied sources. Worth settling once, up front. None of them looks like a blocker.")], { after: 50 }));
const valRef = "val";
[
  "Does the app already record a guard’s first successful sign-in?",
  "Can this exercise be launched on its own, without the usual voice or tap triggers that start other scenarios?",
  "Can SAM run in a practice mode whose conversation and report never enter live reporting?",
  "What is the exact way the app updates an existing report, by voice and by a saved manual edit?",
  "Which live consumers must the new exclusion filter cover (at minimum supervisor alerts, KPIs, and the Pronect Action Tracker, all confirmed live today), and where are reports created so the training flag is set at creation? (The flag exists; the filter does not yet.)",
  "Which permissions and grant levels does the app need (e.g. location “all the time” vs “while using”), can it read each one’s current level to verify it, and can it switch off the phone’s automatic permission removal for unused apps? The app’s Location screen already shows a “Permissions needed” status and an “Open system settings” link, so state-reading and the settings link exist at least for location.",
  "Is there already a place to store “this guard has completed onboarding”, or must one be added?",
].forEach((s) => kids.push(bullet(valRef, s)));
kids.push(gap(10));
kids.push(callout([
  "What is proven today: the reporting console exposes edit and soft-delete/archive on SAM OnSite activities, and SAM can correct a report’s structured fields through the conversation. The archive flag exists and can be used now. Real app screens add more: activities have live consequences (supervisor alerts, KPI impact, items in the Pronect Action Tracker, exactly the surfaces the filter must cover); an activity is a report card plus transcript, matching the message-is-the-report model; permission state is readable and a settings link exists at least on the Location screen (“Permissions needed” / “Open system settings”); and a spoken question was recorded as a Procedure activity, so even the ask step may create a record today. What is not yet built: the filtering that actually excludes archived/training data from live surfaces. It is currently still visible in production and is a later-stage enhancement. No production change has been made.",
], TEALD, "Evidence basis"));

kids.push(new Paragraph({ spacing: { before: 90 }, border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 5, space: 5 } },
  children: [t("Describes intended behaviour for internal build and testing only. No production change is made until Vincent Smeyers has signed off.", { size: 16, it: true, color: GREY })] }));

/* ================= NUMBERING ================= */
const numbering = { config: [] };
[decRef, inRef, outRef, laterRef, rollRef, dodRef, valRef, toneRef].forEach((reference) => numbering.config.push({
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
