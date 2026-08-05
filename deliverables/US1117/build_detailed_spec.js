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
  ["FR-01", "C", "Trigger", "Onboarding launches automatically at sign-in for any active guard who has not completed it, before the home screen is reachable. No detection of a “first” sign-in is needed; not-completed is the condition, and it covers guards whose accounts predate this feature."],
  ["FR-02", "C", "Audience", "Every active guard receives it exactly once, including guards whose accounts existed before this feature."],
  ["FR-03", "C", "Mandatory", "There is no path to the home screen except by completing it, other than four defined exits: the real-incident exit (FR-17), the stuck-guard release (FR-28), the no-connectivity pass-through (FR-28), and leaving a refresher when already complete (FR-04). None of them marks the guard complete."],
  ["FR-04", "C", "Refresher and re-entry", "A completed guard can reopen the training from a named entry point on the home screen, and reopening never clears or regresses completion. A refresher starts a fresh practice run with its own practice report; the one-active-report rule is per run. The same entry point serves a guard released under FR-28 and a guard let through under the no-connectivity rule. Without a specified entry point there is no route back for either."],
  ["FR-05", "C", "Language", "All onboarding voice and screen copy is English for the MVP."],
  ["FR-06", "C", "Location gate", "Only location is gated here. The microphone and the camera prompt themselves at the moment they are first used, so onboarding does not chase them; it just warns the guard a prompt is coming and to allow it. Location is different: the app requests it, but not at the “allow all the time” level, and the app cannot read that level back. So onboarding walks the guard to the second Android screen and names the setting to pick. SAM opens the phone’s settings page for SAM OnSite through a link held inside onboarding, and the guard returns and confirms with a tap. After a restart the flow resumes where it left off, but the guard is told that before they leave, or they will not know to come back. A guard who says location cannot be granted reaches a support screen."],
  ["FR-07", "P", "Transcription check", "A speaking step is credited only after at least one successful transcription; an empty or failed one does not advance it. The first Talk press doubles as the microphone check. Steps are judged on the guard demonstrating the intended action and on the resulting report state; semantically-equivalent English is accepted and an exact phrase is never required. A prompt to speak more fully is not a failure and does not count towards the retry limit in FR-28."],
  ["FR-08", "P", "Ask a question", "The guard asks a spoken question and SAM answers. Nothing live may result. Confirmed on the real app: asking a question writes a row into the activity table, typed as an “other question”. The practice context must suppress that row entirely; flagging it is not sufficient while the exclusion filter does not exist. The answer to the suggested question is written and reviewed in advance. If the guard asks something else, the step is still credited on a successful transcription, and SAM answers if it can or says plainly that it cannot."],
  ["FR-09", "P", "One practice report", "Reporting the fictional incident creates exactly one practice report. Retries or repeated messages must not create additional reports. SAM asks at most two follow-up questions and “I don’t know” always advances, since the guard is inventing details about a laptop that does not exist."],
  ["FR-10", "C", "Voice correction", "A spoken correction updates the same report in place. Confirmed on the real app: the report updates and no second report is created. SAM’s prompt must not assume the guard used any particular word for the laptop, since they were invited to report in their own words."],
  ["FR-11", "C", "Manual correction", "Long-pressing the guard’s report message opens a manual edit that, when saved, updates the same report. Confirmed on the real app. The target message must be identified on screen and not by description alone, since by this point the guard has sent several messages. This is the only step that requires typing, so it must accept a short addition without demanding the whole sentence be retyped."],
  ["FR-12", "C", "Marked at creation", "Every practice report carries the training flag from creation. If it cannot be created already marked, it is not created at all. The card the guard sees carries a visible “Training” label. Whether the flag can be set atomically at creation is an engineering confirmation."],
  ["FR-13", "C", "Kept out of live surfaces", "The practice report must appear in no live surface (see the Data isolation section). The platform has the training flag, but the filter that excludes marked data is not built. Data is currently visible in production, so this exclusion filter is a required enhancement; production is blocked until it is built and verified, and the exercise runs only in DEV/UAT until then."],
  ["FR-14", "P", "Completion integrity", "Completion is recorded server-side and versioned, issued once per guard, and set only after each required action is observed. A quiz or an “I’m done” button does not count. If the server cannot be reached at the finish, completion is recorded locally and synced; the operation is idempotent (FR-18), so a guard who did the work is never made to repeat it."],
  ["FR-15", "P", "Fail-closed", "If safe isolation of practice data cannot be assured at any step, the step fails and completion is not marked; the guard stays in the exercise, subject to the release in FR-28."],
  ["FR-16", "C", "Hand-off", "On confirmed completion the guard reaches the normal home screen (Talk / Capture / Type). The closing screen names the settings now switched on, so a guard who missed one has a last chance to catch it before it becomes a problem on shift."],
  ["FR-17", "P", "Way out for a real incident", "A way to leave onboarding, report a real incident, and return to the saved stage, with completion neither granted nor skipped. Reachable from every step, including the setup gates and the refresher. Crossing it must be unmistakable: the practice banner clears and the guard is told this one is live. Before the microphone permission exists the exit must open typed reporting or tell the guard to call it in, or it is unusable at exactly the point it is most likely to be needed."],
  ["FR-18", "P", "Idempotency / concurrency", "One onboarding record per guard, and one active practice-report reference per practice run. Retries do not duplicate; the newest valid server-side progress wins with no regression; the completion operation is idempotent. One carve-out: if the build implements FR-26 by moving the stage back to the permissions gate, that is deliberate and must not be blocked by the no-regression rule."],
  ["FR-19", "P", "Privacy / telemetry", "Telemetry records only stage and error codes and completion metrics. It does not log raw voice or transcript content beyond existing policy, and never routes fictional incident content into operational analytics."],
  ["FR-20", "C", "Deferred visual support", "Tooltips and highlight animations for dark or noisy sites are deferred; the need is judged during testing. This does not cover the on-screen identification of the edit target in FR-11, which is required for the MVP."],
  ["FR-21", "C", "Tone and engagement", "The experience must feel like being handed something useful, not like sitting a test. SAM introduces itself as the guard’s new assistant, invites them to get to know each other, frames each step as a benefit, and closes warmly. A dry “do this, then that” sequence does not meet this requirement."],
  ["FR-22", "C", "NFC gate", "After permissions, SAM links the guard to the NFC setting; the guard switches it on, returns and confirms. NFC is what reads the checkpoint tags. Not every site uses tags, but having it on costs nothing and avoids a failure later, so it is part of setup for everyone. The copy must not make it conditional on the guard’s site, or a guard on a tag-less site is handed a reason to confirm without acting."],
  ["FR-23", "C", "Background-tracking gate", "The third gate: background tracking, which the reviewer confirms is part of SAM OnSite’s own app settings screen rather than the phone’s. Same shape as the others: SAM explains why it is needed, links straight to that one setting from inside onboarding, and the guard switches it on and returns to confirm."],
  ["FR-24", "C", "Check the settings where the app can", "Engineering has confirmed the app can detect a permission that should be on and is off, with three exceptions: location at “allow all the time”, background activity, and physical activity. So the gates split. Everything checkable is checked, and a gate that reads back as off returns the guard to it instead of advancing. The three exceptions are guided, taken on the guard’s word, and re-checked next session (FR-32). Whether NFC state is readable is still open."],
  ["FR-25", "C", "Tips taught in context", "All eight tips from the deck’s General recommendations slide are carried by SAM’s spoken lines at the step where each becomes useful (see Where each tip lands). Two are taught by doing: the guard performs the action and SAM’s next line names what just happened. A slide of tips listed on a screen or read out in sequence does not meet this requirement."],
  ["FR-26", "C", "Microphone checked, not inferred", "Microphone state is one the app can read (FR-24), so the gate checks it directly instead of waiting for Talk to fail. A guard who confirms without granting is returned to the gate at once. If Talk still produces nothing after that, the remaining likely cause is that the app has not been restarted since the grant, and the guard is told so."],
  ["FR-27", "X", "No NFC hardware", "If the device has no NFC chip, the gate is skipped and the flow continues. A guard is never sent to support over hardware their phone does not have. Whether the app can detect the chip’s presence is an engineering confirmation, as with FR-24."],
  ["FR-28", "P", "Retry floor and release", "After three failed attempts at a spoken step, typing is offered. On any step, a guard who is stuck can say so and reach the home screen with completion unmarked and the stage saved; the release must not depend on a failure counter, since a guard who cannot find the message to long-press never generates a failed attempt. If the server cannot be reached at sign-in, the guard passes through rather than being held. A released guard must not be dropped back onto the same step at every subsequent sign-in: the release carries a cool-down or routes them to a trainer, and re-entry is through FR-04."],
  ["FR-29", "P", "A real incident reported inside practice", "A practice banner stays visible for the whole exercise and not only on the finished report card, and it carries the way out to normal reporting (FR-17). A genuine incident reported during practice would be flagged training and hidden from the supervisors it should have reached. Detecting that what the guard said is not the practice incident is the better answer and is outside MVP scope. The banner cannot be shown while the guard is inside the phone’s settings, so the claim is that it is visible at every step of the flow itself."],
  ["FR-30", "P", "Go-live gate on the filter", "The MVP is built and tested in DEV/UAT only. Before the switch is enabled in production the exclusion filter must be present and verified, enforced by something firmer than memory. The recommended enforcement is a server-side capability check that refuses to start onboarding when the filter is absent, built at go-live, outside the MVP."],
  ["FR-31", "X", "Say what is tracked", "Guards are asked for location and background tracking on a work phone and will ask what is done with it, so one plain line belongs in the flow. The wording must be supplied by Pronect and, where required, agreed with the works council. This specification does not state what a supervisor can or cannot see; that boundary is not ours to draw."],
  ["FR-32", "C", "Check the three settings at the start of a session", "Guards share phones, each with their own login, so the phone being ready cannot be recorded against the guard. Rather than tracking which handset is set up, the three settings are simply checked again when a guard starts a session, whatever phone they are holding, and the guard is walked through anything that is off. Repeating a check that passes costs seconds and needs no per-device record. This is what makes the promise hold on a shared handset."],
];
const AC = [
  ["AC-01", "FR-01 / FR-02", "Given an active guard who has not completed onboarding, when they sign in, then it launches before the home screen is reachable; a guard whose account predates this feature receives it once, and after completion a later sign-in goes straight to home."],
  ["AC-02", "FR-04", "Given a completed guard, or a guard released under FR-28, when they open the named entry point on the home screen, then the training reopens at the saved stage without clearing or regressing completion."],
  ["AC-03", "FR-03", "Given onboarding in progress, when the guard tries to skip, then there is no path to home except completion or one of the four defined exits, none of which marks them complete."],
  ["AC-04", "FR-06", "Given the permissions gate, when the guard follows the link, enables the permissions and returns, then a tap confirms and the flow continues; the link opened the phone’s settings page for SAM OnSite from inside onboarding. Confirmation works before microphone permission is granted. The required grant level is named on screen for each permission. If the guard declares a permission cannot be granted, they reach a support screen and are not trapped."],
  ["AC-05", "FR-07", "Given a speaking step, when transcription fails or returns empty, then the step is not credited and the guard is asked to try again; when the guard says the right thing in different words, the step is credited; a prompt to speak more fully does not count as a failed attempt."],
  ["AC-06", "FR-08 / FR-09", "Given the guard asks a question, then SAM answers and no activity record of any kind is written for that step, verified in the activity table, not by inspecting a flag; the step is credited even if the guard asks their own question. Given the guard then reports the incident, then exactly one practice report exists, after at most two follow-ups, with “I don’t know” advancing."],
  ["AC-07", "FR-11", "Given SAM has identified the target message on screen, when the guard long-presses it and saves it with the location added, then the same report updates and the added text is present."],
  ["AC-08", "FR-10", "Given a practice report the guard has already edited by hand, when they say the item was a MacBook, then the same report updates, no second report is created, and the final report contains both the corrected item and the added location."],
  ["AC-09", "FR-12 / FR-13", "Given a practice report at any stage, then it is flagged training at creation and the card carries a visible “Training” label; once the exclusion filter is built it appears in no live surface; until the filter exists the exercise runs only in DEV/UAT."],
  ["AC-10", "FR-14 / FR-15", "Given every required action observed and isolation assured, when completion runs, then it is recorded server-side, or locally and synced if the server is unreachable; if isolation cannot be assured, completion is withheld."],
  ["AC-11", "FR-16", "Given completion, then the guard reaches the home screen and the closing screen names the settings now switched on."],
  ["AC-12", "FR-18", "Given duplicate concurrent sessions or retried messages within one run, then no duplicate record or report is created and progress does not regress, except a deliberate move back to the permissions gate under FR-26."],
  ["AC-13", "FR-17", "Given a guard leaves onboarding to report a real incident, then the practice banner clears, they are told this one is live, reporting opens, and on return the saved stage resumes with completion neither granted nor skipped. Given the exit is crossed before microphone permission exists, then a typed route or a call-it-in instruction is offered."],
  ["AC-14", "FR-05", "Given any onboarding screen or prompt, then all voice and screen copy is in English."],
  ["AC-15", "FR-19", "Given a run, when telemetry and analytics are inspected, then only stage and error codes and completion metrics are recorded, with no raw voice or transcript content and no fictional incident content in operational analytics."],
  ["AC-16", "FR-21 / FR-25", "Given the full run-through, then it opens with SAM introducing itself as the guard’s new assistant, each step is framed as a benefit, the close is warm, and each of the eight tips is carried by a spoken line at the step named in Where each tip lands, with none read out as a list. Signed off by product on a read-through, and at least 80% of a test group of five or more guards describing it as welcoming and not test-like."],
  ["AC-17", "FR-22 / FR-27", "Given the NFC gate, when the guard follows the link, switches NFC on and returns, then they confirm and the flow continues; given a device with no NFC chip, the gate is skipped and no support screen is shown."],
  ["AC-18", "FR-23", "Given the background-tracking gate, when the guard follows the link to the in-app setting, switches it on and returns, then they confirm and the flow advances to the practice steps. What is observed is the confirmation, not the setting: no gate state is read in the MVP."],
  ["AC-19", "FR-26", "Given a guard who confirmed the permissions gate without granting the microphone, or who granted it without restarting the app, when Talk fails repeatedly, then they are pointed back at the permissions gate with both causes named."],
  ["AC-20", "FR-28", "Given three failed attempts at a spoken step, then typing is offered; given a guard stuck on any step who asks to leave, then they reach the home screen with completion unmarked and the stage saved, and are not dropped back onto that step at the next sign-in; given no connectivity at sign-in, then the guard passes through."],
  ["AC-21", "FR-29", "Given a practice run, then a practice banner is visible at every step of the flow and carries the way out to normal reporting."],
  ["AC-22", "FR-30", "Given production go-live, then the exclusion filter is present and verified before the switch is enabled, and its absence blocks onboarding from starting."],
  ["AC-23", "FR-31", "Given the setup gates, then a line supplied by Pronect states what location and background tracking are used for. No statement about what a supervisor can or cannot see originates in this specification."],
  ["AC-24", "FR-32", "Given a guard who completed the gates on one device and resumes or signs in on another, then the three gates are replayed on the new device before the flow continues; given an already-completed guard on an unrecognised device, then the gates run alone and completion is unaffected."],
];
const UAT = [
  ["UAT-01", "AC-01", "New guard signs in; separately, a pre-feature guard signs in.", "Onboarding shown before home for both; progress record created; issued once."],
  ["UAT-02", "AC-02", "Reopen from the home-screen entry point, as a completed guard and as a released guard.", "Reopens at the saved stage; completion preserved."],
  ["UAT-03", "AC-03", "Attempt to skip onboarding. (negative)", "No skip path to home; only the four defined exits, none marking completion."],
  ["UAT-04", "AC-04", "Complete the permissions gate on a fresh install, following the link and returning.", "Link opens the phone’s settings page for SAM OnSite; grant level named for each permission; confirmation advances the flow."],
  ["UAT-05", "AC-04", "Declare that a permission cannot be granted. (negative)", "Support screen reached from the guard’s own declaration; guard not trapped."],
  ["UAT-06", "AC-04 / AC-19", "Grant permissions on a device where the app restarts on grant. (negative)", "Guard returns to the same gate, is told permissions are kept, and confirms without redoing them."],
  ["UAT-07", "AC-05", "Force a failed transcription, then repeat the step in different wording.", "Not credited on failure; credited on a differently-worded equivalent; a fullness prompt does not count as a failure."],
  ["UAT-08", "AC-06", "Ask the suggested question, then inspect the activity table; separately, ask an unrelated question.", "SAM answers; no activity row written; step credited in both cases."],
  ["UAT-09", "AC-06", "Report the fictional incident, answering “I don’t know” to every follow-up.", "At most two follow-ups; exactly one practice report; step advances."],
  ["UAT-10", "AC-07", "Manual long-press edit and save, with the target message identified on screen.", "Same report updated; added location present."],
  ["UAT-11", "AC-07", "Manual save fails, then retry. (negative)", "Error shown; typed text retained; report not half-edited."],
  ["UAT-12", "AC-08", "Voice-correct the item after the manual edit.", "Same report updated; no duplicate; final report holds both the corrected item and the location."],
  ["UAT-13", "AC-08", "Voice correction misheard, then retry. (negative)", "SAM re-asks; still exactly one report."],
  ["UAT-14", "AC-09", "Inspect the practice report card; after the filter is built, inspect every live surface.", "Visible “Training” label; absent from all listed consumers; before the filter, run confined to DEV/UAT."],
  ["UAT-15", "AC-09 / AC-10", "Force create-with-flag / isolation failure. (negative)", "Report not created; completion withheld (fail-closed)."],
  ["UAT-16", "AC-10 / AC-11", "Complete all actions with isolation assured; separately, complete with the server unreachable.", "Recorded server-side; recorded locally and synced when offline; closing screen names the settings switched on."],
  ["UAT-17", "AC-12", "Duplicate concurrent sessions and retried messages within one run. (negative)", "One record, one report for the run; no progress regression; a deliberate FR-26 move back is not blocked."],
  ["UAT-18", "AC-24", "Complete the gates on device A, then resume on device B. (negative)", "Gates replayed on device B before the flow continues; the guard is not carried past setup on an unconfigured phone."],
  ["UAT-19", "AC-24", "Already-completed guard signs in on an unrecognised device.", "Gates run alone; completion unaffected."],
  ["UAT-20", "AC-13", "Leave to report a real incident, then return; separately, cross the exit before the microphone is granted.", "Banner cleared and the live warning shown; typed or call-it-in route offered pre-microphone; saved stage resumes."],
  ["UAT-21", "AC-14", "Review all onboarding copy.", "All copy is English."],
  ["UAT-22", "AC-15", "Inspect telemetry and analytics during a run. (negative)", "Only stage and error codes plus completion metrics; no raw content; no fictional data in analytics."],
  ["UAT-23", "AC-16", "Full run-through with a test group of five or more guards, reviewing tone and tip placement.", "SAM introduces itself as the guard’s assistant; each of the eight tips lands at its named step; none read as a list; at least 80% describe it as welcoming."],
  ["UAT-24", "AC-17", "NFC gate on a device with a chip; separately, on a device without one. (negative)", "Link opens the NFC setting and confirmation advances; the gate is skipped on hardware without a chip and no support screen appears."],
  ["UAT-25", "AC-18", "Background-tracking gate: follow the link to the in-app setting, switch it on, return and confirm.", "Confirmation advances to the practice steps. No gate state is read; the test observes the confirmation only."],
  ["UAT-26", "AC-04 / AC-17 / AC-18", "Confirm any of the three gates with the setting still switched off. (negative)", "Accepted for the MVP; no gate is checked. Microphone case is covered by UAT-27."],
  ["UAT-27", "AC-19", "Press Talk repeatedly having confirmed permissions without granting the microphone. (negative)", "Pointed back at the permissions gate with both causes named; no endless identical retry."],
  ["UAT-28", "AC-20", "Fail a spoken step three times; separately, ask to leave from a non-spoken step; separately, sign in with no connectivity. (negative)", "Typing offered; release reaches home with completion unmarked and stage saved, and does not recur at the next sign-in; no connectivity does not hold the guard."],
  ["UAT-29", "AC-21", "Walk every step of the flow checking the practice banner.", "Banner visible at every in-flow step and offers the way out."],
  ["UAT-30", "AC-22", "At go-live, launch onboarding with the exclusion filter absent. (negative)", "Refuses to start and reports why. Run at go-live, not during MVP build."],
  ["UAT-31", "AC-23", "Read the setup copy as a first-time guard.", "A Pronect-supplied line states what location and background tracking are used for; no supervisor-visibility claim originates here."],
];

/* ================= CONTENT ================= */
const kids = [];

/* cover */
kids.push(new Paragraph({ spacing: { after: 16 }, children: [t("PRONECT   ·   SAM ONSITE", { size: 16, bold: true, color: TEAL, cs: 40 })] }));
kids.push(new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: "Guard Onboarding", font: SERIF, size: 46, bold: true, color: INK })] }));
kids.push(new Paragraph({ spacing: { after: 140 },
  border: { bottom: { color: TEAL, style: BorderStyle.SINGLE, size: 4, space: 10 } },
  children: [new TextRun({ text: "Detailed specification — engineering and QA companion", font: SERIF, size: 21, italics: true, color: GREY })] }));
kids.push(new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
  rows: [new TableRow({ children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX }, margins: { top: 100, bottom: 100, left: 180, right: 180 },
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
    children: [new Paragraph({ spacing: { after: 0 }, children: [
      t("Product  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("SAM OnSite       ", { size: 18 }),
      t("Audience  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Engineering + QA       ", { size: 18 }),
      t("Status  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Draft for build       ", { size: 18 }),
      t("Sign-off  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Vincent Smeyers", { size: 18 }),
    ] })] })] })] }));
kids.push(gap(40));

/* 01 Purpose and scope */
kids.push(eyebrow("01", "Purpose and scope"));
kids.push(p("A mandatory, one-time voice onboarding at a guard’s first sign-in. Two outputs, both required at the end: the guard can ask a question, report an incident and correct a report, practised on a made-up incident; and permissions, NFC and background tracking are all on.", { after: 70 }));
kids.push(subhead("IN SCOPE"));
const inRef = "insc";
[
  "The onboarding flow itself, launched at sign-in and reopenable, replacing the deck’s app-setup slides: background tracking opt-in, granting permissions, and enabling NFC, with their screenshot walkthroughs. It also absorbs the General recommendations tips.",
  "Device setup inside the flow: location, NFC and background tracking, each reached by a link duplicated into onboarding, and re-checked at the start of every session because guards share phones.",
  "One guided pass through ask, report, correct by hand, correct by voice.",
  "Server-side completion tracking, issued once per guard, with gate state held per device.",
  "Isolation of practice data from every live surface, dependent on the exclusion filter.",
].forEach((s) => kids.push(bullet(inRef, s)));
kids.push(subhead("OUT OF SCOPE"));
const outRef = "outsc";
[
  "Installation and first log-in. These stay with the trainer.",
  "Languages other than English; tooltips and highlight animations.",
  "Customer- or site-specific variants; any change to a live or protected system template.",
  "Physical deletion of practice data, a later task from the logical archive used here.",
].forEach((s) => kids.push(bullet(outRef, s)));
kids.push(subhead("DELIBERATELY LEFT OUT OF THE MVP"));
const laterRef = "later";
[
  "Detecting that what the guard reported is not the practice incident (part of FR-29). Needs content comparison and a threshold nobody has set; the banner and the way out cover the same risk cheaply.",
  "Checking the settings rather than asking (FR-24). Left out only because nobody has established whether the app can read that state.",
  "The capability check that refuses to launch without the exclusion filter (FR-30). Required at go-live, not in a build that never leaves DEV/UAT.",
].forEach((s) => kids.push(bullet(laterRef, s)));

/* 02 End-to-end flow */
kids.push(eyebrow("02", "End-to-end flow"));
kids.push(p("Roughly 3–5 minutes. In SAM, a guard’s message is what creates and updates their report, so a report is fixed either by voice or by editing the message. Copy for every stage is in Tone and script; gate behaviour is at FR-06, FR-22 and FR-23.", { after: 60 }));
const flowRef = "flow";
[
  "The guard signs in. If onboarding is not complete it opens automatically, before the home screen.",
  "SAM introduces itself, says it can be taught the words the guard uses, and sets the ground rules. A practice banner appears and stays for the whole exercise.",
  "Three setup gates in order: location, NFC, background tracking. Each links out to one setting and the guard returns to confirm. Granting may close the app; the guard reopens it and resumes at the same gate. The microphone and camera are not gated — they prompt themselves when first used, and SAM warns the guard to expect them.",
  "The guard presses Talk, waits for the microphone to turn green, and asks the suggested question. That first press doubles as the microphone check. Nothing live is written.",
  "The guard reports the suggested incident in their own words. SAM asks at most two follow-ups, creates exactly one report, and shows it as a card labelled “Training”.",
  "The guard corrects the report by hand first: they long-press their own report message, which SAM identifies on screen, add the location, and save.",
  "Then by voice: the guard says what the item actually was, and the same report updates. Voice runs last because the report always takes the state of the most recent message, even when an earlier one was edited afterwards. Confirmed on the real app.",
  "SAM signs off, names the settings now switched on, and a Go-to-home button lands the guard on the normal home screen. Completion is marked only after each step is observed.",
].forEach((s) => kids.push(numitem(flowRef, s)));

/* 03 Tone and script */
kids.push(eyebrow("03", "Tone and script"));
kids.push(p([chip("P"), t("  Reference copy: the intended tone and the tip placement. Wording is settled with product before build; no line here is a locked string.")], { after: 50 }));
kids.push(table([1400, 4624, 3624], ["Stage", "SAM says", "Guard does"], [
  ["Welcome", "“Hello. I am SAM, your new assistant. From today you talk, and I write the report for you. We will work together on every shift, so it helps if we get to know each other a little. I already know many of the words guards use. If I keep getting a word wrong, tell your trainer which word you use, and we can teach me.”", "Reads, taps Start."],
  ["What this is", "“Nothing here is real. Everything we make together now is marked as training, so it stays out of your real reports. If you have to stop, I remember where you were. If something real happens while we practise, use the button at the top and report it for real.”", "Reads the practice banner, taps Continue."],
  ["What is tracked", "Copy supplied by Pronect. States what location and background tracking are used for. No claim about what a supervisor can or cannot see originates here.", "Reads, taps Continue."],
  ["Gate 1 — location", "“First, location, so your reports say where you are. This link opens the settings for SAM OnSite. Your phone will ask twice. On the second screen choose Allow all the time, or I stop knowing where you are the moment your screen goes dark. If this closes me, just open SAM OnSite again and you will be right back here.”", "Follows the link, sets location to Allow all the time, returns."],
  ["Back from settings", "“You are back at the same step, nothing lost. Tap Done if it is set. If you cannot set it at all, tell me and I will get you help.”", "Taps Done, or says it cannot be granted. Same line whether they returned directly or reopened the app."],
  ["Gate 2 — NFC", "“Next is NFC. That is how your phone reads the checkpoint tags on your round. Switch it on through here, come back, and tap Done.”", "Enables NFC, returns, taps Done. Skipped if the phone has no NFC chip."],
  ["Gate 3 — background tracking", "“Last one, and this setting is mine rather than your phone’s. Background tracking keeps me working when your screen goes dark, so your round keeps counting with the phone in your pocket. Switch it on and come back. After this, when I ask to use your microphone or camera, just say yes — I only ask when I need them.”", "Enables background tracking in the in-app setting, returns, taps Done."],
  ["Ask a question", "“Now we talk. Press Talk and wait until the microphone turns green. Green means I am listening. If you speak before that, I lose your first words. Then ask me: how do I report an incident?”", "Presses Talk, waits for green, asks. First transcription doubles as the mic check."],
  ["After the answer", "“That is your microphone working well. One thing about my listening: noise around you is fine, I can work with that. Two people talking at the same time is hard for me. If a colleague is speaking next to you, wait a moment, then speak.”", "Listens. Nothing live is written."],
  ["Practice report", "“Now report something to me. Say the whole thing, the way you would say it to a colleague. Here is the practice one: a Dell laptop is missing.”", "Speaks a full sentence in their own words. At most two follow-ups; “I don’t know” advances."],
  ["Your report on screen", "“That is it written up. Look at the card. It says Training, so you can see it is not real. What helped me is that you gave me a full sentence. ‘I want to make a theft report’ tells me what you want to do. ‘Theft report’ on its own could be a procedure, a report, something else. One more thing while you have it open: you can add a photo whenever you want, also when I do not ask for one.”", "Looks at the report card, taps Continue."],
  ["Fix it by hand", "“You can fix a report with your hands. Press and hold the message I have marked for you, the one you sent me first. It turns into edit mode. Add that it happened at reception, and save.”", "Long-presses the marked report message, adds “at reception”, saves."],
  ["After the manual fix", "“The location is in the report now. Use that same press and hold for a name I spelled wrong. Hold the text, change the letters, save.”", "Sees the updated card."],
  ["Fix it by voice", "“You can also just tell me. Whatever you called it, tell me now it was a MacBook.”", "Says it was a MacBook. The same report updates."],
  ["After the voice fix", "“Changed. It works the same way when I pick the wrong kind of report. You do not start again. You tell me in the same conversation and I change it.”", "Sees the final report holding both changes."],
  ["Finish", "“That is everything. You can report now, and you can correct me when I get it wrong. Your phone is set up too: permissions, NFC and background tracking are all on. If one of them did not stick, tap here and we will do it again. Tomorrow, when you start, say ‘I’m starting my shift’. A full sentence, not ‘shift start’, that is too short for me. And nothing is closed for good: you can open a conversation again later, change it, or add what you find out afterwards.”", "Taps Go to home once completion is confirmed."],
]));

kids.push(subhead("WHERE EACH TIP LANDS"));
kids.push(p("The eight tips from the deck’s General recommendations slide. Two are practised, not told: the guard performs the action and SAM’s next line names what just happened.", { after: 50 }));
kids.push(table([4200, 2000, 3448], ["Tip", "Stage", "How"], [
  ["Wait for the mic to turn green", "Ask a question", "Said as the guard reaches for Talk, the only moment it means anything."],
  ["Copes with noise, not with several voices at once", "After the answer", "An aside once the guard has heard SAM get it right."],
  ["Your new assistant; it is trained on certain words, tell us the terms you use", "Welcome", "Shapes the opening, and the invitation to send words back makes it a tool the guard can shape."],
  ["Use full sentences, with the deck’s own examples", "Practice report, then the report card, then Finish", "Practised. The guard speaks a full sentence, then SAM names why it worked. The second example lands at the close."],
  ["Take photos even when SAM does not ask", "Your report on screen", "An aside while the guard is looking at the finished report."],
  ["Fix spelling by pressing and holding the text", "Fix it by hand, then after", "Practised. The guard uses the gesture to add a location, then SAM names it as the way to fix a name."],
  ["Change scenario in the same conversation", "After the voice fix", "An aside once a correction has visibly worked."],
  ["Reopen a conversation to change or add", "Finish", "Part of the sign-off: what the guard can do after today."],
]));

/* 04 Functional requirements */
kids.push(eyebrow("04", "Functional requirements"));
kids.push(table([850, 2100, 6698], ["ID", "Requirement", "Detail"],
  FR.map((r) => [cell(r[0], { w: 850, bold: true, size: 17 }), cell(r[2], { w: 2100, size: 18, bold: true }), dcell(r[1], r[3], 6698)])));

/* 05 Progress and completion */
kids.push(eyebrow("05", "Progress and completion model"));
kids.push(p([chip("P"), t("  A recommended record shape. Field and key names to be confirmed in build.")], { after: 50 }));
kids.push(table([2600, 2400, 4648], ["Field", "Example", "Purpose"], [
  ["user_id", "—", "The guard the record belongs to."],
  ["onboarding_key", "sam_onsite_guard", "Identifies this onboarding; leaves room for future ones."],
  ["version", "1", "Version issued; each guard gets version 1 once."],
  ["status", "not_started / in_progress / completed", "Server-side state; never trusted from the phone alone."],
  ["current_stage", "e.g. manual_fix", "Resume point after app close or device change."],
  ["practice_report_ref", "id of the run’s practice report", "Resume target for the correction steps."],
  ["practice_message_ref", "id of the message that created the report", "The message the manual-edit step long-presses. Must be the report-creating message, not a follow-up answer."],
  ["device_gate_state", "per device: permissions / NFC / background confirmed", "Gates are per device, not per guard (FR-32). Absent for an unrecognised handset, which replays the gates."],
  ["started_at / updated_at / completed_at", "timestamps", "Lifecycle timing and idempotency support."],
]));

/* 06 Data isolation */
kids.push(eyebrow("06", "Practice data isolation"));
kids.push(callout([
  [t("Every practice report carries the training flag from creation. "), t("The dependency:", { bold: true }),
   t(" the filter that keeps marked data out of live surfaces is not built, and archived items are still visible in production. It must be built and verified before go-live; until then this runs only in DEV/UAT. The product owner has named where it belongs: the reporting layer must be updated to exclude this data. Permanent deletion is a separate database clean-up at a later stage.")],
], TEAL, "PRACTICE DATA MUST NEVER REACH LIVE OPERATIONS"));
kids.push(p("Surfaces the filter must cover. The first four are confirmed from real transcripts and app screens; the rest are the likely set and must be confirmed with engineering.", { before: 60, after: 50 }));
const consRef = "cons";
[
  "Confirmed: supervisor alerts; KPI and dashboard figures; items in the Pronect Action Tracker; the guard’s own activity list and search.",
  "To confirm: reports and exports; automated integrations and webhooks; any queue or outbox feeding the above; read replicas, caches and search indexes; the conversation and transcript store.",
].forEach((s) => kids.push(bullet(consRef, s)));

/* 07 Error and recovery */
kids.push(eyebrow("07", "Error and recovery behaviour"));
kids.push(table([2600, 7048], ["Situation", "Behaviour"], [
  ["A setting is still off after the guard confirms", "The app can detect a permission that is off, with three exceptions it cannot read: location at “allow all the time”, background activity, and physical activity. Those three are guided, then taken on the guard’s word, and re-checked at the start of the next session (FR-32), which is what stops a miss lasting. Location granted only “while using” is the one that bites, because it silently defeats background tracking with no visible failure. Anything the guard declares ungrantable reaches a support screen."],
  ["Granting permissions closes the app", "Confirmed: the flow resumes where it left off. The guard must be told this before they leave for the settings, or they will not know to come back. They reopen SAM OnSite and land at the same gate, told their permissions are kept."],
  ["Network, speech or save fails", "Progress is kept and a retryable error is shown; the current step and any edit are never silently lost. At the finish, completion is recorded locally and synced."],
  ["A spoken step keeps failing", "After three failed attempts typing is offered. A prompt to speak more fully is not a failed attempt."],
  ["The guard is stuck on a step with no failures to count", "They can say so and reach the home screen with completion unmarked and the stage saved, re-entering later through the home-screen entry point, and not dropped back onto the step at every sign-in."],
  ["Practice data cannot be safely hidden", "The step stops, completion is withheld, and only a stage or error code is recorded, never the practice content."],
  ["A real incident comes up", "The way out is on the banner at every step. The banner clears, the guard is told this one is live, and normal reporting opens. Before the microphone exists, a typed route or a call-it-in instruction is offered."],
]));

/* 08 Acceptance criteria */
kids.push(eyebrow("08", "Acceptance criteria"));
kids.push(table([850, 1500, 7298], ["ID", "Traces to", "Criterion"],
  AC.map((r) => [cell(r[0], { w: 850, bold: true, size: 17 }), cell(r[1], { w: 1500, size: 17 }), cell(r[2], { w: 7298, size: 18 })])));

/* 09 Test matrix */
kids.push(eyebrow("09", "Test matrix (DEV / UAT)"));
kids.push(table([850, 1300, 3900, 3598], ["ID", "Verifies", "Test", "Expected"],
  UAT.map((r) => [cell(r[0], { w: 850, bold: true, size: 17 }), cell(r[1], { w: 1300, size: 17 }), cell(r[2], { w: 3900, size: 18 }), cell(r[3], { w: 3598, size: 18 })])));

/* 10 Traceability */
kids.push(eyebrow("10", "Requirement traceability"));
kids.push(p("Coverage at a glance. Detail is in the tables above.", { after: 50 }));
kids.push(table([2200, 2200, 2200, 3048], ["Requirement", "Acceptance", "Test", "Note"], [
  ["FR-01, FR-02", "AC-01", "UAT-01", ""],
  ["FR-03", "AC-03", "UAT-03", ""],
  ["FR-04", "AC-02", "UAT-02", ""],
  ["FR-05", "AC-14", "UAT-21", ""],
  ["FR-06", "AC-04", "UAT-04/05/06/26", ""],
  ["FR-07", "AC-05", "UAT-07", ""],
  ["FR-08, FR-09", "AC-06", "UAT-08/09", ""],
  ["FR-10", "AC-08", "UAT-12/13", "Open engineering question"],
  ["FR-11", "AC-07", "UAT-10/11", "Open engineering question"],
  ["FR-12, FR-13", "AC-09", "UAT-14/15", ""],
  ["FR-14, FR-15", "AC-10", "UAT-15/16", ""],
  ["FR-16", "AC-11", "UAT-16", ""],
  ["FR-17", "AC-13", "UAT-20", ""],
  ["FR-18", "AC-12", "UAT-17", ""],
  ["FR-19", "AC-15", "UAT-22", ""],
  ["FR-20", "—", "—", "Scope decision"],
  ["FR-21, FR-25", "AC-16", "UAT-23", ""],
  ["FR-22, FR-27", "AC-17", "UAT-24", ""],
  ["FR-23", "AC-18", "UAT-25", ""],
  ["FR-24", "—", "—", "Open engineering question; no MVP test"],
  ["FR-26", "AC-19", "UAT-06/27", ""],
  ["FR-28", "AC-20", "UAT-28", ""],
  ["FR-29", "AC-21", "UAT-29", ""],
  ["FR-30", "AC-22", "UAT-30", "Run at go-live"],
  ["FR-31", "AC-23", "UAT-31", "Copy supplied by Pronect"],
  ["FR-32", "AC-24", "UAT-18/19", ""],
]));

/* 11 Rollout and done */
kids.push(eyebrow("11", "Rollout and definition of done"));
kids.push(p("Monitor completion-failure and isolation-failure rates, and alert on any isolation failure. Rollout, guardrails and go-live gating are in the brief.", { after: 60 }));
kids.push(subhead("DONE MEANS"));
const dodRef = "dod";
[
  "All acceptance criteria pass in UAT, including every negative test.",
  "A guard who completes it ends with a working phone. For the MVP that rests on the guard confirming each gate; FR-24 is the first question to settle.",
  "The exclusion filter is built and the practice report is proven absent from every confirmed consumer; a forced isolation failure withholds completion.",
  "The confirm-with-engineering items are answered, and Vincent Smeyers has signed off.",
].forEach((s) => kids.push(bullet(dodRef, s)));

/* 12 Confirm with engineering */
kids.push(eyebrow("12", "Confirm with engineering before building"));
kids.push(p([chip("X"), t("  Most of this list is now answered. What is left is below, and the first one is the only item that can still change the design.")], { after: 50 }));
const valRef = "val";
[
  "Does SAM need the notifications permission, and does it prompt for itself like the microphone and camera, or does it need gating too?",
  "Can the app read whether NFC is on, and whether the handset has an NFC chip at all?",
  "What is the physical activity permission for, and does SAM actually need it? It appears in the permission list on a fresh install and nothing here establishes why.",
  "Can setup switch off the phone’s automatic permission removal for unused apps?",
  "Can the question step be suppressed from the activity table entirely? It currently writes a row typed as an “other question”.",
  "Which live surfaces must the exclusion filter cover, and where are reports created so the training flag is set there?",
].forEach((s) => kids.push(bullet(valRef, s)));

kids.push(new Paragraph({ spacing: { before: 90 },
  border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 5, space: 5 } },
  children: [t("Internal build and testing only. No production change until Vincent Smeyers has signed off.", { size: 16, it: true, color: GREY })] }));

/* ================= NUMBERING ================= */
const numbering = { config: [] };
[inRef, outRef, laterRef, consRef, dodRef, valRef].forEach((reference) => numbering.config.push({
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
