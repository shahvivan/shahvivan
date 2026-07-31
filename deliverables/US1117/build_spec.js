/*
 * SAM OnSite — Guard Onboarding  |  MVP specification (condensed, 4-page)
 * Node + docx.  Run:  node build_spec.js
 * Out:  SAM_OnSite_Guard_Onboarding_Spec.docx
 *
 * Design goals:
 *  - <= 4 pages, self-contained (no "see section X" cross-references to chase).
 *  - Plain language; every requirement carries its own acceptance/proof so the
 *    reader never has to connect dots.
 *  - Professional, human-authored feel: Georgia serif headings, a restrained
 *    slate + teal palette, eyebrow section numerals, generous but tight spacing.
 *  - Real Word numbering (no literal bullet glyphs); tables with dual DXA widths,
 *    repeated headers, non-splitting rows.
 *  - No ticket / user-story references anywhere in the content.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  Header, Footer, PageNumber, LevelFormat, TabStopType, Tab, VerticalAlign,
} = require("docx");

/* ---------- palette (calm editorial) ---------- */
const INK   = "3A3A38";   // warm charcoal — body & headings
const TEAL  = "4A6B82";   // single restrained accent (dusty blue) — numerals, thin rules
const TEALD = "2E3A44";   // deep blue-grey — subheads
const GREY  = "77746E";   // warm muted grey
const HFILL = "F4F2EE";   // soft warm header fill (light, dark text)
const ZEBRA = "F3F6F5";   // (unused — kept for compatibility)
const BOX   = "F7F6F2";   // soft callout background
const RULE  = "E6E4DF";   // hairline
const GOLD  = "9C7A32";   // (unused)

const SERIF = "Georgia";
const SANS  = "Calibri";

const PAGE_W = 12240, PAGE_H = 15840;
const MARGIN_TB = 1224;   // 0.85"
const MARGIN_LR = 1296;   // 0.90"
const CONTENT_W = PAGE_W - 2 * MARGIN_LR; // 9648 DXA

/* plain-language status labels (self-explanatory — no legend hunting) */
const STATUS = {
  C: { label: "Confirmed",         color: "43724E" }, // agreed decision / instruction
  P: { label: "Proposed",          color: "8A6A1F" }, // our recommended default
  X: { label: "Check w/ eng.",     color: "A05A38" }, // needs an engineering confirmation
};

/* ---------- helpers ---------- */
function t(text, o = {}) {
  return new TextRun({ text, font: o.font || SANS, size: o.size || 19, color: o.color || INK,
    bold: !!o.bold, italics: !!o.it, characterSpacing: o.cs });
}
function statusChip(key) {
  // quiet coloured label (no filled pill) — calmer, still scannable
  const s = STATUS[key];
  return new TextRun({ text: s.label.toUpperCase(), font: SANS, size: 14, bold: true, color: s.color, characterSpacing: 6 });
}
function p(children, o = {}) {
  return new Paragraph({
    spacing: { after: o.after == null ? 92 : o.after, line: o.line || 252, lineRule: "auto" },
    alignment: o.align, keepNext: o.keepNext,
    children: Array.isArray(children) ? children : [t(children, o)],
  });
}
function eyebrow(num, title) {
  // small accent numeral, calm serif title, faint hairline under
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 190, after: 70 }, keepNext: true,
    border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 4, space: 6 } },
    children: [
      new TextRun({ text: num + "   ", font: SANS, size: 19, bold: true, color: TEAL, characterSpacing: 10 }),
      new TextRun({ text: title, font: SERIF, size: 25, color: INK }),
    ],
  });
}
function subhead(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 160, after: 60 }, keepNext: true,
    children: [new TextRun({ text, font: SERIF, size: 20, bold: true, color: TEALD })],
  });
}

/* ---------- table primitives (open editorial style) ---------- */
const HAIR = { style: BorderStyle.SINGLE, size: 2, color: RULE };
const NONE = { style: BorderStyle.NONE };
const HEADRULE = { style: BorderStyle.SINGLE, size: 8, color: TEAL };
function cell(runsOrText, o = {}) {
  const paras = Array.isArray(runsOrText) && runsOrText[0] instanceof Paragraph
    ? runsOrText
    : [new Paragraph({
        spacing: { after: 0, line: 248, lineRule: "auto" }, alignment: o.align,
        children: Array.isArray(runsOrText) ? runsOrText : [t(runsOrText, { size: o.size || 18, bold: o.bold, color: o.color })],
      })];
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA }, verticalAlign: o.va || VerticalAlign.TOP,
    margins: { top: 46, bottom: 46, left: 40, right: 150 },
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
    children: paras,
  });
}
function hcell(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM,
    margins: { top: 26, bottom: 44, left: 40, right: 150 },
    borders: { top: NONE, bottom: HEADRULE, left: NONE, right: NONE },
    children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: text.toUpperCase(), font: SANS, size: 15, bold: true, color: TEAL, characterSpacing: 8 })] })],
  });
}
function table(colW, headers, rows) {
  const head = new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => hcell(h, colW[i])) });
  const body = rows.map((r) => new TableRow({
    cantSplit: true,
    children: r.map((c, ci) => {
      if (c instanceof TableCell) return c;
      if (Array.isArray(c) && (c[0] instanceof TextRun || c[0] instanceof Paragraph)) return cell(c, { w: colW[ci] });
      return cell(String(c == null ? "" : c), { w: colW[ci] });
    }),
  }));
  return new Table({
    columnWidths: colW, width: { size: colW.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: HAIR, insideVertical: NONE },
    rows: [head, ...body],
  });
}

/* callout: soft fill, thin left accent — light touch */
function callout(lines, accent = TEAL) {
  const inner = lines.map((ln, i) => new Paragraph({
    spacing: { after: i === lines.length - 1 ? 0 : 60, line: 258, lineRule: "auto" },
    children: Array.isArray(ln) ? ln : [t(ln)],
  }));
  return new Table({
    columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
    borders: {
      top: NONE, bottom: NONE, right: NONE,
      left: { style: BorderStyle.SINGLE, size: 12, color: accent },
      insideHorizontal: NONE, insideVertical: NONE,
    },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX },
      margins: { top: 96, bottom: 96, left: 170, right: 170 }, children: inner,
    })] })],
  });
}
function bullet(ref, children) {
  return new Paragraph({ numbering: { reference: ref, level: 0 },
    spacing: { after: 48, line: 250, lineRule: "auto" },
    children: Array.isArray(children) ? children : [t(children)] });
}
function numitem(ref, children) {
  return new Paragraph({ numbering: { reference: ref, level: 0 },
    spacing: { after: 42, line: 248, lineRule: "auto" },
    children: Array.isArray(children) ? children : [t(children)] });
}
const gap = (h) => new Paragraph({ spacing: { after: h }, children: [] });

/* ================= CONTENT ================= */
const kids = [];

/* ---- cover ---- */
kids.push(new Paragraph({ spacing: { after: 16 },
  children: [t("PRONECT   ·   SAM ONSITE", { size: 16, bold: true, color: TEAL, cs: 40 })] }));
kids.push(new Paragraph({ spacing: { after: 20 },
  children: [new TextRun({ text: "Guard Onboarding", font: SERIF, size: 46, bold: true, color: INK })] }));
kids.push(new Paragraph({ spacing: { after: 140 },
  border: { bottom: { color: TEAL, style: BorderStyle.SINGLE, size: 4, space: 10 } },
  children: [new TextRun({ text: "Mandatory first-login voice training — MVP specification", font: SERIF, size: 21, italics: true, color: GREY })] }));

/* meta strip (no ticket data) */
kids.push(new Table({
  columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [new TableRow({ children: [new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX },
    margins: { top: 100, bottom: 100, left: 180, right: 180 }, borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
    children: [new Paragraph({ spacing: { after: 0 }, children: [
      t("Product  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("SAM OnSite       ", { size: 18 }),
      t("Applies to  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Every guard, once       ", { size: 18 }),
      t("Language  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("English       ", { size: 18 }),
      t("Status  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Draft for build       ", { size: 18 }),
      t("Sign-off  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Vincent Smeyers", { size: 18 }),
    ] })],
  })] })],
}));
kids.push(gap(70));

/* ---- In one minute ---- */
kids.push(callout([
  [t("In one minute.  ", { bold: true, color: TEALD }), t("The first time a guard signs in, SAM introduces itself, gets the phone set up, and shows what it can do. It should feel like being handed a good tool rather than sitting a test. SAM walks the guard through the three everyday actions (ask a question, report an incident, fix a report) on a made-up incident, so nothing real is affected. It takes a few minutes, saves progress, cannot be skipped, and counts as done only once each step is genuinely done. It replaces the app-onboarding steps currently in the onboarding presentation. Practice data must stay out of every live surface. The filter that enforces this is not built yet, so it runs in the internal test environments until that is ready.")],
]));
kids.push(gap(40));

/* ---- 01 How it works ---- */
kids.push(eyebrow("01", "How it works, step by step"));
kids.push(p([
  t("In SAM, a guard’s message is what creates and updates their report, so you fix it either by voice or by editing the message. Both are practised below. Quoted lines show the intended tone; the wording is settled with product before build."),
], { after: 60 }));
const flow = "flow";
[
  "The guard signs in for the first time. SAM opens by itself, before the home screen, for new and existing guards alike.",
  "SAM opens with “Hello, I’m SAM. From today you talk and I write it up. Give me a few minutes and I’ll show you how.” The screen also sets the ground rules. This is required, nothing in it counts, and progress is saved, so leaving and coming back resumes at the same step.",
  "Next the phone gets set up, in three gates: app permissions, NFC (which reads the checkpoint tags), then background tracking. Each follows the same shape. SAM says what it needs and why, links to the right place from inside onboarding, and asks the guard to switch it on and come back to confirm. The permissions link opens the phone’s settings page for SAM OnSite, where the guard grants each one; the NFC and background-tracking links open a single setting each, and background tracking is the one that lives inside SAM rather than the phone. All three return to onboarding, so the guard is never left in the app’s ordinary navigation. Naming the grant level matters, since Android asks twice for location and the app needs “all the time”. A phone with no NFC chip skips that gate rather than routing the guard to support.",
  "Only now does the guard speak. “Press Talk and try this one: how do I report an incident?” SAM answers with an answer written in advance, and that first press doubles as the microphone check. If the permission was confirmed but never actually granted, the guard is sent back to that gate rather than left pressing a button that will never work. Nothing live is created here.",
  "Then the real thing: SAM asks for a report “the way you’d tell a colleague” and offers one to use, a Dell laptop that has gone missing. SAM asks at most a couple of follow-up questions, accepts “I don’t know” for any of them, writes one practice report, and puts it on screen as a card labelled “Training”. “That’s it written up. Have a look.”",
  "With the report in front of them, the guard learns to correct it. “Say the wrong thing and I’ll fix it. Try it now: tell me it was a MacBook, not a Dell.” The same report and the guard’s own message update. No second report appears, and SAM says so. “Changed. Same report.”",
  "The second way is by hand. SAM names the change it wants and points at the message to edit, since by now the guard has sent several. They long-press their report, add that it happened at reception, and save, and the finished report shows both “MacBook” and “reception”.",
  "SAM signs off. “That’s everything. You can report now, and you can put it right when I get it wrong.” Completion is recorded only after SAM has seen each step done for real. A Go-to-home button lands the guard on the normal home screen.",
].forEach((s) => kids.push(numitem(flow, s)));
kids.push(new Paragraph({ spacing: { before: 20, after: 40 }, children: [
  t("Throughout, the practice report carries the training flag from the instant it is created. That flag is what must keep it out of every live surface.", { size: 18, color: GREY }),
] }));

/* ---- 02 Requirements & proof ---- */
kids.push(eyebrow("02", "What must be true, and how we confirm it"));
kids.push(p([
  t("Each row is a requirement together with how we will confirm it works, so nothing is left to interpretation. ", { size: 18 }),
  statusChip("C"), t("  an agreed behaviour (a few note an implementation detail to check).   ", { size: 16, color: GREY }),
  statusChip("P"), t("  our recommended default.   ", { size: 16, color: GREY }),
  statusChip("X"), t("  needs an engineering confirmation (listed in section 05).", { size: 16, color: GREY }),
], { after: 70 }));

function reqRow(name, key, detail) {
  const left = [new Paragraph({ spacing: { after: 30 }, children: [t(name, { size: 18, bold: true })] }),
                new Paragraph({ spacing: { after: 0 }, children: [statusChip(key)] })];
  return [cell(left, { w: 2340, va: VerticalAlign.TOP }), cell([new Paragraph({ spacing: { after: 0, line: 244, lineRule: "auto" }, children: [t(detail, { size: 18 })] })], { w: 7308, va: VerticalAlign.TOP })];
}
const REQ = [
  ["Feels like a welcome, not a test", "C", "SAM opens with a warm greeting saying what it does for the guard and why that helps, frames each step as a benefit (“That’s it written up”), and signs off warmly. None of it should read as instructions. Confirmed by a product read-through and by four of five test guards calling it welcoming rather than test-like."],
  ["Automatic start & refresher", "C", "Opens by itself at first sign-in, before the home screen is reachable. A finished guard can reopen it any time as a refresher without losing their completed status."],
  ["Everyone, exactly once", "C", "Every active guard receives it one time, including guards who already had accounts before this feature. After finishing, their next sign-in goes straight to the home screen."],
  ["Cannot be skipped", "C", "There is no way to reach the home screen except by completing it."],
  ["Set up the phone: permissions, NFC, background tracking", "C", "Three gates in sequence, each the same shape: SAM explains what it needs and why, links to the right place from inside onboarding, and the guard taps to confirm on return. Every link returns to onboarding, so the guard is never left in the app’s ordinary navigation. Three details decide whether this works. The words on screen must name the grant level, and say that Android asks for location twice. Setup must switch off the phone’s automatic permission removal for unused apps. And a phone with no NFC chip skips that gate rather than being sent to support. Confirmation is a tap, since the microphone may not be on yet. The companion carries the full gate behaviour."],
  ["Ask a question (practice only)", "P", "The guard asks a spoken question and SAM answers. Nothing live may result. Today the app logs even questions as activities, so the practice mode must suppress that record. The answer to the suggested question is written in advance, and the prompt names that question rather than inviting an open one."],
  ["Report one practice incident", "P", "The guard reports the made-up incident and SAM creates exactly one practice report. Repeats or retries must never create extra reports. Follow-up questions are capped and “I don’t know” always advances, since the guard is inventing details about a laptop that does not exist."],
  ["Fix a report by voice", "X", "The guard tells SAM it was a MacBook, the same report updates and no duplicate appears. The prompt must not assume the guard used the word “Dell”, since they were asked to report in their own words. Marked for engineering because how the app updates an existing report is not yet confirmed, and this step teaches the opposite lesson if it creates a second report instead."],
  ["Fix a report by hand", "X", "SAM points at the right message and names the change, and the guard long-presses their report, adds the location, and saves; the finished report ends up with both “MacBook” and “reception” in it. Pointing matters: by this step the guard has sent a question, a report, answers and a correction. Marked for engineering because it is not yet confirmed the app supports this edit."],
  ["Marked as training at creation", "C", "Every practice report carries the training flag (the platform’s archive / soft-delete flag) from the moment it is created. It is set as part of creating the report, so a report never exists without it."],
  ["Kept out of every live surface", "C", "The practice report must appear in no live surface: reports, dashboards, analytics, exports, alerts and integrations, or the guard’s own activity list and search. The filter that excludes training-flagged data does not exist yet, which is why archived data is still visible in production. It must be built and verified before go-live; until then this stays in DEV/UAT."],
  ["Finishing is earned & remembered", "P", "Completion is recorded only after SAM sees each step done for real. A quiz or an “I’m done” button does not count. Stored on the server as version 1, issued once per guard, survives closing the app or switching devices, and cannot be duplicated."],
  ["Tips taught in context", "P", "The tips from the train-the-trainer deck belong inside the practice steps, at the moment each becomes useful, with SAM slipping one in as an aside. A screen of tips read out in order does not meet this. Which tips, and where each fits, once the slide is shared."],
  ["Say what is tracked, and who sees it", "P", "Guards are asked for all-the-time location and background tracking on a work phone. A line in the flow says what each is for and what a supervisor can and cannot see. Wording to be agreed with product and, where required, with the works council."],
  ["Visual aids come later", "C", "Extra visual help (tooltips, highlights) for dark or noisy sites is not built now; the need is judged during testing."],
  ["A way out for a real incident (minor)", "P", "Leave onboarding, report a real incident through normal reporting, return to where they left off, with completion neither granted nor skipped. Reachable from every step. Crossing it must be unmistakable, since the guard has just been told nothing counts: the practice banner clears and they are told this one is live."],
];
kids.push(table([2340, 7308], ["Requirement", "What it means and how we confirm it"], REQ.map((r) => reqRow(r[0], r[1], r[2]))));

/* ---- 03 Practice data, the reverse risk, and completion ---- */
kids.push(eyebrow("03", "What must not go wrong"));
kids.push(subhead("PRACTICE DATA MUST NEVER REACH LIVE OPERATIONS"));
kids.push(p([
  t("A pretend stolen laptop must never pull a real supervisor out of bed at three in the morning. So the practice report is stamped with the training flag at the moment it is created, using the archive / soft-delete flag the platform already has. "),
  t("The catch:", { bold: true }),
  t(" nothing yet acts on that flag. Archived items still show up in production today, so the filter that hides them has to be written first. Until it exists and has been checked, this stays in the internal test environments and does not go near production. Write it once, "),
  t("hide-by-default", { bold: true }),
  t(" and low down where the data is read, and every screen above it inherits the behaviour without being asked. Clearing the practice records out for good is a later job. All that is kept is the guard’s completion status. The MVP never leaves DEV/UAT, so nothing here can reach production during the build. What matters is that switching it on in production later is gated on the filter by something firmer than memory, because a dispatched alert cannot be recalled."),
]));
kids.push(subhead("THE RISK THAT RUNS THE OTHER WAY"));
kids.push(p([
  t("A guard is thirty seconds into being taught that this is how you report things when something actually happens. They press Talk and report it, and the safety mechanism swallows it: flagged training, hidden from the supervisor it should have reached. For the MVP the answer is cheap — a practice banner on every step, carrying the way out to normal reporting. Having SAM notice that what was just said is not the practice incident is better and is listed as deliberately deferred."),
]));
kids.push(subhead("COMPLETION IS SERVER-SIDE, EARNED, AND SAFE UNDER PRESSURE"));
kids.push(p([
  t("Nothing is marked finished on the guard’s word. The system watches each action actually happen, and the record of it sits on the server, never trusted from the phone. One per guard, it survives the app closing or the guard picking up a different handset, and two sessions running at once can never produce duplicate reports or push progress backwards. If practice data cannot be safely hidden, completion is withheld and the guard stays in the exercise. Protecting live data comes before finishing it. What it never does is trap anyone: a guard who cannot get past a step reaches the home screen anyway, uncompleted."),
], { after: 40 }));

/* ---- 04 Unhappy paths (summary; the companion carries the full table) ---- */
kids.push(eyebrow("04", "Handling the unhappy paths"));
kids.push(p("Nothing here may leave a guard stuck on shift. Network, speech and save failures keep progress and offer a retry, and no edit is silently lost. After three failed attempts at a spoken step, typing is offered. If practice data cannot be safely hidden, the step stops, completion is withheld, and only a stage code is recorded, never the practice content; if that persists, the guard is released to the home screen with completion unmarked. A missed setting is the one the MVP accepts: SAM takes the guard’s confirmation and moves on, except the microphone, which is checked at the first Talk press because nothing after it works without one. The companion carries the full table.", { after: 40 }));


/* ---- 05 Confirm before building ---- */
kids.push(eyebrow("05", "Confirm with engineering before building"));
kids.push(p("These depend on how the platform already behaves. They are gathered here to be settled once, up front, and none of them looks like a blocker.", { after: 60 }));
const conf = "conf";
[
  "Does the app already know when a guard signs in for the first time?",
  "Can it be launched on its own, outside whatever normally starts a scenario?",
  "Is there a practice mode in which neither the conversation nor the report enters live reporting?",
  "How exactly does the app update an existing report, by voice and by a saved manual edit?",
  "Which permissions and grant levels does the app need (e.g. location “all the time”), and can it read back permissions, NFC and background tracking to check them rather than asking the guard? Its Location screen already shows a “Permissions needed” status and a settings link, so this partly exists. Can it also stop the phone’s automatic permission removal for unused apps?",
  "The archive / training flag exists. Confirm which live surfaces the new exclusion filter must cover, and where reports are created so the flag is set at creation.",
  "Is there already a place to store a guard’s onboarding completion, or must one be added?",
].forEach((s) => kids.push(bullet(conf, s)));

/* ---- 06 Rollout, guardrails & done ---- */
kids.push(eyebrow("06", "Rollout, guardrails, and done"));
kids.push(p([
  t("Production go-live is "),
  t("blocked", { bold: true }),
  t(" until the exclusion filter that hides training data from every live surface is built and verified; until then it runs only in DEV/UAT with a dedicated test tenant and test guards. Roll out behind a switch: on gradually, off instantly, and never erasing anyone’s completed status. Universal for all guards, no customer-specific setup, no changes to existing templates."),
], { after: 60 }));
kids.push(subhead("DONE MEANS"));
const done = "done";
[
  "A guard, new or existing, is shown the exercise once, completes it by doing the real steps, and lands on the home screen; reopening later keeps their completed status.",
  "The exclusion filter is built, the practice report is proven absent from every live surface, and a forced hide-failure withholds completion.",
  "The confirm-with-engineering items are answered and Vincent Smeyers has signed off.",
].forEach((s) => kids.push(bullet(done, s)));

kids.push(new Paragraph({ spacing: { before: 90 },
  border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 5, space: 5 } },
  children: [t("Internal build and testing only. No production change until Vincent Smeyers has signed off.", { size: 16, it: true, color: GREY })] }));

/* ================= NUMBERING ================= */
const numbering = { config: [] };
["conf", "done"].forEach((reference) => numbering.config.push({
  reference, levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
    style: { run: { font: SANS, size: 19, color: TEAL }, paragraph: { indent: { left: 360, hanging: 220 } } } }],
}));
["flow"].forEach((reference) => numbering.config.push({
  reference, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
    style: { run: { font: SANS, size: 19, bold: true, color: TEAL }, paragraph: { indent: { left: 400, hanging: 260 } } } }],
}));

/* ================= DOCUMENT ================= */
const doc = new Document({
  creator: "Pronect", title: "SAM OnSite — Guard Onboarding (MVP specification)",
  description: "Guard onboarding MVP specification",
  styles: {
    default: { document: { run: { font: SANS, size: 19, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: SERIF, size: 24, bold: true, color: INK } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: SANS, size: 19, bold: true, color: TEALD } },
    ],
  },
  numbering,
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H },
      margin: { top: MARGIN_TB, bottom: MARGIN_TB, left: MARGIN_LR, right: MARGIN_LR } } },
    headers: { default: new Header({ children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 3, space: 3 } },
      spacing: { after: 0 },
      children: [ t("SAM OnSite — Guard Onboarding", { size: 15, color: GREY }),
        new TextRun({ children: [new Tab()], font: SANS, size: 15 }),
        t("MVP specification", { size: 15, color: GREY }) ],
    })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 3, space: 3 } },
      spacing: { before: 0 },
      children: [ t("Confidential — Pronect internal", { size: 15, color: GREY }),
        new TextRun({ children: [new Tab(), "Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: SANS, size: 15, color: GREY }) ],
    })] }) },
    children: kids,
  }],
});

const OUT = path.join(__dirname, "SAM_OnSite_Guard_Onboarding_Spec.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Wrote", OUT, "(" + buf.length + " bytes)"); });
