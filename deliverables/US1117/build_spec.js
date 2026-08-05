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

kids.push(new Table({
  columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
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
    ] }),
    new Paragraph({ spacing: { before: 60, after: 0 }, children: [
      t("Full requirements, acceptance criteria and tests are in the companion document.", { size: 16, it: true, color: GREY })] })],
  })] })],
}));
kids.push(gap(70));

/* ---- In one minute ---- */
kids.push(callout([
  [t("In one minute.  ", { bold: true, color: TEALD }), t("At first sign-in SAM introduces itself, gets the phone set up (permissions, NFC, background tracking) and teaches three actions — ask, report, correct — on a made-up incident. Mandatory, saves progress, complete only when each step is genuinely done. It replaces the deck’s app-setup slides and absorbs its tips; installation and log-in stay with the trainer. Practice data must stay out of every live surface, and the filter that enforces that is not built, so this runs in DEV/UAT until it is.")],
]));
kids.push(gap(40));

/* ---- 01 How it works ---- */
kids.push(eyebrow("01", "How it works, step by step"));
kids.push(p([
  t("In SAM, a guard’s message is what creates and updates their report, so a report is fixed either by voice or by editing the message. Both are practised below."),
], { after: 60 }));
const flow = "flow";
[
  "The guard signs in for the first time. SAM opens by itself, before the home screen, for new and existing guards alike.",
  "SAM opens with “Hello. I am SAM, your new assistant. From today you talk, and I write the report for you.” It says it can be taught the words the guard uses, and sets the ground rules: this is required, nothing in it counts, progress is saved.",
  "Next, three setup gates in order: permissions, NFC (which reads the checkpoint tags), then background tracking, which lives in SAM OnSite’s own settings rather than the phone’s. Each has the same shape: SAM says what it needs and why, links to that one setting from inside onboarding, and the guard switches it on and returns to confirm. Granting permissions can close the app; the guard reopens it and resumes at the same gate.",
  "Only now does the guard speak. SAM asks them to press Talk, wait for the microphone to turn green, and ask how to report an incident. That first press doubles as the microphone check. Nothing live is created here.",
  "Then the practice report. SAM asks for it the way the guard would tell a colleague and offers one to use: a Dell laptop is missing. SAM writes one report and puts it on screen as a card labelled “Training”.",
  "With the report in front of them, the guard corrects it by hand first: they press and hold their own report message, which SAM marks on screen, add that it happened at reception, and save.",
  "Then by voice: the guard says what the item actually was, and the same report updates with no duplicate. Voice runs last so no later edit can overwrite it.",
  "SAM signs off, names the settings now switched on so a guard who missed one can catch it, and a Go-to-home button lands on the normal home screen. Completion is recorded only after each step has been seen done for real.",
].forEach((s) => kids.push(numitem(flow, s)));
kids.push(new Paragraph({ spacing: { before: 20, after: 40 }, children: [
  t("Throughout, the practice report carries the training flag from the instant it is created. That flag is what must keep it out of every live surface.", { size: 18, color: GREY }),
] }));

/* ---- 02 Requirements ---- */
kids.push(eyebrow("02", "What must be true, and how we confirm it"));
kids.push(p([
  statusChip("C"), t("  an agreed behaviour.   ", { size: 16, color: GREY }),
  statusChip("P"), t("  our recommended default.   ", { size: 16, color: GREY }),
  statusChip("X"), t("  needs an engineering confirmation (listed in section 04).", { size: 16, color: GREY }),
], { after: 70 }));

function reqRow(name, key, detail) {
  const left = [new Paragraph({ spacing: { after: 30 }, children: [t(name, { size: 18, bold: true })] }),
                new Paragraph({ spacing: { after: 0 }, children: [statusChip(key)] })];
  return [cell(left, { w: 2340, va: VerticalAlign.TOP }), cell([new Paragraph({ spacing: { after: 0, line: 244, lineRule: "auto" }, children: [t(detail, { size: 18 })] })], { w: 7308, va: VerticalAlign.TOP })];
}
const REQ = [
  ["Feels like a welcome, not a test", "C", "SAM introduces itself as the guard’s new assistant and frames each step as a benefit. None of it reads as instructions. Signed off by product on a read-through, and tested with guards before release."],
  ["Automatic start & refresher", "C", "Opens by itself at first sign-in, before the home screen is reachable. A finished guard can reopen it any time without losing their completed status."],
  ["Everyone, exactly once", "C", "Every active guard receives it one time, including guards who already had accounts. After finishing, their next sign-in goes straight to the home screen."],
  ["Cannot be skipped", "C", "No path to the home screen except completing it, other than four defined exits: a real incident, a guard who cannot get past a step, a guard signing in with no connectivity, and leaving a refresher when already complete. None marks the guard complete."],
  ["Set up the phone: permissions, NFC, background tracking", "C", "Each link opens one setting and returns, so the guard is never left in the app’s ordinary navigation. The grant level must be named on screen for every permission, and the guard is working inside another app where SAM’s words are not visible, so the gate presents them one at a time or leaves a checklist up. Confirmation is a tap, since the microphone may not be on yet."],
  ["A phone with no NFC chip", "X", "The NFC gate is skipped rather than routing the guard to a support screen over hardware they do not have. Whether the app can detect the chip is an engineering confirmation."],
  ["Granting permissions may close the app", "X", "Reported by the reviewer: granting the permissions often requires restarting the app. The flow must survive it — the guard reopens SAM OnSite and resumes at the same gate, losing nothing. Whether the restart is always needed, and whether it can be avoided, is for engineering."],
  ["Ask a question (practice only)", "P", "The guard asks a spoken question and SAM answers. Nothing live may result, so the practice context must suppress the activity record the app writes today. The answer to the suggested question is written in advance."],
  ["Report one practice incident", "P", "The guard reports the made-up incident and SAM creates exactly one practice report. Repeats or retries must never create extra reports. Follow-up questions are capped and “I don’t know” always advances."],
  ["Fix a report by voice", "X", "The guard tells SAM it was a MacBook, the same report updates and no duplicate appears. Marked for engineering: built wrong, this step teaches the opposite of what is intended."],
  ["Fix a report by hand", "X", "SAM names the change and points at the right message; the guard presses and holds their report, adds the location, and saves. The finished report ends up with both “MacBook” and “reception”. The deck confirms the gesture edits text; that a saved edit updates the report is not yet confirmed."],
  ["Marked as training at creation", "C", "Every practice report carries the training flag from the moment it is created, using the platform’s archive / soft-delete flag. If it cannot be created already marked, it is not created at all. Whether the flag can be set atomically at creation is for engineering."],
  ["Kept out of every live surface", "C", "The practice report must appear in no live surface. Four are confirmed: supervisor alerts, KPI and dashboard figures, the Pronect Action Tracker, and the guard’s own activity list. Reports, exports, integrations and the stores behind them are the likely rest and are listed for engineering. The filter that excludes training-flagged data does not exist yet, which is why archived data is still visible in production. It must be built and verified before go-live; until then this stays in DEV/UAT."],
  ["Setup is per phone, training is per guard", "P", "Guards draw pooled handsets. A guard who passes the gates on one phone and resumes on another would otherwise be carried past setup on a phone that never went through it, and then marked complete for life. Gate confirmations are held against the device: an unrecognised handset replays the three gates before the flow continues, and for an already-completed guard the gates run on their own."],
  ["Finishing means the phone works too", "P", "Completion requires both halves: the three actions done for real, and the device set up. For the MVP the second half rests on the guard’s confirmation. A missing microphone is the one miss that shows itself, because Talk produces nothing at the first press. Whether the app can check these settings instead is the first question for engineering."],
  ["Finishing is earned & remembered", "P", "Recorded only after SAM sees each step done for real. A quiz or an “I’m done” button does not count. Held on the server, issued once per guard, surviving a closed app or a change of handset."],
  ["Tips taught in context", "C", "All eight tips from the deck’s General recommendations slide are inside SAM’s spoken lines, at the step where each becomes useful. Two are practised rather than told: pressing and holding text to fix a mistake is the manual-fix step, and speaking in full sentences is what the report prompt asks for."],
  ["Say what is tracked", "X", "Guards are asked for location and background tracking on a work phone and will ask what is done with it. One plain line belongs in the flow. The wording must come from Pronect. What a supervisor sees is not ours to state, and this document does not state it."],
  ["A way out for a real incident", "P", "Leave onboarding, report a real incident, return to the saved stage, with completion neither granted nor skipped. Reachable from every step. Crossing it must be unmistakable: the banner clears and the guard is told this one is live. Before the microphone exists it must offer typed reporting or say to call it in, or it is unusable exactly where it is most needed."],
];
kids.push(table([2340, 7308], ["Requirement", "What it means and how we confirm it"], REQ.map((r) => reqRow(r[0], r[1], r[2]))));

/* ---- 03 What must not go wrong ---- */
kids.push(eyebrow("03", "What must not go wrong"));
kids.push(subhead("PRACTICE DATA MUST NEVER REACH LIVE OPERATIONS"));
kids.push(p([
  t("A pretend stolen laptop must never pull a real supervisor out of bed. The practice report is stamped with the training flag at creation. "),
  t("The catch:", { bold: true }),
  t(" nothing acts on that flag yet, and archived items still show in production. The exclusion has to be written first, into "),
  t("the reporting layer", { bold: true }),
  t(", which is where the product owner has said the exclusion belongs. Until it is built this stays in DEV/UAT, and switching it on in production later must be gated on the filter by something firmer than memory: a dispatched alert cannot be recalled."),
]));
kids.push(subhead("THE RISK THAT RUNS THE OTHER WAY"));
kids.push(p([
  t("A real incident reported during practice would be flagged training and hidden from the supervisor it should reach. For the MVP the answer is cheap: a practice banner on every step, carrying the way out to normal reporting. Detection of what the guard actually said is deliberately deferred."),
]));
kids.push(subhead("NOTHING MAY LEAVE A GUARD STUCK ON SHIFT"));
kids.push(p([
  t("Failures preserve progress and offer a retry. Three failed spoken attempts offer typing. A guard who still cannot proceed reaches the home screen with completion unmarked, because being un-onboarded is a reporting problem and being locked out on shift is a safety one. A missed setting is the one gap the MVP accepts, except the microphone, which fails visibly at the first Talk press."),
], { after: 40 }));

/* ---- 04 Confirm before building ---- */
kids.push(eyebrow("04", "Confirm with engineering before building"));
kids.push(p("None of these looks like a blocker.", { after: 60 }));
const conf = "conf";
[
  "Can it be launched on its own, outside whatever normally starts a scenario?",
  "Is there a practice mode in which neither the conversation nor the report enters live reporting?",
  "How exactly does the app update an existing report, by voice and by a saved manual edit?",
  "Can the app read the state of app permissions? Of the NFC adapter? Of its own background-tracking setting? Three separate questions — a single answer for all three would lose the gates that were cheap. Assume not for now.",
  "Can the app detect whether the handset has an NFC chip?",
  "Which permissions and grant levels does the app need, can setup stop the phone’s automatic permission removal for unused apps, and does granting permissions always require a restart?",
  "Which live surfaces must the exclusion filter cover, and where are reports created so the training flag is set there?",
  "Is there already a place to store a guard’s onboarding completion, or must one be added?",
].forEach((s) => kids.push(bullet(conf, s)));

/* ---- 05 Rollout, guardrails & done ---- */
kids.push(eyebrow("05", "Rollout, guardrails, and done"));
kids.push(p([
  t("Production go-live is "),
  t("blocked", { bold: true }),
  t(" until the exclusion filter is built and verified; until then it runs only in DEV/UAT with a dedicated test tenant and test guards. Roll out behind a switch: on gradually, off instantly, and never erasing anyone’s completed status. Universal for all guards, no customer-specific setup, no changes to existing templates."),
], { after: 60 }));
kids.push(subhead("DONE MEANS"));
const done = "done";
[
  "Every guard, new or existing, completes it once by doing the real steps and ends with a set-up phone; the refresher preserves completion.",
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
