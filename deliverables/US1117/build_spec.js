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
kids.push(new Paragraph({ spacing: { after: 120 },
  border: { bottom: { color: TEAL, style: BorderStyle.SINGLE, size: 4, space: 10 } },
  children: [new TextRun({ text: "Mandatory first-login voice training — MVP specification", font: SERIF, size: 21, italics: true, color: GREY })] }));

kids.push(new Table({
  columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
  rows: [new TableRow({ children: [new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX },
    margins: { top: 90, bottom: 90, left: 180, right: 180 }, borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
    children: [new Paragraph({ spacing: { after: 0 }, children: [
      t("Applies to  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Every guard, once       ", { size: 18 }),
      t("Language  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("English       ", { size: 18 }),
      t("Status  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Draft for build       ", { size: 18 }),
      t("Sign-off  ", { size: 16, bold: true, color: TEAL, cs: 4 }), t("Vincent Smeyers", { size: 18 }),
    ] }),
    new Paragraph({ spacing: { before: 50, after: 0 }, children: [
      t("Requirements, acceptance criteria and tests are in the companion.", { size: 16, it: true, color: GREY })] })],
  })] })],
}));
kids.push(gap(60));

kids.push(callout([
  [t("In one minute.  ", { bold: true, color: TEALD }), t("At first sign-in SAM introduces itself, sets the phone up (location, NFC, background tracking) and teaches three actions — ask, report, correct — on a made-up incident. Roughly three to five minutes, mandatory, saves progress, done only when each step is genuinely done. It replaces the deck’s app-setup slides and absorbs its tips. Practice data must stay out of every live surface; the filter that enforces that is not built, so this runs in DEV/UAT until it is.")],
]));
kids.push(gap(30));

/* ---- 01 How it works ---- */
kids.push(eyebrow("01", "How it works"));
const flow = "flow";
[
  "The guard signs in. SAM opens by itself, before the home screen, for new and existing guards alike.",
  "SAM introduces itself as the guard’s new assistant, says it can be taught the words they use, and sets the ground rules: required, nothing counts, progress saved. “Hi, I’m SAM, your guarding assistant. Any question about patrols, procedures or reporting, talk about it here with me.”",
  "Three setup gates, each reached by a link in the chat: location (set to Allow all the time in settings), NFC, then background tracking, which sits in SAM OnSite’s own settings. SAM says what it needs and why, links straight to the setting, and the guard switches it on and returns. Microphone and camera are not gated; Android prompts for those when first used. Once everything is enabled the guard has to be guided back: restart the app and open the existing chat rather than start a new one, or a Continue onboarding button sits in the open space until onboarding is marked done. A new chat cannot be started before onboarding is complete, except under the urgency override.",
  "Only now does the guard speak. SAM OnSite tells the guard to press Talk, wait for the microphone, and ask how to report an incident. That first press doubles as the microphone check. Nothing live is created.",
  "The practice report. SAM asks for it the way the guard would tell a colleague and offers one: a Dell laptop is missing. One report is written.",
  "The guard is told to fix it by hand: press and hold their own report message, which SAM marks on screen, and edit it.",
  "Then by voice: the same report updates. Voice runs last, because the report takes the state of the most recent message.",
  "SAM signs off, names the settings now on so a miss can be caught, and a button lands the guard on the home screen.",
].forEach((s) => kids.push(numitem(flow, s)));

/* ---- 02 Requirements ---- */
kids.push(eyebrow("02", "What must be true"));
kids.push(p([
  statusChip("C"), t("  agreed.   ", { size: 16, color: GREY }),
  statusChip("P"), t("  our recommendation.   ", { size: 16, color: GREY }),
  statusChip("X"), t("  needs an answer (section 04).", { size: 16, color: GREY }),
], { after: 60 }));

function reqRow(name, key, detail) {
  const left = [new Paragraph({ spacing: { after: 30 }, children: [t(name, { size: 18, bold: true })] }),
                new Paragraph({ spacing: { after: 0 }, children: [statusChip(key)] })];
  return [cell(left, { w: 2340, va: VerticalAlign.TOP }), cell([new Paragraph({ spacing: { after: 0, line: 244, lineRule: "auto" }, children: [t(detail, { size: 18 })] })], { w: 7308, va: VerticalAlign.TOP })];
}
const REQ = [
  ["Feels like a welcome, not a test", "C", "SAM introduces itself as the guard’s new assistant and frames each step as a benefit. Nothing reads as instructions. Signed off by product, and tested with guards before release."],
  ["Set up the phone", "C", "Location, NFC, background tracking, in that order. Microphone and camera are not gated: Android prompts for those when first used, so SAM only warns the guard to expect it. Location is gated because the app asks for it but not at “allow all the time”, so SAM walks the guard to the second Android screen and names the setting."],
  ["Checked again each session", "C", "Guards share handsets, so a set-up phone cannot be recorded against a guard. The three settings are checked again whenever a guard starts a session. The app can read most permission states and will not advance past one that is off. Three it cannot read — location at “allow all the time”, background activity, physical activity — are guided and then taken on the guard’s word, so a miss there lasts until the next session catches it."],
  ["Marked as training at creation", "C", "Every practice report carries the training flag from the moment it is created, using the platform’s archive flag. If it cannot be created already marked, it is not created at all."],
  ["Kept out of every live surface", "C", "Four consumers are confirmed: supervisor alerts, KPIs and dashboards, the Pronect Action Tracker, and the guard’s own activity list. Reports, exports, integrations, and the queues, caches and transcript stores behind them are the likely rest and are not yet enumerated. The filter itself does not exist, which is why archived data is still visible in production."],
  ["Nothing may trap a guard", "P", "There is no way to the home screen except finishing, apart from four exits: a real incident, a guard stuck on a step, no connectivity at sign-in, and leaving a refresher when already complete. None marks the guard complete. The real-incident exit must offer typed reporting before the microphone exists, or it is unusable where it matters most."],
  ["Finishing is earned", "P", "Recorded only after each step is seen done. A quiz or an “I’m done” button does not count. Held on the server, issued once per guard, surviving a closed app or a change of handset."],
  ["Tips taught in context", "C", "All eight tips from the deck’s General recommendations slide sit inside SAM’s spoken lines, at the step where each becomes useful. Two are practised rather than told, by making the guard do the thing the tip describes."],
];
kids.push(table([2340, 7308], ["Requirement", "What it means"], REQ.map((r) => reqRow(r[0], r[1], r[2]))));

/* ---- 03 What must not go wrong ---- */
kids.push(eyebrow("03", "What must not go wrong"));
kids.push(subhead("PRACTICE DATA REACHING LIVE OPERATIONS"));
kids.push(p([
  t("A pretend stolen laptop must never pull a real supervisor out of bed. The report is flagged at creation, but nothing acts on that flag yet and archived items still show in production. The exclusion has to be written into "),
  t("the reporting layer", { bold: true }),
  t(" first. That work sits outside this specification and is not scoped here, and it sets the go-live date. Until it exists this stays in DEV/UAT. Turning it on in production later should be enforced by a server-side check that refuses to start onboarding when the filter is absent, since a dispatched alert cannot be recalled. Practice records are hidden, not deleted; clearing them is a separate database task with no date."),
]));
kids.push(subhead("THE RISK THAT RUNS THE OTHER WAY"));
kids.push(p("A real incident reported during practice would be flagged training and hidden from the supervisor it should reach. The MVP answer is cheap: a practice banner on every step, carrying the way out to normal reporting."));
kids.push(subhead("A GUARD STUCK ON SHIFT"));
kids.push(p("Failures keep progress and offer a retry; three failed spoken attempts offer typing; a guard who still cannot proceed reaches the home screen with completion unmarked. Being un-onboarded is a reporting problem. Being locked out on shift is a safety one.", { after: 30 }));

/* ---- 04 Open questions ---- */
kids.push(eyebrow("04", "Answers still needed"));
kids.push(p("Only the first can still change the design.", { after: 50 }));
const conf = "conf";
[
  "Does SAM need the notifications permission, and does it prompt for itself like the microphone and camera?",
  "What is the physical activity permission for, and does SAM need it at all?",
  "Can the app read whether NFC is on, and whether the handset has an NFC chip?",
  "Can setup stop the phone removing permissions from unused apps?",
  "The practice question writes an activity row today and must be suppressed entirely. Can the platform do that?",
  "Which live surfaces must the exclusion filter cover, and where are reports created so the flag is set there?",
].forEach((s) => kids.push(bullet(conf, s)));

/* ---- 05 Rollout & done ---- */
kids.push(eyebrow("05", "Rollout and done"));
kids.push(p("Go-live is blocked until the exclusion filter is built and verified; until then, DEV/UAT only, with a test tenant and test guards. Behind a switch: on gradually, off instantly, never erasing a completed status. Universal for all guards, no customer-specific setup, no template changes. Trainers are still needed for install and first log-in, and as the fallback for a guard who cannot get past a step.", { after: 50 }));
const done = "done";
[
  "Every guard completes it once by doing the real steps and ends with a set-up phone.",
  "The practice report is proven absent from every live surface, and a forced hide-failure withholds completion.",
  "The open questions are answered and Vincent Smeyers has signed off.",
].forEach((s) => kids.push(bullet(done, s)));

kids.push(new Paragraph({ spacing: { before: 80 },
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
