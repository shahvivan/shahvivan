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
  [t("In one minute.  ", { bold: true, color: TEALD }), t("The first time a guard signs in, SAM introduces itself and shows what it can do. It should feel like unwrapping a good tool, not sitting a test: SAM welcomes them, then walks through the three everyday actions — ask a question, report an incident, fix a report — on a made-up incident so nothing real is affected. It takes a few minutes, saves progress, cannot be skipped, and counts as done only once each step is genuinely done. Practice data must stay out of every live report and dashboard — and because the filter that enforces this is not built yet, it runs in the internal test environments until that is ready.")],
]));
kids.push(gap(40));

/* ---- 01 How it works ---- */
kids.push(eyebrow("01", "How it works, step by step"));
kids.push(p([
  t("In SAM, a guard’s message is what creates and updates their report — fix it by voice or by editing the message; both are practised below. Quoted lines show intended tone, not final copy."),
], { after: 60 }));
const flow = "flow";
[
  "The guard signs in for the first time and SAM opens by itself, before the home screen — for new and existing guards alike.",
  "SAM introduces itself: “Hello, I’m SAM. From today I do the paperwork with you — a few minutes and you’ll know the essentials.” The screen makes three things clear up front: this is required, everything in it is pretend, and progress is saved — leaving and returning resumes where they left off.",
  "SAM gets itself set up — first permissions, then NFC: “Tap here, turn on everything SAM OnSite needs, then come back and type done.” Each link opens the right place in the phone’s settings (the app’s Permissions page, then the NFC switch that reads checkpoint tags); the guard turns things on, comes back, and says or types “done.” SAM checks each item itself — naming anything wrong and linking back, or moving on.",
  "SAM invites their first words: “Ask me anything you’d ask a colleague — press Talk and try: how do I report an incident?” SAM answers and confirms it heard them — the first press of Talk doubles as the microphone check — then adds: “Any time you’re unsure, just ask.” Nothing live is created here.",
  "SAM moves to the useful part: “Now tell me about an incident the way you’d tell a colleague — let’s pretend a Dell laptop was stolen.” SAM asks for anything missing, writes one practice report, and shows it on screen as a card visibly labelled “Training”: “Done. I wrote that up while you talked — have a look.”",
  "With the report in front of them, SAM shows corrections are easy: “Got a detail wrong? Just say so — tell me to change the Dell to a MacBook.” The same report and the guard’s own message update; no second report appears.",
  "SAM offers the other way: “Prefer to type? Press and hold your message and edit it yourself.” The guard edits it to “A MacBook was stolen at reception,” and the finished report shows both “MacBook” and “reception.”",
  "SAM closes warmly: “You’re ready — you can ask, report, and fix reports whenever you need me.” Completion is recorded only after SAM has seen each step done for real, and a Go-to-home button lands them on the normal home screen.",
].forEach((s) => kids.push(numitem(flow, s)));
kids.push(new Paragraph({ spacing: { before: 20, after: 40 }, children: [
  t("Throughout, the practice report carries a training mark from the instant it is created — the mark that must keep it out of every live surface.", { size: 18, color: GREY }),
] }));

/* ---- 02 Requirements & proof ---- */
kids.push(eyebrow("02", "What must be true — and how we confirm it"));
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
  ["Feels like a welcome, not a test", "C", "SAM opens with a warm greeting saying what it does for the guard and why that helps, frames each step as a benefit (“I wrote that up while you talked”), encourages briefly after each, and signs off warmly. Short, friendly, first-person — never a dry list of instructions. Confirmed by a product read-through and by four of five test guards calling it welcoming, not a test."],
  ["Automatic start & refresher", "C", "Opens by itself at first sign-in, before the home screen is reachable. A finished guard can reopen it any time as a refresher without losing their completed status."],
  ["Everyone, exactly once", "C", "Every active guard receives it one time — including guards who already had accounts before this feature. After finishing, their next sign-in goes straight to the home screen."],
  ["Cannot be skipped", "C", "There is no way to reach the home screen except by completing it."],
  ["Permissions and NFC — checked, not assumed", "C", "SAM verifies each permission is at the level the app needs — grants have levels (“only while using”, “ask every time”, location “all the time”) — and names anything wrong rather than taking the guard’s word; the same pattern for NFC. Setup also turns off the phone’s automatic permission removal for unused apps, so access is not silently lost later. The “done” must accept typing, since until the microphone is on the guard cannot speak. What truly cannot be enabled routes to support — never trapped."],
  ["Ask a question (practice only)", "P", "The guard asks a spoken question and SAM answers. Nothing live may result — today the app logs even questions as activities, so either the practice mode suppresses that record or it carries the training mark like everything else here."],
  ["Report one practice incident", "P", "The guard reports the made-up incident and SAM creates exactly one practice report. Repeats or retries must never create extra reports."],
  ["Fix a report by voice", "C", "The guard tells SAM to change “Dell” to “MacBook”; the same report updates and no duplicate appears. (The exact way the app updates the report is a Check-with-engineering item.)"],
  ["Fix a report by hand", "C", "The guard long-presses their own message, edits it, and saves; the finished report contains both “MacBook” and “reception.” (Whether the app already supports this edit is a Check-with-engineering item.)"],
  ["Marked as training at creation", "C", "Every practice report is flagged as training (archive / soft-delete) the moment it is created — the platform already has this flag for exactly this purpose — so a report never exists un-marked."],
  ["Kept out of every live surface", "C", "The practice report must appear in no live surface — reports, dashboards, analytics, exports, automated alerts and integrations, and the guard’s own activity list and search. The filter that excludes training-marked data does not exist yet, so archived data is still visible in production; it must be built and verified before go-live. Until then this runs only in the internal test environments."],
  ["Finishing is earned & remembered", "P", "Completion is recorded only after SAM sees each step done for real — never a quiz or an “I’m done” button. Stored on the server as version 1, issued once per guard, survives closing the app or switching devices, and cannot be duplicated."],
  ["Visual aids come later", "C", "Extra visual help (tooltips, highlights) for dark or noisy sites is not built now; the need is judged during testing."],
  ["A way out for a real incident (minor)", "P", "Deliberately small: leave onboarding, report a real incident, and return to where they left off, without it counting as finishing or skipping. In practice guards handle real emergencies as they always have, and new guards are accompanied early on — a low-priority safeguard, not a focus."],
];
kids.push(table([2340, 7308], ["Requirement", "What it means and how we confirm it"], REQ.map((r) => reqRow(r[0], r[1], r[2]))));

/* ---- 03 Practice data & completion (the two things that must not go wrong) ---- */
kids.push(eyebrow("03", "The two things that must not go wrong"));
kids.push(subhead("PRACTICE DATA MUST NEVER REACH LIVE OPERATIONS"));
kids.push(p([
  t("The made-up incident must never trigger a real response. Every practice report is marked as training the moment it is created — the platform already has this archive / soft-delete flag. "),
  t("The dependency to be clear about:", { bold: true }),
  t(" the filtering that keeps marked data out of live surfaces does not exist yet — today, archived items are still visible in production — so it is a required enhancement, not something that works today. This runs only in the internal test environments until that filter is built and verified, and must not go live before then. When built, it should work "),
  t("hide-by-default", { bold: true }),
  t(" at the shared data layer, so every current and future surface excludes training data automatically. Permanent deletion is a separate clean-up done later; only the guard’s completion status is kept."),
]));
kids.push(subhead("COMPLETION IS SERVER-SIDE, EARNED, AND SAFE UNDER PRESSURE"));
kids.push(p([
  t("Completion is recorded only after the system observes each real action, and lives on the server (never trusted from the phone). It is issued once per guard, resumes safely if the app closes or the guard switches devices, and two sessions at once can never create duplicate reports or roll progress backwards. If practice data cannot be safely hidden, completion is withheld and the guard stays in the exercise — protecting live data outranks finishing."),
], { after: 40 }));

/* ---- 04 Unhappy paths ---- */
kids.push(eyebrow("04", "Handling the unhappy paths"));
kids.push(table([2500, 7148], ["If this happens", "What the guard experiences"], [
  ["A permission or NFC is still off", "SAM names exactly what is missing and links straight back; it never advances on the guard’s word alone. What truly cannot be enabled routes to support."],
  ["Network, speech, or save fails", "Progress is kept and a clear “try again” is shown; the current step and any edit are never silently lost."],
  ["A spoken correction is misheard", "SAM asks again; the guard can retry by voice or switch to the manual long-press edit instead."],
  ["A manual edit fails to save", "The failure is shown and the guard’s typed text is kept for another try; the report is never left half-changed."],
  ["Practice data cannot be safely hidden", "The step stops safely: completion is withheld and the guard stays in the exercise. Only a stage/error code is recorded — never the practice content."],
]));

/* ---- 05 Confirm before building ---- */
kids.push(eyebrow("05", "Confirm with engineering before building"));
kids.push(p("These depend on how the platform already behaves — gathered here so they are settled once, up front. None is expected to be a blocker.", { after: 60 }));
const conf = "conf";
[
  "Does the app already know when a guard signs in for the first time?",
  "Can this exercise be launched on its own, without the usual triggers that start other scenarios?",
  "Can SAM run in a practice mode whose conversation and report never enter live reporting?",
  "How exactly does the app update an existing report — by voice, and by a saved manual edit?",
  "Which permissions and grant levels does the app need (e.g. location “all the time”), can it read each one to verify, and can it stop the phone’s automatic permission removal? Its Location screen already shows a “Permissions needed” status and a settings link, so the mechanism partly exists.",
  "The archive / training flag exists — confirm which live surfaces the new exclusion filter must cover, and where reports are created so the flag is set at creation.",
  "Is there already a place to store a guard’s onboarding completion, or must one be added?",
].forEach((s) => kids.push(bullet(conf, s)));

/* ---- 06 Rollout, guardrails & done ---- */
kids.push(eyebrow("06", "Rollout, guardrails, and done"));
kids.push(p([
  t("Production go-live is "),
  t("blocked", { bold: true }),
  t(" until the exclusion filter that hides training data from every live surface is built and verified; until then it runs only in DEV/UAT with a dedicated test tenant and test guards. Roll out behind a switch — on gradually, off instantly, never erasing completed status. Universal for all guards, no customer-specific setup, no changes to existing templates."),
], { after: 60 }));
kids.push(subhead("DONE MEANS"));
const done = "done";
[
  "A guard — new or existing — is shown the exercise once, completes it by doing the real steps, and lands on the home screen; reopening later keeps their completed status.",
  "The exclusion filter is built, the practice report is proven absent from every live surface, and a forced hide-failure withholds completion.",
  "The confirm-with-engineering items are answered and Vincent Smeyers has signed off.",
].forEach((s) => kids.push(bullet(done, s)));

kids.push(new Paragraph({ spacing: { before: 90 },
  border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 5, space: 5 } },
  children: [t("Internal build and testing only — no production change until Vincent Smeyers has signed off.", { size: 16, it: true, color: GREY })] }));

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
