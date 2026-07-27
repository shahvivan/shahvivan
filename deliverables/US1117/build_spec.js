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

/* ---------- palette ---------- */
const INK   = "1F2A33";   // slate — body & headings
const TEAL  = "2F6E63";   // accent — rules, eyebrows, subheads
const TEALD = "234f47";   // darker teal
const GREY  = "5C6B73";   // muted
const HFILL = "1F2A33";   // table header fill (slate)
const ZEBRA = "F3F6F5";   // faint row band
const BOX   = "F5F8F7";   // callout background
const RULE  = "CDD8D5";   // hairline
const GOLD  = "9C7A32";   // subtle warm accent (used sparingly)

const SERIF = "Georgia";
const SANS  = "Calibri";

const PAGE_W = 12240, PAGE_H = 15840;
const MARGIN_TB = 1224;   // 0.85"
const MARGIN_LR = 1296;   // 0.90"
const CONTENT_W = PAGE_W - 2 * MARGIN_LR; // 9648 DXA

/* plain-language status labels (self-explanatory — no legend hunting) */
const STATUS = {
  C: { label: "Confirmed",         color: "2F6E4E" }, // agreed decision / instruction
  P: { label: "Proposed",          color: "8A6100" }, // our recommended default
  X: { label: "Check w/ eng.",     color: "9A3B2E" }, // needs an engineering confirmation
};

/* ---------- helpers ---------- */
function t(text, o = {}) {
  return new TextRun({ text, font: o.font || SANS, size: o.size || 19, color: o.color || INK,
    bold: !!o.bold, italics: !!o.it, characterSpacing: o.cs });
}
function statusChip(key) {
  const s = STATUS[key];
  return new TextRun({ text: ` ${s.label} `, font: SANS, size: 15, bold: true, color: "FFFFFF",
    shading: { type: ShadingType.CLEAR, color: "auto", fill: s.color } });
}
function p(children, o = {}) {
  return new Paragraph({
    spacing: { after: o.after == null ? 90 : o.after, line: o.line || 250, lineRule: "auto" },
    alignment: o.align, keepNext: o.keepNext,
    children: Array.isArray(children) ? children : [t(children, o)],
  });
}
function eyebrow(num, title) {
  // "01 —  Title"  with teal numeral, slate Georgia title, hairline under
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 200, after: 80 }, keepNext: true,
    border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 6, space: 4 } },
    children: [
      new TextRun({ text: num + "  ", font: SERIF, size: 24, bold: true, color: TEAL }),
      new TextRun({ text: title, font: SERIF, size: 24, bold: true, color: INK }),
    ],
  });
}
function subhead(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 120, after: 50 }, keepNext: true,
    children: [new TextRun({ text, font: SANS, size: 19, bold: true, color: TEALD, characterSpacing: 4 })],
  });
}

/* ---------- table primitives ---------- */
const BRD = { style: BorderStyle.SINGLE, size: 3, color: "D5DEDB" };
function cellBorders() { return { top: BRD, bottom: BRD, left: BRD, right: BRD }; }
function cell(runsOrText, o = {}) {
  const paras = Array.isArray(runsOrText) && runsOrText[0] instanceof Paragraph
    ? runsOrText
    : [new Paragraph({
        spacing: { after: 0, line: 244, lineRule: "auto" }, alignment: o.align,
        children: Array.isArray(runsOrText) ? runsOrText : [t(runsOrText, { size: o.size || 18, bold: o.bold, color: o.color })],
      })];
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA }, verticalAlign: o.va || VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    margins: { top: 54, bottom: 54, left: 96, right: 96 }, borders: cellBorders(),
    children: paras,
  });
}
function hcell(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, color: "auto", fill: HFILL },
    margins: { top: 60, bottom: 60, left: 96, right: 96 }, borders: cellBorders(),
    children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text, font: SANS, size: 18, bold: true, color: "FFFFFF", characterSpacing: 3 })] })],
  });
}
function table(colW, headers, rows) {
  const head = new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => hcell(h, colW[i])) });
  const body = rows.map((r, ri) => new TableRow({
    cantSplit: true,
    children: r.map((c, ci) => {
      const fill = ri % 2 ? ZEBRA : undefined;
      if (c instanceof TableCell) return c;
      if (Array.isArray(c) && (c[0] instanceof TextRun || c[0] instanceof Paragraph)) return cell(c, { w: colW[ci], fill, va: VerticalAlign.TOP });
      return cell(String(c == null ? "" : c), { w: colW[ci], fill, va: VerticalAlign.TOP });
    }),
  }));
  return new Table({
    columnWidths: colW, width: { size: colW.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: { top: BRD, bottom: BRD, left: BRD, right: BRD, insideHorizontal: BRD, insideVertical: BRD },
    rows: [head, ...body],
  });
}

/* callout: left accent bar, soft fill, optional bold lead lines */
function callout(lines, accent = TEAL) {
  const inner = lines.map((ln, i) => new Paragraph({
    spacing: { after: i === lines.length - 1 ? 0 : 60, line: 248, lineRule: "auto" },
    children: Array.isArray(ln) ? ln : [t(ln)],
  }));
  return new Table({
    columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 3, color: RULE },
      bottom: { style: BorderStyle.SINGLE, size: 3, color: RULE },
      left: { style: BorderStyle.SINGLE, size: 22, color: accent },
      right: { style: BorderStyle.SINGLE, size: 3, color: RULE },
      insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX },
      margins: { top: 80, bottom: 80, left: 150, right: 150 }, children: inner,
    })] })],
  });
}
function bullet(ref, children) {
  return new Paragraph({ numbering: { reference: ref, level: 0 },
    spacing: { after: 54, line: 248, lineRule: "auto" },
    children: Array.isArray(children) ? children : [t(children)] });
}
function numitem(ref, children) {
  return new Paragraph({ numbering: { reference: ref, level: 0 },
    spacing: { after: 46, line: 246, lineRule: "auto" },
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
kids.push(new Paragraph({ spacing: { after: 100 },
  border: { bottom: { color: TEAL, style: BorderStyle.SINGLE, size: 14, space: 6 } },
  children: [new TextRun({ text: "Mandatory first-login voice training — MVP specification", font: SERIF, size: 21, italics: true, color: GREY })] }));

/* meta strip (no ticket data) */
kids.push(new Table({
  columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [new TableRow({ children: [new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX },
    margins: { top: 70, bottom: 70, left: 150, right: 150 }, borders: cellBorders(),
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
  [t("In one minute.  ", { bold: true, color: TEALD }), t("The first time a guard signs in, SAM — the in-app voice assistant — runs a short, required voice exercise that teaches the three things they will do every day: ask a question, report an incident, and correct a report. It uses a made-up incident so nothing real is affected, cannot be skipped, takes a few minutes, saves progress, and is marked complete only after the guard actually performs each step. All practice data must be kept out of every live report and dashboard — and because the filter that enforces this is not built yet, the exercise runs in the internal test environments until it is. A guard can reopen the exercise later as a refresher.")],
]));
kids.push(gap(40));

/* ---- 01 How it works ---- */
kids.push(eyebrow("01", "How it works, step by step"));
kids.push(p([
  t("Good to know before reading: in SAM, a guard’s message is what creates and updates their report — so a report can be fixed either by telling SAM out loud or by editing the original message. Both are practised below."),
], { after: 60 }));
const flow = "flow";
[
  "The guard signs in. If they have not completed this onboarding, it opens automatically — before the home screen — for new and existing guards alike.",
  "A welcome screen explains that the exercise is required, uses made-up information, is short, and saves progress.",
  "The guard turns on the microphone. If they decline, SAM keeps asking with clear guidance. If the microphone truly cannot be enabled (broken hardware or company device policy), the guard is taken to a support screen — showing the site’s support contact — that still carries the always-visible urgent-incident action, so they are never trapped.",
  "The guard asks SAM a practice question out loud (for example, “How do I report an incident?”) and SAM answers. This is practice only — no report is created.",
  "The guard reports a made-up incident (“A Dell laptop was stolen”). SAM asks for anything missing and creates one practice report.",
  "The guard says to change “Dell” to “MacBook.” The same report — and the guard’s own message — update to read “MacBook”; no second report appears.",
  "The guard long-presses their own message, edits it to “A MacBook was stolen at reception,” and saves. The finished report now shows both “MacBook” and “reception.”",
  "The exercise is marked complete — but only after SAM has seen the guard do each step for real (a question answered, one practice report created, the report showing each correction) — and the guard arrives at the normal home screen.",
].forEach((s) => kids.push(numitem(flow, s)));
kids.push(new Paragraph({ spacing: { before: 20, after: 40 }, children: [
  t("Throughout, the practice report carries a training mark from the instant it is created; that mark is what must keep it out of every live report, dashboard, analytics view, export, automated alert or integration, and the guard’s own activity list and search. An always-visible ", { size: 18, color: GREY }),
  t("Report an urgent incident", { size: 18, bold: true, color: GREY }),
  t(" action — a safety measure we added — lets a guard handle a real emergency at any moment, even without a microphone, and returns them to where they left off.", { size: 18, color: GREY }),
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
  ["Automatic start & refresher", "C", "It opens by itself the first time a guard signs in, before they can reach the home screen. A guard who has finished can reopen it any time as a refresher without losing their completed status."],
  ["Everyone, exactly once", "C", "Every active guard receives it one time — including guards who already had accounts before this feature. After finishing, their next sign-in goes straight to the home screen."],
  ["Cannot be skipped", "C", "There is no way to reach the home screen except by completing it. The only exception is the urgent-incident action, which is a detour, not a way to finish."],
  ["Microphone required, never a dead end", "C", "The guard must allow the microphone; declining re-prompts with guidance. If it genuinely cannot be enabled (broken hardware or company device policy), the guard reaches a support screen (the site’s support contact and a way to reach them) that still shows the urgent-incident action — confirmed when a blocked-microphone guard can leave the prompt and is never trapped."],
  ["Ask a question (practice only)", "P", "The guard asks SAM a spoken question and SAM answers. Confirmed correct when SAM gives an answer and no report is created by this step."],
  ["Report one practice incident", "P", "The guard reports the made-up incident and SAM creates exactly one practice report. Repeats or retries must never create extra reports."],
  ["Fix a report by voice", "C", "The guard tells SAM to change “Dell” to “MacBook”; the same report updates and no duplicate appears. (The exact way the app updates the report is a Check-with-engineering item.)"],
  ["Fix a report by hand", "C", "The guard long-presses their own message, edits it, and saves; the finished report contains both “MacBook” and “reception.” (Whether the app already supports this edit is a Check-with-engineering item.)"],
  ["Marked as training at creation", "C", "Every practice report is flagged as training (archive / soft-delete) the moment it is created — the platform already has this flag for exactly this purpose — so a report never exists un-marked."],
  ["Kept out of every live surface", "C", "The practice report must appear in no live surface — reports, dashboards, analytics, exports, automated alerts and integrations, and the guard’s own activity list and search. The filter that excludes training-marked data does not exist yet, so archived data is currently still visible in production; this exclusion filter must be built and verified before go-live. Until then the exercise runs only in the internal test environments."],
  ["Finishing is earned & remembered", "P", "Completion is recorded only after SAM sees the guard actually do each step — never a quiz or an “I’m done” button. It is stored on the server as version 1, issued once per guard, survives closing the app or switching devices, and cannot be duplicated."],
  ["Emergencies always get through", "P", "A safety measure we added, not an original requirement: a “Report an urgent incident” action is visible on every onboarding screen, works without a microphone, opens normal reporting, and returns the guard to where they left off. It does not count as finishing or skipping."],
  ["Visual aids come later", "C", "Extra visual help (tooltips, highlights) for dark or noisy sites is not built now; the need is judged during testing."],
];
kids.push(table([2340, 7308], ["Requirement", "What it means and how we confirm it"], REQ.map((r) => reqRow(r[0], r[1], r[2]))));

/* ---- 03 Practice data & completion (the two things that must not go wrong) ---- */
kids.push(eyebrow("03", "The two things that must not go wrong"));
kids.push(subhead("PRACTICE DATA NEVER REACHES LIVE OPERATIONS"));
kids.push(p([
  t("The made-up incident must never trigger a real response. Every practice report is marked as training the moment it is created — the platform already has this archive / soft-delete flag. "),
  t("The dependency to be clear about:", { bold: true }),
  t(" the filtering that actually keeps marked data out of live surfaces does not exist yet — today, archived items are still visible in production — so it is a required enhancement, not something that works today. This exercise therefore runs only in the internal test environments until that exclusion filter is built and verified, and must not go live in production before then. When built, the filter should work "),
  t("hide-by-default", { bold: true }),
  t(" at the shared data layer, so every current and future live surface excludes training data automatically rather than each one having to remember to. Permanent deletion of practice data is a separate database clean-up done later; only the guard’s completion status is kept."),
]));
kids.push(subhead("COMPLETION IS SERVER-SIDE, EARNED, AND SAFE UNDER PRESSURE"));
kids.push(p([
  t("Onboarding is marked complete only after the system observes the guard perform each real action, and that status lives on the server (never trusted from the phone alone). It is issued once per guard, resumes safely if the app closes or the guard switches devices, and two sessions at once can never create duplicate reports or roll progress backwards. If practice data cannot be safely hidden at any point, completion is withheld and the guard stays in the exercise — protecting live data always outranks finishing."),
], { after: 40 }));

/* ---- 04 Unhappy paths ---- */
kids.push(eyebrow("04", "Handling the unhappy paths"));
kids.push(table([2500, 7148], ["If this happens", "What the guard experiences"], [
  ["Microphone declined or unavailable", "SAM keeps asking with clear steps to enable it; a truly blocked microphone routes to support. The urgent-incident action still works without a microphone, so a real emergency is never blocked."],
  ["Network, speech, or save fails", "Progress is kept and a clear “try again” is shown; the current step and any edit are never silently lost."],
  ["A spoken correction is misheard", "SAM asks again; the guard can retry by voice or switch to the manual long-press edit instead."],
  ["A manual edit fails to save", "The failure is shown and the guard’s typed text is kept for another try; the report is never left half-changed."],
  ["Practice data cannot be safely hidden", "The step stops safely: completion is withheld and the guard stays in the exercise. Only a stage/error code is recorded — never the practice content."],
]));

/* ---- 05 Confirm before building ---- */
kids.push(eyebrow("05", "Confirm with engineering before building"));
kids.push(p("A few things depend on how the platform already behaves. They are gathered here so they are settled once, up front. None is expected to be a blocker.", { after: 60 }));
const conf = "conf";
[
  "Does the app already know when a guard signs in for the first time?",
  "Can this exercise be launched on its own, without the usual voice or tap triggers that start other scenarios?",
  "Can SAM run in a practice mode whose conversation and report never enter live reporting?",
  "What is the exact way the app updates an existing report — both when SAM does it by voice and when the guard edits their message by hand?",
  "The archive / training flag exists — confirm exactly which live feeds, reports, alerts, integrations, and views the new exclusion filter must cover, and where reports are created so the flag is set at creation.",
  "Is there already a place to store “this guard has completed onboarding,” or does one need to be added?",
].forEach((s) => kids.push(bullet(conf, s)));

/* ---- 06 Rollout, guardrails & done ---- */
kids.push(eyebrow("06", "Rollout, guardrails, and done"));
kids.push(p([
  t("Production go-live is "),
  t("blocked", { bold: true }),
  t(" until the exclusion filter that hides training data from every live surface is built and verified; until then this runs only in the internal test environments (DEV/UAT) with a dedicated test customer account (tenant) and test guards. Roll out behind a switch so it can be turned on gradually and turned off instantly, without erasing anyone’s completed status. This is universal for all guards — no customer-specific setup — and it does not change any existing templates. Nothing reaches production without Vincent Smeyers’ sign-off."),
], { after: 60 }));
kids.push(subhead("DONE MEANS"));
const done = "done";
[
  "A guard — new or existing — is shown the exercise once, completes it by doing the real steps, and lands on the home screen; reopening later keeps their completed status.",
  "The exclusion filter is built and the practice report is proven absent from every live surface, and a forced hide-failure withholds completion.",
  "The confirm-with-engineering items are answered, and Vincent Smeyers has signed off.",
].forEach((s) => kids.push(bullet(done, s)));

kids.push(new Paragraph({ spacing: { before: 90 },
  border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 5, space: 5 } },
  children: [t("This document describes intended behaviour for internal build and testing. No production change is made until Vincent Smeyers has signed off.", { size: 16, it: true, color: GREY })] }));

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
