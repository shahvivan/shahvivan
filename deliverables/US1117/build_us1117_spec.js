/*
 * Builder for User Story 1117 — SAM OnSite: User onboarding scenario
 * Developer-ready Word specification (US Letter, real Word numbering).
 *
 * Run:  node build_us1117_spec.js
 * Out:  US1117_SAM_OnSite_Guard_Onboarding_MVP_Spec.docx (same folder)
 *
 * Design notes:
 *  - US Letter portrait, 1" margins.
 *  - Calibri business typography, blue heading hierarchy.
 *  - Real Word numbering definitions (LevelFormat.BULLET / DECIMAL) — no
 *    literal bullet glyphs, no manual "1." text. Each restarting list gets a
 *    unique reference.
 *  - Tables use dual DXA widths (table columnWidths + per-cell width), repeated
 *    header rows, and cantSplit so rows never break across pages.
 *  - Evidence discipline: a [Verified] / [Proposed] / [To validate] tag opens
 *    requirement/decision rows so nothing reads as fact unless it is.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  Header, Footer, PageNumber, LevelFormat, TabStopType, TabStopPosition,
  Tab, VerticalAlign,
} = require("docx");

/* ---------- palette & type ---------- */
const BLUE = "1F4E79";       // Heading 1
const BLUE2 = "2E74B5";      // Heading 2
const BLUE3 = "5B9BD5";      // accents
const INK = "1A1A1A";        // body
const GREY = "595959";       // muted
const HEADER_FILL = "1F4E79";
const ZEBRA = "EEF3F9";      // light row band
const CALLOUT = "F2F7FC";    // callout background
const RULE = "BDD3EA";

const FONT = "Calibri";
const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 1440;
const CONTENT_W = PAGE_W - 2 * MARGIN; // 9360 DXA

/* ---------- evidence tags ---------- */
const TAG = {
  V: { label: "Verified", color: "1E7A46" },
  P: { label: "Proposed", color: "9A6700" },
  T: { label: "To validate", color: "9A3412" },
  R: { label: "Recommended", color: "6A4BA6" },
};

/* ---------- small helpers ---------- */
function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: opts.size || 21, color: opts.color || INK, bold: opts.bold || false, italics: opts.italics || false });
}
function body(children, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after == null ? 120 : opts.after, line: 264, lineRule: "auto" },
    alignment: opts.align,
    children: Array.isArray(children) ? children : [run(children)],
    keepNext: opts.keepNext || false,
  });
}
function tagRuns(tagKey) {
  const t = TAG[tagKey];
  if (!t) return [];
  return [new TextRun({ text: `[${t.label}] `, font: FONT, size: 18, bold: true, color: t.color })];
}
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 140 },
    keepNext: true,
    children: [new TextRun({ text, font: FONT, size: 30, bold: true, color: BLUE })],
    border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 8, space: 4 } },
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 90 },
    keepNext: true,
    children: [new TextRun({ text, font: FONT, size: 24, bold: true, color: BLUE2 })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 70 },
    keepNext: true,
    children: [new TextRun({ text, font: FONT, size: 21, bold: true, color: BLUE2 })],
  });
}
// bullet paragraph using a real numbering reference
function bullet(ref, children, level = 0) {
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { after: 80, line: 264, lineRule: "auto" },
    children: Array.isArray(children) ? children : [run(children)],
  });
}
function numItem(ref, children, level = 0) {
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { after: 80, line: 264, lineRule: "auto" },
    children: Array.isArray(children) ? children : [run(children)],
  });
}

/* ---------- table builders ---------- */
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const CELL_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "C9D6E5" };
function cellBorders() {
  return { top: CELL_BORDER, bottom: CELL_BORDER, left: CELL_BORDER, right: CELL_BORDER };
}
function tcell(content, opts = {}) {
  const children = Array.isArray(content) ? content : [
    new Paragraph({
      spacing: { after: 0, line: 252, lineRule: "auto" },
      alignment: opts.align,
      children: (Array.isArray(content) ? content : [content]).map((c) =>
        typeof c === "string" ? new TextRun({ text: c, font: FONT, size: opts.size || 19, bold: opts.bold || false, color: opts.color || INK }) : c
      ),
    }),
  ];
  return new TableCell({
    width: { size: opts.w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: opts.fill ? { type: ShadingType.CLEAR, color: "auto", fill: opts.fill } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    borders: cellBorders(),
    children,
  });
}
function headerCell(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, color: "auto", fill: HEADER_FILL },
    margins: { top: 70, bottom: 70, left: 90, right: 90 },
    borders: cellBorders(),
    children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text, font: FONT, size: 19, bold: true, color: "FFFFFF" })] })],
  });
}
// rich cell: array of TextRuns (for evidence-tagged content)
function richCell(runs, opts = {}) {
  return new TableCell({
    width: { size: opts.w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: opts.fill ? { type: ShadingType.CLEAR, color: "auto", fill: opts.fill } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    borders: cellBorders(),
    children: [new Paragraph({ spacing: { after: 0, line: 252, lineRule: "auto" }, children: runs })],
  });
}
// build a data table: colWidths[], header[], rows[] where each row is array of
// either string | {runs} | {t, tag}
function dataTable(colWidths, headers, rows, opts = {}) {
  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: headers.map((h, i) => headerCell(h, colWidths[i])),
  });
  const bodyRows = rows.map((cells, ri) =>
    new TableRow({
      cantSplit: true,
      children: cells.map((c, ci) => {
        const fill = ri % 2 === 1 ? ZEBRA : undefined;
        if (c && c.runs) return richCell(c.runs, { w: colWidths[ci], fill });
        if (c && typeof c === "object" && c.tag !== undefined) {
          const rr = [...tagRuns(c.tag), new TextRun({ text: c.t, font: FONT, size: 19, color: INK })];
          return richCell(rr, { w: colWidths[ci], fill });
        }
        return tcell(String(c == null ? "" : c), { w: colWidths[ci], fill, bold: ci === 0 && opts.boldFirst });
      }),
    })
  );
  return new Table({
    columnWidths: colWidths,
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: {
      top: CELL_BORDER, bottom: CELL_BORDER, left: CELL_BORDER, right: CELL_BORDER,
      insideHorizontal: CELL_BORDER, insideVertical: CELL_BORDER,
    },
    rows: [headerRow, ...bodyRows],
  });
}

/* ---------- callout box (single-cell table w/ a labelled header row for a11y) ---------- */
function callout(title, lines, accent = BLUE2) {
  const inner = [];
  lines.forEach((ln, i) => {
    inner.push(new Paragraph({
      spacing: { after: i === lines.length - 1 ? 0 : 70, line: 260, lineRule: "auto" },
      children: Array.isArray(ln) ? ln : [run(ln)],
    }));
  });
  return new Table({
    columnWidths: [CONTENT_W],
    width: { size: CONTENT_W, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      left: { style: BorderStyle.SINGLE, size: 18, color: accent },
      right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      insideHorizontal: NO_BORDER, insideVertical: NO_BORDER,
    },
    rows: [
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: [new TableCell({
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, color: "auto", fill: CALLOUT },
          margins: { top: 80, bottom: 40, left: 140, right: 140 },
          children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: title, font: FONT, size: 20, bold: true, color: accent })] })],
        })],
      }),
      new TableRow({
        cantSplit: true,
        children: [new TableCell({
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, color: "auto", fill: CALLOUT },
          margins: { top: 20, bottom: 90, left: 140, right: 140 },
          children: inner.length ? inner : [new Paragraph({ children: [] })],
        })],
        cantSplit: true,
      }),
    ],
  });
}

function spacer(h = 60) { return new Paragraph({ spacing: { after: h }, children: [] }); }

/* ================= CONTENT ================= */
const children = [];

/* ----- masthead ----- */
children.push(new Paragraph({
  spacing: { after: 20 },
  children: [new TextRun({ text: "PRONECT  •  SAM ONSITE", font: FONT, size: 18, bold: true, color: BLUE3, characterSpacing: 30 })],
}));
children.push(new Paragraph({
  spacing: { after: 30 },
  children: [new TextRun({ text: "User Story 1117 — Guard Onboarding Scenario", font: FONT, size: 40, bold: true, color: BLUE })],
}));
children.push(new Paragraph({
  spacing: { after: 120 },
  border: { bottom: { color: BLUE, style: BorderStyle.SINGLE, size: 12, space: 6 } },
  children: [new TextRun({ text: "Developer-ready MVP specification for the first-login voice training scenario", font: FONT, size: 22, italics: true, color: GREY })],
}));
// meta strip table
children.push(dataTable(
  [1560, 3120, 1560, 3120],
  ["Field", "Value", "Field", "Value"],
  [
    ["Story", "US 1117 — SAM OnSite: User onboarding scenario", "Area", "Pronect"],
    ["Iteration", "Pronect\\3.0.12", "State", "New"],
    ["Owner / requester", "Vivan Shah", "Prod. approval", "Vincent Smeyers"],
    ["Audience", "All active guards (new + existing), once", "Language", "English (MVP)"],
  ],
  { boldFirst: false }
));
children.push(spacer(40));
children.push(callout("How to read this document — evidence legend", [
  [new TextRun({ text: "[Verified] ", font: FONT, size: 20, bold: true, color: TAG.V.color }), run("Backed by the ticket, an explicit decision from Vivan, Vincent's written instruction, or read-only DEV/UAT inspection.")],
  [new TextRun({ text: "[Proposed] ", font: FONT, size: 20, bold: true, color: TAG.P.color }), run("A design choice recommended by this spec (fields, keys, timings, model). Sensible default — confirm during build.")],
  [new TextRun({ text: "[To validate] ", font: FONT, size: 20, bold: true, color: TAG.T.color }), run("Depends on platform behaviour not yet inspected. Listed in §16 Engineering validation before it is relied on.")],
  [new TextRun({ text: "[Recommended] ", font: FONT, size: 20, bold: true, color: TAG.R.color }), run("Added safety/quality scope beyond the literal ticket. Flagged, never disguised as an original acceptance criterion.")],
  [new TextRun({ text: "Note: ", font: FONT, size: 20, bold: true, color: GREY }), run("[Verified] certifies where a requirement comes from (its source or decision), not that the platform already implements it. Implementation feasibility for such requirements is tracked separately in §16.")],
], BLUE2));

/* ----- 1. Purpose & confirmed decisions ----- */
children.push(h1("1.  Purpose and confirmed product decisions"));
children.push(body([
  run("SAM OnSite is Pronect's voice-first guard assistant: a guard speaks naturally and SAM captures structured, time-stamped, geo-tagged activity and incident reports. This specification defines a "),
  run("mandatory, one-time voice onboarding scenario", { bold: true }),
  run(" that runs the first time a guard signs in, so the guard practises SAM's core actions — asking a question, reporting an incident, and correcting a report — by following SAM's spoken instructions, using safe fictional data."),
]));
children.push(body([
  run("Original story (from the ticket): "),
  run("“As a new security guard, I want to complete a simple, stage-by-stage voice training template upon my first login, so that I can quickly master core features — like asking questions, editing logs, and generating reports — by simply following SAM's voice instructions.”", { italics: true, color: GREY }),
]));
children.push(h2("1.1  Confirmed decisions (do not reopen)"));
const decRef = "dec-list";
const decisions = [
  ["V", "Audience — every eligible guard receives onboarding once, including existing guards. “New” guard does not mean only newly-created accounts."],
  ["V", "Mandatory — onboarding cannot be skipped."],
  ["V", "“Generate a report” means reporting a fictional incident by talking to SAM OnSite."],
  ["V", "Applicability — all guards. Configuration is universal, not client-specific."],
  ["V", "Refresher — a guard who has completed onboarding can reopen it later without clearing their completion."],
  ["V", "Language — English for the MVP."],
  ["V", "Enhanced visual support (tooltips / highlight animations) may be evaluated later for dark or noisy environments; not required for MVP delivery."],
  ["V", "Microphone denied — keep prompting the guard to enable microphone access."],
  ["V", "“Editing logs” means editing the guard's own incident report / conversation; the final report must reflect the correction."],
  ["P", "Completion logic — completion is inferred from the system observing the guard actually perform each action. The observed-action mechanism is proposed; the rule that completion must never be a quiz or a self-confirmation (“I'm done”) button is a hard requirement, not an option."],
];
decisions.forEach(([tag, txt]) => children.push(bullet(decRef, [...tagRuns(tag), run(txt)])));

children.push(h2("1.2  Worked correction example (from Vivan)"));
children.push(body([
  run("A guard reports a stolen laptop as a "), run("Dell", { bold: true }),
  run(", but it was actually a "), run("MacBook", { bold: true }),
  run(". Two correction paths must both work: (a) "), run("voice", { bold: true }),
  run(" — the guard tells SAM to change Dell to MacBook; and (b) "), run("manual", { bold: true }),
  run(" — the guard long-presses their own sent message, edits it, and saves. In both paths the same Activity/report must update to the corrected information — no duplicate report is created."),
]));

/* ----- 2. Scope ----- */
children.push(h1("2.  Scope"));
children.push(h2("2.1  In scope (MVP)"));
const inRef = "in-scope";
[
  "A universal, application-level onboarding flow that launches automatically after a guard's first successful login and is available later as a refresher.",
  "One guided pass through three core actions: ask SAM a question, report a fictional incident, and correct that report by voice and by manual long-press edit.",
  "Server-side, versioned completion tracking that issues version 1 exactly once to every active guard, new or existing.",
  "Isolation of all fictional practice data from every operational output for the entire life of the practice Activity.",
  "English voice and on-screen copy.",
].forEach((t) => children.push(bullet(inRef, t)));
children.push(h2("2.2  Out of scope (MVP)"));
const outRef = "out-scope";
[
  [ "P", "Languages other than English." ],
  [ "V", "Tooltips / highlight animations and other enhanced visual support — deferred for evaluation during testing." ],
  [ "P", "Client-specific or site-specific onboarding variants." ],
  [ "P", "Any change to a live client template pack or to a protected system template. Onboarding is delivered at the application layer and must not modify packs." ],
  [ "P", "Physical deletion / database cleansing of practice data — a later, separate task from the logical archive used here." ],
].forEach(([tag, t]) => children.push(bullet(outRef, [...tagRuns(tag), run(t)])));

/* ----- 3. Roles & ownership ----- */
children.push(h1("3.  Roles and implementation ownership"));
children.push(body("Ownership is called out so no requirement is orphaned across teams. Names reflect the working group; assignment is a coordination aid, not a contractual RACI."));
children.push(dataTable(
  [1900, 3260, 4200],
  ["Layer", "Owner (indicative)", "Responsibility in this story"],
  [
    ["Mobile app", "SAM OnSite (Jordi)", "First-login detection, onboarding launch/resume, mic-permission loop, long-press manual edit, Home hand-off."],
    ["SAM / AI", "AI (Sartak)", { runs: [...tagRuns("T"), new TextRun({ text: "Scenario prompts in a non-reporting practice context (mode existence to confirm, §16); structured-essential capture and correction (context_set) so a correction updates the same Activity.", font: FONT, size: 19, color: INK })] }],
    ["Backend / data", "Backend / Supabase (Berk)", "Onboarding-progress record; atomic archive/training flag at Activity creation; idempotency and concurrency; completion finalisation."],
    ["Reporting / integrations", "Backend + Reporting", "Reporting-layer exclusion of archived training records across every downstream consumer."],
    ["Product / approval", "Vivan Shah; Vincent Smeyers", "Product decisions and acceptance; Vincent approves any production step."],
    ["QA", "QA", "Execute the UAT matrix, including negative and concurrency tests, in DEV/UAT."],
  ]
));

/* ----- 4. End-to-end flow ----- */
children.push(h1("4.  End-to-end user flow"));
children.push(body([
  ...tagRuns("P"),
  run("Target duration is approximately 3–5 minutes. This is a product proposal, not a measured fact; completion is gated on observed actions, not on elapsed time."),
]));
const flowRef = "flow-list";
[
  "After successful authentication the app checks the guard's onboarding version status.",
  "Any active guard who has not completed onboarding version 1 enters the mandatory flow (new and existing guards alike).",
  "A welcome screen explains that the exercise is mandatory, fictional, short, and that progress is saved.",
  "The guard enables microphone access. If denied, the app keeps prompting with clear guidance and the guard cannot proceed until access is granted.",
  "The system verifies at least one successful voice transcription before crediting the speaking steps.",
  "Using Talk, the guard asks SAM a question (reference prompt: “How do I report an incident?”). This step creates no Activity.",
  "The guard reports a fictional incident (reference: “A Dell laptop was stolen”). SAM may ask for any missing essentials and produces exactly one practice Activity.",
  "The guard corrects Dell to MacBook by voice. The same Activity updates in place — no duplicate is created.",
  "The guard long-presses their own sent message and changes it to “A MacBook was stolen at reception.” The final Activity must contain both “MacBook” and “reception”.",
  "Throughout, the practice Activity is excluded from every operational consumer (it was flagged as training the moment it was created).",
  "Only after every required action succeeds and safe isolation is confirmed is onboarding version 1 marked completed (server-side).",
  "The guard lands on the normal Home screen (Talk / Capture / Type).",
  "A completed guard may voluntarily reopen the training later; reopening never clears the completion record.",
].forEach((t) => children.push(numItem(flowRef, t)));
children.push(body([
  ...tagRuns("P"),
  run("Message vs. Activity: the guard's sent message and the structured Activity are linked but distinct. A voice correction updates the Activity's structured essentials; the manual long-press edits the guard's message text; on save the two reconcile so the final Activity reflects the latest of each. Whether the displayed message text also re-renders after a voice correction is a UX detail to confirm (§16); the acceptance target is the final Activity content, not the intermediate message string."),
]));
children.push(spacer(20));
children.push(callout("Fail-closed rule (at creation and at completion)", [
  [new TextRun({ text: "If the platform cannot create the practice Activity with its isolation flag in a single atomic write, the practice Activity ", font: FONT, size: 21, color: INK }),
   new TextRun({ text: "must not be created", font: FONT, size: 21, bold: true, color: INK }),
   new TextRun({ text: " — it is never created first and flagged afterwards. And if, at any later point, isolation cannot be guaranteed, the stage fails closed: onboarding is ", font: FONT, size: 21, color: INK }),
   new TextRun({ text: "not", font: FONT, size: 21, bold: true, color: INK }),
   new TextRun({ text: " marked complete and the guard is kept in the flow. Safety of operational data outranks finishing the exercise.", font: FONT, size: 21, color: INK })],
], TAG.T.color));

/* ----- 5. Voice & screen script ----- */
children.push(h1("5.  English voice and screen script"));
children.push(body([
  ...tagRuns("P"),
  run("Reference copy. Wording is a starting point for product/UX. Where completion is judged by final system state, the guard need only demonstrate the intended action — semantically-equivalent English is accepted (see §6 for the exact matching rule per step)."),
]));
children.push(dataTable(
  [1500, 3930, 3930],
  ["Stage", "SAM says (voice + screen)", "Guard does"],
  [
    ["Welcome", "“Welcome to SAM. This quick practice is required and uses made-up information. It takes a few minutes and your progress is saved.”", "Reads, taps Start."],
    ["Microphone", "“I need your microphone to hear you. Please enable microphone access to continue.”", "Grants mic permission (re-prompted until granted)."],
    ["Ask a question", "“Try asking me something. For example: ‘How do I report an incident?’”", "Speaks a question to SAM. No report is created."],
    ["Report incident", "“Now report a practice incident. For example: ‘A Dell laptop was stolen.’ I may ask for a few details.”", "Reports the fictional incident; answers any follow-up. One practice Activity is created."],
    ["Voice correction", "“Mistakes happen. Tell me to change Dell to MacBook.”", "Says, e.g., “Change the Dell to a MacBook.” Same Activity updates."],
    ["Manual edit", "“You can also fix it yourself. Long-press your message, edit it, and save.”", "Long-presses own message; sets it to “A MacBook was stolen at reception.” and saves."],
    ["Finish", "“Great work. That's everything — here's your home screen.”", "Continues to Home once completion is confirmed."],
  ]
));

/* ----- 6. Functional requirements ----- */
children.push(h1("6.  Functional requirements"));
children.push(body("Each requirement carries an evidence tag. FR IDs are referenced by the acceptance criteria (§10), the UAT matrix (§11), and the traceability matrix (§12)."));
children.push(body([
  ...tagRuns("P"),
  run("Completion-matching rule: a step is credited on the guard's observed action and the resulting system state, not on exact phrasing. Speaking steps accept semantically-equivalent English; correction steps are judged by the final Activity content (e.g. it contains “MacBook” and “reception”), not by a fixed sentence. Only steps that genuinely require exact wording would state so explicitly — none do in this MVP."),
]));
const FR = [
  ["FR-01", "V", "Trigger", "Onboarding launches automatically immediately after a guard's first successful login, and can be reopened later as a refresher."],
  ["FR-02", "V", "Audience", "Every active guard receives onboarding version 1 exactly once, including guards whose accounts already existed before this feature."],
  ["FR-03", "V", "Mandatory", "The guard cannot skip or dismiss onboarding to reach Home; it must be completed (subject to FR-17)."],
  ["FR-04", "V", "Refresher", "A completed guard can voluntarily reopen the training; reopening does not clear or regress the completion record."],
  ["FR-05", "V", "Language", "All onboarding voice and screen copy is English for the MVP."],
  ["FR-06", "V", "Microphone", "The guard must enable microphone access. If denied, the app keeps prompting with actionable guidance and does not advance the speaking steps until access is granted. If the microphone cannot be granted (hardware fault or device/MDM policy), the guard is routed to support/admin rather than looped indefinitely, and the urgent-incident action (FR-17) stays reachable throughout."],
  ["FR-07", "P", "Transcription check", "Before crediting a speaking step, the system verifies at least one successful voice transcription; an empty or failed transcription does not credit the step (distinct from microphone-denied). Traces to AC-13 / UAT-17."],
  ["FR-08", "P", "Ask-a-question step", "The “ask SAM a question” step runs in a practice (non-reporting) context and creates no Activity."],
  ["FR-09", "P", "One practice Activity", "The “report an incident” step creates exactly one practice Activity. Retries or repeated messages must not create additional Activities."],
  ["FR-10", null, "Voice correction", [
    ["V", "Vivan requires that a spoken correction (Dell → MacBook) update the essentials of the same Activity, and SAM can write corrected essentials via context_set. "],
    ["T", "That the correction mutates the same Activity record with no duplicate emitted downstream is a runtime behaviour to confirm (§16)."],
  ]],
  ["FR-11", null, "Manual correction", [
    ["V", "Vivan requires that long-pressing the guard's own sent message allow a manual edit that, when saved, updates the same Activity so the final Activity contains both “MacBook” and “reception”. "],
    ["T", "Whether a guard-app long-press edit route exists (the DEV Edit surface is the back-office reporting console) is to confirm (§16)."],
  ]],
  ["FR-12", "P", "Isolation at creation", "The practice/training exclusion flag is applied atomically when the Activity is created — before any outbox, event, or downstream processing — not added afterwards. If atomic create-with-flag is not supported, the Activity is not created (fail closed at create)."],
  ["FR-13", null, "Operational exclusion", [
    ["V", "Vincent's instruction: the practice data is archived (soft-delete) and the reporting layer is updated to exclude archived training records — the agreed exclusion approach. "],
    ["R", "Strengthened: exclusion is enforced at the lowest shared data-access layer as a default-exclude, so every current or future consumer — reporting, dashboards, analytics, KPIs, exports, alerts, webhooks, integrations, event outbox, and the guard's own in-app views — inherits it unless it deliberately opts in."],
  ]],
  ["FR-14", "P", "Completion integrity", "Completion is recorded server-side and versioned (per user + onboarding key + version). It is set only after every required action is observed and isolation is confirmed. No client-only “completed” flag is trusted."],
  ["FR-15", "P", "Fail-closed", "If safe isolation of practice data cannot be guaranteed at any step — at creation or later — the stage fails and completion is not marked. The isolation-failure signal carries only stage/error codes, never fictional content."],
  ["FR-16", "V", "Hand-off", "On confirmed completion the guard reaches the normal Home screen (Talk / Capture / Type)."],
  ["FR-17", "R", "Urgent-incident path", "A narrowly-scoped “Report an urgent incident” action is reachable at every onboarding stage — including the welcome and microphone-permission screens — always visible, and usable without microphone access (it routes to normal incident reporting via Capture / Type as well as voice). It does not mark onboarding complete or skipped, and returns the guard to the saved stage afterward. Added safety scope, not an original acceptance criterion."],
  ["FR-18", "P", "Idempotency / concurrency", "Exactly one onboarding record and one active practice-Activity reference per user + key + version. Retries do not duplicate (contingent on the create endpoint supporting an idempotency key or uniqueness constraint — see §16); the newest valid server-side progress wins with no stage regression; the completion operation is idempotent."],
  ["FR-19", "P", "Privacy / telemetry", "Onboarding telemetry records stage and error codes and completion metrics only. It does not store microphone-permission state beyond need, does not log raw voice/transcript content unless existing policy allows, discards the practice session's raw audio after transcription unless policy requires retention, and never routes fictional incident content into operational analytics."],
  ["FR-20", "V", "Deferred visual support", "Tooltips / highlight animations for dark or noisy environments are deferred; the need is evaluated during testing, not built for MVP."],
];
children.push(dataTable(
  [900, 1750, 6710],
  ["ID", "Requirement", "Detail"],
  FR.map(([id, tag, name, detail]) => {
    let detailRuns;
    if (Array.isArray(detail)) {
      detailRuns = [];
      detail.forEach(([tk, txt]) => {
        detailRuns.push(...tagRuns(tk), new TextRun({ text: txt, font: FONT, size: 19, color: INK }));
      });
    } else {
      detailRuns = [...tagRuns(tag), new TextRun({ text: detail, font: FONT, size: 19, color: INK })];
    }
    return [
      id,
      { runs: [new TextRun({ text: name, font: FONT, size: 19, bold: true, color: INK })] },
      { runs: detailRuns },
    ];
  })
));

/* ----- 7. Progress & completion model ----- */
children.push(h1("7.  Progress and completion model"));
children.push(body([
  ...tagRuns("P"),
  run("The following record is a recommended shape, not observed database schema. No onboarding-completion field was visible in the supplied user form, so field and key names must be confirmed during build (§16)."),
]));
children.push(dataTable(
  [2600, 2100, 4660],
  ["Field", "Example", "Purpose"],
  [
    ["user_id", "—", "Guard the record belongs to."],
    ["onboarding_key", "sam_onsite_guard (proposed)", "Identifies this onboarding; allows future onboardings without collision."],
    ["version", "1", "Version issued. Existing + new guards get v1 once."],
    ["status", "not_started / in_progress / completed", "Server-side state; never trusted from the client alone."],
    ["current_stage", "e.g. voice_correction", "Resume point after app close or device change."],
    ["started_at / updated_at / completed_at", "timestamps", "Lifecycle timing and idempotency support."],
  ]
));
children.push(h2("7.1  Integrity rules"));
const intRef = "integrity";
[
  [ "P", "Completion is server-authoritative and versioned; a client-side flag alone never marks completion." ],
  [ "P", "Version 1 is issued exactly once per guard. Re-running the refresher does not re-issue or clear v1." ],
  [ "P", "Each stage advances only on observed system evidence (e.g. a real transcription, a real Activity edit) — not a quiz or an “I'm done” tap." ],
  [ "P", "Resume is safe across app closure and device change: the newest valid server-side progress wins and stages never regress." ],
  [ "P", "Concurrent sessions cannot create duplicate practice Activities or duplicate onboarding records; the completion call is idempotent." ],
].forEach(([tag, t]) => children.push(bullet(intRef, [...tagRuns(tag), run(t)])));

/* ----- 8. Practice Activity lifecycle & isolation ----- */
children.push(h1("8.  Practice Activity lifecycle and data isolation"));
children.push(body([
  run("Vincent's instruction: put the fictional data in "),
  run("archive mode (soft delete)", { bold: true }),
  run(", update the reporting layer to exclude it (the flag exists for exactly that purpose), and do any physical database cleansing later. This section strengthens that into a lifecycle where fictional data can never be operationally visible, "),
  run("even temporarily", { bold: true, italics: true }),
  run("."),
]));
children.push(dataTable(
  [1700, 3400, 4260],
  ["Phase", "State", "Isolation guarantee"],
  [
    [{ runs: [new TextRun({ text: "Create", font: FONT, size: 19, bold: true, color: INK })] }, "Practice Activity created with archive/training flag set atomically in the same write.", { runs: [...tagRuns("P"), new TextRun({ text: "Flag precedes any outbox/event/downstream step. No window exists in which the record is visible without the flag.", font: FONT, size: 19, color: INK })] }],
    [{ runs: [new TextRun({ text: "Edit", font: FONT, size: 19, bold: true, color: INK })] }, "Activity remains editable during onboarding (voice + manual correction).", { runs: [...tagRuns("P"), new TextRun({ text: "Flag persists through every edit and regeneration.", font: FONT, size: 19, color: INK })] }],
    [{ runs: [new TextRun({ text: "Complete", font: FONT, size: 19, bold: true, color: INK })] }, "Completion finalises the archive state and records onboarding completion.", { runs: [...tagRuns("P"), new TextRun({ text: "If isolation cannot be confirmed, completion is withheld (fail-closed, FR-15).", font: FONT, size: 19, color: INK })] }],
    [{ runs: [new TextRun({ text: "Post", font: FONT, size: 19, bold: true, color: INK })] }, "Record stays logically archived; physical cleansing is a later, separate task.", { runs: [...tagRuns("V"), new TextRun({ text: "Reporting-layer exclusion is the agreed approach (Vincent). ", font: FONT, size: 19, color: INK }), ...tagRuns("R"), new TextRun({ text: "Extended to every other consumer via default-exclude below.", font: FONT, size: 19, color: INK })] }],
  ]
));
children.push(spacer(20));
children.push(callout("Enforcement: default-exclude, not an allowlist", [
  [run("Exclusion is enforced at the "), run("lowest shared data-access / query layer as a default-exclude", { bold: true }), run(", so any current or future consumer inherits it unless it deliberately opts in. The enumerated list below is verification "), run("coverage", { italics: true }), run(", not the enforcement mechanism — an unlisted or newly-added consumer must still not see practice data.")],
], TAG.R.color));
children.push(spacer(16));
children.push(callout("Consumers and vectors verified to ignore practice data", [
  "Reporting views and events pages • dashboards and KPIs • analytics and metrics • exports and scheduled reports • alerts and notifications • webhooks and integrations • the event outbox and any fan-out queue.",
  "The guard's own in-app surfaces: Home recent-activity list, the Activity tab, and the search-activity field — the practice incident must not appear in the guard's own history after completion.",
  "Derived stores: read replicas, caches, and search indexes must also exclude it; and because an already-dispatched webhook or push notification cannot be un-sent, nothing may be emitted before the isolation flag is set. The SAM conversation/transcript store is a separate vector from the Activity record and is isolated too.",
  "Exclusion is verified against all of the above before production (§13 rollout and §15 gates).",
], TAG.V.color));

/* ----- 9. Error & recovery ----- */
children.push(h1("9.  Error and recovery behaviour"));
children.push(dataTable(
  [3000, 6360],
  ["Condition", "Behaviour"],
  [
    ["Microphone denied", "Keep prompting with clear guidance on how to enable access; speaking steps do not advance until granted. If the microphone genuinely cannot be granted (hardware fault or device/MDM policy), route the guard to support/admin instead of an infinite loop. The urgent-incident action stays visible and works without a microphone (Capture / Type), so a real emergency is never blocked by the mic gate."],
    ["Network / speech / save failure", "Preserve progress; show a retryable error; never lose the current stage or silently drop the guard's edit."],
    ["Voice correction not understood", "SAM re-asks; the guard may retry by voice or use the manual long-press edit path instead."],
    ["Manual save failure", "Surface the failure and retain the guard's typed text for retry; the Activity is not left half-edited."],
    ["Isolation cannot be confirmed", "Fail closed: hold completion, keep the guard in the flow, raise an isolation-failure signal for engineering (no sensitive content logged)."],
    ["Urgent real incident during onboarding", "The recommended urgent-incident action (FR-17) opens normal reporting and returns to the saved stage; completion is neither granted nor skipped."],
  ]
));

/* ----- 10. Acceptance criteria ----- */
children.push(h1("10.  Acceptance criteria"));
children.push(body("Given/When/Then criteria. Each maps to at least one functional requirement and one UAT test (§11)."));
const AC = [
  ["AC-01", "FR-01", "Given an active guard who has not completed onboarding v1, when they log in successfully, then onboarding launches automatically before Home is reachable."],
  ["AC-02", "FR-02", "Given an existing guard (account created before this feature) who has not completed v1, when they log in, then they receive onboarding v1 exactly once."],
  ["AC-03", "FR-03", "Given onboarding is in progress, when the guard attempts to skip, then the only route to Home is completion. The urgent-incident action (AC-12) is a temporary detour that returns the guard to the saved stage — it is not an alternative exit to Home."],
  ["AC-04", "FR-06", "Given microphone permission is denied, when the guard tries to proceed, then the app re-prompts with guidance and the speaking steps do not advance."],
  ["AC-05", "FR-08/FR-09", "Given the guard asks SAM a question, then no Activity is created; given the guard then reports the fictional incident, then exactly one practice Activity exists."],
  ["AC-06", "FR-10", "Given a practice Activity naming a Dell, when the guard says to change it to a MacBook, then the same Activity is updated and no second Activity is created."],
  ["AC-07", "FR-11", "Given the guard long-presses their own message and saves “A MacBook was stolen at reception,” then the final Activity contains both “MacBook” and “reception”."],
  ["AC-08", "FR-12/FR-13", "Given a practice Activity exists at any lifecycle phase, then it appears in no operational consumer, with the exclusion flag set from creation."],
  ["AC-09", "FR-14/FR-15", "Given all required actions succeeded and isolation is confirmed, when completion runs, then v1 is marked complete server-side; if isolation is not confirmed, completion is withheld."],
  ["AC-10", "FR-04/FR-16", "Given a completed guard, when they finish, then they reach Home; and when they reopen the training later, then completion is preserved."],
  ["AC-11", "FR-18", "Given duplicate concurrent sessions or retried messages, then no duplicate onboarding record or practice Activity is created and progress does not regress."],
  ["AC-12", "FR-17", "Given a real urgent incident during onboarding — including while microphone permission is denied — when the guard uses the always-visible urgent-incident action, then normal reporting opens and is usable without a microphone, and on return the saved stage resumes with completion neither granted nor skipped."],
  ["AC-13", "FR-07", "Given the microphone is granted but a speaking step produces no successful transcription, then the step is not credited and the guard is re-prompted to speak (no false completion)."],
];
children.push(dataTable(
  [900, 1500, 6960],
  ["ID", "Traces to", "Criterion"],
  AC.map(([id, fr, txt]) => [id, fr, txt])
));

/* ----- 11. UAT matrix ----- */
children.push(h1("11.  UAT matrix (DEV / UAT)"));
children.push(body("Positive and negative tests. Every acceptance criterion is covered; negative tests exist for duplication, archive/isolation failure, leakage (including the guard's own in-app views), microphone denial, transcription failure, network failure, manual-save failure, voice-correction failure, concurrent devices, the mic-gate urgent path, and refresher behaviour."));
const UAT = [
  ["UAT-01", "AC-01", "New guard first login triggers onboarding automatically.", "Onboarding shown before Home; progress record created."],
  ["UAT-02", "AC-02", "Existing guard (pre-feature account) is issued v1 once.", "Onboarding runs once; second login after completion goes straight to Home."],
  ["UAT-03", "AC-03", "Attempt to skip onboarding.", "No skip path to Home; only completion or urgent-incident path exits."],
  ["UAT-04", "AC-04", "Deny microphone, then attempt to continue. (negative)", "Re-prompt with guidance; speaking steps blocked until granted."],
  ["UAT-05", "AC-05", "Ask a question, then report the incident.", "No Activity from the question; exactly one practice Activity from the report."],
  ["UAT-06", "AC-06", "Voice-correct Dell to MacBook.", "Same Activity updated; no duplicate Activity."],
  ["UAT-07", "AC-06", "Voice correction misheard/failed, then retry. (negative)", "SAM re-asks; retry or manual edit succeeds; still one Activity."],
  ["UAT-08", "AC-07", "Long-press manual edit to “MacBook … reception” and save.", "Final Activity contains “MacBook” and “reception”."],
  ["UAT-09", "AC-07", "Manual save fails, then retry. (negative)", "Error shown; typed text retained; Activity not left half-edited."],
  ["UAT-10", "AC-08", "Inspect every operational consumer for the practice Activity. (negative)", "Absent from reporting, dashboards, analytics, exports, alerts, webhooks, outbox at all phases."],
  ["UAT-11", "AC-08/AC-09", "Force isolation/archive failure at creation. (negative)", "Activity never operationally visible; completion withheld (fail-closed)."],
  ["UAT-12", "AC-09/AC-10", "Complete all actions with isolation confirmed.", "v1 marked complete server-side; guard reaches Home."],
  ["UAT-13", "AC-11", "Duplicate concurrent sessions / retried messages. (negative)", "One onboarding record, one practice Activity; no progress regression."],
  ["UAT-14", "AC-10/AC-11", "Network drop mid-stage, then resume on another device. (negative)", "Progress preserved; resumes at saved stage; newest valid progress wins; no data loss."],
  ["UAT-15", "AC-10", "Reopen training after completion (refresher).", "Training reopens; completion record preserved, not re-cleared."],
  ["UAT-16", "AC-12", "Deny microphone, then trigger the urgent-incident path. (negative)", "Urgent action visible at the mic gate; normal reporting opens and is usable without a mic; onboarding neither completed nor skipped; returns to saved stage."],
  ["UAT-17", "AC-13", "Mic granted, but a speaking step yields no/failed transcription. (negative)", "Step not credited; guard re-prompted to speak; no false completion."],
  ["UAT-18", "AC-08", "Confirm practice Activity is absent from the guard's own in-app Home recent-activity, Activity tab, and search. (negative)", "Not present in any of the guard's own in-app views after completion."],
];
children.push(dataTable(
  [900, 1150, 3760, 3550],
  ["ID", "Verifies", "Test", "Expected evidence"],
  UAT.map(([id, ac, test, ev]) => [id, ac, test, ev])
));

/* ----- 12. Traceability matrix ----- */
children.push(h1("12.  Requirement traceability matrix"));
children.push(body("Ties each source requirement to its functional requirement, acceptance criterion, UAT test, and the evidence QA should capture. This is the single view that proves the ticket is fully covered."));
children.push(body([
  run("Carve-out: FR-05 (English), FR-19 (privacy/telemetry), and FR-20 (deferred visuals) are cross-cutting or non-functional. They carry no single acceptance criterion by design — FR-05 is verified across the whole UAT set, FR-19 is verified via UAT-10/UAT-18, and FR-20 is a deferral decision with nothing to build for MVP. Every other functional requirement maps to at least one acceptance criterion and one UAT test.", { italics: true, color: GREY }),
]));
const TRACE = [
  ["Ticket — first-login trigger + refresher", "FR-01", "AC-01", "UAT-01", "Onboarding auto-shown; refresher reopens"],
  ["Vivan — existing + new guards, once", "FR-02", "AC-02", "UAT-02", "v1 issued once per guard"],
  ["Vivan — mandatory", "FR-03", "AC-03", "UAT-03", "No skip path; urgent path is a detour"],
  ["Vivan — refresher keeps completion", "FR-04", "AC-10", "UAT-15", "Completion preserved on reopen"],
  ["Vivan — English MVP", "FR-05", "— (carve-out)", "UAT-01–18", "All copy English"],
  ["Vivan — mic denied keeps prompting", "FR-06", "AC-04", "UAT-04", "Re-prompt; steps blocked; support route"],
  ["Strengthened — transcription check", "FR-07", "AC-13", "UAT-17", "Step credited only on real transcription"],
  ["Ticket — ask a question", "FR-08", "AC-05", "UAT-05", "No Activity from question"],
  ["Ticket/Vivan — generate report = report fictional incident", "FR-09", "AC-05", "UAT-05", "Exactly one practice Activity"],
  ["Vivan — voice correction (Dell→MacBook)", "FR-10", "AC-06", "UAT-06/07", "Same Activity updated; no duplicate"],
  ["Vivan — manual long-press edit", "FR-11", "AC-07", "UAT-08/09", "Final has MacBook + reception"],
  ["Vincent — archive/soft-delete + reporting exclusion", "FR-12/13", "AC-08", "UAT-10/11/18", "Absent from all consumers + guard's own views"],
  ["Codex — completion by observed action", "FR-14", "AC-09", "UAT-12", "Server-side completion"],
  ["Strengthened — fail-closed isolation (create + complete)", "FR-15", "AC-09", "UAT-11", "Completion withheld on failure"],
  ["App screenshot — Home hand-off", "FR-16", "AC-10", "UAT-12", "Reaches Home"],
  ["Recommended — urgent-incident safety (mic-independent)", "FR-17", "AC-12", "UAT-16", "Reporting opens without mic; stage resumes"],
  ["Strengthened — idempotency/concurrency", "FR-18", "AC-11", "UAT-13/14", "No duplicates; no regression"],
  ["Strengthened — privacy/telemetry", "FR-19", "— (carve-out)", "UAT-10/18", "No fictional data in analytics or guard views"],
  ["Ticket/Vivan — visual support deferred", "FR-20", "— (carve-out)", "—", "Evaluated in testing"],
];
children.push(dataTable(
  [2900, 1150, 1050, 1600, 2660],
  ["Source requirement", "FR", "AC", "UAT", "Evidence QA captures"],
  TRACE.map((r) => r)
));

/* ----- 13. Rollout & rollback ----- */
children.push(h1("13.  Rollout and rollback"));
const rollRef = "roll-list";
[
  [ "P", "Release behind a feature flag (or equivalent controlled rollout) so automatic launch can be enabled per environment/cohort." ],
  [ "P", "Test first with a dedicated Pronect test client and test guards in DEV, then UAT." ],
  [ "V", "Current DEV pack observed: BLECKMANN-2306-v3 (2), 76 templates, 1 client — evidence only, not a migration target." ],
  [ "V", "No production step without Vincent Smeyers' approval." ],
  [ "V", "Verify the archive filter against every downstream consumer (§8 callout) before enabling in production." ],
  [ "P", "Rollback / disable: turn off automatic launch if a critical failure occurs, without clearing valid completion records." ],
  [ "P", "Monitor completion-failure and isolation-failure rates; alert on isolation failures. Do not log raw voice/report content in monitoring beyond existing policy." ],
].forEach(([tag, t]) => children.push(bullet(rollRef, [...tagRuns(tag), run(t)])));

/* ----- 14. Definition of Done ----- */
children.push(h1("14.  Definition of Done"));
const dodRef = "dod-list";
[
  "All acceptance criteria (AC-01–AC-13) pass in UAT, including every negative test in §11.",
  "Version 1 issued exactly once to a sample of existing and new guards; refresher preserves completion.",
  "Practice data proven absent from every downstream consumer at create, edit, complete, and post phases.",
  "Completion is server-side, versioned, idempotent, and resilient to concurrent sessions and device change.",
  "Fail-closed verified: forced isolation failure withholds completion.",
  "Engineering-validation items (§16) resolved or explicitly accepted before production.",
  "Vincent Smeyers has approved the production step.",
].forEach((t) => children.push(bullet(dodRef, t)));

/* ----- 15. Release-safety gate summary (maps to reviewer gates) ----- */
children.push(h1("15.  Release-safety gates"));
children.push(body("This spec is structured to satisfy the reviewer gates. Summary of where each is met:"));
children.push(dataTable(
  [1650, 7710],
  ["Gate", "Where satisfied"],
  [
    ["Ticket fidelity", "§1 decisions, §4 flow, §6 FRs, §12 traceability — every ticket line and Vivan clarification mapped."],
    ["No unsupported claims", "Evidence legend (source ≠ feasibility) + per-row tags; §7/§16 keep fields, keys, routes as proposed / to validate; FR-10/11/13 split verified decision from unproven runtime."],
    ["Practice-data safety", "§8 isolation at creation, default-exclude across all consumers + guard's own views, fail-closed at create and complete (FR-12/13/15)."],
    ["Completion integrity", "§7 server-side, versioned, once-per-guard, idempotent, concurrency-safe (FR-14/18)."],
    ["UX & operational safety", "§9 recovery, FR-06 mic loop + support route, FR-17 mic-independent urgent path, FR-20 deferred visuals."],
    ["Engineering usefulness", "§3 ownership, §16 one validation list, no pack/system-template change (§2.2)."],
    ["Acceptance & traceability", "Every functional FR→AC→UAT (non-functional carve-out documented); negative tests enumerated (§10–12)."],
    ["Release safety", "§13 DEV/UAT-first, Vincent approval, all-consumer filter check, rollback, monitoring."],
  ]
));

/* ----- 16. Engineering validation ----- */
children.push(h1("16.  Engineering validation before implementation"));
children.push(body([
  ...tagRuns("T"),
  run("These items depend on platform behaviour not provable from the supplied sources. They are collected here so they are validated once, up front — they are not necessarily product blockers, and none should be invented into the spec as fact."),
]));
const valRef = "val-list";
[
  "Is a guard's first successful login already stored anywhere the app can read?",
  "Can the app start a SAM scenario directly by template code, without speech classification or NFC?",
  "Can a SAM conversation run in a non-reporting practice mode (so the question step and practice report never enter operational reporting)?",
  "Which app routes / services / APIs perform (a) SAM-driven correction of an existing Activity and (b) a saved manual long-press edit?",
  "What is the exact archive/training field and the create/edit API endpoint? (Confirmed to exist behaviourally via UAT reporting Edit/Delete and Vincent's archive instruction; exact names not inspected — Azure Repos access was unavailable.)",
  "Does the onboarding-progress record already exist, or must it be added? No completion field was visible in the supplied user form.",
].forEach((t) => children.push(bullet(valRef, t)));
children.push(spacer(20));
children.push(callout("Evidence basis (what is proven today)", [
  "UAT reporting (uat.pronect-it.com/platform-event) exposes Edit and Delete on “Sam OnSite Activities”; Vincent confirmed Delete corresponds to soft-delete/archive. The Bleckmann pack exposes context_get / context_set / checklist_advance / procedure_advance / reclassify_scenario / clarification_needed, and structured essentials — including corrections after a scenario is otherwise done — can be written through context_set.",
  "No supplied pack tool covers login state, app navigation, onboarding completion, or reporting suppression; those remain application-layer work (the questions listed above). No production change was made during inspection.",
], TAG.V.color));

/* ----- footer note ----- */
children.push(spacer(40));
children.push(new Paragraph({
  spacing: { before: 120 },
  border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 6, space: 6 } },
  children: [new TextRun({ text: "Nothing in this document authorises a production change. All work proceeds in DEV/UAT and reaches production only after Vincent Smeyers' approval.", font: FONT, size: 18, italics: true, color: GREY })],
}));

/* ================= NUMBERING ================= */
// unique bullet reference per restarting list + one decimal for the flow list
const bulletRefs = [
  "dec-list", "in-scope", "out-scope", "integrity", "roll-list", "dod-list", "val-list",
];
const numberingConfig = bulletRefs.map((reference) => ({
  reference,
  levels: [{
    level: 0,
    format: LevelFormat.BULLET,
    text: "•",
    alignment: AlignmentType.LEFT,
    style: { run: { font: FONT, size: 21, color: BLUE2 }, paragraph: { indent: { left: 720, hanging: 360 } } },
  }],
}));
// decimal, restarting lists
["flow-list"].forEach((reference) => {
  numberingConfig.push({
    reference,
    levels: [{
      level: 0,
      format: LevelFormat.DECIMAL,
      text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { run: { font: FONT, size: 21, bold: true, color: BLUE2 }, paragraph: { indent: { left: 720, hanging: 360 } } },
    }],
  });
});

/* ================= DOCUMENT ================= */
const doc = new Document({
  creator: "Pronect",
  title: "US1117 SAM OnSite Guard Onboarding MVP Specification",
  description: "Developer-ready specification for User Story 1117",
  styles: {
    default: {
      document: { run: { font: FONT, size: 21, color: INK } },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 30, bold: true, color: BLUE }, paragraph: { spacing: { before: 320, after: 140 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 24, bold: true, color: BLUE2 }, paragraph: { spacing: { before: 220, after: 90 } } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 21, bold: true, color: BLUE2 }, paragraph: { spacing: { before: 160, after: 70 } } },
    ],
  },
  numbering: { config: numberingConfig },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
          border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 4, space: 3 } },
          children: [
            new TextRun({ text: "Pronect • SAM OnSite", font: FONT, size: 16, color: GREY }),
            new TextRun({ children: [new Tab()], font: FONT, size: 16 }),
            new TextRun({ text: "User Story 1117 — Guard Onboarding MVP Specification", font: FONT, size: 16, color: GREY }),
          ],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          tabStops: [{ type: TabStopType.CENTER, position: CONTENT_W / 2 }, { type: TabStopType.RIGHT, position: CONTENT_W }],
          border: { top: { color: RULE, style: BorderStyle.SINGLE, size: 4, space: 3 } },
          children: [
            new TextRun({ text: "Confidential — Pronect internal", font: FONT, size: 16, color: GREY }),
            new TextRun({ children: [new Tab(), "Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GREY }),
            new TextRun({ children: [new Tab()], font: FONT, size: 16 }),
            new TextRun({ text: "US 1117", font: FONT, size: 16, color: GREY }),
          ],
        })],
      }),
    },
    children,
  }],
});

const OUT = path.join(__dirname, "US1117_SAM_OnSite_Guard_Onboarding_MVP_Spec.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("Wrote", OUT, "(" + buf.length + " bytes)");
});
