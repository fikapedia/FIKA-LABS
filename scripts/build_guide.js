/**
 * Builds: docs/claude-code-climate-work-operating-guide.docx
 *
 * Usage:  npm install docx && node scripts/build_guide.js
 * Output is deterministic; re-run after editing content below.
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageOrientation, convertInchesToTwip,
} = require("docx");

const CONTENT_W = 9360;            // 12240 page - 2 x 1440 margins
const INK = "1A1A1A";
const ACCENT = "1F5C4D";           // deep green
const RULE = "C9D6D1";
const HEADFILL = "EAF1EE";
const ZEBRA = "F7FAF9";

/* ---------- helpers ---------- */
const P = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 276 },
  alignment: o.align,
  children: [new TextRun({ text, bold: o.bold, italics: o.italics, color: o.color ?? INK, size: o.size ?? 22 })],
});

const BULLET = (text, o = {}) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 70, line: 276 },
  children: [new TextRun({ text, size: 22, color: INK, bold: o.bold })],
});

const STEP = (text) => new Paragraph({
  numbering: { reference: "steps", level: 0 },
  spacing: { after: 90, line: 276 },
  children: [new TextRun({ text, size: 22, color: INK })],
});

// Paragraph with a bold lead-in, then normal text
const LEADP = (lead, rest, numbered) => new Paragraph({
  numbering: numbered ? { reference: "steps", level: 0 } : undefined,
  spacing: { after: 90, line: 276 },
  children: [
    new TextRun({ text: lead, bold: true, size: 22, color: INK }),
    new TextRun({ text: rest, size: 22, color: INK }),
  ],
});

const MONO = (lines) => lines.map((t, i) => new Paragraph({
  spacing: { after: i === lines.length - 1 ? 60 : 20, line: 240 },
  indent: { left: 240 },
  shading: { type: ShadingType.CLEAR, fill: "F2F4F3" },
  children: [new TextRun({ text: t, font: "Consolas", size: 18, color: "22332E" })],
}));

const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 320, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } },
  children: [new TextRun({ text, bold: true, size: 28, color: ACCENT })],
});

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, bold: true, size: 23, color: INK })],
});

const cell = (text, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  shading: { type: ShadingType.CLEAR, fill: o.fill ?? "FFFFFF" },
  margins: { top: 90, bottom: 90, left: 120, right: 120 },
  children: (Array.isArray(text) ? text : [text]).map((t) => new Paragraph({
    spacing: { after: 0, line: 252 },
    children: [new TextRun({ text: t, bold: o.bold, size: o.size ?? 18, color: o.color ?? INK })],
  })),
});

const table = (widths, header, rows) => new Table({
  columnWidths: widths,
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: {
    top:    { style: BorderStyle.SINGLE, size: 4, color: RULE },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    left:   { style: BorderStyle.SINGLE, size: 4, color: RULE },
    right:  { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideVertical:   { style: BorderStyle.SINGLE, size: 4, color: RULE },
  },
  rows: [
    new TableRow({
      tableHeader: true,
      children: header.map((h, i) => cell(h, widths[i], { bold: true, fill: HEADFILL, color: "144035" })),
    }),
    ...rows.map((r, ri) => new TableRow({
      children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? ZEBRA : "FFFFFF" })),
    })),
  ],
});

/* ---------- document ---------- */
const doc = new Document({
  creator: "FIKA Labs",
  title: "Claude Code for Climate Work - Operating Guide",
  description: "How to produce current, accurate climate deliverables with Claude Code",
  styles: { default: { document: { run: { font: "Calibri", size: 22, color: INK } } } },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 220 } } } }] },
      { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 380, hanging: 380 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      },
    },
    children: [
      /* ---- title block ---- */
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: "CLAUDE CODE FOR CLIMATE WORK", bold: true, size: 20, color: ACCENT, characterSpacing: 40 })],
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: "How to get current, accurate results", bold: true, size: 40, color: INK })],
      }),
      new Paragraph({
        spacing: { after: 300 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
        children: [new TextRun({ text: "Operating guide  ·  Prepared 18 September 2026  ·  Source status verified on that date", size: 19, color: "5A6B66" })],
      }),

      /* ---- 1 ---- */
      H1("1.  Where accuracy actually breaks"),
      P("Ask a model for an emission factor and you will get one. It arrives with the right units, a sensible magnitude and a confident tone. That is exactly the problem. The number came from training data with a cutoff date, it carries no citation, and it will not survive the first reviewer who asks where it came from."),
      P("A wrong number that looks right costs more than a blank cell. It clears internal review, reaches the client, and fails at assurance."),
      P("Better prompting does not fix this. What fixes it is deciding, before you start, what the model is allowed to treat as a fact."),

      H2("Three rules that do most of the work"),
      STEP("The model never types a factor from memory. It reads factors from files you downloaded and checked."),
      STEP("Every number in a deliverable traces to a file, a table and a row."),
      STEP("Anything that cannot be traced is flagged in the draft, not smoothed over."),
      P("The rest of this guide is the mechanics of holding those three rules."),

      /* ---- 2 ---- */
      H1("2.  Set the workspace up first"),
      P("Claude Code works best when the folder structure carries the discipline. Build this before the first prompt."),
      table([2000, 3760, 3600],
        ["Folder", "What goes in it", "Rule"],
        [
          ["data/raw/", "Client files exactly as received", "Read-only. Never edit or overwrite."],
          ["data/factors/", "Factor tables you downloaded", "Vintage in every filename, e.g. egrid2023rev2_subregion.csv"],
          ["sources/", "A PDF or screenshot of each source page, dated", "One file per factor set. This is your audit trail."],
          ["analysis/", "Scripts that do the arithmetic", "Maths lives in code, not in prose."],
          ["deliverables/", "Drafts and final output", "Nothing enters without passing Section 5."],
        ]),
      new Paragraph({ spacing: { after: 120 } }),
      P("Then write the project CLAUDE.md. This version holds up in practice:"),
      ...MONO([
        "Never state an emission factor, GWP or regulatory deadline from memory.",
        "Read it from data/factors/ or sources/ and cite the filename and row.",
        "",
        "If a needed factor is not in data/factors/, stop and say so. Do not estimate.",
        "",
        "All arithmetic goes in a script under analysis/. No mental maths in prose.",
        "",
        "Report every figure with its unit, its factor vintage and its source file.",
        "",
        "Mark each assumption inline as [ASSUMPTION: ...] so it survives to review.",
      ]),

      /* ---- 3 ---- */
      H1("3.  Source register, status at 18 September 2026"),
      P("Check this at the start of every engagement. Two of the US federal sets did not update on schedule this year, so “current” is no longer the same as “latest expected”."),
      table([1750, 1850, 2560, 3200],
        ["Source", "Use it for", "Current version", "What to watch"],
        [
          ["EPA GHG Emission Factors Hub", "US Scope 1 and Scope 3 default factors",
           "January 2025 edition, the last official EPA release",
           "No confirmed official 2026 edition. Community rebuilds circulate; label them as such."],
          ["EPA eGRID", "US Scope 2 location-based grid factors",
           "eGRID2023 rev2, released 12 June 2025",
           "eGRID2024 was expected January 2026 and has not appeared. Production is reported paused."],
          ["IPCC AR6 GWP-100", "Converting CH₄ and N₂O to CO₂e",
           "Fossil CH₄ 29.8 · Non-fossil CH₄ 27.0 · N₂O 273",
           "Confirm which set the prior year used. Mixing AR5 and AR6 breaks the trend line."],
          ["GHG Protocol Scope 2 Guidance", "Market-based and location-based method",
           "2015 Guidance, still in force",
           "Revision in consultation. Hourly matching, deliverability and a marginal-impact metric proposed. Final expected around end-2027."],
          ["SBTi Corporate Net-Zero Standard", "Target setting and validation",
           "V2.0 published 11 June 2026",
           "V1.3.1 governs validation through 2026. V2.0 optional from 1 Feb 2027, required for new submissions from 1 Feb 2028."],
          ["California SB 253", "US Scope 1 and 2 reporting",
           "First reports due 10 November 2026",
           "Moved back from 10 August 2026. Scope 3 and assurance follow in 2027."],
          ["California SB 261", "Climate-related financial risk reporting",
           "1 January 2026 deadline not being enforced",
           "CARB will set an alternate date once the Ninth Circuit appeal resolves."],
          ["CSRD, as amended by Omnibus I", "EU sustainability reporting",
           "Directive (EU) 2026/470, in force 18 March 2026",
           "Scope is now more than 1,000 employees AND turnover above €450m, both required. Revised ESRS still in draft."],
          ["SEC climate disclosure rules", "US securities disclosure",
           "Rescission proposed 29 May 2026",
           "Treat as not in force. Do not build a deliverable around it."],
        ]),

      new Paragraph({ children: [new TextRun({ text: "" })], spacing: { after: 100 } }),
      H2("The US data gap, and how to handle it"),
      P("This is the live issue for any US inventory right now. Both EPA workhorse datasets have stalled: the last official eGRID covers data year 2023, and the last official Emission Factors Hub edition is January 2025. Independent rebuilds exist. One group ran EPA's own code against 2024 inputs and found subregion factors down roughly 3 percent against 2023."),
      P("Handle it in this order:"),
      STEP("Use the last official EPA version and state the vintage plainly in the deliverable."),
      STEP("If the client needs a 2024 view, run the rebuild as a clearly labelled sensitivity case beside the official figure."),
      STEP("Never swap a rebuild in silently. If it moves the total, the client has to know why."),
      P("One line in the methodology note covers you:"),
      ...MONO([
        "Grid factors are eGRID2023 rev2 (EPA, June 2025), the most recent official",
        "release. EPA has not published a 2024 data year.",
      ]),

      /* ---- 4 ---- */
      H1("4.  The run loop"),
      P("Work every deliverable in five passes. Do not jump to the narrative."),
      LEADP("Data audit.  ", "Point Claude Code at data/raw/ and ask what is actually there: row counts, date ranges, units, blanks, duplicates, outliers. Fix the data before any factor touches it. Most bad inventories go wrong here, not in the maths.", true),
      LEADP("Factor binding.  ", "Match each activity line to a factor and record the match in a table: activity, factor value, unit, source file, row. Anything unmatched stays visible rather than defaulting.", true),
      LEADP("Calculation.  ", "Arithmetic goes in a script under analysis/, never in chat. A script can be rerun, diffed and checked by someone else. A chat answer cannot.", true),
      LEADP("Reconciliation.  ", "Tie totals back to something the client already holds: a utility bill, last year's inventory, a spend ledger. Explain every variance above your materiality threshold.", true),
      LEADP("Narrative.  ", "Only now write the memo or disclosure text, and write it off the reconciled numbers.", true),

      /* ---- 5 ---- */
      H1("5.  Verification gates"),
      P("Run these before anything leaves your desk. They take about ten minutes and catch most of what assurance would catch later at a much higher cost."),
      table([1700, 4000, 3660],
        ["Gate", "What you check", "Fail signal"],
        [
          ["Vintage", "Every factor file carries a version and date, and the deliverable states them", "A factor with no vintage anywhere in the pack"],
          ["Hand-check", "Recompute one line per factor set on a calculator", "Your number and the script disagree"],
          ["Units", "kWh against MWh, kg against tonnes, gallons against litres, short against metric tons", "A total off by a factor of 1,000 or 1.102"],
          ["Traceability", "Pick three figures at random and trace each to a file and a row", "You cannot find the row"],
          ["Tie-out", "Category totals sum to the reported total", "Rounding hiding a dropped category"],
          ["Assumptions", "Every [ASSUMPTION] tag is resolved or disclosed", "A tag that quietly vanished between drafts"],
        ]),

      /* ---- 6 ---- */
      H1("6.  Final check before issue"),
      BULLET("Factor vintages stated in the methodology note"),
      BULLET("One line per factor set recomputed by hand and matched"),
      BULLET("Units checked end to end"),
      BULLET("Three figures traced at random back to source rows"),
      BULLET("Variances above threshold explained"),
      BULLET("Rebuilt or non-official data clearly labelled as such"),
      BULLET("Regulatory status re-checked against Section 3 on the day of issue"),
      new Paragraph({ spacing: { before: 140, after: 120 } , children: [new TextRun({
        text: "The status items in Section 3 move. Re-check them at the start of each engagement rather than trusting this page.",
        italics: true, size: 21, color: "5A6B66" })] }),

      /* ---- sources ---- */
      H1("Sources checked"),
      ...[
        "GHG Protocol, Scope 2 Standard advances: ISB approves consultation on market- and location-based revisions — ghgprotocol.org",
        "EPA, Emissions & Generation Resource Integrated Database (eGRID) and eGRID2023 technical guide — epa.gov/egrid",
        "EPA, Emission Factors for Greenhouse Gas Inventories, January 2025 — epa.gov/climateleadership",
        "SBTi, Introducing the Corporate Net-Zero Standard Version 2.0, published 11 June 2026 — sciencebasedtargets.org",
        "Sidley, SB 253 update: CARB delays reporting deadline to November 2026, 30 June 2026 — sidley.com",
        "Greenberg Traurig, CARB adopts initial climate disclosure reporting regulations, March 2026 — gtlaw.com",
        "PwC Viewpoint, Omnibus directive finalised; Directive (EU) 2026/470 — viewpoint.pwc.com",
        "SEC, Proposed rescission of climate-related disclosure rules, 29 May 2026 — sec.gov",
        "Cornerstone Data, Making eGRID 2024 data easily accessible — cornerstonedata.org",
      ].map((s) => new Paragraph({
        spacing: { after: 60, line: 252 },
        indent: { left: 240, hanging: 240 },
        children: [new TextRun({ text: s, size: 18, color: "44544F" })],
      })),
    ],
  }],
});

const out = path.join(__dirname, "..", "docs", "claude-code-climate-work-operating-guide.docx");
fs.mkdirSync(path.dirname(out), { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("wrote", out, "(" + buf.length + " bytes)");
});
