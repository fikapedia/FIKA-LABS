/**
 * Builds: docs/ai-engine-selection-climate-risk.docx
 * Usage:  npm install docx && node scripts/build_engine_brief.js
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat,
} = require("docx");

const CONTENT_W = 9360;
const INK = "1A1A1A";
const ACCENT = "1F5C4D";
const RULE = "C9D6D1";
const HEADFILL = "EAF1EE";
const ZEBRA = "F7FAF9";

const P = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 276 },
  children: [new TextRun({ text, bold: o.bold, italics: o.italics, color: o.color ?? INK, size: o.size ?? 22 })],
});

const BULLET = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 70, line: 276 },
  children: [new TextRun({ text, size: 22, color: INK })],
});

const STEP = (text) => new Paragraph({
  numbering: { reference: "steps", level: 0 },
  spacing: { after: 90, line: 276 },
  children: [new TextRun({ text, size: 22, color: INK })],
});

const LEADB = (lead, rest) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 80, line: 276 },
  children: [
    new TextRun({ text: lead, bold: true, size: 22, color: INK }),
    new TextRun({ text: rest, size: 22, color: INK }),
  ],
});

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

const CALLOUT = (text) => new Paragraph({
  spacing: { before: 140, after: 140, line: 276 },
  indent: { left: 200, right: 200 },
  shading: { type: ShadingType.CLEAR, fill: "EAF1EE" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 10 } },
  children: [new TextRun({ text, size: 21, color: "174036" })],
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
    top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    left: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  },
  rows: [
    new TableRow({ tableHeader: true,
      children: header.map((h, i) => cell(h, widths[i], { bold: true, fill: HEADFILL, color: "144035" })) }),
    ...rows.map((r, ri) => new TableRow({
      children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? ZEBRA : "FFFFFF" })),
    })),
  ],
});

const doc = new Document({
  creator: "FIKA Labs",
  title: "Choosing an AI Engine for Sustainability and Climate Risk",
  description: "Engine, harness and data-layer selection for climate risk assessment",
  styles: { default: { document: { run: { font: "Calibri", size: 22, color: INK } } } },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 360, hanging: 220 } } } }] },
    { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 380, hanging: 380 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    children: [
      new Paragraph({ spacing: { after: 40 },
        children: [new TextRun({ text: "SELECTION BRIEF", bold: true, size: 20, color: ACCENT, characterSpacing: 40 })] }),
      new Paragraph({ spacing: { after: 80 },
        children: [new TextRun({ text: "Choosing an AI engine for sustainability and climate risk", bold: true, size: 36, color: INK })] }),
      new Paragraph({ spacing: { after: 300 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
        children: [new TextRun({ text: "Prepared 18 September 2026  ·  Prices and vendor status verified on that date", size: 19, color: "5A6B66" })] }),

      /* 1 */
      H1("1.  Recommendation at a glance"),
      P("There is no single engine that does climate risk assessment. The decision splits into three layers, and most buyers get the first one right and the third one wrong."),
      table([2050, 2750, 4560],
        ["Layer", "Recommendation", "Why"],
        [
          ["Reasoning engine", "Claude Opus 5", "1M-token context, page-anchored citations, strong tool use. List price $5 / $25 per million tokens."],
          ["Bulk extraction", "Claude Haiku 4.5", "$1 / $5 per million tokens. Use it for parsing invoices, meter reads and supplier files at volume."],
          ["Harness, bespoke work", "Claude Code", "Client-by-client consulting deliverables where a human drives each step."],
          ["Harness, repeatable work", "Managed Agents", "Hosted agent loop, versioned configs and scheduled runs for recurring assessments."],
          ["Physical hazard data", "Two vendors, never one", "Vendor outputs diverge sharply for the same asset. See Section 4."],
          ["Carbon accounting", "Watershed, Persefoni or Sweep", "Pick on organisational shape, not feature lists. See Section 5."],
        ]),

      CALLOUT("Short version: spend your evaluation effort on the data layer. The engine is the easy half of this decision, and the expensive mistakes live downstream of it."),

      /* 2 */
      H1("2.  Why the engine is the easy half"),
      P("A language model reasons. It does not measure. No engine knows the flood depth at a given address, the grid factor for a given subregion, or what your Tier 2 supplier burned last year. It reads what you give it and reasons over that."),
      P("So the engine question reduces to a short list of properties that actually matter for this work:"),
      LEADB("Context window.  ", "Climate work is document-heavy. A CSRD filing, a year of utility data and a hazard report should fit in one pass rather than being chunked and stitched."),
      LEADB("Citation anchoring.  ", "The engine must point at the page it drew a claim from, not paraphrase from memory."),
      LEADB("Tool use.  ", "It has to call your scripts and datasets reliably, because that is where the real numbers live."),
      LEADB("Cost per document.  ", "Scope 3 supplier screening means thousands of documents, so the per-token rate decides whether the workflow is viable."),

      /* 3 */
      H1("3.  The reasoning layer"),
      P("Claude Opus 5 is the sensible default for this domain. The deciding feature is citations: with citations enabled, the response carries the quoted source text and the page number it came from. That turns every assertion into something a reviewer can check against the original PDF, which is exactly what assurance asks for."),
      P("List prices, Anthropic first-party API, mid-2026:"),
      table([2000, 1400, 1800, 4160],
        ["Model", "Context", "Input / Output", "Use it for"],
        [
          ["Claude Opus 5", "1M", "$5 / $25", "Default. Materiality judgement, risk narratives, disclosure drafting, reconciliation."],
          ["Claude Fable 5.1", "1M", "$10 / $50", "Most capable widely released model. Reserve for genuinely hard reasoning; not the everyday choice."],
          ["Claude Sonnet 5", "1M", "$2 / $10", "Mid-tier throughput where judgement is lighter."],
          ["Claude Haiku 4.5", "200K", "$1 / $5", "High-volume extraction: invoices, meter data, supplier questionnaires."],
        ]),
      new Paragraph({ spacing: { after: 120 } }),
      P("Split the work across two models rather than paying top rate for everything. Haiku handles extraction at volume; Opus 5 handles anything that requires a judgement you would defend to a client.", { italics: true }),

      H2("Harness: which one depends on repeatability"),
      LEADB("Claude Code.  ", "Right for consulting deliverables built once per client, where an analyst steers each pass and reviews the output. This is the pattern behind the workshop-style climate builds."),
      LEADB("Managed Agents.  ", "Right when the same assessment runs on a schedule across a portfolio. Anthropic hosts the loop and the workspace, agent configs are versioned, and deployments fire on a cron cadence."),

      /* 4 */
      H1("4.  Where climate risk assessments actually fail"),
      P("The data layer, not the model layer. This is settled by evidence rather than opinion."),
      P("The Global Association of Risk Professionals benchmarked 13 physical risk vendors on the same assets. The spread was wide across both hazard estimates and the damages attached to them. A single property could be rated highly exposed to flooding by one vendor and not exposed at all by another. GARP attributes it to a complexity cascade: modelling choices compound across weather, hydrology, hydraulics, damage functions and scenarios."),
      P("CarbonPlan found the same pattern independently. Comparing a commercial flood model against an open academic model on Los Angeles County properties, the two agreed on roughly one in five."),
      CALLOUT("The practical consequence: no vendor is the accurate one. Treat a single-vendor number as an opinion, not a measurement."),
      P("What to do instead:"),
      STEP("Licence two providers and run both across the portfolio."),
      STEP("Report the spread, not a single figure. Where they agree, confidence is genuine; where they diverge, that is the finding."),
      STEP("Disclose the model, the scenario and the vintage behind every number you publish."),
      STEP("Keep an open-source model as a sanity check on the commercial ones."),

      /* 5 */
      H1("5.  Choosing at the data layer"),
      H2("Physical risk"),
      table([2100, 4000, 3260],
        ["Vendor", "Strength", "Commercial shape"],
        [
          ["Jupiter Intelligence", "22.3bn locations, three scenarios in five-year steps to 2100, adaptation ROI modelling", "Enterprise, six figures annually"],
          ["XDI", "Nine hazards, four scenarios to 2100, resolution to five metres, damage in monetary terms", "Enterprise"],
          ["Moody's (Four Twenty Seven)", "Flood, heat, hurricane, sea level, water stress, wildfire, tied to credit assessment", "Enterprise"],
          ["S&P Climanomics, MSCI", "Portfolio-level financial impact for listed exposure", "Enterprise, six figures"],
          ["Climate X, Repath", "Mid-market coverage with lighter commitments", "Roughly €20k–€60k for 30–100 assets"],
        ]),
      new Paragraph({ spacing: { after: 140 } }),
      H2("Carbon accounting"),
      P("The AI feature gap between these has largely closed. Every serious vendor now does AI-assisted ingestion and Scope 3 estimation, so choose on organisational shape."),
      table([2100, 4000, 3260],
        ["Platform", "Best fit", "Commercial shape"],
        [
          ["Watershed", "Large enterprises with complex supply chains and heavy disclosure load; aligns to 10+ frameworks including CSRD, ISSB, CDP and GRI", "Roughly $55k–$250k a year"],
          ["Persefoni", "Financial institutions and investor-grade reporting", "Roughly $55k–$250k a year"],
          ["Sweep", "Decentralised organisations needing cross-functional accountability", "Enterprise quote"],
          ["IBM Envizi, Microsoft, Salesforce", "Estates already standardised on that vendor's cloud or ERP", "Enterprise quote"],
        ]),

      /* 6 */
      H1("6.  How the layers bind"),
      P("The architecture that holds up in practice is thin at the engine layer and disciplined at the data layer:"),
      STEP("Licensed hazard and factor data lands in files with vintages in the filenames."),
      STEP("Haiku 4.5 extracts structured records from client documents at volume."),
      STEP("Scripts do the arithmetic, so it can be rerun and diffed."),
      STEP("Opus 5 reasons over the reconciled output and drafts the narrative, citing pages."),
      STEP("A human checks the citations and the divergence between vendors before anything ships."),
      P("The engine never supplies a number. It reads, reasons and cites. Every figure originates in a file you licensed or measured.", { italics: true }),

      /* 7 */
      H1("7.  A note on this recommendation"),
      P("I am not a neutral party on the reasoning layer, so weigh Section 3 accordingly. The criteria in Section 2 are the objective part: context window, citation anchoring, tool-use reliability and cost per document. Test any candidate engine against those four on your own documents before committing. Sections 4 and 5 are where the independent evidence sits, and that is the part of the decision I would not shortcut."),

      /* sources */
      H1("Sources checked"),
      ...[
        "GARP Risk Institute, Comparing Climate Risk Vendors: A User's Guide to Physical Risk Assessments, and The Complexity Cascade — garp.org",
        "CarbonPlan, Climate risk companies don't always agree — carbonplan.org",
        "Jupiter Intelligence, ClimateScore Global product documentation — jupiterintel.com",
        "Verdantix, Smart Innovators: Physical Climate Risk Solutions — verdantix.com",
        "Persefoni, The 10 Best Carbon Accounting Software in 2026 — persefoni.com",
        "Sweep, Persefoni alternatives for carbon and ESG reporting — sweep.net",
        "Anthropic model pricing and capabilities, first-party API rates, mid-2026 — anthropic.com",
      ].map((s) => new Paragraph({
        spacing: { after: 60, line: 252 },
        indent: { left: 240, hanging: 240 },
        children: [new TextRun({ text: s, size: 18, color: "44544F" })],
      })),
    ],
  }],
});

const out = path.join(__dirname, "..", "docs", "ai-engine-selection-climate-risk.docx");
fs.mkdirSync(path.dirname(out), { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("wrote", out, "(" + buf.length + " bytes)");
});
