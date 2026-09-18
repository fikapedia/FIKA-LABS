/**
 * Builds: docs/ai-search-tools-climate-risk.docx
 * Usage:  npm install docx && node scripts/build_tools_scan.js
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat,
} = require("docx");

const CONTENT_W = 9360;
const INK = "1A1A1A", ACCENT = "1F5C4D", RULE = "C9D6D1", HEADFILL = "EAF1EE", ZEBRA = "F7FAF9";

const P = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 120, line: 276 },
  children: [new TextRun({ text, bold: o.bold, italics: o.italics, color: o.color ?? INK, size: o.size ?? 22 })],
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
      children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? ZEBRA : "FFFFFF" })) })),
  ],
});

const doc = new Document({
  creator: "FIKA Labs",
  title: "AI Search Tools for Climate Risk",
  description: "Scan of AI-powered search and research tools for climate risk management and strategy",
  styles: { default: { document: { run: { font: "Calibri", size: 22, color: INK } } } },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 360, hanging: 220 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    children: [
      new Paragraph({ spacing: { after: 40 },
        children: [new TextRun({ text: "LANDSCAPE SCAN", bold: true, size: 20, color: ACCENT, characterSpacing: 40 })] }),
      new Paragraph({ spacing: { after: 80 },
        children: [new TextRun({ text: "AI search tools for climate risk and strategy", bold: true, size: 36, color: INK })] }),
      new Paragraph({ spacing: { after: 300 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
        children: [new TextRun({ text: "What exists now, and which ones cite their sources  ·  18 September 2026", size: 19, color: "5A6B66" })] }),

      H1("1.  There is no single engine for this"),
      P("Nothing owns the category the way a general search engine owns the open web. What exists instead splits into three groups that answer different questions, and mixing them up is the common mistake."),
      LEADB("Grounded search over science and policy.  ", "Open tools that retrieve from a fixed, citable corpus. Best for regulatory and scientific questions."),
      LEADB("Commercial disclosure assistants.  ", "Closed tools that read your own reports and pre-fill frameworks. Best for reporting workload."),
      LEADB("Measurement engines.  ", "Queryable datasets rather than search. Best for baselines."),
      CALLOUT("One test separates the useful from the decorative: does it hand back a source you can open? Anything that answers without a retrievable citation is a liability in a climate deliverable, however fluent it sounds."),

      H1("2.  Grounded search over science and policy"),
      table([2300, 3500, 3560],
        ["Tool", "What it searches", "Access and grounding"],
        [
          ["Climate Policy Radar",
           "30,000+ climate laws, policies, litigation records, finance documents and UN submissions from every country",
           "Free and open. UK not-for-profit, around 500,000 annual users. NLP over full document text. The strongest fit for regulatory and strategy questions."],
          ["ClimateQ&A (Ekimetrics)",
           "Roughly 14,000 pages of IPCC and IPBES reports",
           "Open source, published on Hugging Face, multilingual, returns the passages it used. Retrieval is climate-specific; answer writing routes through a general model."],
          ["ChatClimate",
           "IPCC Sixth Assessment Report",
           "Peer reviewed in Nature Communications Earth & Environment, with answers evaluated by IPCC authors. The hybrid mode scored most accurate."],
          ["ClimateGPT (Erasmus.AI, AppTek, EQTY Lab)",
           "Climate corpus of roughly 300 billion climate-specific tokens behind a 7-billion parameter model",
           "Open source on Hugging Face, 20+ languages. Announced January 2024 and no later status found, so confirm it is still maintained before depending on it."],
        ]),

      H1("3.  Commercial disclosure and regulatory assistants"),
      P("These read your documents rather than the public record. They shorten reporting work; they do not answer research questions."),
      table([2100, 3900, 3360],
        ["Tool", "Built for", "Worth knowing"],
        [
          ["Briink",
           "Extracting answers from your existing reports and policies, then pre-filling CDP, CSRD and ESRS, supplier and investor questionnaires",
           "Grounds each answer in a traceable source snippet. Partnered with CDP on AI-assisted disclosure."],
          ["Manifest Climate",
           "Gap analysis and benchmarking against CSRD, IFRS S1 and S2, and TCFD",
           "Positioned as an AI analyst over disclosure documents rather than a search tool."],
          ["Clarity AI",
           "Financial institutions: emissions, net-zero alignment, scenario analysis and regulatory compliance",
           "Risk analytics rather than open search. Their 2026 survey found under 30% of organisations fully ready for this year's disclosure requirements."],
        ]),

      H1("4.  Measurement, not search"),
      P("Climate TRACE is the one worth knowing. It uses satellites, remote sensing and machine learning to publish asset-level emissions across more than 352 million assets, free and open, from a coalition of over 100 universities and research groups. Releases moved to monthly in 2025; the August 2026 release carried data through June 2026, with July due on 24 September 2026."),
      P("Treat it as a cross-check rather than a source of record. An independent study has publicly disputed part of the dataset, so verify any figure you intend to publish against the asset itself.", { italics: true }),

      H1("5.  How to choose"),
      LEADB("Policy, regulation and litigation.  ", "Climate Policy Radar. Open, broad and citable."),
      LEADB("Climate science claims.  ", "ChatClimate or ClimateQ&A, because both return the passage behind the answer."),
      LEADB("Your own reporting burden.  ", "Briink or Manifest Climate, chosen on whether you need questionnaire pre-fill or framework gap analysis."),
      LEADB("Emissions baselines.  ", "Climate TRACE, verified before publication."),
      P("None of these replaces licensed hazard data. Search tools tell you what has been written and what has been measured. They do not tell you the flood depth at an asset, and vendors of that data disagree with each other sharply.", { bold: true }),

      H1("Sources checked"),
      ...[
        "Climate Policy Radar, open climate law and policy database — climatepolicyradar.org",
        "Ekimetrics, ClimateQ&A, and its Hugging Face space — ekimetrics.com; huggingface.co/spaces/Ekimetrics",
        "Vaghefi et al., ChatClimate: Grounding conversational AI in climate science, Nature Communications Earth & Environment — nature.com",
        "AppTek and Erasmus.AI, ClimateGPT announcement, January 2024 — apptek.ai; erasmus.ai",
        "Briink, AI for ESG data, and the CDP partnership — briink.com",
        "Manifest Climate, Manifest Research — manifestclimate.com",
        "Clarity AI, climate risk products and 2026 investor survey — clarity.ai",
        "Climate TRACE, 2026 data releases and open emissions database — climatetrace.org",
      ].map((s) => new Paragraph({
        spacing: { after: 60, line: 252 },
        indent: { left: 240, hanging: 240 },
        children: [new TextRun({ text: s, size: 18, color: "44544F" })],
      })),
    ],
  }],
});

const out = path.join(__dirname, "..", "docs", "ai-search-tools-climate-risk.docx");
fs.mkdirSync(path.dirname(out), { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("wrote", out, "(" + buf.length + " bytes)");
});
