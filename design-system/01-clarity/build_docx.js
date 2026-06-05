// =============================================================================
// Mivada — Concept 01 · CLARITY  →  clarity.docx
// Native Word capability statement (docx-js). A4, Inter (fallback Arial),
// coral eyebrows/headings, ink body, a coral-header KPI table, paragraph
// bottom-border rules (never tables as dividers), header/footer with contact
// and page number.
//
// Run:  NODE_PATH=/opt/node22/lib/node_modules node build_docx.js
// =============================================================================
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, TabStopType, TabStopPosition,
} = require("docx");

// ---- Brand tokens (exact) ---------------------------------------------------
const CORAL = "EA493F";
const INK = "111111";
const GRAPHITE = "323232";
const MIDGREY = "AEAEAE";
const HAIRLINE = "E4E4E0";
const WHITE = "FFFFFF";
const FONT = "Inter"; // Word substitutes Arial if Inter isn't installed

// A4 content width with ~1" margins: 11906 - 2*1418 ≈ 9070 DXA
const CONTENT_W = 9070;
const MARGIN = 1418;

// ---- helpers ---------------------------------------------------------------
const bottomRule = (color = HAIRLINE, size = 6, space = 6) => ({
  bottom: { style: BorderStyle.SINGLE, size, color, space },
});

function eyebrow(text, { color = CORAL, before = 240 } = {}) {
  return new Paragraph({
    spacing: { before, after: 80 },
    children: [
      new TextRun({
        text: text.toUpperCase(),
        font: FONT,
        bold: true,
        size: 16, // 8pt
        color,
        characterSpacing: 28, // tracked (twentieths of a point)
      }),
    ],
  });
}

// a hairline rule as a paragraph bottom-border (never a table)
function rule(color = HAIRLINE, size = 6, after = 160, before = 0) {
  return new Paragraph({
    border: bottomRule(color, size, 1),
    spacing: { before, after, line: 1 },
    children: [new TextRun({ text: "", size: 2 })],
  });
}

function run(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: opts.size ?? 21, // 10.5pt body
    color: opts.color ?? GRAPHITE,
    bold: opts.bold ?? false,
    italics: opts.italics ?? false,
    characterSpacing: opts.spacing,
  });
}

// pillar / step item: bold ink title + grey description on next line
function itemBlock(title, body, { num } = {}) {
  const titleChildren = [];
  if (num) {
    titleChildren.push(new TextRun({ text: num + "  ", font: FONT, bold: true, size: 21, color: CORAL }));
  }
  titleChildren.push(new TextRun({ text: title, font: FONT, bold: true, size: 21, color: INK, characterSpacing: -2 }));
  return [
    new Paragraph({
      spacing: { before: 140, after: 20 },
      border: bottomRule(HAIRLINE, 4, 0),
      children: titleChildren,
    }),
    new Paragraph({
      spacing: { before: 30, after: 60 },
      children: [run(body, { size: 20 })],
    }),
  ];
}

// KPI table cell
function kpiCell(eyebrowText, big, label, width, isHeaderRow) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: { top: 120, bottom: 140, left: 160, right: 160 },
    shading: { fill: WHITE, type: ShadingType.CLEAR, color: "auto" },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
      left: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
      right: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
    },
    children: [
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: eyebrowText.toUpperCase(), font: FONT, bold: true, size: 13, color: CORAL, characterSpacing: 26 })],
      }),
      new Paragraph({
        spacing: { after: 30 },
        children: [new TextRun({ text: big, font: FONT, bold: true, size: 40, color: INK, characterSpacing: -4 })],
      }),
      new Paragraph({
        children: [new TextRun({ text: label, font: FONT, size: 17, color: GRAPHITE })],
      }),
    ],
    verticalAlign: VerticalAlign.TOP,
  });
}

function kpiRow(cells) {
  const w = Math.floor(CONTENT_W / 3);
  return new TableRow({
    children: cells.map((c) => kpiCell(c[0], c[1], c[2], w, false)),
  });
}

// =============================================================================
// DOCUMENT
// =============================================================================
const doc = new Document({
  creator: "Mivada",
  title: "Mivada — Capability statement 2026",
  description: "Capability statement (Clarity concept)",
  styles: {
    default: {
      document: { run: { font: FONT, size: 21, color: INK } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 33, bold: true, color: INK, characterSpacing: -6 },
        paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: INK, characterSpacing: -4 },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "caps",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "–",
          alignment: AlignmentType.LEFT,
          style: { run: { color: CORAL, bold: true }, paragraph: { indent: { left: 360, hanging: 240 } } },
        }],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4 DXA
          margin: { top: MARGIN, right: MARGIN, bottom: 1300, left: MARGIN },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
              spacing: { after: 60 },
              children: [
                new TextRun({ text: "Mivada", font: FONT, bold: true, size: 20, color: INK, characterSpacing: -2 }),
                new TextRun({ text: "\tCapability statement · 2026", font: FONT, bold: true, size: 14, color: GRAPHITE, characterSpacing: 24 }),
              ],
            }),
            rule(CORAL, 8, 0, 0),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            rule(INK, 8, 80, 40),
            new Paragraph({
              tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
              children: [
                new TextRun({ text: "mivada.com · hello@mivada.com · Sydney & Melbourne", font: FONT, size: 16, color: GRAPHITE }),
                new TextRun({ text: "\tPage ", font: FONT, size: 16, color: MIDGREY }),
                new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: MIDGREY }),
              ],
            }),
            new Paragraph({
              spacing: { before: 20 },
              children: [new TextRun({ text: "Figures illustrative.", font: FONT, size: 14, italics: true, color: MIDGREY })],
            }),
          ],
        }),
      },
      children: [
        // ---- Masthead eyebrow + Title -------------------------------------
        eyebrow("Capability overview · 2026", { before: 40 }),
        new Paragraph({
          spacing: { before: 40, after: 160 },
          children: [
            new TextRun({ text: "Technology, ", font: FONT, bold: true, size: 38, color: INK, characterSpacing: -6 }),
            new TextRun({ text: "human", font: FONT, bold: true, size: 38, color: CORAL, characterSpacing: -6 }),
            new TextRun({ text: " first.", font: FONT, bold: true, size: 38, color: INK, characterSpacing: -6 }),
          ],
        }),

        // ---- Intro --------------------------------------------------------
        new Paragraph({
          spacing: { after: 120, line: 300 },
          children: [
            run("Mivada is an ", { size: 22 }),
            run("Australian technology consultancy", { size: 22, bold: true, color: INK }),
            run(" that turns enterprise platforms — Workday, payroll, data and AI — into measurable human outcomes. We pair senior onshore leadership with cost-effective offshore delivery, and we stay on after go-live to measure outcomes, not tickets closed. The name comes from ", { size: 22 }),
            run("my village", { size: 22, bold: true, color: INK }),
            run(" (community) and ", { size: 22 }),
            run("wisdom", { size: 22, bold: true, color: INK }),
            run(" (knowledge); our core value is that we care.", { size: 22 }),
          ],
        }),

        // ---- What we do ---------------------------------------------------
        eyebrow("What we do"),
        rule(HAIRLINE, 6, 120),
        ...itemBlock("Workday & ERP", "Implementation, optimisation and managed support for Workday HCM, Financials and adjacent ERP."),
        ...itemBlock("Payroll consulting", "Compliant, accurate payroll across complex awards and multi-entity structures."),
        ...itemBlock("Data & AI", "Lakehouse architecture, governed analytics and applied AI on top of your people and finance data."),
        ...itemBlock("Intelligent automation", "RPA, ML and process design combined to remove low-value work."),

        // ---- How we work --------------------------------------------------
        eyebrow("How we work — think human first"),
        rule(HAIRLINE, 6, 120),
        ...itemBlock("Listen", "Understand the people and the work before the technology.", { num: "1" }),
        ...itemBlock("Design", "Shape the platform around how teams actually operate.", { num: "2" }),
        ...itemBlock("Deliver", "Implement in weeks-long increments, not multi-year programs.", { num: "3" }),
        ...itemBlock("Care", "Stay on after go-live; measure outcomes, not tickets closed.", { num: "4" }),

        // ---- Selected capabilities — Data & AI ----------------------------
        eyebrow("Selected capabilities — Data & AI"),
        rule(HAIRLINE, 6, 140),
        new Paragraph({
          numbering: { reference: "caps", level: 0 },
          spacing: { after: 70, line: 280 },
          children: [
            run("Certified consultants design "), run("Lakehouse architectures", { bold: true, color: INK }),
            run(" built for governance and scale."),
          ],
        }),
        new Paragraph({
          numbering: { reference: "caps", level: 0 },
          spacing: { after: 70, line: 280 },
          children: [
            run("We build "), run("Delta Lake pipelines", { bold: true, color: INK }),
            run(" — batch and streaming — that are reliable and observable."),
          ],
        }),
        new Paragraph({
          numbering: { reference: "caps", level: 0 },
          spacing: { after: 70, line: 280 },
          children: [
            run("We integrate "), run("Databricks", { bold: true, color: INK }),
            run(" with Workday, payroll and finance to deliver a "),
            run("governed KPI store", { bold: true, color: INK }), run("."),
          ],
        }),
        new Paragraph({
          numbering: { reference: "caps", level: 0 },
          spacing: { after: 70, line: 280 },
          children: [
            run("Leaders get "), run("real-time insight", { bold: true, color: INK }),
            run(" and decide on current data, not last quarter’s."),
          ],
        }),

        // ---- By the numbers (KPI table, coral header) ---------------------
        eyebrow("By the numbers"),
        new Paragraph({ spacing: { after: 60 }, children: [run("Illustrative figures for these design examples.", { size: 18, color: MIDGREY })] }),
        new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          columnWidths: [Math.floor(CONTENT_W / 3), Math.floor(CONTENT_W / 3), CONTENT_W - 2 * Math.floor(CONTENT_W / 3)],
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
            left: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
            right: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
            insideVertical: { style: BorderStyle.SINGLE, size: 4, color: HAIRLINE },
          },
          rows: [
            // coral header row spanning the KPI block
            new TableRow({
              tableHeader: true,
              children: [
                new TableCell({
                  columnSpan: 3,
                  width: { size: CONTENT_W, type: WidthType.DXA },
                  shading: { fill: CORAL, type: ShadingType.CLEAR, color: "auto" },
                  margins: { top: 80, bottom: 80, left: 160, right: 160 },
                  children: [new Paragraph({ children: [new TextRun({ text: "PROOF POINTS", font: FONT, bold: true, size: 15, color: WHITE, characterSpacing: 30 })] })],
                }),
              ],
            }),
            kpiRow([
              ["Established", "2014", "Founded as LJM Infotech."],
              ["Specialists", "120+", "Certified, not generalist."],
              ["Workday", "40+", "Deployments delivered."],
            ]),
            kpiRow([
              ["Retention", "98%", "Client retention rate."],
              ["Pillars", "4", "Workday, payroll, data & AI, automation."],
              ["Delivery", "AU + IN", "Onshore Australia, offshore India."],
            ]),
          ],
        }),

        // ---- Closing line -------------------------------------------------
        new Paragraph({
          spacing: { before: 260, after: 40 },
          children: [
            new TextRun({ text: "Let’s build smarter systems, with people at the ", font: FONT, bold: true, size: 22, color: INK, characterSpacing: -2 }),
            new TextRun({ text: "centre.", font: FONT, bold: true, size: 22, color: CORAL, characterSpacing: -2 }),
          ],
        }),
        new Paragraph({
          spacing: { after: 40 },
          children: [run("Experience the Mivada difference.", { size: 19, color: GRAPHITE })],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("clarity.docx", buffer);
  console.log("Wrote clarity.docx");
});
