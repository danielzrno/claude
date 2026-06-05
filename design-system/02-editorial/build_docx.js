// ============================================================================
// Mivada — Concept 02 · EDITORIAL  ·  native Word (.docx) via docx-js
//
// A4 capability statement, Inter (fallback Arial). Large bold headings, coral
// accents, a black-shaded callout cell, a coral-left-border pull-quote, a KPI
// table. Same copy as document.html, styled on-brand within Word's limits.
//
//   NODE_PATH=/opt/node22/lib/node_modules node build_docx.js   ->  editorial.docx
// ============================================================================

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, TabStopType, TabStopPosition,
} = require("docx");

// ---- Brand tokens (exact) --------------------------------------------------
const CORAL = "EA493F";
const CORAL_DEEP = "C9362B";
const BLACK = "000000";
const INK = "111111";
const GRAPHITE = "323232";
const MID = "AEAEAE";
const HAIR = "E4E4E0";
const WHITE = "FFFFFF";
const OFFWHITE = "FAFAFA";
const FONT = "Inter"; // Office fallback Arial happens automatically if absent

// A4 content width with the section margins below (~11906 - 1300 - 1300).
// Use 9304 so it divides evenly by 2 and 4 for exact table column math.
const CONTENT_W = 9304;

// ---- small helpers ---------------------------------------------------------
const sz = (pt) => pt * 2;            // half-points
const dxa = (inch) => Math.round(inch * 1440);

function run(text, { size = 11, bold = false, color = INK, italic = false, caps = false, spacing } = {}) {
  return new TextRun({
    text, font: FONT, size: sz(size), bold, italics: italic,
    color, allCaps: caps,
    ...(spacing != null ? { characterSpacing: spacing } : {}),
  });
}

// eyebrow: small uppercase tracked coral label
function eyebrow(text, { color = CORAL, before = 0, after = 60 } = {}) {
  return new Paragraph({
    spacing: { before, after },
    children: [run(text, { size: 8.5, bold: true, color, caps: true, spacing: 32 })],
  });
}

// big section headline
function headline(parts, { before = 40, after = 140, size = 18 } = {}) {
  return new Paragraph({
    spacing: { before, after },
    children: parts.map(([t, accent]) =>
      run(t, { size, bold: true, color: accent ? CORAL : INK, spacing: -6 })),
  });
}

// a paragraph that draws a bottom rule (never a table for dividers)
function rule(color = HAIR, size = 6, before = 0, after = 120) {
  return new Paragraph({
    spacing: { before, after },
    border: { bottom: { style: BorderStyle.SINGLE, size, color, space: 1 } },
    children: [run("", { size: 1 })],
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after != null ? opts.after : 80, line: 276 },
    alignment: AlignmentType.LEFT,
    children: [run(text, { size: opts.size || 10.5, color: opts.color || GRAPHITE })],
  });
}

// ============================================================================
// Build
// ============================================================================
function build() {

  // ---- Masthead (wordmark left via coral "chip" run + meta right) ----------
  const masthead = new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    spacing: { after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: INK, space: 6 } },
    children: [
      // M chip rendered as a coral-highlighted bold M
      new TextRun({ text: " M ", font: FONT, size: sz(15), bold: true, color: WHITE,
        shading: { type: ShadingType.CLEAR, fill: CORAL, color: "auto" } }),
      run("  Mivada", { size: 15, bold: true, color: INK, spacing: -4 }),
      new TextRun({ text: "\tCAPABILITY STATEMENT · ", font: FONT, size: sz(8.5),
        bold: true, color: GRAPHITE, allCaps: true, characterSpacing: 28 }),
      new TextRun({ text: "2026", font: FONT, size: sz(8.5), bold: true,
        color: CORAL, allCaps: true, characterSpacing: 28 }),
    ],
  });

  // ---- Oversized lede ------------------------------------------------------
  const lede = [
    eyebrow("Technology, human first", { before: 150, after: 80 }),
    new Paragraph({
      spacing: { after: 110 },
      children: [
        run("We turn enterprise platforms ", { size: 26, bold: true, color: INK, spacing: -8 }),
        run("into ", { size: 26, bold: true, color: INK, spacing: -8 }),
        run("human outcomes.", { size: 26, bold: true, color: CORAL, spacing: -8 }),
      ],
    }),
  ];

  // intro — two-column flow via a borderless table (keeps it on one page band)
  const introCell = (txt) => new TableCell({
    width: { size: CONTENT_W / 2, type: WidthType.DXA },
    margins: { top: 0, bottom: 0, left: 0, right: 180 },
    borders: noBorders(),
    children: [body(txt, { after: 0 })],
  });
  const introTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
    borders: noBorders(),
    rows: [new TableRow({ children: [
      introCell("Mivada is an Australian technology consultancy. We take the platforms that run an enterprise — Workday, payroll, data and AI — and shape them around the people who use them every day."),
      new TableCell({
        width: { size: CONTENT_W / 2, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 180, right: 0 },
        borders: noBorders(),
        children: [body("Founded in 2014 and now 120-plus specialists, we lead onshore in Australia and deliver across India. We measure success in outcomes adopted, not tickets closed.", { after: 0 })],
      }),
    ] })],
  });

  // ---- What we do — numbered editorial items (2 columns) -------------------
  const pillarCell = (idx, title, desc) => new TableCell({
    width: { size: CONTENT_W / 2, type: WidthType.DXA },
    margins: { top: 50, bottom: 80, left: 0, right: 160 },
    borders: { top: { style: BorderStyle.SINGLE, size: 6, color: HAIR } ,
               bottom: noB(), left: noB(), right: noB() },
    children: [
      new Paragraph({ spacing: { after: 30 }, children: [
        run(idx + "  ", { size: 16, bold: true, color: CORAL, spacing: -8 }),
        run(title, { size: 11.5, bold: true, color: INK }),
      ]}),
      new Paragraph({ spacing: { after: 0, line: 264 }, children: [
        run(desc, { size: 9.5, color: GRAPHITE }),
      ]}),
    ],
  });
  const pillarRow = (a, b) => new TableRow({ children: [a, b] });
  const pillars = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
    borders: noBorders(),
    rows: [
      pillarRow(
        pillarCell("01", "Workday & ERP", "Implementation, optimisation and managed support for Workday HCM, Financials and adjacent ERP."),
        pillarCell("02", "Payroll consulting", "Compliant, accurate payroll across complex awards and multi-entity structures."),
      ),
      pillarRow(
        pillarCell("03", "Data & AI", "Lakehouse architecture, governed analytics and applied AI on your people and finance data."),
        pillarCell("04", "Intelligent automation", "RPA, ML and process design combined to remove low-value work."),
      ),
    ],
  });

  // ---- How we work — 4-up bold sequence (table, vertical rules) ------------
  const stepCell = (n, name, desc, first) => new TableCell({
    width: { size: CONTENT_W / 4, type: WidthType.DXA },
    margins: { top: 40, bottom: 40, left: first ? 0 : 150, right: 120 },
    borders: { top: noB(), bottom: noB(), right: noB(),
      left: first ? noB() : { style: BorderStyle.SINGLE, size: 6, color: HAIR } },
    children: [
      new Paragraph({ spacing: { after: 40 }, children: [run(n, { size: 7.5, bold: true, color: CORAL, caps: true, spacing: 30 })] }),
      new Paragraph({ spacing: { after: 40 }, children: [run(name, { size: 15, bold: true, color: INK, spacing: -6 })] }),
      new Paragraph({ spacing: { after: 0, line: 252 }, children: [run(desc, { size: 8.5, color: GRAPHITE })] }),
    ],
  });
  const steps = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W / 4, CONTENT_W / 4, CONTENT_W / 4, CONTENT_W / 4],
    borders: noBorders(),
    rows: [new TableRow({ children: [
      stepCell("Step 01", "Listen", "Understand the people and the work before the technology.", true),
      stepCell("Step 02", "Design", "Shape the platform around how teams actually operate."),
      stepCell("Step 03", "Deliver", "Implement in weeks-long increments, not multi-year programs."),
      stepCell("Step 04", "Care", "Stay on after go-live; measure outcomes, not tickets."),
    ]})],
  });

  // ---- BLACK CALLOUT — Data & AI capabilities (single black-shaded cell) ---
  const capLine = (k, parts) => new Paragraph({
    spacing: { after: 90, line: 260 },
    children: [
      run(k.toUpperCase() + "   ", { size: 8, bold: true, color: CORAL, caps: true, spacing: 26 }),
      ...parts.map(([t, strong]) => run(t, { size: 9.5, bold: strong, color: strong ? WHITE : "E2E2E2" })),
    ],
  });
  const calloutCell = new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: BLACK, color: "auto" },
    margins: { top: 170, bottom: 170, left: 220, right: 220 },
    borders: noBorders(),
    children: [
      eyebrow("Selected capabilities — Data & AI", { after: 70 }),
      new Paragraph({ spacing: { after: 140 }, children: [
        run("The data foundation for current-data decisions.", { size: 15, bold: true, color: WHITE, spacing: -6 }),
      ]}),
      capLine("Architecture", [["Certified consultants design ", false], ["Lakehouse architectures", true], [" on your people and finance data.", false]]),
      capLine("Pipelines", [["Build ", false], ["Delta Lake pipelines", true], [" — batch and streaming — on ", false], ["Databricks", true], [".", false]]),
      capLine("Integration", [["Integrate Databricks with ", false], ["Workday, payroll and finance", true], [".", false]]),
      capLine("Outcome", [["Deliver ", false], ["governed KPI stores", true], [" and ", false], ["real-time insight", true], [".", false]]),
    ],
  });
  const callout = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    borders: noBorders(),
    rows: [new TableRow({ children: [calloutCell] })],
  });

  // ---- Coral pull-quote (coral left border) --------------------------------
  const pullquote = [
    new Paragraph({
      spacing: { before: 130, after: 30, line: 264 },
      indent: { left: 200 },
      border: { left: { style: BorderStyle.SINGLE, size: 30, color: CORAL, space: 14 } },
      children: [run("“They delivered in weeks what we'd scoped for a year — and our team actually uses it.”",
        { size: 15.5, bold: true, color: CORAL, spacing: -6 })],
    }),
    new Paragraph({
      spacing: { after: 40 }, indent: { left: 200 },
      border: { left: { style: BorderStyle.SINGLE, size: 30, color: CORAL, space: 14 } },
      children: [run("People & Payroll Director · national employer (illustrative)", { size: 8.5, bold: true, color: MID })],
    }),
  ];

  // ---- By the numbers — KPI table (coral big numerals) ---------------------
  const kpi = (num, accent, label, first) => new TableCell({
    width: { size: CONTENT_W / 4, type: WidthType.DXA },
    margins: { top: 20, bottom: 20, left: first ? 0 : 150, right: 120 },
    verticalAlign: VerticalAlign.TOP,
    borders: { top: noB(), bottom: noB(), right: noB(),
      left: first ? noB() : { style: BorderStyle.SINGLE, size: 6, color: HAIR } },
    children: [
      new Paragraph({ spacing: { after: 30 }, children: [
        run(num, { size: 23, bold: true, color: INK, spacing: -10 }),
        ...(accent ? [run(accent, { size: 23, bold: true, color: CORAL, spacing: -10 })] : []),
      ]}),
      new Paragraph({ children: [run(label.toUpperCase(), { size: 7.5, bold: true, color: MID, caps: true, spacing: 22 })] }),
    ],
  });
  const numbers = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W / 4, CONTENT_W / 4, CONTENT_W / 4, CONTENT_W / 4],
    borders: noBorders(),
    rows: [new TableRow({ children: [
      kpi("40", "+", "Workday deployments", true),
      kpi("98", "%", "Client retention"),
      kpi("120", "+", "Specialists"),
      kpi("4", "", "Service pillars"),
    ]})],
  });

  // ---- Assemble ------------------------------------------------------------
  const doc = new Document({
    creator: "Mivada",
    title: "Mivada — Capability Statement 2026",
    styles: {
      default: { document: { run: { font: FONT, size: sz(10.5), color: INK } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: sz(18), bold: true, color: INK },
          paragraph: { spacing: { before: 240, after: 140 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: sz(13), bold: true, color: INK },
          paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },        // A4 (DXA)
          margin: { top: 1120, right: 1300, bottom: 1000, left: 1300 },
        },
      },
      headers: {
        default: new Header({ children: [ new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
          spacing: { after: 0 },
          children: [
            run("Mivada", { size: 8.5, bold: true, color: INK, spacing: -2 }),
            new TextRun({ text: "\tTechnology, human first", font: FONT, size: sz(8), color: MID }),
          ],
        }) ] }),
      },
      footers: {
        default: new Footer({ children: [ new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: HAIR, space: 6 } },
          children: [
            run("hello@mivada.com · mivada.com · Sydney & Melbourne", { size: 8, color: GRAPHITE }),
            new TextRun({ text: "\t", font: FONT }),
            new TextRun({ text: "Figures illustrative · ", font: FONT, size: sz(8), color: MID }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: sz(8), color: MID }),
          ],
        }) ] }),
      },
      children: [
        masthead,
        ...lede,
        introTable,
        // What we do
        eyebrow("What we do", { before: 170, after: 26 }),
        headline([["Four practices, one ", false], ["operating belief.", true]], { before: 16, after: 100, size: 15 }),
        pillars,
        // How we work
        eyebrow("How we work — think human first", { before: 150, after: 26 }),
        headline([["Understand the people ", false], ["before the technology.", true]], { before: 16, after: 100, size: 15 }),
        steps,
        // Black callout
        new Paragraph({ spacing: { before: 120, after: 0 }, children: [run("", { size: 2 })] }),
        callout,
        // Pull-quote
        ...pullquote,
        // Numbers
        eyebrow("By the numbers", { before: 110, after: 30 }),
        numbers,
      ],
    }],
  });

  Packer.toBuffer(doc).then((buf) => {
    fs.writeFileSync("editorial.docx", buf);
    console.log("Wrote editorial.docx");
  });
}

// ---- border helpers --------------------------------------------------------
function noB() { return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }; }
function noBorders() {
  return { top: noB(), bottom: noB(), left: noB(), right: noB(),
    insideHorizontal: noB(), insideVertical: noB() };
}

build();
