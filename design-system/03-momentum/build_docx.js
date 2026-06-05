// ============================================================================
// MIVADA · Concept 03 — MOMENTUM · native Word (.docx) builder (docx-js)
//
// Capability Statement, A4, Inter (fallback Arial). Brand-forward within Word's
// constraints: a coral masthead band with a white M-chip, card-like bordered
// sections, a KPI table with a coral header row, pill-style labels, ink body.
//
// Run:  NODE_PATH=/opt/node22/lib/node_modules node build_docx.js
//        -> momentum.docx
// ============================================================================
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, ImageRun, Tab, TabStopType,
  TabStopPosition,
} = require("docx");

const HERE = __dirname;

// ---- Palette (exact) -------------------------------------------------------
const CORAL = "EA493F";
const CORAL_DEEP = "C9362B";
const INK = "111111";
const GRAPHITE = "323232";
const MID = "AEAEAE";
const HAIR = "E4E4E0";
const WARM = "F2F3EE";
const WHITE = "FFFFFF";
const CORAL_SOFT = "FCE7E4";
const ON_DARK_EY = "FFD9D4";
const FONT = "Inter";

// A4 content width (DXA): 11906 - 2*~1080 margins ≈ 9746; we use 15mm margins.
const MARGIN = 850;                 // ~15mm in DXA (1440=1in, 15mm≈850)
const PAGE_W = 11906;
const CONTENT_W = PAGE_W - 2 * MARGIN; // ≈ 10206

const mchip = fs.readFileSync(path.join(HERE, "mchip.png"));
const mchipWhite = fs.readFileSync(path.join(HERE, "mchip_white.png"));

// ---- small builders --------------------------------------------------------
function img(buf, w, h, alt) {
  return new ImageRun({
    type: "png", data: buf,
    transformation: { width: w, height: h },
    altText: { title: alt, description: alt, name: alt },
  });
}

function t(text, opts = {}) {
  return new TextRun({
    text, font: FONT,
    size: opts.size || 21,            // half-points (21 = 10.5pt)
    bold: opts.bold || false,
    color: opts.color || INK,
    allCaps: opts.caps || false,
    characterSpacing: opts.spacing || undefined, // in twentieths of a point
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after === undefined ? 120 : opts.after,
               line: opts.line || 264, lineRule: "auto" },
    indent: opts.indent,
    border: opts.border,
    shading: opts.shading,
    keepNext: opts.keepNext,
  });
}

function eyebrow(text, color = CORAL) {
  return para([t(text, { size: 16, bold: true, color, caps: true, spacing: 28 })],
              { after: 60 });
}

// a "pill" rendered as a single-cell shaded rounded-ish table cell row
// (Word has no true pills; we approximate with a tightly-shaded inline cell run)
function pillCell(label, fill, fg, opts = {}) {
  return new TableCell({
    width: { size: opts.w || 1300, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR, color: "auto" },
    margins: { top: 30, bottom: 30, left: 130, right: 130 },
    verticalAlign: VerticalAlign.CENTER,
    borders: noBorders(fill),
    children: [para([t(label, { size: 17, bold: true, color: fg })],
                    { after: 0, align: AlignmentType.CENTER, line: 240 })],
  });
}

function noBorders(c) {
  const b = { style: BorderStyle.SINGLE, size: 2, color: c };
  return { top: b, bottom: b, left: b, right: b };
}
function cellBorders(c, sz = 4) {
  const b = { style: BorderStyle.SINGLE, size: sz, color: c };
  return { top: b, bottom: b, left: b, right: b };
}

// a row of pills laid out as a borderless table
function pillRow(pills, gap = 120) {
  const cells = [];
  const widths = [];
  pills.forEach((p, i) => {
    cells.push(pillCell(p.label, p.fill, p.fg, { w: p.w || 1500 }));
    widths.push(p.w || 1500);
    if (i < pills.length - 1) {
      cells.push(spacerCell(gap));
      widths.push(gap);
    }
  });
  // trailing spacer to fill row
  const used = widths.reduce((a, b) => a + b, 0);
  if (used < CONTENT_W) { cells.push(spacerCell(CONTENT_W - used)); widths.push(CONTENT_W - used); }
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    borders: noBordersTable(),
    rows: [new TableRow({ children: cells })],
  });
}

function spacerCell(w, fill) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: noBordersTable(),
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children: [para([t("", { size: 2 })], { after: 0, line: 120 })],
  });
}

function noBordersTable() {
  const n = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: n, bottom: n, left: n, right: n, insideHorizontal: n, insideVertical: n };
}

// ============================================================================
// Build
// ============================================================================
function build() {
  const children = [];

  // ---- MASTHEAD (coral band, full content width) ----------------------------
  // M-chip (white) + wordmark + meta, on coral, then headline + service pills.
  const masthead = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    borders: noBordersTable(),
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { fill: CORAL, type: ShadingType.CLEAR, color: "auto" },
            margins: { top: 230, bottom: 250, left: 260, right: 260 },
            borders: noBorders(CORAL),
            children: [
              // wordmark row: chip + name  (… meta on the right via tab)
              new Paragraph({
                spacing: { after: 60, line: 240 },
                tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W - 540 }],
                children: [
                  img(mchipWhite, 30, 30, "Mivada M chip"),
                  new TextRun({ text: "  Mivada", font: FONT, size: 30, bold: true, color: WHITE }),
                  new TextRun({ text: "\t", font: FONT }),
                  new TextRun({ text: "Capability Statement · 2026", font: FONT, size: 17, bold: true, color: ON_DARK_EY, allCaps: true, characterSpacing: 24 }),
                ],
              }),
              para([t("Technology, human first", { size: 16, bold: true, color: ON_DARK_EY, caps: true, spacing: 28 })],
                   { after: 70, before: 60 }),
              para([t("Enterprise platforms, turned into outcomes people feel.",
                      { size: 30, bold: true, color: WHITE })],
                   { after: 0, line: 268 }),
            ],
          }),
        ],
      }),
    ],
  });
  children.push(masthead);
  children.push(spacerPara(70));

  // service pills under the masthead
  children.push(pillRow([
    { label: "Workday & ERP", fill: CORAL_SOFT, fg: CORAL_DEEP, w: 1850 },
    { label: "Payroll", fill: CORAL_SOFT, fg: CORAL_DEEP, w: 1250 },
    { label: "Data & AI", fill: CORAL_SOFT, fg: CORAL_DEEP, w: 1450 },
    { label: "Automation", fill: CORAL_SOFT, fg: CORAL_DEEP, w: 1650 },
  ]));
  children.push(spacerPara(120));

  // ---- INTRO -----------------------------------------------------------------
  children.push(para([
    t("Mivada is an Australian technology consultancy", { bold: true, size: 21 }),
    t(" that turns enterprise platforms — Workday, payroll, data and AI — into measurable human outcomes. Founded in 2014, we pair senior onshore leadership with cost-effective delivery from India, and we stay on after go-live to measure outcomes, not tickets closed. We think human first: people and the work come before the technology.", { color: GRAPHITE, size: 21 }),
  ], { line: 288, after: 160 }));

  children.push(ruleParagraph());

  // ---- WHAT WE DO (4 pillar cards as a 2x2 bordered table) -------------------
  children.push(eyebrow("What we do"));
  children.push(para([t("Four service pillars", { size: 24, bold: true })], { after: 140, keepNext: true }));
  children.push(pillarGrid([
    ["01", "Workday & ERP", "Implementation, optimisation and managed support for Workday HCM, Financials and adjacent ERP."],
    ["02", "Payroll consulting", "Compliant, accurate payroll across complex awards and multi-entity structures."],
    ["03", "Data & AI", "Lakehouse architecture, governed analytics and applied AI on your people and finance data."],
    ["04", "Intelligent automation", "RPA, ML and process design combined to remove low-value work for good."],
  ]));
  children.push(spacerPara(150));

  // ---- HOW WE WORK -----------------------------------------------------------
  children.push(eyebrow("How we work — think human first"));
  children.push(para([t("Listen · Design · Deliver · Care", { size: 24, bold: true })], { after: 120, keepNext: true }));
  children.push(stepGrid([
    ["Listen", "Understand the people and the work before the technology.", false],
    ["Design", "Shape the platform around how teams actually operate.", false],
    ["Deliver", "Implement in weeks-long increments, not multi-year programs.", false],
    ["Care", "Stay on after go-live; measure outcomes, not tickets closed.", true],
  ]));
  children.push(spacerPara(150));

  children.push(ruleParagraph());

  // ---- SELECTED CAPABILITIES — DATA & AI -------------------------------------
  children.push(eyebrow("Selected capabilities — Data & AI"));
  children.push(para([t("Leaders decide on current data", { size: 24, bold: true })], { after: 120, keepNext: true }));
  const caps = [
    ["Lakehouse architecture", "Designed by certified consultants around your platforms."],
    ["Delta Lake pipelines", "Batch and streaming, built to govern and to scale."],
    ["Databricks integration", "Connected to Workday, payroll and finance."],
    ["Governed KPI store", "A single, trusted source for the numbers that matter."],
    ["Real-time insight", "So leaders decide on what is true right now."],
  ];
  caps.forEach(([title, desc], i) => {
    children.push(new Paragraph({
      spacing: { after: 70, line: 252 },
      indent: { left: 360, hanging: 360 },
      children: [
        t(String(i + 1).padStart(2, "0") + "   ", { bold: true, color: CORAL, size: 19 }),
        t(title + " — ", { bold: true, size: 19 }),
        t(desc, { color: GRAPHITE, size: 19 }),
      ],
    }));
  });
  children.push(spacerPara(150));

  // ---- BY THE NUMBERS (KPI table, coral header row) --------------------------
  children.push(eyebrow("By the numbers"));
  children.push(kpiTable([
    ["120+", "Specialists"],
    ["40+", "Workday deployments"],
    ["98%", "Client retention"],
    ["2014", "Founded"],
  ]));

  // ---- DOCUMENT --------------------------------------------------------------
  const doc = new Document({
    creator: "Mivada",
    title: "Mivada — Capability Statement 2026",
    description: "Momentum concept capability statement",
    styles: {
      default: { document: { run: { font: FONT, size: 21, color: INK } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 30, bold: true, font: FONT, color: INK },
          paragraph: { spacing: { before: 200, after: 140 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: FONT, color: INK },
          paragraph: { spacing: { before: 160, after: 120 }, outlineLevel: 1 } },
      ],
    },
    numbering: {
      config: [
        { reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { run: { color: CORAL }, paragraph: { indent: { left: 360, hanging: 240 } } } }] },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: 16838 },
          margin: { top: MARGIN, right: MARGIN, bottom: 980, left: MARGIN },
        },
      },
      headers: {
        default: new Header({ children: [] }),
      },
      footers: {
        default: new Footer({
          children: [
            // top hairline rule
            new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 12, color: INK, space: 6 } },
              spacing: { before: 0, after: 0, line: 240 },
              tabStops: [
                { type: TabStopType.CENTER, position: Math.floor(CONTENT_W / 2) },
                { type: TabStopType.RIGHT, position: CONTENT_W },
              ],
              children: [
                new TextRun({ text: "mivada.com  ·  hello@mivada.com", font: FONT, size: 17, color: GRAPHITE, bold: true }),
                new TextRun({ text: "\tSydney & Melbourne · Onshore AU & India\t", font: FONT, size: 17, color: GRAPHITE }),
                new TextRun({ text: "Figures illustrative.", font: FONT, size: 17, color: MID }),
              ],
            }),
          ],
        }),
      },
      children,
    }],
  });

  return doc;
}

// ---- helpers that build tables --------------------------------------------
function spacerPara(after) {
  return new Paragraph({ children: [t("", { size: 2 })], spacing: { after, before: 0, line: 40 } });
}

function ruleParagraph() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: HAIR, space: 1 } },
    spacing: { before: 40, after: 200, line: 40 },
    children: [t("", { size: 2 })],
  });
}

// 2x2 grid of pillar "cards" with an M-chip and an index
function pillarGrid(items) {
  const colW = Math.floor((CONTENT_W - 200) / 2);
  const rows = [];
  for (let r = 0; r < 2; r++) {
    const cells = [];
    for (let c = 0; c < 2; c++) {
      const idx = r * 2 + c;
      const [num, title, body] = items[idx];
      cells.push(new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: cellBorders(HAIR, 4),
        shading: { fill: WHITE, type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 150, bottom: 160, left: 180, right: 180 },
        verticalAlign: VerticalAlign.TOP,
        children: [
          new Paragraph({
            spacing: { after: 60, line: 240 },
            children: [
              img(mchip, 22, 22, "M"),
              new TextRun({ text: "   " + num + " · " + title.toUpperCase(), font: FONT, size: 15, bold: true, color: CORAL, characterSpacing: 18 }),
            ],
          }),
          para([t(title, { size: 21, bold: true })], { after: 40, line: 240 }),
          para([t(body, { size: 18, color: GRAPHITE })], { after: 0, line: 248 }),
        ],
      }));
      if (c === 0) cells.push(spacerCell(200));
    }
    rows.push(new TableRow({ children: cells }));
    if (r === 0) rows.push(spacerRowGrid([colW, 200, colW]));
  }
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [colW, 200, colW],
    borders: noBordersTable(),
    rows,
  });
}

function spacerRowGrid(widths) {
  return new TableRow({
    children: widths.map(w => new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: noBordersTable(),
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [para([t("", { size: 2 })], { after: 0, line: 120 })],
    })),
  });
}

// 4 step cells in a row; last is coral
function stepGrid(items) {
  const gap = 140;
  const colW = Math.floor((CONTENT_W - gap * 3) / 4);
  const cells = [];
  const widths = [];
  items.forEach(([label, body, coral], i) => {
    const fill = coral ? CORAL : WHITE;
    const labelFill = coral ? WHITE : INK;
    const labelFg = coral ? CORAL_DEEP : WHITE;
    const bodyCol = coral ? WHITE : GRAPHITE;
    cells.push(new TableCell({
      width: { size: colW, type: WidthType.DXA },
      borders: coral ? noBorders(CORAL) : cellBorders(HAIR, 4),
      shading: { fill, type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 150, bottom: 160, left: 160, right: 160 },
      verticalAlign: VerticalAlign.TOP,
      children: [
        // pill label as a mini nested table so it reads as a chip
        miniPill(label, labelFill, labelFg),
        para([t("Step " + String(i + 1).padStart(2, "0"), { size: 15, bold: true, color: coral ? ON_DARK_EY : CORAL, spacing: 16 })],
             { after: 60, before: 60, line: 220 }),
        para([t(body, { size: 17, color: bodyCol })], { after: 0, line: 248 }),
      ],
    }));
    widths.push(colW);
    if (i < items.length - 1) { cells.push(spacerCell(gap)); widths.push(gap); }
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    borders: noBordersTable(),
    rows: [new TableRow({ children: cells })],
  });
}

function miniPill(label, fill, fg) {
  // a one-cell table that visually reads as a pill/tag
  return new Table({
    width: { size: 1100, type: WidthType.DXA },
    columnWidths: [1100],
    borders: noBordersTable(),
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 1100, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR, color: "auto" },
      borders: noBorders(fill),
      margins: { top: 20, bottom: 24, left: 110, right: 110 },
      verticalAlign: VerticalAlign.CENTER,
      children: [para([t(label, { size: 16, bold: true, color: fg })], { after: 0, align: AlignmentType.CENTER, line: 220 })],
    })] })],
  });
}

// KPI table: coral header row "By the numbers", then big coral numbers + labels
function kpiTable(items) {
  const n = items.length;
  const colW = Math.floor(CONTENT_W / n);
  const widths = items.map((_, i) => (i === n - 1 ? CONTENT_W - colW * (n - 1) : colW));

  const headerRow = new TableRow({
    tableHeader: true,
    children: items.map((_, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: CORAL, type: ShadingType.CLEAR, color: "auto" },
      borders: { top: cb(CORAL), bottom: cb(WHITE, 8), left: cb(i === 0 ? CORAL : WHITE, 8), right: cb(CORAL) },
      margins: { top: 60, bottom: 60, left: 160, right: 160 },
      verticalAlign: VerticalAlign.CENTER,
      children: [para([t(i === 0 ? "Mivada by the numbers" : "", { size: 16, bold: true, color: WHITE, caps: true, spacing: 24 })],
                      { after: 0, line: 220 })],
    })),
  });

  const numRow = new TableRow({
    children: items.map(([num, _], i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: WARM, type: ShadingType.CLEAR, color: "auto" },
      borders: { top: cb(WHITE, 8), bottom: cb(WHITE, 4), left: cb(i === 0 ? WARM : WHITE, 8), right: cb(WARM) },
      margins: { top: 140, bottom: 40, left: 160, right: 160 },
      verticalAlign: VerticalAlign.BOTTOM,
      children: [para([t(num, { size: 44, bold: true, color: CORAL })], { after: 0, line: 240 })],
    })),
  });

  const labelRow = new TableRow({
    children: items.map(([_, label], i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: WARM, type: ShadingType.CLEAR, color: "auto" },
      borders: { top: cb(WHITE, 4), bottom: cb(WARM), left: cb(i === 0 ? WARM : WHITE, 8), right: cb(WARM) },
      margins: { top: 20, bottom: 150, left: 160, right: 160 },
      verticalAlign: VerticalAlign.TOP,
      children: [para([t(label, { size: 18, color: GRAPHITE, bold: true })], { after: 0, line: 224 })],
    })),
  });

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, numRow, labelRow],
  });
}

function cb(color, size = 4) {
  return { style: BorderStyle.SINGLE, size, color };
}

// ---- write -----------------------------------------------------------------
const doc = build();
Packer.toBuffer(doc).then((buf) => {
  const out = path.join(HERE, "momentum.docx");
  fs.writeFileSync(out, buf);
  console.log("wrote", out);
});
