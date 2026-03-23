import * as fs from "fs";
import * as path from "path";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  LevelFormat, TabStopType, TabStopPosition, TableOfContents,
} from "docx";

import {
  AnnualReport, ReportMeta,
  TextSection, KpiSection, TableSection, ChartSection,
  ImageSection, BulletSection, TwoColumnSection, DividerSection,
  QuoteSection, TimelineSection, TeamSection, ComparisonSection,
  AwardsSection, FinancialsSection,
  TextTwoColumnSection,
  SubSection, 
} from "./types";
import { renderChartToBuffer } from "./chartRenderer";

// ── Page geometry (A4) ────────────────────────────────────────
const PAGE_W   = 11906;
const PAGE_H   = 16838;
const MARGIN   = 1134;   // ~2 cm
const CW       = PAGE_W - MARGIN * 2;   // content width = 9638 DXA

// ── Half-point font sizes ─────────────────────────────────────
const PT = (n: number) => n * 2;

// ── Border helpers ────────────────────────────────────────────
const nb  = () => ({ style: BorderStyle.NIL,    size: 0,  color: "FFFFFF" });
const sb  = (c: string, sz = 6) => ({ style: BorderStyle.SINGLE, size: sz, color: c });
const borders = (c: string) => ({ top: sb(c), bottom: sb(c), left: sb(c), right: sb(c) });
const noBorders = () => ({ top: nb(), bottom: nb(), left: nb(), right: nb() });

function accentLine(c: string): object {
  return { bottom: sb(c, 16), top: nb(), left: nb(), right: nb() };
}

// ── Spacer ────────────────────────────────────────────────────
const spacer = (before = 0, after = 160) =>
  new Paragraph({ spacing: { before, after }, children: [] });

// ── Page break ───────────────────────────────────────────────
const pgBreak = () => new Paragraph({ children: [new PageBreak()] });

// ── Tint: lighten a hex colour ────────────────────────────────
function tint(hex: string, pct = 0.85): string {
  const h = hex.replace("#", "");
  const t = (c: string) => Math.round(parseInt(c, 16) + (255 - parseInt(c, 16)) * pct)
    .toString(16).padStart(2, "0");
  return `${t(h.slice(0,2))}${t(h.slice(2,4))}${t(h.slice(4,6))}`;
}

// ─────────────────────────────────────────────────────────────
//  COVER PAGE
// ─────────────────────────────────────────────────────────────
// ── Normalize any image to a raw buffer docx can read ─────────
function readImageBuffer(filePath: string): Buffer {
  const ext = path.extname(filePath).toLowerCase();
  const buf = fs.readFileSync(filePath);
  // docx.js auto-detects PNG/JPG from buffer magic bytes — just return raw
  return buf;
}

async function buildCover(meta: ReportMeta, dir: string): Promise<Paragraph[]> {
  const ac  = meta.accentColor  ?? "1B3A6B";
  const dk  = meta.darkColor    ?? "0D1F3C";
  const out: Paragraph[] = [];

  // Logo
  if (meta.logoPath) {
    const abs = path.resolve(dir, meta.logoPath);
    if (fs.existsSync(abs)) {
      const buf=readImageBuffer(abs);
      console.log(`🖼️  Logo found: ${abs} (${buf.length} bytes)`);
      out.push(new Paragraph({
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: readImageBuffer(abs), transformation: { width: 180, height: 90 } })],
      }));
    }
  } else {
    out.push(spacer(240,240));
  }

  // Decorative accent bar
  out.push(new Paragraph({
    spacing: { before: 0, after: 0 },
    border: { bottom: sb(ac, 32), top: nb(), left: nb(), right: nb() },
    children: [],
  }));

  // Company name
  out.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text: meta.companyName, bold: true, size: PT(36), font: "Arial", color: dk })],
  }));

  // Report title
  out.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: meta.reportTitle, size: PT(20), font: "Arial", color: "555555" })],
  }));

  // Year badge row
  out.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 120 },
    children: [
      new TextRun({ text: meta.reportYear, size: PT(16), font: "Arial", color: ac, bold: true }),
      ...(meta.fiscalYear ? [new TextRun({ text: `   |   ${meta.fiscalYear}`, size: PT(11), font: "Arial", color: "888888" })] : []),
    ],
  }));

  // Tagline
  if (meta.tagline) {
    out.push(spacer(360, 0));
    out.push(new Paragraph({
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: meta.tagline, size: PT(13), italics: true, font: "Georgia", color: "555555" })],
    }));
  }

  // Stock/ticker
  if (meta.tickerSymbol) {
    out.push(spacer(240, 0));
    out.push(new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: "Listed on  ", size: PT(10), font: "Arial", color: "AAAAAA" }),
        new TextRun({ text: `${meta.stockExchange ?? ""}  ${meta.tickerSymbol}`, size: PT(10), font: "Arial", bold: true, color: "555555" }),
      ],
    }));
  }

  // Prepared by
  if (meta.preparedBy) {
    out.push(spacer(560, 0));
    out.push(new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: "Prepared by  ", size: PT(10), font: "Arial", color: "AAAAAA" }),
        new TextRun({ text: meta.preparedBy, size: PT(10), font: "Arial", bold: true, color: "333333" }),
      ],
    }));
  }

  // Company details footer on cover
  if (meta.registeredOffice || meta.cin) {
    out.push(spacer(240, 0));
    if (meta.registeredOffice) {
      out.push(new Paragraph({
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: `Registered Office: ${meta.registeredOffice}`, size: PT(9), font: "Arial", color: "AAAAAA" })],
      }));
    }
    if (meta.cin) {
      out.push(new Paragraph({
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: `CIN: ${meta.cin}`, size: PT(9), font: "Arial", color: "AAAAAA" })],
      }));
    }
  }

  out.push(pgBreak());
  return out;
}

// ─────────────────────────────────────────────────────────────
//  TABLE OF CONTENTS
// ─────────────────────────────────────────────────────────────
function buildTOC(ac: string): (Paragraph | TableOfContents)[] {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 0, after: 240 },
      children: [new TextRun({ text: "Contents", font: "Arial", size: PT(20), bold: true, color: ac })],
    }),
    new TableOfContents("TOC", {
      hyperlink: true,
      headingStyleRange: "1-2",
      stylesWithLevels: [
        { styleName: "Heading 1", level: 1 },
        { styleName: "Heading 2", level: 2 },
      ],
    }),
    pgBreak(),
  ];
}

// ─────────────────────────────────────────────────────────────
//  SECTION BUILDERS
// ─────────────────────────────────────────────────────────────

// ── Text ──────────────────────────────────────────────────────
function renderSubsections(subsections: SubSection[], ac: string): Paragraph[] {
  const out: Paragraph[] = [];
  for (const sub of subsections) {
    if (sub.type === "heading") {
      out.push(new Paragraph({
        spacing: { before: 160, after: 60 },
        children: [new TextRun({
          text: sub.text, font: "Arial",
          size: PT(11), bold: true, color: ac,
        })],
      }));
    } else if (sub.type === "paragraph") {
      for (const para of sub.text.split(/\n\n+/)) {
        if (para.trim()) {
          out.push(new Paragraph({
            spacing: { before: 0, after: 160 },
            children: [new TextRun({
              text: para.trim(), font: "Arial",
              size: PT(9), color: "222222",
            })],
          }));
        }
      }
    } else if (sub.type === "bullets") {
      sub.items.forEach(item => {
        out.push(new Paragraph({
          numbering: sub.ordered
            ? { reference: "numbered", level: 0 }
            : { reference: "bullets",  level: 0 },
          spacing: { before: 40, after: 40 },
          children: [new TextRun({
            text: item, font: "Arial",
            size: PT(7), color: "222222",
          })],
        }));
      });
      out.push(spacer(0, 80));
    }
  }
  return out;
}

function buildText(s: TextSection, ac: string, sec: string): Paragraph[] {
  const out: Paragraph[] = [];

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  if (s.subtitle) {
    out.push(new Paragraph({
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: s.subtitle, font: "Arial", size: PT(12), italics: true, color: "666666" })],
    }));
  }

  if (s.highlight) {
    out.push(new Paragraph({
      spacing: { before: 120, after: 120 },
      shading: { fill: "FFF9C4", type: ShadingType.CLEAR },
      border: { left: sb(ac, 16), top: nb(), bottom: nb(), right: nb() },
      indent: { left: 240, right: 240 },
      children: [new TextRun({ text: s.highlight, font: "Arial", size: PT(11), color: "333333" })],
    }));
  }

  // Plain body text
  if (s.body) {
    for (const para of s.body.split(/\n\n+/)) {
      if (para.trim()) {
        out.push(new Paragraph({
          spacing: { before: 0, after: 160 },
          children: [new TextRun({ text: para.trim(), font: "Arial", size: PT(11), color: "222222" })],
        }));
      }
    }
  }

  // Subsections (paragraphs + bullets + sub-headings mixed)
  if (s.subsections) {
    out.push(...renderSubsections(s.subsections, ac));
  }

  if (s.quote) {
    out.push(new Paragraph({
      spacing: { before: 200, after: 200 },
      indent: { left: 800, right: 800 },
      shading: { fill: tint(ac, 0.92), type: ShadingType.CLEAR },
      border: { left: sb(ac, 24), top: nb(), bottom: nb(), right: nb() },
      children: [new TextRun({
        text: `\u201C${s.quote}\u201D`,
        font: "Georgia", size: PT(13), italics: true, color: "333333",
      })],
    }));
  }

  return out;
}

// ── KPI Metrics ───────────────────────────────────────────────
function buildKpi(s: KpiSection, ac: string, sec: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const cols   = Math.min(s.columns ?? 4, s.items.length, 4);
  const colW   = Math.floor(CW / cols);
  const widths = Array(cols).fill(colW);
  const lightBg = tint(ac, 0.9);

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  // Chunk items into rows of `cols`
  for (let r = 0; r < s.items.length; r += cols) {
    const rowItems = s.items.slice(r, r + cols);
    // Pad if last row is short
    while (rowItems.length < cols) rowItems.push({ label: "", value: "" });

    const labelRow = new TableRow({
      children: rowItems.map(it => new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: noBorders(),
        shading: { fill: ac, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            ...(it.icon ? [new TextRun({ text: `${it.icon}  `, font: "Arial", size: PT(10), color: "FFFFFF" })] : []),
            new TextRun({ text: it.label, font: "Arial", size: PT(9), bold: true, color: "FFFFFF" }),
          ],
        })],
      })),
    });

    const valueRow = new TableRow({
      children: rowItems.map(it => new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: noBorders(),
        shading: { fill: lightBg, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 80, left: 160, right: 160 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: it.value, font: "Arial", size: PT(24), bold: true, color: it.value ? "1A1A1A" : lightBg })],
          }),
          ...(it.subLabel ? [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: it.subLabel, font: "Arial", size: PT(8), color: "888888" })],
          })] : []),
          ...(it.delta ? [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: it.delta,
              font: "Arial", size: PT(9), bold: true,
              color: it.deltaPositive === false ? "C62828" : "2E7D32",
            })],
          })] : []),
        ],
      })),
    });

    out.push(new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: widths,
      rows: [labelRow, valueRow],
    }));
    out.push(spacer(0, 160));
  }

  return out;
}

// ── Table ─────────────────────────────────────────────────────
function buildTable(s: TableSection, ac: string, sec: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const cols  = s.headers.length;
  const colW  = Math.floor(CW / cols);
  const widths = Array(cols).fill(colW);
  const lightBg = tint(ac, 0.9);
  const brd = borders("D0D0D0");

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  const align = (i: number) => {
    const a = s.columnAlignments?.[i] ?? "left";
    return a === "right" ? AlignmentType.RIGHT : a === "center" ? AlignmentType.CENTER : AlignmentType.LEFT;
  };

  const headerRow = new TableRow({
    tableHeader: true,
    children: s.headers.map((h, i) => new TableCell({
      width: { size: colW, type: WidthType.DXA },
      borders: brd,
      shading: { fill: ac, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        alignment: align(i),
        children: [new TextRun({ text: String(h), font: "Arial", size: PT(10), bold: true, color: "FFFFFF" })],
      })],
    })),
  });

  const dataRows = s.rows.map((row, ri) => {
    const isLast   = ri === s.rows.length - 1 && s.highlightLastRow;
    const isFirstC = s.highlightFirstColumn;
    const fill     = isLast ? lightBg : ri % 2 === 0 ? "FFFFFF" : "F7F8FA";

    return new TableRow({
      children: row.map((cell, ci) => new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: brd,
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: align(ci),
          children: [new TextRun({
            text: String(cell),
            font: "Arial",
            size: PT(10),
            bold: isLast || (isFirstC && ci === 0),
            color: (isLast || (isFirstC && ci === 0)) ? ac : "1A1A1A",
          })],
        })],
      })),
    });
  });

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths, rows: [headerRow, ...dataRows] }));

  out.push(s.caption
    ? new Paragraph({ spacing: { before: 80, after: 200 }, children: [new TextRun({ text: s.caption, font: "Arial", size: PT(9), italics: true, color: "777777" })] })
    : spacer(0, 200)
  );

  return out;
}

// ── Chart ─────────────────────────────────────────────────────
async function buildChart(s: ChartSection, ac: string): Promise<(Paragraph | Table)[]> {
  const out: (Paragraph | Table)[] = [];

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  try {
    const buf = await renderChartToBuffer(s, 600, 340);
    out.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [new ImageRun({ data: buf, transformation: { width: 510, height: 289 } })],
    }));
  } catch {
    out.push(new Paragraph({
      children: [new TextRun({ text: `[Chart unavailable: ${s.title}]`, italics: true, color: "999999", font: "Arial", size: PT(10) })],
    }));
  }

  out.push(s.caption
    ? new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 200 }, children: [new TextRun({ text: s.caption, font: "Arial", size: PT(9), italics: true, color: "777777" })] })
    : spacer(0, 200)
  );

  return out;
}

// ── Image ─────────────────────────────────────────────────────
async function buildImage(s: ImageSection, ac: string, dir: string): Promise<Paragraph[]> {
  const out: Paragraph[] = [];
  const alignment = s.align === "left" ? AlignmentType.LEFT : s.align === "right" ? AlignmentType.RIGHT : AlignmentType.CENTER;

  if (s.title) {
    out.push(new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 280, after: 120 },
      children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
    }));
  }

  const abs = path.resolve(dir, s.path);
  if (fs.existsSync(abs)) {
    out.push(new Paragraph({
      alignment,
      children: [new ImageRun({ data: fs.readFileSync(abs), transformation: { width: s.width ?? 500, height: s.height ?? 300 } })],
    }));
  } else {
    out.push(new Paragraph({
      children: [new TextRun({ text: `[Image not found: ${s.path}]`, italics: true, color: "999999", font: "Arial", size: PT(10) })],
    }));
  }

  if (s.caption) {
    out.push(new Paragraph({
      alignment,
      spacing: { before: 80, after: 200 },
      children: [new TextRun({ text: s.caption, font: "Arial", size: PT(9), italics: true, color: "777777" })],
    }));
  }

  return out;
}

// ── Bullets ───────────────────────────────────────────────────
function buildBullets(s: BulletSection, ac: string): Paragraph[] {
  const out: Paragraph[] = [];

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  s.items.forEach(item => {
    out.push(new Paragraph({
      numbering: s.ordered ? { reference: "numbered", level: 0 } : { reference: "bullets", level: 0 },
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text: item, font: "Arial", size: PT(11), color: "222222" })],
    }));
  });

  out.push(spacer(0, 160));
  return out;
}



// ── Two Column ────────────────────────────────────────────────
function buildTwoColumn(s: TwoColumnSection, ac: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const halfW = Math.floor(CW / 2) - 160;
  const lightBg = tint(ac, 0.93);

  if (s.title) {
    out.push(new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 280, after: 160 },
      children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
    }));
  }

  const makeCell = (heading: string, body: string) =>
    new TableCell({
      width: { size: halfW, type: WidthType.DXA },
      borders: noBorders(),
      shading: { fill: lightBg, type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 200, right: 200 },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 80 },
          children: [new TextRun({ text: heading, font: "Arial", size: PT(12), bold: true, color: ac })],
        }),
        ...body.split(/\n\n+/).map(p =>
          new Paragraph({
            spacing: { before: 0, after: 80 },
            children: [new TextRun({ text: p.trim(), font: "Arial", size: PT(10), color: "333333" })],
          })
        ),
      ],
    });

  out.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [halfW, halfW],
    rows: [new TableRow({ children: [makeCell(s.left.heading, s.left.body), makeCell(s.right.heading, s.right.body)] })],
  }));

  out.push(spacer(0, 200));
  return out;
}

function buildTextTwoColumn(s: TextTwoColumnSection, ac: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const colW = Math.floor(CW / 2) - 80;

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  // Helper: convert Paragraphs into TableCell children
  const makeCell = (paragraphs: Paragraph[]) =>
    new TableCell({
      width: { size: colW, type: WidthType.DXA },
      borders: noBorders(),
      margins: { top: 0, bottom: 0, left: 160, right: 160 },
      children: paragraphs,
    });

  let leftParas:  Paragraph[] = [];
  let rightParas: Paragraph[] = [];

  if (s.subsections) {
    // Split subsections array in half
    const mid   = Math.ceil(s.subsections.length / 2);
    leftParas   = renderSubsections(s.subsections.slice(0, mid), ac);
    rightParas  = renderSubsections(s.subsections.slice(mid), ac);
  } else if (s.body) {
    // Split body text paragraphs in half
    const paras = s.body.split(/\n\n+/).map(p => p.trim()).filter(Boolean);
    const mid   = Math.ceil(paras.length / 2);
    leftParas   = paras.slice(0, mid).map(p =>
      new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun({ text: p, font: "Arial", size: PT(11), color: "222222" })] })
    );
    rightParas  = paras.slice(mid).map(p =>
      new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun({ text: p, font: "Arial", size: PT(11), color: "222222" })] })
    );
  }

  // Ensure cells are never empty
  if (leftParas.length  === 0) leftParas  = [new Paragraph({ children: [] })];
  if (rightParas.length === 0) rightParas = [new Paragraph({ children: [] })];

  out.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [colW, colW],
    rows: [new TableRow({ children: [makeCell(leftParas), makeCell(rightParas)] })],
  }));

  if (s.quote) {
    out.push(new Paragraph({
      spacing: { before: 200, after: 200 },
      indent: { left: 800, right: 800 },
      shading: { fill: tint(ac, 0.92), type: ShadingType.CLEAR },
      border: { left: sb(ac, 24), top: nb(), bottom: nb(), right: nb() },
      children: [new TextRun({ text: `\u201C${s.quote}\u201D`, font: "Georgia", size: PT(13), italics: true, color: "333333" })],
    }));
  }

  out.push(spacer(0, 200));
  return out;
}

// ── Divider ───────────────────────────────────────────────────
function buildDivider(s: DividerSection, ac: string): Paragraph[] {
  const out: Paragraph[] = [];
  if (s.label) {
    out.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200 },
      border: { bottom: sb(ac, 4), top: nb(), left: nb(), right: nb() },
      children: [new TextRun({ text: s.label, font: "Arial", size: PT(9), bold: true, color: ac })],
    }));
  } else {
    out.push(new Paragraph({
      spacing: { before: 200, after: 200 },
      border: { bottom: sb("CCCCCC", 4), top: nb(), left: nb(), right: nb() },
      children: [],
    }));
  }
  return out;
}

// ── Quote / Testimonial ───────────────────────────────────────
function buildQuote(s: QuoteSection, ac: string): Paragraph[] {
  const out: Paragraph[] = [];
  const lightBg = tint(ac, 0.9);

  out.push(new Paragraph({
    spacing: { before: 240, after: 80 },
    indent: { left: 720, right: 720 },
    shading: { fill: lightBg, type: ShadingType.CLEAR },
    border: { left: sb(ac, 28), top: nb(), bottom: nb(), right: nb() },
    children: [
      new TextRun({ text: "\u201C", font: "Georgia", size: PT(28), bold: true, color: ac }),
      new TextRun({ text: s.text, font: "Georgia", size: PT(14), italics: true, color: "333333" }),
      new TextRun({ text: "\u201D", font: "Georgia", size: PT(28), bold: true, color: ac }),
    ],
  }));

  if (s.attribution) {
    out.push(new Paragraph({
      spacing: { before: 80, after: 200 },
      indent: { left: 720 },
      children: [
        new TextRun({ text: `${s.attribution}`, font: "Arial", size: PT(10), bold: true, color: "444444" }),
        ...(s.role ? [new TextRun({ text: `  |  ${s.role}`, font: "Arial", size: PT(9), color: "888888" })] : []),
      ],
    }));
  }

  return out;
}

// ── Timeline ──────────────────────────────────────────────────
function buildTimeline(s: TimelineSection, ac: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const yearW = 1200;
  const bodyW = CW - yearW;
  const lightBg = tint(ac, 0.9);

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  const rows = s.items.map((item, i) =>
    new TableRow({
      children: [
        new TableCell({
          width: { size: yearW, type: WidthType.DXA },
          borders: noBorders(),
          shading: { fill: i % 2 === 0 ? ac : tint(ac, 0.3), type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 160, right: 160 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: item.year, font: "Arial", size: PT(13), bold: true, color: "FFFFFF" })],
          })],
        }),
        new TableCell({
          width: { size: bodyW, type: WidthType.DXA },
          borders: noBorders(),
          shading: { fill: i % 2 === 0 ? lightBg : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 200, right: 160 },
          children: [
            new Paragraph({
              spacing: { before: 0, after: 40 },
              children: [new TextRun({ text: item.title, font: "Arial", size: PT(11), bold: true, color: ac })],
            }),
            new Paragraph({
              spacing: { before: 0, after: 0 },
              children: [new TextRun({ text: item.description, font: "Arial", size: PT(10), color: "444444" })],
            }),
          ],
        }),
      ],
    })
  );

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [yearW, bodyW], rows }));
  out.push(spacer(0, 200));
  return out;
}

// ── Team / Board ──────────────────────────────────────────────
function buildTeam(s: TeamSection, ac: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const cols  = 3;
  const colW  = Math.floor(CW / cols);
  const lightBg = tint(ac, 0.9);

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  // Chunk into rows of 3
  for (let i = 0; i < s.members.length; i += cols) {
    const group = s.members.slice(i, i + cols);
    while (group.length < cols) group.push({ name: "", role: "" });

    out.push(new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: Array(cols).fill(colW),
      rows: [new TableRow({
        children: group.map(m => new TableCell({
          width: { size: colW, type: WidthType.DXA },
          borders: noBorders(),
          shading: { fill: m.name ? lightBg : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 160, right: 160 },
          children: m.name ? [
            new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: m.name, font: "Arial", size: PT(11), bold: true, color: ac })] }),
            new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: m.role, font: "Arial", size: PT(10), color: "555555" })] }),
            ...(m.tenure ? [new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: m.tenure, font: "Arial", size: PT(9), italics: true, color: "888888" })] })] : []),
            ...(m.bio ? [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: m.bio, font: "Arial", size: PT(9), color: "444444" })] })] : []),
          ] : [new Paragraph({ children: [] })],
        })),
      })],
    }));
    out.push(spacer(0, 120));
  }

  out.push(spacer(0, 80));
  return out;
}

// ── Comparison ────────────────────────────────────────────────
function buildComparison(s: ComparisonSection, ac: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const cols  = s.columns.length;
  const colW  = Math.floor(CW / cols);
  const lightBg = tint(ac, 0.9);
  const brd = borders("D0D0D0");

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  const headerRow = new TableRow({
    tableHeader: true,
    children: s.columns.map((h, i) => new TableCell({
      width: { size: colW, type: WidthType.DXA },
      borders: brd,
      shading: { fill: i === 0 ? "F0F0F0" : ac, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
        children: [new TextRun({ text: h, font: "Arial", size: PT(10), bold: true, color: i === 0 ? "444444" : "FFFFFF" })],
      })],
    })),
  });

  const dataRows = s.rows.map((row, ri) => new TableRow({
    children: [
      new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: brd,
        shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F7F8FA", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: row.label, font: "Arial", size: PT(10), bold: true, color: "333333" })] })],
      }),
      ...row.values.map(v => new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders: brd,
        shading: { fill: ri % 2 === 0 ? lightBg : "FFFFFF", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(v), font: "Arial", size: PT(10), color: "1A1A1A" })] })],
      })),
    ],
  }));

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: Array(cols).fill(colW), rows: [headerRow, ...dataRows] }));
  if (s.caption) out.push(new Paragraph({ spacing: { before: 80, after: 200 }, children: [new TextRun({ text: s.caption, font: "Arial", size: PT(9), italics: true, color: "777777" })] }));
  else out.push(spacer(0, 200));
  return out;
}

// ── Awards ────────────────────────────────────────────────────
function buildAwards(s: AwardsSection, ac: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const cols  = Math.min(s.items.length, 3);
  const colW  = Math.floor(CW / cols);
  const lightBg = tint(ac, 0.9);

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  for (let i = 0; i < s.items.length; i += cols) {
    const group = s.items.slice(i, i + cols);
    while (group.length < cols) group.push({ title: "", body: "" });

    out.push(new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: Array(cols).fill(colW),
      rows: [new TableRow({
        children: group.map(aw => new TableCell({
          width: { size: colW, type: WidthType.DXA },
          borders: { top: nb(), bottom: nb(), left: nb(), right: nb() },
          shading: { fill: aw.title ? lightBg : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 140, bottom: 140, left: 180, right: 180 },
          children: aw.title ? [
            ...(aw.year ? [new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: aw.year, font: "Arial", size: PT(9), bold: true, color: ac })] })] : []),
            new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "\u2605 " + aw.title, font: "Arial", size: PT(11), bold: true, color: "222222" })] }),
            new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: aw.body, font: "Arial", size: PT(10), color: "555555" })] }),
          ] : [new Paragraph({ children: [] })],
        })),
      })],
    }));
    out.push(spacer(0, 120));
  }

  return out;
}

// ── Financials (Income Statement style) ──────────────────────
function buildFinancials(s: FinancialsSection, ac: string): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const periodCols = s.periods.length;
  const labelW = Math.floor(CW * 0.38);
  const valW   = Math.floor((CW - labelW) / periodCols);
  const widths = [labelW, ...Array(periodCols).fill(valW)];
  const lightBg = tint(ac, 0.9);
  const brd = borders("DDDDDD");

  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 80 },
    children: [new TextRun({ text: s.title, font: "Arial", size: PT(15), bold: true, color: ac })],
  }));

  if (s.unit || s.currency) {
    out.push(new Paragraph({
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: `All figures in ${s.currency ?? ""}${s.unit ? ` ${s.unit}` : ""}`, font: "Arial", size: PT(9), italics: true, color: "777777" })],
    }));
  }

  // Header row (periods)
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      new TableCell({
        width: { size: labelW, type: WidthType.DXA },
        borders: brd,
        shading: { fill: ac, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: "Particulars", font: "Arial", size: PT(10), bold: true, color: "FFFFFF" })] })],
      }),
      ...s.periods.map(p => new TableCell({
        width: { size: valW, type: WidthType.DXA },
        borders: brd,
        shading: { fill: ac, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: p, font: "Arial", size: PT(10), bold: true, color: "FFFFFF" })] })],
      })),
    ],
  });

  const dataRows = s.items.map((item, ri) => {
    const fill = item.bold ? lightBg : ri % 2 === 0 ? "FFFFFF" : "F7F8FA";
    const sep  = item.separator ? { top: sb(ac, 6), bottom: nb(), left: nb(), right: nb() } : noBorders();

    return new TableRow({
      children: [
        new TableCell({
          width: { size: labelW, type: WidthType.DXA },
          borders: brd,
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: item.indent ? 360 : 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: item.label, font: "Arial", size: PT(10), bold: item.bold, color: item.bold ? ac : "333333" })],
          })],
        }),
        ...item.values.map(v => new TableCell({
          width: { size: valW, type: WidthType.DXA },
          borders: brd,
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: String(v), font: "Arial", size: PT(10), bold: item.bold, color: item.bold ? ac : "1A1A1A" })],
          })],
        })),
      ],
    });
  });

  out.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths, rows: [headerRow, ...dataRows] }));
  if (s.caption) out.push(new Paragraph({ spacing: { before: 80, after: 200 }, children: [new TextRun({ text: s.caption, font: "Arial", size: PT(9), italics: true, color: "777777" })] }));
  else out.push(spacer(0, 200));
  return out;
}

// ─────────────────────────────────────────────────────────────
//  HEADER & FOOTER
// ─────────────────────────────────────────────────────────────
function buildHeader(meta: ReportMeta): Header {
  const ac = meta.accentColor ?? "1B3A6B";
  return new Header({
    children: [new Paragraph({
      border: { bottom: sb(ac, 6), top: nb(), left: nb(), right: nb() },
      spacing: { before: 0, after: 80 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: meta.companyName, font: "Arial", size: PT(9), bold: true, color: ac }),
        new TextRun({ text: `\t${meta.reportTitle}  ${meta.reportYear}`, font: "Arial", size: PT(9), color: "999999" }),
      ],
    })],
  });
}

function buildFooter(meta: ReportMeta): Footer {
  return new Footer({
    children: [new Paragraph({
      spacing: { before: 80, after: 0 },
      border: { top: sb("CCCCCC", 4), bottom: nb(), left: nb(), right: nb() },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: meta.website ?? meta.companyName, font: "Arial", size: PT(8), color: "BBBBBB" }),
        new TextRun({ children: ["\tPage ", PageNumber.CURRENT], font: "Arial", size: PT(8), color: "888888" } as any),
      ],
    })],
  });
}

// ─────────────────────────────────────────────────────────────
//  MAIN BUILDER
// ─────────────────────────────────────────────────────────────
export async function buildReport(report: AnnualReport, jsonDir: string): Promise<Buffer> {
  const ac  = report.meta.accentColor  ?? "1B3A6B";
  const sec = report.meta.secondaryColor ?? "E8EDF5";

  const all: (Paragraph | Table | TableOfContents)[] = [];

  // Cover
  all.push(...await buildCover(report.meta, jsonDir));

  // TOC
  all.push(...buildTOC(ac));

  // Chapters
  for (const chapter of report.chapters) {
    all.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: true,
      spacing: { before: 0, after: 200 },
      border: accentLine(ac) as any,
      children: [
        new TextRun({ text: chapter.title, font: "Arial", size: PT(20), bold: true, color: ac }),
      ],
    }));

    if (chapter.subtitle) {
      all.push(new Paragraph({
        spacing: { before: 0, after: 240 },
        children: [new TextRun({ text: chapter.subtitle, font: "Arial", size: PT(12), italics: true, color: "777777" })],
      }));
    }

    for (const section of chapter.sections) {
      switch (section.type) {
        case "text":        all.push(...buildText(section, ac, sec)); break;
        case "kpi":         all.push(...buildKpi(section, ac, sec)); break;
        case "table":       all.push(...buildTable(section, ac, sec)); break;
        case "chart":       all.push(...await buildChart(section, ac)); break;
        case "image":       all.push(...await buildImage(section, ac, jsonDir)); break;
        case "bullets":     all.push(...buildBullets(section, ac)); break;
        case "twoColumn":   all.push(...buildTwoColumn(section, ac)); break;
        case "textTwoColumn": all.push(...buildTextTwoColumn(section, ac));break;
        case "divider":     all.push(...buildDivider(section, ac)); break;
        case "quote":       all.push(...buildQuote(section, ac)); break;
        case "timeline":    all.push(...buildTimeline(section, ac)); break;
        case "team":        all.push(...buildTeam(section, ac)); break;
        case "comparison":  all.push(...buildComparison(section, ac)); break;
        case "awards":      all.push(...buildAwards(section, ac)); break;
        case "financials":  all.push(...buildFinancials(section, ac)); break;
      }
    }
  }

  const doc = new Document({
    creator: report.meta.preparedBy ?? report.meta.companyName,
    title: `${report.meta.reportTitle} ${report.meta.reportYear}`,
    description: `Annual Report — ${report.meta.companyName}`,
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
        },
        {
          reference: "numbered",
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
        },
      ],
    },
    styles: {
      default: { document: { run: { font: "Arial", size: PT(11) } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: PT(20), bold: true, font: "Arial", color: ac },
          paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: PT(15), bold: true, font: "Arial", color: ac },
          paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: { default: buildHeader(report.meta) },
      footers: { default: buildFooter(report.meta) },
      children: all as Paragraph[],
    }],
  });

  return Packer.toBuffer(doc);
}
