// ════════════════════════════════════════════════════════════
//  Annual Report JSON Schema — Complete Types
//  Covers every section type found in real annual reports
// ════════════════════════════════════════════════════════════

// ── Chart Types ──────────────────────────────────────────────
export type ChartType = "bar" | "line" | "pie" | "doughnut" | "horizontalBar" | "stackedBar" | "area";

// ── Subsection inside a text body ────────────────────────────
export type SubSection =
  | { type: "paragraph"; text: string }
  | { type: "bullets";   items: string[]; ordered?: boolean }
  | { type: "heading";   text: string }   // small sub-heading inside the body

export interface ChartDataset {
  label: string;
  data: number[];
  color?: string;        // hex e.g. "#1B3A6B"
  fill?: boolean;        // for area/line charts
}

// ── Section: Chart ────────────────────────────────────────────
export interface ChartSection {
  type: "chart";
  title: string;
  chartType: ChartType;
  labels: string[];
  datasets: ChartDataset[];
  caption?: string;
  yAxisLabel?: string;
  xAxisLabel?: string;
  yAxisPrefix?: string;  // e.g. "$"
  yAxisSuffix?: string;  // e.g. "%"
  showLegend?: boolean;
  showValues?: boolean;  // show value labels on bars
}

// ── Section: Table ────────────────────────────────────────────
export interface TableSection {
  type: "table";
  title: string;
  headers: string[];
  rows: (string | number)[][];
  caption?: string;
  highlightLastRow?: boolean;    // totals row
  highlightFirstColumn?: boolean; // row headers
  columnAlignments?: ("left" | "right" | "center")[];
}

// ── Section: KPI Metrics ──────────────────────────────────────
export interface KpiItem {
  label: string;
  value: string;
  delta?: string;              // e.g. "+12%" or "-3 pp"
  deltaPositive?: boolean;     // true = green, false = red
  subLabel?: string;           // small text below value e.g. "USD Millions"
  icon?: string;               // single emoji icon e.g. "📈"
}

export interface KpiSection {
  type: "kpi";
  title: string;
  items: KpiItem[];
  columns?: 2 | 3 | 4;        // default 4
}

// ── Section: Text ─────────────────────────────────────────────
export interface TextSection {
  type: "text";
  title: string;
  body?: string;                // simple body text (optional if subsections used)
  subsections?: SubSection[];   // ← new
  quote?: string;
  highlight?: string;
  subtitle?: string;
}

// ── Section: Bullet / Numbered List ──────────────────────────
export interface BulletSection {
  type: "bullets";
  title: string;
  items: string[];
  ordered?: boolean;           // numbered list if true
  columns?: 1 | 2;            // two-column layout
}

// ── Section: Image ────────────────────────────────────────────
export interface ImageSection {
  type: "image";
  title?: string;
  path: string;                // file path relative to JSON dir
  width?: number;              // display width in px (default 500)
  height?: number;             // display height in px (default 300)
  caption?: string;
  align?: "left" | "center" | "right";
}

// ── Section: Two-Column ───────────────────────────────────────
// Renders two text blocks side by side (useful for segment summaries)
export interface TwoColumnSection {
  type: "twoColumn";
  title?: string;
  left: {
    heading: string;
    body: string;
  };
  right: {
    heading: string;
    body: string;
  };
}

// ── Section: Text Two Column ──────────────────────────────────
export interface TextTwoColumnSection {
  type: "textTwoColumn";
  title: string;
  body?: string;                // auto-split by paragraph count
  subsections?: SubSection[];   // ← new — split across two columns
  quote?: string;
}

// ── Section: Divider / Section Break ─────────────────────────
export interface DividerSection {
  type: "divider";
  label?: string;              // optional centered label on the line
}

// ── Section: Quote / Testimonial ─────────────────────────────
export interface QuoteSection {
  type: "quote";
  text: string;
  attribution?: string;        // e.g. "— CEO, John Smith"
  role?: string;               // e.g. "Chief Executive Officer"
}

// ── Section: Timeline ────────────────────────────────────────
export interface TimelineItem {
  year: string;
  title: string;
  description: string;
}

export interface TimelineSection {
  type: "timeline";
  title: string;
  items: TimelineItem[];
}

// ── Section: Team / Board Members ────────────────────────────
export interface TeamMember {
  name: string;
  role: string;
  bio?: string;
  tenure?: string;
}

export interface TeamSection {
  type: "team";
  title: string;
  members: TeamMember[];
}

// ── Section: Comparison / Side-by-side Table ─────────────────
export interface ComparisonRow {
  label: string;
  values: string[];
}

export interface ComparisonSection {
  type: "comparison";
  title: string;
  columns: string[];           // column headers e.g. ["Feature", "Basic", "Pro", "Enterprise"]
  rows: ComparisonRow[];
  caption?: string;
}

// ── Section: Awards / Recognitions ───────────────────────────
export interface AwardItem {
  title: string;
  body: string;
  year?: string;
}

export interface AwardsSection {
  type: "awards";
  title: string;
  items: AwardItem[];
}

// ── Section: Financials (Income Statement / Balance Sheet) ───
export interface FinancialsSection {
  type: "financials";
  title: string;
  currency?: string;           // e.g. "USD" or "INR"
  unit?: string;               // e.g. "Millions" or "Crores"
  periods: string[];           // column headers e.g. ["FY2022","FY2023","FY2024"]
  items: {
    label: string;
    values: (string | number)[];
    bold?: boolean;            // for section totals
    indent?: boolean;          // for sub-items
    separator?: boolean;       // draws a line above this row
  }[];
  caption?: string;
}

// ── Section: Signature / Closing Block ───────────────────────
// Used at the end of chairman/CEO letters
export interface SignatureSection {
  type: "signature";
  closing?: string;            // e.g. "Warm regards,"
  name: string;                // signatory's full name
  role: string;                // e.g. "Chairman" or "CEO & Managing Director"
  nameColor?: string;          // hex WITHOUT # — defaults to accentColor
}

// ── Section: Rich Text ────────────────────────────────────────
// Paragraphs with inline bold/italic runs — ideal for stats-heavy text
export interface RichRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;              // hex WITHOUT # for colored inline text
}

export interface RichParagraph {
  runs: RichRun[];
}

export interface RichTextSection {
  type: "richText";
  title: string;
  subtitle?: string;
  highlight?: string;          // yellow callout box above the paragraphs
  paragraphs: RichParagraph[];
}

// ── Section: Page Break ───────────────────────────────────────
// Forces the next section to start on a new page in the .docx output
export interface PageBreakSection {
  type: "pageBreak";
}

// ── Union of all Section types ────────────────────────────────
export type Section =
  | TextSection
  | RichTextSection
  | KpiSection
  | TableSection
  | ChartSection
  | ImageSection
  | BulletSection
  | TwoColumnSection
  | TextTwoColumnSection
  | DividerSection
  | PageBreakSection
  | QuoteSection
  | SignatureSection
  | TimelineSection
  | TeamSection
  | ComparisonSection
  | AwardsSection
  | FinancialsSection;

// ── Chapter & Report ─────────────────────────────────────────
export interface ReportChapter {
  title: string;
  subtitle?: string;           // optional subtitle under chapter heading
  sections: Section[];
}

export interface ReportMeta {
  companyName: string;
  tagline?: string;
  reportTitle: string;         // e.g. "Annual Report" or "Integrated Annual Report"
  reportYear: string;
  preparedBy?: string;
  website?: string;
  logoPath?: string;           // path to logo image (png/jpg) relative to JSON
  coverImagePath?: string;     // full-page background cover image
  accentColor?: string;        // hex WITHOUT # e.g. "1B3A6B"
  secondaryColor?: string;     // light tint, hex WITHOUT # e.g. "E8EDF5"
  darkColor?: string;          // very dark version of accent e.g. "0D1F3C"
  stockExchange?: string;      // e.g. "NSE / BSE"
  tickerSymbol?: string;       // e.g. "TCS"
  fiscalYear?: string;         // e.g. "April 2023 – March 2024"
  registeredOffice?: string;
  cin?: string;                // Company Identification Number
}

export interface AnnualReport {
  meta: ReportMeta;
  chapters: ReportChapter[];
}
