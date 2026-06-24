import { Slide, Shape } from "../pptx/types";
import { Finding, RuleContext, makeFindingId, FindingCategory, Severity } from "./types";
import { visibleRuns } from "./util";
import { cellRect } from "../pptx/render";
import { TABLE_COLORS } from "../brand/palette";

interface TextPattern {
  ruleId: string;
  re: RegExp;
  title: string;
  detail: string;
  suggestion: string;
  guideline: string;
  severity: Severity;
  confidence: number;
  category: FindingCategory;
}

// Patterns checked against general slide text (prose, shapes).
const PROSE_PATTERNS: TextPattern[] = [
  {
    ruleId: "term-percentile",
    re: /\bpercentile\b/i,
    title: '"Percentile" should be "%ile"',
    detail: 'ACME labels percentiles as "%ile", not the spelled-out "Percentile".',
    suggestion: 'Replace "Percentile" with "%ile".',
    guideline: 'Tables: Percentiles are labeled with "%ile", not "Percentile"',
    severity: "warning",
    confidence: 0.82,
    category: "terminology",
  },
  {
    ruleId: "term-tgt",
    re: /\bTGT\b/,
    title: '"TGT" should be spelled out as "Target"',
    detail: 'ACME spells out "Target" and does not abbreviate it as "TGT".',
    suggestion: 'Replace "TGT" with "Target".',
    guideline: 'Other: Spell out "Target" (do not abbreviate as "TGT")',
    severity: "warning",
    confidence: 0.9,
    category: "terminology",
  },
];

// Patterns checked specifically against table header cells.
const HEADER_PATTERNS: TextPattern[] = [
  {
    ruleId: "term-incumbent",
    re: /\bincumbent\b/i,
    title: '"Incumbent" should be "Executive"',
    detail: 'Employee-name columns should be titled "Executive", not "Incumbent".',
    suggestion: 'Rename the column header to "Executive".',
    guideline: 'Table Header: Use "Executive" (not "Incumbent") for the employee column',
    severity: "warning",
    confidence: 0.85,
    category: "terminology",
  },
  {
    ruleId: "term-company-name",
    re: /^company name$/i,
    title: '"Company Name" should be "Company"',
    detail: 'Company columns should be labeled "Company" rather than "Company Name".',
    suggestion: 'Rename the column header to "Company".',
    guideline: "Table Header: consistent, concise column terminology",
    severity: "info",
    confidence: 0.6,
    category: "terminology",
  },
];

export function terminologyRules(slide: Slide, _ctx: RuleContext): Finding[] {
  const findings: Finding[] = [];

  for (const shape of slide.shapes) {
    if (shape.kind === "table" && shape.table) {
      scanTable(shape, slide.index, findings);
    } else {
      scanProse(shape, slide.index, findings);
    }
  }

  return findings;
}

function scanProse(shape: Shape, slideIndex: number, findings: Finding[]) {
  const seen = new Set<string>();
  for (const r of visibleRuns(shape)) {
    for (const p of PROSE_PATTERNS) {
      if (seen.has(p.ruleId)) continue;
      if (p.re.test(r.text)) {
        seen.add(p.ruleId);
        findings.push({
          id: makeFindingId(p.ruleId, slideIndex),
          ruleId: p.ruleId,
          category: p.category,
          severity: p.severity,
          confidence: p.confidence,
          slideIndex,
          title: p.title,
          detail: p.detail,
          evidence: r.text.trim().slice(0, 60),
          suggestion: p.suggestion,
          guideline: p.guideline,
          source: "deterministic",
          shapeId: shape.id,
          rect: shape.rect,
        });
      }
    }
  }
}

function rowIsHeader(row: { fill: string | null }[], ri: number): boolean {
  // Header rows are shaded with the brand header color; can span >1 row.
  const headerCells = row.filter((c) => c.fill === TABLE_COLORS.headerFill).length;
  if (row.length > 0 && headerCells / row.length >= 0.5) return true;
  return ri === 0; // fallback: first row
}

function scanTable(shape: Shape, slideIndex: number, findings: Finding[]) {
  const table = shape.table!;
  // Prose patterns apply to every cell; header patterns only to header rows.
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const isHeader = rowIsHeader(row, ri);
    for (let ci = 0; ci < row.length; ci++) {
      const cell = row[ci];
      const text = cell.text.trim();
      if (!text) continue;
      const patterns = isHeader ? [...HEADER_PATTERNS, ...PROSE_PATTERNS] : PROSE_PATTERNS;
      for (const p of patterns) {
        if (p.re.test(text)) {
          findings.push({
            id: makeFindingId(p.ruleId, slideIndex),
            ruleId: p.ruleId,
            category: p.category,
            severity: p.severity,
            confidence: p.confidence,
            slideIndex,
            title: p.title,
            detail: p.detail,
            evidence: text.slice(0, 60),
            suggestion: p.suggestion,
            guideline: p.guideline,
            source: "deterministic",
            shapeId: shape.id,
            rect: cellRect(shape.rect, table, ri, ci) || shape.rect,
            cell: { row: ri, col: ci },
          });
        }
      }
    }
  }
}
