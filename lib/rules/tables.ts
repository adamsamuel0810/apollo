import { Slide, Shape, TableModel } from "../pptx/types";
import { Finding, RuleContext, makeFindingId } from "./types";
import { cellRect } from "../pptx/render";
import { centerX, emuToIn } from "./util";
import { TABLE_COLORS } from "../brand/palette";

const STATS_LABEL_RE = /^(median|mean|average|count|25th|75th|50th|percentile|%ile|p25|p50|p75|min|max)/i;
const HEADER_FILL = TABLE_COLORS.headerFill;

/**
 * A "data table" is a real tabular block (multiple columns and rows), as opposed
 * to small legend/key boxes (typically 2 columns). Brand table rules only apply
 * to data tables, which avoids flagging decorative legend mini-tables.
 */
function isDataTable(table: { rows: unknown[]; colWidthsEmu: unknown[] }): boolean {
  return table.colWidthsEmu.length >= 3 && table.rows.length >= 3;
}

export function tableRules(slide: Slide, _ctx: RuleContext): Finding[] {
  const findings: Finding[] = [];
  const tables = slide.shapes.filter((s) => s.kind === "table" && s.table);

  for (const shape of tables) {
    const table = shape.table!;
    if (!isDataTable(table)) continue;
    fontConsistency(shape, table, slide.index, findings);
    sideBorders(shape, table, slide.index, findings);
    statsRowColor(shape, table, slide.index, findings);
    mergedBodyCells(shape, table, slide.index, findings);
    tableTitleCentered(shape, slide, findings);
  }

  return findings;
}

function firstRunSize(cellRuns: { sizePt: number | null; text: string }[]): number | null {
  const r = cellRuns.find((x) => x.text.trim());
  return r?.sizePt ?? null;
}

function fontConsistency(
  shape: Shape,
  table: TableModel,
  slideIndex: number,
  findings: Finding[]
) {
  const sizes: { size: number; row: number; col: number }[] = [];
  for (let ri = 0; ri < table.rows.length; ri++) {
    for (let ci = 0; ci < table.rows[ri].length; ci++) {
      const cell = table.rows[ri][ci];
      if (cell.hMerge || cell.vMerge) continue;
      const s = firstRunSize(cell.runs);
      if (s != null) sizes.push({ size: s, row: ri, col: ci });
    }
  }
  if (sizes.length < 4) return;

  const freq = new Map<number, number>();
  for (const s of sizes) freq.set(s.size, (freq.get(s.size) || 0) + 1);
  let dominant = sizes[0].size;
  let dn = 0;
  for (const [sz, n] of freq) if (n > dn) ((dn = n), (dominant = sz));

  // Only flag if there is a clear majority and a minority of outliers.
  if (dn / sizes.length < 0.6) return;

  const reported = new Set<number>();
  for (const s of sizes) {
    if (s.size === dominant) continue;
    if (reported.has(s.size)) continue;
    reported.add(s.size);
    findings.push({
      id: makeFindingId("table-font-consistency", slideIndex),
      ruleId: "table-font-consistency",
      category: "table",
      severity: "warning",
      confidence: 0.75,
      slideIndex,
      title: "Inconsistent font size in table",
      detail: `A cell uses ${s.size} pt while most of the table uses ${dominant} pt. Font size should be consistent across a table.`,
      evidence: `${s.size} pt vs ${dominant} pt`,
      suggestion: `Set the cell to ${dominant} pt to match the rest of the table.`,
      guideline: "Tables: Size 10–12 pt, consistent across the table",
      source: "deterministic",
      shapeId: shape.id,
      rect: cellRect(shape.rect, table, s.row, s.col) || shape.rect,
      cell: { row: s.row, col: s.col },
      fix: {
        kind: "font-size-set",
        slideIndex,
        shapeId: shape.id,
        cell: { row: s.row, col: s.col },
        fontSizePt: dominant,
      },
    });
  }
}

function sideBorders(
  shape: Shape,
  table: TableModel,
  slideIndex: number,
  findings: Finding[]
) {
  let hasSide = false;
  for (const row of table.rows) {
    if (row.length === 0) continue;
    const firstReal = row.find((c) => !c.hMerge && !c.vMerge);
    const lastReal = [...row].reverse().find((c) => !c.hMerge && !c.vMerge);
    if (firstReal && firstReal.borders.left) hasSide = true;
    if (lastReal && lastReal.borders.right) hasSide = true;
  }
  if (hasSide) {
    findings.push({
      id: makeFindingId("table-side-border", slideIndex),
      ruleId: "table-side-border",
      category: "table",
      severity: "warning",
      confidence: 0.7,
      slideIndex,
      title: "Table has left/right borders",
      detail: "ACME tables should not have left or right borders.",
      suggestion: "Remove the vertical borders on the outer left and right edges of the table.",
      guideline: "Tables: No left or right border on the table",
      source: "deterministic",
      shapeId: shape.id,
      rect: shape.rect,
      fix: {
        kind: "border-remove",
        slideIndex,
        shapeId: shape.id,
        borderScope: "table-sides",
      },
    });
  }
}

function rowFill(row: { fill: string | null }[]): string | null {
  const counts = new Map<string, number>();
  for (const c of row) if (c.fill) counts.set(c.fill, (counts.get(c.fill) || 0) + 1);
  let best: string | null = null;
  let bn = 0;
  for (const [f, n] of counts) if (n > bn) ((bn = n), (best = f));
  return best;
}

function firstNonEmpty(row: { text: string }[]): { text: string; col: number } | null {
  for (let i = 0; i < row.length; i++) {
    if (row[i].text.trim()) return { text: row[i].text.trim(), col: i };
  }
  return null;
}

function statsRowColor(
  shape: Shape,
  table: TableModel,
  slideIndex: number,
  findings: Finding[]
) {
  // Whether the table even uses the yellow stats convention anywhere.
  const hasYellowStats = table.rows.some(
    (row) => rowFill(row) === TABLE_COLORS.statsRow
  );
  if (!hasYellowStats) return; // no stats-shading convention to enforce

  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const label = firstNonEmpty(row);
    if (!label) continue;
    if (!STATS_LABEL_RE.test(label.text)) continue; // row label must be a statistic
    const fill = rowFill(row);
    if (fill === TABLE_COLORS.statsRow) continue; // already correct

    const isGreen = fill === TABLE_COLORS.clientRow;
    findings.push({
      id: makeFindingId("table-statsrow-color", slideIndex),
      ruleId: "table-statsrow-color",
      category: "table",
      severity: "warning",
      confidence: 0.78,
      slideIndex,
      title: "Statistics row is not shaded yellow",
      detail: isGreen
        ? `The "${label.text}" row contains summary statistics but is shaded green, which is reserved for client rows.`
        : `The "${label.text}" row contains summary statistics but is not shaded yellow like the other statistics rows.`,
      evidence: label.text.slice(0, 40),
      suggestion: `Shade the statistics row yellow (#${TABLE_COLORS.statsRow}).`,
      guideline: "Statistics Row: Shading Yellow, Accent 3 (#FFFFDB); green reserved for clients",
      source: "deterministic",
      shapeId: shape.id,
      rect: cellRect(shape.rect, table, ri, label.col) || shape.rect,
      cell: { row: ri, col: label.col },
      fix: {
        kind: "row-fill-set",
        slideIndex,
        shapeId: shape.id,
        rowIndex: ri,
        fillColor: TABLE_COLORS.statsRow,
      },
    });
  }
}

function mergedBodyCells(
  shape: Shape,
  table: TableModel,
  slideIndex: number,
  findings: Finding[]
) {
  // Vertical merges inside the table body can be ambiguous. Skip header rows
  // (blue fill) and the top two rows, which are commonly multi-line headers.
  for (let ri = 2; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    if (rowFill(row) === HEADER_FILL) continue;
    for (let ci = 0; ci < row.length; ci++) {
      const cell = row[ci];
      if (cell.rowSpan > 1 && cell.text.trim()) {
        findings.push({
          id: makeFindingId("table-merged-body", slideIndex),
          ruleId: "table-merged-body",
          category: "table",
          severity: "info",
          confidence: 0.5,
          slideIndex,
          title: "Merged cells in table body",
          detail: `"${cell.text.trim().slice(0, 40)}" spans multiple rows. Confirm the merge is intentional and clearly readable.`,
          evidence: cell.text.trim().slice(0, 50),
          suggestion: "Verify the merged cell is intentional; otherwise split it per row.",
          guideline: "Tables: favor clarity; merges should be clearly intentional",
          source: "deterministic",
          shapeId: shape.id,
          rect: cellRect(shape.rect, table, ri, ci) || shape.rect,
          cell: { row: ri, col: ci },
        });
        return; // one per table is enough
      }
    }
  }
}

function tableTitleCentered(shape: Shape, slide: Slide, findings: Finding[]) {
  if (!shape.rect) return;
  // Find a bold title-like text shape sitting just above the table. Exclude
  // legend/key labels (short, ending with ":") which are not table titles.
  const candidates = slide.shapes.filter((s) => {
    if (s.kind !== "text" || !s.rect || !s.text.trim()) return false;
    const label = s.text.trim();
    if (label.length < 6 || label.endsWith(":")) return false;
    const above = s.rect.y + s.rect.cy <= shape.rect!.y + 91440; // within ~0.1"
    const near = shape.rect!.y - (s.rect.y + s.rect.cy) < 91440 * 6; // within ~0.6"
    const horizOverlap =
      s.rect.x < shape.rect!.x + shape.rect!.cx && s.rect.x + s.rect.cx > shape.rect!.x;
    const bold = s.paragraphs.some((p) => p.runs.some((r) => r.bold && r.text.trim()));
    // The title's width should be comparable to the table's (not a stray label).
    const widthRatio = s.rect.cx / shape.rect!.cx;
    return above && near && horizOverlap && bold && widthRatio > 0.4;
  });
  if (candidates.length === 0) return;
  const titleShape = candidates.sort(
    (a, b) => b.rect!.y - a.rect!.y
  )[0];

  const tCenter = centerX(shape.rect);
  const titleCenter = centerX(titleShape.rect!);
  const offIn = Math.abs(tCenter - titleCenter) / 914400;
  if (offIn > 0.1) {
    findings.push({
      id: makeFindingId("table-title-centered", slide.index),
      ruleId: "table-title-centered",
      category: "table",
      severity: "info",
      confidence: 0.5,
      slideIndex: slide.index,
      title: "Table title not centered over table",
      detail: `The title "${titleShape.text.trim().slice(0, 40)}" is offset ${offIn.toFixed(2)}" from the table's horizontal center.`,
      evidence: titleShape.text.trim().slice(0, 40),
      suggestion: "Center the table title above the table.",
      guideline: "Table Title: Position Centered above table",
      source: "deterministic",
      shapeId: titleShape.id,
      rect: titleShape.rect,
    });
  }
}
