import { Slide, Shape, Rect } from "../pptx/types";
import { Finding, RuleContext, makeFindingId } from "./types";
import { isOnPalette, nearestPaletteColor, normalizeHex } from "../brand/palette";
import { cellRect } from "../pptx/render";

const MAX_PER_SLIDE = 4;
// Small tolerance so near-identical anti-aliased values don't trip the rule.
const TOLERANCE = 8;

export function colorRules(slide: Slide, _ctx: RuleContext): Finding[] {
  const found = new Map<
    string,
    {
      shape: Shape;
      rect: Rect | null;
      where: string;
      cell?: { row: number; col: number };
    }
  >();

  const consider = (
    hex: string | null,
    shape: Shape,
    rect: Rect | null,
    where: string,
    cell?: { row: number; col: number }
  ) => {
    const h = normalizeHex(hex);
    if (!h) return;
    if (isOnPalette(h, TOLERANCE)) return;
    const key = cell ? `${h}:${cell.row}:${cell.col}` : h;
    if (!found.has(key)) found.set(key, { shape, rect, where, cell });
  };

  for (const shape of slide.shapes) {
    // Shape fill
    consider(shape.fill, shape, shape.rect, `fill of "${shape.name}"`);
    // Run text colors
    for (const p of shape.paragraphs) {
      for (const r of p.runs) {
        if (r.text.trim()) consider(r.color, shape, shape.rect, `text in "${shape.name}"`);
      }
    }
    // Table cell fills + text
    if (shape.table) {
      for (let ri = 0; ri < shape.table.rows.length; ri++) {
        for (let ci = 0; ci < shape.table.rows[ri].length; ci++) {
          const cell = shape.table.rows[ri][ci];
          const rect = cellRect(shape.rect, shape.table, ri, ci) || shape.rect;
          consider(cell.fill, shape, rect, `cell fill`, { row: ri, col: ci });
          for (const r of cell.runs)
            if (r.text.trim()) consider(r.color, shape, rect, `cell text`, { row: ri, col: ci });
        }
      }
    }
  }

  const findings: Finding[] = [];
  let n = 0;
  for (const [hex, info] of found) {
    if (n >= MAX_PER_SLIDE) break;
    n++;
    const near = nearestPaletteColor(hex);
    const isFill = info.where.toLowerCase().includes("fill");
    findings.push({
      id: makeFindingId("off-palette-color", slide.index),
      ruleId: "off-palette-color",
      category: "color",
      severity: "warning",
      confidence: 0.5,
      slideIndex: slide.index,
      title: "Off-palette color",
      detail: `Color #${hex} (${info.where}) is not in the ACME color palette${
        near ? `; nearest brand color is #${near.color}` : ""
      }.`,
      evidence: `#${hex}`,
      suggestion: near
        ? `Use the closest ACME palette color #${near.color}, or another approved brand color.`
        : "Use an approved ACME palette color.",
      guideline: "Colors: Drawn from ACME Color Palette",
      source: "deterministic",
      shapeId: info.shape.id,
      rect: info.rect,
      cell: info.cell,
      fix: near
        ? {
            kind: "color-set",
            slideIndex: slide.index,
            shapeId: info.shape.id,
            fromColor: hex,
            toColor: near.color,
            colorTarget: isFill ? "fill" : "text",
            cell: info.cell,
          }
        : undefined,
    });
  }
  return findings;
}
