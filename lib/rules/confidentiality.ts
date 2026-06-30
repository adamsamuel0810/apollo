import { Slide, Shape } from "../pptx/types";
import { Finding, RuleContext, makeFindingId } from "./types";
import { normalizeText, intersectionArea, shapeText } from "./util";
import { EMU_PER_INCH } from "../brand/guidelines";

export function confidentialityRule(slide: Slide, ctx: RuleContext): Finding[] {
  const findings: Finding[] = [];
  const canonical = normalizeText(ctx.aggregates.confidentialityText || "");
  if (!canonical) return findings;

  // Candidate footer shapes: the ftr placeholder, or any shape whose text
  // resembles the confidentiality statement.
  const footers: Shape[] = slide.shapes.filter(
    (s) =>
      s.placeholderType === "ftr" ||
      (s.text.trim().length > 0 && jaccard(normalizeText(s.text), canonical) >= 0.5)
  );

  const withText = footers.filter((s) => s.text.trim().length > 0);

  if (withText.length === 0) {
    const ftr = slide.shapes.find((s) => s.placeholderType === "ftr");
    findings.push({
      id: makeFindingId("confidentiality-missing", slide.index),
      ruleId: "confidentiality-missing",
      category: "confidentiality",
      severity: "error",
      confidence: 0.9,
      slideIndex: slide.index,
      title: "Confidentiality statement missing",
      detail:
        "Every slide (including title and section headers) must display the confidentiality statement.",
      suggestion: `Add the confidentiality statement: "${ctx.aggregates.confidentialityText}".`,
      guideline:
        "Confidentiality Statement: Appears on every slide, including title and section header slides",
      source: "deterministic",
      rect: ctx.aggregates.footerRect,
      shapeId: ftr?.id,
      fix: ctx.aggregates.footerRect
        ? {
            kind: "ensure-footer",
            slideIndex: slide.index,
            shapeId: ftr?.id,
            newText: ctx.aggregates.confidentialityText || "",
            targetRect: ctx.aggregates.footerRect,
          }
        : undefined,
    });
    return findings;
  }

  const footer = withText[0];
  const footerText = normalizeText(footer.text);

  // Modified / paraphrased legal language.
  if (footerText !== canonical) {
    const sim = jaccard(footerText, canonical);
    if (sim < 0.97) {
      findings.push({
        id: makeFindingId("confidentiality-modified", slide.index),
        ruleId: "confidentiality-modified",
        category: "confidentiality",
        severity: "error",
        confidence: sim > 0.6 ? 0.85 : 0.7,
        slideIndex: slide.index,
        title: "Confidentiality statement modified",
        detail:
          "The confidentiality / legal language differs from the standard wording used elsewhere in the deck and must not be paraphrased.",
        evidence: footer.text.trim().slice(0, 90),
        suggestion: `Restore the exact wording: "${ctx.aggregates.confidentialityText}".`,
        guideline: "Confidentiality Statement: Legal language is not modified or paraphrased",
        source: "deterministic",
        shapeId: footer.id,
        rect: footer.rect,
        fix: {
          kind: "title-rewrite",
          slideIndex: slide.index,
          shapeId: footer.id,
          newText: ctx.aggregates.confidentialityText || "",
        },
      });
    }
  }

  // Position moved from the master anchor.
  const anchor = ctx.aggregates.footerRect;
  if (anchor && footer.rect) {
    const dx = Math.abs(footer.rect.x - anchor.x) / EMU_PER_INCH;
    const dy = Math.abs(footer.rect.y - anchor.y) / EMU_PER_INCH;
    if (dx > 0.08 || dy > 0.08) {
      findings.push({
        id: makeFindingId("confidentiality-moved", slide.index),
        ruleId: "confidentiality-moved",
        category: "confidentiality",
        severity: "warning",
        confidence: 0.75,
        slideIndex: slide.index,
        title: "Confidentiality statement moved",
        detail: `The statement is offset by ${dx.toFixed(2)}" horizontally and ${dy.toFixed(2)}" vertically from its standard position.`,
        suggestion: "Return the confidentiality statement to its master-slide position.",
        guideline: 'Confidentiality Statement: Position same as "Title and Content" master slide',
        source: "deterministic",
        shapeId: footer.id,
        rect: footer.rect,
        fix: anchor
          ? {
              kind: "shape-reposition",
              slideIndex: slide.index,
              shapeId: footer.id,
              targetRect: anchor,
            }
          : undefined,
      });
    }
  }

  // Overlapped by other content.
  if (footer.rect) {
    const footerArea = footer.rect.cx * footer.rect.cy || 1;
    for (const s of slide.shapes) {
      if (s === footer) continue;
      if (s.placeholderType === "sldNum" || s.placeholderType === "dt") continue;
      if (!s.rect) continue;
      if (s.text.trim().length === 0 && s.kind === "text" && !s.fill) continue;
      const overlap = intersectionArea(footer.rect, s.rect) / footerArea;
      if (overlap > 0.45) {
        findings.push({
          id: makeFindingId("confidentiality-overlap", slide.index),
          ruleId: "confidentiality-overlap",
          category: "confidentiality",
          severity: "warning",
          confidence: 0.65,
          slideIndex: slide.index,
          title: "Confidentiality statement overlapped",
          detail: `"${s.name}" overlaps the confidentiality statement, which can obscure the legal text.`,
          evidence: shapeText(s).slice(0, 60),
          suggestion: "Reposition the overlapping element so the statement is fully legible.",
          guideline: "Confidentiality Statement: Position / legibility",
          source: "deterministic",
          shapeId: footer.id,
          rect: footer.rect,
        });
        break;
      }
    }
  }

  return findings;
}

function jaccard(a: string, b: string): number {
  const sa = new Set(a.toLowerCase().split(/\s+/).filter(Boolean));
  const sb = new Set(b.toLowerCase().split(/\s+/).filter(Boolean));
  if (sa.size === 0 && sb.size === 0) return 1;
  let inter = 0;
  for (const w of sa) if (sb.has(w)) inter++;
  return inter / (sa.size + sb.size - inter);
}
