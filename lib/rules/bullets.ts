import { Slide, Shape } from "../pptx/types";
import { Finding, RuleContext, makeFindingId } from "./types";

const TRAILING = /[.;]$/;

export function bulletRules(slide: Slide, _ctx: RuleContext): Finding[] {
  const findings: Finding[] = [];

  for (const shape of slide.shapes) {
    if (shape.kind !== "text") continue;
    // Only consider shapes that are genuine bulleted lists (2+ bulleted paras).
    const bulletParas = shape.paragraphs.filter(
      (p) => (p.bulletType === "char" || p.bulletType === "auto") && p.text.trim()
    );
    if (bulletParas.length < 2) continue;

    let flagged = false;
    for (const p of bulletParas) {
      const text = p.text.trim();
      if (TRAILING.test(text)) {
        flagged = true;
        break;
      }
    }
    if (flagged) {
      findings.push({
        id: makeFindingId("bullet-punct", slide.index),
        ruleId: "bullet-punct",
        category: "bullets",
        severity: "info",
        confidence: 0.55,
        slideIndex: slide.index,
        title: "Bullet ends with punctuation",
        detail: `One or more bullets in "${shape.name}" end with punctuation. Bullet text should not end in punctuation.`,
        suggestion: "Remove trailing punctuation from bullet text.",
        guideline: "Bulleted Lists: Bullet text does not end in punctuation",
        source: "deterministic",
        shapeId: shape.id,
        rect: shape.rect,
      });
    }
  }

  return findings;
}
