import { Slide } from "../pptx/types";
import { Finding, RuleContext, makeFindingId } from "./types";
import {
  isCalibri,
  visibleRuns,
  endsWithPunctuation,
  titleText,
} from "./util";
import { TITLE, TYPEFACE, SIZE_TOLERANCE_PT } from "../brand/guidelines";

export function typographyRules(slide: Slide, _ctx: RuleContext): Finding[] {
  const findings: Finding[] = [];

  // --- Font family must be Calibri across all text ---
  for (const shape of slide.shapes) {
    const offenders = new Set<string>();
    for (const r of visibleRuns(shape)) {
      if (r.font && !isCalibri(r.font)) offenders.add(r.font);
    }
    if (offenders.size > 0) {
      findings.push({
        id: makeFindingId("font-family", slide.index),
        ruleId: "font-family",
        category: "typography",
        severity: "error",
        confidence: 0.95,
        slideIndex: slide.index,
        title: "Non-Calibri font detected",
        detail: `Text in "${shape.name}" uses ${[...offenders].join(", ")}, but ACME decks use ${TYPEFACE}.`,
        evidence: [...offenders].join(", "),
        suggestion: `Change the typeface to ${TYPEFACE}.`,
        guideline: "Typeface: Calibri (Headings/Body)",
        source: "deterministic",
        shapeId: shape.id,
        rect: shape.rect,
      });
    }
  }

  // --- Slide title checks ---
  const title = slide.titleShape;
  const tText = titleText(slide);
  if (title && tText) {
    const runs = visibleRuns(title);

    // Size should be 24pt.
    const badSize = runs.find(
      (r) => r.sizePt != null && Math.abs(r.sizePt - TITLE.sizePt) > SIZE_TOLERANCE_PT
    );
    if (badSize) {
      findings.push({
        id: makeFindingId("title-size", slide.index),
        ruleId: "title-size",
        category: "title",
        severity: "warning",
        confidence: 0.85,
        slideIndex: slide.index,
        title: "Slide title is not 24 pt",
        detail: `The slide title renders at ${badSize.sizePt} pt; ACME titles are ${TITLE.sizePt} pt.`,
        evidence: `${badSize.sizePt} pt`,
        suggestion: `Set the title font size to ${TITLE.sizePt} pt.`,
        guideline: "Slide Title: Size 24 pt",
        source: "deterministic",
        shapeId: title.id,
        rect: title.rect,
      });
    }

    // Bold (content/section slide titles must be Calibri Bold).
    if (slide.kind !== "title") {
      const anyBold = runs.some((r) => r.bold === true);
      if (!anyBold && runs.length > 0) {
        findings.push({
          id: makeFindingId("title-bold", slide.index),
          ruleId: "title-bold",
          category: "title",
          severity: "warning",
          confidence: 0.8,
          slideIndex: slide.index,
          title: "Slide title is not bold",
          detail: "Content slide titles should be Calibri Bold.",
          suggestion: "Apply bold to the slide title.",
          guideline: "Slide Title: Typeface Calibri Bold",
          source: "deterministic",
          shapeId: title.id,
          rect: title.rect,
        });
      }
    }

    // No ending punctuation.
    if (endsWithPunctuation(tText)) {
      findings.push({
        id: makeFindingId("title-punct", slide.index),
        ruleId: "title-punct",
        category: "title",
        severity: "warning",
        confidence: 0.9,
        slideIndex: slide.index,
        title: "Slide title ends with punctuation",
        detail: `The title ends with a trailing "${tText.trim().slice(-1)}".`,
        evidence: tText.trim().slice(-40),
        suggestion: "Remove the trailing punctuation from the title.",
        guideline: "Slide Title: No ending punctuation",
        source: "deterministic",
        shapeId: title.id,
        rect: title.rect,
      });
    }
    // Headline length (> 3 lines) is handled by the AI layer, which can reason
    // about rendered wrapping/autofit far more reliably than a text estimate.
  }

  return findings;
}
