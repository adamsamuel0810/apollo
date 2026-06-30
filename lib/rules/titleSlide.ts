import { Slide } from "../pptx/types";
import { Finding, RuleContext, makeFindingId } from "./types";
import { visibleRuns, shapeText, bodyShapes } from "./util";

// Standalone DRAFT/CONFIDENTIAL-style designators (uppercase token).
const DRAFT_RE = /\b(DRAFT|WIP|INTERNAL ONLY|DO NOT DISTRIBUTE|PRELIMINARY)\b/;

const DATE_RE =
  /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b|\b\d{1,2}\/\d{1,2}\/\d{2,4}\b|\b\d{4}-\d{2}-\d{2}\b/;

export function draftDesignatorRule(slide: Slide, _ctx: RuleContext): Finding[] {
  const findings: Finding[] = [];
  for (const shape of slide.shapes) {
    for (const r of visibleRuns(shape)) {
      const m = DRAFT_RE.exec(r.text);
      if (m) {
        findings.push({
          id: makeFindingId("draft-designator", slide.index),
          ruleId: "draft-designator",
          category: "title",
          severity: "error",
          confidence: slide.kind === "title" ? 0.97 : 0.9,
          slideIndex: slide.index,
          title: `"${m[1]}" designator present`,
          detail: `A "${m[1]}" marker appears in "${shape.name}". Client-ready ACME decks must not contain draft/internal designators.`,
          evidence: r.text.trim().slice(0, 60),
          suggestion: `Remove the "${m[1]}" designator before distribution.`,
          guideline: 'Title Slides: No "DRAFT" or other designator',
          source: "deterministic",
          shapeId: shape.id,
          rect: shape.rect,
          fix: {
            kind: "text-replace",
            slideIndex: slide.index,
            shapeId: shape.id,
            search: `\\b${m[1]}\\b\\s*`,
            replace: "",
            regex: true,
            caseInsensitive: true,
          },
        });
        break;
      }
    }
  }
  return findings;
}

export function titleSlideMetaRule(slide: Slide, _ctx: RuleContext): Finding[] {
  if (slide.kind !== "title") return [];
  const findings: Finding[] = [];

  const allText = slide.shapes.map((s) => shapeText(s)).join(" ");
  if (!DATE_RE.test(allText)) {
    findings.push({
      id: makeFindingId("title-date", slide.index),
      ruleId: "title-date",
      category: "title",
      severity: "info",
      confidence: 0.5,
      slideIndex: slide.index,
      title: "No clear date on title slide",
      detail: "A valid presentation date should be noted on the title slide.",
      suggestion: "Add the presentation date to the title slide.",
      guideline: "Title Slides: Valid date is noted",
      source: "deterministic",
    });
  }

  // Authors: expect at least one non-title text block beyond the date/footer.
  const authorCandidates = bodyShapes(slide).filter((s) => {
    const t = shapeText(s);
    return t.length > 0 && !DATE_RE.test(t) && !DRAFT_RE.test(t);
  });
  if (authorCandidates.length === 0) {
    findings.push({
      id: makeFindingId("title-authors", slide.index),
      ruleId: "title-authors",
      category: "title",
      severity: "info",
      confidence: 0.45,
      slideIndex: slide.index,
      title: "Authors may not be identified",
      detail: "The title slide does not appear to identify the authors.",
      suggestion: "Add the author name(s) and role(s) to the title slide.",
      guideline: "Title Slides: Authors are identified",
      source: "deterministic",
    });
  }

  return findings;
}
