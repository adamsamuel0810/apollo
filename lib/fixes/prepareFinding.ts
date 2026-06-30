import { parseDeck } from "../pptx/parser";
import { computeAggregates } from "../deck/aggregate";
import { Finding } from "../rules/types";
import { inferFix } from "./inferFix";
import { resolveAiFix, resolveAiShapeFix } from "../ai/autofix";
import { titleText, estimateLines, emuToIn, shapeText } from "../rules/util";
import { TITLE } from "../brand/guidelines";

/** Ensure a finding has a resolved fix before applying to the PPTX. */
export async function prepareFindingForApply(
  finding: Finding,
  pptx: Buffer
): Promise<Finding> {
  const deck = await parseDeck(pptx);
  const aggregates = computeAggregates(deck);

  let fix: NonNullable<Finding["fix"]> | null =
    finding.fix ?? inferFix(finding, deck, aggregates);
  if (!fix) return finding;

  if (fix.kind === "ai-rewrite" || (fix.kind === "title-rewrite" && !fix.newText)) {
    const slide = deck.slides.find((s) => s.index === finding.slideIndex);
    if (slide) {
      const shapeId = fix.shapeId;
      const shape = shapeId
        ? slide.shapes.find((s) => s.id === shapeId)
        : slide.titleShape;
      const text = shape ? shapeText(shape) : titleText(slide) || finding.evidence || "";
      const isTitle =
        finding.ruleId === "title-lines" ||
        shape?.placeholderType === "title" ||
        shape?.placeholderType === "ctrTitle";

      if (isTitle) {
        const title = slide.titleShape;
        const tText = titleText(slide) || text;
        const w = title?.rect ? emuToIn(title.rect.cx) : 8.5;
        const lines = tText ? estimateLines(tText, TITLE.sizePt, w) : TITLE.maxLines + 1;
        const resolved = await resolveAiFix(
          { ...fix, kind: "ai-rewrite" },
          tText,
          lines,
          w
        );
        if (resolved) fix = resolved;
      } else {
        const resolved = await resolveAiShapeFix({ ...fix, kind: "ai-rewrite" }, text);
        if (resolved) fix = resolved;
      }
    }
  }

  return { ...finding, fix };
}

export async function prepareFindingsForApply(
  findings: Finding[],
  pptx: Buffer
): Promise<Finding[]> {
  const out: Finding[] = [];
  for (const f of findings) {
    out.push(await prepareFindingForApply(f, pptx));
  }
  return out;
}
