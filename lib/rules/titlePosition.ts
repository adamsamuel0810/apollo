import { Slide } from "../pptx/types";
import { Finding, RuleContext, makeFindingId } from "./types";
import { EMU_PER_INCH } from "../brand/guidelines";

// Threshold (inches) beyond which the title is considered "moved from default".
const MOVE_TOL_IN = 0.12;

export function titlePositionRule(slide: Slide, ctx: RuleContext): Finding[] {
  const title = slide.titleShape;
  if (!title || !title.rect || !slide.layoutName) return [];
  const layoutRect = ctx.aggregates.titleRectByLayout[slide.layoutName];
  if (!layoutRect) return [];

  const dx = Math.abs(title.rect.x - layoutRect.x) / EMU_PER_INCH;
  const dy = Math.abs(title.rect.y - layoutRect.y) / EMU_PER_INCH;
  const dw = Math.abs(title.rect.cx - layoutRect.cx) / EMU_PER_INCH;
  const dh = Math.abs(title.rect.cy - layoutRect.cy) / EMU_PER_INCH;

  const maxDelta = Math.max(dx, dy, dw, dh);
  if (maxDelta <= MOVE_TOL_IN) return [];

  const parts: string[] = [];
  if (dx > MOVE_TOL_IN) parts.push(`${dx.toFixed(2)}" horizontally`);
  if (dy > MOVE_TOL_IN) parts.push(`${dy.toFixed(2)}" vertically`);
  if (dw > MOVE_TOL_IN || dh > MOVE_TOL_IN) parts.push("and was resized");

  return [
    {
      id: makeFindingId("title-moved", slide.index),
      ruleId: "title-moved",
      category: "title",
      severity: "info",
      confidence: 0.45,
      slideIndex: slide.index,
      title: "Title moved from default position",
      detail: `The title box differs from its layout default (${parts.join(", ")}). Titles should keep the master's size and position.`,
      suggestion: "Reset the title placeholder to the layout's default size and position.",
      guideline: 'Slide Title: Size/Position same as "Title and Content" master slide',
      source: "deterministic",
      shapeId: title.id,
      rect: title.rect,
    },
  ];
}
