import { DeckModel, Slide, Rect } from "../pptx/types";

export type Severity = "error" | "warning" | "info";

export type FindingSource = "deterministic" | "ai";

export type FindingCategory =
  | "typography"
  | "title"
  | "confidentiality"
  | "bullets"
  | "table"
  | "color"
  | "placement"
  | "footnote"
  | "terminology"
  | "consistency"
  | "content";

export interface Finding {
  id: string;
  ruleId: string;
  category: FindingCategory;
  severity: Severity;
  /** 0..1 confidence; drives noise filtering and ordering. */
  confidence: number;
  slideIndex: number;
  title: string;
  /** Human-readable detail of what is wrong. */
  detail: string;
  /** The exact offending text/value, when applicable. */
  evidence?: string;
  /** Suggested fix. */
  suggestion?: string;
  /** Guideline citation (short reference into ACME-brand-guidelines.md). */
  guideline: string;
  source: FindingSource;
  /** Shape id (or synthetic locator) the finding points at, for highlighting. */
  shapeId?: string;
  /** Bounding box (EMU) for the highlight overlay; falls back to shape rect. */
  rect?: Rect | null;
  /** Optional table cell location for highlight + context. */
  cell?: { row: number; col: number };
}

export interface RuleContext {
  deck: DeckModel;
  /** Deck-level aggregates computed once and shared across rules. */
  aggregates: DeckAggregates;
}

export interface DeckAggregates {
  /** Dominant client name token (e.g. "Goofy") and all observed variants. */
  clientName: { dominant: string | null; variants: Record<string, number> };
  /** The canonical confidentiality statement observed in the deck. */
  confidentialityText: string | null;
  /** Master footer placeholder rect (confidentiality anchor). */
  footerRect: Rect | null;
  /** Title placeholder rect from the content master/layout. */
  titleRectByLayout: Record<string, Rect | null>;
  /** Observed legend/key item orderings keyed by a normalized signature. */
  legendOrders: { slideIndex: number; label: string; items: string[] }[];
}

/** A deterministic rule: receives the deck + a slide, returns findings. */
export type SlideRule = (slide: Slide, ctx: RuleContext) => Finding[];

/** A deck-level rule that needs to see all slides at once. */
export type DeckRule = (ctx: RuleContext) => Finding[];

let counter = 0;
export function makeFindingId(ruleId: string, slideIndex: number): string {
  counter += 1;
  return `${ruleId}-s${slideIndex}-${counter}`;
}
