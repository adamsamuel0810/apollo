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
  /** Machine-applicable fix metadata; present when Accept can auto-apply. */
  fix?: FixInstruction;
}

/** Describes how to apply an automatic fix to the underlying PPTX. */
export type FixKind =
  | "text-replace"
  | "title-rewrite"
  | "ai-rewrite"
  | "deck-replace"
  | "font-family-set"
  | "font-size-set"
  | "font-bold-set"
  | "color-set"
  | "shape-reposition"
  | "row-fill-set"
  | "border-remove"
  | "ensure-footer";

export interface FixInstruction {
  kind: FixKind;
  slideIndex: number;
  shapeId?: string;
  cell?: { row: number; col: number };
  /** For text-replace fixes. */
  search?: string;
  replace?: string;
  caseInsensitive?: boolean;
  /** When true, `search` is a regex pattern (not escaped). */
  regex?: boolean;
  /** When true, apply text-replace across every slide in the deck. */
  deckWide?: boolean;
  /** For title-rewrite / ai-rewrite fixes. */
  newText?: string;
  /** Extra context for AI-driven rewrites. */
  aiContext?: string;
  /** font-family-set / font-size-set */
  fontFamily?: string;
  fontSizePt?: number;
  /** color-set */
  fromColor?: string;
  toColor?: string;
  colorTarget?: "fill" | "text" | "both";
  /** shape-reposition (EMU) */
  targetRect?: import("../pptx/types").Rect;
  /** row-fill-set */
  rowIndex?: number;
  fillColor?: string;
  /** border-remove */
  borderScope?: "table-sides";
}

export interface FixRecord {
  findingId: string;
  ruleId: string;
  slideIndex: number;
  title: string;
  before: string;
  after: string;
  appliedAt: string;
  source: FindingSource;
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
