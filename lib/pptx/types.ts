// Normalized, render- and rule-friendly model of a PPTX deck.
// All geometry is stored in EMU (English Metric Units, 914400 per inch).

export interface Rect {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

export interface RunStyle {
  /** Effective font size in points (after inheritance resolution). */
  sizePt: number | null;
  bold: boolean | null;
  italic: boolean | null;
  underline: boolean | null;
  strike: boolean | null;
  /** Effective latin typeface (theme tokens resolved to real names). */
  font: string | null;
  /** Hex color (6 chars, no #) or null when theme/unknown. */
  color: string | null;
}

export interface TextRun extends RunStyle {
  text: string;
  /** True when the size/font/etc. was explicitly set on the run (vs inherited). */
  explicit: { size: boolean; bold: boolean; font: boolean; color: boolean };
}

export interface Paragraph {
  level: number; // 0-based indent level
  runs: TextRun[];
  text: string; // concatenated run text
  bulletChar: string | null; // resolved bullet glyph, or null if buNone
  bulletType: "char" | "none" | "auto" | null;
  /** Typeface declared for the bullet glyph (e.g. Wingdings), if any. */
  bulletFont: string | null;
  align: string | null; // l/ctr/r/just
  /** Left margin (marL) in inches. */
  marginLeftIn: number | null;
  /** First-line / hanging indent in inches (negative = hanging bullet). */
  firstLineIndentIn: number | null;
  /** @deprecated use marginLeftIn */
  indentIn: number | null;
  spaceBeforePt: number | null;
  spaceAfterPt: number | null;
  lineSpacingPct: number | null; // 100 = single
}

export type ShapeKind = "text" | "table" | "picture" | "chart" | "group" | "other";

/** Text-frame properties from <a:bodyPr> (vertical anchor + internal insets). */
export interface BodyProps {
  /** Vertical anchor: "t" (top), "ctr" (middle), "b" (bottom). */
  anchor: "t" | "ctr" | "b";
  /** Internal text insets in EMU. */
  lIns: number;
  tIns: number;
  rIns: number;
  bIns: number;
  /** Word wrap: "square" (wrap) or "none". */
  wrap: "square" | "none";
}

export interface TableCell {
  text: string;
  runs: TextRun[];
  rowSpan: number;
  gridSpan: number;
  /** Cell is covered by a merge (horizontal or vertical continuation). */
  merged: boolean;
  hMerge: boolean;
  vMerge: boolean;
  fill: string | null; // hex or null
  align: string | null;
  borders: {
    left: boolean;
    right: boolean;
    top: boolean;
    bottom: boolean;
  };
}

export interface TableModel {
  rows: TableCell[][];
  colWidthsEmu: number[];
  rowHeightsEmu: number[];
}

export interface Shape {
  id: string;
  name: string;
  kind: ShapeKind;
  placeholderType: string | null; // title, ctrTitle, body, ftr, sldNum, ...
  placeholderIdx: string | null;
  rect: Rect | null;
  paragraphs: Paragraph[];
  text: string; // all text in shape
  /** Text-frame anchor/insets, used for faithful vertical alignment + padding. */
  bodyPr: BodyProps | null;
  table: TableModel | null;
  /** For pictures: a data URL the client can render. */
  imageDataUrl: string | null;
  /** Shape fill color (hex) if a solid fill is set. */
  fill: string | null;
  /** Children for groups (flattened shapes also appear top-level for rules). */
  children?: Shape[];
}

export type SlideKind = "title" | "section" | "content";

export interface Slide {
  index: number; // 1-based slide number
  kind: SlideKind;
  layoutName: string | null;
  layoutType: string | null;
  size: { cx: number; cy: number };
  shapes: Shape[];
  /** The title shape (ctrTitle/title) if present. */
  titleShape: Shape | null;
  /** Speaker notes text (used only by the eval harness; never shown to the AI). */
  notes: string;
}

export interface MasterPlaceholder {
  type: string | null;
  idx: string | null;
  rect: Rect | null;
}

export interface DeckModel {
  size: { cx: number; cy: number };
  slides: Slide[];
  master: {
    placeholders: MasterPlaceholder[];
  };
  /** Map of layout name -> placeholders (for "moved from default" checks). */
  layoutPlaceholders: Record<string, MasterPlaceholder[]>;
  theme: {
    majorFont: string;
    minorFont: string;
  };
}
