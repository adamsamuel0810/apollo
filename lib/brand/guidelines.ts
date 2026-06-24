// Numeric/string constants distilled from ACME-brand-guidelines.md.
// Centralized so rules stay declarative and easy to audit.

export const EMU_PER_INCH = 914400;
export const EMU_PER_POINT = 12700;

export const TYPEFACE = "Calibri";

// Allowable point-size tolerance when comparing resolved font sizes.
export const SIZE_TOLERANCE_PT = 0.5;

// Geometry tolerance (inches) for position/centering comparisons.
export const POS_TOLERANCE_IN = 0.05;

export const TITLE = {
  sizePt: 24,
  maxLines: 3,
};

export const BULLETS = {
  primarySizePt: 16,
  subSizePt: 16,
  // Indentation by level (inches). Level 1 has no extra indent baseline.
  indentByLevel: { 2: 0.4, 3: 0.6, 4: 0.6, 5: 0.6 } as Record<number, number>,
  // Spacing-before between bullets (points).
  spacingByLevel: { 1: 6, 2: 0, 3: 0, 4: 0, 5: 0 } as Record<number, number>,
};

export const TABLE = {
  minSizePt: 10,
  maxSizePt: 12,
  titleSizePt: 14,
  firstColLeftMarginIn: 0.03, // "First column should be .03 in from edge"
};

export const FOOTNOTE = {
  sizePt: 10,
  chartSourceSizePt: 9,
};

export const CONFIDENTIALITY = {
  // Canonical legal text; must appear (unmodified) on every slide.
  canonical:
    "These materials are confidential and to be used only by the client to whom they are addressed",
  sizePt: 8,
};

// Content must not be closer to the left edge than 1" below the header.
export const CONTENT_MIN_LEFT_IN = 1.0;

// Square bullet glyphs treated as equivalent for Level 1.
export const L1_BULLET_GLYPHS = ["\u220E" /* ∎ */, "\u25AA" /* ▪ */];
export const L2_BULLET_GLYPHS = ["-", "\u2013" /* – */, "--"];
export const L3_BULLET_GLYPH = "\u279E"; // ➞ three-D top-lighted rightwards arrowhead (Alt-10146)

// Terminology rules.
export const TERMINOLOGY = {
  forbiddenAbbreviations: [{ bad: "TGT", good: "Target" }],
  // Employee-table header expectations.
  incumbentBad: "Incumbent",
  incumbentGood: "Executive",
  percentileBad: "Percentile",
  percentileGood: "%ile",
};
