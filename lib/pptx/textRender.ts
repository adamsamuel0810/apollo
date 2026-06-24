import { EMU_PER_INCH } from "../brand/guidelines";
import { Paragraph, BodyProps } from "./types";

/** CSS font-family stack used when drawing slide text in the browser. */
export const SLIDE_FONT_STACK =
  'var(--slide-font-family, Carlito), Calibri, "Segoe UI", system-ui, sans-serif';

/** Map OOXML typeface names to the slide font stack (Calibri → Carlito on web). */
export function cssFontFamily(typeface: string | null | undefined): string {
  if (!typeface) return SLIDE_FONT_STACK;
  const f = typeface.trim();
  if (/^calibri/i.test(f) || f.startsWith("+mj") || f.startsWith("+mn")) {
    return SLIDE_FONT_STACK;
  }
  return `"${f}", ${SLIDE_FONT_STACK}`;
}

/** Convert points to CSS pixels given the slide scale factor. */
export function ptToPx(sizePt: number, pxPerEmu: number): number {
  return (sizePt / 72) * EMU_PER_INCH * pxPerEmu;
}

/** Convert inches to CSS pixels given the slide scale factor. */
export function inToPx(inches: number, pxPerEmu: number): number {
  return inches * EMU_PER_INCH * pxPerEmu;
}

/** Convert EMU directly to CSS pixels. */
export function emuToPx(emu: number, pxPerEmu: number): number {
  return emu * pxPerEmu;
}

export interface BodyFrameStyle {
  justifyContent: "flex-start" | "center" | "flex-end";
  paddingTop: number;
  paddingRight: number;
  paddingBottom: number;
  paddingLeft: number;
}

/** Map a shape's text-frame props to flex alignment + padding (in px). */
export function bodyFrameStyle(
  body: BodyProps | null,
  pxPerEmu: number
): BodyFrameStyle {
  const anchor = body?.anchor ?? "t";
  return {
    justifyContent:
      anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start",
    paddingTop: emuToPx(body?.tIns ?? 45720, pxPerEmu),
    paddingRight: emuToPx(body?.rIns ?? 91440, pxPerEmu),
    paddingBottom: emuToPx(body?.bIns ?? 45720, pxPerEmu),
    paddingLeft: emuToPx(body?.lIns ?? 91440, pxPerEmu),
  };
}

export interface BulletDisplay {
  char: string;
  fontFamily: string;
}

const SYMBOL_FONT_STACK =
  '"Segoe UI Symbol", "Arial Unicode MS", "Noto Sans Symbols", system-ui, sans-serif';

// Wingdings code point (as the literal char stored in buChar) -> Unicode glyph.
// Keys are the characters PowerPoint writes into buChar for Wingdings bullets.
const WINGDINGS_MAP: Record<string, string> = {
  "\u00A7": "\u25AA", // § -> ▪ small black square (L1)
  "\u00D8": "\u2794", // Ø -> ➔ heavy wide-headed rightwards arrow (L3)
  n: "\u25AA", // n -> ▪ filled square
  l: "\u25CF", // l -> ● black circle
  u: "\u25C6", // u -> ◆ black diamond
  "\u00FC": "\u2713", // ü -> ✓ check
  "\u00D6": "\u2666", // Ö -> ♦
};

// Symbol-font (Adobe Symbol) bullets occasionally used.
const SYMBOL_MAP: Record<string, string> = {
  "\u00B7": "\u2022", // middle dot -> bullet
  "\u00A8": "\u2022",
};

function isWingdings(font: string | null): boolean {
  return !!font && /wingdings/i.test(font);
}
function isSymbolFont(font: string | null): boolean {
  return !!font && /^symbol$/i.test(font);
}

/**
 * Resolve the actual bullet glyph + font for the browser, honoring the bullet
 * font declared in the PPTX (e.g. Wingdings) so symbols don't break.
 */
export function bulletDisplay(para: Paragraph): BulletDisplay | null {
  if (para.bulletType === "none" || para.bulletType == null) return null;

  if (para.bulletType === "auto") {
    return { char: "\u2022", fontFamily: SLIDE_FONT_STACK };
  }

  const raw = para.bulletChar || "";
  if (!raw) return { char: "\u2022", fontFamily: SLIDE_FONT_STACK };

  // Glyph encoded in a symbol font: translate the code point to real Unicode.
  if (isWingdings(para.bulletFont)) {
    const mapped = WINGDINGS_MAP[raw];
    if (mapped) return { char: mapped, fontFamily: SYMBOL_FONT_STACK };
    // Unknown Wingdings code: fall back to a square so it never shows mojibake.
    return { char: "\u25AA", fontFamily: SYMBOL_FONT_STACK };
  }
  if (isSymbolFont(para.bulletFont)) {
    const mapped = SYMBOL_MAP[raw];
    return { char: mapped || "\u2022", fontFamily: SYMBOL_FONT_STACK };
  }

  // Text-font bullet (buFontTx or a normal font): the char is literal Unicode
  // already (e.g. "•", "–"). Render it in a symbol-safe stack for odd glyphs.
  const needsSymbolFont = raw.charCodeAt(0) > 0x2000;
  return {
    char: raw,
    fontFamily: needsSymbolFont ? SYMBOL_FONT_STACK : SLIDE_FONT_STACK,
  };
}

export interface ParagraphLayoutStyle {
  marginTopPx: number;
  marginBottomPx: number;
  lineHeight: number;
  marginLeftPx: number;
  textIndentPx: number;
  textAlign: "left" | "center" | "right" | "justify";
}

export function paragraphLayoutStyle(
  para: Paragraph,
  pxPerEmu: number
): ParagraphLayoutStyle {
  const align =
    para.align === "ctr"
      ? "center"
      : para.align === "r"
        ? "right"
        : para.align === "just"
          ? "justify"
          : "left";

  const marginLeftPx = para.marginLeftIn != null ? inToPx(para.marginLeftIn, pxPerEmu) : 0;
  const textIndentPx =
    para.firstLineIndentIn != null ? inToPx(para.firstLineIndentIn, pxPerEmu) : 0;

  const lineHeight =
    para.lineSpacingPct != null && para.lineSpacingPct > 0
      ? para.lineSpacingPct / 100
      : 1.15;

  return {
    marginTopPx:
      para.spaceBeforePt != null ? ptToPx(para.spaceBeforePt, pxPerEmu) : 0,
    marginBottomPx:
      para.spaceAfterPt != null ? ptToPx(para.spaceAfterPt, pxPerEmu) : 0,
    lineHeight,
    marginLeftPx,
    textIndentPx,
    textAlign: align,
  };
}
