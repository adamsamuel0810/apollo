import { Shape, Slide, Paragraph, TextRun, Rect } from "../pptx/types";
import { EMU_PER_INCH, SIZE_TOLERANCE_PT, TYPEFACE } from "../brand/guidelines";

export function emuToIn(emu: number): number {
  return emu / EMU_PER_INCH;
}
export function inToEmu(inch: number): number {
  return Math.round(inch * EMU_PER_INCH);
}

export function ptApproxEqual(a: number | null, b: number, tol = SIZE_TOLERANCE_PT): boolean {
  if (a == null) return false;
  return Math.abs(a - b) <= tol;
}

export function isCalibri(font: string | null): boolean {
  if (!font) return true; // unknown -> inherited Calibri, don't flag
  const f = font.toLowerCase();
  return f === TYPEFACE.toLowerCase() || f.startsWith("calibri");
}

/** Runs with non-whitespace text. */
export function visibleRuns(shape: Shape): TextRun[] {
  const out: TextRun[] = [];
  for (const p of shape.paragraphs) {
    for (const r of p.runs) if (r.text.trim()) out.push(r);
  }
  return out;
}

export function visibleParagraphs(shape: Shape): Paragraph[] {
  return shape.paragraphs.filter((p) => p.text.trim().length > 0);
}

export function shapeText(shape: Shape): string {
  return shape.text.replace(/\s+/g, " ").trim();
}

const TRAILING_PUNCT = /[.,;]$/;
export function endsWithPunctuation(text: string): boolean {
  return TRAILING_PUNCT.test(text.trim());
}

/** Estimate how many visual lines a string occupies in a box of given width. */
export function estimateLines(
  text: string,
  fontSizePt: number,
  boxWidthIn: number,
  explicitBreaks = 0
): number {
  const effectiveWidthPt = Math.max(0.5, boxWidthIn - 0.2) * 72;
  const avgCharWidthPt = fontSizePt * 0.48; // conservative for Calibri
  const charsPerLine = Math.max(1, Math.floor(effectiveWidthPt / avgCharWidthPt));
  const segments = text.split("\n");
  let lines = 0;
  for (const seg of segments) {
    const len = seg.trim().length;
    lines += Math.max(1, Math.ceil(len / charsPerLine));
  }
  return lines + explicitBreaks;
}

export function countExplicitBreaks(shape: Shape): number {
  let n = 0;
  for (const p of shape.paragraphs) {
    for (const r of p.runs) if (r.text === "\n") n++;
    n += Math.max(0, p.runs.filter((r) => r.text === "\n").length);
  }
  // paragraphs beyond the first also add lines
  return Math.max(0, shape.paragraphs.filter((p) => p.text.trim()).length - 1);
}

export function rectsIntersect(a: Rect, b: Rect, slackEmu = 0): boolean {
  return !(
    a.x + a.cx <= b.x + slackEmu ||
    b.x + b.cx <= a.x + slackEmu ||
    a.y + a.cy <= b.y + slackEmu ||
    b.y + b.cy <= a.y + slackEmu
  );
}

export function intersectionArea(a: Rect, b: Rect): number {
  const x = Math.max(0, Math.min(a.x + a.cx, b.x + b.cx) - Math.max(a.x, b.x));
  const y = Math.max(0, Math.min(a.y + a.cy, b.y + b.cy) - Math.max(a.y, b.y));
  return x * y;
}

export function centerX(r: Rect): number {
  return r.x + r.cx / 2;
}

export function normalizeText(t: string): string {
  return t
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/\s+/g, " ")
    .trim();
}

/** The title text of a slide (from the title placeholder). */
export function titleText(slide: Slide): string | null {
  if (!slide.titleShape) return null;
  return shapeText(slide.titleShape) || null;
}

/** All content shapes excluding footer/slide-number/title placeholders. */
export function bodyShapes(slide: Slide): Shape[] {
  return slide.shapes.filter(
    (s) =>
      s.placeholderType !== "ftr" &&
      s.placeholderType !== "sldNum" &&
      s.placeholderType !== "dt" &&
      s.placeholderType !== "title" &&
      s.placeholderType !== "ctrTitle"
  );
}
