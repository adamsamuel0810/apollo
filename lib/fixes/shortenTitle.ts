import { TITLE } from "../brand/guidelines";

/** Max characters that fit in TITLE.maxLines at 24pt for the given box width. */
export function titleMaxChars(boxWidthIn: number): number {
  const charsPerLine = Math.max(
    20,
    Math.floor(((boxWidthIn - 0.2) * 72) / (TITLE.sizePt * 0.48))
  );
  return charsPerLine * TITLE.maxLines;
}

/**
 * Shorten a title to fit within ACME's 3-line limit without AI.
 * Keeps full sentences where possible, then truncates at a word boundary.
 */
export function deterministicShortenTitle(
  currentTitle: string,
  boxWidthIn: number
): string {
  const maxChars = titleMaxChars(boxWidthIn);
  const normalized = currentTitle.trim();
  if (normalized.length <= maxChars) return normalized;

  const sentences = normalized.split(/(?<=[.!?])\s+/).filter(Boolean);
  let out = "";
  for (const sentence of sentences) {
    const candidate = out ? `${out} ${sentence}` : sentence;
    if (candidate.length <= maxChars) out = candidate;
    else break;
  }
  if (out.length >= Math.min(maxChars * 0.5, normalized.length * 0.35)) {
    return out.replace(/[.,;]+$/, "").trim();
  }

  let cut = normalized.slice(0, maxChars);
  const sp = cut.lastIndexOf(" ");
  if (sp > maxChars * 0.6) cut = cut.slice(0, sp);
  return cut.replace(/[.,;]+$/, "").trim();
}
