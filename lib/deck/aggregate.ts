import { DeckModel, Rect, Shape } from "../pptx/types";
import { DeckAggregates } from "../rules/types";
import { CONFIDENTIALITY } from "../brand/guidelines";

// Words that look capitalized but are not client names.
const NAME_STOPLIST = new Set([
  "ACME",
  "THE",
  "AI",
  "M&A",
  "CEO",
  "LLC",
  "INC",
  "CORP",
  "JANUARY",
  "FEBRUARY",
  "MARCH",
  "APRIL",
  "MAY",
  "JUNE",
  "JULY",
  "AUGUST",
  "SEPTEMBER",
  "OCTOBER",
  "NOVEMBER",
  "DECEMBER",
  "DRAFT",
  "EXECUTIVE",
  "SUMMARY",
  "PEER",
  "GROUP",
  "TITLE",
  "POSITION",
  "COMPANY",
]);

export function computeAggregates(deck: DeckModel): DeckAggregates {
  return {
    clientName: detectClientName(deck),
    confidentialityText: detectConfidentialityText(deck),
    footerRect: masterFooterRect(deck),
    titleRectByLayout: titleRectByLayout(deck),
    legendOrders: detectLegendOrders(deck),
  };
}

function allTextRuns(deck: DeckModel): string[] {
  const texts: string[] = [];
  for (const s of deck.slides) {
    for (const sh of s.shapes) {
      if (sh.text) texts.push(sh.text);
      if (sh.table) {
        for (const row of sh.table.rows) for (const c of row) if (c.text) texts.push(c.text);
      }
    }
  }
  return texts;
}

function detectClientName(deck: DeckModel): DeckAggregates["clientName"] {
  const text = allTextRuns(deck).join(" ").replace(/[\u2018\u2019]/g, "'");
  // Possessive proper nouns: "Word's"
  const counts: Record<string, number> = {};
  const possRe = /\b([A-Z][a-zA-Z]{2,})'s\b/g;
  let m: RegExpExecArray | null;
  while ((m = possRe.exec(text))) {
    const tok = m[1];
    if (NAME_STOPLIST.has(tok.toUpperCase())) continue;
    counts[tok] = (counts[tok] || 0) + 1;
  }
  // Bare occurrences too (lower weight).
  const bareRe = /\b([A-Z][a-zA-Z]{2,})\b/g;
  while ((m = bareRe.exec(text))) {
    const tok = m[1];
    if (NAME_STOPLIST.has(tok.toUpperCase())) continue;
    counts[tok] = (counts[tok] || 0) + 0.25;
  }
  let dominant: string | null = null;
  let best = 0;
  for (const [tok, c] of Object.entries(counts)) {
    if (c > best) {
      best = c;
      dominant = tok;
    }
  }

  // Variants: multiword names starting with the dominant token.
  const variants: Record<string, number> = {};
  if (dominant) {
    const varRe = new RegExp(`\\b${dominant}\\s+([A-Z][a-zA-Z]+)\\b`, "g");
    while ((m = varRe.exec(text))) {
      const phrase = `${dominant} ${m[1]}`;
      variants[phrase] = (variants[phrase] || 0) + 1;
    }
  }
  return { dominant, variants };
}

function detectConfidentialityText(deck: DeckModel): string | null {
  const counts: Record<string, number> = {};
  for (const s of deck.slides) {
    for (const sh of s.shapes) {
      if (sh.placeholderType === "ftr" && sh.text.trim()) {
        const t = sh.text.trim();
        counts[t] = (counts[t] || 0) + 1;
      }
    }
  }
  let best: string | null = null;
  let bestN = 0;
  for (const [t, n] of Object.entries(counts)) {
    if (n > bestN) {
      bestN = n;
      best = t;
    }
  }
  // Fall back to the canonical statement if nothing was found.
  return best || CONFIDENTIALITY.canonical;
}

function masterFooterRect(deck: DeckModel): Rect | null {
  const ftr = deck.master.placeholders.find((p) => p.type === "ftr");
  return ftr?.rect || null;
}

function titleRectByLayout(deck: DeckModel): Record<string, Rect | null> {
  const out: Record<string, Rect | null> = {};
  for (const [name, phs] of Object.entries(deck.layoutPlaceholders)) {
    const title = phs.find((p) => p.type === "title" || p.type === "ctrTitle");
    out[name] = title?.rect || null;
  }
  return out;
}

function detectLegendOrders(deck: DeckModel): DeckAggregates["legendOrders"] {
  // Capture small text boxes whose label is "Key" or "Legend" and the nearby
  // items, so the AI can reason about cross-slide ordering consistency.
  const out: DeckAggregates["legendOrders"] = [];
  for (const s of deck.slides) {
    for (const sh of s.shapes) {
      const label = sh.text.trim().toLowerCase().replace(/:$/, "");
      if (label === "key" || label === "legend") {
        const items = nearbyShortTexts(s.shapes, sh);
        out.push({ slideIndex: s.index, label, items });
      }
    }
  }
  return out;
}

function nearbyShortTexts(shapes: Shape[], anchor: Shape): string[] {
  if (!anchor.rect) return [];
  const ay = anchor.rect.y;
  const items = shapes
    .filter(
      (s) =>
        s !== anchor &&
        s.kind === "text" &&
        s.rect &&
        Math.abs(s.rect.x - anchor.rect!.x) < 914400 * 2 &&
        s.rect.y >= ay - 914400 &&
        s.text.trim().length > 0 &&
        s.text.trim().length < 40
    )
    .sort((a, b) => (a.rect!.y || 0) - (b.rect!.y || 0))
    .map((s) => s.text.trim());
  return items.slice(0, 12);
}
