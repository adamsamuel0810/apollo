// ACME Color Palette (from ACME-brand-guidelines.md).
// Used to validate that every fill/text color in a deck is "on brand".

export const ACME_PALETTE: string[] = [
  // Neutrals
  "FFFFFF", "000000", "F2F2F2", "808080", "D9D9D9", "595959", "BFBFBF",
  "404040", "A6A6A6", "262626", "0D0D0D",
  // Teals / cyans
  "75FDF7", "E1FFFE", "C7FEFB", "ABFEFA", "18FCF0", "03B6AD",
  // Blues
  "006EBE", "BDE3FF", "7DC8FF", "3EADFF", "00518C", "00365E",
  "7DD7FF", "E3F7FF", "CAEFFF", "B0E8FF", "1EBBFF", "0085BD",
  // Greens
  "B9FF6E", "F1FFE1", "E2FFC4", "D6FFA8", "8FFF13", "60B700",
  // Yellows
  "FFFF50", "FFFFDB", "FFFFB9", "FFFF95", "FBFB00", "A8A800",
  // Reds
  "FF3737", "FFD7D7", "FFAEAE", "FF8686", "E80000", "9B0000",
  // Oranges
  "FFCD69", "FFF4DF", "FFEBC1", "FFE1A4", "FFB00D", "B37800",
  // Purples
  "972FFF", "EAD5FF", "D5AAFF", "C082FF", "7100E1", "4B0097",
].map((c) => c.toUpperCase());

const PALETTE_SET = new Set(ACME_PALETTE);

/** Normalize a hex color string to 6 uppercase hex chars (no leading #). */
export function normalizeHex(hex: string | null | undefined): string | null {
  if (!hex) return null;
  let h = hex.trim().replace(/^#/, "").toUpperCase();
  if (h.length === 8) h = h.slice(0, 6); // strip alpha
  if (!/^[0-9A-F]{6}$/.test(h)) return null;
  return h;
}

/** True if the color is part of the ACME palette (with a small tolerance). */
export function isOnPalette(hex: string | null | undefined, tolerance = 0): boolean {
  const h = normalizeHex(hex);
  if (!h) return true; // unknown / theme-driven colors are not flagged here
  if (PALETTE_SET.has(h)) return true;
  if (tolerance <= 0) return false;
  return ACME_PALETTE.some((p) => colorDistance(h, p) <= tolerance);
}

/** Nearest palette color (for "did you mean" suggestions). */
export function nearestPaletteColor(hex: string): { color: string; distance: number } | null {
  const h = normalizeHex(hex);
  if (!h) return null;
  let best: { color: string; distance: number } | null = null;
  for (const p of ACME_PALETTE) {
    const d = colorDistance(h, p);
    if (!best || d < best.distance) best = { color: p, distance: d };
  }
  return best;
}

function colorDistance(a: string, b: string): number {
  const ar = parseInt(a.slice(0, 2), 16);
  const ag = parseInt(a.slice(2, 4), 16);
  const ab = parseInt(a.slice(4, 6), 16);
  const br = parseInt(b.slice(0, 2), 16);
  const bg = parseInt(b.slice(2, 4), 16);
  const bb = parseInt(b.slice(4, 6), 16);
  return Math.sqrt((ar - br) ** 2 + (ag - bg) ** 2 + (ab - bb) ** 2);
}

// Semantic colors referenced by table rules.
export const TABLE_COLORS = {
  headerFill: "006EBE",
  headerFont: "FFFFFF",
  zebra: "F2F2F2",
  statsRow: "FFFFDB", // Yellow, Accent 3
  clientRow: "F1FFE1", // Light Green, Accent 2
} as const;
