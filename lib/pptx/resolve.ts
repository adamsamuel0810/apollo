import { XmlNode, child, children, attr, nodeName } from "./xml";
import { RunStyle } from "./types";

export interface ThemeFonts {
  majorFont: string;
  minorFont: string;
}

type PartialStyle = Partial<RunStyle>;

/** Extract run properties from an rPr / defRPr / endParaRPr node. */
export function extractProps(rPr: XmlNode | undefined, theme: ThemeFonts): PartialStyle {
  if (!rPr) return {};
  const out: PartialStyle = {};

  const sz = attr(rPr, "sz");
  if (sz != null) out.sizePt = parseInt(sz, 10) / 100;

  const b = attr(rPr, "b");
  if (b != null) out.bold = b === "1" || b === "true";

  const i = attr(rPr, "i");
  if (i != null) out.italic = i === "1" || i === "true";

  const u = attr(rPr, "u");
  if (u != null) out.underline = u !== "none";

  const strike = attr(rPr, "strike");
  if (strike != null) out.strike = strike !== "noStrike";

  const latin = child(rPr, "a:latin");
  if (latin) {
    const tf = attr(latin, "typeface");
    if (tf != null) out.font = resolveThemeFont(tf, theme);
  }

  const fill = child(rPr, "a:solidFill");
  if (fill) {
    const srgb = child(fill, "a:srgbClr");
    if (srgb) {
      const val = attr(srgb, "val");
      if (val) out.color = val.toUpperCase();
    }
    // schemeClr left unresolved (null) to avoid false palette flags.
  }

  return out;
}

export function resolveThemeFont(typeface: string, theme: ThemeFonts): string {
  if (typeface.startsWith("+mj")) return theme.majorFont;
  if (typeface.startsWith("+mn")) return theme.minorFont;
  return typeface;
}

/** Find the defRPr for a given 0-based level inside an lstStyle node. */
export function levelDefRPr(lstStyle: XmlNode | undefined, level: number): XmlNode | undefined {
  if (!lstStyle) return undefined;
  const tag = level === 0 ? "a:lvl1pPr" : `a:lvl${level + 1}pPr`;
  let pPr = child(lstStyle, tag);
  // defPPr is used as a generic fallback inside otherStyle.
  if (!pPr && level === 0) pPr = child(lstStyle, "a:defPPr");
  if (!pPr) return undefined;
  return child(pPr, "a:defRPr");
}

/**
 * Resolve effective run style by cascading from least specific (end of chain)
 * to most specific (run rPr). `styleChain` is ordered most-specific first.
 */
export function resolveRunStyle(
  runRPr: XmlNode | undefined,
  styleChain: (XmlNode | undefined)[],
  level: number,
  theme: ThemeFonts
): { style: RunStyle; explicit: { size: boolean; bold: boolean; font: boolean; color: boolean } } {
  const merged: RunStyle = {
    sizePt: null,
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    font: null,
    color: null,
  };

  const apply = (p: PartialStyle) => {
    for (const k of Object.keys(p) as (keyof RunStyle)[]) {
      const v = p[k];
      if (v !== undefined && v !== null && merged[k] === null) {
        // @ts-expect-error index assignment across union
        merged[k] = v;
      }
    }
  };

  // Run-level explicit props first (highest precedence).
  const runProps = extractProps(runRPr, theme);
  const explicit = {
    size: runProps.sizePt != null,
    bold: runProps.bold != null,
    font: runProps.font != null,
    color: runProps.color != null,
  };
  apply(runProps);

  // Then each lstStyle source, most specific first.
  for (const ls of styleChain) {
    apply(extractProps(levelDefRPr(ls, level), theme));
  }

  return { style: merged, explicit };
}

/** Pull the lstStyle node out of a txBody / placeholder shape if present. */
export function lstStyleOf(node: XmlNode | undefined): XmlNode | undefined {
  if (!node) return undefined;
  return child(node, "a:lstStyle");
}

/** Read theme major/minor latin fonts from a theme document root. */
export function readThemeFonts(themeRoot: XmlNode | undefined): ThemeFonts {
  const fallback = { majorFont: "Calibri", minorFont: "Calibri" };
  if (!themeRoot) return fallback;
  const themeEls = child(themeRoot, "a:themeElements");
  const fontScheme = themeEls && child(themeEls, "a:fontScheme");
  if (!fontScheme) return fallback;
  const major = child(fontScheme, "a:majorFont");
  const minor = child(fontScheme, "a:minorFont");
  const majLatin = major && child(major, "a:latin");
  const minLatin = minor && child(minor, "a:latin");
  return {
    majorFont: (majLatin && attr(majLatin, "typeface")) || fallback.majorFont,
    minorFont: (minLatin && attr(minLatin, "typeface")) || fallback.minorFont,
  };
}

/** Map a placeholder type to a master txStyles category. */
export function masterCategoryFor(phType: string | null): "title" | "body" | "other" {
  if (phType === "title" || phType === "ctrTitle") return "title";
  if (phType === "body" || phType === "subTitle" || phType == null) return "body";
  return "other";
}

export { nodeName, children };
