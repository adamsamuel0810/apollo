import { Finding, Severity } from "../rules/types";

const SEVERITY_RANK: Record<Severity, number> = { error: 0, warning: 1, info: 2 };

/**
 * Collapse duplicate findings and merge AI + deterministic results.
 *
 * - Deterministic findings of the same (slide, ruleId) are collapsed into one
 *   (keeping the highest confidence) to avoid repeating the same issue.
 * - When deterministic and AI flag the same (slide, category) at overlapping
 *   locations, the deterministic finding wins (it is precise) and the AI one is
 *   dropped to control noise.
 */
export function mergeFindings(findings: Finding[]): Finding[] {
  // 1. Collapse deterministic duplicates per (slide, ruleId).
  const collapsed: Finding[] = [];
  const byRule = new Map<string, Finding>();
  for (const f of findings) {
    if (f.source !== "deterministic") {
      collapsed.push(f);
      continue;
    }
    const key = `${f.slideIndex}:${f.ruleId}`;
    const existing = byRule.get(key);
    if (!existing) {
      byRule.set(key, { ...f });
    } else {
      if (f.confidence > existing.confidence) {
        byRule.set(key, { ...f, detail: existing.detail });
      }
      // Aggregate distinct evidence snippets into the detail.
      if (f.evidence && existing.evidence && !existing.evidence.includes(f.evidence)) {
        existing.evidence = `${existing.evidence}; ${f.evidence}`;
      }
    }
  }
  collapsed.push(...byRule.values());

  // 2. Drop AI findings that duplicate a deterministic finding on the same slide
  //    and category (deterministic is authoritative).
  const detKeys = new Set(
    collapsed
      .filter((f) => f.source === "deterministic")
      .map((f) => `${f.slideIndex}:${f.category}`)
  );
  const deduped = collapsed.filter((f) => {
    if (f.source !== "ai") return true;
    return !detKeys.has(`${f.slideIndex}:${f.category}`);
  });

  // 3. Sort: by slide, then severity, then confidence desc.
  deduped.sort((a, b) => {
    if (a.slideIndex !== b.slideIndex) return a.slideIndex - b.slideIndex;
    if (SEVERITY_RANK[a.severity] !== SEVERITY_RANK[b.severity])
      return SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity];
    return b.confidence - a.confidence;
  });

  return deduped;
}

/** Default visibility threshold for the "primary" review view. */
export const DEFAULT_CONFIDENCE_THRESHOLD = 0.6;

export function isPrimary(f: Finding, threshold = DEFAULT_CONFIDENCE_THRESHOLD): boolean {
  return f.severity !== "info" && f.confidence >= threshold;
}
