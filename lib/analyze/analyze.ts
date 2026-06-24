import { parseDeck } from "../pptx/parser";
import { runDeterministicRules } from "../rules/engine";
import { runAiChecks } from "../ai/orchestrator";
import { mergeFindings } from "../findings/merge";
import { aiEnabled } from "../ai/client";
import { AnalyzeResult, AnalyzeSlide, AnalyzeSummary } from "./types";
import { Finding, Severity } from "../rules/types";

export async function analyzePptx(
  data: Buffer | ArrayBuffer | Uint8Array,
  fileName: string
): Promise<AnalyzeResult> {
  const deck = await parseDeck(data);

  const { findings: deterministic, ctx } = runDeterministicRules(deck);

  let aiFindings: Finding[] = [];
  try {
    aiFindings = await runAiChecks(deck, ctx.aggregates);
  } catch (err) {
    console.error("AI checks failed; continuing with deterministic only:", err);
  }

  const merged = mergeFindings([...deterministic, ...aiFindings]);

  const slides: AnalyzeSlide[] = deck.slides.map((s) => ({
    index: s.index,
    kind: s.kind,
    size: s.size,
    shapes: s.shapes,
    findings: merged.filter((f) => f.slideIndex === s.index),
  }));

  const summary = summarize(merged);

  return {
    fileName,
    slideCount: deck.slides.length,
    generatedAt: new Date().toISOString(),
    aiEnabled: aiEnabled(),
    summary,
    slides,
  };
}

function summarize(findings: Finding[]): AnalyzeSummary {
  const bySeverity: Record<Severity, number> = { error: 0, warning: 0, info: 0 };
  const slidesWith = new Set<number>();
  for (const f of findings) {
    bySeverity[f.severity]++;
    slidesWith.add(f.slideIndex);
  }
  return {
    total: findings.length,
    bySeverity,
    slidesWithFindings: slidesWith.size,
  };
}
