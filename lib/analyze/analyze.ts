import { parseDeck } from "../pptx/parser";
import { runDeterministicRules } from "../rules/engine";
import { runAiChecks } from "../ai/orchestrator";
import { mergeFindings } from "../findings/merge";
import { aiEnabled } from "../ai/client";
import { convertPptxToImages } from "../pptx/convert";
import { AnalyzeResult, AnalyzeSlide, AnalyzeSummary } from "./types";
import { Finding, Severity } from "../rules/types";

export async function analyzePptx(
  data: Buffer | ArrayBuffer | Uint8Array,
  fileName: string,
  options: { skipImages?: boolean; skipAi?: boolean } = {}
): Promise<AnalyzeResult> {
  const deck = await parseDeck(data);

  const { findings: deterministic, ctx } = runDeterministicRules(deck);

  const [aiFindings, slideImages] = await Promise.all([
    options.skipAi
      ? Promise.resolve([] as Finding[])
      : runAiChecks(deck, ctx.aggregates).catch((err) => {
          console.error("AI checks failed; continuing with deterministic only:", err);
          return [] as Finding[];
        }),
    options.skipImages
      ? Promise.resolve(null)
      : convertPptxToImages(data, fileName).catch((err) => {
          console.error("Slide image rendering failed; falling back to HTML:", err);
          return null;
        }),
  ]);

  const merged = mergeFindings([...deterministic, ...aiFindings]);

  const imageByIndex = new Map<number, string>();
  for (const img of slideImages ?? []) imageByIndex.set(img.index, img.dataUrl);

  const slides: AnalyzeSlide[] = deck.slides.map((s) => ({
    index: s.index,
    kind: s.kind,
    size: s.size,
    shapes: s.shapes,
    findings: merged.filter((f) => f.slideIndex === s.index),
    imageUrl: imageByIndex.get(s.index) ?? null,
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
