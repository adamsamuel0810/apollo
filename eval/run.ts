import { readFileSync, writeFileSync, mkdirSync, existsSync } from "fs";
import path from "path";
import { analyzePptx } from "../lib/analyze/analyze";
import { isPrimary } from "../lib/findings/merge";
import { Finding } from "../lib/rules/types";
import { GROUND_TRUTH, RULE_TO_TAG, AI_CATEGORY_TO_TAGS, Layer } from "./groundTruth";

const ANNOTATED = "Goofy Corp Peer Group Assessment - ANNOTATED.pptx";
const PLUTOS = "Plutos Assessment of Compensation Peers.pptx";

function tagsForFinding(f: Finding): string[] {
  if (f.source === "deterministic") {
    const t = RULE_TO_TAG[f.ruleId];
    return t ? [t] : [];
  }
  return AI_CATEGORY_TO_TAGS[f.category] || [];
}

async function evalAnnotated() {
  const buf = readFileSync(ANNOTATED);
  const result = await analyzePptx(buf, ANNOTATED);

  console.log("=".repeat(70));
  console.log(`EVAL: ${ANNOTATED}`);
  console.log(`AI enabled: ${result.aiEnabled}`);
  console.log("=".repeat(70));

  const byLayer: Record<Layer, { expected: number; covered: number; missed: string[] }> = {
    deterministic: { expected: 0, covered: 0, missed: [] },
    ai: { expected: 0, covered: 0, missed: [] },
    hard: { expected: 0, covered: 0, missed: [] },
  };

  let primaryTotal = 0;
  let primaryMatched = 0;
  const primaryFPs: string[] = [];

  for (const slide of result.slides) {
    const gt = GROUND_TRUTH[slide.index] || [];
    const foundTags = new Set<string>();
    for (const f of slide.findings) for (const t of tagsForFinding(f)) foundTags.add(t);

    for (const { tag, layer } of gt) {
      byLayer[layer].expected++;
      if (foundTags.has(tag)) byLayer[layer].covered++;
      else byLayer[layer].missed.push(`s${slide.index}:${tag}`);
    }

    // Precision on the PRIMARY tier (what the user sees by default).
    const expectedTags = new Set(gt.map((g) => g.tag));
    for (const f of slide.findings) {
      if (!isPrimary(f)) continue;
      primaryTotal++;
      const tags = tagsForFinding(f);
      const matched = tags.some((t) => expectedTags.has(t));
      if (matched) primaryMatched++;
      else primaryFPs.push(`s${slide.index}:${f.ruleId} (${f.title})`);
    }
  }

  const recall = (l: Layer) =>
    byLayer[l].expected === 0 ? 1 : byLayer[l].covered / byLayer[l].expected;

  console.log("\nRECALL by layer (covered / expected):");
  (["deterministic", "ai", "hard"] as Layer[]).forEach((l) => {
    console.log(
      `  ${l.padEnd(14)} ${byLayer[l].covered}/${byLayer[l].expected}  (${(recall(l) * 100).toFixed(0)}%)`
    );
    if (byLayer[l].missed.length) console.log(`      missed: ${byLayer[l].missed.join(", ")}`);
  });

  const precision = primaryTotal === 0 ? 1 : primaryMatched / primaryTotal;
  console.log(`\nPRIMARY-tier precision: ${primaryMatched}/${primaryTotal} (${(precision * 100).toFixed(0)}%)`);
  if (primaryFPs.length) {
    console.log("  potential false positives in primary tier:");
    for (const fp of primaryFPs) console.log(`    - ${fp}`);
  } else {
    console.log("  no false positives in the primary tier.");
  }

  console.log(`\nTotals: ${result.summary.total} findings ` +
    `(E:${result.summary.bySeverity.error} W:${result.summary.bySeverity.warning} I:${result.summary.bySeverity.info}) ` +
    `across ${result.summary.slidesWithFindings} slides`);

  // Persist a sample result for the UI preview / debugging.
  const outDir = path.join("eval", "output");
  if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });
  writeFileSync(path.join(outDir, "annotated-result.json"), JSON.stringify(result, null, 2));
  console.log(`\nWrote eval/output/annotated-result.json`);
}

async function smokePlutos() {
  if (!existsSync(PLUTOS)) return;
  const buf = readFileSync(PLUTOS);
  const result = await analyzePptx(buf, PLUTOS);
  console.log("\n" + "=".repeat(70));
  console.log(`SMOKE: ${PLUTOS}`);
  console.log("=".repeat(70));
  console.log(
    `Slides: ${result.slideCount}, findings: ${result.summary.total} ` +
      `(E:${result.summary.bySeverity.error} W:${result.summary.bySeverity.warning} I:${result.summary.bySeverity.info})`
  );
  const primaryPerSlide = result.slides
    .map((s) => ({ i: s.index, n: s.findings.filter(isPrimary).length }))
    .filter((x) => x.n > 0);
  console.log("Primary findings per slide:", primaryPerSlide.map((x) => `s${x.i}:${x.n}`).join(" "));
  writeFileSync(path.join("eval", "output", "plutos-result.json"), JSON.stringify(result, null, 2));
  console.log("Wrote eval/output/plutos-result.json");
}

async function main() {
  await evalAnnotated();
  await smokePlutos();
}
main().catch((e) => {
  console.error(e);
  process.exit(1);
});
