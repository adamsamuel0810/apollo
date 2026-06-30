import { readFileSync } from "fs";
import { parseDeck } from "../lib/pptx/parser";
import { runDeterministicRules } from "../lib/rules/engine";
import { applyFixToPptx } from "../lib/pptx/writer";
import { prepareFindingForApply } from "../lib/fixes/prepareFinding";

const path = process.argv[2] || "_out-fixed.pptx";
const buf = readFileSync(path);

async function main() {
  const deck = await parseDeck(buf);
  const { findings, ctx } = runDeterministicRules(deck);
  const missing = findings.filter((f) => f.ruleId === "confidentiality-missing");
  console.log("File:", path);
  console.log("Missing:", missing.length, missing.map((f) => f.slideIndex));

  for (const f of missing.slice(0, 5)) {
    const slide = deck.slides.find((s) => s.index === f.slideIndex)!;
    const ftrs = slide.shapes.filter((s) => s.placeholderType === "ftr");
    console.log("\nSlide", f.slideIndex, {
      ftrCount: ftrs.length,
      ftrs: ftrs.map((s) => ({ id: s.id, text: JSON.stringify(s.text) })),
      fixShapeId: f.fix?.shapeId,
      hasFix: !!f.fix,
      footerRect: ctx.aggregates.footerRect,
    });

    const prepared = await prepareFindingForApply(f, buf);
    console.log("Prepared fix:", prepared.fix?.kind, prepared.fix?.shapeId, prepared.fix?.newText?.slice(0, 40));

    if (prepared.fix) {
      const res = await applyFixToPptx(buf, prepared.fix, {
        findingId: f.id,
        ruleId: f.ruleId,
        title: f.title,
        source: "deterministic",
      });
      console.log("Apply result:", res.error || "OK", res.record?.after?.slice(0, 60));
    }
  }
}

main().catch(console.error);
