import { Finding } from "../rules/types";
import { buildBatchFixPlan } from "../ai/batchAutofix";
import { applyFixToPptx, applyFixesToPptx } from "../pptx/writer";
import { aiEnabled } from "../ai/client";
import { prepareFindingForApply } from "./prepareFinding";

export interface FixOutcome {
  buffer: Buffer;
  records: import("../rules/types").FixRecord[];
  errors: string[];
  attemptedCount: number;
}

/** Apply a single accepted finding fix, resolving AI rewrites when needed. */
export async function applyAcceptedFinding(
  pptx: Buffer,
  finding: Finding
): Promise<FixOutcome> {
  const prepared = await prepareFindingForApply(finding, pptx);
  if (!prepared.fix) {
    return {
      buffer: pptx,
      records: [],
      errors: ["This finding has no auto-fix available."],
      attemptedCount: 0,
    };
  }

  const res = await applyFixToPptx(pptx, prepared.fix, {
    findingId: prepared.id,
    ruleId: prepared.ruleId,
    title: prepared.title,
    source: prepared.source,
  });

  return {
    buffer: res.buffer,
    records: res.record ? [res.record] : [],
    errors: res.error && !res.record ? [res.error] : [],
    attemptedCount: 1,
  };
}

/** Apply all fixable findings using Anthropic AI batch plan + deterministic fixes. */
export async function applyAllFixes(
  pptx: Buffer,
  findings: Finding[],
  options: { includeAi?: boolean } = {}
): Promise<FixOutcome> {
  if (options.includeAi && aiEnabled()) {
    const items = await buildBatchFixPlan(pptx, findings);
    const { buffer, records, errors } = await applyFixesToPptx(pptx, items);
    return { buffer, records, errors, attemptedCount: items.length };
  }

  const fixable = findings.filter((f) => f.fix && f.fix.kind !== "ai-rewrite");
  const items: {
    fix: import("../rules/types").FixInstruction;
    findingId: string;
    ruleId: string;
    title: string;
    source: "deterministic" | "ai";
  }[] = [];

  for (const f of fixable) {
    items.push({
      fix: { ...f.fix! },
      findingId: f.id,
      ruleId: f.ruleId,
      title: f.title,
      source: f.source,
    });
  }

  const { buffer, records, errors } = await applyFixesToPptx(pptx, items);
  return { buffer, records, errors, attemptedCount: items.length };
}
