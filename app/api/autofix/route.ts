import { NextRequest, NextResponse } from "next/server";
import { analyzePptx } from "@/lib/analyze/analyze";
import { applyAllFixes } from "@/lib/fixes/apply";
import { mergeFindings } from "@/lib/findings/merge";
import { runDeterministicRules } from "@/lib/rules/engine";
import { runAiChecks } from "@/lib/ai/orchestrator";
import { parseDeck } from "@/lib/pptx/parser";
import { Finding } from "@/lib/rules/types";

export const runtime = "nodejs";
export const maxDuration = 300;

const MAX_BYTES = 30 * 1024 * 1024;

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: "Expected multipart/form-data" }, { status: 400 });
  }

  const file = formData.get("file");
  const useAi = formData.get("useAi") !== "0";
  if (!(file instanceof File)) {
    return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
  }
  if (file.size > MAX_BYTES) {
    return NextResponse.json({ error: "File too large" }, { status: 413 });
  }

  try {
    const buf = Buffer.from(await file.arrayBuffer());
    const deck = await parseDeck(buf);
    const { findings: deterministic, ctx } = runDeterministicRules(deck);
    let aiFindings: Finding[] = [];
    if (useAi) {
      try {
        aiFindings = await runAiChecks(deck, ctx.aggregates);
      } catch (err) {
        console.error("AI checks during autofix:", err);
      }
    }
    const merged = mergeFindings([...deterministic, ...aiFindings]);

    const actionable = merged.filter(
      (f) => f.severity === "error" || f.severity === "warning"
    );

    const { buffer, records, errors, attemptedCount } = await applyAllFixes(buf, merged, {
      includeAi: useAi,
    });

    const result = await analyzePptx(buffer, file.name, { skipImages: true });
    const remaining = result.slides.reduce((n, s) => {
      for (const f of s.findings) {
        if (f.severity === "error" || f.severity === "warning") n++;
      }
      return n;
    }, 0);

    return NextResponse.json({
      pptxBase64: buffer.toString("base64"),
      result,
      fixRecords: records,
      fixErrors: errors,
      appliedCount: records.length,
      attemptedCount,
      totalActionable: actionable.length,
      remainingIssues: remaining,
    });
  } catch (err) {
    console.error("Autofix failed:", err);
    return NextResponse.json({ error: "Auto-fix failed" }, { status: 500 });
  }
}
