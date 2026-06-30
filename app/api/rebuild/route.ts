import { NextRequest, NextResponse } from "next/server";
import { analyzePptx } from "@/lib/analyze/analyze";
import { applyFixesToPptx } from "@/lib/pptx/writer";
import { prepareFindingsForApply } from "@/lib/fixes/prepareFinding";
import { Finding } from "@/lib/rules/types";

export const runtime = "nodejs";
export const maxDuration = 300;

const MAX_BYTES = 30 * 1024 * 1024;

/** Rebuild deck from the original file by applying a list of accepted findings. */
export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: "Expected multipart/form-data" }, { status: 400 });
  }

  const file = formData.get("file");
  const findingsRaw = formData.get("findings");
  if (!(file instanceof File)) {
    return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
  }
  if (typeof findingsRaw !== "string") {
    return NextResponse.json({ error: "Missing findings JSON" }, { status: 400 });
  }
  if (file.size > MAX_BYTES) {
    return NextResponse.json({ error: "File too large" }, { status: 413 });
  }

  let findings: Finding[];
  try {
    findings = JSON.parse(findingsRaw) as Finding[];
  } catch {
    return NextResponse.json({ error: "Invalid findings JSON" }, { status: 400 });
  }

  try {
    const buf = Buffer.from(await file.arrayBuffer());
    const prepared = await prepareFindingsForApply(findings, buf);

    const items: {
      fix: NonNullable<Finding["fix"]>;
      findingId: string;
      ruleId: string;
      title: string;
      source: "deterministic" | "ai";
    }[] = [];

    for (const f of prepared) {
      if (!f.fix) continue;
      items.push({
        fix: f.fix,
        findingId: f.id,
        ruleId: f.ruleId,
        title: f.title,
        source: f.source,
      });
    }

    const { buffer, records, errors } = await applyFixesToPptx(buf, items);
    const result = await analyzePptx(buffer, file.name, { skipImages: true });

    return NextResponse.json({
      result,
      fixRecords: records,
      fixErrors: errors,
      pptxBase64: buffer.toString("base64"),
    });
  } catch (err) {
    console.error("Rebuild failed:", err);
    return NextResponse.json({ error: "Failed to rebuild deck" }, { status: 500 });
  }
}
