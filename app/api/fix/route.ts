import { NextRequest, NextResponse } from "next/server";
import { analyzePptx } from "@/lib/analyze/analyze";
import { applyAcceptedFinding } from "@/lib/fixes/apply";
import { Finding } from "@/lib/rules/types";

export const runtime = "nodejs";
export const maxDuration = 120;

const MAX_BYTES = 30 * 1024 * 1024;

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: "Expected multipart/form-data" }, { status: 400 });
  }

  const file = formData.get("file");
  const findingRaw = formData.get("finding");
  if (!(file instanceof File)) {
    return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
  }
  if (typeof findingRaw !== "string") {
    return NextResponse.json({ error: "Missing finding JSON" }, { status: 400 });
  }
  if (file.size > MAX_BYTES) {
    return NextResponse.json({ error: "File too large" }, { status: 413 });
  }

  let finding: Finding;
  try {
    finding = JSON.parse(findingRaw) as Finding;
  } catch {
    return NextResponse.json({ error: "Invalid finding JSON" }, { status: 400 });
  }

  try {
    const buf = Buffer.from(await file.arrayBuffer());
    const { buffer, records, errors } = await applyAcceptedFinding(buf, finding);
    if (records.length === 0) {
      return NextResponse.json(
        { error: errors[0] || "Fix could not be applied" },
        { status: 422 }
      );
    }

    const result = await analyzePptx(buffer, file.name, { skipImages: true });
    return NextResponse.json({
      pptxBase64: buffer.toString("base64"),
      result,
      fixRecords: records,
    });
  } catch (err) {
    console.error("Fix failed:", err);
    return NextResponse.json({ error: "Failed to apply fix" }, { status: 500 });
  }
}
