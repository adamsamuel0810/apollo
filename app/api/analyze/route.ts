import { NextRequest, NextResponse } from "next/server";
import { analyzePptx } from "@/lib/analyze/analyze";

export const runtime = "nodejs";
// 60s is the Vercel Hobby cap; Pro supports up to 300s for very large decks.
export const maxDuration = 60;

const MAX_BYTES = 30 * 1024 * 1024; // 30 MB

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: "Expected multipart/form-data" }, { status: 400 });
  }

  const file = formData.get("file");
  if (!(file instanceof File)) {
    return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
  }
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    return NextResponse.json({ error: "Please upload a .pptx file" }, { status: 400 });
  }
  if (file.size > MAX_BYTES) {
    return NextResponse.json({ error: "File too large (max 30 MB)" }, { status: 413 });
  }

  try {
    const buf = Buffer.from(await file.arrayBuffer());
    const result = await analyzePptx(buf, file.name);
    return NextResponse.json(result);
  } catch (err) {
    console.error("Analyze failed:", err);
    return NextResponse.json(
      { error: "Failed to analyze the presentation. It may be corrupt or unsupported." },
      { status: 500 }
    );
  }
}
