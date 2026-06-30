import { FixRecord, Finding } from "@/lib/rules/types";
import { AnalyzeResult } from "@/lib/analyze/types";

export function buildFixReport(
  result: AnalyzeResult,
  fixRecords: FixRecord[]
): Record<string, unknown> {
  return {
    fileName: result.fileName,
    generatedAt: result.generatedAt,
    exportedAt: new Date().toISOString(),
    fixesApplied: fixRecords.length,
    fixes: fixRecords.map((r) => ({
      slide: r.slideIndex,
      rule: r.ruleId,
      title: r.title,
      before: r.before,
      after: r.after,
      appliedAt: r.appliedAt,
      source: r.source,
    })),
  };
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function downloadPptx(base64: string, fileName: string) {
  const bytes = base64ToBytes(base64);
  const name = fileName.toLowerCase().endsWith(".pptx")
    ? fileName.replace(/\.pptx$/i, "-fixed.pptx")
    : `${fileName}-fixed.pptx`;
  downloadBlob(
    new Blob([new Uint8Array(bytes)], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }),
    name
  );
}

/** Stable key for tracking accept/reject across re-analyses (finding ids change). */
export function findingKey(f: Finding): string {
  return `${f.ruleId}:${f.slideIndex}:${f.shapeId ?? ""}:${f.cell?.row ?? ""}:${f.cell?.col ?? ""}`;
}

function base64ToBytes(base64: string): Uint8Array {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

export function base64ToFile(base64: string, fileName: string): File {
  const bytes = base64ToBytes(base64);
  return new File([new Uint8Array(bytes)], fileName, {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

/** Collect every finding from an analyze result into a flat list. */
export function allFindings(result: AnalyzeResult): Finding[] {
  return result.slides.flatMap((s) => s.findings);
}
