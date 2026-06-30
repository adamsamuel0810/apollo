"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import {
  ShieldCheck,
  LogOut,
  Download,
  RotateCcw,
  Eye,
  EyeOff,
  AlertCircle,
  AlertTriangle,
  Info,
  CheckCircle2,
  Wand2,
  Loader2,
  FileDown,
} from "lucide-react";
import Uploader from "@/components/Uploader";
import SlideCanvas from "@/components/SlideCanvas";
import FindingsPanel, { Decision } from "@/components/FindingsPanel";
import SlideRail, { RailItem } from "@/components/SlideRail";
import { AnalyzeResult } from "@/lib/analyze/types";
import { Finding, FixRecord } from "@/lib/rules/types";
import { isPrimary } from "@/lib/findings/merge";
import {
  allFindings,
  base64ToFile,
  buildFixReport,
  downloadBlob,
  findingKey,
} from "@/lib/export/report";
import { applyPreviewPatch } from "@/lib/preview/patch";
import { enrichFindingForPreview } from "@/lib/preview/enrichClient";

export default function Home() {
  const router = useRouter();
  const [result, setResult] = useState<AnalyzeResult | null>(null);
  const [originalFile, setOriginalFile] = useState<File | null>(null);
  const [workingFile, setWorkingFile] = useState<File | null>(null);
  /** Snapshots of findings the user accepted (stable across re-analyses). */
  const [acceptedSnapshots, setAcceptedSnapshots] = useState<Finding[]>([]);
  const [fixRecords, setFixRecords] = useState<FixRecord[]>([]);
  const [selected, setSelected] = useState(1);
  const [decisions, setDecisions] = useState<Record<string, Decision>>({});
  const [selectedFinding, setSelectedFinding] = useState<string | null>(null);
  const [showAll, setShowAll] = useState(true);
  const [fixingId, setFixingId] = useState<string | null>(null);
  const [autoFixing, setAutoFixing] = useState(false);
  const [previewKey, setPreviewKey] = useState(0);
  const [exportOpen, setExportOpen] = useState(false);
  const exportRef = useRef<HTMLDivElement>(null);

  const visibleForSlide = useMemo(() => {
    return (idx: number): Finding[] => {
      if (!result) return [];
      const slide = result.slides.find((s) => s.index === idx);
      if (!slide) return [];
      return slide.findings.filter((f) => (showAll ? true : isPrimary(f)));
    };
  }, [result, showAll]);

  const railItems: RailItem[] = useMemo(() => {
    if (!result) return [];
    return result.slides.map((s) => {
      const vis = s.findings.filter((f) => (showAll ? true : isPrimary(f)));
      const counts = { error: 0, warning: 0, info: 0 };
      for (const f of vis) counts[f.severity]++;
      return { index: s.index, kind: s.kind, counts, total: vis.length };
    });
  }, [result, showAll]);

  const totals = useMemo(() => {
    if (!result) return { error: 0, warning: 0, info: 0, accepted: 0, visible: 0, fixed: 0 };
    let error = 0,
      warning = 0,
      info = 0,
      visible = 0;
    for (const s of result.slides) {
      for (const f of s.findings) {
        if (!showAll && !isPrimary(f)) continue;
        visible++;
        if (f.severity === "error") error++;
        else if (f.severity === "warning") warning++;
        else info++;
      }
    }
    const accepted = Object.values(decisions).filter((d) => d === "accepted").length;
    return { error, warning, info, accepted, visible, fixed: fixRecords.length };
  }, [result, decisions, showAll, fixRecords.length]);

  /** Map stable finding keys to current finding ids after re-analyze. */
  const syncDecisions = useCallback((newResult: AnalyzeResult, snapshots: Finding[]) => {
    const acceptedKeys = new Set(snapshots.map(findingKey));
    const next: Record<string, Decision> = {};
    for (const f of allFindings(newResult)) {
      if (acceptedKeys.has(findingKey(f))) next[f.id] = "accepted";
    }
    setDecisions(next);
  }, []);

  function onAnalyzed(r: AnalyzeResult, file: File) {
    setResult(r);
    setOriginalFile(file);
    setWorkingFile(file);
    setAcceptedSnapshots([]);
    setFixRecords([]);
    setDecisions({});
    setSelectedFinding(null);
    setPreviewKey(0);
    const firstWithFindings =
      r.slides.find((s) => s.findings.some((f) => isPrimary(f)))?.index ||
      r.slides[0]?.index ||
      1;
    setSelected(firstWithFindings);
  }

  function optimisticFixRecord(finding: Finding): FixRecord {
    const fix = finding.fix;
    return {
      findingId: finding.id,
      ruleId: finding.ruleId,
      slideIndex: finding.slideIndex,
      title: finding.title,
      before: finding.evidence || finding.detail.slice(0, 120),
      after:
        fix?.newText ||
        (fix?.replace != null && fix.search ? fix.replace : finding.suggestion || "applied"),
      appliedAt: new Date().toISOString(),
      source: finding.source,
    };
  }

  async function acceptFix(finding: Finding) {
    if (!result || !originalFile) return;

    const slide = result.slides.find((s) => s.index === finding.slideIndex);
    const enriched = slide ? enrichFindingForPreview(finding, slide) : finding;

    const key = findingKey(enriched);
    const nextSnapshots = [
      ...acceptedSnapshots.filter((s) => findingKey(s) !== key),
      enriched,
    ];

    const previousResult = result;
    const previousSnapshots = acceptedSnapshots;
    const previousFixRecords = fixRecords;

    // 1. Instant preview — update the canvas immediately (no wait for server).
    const optimistic = applyPreviewPatch(result, enriched);
    for (const s of optimistic.slides) s.imageUrl = null;
    setResult(optimistic);
    setPreviewKey((k) => k + 1);
    setDecisions((prev) => ({ ...prev, [finding.id]: "accepted" }));
    setFixRecords((prev) => [
      ...prev.filter((r) => r.findingId !== finding.id),
      optimisticFixRecord(enriched),
    ]);

    setAcceptedSnapshots(nextSnapshots);
    setFixingId(finding.id);

    try {
      const fd = new FormData();
      fd.append("file", originalFile);
      fd.append("findings", JSON.stringify(nextSnapshots));
      const res = await fetch("/api/rebuild", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Fix failed");

      const newResult = data.result as AnalyzeResult;
      for (const s of newResult.slides) s.imageUrl = null;
      setResult(newResult);
      setWorkingFile(base64ToFile(data.pptxBase64, originalFile.name));
      setFixRecords(data.fixRecords as FixRecord[]);
      syncDecisions(newResult, nextSnapshots);
      setPreviewKey((k) => k + 1);

      if (data.fixErrors?.length) {
        console.warn("Partial fix errors:", data.fixErrors);
      }
    } catch (err) {
      setAcceptedSnapshots(previousSnapshots);
      setFixRecords(previousFixRecords);
      setDecisions((prev) => {
        const n = { ...prev };
        delete n[finding.id];
        return n;
      });
      if (previousResult) {
        setResult(previousResult);
        setPreviewKey((k) => k + 1);
      }
      alert(err instanceof Error ? err.message : "Could not apply fix");
    } finally {
      setFixingId(null);
    }
  }

  async function rejectFix(finding: Finding) {
    const key = findingKey(finding);
    const wasAccepted = acceptedSnapshots.some((s) => findingKey(s) === key);

    if (!wasAccepted) {
      setDecisions((prev) => ({ ...prev, [finding.id]: "rejected" }));
      return;
    }

    const nextSnapshots = acceptedSnapshots.filter((s) => findingKey(s) !== key);
    setAcceptedSnapshots(nextSnapshots);
    setDecisions((prev) => {
      const n = { ...prev };
      delete n[finding.id];
      n[finding.id] = "rejected";
      return n;
    });
    setFixingId(finding.id);

    try {
      const fd = new FormData();
      fd.append("file", originalFile!);
      fd.append("findings", JSON.stringify(nextSnapshots));
      const res = await fetch("/api/rebuild", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Revert failed");

      const newResult = data.result as AnalyzeResult;
      for (const s of newResult.slides) s.imageUrl = null;
      setResult(newResult);
      setWorkingFile(
        nextSnapshots.length
          ? base64ToFile(data.pptxBase64, originalFile!.name)
          : originalFile!
      );
      setFixRecords(data.fixRecords as FixRecord[]);
      syncDecisions(newResult, nextSnapshots);
      setPreviewKey((k) => k + 1);
    } catch (err) {
      alert(err instanceof Error ? err.message : "Could not revert fix");
    } finally {
      setFixingId(null);
    }
  }

  async function runAutoFixAll() {
    if (!originalFile || !result) return;
    if (!confirm("Apply all brand fixes at once using Anthropic AI? This may take a minute.")) return;
    setAutoFixing(true);
    try {
      const fd = new FormData();
      fd.append("file", originalFile);
      fd.append("useAi", "1");
      const res = await fetch("/api/autofix", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Auto-fix failed");

      const newResult = data.result as AnalyzeResult;
      for (const s of newResult.slides) s.imageUrl = null;
      setResult(newResult);
      setWorkingFile(base64ToFile(data.pptxBase64, originalFile.name));
      setFixRecords(data.fixRecords as FixRecord[]);
      const actionable = allFindings(result).filter(
        (f) => f.severity === "error" || f.severity === "warning"
      );
      setAcceptedSnapshots(actionable);
      syncDecisions(newResult, actionable);
      setPreviewKey((k) => k + 1);
      const errNote =
        data.fixErrors?.length > 0 ? `\n\nPartial errors: ${data.fixErrors.slice(0, 5).join("; ")}` : "";
      const remainNote =
        data.remainingIssues > 0
          ? `\n\n${data.remainingIssues} error/warning(s) may remain (layout-only issues that cannot be auto-fixed in PPTX).`
          : "\n\nAll errors and warnings addressed.";
      alert(
        `Applied ${data.appliedCount} of ${data.attemptedCount ?? data.appliedCount} fix(es) (${data.totalActionable ?? "?"} issues found).${remainNote}${errNote}`
      );
    } catch (err) {
      alert(err instanceof Error ? err.message : "Auto-fix failed");
    } finally {
      setAutoFixing(false);
    }
  }

  function exportReviewJson() {
    if (!result) return;
    const payload = {
      fileName: result.fileName,
      generatedAt: result.generatedAt,
      reviewedAt: new Date().toISOString(),
      acceptedCount: acceptedSnapshots.length,
      fixesApplied: fixRecords.length,
      acceptedFindings: acceptedSnapshots.map((f) => ({
        slide: f.slideIndex,
        severity: f.severity,
        title: f.title,
        detail: f.detail,
        evidence: f.evidence,
        suggestion: f.suggestion,
        guideline: f.guideline,
        source: f.source,
        ruleId: f.ruleId,
      })),
      fixRecords,
    };
    downloadBlob(
      new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }),
      `${result.fileName.replace(/\.pptx$/i, "")}-brand-review.json`
    );
    setExportOpen(false);
  }

  function exportFixReport() {
    if (!result) return;
    downloadBlob(
      new Blob([JSON.stringify(buildFixReport(result, fixRecords), null, 2)], {
        type: "application/json",
      }),
      `${result.fileName.replace(/\.pptx$/i, "")}-fix-report.json`
    );
    setExportOpen(false);
  }

  function exportPptx() {
    if (!workingFile || !result) return;
    const name = result.fileName.replace(/\.pptx$/i, "") + "-fixed.pptx";
    downloadBlob(workingFile, name);
    setExportOpen(false);
  }

  async function logout() {
    await fetch("/api/logout", { method: "POST" });
    router.replace("/login");
    router.refresh();
  }

  if (!result) return <Uploader onAnalyzed={onAnalyzed} />;

  const slide = result.slides.find((s) => s.index === selected) || result.slides[0];
  const slideFindings = visibleForSlide(slide.index);
  const previewBusy = fixingId != null || autoFixing;

  return (
    <div className="flex h-screen flex-col bg-slate-50">
      <header className="flex flex-none items-center gap-4 border-b border-slate-200 bg-white px-4 py-2.5">
        <div className="flex items-center gap-2">
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-brand text-white">
            <ShieldCheck className="h-4 w-4" />
          </div>
          <div className="leading-tight">
            <p className="text-sm font-semibold text-slate-800">Brand Compliance</p>
            <p className="max-w-[220px] truncate text-[11px] text-slate-400">{result.fileName}</p>
          </div>
        </div>

        <div className="ml-2 flex items-center gap-2">
          <Stat icon={<AlertCircle className="h-3.5 w-3.5 text-red-500" />} value={totals.error} label="errors" />
          <Stat icon={<AlertTriangle className="h-3.5 w-3.5 text-amber-500" />} value={totals.warning} label="warnings" />
          <Stat icon={<Info className="h-3.5 w-3.5 text-sky-500" />} value={totals.info} label="info" />
          <Stat icon={<CheckCircle2 className="h-3.5 w-3.5 text-emerald-500" />} value={totals.accepted} label="accepted" />
          {totals.fixed > 0 && (
            <Stat icon={<Wand2 className="h-3.5 w-3.5 text-brand" />} value={totals.fixed} label="fixed" />
          )}
        </div>

        <div className="ml-auto flex items-center gap-2">
          <span
            className="hidden items-center gap-1 rounded-full bg-slate-100 px-2 py-1 text-[11px] font-medium text-slate-500 sm:flex"
            title={result.aiEnabled ? "AI review enabled" : "Deterministic-only (no AI key configured)"}
          >
            {result.aiEnabled ? "AI + Rules" : "Rules only"}
          </span>
          <button
            onClick={runAutoFixAll}
            disabled={autoFixing || previewBusy}
            className="flex items-center gap-1.5 rounded-lg border border-brand/30 bg-brand/5 px-2.5 py-1.5 text-xs font-medium text-brand hover:bg-brand/10 disabled:opacity-60"
          >
            {autoFixing ? (
              <Loader2 className="h-3.5 w-3.5 animate-spin" />
            ) : (
              <Wand2 className="h-3.5 w-3.5" />
            )}
            {autoFixing ? "Fixing…" : "Auto-fix all"}
          </button>
          <button
            onClick={() => setShowAll((v) => !v)}
            className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
          >
            {showAll ? <EyeOff className="h-3.5 w-3.5" /> : <Eye className="h-3.5 w-3.5" />}
            {showAll ? "Primary only" : "Show all flags"}
          </button>
          <div className="relative" ref={exportRef}>
            <button
              onClick={() => setExportOpen((v) => !v)}
              className="flex items-center gap-1.5 rounded-lg bg-brand px-2.5 py-1.5 text-xs font-medium text-white hover:bg-brand-dark"
            >
              <Download className="h-3.5 w-3.5" /> Export
            </button>
            {exportOpen && (
              <div className="absolute right-0 top-full z-30 mt-1 w-52 rounded-lg border border-slate-200 bg-white py-1 shadow-lg">
                <ExportMenuItem
                  icon={<FileDown className="h-3.5 w-3.5" />}
                  label="Updated PPTX"
                  hint={fixRecords.length ? `${fixRecords.length} fix(es) in file` : "No fixes applied yet"}
                  onClick={exportPptx}
                />
                <ExportMenuItem
                  icon={<Download className="h-3.5 w-3.5" />}
                  label="Fix report (JSON)"
                  hint={`${fixRecords.length} change(s) logged`}
                  onClick={exportFixReport}
                />
                <ExportMenuItem
                  icon={<Download className="h-3.5 w-3.5" />}
                  label="Review decisions (JSON)"
                  hint="Accepted/rejected flags"
                  onClick={exportReviewJson}
                />
              </div>
            )}
          </div>
          <button
            onClick={() => {
              setResult(null);
              setOriginalFile(null);
              setWorkingFile(null);
              setAcceptedSnapshots([]);
              setFixRecords([]);
            }}
            className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
          >
            <RotateCcw className="h-3.5 w-3.5" /> New
          </button>
          <button
            onClick={logout}
            className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
          >
            <LogOut className="h-3.5 w-3.5" /> Sign out
          </button>
        </div>
      </header>

      <div className="grid min-h-0 flex-1 grid-cols-[140px_1fr_400px]">
        <aside className="min-h-0 border-r border-slate-200 bg-white">
          <SlideRail items={railItems} selected={selected} onSelect={setSelected} />
        </aside>

        <main className="min-h-0 overflow-auto p-6">
          <div className="mx-auto max-w-4xl">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-sm font-semibold text-slate-700">
                Slide {slide.index}{" "}
                <span className="text-xs font-normal text-slate-400">/ {result.slideCount}</span>
              </h2>
              <span className="text-xs text-slate-400">
                {slideFindings.length} {slideFindings.length === 1 ? "flag" : "flags"} shown
                {previewBusy && " · updating preview…"}
              </span>
            </div>
            <div className={previewBusy ? "opacity-70 transition-opacity" : ""}>
              <SlideCanvas
                key={`${slide.index}-${previewKey}`}
                slide={slide}
                findings={slideFindings}
                selectedId={selectedFinding}
                onSelectFinding={setSelectedFinding}
              />
            </div>
            <p className="mt-3 text-center text-xs text-slate-400">
              Accept &amp; Fix updates the live preview immediately. Reject restores the original
              for that item. Export the PPTX to open in PowerPoint.
            </p>
          </div>
        </main>

        <aside className="scrollbar-thin min-h-0 overflow-y-auto border-l border-slate-200 bg-slate-50">
          <FindingsPanel
            slideIndex={slide.index}
            findings={slideFindings}
            decisions={decisions}
            selectedId={selectedFinding}
            fixingId={fixingId}
            onSelect={setSelectedFinding}
            onDecision={(id, d) =>
              setDecisions((prev) => {
                const next = { ...prev };
                if (d == null) delete next[id];
                else next[id] = d;
                return next;
              })
            }
            onAcceptFix={acceptFix}
            onRejectFix={rejectFix}
          />
        </aside>
      </div>
    </div>
  );
}

function Stat({ icon, value, label }: { icon: React.ReactNode; value: number; label: string }) {
  return (
    <span className="flex items-center gap-1 rounded-lg bg-slate-50 px-2 py-1 text-xs">
      {icon}
      <span className="font-semibold text-slate-700">{value}</span>
      <span className="hidden text-slate-400 md:inline">{label}</span>
    </span>
  );
}

function ExportMenuItem({
  icon,
  label,
  hint,
  onClick,
}: {
  icon: React.ReactNode;
  label: string;
  hint: string;
  onClick: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className="flex w-full flex-col items-start gap-0.5 px-3 py-2 text-left hover:bg-slate-50"
    >
      <span className="flex items-center gap-1.5 text-xs font-medium text-slate-700">
        {icon}
        {label}
      </span>
      <span className="pl-5 text-[10px] text-slate-400">{hint}</span>
    </button>
  );
}
