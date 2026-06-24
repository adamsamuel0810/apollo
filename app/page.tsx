"use client";

import { useMemo, useState } from "react";
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
} from "lucide-react";
import Uploader from "@/components/Uploader";
import SlideCanvas from "@/components/SlideCanvas";
import FindingsPanel, { Decision } from "@/components/FindingsPanel";
import SlideRail, { RailItem } from "@/components/SlideRail";
import { AnalyzeResult } from "@/lib/analyze/types";
import { Finding } from "@/lib/rules/types";
import { isPrimary } from "@/lib/findings/merge";

export default function Home() {
  const router = useRouter();
  const [result, setResult] = useState<AnalyzeResult | null>(null);
  const [selected, setSelected] = useState(1);
  const [decisions, setDecisions] = useState<Record<string, Decision>>({});
  const [selectedFinding, setSelectedFinding] = useState<string | null>(null);
  const [showAll, setShowAll] = useState(false);

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
    if (!result) return { error: 0, warning: 0, info: 0, accepted: 0, visible: 0 };
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
    return { error, warning, info, accepted, visible };
  }, [result, decisions, showAll]);

  function onAnalyzed(r: AnalyzeResult) {
    setResult(r);
    setDecisions({});
    setSelectedFinding(null);
    const firstWithFindings =
      r.slides.find((s) => s.findings.some((f) => isPrimary(f)))?.index ||
      r.slides[0]?.index ||
      1;
    setSelected(firstWithFindings);
  }

  function setDecision(id: string, decision: Decision | null) {
    setDecisions((prev) => {
      const next = { ...prev };
      if (decision == null) delete next[id];
      else next[id] = decision;
      return next;
    });
  }

  function exportReport() {
    if (!result) return;
    const accepted: Finding[] = [];
    for (const s of result.slides)
      for (const f of s.findings) if (decisions[f.id] === "accepted") accepted.push(f);
    const payload = {
      fileName: result.fileName,
      generatedAt: result.generatedAt,
      reviewedAt: new Date().toISOString(),
      acceptedCount: accepted.length,
      acceptedFindings: accepted.map((f) => ({
        slide: f.slideIndex,
        severity: f.severity,
        title: f.title,
        detail: f.detail,
        evidence: f.evidence,
        suggestion: f.suggestion,
        guideline: f.guideline,
        source: f.source,
      })),
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${result.fileName.replace(/\.pptx$/i, "")}-brand-review.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function logout() {
    await fetch("/api/logout", { method: "POST" });
    router.replace("/login");
    router.refresh();
  }

  if (!result) return <Uploader onAnalyzed={onAnalyzed} />;

  const slide = result.slides.find((s) => s.index === selected) || result.slides[0];
  const slideFindings = visibleForSlide(slide.index);

  return (
    <div className="flex h-screen flex-col bg-slate-50">
      {/* Top bar */}
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
        </div>

        <div className="ml-auto flex items-center gap-2">
          <span
            className="hidden items-center gap-1 rounded-full bg-slate-100 px-2 py-1 text-[11px] font-medium text-slate-500 sm:flex"
            title={result.aiEnabled ? "AI review enabled" : "Deterministic-only (no AI key configured)"}
          >
            {result.aiEnabled ? "AI + Rules" : "Rules only"}
          </span>
          <button
            onClick={() => setShowAll((v) => !v)}
            className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
          >
            {showAll ? <EyeOff className="h-3.5 w-3.5" /> : <Eye className="h-3.5 w-3.5" />}
            {showAll ? "Hide low-confidence" : "Show all"}
          </button>
          <button
            onClick={exportReport}
            className="flex items-center gap-1.5 rounded-lg bg-brand px-2.5 py-1.5 text-xs font-medium text-white hover:bg-brand-dark"
          >
            <Download className="h-3.5 w-3.5" /> Export
          </button>
          <button
            onClick={() => setResult(null)}
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

      {/* Body */}
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
              </span>
            </div>
            <SlideCanvas
              slide={slide}
              findings={slideFindings}
              selectedId={selectedFinding}
              onSelectFinding={setSelectedFinding}
            />
            <p className="mt-3 text-center text-xs text-slate-400">
              Numbered boxes mark flagged elements. Click a box or a card to focus it.
            </p>
          </div>
        </main>

        <aside className="scrollbar-thin min-h-0 overflow-y-auto border-l border-slate-200 bg-slate-50">
          <FindingsPanel
            slideIndex={slide.index}
            findings={slideFindings}
            decisions={decisions}
            selectedId={selectedFinding}
            onSelect={setSelectedFinding}
            onDecision={setDecision}
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
