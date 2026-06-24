"use client";

import { Finding } from "@/lib/rules/types";
import { Check, X, Sparkles, Cog, BookOpen, CornerDownRight, CircleCheck } from "lucide-react";
import { cn } from "@/lib/utils";

export type Decision = "accepted" | "rejected";

interface Props {
  slideIndex: number;
  findings: Finding[]; // already in display order (matches canvas numbering)
  decisions: Record<string, Decision>;
  selectedId: string | null;
  onSelect: (id: string) => void;
  onDecision: (id: string, decision: Decision | null) => void;
}

const SEV_STYLES: Record<string, { dot: string; chip: string; label: string }> = {
  error: { dot: "bg-red-500", chip: "bg-red-50 text-red-700 ring-red-200", label: "Error" },
  warning: { dot: "bg-amber-500", chip: "bg-amber-50 text-amber-700 ring-amber-200", label: "Warning" },
  info: { dot: "bg-sky-500", chip: "bg-sky-50 text-sky-700 ring-sky-200", label: "Info" },
};

export default function FindingsPanel({
  slideIndex,
  findings,
  decisions,
  selectedId,
  onSelect,
  onDecision,
}: Props) {
  if (findings.length === 0) {
    return (
      <div className="flex h-full flex-col items-center justify-center px-6 text-center">
        <CircleCheck className="mb-3 h-10 w-10 text-emerald-500" />
        <p className="text-sm font-medium text-slate-700">No flags on slide {slideIndex}</p>
        <p className="mt-1 text-xs text-slate-400">
          This slide passed the visible brand checks.
        </p>
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-3 p-4">
      {findings.map((f, idx) => {
        const sev = SEV_STYLES[f.severity] || SEV_STYLES.info;
        const decision = decisions[f.id];
        const selected = f.id === selectedId;
        return (
          <div
            key={f.id}
            onClick={() => onSelect(f.id)}
            className={cn(
              "cursor-pointer rounded-xl border bg-white p-3.5 shadow-sm transition",
              selected ? "border-brand ring-2 ring-brand/20" : "border-slate-200 hover:border-slate-300",
              decision === "rejected" && "opacity-50",
              decision === "accepted" && "border-emerald-300 bg-emerald-50/40"
            )}
          >
            <div className="flex items-start gap-2.5">
              <span
                className={cn(
                  "mt-0.5 flex h-5 w-5 flex-none items-center justify-center rounded-full text-[10px] font-bold text-white",
                  sev.dot
                )}
              >
                {idx + 1}
              </span>
              <div className="min-w-0 flex-1">
                <div className="flex flex-wrap items-center gap-1.5">
                  <span
                    className={cn(
                      "rounded-full px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-wide ring-1 ring-inset",
                      sev.chip
                    )}
                  >
                    {sev.label}
                  </span>
                  <span
                    className="flex items-center gap-1 rounded-full bg-slate-100 px-1.5 py-0.5 text-[10px] font-medium text-slate-500"
                    title={f.source === "ai" ? "Flagged by AI review" : "Flagged by deterministic check"}
                  >
                    {f.source === "ai" ? (
                      <Sparkles className="h-2.5 w-2.5" />
                    ) : (
                      <Cog className="h-2.5 w-2.5" />
                    )}
                    {f.source === "ai" ? "AI" : "Auto"}
                  </span>
                  <span className="ml-auto text-[10px] font-medium text-slate-400">
                    {Math.round(f.confidence * 100)}%
                  </span>
                </div>

                <h3 className="mt-1.5 text-sm font-semibold leading-snug text-slate-800">
                  {f.title}
                </h3>
                <p className="mt-1 text-xs leading-relaxed text-slate-600">{f.detail}</p>

                {f.evidence && (
                  <p className="mt-1.5 rounded-md bg-slate-50 px-2 py-1 text-[11px] font-mono text-slate-500">
                    “{f.evidence}”
                  </p>
                )}

                {f.suggestion && (
                  <p className="mt-1.5 flex items-start gap-1 text-[11px] text-emerald-700">
                    <CornerDownRight className="mt-0.5 h-3 w-3 flex-none" />
                    {f.suggestion}
                  </p>
                )}

                <p className="mt-1.5 flex items-start gap-1 text-[10px] italic text-slate-400">
                  <BookOpen className="mt-0.5 h-3 w-3 flex-none" />
                  {f.guideline}
                </p>

                <div className="mt-2.5 flex items-center gap-2">
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onDecision(f.id, decision === "accepted" ? null : "accepted");
                    }}
                    className={cn(
                      "flex items-center gap-1 rounded-lg px-2.5 py-1 text-xs font-medium transition",
                      decision === "accepted"
                        ? "bg-emerald-600 text-white"
                        : "bg-emerald-50 text-emerald-700 hover:bg-emerald-100"
                    )}
                  >
                    <Check className="h-3.5 w-3.5" />
                    {decision === "accepted" ? "Accepted" : "Accept"}
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onDecision(f.id, decision === "rejected" ? null : "rejected");
                    }}
                    className={cn(
                      "flex items-center gap-1 rounded-lg px-2.5 py-1 text-xs font-medium transition",
                      decision === "rejected"
                        ? "bg-slate-700 text-white"
                        : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                    )}
                  >
                    <X className="h-3.5 w-3.5" />
                    {decision === "rejected" ? "Rejected" : "Reject"}
                  </button>
                </div>
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}
