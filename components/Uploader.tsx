"use client";

import { useCallback, useRef, useState } from "react";
import { UploadCloud, FileText, Loader2, ShieldCheck, Wand2 } from "lucide-react";
import { cn } from "@/lib/utils";
import { AnalyzeResult } from "@/lib/analyze/types";

interface Props {
  onAnalyzed: (result: AnalyzeResult, file: File) => void;
}

export default function Uploader({ onAnalyzed }: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");

  const upload = useCallback(
    async (file: File) => {
      setError("");
      if (!file.name.toLowerCase().endsWith(".pptx")) {
        setError("Please upload a .pptx file.");
        return;
      }
      setLoading(true);
      setStatus("Parsing slides and running brand checks…");
      try {
        const fd = new FormData();
        fd.append("file", file);
        const res = await fetch("/api/analyze", { method: "POST", body: fd });
        if (!res.ok) {
          const data = await res.json().catch(() => ({}));
          throw new Error(data.error || `Analysis failed (${res.status})`);
        }
        const result = (await res.json()) as AnalyzeResult;
        onAnalyzed(result, file);
      } catch (e) {
        setError(e instanceof Error ? e.message : "Something went wrong.");
      } finally {
        setLoading(false);
        setStatus("");
      }
    },
    [onAnalyzed]
  );

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const file = e.dataTransfer.files?.[0];
      if (file) upload(file);
    },
    [upload]
  );

  return (
    <div className="mx-auto flex min-h-screen max-w-2xl flex-col items-center justify-center px-4">
      <div className="mb-8 text-center">
        <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-brand text-white shadow-lg">
          <ShieldCheck className="h-7 w-7" />
        </div>
        <h1 className="text-2xl font-semibold tracking-tight text-slate-900">
          ACME Brand Compliance Checker
        </h1>
        <p className="mt-2 text-sm text-slate-500">
          Upload a PowerPoint to flag brand issues, accept fixes to update the deck
          automatically, and export the corrected file.
        </p>
      </div>

      <div
        onClick={() => !loading && inputRef.current?.click()}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        className={cn(
          "flex w-full cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed bg-white px-6 py-14 text-center transition",
          dragging ? "border-brand bg-brand/5" : "border-slate-300 hover:border-brand/60",
          loading && "pointer-events-none opacity-80"
        )}
      >
        {loading ? (
          <>
            <Loader2 className="mb-3 h-8 w-8 animate-spin text-brand" />
            <p className="text-sm font-medium text-slate-700">{status}</p>
            <p className="mt-1 text-xs text-slate-400">
              This can take up to a minute for large decks.
            </p>
          </>
        ) : (
          <>
            <UploadCloud className="mb-3 h-9 w-9 text-slate-400" />
            <p className="text-sm font-medium text-slate-700">
              Drag &amp; drop a .pptx here, or click to browse
            </p>
            <p className="mt-1 flex items-center gap-1 text-xs text-slate-400">
              <FileText className="h-3.5 w-3.5" /> PowerPoint (.pptx) up to 30 MB
            </p>
          </>
        )}
        <input
          ref={inputRef}
          type="file"
          accept=".pptx"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) upload(f);
            e.target.value = "";
          }}
        />
      </div>

      {error && (
        <p className="mt-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</p>
      )}

      <p className="mt-6 flex items-center justify-center gap-1.5 text-center text-xs text-slate-400">
        <Wand2 className="h-3.5 w-3.5" />
        Accept a flag to auto-fix the deck · Export corrected PPTX + fix report
      </p>
    </div>
  );
}
