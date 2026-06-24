"use client";

import { cn } from "@/lib/utils";

export interface RailItem {
  index: number;
  kind: string;
  counts: { error: number; warning: number; info: number };
  total: number;
}

interface Props {
  items: RailItem[];
  selected: number;
  onSelect: (index: number) => void;
}

export default function SlideRail({ items, selected, onSelect }: Props) {
  return (
    <div className="scrollbar-thin flex h-full flex-col gap-1 overflow-y-auto p-2">
      {items.map((it) => {
        const isSel = it.index === selected;
        return (
          <button
            key={it.index}
            onClick={() => onSelect(it.index)}
            className={cn(
              "flex items-center justify-between rounded-lg px-3 py-2 text-left text-sm transition",
              isSel ? "bg-brand text-white" : "text-slate-700 hover:bg-slate-100"
            )}
          >
            <span className="flex items-center gap-2">
              <span className={cn("font-medium", isSel ? "text-white" : "text-slate-800")}>
                {it.index}
              </span>
              <span className={cn("text-[10px] uppercase", isSel ? "text-white/70" : "text-slate-400")}>
                {it.kind}
              </span>
            </span>
            <span className="flex items-center gap-1">
              {it.counts.error > 0 && (
                <Badge sel={isSel} color="bg-red-500">
                  {it.counts.error}
                </Badge>
              )}
              {it.counts.warning > 0 && (
                <Badge sel={isSel} color="bg-amber-500">
                  {it.counts.warning}
                </Badge>
              )}
              {it.counts.info > 0 && (
                <Badge sel={isSel} color="bg-sky-500">
                  {it.counts.info}
                </Badge>
              )}
              {it.total === 0 && (
                <span className={cn("text-[10px]", isSel ? "text-white/60" : "text-emerald-500")}>
                  ✓
                </span>
              )}
            </span>
          </button>
        );
      })}
    </div>
  );
}

function Badge({ children, color, sel }: { children: React.ReactNode; color: string; sel: boolean }) {
  return (
    <span
      className={cn(
        "flex h-4 min-w-4 items-center justify-center rounded-full px-1 text-[10px] font-bold text-white",
        sel ? "bg-white/30" : color
      )}
    >
      {children}
    </span>
  );
}
