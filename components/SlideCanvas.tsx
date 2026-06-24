"use client";

import { useEffect, useRef, useState } from "react";
import { Shape, Paragraph, Rect } from "@/lib/pptx/types";
import { Finding } from "@/lib/rules/types";
import { AnalyzeSlide } from "@/lib/analyze/types";
import { cn } from "@/lib/utils";
import {
  bulletDisplay,
  bodyFrameStyle,
  cssFontFamily,
  paragraphLayoutStyle,
  ptToPx,
} from "@/lib/pptx/textRender";

interface Props {
  slide: AnalyzeSlide;
  findings: Finding[]; // findings to highlight (already filtered)
  selectedId?: string | null;
  onSelectFinding?: (id: string) => void;
}

const SEVERITY_BORDER: Record<string, string> = {
  error: "border-red-500",
  warning: "border-amber-500",
  info: "border-sky-500",
};
const SEVERITY_BG: Record<string, string> = {
  error: "bg-red-500",
  warning: "bg-amber-500",
  info: "bg-sky-500",
};

export default function SlideCanvas({
  slide,
  findings,
  selectedId,
  onSelectFinding,
}: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(960);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const ro = new ResizeObserver((entries) => {
      for (const e of entries) setWidth(e.contentRect.width);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const { cx, cy } = slide.size;
  const pxPerEmu = width / cx;
  const heightPx = cy * pxPerEmu;

  return (
    <div
      ref={ref}
      className="relative w-full overflow-hidden rounded-lg border border-slate-300 bg-white shadow-sm"
      style={{ height: heightPx }}
    >
      {slide.imageUrl ? (
        // Pixel-identical slide from the cloud renderer.
        // eslint-disable-next-line @next/next/no-img-element
        <img
          src={slide.imageUrl}
          alt={`Slide ${slide.index}`}
          className="absolute inset-0 h-full w-full select-none"
          draggable={false}
        />
      ) : (
        slide.shapes.map((s, i) => (
          <ShapeView key={s.id || i} shape={s} pxPerEmu={pxPerEmu} />
        ))
      )}

      {/* Highlight overlays */}
      {findings.map((f, idx) => {
        const r = f.rect || shapeRect(slide.shapes, f.shapeId);
        if (!r) return null;
        const selected = f.id === selectedId;
        return (
          <button
            key={f.id}
            onClick={() => onSelectFinding?.(f.id)}
            className={cn(
              "group absolute z-20 rounded-[3px] border-2 transition-all",
              SEVERITY_BORDER[f.severity] || "border-red-500",
              selected
                ? "bg-amber-300/20 ring-2 ring-offset-1 ring-amber-400"
                : "bg-transparent hover:bg-amber-200/10"
            )}
            style={{
              left: r.x * pxPerEmu,
              top: r.y * pxPerEmu,
              width: Math.max(8, r.cx * pxPerEmu),
              height: Math.max(8, r.cy * pxPerEmu),
            }}
            title={f.title}
          >
            <span
              className={cn(
                "absolute -left-2 -top-2 flex h-4 w-4 items-center justify-center rounded-full text-[9px] font-bold text-white shadow",
                SEVERITY_BG[f.severity] || "bg-red-500"
              )}
            >
              {idx + 1}
            </span>
          </button>
        );
      })}
    </div>
  );
}

function shapeRect(shapes: Shape[], id?: string): Rect | null {
  if (!id) return null;
  const s = shapes.find((x) => x.id === id);
  return s?.rect || null;
}

function ShapeView({ shape, pxPerEmu }: { shape: Shape; pxPerEmu: number }) {
  const r = shape.rect;
  if (!r) return null;
  const style: React.CSSProperties = {
    left: r.x * pxPerEmu,
    top: r.y * pxPerEmu,
    width: r.cx * pxPerEmu,
    height: r.cy * pxPerEmu,
  };

  if (shape.kind === "picture" && shape.imageDataUrl) {
    return (
      // eslint-disable-next-line @next/next/no-img-element
      <img
        src={shape.imageDataUrl}
        alt={shape.name}
        className="absolute object-contain"
        style={style}
      />
    );
  }

  if (shape.kind === "chart") {
    return (
      <div
        className="absolute flex items-center justify-center rounded border border-dashed border-slate-300 bg-slate-50 text-[10px] text-slate-400"
        style={style}
      >
        Chart / Visual
      </div>
    );
  }

  if (shape.kind === "table" && shape.table) {
    return <TableView shape={shape} pxPerEmu={pxPerEmu} style={style} />;
  }

  // Text / shape-with-fill
  const bg = shape.fill ? `#${shape.fill}` : undefined;
  const frame = bodyFrameStyle(shape.bodyPr, pxPerEmu);
  return (
    <div
      className="absolute overflow-hidden"
      style={{ ...style, backgroundColor: bg }}
    >
      <div
        className="flex h-full w-full flex-col"
        style={{
          justifyContent: frame.justifyContent,
          paddingTop: frame.paddingTop,
          paddingRight: frame.paddingRight,
          paddingBottom: frame.paddingBottom,
          paddingLeft: frame.paddingLeft,
        }}
      >
        {shape.paragraphs.map((p, i) => (
          <ParagraphView key={i} para={p} pxPerEmu={pxPerEmu} />
        ))}
      </div>
    </div>
  );
}

function ParagraphView({ para, pxPerEmu }: { para: Paragraph; pxPerEmu: number }) {
  const layout = paragraphLayoutStyle(para, pxPerEmu);
  const bullet = bulletDisplay(para);
  const defaultSizePt = para.runs.find((r) => r.sizePt)?.sizePt ?? 16;

  // Hanging indent: PowerPoint uses negative firstLineIndent for bullet glyphs.
  const hangPx =
    bullet != null
      ? layout.textIndentPx < 0
        ? -layout.textIndentPx
        : Math.max(ptToPx(defaultSizePt, pxPerEmu) * 0.38, 6)
      : 0;
  const padLeft = layout.marginLeftPx + (bullet ? hangPx : 0);

  return (
    <p
      style={{
        textAlign: layout.textAlign,
        margin: `${layout.marginTopPx}px 0 ${layout.marginBottomPx}px 0`,
        paddingLeft: padLeft,
        textIndent: bullet ? -hangPx : layout.textIndentPx,
        lineHeight: layout.lineHeight,
      }}
    >
      {bullet && (
        <span
          style={{
            fontFamily: bullet.fontFamily,
            fontWeight: 400,
            marginRight: hangPx * 0.15,
          }}
        >
          {bullet.char}
        </span>
      )}
      {para.runs
        .filter((r) => r.text !== "\n")
        .map((run, i) => {
          const sizePx = run.sizePt
            ? ptToPx(run.sizePt, pxPerEmu)
            : ptToPx(defaultSizePt, pxPerEmu);
          return (
            <span
              key={i}
              style={{
                fontSize: sizePx,
                fontWeight: run.bold ? 700 : 400,
                fontStyle: run.italic ? "italic" : "normal",
                textDecoration: run.underline
                  ? "underline"
                  : run.strike
                    ? "line-through"
                    : "none",
                fontFamily: cssFontFamily(run.font),
                color: run.color ? `#${run.color}` : "#1a1a1a",
                whiteSpace: "pre-wrap",
              }}
            >
              {run.text}
            </span>
          );
        })}
    </p>
  );
}

function TableView({
  shape,
  pxPerEmu,
  style,
}: {
  shape: Shape;
  pxPerEmu: number;
  style: React.CSSProperties;
}) {
  const table = shape.table!;
  return (
    <div className="absolute overflow-hidden" style={style}>
      <table
        className="border-collapse"
        style={{ width: "100%", height: "100%", tableLayout: "fixed" }}
      >
        <tbody>
          {table.rows.map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => {
                if (cell.hMerge || cell.vMerge) return null;
                const run = cell.runs.find((r) => r.text.trim());
                const sizePx = run?.sizePt
                  ? ptToPx(run.sizePt, pxPerEmu)
                  : ptToPx(10, pxPerEmu);
                return (
                  <td
                    key={ci}
                    colSpan={cell.gridSpan > 1 ? cell.gridSpan : undefined}
                    rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                    style={{
                      backgroundColor: cell.fill ? `#${cell.fill}` : undefined,
                      borderTop: cell.borders.top ? "1px solid #000" : "none",
                      borderBottom: cell.borders.bottom ? "1px solid #000" : "none",
                      borderLeft: cell.borders.left ? "1px solid #000" : "none",
                      borderRight: cell.borders.right ? "1px solid #000" : "none",
                      fontSize: sizePx,
                      fontWeight: run?.bold ? 700 : 400,
                      fontStyle: run?.italic ? "italic" : "normal",
                      color: run?.color ? `#${run.color}` : "#1a1a1a",
                      textAlign:
                        (cell.align === "ctr"
                          ? "center"
                          : cell.align === "r"
                            ? "right"
                            : "left") as React.CSSProperties["textAlign"],
                      padding: "1px 3px",
                      fontFamily: cssFontFamily(run?.font ?? "Calibri"),
                      overflow: "hidden",
                      whiteSpace: "nowrap",
                      lineHeight: 1.15,
                    }}
                  >
                    {cell.text}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
