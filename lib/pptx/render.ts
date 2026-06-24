import { Slide, Shape, Rect, TableModel } from "./types";

/** Strip the heavy notes field and keep only render-relevant slide data. */
export function toRenderSlide(slide: Slide) {
  return {
    index: slide.index,
    kind: slide.kind,
    size: slide.size,
    shapes: slide.shapes,
  };
}

/** Compute the bounding rect of a specific table cell (EMU). */
export function cellRect(
  tableRect: Rect | null,
  table: TableModel,
  row: number,
  col: number
): Rect | null {
  if (!tableRect) return null;
  const colW = table.colWidthsEmu;
  const rowH = table.rowHeightsEmu;
  if (col >= colW.length || row >= rowH.length) return null;
  let x = tableRect.x;
  for (let c = 0; c < col; c++) x += colW[c] || 0;
  let y = tableRect.y;
  for (let r = 0; r < row; r++) y += rowH[r] || 0;
  return { x, y, cx: colW[col] || 0, cy: rowH[row] || 0 };
}

/** Whether a shape carries any visible content (used to skip empty overlays). */
export function shapeHasContent(s: Shape): boolean {
  return (
    s.text.trim().length > 0 ||
    s.kind === "table" ||
    s.kind === "picture" ||
    s.kind === "chart" ||
    s.fill != null
  );
}
