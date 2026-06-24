import JSZip from "jszip";
import {
  XmlNode,
  parseXml,
  root,
  child,
  children,
  childrenNamed,
  attr,
  find,
  findAll,
  allText,
  nodeName,
} from "./xml";
import {
  ThemeFonts,
  readThemeFonts,
  resolveRunStyle,
  extractProps,
  lstStyleOf,
  masterCategoryFor,
  levelDefRPr,
} from "./resolve";
import {
  DeckModel,
  Slide,
  Shape,
  Paragraph,
  TextRun,
  TableModel,
  TableCell,
  Rect,
  MasterPlaceholder,
  ShapeKind,
  SlideKind,
} from "./types";
import { EMU_PER_INCH } from "../brand/guidelines";

const A = (n: string) => `a:${n}`;
const P = (n: string) => `p:${n}`;

interface PlaceholderStyle {
  type: string | null;
  idx: string | null;
  rect: Rect | null;
  lstStyle: XmlNode | undefined;
}

interface ParsedMaster {
  placeholders: PlaceholderStyle[];
  categoryStyles: { title?: XmlNode; body?: XmlNode; other?: XmlNode };
}

interface ParsedLayout {
  name: string;
  type: string | null;
  placeholders: PlaceholderStyle[];
}

// ---------------------------------------------------------------------------
// Rels handling
// ---------------------------------------------------------------------------

interface Rel {
  id: string;
  type: string;
  target: string;
}

function parseRels(xml: string): Record<string, Rel> {
  const out: Record<string, Rel> = {};
  const doc = parseXml(xml);
  const rels = root(doc, "Relationships");
  if (!rels) return out;
  for (const r of childrenNamed(rels, "Relationship")) {
    const id = attr(r, "Id");
    const type = attr(r, "Type") || "";
    const target = attr(r, "Target") || "";
    if (id) out[id] = { id, type, target };
  }
  return out;
}

/** Resolve a relationship target relative to the part's directory. */
function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir + "/" + target).split("/");
  const stack: string[] = [];
  for (const p of parts) {
    if (p === "." || p === "") continue;
    if (p === "..") stack.pop();
    else stack.push(p);
  }
  return stack.join("/");
}

function dirOf(path: string): string {
  const i = path.lastIndexOf("/");
  return i === -1 ? "" : path.slice(0, i);
}

function relsPathFor(partPath: string): string {
  const dir = dirOf(partPath);
  const file = partPath.slice(dir.length + 1);
  return `${dir}/_rels/${file}.rels`;
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

function readXfrm(spPr: XmlNode | undefined): Rect | null {
  if (!spPr) return null;
  const xfrm = child(spPr, A("xfrm")) || child(spPr, P("xfrm"));
  if (!xfrm) return null;
  const off = child(xfrm, A("off"));
  const ext = child(xfrm, A("ext"));
  if (!off || !ext) return null;
  const x = parseInt(attr(off, "x") || "0", 10);
  const y = parseInt(attr(off, "y") || "0", 10);
  const cx = parseInt(attr(ext, "cx") || "0", 10);
  const cy = parseInt(attr(ext, "cy") || "0", 10);
  return { x, y, cx, cy };
}

type Transform = (r: Rect) => Rect;
const IDENTITY: Transform = (r) => r;

function groupTransform(parent: Transform, grpSpPr: XmlNode | undefined): Transform {
  if (!grpSpPr) return parent;
  const xfrm = child(grpSpPr, A("xfrm"));
  if (!xfrm) return parent;
  const off = child(xfrm, A("off"));
  const ext = child(xfrm, A("ext"));
  const chOff = child(xfrm, A("chOff"));
  const chExt = child(xfrm, A("chExt"));
  if (!off || !ext || !chOff || !chExt) return parent;
  const ox = parseInt(attr(off, "x") || "0", 10);
  const oy = parseInt(attr(off, "y") || "0", 10);
  const ecx = parseInt(attr(ext, "cx") || "0", 10);
  const ecy = parseInt(attr(ext, "cy") || "0", 10);
  const cox = parseInt(attr(chOff, "x") || "0", 10);
  const coy = parseInt(attr(chOff, "y") || "0", 10);
  const cecx = parseInt(attr(chExt, "cx") || "1", 10) || 1;
  const cecy = parseInt(attr(chExt, "cy") || "1", 10) || 1;
  const sx = ecx / cecx;
  const sy = ecy / cecy;
  return (r: Rect) =>
    parent({
      x: ox + (r.x - cox) * sx,
      y: oy + (r.y - coy) * sy,
      cx: r.cx * sx,
      cy: r.cy * sy,
    });
}

// ---------------------------------------------------------------------------
// Text / paragraph parsing
// ---------------------------------------------------------------------------

function emuToInches(emu: number): number {
  return emu / EMU_PER_INCH;
}

function resolveParagraph(
  pNode: XmlNode,
  styleChain: (XmlNode | undefined)[],
  theme: ThemeFonts
): Paragraph {
  const pPr = child(pNode, A("pPr"));
  const level = parseInt(attr(pPr, "lvl") || "0", 10);

  const runs: TextRun[] = [];
  let text = "";
  for (const c of children(pNode)) {
    const nm = nodeName(c);
    if (nm === A("r")) {
      const rPr = child(c, A("rPr"));
      const tNode = child(c, A("t"));
      const runText = tNode ? allText(tNode) : "";
      const { style, explicit } = resolveRunStyle(rPr, styleChain, level, theme);
      runs.push({ ...style, text: runText, explicit });
      text += runText;
    } else if (nm === A("br")) {
      text += "\n";
      runs.push({
        text: "\n",
        sizePt: null,
        bold: null,
        italic: null,
        underline: null,
        strike: null,
        font: null,
        color: null,
        explicit: { size: false, bold: false, font: false, color: false },
      });
    } else if (nm === A("fld")) {
      const tNode = child(c, A("t"));
      const runText = tNode ? allText(tNode) : "";
      const rPr = child(c, A("rPr"));
      const { style, explicit } = resolveRunStyle(rPr, styleChain, level, theme);
      runs.push({ ...style, text: runText, explicit });
      text += runText;
    }
  }

  // Resolve paragraph-level properties via cascade.
  const para: Paragraph = {
    level,
    runs,
    text,
    bulletChar: null,
    bulletType: null,
    align: attr(pPr, "algn") || null,
    indentIn: null,
    spaceBeforePt: null,
    spaceAfterPt: null,
    lineSpacingPct: null,
  };

  const marL = attr(pPr, "marL");
  if (marL != null) para.indentIn = emuToInches(parseInt(marL, 10));

  applyParaSpacing(pPr, para);
  applyBullet(pPr, para);

  // Cascade for any unset paragraph props from level pPr in the style chain.
  for (const ls of styleChain) {
    const lvlPPr = levelPPr(ls, level);
    if (!lvlPPr) continue;
    if (para.align == null) para.align = attr(lvlPPr, "algn") || null;
    if (para.indentIn == null) {
      const m = attr(lvlPPr, "marL");
      if (m != null) para.indentIn = emuToInches(parseInt(m, 10));
    }
    if (para.spaceBeforePt == null || para.spaceAfterPt == null || para.lineSpacingPct == null) {
      applyParaSpacing(lvlPPr, para);
    }
    if (para.bulletType == null) applyBullet(lvlPPr, para);
  }

  return para;
}

function levelPPr(lstStyle: XmlNode | undefined, level: number): XmlNode | undefined {
  if (!lstStyle) return undefined;
  const tag = level === 0 ? A("lvl1pPr") : A(`lvl${level + 1}pPr`);
  return child(lstStyle, tag) || (level === 0 ? child(lstStyle, A("defPPr")) : undefined);
}

function applyParaSpacing(pPr: XmlNode | undefined, para: Paragraph) {
  if (!pPr) return;
  if (para.spaceBeforePt == null) {
    const bef = child(pPr, A("spcBef"));
    const pts = bef && child(bef, A("spcPts"));
    if (pts) para.spaceBeforePt = parseInt(attr(pts, "val") || "0", 10) / 100;
  }
  if (para.spaceAfterPt == null) {
    const aft = child(pPr, A("spcAft"));
    const pts = aft && child(aft, A("spcPts"));
    if (pts) para.spaceAfterPt = parseInt(attr(pts, "val") || "0", 10) / 100;
  }
  if (para.lineSpacingPct == null) {
    const ln = child(pPr, A("lnSpc"));
    const pct = ln && child(ln, A("spcPct"));
    if (pct) para.lineSpacingPct = parseInt(attr(pct, "val") || "0", 10) / 1000;
  }
}

function applyBullet(pPr: XmlNode | undefined, para: Paragraph) {
  if (!pPr || para.bulletType != null) return;
  if (child(pPr, A("buNone"))) {
    para.bulletType = "none";
    para.bulletChar = null;
  } else if (child(pPr, A("buChar"))) {
    para.bulletType = "char";
    para.bulletChar = attr(child(pPr, A("buChar")), "char") || null;
  } else if (child(pPr, A("buAutoNum"))) {
    para.bulletType = "auto";
  }
}

// ---------------------------------------------------------------------------
// Shape parsing
// ---------------------------------------------------------------------------

function placeholderInfo(sp: XmlNode): { type: string | null; idx: string | null } {
  const nvPr = find(sp, P("nvPr"));
  const ph = nvPr && child(nvPr, P("ph"));
  if (!ph) return { type: null, idx: null };
  return { type: attr(ph, "type") || null, idx: attr(ph, "idx") || null };
}

function shapeName(sp: XmlNode): { id: string; name: string } {
  const cNvPr = find(sp, P("cNvPr"));
  return {
    id: (cNvPr && attr(cNvPr, "id")) || "",
    name: (cNvPr && attr(cNvPr, "name")) || "",
  };
}

function solidFillColor(spPr: XmlNode | undefined): string | null {
  if (!spPr) return null;
  const fill = child(spPr, A("solidFill"));
  if (!fill) return null;
  const srgb = child(fill, A("srgbClr"));
  if (srgb) return (attr(srgb, "val") || "").toUpperCase() || null;
  return null;
}

function buildStyleChain(
  txBody: XmlNode | undefined,
  layoutPh: PlaceholderStyle | undefined,
  masterPh: PlaceholderStyle | undefined,
  masterCategory: XmlNode | undefined,
  presentationDefault: XmlNode | undefined
): (XmlNode | undefined)[] {
  return [
    lstStyleOf(txBody),
    layoutPh?.lstStyle,
    masterPh?.lstStyle,
    masterCategory,
    presentationDefault,
  ];
}

interface SlideContext {
  theme: ThemeFonts;
  master: ParsedMaster;
  layout: ParsedLayout | undefined;
  presentationDefault: XmlNode | undefined;
  slideRels: Record<string, Rel>;
  slidePath: string;
  zip: JSZip;
  mediaCache: Map<string, string | null>;
}

async function parseSpShape(
  sp: XmlNode,
  ctx: SlideContext,
  transform: Transform
): Promise<Shape> {
  const { type: phType, idx: phIdx } = placeholderInfo(sp);
  const { id, name } = shapeName(sp);
  const spPr = child(sp, P("spPr"));
  const localRect = readXfrm(spPr);
  const rect = localRect ? transform(localRect) : matchPlaceholderRect(phType, phIdx, ctx);

  const layoutPh = findPlaceholder(ctx.layout?.placeholders, phType, phIdx);
  const masterPh = findPlaceholder(ctx.master.placeholders, phType, phIdx);
  const category = ctx.master.categoryStyles[masterCategoryFor(phType)];

  const txBody = child(sp, P("txBody"));
  const styleChain = buildStyleChain(
    txBody,
    layoutPh,
    masterPh,
    category,
    ctx.presentationDefault
  );

  const paragraphs: Paragraph[] = [];
  if (txBody) {
    for (const p of childrenNamed(txBody, A("p"))) {
      paragraphs.push(resolveParagraph(p, styleChain, ctx.theme));
    }
  }
  const text = paragraphs.map((p) => p.text).join("\n");

  return {
    id,
    name,
    kind: "text",
    placeholderType: phType,
    placeholderIdx: phIdx,
    rect,
    paragraphs,
    text,
    table: null,
    imageDataUrl: null,
    fill: solidFillColor(spPr),
  };
}

function matchPlaceholderRect(
  phType: string | null,
  phIdx: string | null,
  ctx: SlideContext
): Rect | null {
  const layoutPh = findPlaceholder(ctx.layout?.placeholders, phType, phIdx);
  if (layoutPh?.rect) return layoutPh.rect;
  const masterPh = findPlaceholder(ctx.master.placeholders, phType, phIdx);
  return masterPh?.rect || null;
}

function findPlaceholder(
  list: PlaceholderStyle[] | undefined,
  type: string | null,
  idx: string | null
): PlaceholderStyle | undefined {
  if (!list) return undefined;
  // Exact match on type + idx first.
  let m = list.find((p) => p.type === type && p.idx === idx);
  if (m) return m;
  // Then type only.
  m = list.find((p) => p.type === type);
  if (m) return m;
  // Then idx only.
  if (idx != null) m = list.find((p) => p.idx === idx);
  return m;
}

async function parseTableShape(
  gf: XmlNode,
  ctx: SlideContext,
  transform: Transform,
  tbl: XmlNode
): Promise<Shape> {
  const { id, name } = shapeName(gf);
  const xfrm = child(gf, P("xfrm"));
  const localRect = xfrm ? readXfrmFromNode(xfrm) : null;
  const rect = localRect ? transform(localRect) : null;

  const tblGrid = child(tbl, A("tblGrid"));
  const colWidthsEmu = tblGrid
    ? childrenNamed(tblGrid, A("gridCol")).map((g) => parseInt(attr(g, "w") || "0", 10))
    : [];

  const rows: TableCell[][] = [];
  const rowHeightsEmu: number[] = [];
  const category = ctx.master.categoryStyles.other;

  for (const tr of childrenNamed(tbl, A("tr"))) {
    rowHeightsEmu.push(parseInt(attr(tr, "h") || "0", 10));
    const cells: TableCell[] = [];
    for (const tc of childrenNamed(tr, A("tc"))) {
      const gridSpan = parseInt(attr(tc, "gridSpan") || "1", 10);
      const rowSpan = parseInt(attr(tc, "rowSpan") || "1", 10);
      const hMerge = attr(tc, "hMerge") === "1";
      const vMerge = attr(tc, "vMerge") === "1";
      const tcPr = child(tc, A("tcPr"));
      const txBody = child(tc, A("txBody"));
      const styleChain = [lstStyleOf(txBody), category, ctx.presentationDefault];
      const runs: TextRun[] = [];
      let cellText = "";
      if (txBody) {
        for (const p of childrenNamed(txBody, A("p"))) {
          const para = resolveParagraph(p, styleChain, ctx.theme);
          runs.push(...para.runs.filter((r) => r.text !== "\n"));
          cellText += (cellText ? "\n" : "") + para.text;
        }
      }
      cells.push({
        text: cellText,
        runs,
        rowSpan,
        gridSpan,
        merged: hMerge || vMerge,
        hMerge,
        vMerge,
        fill: cellFill(tcPr),
        align: cellAlign(txBody),
        borders: cellBorders(tcPr),
      });
    }
    rows.push(cells);
  }

  const table: TableModel = { rows, colWidthsEmu, rowHeightsEmu };

  return {
    id,
    name,
    kind: "table",
    placeholderType: null,
    placeholderIdx: null,
    rect,
    paragraphs: [],
    text: rows.map((r) => r.map((c) => c.text).join("\t")).join("\n"),
    table,
    imageDataUrl: null,
    fill: null,
  };
}

function readXfrmFromNode(xfrm: XmlNode): Rect | null {
  const off = child(xfrm, A("off"));
  const ext = child(xfrm, A("ext"));
  if (!off || !ext) return null;
  return {
    x: parseInt(attr(off, "x") || "0", 10),
    y: parseInt(attr(off, "y") || "0", 10),
    cx: parseInt(attr(ext, "cx") || "0", 10),
    cy: parseInt(attr(ext, "cy") || "0", 10),
  };
}

function cellFill(tcPr: XmlNode | undefined): string | null {
  if (!tcPr) return null;
  const fill = child(tcPr, A("solidFill"));
  if (!fill) return null;
  const srgb = child(fill, A("srgbClr"));
  if (srgb) return (attr(srgb, "val") || "").toUpperCase() || null;
  return null;
}

function cellAlign(txBody: XmlNode | undefined): string | null {
  if (!txBody) return null;
  const p = child(txBody, A("p"));
  const pPr = p && child(p, A("pPr"));
  return (pPr && attr(pPr, "algn")) || null;
}

function hasBorder(tcPr: XmlNode | undefined, tag: string): boolean {
  if (!tcPr) return false;
  const ln = child(tcPr, A(tag));
  if (!ln) return false;
  if (child(ln, A("noFill"))) return false;
  // A line with a solidFill (or default) counts as a visible border.
  return true;
}

function cellBorders(tcPr: XmlNode | undefined) {
  return {
    left: hasBorder(tcPr, "lnL"),
    right: hasBorder(tcPr, "lnR"),
    top: hasBorder(tcPr, "lnT"),
    bottom: hasBorder(tcPr, "lnB"),
  };
}

async function mediaDataUrl(
  embedId: string,
  ctx: SlideContext
): Promise<string | null> {
  const rel = ctx.slideRels[embedId];
  if (!rel) return null;
  const path = resolvePath(dirOf(ctx.slidePath), rel.target);
  if (ctx.mediaCache.has(path)) return ctx.mediaCache.get(path)!;
  const file = ctx.zip.file(path);
  if (!file) {
    ctx.mediaCache.set(path, null);
    return null;
  }
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  const mime =
    ext === "png"
      ? "image/png"
      : ext === "jpg" || ext === "jpeg"
        ? "image/jpeg"
        : ext === "gif"
          ? "image/gif"
          : ext === "svg"
            ? "image/svg+xml"
            : ext === "emf" || ext === "wmf"
              ? null
              : "application/octet-stream";
  if (!mime) {
    ctx.mediaCache.set(path, null);
    return null;
  }
  const b64 = await file.async("base64");
  const url = `data:${mime};base64,${b64}`;
  ctx.mediaCache.set(path, url);
  return url;
}

async function parsePicShape(
  pic: XmlNode,
  ctx: SlideContext,
  transform: Transform
): Promise<Shape> {
  const { id, name } = shapeName(pic);
  const spPr = child(pic, P("spPr"));
  const localRect = readXfrm(spPr);
  const rect = localRect ? transform(localRect) : null;
  const blip = find(pic, A("blip"));
  const embed = blip && attr(blip, "r:embed");
  let imageDataUrl: string | null = null;
  if (embed) imageDataUrl = await mediaDataUrl(embed, ctx);
  return {
    id,
    name,
    kind: "picture",
    placeholderType: null,
    placeholderIdx: null,
    rect,
    paragraphs: [],
    text: "",
    table: null,
    imageDataUrl,
    fill: null,
  };
}

function parseChartShape(gf: XmlNode, transform: Transform): Shape {
  const { id, name } = shapeName(gf);
  const xfrm = child(gf, P("xfrm"));
  const localRect = xfrm ? readXfrmFromNode(xfrm) : null;
  const rect = localRect ? transform(localRect) : null;
  return {
    id,
    name,
    kind: "chart",
    placeholderType: null,
    placeholderIdx: null,
    rect,
    paragraphs: [],
    text: "",
    table: null,
    imageDataUrl: null,
    fill: null,
  };
}

async function walkShapes(
  container: XmlNode,
  ctx: SlideContext,
  transform: Transform,
  out: Shape[]
): Promise<void> {
  for (const node of children(container)) {
    const nm = nodeName(node);
    if (nm === P("sp")) {
      out.push(await parseSpShape(node, ctx, transform));
    } else if (nm === P("pic")) {
      out.push(await parsePicShape(node, ctx, transform));
    } else if (nm === P("graphicFrame")) {
      const tbl = find(node, A("tbl"));
      if (tbl) {
        out.push(await parseTableShape(node, ctx, transform, tbl));
      } else {
        const gd = find(node, A("graphicData"));
        const uri = gd ? attr(gd, "uri") || "" : "";
        if (uri.includes("chart")) {
          out.push(parseChartShape(node, transform));
        } else {
          out.push(parseChartShape(node, transform)); // treat unknown graphic as chart-like visual
        }
      }
    } else if (nm === P("grpSp")) {
      const grpSpPr = child(node, P("grpSpPr"));
      const childTransform = groupTransform(transform, grpSpPr);
      await walkShapes(node, ctx, childTransform, out);
    }
  }
}

// ---------------------------------------------------------------------------
// Master / layout parsing
// ---------------------------------------------------------------------------

function parsePlaceholders(spTree: XmlNode): PlaceholderStyle[] {
  const out: PlaceholderStyle[] = [];
  for (const sp of findAll(spTree, P("sp"))) {
    const { type, idx } = placeholderInfo(sp);
    if (type == null && idx == null) continue;
    const spPr = child(sp, P("spPr"));
    const txBody = child(sp, P("txBody"));
    out.push({
      type,
      idx,
      rect: readXfrm(spPr),
      lstStyle: lstStyleOf(txBody),
    });
  }
  return out;
}

function parseMaster(xml: string): ParsedMaster {
  const doc = parseXml(xml);
  const masterRoot = root(doc, P("sldMaster"));
  const cSld = masterRoot && child(masterRoot, P("cSld"));
  const spTree = cSld && child(cSld, P("spTree"));
  const placeholders = spTree ? parsePlaceholders(spTree) : [];
  const txStyles = masterRoot && child(masterRoot, P("txStyles"));
  const categoryStyles = {
    title: txStyles ? child(txStyles, P("titleStyle")) : undefined,
    body: txStyles ? child(txStyles, P("bodyStyle")) : undefined,
    other: txStyles ? child(txStyles, P("otherStyle")) : undefined,
  };
  return { placeholders, categoryStyles };
}

function parseLayout(xml: string, name: string): ParsedLayout {
  const doc = parseXml(xml);
  const layoutRoot = root(doc, P("sldLayout"));
  const type = layoutRoot ? attr(layoutRoot, "type") || null : null;
  const cSld = layoutRoot && child(layoutRoot, P("cSld"));
  const spTree = cSld && child(cSld, P("spTree"));
  const placeholders = spTree ? parsePlaceholders(spTree) : [];
  return { name, type, placeholders };
}

function layoutKind(type: string | null): SlideKind {
  if (type === "title") return "title";
  if (type === "secHead") return "section";
  return "content";
}

// ---------------------------------------------------------------------------
// Top-level parse
// ---------------------------------------------------------------------------

export async function parseDeck(data: ArrayBuffer | Buffer | Uint8Array): Promise<DeckModel> {
  const zip = await JSZip.loadAsync(data);

  const readText = async (path: string): Promise<string | null> => {
    const f = zip.file(path);
    return f ? f.async("string") : null;
  };

  // Presentation + size.
  const presXml = (await readText("ppt/presentation.xml")) || "";
  const presDoc = parseXml(presXml);
  const presRoot = root(presDoc, P("presentation"));
  const sldSz = presRoot && child(presRoot, P("sldSz"));
  const size = {
    cx: sldSz ? parseInt(attr(sldSz, "cx") || "9144000", 10) : 9144000,
    cy: sldSz ? parseInt(attr(sldSz, "cy") || "6858000", 10) : 6858000,
  };
  const presentationDefault = presRoot && child(presRoot, P("defaultTextStyle"));

  const presRels = parseRels((await readText("ppt/_rels/presentation.xml.rels")) || "");

  // Ordered slide paths from sldIdLst.
  const sldIdLst = presRoot && child(presRoot, P("sldIdLst"));
  const slidePaths: string[] = [];
  if (sldIdLst) {
    for (const sldId of childrenNamed(sldIdLst, P("sldId"))) {
      const rid = attr(sldId, "r:id");
      if (rid && presRels[rid]) {
        slidePaths.push(resolvePath("ppt", presRels[rid].target));
      }
    }
  }

  // Parse the first master + theme (decks here use a single master).
  const masterRels = parseRels((await readText("ppt/slideMasters/_rels/slideMaster1.xml.rels")) || "");
  const masterXml = (await readText("ppt/slideMasters/slideMaster1.xml")) || "";
  const master = parseMaster(masterXml);

  let themeFonts: ThemeFonts = { majorFont: "Calibri", minorFont: "Calibri" };
  for (const rel of Object.values(masterRels)) {
    if (rel.type.endsWith("/theme")) {
      const themePath = resolvePath("ppt/slideMasters", rel.target);
      const themeXml = await readText(themePath);
      if (themeXml) {
        const themeDoc = parseXml(themeXml);
        themeFonts = readThemeFonts(root(themeDoc, "a:theme"));
      }
      break;
    }
  }

  const masterPlaceholders: MasterPlaceholder[] = master.placeholders.map((p) => ({
    type: p.type,
    idx: p.idx,
    rect: p.rect,
  }));

  // Layout cache.
  const layoutCache = new Map<string, ParsedLayout>();
  const layoutPlaceholders: Record<string, MasterPlaceholder[]> = {};

  const mediaCache = new Map<string, string | null>();
  const slides: Slide[] = [];

  for (let i = 0; i < slidePaths.length; i++) {
    const slidePath = slidePaths[i];
    const slideXml = await readText(slidePath);
    if (!slideXml) continue;
    const slideRels = parseRels((await readText(relsPathFor(slidePath))) || "");

    // Resolve layout for this slide.
    let layout: ParsedLayout | undefined;
    for (const rel of Object.values(slideRels)) {
      if (rel.type.endsWith("/slideLayout")) {
        const layoutPath = resolvePath(dirOf(slidePath), rel.target);
        if (!layoutCache.has(layoutPath)) {
          const layoutXml = await readText(layoutPath);
          const parsed = parseLayout(layoutXml || "", layoutPath.split("/").pop() || layoutPath);
          layoutCache.set(layoutPath, parsed);
          layoutPlaceholders[parsed.name] = parsed.placeholders.map((p) => ({
            type: p.type,
            idx: p.idx,
            rect: p.rect,
          }));
        }
        layout = layoutCache.get(layoutPath);
        break;
      }
    }

    const ctx: SlideContext = {
      theme: themeFonts,
      master,
      layout,
      presentationDefault,
      slideRels,
      slidePath,
      zip,
      mediaCache,
    };

    const slideDoc = parseXml(slideXml);
    const slideRoot = root(slideDoc, P("sld"));
    const cSld = slideRoot && child(slideRoot, P("cSld"));
    const spTree = cSld && child(cSld, P("spTree"));

    const shapes: Shape[] = [];
    if (spTree) await walkShapes(spTree, ctx, IDENTITY, shapes);

    const titleShape =
      shapes.find(
        (s) => s.placeholderType === "title" || s.placeholderType === "ctrTitle"
      ) || null;

    const hasCtrTitle = shapes.some((s) => s.placeholderType === "ctrTitle");
    const slideKind: SlideKind = hasCtrTitle ? "title" : layoutKind(layout?.type || null);

    // Notes (eval only).
    let notes = "";
    for (const rel of Object.values(slideRels)) {
      if (rel.type.endsWith("/notesSlide")) {
        const notesPath = resolvePath(dirOf(slidePath), rel.target);
        const notesXml = await readText(notesPath);
        if (notesXml) notes = extractNotesText(notesXml);
        break;
      }
    }

    slides.push({
      index: i + 1,
      kind: slideKind,
      layoutName: layout?.name || null,
      layoutType: layout?.type || null,
      size,
      shapes,
      titleShape,
      notes,
    });
  }

  return {
    size,
    slides,
    master: { placeholders: masterPlaceholders },
    layoutPlaceholders,
    theme: themeFonts,
  };
}

function extractNotesText(xml: string): string {
  const doc = parseXml(xml);
  const r = root(doc, P("notes"));
  if (!r) return "";
  // Collect text from body placeholders (skip slide-number placeholder).
  const parts: string[] = [];
  for (const sp of findAll(r, P("sp"))) {
    const { type } = placeholderInfo(sp);
    if (type === "sldNum") continue;
    const txBody = child(sp, P("txBody"));
    if (!txBody) continue;
    for (const p of childrenNamed(txBody, A("p"))) {
      const t = allText(p).trim();
      if (t) parts.push(t);
    }
  }
  return parts.join("\n");
}
