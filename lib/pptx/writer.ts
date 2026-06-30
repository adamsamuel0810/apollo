import JSZip from "jszip";
import {
  XmlNode,
  parseXml,
  buildXml,
  root,
  children,
  childrenNamed,
  attr,
  find,
  findAll,
  allText,
  nodeName,
  isText,
  textValue,
} from "./xml";
import { FixInstruction, FixRecord } from "../rules/types";
import {
  applyFontFamilySet,
  applyFontSizeSet,
  applyFontBoldSet,
  applyColorSet,
  applyShapeReposition,
  applyRowFillSet,
  applyTableSideBorderRemove,
} from "./ooxmlFixes";
import { ensureFooterText } from "./footerShape";

const P = (n: string) => `p:${n}`;
const A = (n: string) => `a:${n}`;

export interface ApplyFixResult {
  buffer: Buffer;
  record: FixRecord | null;
  error?: string;
}

function shapeIdOf(node: XmlNode): string | null {
  const cNvPr =
    find(node, P("cNvPr")) ||
    find(node, A("cNvPr"));
  return cNvPr ? attr(cNvPr, "id") || null : null;
}

/** Walk all shapes including groups and graphic frames. */
function* iterShapes(nodes: XmlNode[]): Generator<XmlNode> {
  for (const n of nodes) {
    const nm = nodeName(n);
    if (
      nm === P("sp") ||
      nm === P("graphicFrame") ||
      nm === P("grpSp") ||
      nm === P("pic")
    ) {
      yield n;
    }
    for (const c of children(n)) yield* iterShapes([c]);
  }
}

function findShapeById(nodes: XmlNode[], shapeId: string): XmlNode | undefined {
  for (const sp of iterShapes(nodes)) {
    const id = shapeIdOf(sp);
    if (id === shapeId) return sp;
  }
  return undefined;
}

function textLeaves(node: XmlNode): XmlNode[] {
  const out: XmlNode[] = [];
  for (const t of findAll(node, A("t"))) {
    const kids = children(t);
    const textChild = kids.find((c) => isText(c));
    if (textChild) out.push(textChild);
    else if (kids.length === 0) out.push(t);
  }
  return out;
}

function setTextOnLeaf(leaf: XmlNode, value: string) {
  if (isText(leaf)) {
    leaf["#text"] = value;
    return;
  }
  const kids = children(leaf);
  const textChild = kids.find((c) => isText(c));
  if (textChild) textChild["#text"] = value;
  else {
    const arr = leaf[A("t")];
    if (Array.isArray(arr)) arr.push({ "#text": value });
    else leaf[A("t")] = [{ "#text": value }];
  }
}

function replaceInSubtree(
  node: XmlNode,
  replacer: (text: string) => string
): { before: string; after: string; changed: boolean } {
  const before = allText(node);
  let changed = false;
  for (const leaf of textLeaves(node)) {
    const old = isText(leaf) ? textValue(leaf) : allText(leaf);
    const next = replacer(old);
    if (next !== old) {
      changed = true;
      setTextOnLeaf(leaf, next);
    }
  }
  return { before, after: allText(node), changed };
}

function setShapePlainText(node: XmlNode, newText: string): { before: string; after: string } {
  const before = allText(node);
  let leaves = textLeaves(node);
  if (leaves.length === 0) {
    writeTextToEmptyShape(node, newText);
    return { before, after: newText };
  }
  setTextOnLeaf(leaves[0], newText);
  for (let i = 1; i < leaves.length; i++) setTextOnLeaf(leaves[i], "");
  return { before, after: newText };
}

/** Insert minimal txBody/a:p/a:r/a:t structure when a shape has no text yet. */
function writeTextToEmptyShape(shape: XmlNode, text: string) {
  let txBody = find(shape, P("txBody"));
  if (!txBody) {
    txBody = { "p:txBody": [] };
    const nm = nodeName(shape);
    if (nm) {
      const arr = shape[nm];
      if (Array.isArray(arr)) arr.push(txBody);
    }
  }
  const txKey = nodeName(txBody)!;
  const txKids = (txBody[txKey] as XmlNode[]) || [];
  if (!txKids.some((c) => nodeName(c) === A("bodyPr"))) {
    txKids.unshift({ "a:bodyPr": [] });
  }
  const filtered = txKids.filter((c) => nodeName(c) !== A("p"));
  filtered.push({
    "a:p": [
      {
        "a:r": [{ "a:rPr": [] }, { "a:t": [{ "#text": text }] }],
      },
    ],
  });
  txBody[txKey] = filtered;
}

function tableCellNode(shape: XmlNode, row: number, col: number): XmlNode | undefined {
  const tbl = find(shape, A("tbl"));
  if (!tbl) return undefined;
  const rows = childrenNamed(tbl, A("tr"));
  const tr = rows[row];
  if (!tr) return undefined;
  const cells = childrenNamed(tr, A("tc"));
  return cells[col];
}

async function loadSlide(zip: JSZip, slideIndex: number): Promise<{ path: string; doc: XmlNode[] } | null> {
  const path = `ppt/slides/slide${slideIndex}.xml`;
  const file = zip.file(path);
  if (!file) return null;
  const xml = await file.async("string");
  return { path, doc: parseXml(xml) };
}

function slidePart(nodes: XmlNode[]): XmlNode | undefined {
  return root(nodes, P("sld"));
}

function makeReplacer(
  search: string,
  replace: string,
  caseInsensitive?: boolean,
  regex?: boolean
) {
  const re = regex
    ? new RegExp(search, caseInsensitive ? "gi" : "g")
    : new RegExp(
        search.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
        caseInsensitive ? "gi" : "g"
      );
  return (t: string) => {
    const next = t.replace(re, replace);
    if (regex) return next.replace(/\s{2,}/g, " ").trim();
    return next;
  };
}

/** Replace text in every slide part of the deck. */
async function applyDeckWideReplace(
  zip: JSZip,
  search: string,
  replace: string,
  caseInsensitive?: boolean,
  regex?: boolean
): Promise<{ changed: boolean; before: string; after: string; touchedSlides: number[] }> {
  const replacer = makeReplacer(search, replace, caseInsensitive, regex);
  const touchedSlides: number[] = [];
  let beforeSample = "";
  let afterSample = "";
  let anyChanged = false;

  const slidePaths = Object.keys(zip.files).filter((p) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(p)
  );

  for (const path of slidePaths) {
    const file = zip.file(path);
    if (!file) continue;
    const doc = parseXml(await file.async("string"));
    const sld = slidePart(doc);
    if (!sld) continue;
    const spTree = find(sld, P("spTree"));
    if (!spTree) continue;
    const result = replaceInSubtree(spTree, replacer);
    if (result.changed) {
      anyChanged = true;
      const idx = parseInt(path.match(/slide(\d+)/)?.[1] || "0", 10);
      if (idx) touchedSlides.push(idx);
      if (!beforeSample) beforeSample = result.before.slice(0, 120);
      afterSample = result.after.slice(0, 120);
      zip.file(path, buildXml(doc));
    }
  }

  return { changed: anyChanged, before: beforeSample, after: afterSample, touchedSlides };
}

async function finalizeZip(zip: JSZip): Promise<Buffer> {
  return Buffer.from(
    await zip.generateAsync({
      type: "uint8array",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    })
  );
}

export async function applyFixToPptx(
  data: Buffer | Uint8Array,
  fix: FixInstruction,
  meta: { findingId: string; ruleId: string; title: string; source: "deterministic" | "ai" }
): Promise<ApplyFixResult> {
  const zip = await JSZip.loadAsync(data);

  // Deck-wide terminology replacements (all slides at once).
  if (
    (fix.kind === "deck-replace" || (fix.kind === "text-replace" && fix.deckWide)) &&
    fix.search != null &&
    fix.replace != null
  ) {
    const deckResult = await applyDeckWideReplace(
      zip,
      fix.search,
      fix.replace,
      fix.caseInsensitive,
      fix.regex
    );
    if (!deckResult.changed) {
      return { buffer: Buffer.from(data), record: null, error: "No change applied" };
    }
    const out = await finalizeZip(zip);
    return {
      buffer: out,
      record: {
        findingId: meta.findingId,
        ruleId: meta.ruleId,
        slideIndex: fix.slideIndex,
        title: meta.title,
        before: deckResult.before,
        after: deckResult.after,
        appliedAt: new Date().toISOString(),
        source: meta.source,
      },
    };
  }

  const loaded = await loadSlide(zip, fix.slideIndex);
  if (!loaded) return { buffer: Buffer.from(data), record: null, error: "Slide not found" };

  const { path, doc } = loaded;
  const sld = slidePart(doc);
  if (!sld) return { buffer: Buffer.from(data), record: null, error: "Invalid slide XML" };

  const spTree = find(sld, P("spTree"));
  if (!spTree) return { buffer: Buffer.from(data), record: null, error: "No shape tree" };

  let before = "";
  let after = "";
  let changed = false;

  if (fix.kind === "ensure-footer" && fix.newText != null && fix.targetRect) {
    const result = ensureFooterText(spTree, fix.targetRect, fix.newText, setShapePlainText);
    before = result.before;
    after = result.after;
    changed = result.changed;
    if (!changed) {
      return { buffer: Buffer.from(data), record: null, error: "No change applied" };
    }
    zip.file(path, buildXml(doc));
    const out = await finalizeZip(zip);
    return {
      buffer: out,
      record: {
        findingId: meta.findingId,
        ruleId: meta.ruleId,
        slideIndex: fix.slideIndex,
        title: meta.title,
        before: before.slice(0, 200),
        after: after.slice(0, 200),
        appliedAt: new Date().toISOString(),
        source: meta.source,
      },
    };
  }

  let target: XmlNode | undefined;
  if (fix.shapeId) target = findShapeById([spTree], fix.shapeId);
  if (!target && fix.shapeId) {
    return { buffer: Buffer.from(data), record: null, error: "Shape not found" };
  }
  if (fix.cell && target) {
    const cell = tableCellNode(target, fix.cell.row, fix.cell.col);
    if (cell) target = cell;
  }

  if (fix.kind === "title-rewrite" && fix.newText != null && target) {
    ({ before, after } = setShapePlainText(target, fix.newText));
    changed = before !== after;
  } else if (fix.kind === "font-family-set" && target) {
    changed = applyFontFamilySet(target, fix.fontFamily || "Calibri");
    before = allText(target).slice(0, 120);
    after = before;
  } else if (fix.kind === "font-size-set" && target) {
    changed = applyFontSizeSet(target, fix.fontSizePt || 24);
    before = allText(target).slice(0, 120);
    after = before;
  } else if (fix.kind === "font-bold-set" && target) {
    changed = applyFontBoldSet(target);
    before = allText(target).slice(0, 120);
    after = before;
  } else if (
    fix.kind === "color-set" &&
    target &&
    fix.fromColor &&
    fix.toColor
  ) {
    changed = applyColorSet(
      target,
      fix.fromColor,
      fix.toColor,
      fix.colorTarget || "both"
    );
    before = fix.fromColor;
    after = fix.toColor;
  } else if (fix.kind === "shape-reposition" && target && fix.targetRect) {
    changed = applyShapeReposition(target, fix.targetRect);
    before = "moved";
    after = "repositioned";
  } else if (
    fix.kind === "row-fill-set" &&
    target &&
    fix.rowIndex != null &&
    fix.fillColor
  ) {
    changed = applyRowFillSet(target, fix.rowIndex, fix.fillColor);
    before = "row fill";
    after = fix.fillColor;
  } else if (fix.kind === "border-remove" && target && fix.borderScope === "table-sides") {
    changed = applyTableSideBorderRemove(target);
    before = "borders";
    after = "removed";
  } else if (fix.kind === "text-replace" && fix.search != null && fix.replace != null) {
    const replacer = makeReplacer(
      fix.search,
      fix.replace,
      fix.caseInsensitive,
      fix.regex
    );
    const scope = target || spTree;
    const result = replaceInSubtree(scope, replacer);
    before = result.before;
    after = result.after;
    changed = result.changed;
  } else if (fix.kind === "ai-rewrite" && fix.newText != null && target) {
    ({ before, after } = setShapePlainText(target, fix.newText));
    changed = before !== after;
  }

  if (!changed) {
    return { buffer: Buffer.from(data), record: null, error: "No change applied" };
  }

  zip.file(path, buildXml(doc));
  const out = await finalizeZip(zip);

  return {
    buffer: out,
    record: {
      findingId: meta.findingId,
      ruleId: meta.ruleId,
      slideIndex: fix.slideIndex,
      title: meta.title,
      before: before.slice(0, 200),
      after: after.slice(0, 200),
      appliedAt: new Date().toISOString(),
      source: meta.source,
    },
  };
}

/** Apply multiple fixes sequentially; skips duplicate deck-wide ruleIds. */
export async function applyFixesToPptx(
  data: Buffer | Uint8Array,
  items: { fix: FixInstruction; findingId: string; ruleId: string; title: string; source: "deterministic" | "ai" }[]
): Promise<{ buffer: Buffer; records: FixRecord[]; errors: string[] }> {
  let buf = Buffer.from(data);
  const records: FixRecord[] = [];
  const errors: string[] = [];
  const deckWideDone = new Set<string>();

  for (const item of items) {
    const isDeck =
      item.fix.kind === "deck-replace" ||
      (item.fix.kind === "text-replace" && item.fix.deckWide);
    if (isDeck) {
      const key = `${item.ruleId}:${item.fix.search}:${item.fix.replace}`;
      if (deckWideDone.has(key)) continue;
      deckWideDone.add(key);
    }

    const res = await applyFixToPptx(buf, item.fix, {
      findingId: item.findingId,
      ruleId: item.ruleId,
      title: item.title,
      source: item.source,
    });
    if (res.record) {
      records.push(res.record);
      buf = Buffer.from(res.buffer);
    } else if (res.error) {
      errors.push(`${item.ruleId} (slide ${item.fix.slideIndex}): ${res.error}`);
    }
  }
  return { buffer: buf, records, errors };
}
