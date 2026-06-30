import {
  XmlNode,
  children,
  childrenNamed,
  attr,
  find,
  findAll,
  nodeName,
} from "./xml";
import { Rect } from "./types";

const A = (n: string) => `a:${n}`;
const P = (n: string) => `p:${n}`;

function setAttr(node: XmlNode, name: string, value: string) {
  if (!node[":@"]) node[":@"] = {};
  node[":@"]["@_" + name] = value;
}

function pushChild(parent: XmlNode, child: XmlNode) {
  const nm = nodeName(parent);
  if (!nm) return;
  const arr = parent[nm];
  if (Array.isArray(arr)) arr.push(child);
}

function ensureChild(parent: XmlNode, tag: string): XmlNode {
  let c = find(parent, tag);
  if (c) return c;
  c = { [tag]: [] };
  pushChild(parent, c);
  return c;
}

function szValue(pt: number): string {
  return String(Math.round(pt * 100));
}

export function applyFontFamilySet(node: XmlNode, fontFamily: string): boolean {
  let changed = false;
  for (const rPr of findAll(node, A("rPr"))) {
    for (const tag of [A("latin"), A("cs")]) {
      let el = find(rPr, tag);
      if (!el) {
        el = { [tag]: [] };
        pushChild(rPr, el);
      }
      if (attr(el, "typeface") !== fontFamily) {
        setAttr(el, "typeface", fontFamily);
        changed = true;
      }
    }
  }
  return changed;
}

export function applyFontSizeSet(node: XmlNode, sizePt: number): boolean {
  const sz = szValue(sizePt);
  let changed = false;
  for (const rPr of findAll(node, A("rPr"))) {
    if (attr(rPr, "sz") !== sz) {
      setAttr(rPr, "sz", sz);
      changed = true;
    }
  }
  return changed;
}

export function applyFontBoldSet(node: XmlNode): boolean {
  let changed = false;
  for (const rPr of findAll(node, A("rPr"))) {
    if (attr(rPr, "b") !== "1") {
      setAttr(rPr, "b", "1");
      changed = true;
    }
  }
  return changed;
}

function normalizeColor(hex: string): string {
  return hex.replace(/^#/, "").toUpperCase();
}

function setSrgbFill(parent: XmlNode, color: string): boolean {
  const solid = ensureChild(parent, A("solidFill"));
  let srgb = find(solid, A("srgbClr"));
  if (!srgb) {
    srgb = { [A("srgbClr")]: [] };
    pushChild(solid, srgb);
  }
  const val = normalizeColor(color);
  if (attr(srgb, "val") !== val) {
    setAttr(srgb, "val", val);
    return true;
  }
  return false;
}

function replaceChild(parent: XmlNode, tag: string, replacement: XmlNode) {
  const nm = nodeName(parent);
  if (!nm) return;
  const arr = (parent[nm] as XmlNode[]) || [];
  const next = arr.filter((c) => nodeName(c) !== tag);
  next.push(replacement);
  parent[nm] = next;
}

export function applyColorSet(
  node: XmlNode,
  fromColor: string,
  toColor: string,
  target: "fill" | "text" | "both"
): boolean {
  const from = normalizeColor(fromColor);
  const to = normalizeColor(toColor);
  let changed = false;

  const replaceSrgb = (srgb: XmlNode): boolean => {
    const val = attr(srgb, "val")?.toUpperCase();
    if (val === from) {
      setAttr(srgb, "val", to);
      return true;
    }
    return false;
  };

  if (target === "fill" || target === "both") {
    for (const sf of findAll(node, A("solidFill"))) {
      for (const srgb of findAll(sf, A("srgbClr"))) {
        if (replaceSrgb(srgb)) changed = true;
      }
    }
  }

  if (target === "text" || target === "both") {
    for (const rPr of findAll(node, A("rPr"))) {
      const sf = find(rPr, A("solidFill"));
      if (sf) {
        for (const srgb of findAll(sf, A("srgbClr"))) {
          if (replaceSrgb(srgb)) changed = true;
        }
      }
    }
  }

  return changed;
}

export function applyShapeReposition(node: XmlNode, rect: Rect): boolean {
  const spPr = find(node, P("spPr")) || find(node, A("spPr"));
  if (!spPr) return false;
  const xfrm = ensureChild(spPr, A("xfrm"));
  const off = ensureChild(xfrm, A("off"));
  const ext = ensureChild(xfrm, A("ext"));

  const before = `${attr(off, "x")},${attr(off, "y")},${attr(ext, "cx")},${attr(ext, "cy")}`;
  setAttr(off, "x", String(rect.x));
  setAttr(off, "y", String(rect.y));
  setAttr(ext, "cx", String(rect.cx));
  setAttr(ext, "cy", String(rect.cy));
  const after = `${rect.x},${rect.y},${rect.cx},${rect.cy}`;
  return before !== after;
}

export function applyRowFillSet(shape: XmlNode, rowIndex: number, fillColor: string): boolean {
  const tbl = find(shape, A("tbl"));
  if (!tbl) return false;
  const rows = childrenNamed(tbl, A("tr"));
  const tr = rows[rowIndex];
  if (!tr) return false;
  let changed = false;
  for (const tc of childrenNamed(tr, A("tc"))) {
    const tcPr = ensureChild(tc, A("tcPr"));
    replaceChild(tcPr, A("solidFill"), { [A("solidFill")]: [] });
    if (setSrgbFill(tcPr, fillColor)) changed = true;
  }
  return changed;
}

export function applyTableSideBorderRemove(shape: XmlNode): boolean {
  const tbl = find(shape, A("tbl"));
  if (!tbl) return false;
  let changed = false;
  for (const tr of childrenNamed(tbl, A("tr"))) {
    const cells = childrenNamed(tr, A("tc"));
    if (cells.length === 0) continue;
    const first = cells[0];
    const last = cells[cells.length - 1];
    for (const tc of [first, last]) {
      const tcPr = ensureChild(tc, A("tcPr"));
      const side = tc === first ? A("lnL") : A("lnR");
      replaceChild(tcPr, side, { [side]: [{ [A("noFill")]: [] }] });
      changed = true;
    }
  }
  return changed;
}
