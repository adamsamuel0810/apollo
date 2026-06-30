import {
  XmlNode,
  children,
  find,
  findAll,
  attr,
  nodeName,
  allText,
} from "./xml";
import { Rect } from "./types";

const P = (n: string) => `p:${n}`;
const A = (n: string) => `a:${n}`;

function pushChild(parent: XmlNode, childNode: XmlNode) {
  const nm = nodeName(parent);
  if (!nm) return;
  const arr = parent[nm];
  if (Array.isArray(arr)) arr.push(childNode);
}

function maxShapeId(spTree: XmlNode): number {
  let max = 1;
  for (const tag of [P("sp"), P("pic"), P("graphicFrame")]) {
    for (const node of findAll(spTree, tag)) {
      const cNvPr = find(node, P("cNvPr")) || find(node, A("cNvPr"));
      const id = parseInt(attr(cNvPr, "id") || "0", 10);
      if (id > max) max = id;
    }
  }
  return max;
}

function findFooterShape(spTree: XmlNode): XmlNode | undefined {
  for (const sp of findAll(spTree, P("sp"))) {
    const ph = find(sp, P("ph"));
    if (attr(ph, "type") === "ftr") return sp;
  }
  return undefined;
}

function buildFooterShape(id: number, rect: Rect, text: string): XmlNode {
  return {
    "p:sp": [
      {
        "p:nvSpPr": [
          {
            "p:cNvPr": [],
            ":@": { "@_id": String(id), "@_name": `Footer Placeholder ${id}` },
          },
          {
            "p:cNvSpPr": [{ "a:spLocks": [], ":@": { "@_noGrp": "1" } }],
          },
          {
            "p:nvPr": [
              {
                "p:ph": [],
                ":@": { "@_type": "ftr", "@_sz": "quarter", "@_idx": "16" },
              },
            ],
          },
        ],
      },
      {
        "p:spPr": [
          {
            "a:xfrm": [
              {
                "a:off": [],
                ":@": { "@_x": String(rect.x), "@_y": String(rect.y) },
              },
              {
                "a:ext": [],
                ":@": { "@_cx": String(rect.cx), "@_cy": String(rect.cy) },
              },
            ],
          },
        ],
      },
      {
        "p:txBody": [
          { "a:bodyPr": [] },
          { "a:lstStyle": [] },
          {
            "a:p": [
              {
                "a:r": [
                  { "a:rPr": [], ":@": { "@_lang": "en-US", "@_sz": "800" } },
                  { "a:t": [{ "#text": text }] },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

/** Write confidentiality/footer text — creates the ftr shape when the slide has none. */
export function ensureFooterText(
  spTree: XmlNode,
  rect: Rect,
  text: string,
  writeText: (shape: XmlNode, value: string) => { before: string; after: string }
): { before: string; after: string; changed: boolean } {
  const existing = findFooterShape(spTree);
  if (existing) {
    const result = writeText(existing, text);
    return { ...result, changed: result.after !== result.before || result.after === text };
  }

  const newId = maxShapeId(spTree) + 1;
  pushChild(spTree, buildFooterShape(newId, rect, text));
  return { before: "", after: text, changed: true };
}

export function footerShapeText(spTree: XmlNode): string {
  const ftr = findFooterShape(spTree);
  return ftr ? allText(ftr) : "";
}
