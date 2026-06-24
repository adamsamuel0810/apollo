import { XMLParser } from "fast-xml-parser";

// We parse in preserveOrder mode so sibling ordering across different element
// types (e.g. <a:r> interleaved with <a:br>) is retained. Each node looks like:
//   { "a:p": [ ...childNodes... ], ":@": { "@_lvl": "1" } }
// Text nodes look like { "#text": "value" }.

export type XmlNode = Record<string, unknown> & { ":@"?: Record<string, string> };

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  // Keep namespace prefixes (a:, p:, r:) intact.
  removeNSPrefix: false,
});

export function parseXml(xml: string): XmlNode[] {
  return parser.parse(xml) as XmlNode[];
}

/** The element tag name of a node, or null for text nodes. */
export function nodeName(node: XmlNode): string | null {
  for (const key of Object.keys(node)) {
    if (key === ":@" || key === "#text") continue;
    return key;
  }
  return null;
}

/** Children array of an element node. */
export function children(node: XmlNode): XmlNode[] {
  const name = nodeName(node);
  if (!name) return [];
  const val = node[name];
  return Array.isArray(val) ? (val as XmlNode[]) : [];
}

export function attr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  const a = node[":@"];
  if (!a) return undefined;
  return a["@_" + name];
}

export function isText(node: XmlNode): boolean {
  return Object.prototype.hasOwnProperty.call(node, "#text");
}

export function textValue(node: XmlNode): string {
  const t = node["#text"];
  return t == null ? "" : String(t);
}

/** First direct child with the given tag name. */
export function child(node: XmlNode, name: string): XmlNode | undefined {
  return children(node).find((c) => nodeName(c) === name);
}

/** All direct children with the given tag name. */
export function childrenNamed(node: XmlNode, name: string): XmlNode[] {
  return children(node).filter((c) => nodeName(c) === name);
}

/** Depth-first search for the first descendant (or self) with the tag name. */
export function find(node: XmlNode, name: string): XmlNode | undefined {
  if (nodeName(node) === name) return node;
  for (const c of children(node)) {
    const r = find(c, name);
    if (r) return r;
  }
  return undefined;
}

/** Depth-first collection of all descendants with the given tag name. */
export function findAll(node: XmlNode, name: string, acc: XmlNode[] = []): XmlNode[] {
  if (nodeName(node) === name) acc.push(node);
  for (const c of children(node)) findAll(c, name, acc);
  return acc;
}

/** Find the first descendant matching tag name within a list of roots. */
export function findIn(roots: XmlNode[], name: string): XmlNode | undefined {
  for (const r of roots) {
    const f = find(r, name);
    if (f) return f;
  }
  return undefined;
}

/** Concatenate all text within a subtree. */
export function allText(node: XmlNode): string {
  if (isText(node)) return textValue(node);
  let out = "";
  for (const c of children(node)) out += allText(c);
  return out;
}

/** Walk to the document root element (first non-meta node). */
export function root(nodes: XmlNode[], name: string): XmlNode | undefined {
  for (const n of nodes) {
    if (nodeName(n) === name) return n;
  }
  // fall back to first element node
  return nodes.find((n) => nodeName(n) != null);
}
