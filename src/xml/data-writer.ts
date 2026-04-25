// ── XML Data Writer ──────────────────────────────────────────────────
// Serialize an array of row objects to XML. Keys starting with `attrPrefix`
// (default `@`) are emitted as XML attributes; everything else becomes a
// child element. Nested dot-paths (e.g. `Pricing.Cost`) are reconstructed
// into a tree.

import type { CellValue } from "../_types";
import { ParseError } from "../errors";

export interface XmlWriteOptions {
  /** Root element tag. Default: "root". */
  rootTag?: string;
  /** Per-row element tag. Default: "row". */
  rowTag?: string;
  /** Prefix marking a key as an XML attribute. Default: "@". */
  attrPrefix?: string;
  /** Mixed-content text key. Default: "#text". */
  textKey?: string;
  /** Emit `<?xml version="1.0" encoding="UTF-8"?>` declaration. Default: true. */
  declaration?: boolean;
  /** Pretty-print with indentation. Default: false. */
  pretty?: boolean;
  /** Indent string when `pretty` is true. Default: "  ". */
  indent?: string;
}

const VALID_NAME_RE = /^[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?$/;

function escapeText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}

function valueToString(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

interface TreeNode {
  attrs: Record<string, string>;
  text?: string;
  children: Map<string, TreeNode>;
}

function makeNode(): TreeNode {
  return { attrs: {}, children: new Map() };
}

/**
 * Serialize an array of flat objects to XML.
 *
 * Keys with the `attrPrefix` (default `@`) become XML attributes on the
 * containing element. Dot-separated keys (e.g. `Pricing.Cost`) reconstruct
 * a nested element tree. The `textKey` (default `#text`) emits text content
 * inside an element that also has attributes or children.
 *
 * Throws {@link ParseError} when a key cannot be serialized as a valid XML
 * element name.
 */
export function writeXml(data: Record<string, CellValue>[], options?: XmlWriteOptions): string {
  const rootTag = options?.rootTag ?? "root";
  const rowTag = options?.rowTag ?? "row";
  const attrPrefix = options?.attrPrefix ?? "@";
  const textKey = options?.textKey ?? "#text";
  const declaration = options?.declaration ?? true;
  const pretty = options?.pretty ?? false;
  const indent = options?.indent ?? "  ";

  validateName(rootTag, "rootTag");
  validateName(rowTag, "rowTag");

  const parts: string[] = [];
  if (declaration) {
    parts.push('<?xml version="1.0" encoding="UTF-8"?>');
    if (pretty) parts.push("\n");
  }

  const rowDepth = 1;
  const sep = pretty ? "\n" : "";
  const pad = (d: number): string => (pretty ? indent.repeat(d) : "");

  parts.push(`<${rootTag}>`);
  parts.push(sep);

  for (const row of data) {
    const tree = buildTree(row, attrPrefix, textKey);
    parts.push(pad(rowDepth));
    parts.push(renderElement(rowTag, tree, pretty, indent, rowDepth));
    parts.push(sep);
  }

  parts.push(`</${rootTag}>`);
  if (pretty) parts.push("\n");
  return parts.join("");
}

function validateName(name: string, label: string): void {
  if (!VALID_NAME_RE.test(name)) {
    throw new ParseError(`Invalid XML name for ${label}: "${name}"`);
  }
}

function buildTree(row: Record<string, CellValue>, attrPrefix: string, textKey: string): TreeNode {
  const root = makeNode();

  for (const [rawKey, rawVal] of Object.entries(row)) {
    if (rawVal === undefined) continue;
    insert(root, rawKey, rawVal, attrPrefix, textKey);
  }

  return root;
}

function insert(
  node: TreeNode,
  key: string,
  value: CellValue,
  attrPrefix: string,
  textKey: string,
): void {
  const path = key.split(".");
  let current = node;

  for (let i = 0; i < path.length; i++) {
    const segment = path[i]!;
    const isLast = i === path.length - 1;

    if (segment.startsWith(attrPrefix)) {
      const attrName = segment.slice(attrPrefix.length);
      validateName(attrName, `attribute "${segment}"`);
      current.attrs[attrName] = valueToString(value);
      return;
    }

    if (segment === textKey) {
      current.text = valueToString(value);
      return;
    }

    validateName(segment, `element "${segment}"`);

    if (isLast) {
      let child = current.children.get(segment);
      if (!child) {
        child = makeNode();
        current.children.set(segment, child);
      }
      child.text = valueToString(value);
      return;
    }

    let child = current.children.get(segment);
    if (!child) {
      child = makeNode();
      current.children.set(segment, child);
    }
    current = child;
  }
}

function renderElement(
  tag: string,
  node: TreeNode,
  pretty: boolean,
  indent: string,
  depth: number,
): string {
  const sep = pretty ? "\n" : "";
  const pad = (d: number): string => (pretty ? indent.repeat(d) : "");

  let attrStr = "";
  for (const [name, val] of Object.entries(node.attrs)) {
    attrStr += ` ${name}="${escapeAttr(val)}"`;
  }

  const hasChildren = node.children.size > 0;
  const text = node.text ?? "";
  const hasText = text !== "";

  if (!hasChildren && !hasText) {
    return `<${tag}${attrStr}/>`;
  }

  if (!hasChildren) {
    return `<${tag}${attrStr}>${escapeText(text)}</${tag}>`;
  }

  const inner: string[] = [];
  for (const [childTag, childNode] of node.children) {
    inner.push(pad(depth + 1));
    inner.push(renderElement(childTag, childNode, pretty, indent, depth + 1));
    inner.push(sep);
  }

  if (hasText) {
    inner.push(pad(depth + 1));
    inner.push(escapeText(text));
    inner.push(sep);
  }

  return `<${tag}${attrStr}>${sep}${inner.join("")}${pad(depth)}</${tag}>`;
}
