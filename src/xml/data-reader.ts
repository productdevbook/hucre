// ── XML Data Reader ──────────────────────────────────────────────────
// Read tabular data from XML using SAX so that large product feeds
// (GS1, Trendyol, SAP B1, Logo GO, Netsis) don't pay full-DOM memory cost.

import type { CellValue } from "../_types";
import { ParseError } from "../errors";
import { parseSax } from "./parser";

/**
 * Options for {@link readXml}.
 */
export interface XmlReadOptions {
  /**
   * Local tag name (or `prefix:local`) to treat as a single row.
   * Default: auto-detect the most frequently repeating direct child of root.
   */
  rowTag?: string;
  /** Prefix for attribute keys. Default: "@". */
  attrPrefix?: string;
  /** Flatten nested child elements into dot-path keys. Default: true. */
  flatten?: boolean;
  /** Strip namespace prefixes from tag names. Default: false. */
  stripNamespaces?: boolean;
  /** Key under which mixed-content text is stored. Default: "#text". */
  textKey?: string;
  /** Transform header keys. */
  transformHeader?: (header: string, index: number) => string;
  /** Transform cell values. */
  transformValue?: (value: CellValue, header: string, rowIndex: number) => CellValue;
  /** Maximum number of rows. */
  maxRows?: number;
}

export interface XmlReadResult<T extends Record<string, CellValue> = Record<string, CellValue>> {
  data: T[];
  headers: string[];
  /** The detected (or specified) row tag, after namespace handling. */
  rowTag: string;
}

interface MiniElement {
  tag: string;
  local: string;
  prefix: string;
  attrs: Record<string, string>;
  children: MiniNode[];
  text: string;
}

type MiniNode = MiniElement | { __text: string };

function splitTag(tag: string): { local: string; prefix: string } {
  const colon = tag.indexOf(":");
  if (colon === -1) return { local: tag, prefix: "" };
  return { prefix: tag.slice(0, colon), local: tag.slice(colon + 1) };
}

/**
 * Read XML and return tabular rows. Auto-detects the repeating element when
 * `rowTag` is omitted: it counts the direct children of the root element and
 * picks the most-frequent tag.
 */
export function readXml<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: string,
  options?: XmlReadOptions,
): XmlReadResult<T> {
  if (input.trim() === "") {
    return { data: [], headers: [], rowTag: options?.rowTag ?? "" };
  }

  const stripNs = options?.stripNamespaces ?? false;
  const attrPrefix = options?.attrPrefix ?? "@";
  const flatten = options?.flatten ?? true;
  const textKey = options?.textKey ?? "#text";

  const requestedRowTag = options?.rowTag;
  const rowTag = requestedRowTag ?? detectRowTag(input, stripNs);

  if (!rowTag) {
    return { data: [], headers: [], rowTag: "" };
  }

  const rows = collectRows(input, rowTag, stripNs);

  const flatOpts = { attrPrefix, flatten, textKey, stripNs };
  const limit = options?.maxRows ?? Infinity;

  const flatRows: Record<string, CellValue>[] = [];
  for (const el of rows) {
    if (flatRows.length >= limit) break;
    const flat: Record<string, CellValue> = {};
    elementToFlat(el, flatOpts, "", flat);
    flatRows.push(flat);
  }

  // Header collection: union of keys, first-seen order
  const seen = new Set<string>();
  let headers: string[] = [];
  for (const row of flatRows) {
    for (const key of Object.keys(row)) {
      if (!seen.has(key)) {
        seen.add(key);
        headers.push(key);
      }
    }
  }

  if (options?.transformHeader) {
    const orig = headers;
    headers = orig.map((h, i) => options.transformHeader!(h, i));
    const map = new Map(orig.map((h, i) => [h, headers[i]!]));
    for (let r = 0; r < flatRows.length; r++) {
      const src = flatRows[r]!;
      const next: Record<string, CellValue> = {};
      for (const o of orig) {
        next[map.get(o)!] = src[o] ?? null;
      }
      flatRows[r] = next;
    }
  }

  const data: T[] = [];
  for (let r = 0; r < flatRows.length; r++) {
    const src = flatRows[r]!;
    const obj: Record<string, CellValue> = {};
    for (const h of headers) {
      let val = src[h] ?? null;
      if (options?.transformValue) {
        val = options.transformValue(val, h, r);
      }
      obj[h] = val;
    }
    data.push(obj as T);
  }

  return { data, headers, rowTag };
}

// ── Row tag auto-detection ────────────────────────────────────────────

function detectRowTag(input: string, stripNs: boolean): string {
  const freq = new Map<string, number>();
  const order: string[] = [];
  let depth = 0;

  parseSax(input, {
    onOpenTag(tag) {
      depth++;
      // Only count direct children of the root (depth === 2)
      if (depth === 2) {
        const key = stripNs ? splitTag(tag).local : tag;
        if (!freq.has(key)) order.push(key);
        freq.set(key, (freq.get(key) ?? 0) + 1);
      }
    },
    onCloseTag() {
      depth--;
    },
  });

  if (freq.size === 0) {
    throw new ParseError("XML root has no child elements");
  }

  // Most frequent; ties broken by first-seen
  let best = order[0]!;
  let bestCount = freq.get(best)!;
  for (const tag of order) {
    const count = freq.get(tag)!;
    if (count > bestCount) {
      best = tag;
      bestCount = count;
    }
  }
  return best;
}

// ── Row collection ─────────────────────────────────────────────────────

function collectRows(input: string, rowTag: string, stripNs: boolean): MiniElement[] {
  const rows: MiniElement[] = [];
  // Stack of elements we are currently building. The first entry (when set)
  // is the root row element; nested elements are pushed/popped during traversal.
  const stack: MiniElement[] = [];
  let depth = 0;
  // Suppress the next onText invocation when we just handled a CDATA block —
  // the SAX parser fires both `onCData` and `onText` for the same content.
  let suppressNextText = false;

  const matches = (tag: string): boolean => {
    if (stripNs) return splitTag(tag).local === rowTag;
    return tag === rowTag;
  };

  parseSax(input, {
    onOpenTag(tag, attrs) {
      depth++;
      if (stack.length === 0) {
        // Not currently inside a row — only enter when we see the row tag at depth 2
        if (depth >= 2 && matches(tag)) {
          const { local, prefix } = splitTag(tag);
          stack.push({ tag, local, prefix, attrs, children: [], text: "" });
        }
        return;
      }
      // Inside a row: push child element
      const { local, prefix } = splitTag(tag);
      const el: MiniElement = { tag, local, prefix, attrs, children: [], text: "" };
      stack[stack.length - 1]!.children.push(el);
      stack.push(el);
    },
    onCloseTag() {
      const el = stack[stack.length - 1];
      if (el) stack.pop();
      depth--;
      if (stack.length === 0 && el && matches(el.tag)) {
        rows.push(el);
      }
    },
    onText(text) {
      if (suppressNextText) {
        suppressNextText = false;
        return;
      }
      if (stack.length === 0) return;
      if (text.trim() === "") return;
      stack[stack.length - 1]!.text += text;
    },
    onCData(text) {
      if (stack.length === 0) {
        suppressNextText = true;
        return;
      }
      stack[stack.length - 1]!.text += text;
      suppressNextText = true;
    },
  });

  return rows;
}

// ── Flatten a row element into dot-path keys ───────────────────────────

interface FlattenCtx {
  attrPrefix: string;
  flatten: boolean;
  textKey: string;
  stripNs: boolean;
}

function elementToFlat(
  el: MiniElement,
  ctx: FlattenCtx,
  prefix: string,
  out: Record<string, CellValue>,
): void {
  // Attributes always emit as `<prefix>.<attrPrefix><name>` (or top-level when prefix is empty)
  for (const [name, val] of Object.entries(el.attrs)) {
    const key = prefix ? `${prefix}.${ctx.attrPrefix}${name}` : `${ctx.attrPrefix}${name}`;
    out[key] = val;
  }

  const elementChildren = el.children.filter(isElement);

  if (elementChildren.length === 0) {
    // Leaf: write text content under prefix (or under textKey when prefix is empty
    // and there are no attrs — i.e. for the row root with text-only content)
    const text = el.text.trim();
    if (prefix) {
      out[prefix] = text === "" ? null : text;
    } else if (text !== "") {
      out[ctx.textKey] = text;
    }
    return;
  }

  // When flatten is disabled we still expand the row root into its top-level
  // children — but each child's nested content is JSON-stringified rather
  // than recursed.
  if (!ctx.flatten) {
    if (!prefix) {
      // Row root: keep top-level child names visible, stringify their content
      for (const child of elementChildren) {
        const childTag = ctx.stripNs ? child.local : child.tag;
        const childElementChildren = child.children.filter(isElement);
        if (childElementChildren.length === 0 && Object.keys(child.attrs).length === 0) {
          const text = child.text.trim();
          out[childTag] = text === "" ? null : text;
        } else {
          out[childTag] = serializeMini(child);
        }
      }
      const mixedText = el.text.trim();
      if (mixedText !== "") out[ctx.textKey] = mixedText;
      return;
    }
    // Nested non-flattened branch: stringify whole subtree under its dot-path
    out[prefix] = serializeMini(el);
    return;
  }

  for (const child of elementChildren) {
    const childTag = ctx.stripNs ? child.local : child.tag;
    const nextKey = prefix ? `${prefix}.${childTag}` : childTag;
    elementToFlat(child, ctx, nextKey, out);
  }

  // Capture mixed content (element-with-text)
  const mixed = el.text.trim();
  if (mixed !== "") {
    out[prefix ? `${prefix}.${ctx.textKey}` : ctx.textKey] = mixed;
  }
}

function isElement(node: MiniNode): node is MiniElement {
  return (node as MiniElement).tag !== undefined;
}

function serializeMini(el: MiniElement): string {
  // Compact JSON-ish representation for flatten:false fallback
  const result: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(el.attrs)) result[`@${k}`] = v;
  const elementChildren = el.children.filter(isElement);
  if (elementChildren.length === 0) {
    if (el.text) result["#text"] = el.text.trim();
  } else {
    for (const child of elementChildren) {
      const tag = child.tag;
      const existing = result[tag];
      const sub = serializeMiniValue(child);
      if (existing === undefined) result[tag] = sub;
      else if (Array.isArray(existing)) (existing as unknown[]).push(sub);
      else result[tag] = [existing, sub];
    }
  }
  return JSON.stringify(result);
}

function serializeMiniValue(el: MiniElement): unknown {
  const result: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(el.attrs)) result[`@${k}`] = v;
  const children = el.children.filter(isElement);
  if (children.length === 0) {
    const text = el.text.trim();
    if (Object.keys(result).length === 0) return text;
    if (text) result["#text"] = text;
    return result;
  }
  for (const child of children) {
    const existing = result[child.tag];
    const sub = serializeMiniValue(child);
    if (existing === undefined) result[child.tag] = sub;
    else if (Array.isArray(existing)) (existing as unknown[]).push(sub);
    else result[child.tag] = [existing, sub];
  }
  return result;
}
