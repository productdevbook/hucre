// ── Chart Util ─────────────────────────────────────────────────────
// Tiny XML walk + bool-attr helpers shared by every per-host parser
// module. The chart-reader, chart/text and chart/shape modules used to
// each carry their own private copies of `findChild` because the
// underlying `XmlElement` walker did not expose one — the helpers here
// give the per-host modules a single source of truth without disturbing
// the parser surface.

import type { XmlElement } from "../../xml/parser";

export function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

export function findDescendant(el: XmlElement, localName: string): XmlElement | undefined {
  if (el.local === localName) return el;
  for (const c of el.children) {
    if (typeof c === "string") continue;
    const hit = findDescendant(c, localName);
    if (hit) return hit;
  }
  return undefined;
}

export function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}

export function collectTextRuns(el: XmlElement, out: string[]): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    if (child.local === "t") {
      out.push(elementText(child));
    } else {
      collectTextRuns(child, out);
    }
  }
}

export function elementText(el: XmlElement): string {
  let buf = "";
  for (const child of el.children) {
    if (typeof child === "string") buf += child;
    else buf += elementText(child);
  }
  return buf;
}

/** Coerce an XML boolean attribute (`"0"`, `"1"`, `"true"`, `"false"`). */
export function parseBoolAttr(value: unknown): boolean | undefined {
  if (typeof value !== "string") return undefined;
  const v = value.trim().toLowerCase();
  if (v === "1" || v === "true") return true;
  if (v === "0" || v === "false") return false;
  return undefined;
}

export function readBoolAttr(el: XmlElement): boolean | undefined {
  const v = el.attrs.val;
  if (typeof v !== "string") return undefined;
  return v === "1" || v.toLowerCase() === "true";
}

export function readBoolVal(raw: string | undefined): boolean | undefined {
  if (raw === undefined) return undefined;
  if (raw === "1" || raw === "true") return true;
  if (raw === "0" || raw === "false") return false;
  return undefined;
}

/**
 * Walk `<c:val>` / `<c:cat>` / `<c:xVal>` / `<c:yVal>` to its inner
 * `<c:f>` formula text. Returns `undefined` for embedded `<c:numLit>`
 * literal data (no formula) or when the element is absent.
 */
export function formulaText(wrapper: XmlElement | undefined): string | undefined {
  if (!wrapper) return undefined;
  const numRef = findChild(wrapper, "numRef") ?? findChild(wrapper, "strRef");
  if (numRef) {
    const f = findChild(numRef, "f");
    if (f) {
      const text = elementText(f).trim();
      if (text.length > 0) return text;
    }
  }
  // Some writers put <c:f> directly under <c:strRef> (already handled
  // above via numRef fallback) or under the wrapper itself.
  const direct = findChild(wrapper, "f");
  if (direct) {
    const text = elementText(direct).trim();
    if (text.length > 0) return text;
  }
  return undefined;
}

export function parseNumericChildVal(parent: XmlElement, localName: string): number | undefined {
  const child = findChild(parent, localName);
  if (!child) return undefined;
  const raw = child.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : undefined;
}

/**
 * Generic clone-override resolver. `undefined` keeps the source value
 * (inherit), `null` drops it (suppress), and any other value replaces.
 * Used by every per-field clone resolver that does not need additional
 * normalization on the value.
 */
export function applyOverride<T>(
  sourceValue: T | undefined,
  override: T | null | undefined,
): T | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}
