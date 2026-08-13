// ── XML Tree Accessors ───────────────────────────────────────────────
// Reading a parsed `XmlElement`: find a child by local name, list the
// element children, search a subtree, pull an element's direct text.
//
// `parseXml` returns the tree and nothing else, so every part reader
// wrote its own copy of these four — nine copies of `findChild` alone,
// identical to the character. They are here so there is one, and so a
// fix to one is a fix to all.
//
// `src/ods/reader.ts` keeps its own pair: they fall back to `tag` when
// `local` is empty, and that difference has not been shown to be
// unnecessary.

import type { XmlElement, XmlNode } from "./parser"

/** First element child with this local name, or `undefined`. */
export function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c
  }
  return undefined
}

/** Every element child, with text nodes dropped. */
export function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = []
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c)
  }
  return out
}

/**
 * First element at or below `el` with this local name — `el` itself
 * counts, so a hit on the root is returned rather than skipped.
 */
export function findDescendant(el: XmlElement, localName: string): XmlElement | undefined {
  if (el.local === localName) return el
  for (const c of el.children) {
    if (typeof c === "string") continue
    const hit = findDescendant(c, localName)
    if (hit) return hit
  }
  return undefined
}

/**
 * Direct text of the named child, or `""` when it is absent. Text inside
 * nested elements is not collected — the callers use this on leaf
 * elements whose content is a single run.
 */
export function readChildText(el: XmlElement, localName: string): string {
  const child = findChild(el, localName)
  if (!child) return ""
  let text = ""
  for (const c of child.children as XmlNode[]) {
    if (typeof c === "string") text += c
  }
  return text
}

/** Integer from an attribute value, falling back when absent or not a number. */
export function parseIntSafe(s: string | undefined, fallback: number): number {
  if (s === undefined) return fallback
  const n = parseInt(s, 10)
  return Number.isNaN(n) ? fallback : n
}
