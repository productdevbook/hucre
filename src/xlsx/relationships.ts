// ── Relationships Parser ──────────────────────────────────────────────
// Parses .rels files (OPC relationship parts) from an XLSX package, and
// the two path rules that go with them: where a relationship's target
// lands, and which attribute carries the rId.
//
// `readXlsx`, `streamXlsxRows` and `readXlsb` all walk a package the same
// way and each used to carry its own copy of these. One copy, so a
// package the buffered reader can open is one the streaming reader can.

import { parseXml } from "../xml/parser"

export interface Relationship {
  id: string
  type: string
  target: string
  targetMode?: string
}

/**
 * Parse a .rels XML file and return an array of relationships.
 */
export function parseRelationships(xml: string): Relationship[] {
  const doc = parseXml(xml)
  const rels: Relationship[] = []

  for (const child of doc.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag

    if (local === "Relationship") {
      const id = child.attrs["Id"] ?? ""
      const type = child.attrs["Type"] ?? ""
      const target = child.attrs["Target"] ?? ""
      const targetMode = child.attrs["TargetMode"]
      if (id && type && target) {
        const rel: Relationship = { id, type, target }
        if (targetMode) rel.targetMode = targetMode
        rels.push(rel)
      }
    }
  }

  return rels
}

/**
 * Resolve a relative target path against a base directory.
 * E.g. resolvePath("xl/_rels", "../worksheets/sheet1.xml") → "xl/worksheets/sheet1.xml"
 */
export function resolvePath(base: string, target: string): string {
  // If target starts with /, it's absolute from the package root
  if (target.startsWith("/")) return target.slice(1)

  const baseParts = base.split("/").filter(Boolean)
  const targetParts = target.split("/").filter(Boolean)

  for (const part of targetParts) {
    if (part === "..") {
      baseParts.pop()
    } else if (part !== ".") {
      baseParts.push(part)
    }
  }

  return baseParts.join("/")
}

/**
 * Get the directory portion of a path.
 * E.g. "xl/workbook.xml" → "xl"
 */
export function dirname(path: string): string {
  const idx = path.lastIndexOf("/")
  return idx === -1 ? "" : path.slice(0, idx)
}

/**
 * The rId an element points at, when the `r:` namespace prefix is not
 * spelled `r`. Producers bind the relationship namespace to whatever
 * prefix they like, so the attribute is found by its shape — a name
 * ending `:id` whose value looks like an rId — rather than by its name.
 */
export function findRIdAttr(attrs: Record<string, string>): string | undefined {
  for (const key of Object.keys(attrs)) {
    if (key.endsWith(":id") && attrs[key].startsWith("rId")) {
      return attrs[key]
    }
  }
  return undefined
}
