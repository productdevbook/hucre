// ── Relationships Parser ──────────────────────────────────────────────
// Parses .rels files (OPC relationship parts) from an XLSX package.

import { parseXml } from "../xml/parser";

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/**
 * Parse a .rels XML file and return an array of relationships.
 */
export function parseRelationships(xml: string): Relationship[] {
  const doc = parseXml(xml);
  const rels: Relationship[] = [];

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "Relationship") {
      const id = child.attrs["Id"] ?? "";
      const type = child.attrs["Type"] ?? "";
      const target = child.attrs["Target"] ?? "";
      const targetMode = child.attrs["TargetMode"];
      if (id && type && target) {
        const rel: Relationship = { id, type, target };
        if (targetMode) rel.targetMode = targetMode;
        rels.push(rel);
      }
    }
  }

  return rels;
}
