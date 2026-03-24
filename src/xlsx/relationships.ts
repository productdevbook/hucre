// ── Relationships Parser ──────────────────────────────────────────────
// Parses .rels files (OPC relationship parts) from an XLSX package.

import { parseXml } from "../xml/parser";

export interface Relationship {
  id: string;
  type: string;
  target: string;
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
      if (id && type && target) {
        rels.push({ id, type, target });
      }
    }
  }

  return rels;
}
