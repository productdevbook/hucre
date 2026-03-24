// ── Content Types Parser ──────────────────────────────────────────────
// Parses [Content_Types].xml from an XLSX (OOXML) package.

import { parseXml } from "../xml/parser";

export interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

/**
 * Parse [Content_Types].xml.
 * Returns default extension→contentType mappings and part-specific overrides.
 */
export function parseContentTypes(xml: string): ContentTypes {
  const doc = parseXml(xml);
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "Default") {
      const ext = child.attrs["Extension"];
      const ct = child.attrs["ContentType"];
      if (ext && ct) {
        defaults.set(ext, ct);
      }
    } else if (local === "Override") {
      const partName = child.attrs["PartName"];
      const ct = child.attrs["ContentType"];
      if (partName && ct) {
        overrides.set(partName, ct);
      }
    }
  }

  return { defaults, overrides };
}
