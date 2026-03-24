// ── Document Properties Reader ──────────────────────────────────────
// Parses docProps/core.xml and docProps/app.xml from XLSX packages.

import type { WorkbookProperties } from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";

// ── Helpers ─────────────────────────────────────────────────────────

function getChildText(parent: XmlElement, localName: string): string | undefined {
  for (const child of parent.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    if (local === localName) {
      const text = child.children.filter((c: unknown) => typeof c === "string").join("");
      return text || undefined;
    }
  }
  return undefined;
}

function parseW3CDTF(value: string): Date | undefined {
  if (!value) return undefined;
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return undefined;
  return d;
}

// ── core.xml parsing ────────────────────────────────────────────────

/**
 * Parse docProps/core.xml into WorkbookProperties fields.
 */
export function parseCoreProperties(xml: string): Partial<WorkbookProperties> {
  const doc = parseXml(xml);
  const props: Partial<WorkbookProperties> = {};

  // core.xml uses namespaced tags: dc:title, dc:subject, dc:creator,
  // cp:keywords, dc:description, cp:lastModifiedBy, cp:category,
  // dcterms:created, dcterms:modified
  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;
    const text = child.children.filter((c: unknown) => typeof c === "string").join("");

    switch (local) {
      case "title":
        if (text) props.title = text;
        break;
      case "subject":
        if (text) props.subject = text;
        break;
      case "creator":
        if (text) props.creator = text;
        break;
      case "keywords":
        if (text) props.keywords = text;
        break;
      case "description":
        if (text) props.description = text;
        break;
      case "lastModifiedBy":
        if (text) props.lastModifiedBy = text;
        break;
      case "category":
        if (text) props.category = text;
        break;
      case "created":
        if (text) {
          const d = parseW3CDTF(text);
          if (d) props.created = d;
        }
        break;
      case "modified":
        if (text) {
          const d = parseW3CDTF(text);
          if (d) props.modified = d;
        }
        break;
    }
  }

  return props;
}

// ── app.xml parsing ─────────────────────────────────────────────────

/**
 * Parse docProps/app.xml into WorkbookProperties fields.
 */
export function parseAppProperties(xml: string): Partial<WorkbookProperties> {
  const doc = parseXml(xml);
  const props: Partial<WorkbookProperties> = {};

  const company = getChildText(doc, "Company");
  if (company) props.company = company;

  const manager = getChildText(doc, "Manager");
  if (manager) props.manager = manager;

  return props;
}
