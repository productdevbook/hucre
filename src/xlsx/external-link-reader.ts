// ── External Link Reader ──────────────────────────────────────────
// Parses xl/externalLinks/externalLinkN.xml plus its sibling
// `_rels/externalLinkN.xml.rels` into a structured ExternalLink so
// callers can inspect linked workbooks and their cached cell values.
//
// OOXML reference: ECMA-376 Part 1, §18.14 (External Workbook References).

import type {
  ExternalCachedCell,
  ExternalCellType,
  ExternalDefinedName,
  ExternalLink,
  ExternalSheetData,
} from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement, XmlNode } from "../xml/parser";
import { parseRelationships } from "./relationships";

const VALID_TYPES: ReadonlySet<ExternalCellType> = new Set(["n", "s", "b", "e", "str"]);

const REL_EXTERNAL_LINK_PATH =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
const REL_EXTERNAL_LINK_PATH_STRICT =
  "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath";

/**
 * Parse a single external link.
 *
 * @param xml      Raw XML of `xl/externalLinks/externalLinkN.xml`.
 * @param relsXml  Optional XML of `xl/externalLinks/_rels/externalLinkN.xml.rels`,
 *                 used to resolve the external workbook's target path. When
 *                 omitted the returned `target` is the empty string.
 */
export function parseExternalLink(xml: string, relsXml?: string): ExternalLink {
  const root = parseXml(xml);
  const externalBook = findChild(root, "externalBook");

  let target = "";
  let targetMode: "External" | "Internal" | undefined;
  if (externalBook && relsXml) {
    const rels = parseRelationships(relsXml);
    const rId = externalBook.attrs["r:id"] ?? externalBook.attrs.id;
    const rel = rId ? rels.find((r) => r.id === rId) : undefined;
    if (rel) {
      target = rel.target;
      if (rel.targetMode === "External" || rel.targetMode === "Internal") {
        targetMode = rel.targetMode;
      }
    }
  }

  const sheetNames = parseSheetNames(externalBook);
  const sheetData = parseSheetDataSet(externalBook);
  const definedNames = parseDefinedNames(externalBook);

  const link: ExternalLink = { target, sheetNames, sheetData };
  if (targetMode) link.targetMode = targetMode;
  if (definedNames && definedNames.length > 0) link.definedNames = definedNames;
  return link;
}

// ── Internals ─────────────────────────────────────────────────────

function parseSheetNames(externalBook: XmlElement | undefined): string[] {
  const sheetNames = externalBook ? findChild(externalBook, "sheetNames") : undefined;
  if (!sheetNames) return [];
  const names: string[] = [];
  for (const child of childElements(sheetNames)) {
    if (child.local === "sheetName") names.push(child.attrs.val ?? "");
  }
  return names;
}

function parseSheetDataSet(externalBook: XmlElement | undefined): ExternalSheetData[] {
  const dataSet = externalBook ? findChild(externalBook, "sheetDataSet") : undefined;
  if (!dataSet) return [];
  const result: ExternalSheetData[] = [];
  for (const sheetData of childElements(dataSet)) {
    if (sheetData.local !== "sheetData") continue;
    const sheetId = parseIntSafe(sheetData.attrs.sheetId, 0);
    const cells: ExternalCachedCell[] = [];
    for (const row of childElements(sheetData)) {
      if (row.local !== "row") continue;
      for (const cell of childElements(row)) {
        if (cell.local !== "cell") continue;
        const ref = cell.attrs.r ?? "";
        if (!ref) continue;
        const rawType = (cell.attrs.t ?? "n") as ExternalCellType;
        const type: ExternalCellType = VALID_TYPES.has(rawType) ? rawType : "n";
        const valueText = readChildText(cell, "v");
        cells.push({ ref, type, value: coerceValue(type, valueText) });
      }
    }
    result.push({ sheetId, cells });
  }
  return result;
}

function parseDefinedNames(externalBook: XmlElement | undefined): ExternalDefinedName[] {
  const dn = externalBook ? findChild(externalBook, "definedNames") : undefined;
  if (!dn) return [];
  const result: ExternalDefinedName[] = [];
  for (const child of childElements(dn)) {
    if (child.local !== "definedName") continue;
    const entry: ExternalDefinedName = { name: child.attrs.name ?? "" };
    if (child.attrs.refersTo) entry.refersTo = child.attrs.refersTo;
    if (child.attrs.sheetId !== undefined) {
      const id = parseIntSafe(child.attrs.sheetId, NaN);
      if (!Number.isNaN(id)) entry.sheetId = id;
    }
    if (entry.name) result.push(entry);
  }
  return result;
}

function coerceValue(type: ExternalCellType, text: string): string | number | boolean {
  switch (type) {
    case "n":
    case "s": {
      // `s` here is the shared-string index of the *external* workbook;
      // the index is meaningless without that workbook so we keep it as
      // a number for fidelity. Callers wanting the resolved string need
      // the linked workbook itself.
      const n = Number(text);
      return Number.isFinite(n) ? n : 0;
    }
    case "b":
      return text === "1" || text === "true";
    case "e":
    case "str":
      return text;
  }
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}

function readChildText(el: XmlElement, localName: string): string {
  const child = findChild(el, localName);
  if (!child) return "";
  let text = "";
  for (const c of child.children as XmlNode[]) {
    if (typeof c === "string") text += c;
  }
  return text;
}

function parseIntSafe(s: string | undefined, fallback: number): number {
  if (s === undefined) return fallback;
  const n = parseInt(s, 10);
  return Number.isNaN(n) ? fallback : n;
}

// Deliberately exported but not used internally — exposed for callers
// that already extracted the relationship list and just want the body.
export const REL_EXTERNAL_LINK_PATH_TYPES = [
  REL_EXTERNAL_LINK_PATH,
  REL_EXTERNAL_LINK_PATH_STRICT,
] as const;
