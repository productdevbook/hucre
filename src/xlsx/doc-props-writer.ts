// ── Document Properties Writer ──────────────────────────────────────
// Generates docProps/core.xml and docProps/app.xml for XLSX packages.

import type { WorkbookProperties } from "../_types";
import { xmlDocument, xmlElement, xmlEscape } from "../xml/writer";

// ── Namespaces ──────────────────────────────────────────────────────

const NS_CORE_PROPERTIES =
  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const NS_DC = "http://purl.org/dc/elements/1.1/";
const NS_DCTERMS = "http://purl.org/dc/terms/";
const NS_XSI = "http://www.w3.org/2001/XMLSchema-instance";

const NS_EXTENDED_PROPERTIES =
  "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

// ── Helpers ─────────────────────────────────────────────────────────

function formatW3CDTF(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

// ── core.xml ────────────────────────────────────────────────────────

/**
 * Generate docProps/core.xml from workbook properties.
 * Always includes a modified date (defaults to now).
 */
export function writeCoreProperties(props?: WorkbookProperties): string {
  const children: string[] = [];

  if (props?.title) {
    children.push(xmlElement("dc:title", undefined, xmlEscape(props.title)));
  }

  if (props?.subject) {
    children.push(xmlElement("dc:subject", undefined, xmlEscape(props.subject)));
  }

  if (props?.creator) {
    children.push(xmlElement("dc:creator", undefined, xmlEscape(props.creator)));
  }

  if (props?.keywords) {
    children.push(xmlElement("cp:keywords", undefined, xmlEscape(props.keywords)));
  }

  if (props?.description) {
    children.push(xmlElement("dc:description", undefined, xmlEscape(props.description)));
  }

  if (props?.lastModifiedBy) {
    children.push(xmlElement("cp:lastModifiedBy", undefined, xmlEscape(props.lastModifiedBy)));
  }

  if (props?.category) {
    children.push(xmlElement("cp:category", undefined, xmlEscape(props.category)));
  }

  if (props?.created) {
    children.push(
      xmlElement("dcterms:created", { "xsi:type": "dcterms:W3CDTF" }, formatW3CDTF(props.created)),
    );
  }

  // Always include modified date
  const modified = props?.modified ?? new Date();
  children.push(
    xmlElement("dcterms:modified", { "xsi:type": "dcterms:W3CDTF" }, formatW3CDTF(modified)),
  );

  return xmlDocument(
    "cp:coreProperties",
    {
      "xmlns:cp": NS_CORE_PROPERTIES,
      "xmlns:dc": NS_DC,
      "xmlns:dcterms": NS_DCTERMS,
      "xmlns:xsi": NS_XSI,
    },
    children,
  );
}

// ── app.xml ─────────────────────────────────────────────────────────

/**
 * Generate docProps/app.xml from workbook properties.
 * Always includes Application: "defter".
 */
export function writeAppProperties(props?: WorkbookProperties): string {
  const children: string[] = [];

  // Always include the application name
  children.push(xmlElement("Application", undefined, "defter"));

  if (props?.company) {
    children.push(xmlElement("Company", undefined, xmlEscape(props.company)));
  }

  if (props?.manager) {
    children.push(xmlElement("Manager", undefined, xmlEscape(props.manager)));
  }

  return xmlDocument(
    "Properties",
    {
      xmlns: NS_EXTENDED_PROPERTIES,
    },
    children,
  );
}
