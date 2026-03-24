// ── Content Types Writer ──────────────────────────────────────────────
// Generates [Content_Types].xml for an XLSX package.

import { xmlDocument, xmlSelfClose } from "../xml/writer";

const NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

const CT_RELS = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML = "application/xml";
const CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SHARED_STRINGS =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

/** Generate [Content_Types].xml for XLSX */
export function writeContentTypes(sheetCount: number, hasSharedStrings: boolean): string {
  const children: string[] = [];

  // Default extension mappings
  children.push(xmlSelfClose("Default", { Extension: "rels", ContentType: CT_RELS }));
  children.push(xmlSelfClose("Default", { Extension: "xml", ContentType: CT_XML }));

  // Override for workbook
  children.push(
    xmlSelfClose("Override", {
      PartName: "/xl/workbook.xml",
      ContentType: CT_WORKBOOK,
    }),
  );

  // Override for each worksheet
  for (let i = 1; i <= sheetCount; i++) {
    children.push(
      xmlSelfClose("Override", {
        PartName: `/xl/worksheets/sheet${i}.xml`,
        ContentType: CT_WORKSHEET,
      }),
    );
  }

  // Override for styles
  children.push(
    xmlSelfClose("Override", {
      PartName: "/xl/styles.xml",
      ContentType: CT_STYLES,
    }),
  );

  // Override for shared strings (if present)
  if (hasSharedStrings) {
    children.push(
      xmlSelfClose("Override", {
        PartName: "/xl/sharedStrings.xml",
        ContentType: CT_SHARED_STRINGS,
      }),
    );
  }

  return xmlDocument("Types", { xmlns: NS_CONTENT_TYPES }, children);
}
