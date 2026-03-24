// ── Workbook XML Writer ──────────────────────────────────────────────
// Generates xl/workbook.xml, xl/_rels/workbook.xml.rels, and _rels/.rels

import type { WriteSheet } from "../_types";
import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer";

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const REL_WORKSHEET =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_SHARED_STRINGS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_WORKBOOK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** Generate xl/workbook.xml */
export function writeWorkbookXml(sheets: WriteSheet[]): string {
  const sheetElements: string[] = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const attrs: Record<string, string | number> = {
      name: sheet.name,
      sheetId: i + 1,
      "r:id": `rId${i + 1}`,
    };
    if (sheet.hidden) {
      attrs["state"] = "hidden";
    } else if (sheet.veryHidden) {
      attrs["state"] = "veryHidden";
    }
    sheetElements.push(xmlSelfClose("sheet", attrs));
  }

  const sheetsXml = xmlElement("sheets", undefined, sheetElements);

  return xmlDocument("workbook", { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R }, [sheetsXml]);
}

/** Generate xl/_rels/workbook.xml.rels */
export function writeWorkbookRels(sheetCount: number, hasSharedStrings: boolean): string {
  const children: string[] = [];

  // Worksheet relationships
  for (let i = 1; i <= sheetCount; i++) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${i}`,
        Type: REL_WORKSHEET,
        Target: `worksheets/sheet${i}.xml`,
      }),
    );
  }

  // Styles relationship
  const stylesRid = sheetCount + 1;
  children.push(
    xmlSelfClose("Relationship", {
      Id: `rId${stylesRid}`,
      Type: REL_STYLES,
      Target: "styles.xml",
    }),
  );

  // Shared strings relationship (if present)
  if (hasSharedStrings) {
    const ssRid = sheetCount + 2;
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${ssRid}`,
        Type: REL_SHARED_STRINGS,
        Target: "sharedStrings.xml",
      }),
    );
  }

  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, children);
}

/** Generate _rels/.rels */
export function writeRootRels(): string {
  const children: string[] = [];

  children.push(
    xmlSelfClose("Relationship", {
      Id: "rId1",
      Type: REL_WORKBOOK,
      Target: "xl/workbook.xml",
    }),
  );

  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, children);
}
