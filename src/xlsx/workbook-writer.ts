// ── Workbook XML Writer ──────────────────────────────────────────────
// Generates xl/workbook.xml, xl/_rels/workbook.xml.rels, and _rels/.rels

import type { WriteSheet, NamedRange } from "../_types";
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer";

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const REL_WORKSHEET =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_SHARED_STRINGS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_WORKBOOK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** Generate xl/workbook.xml */
export function writeWorkbookXml(sheets: WriteSheet[], namedRanges?: NamedRange[]): string {
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

  const parts: string[] = [];
  parts.push(xmlElement("sheets", undefined, sheetElements));

  // ── Defined Names (named ranges + print area/titles) ──
  if (namedRanges && namedRanges.length > 0) {
    // Build sheet name → index map for resolving scoped named ranges
    const sheetIndexMap = new Map<string, number>();
    for (let i = 0; i < sheets.length; i++) {
      sheetIndexMap.set(sheets[i].name, i);
    }

    const dnElements: string[] = [];

    for (const nr of namedRanges) {
      const attrs: Record<string, string | number> = {
        name: nr.name,
      };

      // Resolve scope: if scope is a sheet name, convert to localSheetId (0-based index)
      if (nr.scope !== undefined) {
        const idx = sheetIndexMap.get(nr.scope);
        if (idx !== undefined) {
          attrs["localSheetId"] = idx;
        }
      }

      if (nr.comment) {
        attrs["comment"] = nr.comment;
      }

      dnElements.push(xmlElement("definedName", attrs, xmlEscape(nr.range)));
    }

    parts.push(xmlElement("definedNames", undefined, dnElements));
  }

  return xmlDocument("workbook", { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R }, parts);
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

  // Theme relationship
  const themeRid = sheetCount + (hasSharedStrings ? 3 : 2);
  children.push(
    xmlSelfClose("Relationship", {
      Id: `rId${themeRid}`,
      Type: REL_THEME,
      Target: "theme/theme1.xml",
    }),
  );

  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, children);
}

/** Generate _rels/.rels (with optional docProps) */
export function writeRootRels(options?: { hasCoreProps?: boolean; hasAppProps?: boolean }): string {
  const children: string[] = [];

  children.push(
    xmlSelfClose("Relationship", {
      Id: "rId1",
      Type: REL_WORKBOOK,
      Target: "xl/workbook.xml",
    }),
  );

  let nextRId = 2;

  if (options?.hasCoreProps) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRId++}`,
        Type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        Target: "docProps/core.xml",
      }),
    );
  }

  if (options?.hasAppProps) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRId++}`,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        Target: "docProps/app.xml",
      }),
    );
  }

  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, children);
}
