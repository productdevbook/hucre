// ── XLSX Writer ──────────────────────────────────────────────────────
// Generates valid Office Open XML spreadsheet files (XLSX).

import type { WriteOptions, WriteOutput } from "../_types";
import { ZipWriter } from "../zip/writer";
import { writeContentTypes } from "./content-types-writer";
import { writeRootRels, writeWorkbookXml, writeWorkbookRels } from "./workbook-writer";
import { createStylesCollector } from "./styles-writer";
import { createSharedStrings, writeSharedStringsXml, writeWorksheetXml } from "./worksheet-writer";
import type { WorksheetResult } from "./worksheet-writer";
import { xmlDocument, xmlSelfClose } from "../xml/writer";

const encoder = /* @__PURE__ */ new TextEncoder();

const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/**
 * Write a Workbook to XLSX format.
 * Returns a Uint8Array containing the ZIP archive.
 */
export async function writeXlsx(options: WriteOptions): Promise<WriteOutput> {
  const { sheets, defaultFont, dateSystem } = options;

  // Create shared collectors
  const styles = createStylesCollector(defaultFont);
  const sharedStrings = createSharedStrings();

  // Generate worksheet XMLs (also populates styles and shared strings)
  const worksheetResults: WorksheetResult[] = [];
  for (const sheet of sheets) {
    const result = writeWorksheetXml(sheet, styles, sharedStrings, dateSystem);
    worksheetResults.push(result);
  }

  const hasSharedStrings = sharedStrings.count() > 0;

  // Build ZIP archive
  const zip = new ZipWriter();

  // [Content_Types].xml
  zip.add(
    "[Content_Types].xml",
    encoder.encode(writeContentTypes(sheets.length, hasSharedStrings)),
  );

  // _rels/.rels
  zip.add("_rels/.rels", encoder.encode(writeRootRels()));

  // xl/workbook.xml
  zip.add("xl/workbook.xml", encoder.encode(writeWorkbookXml(sheets)));

  // xl/_rels/workbook.xml.rels
  zip.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(writeWorkbookRels(sheets.length, hasSharedStrings)),
  );

  // xl/styles.xml
  zip.add("xl/styles.xml", encoder.encode(styles.toXml()));

  // xl/sharedStrings.xml (if any strings)
  if (hasSharedStrings) {
    zip.add("xl/sharedStrings.xml", encoder.encode(writeSharedStringsXml(sharedStrings)));
  }

  // xl/worksheets/sheetN.xml + optional xl/worksheets/_rels/sheetN.xml.rels
  for (let i = 0; i < worksheetResults.length; i++) {
    const result = worksheetResults[i];
    zip.add(`xl/worksheets/sheet${i + 1}.xml`, encoder.encode(result.xml));

    // Generate worksheet .rels for external hyperlinks
    if (result.hyperlinkRelationships.length > 0) {
      const relElements: string[] = [];
      for (const rel of result.hyperlinkRelationships) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: rel.id,
            Type: REL_HYPERLINK,
            Target: rel.target,
            TargetMode: "External",
          }),
        );
      }
      const relsXml = xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, relElements);
      zip.add(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, encoder.encode(relsXml));
    }
  }

  return zip.build();
}
