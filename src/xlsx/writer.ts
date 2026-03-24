// ── XLSX Writer ──────────────────────────────────────────────────────
// Generates valid Office Open XML spreadsheet files (XLSX).

import type { WriteOptions, WriteOutput } from "../_types";
import { ZipWriter } from "../zip/writer";
import { writeContentTypes } from "./content-types-writer";
import { writeRootRels, writeWorkbookXml, writeWorkbookRels } from "./workbook-writer";
import { createStylesCollector } from "./styles-writer";
import { createSharedStrings, writeSharedStringsXml, writeWorksheetXml } from "./worksheet-writer";

const encoder = /* @__PURE__ */ new TextEncoder();

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
  const worksheetXmls: string[] = [];
  for (const sheet of sheets) {
    const xml = writeWorksheetXml(sheet, styles, sharedStrings, dateSystem);
    worksheetXmls.push(xml);
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

  // xl/worksheets/sheetN.xml
  for (let i = 0; i < worksheetXmls.length; i++) {
    zip.add(`xl/worksheets/sheet${i + 1}.xml`, encoder.encode(worksheetXmls[i]));
  }

  return zip.build();
}
