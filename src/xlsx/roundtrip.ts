// ── XLSX Round-Trip Preservation ─────────────────────────────────────
// Read an XLSX file, modify cells, write it back without losing charts,
// images, macros, shapes, or other features that defter doesn't natively
// understand.

import type { Workbook, ReadOptions, WriteSheet, NamedRange } from "../_types";
import { readXlsx } from "./reader";
import { ZipReader } from "../zip/reader";
import { ZipWriter } from "../zip/writer";
import { writeContentTypes } from "./content-types-writer";
import type { ContentTypesOptions } from "./content-types-writer";
import { writeRootRels, writeWorkbookXml, writeWorkbookRels } from "./workbook-writer";
import { createStylesCollector } from "./styles-writer";
import { createSharedStrings, writeSharedStringsXml, writeWorksheetXml } from "./worksheet-writer";
import type { WorksheetResult } from "./worksheet-writer";
import { writeDrawing } from "./drawing-writer";
import type { DrawingResult } from "./drawing-writer";
import { writeComments } from "./comments-writer";
import type { CommentsResult } from "./comments-writer";
import { writeTable } from "./table-writer";
import { colToLetter } from "./worksheet-writer";
import { xmlDocument, xmlSelfClose } from "../xml/writer";
import { writeCoreProperties, writeAppProperties } from "./doc-props-writer";

// ── Types ────────────────────────────────────────────────────────────

export interface RoundtripWorkbook extends Workbook {
  /** Raw ZIP entries from the original file (for preservation) */
  _rawEntries: Map<string, Uint8Array>;
  /** Paths of parts that were modified and need regeneration */
  _modifiedParts: Set<string>;
  /** Original content types XML */
  _contentTypes: string;
  /** Original root rels XML */
  _rootRels: string;
}

// ── Constants ────────────────────────────────────────────────────────

const encoder = /* @__PURE__ */ new TextEncoder();

const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_VML_DRAWING =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

/**
 * Parts that defter regenerates from parsed data.
 * Matched case-insensitively with normalized paths.
 */
const REGENERATED_PREFIXES = [
  "xl/workbook.xml",
  "xl/worksheets/",
  "xl/sharedstrings.xml",
  "xl/styles.xml",
  "[content_types].xml",
  "_rels/.rels",
  "xl/_rels/workbook.xml.rels",
  "docprops/core.xml",
  "docprops/app.xml",
];

/**
 * Additional parts regenerated per-sheet (drawings, comments, tables).
 * These are regenerated when a sheet has images/comments/tables.
 */
const REGENERATED_SHEET_PREFIXES = [
  "xl/drawings/drawing",
  "xl/drawings/_rels/drawing",
  "xl/drawings/vmldrawing",
  "xl/comments",
  "xl/tables/table",
  "xl/media/image",
  "xl/worksheets/_rels/sheet",
];

// ── openXlsx ─────────────────────────────────────────────────────────

/**
 * Read an XLSX file and return a Workbook with preserved raw data
 * for round-trip writing.
 */
export async function openXlsx(
  input: Uint8Array | ArrayBuffer,
  options?: ReadOptions,
): Promise<RoundtripWorkbook> {
  const data = input instanceof Uint8Array ? input : new Uint8Array(input);

  // 1. Parse the workbook normally
  const workbook = await readXlsx(data, options);

  // 2. Extract ALL raw ZIP entries
  const zip = new ZipReader(data);
  const rawEntries = await zip.extractAll();

  // 3. Read content types and root rels for preservation
  const decoder = new TextDecoder("utf-8");
  const contentTypes = zip.has("[Content_Types].xml")
    ? decoder.decode(await zip.extract("[Content_Types].xml"))
    : "";
  const rootRels = zip.has("_rels/.rels") ? decoder.decode(await zip.extract("_rels/.rels")) : "";

  // 4. Build RoundtripWorkbook
  const rtWorkbook: RoundtripWorkbook = {
    ...workbook,
    _rawEntries: rawEntries,
    _modifiedParts: new Set<string>(),
    _contentTypes: contentTypes,
    _rootRels: rootRels,
  };

  return rtWorkbook;
}

// ── saveXlsx ─────────────────────────────────────────────────────────

/**
 * Write a RoundtripWorkbook back to XLSX, preserving unmodified parts.
 */
export async function saveXlsx(workbook: RoundtripWorkbook): Promise<Uint8Array> {
  const { sheets, properties, namedRanges, dateSystem, defaultFont, activeSheet } = workbook;

  // Convert Sheet[] to WriteSheet[] for the writer infrastructure
  const writeSheets: WriteSheet[] = sheets.map((sheet) => ({
    name: sheet.name,
    columns: sheet.columns,
    rows: sheet.rows,
    cells: sheet.cells,
    merges: sheet.merges,
    dataValidations: sheet.dataValidations,
    conditionalRules: sheet.conditionalRules,
    autoFilter: sheet.autoFilter,
    freezePane: sheet.freezePane,
    images: sheet.images,
    protection: sheet.protection,
    pageSetup: sheet.pageSetup,
    headerFooter: sheet.headerFooter,
    view: sheet.view,
    hidden: sheet.hidden,
    veryHidden: sheet.veryHidden,
    tables: sheet.tables,
    rowDefs: sheet.rowDefs,
  }));

  // Create shared collectors
  const styles = createStylesCollector(defaultFont);
  const sharedStrings = createSharedStrings();

  // Pre-compute global table start indices per sheet
  let globalTableCounter = 1;
  const sheetTableStartIndices: Array<number | undefined> = [];
  for (const sheet of writeSheets) {
    if (sheet.tables && sheet.tables.length > 0) {
      sheetTableStartIndices.push(globalTableCounter);
      globalTableCounter += sheet.tables.length;
    } else {
      sheetTableStartIndices.push(undefined);
    }
  }

  // Generate worksheet XMLs (also populates styles and shared strings)
  const worksheetResults: WorksheetResult[] = [];
  for (let i = 0; i < writeSheets.length; i++) {
    const sheet = writeSheets[i];
    const result = writeWorksheetXml(
      sheet,
      styles,
      sharedStrings,
      dateSystem,
      sheetTableStartIndices[i],
    );
    worksheetResults.push(result);
  }

  const hasSharedStrings = sharedStrings.count() > 0;

  // Generate drawing data for sheets that have images
  const drawingResults: Array<DrawingResult | null> = [];
  const drawingIndices: number[] = [];
  const imageExtensions = new Set<string>();
  let globalImageIndex = 1;

  for (let i = 0; i < writeSheets.length; i++) {
    const sheet = writeSheets[i];
    if (sheet.images && sheet.images.length > 0) {
      const result = writeDrawing(sheet.images, globalImageIndex);
      drawingResults.push(result);
      drawingIndices.push(i + 1);
      for (const img of result.images) {
        const ext = img.path.split(".").pop();
        if (ext) imageExtensions.add(ext);
      }
      globalImageIndex += sheet.images.length;
    } else {
      drawingResults.push(null);
    }
  }

  // Generate comments data for sheets that have comments
  const commentsResults: Array<CommentsResult | null> = [];
  const commentIndices: number[] = [];

  for (let i = 0; i < writeSheets.length; i++) {
    const sheet = writeSheets[i];
    if (sheet.cells) {
      const result = writeComments(sheet.cells, i);
      if (result) {
        commentsResults.push(result);
        commentIndices.push(i + 1);
      } else {
        commentsResults.push(null);
      }
    } else {
      commentsResults.push(null);
    }
  }

  // Collect all table indices for content types
  const allTableIndices: number[] = [];
  for (const result of worksheetResults) {
    for (const t of result.tables) {
      allTableIndices.push(t.globalTableIndex);
    }
  }

  // Build the set of paths we will regenerate
  const regeneratedPaths = new Set<string>();

  // Core parts always regenerated
  regeneratedPaths.add("[Content_Types].xml");
  regeneratedPaths.add("_rels/.rels");
  regeneratedPaths.add("xl/workbook.xml");
  regeneratedPaths.add("xl/_rels/workbook.xml.rels");
  regeneratedPaths.add("xl/styles.xml");
  regeneratedPaths.add("docProps/core.xml");
  regeneratedPaths.add("docProps/app.xml");
  if (hasSharedStrings) {
    regeneratedPaths.add("xl/sharedStrings.xml");
  }

  // Per-sheet parts
  for (let i = 0; i < worksheetResults.length; i++) {
    const idx = i + 1;
    regeneratedPaths.add(`xl/worksheets/sheet${idx}.xml`);
    regeneratedPaths.add(`xl/worksheets/_rels/sheet${idx}.xml.rels`);

    const drawing = drawingResults[i];
    if (drawing) {
      regeneratedPaths.add(`xl/drawings/drawing${idx}.xml`);
      regeneratedPaths.add(`xl/drawings/_rels/drawing${idx}.xml.rels`);
      for (const img of drawing.images) {
        regeneratedPaths.add(img.path);
      }
    }

    const comments = commentsResults[i];
    if (comments) {
      regeneratedPaths.add(`xl/comments${idx}.xml`);
      regeneratedPaths.add(`xl/drawings/vmlDrawing${idx}.vml`);
    }

    const wsResult = worksheetResults[i];
    for (const tableEntry of wsResult.tables) {
      regeneratedPaths.add(`xl/tables/table${tableEntry.globalTableIndex}.xml`);
    }
  }

  // Build ZIP archive
  const zip = new ZipWriter();

  // 1. Add all preserved raw entries (parts we don't regenerate)
  for (const [path, data] of workbook._rawEntries) {
    // Remove calcChain.xml — it becomes stale when formulas change.
    // Excel rebuilds it automatically when opening the file.
    if (path.toLowerCase() === "xl/calcchain.xml") continue;

    if (!regeneratedPaths.has(path)) {
      // Check if this path matches any regenerated prefix pattern (case-insensitive)
      const lowerPath = path.toLowerCase();
      let isRegenerated = false;
      for (const prefix of REGENERATED_PREFIXES) {
        if (lowerPath === prefix || lowerPath.startsWith(prefix)) {
          isRegenerated = true;
          break;
        }
      }
      if (!isRegenerated) {
        for (const prefix of REGENERATED_SHEET_PREFIXES) {
          if (lowerPath.startsWith(prefix)) {
            isRegenerated = true;
            break;
          }
        }
      }
      if (!isRegenerated) {
        // Preserve this entry as-is (don't compress, keep original bytes)
        zip.add(path, data, { compress: false });
      }
    }
  }

  // 2. Generate and add regenerated parts

  // [Content_Types].xml
  const ctOpts: ContentTypesOptions = {
    sheetCount: writeSheets.length,
    hasSharedStrings,
    drawingIndices: drawingIndices.length > 0 ? drawingIndices : undefined,
    imageExtensions: imageExtensions.size > 0 ? imageExtensions : undefined,
    commentIndices: commentIndices.length > 0 ? commentIndices : undefined,
    tableIndices: allTableIndices.length > 0 ? allTableIndices : undefined,
    hasCoreProps: true,
    hasAppProps: true,
  };
  zip.add("[Content_Types].xml", encoder.encode(writeContentTypes(ctOpts)));

  // _rels/.rels
  zip.add("_rels/.rels", encoder.encode(writeRootRels({ hasCoreProps: true, hasAppProps: true })));

  // docProps
  zip.add("docProps/core.xml", encoder.encode(writeCoreProperties(properties)));
  zip.add("docProps/app.xml", encoder.encode(writeAppProperties(properties)));

  // xl/workbook.xml
  const allNamedRanges = buildNamedRanges(writeSheets, namedRanges);
  zip.add(
    "xl/workbook.xml",
    encoder.encode(
      writeWorkbookXml(
        writeSheets,
        allNamedRanges.length > 0 ? allNamedRanges : undefined,
        dateSystem,
        activeSheet,
      ),
    ),
  );

  // xl/_rels/workbook.xml.rels
  zip.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(writeWorkbookRels(writeSheets.length, hasSharedStrings)),
  );

  // xl/styles.xml
  zip.add("xl/styles.xml", encoder.encode(styles.toXml()));

  // xl/sharedStrings.xml
  if (hasSharedStrings) {
    zip.add("xl/sharedStrings.xml", encoder.encode(writeSharedStringsXml(sharedStrings)));
  }

  // xl/worksheets/sheetN.xml + rels + drawings + comments + tables
  for (let i = 0; i < worksheetResults.length; i++) {
    const result = worksheetResults[i];
    const drawing = drawingResults[i];
    const comments = commentsResults[i];

    zip.add(`xl/worksheets/sheet${i + 1}.xml`, encoder.encode(result.xml));

    // Generate worksheet .rels if needed
    const hasHyperlinks = result.hyperlinkRelationships.length > 0;
    const hasDrawing = drawing !== null && result.drawingRId !== null;
    const hasComments = comments !== null && result.legacyDrawingRId !== null;
    const hasTables = result.tables.length > 0;

    if (hasHyperlinks || hasDrawing || hasComments || hasTables) {
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

      if (hasDrawing && result.drawingRId) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.drawingRId,
            Type: REL_DRAWING,
            Target: `../drawings/drawing${i + 1}.xml`,
          }),
        );
      }

      if (hasComments && result.legacyDrawingRId && result.commentsRId) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.legacyDrawingRId,
            Type: REL_VML_DRAWING,
            Target: `../drawings/vmlDrawing${i + 1}.vml`,
          }),
        );
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.commentsRId,
            Type: REL_COMMENTS,
            Target: `../comments${i + 1}.xml`,
          }),
        );
      }

      for (const tableEntry of result.tables) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: tableEntry.rId,
            Type: REL_TABLE,
            Target: `../tables/table${tableEntry.globalTableIndex}.xml`,
          }),
        );
      }

      const relsXml = xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, relElements);
      zip.add(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, encoder.encode(relsXml));
    }

    // Add drawing files
    if (drawing) {
      zip.add(`xl/drawings/drawing${i + 1}.xml`, encoder.encode(drawing.drawingXml));
      zip.add(`xl/drawings/_rels/drawing${i + 1}.xml.rels`, encoder.encode(drawing.drawingRels));
      for (const img of drawing.images) {
        zip.add(img.path, img.data, { compress: false });
      }
    }

    // Add comments and VML drawing files
    if (comments) {
      zip.add(`xl/comments${i + 1}.xml`, encoder.encode(comments.commentsXml));
      zip.add(`xl/drawings/vmlDrawing${i + 1}.vml`, encoder.encode(comments.vmlXml));
    }

    // Add table XML files
    const sheet = writeSheets[i];
    if (sheet.tables && sheet.tables.length > 0) {
      for (let t = 0; t < sheet.tables.length; t++) {
        const tableDef = sheet.tables[t];
        const tableEntry = result.tables[t];
        const globalIdx = tableEntry.globalTableIndex;

        let tableRange = tableDef.range;
        if (!tableRange) {
          tableRange = computeTableRange(tableDef, sheet);
        }

        const tableDefWithRange = { ...tableDef, range: tableRange };
        const tableResult = writeTable(tableDefWithRange, globalIdx, globalIdx);
        zip.add(`xl/tables/table${globalIdx}.xml`, encoder.encode(tableResult.tableXml));
      }
    }
  }

  return zip.build();
}

// ── Helpers ──────────────────────────────────────────────────────────

/**
 * Build the full list of named ranges, merging user-defined ranges with
 * auto-generated _xlnm.Print_Area and _xlnm.Print_Titles from sheet pageSetup.
 */
function buildNamedRanges(sheets: WriteSheet[], userRanges?: NamedRange[]): NamedRange[] {
  const result: NamedRange[] = userRanges ? [...userRanges] : [];

  for (const sheet of sheets) {
    const ps = sheet.pageSetup;
    if (!ps) continue;

    if (ps.printArea) {
      result.push({
        name: "_xlnm.Print_Area",
        range: `${sheet.name}!${ps.printArea}`,
        scope: sheet.name,
      });
    }

    const titleParts: string[] = [];
    if (ps.printTitlesRow) {
      titleParts.push(`${sheet.name}!${ps.printTitlesRow}`);
    }
    if (ps.printTitlesColumn) {
      titleParts.push(`${sheet.name}!${ps.printTitlesColumn}`);
    }
    if (titleParts.length > 0) {
      result.push({
        name: "_xlnm.Print_Titles",
        range: titleParts.join(","),
        scope: sheet.name,
      });
    }
  }

  return result;
}

/**
 * Auto-calculate table range from sheet data and table column count.
 */
function computeTableRange(table: import("../_types").TableDefinition, sheet: WriteSheet): string {
  const colCount = table.columns.length;
  let rowCount = 0;

  if (sheet.rows) {
    rowCount = sheet.rows.length;
  } else if (sheet.data) {
    const hasHeaders = sheet.columns?.some((c) => c.header);
    rowCount = sheet.data.length + (hasHeaders ? 1 : 0);
  }

  if (table.showTotalRow) {
    rowCount += 1;
  }

  if (rowCount < 1) rowCount = 1;

  const startCol = colToLetter(0);
  const endCol = colToLetter(colCount - 1);
  return `${startCol}1:${endCol}${rowCount}`;
}
