// ── XLSX Round-Trip Preservation ─────────────────────────────────────
// Read an XLSX file, modify cells, write it back without losing charts,
// images, macros, shapes, or other features that defter doesn't natively
// understand.

import type { Sheet, Workbook, ReadOptions, WriteSheet, NamedRange } from "../_types";
import { readXlsx } from "./reader";
import { ZipReader } from "../zip/reader";
import { ZipWriter } from "../zip/writer";
import { writeContentTypes } from "./content-types-writer";
import type { ContentTypesOptions } from "./content-types-writer";
import {
  writeRootRels,
  writeWorkbookXml,
  writeWorkbookRels,
  type PivotCacheRef,
  type PivotCacheRel,
} from "./workbook-writer";
import { parseRelationships } from "./relationships";
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
  /**
   * Whether the workbook contains VBA macros (xl/vbaProject.bin).
   * When true, saveXlsx uses XLSM content types
   * (`application/vnd.ms-excel.sheet.macroEnabled.12`).
   * The output should be saved with an `.xlsm` extension.
   */
  hasMacros?: boolean;
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
const REL_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer";
const REL_TIMELINE = "http://schemas.microsoft.com/office/2011/relationships/timeline";
const REL_THREADED_COMMENT =
  "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
const REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person";
const REL_PIVOT_TABLE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

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

  // 4. Detect VBA macros
  const hasMacros = rawEntries.has("xl/vbaProject.bin");

  // 5. Build RoundtripWorkbook
  const rtWorkbook: RoundtripWorkbook = {
    ...workbook,
    _rawEntries: rawEntries,
    _modifiedParts: new Set<string>(),
    _contentTypes: contentTypes,
    _rootRels: rootRels,
    hasMacros,
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

  // Detect Excel 365 threaded comments + persons surviving in raw entries.
  // Threaded comments live at xl/threadedComments/threadedCommentN.xml where
  // N matches the worksheet's 1-based index, so we just probe each sheet.
  const threadedCommentSheetIndices: number[] = [];
  for (let i = 0; i < worksheetResults.length; i++) {
    const probe = `xl/threadedComments/threadedComment${i + 1}.xml`;
    if (workbook._rawEntries.has(probe)) threadedCommentSheetIndices.push(i + 1);
  }
  const hasPersons = workbook._rawEntries.has("xl/persons/person.xml");

  // Collect external link parts that survived in the raw entries.
  // Roundtrip preserves the externalLinkN.xml bodies and their _rels;
  // the workbook.xml + workbook.xml.rels are regenerated and need to
  // re-declare each link so Excel keeps the references.
  const externalLinkIndices: number[] = [];
  for (const path of workbook._rawEntries.keys()) {
    const m = path.match(/^xl\/externalLinks\/externalLink(\d+)\.xml$/i);
    if (m) externalLinkIndices.push(parseInt(m[1], 10));
  }
  externalLinkIndices.sort((a, b) => a - b);

  // Collect pivot cache definitions, pivot cache records, and per-sheet
  // pivot tables that survived in raw entries. Each survives as a body
  // file, but the references that wire them into the workbook get
  // regenerated from scratch — so we must declare them here.
  const pivotCacheDefinitionIndices: number[] = [];
  const pivotCacheRecordIndices: number[] = [];
  const pivotTableIndices: number[] = [];
  // Collect slicer cache parts (xl/slicerCaches/slicerCacheN.xml) and
  // timeline cache parts (xl/timelineCaches/timelineCacheN.xml). These
  // live in the workbook rels and must be re-declared so Excel keeps
  // them; workbook.xml also gains an extLst pointing at them.
  const slicerCacheIndices: number[] = [];
  const timelineCacheIndices: number[] = [];
  // Per-sheet slicer / timeline parts. Indices are global across the
  // workbook (xl/slicers/slicer3.xml may belong to sheet2, etc.).
  const slicerIndices: number[] = [];
  const timelineIndices: number[] = [];
  for (const path of workbook._rawEntries.keys()) {
    let m = path.match(/^xl\/pivotCache\/pivotCacheDefinition(\d+)\.xml$/i);
    if (m) {
      pivotCacheDefinitionIndices.push(parseInt(m[1], 10));
      continue;
    }
    m = path.match(/^xl\/pivotCache\/pivotCacheRecords(\d+)\.xml$/i);
    if (m) {
      pivotCacheRecordIndices.push(parseInt(m[1], 10));
      continue;
    }
    m = path.match(/^xl\/pivotTables\/pivotTable(\d+)\.xml$/i);
    if (m) {
      pivotTableIndices.push(parseInt(m[1], 10));
      continue;
    }
    m = path.match(/^xl\/slicerCaches\/slicerCache(\d+)\.xml$/i);
    if (m) {
      slicerCacheIndices.push(parseInt(m[1], 10));
      continue;
    }
    m = path.match(/^xl\/timelineCaches\/timelineCache(\d+)\.xml$/i);
    if (m) {
      timelineCacheIndices.push(parseInt(m[1], 10));
      continue;
    }
    m = path.match(/^xl\/slicers\/slicer(\d+)\.xml$/i);
    if (m) {
      slicerIndices.push(parseInt(m[1], 10));
      continue;
    }
    m = path.match(/^xl\/timelines\/timeline(\d+)\.xml$/i);
    if (m) timelineIndices.push(parseInt(m[1], 10));
  }
  pivotCacheDefinitionIndices.sort((a, b) => a - b);
  pivotCacheRecordIndices.sort((a, b) => a - b);
  pivotTableIndices.sort((a, b) => a - b);
  slicerCacheIndices.sort((a, b) => a - b);
  timelineCacheIndices.sort((a, b) => a - b);
  slicerIndices.sort((a, b) => a - b);
  timelineIndices.sort((a, b) => a - b);

  // rIds for external link relationships: assigned after all
  // sheet/styles/sharedStrings/theme/macros/featurePropertyBag/persons rIds.
  let nextWorkbookRelId = computeExternalLinkRelStart(
    writeSheets.length,
    hasSharedStrings,
    !!workbook.hasMacros,
    false, // featurePropertyBag — not yet roundtripped
    hasPersons,
  );
  const externalLinkRels = externalLinkIndices.map((idx) => ({
    rId: `rId${nextWorkbookRelId++}`,
    target: `externalLinks/externalLink${idx}.xml`,
  }));
  // rIds for pivot cache definition relationships. The same rIds also
  // flow into the workbook.xml `<pivotCaches>` block so cache references
  // stay sound.
  const pivotCacheRels: PivotCacheRel[] = pivotCacheDefinitionIndices.map((idx) => ({
    rId: `rId${nextWorkbookRelId++}`,
    target: `pivotCache/pivotCacheDefinition${idx}.xml`,
  }));
  // Workbook-level pivotCaches block. cacheId is 0-based and must match
  // the per-pivot-table `cacheId` attribute. We assign them in the
  // order the cache definitions appear in the package — the original
  // cacheIds aren't recoverable without parsing workbook.xml here, but
  // Excel only requires cacheId/rId pairings to be self-consistent, so
  // a 0..N-1 sequence is safe for roundtrip.
  const pivotCacheRefs: PivotCacheRef[] = pivotCacheRels.map((rel, i) => ({
    cacheId: i,
    rId: rel.rId,
  }));
  const slicerCacheRels = slicerCacheIndices.map((idx) => ({
    rId: `rId${nextWorkbookRelId++}`,
    target: `slicerCaches/slicerCache${idx}.xml`,
  }));
  const timelineCacheRels = timelineCacheIndices.map((idx) => ({
    rId: `rId${nextWorkbookRelId++}`,
    target: `timelineCaches/timelineCache${idx}.xml`,
  }));

  // Per-sheet slicer / timeline relationships are recovered from each
  // sheet's original rels (xl/worksheets/_rels/sheetN.xml.rels) so the
  // regenerated rels can re-declare them. We only need the (sheetIndex
  // → list of {target}) mapping; rIds are reassigned per sheet below.
  const sheetSlicerTargets = collectSheetCacheTargets(workbook, sheets, "slicer");
  const sheetTimelineTargets = collectSheetCacheTargets(workbook, sheets, "timeline");

  // Map each pivot table to the sheet that hosts it. The mapping is
  // recovered by walking each sheet's original _rels file — that's
  // where Excel stored the pivotTable -> sheet wiring originally.
  const sheetPivotTableTargets = collectSheetPivotTableTargets(workbook, sheets);

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
    threadedCommentSheetIndices:
      threadedCommentSheetIndices.length > 0 ? threadedCommentSheetIndices : undefined,
    hasPersons: hasPersons || undefined,
    externalLinkIndices: externalLinkIndices.length > 0 ? externalLinkIndices : undefined,
    pivotTableIndices: pivotTableIndices.length > 0 ? pivotTableIndices : undefined,
    pivotCacheDefinitionIndices:
      pivotCacheDefinitionIndices.length > 0 ? pivotCacheDefinitionIndices : undefined,
    pivotCacheRecordIndices:
      pivotCacheRecordIndices.length > 0 ? pivotCacheRecordIndices : undefined,
    slicerIndices: slicerIndices.length > 0 ? slicerIndices : undefined,
    slicerCacheIndices: slicerCacheIndices.length > 0 ? slicerCacheIndices : undefined,
    timelineIndices: timelineIndices.length > 0 ? timelineIndices : undefined,
    timelineCacheIndices: timelineCacheIndices.length > 0 ? timelineCacheIndices : undefined,
    hasCoreProps: true,
    hasAppProps: true,
    hasMacros: workbook.hasMacros,
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
        undefined,
        externalLinkRels.length > 0 ? externalLinkRels : undefined,
        pivotCacheRefs.length > 0 ? pivotCacheRefs : undefined,
        slicerCacheRels.length > 0 ? slicerCacheRels : undefined,
        timelineCacheRels.length > 0 ? timelineCacheRels : undefined,
      ),
    ),
  );

  // xl/_rels/workbook.xml.rels
  zip.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(
      writeWorkbookRels(
        writeSheets.length,
        hasSharedStrings,
        workbook.hasMacros,
        false, // hasFeaturePropertyBag — not yet roundtripped
        hasPersons,
        externalLinkRels.length > 0 ? externalLinkRels : undefined,
        pivotCacheRels.length > 0 ? pivotCacheRels : undefined,
        slicerCacheRels.length > 0 ? slicerCacheRels : undefined,
        timelineCacheRels.length > 0 ? timelineCacheRels : undefined,
      ),
    ),
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
    const slicerTargets = sheetSlicerTargets[i] ?? [];
    const timelineTargets = sheetTimelineTargets[i] ?? [];
    const hasSlicers = slicerTargets.length > 0;
    const hasTimelines = timelineTargets.length > 0;
    const hasThreadedComments = threadedCommentSheetIndices.includes(i + 1);
    const sheetPivotTargets = sheetPivotTableTargets[i] ?? [];
    const hasSheetPivotTables = sheetPivotTargets.length > 0;

    if (
      hasHyperlinks ||
      hasDrawing ||
      hasComments ||
      hasTables ||
      hasSlicers ||
      hasTimelines ||
      hasThreadedComments ||
      hasSheetPivotTables
    ) {
      const relElements: string[] = [];
      // Track the highest existing rId so newly added slicer/timeline
      // relationships pick a number that doesn't collide with anything
      // the worksheet writer already assigned.
      let nextSheetRid = 1;
      const bumpToAfter = (rId: string): void => {
        const m = rId.match(/(\d+)$/);
        if (m) {
          const n = parseInt(m[1], 10);
          if (n + 1 > nextSheetRid) nextSheetRid = n + 1;
        }
      };

      for (const rel of result.hyperlinkRelationships) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: rel.id,
            Type: REL_HYPERLINK,
            Target: rel.target,
            TargetMode: "External",
          }),
        );
        bumpToAfter(rel.id);
      }

      if (hasDrawing && result.drawingRId) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.drawingRId,
            Type: REL_DRAWING,
            Target: `../drawings/drawing${i + 1}.xml`,
          }),
        );
        bumpToAfter(result.drawingRId);
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
        bumpToAfter(result.legacyDrawingRId);
        bumpToAfter(result.commentsRId);
      }

      for (const tableEntry of result.tables) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: tableEntry.rId,
            Type: REL_TABLE,
            Target: `../tables/table${tableEntry.globalTableIndex}.xml`,
          }),
        );
        bumpToAfter(tableEntry.rId);
      }

      // Re-emit slicer relationships read from the original sheet rels.
      // The rIds shift to avoid collisions; they don't need to match the
      // original because hucre regenerates the worksheet body without
      // the `<x14:slicerList>` extension that referenced them.
      for (const target of slicerTargets) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: `rId${nextSheetRid++}`,
            Type: REL_SLICER,
            Target: target,
          }),
        );
      }

      for (const target of timelineTargets) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: `rId${nextSheetRid++}`,
            Type: REL_TIMELINE,
            Target: target,
          }),
        );
      }

      // Threaded comments (Excel 365). The rId only needs to be unique
      // within this rels file — `nextSheetRid` already tracks the next
      // free rId past every relationship emitted above (including the
      // slicer/timeline ones).
      if (hasThreadedComments) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: `rId${nextSheetRid++}`,
            Type: REL_THREADED_COMMENT,
            Target: `../threadedComments/threadedComment${i + 1}.xml`,
          }),
        );
      }

      // Pivot tables hosted on this sheet. Re-emit each one we recovered
      // from the original sheet rels — Excel needs the sheet -> pivot
      // table wiring or it won't render the pivot at all.
      for (const target of sheetPivotTargets) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: `rId${nextSheetRid++}`,
            Type: REL_PIVOT_TABLE,
            Target: target,
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

const REL_TYPE_SLICER = /\/relationships\/slicer$/;
const REL_TYPE_TIMELINE = /\/relationships\/timeline$/;

/**
 * Walk each sheet's original `xl/worksheets/_rels/sheetN.xml.rels` (when
 * present in the raw entries) and pull out the `Target` of every slicer
 * or timeline relationship so the regenerated rels can re-emit them
 * pointing at the same parts.
 *
 * Targets are returned relative to the sheet rels file (e.g.
 * `"../slicers/slicer1.xml"`).
 */
function collectSheetCacheTargets(
  workbook: { _rawEntries: Map<string, Uint8Array> },
  sheets: Sheet[],
  kind: "slicer" | "timeline",
): string[][] {
  const decoder = new TextDecoder("utf-8");
  const out: string[][] = [];
  const matcher = kind === "slicer" ? REL_TYPE_SLICER : REL_TYPE_TIMELINE;
  for (let i = 0; i < sheets.length; i++) {
    // Sheet rels are emitted as xl/worksheets/_rels/sheetN.xml.rels in
    // the regenerated output, but the original file may use a different
    // case — match case-insensitively.
    const expected = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
    let bytes: Uint8Array | undefined;
    for (const [path, data] of workbook._rawEntries) {
      if (path.toLowerCase() === expected) {
        bytes = data;
        break;
      }
    }
    if (!bytes) {
      out.push([]);
      continue;
    }
    const rels = parseRelationships(decoder.decode(bytes));
    const targets: string[] = [];
    for (const rel of rels) {
      if (matcher.test(rel.type)) targets.push(rel.target);
    }
    out.push(targets);
  }
  return out;
}

/**
 * Walk each preserved sheet rels file and pull out its pivot table
 * targets so we can re-emit the sheet -> pivot wiring after the rels
 * are regenerated. Returns one entry per sheet (in workbook order),
 * each a list of relative `Target` paths suitable for plugging back
 * into the regenerated rels.
 */
function collectSheetPivotTableTargets(
  workbook: RoundtripWorkbook,
  sheets: ReadonlyArray<{ name: string }>,
): string[][] {
  const decoder = new TextDecoder("utf-8");
  const out: string[][] = sheets.map(() => []);
  for (let i = 0; i < sheets.length; i++) {
    const relsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
    const data = workbook._rawEntries.get(relsPath);
    if (!data) continue;
    let rels;
    try {
      rels = parseRelationships(decoder.decode(data));
    } catch {
      continue;
    }
    for (const rel of rels) {
      // Match by suffix to tolerate Strict/Transitional namespaces.
      if (rel.type.endsWith("/pivotTable")) {
        out[i].push(rel.target);
      }
    }
  }
  return out;
}

/**
 * Mirror the `nextRid` counter inside `writeWorkbookRels` to determine
 * the starting rId for external link relationships. Keep this in sync
 * with `writeWorkbookRels` — order is: worksheets, styles, optional
 * sharedStrings, theme, optional vbaProject, optional FeaturePropertyBag,
 * optional persons, then externalLinks.
 */
function computeExternalLinkRelStart(
  sheetCount: number,
  hasSharedStrings: boolean,
  hasMacros: boolean,
  hasFeaturePropertyBag: boolean,
  hasPersons: boolean,
): number {
  let next = sheetCount + 1; // worksheets occupy rId1..rId{sheetCount}
  next++; // styles
  if (hasSharedStrings) next++;
  next++; // theme
  if (hasMacros) next++;
  if (hasFeaturePropertyBag) next++;
  if (hasPersons) next++;
  return next;
}

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
