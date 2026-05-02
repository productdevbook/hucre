// ── XLSX Writer ──────────────────────────────────────────────────────
// Generates valid Office Open XML spreadsheet files (XLSX).

import type {
  CellValue,
  NamedRange,
  WorkbookProperties,
  WriteOptions,
  WriteOutput,
  WriteSheet,
} from "../_types";
import { ZipWriter } from "../zip/writer";
import { writeContentTypes } from "./content-types-writer";
import { writeFeaturePropertyBagXml } from "./feature-property-bag";
import type { ContentTypesOptions } from "./content-types-writer";
import { writeRootRels, writeWorkbookXml, writeWorkbookRels } from "./workbook-writer";
import type { PivotCacheRef, PivotCacheRel } from "./workbook-writer";
import { createStylesCollector } from "./styles-writer";
import { createSharedStrings, writeSharedStringsXml, writeWorksheetXml } from "./worksheet-writer";
import type { WorksheetResult } from "./worksheet-writer";
import { writeDrawing } from "./drawing-writer";
import type { DrawingResult } from "./drawing-writer";
import { writeChart } from "./chart-writer";
import { writeComments } from "./comments-writer";
import type { CommentsResult } from "./comments-writer";
import { writeTable } from "./table-writer";
import { colToLetter } from "./worksheet-writer";
import { writePivotTable as writePivotTableParts, resolvePivotSource } from "./pivot-writer";
import type { PivotWriteResult } from "./pivot-writer";
import { xmlDocument, xmlSelfClose } from "../xml/writer";
import { writeCoreProperties, writeAppProperties, writeCustomProperties } from "./doc-props-writer";
import { writeThemeXml } from "./theme-writer";

const encoder = /* @__PURE__ */ new TextEncoder();

const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_VML_DRAWING =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_PIVOT_TABLE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

/**
 * Promote the first non-empty `sheet.a11y.summary` to
 * `properties.description` when the workbook does not already declare one.
 * This is what screen readers announce when the file is opened.
 */
function effectiveProperties(options: WriteOptions): WorkbookProperties | undefined {
  const props = options.properties;
  if (props?.description) return props;

  for (const sheet of options.sheets) {
    const summary = sheet.a11y?.summary;
    if (summary && summary.trim().length > 0) {
      return { ...props, description: summary };
    }
  }
  return props;
}

/**
 * Write a Workbook to XLSX format.
 * Returns a Uint8Array containing the ZIP archive.
 */
export async function writeXlsx(options: WriteOptions): Promise<WriteOutput> {
  const { sheets, defaultFont, dateSystem, namedRanges, activeSheet, workbookProtection } = options;

  const properties = effectiveProperties(options);

  // Create shared collectors
  const styles = createStylesCollector(defaultFont);
  const sharedStrings = createSharedStrings();

  // Pre-compute global table start indices per sheet
  let globalTableCounter = 1;
  const sheetTableStartIndices: Array<number | undefined> = [];
  for (const sheet of sheets) {
    if (sheet.tables && sheet.tables.length > 0) {
      sheetTableStartIndices.push(globalTableCounter);
      globalTableCounter += sheet.tables.length;
    } else {
      sheetTableStartIndices.push(undefined);
    }
  }

  // Pre-compute global pivot-table start indices per sheet. Pivot
  // tables, cache definitions, and cache records all share the same
  // global numbering because each pivot in Phase 1 owns exactly one
  // cache.
  let globalPivotCounter = 1;
  const sheetPivotStartIndices: Array<number | undefined> = [];
  for (const sheet of sheets) {
    if (sheet.pivotTables && sheet.pivotTables.length > 0) {
      sheetPivotStartIndices.push(globalPivotCounter);
      globalPivotCounter += sheet.pivotTables.length;
    } else {
      sheetPivotStartIndices.push(undefined);
    }
  }

  // Generate worksheet XMLs (also populates styles and shared strings)
  const worksheetResults: WorksheetResult[] = [];
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const result = writeWorksheetXml(
      sheet,
      styles,
      sharedStrings,
      dateSystem,
      sheetTableStartIndices[i],
      options.stringMode === "inline",
      sheetPivotStartIndices[i],
    );
    worksheetResults.push(result);
  }

  const hasSharedStrings = sharedStrings.count() > 0;

  // Generate drawing data for sheets that have images, text boxes, or charts
  const drawingResults: Array<DrawingResult | null> = [];
  const drawingIndices: number[] = [];
  const imageExtensions = new Set<string>();
  let globalImageIndex = 1;
  let globalChartIndex = 1;

  // Per-sheet chart payloads, parallel to sheets[]. Each entry holds the
  // serialized chart bodies the sheet contributes, keyed by global
  // chart number so the ZIP layer can write
  // `xl/charts/chart{n}.xml` and `xl/charts/_rels/chart{n}.xml.rels`.
  type ChartFileEntry = { globalIndex: number; xml: string; rels: string };
  const sheetChartFiles: Array<ChartFileEntry[]> = [];
  const allChartIndices: number[] = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const hasImages = sheet.images && sheet.images.length > 0;
    const hasTextBoxes = sheet.textBoxes && sheet.textBoxes.length > 0;
    const hasCharts = sheet.charts && sheet.charts.length > 0;
    if (hasImages || hasTextBoxes || hasCharts) {
      const result = writeDrawing(
        sheet.images ?? [],
        globalImageIndex,
        sheet.textBoxes,
        sheet.charts,
        globalChartIndex,
      );
      drawingResults.push(result);
      drawingIndices.push(i + 1); // 1-based drawing index matches sheet index

      // Track image extensions and advance global counter
      for (const img of result.images) {
        const ext = img.path.split(".").pop();
        if (ext) imageExtensions.add(ext);
      }
      if (sheet.images) {
        globalImageIndex += sheet.images.length;
      }

      // Generate the chart XML bodies. We do this here rather than
      // inside writeDrawing to keep the drawing layer free of
      // chart-specific OOXML dependencies.
      const chartFiles: ChartFileEntry[] = [];
      if (sheet.charts) {
        for (let c = 0; c < sheet.charts.length; c++) {
          const written = writeChart(sheet.charts[c], sheet.name);
          const idx = globalChartIndex + c;
          chartFiles.push({ globalIndex: idx, xml: written.chartXml, rels: written.chartRels });
          allChartIndices.push(idx);
        }
        globalChartIndex += sheet.charts.length;
      }
      sheetChartFiles.push(chartFiles);
    } else {
      drawingResults.push(null);
      sheetChartFiles.push([]);
    }
  }

  // Track background image paths per sheet (for picture relationships)
  const backgroundImagePaths: Array<string | null> = [];
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (sheet.backgroundImage) {
      // Background images are stored as PNG by default
      const bgPath = `xl/media/image${globalImageIndex}.png`;
      backgroundImagePaths.push(bgPath);
      imageExtensions.add("png");
      globalImageIndex++;
    } else {
      backgroundImagePaths.push(null);
    }
  }

  // Generate comments data for sheets that have comments
  const commentsResults: Array<CommentsResult | null> = [];
  const commentIndices: number[] = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
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

  // ── Pivot Tables ──
  // Build cache + table OOXML for every pivot declared on every sheet.
  // Each pivot owns one cache (Phase 1 — multiple pivots cannot share a
  // cache yet), so the indices line up 1:1.
  interface PivotEntry {
    parts: PivotWriteResult;
    globalIndex: number;
  }
  const allPivotEntries: PivotEntry[] = [];
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (!sheet.pivotTables || sheet.pivotTables.length === 0) continue;
    const startIdx = sheetPivotStartIndices[i];
    if (startIdx === undefined) continue;

    for (let p = 0; p < sheet.pivotTables.length; p++) {
      const pivot = sheet.pivotTables[p];
      const sourceSheetName = pivot.sourceSheet ?? sheet.name;
      const sourceSheet = sheets.find((s) => s.name === sourceSheetName);
      if (!sourceSheet) {
        throw new Error(
          `Pivot "${pivot.name}" sourceSheet "${sourceSheetName}" not found in workbook`,
        );
      }
      const sourceRows = collectSourceRows(sourceSheet);
      const resolved = resolvePivotSource(pivot, sourceSheetName, sourceRows);
      const globalIndex = startIdx + p;
      // cacheId is workbook-wide and 0-based, mirroring Excel's own
      // numbering. It also matches the pivot table's `cacheId` attr.
      const cacheId = globalIndex - 1;
      const parts = writePivotTableParts(pivot, resolved, cacheId, globalIndex);
      allPivotEntries.push({ parts, globalIndex });
    }
  }
  const allPivotIndices = allPivotEntries.map((e) => e.globalIndex);
  const pivotCacheRels: PivotCacheRel[] = allPivotEntries.map((e) => ({
    rId: `rIdPivot${e.globalIndex}`,
    target: `pivotCache/pivotCacheDefinition${e.globalIndex}.xml`,
  }));
  const pivotCacheRefs: PivotCacheRef[] = allPivotEntries.map((e, i) => ({
    cacheId: e.globalIndex - 1,
    rId: pivotCacheRels[i].rId,
  }));

  // Build ZIP archive
  const zip = new ZipWriter();

  // Generate custom properties XML (if any)
  const customPropsXml = writeCustomProperties(properties);
  const hasCustomProps = customPropsXml !== null;

  // [Content_Types].xml
  const hasMacros = options.vbaProject !== undefined && options.vbaProject.length > 0;
  const hasFeaturePropertyBag = styles.hasCheckboxFeature();

  const ctOpts: ContentTypesOptions = {
    sheetCount: sheets.length,
    hasSharedStrings,
    drawingIndices: drawingIndices.length > 0 ? drawingIndices : undefined,
    chartIndices: allChartIndices.length > 0 ? allChartIndices : undefined,
    imageExtensions: imageExtensions.size > 0 ? imageExtensions : undefined,
    commentIndices: commentIndices.length > 0 ? commentIndices : undefined,
    tableIndices: allTableIndices.length > 0 ? allTableIndices : undefined,
    pivotTableIndices: allPivotIndices.length > 0 ? allPivotIndices : undefined,
    pivotCacheDefinitionIndices: allPivotIndices.length > 0 ? allPivotIndices : undefined,
    pivotCacheRecordIndices: allPivotIndices.length > 0 ? allPivotIndices : undefined,
    hasCoreProps: true,
    hasAppProps: true,
    hasCustomProps,
    hasMacros,
    hasFeaturePropertyBag,
  };
  zip.add("[Content_Types].xml", encoder.encode(writeContentTypes(ctOpts)));

  // _rels/.rels (with docProps relationships)
  zip.add(
    "_rels/.rels",
    encoder.encode(writeRootRels({ hasCoreProps: true, hasAppProps: true, hasCustomProps })),
  );

  // docProps/core.xml
  zip.add("docProps/core.xml", encoder.encode(writeCoreProperties(properties)));

  // docProps/app.xml
  zip.add("docProps/app.xml", encoder.encode(writeAppProperties(properties)));

  // docProps/custom.xml (if custom properties exist)
  if (customPropsXml) {
    zip.add("docProps/custom.xml", encoder.encode(customPropsXml));
  }

  // xl/workbook.xml — merge user named ranges with auto-generated print area/titles
  const allNamedRanges = buildNamedRanges(sheets, namedRanges);
  zip.add(
    "xl/workbook.xml",
    encoder.encode(
      writeWorkbookXml(
        sheets,
        allNamedRanges.length > 0 ? allNamedRanges : undefined,
        dateSystem,
        activeSheet,
        workbookProtection,
        undefined,
        pivotCacheRefs.length > 0 ? pivotCacheRefs : undefined,
      ),
    ),
  );

  // xl/_rels/workbook.xml.rels
  zip.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(
      writeWorkbookRels(
        sheets.length,
        hasSharedStrings,
        hasMacros,
        hasFeaturePropertyBag,
        undefined,
        undefined,
        pivotCacheRels.length > 0 ? pivotCacheRels : undefined,
      ),
    ),
  );

  // xl/styles.xml
  zip.add("xl/styles.xml", encoder.encode(styles.toXml()));

  // xl/theme/theme1.xml
  zip.add("xl/theme/theme1.xml", encoder.encode(writeThemeXml()));

  // xl/sharedStrings.xml (if any strings)
  if (hasSharedStrings) {
    zip.add("xl/sharedStrings.xml", encoder.encode(writeSharedStringsXml(sharedStrings)));
  }

  // xl/vbaProject.bin (if macros provided)
  if (hasFeaturePropertyBag) {
    zip.add(
      "xl/featurePropertyBag/featurePropertyBag.xml",
      encoder.encode(writeFeaturePropertyBagXml()),
    );
  }

  if (hasMacros) {
    zip.add("xl/vbaProject.bin", options.vbaProject!);
  }

  // xl/worksheets/sheetN.xml + optional xl/worksheets/_rels/sheetN.xml.rels
  for (let i = 0; i < worksheetResults.length; i++) {
    const result = worksheetResults[i];
    const drawing = drawingResults[i];
    const comments = commentsResults[i];

    zip.add(`xl/worksheets/sheet${i + 1}.xml`, encoder.encode(result.xml));

    // Generate worksheet .rels if there are hyperlinks, a drawing, comments, tables, picture, or pivots
    const hasHyperlinks = result.hyperlinkRelationships.length > 0;
    const hasDrawing = drawing !== null && result.drawingRId !== null;
    const hasComments = comments !== null && result.legacyDrawingRId !== null;
    const hasTables = result.tables.length > 0;
    const hasPicture = result.pictureRId !== null && backgroundImagePaths[i] !== null;
    const hasPivots = result.pivotTables.length > 0;

    if (hasHyperlinks || hasDrawing || hasComments || hasTables || hasPicture || hasPivots) {
      const relElements: string[] = [];

      // Hyperlink relationships
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

      // Drawing relationship
      if (hasDrawing && result.drawingRId) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.drawingRId,
            Type: REL_DRAWING,
            Target: `../drawings/drawing${i + 1}.xml`,
          }),
        );
      }

      // Comments relationships (VML drawing + comments file)
      if (hasComments && result.legacyDrawingRId && result.commentsRId) {
        // Legacy drawing (VML) relationship
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.legacyDrawingRId,
            Type: REL_VML_DRAWING,
            Target: `../drawings/vmlDrawing${i + 1}.vml`,
          }),
        );

        // Comments file relationship
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.commentsRId,
            Type: REL_COMMENTS,
            Target: `../comments${i + 1}.xml`,
          }),
        );
      }

      // Table relationships
      for (const tableEntry of result.tables) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: tableEntry.rId,
            Type: REL_TABLE,
            Target: `../tables/table${tableEntry.globalTableIndex}.xml`,
          }),
        );
      }

      // Background image (picture) relationship
      if (hasPicture && result.pictureRId && backgroundImagePaths[i]) {
        const bgMediaPath = backgroundImagePaths[i]!;
        const relTarget = `../${bgMediaPath.slice(3)}`; // Remove "xl/" prefix → "../media/imageN.png"
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.pictureRId,
            Type: REL_IMAGE,
            Target: relTarget,
          }),
        );
      }

      // Pivot table relationships
      for (const pivotEntry of result.pivotTables) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: pivotEntry.rId,
            Type: REL_PIVOT_TABLE,
            Target: `../pivotTables/pivotTable${pivotEntry.globalPivotIndex}.xml`,
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

      // Add image files to ZIP (store, don't compress — images are already compressed)
      for (const img of drawing.images) {
        zip.add(img.path, img.data, { compress: false });
      }
    }

    // Add chart files for this sheet
    const chartFiles = sheetChartFiles[i];
    if (chartFiles && chartFiles.length > 0) {
      for (const cf of chartFiles) {
        zip.add(`xl/charts/chart${cf.globalIndex}.xml`, encoder.encode(cf.xml));
        zip.add(`xl/charts/_rels/chart${cf.globalIndex}.xml.rels`, encoder.encode(cf.rels));
      }
    }

    // Add background image file to ZIP
    if (hasPicture && backgroundImagePaths[i] && sheets[i].backgroundImage) {
      zip.add(backgroundImagePaths[i]!, sheets[i].backgroundImage!, { compress: false });
    }

    // Add comments and VML drawing files
    if (comments) {
      zip.add(`xl/comments${i + 1}.xml`, encoder.encode(comments.commentsXml));
      zip.add(`xl/drawings/vmlDrawing${i + 1}.vml`, encoder.encode(comments.vmlXml));
    }

    // Add table XML files
    const sheet = sheets[i];
    if (sheet.tables && sheet.tables.length > 0) {
      for (let t = 0; t < sheet.tables.length; t++) {
        const tableDef = sheet.tables[t];
        const tableEntry = result.tables[t];
        const globalIdx = tableEntry.globalTableIndex;

        // Auto-calculate range if not provided
        let tableRange = tableDef.range;
        if (!tableRange) {
          tableRange = computeTableRange(tableDef, sheet);
        }

        // Write table XML with resolved range
        const tableDefWithRange = { ...tableDef, range: tableRange };
        const tableResult = writeTable(tableDefWithRange, globalIdx, globalIdx);
        zip.add(`xl/tables/table${globalIdx}.xml`, encoder.encode(tableResult.tableXml));
      }
    }
  }

  // ── Pivot table parts (cache definition + records + table) ──
  for (const entry of allPivotEntries) {
    const idx = entry.globalIndex;
    zip.add(
      `xl/pivotCache/pivotCacheDefinition${idx}.xml`,
      encoder.encode(entry.parts.cacheDefinitionXml),
    );
    zip.add(
      `xl/pivotCache/_rels/pivotCacheDefinition${idx}.xml.rels`,
      encoder.encode(entry.parts.cacheDefinitionRels),
    );
    zip.add(
      `xl/pivotCache/pivotCacheRecords${idx}.xml`,
      encoder.encode(entry.parts.cacheRecordsXml),
    );
    zip.add(`xl/pivotTables/pivotTable${idx}.xml`, encoder.encode(entry.parts.pivotTableXml));
    zip.add(
      `xl/pivotTables/_rels/pivotTable${idx}.xml.rels`,
      encoder.encode(entry.parts.pivotTableRels),
    );
  }

  return zip.build();
}

// ── Pivot Source Resolution ────────────────────────────────────────────

/**
 * Pull the source data out of a `WriteSheet`. Pivot tables can source
 * from either `rows` (raw 2-D arrays) or `data` (objects keyed by
 * `columns[].key`); we normalise both shapes into a single `CellValue[][]`.
 *
 * Returns `[]` when the sheet has no row-shaped data — `resolvePivotSource`
 * will throw a clearer error in that case.
 */
function collectSourceRows(sheet: WriteSheet): CellValue[][] {
  if (sheet.rows && sheet.rows.length > 0) {
    return sheet.rows.map((row) => [...row]);
  }
  if (sheet.data && sheet.data.length > 0 && sheet.columns && sheet.columns.length > 0) {
    const out: CellValue[][] = [];
    const headerRow: CellValue[] = sheet.columns.map((c) => c.header ?? c.key ?? "");
    out.push(headerRow);
    for (const obj of sheet.data) {
      const row: CellValue[] = sheet.columns.map((c) => {
        if (!c.key) return null;
        const v = obj[c.key];
        return v === undefined ? null : (v as CellValue);
      });
      out.push(row);
    }
    return out;
  }
  return [];
}

// ── Named Range Builder ────────────────────────────────────────────────

/**
 * Build the full list of named ranges, merging user-defined ranges with
 * auto-generated _xlnm.Print_Area and _xlnm.Print_Titles from sheet pageSetup.
 */
function buildNamedRanges(sheets: WriteOptions["sheets"], userRanges?: NamedRange[]): NamedRange[] {
  const result: NamedRange[] = userRanges ? [...userRanges] : [];

  for (const sheet of sheets) {
    const ps = sheet.pageSetup;
    if (!ps) continue;

    // Print area → _xlnm.Print_Area
    if (ps.printArea) {
      result.push({
        name: "_xlnm.Print_Area",
        range: `${sheet.name}!${ps.printArea}`,
        scope: sheet.name,
      });
    }

    // Print titles (repeat rows and/or columns)
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

// ── Table Range Computation ──────────────────────────────────────────

/**
 * Auto-calculate table range from sheet data and table column count.
 * Assumes header row is row 1 and data fills remaining rows.
 */
function computeTableRange(
  table: import("../_types").TableDefinition,
  sheet: import("../_types").WriteSheet,
): string {
  const colCount = table.columns.length;
  let rowCount = 0;

  if (sheet.rows) {
    rowCount = sheet.rows.length;
  } else if (sheet.data) {
    // Object data: data rows + 1 header row (if columns have headers)
    const hasHeaders = sheet.columns?.some((c) => c.header);
    rowCount = sheet.data.length + (hasHeaders ? 1 : 0);
  }

  // Add total row if requested
  if (table.showTotalRow) {
    rowCount += 1;
  }

  // Minimum: 1 header row + 0 data rows = 1 row
  if (rowCount < 1) rowCount = 1;

  const startCol = colToLetter(0);
  const endCol = colToLetter(colCount - 1);
  return `${startCol}1:${endCol}${rowCount}`;
}
