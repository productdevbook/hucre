// ── XLSX Reader ──────────────────────────────────────────────────────
// Reads Office Open XML (.xlsx) spreadsheet files.

import type {
  Sheet,
  SheetKind,
  Workbook,
  ReadOptions,
  ReadInput,
  SheetImage,
  SheetTextBox,
  NamedRange,
  TableDefinition,
  TableColumn,
  ThreadedCommentPerson,
  ExternalLink,
  CellImage,
  PivotCache,
  PivotTable,
  SlicerCache,
  TimelineCache,
  ChartAnchor,
} from "../_types"
import { parsePersons, parseThreadedComments } from "./threaded-comments-reader"
import { parseExternalLink } from "./external-link-reader"
import { assembleCellImages, parseCellImages, REL_CELL_IMAGES } from "./cell-images-reader"
import { attachPivotCacheFields, parsePivotCacheDefinition, parsePivotTable } from "./pivot-reader"
import { parseSlicers, parseSlicerCache, parseTimelines, parseTimelineCache } from "./slicer-reader"
import { parseChart } from "./chart-reader"
import { EncryptedFileError, ParseError, ZipError } from "../errors"
import { isOle2Container, readInputToUint8Array } from "../_input"
import { decryptAgile } from "./crypto/agile"
import { ZipReader } from "../zip/reader"
import { parseXml } from "../xml/parser"
import { parseContentTypes } from "./content-types"
import { parseRelationships } from "./relationships"
import { parseSharedStrings } from "./shared-strings"
import { parseStyles } from "./styles"
import { parseWorksheet } from "./worksheet"
import { parseDynamicArrayCellMetadata } from "./metadata"
import type { ParsedStyles } from "./styles"
import type { SharedString } from "./shared-strings"
import type { Relationship } from "./relationships"
import { parseComments } from "./comments-reader"
import { parseCellRef } from "./worksheet"
import { parseCoreProperties, parseAppProperties, parseCustomProperties } from "./doc-props-reader"
import { parseThemeColors } from "./theme"

// ── OOXML Relationship Types ─────────────────────────────────────────

// Transitional namespace (OOXML 2006/Transitional — most common)
const NS_TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
// Strict namespace (OOXML Strict — Excel 2013+ "Strict Open XML" save mode)
const NS_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships"

/**
 * Match a relationship type against both Transitional and Strict OOXML namespaces.
 * Excel 2013+ can save in Strict mode which uses different namespace URIs.
 */
export function matchesRelType(rel: string, type: string): boolean {
  return (
    rel === `${NS_TRANSITIONAL}/${type}` ||
    rel === `${NS_STRICT}/${type}` ||
    rel.endsWith("/" + type)
  )
}

// ── Helpers ──────────────────────────────────────────────────────────

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data)
}

/**
 * Resolve a relative target path against a base directory.
 * E.g. resolve("xl/_rels", "../worksheets/sheet1.xml") → "xl/worksheets/sheet1.xml"
 */
function resolvePath(base: string, target: string): string {
  // If target starts with /, it's absolute from the package root
  if (target.startsWith("/")) return target.slice(1)

  const baseParts = base.split("/").filter(Boolean)
  const targetParts = target.split("/").filter(Boolean)

  for (const part of targetParts) {
    if (part === "..") {
      baseParts.pop()
    } else if (part !== ".") {
      baseParts.push(part)
    }
  }

  return baseParts.join("/")
}

/**
 * Get the directory portion of a path.
 * E.g. "xl/workbook.xml" → "xl"
 */
function dirname(path: string): string {
  const idx = path.lastIndexOf("/")
  return idx === -1 ? "" : path.slice(0, idx)
}

// ── Main Reader ──────────────────────────────────────────────────────

/**
 * Read an XLSX file and return a Workbook.
 * Input can be Uint8Array, ArrayBuffer, or ReadableStream&lt;Uint8Array&gt;.
 *
 * For ReadableStream input, the stream is fully buffered before parsing
 * because the ZIP central directory lives at the end of the archive —
 * true streaming requires random access. Use {@link streamXlsxRows} when
 * you need row-level streaming with low per-row memory.
 */
export async function readXlsx(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  let data = await readInputToUint8Array(input, options?.maxInputBytes)

  // Password-protected workbooks arrive as an OLE2/CFB envelope. With a
  // password we decrypt the inner OOXML ZIP and continue; without one we
  // surface a typed `EncryptedFileError` instead of a generic ZIP failure.
  if (isOle2Container(data)) {
    if (options?.password) {
      data = await decryptAgile(data, options.password, options.maxSpinCount)
    } else {
      throw new EncryptedFileError("xlsx")
    }
  }

  // 1. Open ZIP archive
  let zip: ZipReader
  try {
    zip = new ZipReader(data, options?.maxDecompressedBytes)
  } catch (err) {
    if (err instanceof ZipError) throw err
    throw new ParseError("Failed to open XLSX file: not a valid ZIP archive", undefined, {
      cause: err,
    })
  }

  // 2. Parse [Content_Types].xml (validate it exists)
  if (!zip.has("[Content_Types].xml")) {
    throw new ParseError("Invalid XLSX: missing [Content_Types].xml")
  }
  const contentTypesXml = decodeUtf8(await zip.extract("[Content_Types].xml"))
  parseContentTypes(contentTypesXml) // Validate, not strictly needed for reading

  // 3. Parse _rels/.rels to find the workbook path
  if (!zip.has("_rels/.rels")) {
    throw new ParseError("Invalid XLSX: missing _rels/.rels")
  }
  const rootRelsXml = decodeUtf8(await zip.extract("_rels/.rels"))
  const rootRels = parseRelationships(rootRelsXml)
  const workbookRel = rootRels.find((r) => matchesRelType(r.type, "officeDocument"))
  if (!workbookRel) {
    throw new ParseError("Invalid XLSX: cannot find workbook relationship in _rels/.rels")
  }

  const workbookPath = workbookRel.target.startsWith("/")
    ? workbookRel.target.slice(1)
    : workbookRel.target

  // 4. Parse workbook relationships (xl/_rels/workbook.xml.rels)
  const workbookDir = dirname(workbookPath)
  const workbookRelsPath = workbookDir
    ? `${workbookDir}/_rels/${workbookPath.slice(workbookDir.length + 1)}.rels`
    : `_rels/${workbookPath}.rels`

  let workbookRels: Relationship[] = []
  if (zip.has(workbookRelsPath)) {
    const wbRelsXml = decodeUtf8(await zip.extract(workbookRelsPath))
    workbookRels = parseRelationships(wbRelsXml)
  }

  // 5. Parse xl/workbook.xml for sheet names, order, and date system
  if (!zip.has(workbookPath)) {
    throw new ParseError(`Invalid XLSX: missing workbook at ${workbookPath}`)
  }
  const workbookXml = decodeUtf8(await zip.extract(workbookPath))
  const {
    sheets: sheetInfos,
    dateSystem,
    namedRanges,
    workbookProtection,
    activeSheet,
    pivotCacheRefs,
  } = parseWorkbookXml(workbookXml, options)

  // 6. Parse shared strings if present
  let sharedStrings: SharedString[] = []
  const ssRel = workbookRels.find((r) => matchesRelType(r.type, "sharedStrings"))
  if (ssRel) {
    const ssPath = resolvePath(workbookDir, ssRel.target)
    if (zip.has(ssPath)) {
      const ssXml = decodeUtf8(await zip.extract(ssPath))
      sharedStrings = parseSharedStrings(ssXml)
    }
  }

  // 7. Parse styles if needed (for date detection or if readStyles is true)
  let parsedStyles: ParsedStyles | null = null
  const stylesRel = workbookRels.find((r) => matchesRelType(r.type, "styles"))
  if (stylesRel) {
    const stylesPath = resolvePath(workbookDir, stylesRel.target)
    if (zip.has(stylesPath)) {
      const stylesXml = decodeUtf8(await zip.extract(stylesPath))
      parsedStyles = parseStyles(stylesXml)
    }
  }

  // 7b. Parse theme colors if theme1.xml exists
  let themeColors: string[] | undefined
  const themePath = workbookDir ? `${workbookDir}/theme/theme1.xml` : "theme/theme1.xml"
  if (zip.has(themePath)) {
    const themeXml = decodeUtf8(await zip.extract(themePath))
    themeColors = parseThemeColors(themeXml)
  }

  // 7c. Parse the workbook-wide threaded-comments person directory
  // (xl/persons/person.xml). Linked from workbook.xml.rels by Type=".../person".
  let persons: ThreadedCommentPerson[] | undefined
  const personsRel = workbookRels.find((r) => matchesRelType(r.type, "person"))
  if (personsRel) {
    const personsPath = resolvePath(workbookDir, personsRel.target)
    if (zip.has(personsPath)) {
      const personsXml = decodeUtf8(await zip.extract(personsPath))
      persons = parsePersons(personsXml)
    }
  }

  // 7d. Parse external workbook links (xl/externalLinks/externalLinkN.xml).
  // The workbook.xml.rels file declares them with Type=".../externalLink";
  // resolve each one in declaration order so the index lines up with
  // the `[N]` prefix used in formulas.
  const externalLinkRels = workbookRels
    .filter((r) => matchesRelType(r.type, "externalLink"))
    .sort((a, b) => relIdNum(a.id) - relIdNum(b.id))
  const externalLinks: ExternalLink[] = []
  for (const rel of externalLinkRels) {
    const linkPath = resolvePath(workbookDir, rel.target)
    if (!zip.has(linkPath)) continue
    const linkXml = decodeUtf8(await zip.extract(linkPath))
    const linkRelsPath = relsPathFor(linkPath)
    const linkRelsXml = zip.has(linkRelsPath)
      ? decodeUtf8(await zip.extract(linkRelsPath))
      : undefined
    externalLinks.push(parseExternalLink(linkXml, linkRelsXml))
  }

  // 7e. Parse WPS-style cell-embedded images (xl/cellimages.xml).
  // The workbook.xml.rels file declares the part with a WPS namespace
  // type; cells reference each entry through `=_xlfn.DISPIMG("<id>", 1)`
  // formulas. Real-world XLSX files produced by WPS / Kingsoft Office
  // routinely carry this part and recent Excel versions round-trip it.
  let cellImages: CellImage[] | undefined
  const cellImagesRel = workbookRels.find((r) => r.type === REL_CELL_IMAGES)
  if (cellImagesRel) {
    const cellImagesPath = resolvePath(workbookDir, cellImagesRel.target)
    if (zip.has(cellImagesPath)) {
      const ciXml = decodeUtf8(await zip.extract(cellImagesPath))
      const refs = parseCellImages(ciXml)

      // Resolve each embed rId against the sibling _rels file and
      // pre-load the binaries so `assembleCellImages` stays sync.
      const ciRelsPath = relsPathFor(cellImagesPath)
      const ciDir = dirname(cellImagesPath)
      const media = new Map<string, { data: Uint8Array; type: SheetImage["type"] }>()
      if (zip.has(ciRelsPath)) {
        const ciRelsXml = decodeUtf8(await zip.extract(ciRelsPath))
        for (const rel of parseRelationships(ciRelsXml)) {
          if (!matchesRelType(rel.type, "image")) continue
          const mediaPath = resolvePath(ciDir, rel.target)
          if (!zip.has(mediaPath)) continue
          const ext = mediaPath.split(".").pop()?.toLowerCase() ?? ""
          const type = EXT_TO_IMAGE_TYPE[ext]
          if (!type) continue
          media.set(rel.id, { data: await zip.extract(mediaPath), type })
        }
      }

      const assembled = assembleCellImages(refs, media)
      if (assembled.length > 0) cellImages = assembled
    }
  }

  // 7f. Parse pivot caches (xl/pivotCache/pivotCacheDefinitionN.xml).
  // The workbook's <pivotCaches> block ties each cacheId to an rId in
  // workbook.xml.rels; we walk that pairing and resolve each cache.
  // The cache definitions also surface a `hasRecords` flag based on
  // whether the sibling rels declare a pivotCacheRecords part.
  const pivotCachesByCacheId = new Map<number, PivotCache>()
  const pivotCachesByRId = new Map<string, PivotCache>()
  for (const ref of pivotCacheRefs) {
    const rel = workbookRels.find((r) => r.id === ref.rId)
    if (!rel) continue
    const cachePath = resolvePath(workbookDir, rel.target)
    if (!zip.has(cachePath)) continue
    const cacheXml = decodeUtf8(await zip.extract(cachePath))
    const cache = parsePivotCacheDefinition(cacheXml)
    if (!cache) continue
    cache.cacheId = ref.cacheId
    // Detect a sibling pivotCacheRecords part via the cache's _rels.
    const cacheRelsPath = relsPathFor(cachePath)
    if (zip.has(cacheRelsPath)) {
      const cacheRelsXml = decodeUtf8(await zip.extract(cacheRelsPath))
      const cacheRels = parseRelationships(cacheRelsXml)
      if (cacheRels.some((r) => matchesRelType(r.type, "pivotCacheRecords"))) {
        cache.hasRecords = true
      }
    }
    pivotCachesByCacheId.set(ref.cacheId, cache)
    pivotCachesByRId.set(ref.rId, cache)
  }
  // Preserve the workbook's declaration order so caller-side cacheId
  // lookups by array index match what Excel sees.
  const pivotCaches: PivotCache[] = pivotCacheRefs
    .map((ref) => pivotCachesByCacheId.get(ref.cacheId))
    .filter((c): c is PivotCache => c !== undefined)

  // 7g. Parse slicer caches (xl/slicerCaches/slicerCacheN.xml).
  // Excel 2010+ slicers — declared from workbook.xml.rels with
  // Type=".../slicerCache". Sort by trailing index so the array order is
  // stable across files.
  const slicerCacheRels = workbookRels
    .filter((r) => matchesRelType(r.type, "slicerCache"))
    .sort((a, b) => slicerNumFromTarget(a.target) - slicerNumFromTarget(b.target))
  const slicerCaches: SlicerCache[] = []
  for (const rel of slicerCacheRels) {
    const cachePath = resolvePath(workbookDir, rel.target)
    if (!zip.has(cachePath)) continue
    const cacheXml = decodeUtf8(await zip.extract(cachePath))
    const cache = parseSlicerCache(cacheXml)
    if (cache) slicerCaches.push(cache)
  }

  // 7h. Parse timeline caches (xl/timelineCaches/timelineCacheN.xml).
  // Excel 2013+ timeline slicers — Type=".../timelineCache".
  const timelineCacheRels = workbookRels
    .filter((r) => matchesRelType(r.type, "timelineCache"))
    .sort((a, b) => slicerNumFromTarget(a.target) - slicerNumFromTarget(b.target))
  const timelineCaches: TimelineCache[] = []
  for (const rel of timelineCacheRels) {
    const cachePath = resolvePath(workbookDir, rel.target)
    if (!zip.has(cachePath)) continue
    const cacheXml = decodeUtf8(await zip.extract(cachePath))
    const cache = parseTimelineCache(cacheXml)
    if (cache) timelineCaches.push(cache)
  }

  // 7i. Parse cell metadata (xl/metadata.xml). Cells point into its
  // cellMetadata collection with `cm`; for hucre the only records that
  // matter are the dynamic-array (XLDAPR) ones. Resolved once for the
  // whole package because the part is workbook-level.
  let dynamicArrayCm: Set<number> | undefined
  const metadataRel = workbookRels.find((r) => matchesRelType(r.type, "sheetMetadata"))
  if (metadataRel) {
    const metadataPath = resolvePath(workbookDir, metadataRel.target)
    if (zip.has(metadataPath)) {
      dynamicArrayCm = parseDynamicArrayCellMetadata(decodeUtf8(await zip.extract(metadataPath)))
    }
  }

  // 8. Build a map of rId → sheet relationship for worksheet paths
  const sheetRelMap = new Map<string, string>()
  // …and one for the tabs that are not worksheets. `<sheets>` lists
  // every tab whatever its kind and the relationship type is what tells
  // them apart, so a chart sheet's rId simply never entered the map
  // above — and the lookup below threw, taking every ordinary worksheet
  // in the book down with it. See #499.
  const nonWorksheetKinds = new Map<string, SheetKind>()
  for (const rel of workbookRels) {
    if (matchesRelType(rel.type, "worksheet")) {
      sheetRelMap.set(rel.id, resolvePath(workbookDir, rel.target))
    } else if (matchesRelType(rel.type, "chartsheet")) {
      nonWorksheetKinds.set(rel.id, "chartsheet")
    } else if (matchesRelType(rel.type, "dialogsheet")) {
      nonWorksheetKinds.set(rel.id, "dialogsheet")
    }
  }

  // 9. Filter sheets if options specify which ones to read
  const sheetsToRead = filterSheets(sheetInfos, options?.sheets)

  // 10. Parse each worksheet
  const readStyles = options?.readStyles ?? false

  const sheets = []
  for (const info of sheetsToRead) {
    // A tab that is not a worksheet has no cells to read. It becomes an
    // empty sheet rather than being skipped, so `sheets: [2]` still
    // selects what Excel's third tab is — renumbering would be a quieter
    // kind of wrong.
    const kind = nonWorksheetKinds.get(info.rId)
    if (kind !== undefined) {
      const placeholder: Sheet = { name: info.name, rows: [], kind }
      if (info.state === "hidden") placeholder.hidden = true
      if (info.state === "veryHidden") placeholder.veryHidden = true
      sheets.push(placeholder)
      continue
    }

    const wsPath = sheetRelMap.get(info.rId)
    if (!wsPath || !zip.has(wsPath)) {
      throw new ParseError(`Invalid XLSX: missing worksheet file for sheet "${info.name}"`)
    }

    // Check for worksheet-level relationships (hyperlinks, etc.).
    // relsPathFor handles a root-level part; the hand-rolled slice here
    // ate the first character when dirname returned "", so sheet1.xml
    // looked for _rels/heet1.xml.rels and every relationship-reached
    // feature silently vanished. See #391.
    const wsDir = dirname(wsPath)
    const wsRelsPath = relsPathFor(wsPath)
    let worksheetRels: Relationship[] | undefined
    if (zip.has(wsRelsPath)) {
      const wsRelsXml = decodeUtf8(await zip.extract(wsRelsPath))
      worksheetRels = parseRelationships(wsRelsXml)
    }

    const worksheetCtx = {
      sharedStrings,
      styles: parsedStyles,
      readStyles,
      dateSystem,
      worksheetRels,
      maxRows: options?.maxRows,
      range: options?.range,
      maxTotalCells: options?.maxTotalCells,
      sparse: options?.sparse,
      dynamicArrayCm,
      sheetName: info.name,
      onWarning: options?.onWarning,
    }

    const wsXml = decodeUtf8(await zip.extract(wsPath))
    const sheet = parseWorksheet(wsXml, info.name, worksheetCtx)
    if (info.state === "hidden") sheet.hidden = true
    if (info.state === "veryHidden") sheet.veryHidden = true

    // Extract images and textboxes from drawing if present
    if (worksheetRels) {
      const drawingRel = worksheetRels.find((r) => matchesRelType(r.type, "drawing"))
      if (drawingRel) {
        const drawingPath = resolvePath(wsDir, drawingRel.target)
        const drawing = await extractSheetDrawing(zip, drawingPath)
        if (drawing.images.length > 0) {
          sheet.images = drawing.images
        }
        if (drawing.textBoxes.length > 0) {
          sheet.textBoxes = drawing.textBoxes
        }
        // Resolve chart parts referenced from the drawing. Each
        // graphicFrame's chart rel was resolved against the drawing's
        // package directory in extractSheetDrawing; here we simply
        // parse the chart bodies and pin each chart back to the cell
        // anchor reported by the drawing layer.
        if (drawing.chartRefs.length > 0) {
          const charts: import("../_types").Chart[] = []
          for (const chartRef of drawing.chartRefs) {
            if (!zip.has(chartRef.path)) continue
            const chartXml = decodeUtf8(await zip.extract(chartRef.path))
            const chart = parseChart(chartXml)
            if (!chart) continue
            if (chartRef.anchor) chart.anchor = chartRef.anchor
            if (chartRef.altText) chart.altText = chartRef.altText
            if (chartRef.frameTitle) chart.frameTitle = chartRef.frameTitle
            charts.push(chart)
          }
          if (charts.length > 0) sheet.charts = charts
        }
      }
    }

    // Extract comments if present
    if (worksheetRels) {
      const commentsRel = worksheetRels.find((r) => matchesRelType(r.type, "comments"))
      if (commentsRel) {
        const commentsPath = resolvePath(wsDir, commentsRel.target)
        if (zip.has(commentsPath)) {
          const commentsXml = decodeUtf8(await zip.extract(commentsPath))
          const commentsMap = parseComments(commentsXml)

          // Attach comments to cell objects
          if (commentsMap.size > 0) {
            if (!sheet.cells) {
              sheet.cells = new Map()
            }
            for (const [cellRefStr, comment] of commentsMap) {
              const pos = parseCellRef(cellRefStr)
              const key = `${pos.row},${pos.col}`
              let cell = sheet.cells.get(key)
              if (!cell) {
                cell = {
                  value: (sheet.rows[pos.row] && sheet.rows[pos.row][pos.col]) ?? null,
                  type: "string",
                }
                sheet.cells.set(key, cell)
              }
              cell.comment = comment
            }
          }
        }
      }

      // Extract Excel 365 threaded comments if present.
      // Sheets can have BOTH legacy comments and threaded comments — Excel
      // writes a legacy stub for backward compat, so we treat them as
      // independent surfaces rather than overwriting each other.
      const threadedRel = worksheetRels.find((r) => matchesRelType(r.type, "threadedComment"))
      if (threadedRel) {
        const tcPath = resolvePath(wsDir, threadedRel.target)
        if (zip.has(tcPath)) {
          const tcXml = decodeUtf8(await zip.extract(tcPath))
          const threaded = parseThreadedComments(tcXml)
          if (threaded.length > 0) sheet.threadedComments = threaded
        }
      }
    }

    // Extract tables if present
    if (worksheetRels) {
      const tableRels = worksheetRels.filter((r) => matchesRelType(r.type, "table"))
      if (tableRels.length > 0) {
        const tables: TableDefinition[] = []
        for (const tableRel of tableRels) {
          const tablePath = resolvePath(wsDir, tableRel.target)
          if (zip.has(tablePath)) {
            const tableXml = decodeUtf8(await zip.extract(tablePath))
            const tableDef = parseTableXml(tableXml)
            if (tableDef) {
              tables.push(tableDef)
            }
          }
        }
        if (tables.length > 0) {
          sheet.tables = tables
        }
      }
    }

    // Extract background image (picture relationship) if present
    if (worksheetRels) {
      const pictureRel = worksheetRels.find((r) => matchesRelType(r.type, "image"))
      if (pictureRel) {
        const picturePath = resolvePath(wsDir, pictureRel.target)
        if (zip.has(picturePath)) {
          sheet.backgroundImage = await zip.extract(picturePath)
        }
      }
    }

    // Extract pivot tables hosted on this sheet, if any. The sheet's
    // _rels file declares `Type=".../pivotTable"` per instance; the
    // body lives in xl/pivotTables/pivotTableN.xml and points at the
    // owning cache through its sibling _rels file.
    if (worksheetRels) {
      const pivotTableRels = worksheetRels.filter((r) => matchesRelType(r.type, "pivotTable"))
      if (pivotTableRels.length > 0) {
        const pivotTables: PivotTable[] = []
        for (const ptRel of pivotTableRels) {
          const ptPath = resolvePath(wsDir, ptRel.target)
          if (!zip.has(ptPath)) continue
          const ptXml = decodeUtf8(await zip.extract(ptPath))
          const pivot = parsePivotTable(ptXml)
          if (!pivot) continue
          // Resolve the pivot's owning cache via its sibling _rels —
          // pivot tables don't carry the rId in the body, only in the
          // companion .rels file. Match by relative target so we
          // stay tolerant of caches living anywhere under xl/.
          const ptRelsPath = relsPathFor(ptPath)
          if (zip.has(ptRelsPath)) {
            const ptRelsXml = decodeUtf8(await zip.extract(ptRelsPath))
            const ptInternalRels = parseRelationships(ptRelsXml)
            const cacheRel = ptInternalRels.find((r) =>
              matchesRelType(r.type, "pivotCacheDefinition"),
            )
            if (cacheRel) {
              const resolvedCachePath = resolvePath(dirname(ptPath), cacheRel.target)
              for (const ref of pivotCacheRefs) {
                const wbRel = workbookRels.find((r) => r.id === ref.rId)
                if (!wbRel) continue
                const wbCachePath = resolvePath(workbookDir, wbRel.target)
                if (wbCachePath === resolvedCachePath) {
                  pivot.cacheId = ref.cacheId
                  break
                }
              }
            }
          }
          // Overlay the cache's field names so consumers see the real
          // names instead of the synthetic field1/field2 placeholders
          // emitted by parsePivotTable when it has no cache context.
          const owningCache = pivotCachesByCacheId.get(pivot.cacheId)
          if (owningCache) attachPivotCacheFields(pivot, owningCache)
          pivotTables.push(pivot)
        }
        if (pivotTables.length > 0) sheet.pivotTables = pivotTables
      }

      // Extract slicers attached to this sheet (Excel 2010+).
      // The worksheet rels declare them with Type=".../slicer".
      const slicerRels = worksheetRels.filter((r) => matchesRelType(r.type, "slicer"))
      if (slicerRels.length > 0) {
        const slicers: import("../_types").Slicer[] = []
        for (const rel of slicerRels) {
          const path = resolvePath(wsDir, rel.target)
          if (!zip.has(path)) continue
          const slicerXml = decodeUtf8(await zip.extract(path))
          for (const s of parseSlicers(slicerXml)) slicers.push(s)
        }
        if (slicers.length > 0) sheet.slicers = slicers
      }

      // Timeline slicers (Excel 2013+) — Type=".../timeline".
      const timelineRels = worksheetRels.filter((r) => matchesRelType(r.type, "timeline"))
      if (timelineRels.length > 0) {
        const timelines: import("../_types").Timeline[] = []
        for (const rel of timelineRels) {
          const path = resolvePath(wsDir, rel.target)
          if (!zip.has(path)) continue
          const tlXml = decodeUtf8(await zip.extract(path))
          for (const t of parseTimelines(tlXml)) timelines.push(t)
        }
        if (timelines.length > 0) sheet.timelines = timelines
      }
    }

    sheets.push(sheet)
  }

  // 11. Parse document properties (if present)
  let properties: import("../_types").WorkbookProperties | undefined

  if (zip.has("docProps/core.xml")) {
    const coreXml = decodeUtf8(await zip.extract("docProps/core.xml"))
    const coreProps = parseCoreProperties(coreXml)
    if (Object.keys(coreProps).length > 0) {
      properties = { ...coreProps }
    }
  }

  if (zip.has("docProps/app.xml")) {
    const appXml = decodeUtf8(await zip.extract("docProps/app.xml"))
    const appProps = parseAppProperties(appXml)
    if (Object.keys(appProps).length > 0) {
      properties = { ...properties, ...appProps }
    }
  }

  if (zip.has("docProps/custom.xml")) {
    const customXml = decodeUtf8(await zip.extract("docProps/custom.xml"))
    const customProps = parseCustomProperties(customXml)
    if (Object.keys(customProps).length > 0) {
      if (!properties) properties = {}
      properties.custom = customProps
    }
  }

  // 12. Build workbook
  const workbook: Workbook = {
    sheets,
    dateSystem,
  }

  const remainingNames = applyPrintDefinedNames(sheets, namedRanges)
  if (remainingNames.length > 0) {
    workbook.namedRanges = remainingNames
  }

  if (properties) {
    workbook.properties = properties
  }

  if (themeColors) {
    workbook.themeColors = themeColors
  }

  if (workbookProtection) {
    workbook.workbookProtection = workbookProtection
  }

  if (activeSheet !== undefined) {
    workbook.activeSheet = activeSheet
  }

  // The workbook's default font is fonts[0] in styles.xml — the entry
  // every xf inherits from unless it names another. Surfacing it closes
  // the WriteOptions.defaultFont round trip.
  const baseFont = parsedStyles?.fonts[0]
  if (baseFont && Object.keys(baseFont).length > 0) {
    workbook.defaultFont = baseFont
  }

  if (persons && persons.length > 0) {
    workbook.persons = persons
  }

  if (externalLinks.length > 0) {
    workbook.externalLinks = externalLinks
  }

  if (cellImages && cellImages.length > 0) {
    workbook.cellImages = cellImages
  }

  if (pivotCaches.length > 0) {
    workbook.pivotCaches = pivotCaches
  }

  if (slicerCaches.length > 0) {
    workbook.slicerCaches = slicerCaches
  }

  if (timelineCaches.length > 0) {
    workbook.timelineCaches = timelineCaches
  }

  return workbook
}

/**
 * Trailing numeric suffix of a relationship target like
 * `slicerCaches/slicerCache3.xml` or `timelines/timeline12.xml`. Used to
 * sort caches in declaration order regardless of the rId order. Returns
 * `+Infinity` for paths with no trailing number so malformed entries
 * sort last instead of throwing.
 */
function slicerNumFromTarget(target: string): number {
  const m = target.match(/(\d+)\.xml$/i)
  return m ? parseInt(m[1], 10) : Number.POSITIVE_INFINITY
}

/**
 * Numeric value for the trailing digits of an `rIdNN` identifier so we
 * can sort external link relationships in declaration order. Falls
 * back to `Infinity` when the id has no digits — keeps malformed
 * entries last instead of throwing.
 */
function relIdNum(rId: string): number {
  const m = rId.match(/(\d+)$/)
  return m ? parseInt(m[1], 10) : Number.POSITIVE_INFINITY
}

/**
 * Path of the `_rels` file belonging to `partPath`. Returns
 * `xl/externalLinks/_rels/externalLink1.xml.rels` for input
 * `xl/externalLinks/externalLink1.xml`.
 */
function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/")
  if (slash === -1) return `_rels/${partPath}.rels`
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`
}

// ── Drawing / Image Extraction ────────────────────────────────────────

/** Extension → SheetImage type mapping */
const EXT_TO_IMAGE_TYPE: Record<string, SheetImage["type"]> = {
  png: "png",
  jpg: "jpeg",
  jpeg: "jpeg",
  gif: "gif",
  svg: "svg",
  webp: "webp",
}

/**
 * Pairing of a chart part path with the cell anchor that pins the
 * chart to its host sheet. The anchor mirrors the drawing-layer
 * `<xdr:from>` / `<xdr:to>` pair, dropped through to {@link Chart.anchor}
 * by the reader.
 *
 * `anchor` is omitted when the drawing positions the chart via
 * `<xdr:absoluteAnchor>` (EMU-positioned, no cell anchor) or when the
 * `<xdr:from>` element itself is missing — both are rare but possible.
 */
interface DrawingChartRef {
  /** Path to `xl/charts/chartN.xml`, resolved against the package root. */
  path: string
  /** Cell anchor surfaced through {@link Chart.anchor}. */
  anchor?: ChartAnchor
  /** `<xdr:cNvPr @descr>` surfaced through {@link Chart.altText}. */
  altText?: string
  /** `<xdr:cNvPr @title>` surfaced through {@link Chart.frameTitle}. */
  frameTitle?: string
}

interface DrawingExtraction {
  images: SheetImage[]
  textBoxes: SheetTextBox[]
  /**
   * Chart parts referenced by this drawing, paired with the cell
   * anchor that pins each chart to its host sheet. Empty when the
   * drawing has no `c:chart` graphic frames or the rels file is missing.
   */
  chartRefs: DrawingChartRef[]
}

/**
 * Extract images and textboxes from a drawing XML and the ZIP archive.
 * Parses the drawing XML to find image anchors and textbox shapes,
 * resolves their relationships to media files, and extracts the binary data.
 */
async function extractSheetDrawing(
  zip: ZipReader,
  drawingPath: string,
): Promise<DrawingExtraction> {
  if (!zip.has(drawingPath)) return { images: [], textBoxes: [], chartRefs: [] }

  const drawingXml = decodeUtf8(await zip.extract(drawingPath))

  // Parse drawing relationships. Same fix as the worksheet path — the
  // hand-rolled slice dropped the first character for a root-level
  // drawing. See #391.
  const drawDir = dirname(drawingPath)
  const drawRelsPath = relsPathFor(drawingPath)

  const imageRelMap = new Map<string, string>()
  // Chart relationships are resolved by the same .rels file. We collect
  // them as `rId → resolved package path` so the drawing-level
  // graphicFrame walk below can map each chart reference to a file.
  const chartRelMap = new Map<string, string>()
  if (zip.has(drawRelsPath)) {
    const drawRelsXml = decodeUtf8(await zip.extract(drawRelsPath))
    const drawRels = parseRelationships(drawRelsXml)
    for (const rel of drawRels) {
      if (matchesRelType(rel.type, "image")) {
        imageRelMap.set(rel.id, resolvePath(drawDir, rel.target))
      } else if (matchesRelType(rel.type, "chart")) {
        chartRelMap.set(rel.id, resolvePath(drawDir, rel.target))
      }
    }
  }

  // Parse the drawing XML to find twoCellAnchor and oneCellAnchor elements with images/textboxes
  const doc = parseXml(drawingXml)
  const images: SheetImage[] = []
  const textBoxes: SheetTextBox[] = []
  const chartRefs: DrawingChartRef[] = []

  for (const child of doc.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag

    if (local === "twoCellAnchor" || local === "oneCellAnchor" || local === "absoluteAnchor") {
      // Charts are anchored via <xdr:graphicFrame> wrapping a
      // <a:graphic>/<a:graphicData uri="...chart"><c:chart r:id="rIdN"/>
      // — collect any chart references regardless of anchor flavor so
      // the roundtrip can declare the chart parts later.
      const chartRid = findChartRid(child)
      if (chartRid) {
        const chartPath = chartRelMap.get(chartRid)
        if (chartPath && !chartRefs.some((r) => r.path === chartPath)) {
          // absoluteAnchor positions in EMU rather than cells — skip
          // its anchor extraction so we don't fabricate a (0,0) cell
          // anchor that doesn't match the underlying placement.
          const anchor = local === "absoluteAnchor" ? undefined : parseChartCellAnchor(child, local)
          const ref: DrawingChartRef = { path: chartPath }
          if (anchor) ref.anchor = anchor
          // Accessibility metadata sits on the frame in the drawing, not
          // in the chart part — collect it while we still have the anchor.
          const frame = findChildEl(child, "graphicFrame")
          if (frame) {
            const meta = findCNvPrMeta(frame, "nvGraphicFramePr")
            if (meta.altText) ref.altText = meta.altText
            if (meta.title) ref.frameTitle = meta.title
          }
          chartRefs.push(ref)
        }
      }
    }

    if (local === "twoCellAnchor") {
      // Check if this anchor contains a textbox shape
      const textBox = parseTwoCellAnchorTextBox(child)
      if (textBox) {
        textBoxes.push(textBox)
        continue
      }

      const imageInfo = parseTwoCellAnchor(child, imageRelMap)
      if (imageInfo) {
        // Extract image data from ZIP
        const imagePath = imageInfo.mediaPath
        if (zip.has(imagePath)) {
          const data = await zip.extract(imagePath)
          const img: SheetImage = {
            data,
            type: imageInfo.type,
            anchor: imageInfo.anchor,
          }
          if (imageInfo.width !== undefined) img.width = imageInfo.width
          if (imageInfo.height !== undefined) img.height = imageInfo.height
          if (imageInfo.altText !== undefined) img.altText = imageInfo.altText
          if (imageInfo.title !== undefined) img.title = imageInfo.title
          images.push(img)
        }
      }
    } else if (local === "oneCellAnchor") {
      const imageInfo = parseOneCellAnchor(child, imageRelMap)
      if (imageInfo) {
        const imagePath = imageInfo.mediaPath
        if (zip.has(imagePath)) {
          const data = await zip.extract(imagePath)
          const img: SheetImage = {
            data,
            type: imageInfo.type,
            anchor: imageInfo.anchor,
          }
          if (imageInfo.width !== undefined) img.width = imageInfo.width
          if (imageInfo.height !== undefined) img.height = imageInfo.height
          if (imageInfo.altText !== undefined) img.altText = imageInfo.altText
          if (imageInfo.title !== undefined) img.title = imageInfo.title
          images.push(img)
        }
      }
    }
  }

  return { images, textBoxes, chartRefs }
}

/**
 * Extract the cell anchor (`<xdr:from>` / `<xdr:to>`) from a drawing
 * anchor element that wraps a chart graphic frame. Used by the reader
 * to surface {@link Chart.anchor}.
 *
 * For `xdr:twoCellAnchor` we read both `<xdr:from>` and `<xdr:to>`.
 * For `xdr:oneCellAnchor` we read `<xdr:from>` only — the chart is
 * pinned to a single cell with intrinsic width/height stored in
 * `<xdr:ext>`, so a `to` cell is not meaningful.
 *
 * Returns `undefined` when no `<xdr:from>` block is present.
 */
function parseChartCellAnchor(
  el: { children: Array<unknown> },
  anchorKind: string,
): ChartAnchor | undefined {
  let from: { row: number; col: number } | undefined
  let to: { row: number; col: number } | undefined

  for (const child of el.children) {
    if (typeof child === "string") continue
    const c = child as { local?: string; tag: string; children: Array<unknown> }
    const local = c.local || c.tag
    if (local === "from") from = parseAnchorPosition(c)
    else if (local === "to") to = parseAnchorPosition(c)
  }

  if (!from) return undefined
  if (anchorKind === "twoCellAnchor" && to) return { from, to }
  return { from }
}

/**
 * Walk an anchor element looking for `<c:chart r:id="...">` (the
 * standard chart embed inside `<xdr:graphicFrame>/<a:graphic>/<a:graphicData>`).
 * Returns the rId attribute or undefined if no chart is found.
 */
function findChartRid(el: { children: Array<unknown> }): string | undefined {
  for (const child of el.children) {
    if (typeof child === "string") continue
    const c = child as {
      local?: string
      tag: string
      children: Array<unknown>
      attrs: Record<string, string>
    }
    const local = c.local || c.tag
    if (local === "chart") {
      // r:id is namespaced; try common attribute spellings.
      const rid = c.attrs["r:id"] ?? c.attrs["id"]
      if (rid) return rid
    }
    const nested = findChartRid(c)
    if (nested) return nested
  }
  return undefined
}

/** Parse a twoCellAnchor element to extract image position and reference */
function parseTwoCellAnchor(
  el: { children: Array<unknown> },
  relMap: Map<string, string>,
): {
  mediaPath: string
  type: SheetImage["type"]
  anchor: SheetImage["anchor"]
  width?: number
  height?: number
  altText?: string
  title?: string
} | null {
  let fromRow = 0
  let fromCol = 0
  let toRow = 0
  let toCol = 0
  let embedId: string | undefined
  let altText: string | undefined
  let title: string | undefined
  let size: { width: number; height: number } | undefined

  for (const child of el.children) {
    if (typeof child === "string") continue
    const c = child as {
      local?: string
      tag: string
      children: Array<unknown>
      attrs: Record<string, string>
    }
    const local = c.local || c.tag

    if (local === "from") {
      const pos = parseAnchorPosition(c)
      fromRow = pos.row
      fromCol = pos.col
    } else if (local === "to") {
      const pos = parseAnchorPosition(c)
      toRow = pos.row
      toCol = pos.col
    } else if (local === "pic") {
      embedId = findBlipEmbed(c)
      const meta = findCNvPrMeta(c, "nvPicPr")
      altText = meta.altText
      title = meta.title
      size = findShapeExtent(c)
    }
  }

  if (!embedId) return null

  const mediaPath = relMap.get(embedId)
  if (!mediaPath) return null

  // Determine image type from file extension
  const ext = mediaPath.split(".").pop()?.toLowerCase() ?? ""
  const imageType = EXT_TO_IMAGE_TYPE[ext] ?? "png"

  const result: {
    mediaPath: string
    type: SheetImage["type"]
    anchor: SheetImage["anchor"]
    width?: number
    height?: number
    altText?: string
    title?: string
  } = {
    mediaPath,
    type: imageType,
    anchor: {
      from: { row: fromRow, col: fromCol },
      to: { row: toRow, col: toCol },
    },
  }
  if (size) {
    result.width = size.width
    result.height = size.height
  }
  if (altText) result.altText = altText
  if (title) result.title = title
  return result
}

/**
 * Pull the rendered size off a `<xdr:pic>` / `<xdr:sp>` shape's
 * `<xdr:spPr><a:xfrm><a:ext cx cy>`, converted from EMU to pixels.
 *
 * On a `twoCellAnchor` the from/to cells are what Excel actually honours,
 * but `a:ext` is still required by the schema and every writer — hucre's
 * included — records the intended size there. Reading it is the only way
 * to recover {@link SheetImage.width} from a hucre-written file, which
 * always uses `twoCellAnchor` (#407).
 */
function findShapeExtent(shapeEl: {
  children: Array<unknown>
}): { width: number; height: number } | undefined {
  const spPr = findChildEl(shapeEl, "spPr")
  if (!spPr) return undefined
  const xfrm = findChildEl(spPr, "xfrm")
  if (!xfrm) return undefined
  const ext = findChildEl(xfrm, "ext")
  if (!ext) return undefined
  const cx = Number(ext.attrs["cx"])
  const cy = Number(ext.attrs["cy"])
  // A zero / absent extent is a placeholder, not a 0×0 shape — the chart
  // writer emits exactly that. Report nothing rather than a size of 0.
  if (!(cx > 0) || !(cy > 0)) return undefined
  return { width: Math.round(cx / EMU_PER_PIXEL), height: Math.round(cy / EMU_PER_PIXEL) }
}

/** Parse a twoCellAnchor element that contains a textbox shape (sp with txBox="1") */
function parseTwoCellAnchorTextBox(el: { children: Array<unknown> }): SheetTextBox | null {
  let fromRow = 0
  let fromCol = 0
  let toRow = 0
  let toCol = 0
  let spElement: any = null

  for (const child of el.children) {
    if (typeof child === "string") continue
    const c = child as {
      local?: string
      tag: string
      children: Array<unknown>
      attrs: Record<string, string>
    }
    const local = c.local || c.tag

    if (local === "from") {
      const pos = parseAnchorPosition(c)
      fromRow = pos.row
      fromCol = pos.col
    } else if (local === "to") {
      const pos = parseAnchorPosition(c)
      toRow = pos.row
      toCol = pos.col
    } else if (local === "sp") {
      // Check if this is a textbox shape
      const nvSpPr = findChildEl(c, "nvSpPr")
      if (nvSpPr) {
        const cNvSpPr = findChildEl(nvSpPr, "cNvSpPr")
        if (cNvSpPr && (cNvSpPr.attrs["txBox"] === "1" || cNvSpPr.attrs["txBox"] === "true")) {
          spElement = c
        }
      }
    }
  }

  if (!spElement) return null

  // Extract text from txBody
  const txBody = findChildEl(spElement, "txBody")
  let text = ""
  let fontSize: number | undefined
  let bold: boolean | undefined
  let color: string | undefined

  if (txBody) {
    // Collect text from all paragraphs
    const paragraphs: string[] = []
    for (const pChild of txBody.children) {
      if (typeof pChild === "string") continue
      const pLocal = (pChild as any).local || (pChild as any).tag
      if (pLocal === "p") {
        const pText = extractParagraphText(pChild as any)
        paragraphs.push(pText.text)
        // Get style from first run with properties
        if (pText.fontSize !== undefined && fontSize === undefined) fontSize = pText.fontSize
        if (pText.bold !== undefined && bold === undefined) bold = pText.bold
        if (pText.color !== undefined && color === undefined) color = pText.color
      }
    }
    text = paragraphs.join("\n")
  }

  if (!text) return null

  // Extract fill and border colors from spPr
  let fillColor: string | undefined
  let borderColor: string | undefined
  const spPr = findChildEl(spElement, "spPr")
  if (spPr) {
    const solidFill = findChildEl(spPr, "solidFill")
    if (solidFill) {
      const srgbClr = findChildEl(solidFill, "srgbClr")
      if (srgbClr && srgbClr.attrs["val"]) {
        fillColor = srgbClr.attrs["val"]
      }
    }
    const ln = findChildEl(spPr, "ln")
    if (ln) {
      const lnFill = findChildEl(ln, "solidFill")
      if (lnFill) {
        const lnClr = findChildEl(lnFill, "srgbClr")
        if (lnClr && lnClr.attrs["val"]) {
          borderColor = lnClr.attrs["val"]
        }
      }
    }
  }

  const tb: SheetTextBox = {
    text,
    anchor: {
      from: { row: fromRow, col: fromCol },
      to: { row: toRow, col: toCol },
    },
  }

  const size = findShapeExtent(spElement)
  if (size) {
    tb.width = size.width
    tb.height = size.height
  }

  // Pull alt text / title off cNvPr so screen-reader metadata round-trips.
  const meta = findCNvPrMeta(spElement, "nvSpPr")
  if (meta.altText) tb.altText = meta.altText
  if (meta.title) tb.title = meta.title

  const style: SheetTextBox["style"] = {}
  let hasStyle = false
  if (fontSize !== undefined) {
    style.fontSize = fontSize
    hasStyle = true
  }
  if (bold !== undefined) {
    style.bold = bold
    hasStyle = true
  }
  if (color !== undefined) {
    style.color = color
    hasStyle = true
  }
  if (fillColor !== undefined) {
    style.fillColor = fillColor
    hasStyle = true
  }
  if (borderColor !== undefined) {
    style.borderColor = borderColor
    hasStyle = true
  }
  if (hasStyle) tb.style = style

  return tb
}

/** Find a child element by local name */
function findChildEl(
  el: { children: Array<unknown> },
  localName: string,
): { local?: string; tag: string; children: Array<unknown>; attrs: Record<string, string> } | null {
  for (const child of el.children) {
    if (typeof child === "string") continue
    const c = child as {
      local?: string
      tag: string
      children: Array<unknown>
      attrs: Record<string, string>
    }
    const local = c.local || c.tag
    if (local === localName) return c
  }
  return null
}

/** Extract text content from a DrawingML <a:p> paragraph element */
function extractParagraphText(pEl: { children: Array<unknown> }): {
  text: string
  fontSize?: number
  bold?: boolean
  color?: string
} {
  let text = ""
  let fontSize: number | undefined
  let bold: boolean | undefined
  let color: string | undefined

  for (const child of pEl.children) {
    if (typeof child === "string") continue
    const c = child as {
      local?: string
      tag: string
      children: Array<unknown>
      attrs: Record<string, string>
    }
    const local = c.local || c.tag

    if (local === "r") {
      // Run element: extract rPr and t
      const rPr = findChildEl(c, "rPr")
      if (rPr) {
        if (rPr.attrs["sz"]) {
          fontSize = Number(rPr.attrs["sz"]) / 100
        }
        if (rPr.attrs["b"] === "1" || rPr.attrs["b"] === "true") {
          bold = true
        }
        // Check for color in solidFill child
        const solidFill = findChildEl(rPr, "solidFill")
        if (solidFill) {
          const srgbClr = findChildEl(solidFill, "srgbClr")
          if (srgbClr && srgbClr.attrs["val"]) {
            color = srgbClr.attrs["val"]
          }
        }
      }
      const tEl = findChildEl(c, "t")
      if (tEl) {
        text += tEl.children.filter((ch: unknown) => typeof ch === "string").join("")
      }
    }
  }

  return { text, fontSize, bold, color }
}

/** EMU per pixel (at 96 DPI) */
const EMU_PER_PIXEL = 9525

/** Parse a oneCellAnchor element to extract image position, dimensions, and reference */
function parseOneCellAnchor(
  el: { children: Array<unknown> },
  relMap: Map<string, string>,
): {
  mediaPath: string
  type: SheetImage["type"]
  anchor: SheetImage["anchor"]
  width?: number
  height?: number
  altText?: string
  title?: string
} | null {
  let fromRow = 0
  let fromCol = 0
  let widthEmu = 0
  let heightEmu = 0
  let embedId: string | undefined
  let altText: string | undefined
  let title: string | undefined

  for (const child of el.children) {
    if (typeof child === "string") continue
    const c = child as {
      local?: string
      tag: string
      children: Array<unknown>
      attrs: Record<string, string>
    }
    const local = c.local || c.tag

    if (local === "from") {
      const pos = parseAnchorPosition(c)
      fromRow = pos.row
      fromCol = pos.col
    } else if (local === "ext") {
      // <xdr:ext cx="..." cy="..."/>
      widthEmu = Number(c.attrs["cx"]) || 0
      heightEmu = Number(c.attrs["cy"]) || 0
    } else if (local === "pic") {
      embedId = findBlipEmbed(c)
      const meta = findCNvPrMeta(c, "nvPicPr")
      altText = meta.altText
      title = meta.title
    }
  }

  if (!embedId) return null

  const mediaPath = relMap.get(embedId)
  if (!mediaPath) return null

  const ext = mediaPath.split(".").pop()?.toLowerCase() ?? ""
  const imageType = EXT_TO_IMAGE_TYPE[ext] ?? "png"

  const result: {
    mediaPath: string
    type: SheetImage["type"]
    anchor: SheetImage["anchor"]
    width?: number
    height?: number
    altText?: string
    title?: string
  } = {
    mediaPath,
    type: imageType,
    anchor: {
      from: { row: fromRow, col: fromCol },
    },
  }

  if (widthEmu > 0) {
    result.width = Math.round(widthEmu / EMU_PER_PIXEL)
  }
  if (heightEmu > 0) {
    result.height = Math.round(heightEmu / EMU_PER_PIXEL)
  }
  if (altText) result.altText = altText
  if (title) result.title = title

  return result
}

/**
 * Walk a `<xdr:pic>`, `<xdr:sp>` or `<xdr:graphicFrame>` element and
 * extract `descr=`/`title=` from its `xdr:cNvPr`. The cNvPr element lives
 * inside an `nv*Pr` wrapper named `nvPicPr` (pictures), `nvSpPr` (shapes)
 * or `nvGraphicFramePr` (charts). Returns empty fields when neither
 * attribute is present.
 */
function findCNvPrMeta(
  parentEl: { children: Array<unknown> },
  wrapperName: "nvPicPr" | "nvSpPr" | "nvGraphicFramePr",
): { altText?: string; title?: string } {
  const wrapper = findChildEl(parentEl, wrapperName)
  if (!wrapper) return {}
  const cNvPr = findChildEl(wrapper, "cNvPr")
  if (!cNvPr) return {}
  const out: { altText?: string; title?: string } = {}
  if (cNvPr.attrs["descr"]) out.altText = cNvPr.attrs["descr"]
  if (cNvPr.attrs["title"]) out.title = cNvPr.attrs["title"]
  return out
}

/** Parse row/col from an anchor position element (from or to) */
function parseAnchorPosition(el: { children: Array<unknown> }): { row: number; col: number } {
  let row = 0
  let col = 0

  for (const child of el.children) {
    if (typeof child === "string") continue
    const c = child as { local?: string; tag: string; children: Array<unknown> }
    const local = c.local || c.tag
    const text = c.children.filter((ch: unknown) => typeof ch === "string").join("")

    if (local === "row") {
      row = Number(text) || 0
    } else if (local === "col") {
      col = Number(text) || 0
    }
  }

  return { row, col }
}

/** Find the r:embed attribute on the blip element inside a pic element */
function findBlipEmbed(picEl: { children: Array<unknown> }): string | undefined {
  for (const child of picEl.children) {
    if (typeof child === "string") continue
    const c = child as {
      local?: string
      tag: string
      children: Array<unknown>
      attrs: Record<string, string>
    }
    const local = c.local || c.tag

    if (local === "blipFill") {
      for (const blipChild of c.children) {
        if (typeof blipChild === "string") continue
        const bc = blipChild as { local?: string; tag: string; attrs: Record<string, string> }
        const blipLocal = bc.local || bc.tag
        if (blipLocal === "blip") {
          // Look for r:embed attribute (namespace prefix may vary)
          return bc.attrs["r:embed"] ?? bc.attrs["R:embed"] ?? findEmbedAttr(bc.attrs)
        }
      }
    }
  }
  return undefined
}

/** Find an embed attribute regardless of namespace prefix */
function findEmbedAttr(attrs: Record<string, string>): string | undefined {
  for (const key of Object.keys(attrs)) {
    if (key.endsWith(":embed") && attrs[key].startsWith("rId")) {
      return attrs[key]
    }
  }
  return undefined
}

// ── Print Defined Names ──────────────────────────────────────────────

/**
 * Fold the reserved `_xlnm.Print_Area` / `_xlnm.Print_Titles` defined
 * names back into the owning sheet's {@link PageSetup} and return the
 * defined names that remain.
 *
 * The writer only ever *derives* these two names from `pageSetup`
 * (`buildNamedRanges`), so `pageSetup` is the single authoring surface.
 * Leaving them in `namedRanges` as well would give the same setting two
 * unconnected representations depending on the direction of travel
 * (#407) — and would make `openXlsx` → `saveXlsx` emit each name twice,
 * once from the carried-over `namedRanges` entry and once re-derived
 * from `pageSetup`. So they are consumed, not copied.
 *
 * A name we cannot attribute to a sheet we actually read — no
 * `localSheetId`, or a sheet skipped by `ReadOptions.sheets` — is left
 * alone; dropping it would lose information with nowhere else to go.
 */
function applyPrintDefinedNames(sheets: Sheet[], namedRanges: NamedRange[]): NamedRange[] {
  if (namedRanges.length === 0) return namedRanges

  const byName = new Map<string, Sheet>()
  for (const sheet of sheets) byName.set(sheet.name, sheet)

  const remaining: NamedRange[] = []
  for (const nr of namedRanges) {
    const isPrintArea = nr.name === "_xlnm.Print_Area"
    const isPrintTitles = nr.name === "_xlnm.Print_Titles"
    const sheet = nr.scope !== undefined ? byName.get(nr.scope) : undefined
    if ((!isPrintArea && !isPrintTitles) || !sheet) {
      remaining.push(nr)
      continue
    }

    const pageSetup = sheet.pageSetup ?? (sheet.pageSetup = {})
    if (isPrintArea) {
      const area = nr.range
        .split(",")
        .map(stripSheetQualifier)
        .filter((part) => part.length > 0)
        .join(",")
      if (area) pageSetup.printArea = area
    } else {
      // Print_Titles packs the repeat-rows and repeat-columns ranges into
      // one comma-separated value, in either order. A row range is all
      // digits ("$1:$1"), a column range all letters ("$A:$A").
      for (const raw of nr.range.split(",")) {
        const part = stripSheetQualifier(raw)
        if (!part) continue
        if (/\d/.test(part)) pageSetup.printTitlesRow = part
        else pageSetup.printTitlesColumn = part
      }
    }
  }

  return remaining
}

/**
 * Drop the `Sheet1!` / `'My Sheet'!` prefix from one range of a defined
 * name, leaving the bare A1 reference that {@link PageSetup} stores.
 */
function stripSheetQualifier(range: string): string {
  const trimmed = range.trim()
  const bang = trimmed.lastIndexOf("!")
  return bang === -1 ? trimmed : trimmed.slice(bang + 1)
}

// ── Workbook XML Parsing ─────────────────────────────────────────────

interface SheetInfo {
  name: string
  sheetId: number
  rId: string
  state?: "visible" | "hidden" | "veryHidden"
}

function parseWorkbookXml(
  xml: string,
  options?: ReadOptions,
): {
  sheets: SheetInfo[]
  dateSystem: "1900" | "1904"
  namedRanges: NamedRange[]
  workbookProtection?: { lockStructure?: boolean; lockWindows?: boolean }
  /**
   * `<bookViews><workbookView activeTab>` — the tab Excel opens on.
   * Omitted when the file declares tab 0, which is the OOXML default and
   * therefore indistinguishable from "the file said nothing"; the same
   * collapse the chart reader applies to `<c:overlay val="0"/>`.
   */
  activeSheet?: number
  /**
   * Pivot cache wiring read off the workbook's `<pivotCaches>` block.
   * Each entry maps a cacheId (Excel's stable handle) to an rId in
   * `xl/_rels/workbook.xml.rels`.
   */
  pivotCacheRefs: Array<{ cacheId: number; rId: string }>
} {
  const doc = parseXml(xml)

  const sheets: SheetInfo[] = []
  const namedRanges: NamedRange[] = []
  const pivotCacheRefs: Array<{ cacheId: number; rId: string }> = []
  let dateSystem: "1900" | "1904" = "1900"

  // Check date system override from options
  if (options?.dateSystem === "1904") {
    dateSystem = "1904"
  } else if (options?.dateSystem === "1900") {
    dateSystem = "1900"
  }

  let wbProtection: { lockStructure?: boolean; lockWindows?: boolean } | undefined
  let activeSheet: number | undefined

  // First pass: collect sheets (needed for resolving localSheetId)
  for (const child of doc.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag

    if (local === "workbookPr") {
      // Check for 1904 date system
      if (child.attrs["date1904"] === "1" || child.attrs["date1904"] === "true") {
        // Only override if auto or not set
        if (!options?.dateSystem || options.dateSystem === "auto") {
          dateSystem = "1904"
        }
      }
    }

    if (local === "workbookProtection") {
      const lockStructure =
        child.attrs["lockStructure"] === "1" || child.attrs["lockStructure"] === "true"
      const lockWindows =
        child.attrs["lockWindows"] === "1" || child.attrs["lockWindows"] === "true"
      if (lockStructure || lockWindows) {
        wbProtection = {}
        if (lockStructure) wbProtection.lockStructure = true
        if (lockWindows) wbProtection.lockWindows = true
      }
    }

    if (local === "bookViews") {
      for (const viewChild of child.children) {
        if (typeof viewChild === "string") continue
        const viewLocal = viewChild.local || viewChild.tag
        if (viewLocal !== "workbookView") continue
        const tab = Number(viewChild.attrs["activeTab"])
        // Excel writes one workbookView per window; the first is the one
        // that decides the opening tab.
        if (Number.isInteger(tab) && tab > 0 && activeSheet === undefined) activeSheet = tab
      }
    }

    if (local === "sheets") {
      for (const sheetChild of child.children) {
        if (typeof sheetChild === "string") continue
        const sheetLocal = sheetChild.local || sheetChild.tag
        if (sheetLocal === "sheet") {
          const name = sheetChild.attrs["name"] ?? ""
          const sheetId = Number(sheetChild.attrs["sheetId"] ?? "0")
          // r:id attribute — the namespace prefix may vary
          const rId =
            sheetChild.attrs["r:id"] ??
            sheetChild.attrs["R:id"] ??
            findRIdAttr(sheetChild.attrs) ??
            ""
          const stateRaw = sheetChild.attrs["state"]
          let state: SheetInfo["state"] = "visible"
          if (stateRaw === "hidden") state = "hidden"
          else if (stateRaw === "veryHidden") state = "veryHidden"

          if (name && rId) {
            sheets.push({ name, sheetId, rId, state })
          }
        }
      }
    }

    if (local === "pivotCaches") {
      for (const pcChild of child.children) {
        if (typeof pcChild === "string") continue
        const pcLocal = pcChild.local || pcChild.tag
        if (pcLocal === "pivotCache") {
          const cacheIdRaw = pcChild.attrs["cacheId"]
          const cacheId = cacheIdRaw === undefined ? NaN : Number(cacheIdRaw)
          const rId =
            pcChild.attrs["r:id"] ?? pcChild.attrs["R:id"] ?? findRIdAttr(pcChild.attrs) ?? ""
          if (rId && !Number.isNaN(cacheId)) {
            pivotCacheRefs.push({ cacheId, rId })
          }
        }
      }
    }
  }

  // Second pass: collect defined names (named ranges)
  for (const child of doc.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag

    if (local === "definedNames") {
      for (const dnChild of child.children) {
        if (typeof dnChild === "string") continue
        const dnLocal = dnChild.local || dnChild.tag
        if (dnLocal === "definedName") {
          const name = dnChild.attrs["name"] ?? ""
          const rangeText = dnChild.children.filter((c: unknown) => typeof c === "string").join("")

          if (name && rangeText) {
            const nr: NamedRange = { name, range: rangeText }

            // Resolve localSheetId to sheet name
            const localSheetId = dnChild.attrs["localSheetId"]
            if (localSheetId !== undefined) {
              const idx = Number(localSheetId)
              if (idx >= 0 && idx < sheets.length) {
                nr.scope = sheets[idx].name
              }
            }

            // Comment attribute
            if (dnChild.attrs["comment"]) {
              nr.comment = dnChild.attrs["comment"]
            }

            namedRanges.push(nr)
          }
        }
      }
    }
  }

  return {
    sheets,
    dateSystem,
    namedRanges,
    workbookProtection: wbProtection,
    activeSheet,
    pivotCacheRefs,
  }
}

/** Find an r:id attribute regardless of namespace prefix */
function findRIdAttr(attrs: Record<string, string>): string | undefined {
  for (const key of Object.keys(attrs)) {
    // Match any prefix:id where the value looks like an rId
    if (key.endsWith(":id") && attrs[key].startsWith("rId")) {
      return attrs[key]
    }
  }
  return undefined
}

/** Filter sheet infos based on user-specified sheets option */
function filterSheets(allSheets: SheetInfo[], filter?: ReadOptions["sheets"]): SheetInfo[] {
  if (filter === undefined) return allSheets

  if (typeof filter === "function") {
    const result: SheetInfo[] = []
    for (let i = 0; i < allSheets.length; i++) {
      const info = allSheets[i]!
      const decision = filter(
        {
          name: info.name,
          index: i,
          hidden: info.state === "hidden",
          veryHidden: info.state === "veryHidden",
        },
        i,
      )
      if (decision) result.push(info)
    }
    return result
  }

  if (filter.length === 0) return allSheets

  const result: SheetInfo[] = []
  for (const spec of filter) {
    if (typeof spec === "number") {
      if (spec >= 0 && spec < allSheets.length) {
        result.push(allSheets[spec])
      }
    } else {
      const found = allSheets.find((s) => s.name === spec)
      if (found) result.push(found)
    }
  }

  return result
}

// ── Table XML Parsing ────────────────────────────────────────────────

/**
 * Parse a table XML file (xl/tables/tableN.xml) into a TableDefinition.
 */
function parseTableXml(xml: string): TableDefinition | null {
  const doc = parseXml(xml)

  // Root element should be <table>
  const name = doc.attrs["name"] ?? ""
  const displayName = doc.attrs["displayName"] ?? name
  const ref = doc.attrs["ref"] ?? ""

  if (!name) return null

  // Determine showTotalRow from totalsRowCount or totalsRowShown
  const totalsRowCount = doc.attrs["totalsRowCount"]
  const showTotalRow = totalsRowCount !== undefined && totalsRowCount !== "0"

  const columns: TableColumn[] = []
  let style: string | undefined
  let showRowStripes: boolean | undefined
  let showColumnStripes: boolean | undefined
  // The filter dropdowns exist only if <autoFilter> does — a table part
  // with no such child renders none. Starting from `true` made a table
  // written with `showAutoFilter: false` read back as `true` (#407).
  let showAutoFilter = false

  for (const child of doc.children) {
    if (typeof child === "string") continue
    const local = child.local || child.tag

    if (local === "autoFilter") {
      showAutoFilter = true
    } else if (local === "tableColumns") {
      for (const colChild of child.children) {
        if (typeof colChild === "string") continue
        const colLocal = colChild.local || colChild.tag
        if (colLocal === "tableColumn") {
          const col: TableColumn = {
            name: colChild.attrs["name"] ?? "",
          }
          if (colChild.attrs["totalsRowFunction"]) {
            col.totalFunction = colChild.attrs["totalsRowFunction"]
          }
          if (colChild.attrs["totalsRowLabel"]) {
            col.totalLabel = colChild.attrs["totalsRowLabel"]
          }
          // Parse totalsRowFormula child element
          for (const formulaChild of colChild.children) {
            if (typeof formulaChild === "string") continue
            const formulaLocal = formulaChild.local || formulaChild.tag
            if (formulaLocal === "totalsRowFormula") {
              const formulaText = formulaChild.children
                .filter((c: unknown) => typeof c === "string")
                .join("")
              if (formulaText) {
                col.totalFormula = formulaText
              }
            }
          }
          columns.push(col)
        }
      }
    } else if (local === "tableStyleInfo") {
      style = child.attrs["name"]
      showRowStripes = child.attrs["showRowStripes"] === "1"
      showColumnStripes = child.attrs["showColumnStripes"] === "1"
    }
  }

  const tableDef: TableDefinition = {
    name,
    columns,
  }

  if (displayName && displayName !== name) {
    tableDef.displayName = displayName
  }
  if (ref) {
    tableDef.range = ref
  }
  if (style) {
    tableDef.style = style
  }
  if (showRowStripes !== undefined) {
    tableDef.showRowStripes = showRowStripes
  }
  if (showColumnStripes !== undefined) {
    tableDef.showColumnStripes = showColumnStripes
  }
  tableDef.showAutoFilter = showAutoFilter
  if (showTotalRow) {
    tableDef.showTotalRow = true
  }

  return tableDef
}
