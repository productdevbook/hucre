// ── XLSX Reader ──────────────────────────────────────────────────────
// Reads Office Open XML (.xlsx) spreadsheet files.

import type {
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
  PivotCache,
  PivotTable,
} from "../_types";
import { parsePersons, parseThreadedComments } from "./threaded-comments-reader";
import { parseExternalLink } from "./external-link-reader";
import { attachPivotCacheFields, parsePivotCacheDefinition, parsePivotTable } from "./pivot-reader";
import { ParseError, ZipError } from "../errors";
import { ZipReader } from "../zip/reader";
import { parseXml } from "../xml/parser";
import { parseContentTypes } from "./content-types";
import { parseRelationships } from "./relationships";
import { parseSharedStrings } from "./shared-strings";
import { parseStyles } from "./styles";
import { parseWorksheet } from "./worksheet";
import type { ParsedStyles } from "./styles";
import type { SharedString } from "./shared-strings";
import type { Relationship } from "./relationships";
import { parseComments } from "./comments-reader";
import { parseCellRef } from "./worksheet";
import { parseCoreProperties, parseAppProperties, parseCustomProperties } from "./doc-props-reader";
import { parseThemeColors } from "./theme";

// ── OOXML Relationship Types ─────────────────────────────────────────

// Transitional namespace (OOXML 2006/Transitional — most common)
const NS_TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// Strict namespace (OOXML Strict — Excel 2013+ "Strict Open XML" save mode)
const NS_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/**
 * Match a relationship type against both Transitional and Strict OOXML namespaces.
 * Excel 2013+ can save in Strict mode which uses different namespace URIs.
 */
function matchesRelType(rel: string, type: string): boolean {
  return (
    rel === `${NS_TRANSITIONAL}/${type}` ||
    rel === `${NS_STRICT}/${type}` ||
    rel.endsWith("/" + type)
  );
}

// ── Helpers ──────────────────────────────────────────────────────────

function toUint8Array(input: ReadInput): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  throw new ParseError("Unsupported input type. Expected Uint8Array or ArrayBuffer.");
}

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data);
}

/**
 * Resolve a relative target path against a base directory.
 * E.g. resolve("xl/_rels", "../worksheets/sheet1.xml") → "xl/worksheets/sheet1.xml"
 */
function resolvePath(base: string, target: string): string {
  // If target starts with /, it's absolute from the package root
  if (target.startsWith("/")) return target.slice(1);

  const baseParts = base.split("/").filter(Boolean);
  const targetParts = target.split("/").filter(Boolean);

  for (const part of targetParts) {
    if (part === "..") {
      baseParts.pop();
    } else if (part !== ".") {
      baseParts.push(part);
    }
  }

  return baseParts.join("/");
}

/**
 * Get the directory portion of a path.
 * E.g. "xl/workbook.xml" → "xl"
 */
function dirname(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

// ── Main Reader ──────────────────────────────────────────────────────

/**
 * Read an XLSX file and return a Workbook.
 * Input can be Uint8Array or ArrayBuffer.
 */
export async function readXlsx(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = toUint8Array(input);

  // 1. Open ZIP archive
  let zip: ZipReader;
  try {
    zip = new ZipReader(data);
  } catch (err) {
    if (err instanceof ZipError) throw err;
    throw new ParseError("Failed to open XLSX file: not a valid ZIP archive", undefined, {
      cause: err,
    });
  }

  // 2. Parse [Content_Types].xml (validate it exists)
  if (!zip.has("[Content_Types].xml")) {
    throw new ParseError("Invalid XLSX: missing [Content_Types].xml");
  }
  const contentTypesXml = decodeUtf8(await zip.extract("[Content_Types].xml"));
  parseContentTypes(contentTypesXml); // Validate, not strictly needed for reading

  // 3. Parse _rels/.rels to find the workbook path
  if (!zip.has("_rels/.rels")) {
    throw new ParseError("Invalid XLSX: missing _rels/.rels");
  }
  const rootRelsXml = decodeUtf8(await zip.extract("_rels/.rels"));
  const rootRels = parseRelationships(rootRelsXml);
  const workbookRel = rootRels.find((r) => matchesRelType(r.type, "officeDocument"));
  if (!workbookRel) {
    throw new ParseError("Invalid XLSX: cannot find workbook relationship in _rels/.rels");
  }

  const workbookPath = workbookRel.target.startsWith("/")
    ? workbookRel.target.slice(1)
    : workbookRel.target;

  // 4. Parse workbook relationships (xl/_rels/workbook.xml.rels)
  const workbookDir = dirname(workbookPath);
  const workbookRelsPath = workbookDir
    ? `${workbookDir}/_rels/${workbookPath.slice(workbookDir.length + 1)}.rels`
    : `_rels/${workbookPath}.rels`;

  let workbookRels: Relationship[] = [];
  if (zip.has(workbookRelsPath)) {
    const wbRelsXml = decodeUtf8(await zip.extract(workbookRelsPath));
    workbookRels = parseRelationships(wbRelsXml);
  }

  // 5. Parse xl/workbook.xml for sheet names, order, and date system
  if (!zip.has(workbookPath)) {
    throw new ParseError(`Invalid XLSX: missing workbook at ${workbookPath}`);
  }
  const workbookXml = decodeUtf8(await zip.extract(workbookPath));
  const {
    sheets: sheetInfos,
    dateSystem,
    namedRanges,
    workbookProtection,
    pivotCacheRefs,
  } = parseWorkbookXml(workbookXml, options);

  // 6. Parse shared strings if present
  let sharedStrings: SharedString[] = [];
  const ssRel = workbookRels.find((r) => matchesRelType(r.type, "sharedStrings"));
  if (ssRel) {
    const ssPath = resolvePath(workbookDir, ssRel.target);
    if (zip.has(ssPath)) {
      const ssXml = decodeUtf8(await zip.extract(ssPath));
      sharedStrings = parseSharedStrings(ssXml);
    }
  }

  // 7. Parse styles if needed (for date detection or if readStyles is true)
  let parsedStyles: ParsedStyles | null = null;
  const stylesRel = workbookRels.find((r) => matchesRelType(r.type, "styles"));
  if (stylesRel) {
    const stylesPath = resolvePath(workbookDir, stylesRel.target);
    if (zip.has(stylesPath)) {
      const stylesXml = decodeUtf8(await zip.extract(stylesPath));
      parsedStyles = parseStyles(stylesXml);
    }
  }

  // 7b. Parse theme colors if theme1.xml exists
  let themeColors: string[] | undefined;
  const themePath = workbookDir ? `${workbookDir}/theme/theme1.xml` : "theme/theme1.xml";
  if (zip.has(themePath)) {
    const themeXml = decodeUtf8(await zip.extract(themePath));
    themeColors = parseThemeColors(themeXml);
  }

  // 7c. Parse the workbook-wide threaded-comments person directory
  // (xl/persons/person.xml). Linked from workbook.xml.rels by Type=".../person".
  let persons: ThreadedCommentPerson[] | undefined;
  const personsRel = workbookRels.find((r) => matchesRelType(r.type, "person"));
  if (personsRel) {
    const personsPath = resolvePath(workbookDir, personsRel.target);
    if (zip.has(personsPath)) {
      const personsXml = decodeUtf8(await zip.extract(personsPath));
      persons = parsePersons(personsXml);
    }
  }

  // 7d. Parse external workbook links (xl/externalLinks/externalLinkN.xml).
  // The workbook.xml.rels file declares them with Type=".../externalLink";
  // resolve each one in declaration order so the index lines up with
  // the `[N]` prefix used in formulas.
  const externalLinkRels = workbookRels
    .filter((r) => matchesRelType(r.type, "externalLink"))
    .sort((a, b) => relIdNum(a.id) - relIdNum(b.id));
  const externalLinks: ExternalLink[] = [];
  for (const rel of externalLinkRels) {
    const linkPath = resolvePath(workbookDir, rel.target);
    if (!zip.has(linkPath)) continue;
    const linkXml = decodeUtf8(await zip.extract(linkPath));
    const linkRelsPath = relsPathFor(linkPath);
    const linkRelsXml = zip.has(linkRelsPath)
      ? decodeUtf8(await zip.extract(linkRelsPath))
      : undefined;
    externalLinks.push(parseExternalLink(linkXml, linkRelsXml));
  }

  // 7e. Parse pivot caches (xl/pivotCache/pivotCacheDefinitionN.xml).
  // The workbook's <pivotCaches> block ties each cacheId to an rId in
  // workbook.xml.rels; we walk that pairing and resolve each cache.
  // The cache definitions also surface a `hasRecords` flag based on
  // whether the sibling rels declare a pivotCacheRecords part.
  const pivotCachesByCacheId = new Map<number, PivotCache>();
  const pivotCachesByRId = new Map<string, PivotCache>();
  for (const ref of pivotCacheRefs) {
    const rel = workbookRels.find((r) => r.id === ref.rId);
    if (!rel) continue;
    const cachePath = resolvePath(workbookDir, rel.target);
    if (!zip.has(cachePath)) continue;
    const cacheXml = decodeUtf8(await zip.extract(cachePath));
    const cache = parsePivotCacheDefinition(cacheXml);
    if (!cache) continue;
    cache.cacheId = ref.cacheId;
    // Detect a sibling pivotCacheRecords part via the cache's _rels.
    const cacheRelsPath = relsPathFor(cachePath);
    if (zip.has(cacheRelsPath)) {
      const cacheRelsXml = decodeUtf8(await zip.extract(cacheRelsPath));
      const cacheRels = parseRelationships(cacheRelsXml);
      if (cacheRels.some((r) => matchesRelType(r.type, "pivotCacheRecords"))) {
        cache.hasRecords = true;
      }
    }
    pivotCachesByCacheId.set(ref.cacheId, cache);
    pivotCachesByRId.set(ref.rId, cache);
  }
  // Preserve the workbook's declaration order so caller-side cacheId
  // lookups by array index match what Excel sees.
  const pivotCaches: PivotCache[] = pivotCacheRefs
    .map((ref) => pivotCachesByCacheId.get(ref.cacheId))
    .filter((c): c is PivotCache => c !== undefined);

  // 8. Build a map of rId → sheet relationship for worksheet paths
  const sheetRelMap = new Map<string, string>();
  for (const rel of workbookRels) {
    if (matchesRelType(rel.type, "worksheet")) {
      sheetRelMap.set(rel.id, resolvePath(workbookDir, rel.target));
    }
  }

  // 9. Filter sheets if options specify which ones to read
  const sheetsToRead = filterSheets(sheetInfos, options?.sheets);

  // 10. Parse each worksheet
  const readStyles = options?.readStyles ?? false;

  const sheets = [];
  for (const info of sheetsToRead) {
    const wsPath = sheetRelMap.get(info.rId);
    if (!wsPath || !zip.has(wsPath)) {
      throw new ParseError(`Invalid XLSX: missing worksheet file for sheet "${info.name}"`);
    }

    // Check for worksheet-level relationships (hyperlinks, etc.)
    const wsDir = dirname(wsPath);
    const wsFileName = wsPath.slice(wsDir.length + 1);
    const wsRelsPath = wsDir ? `${wsDir}/_rels/${wsFileName}.rels` : `_rels/${wsFileName}.rels`;
    let worksheetRels: Relationship[] | undefined;
    if (zip.has(wsRelsPath)) {
      const wsRelsXml = decodeUtf8(await zip.extract(wsRelsPath));
      worksheetRels = parseRelationships(wsRelsXml);
    }

    const worksheetCtx = {
      sharedStrings,
      styles: parsedStyles,
      readStyles,
      dateSystem,
      worksheetRels,
      maxRows: options?.maxRows,
      range: options?.range,
    };

    const wsXml = decodeUtf8(await zip.extract(wsPath));
    const sheet = parseWorksheet(wsXml, info.name, worksheetCtx);
    if (info.state === "hidden") sheet.hidden = true;
    if (info.state === "veryHidden") sheet.veryHidden = true;

    // Extract images and textboxes from drawing if present
    if (worksheetRels) {
      const drawingRel = worksheetRels.find((r) => matchesRelType(r.type, "drawing"));
      if (drawingRel) {
        const drawingPath = resolvePath(wsDir, drawingRel.target);
        const drawing = await extractSheetDrawing(zip, drawingPath);
        if (drawing.images.length > 0) {
          sheet.images = drawing.images;
        }
        if (drawing.textBoxes.length > 0) {
          sheet.textBoxes = drawing.textBoxes;
        }
      }
    }

    // Extract comments if present
    if (worksheetRels) {
      const commentsRel = worksheetRels.find((r) => matchesRelType(r.type, "comments"));
      if (commentsRel) {
        const commentsPath = resolvePath(wsDir, commentsRel.target);
        if (zip.has(commentsPath)) {
          const commentsXml = decodeUtf8(await zip.extract(commentsPath));
          const commentsMap = parseComments(commentsXml);

          // Attach comments to cell objects
          if (commentsMap.size > 0) {
            if (!sheet.cells) {
              sheet.cells = new Map();
            }
            for (const [cellRefStr, comment] of commentsMap) {
              const pos = parseCellRef(cellRefStr);
              const key = `${pos.row},${pos.col}`;
              let cell = sheet.cells.get(key);
              if (!cell) {
                cell = {
                  value: (sheet.rows[pos.row] && sheet.rows[pos.row][pos.col]) ?? null,
                  type: "string",
                };
                sheet.cells.set(key, cell);
              }
              cell.comment = comment;
            }
          }
        }
      }

      // Extract Excel 365 threaded comments if present.
      // Sheets can have BOTH legacy comments and threaded comments — Excel
      // writes a legacy stub for backward compat, so we treat them as
      // independent surfaces rather than overwriting each other.
      const threadedRel = worksheetRels.find((r) => matchesRelType(r.type, "threadedComment"));
      if (threadedRel) {
        const tcPath = resolvePath(wsDir, threadedRel.target);
        if (zip.has(tcPath)) {
          const tcXml = decodeUtf8(await zip.extract(tcPath));
          const threaded = parseThreadedComments(tcXml);
          if (threaded.length > 0) sheet.threadedComments = threaded;
        }
      }
    }

    // Extract tables if present
    if (worksheetRels) {
      const tableRels = worksheetRels.filter((r) => matchesRelType(r.type, "table"));
      if (tableRels.length > 0) {
        const tables: TableDefinition[] = [];
        for (const tableRel of tableRels) {
          const tablePath = resolvePath(wsDir, tableRel.target);
          if (zip.has(tablePath)) {
            const tableXml = decodeUtf8(await zip.extract(tablePath));
            const tableDef = parseTableXml(tableXml);
            if (tableDef) {
              tables.push(tableDef);
            }
          }
        }
        if (tables.length > 0) {
          sheet.tables = tables;
        }
      }
    }

    // Extract background image (picture relationship) if present
    if (worksheetRels) {
      const pictureRel = worksheetRels.find((r) => matchesRelType(r.type, "image"));
      if (pictureRel) {
        const picturePath = resolvePath(wsDir, pictureRel.target);
        if (zip.has(picturePath)) {
          sheet.backgroundImage = await zip.extract(picturePath);
        }
      }
    }

    // Extract pivot tables hosted on this sheet, if any. The sheet's
    // _rels file declares `Type=".../pivotTable"` per instance; the
    // body lives in xl/pivotTables/pivotTableN.xml and points at the
    // owning cache through its sibling _rels file.
    if (worksheetRels) {
      const pivotTableRels = worksheetRels.filter((r) => matchesRelType(r.type, "pivotTable"));
      if (pivotTableRels.length > 0) {
        const pivotTables: PivotTable[] = [];
        for (const ptRel of pivotTableRels) {
          const ptPath = resolvePath(wsDir, ptRel.target);
          if (!zip.has(ptPath)) continue;
          const ptXml = decodeUtf8(await zip.extract(ptPath));
          const pivot = parsePivotTable(ptXml);
          if (!pivot) continue;
          // Resolve the pivot's owning cache via its sibling _rels —
          // pivot tables don't carry the rId in the body, only in the
          // companion .rels file. Match by relative target so we
          // stay tolerant of caches living anywhere under xl/.
          const ptRelsPath = relsPathFor(ptPath);
          if (zip.has(ptRelsPath)) {
            const ptRelsXml = decodeUtf8(await zip.extract(ptRelsPath));
            const ptInternalRels = parseRelationships(ptRelsXml);
            const cacheRel = ptInternalRels.find((r) =>
              matchesRelType(r.type, "pivotCacheDefinition"),
            );
            if (cacheRel) {
              const resolvedCachePath = resolvePath(dirname(ptPath), cacheRel.target);
              for (const ref of pivotCacheRefs) {
                const wbRel = workbookRels.find((r) => r.id === ref.rId);
                if (!wbRel) continue;
                const wbCachePath = resolvePath(workbookDir, wbRel.target);
                if (wbCachePath === resolvedCachePath) {
                  pivot.cacheId = ref.cacheId;
                  break;
                }
              }
            }
          }
          // Overlay the cache's field names so consumers see the real
          // names instead of the synthetic field1/field2 placeholders
          // emitted by parsePivotTable when it has no cache context.
          const owningCache = pivotCachesByCacheId.get(pivot.cacheId);
          if (owningCache) attachPivotCacheFields(pivot, owningCache);
          pivotTables.push(pivot);
        }
        if (pivotTables.length > 0) sheet.pivotTables = pivotTables;
      }
    }

    sheets.push(sheet);
  }

  // 11. Parse document properties (if present)
  let properties: import("../_types").WorkbookProperties | undefined;

  if (zip.has("docProps/core.xml")) {
    const coreXml = decodeUtf8(await zip.extract("docProps/core.xml"));
    const coreProps = parseCoreProperties(coreXml);
    if (Object.keys(coreProps).length > 0) {
      properties = { ...coreProps };
    }
  }

  if (zip.has("docProps/app.xml")) {
    const appXml = decodeUtf8(await zip.extract("docProps/app.xml"));
    const appProps = parseAppProperties(appXml);
    if (Object.keys(appProps).length > 0) {
      properties = { ...properties, ...appProps };
    }
  }

  if (zip.has("docProps/custom.xml")) {
    const customXml = decodeUtf8(await zip.extract("docProps/custom.xml"));
    const customProps = parseCustomProperties(customXml);
    if (Object.keys(customProps).length > 0) {
      if (!properties) properties = {};
      properties.custom = customProps;
    }
  }

  // 12. Build workbook
  const workbook: Workbook = {
    sheets,
    dateSystem,
  };

  if (namedRanges.length > 0) {
    workbook.namedRanges = namedRanges;
  }

  if (properties) {
    workbook.properties = properties;
  }

  if (themeColors) {
    workbook.themeColors = themeColors;
  }

  if (workbookProtection) {
    workbook.workbookProtection = workbookProtection;
  }

  if (persons && persons.length > 0) {
    workbook.persons = persons;
  }

  if (externalLinks.length > 0) {
    workbook.externalLinks = externalLinks;
  }

  if (pivotCaches.length > 0) {
    workbook.pivotCaches = pivotCaches;
  }

  return workbook;
}

/**
 * Numeric value for the trailing digits of an `rIdNN` identifier so we
 * can sort external link relationships in declaration order. Falls
 * back to `Infinity` when the id has no digits — keeps malformed
 * entries last instead of throwing.
 */
function relIdNum(rId: string): number {
  const m = rId.match(/(\d+)$/);
  return m ? parseInt(m[1], 10) : Number.POSITIVE_INFINITY;
}

/**
 * Path of the `_rels` file belonging to `partPath`. Returns
 * `xl/externalLinks/_rels/externalLink1.xml.rels` for input
 * `xl/externalLinks/externalLink1.xml`.
 */
function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  if (slash === -1) return `_rels/${partPath}.rels`;
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
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
};

interface DrawingExtraction {
  images: SheetImage[];
  textBoxes: SheetTextBox[];
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
  if (!zip.has(drawingPath)) return { images: [], textBoxes: [] };

  const drawingXml = decodeUtf8(await zip.extract(drawingPath));

  // Parse drawing relationships
  const drawDir = dirname(drawingPath);
  const drawFileName = drawingPath.slice(drawDir.length + 1);
  const drawRelsPath = drawDir
    ? `${drawDir}/_rels/${drawFileName}.rels`
    : `_rels/${drawFileName}.rels`;

  const imageRelMap = new Map<string, string>();
  if (zip.has(drawRelsPath)) {
    const drawRelsXml = decodeUtf8(await zip.extract(drawRelsPath));
    const drawRels = parseRelationships(drawRelsXml);
    for (const rel of drawRels) {
      if (matchesRelType(rel.type, "image")) {
        imageRelMap.set(rel.id, resolvePath(drawDir, rel.target));
      }
    }
  }

  // Parse the drawing XML to find twoCellAnchor and oneCellAnchor elements with images/textboxes
  const doc = parseXml(drawingXml);
  const images: SheetImage[] = [];
  const textBoxes: SheetTextBox[] = [];

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "twoCellAnchor") {
      // Check if this anchor contains a textbox shape
      const textBox = parseTwoCellAnchorTextBox(child);
      if (textBox) {
        textBoxes.push(textBox);
        continue;
      }

      const imageInfo = parseTwoCellAnchor(child, imageRelMap);
      if (imageInfo) {
        // Extract image data from ZIP
        const imagePath = imageInfo.mediaPath;
        if (zip.has(imagePath)) {
          const data = await zip.extract(imagePath);
          const img: SheetImage = {
            data,
            type: imageInfo.type,
            anchor: imageInfo.anchor,
          };
          if (imageInfo.altText !== undefined) img.altText = imageInfo.altText;
          if (imageInfo.title !== undefined) img.title = imageInfo.title;
          images.push(img);
        }
      }
    } else if (local === "oneCellAnchor") {
      const imageInfo = parseOneCellAnchor(child, imageRelMap);
      if (imageInfo) {
        const imagePath = imageInfo.mediaPath;
        if (zip.has(imagePath)) {
          const data = await zip.extract(imagePath);
          const img: SheetImage = {
            data,
            type: imageInfo.type,
            anchor: imageInfo.anchor,
          };
          if (imageInfo.width !== undefined) img.width = imageInfo.width;
          if (imageInfo.height !== undefined) img.height = imageInfo.height;
          if (imageInfo.altText !== undefined) img.altText = imageInfo.altText;
          if (imageInfo.title !== undefined) img.title = imageInfo.title;
          images.push(img);
        }
      }
    }
  }

  return { images, textBoxes };
}

/** Parse a twoCellAnchor element to extract image position and reference */
function parseTwoCellAnchor(
  el: { children: Array<unknown> },
  relMap: Map<string, string>,
): {
  mediaPath: string;
  type: SheetImage["type"];
  anchor: SheetImage["anchor"];
  altText?: string;
  title?: string;
} | null {
  let fromRow = 0;
  let fromCol = 0;
  let toRow = 0;
  let toCol = 0;
  let embedId: string | undefined;
  let altText: string | undefined;
  let title: string | undefined;

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const c = child as {
      local?: string;
      tag: string;
      children: Array<unknown>;
      attrs: Record<string, string>;
    };
    const local = c.local || c.tag;

    if (local === "from") {
      const pos = parseAnchorPosition(c);
      fromRow = pos.row;
      fromCol = pos.col;
    } else if (local === "to") {
      const pos = parseAnchorPosition(c);
      toRow = pos.row;
      toCol = pos.col;
    } else if (local === "pic") {
      embedId = findBlipEmbed(c);
      const meta = findCNvPrMeta(c, "nvPicPr");
      altText = meta.altText;
      title = meta.title;
    }
  }

  if (!embedId) return null;

  const mediaPath = relMap.get(embedId);
  if (!mediaPath) return null;

  // Determine image type from file extension
  const ext = mediaPath.split(".").pop()?.toLowerCase() ?? "";
  const imageType = EXT_TO_IMAGE_TYPE[ext] ?? "png";

  const result: {
    mediaPath: string;
    type: SheetImage["type"];
    anchor: SheetImage["anchor"];
    altText?: string;
    title?: string;
  } = {
    mediaPath,
    type: imageType,
    anchor: {
      from: { row: fromRow, col: fromCol },
      to: { row: toRow, col: toCol },
    },
  };
  if (altText) result.altText = altText;
  if (title) result.title = title;
  return result;
}

/** Parse a twoCellAnchor element that contains a textbox shape (sp with txBox="1") */
function parseTwoCellAnchorTextBox(el: { children: Array<unknown> }): SheetTextBox | null {
  let fromRow = 0;
  let fromCol = 0;
  let toRow = 0;
  let toCol = 0;
  let spElement: any = null;

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const c = child as {
      local?: string;
      tag: string;
      children: Array<unknown>;
      attrs: Record<string, string>;
    };
    const local = c.local || c.tag;

    if (local === "from") {
      const pos = parseAnchorPosition(c);
      fromRow = pos.row;
      fromCol = pos.col;
    } else if (local === "to") {
      const pos = parseAnchorPosition(c);
      toRow = pos.row;
      toCol = pos.col;
    } else if (local === "sp") {
      // Check if this is a textbox shape
      const nvSpPr = findChildEl(c, "nvSpPr");
      if (nvSpPr) {
        const cNvSpPr = findChildEl(nvSpPr, "cNvSpPr");
        if (cNvSpPr && (cNvSpPr.attrs["txBox"] === "1" || cNvSpPr.attrs["txBox"] === "true")) {
          spElement = c;
        }
      }
    }
  }

  if (!spElement) return null;

  // Extract text from txBody
  const txBody = findChildEl(spElement, "txBody");
  let text = "";
  let fontSize: number | undefined;
  let bold: boolean | undefined;
  let color: string | undefined;

  if (txBody) {
    // Collect text from all paragraphs
    const paragraphs: string[] = [];
    for (const pChild of txBody.children) {
      if (typeof pChild === "string") continue;
      const pLocal = (pChild as any).local || (pChild as any).tag;
      if (pLocal === "p") {
        const pText = extractParagraphText(pChild as any);
        paragraphs.push(pText.text);
        // Get style from first run with properties
        if (pText.fontSize !== undefined && fontSize === undefined) fontSize = pText.fontSize;
        if (pText.bold !== undefined && bold === undefined) bold = pText.bold;
        if (pText.color !== undefined && color === undefined) color = pText.color;
      }
    }
    text = paragraphs.join("\n");
  }

  if (!text) return null;

  // Extract fill and border colors from spPr
  let fillColor: string | undefined;
  let borderColor: string | undefined;
  const spPr = findChildEl(spElement, "spPr");
  if (spPr) {
    const solidFill = findChildEl(spPr, "solidFill");
    if (solidFill) {
      const srgbClr = findChildEl(solidFill, "srgbClr");
      if (srgbClr && srgbClr.attrs["val"]) {
        fillColor = srgbClr.attrs["val"];
      }
    }
    const ln = findChildEl(spPr, "ln");
    if (ln) {
      const lnFill = findChildEl(ln, "solidFill");
      if (lnFill) {
        const lnClr = findChildEl(lnFill, "srgbClr");
        if (lnClr && lnClr.attrs["val"]) {
          borderColor = lnClr.attrs["val"];
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
  };

  // Pull alt text / title off cNvPr so screen-reader metadata round-trips.
  const meta = findCNvPrMeta(spElement, "nvSpPr");
  if (meta.altText) tb.altText = meta.altText;
  if (meta.title) tb.title = meta.title;

  const style: SheetTextBox["style"] = {};
  let hasStyle = false;
  if (fontSize !== undefined) {
    style.fontSize = fontSize;
    hasStyle = true;
  }
  if (bold !== undefined) {
    style.bold = bold;
    hasStyle = true;
  }
  if (color !== undefined) {
    style.color = color;
    hasStyle = true;
  }
  if (fillColor !== undefined) {
    style.fillColor = fillColor;
    hasStyle = true;
  }
  if (borderColor !== undefined) {
    style.borderColor = borderColor;
    hasStyle = true;
  }
  if (hasStyle) tb.style = style;

  return tb;
}

/** Find a child element by local name */
function findChildEl(
  el: { children: Array<unknown> },
  localName: string,
): { local?: string; tag: string; children: Array<unknown>; attrs: Record<string, string> } | null {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    const c = child as {
      local?: string;
      tag: string;
      children: Array<unknown>;
      attrs: Record<string, string>;
    };
    const local = c.local || c.tag;
    if (local === localName) return c;
  }
  return null;
}

/** Extract text content from a DrawingML <a:p> paragraph element */
function extractParagraphText(pEl: { children: Array<unknown> }): {
  text: string;
  fontSize?: number;
  bold?: boolean;
  color?: string;
} {
  let text = "";
  let fontSize: number | undefined;
  let bold: boolean | undefined;
  let color: string | undefined;

  for (const child of pEl.children) {
    if (typeof child === "string") continue;
    const c = child as {
      local?: string;
      tag: string;
      children: Array<unknown>;
      attrs: Record<string, string>;
    };
    const local = c.local || c.tag;

    if (local === "r") {
      // Run element: extract rPr and t
      const rPr = findChildEl(c, "rPr");
      if (rPr) {
        if (rPr.attrs["sz"]) {
          fontSize = Number(rPr.attrs["sz"]) / 100;
        }
        if (rPr.attrs["b"] === "1" || rPr.attrs["b"] === "true") {
          bold = true;
        }
        // Check for color in solidFill child
        const solidFill = findChildEl(rPr, "solidFill");
        if (solidFill) {
          const srgbClr = findChildEl(solidFill, "srgbClr");
          if (srgbClr && srgbClr.attrs["val"]) {
            color = srgbClr.attrs["val"];
          }
        }
      }
      const tEl = findChildEl(c, "t");
      if (tEl) {
        text += tEl.children.filter((ch: unknown) => typeof ch === "string").join("");
      }
    }
  }

  return { text, fontSize, bold, color };
}

/** EMU per pixel (at 96 DPI) */
const EMU_PER_PIXEL = 9525;

/** Parse a oneCellAnchor element to extract image position, dimensions, and reference */
function parseOneCellAnchor(
  el: { children: Array<unknown> },
  relMap: Map<string, string>,
): {
  mediaPath: string;
  type: SheetImage["type"];
  anchor: SheetImage["anchor"];
  width?: number;
  height?: number;
  altText?: string;
  title?: string;
} | null {
  let fromRow = 0;
  let fromCol = 0;
  let widthEmu = 0;
  let heightEmu = 0;
  let embedId: string | undefined;
  let altText: string | undefined;
  let title: string | undefined;

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const c = child as {
      local?: string;
      tag: string;
      children: Array<unknown>;
      attrs: Record<string, string>;
    };
    const local = c.local || c.tag;

    if (local === "from") {
      const pos = parseAnchorPosition(c);
      fromRow = pos.row;
      fromCol = pos.col;
    } else if (local === "ext") {
      // <xdr:ext cx="..." cy="..."/>
      widthEmu = Number(c.attrs["cx"]) || 0;
      heightEmu = Number(c.attrs["cy"]) || 0;
    } else if (local === "pic") {
      embedId = findBlipEmbed(c);
      const meta = findCNvPrMeta(c, "nvPicPr");
      altText = meta.altText;
      title = meta.title;
    }
  }

  if (!embedId) return null;

  const mediaPath = relMap.get(embedId);
  if (!mediaPath) return null;

  const ext = mediaPath.split(".").pop()?.toLowerCase() ?? "";
  const imageType = EXT_TO_IMAGE_TYPE[ext] ?? "png";

  const result: {
    mediaPath: string;
    type: SheetImage["type"];
    anchor: SheetImage["anchor"];
    width?: number;
    height?: number;
    altText?: string;
    title?: string;
  } = {
    mediaPath,
    type: imageType,
    anchor: {
      from: { row: fromRow, col: fromCol },
    },
  };

  if (widthEmu > 0) {
    result.width = Math.round(widthEmu / EMU_PER_PIXEL);
  }
  if (heightEmu > 0) {
    result.height = Math.round(heightEmu / EMU_PER_PIXEL);
  }
  if (altText) result.altText = altText;
  if (title) result.title = title;

  return result;
}

/**
 * Walk a `<xdr:pic>` or `<xdr:sp>` element and extract `descr=`/`title=`
 * from its `xdr:cNvPr`. The cNvPr element lives inside an `nv*Pr`
 * wrapper named `nvPicPr` (pictures) or `nvSpPr` (shapes). Returns
 * empty fields when neither attribute is present.
 */
function findCNvPrMeta(
  parentEl: { children: Array<unknown> },
  wrapperName: "nvPicPr" | "nvSpPr",
): { altText?: string; title?: string } {
  const wrapper = findChildEl(parentEl, wrapperName);
  if (!wrapper) return {};
  const cNvPr = findChildEl(wrapper, "cNvPr");
  if (!cNvPr) return {};
  const out: { altText?: string; title?: string } = {};
  if (cNvPr.attrs["descr"]) out.altText = cNvPr.attrs["descr"];
  if (cNvPr.attrs["title"]) out.title = cNvPr.attrs["title"];
  return out;
}

/** Parse row/col from an anchor position element (from or to) */
function parseAnchorPosition(el: { children: Array<unknown> }): { row: number; col: number } {
  let row = 0;
  let col = 0;

  for (const child of el.children) {
    if (typeof child === "string") continue;
    const c = child as { local?: string; tag: string; children: Array<unknown> };
    const local = c.local || c.tag;
    const text = c.children.filter((ch: unknown) => typeof ch === "string").join("");

    if (local === "row") {
      row = Number(text) || 0;
    } else if (local === "col") {
      col = Number(text) || 0;
    }
  }

  return { row, col };
}

/** Find the r:embed attribute on the blip element inside a pic element */
function findBlipEmbed(picEl: { children: Array<unknown> }): string | undefined {
  for (const child of picEl.children) {
    if (typeof child === "string") continue;
    const c = child as {
      local?: string;
      tag: string;
      children: Array<unknown>;
      attrs: Record<string, string>;
    };
    const local = c.local || c.tag;

    if (local === "blipFill") {
      for (const blipChild of c.children) {
        if (typeof blipChild === "string") continue;
        const bc = blipChild as { local?: string; tag: string; attrs: Record<string, string> };
        const blipLocal = bc.local || bc.tag;
        if (blipLocal === "blip") {
          // Look for r:embed attribute (namespace prefix may vary)
          return bc.attrs["r:embed"] ?? bc.attrs["R:embed"] ?? findEmbedAttr(bc.attrs);
        }
      }
    }
  }
  return undefined;
}

/** Find an embed attribute regardless of namespace prefix */
function findEmbedAttr(attrs: Record<string, string>): string | undefined {
  for (const key of Object.keys(attrs)) {
    if (key.endsWith(":embed") && attrs[key].startsWith("rId")) {
      return attrs[key];
    }
  }
  return undefined;
}

// ── Workbook XML Parsing ─────────────────────────────────────────────

interface SheetInfo {
  name: string;
  sheetId: number;
  rId: string;
  state?: "visible" | "hidden" | "veryHidden";
}

function parseWorkbookXml(
  xml: string,
  options?: ReadOptions,
): {
  sheets: SheetInfo[];
  dateSystem: "1900" | "1904";
  namedRanges: NamedRange[];
  workbookProtection?: { lockStructure?: boolean; lockWindows?: boolean };
  /**
   * Pivot cache wiring read off the workbook's `<pivotCaches>` block.
   * Each entry maps a cacheId (Excel's stable handle) to an rId in
   * `xl/_rels/workbook.xml.rels`.
   */
  pivotCacheRefs: Array<{ cacheId: number; rId: string }>;
} {
  const doc = parseXml(xml);

  const sheets: SheetInfo[] = [];
  const namedRanges: NamedRange[] = [];
  const pivotCacheRefs: Array<{ cacheId: number; rId: string }> = [];
  let dateSystem: "1900" | "1904" = "1900";

  // Check date system override from options
  if (options?.dateSystem === "1904") {
    dateSystem = "1904";
  } else if (options?.dateSystem === "1900") {
    dateSystem = "1900";
  }

  let wbProtection: { lockStructure?: boolean; lockWindows?: boolean } | undefined;

  // First pass: collect sheets (needed for resolving localSheetId)
  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "workbookPr") {
      // Check for 1904 date system
      if (child.attrs["date1904"] === "1" || child.attrs["date1904"] === "true") {
        // Only override if auto or not set
        if (!options?.dateSystem || options.dateSystem === "auto") {
          dateSystem = "1904";
        }
      }
    }

    if (local === "workbookProtection") {
      const lockStructure =
        child.attrs["lockStructure"] === "1" || child.attrs["lockStructure"] === "true";
      const lockWindows =
        child.attrs["lockWindows"] === "1" || child.attrs["lockWindows"] === "true";
      if (lockStructure || lockWindows) {
        wbProtection = {};
        if (lockStructure) wbProtection.lockStructure = true;
        if (lockWindows) wbProtection.lockWindows = true;
      }
    }

    if (local === "sheets") {
      for (const sheetChild of child.children) {
        if (typeof sheetChild === "string") continue;
        const sheetLocal = sheetChild.local || sheetChild.tag;
        if (sheetLocal === "sheet") {
          const name = sheetChild.attrs["name"] ?? "";
          const sheetId = Number(sheetChild.attrs["sheetId"] ?? "0");
          // r:id attribute — the namespace prefix may vary
          const rId =
            sheetChild.attrs["r:id"] ??
            sheetChild.attrs["R:id"] ??
            findRIdAttr(sheetChild.attrs) ??
            "";
          const stateRaw = sheetChild.attrs["state"];
          let state: SheetInfo["state"] = "visible";
          if (stateRaw === "hidden") state = "hidden";
          else if (stateRaw === "veryHidden") state = "veryHidden";

          if (name && rId) {
            sheets.push({ name, sheetId, rId, state });
          }
        }
      }
    }

    if (local === "pivotCaches") {
      for (const pcChild of child.children) {
        if (typeof pcChild === "string") continue;
        const pcLocal = pcChild.local || pcChild.tag;
        if (pcLocal === "pivotCache") {
          const cacheIdRaw = pcChild.attrs["cacheId"];
          const cacheId = cacheIdRaw === undefined ? NaN : Number(cacheIdRaw);
          const rId =
            pcChild.attrs["r:id"] ?? pcChild.attrs["R:id"] ?? findRIdAttr(pcChild.attrs) ?? "";
          if (rId && !Number.isNaN(cacheId)) {
            pivotCacheRefs.push({ cacheId, rId });
          }
        }
      }
    }
  }

  // Second pass: collect defined names (named ranges)
  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "definedNames") {
      for (const dnChild of child.children) {
        if (typeof dnChild === "string") continue;
        const dnLocal = dnChild.local || dnChild.tag;
        if (dnLocal === "definedName") {
          const name = dnChild.attrs["name"] ?? "";
          const rangeText = dnChild.children.filter((c: unknown) => typeof c === "string").join("");

          if (name && rangeText) {
            const nr: NamedRange = { name, range: rangeText };

            // Resolve localSheetId to sheet name
            const localSheetId = dnChild.attrs["localSheetId"];
            if (localSheetId !== undefined) {
              const idx = Number(localSheetId);
              if (idx >= 0 && idx < sheets.length) {
                nr.scope = sheets[idx].name;
              }
            }

            // Comment attribute
            if (dnChild.attrs["comment"]) {
              nr.comment = dnChild.attrs["comment"];
            }

            namedRanges.push(nr);
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
    pivotCacheRefs,
  };
}

/** Find an r:id attribute regardless of namespace prefix */
function findRIdAttr(attrs: Record<string, string>): string | undefined {
  for (const key of Object.keys(attrs)) {
    // Match any prefix:id where the value looks like an rId
    if (key.endsWith(":id") && attrs[key].startsWith("rId")) {
      return attrs[key];
    }
  }
  return undefined;
}

/** Filter sheet infos based on user-specified sheets option */
function filterSheets(allSheets: SheetInfo[], filter?: ReadOptions["sheets"]): SheetInfo[] {
  if (filter === undefined) return allSheets;

  if (typeof filter === "function") {
    const result: SheetInfo[] = [];
    for (let i = 0; i < allSheets.length; i++) {
      const info = allSheets[i]!;
      const decision = filter(
        {
          name: info.name,
          index: i,
          hidden: info.state === "hidden",
          veryHidden: info.state === "veryHidden",
        },
        i,
      );
      if (decision) result.push(info);
    }
    return result;
  }

  if (filter.length === 0) return allSheets;

  const result: SheetInfo[] = [];
  for (const spec of filter) {
    if (typeof spec === "number") {
      if (spec >= 0 && spec < allSheets.length) {
        result.push(allSheets[spec]);
      }
    } else {
      const found = allSheets.find((s) => s.name === spec);
      if (found) result.push(found);
    }
  }

  return result;
}

// ── Table XML Parsing ────────────────────────────────────────────────

/**
 * Parse a table XML file (xl/tables/tableN.xml) into a TableDefinition.
 */
function parseTableXml(xml: string): TableDefinition | null {
  const doc = parseXml(xml);

  // Root element should be <table>
  const name = doc.attrs["name"] ?? "";
  const displayName = doc.attrs["displayName"] ?? name;
  const ref = doc.attrs["ref"] ?? "";

  if (!name) return null;

  // Determine showTotalRow from totalsRowCount or totalsRowShown
  const totalsRowCount = doc.attrs["totalsRowCount"];
  const showTotalRow = totalsRowCount !== undefined && totalsRowCount !== "0";

  const columns: TableColumn[] = [];
  let style: string | undefined;
  let showRowStripes: boolean | undefined;
  let showColumnStripes: boolean | undefined;
  let showAutoFilter = true;

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "autoFilter") {
      showAutoFilter = true;
    } else if (local === "tableColumns") {
      for (const colChild of child.children) {
        if (typeof colChild === "string") continue;
        const colLocal = colChild.local || colChild.tag;
        if (colLocal === "tableColumn") {
          const col: TableColumn = {
            name: colChild.attrs["name"] ?? "",
          };
          if (colChild.attrs["totalsRowFunction"]) {
            col.totalFunction = colChild.attrs["totalsRowFunction"];
          }
          if (colChild.attrs["totalsRowLabel"]) {
            col.totalLabel = colChild.attrs["totalsRowLabel"];
          }
          // Parse totalsRowFormula child element
          for (const formulaChild of colChild.children) {
            if (typeof formulaChild === "string") continue;
            const formulaLocal = formulaChild.local || formulaChild.tag;
            if (formulaLocal === "totalsRowFormula") {
              const formulaText = formulaChild.children
                .filter((c: unknown) => typeof c === "string")
                .join("");
              if (formulaText) {
                col.totalFormula = formulaText;
              }
            }
          }
          columns.push(col);
        }
      }
    } else if (local === "tableStyleInfo") {
      style = child.attrs["name"];
      showRowStripes = child.attrs["showRowStripes"] === "1";
      showColumnStripes = child.attrs["showColumnStripes"] === "1";
    }
  }

  const tableDef: TableDefinition = {
    name,
    columns,
  };

  if (displayName && displayName !== name) {
    tableDef.displayName = displayName;
  }
  if (ref) {
    tableDef.range = ref;
  }
  if (style) {
    tableDef.style = style;
  }
  if (showRowStripes !== undefined) {
    tableDef.showRowStripes = showRowStripes;
  }
  if (showColumnStripes !== undefined) {
    tableDef.showColumnStripes = showColumnStripes;
  }
  if (showAutoFilter !== undefined) {
    tableDef.showAutoFilter = showAutoFilter;
  }
  if (showTotalRow) {
    tableDef.showTotalRow = true;
  }

  return tableDef;
}
