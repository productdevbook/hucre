// ── XLSX Reader ──────────────────────────────────────────────────────
// Reads Office Open XML (.xlsx) spreadsheet files.

import type { Workbook, ReadOptions, ReadInput } from "../_types";
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

// ── OOXML Relationship Types ─────────────────────────────────────────

const REL_WORKBOOK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SHARED_STRINGS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

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
  const workbookRel = rootRels.find((r) => r.type === REL_WORKBOOK);
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
  const { sheets: sheetInfos, dateSystem } = parseWorkbookXml(workbookXml, options);

  // 6. Parse shared strings if present
  let sharedStrings: SharedString[] = [];
  const ssRel = workbookRels.find((r) => r.type === REL_SHARED_STRINGS);
  if (ssRel) {
    const ssPath = resolvePath(workbookDir, ssRel.target);
    if (zip.has(ssPath)) {
      const ssXml = decodeUtf8(await zip.extract(ssPath));
      sharedStrings = parseSharedStrings(ssXml);
    }
  }

  // 7. Parse styles if needed (for date detection or if readStyles is true)
  let parsedStyles: ParsedStyles | null = null;
  const stylesRel = workbookRels.find((r) => r.type === REL_STYLES);
  if (stylesRel) {
    const stylesPath = resolvePath(workbookDir, stylesRel.target);
    if (zip.has(stylesPath)) {
      const stylesXml = decodeUtf8(await zip.extract(stylesPath));
      parsedStyles = parseStyles(stylesXml);
    }
  }

  // 8. Build a map of rId → sheet relationship for worksheet paths
  const sheetRelMap = new Map<string, string>();
  for (const rel of workbookRels) {
    if (rel.type === REL_WORKSHEET) {
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
    };

    const wsXml = decodeUtf8(await zip.extract(wsPath));
    const sheet = parseWorksheet(wsXml, info.name, worksheetCtx);
    if (info.state === "hidden") sheet.hidden = true;
    if (info.state === "veryHidden") sheet.veryHidden = true;
    sheets.push(sheet);
  }

  // 11. Build workbook
  const workbook: Workbook = {
    sheets,
    dateSystem,
  };

  return workbook;
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
): { sheets: SheetInfo[]; dateSystem: "1900" | "1904" } {
  const doc = parseXml(xml);

  const sheets: SheetInfo[] = [];
  let dateSystem: "1900" | "1904" = "1900";

  // Check date system override from options
  if (options?.dateSystem === "1904") {
    dateSystem = "1904";
  } else if (options?.dateSystem === "1900") {
    dateSystem = "1900";
  }

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
  }

  return { sheets, dateSystem };
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
function filterSheets(allSheets: SheetInfo[], filter?: Array<number | string>): SheetInfo[] {
  if (!filter || filter.length === 0) return allSheets;

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
