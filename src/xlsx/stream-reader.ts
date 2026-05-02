// ── Streaming XLSX Reader ────────────────────────────────────────────
// Yields rows one at a time from an XLSX file via SAX parsing.
// Parses shared strings and styles upfront (small), then streams
// worksheet rows without buffering the entire sheet in memory.

import type { CellValue, ReadOptions } from "../_types";
import type { SharedString } from "./shared-strings";
import type { ParsedStyles } from "./styles";
import type { Relationship } from "./relationships";
import { ParseError, ZipError } from "../errors";
import { assertNotEncrypted, bufferReadableStream } from "../_input";
import { ZipReader } from "../zip/reader";
import { parseXml, parseSaxStream, decodeOoxmlEscapes } from "../xml/parser";
import { parseContentTypes } from "./content-types";
import { parseRelationships } from "./relationships";
import { parseSharedStrings } from "./shared-strings";
import { parseStyles, isDateStyle } from "./styles";
import { parseCellRef } from "./worksheet";
import { serialToDate } from "../_date";

// ── Types ────────────────────────────────────────────────────────────

export interface StreamRow {
  /** 0-based row index */
  index: number;
  /** Cell values for this row */
  values: CellValue[];
}

// ── Range filter ────────────────────────────────────────────────────

interface RangeFilter {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

/**
 * Parse a range reference like "A1:D10" into 0-based row/col bounds.
 * Single-cell refs like "B2" are also accepted (start == end).
 */
function parseRangeFilter(ref: string): RangeFilter {
  const parts = ref.split(":");
  if (parts.length === 0 || parts.length > 2) {
    throw new ParseError(`Invalid range reference: "${ref}"`);
  }
  const start = parseCellRef(parts[0]!);
  const end = parts.length > 1 ? parseCellRef(parts[1]!) : start;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

// ── OOXML Relationship Types ─────────────────────────────────────────

const REL_WORKBOOK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SHARED_STRINGS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

// ── Helpers ──────────────────────────────────────────────────────────

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data);
}

function resolvePath(base: string, target: string): string {
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

function dirname(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

// ── Workbook XML Parsing (minimal — just sheet info + date system) ───

interface SheetInfo {
  name: string;
  sheetId: number;
  rId: string;
}

function parseWorkbookXml(
  xml: string,
  options?: ReadOptions,
): { sheets: SheetInfo[]; dateSystem: "1900" | "1904" } {
  const doc = parseXml(xml);
  const sheets: SheetInfo[] = [];
  let dateSystem: "1900" | "1904" = "1900";

  if (options?.dateSystem === "1904") {
    dateSystem = "1904";
  } else if (options?.dateSystem === "1900") {
    dateSystem = "1900";
  }

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "workbookPr") {
      if (child.attrs["date1904"] === "1" || child.attrs["date1904"] === "true") {
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
          const rId =
            sheetChild.attrs["r:id"] ??
            sheetChild.attrs["R:id"] ??
            findRIdAttr(sheetChild.attrs) ??
            "";

          if (name && rId) {
            sheets.push({ name, sheetId, rId });
          }
        }
      }
    }
  }

  return { sheets, dateSystem };
}

function findRIdAttr(attrs: Record<string, string>): string | undefined {
  for (const key of Object.keys(attrs)) {
    if (key.endsWith(":id") && attrs[key].startsWith("rId")) {
      return attrs[key];
    }
  }
  return undefined;
}

// ── Resolve target sheet ────────────────────────────────────────────

function resolveTargetSheet(allSheets: SheetInfo[], sheetSpec?: number | string): SheetInfo | null {
  if (sheetSpec === undefined) {
    // Default: first sheet
    return allSheets[0] ?? null;
  }

  if (typeof sheetSpec === "number") {
    return sheetSpec >= 0 && sheetSpec < allSheets.length ? allSheets[sheetSpec] : null;
  }

  return allSheets.find((s) => s.name === sheetSpec) ?? null;
}

// ── Worksheet SAX handlers (shared between sync and streaming paths) ─

interface RowSaxState {
  inSheetData: boolean;
  inRow: boolean;
  inCell: boolean;
  inValue: boolean;
  inFormula: boolean;
  inInlineStr: boolean;
  inInlineT: boolean;
  inInlineR: boolean;
  inInlineRT: boolean;
  currentRowIndex: number;
  currentRowCells: Array<{ col: number; value: CellValue }>;
  cellRef: string;
  cellType: string;
  cellStyleIndex: number;
  cellValueText: string;
  inlineText: string;
  inlineRichTextParts: string[];
  currentRunText: string;
  /** Implicit column counter for cells without r attribute */
  implicitCol: number;
}

function createRowSaxState(): RowSaxState {
  return {
    inSheetData: false,
    inRow: false,
    inCell: false,
    inValue: false,
    inFormula: false,
    inInlineStr: false,
    inInlineT: false,
    inInlineR: false,
    inInlineRT: false,
    currentRowIndex: -1,
    currentRowCells: [],
    cellRef: "",
    cellType: "",
    cellStyleIndex: -1,
    cellValueText: "",
    inlineText: "",
    inlineRichTextParts: [],
    currentRunText: "",
    implicitCol: 0,
  };
}

function buildRowFromCells(cells: Array<{ col: number; value: CellValue }>): CellValue[] {
  // Use reduce instead of Math.max(...spread) to avoid RangeError on wide rows (>65K cols)
  const maxCol = cells.reduce((m, c) => (c.col > m ? c.col : m), -1);
  const values: CellValue[] = maxCol >= 0 ? Array.from({ length: maxCol + 1 }, () => null) : [];
  for (const cell of cells) {
    values[cell.col] = cell.value;
  }
  return values;
}

function handleOpenTag(tag: string, attrs: Record<string, string>, s: RowSaxState): void {
  const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

  switch (local) {
    case "sheetData":
      s.inSheetData = true;
      break;
    case "row":
      if (s.inSheetData) {
        s.inRow = true;
        s.currentRowIndex = attrs["r"] ? Number(attrs["r"]) - 1 : s.currentRowIndex + 1;
        s.currentRowCells = [];
        s.implicitCol = 0;
      }
      break;
    case "c":
      if (s.inRow) {
        s.inCell = true;
        s.cellRef = attrs["r"] ?? "";
        s.cellType = attrs["t"] ?? "";
        s.cellStyleIndex = attrs["s"] ? Number(attrs["s"]) : -1;
        s.cellValueText = "";
        s.inlineText = "";
        s.inlineRichTextParts = [];
      }
      break;
    case "v":
      if (s.inCell) s.inValue = true;
      break;
    case "f":
      if (s.inCell) s.inFormula = true;
      break;
    case "is":
      if (s.inCell) s.inInlineStr = true;
      break;
    case "t":
      if (s.inInlineStr && !s.inInlineR) {
        s.inInlineT = true;
      } else if (s.inInlineR) {
        s.inInlineRT = true;
      }
      break;
    case "r":
      if (s.inInlineStr) {
        s.inInlineR = true;
        s.currentRunText = "";
      }
      break;
  }
}

function handleText(text: string, s: RowSaxState): void {
  if (s.inValue) {
    s.cellValueText += text;
  } else if (s.inInlineT) {
    s.inlineText += text;
  } else if (s.inInlineRT) {
    s.currentRunText += text;
  }
}

/**
 * Handle a closing tag. Returns a completed StreamRow if a row just ended, otherwise null.
 */
function handleCloseTag(
  tag: string,
  s: RowSaxState,
  sharedStrings: SharedString[],
  styles: ParsedStyles | null,
  dateSystem: "1900" | "1904",
): StreamRow | null {
  const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

  switch (local) {
    case "sheetData":
      s.inSheetData = false;
      break;
    case "row":
      if (s.inRow) {
        const values = buildRowFromCells(s.currentRowCells);
        const row: StreamRow = { index: s.currentRowIndex, values };
        s.inRow = false;
        return row;
      }
      break;
    case "c":
      if (s.inCell) {
        const value = resolveStreamCellValue(
          s.cellType,
          s.cellStyleIndex,
          s.cellValueText,
          s.inlineText,
          s.inlineRichTextParts,
          sharedStrings,
          styles,
          dateSystem,
        );
        if (s.cellRef) {
          const pos = parseCellRef(s.cellRef);
          s.currentRowCells.push({ col: pos.col, value });
          s.implicitCol = pos.col + 1;
        } else {
          // Fallback: cells without r attribute use implicit column ordering
          s.currentRowCells.push({ col: s.implicitCol, value });
          s.implicitCol++;
        }
        s.inCell = false;
      }
      break;
    case "v":
      s.inValue = false;
      break;
    case "f":
      s.inFormula = false;
      break;
    case "is":
      s.inInlineStr = false;
      break;
    case "t":
      if (s.inInlineRT) {
        s.inInlineRT = false;
      } else if (s.inInlineT) {
        s.inInlineT = false;
      }
      break;
    case "r":
      if (s.inInlineR) {
        s.inlineRichTextParts.push(decodeOoxmlEscapes(s.currentRunText));
        s.inInlineR = false;
      }
      break;
  }
  return null;
}

// ── Filter application ─────────────────────────────────────────────

/**
 * Apply the range filter to a freshly-yielded row. Returns the row to emit
 * (with cells outside the column range nulled out) or `null` if the row
 * itself falls outside the row range.
 *
 * Mirrors the batch reader (`parseWorksheet`): values stay aligned to
 * their original column index, and cells outside the column window are
 * masked to `null` rather than removed, so callers can still address
 * `row.values[colIndex]` for columns inside the range.
 */
function applyRangeFilter(row: StreamRow, range: RangeFilter): StreamRow | null {
  if (row.index < range.startRow || row.index > range.endRow) return null;
  const len = Math.max(row.values.length, range.endCol + 1);
  const out: CellValue[] = Array.from({ length: len }, () => null);
  const upper = Math.min(row.values.length - 1, range.endCol);
  for (let c = range.startCol; c <= upper; c++) {
    out[c] = row.values[c] ?? null;
  }
  return { index: row.index, values: out };
}

// ── Streaming row parser via SAX (async — ReadableStream) ──────────

async function* parseWorksheetRowsStreaming(
  stream: ReadableStream<Uint8Array>,
  sharedStrings: SharedString[],
  styles: ParsedStyles | null,
  dateSystem: "1900" | "1904",
  filters: { range?: RangeFilter; maxRows?: number } = {},
): AsyncGenerator<StreamRow, void, undefined> {
  const s = createRowSaxState();
  const pendingRows: StreamRow[] = [];
  let resolve: (() => void) | null = null;
  let done = false;
  let aborted = false;

  // Wrap the source reader so we can short-circuit chunk pulls (and cancel
  // the underlying ZIP/decompression stream) once a stop condition fires.
  // We hold the source reader exclusively here so that calling
  // `cancel(reason)` propagates without conflicting with locks.
  const sourceReader = stream.getReader();
  const cancelSource = (reason?: unknown): void => {
    sourceReader.cancel(reason).catch(() => {});
  };
  const cancellable = new ReadableStream<Uint8Array>({
    async pull(controller) {
      if (aborted) {
        controller.close();
        return;
      }
      try {
        const { done: rDone, value } = await sourceReader.read();
        if (rDone) {
          controller.close();
          return;
        }
        controller.enqueue(value);
      } catch (err) {
        controller.error(err);
      }
    },
    cancel(reason) {
      cancelSource(reason);
    },
  });

  let emittedDataRows = 0;
  const maxRows = filters.maxRows ?? 0;
  const range = filters.range;

  const parsePromise = parseSaxStream(cancellable, {
    onOpenTag(tag, attrs) {
      if (aborted) return;
      handleOpenTag(tag, attrs, s);
    },
    onText(text) {
      if (aborted) return;
      handleText(text, s);
    },
    onCloseTag(tag) {
      if (aborted) return;
      const row = handleCloseTag(tag, s, sharedStrings, styles, dateSystem);
      if (row) {
        // If the SAX-emitted row is past the range end, we can stop now —
        // worksheet rows are written in ascending order in valid OOXML.
        if (range && row.index > range.endRow) {
          aborted = true;
          cancelSource();
          if (resolve) {
            resolve();
            resolve = null;
          }
          return;
        }
        const filtered = range ? applyRangeFilter(row, range) : row;
        if (filtered) {
          pendingRows.push(filtered);
          emittedDataRows++;
          if (resolve) {
            resolve();
            resolve = null;
          }
          if (maxRows > 0 && emittedDataRows >= maxRows) {
            aborted = true;
            cancelSource();
          }
        }
      }
    },
  }).then(() => {
    done = true;
    if (resolve) {
      resolve();
      resolve = null;
    }
  });

  try {
    while (!done || pendingRows.length > 0) {
      if (pendingRows.length > 0) {
        yield pendingRows.shift()!;
      } else if (!done) {
        await new Promise<void>((r) => {
          resolve = r;
        });
      }
    }
  } finally {
    // Release the upstream reader if the consumer abandoned the generator
    // before the stream finished. Cancellation is idempotent — if we've
    // already cancelled because of maxRows/range, this is a no-op.
    aborted = true;
    cancelSource();
  }

  await parsePromise.catch(() => {});
}

// ── Cell value resolution (streaming — no Cell objects) ──────────────

function resolveStreamCellValue(
  type: string,
  styleIndex: number,
  valueText: string,
  inlineText: string,
  inlineRichTextParts: string[],
  sharedStrings: SharedString[],
  styles: ParsedStyles | null,
  dateSystem: "1900" | "1904",
): CellValue {
  switch (type) {
    case "s": {
      // Shared string
      const idx = Number(valueText);
      if (!Number.isNaN(idx) && idx >= 0 && idx < sharedStrings.length) {
        return sharedStrings[idx].text;
      }
      return null; // Out-of-bounds SST index — return null, not the raw index string
    }
    case "str": {
      // Inline formula string result
      return decodeOoxmlEscapes(valueText);
    }
    case "inlineStr": {
      // Inline string with <is> element
      if (inlineRichTextParts.length > 0) {
        return inlineRichTextParts.join("");
      }
      return decodeOoxmlEscapes(inlineText);
    }
    case "b": {
      // Boolean
      return valueText === "1" || valueText.toLowerCase() === "true";
    }
    case "e": {
      // Error
      return valueText;
    }
    case "n":
    default: {
      // Number (explicit or implied)
      if (valueText === "") {
        return null;
      }

      const num = Number(valueText);
      if (!Number.isNaN(num)) {
        // Check if this is a date via style
        if (styles && styleIndex >= 0 && isDateStyle(styles, styleIndex)) {
          return serialToDate(num, dateSystem === "1904");
        }
        return num;
      }
      return valueText || null;
    }
  }
}

// ── Main streaming reader ───────────────────────────────────────────

/**
 * Create an async iterable that yields rows one at a time.
 * Parses shared strings and styles upfront (they're small),
 * then streams worksheet rows via SAX parsing.
 *
 * Accepts Uint8Array, ArrayBuffer, or ReadableStream<Uint8Array>.
 * For ReadableStream input, the stream is buffered to read the ZIP
 * central directory, then the worksheet entry is stream-decompressed
 * and piped through the SAX parser in chunks.
 *
 * Honored {@link ReadOptions} fields:
 * - `sheet` — target sheet (number index or name). Default: first sheet.
 * - `dateSystem` — `"1900"` / `"1904"` / `"auto"`. Default: auto-detect.
 * - `range` — A1-style range filter (e.g. `"B2:D10"`). Rows outside the
 *   row span are skipped; cells outside the column span are nulled out.
 *   Parsing stops once a row past the end-row is observed.
 * - `maxRows` — caps the number of rows yielded. Once the cap is hit the
 *   underlying ZIP/SAX stream is cancelled so no further work is done.
 */
export async function* streamXlsxRows(
  input: Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>,
  options?: ReadOptions & { sheet?: number | string },
): AsyncGenerator<StreamRow, void, undefined> {
  // Normalize input to Uint8Array for ZIP central directory parsing.
  // ReadableStream must be fully buffered because ZIP central directory
  // is at the end of the file.
  let data: Uint8Array;
  if (input instanceof Uint8Array) {
    data = input;
  } else if (input instanceof ArrayBuffer) {
    data = new Uint8Array(input);
  } else {
    data = await bufferReadableStream(input);
  }

  // Detect password-protected workbooks (OLE2/CFB envelope) up front so
  // streamers fail fast with a typed `EncryptedFileError` instead of a
  // generic ZIP ParseError. Decryption is tracked in #156.
  assertNotEncrypted(data, "xlsx");

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

  // 2. Validate content types
  if (!zip.has("[Content_Types].xml")) {
    throw new ParseError("Invalid XLSX: missing [Content_Types].xml");
  }
  const contentTypesXml = decodeUtf8(await zip.extract("[Content_Types].xml"));
  parseContentTypes(contentTypesXml);

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

  // 4. Parse workbook relationships
  const workbookDir = dirname(workbookPath);
  const workbookRelsPath = workbookDir
    ? `${workbookDir}/_rels/${workbookPath.slice(workbookDir.length + 1)}.rels`
    : `_rels/${workbookPath}.rels`;

  let workbookRels: Relationship[] = [];
  if (zip.has(workbookRelsPath)) {
    const wbRelsXml = decodeUtf8(await zip.extract(workbookRelsPath));
    workbookRels = parseRelationships(wbRelsXml);
  }

  // 5. Parse workbook XML for sheet names and date system
  if (!zip.has(workbookPath)) {
    throw new ParseError(`Invalid XLSX: missing workbook at ${workbookPath}`);
  }
  const workbookXml = decodeUtf8(await zip.extract(workbookPath));
  const { sheets: sheetInfos, dateSystem } = parseWorkbookXml(workbookXml, options);

  // 6. Parse shared strings (small, needed for cell resolution)
  let sharedStrings: SharedString[] = [];
  const ssRel = workbookRels.find((r) => r.type === REL_SHARED_STRINGS);
  if (ssRel) {
    const ssPath = resolvePath(workbookDir, ssRel.target);
    if (zip.has(ssPath)) {
      const ssXml = decodeUtf8(await zip.extract(ssPath));
      sharedStrings = parseSharedStrings(ssXml);
    }
  }

  // 7. Parse styles (needed for date detection)
  let parsedStyles: ParsedStyles | null = null;
  const stylesRel = workbookRels.find((r) => r.type === REL_STYLES);
  if (stylesRel) {
    const stylesPath = resolvePath(workbookDir, stylesRel.target);
    if (zip.has(stylesPath)) {
      const stylesXml = decodeUtf8(await zip.extract(stylesPath));
      parsedStyles = parseStyles(stylesXml);
    }
  }

  // 8. Build rId → worksheet path map
  const sheetRelMap = new Map<string, string>();
  for (const rel of workbookRels) {
    if (rel.type === REL_WORKSHEET) {
      sheetRelMap.set(rel.id, resolvePath(workbookDir, rel.target));
    }
  }

  // 9. Resolve the target sheet
  const targetSheet = resolveTargetSheet(sheetInfos, options?.sheet);
  if (!targetSheet) {
    return; // No matching sheet — yield nothing
  }

  const wsPath = sheetRelMap.get(targetSheet.rId);
  if (!wsPath || !zip.has(wsPath)) {
    throw new ParseError(`Invalid XLSX: missing worksheet file for sheet "${targetSheet.name}"`);
  }

  // 10. Build optional row/cell filters from ReadOptions.
  // `range` and `maxRows` mirror the batch-reader semantics: range filters
  // both rows (skip outside) and cells (mask outside columns), maxRows
  // caps the number of yielded rows. Both stop pulling from the worksheet
  // stream as soon as no more rows can be emitted.
  let rangeFilter: RangeFilter | undefined;
  if (options?.range) {
    rangeFilter = parseRangeFilter(options.range);
  }
  const maxRowsLimit =
    typeof options?.maxRows === "number" && options.maxRows > 0 ? options.maxRows : 0;

  // 11. Stream worksheet rows
  // Use streaming decompression: pipe ZIP entry through DecompressionStream
  // directly into the SAX parser, yielding rows as they complete.
  const wsStream = zip.extractStream(wsPath);
  yield* parseWorksheetRowsStreaming(wsStream, sharedStrings, parsedStyles, dateSystem, {
    range: rangeFilter,
    maxRows: maxRowsLimit > 0 ? maxRowsLimit : undefined,
  });
}
