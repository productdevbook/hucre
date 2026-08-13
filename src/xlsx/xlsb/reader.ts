// ── XLSB Reader ──────────────────────────────────────────────────────
// Read Excel Binary Workbook (.xlsb) files. XLSB shares XLSX's ZIP package
// and XML relationships/[Content_Types], but stores the workbook, shared
// strings, styles, and worksheets as binary `.bin` record streams instead
// of XML. We reuse the ZIP layer + the XML rels parser, and decode the
// `.bin` parts with the record reader. Read-only (MS-XLSB).

import type { CellValue, MergeRange, ReadOptions, Sheet, Workbook } from "../../_types"
import { EncryptedFileError, ParseError, ZipError } from "../../errors"
import { isOle2Container, readInputToUint8Array } from "../../_input"
import { decryptAgile } from "../crypto/agile"
import { ZipReader } from "../../zip/reader"
import { decodePart } from "../../_decode"
import { parseRelationships } from "../relationships"
import { matchesRelType } from "../reader"
import { densify } from "../../xls/reader"
import { isBuiltinDateFormatId, isDateFormat, serialToDate } from "../../_date"
import { Cursor, decodeRk, iterateRecords } from "./record"
import { MAX_COL_INDEX, MAX_ROW_INDEX } from "../../limits"

// Record ids we care about (MS-XLSB §2.4).
const BrtRowHdr = 0
const BrtCellBlank = 1
const BrtCellRk = 2
const BrtCellError = 3
const BrtCellBool = 4
const BrtCellReal = 5
const BrtCellSt = 6
const BrtCellIsst = 7
const BrtFmlaString = 8
const BrtFmlaNum = 9
const BrtFmlaBool = 10
const BrtFmlaError = 11
const BrtSSTItem = 19
const BrtFmt = 44
const BrtXF = 47
const BrtWbProp = 153
const BrtBundleSh = 156
const BrtMergeCell = 176
const BrtBeginCellXFs = 617
const BrtEndCellXFs = 618

// Short relationship type names; matched leniently (Transitional + Strict)
// via matchesRelType, consistent with the batch XLSX reader.
const REL_WORKBOOK = "officeDocument"
const REL_WORKSHEET = "worksheet"
const REL_SHARED_STRINGS = "sharedStrings"
const REL_STYLES = "styles"

const ERROR_TEXT: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
  0x2b: "#GETTING_DATA",
}

const decoder = new TextDecoder("utf-8")

function decodeUtf8(d: Uint8Array, path = "(unknown)"): string {
  return decodePart(d, path)
}

function dirname(path: string): string {
  const i = path.lastIndexOf("/")
  return i === -1 ? "" : path.slice(0, i)
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1)
  const parts = base.split("/").filter(Boolean)
  for (const seg of target.split("/").filter(Boolean)) {
    if (seg === "..") parts.pop()
    else if (seg !== ".") parts.push(seg)
  }
  return parts.join("/")
}

/** Whether an .xlsb-shaped ZIP is present (used for format auto-detection). */
export function looksLikeXlsb(zip: ZipReader): boolean {
  return zip.has("xl/workbook.bin")
}

/**
 * Read an XLSB workbook into the same {@link Workbook} model `readXlsx`
 * produces (sheet names + typed cell values in `rows`).
 */
export async function readXlsb(
  input: Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>,
  options?: ReadOptions,
): Promise<Workbook> {
  let data = await readInputToUint8Array(input, options?.maxInputBytes)
  if (isOle2Container(data)) {
    if (options?.password) data = await decryptAgile(data, options.password, options.maxSpinCount)
    else throw new EncryptedFileError("xlsx")
  }

  let zip: ZipReader
  try {
    zip = new ZipReader(data, options?.maxDecompressedBytes)
  } catch (err) {
    if (err instanceof ZipError) throw err
    throw new ParseError("Failed to open XLSB: not a valid ZIP archive", undefined, { cause: err })
  }

  // Locate the workbook part via the root relationships (fallback to the
  // conventional path).
  let workbookPath = "xl/workbook.bin"
  if (zip.has("_rels/.rels")) {
    const rootRels = parseRelationships(decodeUtf8(await zip.extract("_rels/.rels"), "_rels/.rels"))
    const wbRel = rootRels.find((r) => matchesRelType(r.type, REL_WORKBOOK))
    if (wbRel) workbookPath = wbRel.target.startsWith("/") ? wbRel.target.slice(1) : wbRel.target
  }
  if (!zip.has(workbookPath)) {
    throw new ParseError(`Invalid XLSB: missing workbook at ${workbookPath}`)
  }
  const workbookDir = dirname(workbookPath)

  // Binary record parsing reads many length-prefixed fields; a truncated or
  // hostile .bin can make the Cursor/DataView throw a raw RangeError. Wrap the
  // parse so malformed input surfaces as the library's ParseError.
  try {
    return await parseXlsbParts(zip, workbookPath, workbookDir, options?.dateSystem)
  } catch (err) {
    if (err instanceof ParseError || err instanceof EncryptedFileError || err instanceof ZipError) {
      throw err
    }
    throw new ParseError("Failed to parse XLSB workbook (malformed or truncated)", undefined, {
      cause: err,
    })
  }
}

async function parseXlsbParts(
  zip: ZipReader,
  workbookPath: string,
  workbookDir: string,
  dateSystem: ReadOptions["dateSystem"],
): Promise<Workbook> {
  // Sheets (name + relId) and the workbook's own date system from the .bin.
  const { sheets: sheetEntries, file1904 } = parseWorkbookBin(await zip.extract(workbookPath))

  // The file's flag wins unless the caller pinned a system, exactly as the
  // XLSX and XLS readers do — a Mac-authored 1904 workbook read with the
  // documented default ("auto") was otherwise 1462 days out on every date.
  const date1904 = !dateSystem || dateSystem === "auto" ? file1904 : dateSystem === "1904"

  // Workbook relationships (XML) → relId → part path.
  const wbRelsPath = workbookDir
    ? `${workbookDir}/_rels/${workbookPath.slice(workbookDir.length + 1)}.rels`
    : `_rels/${workbookPath}.rels`
  const relMap = new Map<string, string>()
  let sharedStringsPath: string | undefined
  let stylesPath: string | undefined
  if (zip.has(wbRelsPath)) {
    for (const rel of parseRelationships(decodeUtf8(await zip.extract(wbRelsPath), wbRelsPath))) {
      const target = resolvePath(workbookDir, rel.target)
      if (matchesRelType(rel.type, REL_WORKSHEET)) relMap.set(rel.id, target)
      else if (matchesRelType(rel.type, REL_SHARED_STRINGS)) sharedStringsPath = target
      else if (matchesRelType(rel.type, REL_STYLES)) stylesPath = target
    }
  }

  const sharedStrings =
    sharedStringsPath && zip.has(sharedStringsPath)
      ? parseSharedStringsBin(await zip.extract(sharedStringsPath))
      : []
  const dateXf =
    stylesPath && zip.has(stylesPath) ? parseStylesBin(await zip.extract(stylesPath)) : []

  const sheets: Sheet[] = []
  for (const entry of sheetEntries) {
    const path = entry.relId ? relMap.get(entry.relId) : undefined
    if (!path || !zip.has(path)) {
      sheets.push({ name: entry.name, rows: [] })
      continue
    }
    const { rows, merges } = parseWorksheetBin(
      await zip.extract(path),
      sharedStrings,
      dateXf,
      date1904,
    )
    const sheet: Sheet = { name: entry.name, rows }
    if (merges.length > 0) sheet.merges = merges
    sheets.push(sheet)
  }

  return { sheets }
}

// ── workbook.bin ─────────────────────────────────────────────────────

interface SheetEntry {
  name: string
  relId: string
}

function parseWorkbookBin(bin: Uint8Array): { sheets: SheetEntry[]; file1904: boolean } {
  const out: SheetEntry[] = []
  let file1904 = false
  for (const rec of iterateRecords(bin)) {
    if (rec.id === BrtWbProp) {
      // BrtWbProp (record 153, MS-XLSB §2.4): a 32-bit flag word, then
      // dwThemeVersion, then strName. Bit 0 of the flag word is f1904 —
      // the 1904 date system, XLSB's equivalent of workbookPr/@date1904 in
      // XLSX and the DATEMODE record in BIFF8. No record means 1900.
      file1904 = (new Cursor(rec.data).u32() & 0x01) !== 0
    } else if (rec.id === BrtBundleSh) {
      const c = new Cursor(rec.data)
      c.u32() // hsState
      c.u32() // iTabID
      const relId = c.nullableWideString()
      const name = c.wideString()
      out.push({ name, relId })
    }
  }
  return { sheets: out, file1904 }
}

// ── sharedStrings.bin ────────────────────────────────────────────────

function parseSharedStringsBin(bin: Uint8Array): string[] {
  const out: string[] = []
  for (const rec of iterateRecords(bin)) {
    if (rec.id !== BrtSSTItem) continue
    const c = new Cursor(rec.data)
    c.u8() // flags (rich/phonetic) — we only need the plain text
    out.push(c.wideString())
  }
  return out
}

// ── styles.bin (cell xf → is-date) ───────────────────────────────────

function parseStylesBin(bin: Uint8Array): boolean[] {
  const numFmtCodes = new Map<number, string>()
  const cellXfFmtIds: number[] = []
  let inCellXfs = false
  for (const rec of iterateRecords(bin)) {
    if (rec.id === BrtFmt) {
      const c = new Cursor(rec.data)
      const ifmt = c.u16()
      numFmtCodes.set(ifmt, c.wideString())
    } else if (rec.id === BrtBeginCellXFs) {
      inCellXfs = true
    } else if (rec.id === BrtEndCellXFs) {
      inCellXfs = false
    } else if (rec.id === BrtXF && inCellXfs) {
      const c = new Cursor(rec.data)
      c.u16() // ixfeParent
      cellXfFmtIds.push(c.u16()) // iFmt
    }
  }
  return cellXfFmtIds.map((id) => {
    if (isBuiltinDateFormatId(id)) return true
    const code = numFmtCodes.get(id)
    return code ? isDateFormat(code) : false
  })
}

// ── worksheet.bin ────────────────────────────────────────────────────

function parseWorksheetBin(
  bin: Uint8Array,
  sst: string[],
  dateXf: boolean[],
  date1904: boolean,
): { rows: CellValue[][]; merges: MergeRange[] } {
  const rows: CellValue[][] = []
  const merges: MergeRange[] = []
  let row = 0

  let widestCol = 0
  const setCell = (col: number, value: CellValue): void => {
    // Reject out-of-range column indices from untrusted u32 fields, which
    // would otherwise grow a row to billions of null slots and OOM.
    if (col < 0 || col > MAX_COL_INDEX) {
      throw new ParseError(
        `Cell column ${col} is outside the supported sheet bounds (max ${MAX_COL_INDEX + 1})`,
      )
    }
    if (col >= widestCol) widestCol = col + 1
    let r = rows[row]
    if (!r) r = rows[row] = []
    while (r.length < col) r.push(null)
    r[col] = value
  }
  const numericCell = (col: number, styleRef: number, num: number): void => {
    setCell(col, dateXf[styleRef] ? serialToDate(num, date1904) : num)
  }

  for (const rec of iterateRecords(bin)) {
    switch (rec.id) {
      case BrtRowHdr: {
        row = new Cursor(rec.data).u32()
        if (row < 0 || row > MAX_ROW_INDEX) {
          throw new ParseError(
            `Cell row ${row} is outside the supported sheet bounds (max ${MAX_ROW_INDEX + 1})`,
          )
        }
        break
      }
      case BrtCellBlank: {
        const c = new Cursor(rec.data)
        c.u32() // col — record present but no value
        break
      }
      case BrtCellRk: {
        const c = new Cursor(rec.data)
        const col = c.u32()
        const styleRef = c.u32() & 0xffffff
        numericCell(col, styleRef, decodeRk(c.u32()))
        break
      }
      case BrtCellReal:
      case BrtFmlaNum: {
        const c = new Cursor(rec.data)
        const col = c.u32()
        const styleRef = c.u32() & 0xffffff
        numericCell(col, styleRef, c.f64())
        break
      }
      case BrtCellBool:
      case BrtFmlaBool: {
        const c = new Cursor(rec.data)
        const col = c.u32()
        c.u32() // styleRef
        setCell(col, c.u8() !== 0)
        break
      }
      case BrtCellError:
      case BrtFmlaError: {
        const c = new Cursor(rec.data)
        const col = c.u32()
        c.u32() // styleRef
        setCell(col, ERROR_TEXT[c.u8()] ?? "#ERR!")
        break
      }
      case BrtCellSt:
      case BrtFmlaString: {
        const c = new Cursor(rec.data)
        const col = c.u32()
        c.u32() // styleRef
        setCell(col, c.wideString())
        break
      }
      case BrtCellIsst: {
        const c = new Cursor(rec.data)
        const col = c.u32()
        c.u32() // styleRef
        const idx = c.u32()
        setCell(col, sst[idx] ?? "")
        break
      }
      case BrtMergeCell: {
        // BrtMergeCell (record 176, MS-XLSB §2.4) carries one UncheckedRfX
        // (§2.5): rwFirst, rwLast, colFirst, colLast as four u32s — one
        // record per merged range, unlike BIFF8's MERGECELLS, which packs
        // a count and the whole list into a single record.
        const c = new Cursor(rec.data)
        merges.push({
          startRow: c.u32(),
          endRow: c.u32(),
          startCol: c.u32(),
          endCol: c.u32(),
        })
        break
      }
      default:
        break
    }
  }

  // Same shape as the XLS reader: rows are built only where cells landed,
  // so a sheet came back ragged and a row with no cell records came back
  // as an `undefined` hole. See #494.
  densify(rows, widestCol)

  return { rows, merges }
}
