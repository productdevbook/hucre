// ── XLS (BIFF5 / BIFF8) Reader ───────────────────────────────────────
// Read legacy Excel .xls files: an OLE2/CFB container whose "Workbook"
// (BIFF8, Excel 97-2003) or "Book" (BIFF5/7, Excel 5.0/95) stream is a
// BIFF record sequence. Reuses the CFB reader (shared with encryption)
// and decodes the records into the standard Workbook model. Read-only
// (MS-XLS).

import type { CellValue, MergeRange, ReadOptions, Sheet, Workbook } from "../_types"
import { ParseError } from "../errors"
import { MAX_COL_INDEX, MAX_ROW_INDEX, MAX_TOTAL_CELLS } from "../limits"
import { readInputToUint8Array } from "../_input"
import { readCfb } from "../xlsx/crypto/cfb"
import { isBuiltinDateFormatId, isDateFormat, serialToDate } from "../_date"
import {
  decodeChars,
  decodeRk,
  parseSst,
  Reader,
  RecordCursor,
  recordStarts,
  SID,
  type BiffRecord,
} from "./biff"
import { codepageDecoder } from "./codepage"

const ERROR_TEXT: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
}

const BIFF5 = 0x0500
const BIFF8 = 0x0600

/** The code page assumed for a BIFF5 workbook that names none. */
const DEFAULT_CODEPAGE = 1252

/** Whether a CFB container holds a BIFF Workbook stream (.xls). */
export function looksLikeXls(streams: Map<string, Uint8Array>): boolean {
  return streams.has("Workbook") || streams.has("Book")
}

/** Read a BIFF5 or BIFF8 .xls workbook into the standard {@link Workbook} model. */
export async function readXls(
  input: Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>,
  options?: ReadOptions,
): Promise<Workbook> {
  const stream = await openWorkbookStream(input, options)
  try {
    return parseWorkbookRecords(stream, options)
  } catch (err) {
    throw asParseError(err)
  }
}

/** The container's Workbook (BIFF8) or Book (BIFF5) stream, read whole. */
export async function openWorkbookStream(
  input: Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>,
  options?: ReadOptions,
): Promise<Uint8Array> {
  const data = await readInputToUint8Array(input, options?.maxInputBytes)
  let streams: Map<string, Uint8Array>
  try {
    streams = readCfb(data)
  } catch (err) {
    throw new ParseError("Failed to open XLS: not a valid OLE2 container", undefined, {
      cause: err,
    })
  }
  const stream = streams.get("Workbook") ?? streams.get("Book")
  if (!stream) throw new ParseError("Invalid XLS: missing Workbook stream")
  return stream
}

/**
 * Record parsing reads many length-prefixed binary fields; a truncated or
 * hostile file can make DataView accessors throw a raw RangeError. Each
 * reader wraps its whole pass in this, so malformed input surfaces as the
 * library's ParseError.
 */
export function asParseError(err: unknown): ParseError {
  if (err instanceof ParseError) return err
  return new ParseError("Failed to parse XLS workbook (malformed or truncated)", undefined, {
    cause: err,
  })
}

/**
 * How the two BIFF generations spell a string. BIFF8 writes Unicode with a
 * per-string flag byte (compressed = Latin-1 bytes, else UTF-16LE); BIFF5/7
 * writes bytes in the workbook's one code page and no flag at all. The
 * length prefix is a u16 on cell text and a u8 on names and format codes
 * in both.
 */
export interface BiffStrings {
  long(r: Reader): string
  short(r: Reader): string
}

const BIFF8_STRINGS: BiffStrings = {
  long: readXLString,
  short: readShortString,
}

function biff5Strings(decode: (bytes: Uint8Array) => string): BiffStrings {
  const bytes = (r: Reader, cch: number): string => {
    // `subarray` clamps, so an over-long count would hand back a silently
    // shortened string instead of the error a truncated record deserves.
    if (cch > r.remaining()) {
      throw new ParseError(
        `Invalid XLS: string claims ${cch} bytes but its record has ${r.remaining()} left`,
      )
    }
    const s = decode(r.buf.subarray(r.pos, r.pos + cch))
    r.skip(cch)
    return s
  }
  return {
    long: (r) => bytes(r, r.u16()),
    short: (r) => bytes(r, r.u8()),
  }
}

/**
 * The globals' CODEPAGE record, if the workbook carries one. Scanned ahead
 * of the globals pass because the strings that pass decodes — sheet names,
 * format codes — may in principle precede it.
 */
function findCodepage(stream: Uint8Array): number | undefined {
  const cursor = new RecordCursor(stream)
  for (let rec = cursor.next(); rec; rec = cursor.next()) {
    if (rec.id === SID.EOF) return undefined
    if (rec.id === SID.CODEPAGE) {
      if (rec.data.length < 2) throw new ParseError("Invalid XLS: CODEPAGE record is too short")
      return new Reader(rec.data).u16()
    }
  }
  return undefined
}

/** What the globals substream settles for every sheet: how to read text and dates, and where the sheets are. */
export interface XlsGlobals {
  strings: BiffStrings
  sst: string[]
  isDate: (ixfe: number) => boolean
  date1904: boolean
  boundSheets: Array<{ name: string; pos: number }>
}

function parseWorkbookRecords(stream: Uint8Array, options?: ReadOptions): Workbook {
  const globals = readGlobals(stream, options)

  // ── Sheet substreams ──
  const starts = recordStarts(
    stream,
    globals.boundSheets.map((bs) => bs.pos),
  )
  const sheets: Sheet[] = []
  for (const bs of globals.boundSheets) {
    if (!starts.has(bs.pos)) {
      sheets.push({ name: bs.name, rows: [] })
      continue
    }
    sheets.push(
      parseSheet(
        new RecordCursor(stream, bs.pos),
        bs.name,
        globals,
        options?.maxTotalCells ?? MAX_TOTAL_CELLS,
      ),
    )
  }

  return { sheets }
}

/** The globals substream (BOF … first EOF), read once for any reader. */
export function readGlobals(stream: Uint8Array, options?: ReadOptions): XlsGlobals {
  const cursor = new RecordCursor(stream)

  // ── BIFF version gate ──
  // The first record is the workbook globals BOF; its first u16 is the BIFF
  // version. BIFF8 (0x0600) is Excel 97-2003; BIFF5 and BIFF7 (both 0x0500)
  // are Excel 5.0 and 95 — the same record layout except that strings are
  // code-page bytes and there is no shared-string table. Anything else is
  // rejected rather than misread.
  const bof = cursor.next()
  if (!bof || bof.id !== SID.BOF) {
    throw new ParseError("Invalid XLS: missing BOF record at start of Workbook stream")
  }
  // A BOF too short to carry a version is read as BIFF8, as it always was.
  const biffVersion = bof.data.length >= 2 ? new Reader(bof.data).u16() : BIFF8
  if (biffVersion !== BIFF8 && biffVersion !== BIFF5) {
    throw new ParseError(
      `Unsupported XLS version (BIFF 0x${biffVersion.toString(16)}). ` +
        "Only BIFF5 (Excel 5.0/95) and BIFF8 (Excel 97-2003) are supported; re-save the file as .xlsx.",
    )
  }

  // A BIFF5 workbook's strings are bytes in one code page: the file's own
  // CODEPAGE record names it, the caller's `codepage` stands in for a file
  // that has none, and Windows-1252 is the last resort. Resolved before
  // any string is read, so a bad page fails here rather than mid-sheet.
  const strings: BiffStrings =
    biffVersion === BIFF5
      ? biff5Strings(codepageDecoder(findCodepage(stream) ?? options?.codepage ?? DEFAULT_CODEPAGE))
      : BIFF8_STRINGS

  // ── Globals substream (BOF … first EOF) ──
  let date1904 = options?.dateSystem === "1904"
  const xfFmtIds: number[] = []
  const fmtCodes = new Map<number, string>()
  const boundSheets: Array<{ name: string; pos: number }> = []
  const sst: string[] = []

  for (let rec: BiffRecord | undefined = bof; rec; rec = cursor.next()) {
    if (rec.id === SID.EOF) break
    switch (rec.id) {
      case SID.DATEMODE: {
        if (!options?.dateSystem || options.dateSystem === "auto") {
          date1904 = new Reader(rec.data).u16() === 1
        }
        break
      }
      case SID.FORMAT: {
        const r = new Reader(rec.data)
        const ifmt = r.u16()
        // BIFF8 spells the code as an XLUnicodeString (u16 count); BIFF5
        // as a u8-counted byte string.
        fmtCodes.set(ifmt, biffVersion === BIFF5 ? strings.short(r) : strings.long(r))
        break
      }
      case SID.XF: {
        const r = new Reader(rec.data)
        r.u16() // ifnt
        xfFmtIds.push(r.u16()) // ifmt
        break
      }
      case SID.BOUNDSHEET: {
        const r = new Reader(rec.data)
        const pos = r.u32()
        r.u8() // hsState (visibility)
        r.u8() // dt (sheet type)
        boundSheets.push({ name: strings.short(r), pos })
        break
      }
      case SID.SST: {
        const blocks: Uint8Array[] = [rec.data]
        // Gather trailing CONTINUE records belonging to the SST.
        for (let next = cursor.peek(); next?.id === SID.CONTINUE; next = cursor.peek()) {
          blocks.push(next.data)
          cursor.next()
        }
        // A loop, not `push(...parseSst(blocks))`: spreading an array as
        // arguments puts one stack slot per element, so a workbook with a
        // few hundred thousand shared strings — an ordinary large .xls —
        // threw `RangeError: Maximum call stack size exceeded`, which the
        // caller reported as "malformed or truncated".
        for (const s of parseSst(blocks)) sst.push(s)
        break
      }
      default:
        break
    }
  }

  // The workbook's own FORMAT records win over the built-in id table: a
  // file may redefine a built-in id (ECMA-376 §18.8.30), and 1C allocates
  // its custom formats from id 50 rather than 164 — landing on the
  // Thai/Chinese/Korean date block. Asking the built-in table first meant
  // the FORMAT record was never consulted, so `'000000000000'`, the mask
  // for a leading-zero barcode, read back as `Invalid Date` and the number
  // was gone. `isDateStyle` in xlsx/styles.ts has always had this order;
  // the two BIFF readers disagreed with it. See #568.
  const dateXf = xfFmtIds.map((id) => {
    const code = fmtCodes.get(id)
    if (code !== undefined) return isDateFormat(code)
    return isBuiltinDateFormatId(id)
  })

  // Under `dates: "serial"` no format counts as a date, so every number
  // reads as the serial the file holds — a Date rounds to the
  // millisecond and puts serial 60 on the same day as 59. See #600.
  const isDate =
    options?.dates === "serial" ? () => false : (ixfe: number): boolean => dateXf[ixfe] === true

  return { strings, sst, isDate, date1904, boundSheets }
}

function parseSheet(
  cursor: RecordCursor,
  name: string,
  globals: XlsGlobals,
  cellLimit: number,
): Sheet {
  const rows: CellValue[][] = []
  const merges: MergeRange[] = []

  // BIFF row/col are u16, so each is bounded at 65,535 on its own — but
  // their product is not, and `rows` is a dense rectangle. 65,535 rows of
  // 65,536 slots is 4.3e9 allocations from a few hundred KB of input.
  // The XLSB reader already guards its coordinates this way; this one did
  // not. See #363.
  let widestCol = 0
  const setCell = (row: number, col: number, value: CellValue): void => {
    assertInSheetBounds(row, col)
    if (col >= widestCol) widestCol = col + 1
    const boundingBox = Math.max(rows.length, row + 1) * widestCol
    if (boundingBox > cellLimit) {
      throw new ParseError(
        `Sheet spans ${boundingBox} cells, over the ${cellLimit} limit. ` +
          "Raise `maxTotalCells` if the sheet really is this large.",
      )
    }
    let r = rows[row]
    if (!r) r = rows[row] = []
    placeCell(r, col, value)
  }

  cursor.skip() // the sheet's BOF
  for (let rec = cursor.next(); rec; rec = cursor.next()) {
    if (rec.id === SID.EOF) break
    if (rec.id !== SID.MERGECELLS) {
      decodeCells(rec, cursor, globals, setCell)
      continue
    }
    const r = new Reader(rec.data)
    const cmcs = r.u16()
    for (let k = 0; k < cmcs; k++) {
      const rwFirst = r.u16(),
        rwLast = r.u16(),
        colFirst = r.u16(),
        colLast = r.u16()
      merges.push({ startRow: rwFirst, endRow: rwLast, startCol: colFirst, endCol: colLast })
    }
  }

  // `rows` is a dense rectangle — the bounding-box guard above is sized
  // on that assumption and `CellValue` has no `undefined` member — but
  // `setCell` only pads a row up to its *own* last written column, and
  // never allocates a row with no cell records at all. So a sheet came
  // back ragged, and a row Excel left empty came back as a hole rather
  // than a row. `readXlsx` normalizes at the end of its parse; this and
  // the XLSB reader did not. See #494.
  densify(rows, widestCol)

  const sheet: Sheet = { name, rows }
  if (merges.length > 0) sheet.merges = merges
  return sheet
}

/** BIFF rows and columns are u16s; the model is narrower, and a cell past it is refused, not dropped. */
export function assertInSheetBounds(row: number, col: number): void {
  if (row < 0 || row > MAX_ROW_INDEX) {
    throw new ParseError(
      `Cell row ${row} is outside the supported sheet bounds (max ${MAX_ROW_INDEX + 1})`,
    )
  }
  if (col < 0 || col > MAX_COL_INDEX) {
    throw new ParseError(
      `Cell column ${col} is outside the supported sheet bounds (max ${MAX_COL_INDEX + 1})`,
    )
  }
}

/** Write `value` at `col`, padding the row with nulls up to it. */
export function placeCell(row: CellValue[], col: number, value: CellValue): void {
  while (row.length < col) row.push(null)
  row[col] = value
}

/**
 * The cells one record carries, each handed to `set` in the record's own
 * order — every reader decodes values here, so they cannot drift apart.
 * Records that carry no value (BLANK, MERGECELLS, formatting) pass through.
 * A string formula's cached value is in the STRING record after it, which
 * `cursor` still points at.
 */
export function decodeCells(
  rec: BiffRecord,
  cursor: RecordCursor,
  globals: XlsGlobals,
  set: (row: number, col: number, value: CellValue) => void,
): void {
  const { sst, strings, isDate, date1904 } = globals
  const numeric = (row: number, col: number, ixfe: number, n: number): void => {
    set(row, col, isDate(ixfe) ? serialToDate(n, date1904) : n)
  }
  const r = new Reader(rec.data)
  switch (rec.id) {
    case SID.LABELSST: {
      const row = r.u16(),
        col = r.u16()
      r.u16() // ixfe
      set(row, col, sst[r.u32()] ?? "")
      break
    }
    case SID.RK: {
      const row = r.u16(),
        col = r.u16(),
        ixfe = r.u16()
      numeric(row, col, ixfe, decodeRk(r.u32()))
      break
    }
    case SID.NUMBER: {
      const row = r.u16(),
        col = r.u16(),
        ixfe = r.u16()
      numeric(row, col, ixfe, r.f64())
      break
    }
    case SID.MULRK: {
      const row = r.u16()
      const colFirst = r.u16()
      const count = (rec.data.length - 6) / 6
      for (let k = 0; k < count; k++) {
        const ixfe = r.u16()
        numeric(row, colFirst + k, ixfe, decodeRk(r.u32()))
      }
      break
    }
    case SID.BOOLERR: {
      const row = r.u16(),
        col = r.u16()
      r.u16() // ixfe
      const val = r.u8()
      const isError = r.u8() === 1
      set(row, col, isError ? (ERROR_TEXT[val] ?? "#ERR!") : val !== 0)
      break
    }
    case SID.LABEL:
    case SID.RSTRING: {
      // Same head; RSTRING (BIFF5 rich text) trails formatting runs the
      // value does not carry.
      const row = r.u16(),
        col = r.u16()
      r.u16() // ixfe
      const text = strings.long(r)
      if (rec.id === SID.RSTRING) skipBiff5Runs(r)
      set(row, col, text)
      break
    }
    case SID.FORMULA: {
      const row = r.u16(),
        col = r.u16(),
        ixfe = r.u16()
      const b = rec.data.subarray(r.pos, r.pos + 8)
      if (b[6] === 0xff && b[7] === 0xff) {
        const kind = b[0]
        if (kind === 1)
          set(row, col, b[2] !== 0) // boolean
        else if (kind === 2)
          set(row, col, ERROR_TEXT[b[2]] ?? "#ERR!") // error
        else if (kind === 0) {
          // string: value is in the following STRING record
          const next = cursor.peek()
          if (next && next.id === SID.STRING) set(row, col, strings.long(new Reader(next.data)))
        }
        // kind === 3 → blank/empty
      } else {
        const num = new DataView(b.buffer, b.byteOffset, 8).getFloat64(0, true)
        numeric(row, col, ixfe, num)
      }
      break
    }
    default:
      break
  }
}

/**
 * The BIFF5 RSTRING trailer: a u8 run count, then two bytes per run (first
 * character, font). Narrower than BIFF8's FormatRun (u16 + u16), which is
 * what the SST path skips. Read for the same reason readSstString skips
 * its runs: a record that stops short of what it announces is malformed.
 */
function skipBiff5Runs(r: Reader): void {
  if (r.remaining() < 1) throw new ParseError("Invalid XLS: RSTRING record has no run count")
  const runs = r.u8()
  if (r.remaining() < runs * 2) {
    throw new ParseError(
      `Invalid XLS: RSTRING claims ${runs} formatting runs but its record has ${r.remaining()} bytes left`,
    )
  }
  r.skip(runs * 2)
}

// ── BIFF8 string helpers ─────────────────────────────────────────────

/** XLUnicodeString: u16 char count + 1 grbit byte + chars. */
function readXLString(r: Reader): string {
  const cch = r.u16()
  return readChars(r, cch)
}

/** ShortXLUnicodeString: u8 char count + 1 grbit byte + chars. */
function readShortString(r: Reader): string {
  const cch = r.u8()
  return readChars(r, cch)
}

function readChars(r: Reader, cch: number): string {
  const grbit = r.u8()
  const compressed = (grbit & 0x01) === 0
  const bytes = compressed ? cch : cch * 2
  if (bytes <= r.remaining()) {
    const s = decodeChars(r.buf, r.pos, cch, compressed)
    r.skip(bytes)
    return s
  }
  // A string running past its record — truncated, or a long formula result
  // continued into a CONTINUE record, which is not followed here — reads
  // as it always has, a character at a time.
  let s = ""
  for (let i = 0; i < cch; i++) s += String.fromCharCode(compressed ? r.u8() : r.u16())
  return s
}

/**
 * Fill a sparsely-built row array out to a rectangle.
 *
 * Two separate holes, both from building rows only where cells landed: a
 * row index never touched is `undefined` — which `CellValue` cannot
 * express — and a row that ended early is shorter than the sheet. See
 * #494.
 */
export function densify(rows: CellValue[][], width: number): void {
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] ?? (rows[r] = [])
    while (row.length < width) row.push(null)
  }
}
