// ── XLS (BIFF8) Reader ───────────────────────────────────────────────
// Read legacy Excel 97-2003 .xls files: an OLE2/CFB container whose
// "Workbook" stream is a BIFF8 record sequence. Reuses the CFB reader
// (shared with encryption) and decodes the records into the standard
// Workbook model. Read-only (MS-XLS).

import type { CellValue, MergeRange, ReadOptions, Sheet, Workbook } from "../_types"
import { ParseError } from "../errors"
import { readInputToUint8Array } from "../_input"
import { readCfb } from "../xlsx/crypto/cfb"
import { isDateFormat, serialToDate } from "../_date"
import { decodeRk, parseRecords, parseSst, Reader, SID, type BiffRecord } from "./biff"

const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47])

const ERROR_TEXT: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
}

/** Whether a CFB container holds a BIFF Workbook stream (.xls). */
export function looksLikeXls(streams: Map<string, Uint8Array>): boolean {
  return streams.has("Workbook") || streams.has("Book")
}

/** Read a BIFF8 .xls workbook into the standard {@link Workbook} model. */
export async function readXls(
  input: Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>,
  options?: ReadOptions,
): Promise<Workbook> {
  const data = await readInputToUint8Array(input)
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

  // Record parsing reads many length-prefixed binary fields; a truncated or
  // hostile file can make DataView accessors throw a raw RangeError. Wrap the
  // whole pass so malformed input surfaces as the library's ParseError.
  try {
    return parseWorkbookRecords(stream, options)
  } catch (err) {
    if (err instanceof ParseError) throw err
    throw new ParseError("Failed to parse XLS workbook (malformed or truncated)", undefined, {
      cause: err,
    })
  }
}

function parseWorkbookRecords(stream: Uint8Array, options?: ReadOptions): Workbook {
  const records = parseRecords(stream)
  const offsetToIndex = new Map<number, number>()
  for (let i = 0; i < records.length; i++) offsetToIndex.set(records[i].offset, i)

  // ── Globals substream (records[0] = BOF … first EOF) ──
  let date1904 = options?.dateSystem === "1904"
  const xfFmtIds: number[] = []
  const fmtCodes = new Map<number, string>()
  const boundSheets: Array<{ name: string; pos: number }> = []
  const sst: string[] = []

  let gi = 0
  for (; gi < records.length; gi++) {
    const rec = records[gi]
    if (rec.id === SID.EOF) {
      gi++
      break
    }
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
        fmtCodes.set(ifmt, readXLString(r))
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
        boundSheets.push({ name: readShortString(r), pos })
        break
      }
      case SID.SST: {
        const blocks: Uint8Array[] = [rec.data]
        // Gather trailing CONTINUE records belonging to the SST.
        for (let j = gi + 1; j < records.length; j++) {
          if (records[j].id !== SID.CONTINUE) break
          blocks.push(records[j].data)
        }
        sst.push(...parseSst(blocks))
        break
      }
      default:
        break
    }
  }

  const dateXf = xfFmtIds.map((id) => {
    if (BUILTIN_DATE_IDS.has(id)) return true
    const code = fmtCodes.get(id)
    return code ? isDateFormat(code) : false
  })

  const isDate = (ixfe: number): boolean => dateXf[ixfe] === true

  // ── Sheet substreams ──
  const sheets: Sheet[] = []
  for (const bs of boundSheets) {
    const startIdx = offsetToIndex.get(bs.pos)
    if (startIdx === undefined) {
      sheets.push({ name: bs.name, rows: [] })
      continue
    }
    sheets.push(parseSheet(records, startIdx, bs.name, sst, isDate, date1904))
  }

  return { sheets }
}

function parseSheet(
  records: BiffRecord[],
  startIdx: number,
  name: string,
  sst: string[],
  isDate: (ixfe: number) => boolean,
  date1904: boolean,
): Sheet {
  const rows: CellValue[][] = []
  const merges: MergeRange[] = []

  const setCell = (row: number, col: number, value: CellValue): void => {
    let r = rows[row]
    if (!r) r = rows[row] = []
    while (r.length < col) r.push(null)
    r[col] = value
  }
  const numeric = (row: number, col: number, ixfe: number, n: number): void => {
    setCell(row, col, isDate(ixfe) ? serialToDate(n, date1904) : n)
  }

  for (let i = startIdx + 1; i < records.length; i++) {
    const rec = records[i]
    if (rec.id === SID.EOF) break
    const r = new Reader(rec.data)
    switch (rec.id) {
      case SID.LABELSST: {
        const row = r.u16(),
          col = r.u16()
        r.u16() // ixfe
        setCell(row, col, sst[r.u32()] ?? "")
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
        setCell(row, col, isError ? (ERROR_TEXT[val] ?? "#ERR!") : val !== 0)
        break
      }
      case SID.LABEL: {
        const row = r.u16(),
          col = r.u16()
        r.u16() // ixfe
        setCell(row, col, readXLString(r))
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
            setCell(row, col, b[2] !== 0) // boolean
          else if (kind === 2)
            setCell(row, col, ERROR_TEXT[b[2]] ?? "#ERR!") // error
          else if (kind === 0) {
            // string: value is in the following STRING record
            const next = records[i + 1]
            if (next && next.id === SID.STRING)
              setCell(row, col, readXLString(new Reader(next.data)))
          }
          // kind === 3 → blank/empty
        } else {
          const num = new DataView(b.buffer, b.byteOffset, 8).getFloat64(0, true)
          numeric(row, col, ixfe, num)
        }
        break
      }
      case SID.MERGECELLS: {
        const cmcs = r.u16()
        for (let k = 0; k < cmcs; k++) {
          const rwFirst = r.u16(),
            rwLast = r.u16(),
            colFirst = r.u16(),
            colLast = r.u16()
          merges.push({ startRow: rwFirst, endRow: rwLast, startCol: colFirst, endCol: colLast })
        }
        break
      }
      default:
        break
    }
  }

  const sheet: Sheet = { name, rows }
  if (merges.length > 0) sheet.merges = merges
  return sheet
}

// ── String helpers ───────────────────────────────────────────────────

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
  let s = ""
  for (let i = 0; i < cch; i++) s += String.fromCharCode(compressed ? r.u8() : r.u16())
  return s
}
