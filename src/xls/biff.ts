// ── BIFF8 Record Stream ──────────────────────────────────────────────
// The "Workbook" stream inside an .xls OLE2 container is a flat sequence
// of BIFF records: 2-byte record id + 2-byte length + `length` bytes of
// data (each record body is ≤ 8224 bytes; longer data overflows into
// CONTINUE records). See [MS-XLS].

import { ParseError } from "../errors"

export interface BiffRecord {
  /** Record id (sid). */
  id: number
  data: Uint8Array
  /** Byte offset of this record's header within the stream. */
  offset: number
}

export const SID = {
  FORMULA: 0x0006,
  EOF: 0x000a,
  CONTINUE: 0x003c,
  DATEMODE: 0x0022,
  CODEPAGE: 0x0042,
  BLANK: 0x0201,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  STRING: 0x0207,
  ROW: 0x0208,
  INDEX: 0x020b,
  RK: 0x027e,
  MULRK: 0x00bd,
  MULBLANK: 0x00be,
  LABELSST: 0x00fd,
  // RSTRING is how BIFF5/BIFF7 writes a rich-text cell: rw, col, ixfe, a
  // codepage byte string, then a run count and its runs. BIFF8 dropped it
  // (rich text moved into the SST, where readSstString skips the runs), so
  // the reader meets it only on the BIFF5 path — and takes the text alone,
  // as it does everywhere else. See #411.
  RSTRING: 0x00d6,
  SST: 0x00fc,
  XF: 0x00e0,
  FORMAT: 0x041e,
  BOUNDSHEET: 0x0085,
  MERGECELLS: 0x00e5,
  BOF: 0x0809,
} as const

/**
 * Walks a stream's records in place, one at a time.
 *
 * The reader used to materialize every record before decoding the first:
 * an object and a view apiece, ~155 bytes each, and a map from offset to
 * index on top. A 17 MB workbook is ~960k records, so that was ~150 MB
 * held for the whole read. Framing is unchanged — a record running past
 * the end of the stream is cut short, and a zero id with a zero length is
 * padding that ends it.
 */
export class RecordCursor {
  private view: DataView
  private ended = false

  constructor(
    private stream: Uint8Array,
    private pos = 0,
  ) {
    this.view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength)
  }

  /** Move past the record at the cursor and return its offset, or -1 at the end. */
  skip(): number {
    if (this.ended || this.pos + 4 > this.stream.length) return -1
    const offset = this.pos
    const id = this.view.getUint16(offset, true)
    const len = this.view.getUint16(offset + 2, true)
    this.pos = offset + 4 + len
    // A zero id past real data means padding — stop.
    if (id === 0 && len === 0) this.ended = true
    return offset
  }

  next(): BiffRecord | undefined {
    const offset = this.skip()
    if (offset < 0) return undefined
    const len = this.view.getUint16(offset + 2, true)
    return {
      id: this.view.getUint16(offset, true),
      data: this.stream.subarray(offset + 4, offset + 4 + len),
      offset,
    }
  }

  /** The record `next` would return, without moving past it. */
  peek(): BiffRecord | undefined {
    const { pos, ended } = this
    const rec = this.next()
    this.pos = pos
    this.ended = ended
    return rec
  }
}

/**
 * Which of `offsets` begin a record. A BOUNDSHEET names its sheet by the
 * byte offset of the sheet's BOF, and one landing inside a record — a
 * damaged file — has always read as an empty sheet rather than as records
 * taken out of phase.
 */
export function recordStarts(stream: Uint8Array, offsets: Iterable<number>): Set<number> {
  const wanted = new Set(offsets)
  const found = new Set<number>()
  const cursor = new RecordCursor(stream)
  for (let at = cursor.skip(); at >= 0 && found.size < wanted.size; at = cursor.skip()) {
    if (wanted.has(at)) found.add(at)
  }
  return found
}

/** Little-endian cursor over a record body. */
export class Reader {
  pos = 0
  private view: DataView

  constructor(public buf: Uint8Array) {
    this.view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  }

  u8(): number {
    return this.buf[this.pos++]
  }
  u16(): number {
    const v = this.view.getUint16(this.pos, true)
    this.pos += 2
    return v
  }
  i16(): number {
    const v = this.view.getInt16(this.pos, true)
    this.pos += 2
    return v
  }
  u32(): number {
    const v = this.view.getUint32(this.pos, true)
    this.pos += 4
    return v
  }
  f64(): number {
    const v = this.view.getFloat64(this.pos, true)
    this.pos += 8
    return v
  }
  skip(n: number): void {
    this.pos += n
  }
  remaining(): number {
    return this.buf.length - this.pos
  }
}

/** The most characters handed to one `String.fromCharCode` call. */
const CHARS_PER_CALL = 8192

/**
 * `count` characters from `start`: a byte each when compressed (the byte
 * is the code point), else UTF-16LE code units. The caller guarantees the
 * bytes are there.
 *
 * Built a character at a time with `+=`, a string is a rope, one node per
 * character until something flattens it, and nothing in a read did. A
 * workbook's 2.3 million shared characters held ~70 MB that way, against
 * ~7 MB decoded flat. TextDecoder would be wrong twice over: the Encoding
 * Standard maps "latin1" to windows-1252, which moves 0x80–0x9F, and
 * "utf-16le" turns a lone surrogate into U+FFFD — which is what each half
 * of a pair becomes when a CONTINUE boundary splits it.
 */
export function decodeChars(
  buf: Uint8Array,
  start: number,
  count: number,
  compressed: boolean,
): string {
  const units = compressed ? buf.subarray(start, start + count) : utf16Units(buf, start, count)
  if (units.length <= CHARS_PER_CALL) return String.fromCharCode(...units)
  const parts: string[] = []
  for (let i = 0; i < units.length; i += CHARS_PER_CALL) {
    parts.push(String.fromCharCode(...units.subarray(i, i + CHARS_PER_CALL)))
  }
  return parts.join("")
}

function utf16Units(buf: Uint8Array, start: number, count: number): Uint16Array {
  const units = new Uint16Array(count)
  for (let i = 0, p = start; i < count; i++, p += 2) units[i] = buf[p] | (buf[p + 1] << 8)
  return units
}

/** Decode an RK number (MS-XLS §2.5.166). */
export function decodeRk(rk: number): number {
  const fX100 = (rk & 1) !== 0
  const fInt = (rk & 2) !== 0
  let value: number
  if (fInt) {
    value = (rk | 0) >> 2
  } else {
    const buf = new ArrayBuffer(8)
    const dv = new DataView(buf)
    dv.setUint32(4, rk & 0xfffffffc, true)
    value = dv.getFloat64(0, true)
  }
  return fX100 ? value / 100 : value
}

// ── SST (shared string table) with CONTINUE handling ─────────────────
// Strings split across CONTINUE records on character boundaries; each
// CONTINUE that resumes a string's character array restarts with a 1-byte
// option flag (fHighByte). Header fields and rich/phonetic trailers are
// assumed not to straddle a boundary (how Excel writes them).

/** Reads across the SST record + its trailing CONTINUE blocks. */
class BlockStream {
  private bi = 0
  private pos = 0
  constructor(private blocks: Uint8Array[]) {}

  private cur(): Uint8Array {
    return this.blocks[this.bi]
  }
  remainingInBlock(): number {
    const b = this.cur()
    return b ? b.length - this.pos : 0
  }
  atEnd(): boolean {
    return this.bi >= this.blocks.length
  }
  /** Total bytes across all blocks — an upper bound on decodable records. */
  totalBytes(): number {
    let n = 0
    for (const b of this.blocks) n += b.length
    return n
  }
  /** Move to the next block (used when a string's chars continue). */
  nextBlock(): void {
    this.bi++
    this.pos = 0
  }
  ensure(): void {
    while (!this.atEnd() && this.remainingInBlock() === 0) this.nextBlock()
  }
  u8(): number {
    return this.cur()[this.pos++]
  }
  u16(): number {
    const v = this.cur()[this.pos] | (this.cur()[this.pos + 1] << 8)
    this.pos += 2
    return v
  }
  u32(): number {
    const b = this.cur()
    const v =
      (b[this.pos] | (b[this.pos + 1] << 8) | (b[this.pos + 2] << 16) | (b[this.pos + 3] << 24)) >>>
      0
    this.pos += 4
    return v
  }
  /** `count` characters from the current block, which holds them all. */
  chars(count: number, compressed: boolean): string {
    const s = decodeChars(this.cur(), this.pos, count, compressed)
    this.pos += compressed ? count : count * 2
    return s
  }
  skip(n: number): void {
    // skip may cross blocks (rich/phonetic trailers), no grbit byte
    let left = n
    while (left > 0) {
      this.ensure()
      const take = Math.min(left, this.remainingInBlock())
      // Once the blocks are exhausted, remainingInBlock() returns 0 and
      // `left` would never decrease — a live lock no try/catch can
      // interrupt. The callers feed this untrusted lengths (`cRun * 4`
      // from a u16, `cbExt` from a u32), so a file claiming more
      // trailer bytes than it carries used to hang the process. See #389.
      if (take === 0) {
        throw new ParseError(
          `Invalid XLS: string record claims ${n} trailing bytes but the stream ends after ${n - left}`,
        )
      }
      this.pos += take
      left -= take
    }
  }
}

/**
 * Parse an SST record (plus its following CONTINUE records) into the
 * shared-string array. `blocks` is `[sstData, ...continueDatas]`.
 */
export function parseSst(blocks: Uint8Array[]): string[] {
  const s = new BlockStream(blocks)
  s.skip(4) // cstTotal
  const cstUnique = s.u32()
  // cstUnique is an untrusted u32. Each string costs at least 3 bytes
  // (2-byte cch + 1-byte grbit), so a string count larger than the total
  // available bytes is a corrupt / hostile header — cap the loop by the
  // bytes actually present so we don't spin allocating empty strings.
  const maxStrings = Math.min(cstUnique, s.totalBytes())
  const out: string[] = []
  for (let i = 0; i < maxStrings; i++) {
    s.ensure()
    if (s.atEnd()) break
    out.push(readSstString(s))
  }
  return out
}

function readSstString(s: BlockStream): string {
  const cch = s.u16()
  let grbit = s.u8()
  let compressed = (grbit & 0x01) === 0
  const rich = (grbit & 0x08) !== 0
  const phonetic = (grbit & 0x04) !== 0
  const cRun = rich ? s.u16() : 0
  const cbExt = phonetic ? s.u32() : 0

  // A block's worth of characters at a time — see decodeChars.
  let str = ""
  let read = 0
  while (read < cch) {
    if (s.remainingInBlock() === 0) {
      // Continue into the next block: it restarts with an option flag for
      // the remaining characters.
      s.nextBlock()
      grbit = s.u8()
      compressed = (grbit & 0x01) === 0
    }
    const fit = compressed ? s.remainingInBlock() : s.remainingInBlock() >> 1
    if (fit > 0) {
      const take = Math.min(cch - read, fit)
      str += s.chars(take, compressed)
      read += take
    } else {
      // A two-byte character straddling the end of a block is malformed;
      // it reads exactly as it always has.
      str += String.fromCharCode(compressed ? s.u8() : s.u16())
      read++
    }
  }
  if (rich) s.skip(cRun * 4)
  if (phonetic) s.skip(cbExt)
  return str
}
