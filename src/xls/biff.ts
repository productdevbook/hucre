// ── BIFF8 Record Stream ──────────────────────────────────────────────
// The "Workbook" stream inside an .xls OLE2 container is a flat sequence
// of BIFF records: 2-byte record id + 2-byte length + `length` bytes of
// data (each record body is ≤ 8224 bytes; longer data overflows into
// CONTINUE records). See [MS-XLS].

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
  RSTRING: 0x00d6,
  SST: 0x00fc,
  XF: 0x00e0,
  FORMAT: 0x041e,
  BOUNDSHEET: 0x0085,
  MERGECELLS: 0x00e5,
  BOF: 0x0809,
} as const

/** Parse a stream into records, recording each one's byte offset. */
export function parseRecords(stream: Uint8Array): BiffRecord[] {
  const out: BiffRecord[] = []
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength)
  let pos = 0
  while (pos + 4 <= stream.length) {
    const id = view.getUint16(pos, true)
    const len = view.getUint16(pos + 2, true)
    const start = pos
    pos += 4
    out.push({ id, data: stream.subarray(pos, pos + len), offset: start })
    pos += len
    // A zero id past real data means padding — stop.
    if (id === 0 && len === 0) break
  }
  return out
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
  skip(n: number): void {
    // skip may cross blocks (rich/phonetic trailers), no grbit byte
    let left = n
    while (left > 0) {
      this.ensure()
      const take = Math.min(left, this.remainingInBlock())
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
  const out: string[] = []
  for (let i = 0; i < cstUnique; i++) out.push(readSstString(s))
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
    str += String.fromCharCode(compressed ? s.u8() : s.u16())
    read++
  }
  if (rich) s.skip(cRun * 4)
  if (phonetic) s.skip(cbExt)
  return str
}
