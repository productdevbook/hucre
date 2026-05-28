// ── XLSB Binary Record Stream ────────────────────────────────────────
// XLSB `.bin` parts are a flat sequence of records:
//   record id (1–2 bytes, 7-bit varint) + size (1–4 bytes, 7-bit varint)
//   + `size` bytes of payload.
// See [MS-XLSB] §2.1.4.

export interface XlsbRecord {
  id: number
  data: Uint8Array
}

/** Iterate every record in a `.bin` part, yielding its id and raw payload. */
export function* iterateRecords(bin: Uint8Array): Generator<XlsbRecord> {
  let pos = 0
  const len = bin.length
  while (pos < len) {
    let id = bin[pos++]
    if (id & 0x80) {
      id = (id & 0x7f) | ((bin[pos++] & 0x7f) << 7)
    }
    let size = 0
    for (let k = 0; k < 4; k++) {
      const b = bin[pos++]
      size |= (b & 0x7f) << (7 * k)
      if ((b & 0x80) === 0) break
    }
    yield { id, data: bin.subarray(pos, pos + size) }
    pos += size
  }
}

/** Little-endian cursor over a record payload. */
export class Cursor {
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

  u32(): number {
    const v = this.view.getUint32(this.pos, true)
    this.pos += 4
    return v
  }

  i32(): number {
    const v = this.view.getInt32(this.pos, true)
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

  /** XLWideString: UInt32 char count + UTF-16LE units. */
  wideString(): string {
    const cch = this.u32()
    return this.readUtf16(cch)
  }

  /** XLNullableWideString: like {@link wideString} but 0xFFFFFFFF means absent. */
  nullableWideString(): string {
    const cch = this.u32()
    if (cch === 0xffffffff) return ""
    return this.readUtf16(cch)
  }

  private readUtf16(cch: number): string {
    let s = ""
    for (let i = 0; i < cch; i++) {
      s += String.fromCharCode(this.view.getUint16(this.pos, true))
      this.pos += 2
    }
    return s
  }
}

/** Decode an RK number (MS-XLSB §2.5.122 RkNumber). */
export function decodeRk(rk: number): number {
  const fX100 = (rk & 1) !== 0
  const fInt = (rk & 2) !== 0
  let value: number
  if (fInt) {
    value = (rk | 0) >> 2 // arithmetic shift — signed 30-bit integer
  } else {
    // The 30 high bits are the most-significant bits of an IEEE-754 double;
    // the low 34 bits are zero. Place them in the high word.
    const buf = new ArrayBuffer(8)
    const dv = new DataView(buf)
    dv.setUint32(4, rk & 0xfffffffc, true)
    value = dv.getFloat64(0, true)
  }
  return fX100 ? value / 100 : value
}
