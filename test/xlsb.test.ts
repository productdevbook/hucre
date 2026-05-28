import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsb } from "../src/xlsx/xlsb/reader"
import { read } from "../src/defter"

// ── Minimal XLSB builder (test-only) ─────────────────────────────────
// Emits valid MS-XLSB binary records so the reader can be round-tripped
// without an external fixture.

const enc = new TextEncoder()

function concat(parts: Array<number[] | Uint8Array>): Uint8Array {
  let len = 0
  for (const p of parts) len += p.length
  const out = new Uint8Array(len)
  let off = 0
  for (const p of parts) {
    out.set(p instanceof Uint8Array ? p : new Uint8Array(p), off)
    off += p.length
  }
  return out
}

const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]
const u32 = (n: number): number[] => [
  n & 0xff,
  (n >> 8) & 0xff,
  (n >> 16) & 0xff,
  (n >>> 24) & 0xff,
]
function f64(n: number): Uint8Array {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return b
}
function wstr(s: string): Uint8Array {
  const out = new Uint8Array(4 + s.length * 2)
  const dv = new DataView(out.buffer)
  dv.setUint32(0, s.length, true)
  for (let i = 0; i < s.length; i++) dv.setUint16(4 + i * 2, s.charCodeAt(i), true)
  return out
}
const nwstr = (s: string | null): Uint8Array =>
  s === null ? new Uint8Array(u32(0xffffffff)) : wstr(s)
function varint(n: number): number[] {
  const out: number[] = []
  let s = n
  do {
    let b = s & 0x7f
    s >>>= 7
    if (s) b |= 0x80
    out.push(b)
  } while (s)
  return out
}
function rec(id: number, payload: Uint8Array | number[]): Uint8Array {
  const body = payload instanceof Uint8Array ? payload : new Uint8Array(payload)
  const idBytes = id < 0x80 ? [id] : [(id & 0x7f) | 0x80, (id >> 7) & 0x7f]
  return concat([idBytes, varint(body.length), body])
}
const rkInt = (v: number): number[] => u32(((v << 2) | 2) >>> 0) // fInt set
const cellPrefix = (col: number, style: number): number[] => [...u32(col), ...u32(style & 0xffffff)]

// record ids (MS-XLSB §2.4)
const BrtRowHdr = 0,
  BrtCellRk = 2,
  BrtCellError = 3,
  BrtCellBool = 4,
  BrtCellReal = 5,
  BrtCellSt = 6,
  BrtCellIsst = 7,
  BrtSSTItem = 19,
  BrtXF = 47,
  BrtBundleSh = 156,
  BrtBeginCellXFs = 617,
  BrtEndCellXFs = 618

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const NS = "http://schemas.openxmlformats.org/package/2006/relationships"

async function buildXlsb(): Promise<Uint8Array> {
  // Shared strings: 0:"Name" 1:"Score" 2:"Ada"
  const sst = concat([
    rec(BrtSSTItem, concat([[0], wstr("Name")])),
    rec(BrtSSTItem, concat([[0], wstr("Score")])),
    rec(BrtSSTItem, concat([[0], wstr("Ada")])),
  ])
  // Styles: xf0 general (iFmt 0), xf1 date (iFmt 14 builtin)
  const styles = concat([
    rec(BrtBeginCellXFs, u32(2)),
    rec(BrtXF, concat([u16(0), u16(0), u16(0), u16(0), u16(0), [0, 0]])),
    rec(BrtXF, concat([u16(0), u16(14), u16(0), u16(0), u16(0), [0, 0]])),
    rec(BrtEndCellXFs, []),
  ])
  // Worksheet rows.
  const ws = concat([
    rec(BrtRowHdr, u32(0)),
    rec(BrtCellIsst, concat([cellPrefix(0, 0), u32(0)])),
    rec(BrtCellIsst, concat([cellPrefix(1, 0), u32(1)])),
    rec(BrtRowHdr, u32(1)),
    rec(BrtCellIsst, concat([cellPrefix(0, 0), u32(2)])),
    rec(BrtCellRk, concat([cellPrefix(1, 0), rkInt(95)])),
    rec(BrtCellReal, concat([cellPrefix(2, 0), f64(3.14)])),
    rec(BrtCellReal, concat([cellPrefix(3, 1), f64(45000)])), // date serial via date xf
    rec(BrtRowHdr, u32(2)),
    rec(BrtCellSt, concat([cellPrefix(0, 0), wstr("Hi")])),
    rec(BrtCellBool, concat([cellPrefix(1, 0), [1]])),
    rec(BrtCellError, concat([cellPrefix(2, 0), [0x07]])),
  ])
  const wb = rec(BrtBundleSh, concat([u32(0), u32(0), nwstr("rId1"), wstr("Sheet1")]))
  const rels = `<?xml version="1.0"?><Relationships xmlns="${NS}"><Relationship Id="rIdWb" Type="${REL}/officeDocument" Target="xl/workbook.bin"/></Relationships>`
  const wbRels =
    `<?xml version="1.0"?><Relationships xmlns="${NS}">` +
    `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.bin"/>` +
    `<Relationship Id="rId2" Type="${REL}/sharedStrings" Target="sharedStrings.bin"/>` +
    `<Relationship Id="rId3" Type="${REL}/styles" Target="styles.bin"/></Relationships>`
  const ct = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`

  const zw = new ZipWriter()
  zw.add("[Content_Types].xml", enc.encode(ct))
  zw.add("_rels/.rels", enc.encode(rels))
  zw.add("xl/workbook.bin", wb)
  zw.add("xl/_rels/workbook.bin.rels", enc.encode(wbRels))
  zw.add("xl/sharedStrings.bin", sst)
  zw.add("xl/styles.bin", styles)
  zw.add("xl/worksheets/sheet1.bin", ws)
  return zw.build()
}

describe("XLSB reader", () => {
  it("decodes shared strings, RK ints, reals, inline strings, bools, errors, and dates", async () => {
    const wb = await readXlsb(await buildXlsb())
    expect(wb.sheets.length).toBe(1)
    expect(wb.sheets[0].name).toBe("Sheet1")
    const rows = wb.sheets[0].rows
    expect(rows[0]).toEqual(["Name", "Score"])
    expect(rows[1][0]).toBe("Ada")
    expect(rows[1][1]).toBe(95)
    expect(rows[1][2]).toBeCloseTo(3.14, 5)
    expect(rows[1][3]).toBeInstanceOf(Date)
    expect(rows[2][0]).toBe("Hi")
    expect(rows[2][1]).toBe(true)
    expect(rows[2][2]).toBe("#DIV/0!")
  })

  it("is auto-detected by read()", async () => {
    const wb = await read(await buildXlsb())
    expect(wb.sheets[0].rows[1][1]).toBe(95)
    expect(wb.sheets[0].rows[1][0]).toBe("Ada")
  })

  it("decodes RK fractional (x100) numbers", async () => {
    // rkInt with x100: value 1234 with fX100 → 12.34
    const rkX100 = (cents: number): number[] => u32(((cents << 2) | 2 | 1) >>> 0)
    const ws = concat([
      rec(BrtRowHdr, u32(0)),
      rec(BrtCellRk, concat([cellPrefix(0, 0), rkX100(1234)])),
    ])
    const wb = rec(BrtBundleSh, concat([u32(0), u32(0), nwstr("rId1"), wstr("S")]))
    const rels = `<?xml version="1.0"?><Relationships xmlns="${NS}"><Relationship Id="r" Type="${REL}/officeDocument" Target="xl/workbook.bin"/></Relationships>`
    const wbRels = `<?xml version="1.0"?><Relationships xmlns="${NS}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.bin"/></Relationships>`
    const zw = new ZipWriter()
    zw.add("_rels/.rels", enc.encode(rels))
    zw.add("xl/workbook.bin", wb)
    zw.add("xl/_rels/workbook.bin.rels", enc.encode(wbRels))
    zw.add("xl/worksheets/sheet1.bin", ws)
    const out = await readXlsb(await zw.build())
    expect(out.sheets[0].rows[0][0]).toBeCloseTo(12.34, 5)
  })
})
