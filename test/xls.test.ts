import { describe, expect, it } from "vitest"
import { writeCfb } from "../src/xlsx/crypto/cfb"
import { readXls } from "../src/xls/reader"
import { read } from "../src/defter"

// ── Minimal BIFF8 .xls builder (test-only) ───────────────────────────

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
function f64(n: number): number[] {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return [...b]
}
// XLUnicodeString (u16 cch) / ShortXLUnicodeString (u8 cch), compressed
const xlStr = (s: string): number[] => [...u16(s.length), 0, ...[...s].map((c) => c.charCodeAt(0))]
const shortStr = (s: string): number[] => [s.length, 0, ...[...s].map((c) => c.charCodeAt(0))]
function record(sid: number, data: number[]): number[] {
  return [...u16(sid), ...u16(data.length), ...data]
}
const rkInt = (v: number): number[] => u32(((v << 2) | 2) >>> 0)

const SID = {
  FORMULA: 0x0006,
  EOF: 0x000a,
  DATEMODE: 0x0022,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  RK: 0x027e,
  MULRK: 0x00bd,
  LABELSST: 0x00fd,
  SST: 0x00fc,
  XF: 0x00e0,
  BOUNDSHEET: 0x0085,
  MERGECELLS: 0x00e5,
  BOF: 0x0809,
}

const bof = (dt: number): number[] =>
  record(SID.BOF, [...u16(0x0600), ...u16(dt), ...u16(0), ...u16(0), ...u32(0), ...u32(0)])
const eof = (): number[] => record(SID.EOF, [])

function sstRecord(strings: string[]): number[] {
  const body: number[] = [...u32(strings.length), ...u32(strings.length)]
  for (const s of strings) body.push(...u16(s.length), 0, ...[...s].map((c) => c.charCodeAt(0)))
  return record(SID.SST, body)
}

function buildXls(): Uint8Array {
  const strings = ["Name", "Score", "Ada"]

  const sheet = concat([
    bof(0x0010),
    record(SID.LABELSST, [...u16(0), ...u16(0), ...u16(0), ...u32(0)]),
    record(SID.LABELSST, [...u16(0), ...u16(1), ...u16(0), ...u32(1)]),
    record(SID.LABELSST, [...u16(1), ...u16(0), ...u16(0), ...u32(2)]),
    record(SID.RK, [...u16(1), ...u16(1), ...u16(0), ...rkInt(95)]),
    record(SID.NUMBER, [...u16(1), ...u16(2), ...u16(0), ...f64(3.14)]),
    record(SID.NUMBER, [...u16(1), ...u16(3), ...u16(1), ...f64(45000)]), // date xf
    record(SID.LABEL, [...u16(2), ...u16(0), ...u16(0), ...xlStr("Hi")]),
    record(SID.BOOLERR, [...u16(2), ...u16(1), ...u16(0), 1, 0]), // true
    record(SID.BOOLERR, [...u16(2), ...u16(2), ...u16(0), 0x07, 1]), // #DIV/0!
    record(SID.MULRK, [
      ...u16(3),
      ...u16(0),
      ...u16(0),
      ...rkInt(10),
      ...u16(0),
      ...rkInt(20),
      ...u16(1),
    ]),
    record(SID.MERGECELLS, [...u16(1), ...u16(0), ...u16(0), ...u16(0), ...u16(1)]),
    eof(),
  ])

  // Globals — BOUNDSHEET position is filled in once the globals size is known.
  const makeGlobals = (sheetPos: number): Uint8Array =>
    concat([
      bof(0x0005),
      record(SID.DATEMODE, u16(0)),
      record(SID.XF, [...u16(0), ...u16(0), ...Array.from({ length: 16 }, () => 0)]), // general
      record(SID.XF, [...u16(0), ...u16(14), ...Array.from({ length: 16 }, () => 0)]), // date (builtin 14)
      sstRecord(strings),
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr("Sheet1")]),
      eof(),
    ])

  const globalsLen = makeGlobals(0).length
  const globals = makeGlobals(globalsLen)
  const workbookStream = concat([globals, sheet])

  return writeCfb([{ name: "Workbook", data: workbookStream }])
}

describe("XLS (BIFF8) reader", () => {
  it("decodes SST labels, RK, MULRK, numbers, bools, errors, dates, and merges", async () => {
    const wb = await readXls(buildXls())
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
    expect(rows[3][0]).toBe(10)
    expect(rows[3][1]).toBe(20)
    expect(wb.sheets[0].merges).toEqual([{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }])
  })

  it("is auto-detected by read()", async () => {
    const wb = await read(buildXls())
    expect(wb.sheets[0].rows[1][0]).toBe("Ada")
    expect(wb.sheets[0].rows[1][1]).toBe(95)
  })
})
