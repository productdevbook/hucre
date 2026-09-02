import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsb } from "../src/xlsx/xlsb/reader"
import { read } from "../src/defter"
import { ParseError } from "../src/errors"

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
  BrtFmt = 44,
  BrtXF = 47,
  BrtWbProp = 153,
  BrtBundleSh = 156,
  BrtMergeCell = 176,
  BrtBeginCellXFs = 617,
  BrtEndCellXFs = 618

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const NS = "http://schemas.openxmlformats.org/package/2006/relationships"

/**
 * BrtWbProp (record 153): 32-bit flag word (bit 0 = f1904), then
 * dwThemeVersion, then strName — here the "absent" nullable string.
 */
const wbProp = (date1904: boolean): Uint8Array =>
  rec(BrtWbProp, concat([u32(date1904 ? 1 : 0), u32(0), nwstr(null)]))

async function buildXlsb(
  opts: {
    date1904?: boolean
    dateFmtId?: number
    fmtCodes?: Array<[number, string]>
    value?: number
  } = {},
): Promise<Uint8Array> {
  // Shared strings: 0:"Name" 1:"Score" 2:"Ada"
  const sst = concat([
    rec(BrtSSTItem, concat([[0], wstr("Name")])),
    rec(BrtSSTItem, concat([[0], wstr("Score")])),
    rec(BrtSSTItem, concat([[0], wstr("Ada")])),
  ])
  // Styles: xf0 general (iFmt 0), xf1 date (a built-in date iFmt).
  // The id is a parameter because the built-in date set is wider than the
  // familiar 14-22 block — see the CJK case below.
  const dateFmtId = opts.dateFmtId ?? 14
  const styles = concat([
    // BrtFmt — the workbook's own number-format definitions, which may
    // redefine a built-in id. See #568.
    ...(opts.fmtCodes ?? []).map(([id, code]) => rec(BrtFmt, concat([u16(id), wstr(code)]))),
    rec(BrtBeginCellXFs, u32(2)),
    rec(BrtXF, concat([u16(0), u16(0), u16(0), u16(0), u16(0), [0, 0]])),
    rec(BrtXF, concat([u16(0), u16(dateFmtId), u16(0), u16(0), u16(0), [0, 0]])),
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
    rec(BrtCellReal, concat([cellPrefix(3, 1), f64(opts.value ?? 45000)])), // date serial via date xf
    rec(BrtRowHdr, u32(2)),
    rec(BrtCellSt, concat([cellPrefix(0, 0), wstr("Hi")])),
    rec(BrtCellBool, concat([cellPrefix(1, 0), [1]])),
    rec(BrtCellError, concat([cellPrefix(2, 0), [0x07]])),
    // BrtMergeCell: one UncheckedRfX — rwFirst, rwLast, colFirst, colLast.
    rec(BrtMergeCell, concat([u32(0), u32(0), u32(0), u32(1)])),
  ])
  const wb = concat([
    ...(opts.date1904 === undefined ? [] : [wbProp(opts.date1904)]),
    rec(BrtBundleSh, concat([u32(0), u32(0), nwstr("rId1"), wstr("Sheet1")])),
  ])
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
    // Padded to the sheet width, not to this row's own last cell:
    // `rows` is a dense rectangle, which these readers used to leave
    // ragged while readXlsx did not. See #494.
    expect(rows[0]).toEqual(["Name", "Score", null, null])
    expect(rows[1][0]).toBe("Ada")
    expect(rows[1][1]).toBe(95)
    expect(rows[1][2]).toBeCloseTo(3.14, 5)
    expect(rows[1][3]).toBeInstanceOf(Date)
    expect(rows[2][0]).toBe("Hi")
    expect(rows[2][1]).toBe(true)
    expect(rows[2][2]).toBe("#DIV/0!")
  })

  it("honours maxTotalCells — the one reader that used to have no ceiling", async () => {
    // 3 rows × 4 columns is 12 slots; a cap under that is refused before
    // the grid is allocated, as readXlsx / readOds / readXls already did.
    await expect(readXlsb(await buildXlsb(), { maxTotalCells: 4 })).rejects.toThrow(ParseError)
    await expect(readXlsb(await buildXlsb(), { maxTotalCells: 4 })).rejects.toThrow(/maxTotalCells/)
    const wb = await readXlsb(await buildXlsb(), { maxTotalCells: 12 })
    expect(wb.sheets[0].rows).toHaveLength(3)
  })

  it("reads merged ranges from BrtMergeCell", async () => {
    // XLS has read merges since it landed; XLSB ignored record 176
    // entirely, so a converted workbook lost its merge layout. See #411.
    const wb = await readXlsb(await buildXlsb())
    expect(wb.sheets[0].merges).toEqual([{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }])
  })

  it("leaves merges undefined when the sheet has none", async () => {
    const ws = concat([
      rec(BrtRowHdr, u32(0)),
      rec(BrtCellSt, concat([cellPrefix(0, 0), wstr("x")])),
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
    expect(out.sheets[0].merges).toBeUndefined()
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

  // ── Date system (#411) ────────────────────────────────────────────
  // Serial 45000 with the builtin date format: 1900 → 2023-03-15,
  // 1904 → 2027-03-16. The two systems are 1462 days apart, so reading a
  // Mac-authored workbook with the wrong one is silent and always wrong.
  const serial45000 = (wb: { sheets: Array<{ rows: unknown[][] }> }): string =>
    (wb.sheets[0].rows[1][3] as Date).toISOString()

  describe("date system", () => {
    it("honours the 1904 flag in BrtWbProp by default", async () => {
      const wb = await readXlsb(await buildXlsb({ date1904: true }))
      expect(serial45000(wb)).toBe("2027-03-16T00:00:00.000Z")
    })

    it("uses 1900 when BrtWbProp says so", async () => {
      const wb = await readXlsb(await buildXlsb({ date1904: false }))
      expect(serial45000(wb)).toBe("2023-03-15T00:00:00.000Z")
    })

    it("uses 1900 when the workbook has no BrtWbProp at all", async () => {
      const wb = await readXlsb(await buildXlsb())
      expect(serial45000(wb)).toBe("2023-03-15T00:00:00.000Z")
    })

    it("detects the file's system under an explicit dateSystem: auto", async () => {
      const wb = await readXlsb(await buildXlsb({ date1904: true }), { dateSystem: "auto" })
      expect(serial45000(wb)).toBe("2027-03-16T00:00:00.000Z")
    })

    it("lets the caller pin a system that contradicts the file", async () => {
      const pinned1900 = await readXlsb(await buildXlsb({ date1904: true }), {
        dateSystem: "1900",
      })
      expect(serial45000(pinned1900)).toBe("2023-03-15T00:00:00.000Z")

      const pinned1904 = await readXlsb(await buildXlsb({ date1904: false }), {
        dateSystem: "1904",
      })
      expect(serial45000(pinned1904)).toBe("2027-03-16T00:00:00.000Z")
    })
  })

  // ── #439: the built-in date set is wider than 14-22 / 45-47 ──────────
  // The CJK block (27-36) and the Thai/Chinese/Korean block (50-58) are
  // date and time formats too, and they carry no formatCode in the file —
  // so a reader that does not know them has no fallback and hands back the
  // raw serial. This reader used to keep a 12-entry table of its own.
  describe("built-in date format ids outside the familiar block", () => {
    const CJK_AND_EXTENDED = [
      27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58,
    ]

    for (const id of CJK_AND_EXTENDED) {
      it(`reads a cell styled with built-in format ${id} as a Date`, async () => {
        const wb = await readXlsb(await buildXlsb({ dateFmtId: id }))

        expect(wb.sheets[0].rows[1][3]).toBeInstanceOf(Date)
      })
    }

    it("still treats a non-date built-in as a number", async () => {
      // 3 is "#,##0" — a numeric built-in, and one that carries no
      // formatCode either, so it exercises the same fallback path.
      const wb = await readXlsb(await buildXlsb({ dateFmtId: 3 }))

      expect(wb.sheets[0].rows[1][3]).toBe(45000)
    })
  })

  // ── #568: a BrtFmt record redefining a built-in id ──────────────────
  // The built-in table was consulted first, so the file's own BrtFmt for
  // that id was never read — the same disagreement with readXlsx that
  // xls/reader.ts had.
  describe("a BrtFmt record redefines a built-in id", () => {
    it("reads a number when the file redefines a built-in date id numerically", async () => {
      const wb = await readXlsb(
        await buildXlsb({ dateFmtId: 50, fmtCodes: [[50, "000000000000"]], value: 81227827687 }),
      )

      expect(wb.sheets[0].rows[1][3]).toBe(81227827687)
    })

    it("reads a number when the file redefines id 14 as '#,##0'", async () => {
      const wb = await readXlsb(await buildXlsb({ dateFmtId: 14, fmtCodes: [[14, "#,##0"]] }))

      expect(wb.sheets[0].rows[1][3]).toBe(45000)
    })

    it("reads a Date when the file redefines a numeric built-in id as a date", async () => {
      const wb = await readXlsb(await buildXlsb({ dateFmtId: 3, fmtCodes: [[3, "yyyy-mm-dd"]] }))

      expect(wb.sheets[0].rows[1][3]).toBeInstanceOf(Date)
    })

    it("keeps the built-in meaning when the file redefines some other id", async () => {
      const wb = await readXlsb(
        await buildXlsb({ dateFmtId: 50, fmtCodes: [[164, "000000000000"]] }),
      )

      expect(wb.sheets[0].rows[1][3]).toBeInstanceOf(Date)
    })
  })

  it("surfaces a malformed workbook.bin as ParseError, not a raw RangeError", async () => {
    // A BrtBundleSh record whose body is truncated (only the 4-byte hsState,
    // missing iTabID/relId/name) makes the Cursor read past the end.
    const wb = rec(BrtBundleSh, u32(0))
    const rels = `<?xml version="1.0"?><Relationships xmlns="${NS}"><Relationship Id="r" Type="${REL}/officeDocument" Target="xl/workbook.bin"/></Relationships>`
    const ct = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`
    const zw = new ZipWriter()
    zw.add("[Content_Types].xml", enc.encode(ct))
    zw.add("_rels/.rels", enc.encode(rels))
    zw.add("xl/workbook.bin", wb)

    await expect(readXlsb(await zw.build())).rejects.toBeInstanceOf(ParseError)
  })
})
