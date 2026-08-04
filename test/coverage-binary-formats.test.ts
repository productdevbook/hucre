import { describe, expect, it, vi } from "vitest"

import { readXls } from "../src/xls/reader"
import { readXlsb } from "../src/xlsx/xlsb/reader"
import { readCfb, writeCfb } from "../src/xlsx/crypto/cfb"
import { decryptAgile, encryptAgile } from "../src/xlsx/crypto/agile"
import { parseRelationships } from "../src/xlsx/relationships"
import { parseContentTypes } from "../src/xlsx/content-types"
import { fetchCsv } from "../src/csv/fetch"
import { ZipWriter } from "../src/zip/writer"
import { DecryptionError, EncryptedFileError, ZipError } from "../src/errors"

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
function f64(n: number): number[] {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return [...b]
}

// ═══════════════════════════════════════════════════════════════════════
// XLS (BIFF8)
// ═══════════════════════════════════════════════════════════════════════

// ── Minimal BIFF8 builder ────────────────────────────────────────────

const SID = {
  FORMULA: 0x0006,
  EOF: 0x000a,
  CONTINUE: 0x003c,
  DATEMODE: 0x0022,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  STRING: 0x0207,
  ROW: 0x0208,
  RK: 0x027e,
  MULRK: 0x00bd,
  LABELSST: 0x00fd,
  SST: 0x00fc,
  XF: 0x00e0,
  FORMAT: 0x041e,
  BOUNDSHEET: 0x0085,
  BOF: 0x0809,
}

function record(sid: number, data: number[]): number[] {
  return [...u16(sid), ...u16(data.length), ...data]
}
const bof = (dt: number): number[] =>
  record(SID.BOF, [...u16(0x0600), ...u16(dt), ...u16(0), ...u16(0), ...u32(0), ...u32(0)])
const eof = (): number[] => record(SID.EOF, [])
const chars = (s: string): number[] => [...s].map((c) => c.charCodeAt(0))
/** XLUnicodeString (u16 cch + grbit + chars), compressed 8-bit. */
const xlStr = (s: string): number[] => [...u16(s.length), 0, ...chars(s)]
/** ShortXLUnicodeString (u8 cch + grbit + chars), compressed 8-bit. */
const shortStr = (s: string): number[] => [s.length, 0, ...chars(s)]
const xf = (ifmt: number): number[] =>
  record(SID.XF, [...u16(0), ...u16(ifmt), ...Array.from({ length: 16 }, () => 0)])

/**
 * Assemble a Workbook stream: globals substream (BOF … EOF, with one
 * BOUNDSHEET per sheet) followed by each sheet substream. BOUNDSHEET
 * positions are resolved in a second pass once the globals size is known.
 */
function buildWorkbookStream(opts: {
  globals?: number[]
  sheets?: Array<{ name: string; records: number[]; positionOverride?: number }>
}): Uint8Array {
  const sheets = opts.sheets ?? []
  const make = (positions: number[]): number[] => [
    ...bof(0x0005),
    ...(opts.globals ?? []),
    ...sheets
      .map((s, i) =>
        record(SID.BOUNDSHEET, [
          ...u32(s.positionOverride ?? positions[i]),
          0,
          0,
          ...shortStr(s.name),
        ]),
      )
      .flat(),
    ...eof(),
  ]
  const globalsLen = make(sheets.map(() => 0)).length
  const positions: number[] = []
  let pos = globalsLen
  for (const s of sheets) {
    positions.push(pos)
    pos += s.records.length
  }
  return concat([make(positions), ...sheets.map((s) => s.records)])
}

function xlsFile(opts: Parameters<typeof buildWorkbookStream>[0], streamName = "Workbook") {
  return writeCfb([{ name: streamName, data: buildWorkbookStream(opts) }])
}

describe("readXls — container and version gates", () => {
  it("reads a workbook stored under the legacy 'Book' stream name", async () => {
    // Excel 95 named the stream "Book"; some BIFF8 producers kept the name.
    const data = xlsFile(
      {
        sheets: [
          {
            name: "S",
            records: [
              ...bof(0x0010),
              ...record(SID.LABEL, [...u16(0), ...u16(0), ...u16(0), ...xlStr("Hi")]),
              ...eof(),
            ],
          },
        ],
      },
      "Book",
    )
    const wb = await readXls(data)
    expect(wb.sheets[0].rows[0][0]).toBe("Hi")
  })

  it("rejects an OLE2 container with no workbook stream", async () => {
    const data = writeCfb([{ name: "SummaryInformation", data: new Uint8Array(64) }])
    await expect(readXls(data)).rejects.toThrow(/missing Workbook stream/)
  })

  it("rejects a workbook stream that does not open with a BOF record", async () => {
    const data = writeCfb([{ name: "Workbook", data: new Uint8Array(record(SID.EOF, [])) }])
    await expect(readXls(data)).rejects.toThrow(/missing BOF record/)
  })

  it("accepts a BOF record too short to carry a version field", async () => {
    // Nothing to gate on — parse it rather than guessing at the version.
    const stream = concat([record(SID.BOF, []), eof()])
    const wb = await readXls(writeCfb([{ name: "Workbook", data: stream }]))
    expect(wb.sheets).toEqual([])
  })

  it("surfaces a non-OLE2 input as a ParseError", async () => {
    await expect(readXls(new Uint8Array(600))).rejects.toThrow(/not a valid OLE2 container/)
  })
})

describe("readXls — globals substream", () => {
  it("honours the DATEMODE record when the caller asks for auto detection", async () => {
    // dateSystem:"auto" must consult DATEMODE; 1 selects the 1904 epoch.
    const data = xlsFile({
      globals: [...record(SID.DATEMODE, u16(1)), ...xf(14)],
      sheets: [
        {
          name: "S",
          records: [
            ...bof(0x0010),
            ...record(SID.NUMBER, [...u16(0), ...u16(0), ...u16(0), ...f64(1)]),
            ...eof(),
          ],
        },
      ],
    })
    const auto = await readXls(data, { dateSystem: "auto" })
    const forced = await readXls(data, { dateSystem: "1900" })
    expect(auto.sheets[0].rows[0][0]).toBeInstanceOf(Date)
    expect((auto.sheets[0].rows[0][0] as Date).getUTCFullYear()).toBe(1904)
    expect((forced.sheets[0].rows[0][0] as Date).getUTCFullYear()).toBe(1900)
    expect(auto.sheets[0].rows[0][0]).not.toEqual(forced.sheets[0].rows[0][0])
  })

  it("treats a cell as a date when a custom FORMAT record says so", async () => {
    // Format ids >= 164 are workbook-defined and only resolvable through
    // the FORMAT records.
    const data = xlsFile({
      globals: [...record(SID.FORMAT, [...u16(200), ...xlStr("dd/mm/yyyy")]), ...xf(200), ...xf(0)],
      sheets: [
        {
          name: "S",
          records: [
            ...bof(0x0010),
            ...record(SID.NUMBER, [...u16(0), ...u16(0), ...u16(0), ...f64(45000)]),
            ...record(SID.NUMBER, [...u16(0), ...u16(1), ...u16(1), ...f64(45000)]),
            ...eof(),
          ],
        },
      ],
    })
    const wb = await readXls(data)
    expect(wb.sheets[0].rows[0][0]).toBeInstanceOf(Date)
    expect(wb.sheets[0].rows[0][1]).toBe(45000)
  })

  it("returns an empty sheet when BOUNDSHEET points at no record", async () => {
    const data = xlsFile({
      sheets: [{ name: "Ghost", records: [...bof(0x0010), ...eof()], positionOverride: 0xdead }],
    })
    const wb = await readXls(data)
    expect(wb.sheets).toEqual([{ name: "Ghost", rows: [] }])
  })
})

describe("readXls — cell records", () => {
  function sheetWith(records: number[], globals?: number[]) {
    return xlsFile({
      globals: globals ?? [...xf(0)],
      sheets: [{ name: "S", records: [...bof(0x0010), ...records, ...eof()] }],
    })
  }

  it("falls back to an empty string for an out-of-range shared-string index", async () => {
    const data = sheetWith(record(SID.LABELSST, [...u16(0), ...u16(0), ...u16(0), ...u32(99)]))
    expect((await readXls(data)).sheets[0].rows[0][0]).toBe("")
  })

  it("falls back to #ERR! for an error code it does not know", async () => {
    const data = sheetWith(record(SID.BOOLERR, [...u16(0), ...u16(0), ...u16(0), 0x55, 1]))
    expect((await readXls(data)).sheets[0].rows[0][0]).toBe("#ERR!")
  })

  it("decodes an RK number stored as a truncated double", async () => {
    // fInt clear: the 30 high bits are the top of an IEEE-754 double.
    // 0x3FF00000 is 1.0, and the fX100 flag divides it by 100.
    const data = sheetWith(
      concat([
        record(SID.RK, [...u16(0), ...u16(0), ...u16(0), ...u32(0x3ff00000)]),
        record(SID.RK, [...u16(0), ...u16(1), ...u16(0), ...u32(0x3ff00001)]),
      ]) as unknown as number[],
    )
    const rows = (await readXls(data)).sheets[0].rows
    expect(rows[0][0]).toBe(1)
    expect(rows[0][1]).toBeCloseTo(0.01, 6)
  })

  it("decodes an uncompressed (UTF-16) label", async () => {
    // grbit bit 0 set means the characters are 16-bit, not codepage bytes.
    const wide = [...u16(2), 1, ...u16(0x00c9), ...u16(0x011f)] // "Éğ"
    const data = sheetWith(record(SID.LABEL, [...u16(0), ...u16(0), ...u16(0), ...wide]))
    expect((await readXls(data)).sheets[0].rows[0][0]).toBe("Éğ")
  })

  it("ignores record types it has no decoder for", async () => {
    const data = sheetWith(
      concat([
        record(SID.ROW, [...u16(0), ...u16(0), ...u16(1), ...u16(255), ...u32(0), ...u32(0)]),
        record(SID.LABEL, [...u16(0), ...u16(0), ...u16(0), ...xlStr("ok")]),
      ]) as unknown as number[],
    )
    expect((await readXls(data)).sheets[0].rows[0][0]).toBe("ok")
  })

  it("stops parsing at the padding record that ends the stream", async () => {
    // Trailing zero bytes in the CFB sector look like a record with id 0 and
    // length 0; treating them as data would append junk records forever.
    const stream = concat([
      buildWorkbookStream({
        globals: [...xf(0)],
        sheets: [
          {
            name: "S",
            records: [
              ...bof(0x0010),
              ...record(SID.LABEL, [...u16(0), ...u16(0), ...u16(0), ...xlStr("ok")]),
              ...eof(),
            ],
          },
        ],
      }),
      new Uint8Array(64),
    ])
    const wb = await readXls(writeCfb([{ name: "Workbook", data: stream }]))
    expect(wb.sheets[0].rows[0][0]).toBe("ok")
  })

  it("rejects a column index outside the supported sheet bounds", async () => {
    const data = sheetWith(record(SID.LABEL, [...u16(0), ...u16(20000), ...u16(0), ...xlStr("x")]))
    await expect(readXls(data)).rejects.toThrow(/outside the supported sheet bounds/)
  })

  it("rejects a sheet whose bounding box exceeds the cell budget", async () => {
    // Two legal BIFF coordinates can describe a rectangle of 20M+ slots.
    const data = sheetWith(
      concat([
        record(SID.LABEL, [...u16(0), ...u16(16000), ...u16(0), ...xlStr("x")]),
        record(SID.LABEL, [...u16(1300), ...u16(0), ...u16(0), ...xlStr("y")]),
      ]) as unknown as number[],
    )
    await expect(readXls(data)).rejects.toThrow(/over the 20000000 limit/)
  })
})

describe("readXls — FORMULA cached values", () => {
  /** A FORMULA record whose 8-byte cached value is given verbatim. */
  const formula = (row: number, col: number, cached: number[]): number[] =>
    record(SID.FORMULA, [
      ...u16(row),
      ...u16(col),
      ...u16(0),
      ...cached,
      ...u16(0),
      ...u32(0),
      ...u16(0),
    ])

  async function read(records: number[]) {
    const data = xlsFile({
      globals: [...xf(0)],
      sheets: [{ name: "S", records: [...bof(0x0010), ...records, ...eof()] }],
    })
    return (await readXls(data)).sheets[0].rows
  }

  it("decodes a cached boolean", async () => {
    const rows = await read(formula(0, 0, [1, 0, 1, 0, 0, 0, 0xff, 0xff]))
    expect(rows[0][0]).toBe(true)
  })

  it("decodes a cached error", async () => {
    const rows = await read(formula(0, 0, [2, 0, 0x17, 0, 0, 0, 0xff, 0xff]))
    expect(rows[0][0]).toBe("#REF!")
  })

  it("falls back to #ERR! for an unknown cached error code", async () => {
    const rows = await read(formula(0, 0, [2, 0, 0x55, 0, 0, 0, 0xff, 0xff]))
    expect(rows[0][0]).toBe("#ERR!")
  })

  it("takes a cached string from the STRING record that follows", async () => {
    const rows = await read([
      ...formula(0, 0, [0, 0, 0, 0, 0, 0, 0xff, 0xff]),
      ...record(SID.STRING, xlStr("computed")),
    ])
    expect(rows[0][0]).toBe("computed")
  })

  it("leaves the cell empty when the STRING record is missing", async () => {
    const rows = await read(formula(0, 0, [0, 0, 0, 0, 0, 0, 0xff, 0xff]))
    expect(rows[0]).toBeUndefined()
  })

  it("leaves the cell empty for a cached blank", async () => {
    const rows = await read(formula(0, 0, [3, 0, 0, 0, 0, 0, 0xff, 0xff]))
    expect(rows[0]).toBeUndefined()
  })

  it("decodes a cached number", async () => {
    const rows = await read(formula(0, 0, f64(12.5)))
    expect(rows[0][0]).toBeCloseTo(12.5, 6)
  })
})

describe("parseSst — rich, phonetic, wide and continued strings", () => {
  it("decodes every SST string flavour, including one split across CONTINUE", async () => {
    // Layout per MS-XLS §2.4.265: cch, grbit, [cRun], [cbExt], characters,
    // then the rich-run and phonetic trailers the reader has to skip.
    const plain = [...u16(4), 0, ...chars("Name")]
    const rich = [...u16(4), 0x08, ...u16(1), ...chars("Rich"), 0, 0, 0, 0]
    const phonetic = [...u16(4), 0x04, ...u32(6), ...chars("Phon"), 0, 0, 0, 0, 0, 0]
    const wide = [...u16(3), 0x01, ...u16(0x00dc), ...u16(0x011f), ...u16(0x015f)]
    // Last in the SST record: only the first three characters fit, the rest
    // resume in the CONTINUE record behind a fresh option byte.
    const splitHead = [...u16(6), 0, ...chars("Con")]
    const splitTail = [0, ...chars("tin")]

    const sst = record(SID.SST, [
      ...u32(5),
      ...u32(5),
      ...plain,
      ...rich,
      ...phonetic,
      ...wide,
      ...splitHead,
    ])
    const cont = record(SID.CONTINUE, splitTail)

    const labels = [0, 1, 2, 3, 4].map((i) =>
      record(SID.LABELSST, [...u16(0), ...u16(i), ...u16(0), ...u32(i)]),
    )
    const data = xlsFile({
      globals: [...xf(0), ...sst, ...cont],
      sheets: [{ name: "S", records: [...bof(0x0010), ...labels.flat(), ...eof()] }],
    })

    const rows = (await readXls(data)).sheets[0].rows
    expect(rows[0]).toEqual(["Name", "Rich", "Phon", "Üğş", "Contin"])
  })

  // BUG (reported): BlockStream.skip in src/xls/biff.ts:166-175 spins
  // forever once every block is exhausted — `ensure()` cannot advance past
  // the last block, `remainingInBlock()` then returned 0 and `left` never
  // decreased. readSstString reaches skip() with an untrusted cRun (a u16)
  // or cbExt (a u32), so an .xls claiming more trailer bytes than it
  // carries used to hang the process — a live lock readXls's try/catch
  // could not interrupt. Fixed in #389; the timeout keeps a regression a
  // failure rather than a hung CI job.
  it(
    "does not hang on a rich-run count that overruns the record",
    { timeout: 15_000 },
    async () => {
      const rich = [...u16(1), 0x08, ...u16(0xffff), ...chars("A")]
      const sst = record(SID.SST, [...u32(1), ...u32(1), ...rich])
      const data = xlsFile({
        globals: [...xf(0), ...sst],
        sheets: [{ name: "S", records: [...bof(0x0010), ...eof()] }],
      })
      await expect(readXls(data)).rejects.toThrow()
    },
  )
})

// ═══════════════════════════════════════════════════════════════════════
// XLSB
// ═══════════════════════════════════════════════════════════════════════

// ── Minimal XLSB builder ─────────────────────────────────────────────

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const NS = "http://schemas.openxmlformats.org/package/2006/relationships"

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
const cellPrefix = (col: number, style: number): number[] => [...u32(col), ...u32(style & 0xffffff)]

const Brt = {
  RowHdr: 0,
  CellBlank: 1,
  CellRk: 2,
  CellError: 3,
  CellBool: 4,
  CellReal: 5,
  CellSt: 6,
  CellIsst: 7,
  SSTItem: 19,
  Fmt: 44,
  XF: 47,
  BundleSh: 156,
  BeginCellXFs: 617,
  EndCellXFs: 618,
  BeginSst: 159,
}

const relsXml = (rels: string) =>
  `<?xml version="1.0"?><Relationships xmlns="${NS}">${rels}</Relationships>`

async function buildXlsbPackage(parts: Record<string, Uint8Array | string>): Promise<Uint8Array> {
  const zw = new ZipWriter()
  for (const [path, body] of Object.entries(parts)) {
    zw.add(path, typeof body === "string" ? enc.encode(body) : body)
  }
  return zw.build()
}

describe("readXlsb — package layout", () => {
  it("resolves a workbook part that lives at the package root", async () => {
    const pkg = await buildXlsbPackage({
      "_rels/.rels": relsXml(
        `<Relationship Id="r" Type="${REL}/officeDocument" Target="workbook.bin"/>`,
      ),
      // A BrtBeginBook-style record ahead of the sheet bundle: the reader
      // must walk past records it does not care about.
      "workbook.bin": concat([
        rec(131, []),
        rec(Brt.BundleSh, concat([u32(0), u32(0), nwstr("rId1"), wstr("Sheet1")])),
      ]),
      "_rels/workbook.bin.rels": relsXml(
        `<Relationship Id="rId1" Type="${REL}/worksheet" Target="sheet1.bin"/>`,
      ),
      "sheet1.bin": concat([
        rec(Brt.RowHdr, u32(0)),
        rec(Brt.CellSt, concat([cellPrefix(0, 0), wstr("root")])),
      ]),
    })
    const wb = await readXlsb(pkg)
    expect(wb.sheets[0].rows[0][0]).toBe("root")
  })

  it("resolves package-absolute relationship targets", async () => {
    const pkg = await buildXlsbPackage({
      "_rels/.rels": relsXml(
        `<Relationship Id="r" Type="${REL}/officeDocument" Target="/xl/workbook.bin"/>`,
      ),
      "xl/workbook.bin": rec(Brt.BundleSh, concat([u32(0), u32(0), nwstr("rId1"), wstr("Sheet1")])),
      "xl/_rels/workbook.bin.rels": relsXml(
        `<Relationship Id="rId1" Type="${REL}/worksheet" Target="/xl/worksheets/sheet1.bin"/>`,
      ),
      "xl/worksheets/sheet1.bin": concat([
        rec(Brt.RowHdr, u32(0)),
        rec(Brt.CellSt, concat([cellPrefix(0, 0), wstr("absolute")])),
      ]),
    })
    expect((await readXlsb(pkg)).sheets[0].rows[0][0]).toBe("absolute")
  })

  it("normalises '..' and '.' segments in a relationship target", async () => {
    const pkg = await buildXlsbPackage({
      "_rels/.rels": relsXml(
        `<Relationship Id="r" Type="${REL}/officeDocument" Target="xl/workbook.bin"/>`,
      ),
      "xl/workbook.bin": rec(Brt.BundleSh, concat([u32(0), u32(0), nwstr("rId1"), wstr("Sheet1")])),
      "xl/_rels/workbook.bin.rels": relsXml(
        `<Relationship Id="rId1" Type="${REL}/worksheet" Target="../xl/./worksheets/sheet1.bin"/>`,
      ),
      "xl/worksheets/sheet1.bin": concat([
        rec(Brt.RowHdr, u32(0)),
        rec(Brt.CellSt, concat([cellPrefix(0, 0), wstr("relative")])),
      ]),
    })
    expect((await readXlsb(pkg)).sheets[0].rows[0][0]).toBe("relative")
  })

  it("rejects a package whose workbook part is missing", async () => {
    const pkg = await buildXlsbPackage({
      "_rels/.rels": relsXml(
        `<Relationship Id="r" Type="${REL}/officeDocument" Target="xl/workbook.bin"/>`,
      ),
    })
    await expect(readXlsb(pkg)).rejects.toThrow(/missing workbook at xl\/workbook\.bin/)
  })

  it("surfaces a non-ZIP input as a ZipError", async () => {
    await expect(readXlsb(enc.encode("not a zip at all"))).rejects.toBeInstanceOf(ZipError)
  })

  it("refuses an encrypted container without a password", async () => {
    const container = writeCfb([
      { name: "EncryptionInfo", data: new Uint8Array(64) },
      { name: "EncryptedPackage", data: new Uint8Array(64) },
    ])
    await expect(readXlsb(container)).rejects.toBeInstanceOf(EncryptedFileError)
    await expect(readXlsb(container, { password: "pw" })).rejects.toBeInstanceOf(DecryptionError)
  })

  it("returns an empty sheet when the relationship does not resolve", async () => {
    const pkg = await buildXlsbPackage({
      "_rels/.rels": relsXml(
        `<Relationship Id="r" Type="${REL}/officeDocument" Target="xl/workbook.bin"/>`,
      ),
      "xl/workbook.bin": concat([
        rec(Brt.BundleSh, concat([u32(0), u32(0), nwstr("rIdGone"), wstr("Dangling")])),
        rec(Brt.BundleSh, concat([u32(0), u32(0), nwstr(null), wstr("NoRel")])),
      ]),
      "xl/_rels/workbook.bin.rels": relsXml(""),
    })
    const wb = await readXlsb(pkg)
    expect(wb.sheets).toEqual([
      { name: "Dangling", rows: [] },
      { name: "NoRel", rows: [] },
    ])
  })
})

describe("readXlsb — record decoding", () => {
  async function sheetPackage(
    sheetBin: Uint8Array,
    extra?: { sst?: Uint8Array; styles?: Uint8Array },
  ) {
    const wbRels = [
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.bin"/>`,
      extra?.sst
        ? `<Relationship Id="rId2" Type="${REL}/sharedStrings" Target="sharedStrings.bin"/>`
        : "",
      extra?.styles ? `<Relationship Id="rId3" Type="${REL}/styles" Target="styles.bin"/>` : "",
    ].join("")
    const parts: Record<string, Uint8Array | string> = {
      "_rels/.rels": relsXml(
        `<Relationship Id="r" Type="${REL}/officeDocument" Target="xl/workbook.bin"/>`,
      ),
      "xl/workbook.bin": rec(Brt.BundleSh, concat([u32(0), u32(0), nwstr("rId1"), wstr("Sheet1")])),
      "xl/_rels/workbook.bin.rels": relsXml(wbRels),
      "xl/worksheets/sheet1.bin": sheetBin,
    }
    if (extra?.sst) parts["xl/sharedStrings.bin"] = extra.sst
    if (extra?.styles) parts["xl/styles.bin"] = extra.styles
    return readXlsb(await buildXlsbPackage(parts))
  }

  it("skips records it has no decoder for, including blank cells", async () => {
    const wb = await sheetPackage(
      concat([
        rec(Brt.RowHdr, u32(0)),
        rec(Brt.CellBlank, cellPrefix(0, 0)),
        rec(999, [1, 2, 3]),
        rec(Brt.CellSt, concat([cellPrefix(1, 0), wstr("kept")])),
      ]),
    )
    expect(wb.sheets[0].rows[0]).toEqual([null, "kept"])
  })

  it("decodes an RK number stored as a truncated double", async () => {
    const wb = await sheetPackage(
      concat([
        rec(Brt.RowHdr, u32(0)),
        rec(Brt.CellRk, concat([cellPrefix(0, 0), u32(0x3ff00000)])),
      ]),
    )
    expect(wb.sheets[0].rows[0][0]).toBe(1)
  })

  it("falls back to #ERR! for an unknown error code", async () => {
    const wb = await sheetPackage(
      concat([rec(Brt.RowHdr, u32(0)), rec(Brt.CellError, concat([cellPrefix(0, 0), [0x55]]))]),
    )
    expect(wb.sheets[0].rows[0][0]).toBe("#ERR!")
  })

  it("falls back to an empty string for an out-of-range shared-string index", async () => {
    const wb = await sheetPackage(
      concat([rec(Brt.RowHdr, u32(0)), rec(Brt.CellIsst, concat([cellPrefix(0, 0), u32(42)]))]),
      { sst: rec(Brt.SSTItem, concat([[0], wstr("only")])) },
    )
    expect(wb.sheets[0].rows[0][0]).toBe("")
  })

  it("reads shared strings longer than a one-byte record size", async () => {
    // Record sizes are 7-bit varints; a 200-character string needs two bytes.
    const long = "s".repeat(200)
    const wb = await sheetPackage(
      concat([rec(Brt.RowHdr, u32(0)), rec(Brt.CellIsst, concat([cellPrefix(0, 0), u32(1)]))]),
      {
        sst: concat([
          rec(Brt.BeginSst, u32(2)),
          rec(Brt.SSTItem, concat([[0], wstr("short")])),
          rec(Brt.SSTItem, concat([[0], wstr(long)])),
        ]),
      },
    )
    expect(wb.sheets[0].rows[0][0]).toBe(long)
  })

  it("treats a cell as a date when a custom number format says so", async () => {
    // Only the XF records between BeginCellXFs/EndCellXFs are cell formats;
    // the ones outside describe named styles and must not shift the indices.
    const styles = concat([
      rec(Brt.Fmt, concat([u16(200), wstr("yyyy-mm-dd")])),
      rec(Brt.XF, concat([u16(0), u16(14), u16(0), u16(0), u16(0), [0, 0]])), // style xf, ignored
      rec(Brt.BeginCellXFs, u32(1)),
      rec(Brt.XF, concat([u16(0), u16(200), u16(0), u16(0), u16(0), [0, 0]])),
      rec(Brt.EndCellXFs, []),
      rec(Brt.XF, concat([u16(0), u16(14), u16(0), u16(0), u16(0), [0, 0]])), // also ignored
    ])
    const wb = await sheetPackage(
      concat([
        rec(Brt.RowHdr, u32(0)),
        rec(Brt.CellReal, concat([cellPrefix(0, 0), new Uint8Array(f64(45000))])),
      ]),
      { styles },
    )
    expect(wb.sheets[0].rows[0][0]).toBeInstanceOf(Date)
  })

  it("rejects a column index outside the supported sheet bounds", async () => {
    await expect(
      sheetPackage(
        concat([
          rec(Brt.RowHdr, u32(0)),
          rec(Brt.CellSt, concat([cellPrefix(20000, 0), wstr("x")])),
        ]),
      ),
    ).rejects.toThrow(/Cell column 20000 is outside the supported sheet bounds/)
  })

  it("rejects a row index outside the supported sheet bounds", async () => {
    await expect(sheetPackage(rec(Brt.RowHdr, u32(2_000_000)))).rejects.toThrow(
      /Cell row 2000000 is outside the supported sheet bounds/,
    )
  })
})

// ═══════════════════════════════════════════════════════════════════════
// CFB container
// ═══════════════════════════════════════════════════════════════════════

describe("readCfb / writeCfb", () => {
  it("rejects a file that is large enough but not a CFB", () => {
    expect(() => readCfb(new Uint8Array(1024))).toThrow(/CFB: bad signature/)
  })

  it("rejects a file too small to hold a header", () => {
    expect(() => readCfb(new Uint8Array(16))).toThrow(/CFB: file too small/)
  })

  it("round-trips a container with no streams at all", () => {
    expect(readCfb(writeCfb([])).size).toBe(0)
  })

  it("falls back to the standard 4096-byte mini cutoff when the header says 0", () => {
    // Some writers leave the mini-stream cutoff field zeroed; assuming 0
    // would route every small stream through the regular FAT and read
    // garbage.
    const file = writeCfb([{ name: "Small", data: enc.encode("tiny payload") }])
    new DataView(file.buffer, file.byteOffset, file.byteLength).setUint32(56, 0, true)
    expect(new TextDecoder().decode(readCfb(file).get("Small")!)).toBe("tiny payload")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Agile encryption
// ═══════════════════════════════════════════════════════════════════════

const FAST = { spinCount: 64 }

/** Rebuild an encrypted container with its EncryptionInfo XML rewritten. */
function rewriteEncryptionInfo(container: Uint8Array, edit: (xml: string) => string): Uint8Array {
  const streams = readCfb(container)
  const info = streams.get("EncryptionInfo")!
  const xml = edit(new TextDecoder().decode(info.subarray(8)))
  return writeCfb([
    { name: "EncryptionInfo", data: concat([info.subarray(0, 8), enc.encode(xml)]) },
    { name: "EncryptedPackage", data: streams.get("EncryptedPackage")! },
  ])
}

describe("decryptAgile — malformed containers", () => {
  it("rejects input that is not a CFB container", async () => {
    await expect(decryptAgile(enc.encode("plain text"), "pw")).rejects.toThrow(
      /Not a valid encrypted workbook container/,
    )
  })

  it("rejects a container missing the encryption streams", async () => {
    const container = writeCfb([{ name: "Workbook", data: new Uint8Array(64) }])
    await expect(decryptAgile(container, "pw")).rejects.toThrow(
      /missing EncryptionInfo\/EncryptedPackage/,
    )
  })

  it("rejects a non-Agile encryption version", async () => {
    // 3.2 is ECMA-376 Standard encryption (RC4/AES with a different layout).
    const info = new Uint8Array(8)
    const dv = new DataView(info.buffer)
    dv.setUint16(0, 3, true)
    dv.setUint16(2, 2, true)
    const container = writeCfb([
      { name: "EncryptionInfo", data: info },
      { name: "EncryptedPackage", data: new Uint8Array(64) },
    ])
    await expect(decryptAgile(container, "pw")).rejects.toThrow(
      /Unsupported encryption \(version 3\.2\)/,
    )
  })
})

describe("decryptAgile — EncryptionInfo XML", () => {
  async function encrypted(): Promise<Uint8Array> {
    return encryptAgile(enc.encode("PK" + "payload ".repeat(600)), "pw", FAST)
  }

  it("rejects EncryptionInfo with no password key encryptor", async () => {
    const bad = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(/<p:encryptedKey[^>]*\/>/, ""),
    )
    await expect(decryptAgile(bad, "pw")).rejects.toThrow(/missing encryptedKey/)
  })

  it("rejects EncryptionInfo with no keyData element", async () => {
    const bad = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(/<keyData[^>]*\/>/, ""),
    )
    await expect(decryptAgile(bad, "pw")).rejects.toThrow(/missing keyData/)
  })

  it("rejects a spinCount that is not a number", async () => {
    const bad = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(/spinCount="\d+"/, `spinCount="many"`),
    )
    await expect(decryptAgile(bad, "pw")).rejects.toThrow(/invalid spinCount/)
  })

  it("assumes SHA-512 for a hash name it does not recognise", async () => {
    // The hash name is only a lookup key and SHA-512 is the documented
    // default, so an unrecognised name must not crash the parse — this file
    // really is SHA-512 and still decrypts.
    const odd = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(
        /hashAlgorithm="SHA512" saltValue="([^"]*)" encryptedVerifierHashInput/,
        `hashAlgorithm="RIPEMD160" saltValue="$1" encryptedVerifierHashInput`,
      ),
    )
    expect(new TextDecoder().decode(await decryptAgile(odd, "pw"))).toContain("payload")
  })

  it("treats a missing verifier attribute as empty rather than crashing", async () => {
    const bad = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(/encryptedVerifierHashInput="[^"]*" /, ""),
    )
    await expect(decryptAgile(bad, "pw")).rejects.toThrow(DecryptionError)
  })
})

describe("decryptAgile — data integrity", () => {
  async function encrypted(): Promise<Uint8Array> {
    return encryptAgile(enc.encode("PK" + "payload ".repeat(600)), "pw", FAST)
  }

  it("skips the HMAC check when no dataIntegrity element is present", async () => {
    // Older writers omit it entirely; that is not an error.
    const without = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(/<dataIntegrity[^>]*\/>/, ""),
    )
    expect(new TextDecoder().decode(await decryptAgile(without, "pw"))).toContain("payload")
  })

  it("skips the HMAC check when the integrity attributes are empty", async () => {
    const empty = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(/encryptedHmacKey="[^"]*"/, `encryptedHmacKey=""`),
    )
    expect(new TextDecoder().decode(await decryptAgile(empty, "pw"))).toContain("payload")
  })

  it("skips the HMAC check when keyData uses a hash other than SHA-512", async () => {
    // Only SHA-512 HMAC is implemented, so the check is bypassed — the
    // package still decrypts rather than failing integrity.
    const other = rewriteEncryptionInfo(await encrypted(), (xml) =>
      xml.replace(/<keyData([^>]*)hashAlgorithm="SHA512"/, `<keyData$1hashAlgorithm="SHA384"`),
    )
    await expect(decryptAgile(other, "pw")).resolves.toBeInstanceOf(Uint8Array)
  })

  it("rejects a package whose ciphertext has been tampered with", async () => {
    const container = await encrypted()
    const streams = readCfb(container)
    const pkg = streams.get("EncryptedPackage")!.slice()
    pkg[pkg.length - 1] ^= 0xff
    const tampered = writeCfb([
      { name: "EncryptionInfo", data: streams.get("EncryptionInfo")! },
      { name: "EncryptedPackage", data: pkg },
    ])
    await expect(decryptAgile(tampered, "pw")).rejects.toThrow(/failed integrity check/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// OPC part parsers
// ═══════════════════════════════════════════════════════════════════════

describe("parseRelationships", () => {
  it("drops relationships that are missing a required attribute", () => {
    const xml = relsXml(
      `<Relationship Type="${REL}/worksheet" Target="sheet1.xml"/>` +
        `<Relationship Id="rId2" Target="sheet2.xml"/>` +
        `<Relationship Id="rId3" Type="${REL}/worksheet"/>` +
        `<Relationship Id="rId4" Type="${REL}/worksheet" Target="sheet4.xml"/>`,
    )
    expect(parseRelationships(xml)).toEqual([
      { id: "rId4", type: `${REL}/worksheet`, target: "sheet4.xml" },
    ])
  })

  it("keeps TargetMode for external relationships", () => {
    const xml = relsXml(
      `<Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://example.com" TargetMode="External"/>`,
    )
    expect(parseRelationships(xml)[0].targetMode).toBe("External")
  })

  it("ignores elements that are not relationships", () => {
    const xml = relsXml(
      `<Note>ignored</Note><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`,
    )
    expect(parseRelationships(xml).map((r) => r.id)).toEqual(["rId1"])
  })
})

describe("parseContentTypes", () => {
  const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

  it("drops Default and Override entries that are missing an attribute", () => {
    const xml =
      `<?xml version="1.0"?><Types xmlns="${CT_NS}">` +
      `<Default Extension="bin"/>` +
      `<Default ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml"/>` +
      `<Override ContentType="application/xml"/>` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/styles"/>` +
      `</Types>`
    const ct = parseContentTypes(xml)
    expect([...ct.defaults.keys()]).toEqual(["rels"])
    expect([...ct.overrides.keys()]).toEqual(["/xl/styles.xml"])
  })

  it("ignores elements that are neither Default nor Override", () => {
    const xml = `<?xml version="1.0"?><Types xmlns="${CT_NS}"><Comment>hi</Comment></Types>`
    const ct = parseContentTypes(xml)
    expect(ct.defaults.size).toBe(0)
    expect(ct.overrides.size).toBe(0)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// fetchCsv
// ═══════════════════════════════════════════════════════════════════════

describe("fetchCsv", () => {
  // A `data:` URL is a real fetch with a real Response, so this exercises
  // the function end to end without depending on a network peer.
  const url = (csv: string) => `data:text/csv,${encodeURIComponent(csv)}`

  it("parses a CSV fetched from a URL", async () => {
    expect(await fetchCsv(url("Name,Score\nAda,95\n"))).toEqual([
      ["Name", "Score"],
      ["Ada", "95"],
    ])
  })

  it("passes reader options through to the parser", async () => {
    expect(
      await fetchCsv(url("Name;Score\nAda;95\n"), { delimiter: ";", typeInference: true }),
    ).toEqual([
      ["Name", "Score"],
      ["Ada", 95],
    ])
  })

  // The failure path is what a caller actually has to handle, and the status
  // code is the only diagnostic it gets. No URL scheme Node's fetch accepts
  // can produce a non-2xx `Response` without a live peer, so this is the one
  // place the global is stubbed — with a *real* `Response`, so everything
  // after the `ok` check is still the genuine WHATWG object.
  it("reports the HTTP status when the server refuses the request", async () => {
    vi.stubGlobal("fetch", async () => new Response("nope", { status: 503 }))
    try {
      await expect(fetchCsv("https://example.invalid/data.csv")).rejects.toThrow(
        "Failed to fetch: 503",
      )
    } finally {
      vi.unstubAllGlobals()
    }
  })
})
