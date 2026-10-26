import { describe, expect, it } from "vitest"
import { writeCfb } from "../src/xlsx/crypto/cfb"
import { readXls } from "../src/xls/reader"
import { streamXlsRows } from "../src/xls/stream-reader"
import { read } from "../src/defter"
import { ParseError } from "../src/errors"
import { expectStreamMatchesReadXls, streamed } from "./_xls-stream"

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
  CONTINUE: 0x003c,
  DATEMODE: 0x0022,
  FORMAT: 0x041e,
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
  BLANK: 0x0201,
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

function buildXls(
  opts: { dateFmtId?: number; fmtCodes?: Array<[number, string]>; value?: number } = {},
): Uint8Array {
  const strings = ["Name", "Score", "Ada"]

  const sheet = concat([
    bof(0x0010),
    record(SID.LABELSST, [...u16(0), ...u16(0), ...u16(0), ...u32(0)]),
    record(SID.LABELSST, [...u16(0), ...u16(1), ...u16(0), ...u32(1)]),
    record(SID.LABELSST, [...u16(1), ...u16(0), ...u16(0), ...u32(2)]),
    record(SID.RK, [...u16(1), ...u16(1), ...u16(0), ...rkInt(95)]),
    record(SID.NUMBER, [...u16(1), ...u16(2), ...u16(0), ...f64(3.14)]),
    record(SID.NUMBER, [...u16(1), ...u16(3), ...u16(1), ...f64(opts.value ?? 45000)]), // date xf
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
      // FORMAT records — the workbook's own number-format definitions,
      // which may redefine a built-in id. See #568.
      ...(opts.fmtCodes ?? []).map(([id, code]) =>
        record(SID.FORMAT, [...u16(id), ...xlStr(code)]),
      ),
      record(SID.XF, [...u16(0), ...u16(0), ...Array.from({ length: 16 }, () => 0)]), // general
      // The date xf's built-in format id is a parameter: the built-in date
      // set is wider than the familiar 14-22 block. See the CJK case below.
      record(SID.XF, [
        ...u16(0),
        ...u16(opts.dateFmtId ?? 14),
        ...Array.from({ length: 16 }, () => 0),
      ]),
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
    expect(rows[3][0]).toBe(10)
    expect(rows[3][1]).toBe(20)
    expect(wb.sheets[0].merges).toEqual([{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }])
  })

  // ── #439: the built-in date set is wider than 14-22 / 45-47 ──────────
  // Built-ins 27-36 (CJK) and 50-58 (Thai/Chinese/Korean) are date and
  // time formats, and they carry no FORMAT record — so a reader that does
  // not know them falls through to "not a date" and hands back the raw
  // serial. This reader used to keep a 12-entry table of its own.
  describe("built-in date format ids outside the familiar block", () => {
    const CJK_AND_EXTENDED = [
      27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58,
    ]

    for (const id of CJK_AND_EXTENDED) {
      it(`reads a cell styled with built-in format ${id} as a Date`, async () => {
        const wb = await readXls(buildXls({ dateFmtId: id }))

        expect(wb.sheets[0].rows[1][3]).toBeInstanceOf(Date)
      })
    }

    it("still treats a non-date built-in as a number", async () => {
      const wb = await readXls(buildXls({ dateFmtId: 3 }))

      expect(wb.sheets[0].rows[1][3]).toBe(45000)
    })
  })

  // ── #568: a FORMAT record redefining a built-in id ──────────────────
  // The built-in table was consulted first, so a workbook's own FORMAT
  // record for that id was never read. readXlsx has always resolved it
  // the other way round; these are the cases where the three disagreed.
  describe("a FORMAT record redefines a built-in id", () => {
    it("reads a number when the file redefines a built-in date id numerically", async () => {
      // 1C allocates its custom formats from id 50 — inside the
      // Thai/Chinese/Korean date block — and '000000000000' is its mask
      // for a leading-zero barcode. Read as a date serial this value is
      // out of range, so the cell came back Invalid Date and the barcode
      // was unrecoverable through the API.
      const wb = await readXls(
        buildXls({ dateFmtId: 50, fmtCodes: [[50, "000000000000"]], value: 81227827687 }),
      )

      expect(wb.sheets[0].rows[1][3]).toBe(81227827687)
    })

    it("reads a number when the file redefines id 14 as '#,##0'", async () => {
      // The xlsx reader's own note names this case: redefining 14 made it
      // resolve to a numeric format *and* report as a date, so the serial
      // was converted and then formatted numerically.
      const wb = await readXls(buildXls({ dateFmtId: 14, fmtCodes: [[14, "#,##0"]] }))

      expect(wb.sheets[0].rows[1][3]).toBe(45000)
    })

    it("reads a Date when the file redefines a numeric built-in id as a date", async () => {
      const wb = await readXls(buildXls({ dateFmtId: 3, fmtCodes: [[3, "yyyy-mm-dd"]] }))

      expect(wb.sheets[0].rows[1][3]).toBeInstanceOf(Date)
    })

    it("keeps the built-in meaning when the file redefines some other id", async () => {
      const wb = await readXls(buildXls({ dateFmtId: 50, fmtCodes: [[164, "000000000000"]] }))

      expect(wb.sheets[0].rows[1][3]).toBeInstanceOf(Date)
    })
  })

  it("is auto-detected by read()", async () => {
    const wb = await read(buildXls())
    expect(wb.sheets[0].rows[1][0]).toBe("Ada")
    expect(wb.sheets[0].rows[1][1]).toBe(95)
  })

  it("rejects a BIFF version it does not read with a clear error", async () => {
    const bof3 = record(SID.BOF, [
      ...u16(0x0300),
      ...u16(0x0005),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(0),
    ])
    const stream = concat([bof3, eof()])
    const data = writeCfb([{ name: "Workbook", data: stream }])
    await expect(readXls(data)).rejects.toThrow(/BIFF 0x300/)
  })
})

// ── Strings decoded a run at a time ──────────────────────────────────
// Text is decoded a record's worth of characters at a time, not one by
// one. Excel splits a long shared string wherever its record fills, and
// each CONTINUE that resumes one restarts with its own flag byte, so a
// string can change width mid-way and a surrogate pair can be cut in two.
// These pass on main as well: they pin what the faster path must keep.

const latin = (s: string): number[] => [...s].map((c) => c.charCodeAt(0))
function utf16(s: string): number[] {
  const out: number[] = []
  for (let i = 0; i < s.length; i++) out.push(...u16(s.charCodeAt(i)))
  return out
}
const generalXf = (): number[] =>
  record(SID.XF, [...u16(0), ...u16(0), ...Array.from({ length: 16 }, () => 0)])

/** One sheet after the given globals records, its BOUNDSHEET pointing at it. */
function oneSheetXls(globals: number[][], sheet: number[][], sheetName = "Sheet1"): Uint8Array {
  const makeGlobals = (sheetPos: number): Uint8Array =>
    concat([
      bof(0x0005),
      ...globals,
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr(sheetName)]),
      eof(),
    ])
  const globalsStream = makeGlobals(makeGlobals(0).length)
  return writeCfb([
    { name: "Workbook", data: concat([globalsStream, bof(0x0010), ...sheet, eof()]) },
  ])
}

describe("XLS (BIFF8) strings", () => {
  it("joins a shared string's runs across CONTINUE records, whatever their width", async () => {
    // 100 one-byte characters in the SST record, a CONTINUE record full of
    // them (more than one String.fromCharCode call takes), then a switch
    // to UTF-16 for the rest. The next string's surrogate pair straddles
    // the following boundary.
    const head = "A".repeat(100)
    const middle = "B".repeat(8223)
    const tail = "Ж".repeat(1677)
    const emoji = "ab😀cd"
    const xls = oneSheetXls(
      [
        generalXf(),
        record(SID.SST, [...u32(3), ...u32(3), ...u16(10_000), 0, ...latin(head)]),
        record(SID.CONTINUE, [0, ...latin(middle)]),
        record(SID.CONTINUE, [
          1,
          ...utf16(tail),
          ...u16(emoji.length),
          1,
          ...utf16(emoji.slice(0, 3)),
        ]),
        record(SID.CONTINUE, [1, ...utf16(emoji.slice(3)), ...u16(3), 0, ...latin("end")]),
      ],
      [0, 1, 2].map((i) => record(SID.LABELSST, [...u16(0), ...u16(i), ...u16(0), ...u32(i)])),
    )

    const wb = await readXls(xls)

    expect(wb.sheets[0].rows[0]).toEqual([head + middle + tail, emoji, "end"])
  })

  it("reads inline text in either width, longer than one String.fromCharCode call takes", async () => {
    const long = "x".repeat(8200)
    const xls = oneSheetXls(
      [generalXf()],
      [
        record(SID.LABEL, [...u16(0), ...u16(0), ...u16(0), ...xlStr(long)]),
        record(SID.LABEL, [...u16(0), ...u16(1), ...u16(0), ...u16(5), 1, ...utf16("Ωmega")]),
      ],
    )

    const wb = await readXls(xls)

    expect(wb.sheets[0].rows[0]).toEqual([long, "Ωmega"])
  })
})

describe("XLS (BIFF8) sheet offsets", () => {
  it("reads a sheet whose BOUNDSHEET offset misses a record boundary as empty", async () => {
    // The offset is the only link from a BOUNDSHEET to its sheet. Records
    // are walked in place now rather than indexed by offset, and one that
    // lands inside a record still means no rows, not records out of phase.
    const makeGlobals = (sheetPos: number): Uint8Array =>
      concat([
        bof(0x0005),
        generalXf(),
        record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr("Good")]),
        record(SID.BOUNDSHEET, [...u32(sheetPos + 2), 0, 0, ...shortStr("Off")]),
        eof(),
      ])
    const globals = makeGlobals(makeGlobals(0).length)
    const sheet = concat([
      bof(0x0010),
      record(SID.LABEL, [...u16(0), ...u16(0), ...u16(0), ...xlStr("Hi")]),
      eof(),
    ])

    const wb = await readXls(writeCfb([{ name: "Workbook", data: concat([globals, sheet]) }]))

    expect(wb.sheets.map((s) => s.name)).toEqual(["Good", "Off"])
    expect(wb.sheets[0].rows).toEqual([["Hi"]])
    expect(wb.sheets[1].rows).toEqual([])
  })
})

// ── Minimal BIFF5 .xls builder (test-only) ───────────────────────────
// Excel 5.0/95: the same records, but every string is bytes in the
// workbook's code page with no per-string flag, cell text is written
// inline (LABEL / RSTRING — there is no SST), and the stream is "Book".

/** Windows-1251 bytes for the Cyrillic and ASCII this file needs. */
function cp1251(s: string): number[] {
  const extras: Record<string, number> = {
    Є: 0xaa,
    І: 0xb2,
    Ї: 0xaf,
    є: 0xba,
    і: 0xb3,
    ї: 0xbf,
  }
  return [...s].map((ch) => {
    const code = ch.charCodeAt(0)
    if (code < 0x80) return code
    if (code >= 0x410 && code <= 0x44f) return 0xc0 + (code - 0x410)
    const extra = extras[ch]
    if (extra === undefined) throw new Error(`no cp1251 byte for ${ch}`)
    return extra
  })
}
const bytesStr = (bytes: number[]): number[] => [...u16(bytes.length), ...bytes]
const shortBytesStr = (bytes: number[]): number[] => [bytes.length, ...bytes]

const SID5 = { CODEPAGE: 0x0042, RSTRING: 0x00d6, STRING: 0x0207 }

const bof5 = (dt: number): number[] =>
  record(SID.BOF, [...u16(0x0500), ...u16(dt), ...u16(0), ...u16(0)])

function buildXls5(
  opts: {
    codepage?: number
    /** Raw CODEPAGE payload, for a malformed record. */
    codepageRecord?: number[]
    text?: number[]
    /** The count the first LABEL claims, when it should lie about its bytes. */
    textCount?: number
    /** The RSTRING trailer: a u8 run count, then (first char, font) byte pairs. */
    richRuns?: number[]
    stream?: string
  } = {},
): Uint8Array {
  const text = opts.text ?? cp1251("Видаткова накладна")
  const richRuns = opts.richRuns ?? [1, 0, 0]
  const sheet = concat([
    bof5(0x0010),
    record(SID.LABEL, [
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(opts.textCount ?? text.length),
      ...text,
    ]),
    // Rich text: the string, then the runs (u8 count, u8 first char + u8 font each).
    record(SID5.RSTRING, [
      ...u16(0),
      ...u16(1),
      ...u16(0),
      ...bytesStr(cp1251("Ціна")),
      ...richRuns,
    ]),
    record(SID.RK, [...u16(1), ...u16(0), ...u16(0), ...rkInt(95)]),
    record(SID.NUMBER, [...u16(1), ...u16(1), ...u16(1), ...f64(45000)]), // date xf
    record(SID.NUMBER, [...u16(1), ...u16(2), ...u16(2), ...f64(481227827687)]), // redefined id 50
    // A string formula: the cached value rides the STRING record after it.
    record(SID.FORMULA, [
      ...u16(2),
      ...u16(0),
      ...u16(0),
      0,
      0,
      0,
      0,
      0,
      0,
      0xff,
      0xff,
      ...u16(0),
      ...u32(0),
      ...u16(0),
    ]),
    record(SID5.STRING, bytesStr(cp1251("Разом"))),
    record(SID.BOOLERR, [...u16(2), ...u16(1), ...u16(0), 1, 0]),
    record(SID.MULRK, [
      ...u16(3),
      ...u16(0),
      ...u16(0),
      ...rkInt(10),
      ...u16(0),
      ...rkInt(20),
      ...u16(1),
    ]),
    // Written by 1C into BIFF5 files although the record is BIFF8's.
    record(SID.MERGECELLS, [...u16(1), ...u16(0), ...u16(0), ...u16(0), ...u16(1)]),
    eof(),
  ])

  const makeGlobals = (sheetPos: number): Uint8Array =>
    concat([
      bof5(0x0005),
      ...(opts.codepageRecord
        ? [record(SID5.CODEPAGE, opts.codepageRecord)]
        : opts.codepage === undefined
          ? []
          : [record(SID5.CODEPAGE, u16(opts.codepage))]),
      record(SID.DATEMODE, u16(0)),
      // BIFF5 FORMAT: u16 id + u8-counted byte string — parsed as BIFF8's
      // u16-counted Unicode string this would read as neither id nor code.
      record(SID.FORMAT, [...u16(50), ...shortBytesStr(cp1251("000000000000"))]),
      record(SID.XF, [...u16(0), ...u16(0), ...Array.from({ length: 12 }, () => 0)]),
      record(SID.XF, [...u16(0), ...u16(14), ...Array.from({ length: 12 }, () => 0)]),
      record(SID.XF, [...u16(0), ...u16(50), ...Array.from({ length: 12 }, () => 0)]),
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortBytesStr(cp1251("Лист1"))]),
      eof(),
    ])

  const globalsLen = makeGlobals(0).length
  const globals = makeGlobals(globalsLen)
  return writeCfb([{ name: opts.stream ?? "Book", data: concat([globals, sheet]) }])
}

describe("XLS (BIFF5) reader", () => {
  it("reads sheet names, cell text, format codes and values through the CODEPAGE record", async () => {
    const wb = await readXls(buildXls5({ codepage: 1251 }))

    expect(wb.sheets[0].name).toBe("Лист1")
    const rows = wb.sheets[0].rows
    expect(rows[0]).toEqual(["Видаткова накладна", "Ціна", null])
    expect(rows[1][0]).toBe(95)
    expect(rows[1][1]).toBeInstanceOf(Date)
    // The BIFF5 FORMAT record was read, so id 50 is the barcode mask, not a date.
    expect(rows[1][2]).toBe(481227827687)
    expect(rows[2]).toEqual(["Разом", true, null])
    expect(rows[3]).toEqual([10, 20, null])
    expect(wb.sheets[0].merges).toEqual([{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }])
  })

  it("falls back to `codepage` when the workbook carries no CODEPAGE record", async () => {
    // 1C exports write Cyrillic text and no CODEPAGE record at all.
    const wb = await readXls(buildXls5(), { codepage: 1251 })

    expect(wb.sheets[0].name).toBe("Лист1")
    expect(wb.sheets[0].rows[0][0]).toBe("Видаткова накладна")
  })

  it("assumes Windows-1252 when neither the file nor the caller names a page", async () => {
    // 0x80 is € in Windows-1252 and a control character in Latin-1, so a
    // byte-for-code-point fallback cannot pass this by accident.
    const wb = await readXls(buildXls5({ text: [0x80, 0x31, 0x32] }))

    expect(wb.sheets[0].rows[0][0]).toBe("€12")
  })

  it("refuses a string that claims more bytes than its record holds", async () => {
    // `subarray` clamps, so without the check this would read back as a
    // silently shortened string rather than an error.
    await expect(readXls(buildXls5({ codepage: 1251, textCount: 500 }))).rejects.toThrow(
      /claims 500 bytes/,
    )
  })

  it("refuses an RSTRING whose formatting runs stop short of the count", async () => {
    await expect(readXls(buildXls5({ codepage: 1251, richRuns: [3, 0, 0] }))).rejects.toThrow(
      /claims 3 formatting runs/,
    )
    await expect(readXls(buildXls5({ codepage: 1251, richRuns: [] }))).rejects.toThrow(
      /no run count/,
    )
  })

  it("refuses a CODEPAGE record too short to name a page", async () => {
    await expect(readXls(buildXls5({ codepageRecord: [0xe3] }))).rejects.toThrow(
      /CODEPAGE record is too short/,
    )
  })

  it("lets the file's CODEPAGE record win over the option", async () => {
    const wb = await readXls(buildXls5({ codepage: 1251 }), { codepage: 1252 })

    expect(wb.sheets[0].rows[0][0]).toBe("Видаткова накладна")
  })

  it("refuses a code page it cannot decode rather than guessing", async () => {
    await expect(readXls(buildXls5(), { codepage: 437 })).rejects.toThrow(/codepage 437/)
    await expect(readXls(buildXls5({ codepage: 437 }))).rejects.toThrow(/codepage 437/)
  })

  it("is auto-detected by read() through the Book stream", async () => {
    const wb = await read(buildXls5({ codepage: 1251 }))

    expect(wb.sheets[0].rows[0][0]).toBe("Видаткова накладна")
  })
})

// ── streamXlsRows ────────────────────────────────────────────────────
// The stream decodes values with readXls's own code, so most of what
// follows is about shape: which sheet, which rows, in what order, and
// when it stops reading.

const label = (row: number, col: number, text: string): number[] =>
  record(SID.LABEL, [...u16(row), ...u16(col), ...u16(0), ...xlStr(text)])

/** Two sheets, each holding its own name in A1. */
function twoSheetXls(): Uint8Array {
  const sheets = ["First", "Second"].map((name) => concat([bof(0x0010), label(0, 0, name), eof()]))
  const makeGlobals = (firstPos: number): Uint8Array =>
    concat([
      bof(0x0005),
      generalXf(),
      record(SID.BOUNDSHEET, [...u32(firstPos), 0, 0, ...shortStr("First")]),
      record(SID.BOUNDSHEET, [...u32(firstPos + sheets[0].length), 0, 0, ...shortStr("Second")]),
      eof(),
    ])
  const globals = makeGlobals(makeGlobals(0).length)
  return writeCfb([{ name: "Workbook", data: concat([globals, ...sheets]) }])
}

/** Rows 0–4, the cell in column c of row r reading `${r}${c}`; rows 1 and 2 four columns wide. */
const gridRecords = (): number[][] => [
  label(0, 0, "00"),
  ...[0, 1, 2, 3].map((c) => label(1, c, `1${c}`)),
  ...[0, 1, 2, 3].map((c) => label(2, c, `2${c}`)),
  label(3, 0, "30"),
  label(4, 0, "40"),
]

/** A LABEL record cut off after its row — malformed wherever it sits. */
const truncatedLabel = (row: number): number[] => record(SID.LABEL, u16(row))

describe("streamXlsRows", () => {
  describe("streams what readXls reads", () => {
    it("BIFF8: SST labels, RK, MULRK, numbers, bools, errors and dates", async () => {
      await expectStreamMatchesReadXls(buildXls())
    })

    it("BIFF8: a FORMAT record that redefines a built-in id", async () => {
      await expectStreamMatchesReadXls(
        buildXls({ dateFmtId: 50, fmtCodes: [[50, "000000000000"]], value: 81227827687 }),
      )
    })

    it("BIFF8: the 1904 date system, when asked for", async () => {
      await expectStreamMatchesReadXls(buildXls(), { dateSystem: "1904" })
    })

    it("BIFF5: code-page text, RSTRING, and a string formula's STRING record", async () => {
      await expectStreamMatchesReadXls(buildXls5({ codepage: 1251 }))
      await expectStreamMatchesReadXls(buildXls5(), { codepage: 1251 })
    })
  })

  it("streams the first sheet, or the one named by index or by name", async () => {
    const xls = twoSheetXls()

    expect(await streamed(xls)).toEqual([{ index: 0, values: ["First"] }])
    expect(await streamed(xls, { sheet: 1 })).toEqual([{ index: 0, values: ["Second"] }])
    expect(await streamed(xls, { sheet: "Second" })).toEqual([{ index: 0, values: ["Second"] }])
  })

  it("yields nothing for a sheet that is not there", async () => {
    const xls = twoSheetXls()

    expect(await streamed(xls, { sheet: 2 })).toEqual([])
    expect(await streamed(xls, { sheet: "Third" })).toEqual([])
  })

  it("skips an empty row and keeps each row's own index", async () => {
    // Row 1 carries only a BLANK — formatting, no value — and row 2 nothing at all.
    const xls = oneSheetXls(
      [generalXf()],
      [label(0, 0, "a"), record(SID.BLANK, [...u16(1), ...u16(0), ...u16(0)]), label(3, 2, "d")],
    )

    expect(await streamed(xls)).toEqual([
      { index: 0, values: ["a"] },
      { index: 3, values: [null, null, "d"] },
    ])
  })

  it("applies range as streamXlsxRows does: rows outside skipped, columns outside masked", async () => {
    const xls = oneSheetXls([generalXf()], gridRecords())

    expect(await streamed(xls, { range: "B2:C3" })).toEqual([
      { index: 1, values: [null, "11", "12", null] },
      { index: 2, values: [null, "21", "22", null] },
    ])
  })

  it("stops reading once range or maxRows can emit nothing more", async () => {
    // Everything after row 4 is damaged: reading that far is an error, so
    // a read that succeeds is one that stopped before it.
    const xls = oneSheetXls([generalXf()], [...gridRecords(), truncatedLabel(5)])

    await expect(streamed(xls)).rejects.toThrow(ParseError)
    expect((await streamed(xls, { range: "A1:D3" })).map((row) => row.index)).toEqual([0, 1, 2])
    expect((await streamed(xls, { maxRows: 2 })).map((row) => row.index)).toEqual([0, 1])
  })

  it("buffers a sheet whose cells go back a row, and yields each row once, in order", async () => {
    // Streaming such a sheet would have to yield row 0 twice, or split it.
    const xls = oneSheetXls(
      [generalXf()],
      [label(2, 0, "c"), label(0, 0, "a"), label(1, 0, "b"), label(0, 1, "a2")],
    )

    expect(await streamed(xls)).toEqual([
      { index: 0, values: ["a", "a2"] },
      { index: 1, values: ["b"] },
      { index: 2, values: ["c"] },
    ])
    await expectStreamMatchesReadXls(xls)
  })

  it("bounds a buffered sheet by maxTotalCells as readXls bounds it, and a streamed one not at all", async () => {
    const backwards = oneSheetXls([generalXf()], [label(2, 1, "c"), label(0, 0, "a")])
    const forwards = oneSheetXls([generalXf()], [label(0, 0, "a"), label(2, 1, "c")])

    await expect(readXls(backwards, { maxTotalCells: 4 })).rejects.toThrow(/maxTotalCells/)
    await expect(streamed(backwards, { maxTotalCells: 4 })).rejects.toThrow(/maxTotalCells/)
    // Nothing is held while streaming, so there is nothing for the bound to bound.
    expect(await streamed(forwards, { maxTotalCells: 4 })).toHaveLength(2)
  })

  it("reports damage as a ParseError, before the first row and mid-sheet alike", async () => {
    await expect(streamed(new Uint8Array([1, 2, 3, 4]))).rejects.toThrow(ParseError)
    const xls = oneSheetXls([generalXf()], [label(0, 0, "a"), truncatedLabel(1)])
    await expect(streamed(xls)).rejects.toThrow(ParseError)
  })

  it("refuses a column past the sheet's bounds, as readXls does", async () => {
    const xls = oneSheetXls([generalXf()], [label(0, 20_000, "far")])

    await expect(readXls(xls)).rejects.toThrow(/outside the supported sheet bounds/)
    await expect(streamed(xls)).rejects.toThrow(/outside the supported sheet bounds/)
  })

  it("takes a ReadableStream, as the other readers do, bounded by maxInputBytes", async () => {
    const bytes = buildXls()
    const input = (): ReadableStream<Uint8Array> =>
      new ReadableStream({
        start(controller) {
          controller.enqueue(bytes)
          controller.close()
        },
      })
    const rows = []
    for await (const row of streamXlsRows(input())) rows.push(row)

    expect(rows).toEqual(await streamed(bytes))
    await expect(readXls(input(), { maxInputBytes: 64 })).rejects.toThrow(ParseError)
    await expect(streamXlsRows(input(), { maxInputBytes: 64 }).next()).rejects.toThrow(ParseError)
  })

  it("can be abandoned after its first row", async () => {
    const rows = streamXlsRows(buildXls())

    expect((await rows.next()).value).toEqual({ index: 0, values: ["Name", "Score"] })
    expect(await rows.return(undefined)).toEqual({ done: true, value: undefined })
    expect(await rows.next()).toEqual({ done: true, value: undefined })
  })
})
