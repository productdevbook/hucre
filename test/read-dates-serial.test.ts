import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import type { ReadOptions } from "../src/_types"
import { isDateFormat } from "../src/_date"
import { read } from "../src/defter"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"
import { readXls } from "../src/xls/reader"
import { readXlsb } from "../src/xlsx/xlsb/reader"
import { readXlsx } from "../src/xlsx/reader"
import { readXlsxObjects } from "../src/xlsx/objects"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { writeXlsx } from "../src/xlsx/writer"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// ═══════════════════════════════════════════════════════════════════════
// #600 — a numeric cell under a date format always read as a `Date`.
// That is the right default and stays the default, but it cannot be
// undone for a caller who needs the number the file holds:
//
// - the conversion rounds to the millisecond, so 45306.123456789 comes
//   back through `dateToSerial` as 45306.123456794;
// - serial 60, the Lotus phantom 29 February 1900, reads as the same
//   `Date` as 59 (docs/PARITY.md, "Serial 60 collapses onto 59");
// - a time-only format yields a `Date` on the epoch day, when what the
//   cell holds is a fraction of a day.
//
// `dates: "serial"` returns the number, with `type: "number"`. The
// number format is still on the style, so `isDateFormat` can say which
// numbers were dates.
// ═══════════════════════════════════════════════════════════════════════

const SERIALS = [45306, 45306.5, 45306.123456789, 59, 60, 61]

/** One column of serials, every one under a date-time format. */
const datedColumn = (serials: number[], numFmt = "yyyy-mm-dd hh:mm:ss"): Promise<Uint8Array> =>
  writeXlsx({
    sheets: [{ name: "S", rows: serials.map((value) => [{ value, style: { numFmt } }]) }],
  })

const column = (rows: unknown[][]): unknown[] => rows.map((r) => r[0])

const collect = async <T>(gen: AsyncIterable<T>): Promise<T[]> => {
  const out: T[] = []
  for await (const item of gen) out.push(item)
  return out
}

const asStream = (bytes: Uint8Array): ReadableStream<Uint8Array> =>
  new ReadableStream({
    start(controller) {
      controller.enqueue(bytes)
      controller.close()
    },
  })

const enc = new TextEncoder()
const dec = new TextDecoder()

/** A workbook whose row 1 is raw `<c>` XML of our choosing (see iso-date-cells.test.ts). */
async function withCells(cellXml: string, dateSystem?: "1904"): Promise<Uint8Array> {
  // Two styled cells so that styles.xml carries a date xf at s="1".
  const base = await writeXlsx({
    sheets: [{ name: "S", rows: [[1, { value: 1, style: { numFmt: "yyyy-mm-dd" } }]] }],
    dateSystem,
  })
  const all = await new ZipReader(base).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? enc.encode(
            dec
              .decode(data)
              .replace(
                /<sheetData>.*<\/sheetData>/,
                `<sheetData><row r="1">${cellXml}</row></sheetData>`,
              ),
          )
        : data,
    )
  }
  return zw.build()
}

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)))

const SERIAL: ReadOptions = { dates: "serial" }

describe("readXlsx", () => {
  it("still reads a date-formatted number as a Date by default", async () => {
    const values = column((await readXlsx(await datedColumn(SERIALS))).sheets[0]!.rows)

    for (const v of values) expect(v).toBeInstanceOf(Date)
    expect((values[0] as Date).toISOString()).toBe("2024-01-15T00:00:00.000Z")
  })

  it('reads the same with dates: "date" as with no option', async () => {
    const bytes = await datedColumn(SERIALS)

    expect((await readXlsx(bytes, { dates: "date" })).sheets[0]!.rows).toEqual(
      (await readXlsx(bytes)).sheets[0]!.rows,
    )
  })

  it('returns the serial the file holds with dates: "serial"', async () => {
    const values = column((await readXlsx(await datedColumn(SERIALS), SERIAL)).sheets[0]!.rows)

    // Exactly, not to the millisecond: 45306.123456789 survives.
    expect(values).toEqual(SERIALS)
  })

  it("keeps serial 60 apart from 59", async () => {
    const values = column((await readXlsx(await datedColumn([59, 60]), SERIAL)).sheets[0]!.rows)

    expect(values).toEqual([59, 60])
  })

  it("returns a time-only cell as the fraction of a day it is", async () => {
    const time = (13 * 3600 + 45 * 60 + 30) / 86400
    const values = column(
      (await readXlsx(await datedColumn([time], "hh:mm:ss"), SERIAL)).sheets[0]!.rows,
    )

    expect(values).toEqual([time])
  })

  it('types the cell "number" and keeps its date format on the style', async () => {
    const wb = await readXlsx(await datedColumn([45306.5]), { ...SERIAL, readStyles: true })
    const cell = wb.sheets[0]!.cells!.get("0,0")!

    expect(cell.type).toBe("number")
    expect(cell.value).toBe(45306.5)
    expect(cell.style?.numFmt).toBe("yyyy-mm-dd hh:mm:ss")
    expect(isDateFormat(cell.style!.numFmt!)).toBe(true)
  })

  it("returns the serial in sparse mode too", async () => {
    const wb = await readXlsx(await datedColumn([60]), { ...SERIAL, sparse: true })

    expect(wb.sheets[0]!.cells!.get("0,0")!.value).toBe(60)
  })

  it("keeps a formula's cached date result a number", async () => {
    const bytes = await withCells('<c r="A1" s="1"><f>TODAY()</f><v>45306.5</v></c>')

    const dated = (await readXlsx(bytes)).sheets[0]!.cells!.get("0,0")!
    expect(dated.formulaResult).toBeInstanceOf(Date)

    const wb = await readXlsx(bytes, SERIAL)
    const cell = wb.sheets[0]!.cells!.get("0,0")!
    expect(wb.sheets[0]!.rows[0]![0]).toBe(45306.5)
    expect(cell.type).toBe("formula")
    expect(cell.formula).toBe("TODAY()")
    expect(cell.formulaResult).toBe(45306.5)
  })

  // A `t="d"` cell holds ISO text, not a serial. Converting it means every
  // date in a workbook reads the same way under this option, whichever
  // way the producer chose to store it.
  it('converts a t="d" ISO cell to the serial for the same day', async () => {
    const bytes = await withCells(
      '<c r="A1" s="1" t="d"><v>2024-03-17</v></c><c r="B1" t="d"><v>2024-03-17T12:00:00Z</v></c>',
    )

    const wb = await readXlsx(bytes, { ...SERIAL, readStyles: true })

    expect(wb.sheets[0]!.rows[0]).toEqual([45368, 45368.5])
    expect(wb.sheets[0]!.cells!.get("0,0")!.type).toBe("number")
  })

  it('converts a t="d" cell against the 1904 epoch in a 1904 workbook', async () => {
    const bytes = await withCells('<c r="A1" t="d"><v>2024-03-17</v></c>', "1904")

    // 45368 - 1462: the same day, counted from 1 January 1904.
    expect((await readXlsx(bytes, SERIAL)).sheets[0]!.rows[0]).toEqual([43906])
  })

  it('leaves a t="d" bare time as text, as the default mode does', async () => {
    const bytes = await withCells('<c r="A1" t="d"><v>13:45:30</v></c>')

    expect((await readXlsx(bytes, SERIAL)).sheets[0]!.rows[0]).toEqual(["13:45:30"])
  })

  it("returns the serial of a 1904 workbook unconverted", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[{ value: 43906, style: { numFmt: "yyyy-mm-dd" } }]] }],
      dateSystem: "1904",
    })

    expect((await readXlsx(bytes, SERIAL)).sheets[0]!.rows[0]).toEqual([43906])
    expect(((await readXlsx(bytes)).sheets[0]!.rows[0]![0] as Date).toISOString()).toBe(
      "2024-03-17T00:00:00.000Z",
    )
  })

  it("is forwarded by readXlsxObjects", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [["when"], [{ value: 60, style: { numFmt: "yyyy-mm-dd" } }]] }],
    })

    expect((await readXlsxObjects(bytes, SERIAL)).data).toEqual([{ when: 60 }])
  })
})

// saveXlsx writes a number back as a number, so the date format has to
// come from the read style — hence readStyles. Without it a Date gets the
// writer's default date format and a serial gets none; docs/PARITY.md says so.
describe("openXlsx / saveXlsx", () => {
  it("round-trips the serial with its date format", async () => {
    const wb = await openXlsx(await datedColumn(SERIALS), { ...SERIAL, readStyles: true })
    expect(column(wb.sheets[0]!.rows)).toEqual(SERIALS)

    const saved = await saveXlsx(wb)

    // Serial 60 included: written back as the number, it is not moved
    // onto 59 the way a Date would be.
    expect(column((await readXlsx(saved, SERIAL)).sheets[0]!.rows)).toEqual(SERIALS)
    const style = (await readXlsx(saved, { readStyles: true })).sheets[0]!.cells!.get("0,0")!
    expect(style.style?.numFmt).toBe("yyyy-mm-dd hh:mm:ss")
  })
})

describe("streamXlsxRows", () => {
  it('yields the serial with dates: "serial"', async () => {
    const rows = await collect(streamXlsxRows(await datedColumn(SERIALS), SERIAL))

    expect(rows.map((r) => r.values[0])).toEqual(SERIALS)
  })

  it("yields the serial from a ReadableStream input", async () => {
    const rows = await collect(streamXlsxRows(asStream(await datedColumn(SERIALS)), SERIAL))

    expect(rows.map((r) => r.values[0])).toEqual(SERIALS)
  })

  it("still yields Dates by default", async () => {
    const rows = await collect(streamXlsxRows(await datedColumn([60])))

    expect(rows[0]!.values[0]).toBeInstanceOf(Date)
  })

  it('converts a t="d" cell the way readXlsx does', async () => {
    const bytes = await withCells('<c r="A1" t="d"><v>2024-03-17</v></c>', "1904")

    const rows = await collect(streamXlsxRows(bytes, SERIAL))

    expect(rows[0]!.values).toEqual([43906])
    expect(rows[0]!.values).toEqual((await readXlsx(bytes, SERIAL)).sheets[0]!.rows[0])
  })
})

// Excel-written files from the #464 corpus. The expected serials are
// counted by hand from what scripts/fixtures/make-fixtures.vbs authored:
// DateSerial(2024, 3, 17) is 45368, DateSerial(1999, 12, 31) is 36525,
// DateSerial(2024, 2, 29) is 45351, TimeSerial(13, 45, 30) is 49530 s.
describe("readXls", () => {
  it('returns serials from an Excel-written .xls with dates: "serial"', async () => {
    const rows = (await readXls(fixture("excel-dates.xls"), SERIAL)).sheets[0]!.rows
    const time = 49530 / 86400

    // iso, builtin 14, builtin 15, japanese, japanese era, korean
    expect(rows.slice(0, 6).map((r) => r[1])).toEqual([45368, 45368, 45368, 45368, 45368, 45368])
    expect(rows[6]![1]).toBeCloseTo(time, 12)
    expect(rows[7]![1]).toBeCloseTo(45368 + time, 9)
    // Numbers that never were dates are untouched.
    expect(rows.slice(8).map((r) => r[1])).toEqual([1234567.891, 0.125, "0042"])
  })

  it("still reads Dates by default", async () => {
    const rows = (await readXls(fixture("excel-dates.xls"))).sheets[0]!.rows

    expect(rows[0]![1]).toBeInstanceOf(Date)
  })
})

describe("readXlsb", () => {
  it('returns serials from an Excel-written .xlsb with dates: "serial"', async () => {
    const rows = (await readXlsb(fixture("excel-basic.xlsb"), SERIAL)).sheets[0]!.rows

    expect(rows.slice(1).map((r) => r[2])).toEqual([45368, 36525, 45351, "xy"])
  })

  it("still reads Dates by default", async () => {
    const rows = (await readXlsb(fixture("excel-basic.xlsb"))).sheets[0]!.rows

    expect(rows[1]![2]).toBeInstanceOf(Date)
  })
})

describe("one sheet, three containers", () => {
  it("returns the same serials whichever container Excel saved it in", async () => {
    const xlsx = (await readXlsx(fixture("excel-basic.xlsx"), SERIAL)).sheets[0]!.rows

    expect(xlsx.slice(1).map((r) => r[2])).toEqual([45368, 36525, 45351, "xy"])
    expect((await readXls(fixture("excel-basic.xls"), SERIAL)).sheets[0]!.rows).toEqual(xlsx)
    expect((await readXlsb(fixture("excel-basic.xlsb"), SERIAL)).sheets[0]!.rows).toEqual(xlsx)
  })
})

describe("read()", () => {
  it("passes dates through to the format reader", async () => {
    expect(column((await read(await datedColumn([60]), SERIAL)).sheets[0]!.rows)).toEqual([60])
    expect((await read(fixture("excel-dates.xls"), SERIAL)).sheets[0]!.rows[0]![1]).toBe(45368)
    expect((await read(fixture("excel-basic.xlsb"), SERIAL)).sheets[0]!.rows[1]![2]).toBe(45368)
  })
})

// ODS stores a date as ISO text (`office:date-value`), not as a number of
// days, so there is no serial to return — the same reason `dateSystem` is
// n/a there. Pinned so the documented behaviour is also the tested one.
describe("readOds", () => {
  it("is n/a: an ODS date is ISO text, and still reads as a Date", async () => {
    const day = new Date(Date.UTC(2024, 2, 17))
    const bytes = await writeOds({ sheets: [{ name: "S", rows: [[day]] }] })

    const value = (await readOds(bytes, SERIAL)).sheets[0]!.rows[0]![0]

    expect(value).toBeInstanceOf(Date)
    expect((value as Date).toISOString()).toBe(day.toISOString())
  })
})
