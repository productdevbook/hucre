import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// ═══════════════════════════════════════════════════════════════════════
// #496 — ECMA-376 §18.18.11 `ST_CellType` has seven members and the
// reader's switch had six. `d` — "Cell containing a date in the ISO 8601
// format" — fell through to `n`, where `Number("2024-03-17")` is NaN, and
// landed in the arm commented "shouldn't happen, but be safe". It does
// happen: openpyxl writes it whenever `iso_dates=True`.
//
// So the same day, under the same number format, came back as a `Date`
// when stored as a serial and as a **string** when stored as ISO text.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

/** A workbook whose A1 is a raw `<c>` element of our choosing. */
async function withCell(cellXml: string): Promise<Uint8Array> {
  const base = await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] })
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

async function valueOf(cellXml: string): Promise<unknown> {
  return (await readXlsx(await withCell(cellXml))).sheets[0]!.rows[0]![0]
}

describe('t="d" cells read as dates', () => {
  it("a plain ISO date", async () => {
    const value = await valueOf('<c r="A1" t="d"><v>2024-03-17</v></c>')

    expect(value).toBeInstanceOf(Date)
    expect((value as Date).toISOString()).toBe("2024-03-17T00:00:00.000Z")
  })

  it("an ISO date-time, read as UTC when it names no zone", async () => {
    // Same reason as the ODS and docProps readers: every format hucre
    // reads records an absolute moment, and local time would make one
    // file mean different things on different machines. See #415, #474.
    const value = await valueOf('<c r="A1" t="d"><v>2024-03-17T13:45:30</v></c>')

    expect((value as Date).toISOString()).toBe("2024-03-17T13:45:30.000Z")
  })

  it("honours a zone the cell does name", async () => {
    expect(
      (
        (await valueOf('<c r="A1" t="d"><v>2024-03-17T13:45:30+02:00</v></c>')) as Date
      ).toISOString(),
    ).toBe("2024-03-17T11:45:30.000Z")
    expect(
      ((await valueOf('<c r="A1" t="d"><v>2024-03-17T13:45:30Z</v></c>')) as Date).toISOString(),
    ).toBe("2024-03-17T13:45:30.000Z")
  })

  it("takes fractional seconds", async () => {
    expect(
      (
        (await valueOf('<c r="A1" t="d"><v>2024-03-17T13:45:30.250Z</v></c>')) as Date
      ).toISOString(),
    ).toBe("2024-03-17T13:45:30.250Z")
  })

  it("does not apply the 1904 epoch to it", async () => {
    // The value is an instant, not an offset from an epoch. The serial
    // path shifts by 1,462 days between the two systems; this one must
    // not move at all.
    const bytes = await withCell('<c r="A1" t="d"><v>2024-03-17</v></c>')
    const all = await new ZipReader(bytes).extractAll()
    const zw = new ZipWriter()
    for (const [name, data] of all) {
      zw.add(
        name,
        name === "xl/workbook.xml"
          ? enc.encode(dec.decode(data).replace("<workbookPr", '<workbookPr date1904="1"'))
          : data,
      )
    }
    const wb = await readXlsx(await zw.build())

    expect((wb.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2024-03-17T00:00:00.000Z")
  })
})

describe("what stays text, and why", () => {
  it("a bare time has no day to anchor it", async () => {
    // openpyxl writes this for a `datetime.time`. Guessing it onto an
    // epoch would invent a date the file does not contain.
    expect(await valueOf('<c r="A1" t="d"><v>13:45:30</v></c>')).toBe("13:45:30")
  })

  it("anything that is not an ISO date is left alone", async () => {
    // `new Date(text)` accepts a great deal that is not ISO 8601, so a
    // loose parse here would turn arbitrary cell text into dates.
    for (const text of ["March 17 2024", "17/03/2024", "2024", "not a date", "2024-13-45"]) {
      expect(await valueOf(`<c r="A1" t="d"><v>${text}</v></c>`), text).toBe(text)
    }
  })

  it('an empty t="d" cell is empty, not Invalid Date', async () => {
    expect(await valueOf('<c r="A1" t="d"><v></v></c>')).toBeNull()
  })
})

describe("the rest of ST_CellType still behaves", () => {
  it("a serial under a date format is still a Date, and still epoch-aware", async () => {
    const wb = await readXlsx(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [[new Date("2024-03-17T00:00:00Z")]],
          },
        ],
      }),
    )

    expect((wb.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2024-03-17T00:00:00.000Z")
  })

  it("a plain number is still a number", async () => {
    expect(await valueOf('<c r="A1"><v>42</v></c>')).toBe(42)
  })

  it("an error cell is still an error", async () => {
    expect(await valueOf('<c r="A1" t="e"><v>#DIV/0!</v></c>')).toBe("#DIV/0!")
  })
})
