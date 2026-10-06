import { getCell } from "../src/cell-store"
import { describe, expect, it } from "vitest"
import { ParseError } from "../src/errors"
import { parseWorksheet, parseWorksheetStream } from "../src/xlsx/worksheet"
import type { WorksheetContext } from "../src/xlsx/worksheet"

const enc = new TextEncoder()
const context: WorksheetContext = {
  sharedStrings: [],
  styles: null,
  readStyles: false,
  dateSystem: "1900",
  maxTotalCells: 64,
}

function worksheet(data: string): string {
  return `<worksheet><sheetData>${data}</sheetData></worksheet>`
}

function byteStream(xml: string): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(enc.encode(xml))
      controller.close()
    },
  })
}

describe.each(["buffered", "streamed"] as const)("%s worksheet allocation limit", (driver) => {
  async function read(xml: string, ctx = context) {
    return driver === "buffered"
      ? parseWorksheet(xml, "S", ctx)
      : parseWorksheetStream(byteStream(xml), "S", ctx)
  }

  it("rejects a far-column cell before allocating its row or padding", async () => {
    // The old finish-only check paid for every touched row first. Stop
    // padding at a tiny budget so a regression fails without exhausting
    // the test runner's memory; the rejection must be hucre's ParseError.
    let slots = 0
    const originalPush = Array.prototype.push
    Array.prototype.push = function (this: unknown[], ...items: unknown[]): number {
      if (items.length === 1 && items[0] === null && ++slots > 64) {
        throw new Error("The reader allocated past maxTotalCells before rejecting the sheet")
      }
      return Reflect.apply(originalPush, this, items)
    }
    let failure: unknown
    try {
      await read(worksheet('<row><c r="XFD1"><v>1</v></c></row>'))
    } catch (error) {
      failure = error
    } finally {
      Array.prototype.push = originalPush
    }
    expect(failure).toBeInstanceOf(ParseError)
    expect(slots).toBe(0)
  })

  it("rejects a far-row cell before processing its value", async () => {
    let warnings = 0
    await expect(
      read(worksheet('<row><c r="A1048576" t="s"><v>0</v></c></row>'), {
        ...context,
        onWarning: () => {
          warnings++
        },
      }),
    ).rejects.toThrow(ParseError)
    expect(warnings).toBe(0)
  })

  it("bounds the combined rectangle when a later cell widens earlier rows", async () => {
    const xml = worksheet(
      '<row r="8"><c r="H8"><v>1</v></c></row>' + '<row r="1"><c r="I1"><v>2</v></c></row>',
    )
    await expect(read(xml)).rejects.toThrow(/8 rows x 9 columns.*over the 64 limit/)
  })

  it("still accepts the exact boundary and pads missing rows", async () => {
    const sheet = await read(worksheet('<row r="8"><c r="H8"><v>1</v></c></row>'))
    expect(sheet.rows).toHaveLength(8)
    expect(sheet.rows.every((row) => row.length === 8)).toBe(true)
    expect(sheet.rows[0][0]).toBeNull()
    expect(sheet.rows[7][7]).toBe(1)
  })

  it("does not apply the dense limit to sparse cells", async () => {
    const sheet = await read(worksheet('<row><c r="XFD8"><v>1</v></c></row>'), {
      ...context,
      sparse: true,
    })
    expect(sheet.rows).toEqual([])
    expect(getCell(sheet.cells, 7, 16383)?.value).toBe(1)
  })

  it("applies the bound only to cells admitted by the range", async () => {
    const sheet = await read(
      worksheet('<row><c r="A1"><v>1</v></c><c r="XFD1"><v>2</v></c></row>'),
      { ...context, range: "A1:A1" },
    )
    expect(sheet.rows).toEqual([[1]])
  })

  it("does not allocate cells excluded by maxRows", async () => {
    const sheet = await read(
      worksheet('<row><c r="A1"><v>1</v></c></row><row><c r="XFD2"><v>2</v></c></row>'),
      { ...context, maxRows: 1 },
    )
    expect(sheet.rows).toEqual([[1]])
  })

  it("keeps coordinate errors distinct from the dense limit", async () => {
    await expect(read(worksheet('<row><c r="XFE1"><v>1</v></c></row>'))).rejects.toThrow(
      /outside the supported sheet bounds/,
    )
  })
})
