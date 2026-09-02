import { describe, expect, it } from "vitest"
import { read } from "../src/defter"
import { HucreError, InvalidArgumentError, UnsupportedFormatError } from "../src/errors"
import { moveSheet, removeSheet } from "../src/sheet-ops"
import { addChart } from "../src/xlsx/chart-helpers"
import { openXlsx } from "../src/xlsx/roundtrip"
import { writeXlsx } from "../src/xlsx/writer"
import { ZipWriter } from "../src/zip/writer"

const enc = new TextEncoder()

describe("read() names a ZIP that is not a spreadsheet", () => {
  it("refuses a plain archive", async () => {
    const zip = new ZipWriter()
    zip.add("hello.txt", enc.encode("hi"))
    await expect(read(await zip.build())).rejects.toThrow(UnsupportedFormatError)
  })

  it("refuses an Office package that is not a spreadsheet", async () => {
    const zip = new ZipWriter()
    zip.add(
      "[Content_Types].xml",
      enc.encode(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
    )
    zip.add("word/document.xml", enc.encode("<w:document/>"))
    await expect(read(await zip.build())).rejects.toThrow(/not a spreadsheet/)
  })

  it("still reads a real workbook", async () => {
    const wb = await read(await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] }))
    expect(wb.sheets[0]!.rows).toEqual([[1]])
  })
})

describe("every error the library throws is a HucreError", () => {
  it("moveSheet and removeSheet check their indexes instead of splicing undefined", () => {
    const wb = { sheets: [{ name: "A", rows: [] }] }
    expect(() => moveSheet(wb, 0, 3)).toThrow(InvalidArgumentError)
    expect(() => removeSheet(wb, 1)).toThrow(InvalidArgumentError)
    expect(wb.sheets).toHaveLength(1)
  })

  it("argument misuse in the chart helpers is an InvalidArgumentError, not a TypeError", () => {
    try {
      addChart({ name: "S", rows: [] }, undefined as never)
      expect.unreachable()
    } catch (e) {
      expect(e).toBeInstanceOf(InvalidArgumentError)
      expect(e).toBeInstanceOf(HucreError)
    }
  })
})

describe("openXlsx takes ReadInput like readXlsx", () => {
  it("accepts a ReadableStream", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        c.enqueue(bytes)
        c.close()
      },
    })
    const wb = await openXlsx(stream)
    expect(wb.sheets[0]!.rows).toEqual([["a"]])
  })
})
