import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { writeXlsxStream, XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import type { ColumnDef } from "../src/_types"

// #616: writeXlsx and writeXlsxStream resolved a styled cell in a formatted
// column differently, and the stream writer never wrote `<col style>`.

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
}

const FMT = '"$"#,##0.00'

describe("column style parity between writeXlsx and writeXlsxStream (#616)", () => {
  it("a cell with its own style replaces the column's, in both writers", async () => {
    const columns: ColumnDef[] = [{ width: 12, numFmt: FMT }]
    const own = { font: { bold: true } }

    const buffered = await writeXlsx({
      sheets: [
        {
          name: "S",
          columns,
          rows: [[1.5]],
          cells: new Map([["0,0", { value: 1.5, style: own }]]),
        },
      ],
    })
    const streamed = await collect(
      writeXlsxStream([[{ value: 1.5, style: own }]], { name: "S", columns }),
    )

    for (const bytes of [buffered, streamed]) {
      const wb = await readXlsx(bytes, { readStyles: true })
      const style = wb.sheets[0].cells?.get("0,0")?.style
      expect(style?.font?.bold).toBe(true)
      expect(style?.numFmt).toBeUndefined()
    }
  })

  it("an unstyled cell in a formatted column gets the column format, in both", async () => {
    const columns: ColumnDef[] = [{ width: 12, numFmt: FMT }]
    const buffered = await writeXlsx({
      sheets: [{ name: "S", columns, rows: [[1.5]] }],
    })
    const streamed = await collect(writeXlsxStream([[1.5]], { name: "S", columns }))
    expect(await part(streamed, "xl/worksheets/sheet1.xml")).toContain('s="1"')
    expect(await part(streamed, "xl/styles.xml")).toContain("&quot;$&quot;#,##0.00")
    expect(await part(buffered, "xl/worksheets/sheet1.xml")).toContain('s="1"')
  })

  it("writeXlsxStream writes <col style> that reads back", async () => {
    const columns: ColumnDef[] = [
      { width: 12, style: { numFmt: FMT } },
      { style: { numFmt: "0.0%" } },
    ]
    const streamed = await collect(writeXlsxStream([[1.5, 0.5]], { name: "S", columns }))
    const xml = await part(streamed, "xl/worksheets/sheet1.xml")
    expect(xml).toMatch(/<col min="1" max="1" width="12" customWidth="true" style="\d+"\/>/)
    // a column with only a style still gets a <col>, and a visible width
    expect(xml).toMatch(/<col min="2" max="2" width="[\d.]+" style="\d+"\/>/)

    const wb = await readXlsx(streamed, { readStyles: true })
    expect(wb.sheets[0].columns?.[0]).toMatchObject({
      width: 12,
      style: { numFmt: FMT },
    })
    expect(wb.sheets[0].columns?.[1]?.style?.numFmt).toBe("0.0%")
  })

  it("the incremental XlsxStreamWriter writes <col style> too", async () => {
    const w = new XlsxStreamWriter({ name: "S", columns: [{ width: 12, numFmt: FMT }] })
    w.addRow([1.5])
    const bytes = await w.finish()
    const wb = await readXlsx(bytes, { readStyles: true })
    expect(wb.sheets[0].columns?.[0]?.style?.numFmt).toBe(FMT)
  })
})
