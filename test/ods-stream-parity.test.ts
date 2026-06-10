import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { writeOds } from "../src/ods/writer"
import { streamOdsRows } from "../src/ods/stream"
import { readOds } from "../src/ods/reader"

const enc = new TextEncoder()

async function odsFromContent(bodyXml: string): Promise<Uint8Array> {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>${bodyXml}</office:spreadsheet></office:body>
</office:document-content>`
  const zip = new ZipWriter()
  zip.add("mimetype", enc.encode("application/vnd.oasis.opendocument.spreadsheet"), {
    compress: false,
  })
  zip.add("content.xml", enc.encode(content))
  return zip.build()
}

describe("streamOdsRows — text element parity with batch reader", () => {
  it("expands text:s / line-break / tab like collectText", async () => {
    const buf = await odsFromContent(
      `<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string">` +
        `<text:p>a<text:s text:c="2"/>b<text:tab/>c<text:line-break/>d</text:p>` +
        `</table:table-cell></table:table-row></table:table>`,
    )
    const rows = []
    for await (const row of streamOdsRows(buf)) rows.push(row)
    expect(rows[0].values[0]).toBe("a  b\tc\nd")
  })
})

describe("streamOdsRows — sheet index", () => {
  it("tags rows with their sheet index", async () => {
    const buf = await writeOds({
      sheets: [
        { name: "One", rows: [["a"]] },
        { name: "Two", rows: [["b"], ["c"]] },
      ],
    })
    const rows = []
    for await (const row of streamOdsRows(buf)) rows.push(row)
    expect(rows.map((r) => [r.sheetIndex, r.index, r.values[0]])).toEqual([
      [0, 0, "a"],
      [1, 0, "b"],
      [1, 1, "c"],
    ])
  })
})

describe("streamOdsRows — number-rows-repeated DoS cap", () => {
  it("does not allocate millions of rows for a non-empty repeated row", async () => {
    const buf = await odsFromContent(
      `<table:table table:name="S"><table:table-row table:number-rows-repeated="5000000">` +
        `<table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table>`,
    )
    let count = 0
    for await (const _row of streamOdsRows(buf)) {
      count++
      if (count > 1_048_576) break // safety
    }
    expect(count).toBeLessThanOrEqual(1_048_576)
    expect(count).toBeGreaterThan(0)
  })

  it("batch readOds also caps a non-empty repeated row", async () => {
    const buf = await odsFromContent(
      `<table:table table:name="S"><table:table-row table:number-rows-repeated="5000000">` +
        `<table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table>`,
    )
    const wb = await readOds(buf)
    expect(wb.sheets[0].rows.length).toBeLessThanOrEqual(1_048_576)
    expect(wb.sheets[0].rows.length).toBeGreaterThan(0)
  })
})
