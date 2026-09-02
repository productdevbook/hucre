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

  it("joins consecutive text:p with a newline, the way collectText does", async () => {
    // A multi-line cell has two spellings in ODF, and this file had only
    // ever been tested on one of them. `<text:line-break/>` inside a
    // single paragraph was covered above; **separate paragraphs** were
    // not, and the streaming reader ran them together — "linebreak" for
    // a cell the batch reader read as "line\nbreak".
    //
    // Which spelling you get depends on the writer: hucre emits
    // `<text:line-break/>`, so a suite that only ever parsed hucre's own
    // output could not see this. SheetJS emits paragraphs, and the #464
    // corpus is what found it.
    const buf = await odsFromContent(
      `<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string">` +
        `<text:p>line</text:p><text:p>break</text:p>` +
        `</table:table-cell></table:table-row></table:table>`,
    )
    const rows = []
    for await (const row of streamOdsRows(buf)) rows.push(row)

    expect(rows[0].values[0]).toBe("line\nbreak")
    expect((await readOds(buf)).sheets[0]!.rows[0]![0]).toBe("line\nbreak")
  })

  it("counts an empty paragraph as a line, not as nothing", async () => {
    // `join("\n")` over ["a", "", "b"] is "a\n\nb". A fix that appended a
    // newline only when the text so far was non-empty would give "a\nb"
    // and lose the blank line.
    const buf = await odsFromContent(
      `<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string">` +
        `<text:p>a</text:p><text:p></text:p><text:p>b</text:p>` +
        `</table:table-cell></table:table-row></table:table>`,
    )
    const rows = []
    for await (const row of streamOdsRows(buf)) rows.push(row)

    expect(rows[0].values[0]).toBe("a\n\nb")
    expect((await readOds(buf)).sheets[0]!.rows[0]![0]).toBe("a\n\nb")
  })

  it("does not put a newline before the first paragraph", async () => {
    const buf = await odsFromContent(
      `<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string">` +
        `<text:p>only</text:p>` +
        `</table:table-cell></table:table-row></table:table>`,
    )
    const rows = []
    for await (const row of streamOdsRows(buf)) rows.push(row)

    expect(rows[0].values[0]).toBe("only")
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
    for await (const row of streamOdsRows(buf, { sheet: "all" })) rows.push(row)
    expect(rows.map((r) => [r.sheet, r.index, r.values[0]])).toEqual([
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
