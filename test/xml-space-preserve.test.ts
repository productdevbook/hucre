import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { ZipReader } from "../src/zip/reader"

// ═══════════════════════════════════════════════════════════════════════
// A `<t>` whose text starts or ends with a space, or holds a newline or a
// tab, has to declare `xml:space="preserve"` — without it an XML consumer
// is entitled to collapse the whitespace, and Excel does.
//
// The check used to be copy-pasted at each of the four sites that emit a
// `<t>`. Three had it; the inline-string branch of the worksheet writer
// did not, so `writeXlsx({ stringMode: "inline" })` shipped padding Excel
// silently trimmed. hucre's own reader does not trim, so a round-trip
// assertion could never see it — these read the emitted XML instead.
// ═══════════════════════════════════════════════════════════════════════

const PADDED = "  padded  "
const TABBED = "a\tb"
const LINES = "a\nb"

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    total += value.length
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

describe("xml:space=preserve on every <t> that needs it", () => {
  it("declares it on an inline string", async () => {
    const bytes = await writeXlsx({
      stringMode: "inline",
      sheets: [{ name: "S", rows: [[PADDED]] }],
    })

    const xml = await part(bytes, "xl/worksheets/sheet1.xml")

    expect(xml).toContain(`<t xml:space="preserve">${PADDED}</t>`)
  })

  it("declares it on a shared string", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[PADDED]] }] })

    const xml = await part(bytes, "xl/sharedStrings.xml")

    expect(xml).toContain(`<t xml:space="preserve">${PADDED}</t>`)
  })

  it("declares it on a rich-text run", async () => {
    const cells = new Map([["0,0", { richText: [{ text: PADDED, font: { bold: true } }] }]])
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[null]], cells }] })

    const xml = await part(bytes, "xl/worksheets/sheet1.xml")

    expect(xml).toContain(`<t xml:space="preserve">${PADDED}</t>`)
  })

  it("declares it on a streamed inline string", async () => {
    const bytes = await collect(writeXlsxStream([[PADDED]], { name: "S" }))

    const xml = await part(bytes, "xl/worksheets/sheet1.xml")

    expect(xml).toContain(`<t xml:space="preserve">${PADDED}</t>`)
  })

  it("covers tabs and newlines, not only leading and trailing spaces", async () => {
    const bytes = await writeXlsx({
      stringMode: "inline",
      sheets: [{ name: "S", rows: [[TABBED], [LINES]] }],
    })

    const xml = await part(bytes, "xl/worksheets/sheet1.xml")

    expect(xml).toContain(`<t xml:space="preserve">${TABBED}</t>`)
    expect(xml).toContain(`<t xml:space="preserve">${LINES}</t>`)
  })

  it("leaves it off text that does not need it", async () => {
    const bytes = await writeXlsx({
      stringMode: "inline",
      sheets: [{ name: "S", rows: [["plain"]] }],
    })

    const xml = await part(bytes, "xl/worksheets/sheet1.xml")

    expect(xml).toContain("<t>plain</t>")
    expect(xml).not.toContain('xml:space="preserve"')
  })
})
