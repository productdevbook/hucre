import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { writeXlsxStream, XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// ═══════════════════════════════════════════════════════════════════════
// An empty string read back as `null`, but only when it had been written
// as an *inline* string:
//
//   <c t="s"><v>1</v></c>            shared, entry 1 is ""   →  ""
//   <c t="inlineStr"><is><t/></is></c>                       →  null
//
// Both are a producer saying "this cell holds the empty string". The
// asymmetry was accidental: #492 skips a cell that carries no
// information, and it decided that from the *text* it had collected —
// which for a shared string is the index `"1"`, non-empty, and for an
// inline string is the string itself, empty.
//
// #492 is about `<c r="WVF45" s="3"/>`: self-closing, no `t`, no
// content, written by Excel for every position formatting ever touched.
// A cell that declares `t="inlineStr"` and carries an `<is>` is not that.
// The producer wrote an element to say so.
//
// It matters because `writeXlsxStream` defaults to inline strings, so
// hucre's own three XLSX writers disagreed with each other on the same
// input — found by a property test comparing them.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}

/** A workbook whose sheet body is exactly the rows given. */
async function sheetWith(rowsXml: string): Promise<Uint8Array> {
  const base = await writeXlsx({ sheets: [{ name: "S", rows: [["seed"]] }] })
  const all = await new ZipReader(base).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? enc.encode(
            dec
              .decode(data)
              .replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${rowsXml}</sheetData>`),
          )
        : data,
    )
  }
  return zw.build()
}

describe("an empty inline string is a value", () => {
  it('reads as "", not null', async () => {
    const bytes = await sheetWith(
      '<row r="1">' +
        '<c r="A1" t="inlineStr"><is><t>a</t></is></c>' +
        '<c r="B1" t="inlineStr"><is><t/></is></c>' +
        '<c r="C1" t="inlineStr"><is><t>c</t></is></c>' +
        "</row>",
    )

    expect((await readXlsx(bytes)).sheets[0]!.rows[0]).toEqual(["a", "", "c"])
  })

  it("the same as the shared-string spelling of it always did", async () => {
    // The two spellings of one value have to agree; this is the pair.
    const inline = await sheetWith('<row r="1"><c r="A1" t="inlineStr"><is><t/></is></c></row>')
    const shared = await writeXlsx({ sheets: [{ name: "S", rows: [[""]] }] })

    expect((await readXlsx(inline)).sheets[0]!.rows[0]![0]).toBe(
      (await readXlsx(shared)).sheets[0]!.rows[0]![0],
    )
  })

  it("including an <is> with no <t> at all", async () => {
    const bytes = await sheetWith(
      '<row r="1"><c r="A1" t="inlineStr"><is/></c><c r="B1" t="inlineStr"><is><t>b</t></is></c></row>',
    )

    expect((await readXlsx(bytes)).sheets[0]!.rows[0]).toEqual(["", "b"])
  })
})

describe("hucre's three XLSX writers agree with each other", () => {
  // The property test that found this compared them on random grids;
  // this is the reduced case, kept because the agreement is the point.
  const ROWS = [["a", "", "c"]]

  it("on an empty string in the middle of a row", async () => {
    const incremental = new XlsxStreamWriter({ name: "S" })
    for (const row of ROWS) incremental.addRow(row)

    const variants = [
      await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] }),
      await drain(writeXlsxStream(ROWS, { name: "S" })),
      await drain(writeXlsxStream(ROWS, { name: "S", stringMode: "inline" })),
      await incremental.finish(),
    ]

    for (const bytes of variants) {
      expect((await readXlsx(bytes)).sheets[0]!.rows[0]).toEqual(["a", "", "c"])
    }
  })

  it("and writeXlsxStream really is writing the inline spelling", async () => {
    // Otherwise the test above could pass for the wrong reason.
    const bytes = await drain(writeXlsxStream(ROWS, { name: "S" }))
    const xml = dec.decode(await new ZipReader(bytes).extract("xl/worksheets/sheet1.xml"))

    expect(xml).toContain('t="inlineStr"')
    expect(xml).toContain("<t/>")
  })
})

describe("the #492 cells are still skipped", () => {
  it("a self-closing style-only cell carries nothing", async () => {
    // The guard this fix relaxes exists for these. A cell with no `t`
    // and no content is not a producer saying "empty string" — it is
    // Excel remembering that someone once formatted the position.
    const bytes = await sheetWith(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="WVF1" s="3"/></row>',
    )

    expect((await readXlsx(bytes)).sheets[0]!.rows[0]).toEqual(["a"])
  })

  it("but is kept when its style is being read", async () => {
    const bytes = await sheetWith(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="D1" s="0"/></row>',
    )
    const rows = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.rows

    expect(rows[0]).toHaveLength(4)
  })
})
