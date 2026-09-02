import { cellError } from "../src/cell-error"
import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"

// ═══════════════════════════════════════════════════════════════════════
// #439 §B — XLSX had two independent implementations of `<c>`, and they
// had drifted in *both* directions:
//
//   the streaming copy could not write   an error value (t="e"), rich
//                                        text, a checkbox xf, a shared or
//                                        array formula, the dynamic-array cm
//   the authoring copy could not write   xml:space="preserve" on an
//                                        inline string (fixed in #441),
//                                        and had no non-finite guard on a
//                                        cached formula result
//
// Every feature added to one had to be re-derived in the other, and #436
// had to do exactly that. There is one implementation now; these pin what
// the streaming path gained, and that the authoring path did not regress.
// ═══════════════════════════════════════════════════════════════════════

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

async function sheetXml(bytes: Uint8Array): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract("xl/worksheets/sheet1.xml"))
}

describe("a streamed error value is an error, not a string", () => {
  it('writes t="e"', async () => {
    const bytes = await collect(
      writeXlsxStream([[cellError("#REF!"), cellError("#DIV/0!")]], { name: "S" }),
    )

    const xml = await sheetXml(bytes)

    expect(xml).toContain('t="e"')
    expect(xml).toContain("<v>#REF!</v>")
  })

  it("reads back as the error it was, not as text", async () => {
    const bytes = await collect(
      writeXlsxStream(
        [[cellError("#VALUE!"), cellError("#N/A"), cellError("#SPILL!"), "ordinary"]],
        {
          name: "S",
        },
      ),
    )

    const cells = (await readXlsx(bytes)).sheets[0]!.cells!

    expect(cells.get("0,0")!.type).toBe("error")
    expect(cells.get("0,1")!.type).toBe("error")
    // #SPILL! is one of the two dynamic-array errors ECMA does not list;
    // hucre knows them because #423 needed it to.
    expect(cells.get("0,2")!.type).toBe("error")
    // An ordinary string needs no cell entry at all, so check the value.
    expect((await readXlsx(bytes)).sheets[0]!.rows[0]![3]).toBe("ordinary")
  })

  it("agrees with the authoring writer, cell for cell", async () => {
    const rows = [[cellError("#REF!"), 1, "text", true, null]]

    const streamed = await readXlsx(await collect(writeXlsxStream(rows, { name: "S" })))
    const authored = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }))

    expect(streamed.sheets[0]!.rows).toEqual(authored.sheets[0]!.rows)
    expect(streamed.sheets[0]!.cells!.get("0,0")!.type).toBe(
      authored.sheets[0]!.cells!.get("0,0")!.type,
    )
  })
})

describe("a non-finite cached result is dropped by both writers", () => {
  // #436 added this guard to the streaming copy only. Sharing the
  // serializer surfaced that the authoring path never had it — NaN went
  // out as `<v>NaN</v>`, which no reader can parse as a number.
  // On the authoring path the cached result is `formulaResult`; `value`
  // is the cell's own value. The streaming path has one slot for both,
  // which is why StreamStyledCell documents `value` as the cache.
  const cells = new Map([
    ["0,0", { value: null, formula: "0/0", formulaResult: Number.NaN }],
    ["0,1", { value: null, formula: "1/0", formulaResult: Number.POSITIVE_INFINITY }],
    ["0,2", { value: null, formula: "1+1", formulaResult: 2 }],
  ])

  it("drops it on the authoring path", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[null, null, null]], cells }],
    })

    const xml = await sheetXml(bytes)
    expect(xml).not.toContain("NaN")
    expect(xml).not.toContain("Infinity")

    const back = (await readXlsx(bytes)).sheets[0]!.cells!
    expect(back.get("0,0")!.formula).toBe("0/0")
    expect(back.get("0,0")!.value).toBeNull()
    expect(back.get("0,2")!.value).toBe(2)
  })

  it("drops it on the streaming path", async () => {
    const bytes = await collect(
      writeXlsxStream(
        [
          [
            { value: Number.NaN, formula: "0/0" },
            { value: 2, formula: "1+1" },
          ],
        ],
        { name: "S" },
      ),
    )

    const xml = await sheetXml(bytes)
    expect(xml).not.toContain("NaN")

    const back = (await readXlsx(bytes)).sheets[0]!.cells!
    expect(back.get("0,0")!.value).toBeNull()
    expect(back.get("0,1")!.value).toBe(2)
  })
})

describe("the authoring path did not regress", () => {
  it("still writes rich text, checkboxes and shared formulas", async () => {
    const cells = new Map<string, Record<string, unknown>>([
      ["0,0", { richText: [{ text: "bold", font: { bold: true } }] }],
      ["1,0", { value: true, checkbox: true }],
      ["2,0", { value: 3, formula: "A1+A2", formulaType: "shared", formulaSharedIndex: 0 }],
      ["3,0", { value: 1, formula: "SEQUENCE(3)", formulaDynamic: true }],
    ])
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[null], [null], [null], [null]], cells: cells as never }],
    })

    const xml = await sheetXml(bytes)

    expect(xml).toContain("<is>")
    expect(xml).toContain('t="shared"')
    expect(xml).toContain('si="0"')
    expect(xml).toMatch(/cm="\d+"/)
  })

  it("still round-trips a plain sheet unchanged", async () => {
    const rows = [
      ["Name", "Amount", "When"],
      ["Ada", 1234.5, new Date(Date.UTC(2024, 0, 15))],
    ]

    const back = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows }] }))

    expect(back.sheets[0]!.rows[1]![0]).toBe("Ada")
    expect(back.sheets[0]!.rows[1]![1]).toBe(1234.5)
    expect(back.sheets[0]!.rows[1]![2]).toBeInstanceOf(Date)
  })
})
