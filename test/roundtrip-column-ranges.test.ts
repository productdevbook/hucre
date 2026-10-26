import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// #622: `<col min="2" max="16384">` was read as 16,383 ColumnDefs and written
// back as 16,383 `<col>` elements. The writer now merges adjacent identical
// columns back into one range.

async function sheetXml(bytes: Uint8Array): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract("xl/worksheets/sheet1.xml"))
}

async function rewrite(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string,
): Promise<Uint8Array> {
  const all = await new ZipReader(bytes).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === path ? new TextEncoder().encode(edit(new TextDecoder().decode(data))) : data,
    )
  }
  return zw.build()
}

const cols = (xml: string) => xml.match(/<col [^>]*\/>/g) ?? []

async function template(): Promise<Uint8Array> {
  const base = await writeXlsx({
    sheets: [{ name: "S", rows: [["x"]], columns: [{ width: 20 }, { width: 9 }] }],
  })
  return rewrite(base, "xl/worksheets/sheet1.xml", (x) =>
    x.replace(
      /<cols>[\s\S]*<\/cols>/,
      '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="16384" width="9" hidden="1" customWidth="1"/></cols>',
    ),
  )
}

describe("column ranges (#622)", () => {
  it("writes a read B:XFD range back as one <col>", async () => {
    const wb = await openXlsx(await template())
    expect(wb.sheets[0]!.columns).toHaveLength(16384)
    const xml = await sheetXml(await saveXlsx(wb))
    expect(cols(xml)).toHaveLength(2)
    expect(xml).toMatch(/<col min="2" max="16384" [^>]*hidden="(1|true)"/)
  })

  it("merges caller-authored adjacent identical columns, and only those", async () => {
    const xml = await sheetXml(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [[1]],
            columns: [
              { width: 10 },
              { width: 10 },
              { width: 10 },
              { width: 12 },
              {},
              { width: 10 },
            ],
          },
        ],
      }),
    )
    expect(cols(xml).map((c) => /min="(\d+)" max="(\d+)"/.exec(c)!.slice(1).join("-"))).toEqual([
      "1-3",
      "4-4",
      "6-6",
    ])
  })

  it("merges in the stream writer too", async () => {
    const stream = writeXlsxStream([[1, 2, 3]], {
      name: "S",
      columns: [{ width: 10 }, { width: 10 }, { width: 10 }],
    })
    const total = new Uint8Array(await new Response(stream).arrayBuffer())
    expect(cols(await sheetXml(total))).toHaveLength(1)
  })

  it("reads a 16384-wide range without a per-column style copy blowing up", async () => {
    const t0 = performance.now()
    const wb = await readXlsx(await template())
    expect(wb.sheets[0]!.columns).toHaveLength(16384)
    expect(performance.now() - t0).toBeLessThan(2000)
  })
})
