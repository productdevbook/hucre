import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import type { CellStyle, RowDef, WriteSheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #601 — a row format.
//
// Excel writes a formatted row as `<row s="N" customFormat="1">`
// (ECMA-376 Part 1 §18.3.1.73), and stamps the same `s` onto the cells
// already in the row. `RowDef` had height, hidden and outline but no
// style, so the writer had nothing to say and the reader dropped the
// attribute: a row format in an Excel file was lost on the way through.
// Columns have had the equivalent since #439 §W.
// ═══════════════════════════════════════════════════════════════════════

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
}

async function sheetXml(sheet: WriteSheet): Promise<string> {
  return part(await writeXlsx({ sheets: [sheet] }), "xl/worksheets/sheet1.xml")
}

/** Rebuild an archive with one part rewritten, to stand in for another tool's output. */
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

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

const TOTAL: CellStyle = {
  numFmt: "0.00",
  fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFDDEBF7" } },
}
const ITALIC: CellStyle = { font: { italic: true } }

/** The `s` a `<row r="N">` carries, or undefined. */
const rowS = (xml: string, r: number): string | undefined =>
  xml.match(new RegExp(`<row r="${r}"[^>]*? s="(\\d+)"`))?.[1]

describe("a row format reaches the file", () => {
  it("writes s and customFormat on <row>, beside the height", async () => {
    const xml = await sheetXml({
      name: "S",
      rows: [
        ["Name", "Qty"],
        ["Total", 3],
      ],
      rowDefs: new Map<number, RowDef>([[1, { height: 20, style: TOTAL }]]),
    })

    expect(xml).toMatch(/<row r="2" ht="20" customHeight="1" s="[1-9]\d*" customFormat="1">/)
    expect(xml).toMatch(/<row r="1">/)
  })

  it("writes a row that has a format and no cells", async () => {
    const xml = await sheetXml({
      name: "S",
      rows: [["a"], [], ["c"]],
      rowDefs: new Map<number, RowDef>([[1, { style: TOTAL }]]),
    })

    expect(xml).toMatch(/<row r="2" s="[1-9]\d*" customFormat="1"\/>/)
  })

  it("stamps the cells in the row that carry no style of their own", async () => {
    const xml = await sheetXml({
      name: "S",
      rows: [["Total", 3]],
      cells: new Map([["0,1", { value: 3, style: ITALIC }]]),
      rowDefs: new Map<number, RowDef>([[0, { style: TOTAL }]]),
    })

    const s = rowS(xml, 1)
    expect(s).toBeDefined()
    expect(xml).toContain(`<c r="A1" t="s" s="${s}">`)
    // B1 has a style of its own, so it keeps it.
    expect(xml).toMatch(/<c r="B1" s="(\d+)">/)
    expect(xml.match(/<c r="B1" s="(\d+)">/)![1]).not.toBe(s)
  })

  it("lets the row format win over a column format on a cell neither styles", async () => {
    // Excel applies whichever format came last; a row format is applied
    // after the sheet's column formats here. Not verified in Excel.
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"], ["b"]],
          columns: [{ style: ITALIC }],
          rowDefs: new Map<number, RowDef>([[1, { style: TOTAL }]]),
        },
      ],
    })

    const cells = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells!
    expect(cells.get("0,0")!.style!.font!.italic).toBe(true)
    expect(cells.get("1,0")!.style!.numFmt).toBe("0.00")
    expect(cells.get("1,0")!.style!.font?.italic).toBeUndefined()
  })

  it("stamps the cells written from data[] as well as rows", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          columns: [{ key: "n", header: "N" }],
          data: [{ n: 1 }, { n: 2 }],
          rowDefs: new Map<number, RowDef>([[2, { style: TOTAL }]]),
        },
      ],
    })

    const cells = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells!
    expect(cells.get("2,0")!.style!.numFmt).toBe("0.00")
    expect(cells.get("1,0")?.style?.numFmt).toBeUndefined()
  })

  it("stamps a cell a cells override adds to the row", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          cells: new Map([["0,3", { value: 4 }]]),
          rowDefs: new Map<number, RowDef>([[0, { style: TOTAL }]]),
        },
      ],
    })

    const cells = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells!
    expect(cells.get("0,3")!.style!.numFmt).toBe("0.00")
  })

  it("is written by the streaming writer too", async () => {
    const bytes = await collect(
      writeXlsxStream([["Total", 3], [], ["x"]], {
        name: "S",
        rowDefs: new Map<number, RowDef>([
          [0, { style: TOTAL }],
          [1, { style: TOTAL }],
        ]),
      }),
    )

    const xml = await part(bytes, "xl/worksheets/sheet1.xml")
    const s = rowS(xml, 1)
    expect(s).toBeDefined()
    expect(xml).toMatch(/<row r="1" s="\d+" customFormat="1">/)
    expect(xml).toContain(`<row r="2" s="${s}" customFormat="1"/>`)
    expect(xml).toContain(`<c r="B1" s="${s}">`)
    expect(xml).toMatch(/<row r="3">/)
  })

  it("lets the row format win over a column number format, streamed or not", async () => {
    // The streaming writer folds a column's numFmt into whatever style a
    // cell ends up with. A row format replaces the column's outright in
    // writeXlsx, so it has to do the same here or the two writers disagree.
    const columns = [{ numFmt: "0%" }]
    const rowDefs = new Map<number, RowDef>([[1, { style: ITALIC }]])
    const streamed = await collect(writeXlsxStream([[0.5], [0.5]], { name: "S", columns, rowDefs }))
    const buffered = await writeXlsx({
      sheets: [{ name: "S", rows: [[0.5], [0.5]], columns, rowDefs }],
    })

    for (const bytes of [streamed, buffered]) {
      const cells = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells!
      expect(cells.get("0,0")!.style!.numFmt).toBe("0%")
      expect(cells.get("1,0")!.style!.font!.italic).toBe(true)
      expect(cells.get("1,0")!.style!.numFmt).toBeUndefined()
    }
  })

  it("leaves a row with no format byte-for-byte as it was", async () => {
    // The new attributes go out only for a style; a row definition without
    // one writes exactly the `<row>` and cells it wrote before.
    const xml = await sheetXml({
      name: "S",
      rows: [
        ["a", 1],
        ["b", 2],
      ],
      rowDefs: new Map<number, RowDef>([[1, { height: 20 }]]),
    })

    expect(xml).toContain(
      '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>' +
        '<row r="2" ht="20" customHeight="1"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>2</v></c></row></sheetData>',
    )
    expect(xml).not.toContain("customFormat")
  })
})

describe("a row format survives a read", () => {
  it("reads <row s customFormat> back as RowDef.style", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["Name"], ["Total"]],
          rowDefs: new Map<number, RowDef>([[1, { height: 20, style: TOTAL }]]),
        },
      ],
    })

    const def = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.rowDefs!.get(1)!

    expect(def.height).toBe(20)
    expect(def.style!.numFmt).toBe("0.00")
    expect(def.style!.fill).toMatchObject({ pattern: "solid", fgColor: { rgb: "DDEBF7" } })
  })

  it("reads a row that has only a format", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"], [], ["c"]], rowDefs: new Map([[1, { style: TOTAL }]]) }],
    })

    const def = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.rowDefs!.get(1)

    expect(def?.style?.numFmt).toBe("0.00")
  })

  it("ignores s when customFormat does not say the row is formatted", async () => {
    // §18.3.1.73: `s` applies to the row only when `customFormat` is 1.
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]], rowDefs: new Map([[0, { style: TOTAL }]]) }],
    })
    const patched = await rewrite(bytes, "xl/worksheets/sheet1.xml", (xml) =>
      xml.replace('customFormat="1"', 'customFormat="0"'),
    )

    const sheet = (await readXlsx(patched, { readStyles: true })).sheets[0]!

    expect(sheet.rowDefs?.get(0)?.style).toBeUndefined()
  })

  it("leaves the format out when styles are not being read", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          rowDefs: new Map<number, RowDef>([[0, { height: 20, style: TOTAL }]]),
        },
      ],
    })

    const def = (await readXlsx(bytes)).sheets[0]!.rowDefs!.get(0)!

    expect(def).toEqual({ height: 20 })
  })

  it("gives each row its own style object", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"], ["b"]],
          rowDefs: new Map<number, RowDef>([
            [0, { style: TOTAL }],
            [1, { style: TOTAL }],
          ]),
        },
      ],
    })

    const rowDefs = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.rowDefs!

    expect(rowDefs.get(0)!.style).not.toBe(rowDefs.get(1)!.style)
    rowDefs.get(0)!.style!.numFmt = "changed"
    expect(rowDefs.get(1)!.style!.numFmt).toBe("0.00")
  })
})
