import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §W and §X — two things that describe a whole column or a whole
// sheet, and that never reached the file.
//
// `ColumnDef.style` was stamped onto the cells hucre wrote and nowhere
// else, so it looked right for those rows and vanished everywhere else:
// in Excel a column format applies to every cell in the column including
// ones nobody has typed in, and on read there was only `<col>` to look at.
//
// `<sheetFormatPr defaultRowHeight>` was hard-coded to 15 and never read,
// so a workbook whose default was 24 came back with every unstyled row
// shortened.
// ═══════════════════════════════════════════════════════════════════════

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
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

const CURRENCY: CellStyle = { numFmt: '"$"#,##0.00', font: { italic: true } }

describe("a column format reaches the file", () => {
  it("writes it onto <col style>", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], columns: [{ width: 12, style: CURRENCY }] }],
    })

    const xml = await part(bytes, "xl/worksheets/sheet1.xml")

    expect(xml).toMatch(/<col [^>]*style="\d+"/)
  })

  it("survives a round trip", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], columns: [{ width: 12, style: CURRENCY }] }],
    })

    const back = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.columns![0]!

    expect(back.width).toBe(12)
    expect(back.style!.numFmt).toBe('"$"#,##0.00')
    expect(back.style!.font!.italic).toBe(true)
  })

  it("folds ColumnDef.numFmt into the column style, as it does for cells", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], columns: [{ numFmt: "0.000" }] }],
    })

    const back = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.columns![0]!

    expect(back.style!.numFmt).toBe("0.000")
  })

  it("lets an explicit style.numFmt win over the shorthand", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], columns: [{ numFmt: "0.000", style: { numFmt: "0%" } }] }],
    })

    const back = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.columns![0]!

    expect(back.style!.numFmt).toBe("0%")
  })

  it("still stamps the cells, so a written row carries the format too", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], columns: [{ style: CURRENCY }] }],
    })

    const cell = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells!.get("0,0")!

    expect(cell.style!.numFmt).toBe('"$"#,##0.00')
  })

  it("gives each column its own style object on read", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1, 2]], columns: [{ style: CURRENCY }, { style: CURRENCY }] }],
    })

    const columns = (await readXlsx(bytes, { readStyles: true })).sheets[0]!.columns!

    expect(columns[0]!.style).not.toBe(columns[1]!.style)
    columns[0]!.style!.numFmt = "changed"
    expect(columns[1]!.style!.numFmt).toBe('"$"#,##0.00')
  })

  it("reads another tool's bestFit back as autoWidth", async () => {
    // hucre resolves `autoWidth` to a concrete width on write, so its own
    // files carry a width rather than the flag. Excel writes the flag, and
    // it used to be read by nobody.
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [["a value"]], columns: [{ width: 10 }] }],
    })

    const patched = await rewrite(bytes, "xl/worksheets/sheet1.xml", (xml) =>
      xml.replace("<col ", '<col bestFit="1" '),
    )

    expect((await readXlsx(patched)).sheets[0]!.columns![0]!.autoWidth).toBe(true)
  })

  it("emits no <col> at all when a column says nothing", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[1]], columns: [{}] }] })

    expect(await part(bytes, "xl/worksheets/sheet1.xml")).not.toContain("<cols>")
  })
})

describe("sheet format defaults reach the file", () => {
  it("writes and reads back a default row height", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], defaultRowHeight: 24 }],
    })

    expect(await part(bytes, "xl/worksheets/sheet1.xml")).toContain('defaultRowHeight="24"')
    expect((await readXlsx(bytes)).sheets[0]!.defaultRowHeight).toBe(24)
  })

  it("writes and reads back a default column width", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [[1]], defaultColWidth: 18 }],
    })

    expect(await part(bytes, "xl/worksheets/sheet1.xml")).toContain('defaultColWidth="18"')
    expect((await readXlsx(bytes)).sheets[0]!.defaultColWidth).toBe(18)
  })

  it("survives readXlsx → writeXlsx", async () => {
    const first = await writeXlsx({ sheets: [{ name: "S", rows: [[1]], defaultRowHeight: 30 }] })
    const wb = await readXlsx(first)

    const second = await writeXlsx({
      sheets: [
        { name: "S", rows: wb.sheets[0]!.rows, defaultRowHeight: wb.sheets[0]!.defaultRowHeight },
      ],
    })

    expect((await readXlsx(second)).sheets[0]!.defaultRowHeight).toBe(30)
  })

  it("still emits the schema-required attribute when the sheet says nothing", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] })

    expect(await part(bytes, "xl/worksheets/sheet1.xml")).toContain('defaultRowHeight="15"')
  })

  it("does not surface Excel's own default as a setting the author made", async () => {
    // Every file carries defaultRowHeight="15" whether or not it means
    // anything, so reporting it would put a value on every sheet read.
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[1]] }] })

    expect((await readXlsx(bytes)).sheets[0]!.defaultRowHeight).toBeUndefined()
  })
})
