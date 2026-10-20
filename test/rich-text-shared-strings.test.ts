import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { ZipReader } from "../src/zip/reader"
import type { Cell, RichTextRun, WriteOptions, WriteSheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// Rich text ignored `stringMode`. Plain text went to xl/sharedStrings.xml
// in the default "shared" mode, but a rich text cell was always written
// inline:
//
//   <c r="A2" t="inlineStr"><is><r><rPr><b/></rPr><t>Bold</t></r>…</is></c>
//
// That is valid, but anything that only reads the shared string table —
// template engines, find/replace, string indexing — never saw the rich
// cells, and it is not what Excel writes. In shared mode a rich string is
// now an `<si>` holding the same `<r>` runs (ECMA-376 Part 1 §18.4.8) and
// the cell is `t="s"`. Inline mode is unchanged. See #602.
// ═══════════════════════════════════════════════════════════════════════

const decoder = new TextDecoder("utf-8")

async function part(buf: Uint8Array, path: string): Promise<string> {
  return decoder.decode(await new ZipReader(buf).extract(path))
}

const BOLD_RUNS: RichTextRun[] = [{ text: "Bold", font: { bold: true } }, { text: " normal" }]
const BOLD_RUNS_XML = '<r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> normal</t></r>'

function write(
  cells: Array<[string, Partial<Cell>]>,
  rows: WriteSheet["rows"] = [],
  options: Omit<WriteOptions, "sheets"> = {},
): Promise<Uint8Array> {
  return writeXlsx({ sheets: [{ name: "S", rows, cells: new Map(cells) }], ...options })
}

describe("rich text in shared string mode (#602)", () => {
  it('writes the runs into the shared string table and the cell as t="s"', async () => {
    const buf = await write([["1,0", { richText: BOLD_RUNS }]], [["Plain"]])

    const sheet = await part(buf, "xl/worksheets/sheet1.xml")
    expect(sheet).toContain('<c r="A2" t="s"><v>1</v></c>')
    expect(sheet).not.toContain("inlineStr")

    const sst = await part(buf, "xl/sharedStrings.xml")
    expect(sst).toContain('count="2" uniqueCount="2"')
    expect(sst).toContain(`<si><t>Plain</t></si><si>${BOLD_RUNS_XML}</si>`)
  })

  it("keeps the style index on a shared rich text cell", async () => {
    const buf = await write([["0,0", { richText: BOLD_RUNS, style: { font: { italic: true } } }]])

    const sheet = await part(buf, "xl/worksheets/sheet1.xml")
    expect(sheet).toMatch(/<c r="A1" t="s" s="[1-9]\d*"><v>0<\/v><\/c>/)
  })

  it("shares one entry between identical rich strings, apart from plain text", async () => {
    const buf = await write([
      ["0,0", { richText: BOLD_RUNS }],
      ["0,1", { value: "Bold normal" }],
      ["0,2", { richText: [{ text: "Bold", font: { bold: true } }, { text: " normal" }] }],
      ["0,3", { richText: [{ text: "Bold", font: { italic: true } }, { text: " normal" }] }],
    ])

    const sheet = await part(buf, "xl/worksheets/sheet1.xml")
    expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>')
    // Same text, no runs: a different string as far as the table goes.
    expect(sheet).toContain('<c r="B1" t="s"><v>1</v></c>')
    // Equal runs, separate array: the same entry.
    expect(sheet).toContain('<c r="C1" t="s"><v>0</v></c>')
    // Same text, different formatting: a new entry.
    expect(sheet).toContain('<c r="D1" t="s"><v>2</v></c>')

    const sst = await part(buf, "xl/sharedStrings.xml")
    expect(sst).toContain('count="3" uniqueCount="3"')
  })

  it("writes the shared string part when rich text is the only text", async () => {
    const buf = await write([["0,0", { richText: BOLD_RUNS }]])
    const zip = new ZipReader(buf)

    expect(zip.has("xl/sharedStrings.xml")).toBe(true)
    expect(await part(buf, "[Content_Types].xml")).toContain('PartName="/xl/sharedStrings.xml"')
    expect(await part(buf, "xl/_rels/workbook.xml.rels")).toContain('Target="sharedStrings.xml"')
    expect(await part(buf, "xl/sharedStrings.xml")).toContain(`<si>${BOLD_RUNS_XML}</si>`)
  })

  it("reads the rich text back from the shared string table", async () => {
    const buf = await write([["0,0", { richText: BOLD_RUNS }]])

    const cell = (await readXlsx(buf)).sheets[0].cells?.get("0,0")
    expect(cell?.type).toBe("richText")
    expect(cell?.value).toBe("Bold normal")
    expect(cell?.richText).toEqual(BOLD_RUNS)
  })

  it("leaves inline mode as it was", async () => {
    const buf = await write([["1,0", { richText: BOLD_RUNS }]], [["Plain"]], {
      stringMode: "inline",
    })

    const sheet = await part(buf, "xl/worksheets/sheet1.xml")
    expect(sheet).toContain(`<c r="A2" t="inlineStr"><is>${BOLD_RUNS_XML}</is></c>`)
    expect(new ZipReader(buf).has("xl/sharedStrings.xml")).toBe(false)
  })

  it("saveXlsx writes rich text into the shared string table too", async () => {
    // saveXlsx always writes shared strings, so a rich cell read from an
    // inline-only file moves into the table.
    const original = await write([["0,0", { richText: BOLD_RUNS }]], [], { stringMode: "inline" })
    const saved = await saveXlsx(await openXlsx(original))

    expect(await part(saved, "xl/worksheets/sheet1.xml")).toContain('<c r="A1" t="s"><v>0</v></c>')
    expect(await part(saved, "xl/sharedStrings.xml")).toContain(`<si>${BOLD_RUNS_XML}</si>`)

    const cell = (await readXlsx(saved)).sheets[0].cells?.get("0,0")
    expect(cell?.richText).toEqual(BOLD_RUNS)
  })
})
