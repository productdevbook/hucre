import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import type { ReadOptions, StreamRow } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// A row the buffered reader puts nothing in is not a row the stream
// yields.
//
// `streamXlsxRows` used to yield every `<row>` element in `sheetData`.
// Excel writes one for each row that ever had a format, a height or a
// hidden flag, often with a self-closing `<c r="B7" s="3"/>` per styled
// position and no value anywhere — and a small Excel-saved workbook was
// reported to stream 722 rows where `read()` returned 121.
// `readXlsx` had already learnt to ignore those cells (#492); the stream
// had not, and `docs/PARITY.md` said it skipped empty rows when it did not.
//
// The rule is the buffered reader's: a row is yielded when one of its
// cells is one `readXlsx` would count toward the bounding box. That is
// decided from the XML, not the resolved value — an out-of-range shared
// string resolves to null and still counts; `<c t="b"/>` resolves to
// `false` and does not.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

function rels(entries: Array<[id: string, type: string, target: string]>): string {
  const items = entries
    .map(
      ([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`,
    )
    .join("")
  return `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`
}

async function workbook(sheetData: string): Promise<Uint8Array> {
  const zip = new ZipWriter()
  const parts: Record<string, string> = {
    "[Content_Types].xml":
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/></Types>`,
    "_rels/.rels": rels([["rId1", "officeDocument", "xl/workbook.xml"]]),
    "xl/workbook.xml": `<?xml version="1.0"?><workbook ${NS} ${R}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheet", "worksheets/sheet1.xml"]]),
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet ${NS} ${R}><sheetData>${sheetData}</sheetData></worksheet>`,
  }
  for (const [path, content] of Object.entries(parts)) zip.add(path, enc.encode(content))
  return zip.build()
}

async function streamed(bytes: Uint8Array, options?: ReadOptions): Promise<StreamRow[]> {
  const out: StreamRow[] = []
  for await (const row of streamXlsxRows(bytes, options)) out.push(row)
  return out
}

/** The rows `readXlsx` has a cell in — what the stream has to yield. */
async function bufferedRows(bytes: Uint8Array, options?: ReadOptions): Promise<number[]> {
  const sheet = (await readXlsx(bytes, { ...options, sparse: true })).sheets[0]!
  const rows = new Set<number>()
  for (const key of sheet.cells?.keys() ?? []) rows.add(Number(key.split(",")[0]))
  return [...rows].sort((a, b) => a - b)
}

const VALUE = (ref: string, v: number): string => `<c r="${ref}"><v>${v}</v></c>`

describe("rows that carry no cell value are not yielded", () => {
  it.each([
    ["a row of style-only cells", `<row r="2"><c r="A2" s="3"/><c r="B2" s="3"/></row>`],
    ["a row with no cells, only a format", `<row r="2" s="2" customFormat="1"/>`],
    ["a row with only a custom height", `<row r="2" ht="30" customHeight="1"/>`],
    ["a hidden row", `<row r="2" hidden="1"/>`],
    ["an empty boolean cell", `<row r="2"><c r="A2" t="b"/></row>`],
    ["an empty formula-string result", `<row r="2"><c r="A2" t="str"><v></v></c></row>`],
    ["an empty `<v/>`", `<row r="2"><c r="A2"><v/></c></row>`],
    [
      "a shared-formula cell with neither formula text nor value",
      `<row r="2"><c r="A2"><f t="shared" si="0"/></c></row>`,
    ],
  ])("%s", async (_label, row2) => {
    const bytes = await workbook(
      `<row r="1">${VALUE("A1", 1)}</row>${row2}<row r="3">${VALUE("A3", 3)}</row>`,
    )

    const rows = await streamed(bytes)

    expect(rows.map((r) => r.index)).toEqual([0, 2])
    expect(rows.map((r) => r.index)).toEqual(await bufferedRows(bytes))
  })
})

describe("rows that carry something are yielded, even when it resolves to null", () => {
  it.each([
    // openpyxl writes formulas with no cached value; `readXlsx` keeps the
    // formula, so the row is not empty to it.
    ["a formula with no cached value", `<row r="2"><c r="A2"><f>A1*3</f><v /></c></row>`, null],
    // No sharedStrings part: the index is out of range and resolves to
    // null, but the producer wrote a string cell.
    ["an out-of-range shared string", `<row r="2"><c r="A2" t="s"><v>7</v></c></row>`, null],
    // An empty inline string is still a string (see #492's carve-out).
    ["an empty inline string", `<row r="2"><c r="A2" t="inlineStr"><is><t></t></is></c></row>`, ""],
    ["an error cell", `<row r="2"><c r="A2" t="e"><v>#N/A</v></c></row>`, "#N/A"],
  ])("%s", async (_label, row2, value) => {
    const bytes = await workbook(`<row r="1">${VALUE("A1", 1)}</row>${row2}`)

    const rows = await streamed(bytes)

    expect(rows.map((r) => r.index)).toEqual([0, 1])
    expect(rows[1]!.values[0]).toEqual(value)
    expect(rows.map((r) => r.index)).toEqual(await bufferedRows(bytes))
  })

  it("keeps the style-only cells of a row that has a value", async () => {
    // Only whether the row is yielded changed; its width did not.
    const rows = await streamed(
      await workbook(`<row r="1">${VALUE("A1", 1)}<c r="C1" s="3"/></row>`),
    )

    expect(rows).toEqual([{ index: 0, values: [1, null, null] }])
  })
})

describe("maxRows and range count only the rows that are yielded", () => {
  const sheet =
    `<row r="1">${VALUE("A1", 1)}</row>` +
    `<row r="2" ht="30" customHeight="1"/>` +
    `<row r="3"><c r="A3" s="3"/></row>` +
    `<row r="4">${VALUE("A4", 4)}</row>` +
    `<row r="5">${VALUE("A5", 5)}</row>`

  it("maxRows does not spend its budget on a skipped row", async () => {
    const rows = await streamed(await workbook(sheet), { maxRows: 2 })

    expect(rows.map((r) => r.index)).toEqual([0, 3])
  })

  it("a row whose only value is outside the range's columns is skipped", async () => {
    // `readXlsx` drops the cells outside the range before deciding a row
    // has content, so the stream has to look at the same columns.
    const bytes = await workbook(
      `<row r="1">${VALUE("A1", 1)}${VALUE("B1", 2)}</row>` +
        `<row r="2">${VALUE("A2", 3)}</row>` +
        `<row r="3">${VALUE("B3", 4)}</row>`,
    )

    const rows = await streamed(bytes, { range: "B1:B3" })

    expect(rows.map((r) => r.index)).toEqual([0, 2])
    expect(rows.map((r) => r.index)).toEqual(await bufferedRows(bytes, { range: "B1:B3" }))
  })

  it("an empty row past the range still ends the read", async () => {
    // The range-end check runs on every `<row>` element, before the
    // empty-row skip. Were it the other way round, the styled row 5 would
    // be skipped without stopping, and row 6 — a column past the sheet's
    // bounds — would be parsed and throw.
    const rows = await streamed(
      await workbook(
        `<row r="1">${VALUE("A1", 1)}</row>` +
          `<row r="5" s="1" customFormat="1"/>` +
          `<row r="6"><c r="AAAAAAA6"><v>1</v></c></row>`,
      ),
      { range: "A1:A2" },
    )

    expect(rows.map((r) => r.index)).toEqual([0])
  })
})
