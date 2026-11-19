import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import type { ReadOptions, StreamRow } from "../src/_types"
import { bufferedRowIndexes } from "./_xlsx-rows"

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

/** Optional parts: a `styles.xml` and a `sharedStrings.xml`, given as their inner XML. */
interface Extras {
  cellXfs?: string
  sharedStrings?: string[]
}

async function workbook(sheetData: string, extras: Extras = {}): Promise<Uint8Array> {
  const zip = new ZipWriter()
  const workbookRels: Array<[string, string, string]> = [
    ["rId1", "worksheet", "worksheets/sheet1.xml"],
  ]
  const parts: Record<string, string> = {
    "[Content_Types].xml":
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/></Types>`,
    "_rels/.rels": rels([["rId1", "officeDocument", "xl/workbook.xml"]]),
    "xl/workbook.xml": `<?xml version="1.0"?><workbook ${NS} ${R}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet ${NS} ${R}><sheetData>${sheetData}</sheetData></worksheet>`,
  }
  if (extras.cellXfs !== undefined) {
    workbookRels.push(["rId2", "styles", "styles.xml"])
    parts["xl/styles.xml"] =
      `<?xml version="1.0"?><styleSheet ${NS}><cellXfs>${extras.cellXfs}</cellXfs></styleSheet>`
  }
  if (extras.sharedStrings !== undefined) {
    workbookRels.push(["rId3", "sharedStrings", "sharedStrings.xml"])
    const items = extras.sharedStrings.map((t) => `<si><t>${t}</t></si>`).join("")
    parts["xl/sharedStrings.xml"] = `<?xml version="1.0"?><sst ${NS}>${items}</sst>`
  }
  parts["xl/_rels/workbook.xml.rels"] = rels(workbookRels)
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
  return bufferedRowIndexes((await readXlsx(bytes, { ...options, sparse: true })).sheets[0]!)
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

describe("every arm of the shared cellCarriesData is held by a test", () => {
  // Both readers call `cellCarriesData` in `worksheet.ts`. A comparison
  // of one reader against the other cannot see an arm deleted from it —
  // both move together — so these pin the expected rows literally.

  // xf 1 carries the Excel 2024 checkbox feature; xf 0 is plain.
  const CHECKBOX_XFS =
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0"><extLst>` +
    `<ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag">` +
    `<xfpb:xfComplement i="0"/></ext></extLst></xf>`

  it("a style-only cell on a checkbox xf yields its row; on a plain xf it does not", async () => {
    // An unchecked checkbox Excel writes as `<c r="A2" s="1"/>`: no value,
    // and the checkbox is the whole content. `readXlsx` keeps it as
    // `{ value: null, checkbox: true }`.
    const bytes = await workbook(
      `<row r="1">${VALUE("A1", 1)}</row>` +
        `<row r="2"><c r="A2" s="1"/></row>` +
        `<row r="3"><c r="A3" s="0"/></row>`,
      { cellXfs: CHECKBOX_XFS },
    )

    const rows = await streamed(bytes)

    expect(rows).toEqual([
      { index: 0, values: [1] },
      { index: 1, values: [null] },
    ])
    expect(await bufferedRows(bytes)).toEqual([0, 1])
  })

  it("an inline run with no text and no t yields its row", async () => {
    // Not subsumed by the `t="inlineStr"` arm: `<is>` is read whatever
    // `t` says, and this cell has a run, no text, and no inline type.
    // Only the run count says the producer wrote something here.
    const bytes = await workbook(
      `<row r="1">${VALUE("A1", 1)}</row><row r="2"><c r="A2"><is><r><t/></r></is></c></row>`,
    )

    const rows = await streamed(bytes)

    expect(rows.map((r) => r.index)).toEqual([0, 1])
    expect(await bufferedRows(bytes)).toEqual([0, 1])
  })
})

describe("a cell readXlsx does not count has no value in the stream either", () => {
  // In a row that has other data, these used to come out with a value
  // where `readXlsx` had nothing: the row was yielded for its other cell,
  // and the empty one was resolved as if it had text.
  it.each([
    ["an empty boolean cell", `<c r="B1" t="b"/>`],
    ["an empty formula-string result", `<c r="B1" t="str"><v></v></c>`],
    ["a formula-string cell with no <v>", `<c r="B1" t="str"/>`],
    // `Number("")` is 0, so this resolved to the first shared string.
    ["a shared-string cell with no <v>", `<c r="B1" t="s"/>`],
  ])("%s", async (_label, cell) => {
    const bytes = await workbook(`<row r="1">${VALUE("A1", 1)}${cell}${VALUE("C1", 3)}</row>`, {
      sharedStrings: ["first"],
    })

    const rows = await streamed(bytes)
    const buffered = (await readXlsx(bytes)).sheets[0]!.rows

    expect(rows).toEqual([{ index: 0, values: [1, null, 3] }])
    expect(buffered).toEqual([[1, null, 3]])
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
