import { describe, expect, it } from "vitest"
import type { OutlineProperties } from "../src/_types"
import {
  XlsxStreamWriter,
  writeXlsxStream,
  writeXlsxStreamSheets,
  type XlsxWriteStreamOptions,
} from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

async function sheetXml(bytes: Uint8Array, index = 1): Promise<string> {
  return new TextDecoder().decode(
    await new ZipReader(bytes).extract(`xl/worksheets/sheet${index}.xml`),
  )
}

const writers = [
  {
    name: "writeXlsxStream",
    write: (rows: string[][], options: XlsxWriteStreamOptions) =>
      drain(writeXlsxStream(rows, options)),
  },
  {
    name: "XlsxStreamWriter",
    write: (rows: string[][], options: XlsxWriteStreamOptions) => {
      const writer = new XlsxStreamWriter(options)
      for (const row of rows) writer.addRow(row)
      return writer.finish()
    },
  },
]

// #592: row outline levels survived streaming, but the sheet-level placement
// did not, so parent-first trees were interpreted as summary-below outlines.
describe.each(writers)("$name outline placement", ({ write }) => {
  it.each<{ properties: OutlineProperties; xml: string }>([
    {
      properties: { summaryBelow: false, summaryRight: false },
      xml: '<outlinePr summaryBelow="0" summaryRight="0"/>',
    },
    {
      properties: { summaryBelow: true, summaryRight: true },
      xml: '<outlinePr summaryBelow="1" summaryRight="1"/>',
    },
    { properties: { summaryBelow: false }, xml: '<outlinePr summaryBelow="0"/>' },
    { properties: { summaryRight: false }, xml: '<outlinePr summaryRight="0"/>' },
    { properties: {}, xml: "<outlinePr/>" },
  ])("preserves $properties without inventing omitted attributes", async ({ properties, xml }) => {
    const bytes = await write([["Parent"], ["Child"]], {
      name: "Tree",
      outlineProperties: properties,
      rowDefs: new Map([[1, { outlineLevel: 1 }]]),
      freezePane: { rows: 1 },
    })
    const sheet = (await readXlsx(bytes)).sheets[0]!
    // The reader omits an empty outlinePr; XML below still checks it is emitted.
    expect(sheet.outlineProperties ?? {}).toEqual(properties)
    expect(sheet.rowDefs?.get(1)?.outlineLevel).toBe(1)
    expect(sheet.freezePane).toEqual({ rows: 1 })
    const worksheet = await sheetXml(bytes)
    // sheetPr must precede sheetViews in the OOXML worksheet sequence.
    expect(worksheet).toContain(`<sheetPr>${xml}</sheetPr><sheetViews>`)
  })

  it("keeps the existing default when outlineProperties is omitted", async () => {
    const bytes = await write([["Parent"], ["Child"]], {
      name: "Tree",
      rowDefs: new Map([[1, { outlineLevel: 1 }]]),
    })
    expect((await readXlsx(bytes)).sheets[0]!.outlineProperties).toBeUndefined()
    expect(await sheetXml(bytes)).not.toContain("<sheetPr")
  })

  it("preserves placement on every rollover sheet alongside repeated headers", async () => {
    const properties = { summaryBelow: false, summaryRight: false }
    const bytes = await write([["Parent"], ["Child"], ["Sibling"]], {
      name: "Tree",
      columns: [{ header: "Title" }],
      maxRowsPerSheet: 3,
      outlineProperties: properties,
      rowDefs: new Map([
        [2, { outlineLevel: 1 }],
        [3, { outlineLevel: 1 }],
      ]),
    })
    const workbook = await readXlsx(bytes)
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Tree", "Tree_2"])
    for (const sheet of workbook.sheets) expect(sheet.outlineProperties).toEqual(properties)
    expect(workbook.sheets[1]!.rows).toEqual([["Title"], ["Sibling"]])
    expect(workbook.sheets[1]!.rowDefs?.get(1)?.outlineLevel).toBe(1)
  })
})

it("keeps outline placement local to each logical sheet, including rollover", async () => {
  const bytes = await drain(
    writeXlsxStreamSheets(
      [
        {
          name: "Above",
          rows: [["Parent"], ["Child"], ["Sibling"]],
          outlineProperties: { summaryBelow: false, summaryRight: false },
        },
        {
          name: "Below",
          rows: [["Child"], ["Parent"]],
          outlineProperties: { summaryBelow: true, summaryRight: true },
        },
        { name: "Default", rows: [["Plain"]] },
      ],
      { maxRowsPerSheet: 2, repeatHeaders: false },
    ),
  )
  const workbook = await readXlsx(bytes)
  expect(workbook.sheets.map((sheet) => sheet.name)).toEqual([
    "Above",
    "Above_2",
    "Below",
    "Default",
  ])
  expect(workbook.sheets.map((sheet) => sheet.outlineProperties)).toEqual([
    { summaryBelow: false, summaryRight: false },
    { summaryBelow: false, summaryRight: false },
    { summaryBelow: true, summaryRight: true },
    undefined,
  ])
  expect(await sheetXml(bytes, 4)).not.toContain("<sheetPr")
})
