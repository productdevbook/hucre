import { describe, expect, it } from "vitest"
import type { CellInput, CellValue, WriteModelDrop } from "../src/_types"
import { createCellStore, getCell } from "../src/cell-store"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { XlsxStreamWriter, writeXlsxStreamSheets } from "../src/xlsx/stream-writer"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { OdsStreamWriter } from "../src/ods/incremental-writer"
import { writeOdsStream } from "../src/ods/stream-writer"
import { ZipReader } from "../src/zip/reader"
import { link } from "../src/xlsx/hyperlink"

import { drain, writerPaths as paths } from "./support/writer-paths"

// The old adapters reconstructed only { value, style, formula }, dropping
// an explicit cache and metadata that the shared serializer already knew.
// Literal expectations across every adapter prevent another selective fix.
for (const path of paths) {
  describe(path.name, () => {
    it("accepts the same link helper in positional rows as in object data", async () => {
      const rows: CellInput[][] = [
        [link("Open", "https://example.com/a", "Hint"), link("Internal", "#S!A1")],
      ]
      const before = structuredClone(rows)
      const sheet = (await path.read(await path.write(rows))).sheets[0]
      expect(sheet.rows).toEqual([["Open", "Internal"]])
      expect(getCell(sheet.cells, 0, 0)?.hyperlink?.target).toBe("https://example.com/a")
      const internal = getCell(sheet.cells, 0, 1)?.hyperlink
      expect(internal?.location ?? internal?.target).toBe(
        path.name.includes("XLSX") ? "S!A1" : "#S!A1",
      )
      expect(rows).toEqual(before)
    })
    it("keeps physical columns when the positional array contains a hole", async () => {
      const row: CellInput[] = []
      row[1] = 42
      expect((await path.read(await path.write([row]))).sheets[0].rows).toEqual([[null, 42]])
    })
    for (const [label, cache] of [
      ["number", 42],
      ["string", "cached"],
      ["boolean", false],
      ["empty", null],
      ["error", { error: "#DIV/0!" }],
      ["date", new Date("2024-03-17T00:00:00Z")],
    ] as const) {
      it(`uses the explicit ${label} cache before the scalar value`, async () => {
        const rows: CellInput[][] = [[{ value: 1, formula: "6*7", formulaResult: cache }]]
        const before = structuredClone(rows)
        const sheet = (await path.read(await path.write(rows), { readStyles: true })).sheets[0]
        expect(sheet.rows).toEqual([[cache]])
        expect(getCell(sheet.cells, 0, 0)?.formula).toBe("6*7")
        if (cache !== null) expect(getCell(sheet.cells, 0, 0)?.formulaResult).toEqual(cache)
        expect(rows).toEqual(before)
      })
    }

    it("uses value as the legacy cache only when formulaResult is omitted", async () => {
      const sheet = (await path.read(await path.write([[{ value: 3, formula: "1+2" }]]))).sheets[0]
      expect(sheet.rows).toEqual([[3]])
      expect(getCell(sheet.cells, 0, 0)?.formulaResult).toBe(3)
    })

    it("retains rich-text display content without needing a scalar value", async () => {
      const sheet = (
        await path.read(await path.write([[{ richText: [{ text: " A & " }, { text: "B " }] }]]))
      ).sheets[0]
      expect(sheet.rows).toEqual([[" A & B "]])
    })
  })
}

for (const path of paths.filter((path) => path.name.includes("XLSX"))) {
  describe(`${path.name} inline metadata`, () => {
    it("keeps rich-text comment fonts and whitespace through the shared comment writer", async () => {
      const runs = [{ text: " Bold & ", font: { bold: true } }, { text: " plain " }]
      const bytes = await path.write([
        [{ value: "Cell", comment: { author: "Ada", text: " Bold &  plain ", richText: runs } }],
      ])
      const xml = new TextDecoder().decode(await new ZipReader(bytes).extract("xl/comments1.xml"))
      expect(xml).toContain("<b/>")
      expect(xml).toContain('xml:space="preserve"')
      expect(getCell((await path.read(bytes)).sheets[0].cells, 0, 0)?.comment?.text).toBe(
        " Bold &  plain ",
      )
    })
    it("retains rich-text fonts and XML whitespace", async () => {
      const runs = [{ text: " Bold & ", font: { bold: true } }, { text: " plain " }]
      const sheet = (
        await path.read(await path.write([[{ richText: runs }]]), { readStyles: true })
      ).sheets[0]
      expect(getCell(sheet.cells, 0, 0)?.richText).toEqual(runs)
    })

    it("keeps shared formula masters and followers, array ranges and dynamic metadata parts", async () => {
      const rows: CellInput[][] = [
        [
          {
            formula: "B1+1",
            formulaType: "shared",
            formulaSharedIndex: 7,
            formulaRef: "A1:A2",
            formulaResult: 2,
          },
        ],
        [{ formula: "", formulaType: "shared", formulaSharedIndex: 7, formulaResult: 3 }],
        [
          {
            formula: "SEQUENCE(2)",
            formulaType: "array",
            formulaRef: "A3:A4",
            formulaDynamic: true,
            formulaResult: 1,
          },
        ],
      ]
      const bytes = await path.write(rows)
      const sheet = (await path.read(bytes)).sheets[0]
      expect(getCell(sheet.cells, 0, 0)).toMatchObject({
        formulaType: "shared",
        formulaSharedIndex: 7,
        formulaRef: "A1:A2",
        formulaResult: 2,
      })
      expect(getCell(sheet.cells, 1, 0)).toMatchObject({
        formulaType: "shared",
        formulaSharedIndex: 7,
        formulaResult: 3,
      })
      expect(getCell(sheet.cells, 2, 0)).toMatchObject({
        formulaType: "array",
        formulaRef: "A3:A4",
        formulaDynamic: true,
        formulaResult: 1,
      })
      const zip = new ZipReader(bytes)
      expect(zip.entries()).toContain("xl/metadata.xml")
      const rels = new TextDecoder().decode(await zip.extract("xl/_rels/workbook.xml.rels"))
      expect(rels).toContain('Target="metadata.xml"')
    })

    it("ships the property bag that backs native checkbox styles", async () => {
      const bytes = await path.write([
        [{ value: true, checkbox: true, style: { font: { bold: true } } }],
      ])
      const sheet = (await path.read(bytes, { readStyles: true })).sheets[0]
      expect(getCell(sheet.cells, 0, 0)?.checkbox).toBe(true)
      expect(getCell(sheet.cells, 0, 0)?.style?.font?.bold).toBe(true)
      expect(new ZipReader(bytes).entries()).toContain(
        "xl/featurePropertyBag/featurePropertyBag.xml",
      )
    })

    it("writes external and internal hyperlink relationships and comment parts", async () => {
      const bytes = await path.write([
        [
          {
            value: "External",
            hyperlink: {
              target: "https://example.com/a?x=1&y=2",
              display: "External",
              tooltip: "A & B",
            },
            comment: { author: "Ada", text: "An & observation" },
          },
          { value: "Internal", hyperlink: { target: "", location: "S!A1", display: "Internal" } },
        ],
      ])
      const sheet = (await path.read(bytes)).sheets[0]
      expect(getCell(sheet.cells, 0, 0)?.hyperlink).toMatchObject({
        target: "https://example.com/a?x=1&y=2",
        display: "External",
        tooltip: "A & B",
      })
      expect(getCell(sheet.cells, 0, 1)?.hyperlink).toMatchObject({
        location: "S!A1",
        display: "Internal",
      })
      expect(getCell(sheet.cells, 0, 0)?.comment).toMatchObject({
        author: "Ada",
        text: "An & observation",
      })
      const entries = new ZipReader(bytes).entries()
      expect(entries).toContain("xl/comments1.xml")
      expect(entries).toContain("xl/drawings/vmlDrawing1.vml")
      expect(entries).toContain("xl/worksheets/_rels/sheet1.xml.rels")
    })
  })
}

it("incremental ODS keeps hyperlinks and formatted rich-text runs", async () => {
  const writer = new OdsStreamWriter({ name: "S" })
  writer.addRow([
    { hyperlink: { target: "https://example.com", display: "Open" } },
    { richText: [{ text: "Bold", font: { bold: true } }] },
  ])
  const bytes = await writer.finish()
  const sheet = (await readOds(bytes, { readStyles: true })).sheets[0]
  expect(sheet.rows).toEqual([["Open", "Bold"]])
  expect(getCell(sheet.cells, 0, 0)?.hyperlink?.target).toBe("https://example.com")
  // The ODS reader flattens runs. Verify the emitted font declaration and
  // its span reference instead of claiming a reader rich-text round trip.
  const xml = new TextDecoder().decode(await new ZipReader(bytes).extract("content.xml"))
  expect(xml).toContain('<style:text-properties fo:font-weight="bold"/>')
  expect(xml).toContain('<text:span text:style-name="T1">Bold</text:span>')
})

it("streamed ODS retains hyperlink display content and reports style losses with coordinates", async () => {
  const drops: WriteModelDrop[] = []
  const bytes = await drain(
    writeOdsStream(
      [
        [
          { hyperlink: { target: "https://example.com", display: "Open" } },
          { richText: [{ text: "Bold", font: { bold: true } }], style: { numFmt: "0.00" } },
        ],
      ],
      { name: "S", onDrop: (drop) => drops.push(drop) },
    ),
  )
  const sheet = (await readOds(bytes)).sheets[0]
  expect(sheet.rows).toEqual([["Open", "Bold"]])
  expect(getCell(sheet.cells, 0, 0)?.hyperlink?.target).toBe("https://example.com")
  expect(drops).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ field: "cells.style", sheet: "S", cell: "B1" }),
      expect.objectContaining({ field: "cells.richText.font", sheet: "S", cell: "B1" }),
    ]),
  )
})

it("rollover metadata keeps physical sheet coordinates and repeated headers are snapshots", async () => {
  const cell: Exclude<CellInput, CellValue> = {
    value: "H",
    comment: { text: "Header" },
    hyperlink: { target: "https://example.com/header" },
  }
  const writer = new XlsxStreamWriter({ name: "S", maxRowsPerSheet: 2 })
  writer.addRow([cell])
  cell.comment!.text = "Changed"
  cell.hyperlink!.target = "https://example.com/changed"
  writer.addRow(["First"])
  writer.addRow([{ value: "Second", comment: { text: "Data" } }])
  const bytes = await writer.finish()
  const sheets = (await readXlsx(bytes)).sheets
  expect(getCell(sheets[1].cells, 0, 0)?.comment?.text).toBe("Header")
  expect(getCell(sheets[1].cells, 0, 0)?.hyperlink?.target).toBe("https://example.com/header")
  expect(getCell(sheets[1].cells, 1, 0)?.comment?.text).toBe("Data")
  expect(new ZipReader(bytes).entries()).toContain("xl/comments2.xml")
})

it("multi-sheet and rollover packages keep comment indexes and hyperlink ids independent", async () => {
  const bytes = await drain(
    writeXlsxStreamSheets([
      {
        name: "S",
        maxRowsPerSheet: 2,
        repeatHeaders: false,
        rows: [
          ["Header"],
          ["First"],
          [
            {
              value: "Second",
              comment: { text: "Rollover" },
              hyperlink: { target: "https://example.com/second" },
            },
          ],
        ],
      },
      { name: "Other", rows: [[{ value: "Third", comment: { text: "Other" } }]] },
    ]),
  )
  const sheets = (await readXlsx(bytes)).sheets
  expect(getCell(sheets[1].cells, 0, 0)?.comment?.text).toBe("Rollover")
  expect(getCell(sheets[1].cells, 0, 0)?.hyperlink?.target).toBe("https://example.com/second")
  expect(getCell(sheets[2].cells, 0, 0)?.comment?.text).toBe("Other")
  const types = new TextDecoder().decode(await new ZipReader(bytes).extract("[Content_Types].xml"))
  expect(types).toContain('PartName="/xl/comments2.xml"')
  expect(types).toContain('PartName="/xl/comments3.xml"')
  expect(types).not.toContain('PartName="/xl/comments1.xml"')
})

it("buffered explicit sparse metadata inherits the same inline cache", async () => {
  const rows: CellInput[][] = [[{ value: 3, formula: "1+2" }]]
  const cells = createCellStore([[0, 0, { style: { font: { bold: true } } }]])
  for (const [write, read] of [
    [writeXlsx, readXlsx],
    [writeOds, readOds],
  ] as const) {
    const sheet = (await read(await write({ sheets: [{ name: "S", rows, cells }] }))).sheets[0]
    expect(sheet.rows).toEqual([[3]])
    expect(getCell(sheet.cells, 0, 0)?.formulaResult).toBe(3)
  }
})

it.each([
  ["XLSX", writeXlsx, readXlsx],
  ["ODS", writeOds, readOds],
] as const)(
  "sparse %s caches determine the value and default date format",
  async (_format, write, read) => {
    const date = new Date("2024-03-17T00:00:00Z")
    const cells = createCellStore([
      [
        0,
        0,
        {
          value: 1,
          formula: "DATE(2024,3,17)",
          formulaResult: date,
          style: { font: { bold: true } },
        },
      ],
    ])
    const sheet = (
      await read(await write({ sheets: [{ name: "S", rows: [[1]], cells }] }), { readStyles: true })
    ).sheets[0]
    expect(sheet.rows).toEqual([[date]])
    expect(getCell(sheet.cells, 0, 0)?.formulaResult).toEqual(date)
  },
)

for (const mode of ["buffered", "incremental", "streamed"] as const) {
  it(`${mode} ODS reports every unsupported populated inline cell feature`, async () => {
    const drops: WriteModelDrop[] = []
    const cell: CellInput = {
      value: 1,
      formula: "6*7",
      formulaResult: 42,
      checkbox: true,
      formulaType: "array",
      formulaRef: "A1:A2",
      formulaSharedIndex: 7,
      formulaDynamic: true,
      comment: { text: "Note" },
    }
    const options = { name: "S", onDrop: (drop: WriteModelDrop) => drops.push(drop) }
    let bytes: Uint8Array
    if (mode === "buffered")
      bytes = await writeOds({ sheets: [{ name: "S", rows: [[cell]] }] }, options)
    else if (mode === "streamed") bytes = await drain(writeOdsStream([[cell]], options))
    else {
      const writer = new OdsStreamWriter(options)
      writer.addRow([cell])
      bytes = await writer.finish()
    }
    expect((await readOds(bytes)).sheets[0].rows).toEqual([[42]])
    expect(drops.map((drop) => drop.field).sort()).toEqual([
      "cells.checkbox",
      "cells.comment",
      "cells.formulaDynamic",
      "cells.formulaRef",
      "cells.formulaSharedIndex",
      "cells.formulaType",
    ])
    expect(
      drops.every((drop) => drop.sheet === "S" && drop.cell === "A1" && drop.reason.length > 0),
    ).toBe(true)
  })
}

it("incremental XLSX rejects additions after finish", async () => {
  const writer = new XlsxStreamWriter({ name: "S" })
  writer.addRow(["First"])
  await writer.finish()
  expect(() => writer.addRow(["Second"])).toThrow(/after finish/)
})

it("incremental XLSX emits configured filters and registers conditional formats before styles", async () => {
  const writer = new XlsxStreamWriter({
    name: "S",
    autoFilter: { range: "A1:A2" },
    conditionalRules: [
      {
        range: "A2",
        type: "cellIs",
        priority: 1,
        operator: "greaterThan",
        formula: "0",
        style: { font: { bold: true } },
      },
    ],
  })
  writer.addRow(["Header"])
  writer.addRow([1])
  const sheet = (await readXlsx(await writer.finish(), { readStyles: true })).sheets[0]
  expect(sheet.autoFilter?.range).toBe("A1:A2")
  expect(sheet.conditionalRules?.[0]?.style?.font?.bold).toBe(true)
})
