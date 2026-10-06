import { describe, expect, it } from "vitest"
import type { CellValue, SheetInput } from "../src/_types"
import { write } from "../src/defter"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { getCell, createCellStore } from "../src/cell-store"
import { InvalidArgumentError } from "../src/errors"
import { link } from "../src/xlsx/hyperlink"
import { ParseError } from "../src/errors"
import { ZipReader } from "../src/zip/reader"

const decoder = new TextDecoder()

describe("one authoring row source", () => {
  it("checks object-data text output bounds before reading any data value", async () => {
    let reads = 0
    const sheet: SheetInput = {
      name: "S",
      columns: [{ key: "name", header: "Name" }],
      data: [
        {
          get name() {
            reads++
            return "Ada"
          },
        },
      ],
    }
    await expect(
      write({ sheets: [sheet] }, { format: "csv", maxTotalCells: 1 }),
    ).rejects.toBeInstanceOf(ParseError)
    expect(reads).toBe(0)
  })

  it("pivots use effective authored headers and cached values from their worksheet", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "Source",
          data: [{ Region: "EU", Amount: 1 }],
          cells: createCellStore([
            [0, 1, { value: "Revenue" }],
            [1, 1, { formula: "6*7", formulaResult: 42 }],
          ]),
        },
        {
          name: "Pivot",
          pivotTables: [
            { name: "P", sourceSheet: "Source", rows: ["Region"], values: [{ field: "Revenue" }] },
          ],
        },
      ],
    })
    expect((await readXlsx(bytes)).sheets[0].rows).toEqual([
      ["Region", "Revenue"],
      ["EU", 42],
    ])
    const zip = new ZipReader(bytes)
    const cache = decoder.decode(await zip.extract("xl/pivotCache/pivotCacheDefinition1.xml"))
    const records = decoder.decode(await zip.extract("xl/pivotCache/pivotCacheRecords1.xml"))
    expect(cache).toContain('<cacheField name="Revenue"')
    expect(records).toContain('<n v="42"/>')
    expect(records).not.toContain('<n v="1"/>')
  })

  for (const format of ["csv", "tsv", "json", "ndjson", "xml", "html", "markdown"] as const) {
    it(`${format} writes object data and link display values without mutating input`, async () => {
      const sheet: SheetInput = {
        name: "S",
        columns: [
          { key: "name", header: "Name" },
          { key: "score", header: "Score" },
        ],
        data: [{ name: link("Ada", "https://example.com/ada"), score: 42 }],
      }
      const before = structuredClone(sheet)
      const text = decoder.decode(await write({ sheets: [sheet] }, { format }))
      expect(text).toContain("Name")
      expect(text).toContain("Ada")
      expect(text).toContain("42")
      expect(text).not.toContain("[object Object]")
      if (format === "json" || format === "ndjson") {
        expect(JSON.parse(text)).toEqual(
          format === "json" ? [{ Name: "Ada", Score: 42 }] : { Name: "Ada", Score: 42 },
        )
      }
      expect(sheet).toEqual(before)
    })
  }

  for (const format of [
    "xlsx",
    "ods",
    "csv",
    "tsv",
    "json",
    "ndjson",
    "xml",
    "html",
    "markdown",
  ] as const) {
    it(`${format} rejects two competing row sources`, async () => {
      const sheet: SheetInput = {
        name: "S",
        rows: [["stale"]],
        columns: [{ key: "name", header: "Name" }],
        data: [{ name: "Ada" }],
      }
      const before = structuredClone(sheet)
      await expect(write({ sheets: [sheet] }, { format })).rejects.toBeInstanceOf(
        InvalidArgumentError,
      )
      expect(sheet).toEqual(before)
    })
  }

  for (const [format, writer, reader] of [
    ["xlsx", writeXlsx, readXlsx],
    ["ods", writeOds, readOds],
  ] as const) {
    it(`${format} infers every own key in first-seen order`, async () => {
      const sheet: SheetInput = { name: "S", data: [{ a: 1 }, { b: 2 }, { a: 3, c: 4 }] }
      const before = structuredClone(sheet)
      expect((await reader(await writer({ sheets: [sheet] }))).sheets[0].rows).toEqual([
        ["a", "b", "c"],
        [1, null, null],
        [null, 2, null],
        [3, null, 4],
      ])
      expect(sheet).toEqual(before)
    })

    it(`${format} resolves header-only columns consistently`, async () => {
      const sheet: SheetInput = {
        name: "S",
        columns: [{ header: "Name" }],
        data: [{ Name: "Ada" }],
      }
      expect((await reader(await writer({ sheets: [sheet] }))).sheets[0].rows).toEqual([
        ["Name"],
        ["Ada"],
      ])
    })

    it(`${format} keeps an explicitly blank header`, async () => {
      const sheet: SheetInput = {
        name: "S",
        columns: [{ key: "a", header: "" }, { key: "b" }],
        data: [{ a: 1, b: 2 }],
      }
      expect((await reader(await writer({ sheets: [sheet] }))).sheets[0].rows).toEqual([
        ["", "b"],
        [1, 2],
      ])
    })

    it(`${format} retains an unnamed column as a blank header and null data`, async () => {
      const sheet: SheetInput = {
        name: "S",
        columns: [{ key: "a", header: "A" }, { key: "b" }, { width: 10 }],
        data: [{ a: 1, b: 2 }],
      }
      expect((await reader(await writer({ sheets: [sheet] }))).sheets[0].rows).toEqual([
        ["A", "b", ""],
        [1, 2, null],
      ])
    })

    it(`${format} reads only own data fields even with explicit columns`, async () => {
      const record: Record<string, CellValue> = Object.assign(Object.create({ secret: 7 }), {
        safe: 1,
      })
      const sheet: SheetInput = {
        name: "S",
        columns: [
          { key: "safe", header: "Safe" },
          { key: "secret", header: "Secret" },
        ],
        data: [record],
      }
      expect((await reader(await writer({ sheets: [sheet] }))).sheets[0].rows).toEqual([
        ["Safe", "Secret"],
        [1, null],
      ])
    })

    it(`${format} keeps data hyperlinks under partial metadata overlays`, async () => {
      const sheet: SheetInput = {
        name: "S",
        columns: [{ key: "url", header: "Link" }],
        data: [{ url: link("Open", "https://example.com/item", "Details") }],
        cells: createCellStore([[1, 0, { style: { font: { bold: true } } }]]),
      }
      const result = (await reader(await writer({ sheets: [sheet] }), { readStyles: true }))
        .sheets[0]
      expect(result.rows).toEqual([["Link"], ["Open"]])
      expect(getCell(result.cells, 1, 0)?.hyperlink?.target).toBe("https://example.com/item")
      expect(getCell(result.cells, 1, 0)?.style?.font?.bold).toBe(true)
    })

    it(`${format} applies column styles to array rows, including trailing null cells`, async () => {
      const sheet: SheetInput = {
        name: "S",
        rows: [[1, null]],
        columns: [{ style: { font: { bold: true } } }, { numFmt: "0.00" }],
      }
      const result = (await reader(await writer({ sheets: [sheet] }), { readStyles: true }))
        .sheets[0]
      expect(getCell(result.cells, 0, 0)?.style?.font?.bold).toBe(true)
      expect(getCell(result.cells, 0, 1)?.style?.numFmt).toBe("0.00")
    })

    it(`${format} applies column styles to data but exempts generated headers from column numFmt`, async () => {
      const sheet: SheetInput = {
        name: "S",
        columns: [
          { key: "amount", header: "Amount", numFmt: "0.00", style: { font: { bold: true } } },
        ],
        data: [{ amount: 12.5 }],
      }
      const result = (await reader(await writer({ sheets: [sheet] }), { readStyles: true }))
        .sheets[0]
      expect(getCell(result.cells, 0, 0)?.style?.font?.bold).toBe(true)
      expect(getCell(result.cells, 0, 0)?.style?.numFmt ?? "General").toBe("General")
      expect(getCell(result.cells, 1, 0)?.style?.numFmt).toBe("0.00")
    })
  }
})
