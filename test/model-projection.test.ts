import { describe, expect, it } from "vitest"
import type { Sheet, Workbook } from "../src/_types"
import { createCellStore } from "../src/cell-store"
import { cellError } from "../src/cell-error"
import { readObjects, write } from "../src/defter"
import { readXlsxObjects } from "../src/xlsx/objects"
import { sheetToArrays, sheetToObjects } from "../src/sheet-utils"
import { toJson } from "../src/export/json"
import { toHtml } from "../src/export/html"
import { toMarkdown } from "../src/export/markdown"
import { workbookToJson } from "../src/json/writer"
import { ParseError } from "../src/errors"
import { audit } from "../src/a11y"
import { parseCsvObjects } from "../src/csv/reader"
import { xlsxWithCells } from "./support/xlsx"

const tableXml =
  '<c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Score</t></is></c><c r="A3" t="inlineStr"><is><t>Ada</t></is></c><c r="B3"><f>6*7</f><v>42</v></c><c r="A5" t="inlineStr"><is><t>Lin</t></is></c><c r="B5"><v>8</v></c>'

function overridden(): Sheet {
  return {
    name: "S",
    rows: [
      ["Old", "Score"],
      ["stale", 1],
      ["erase", 2],
    ],
    cells: createCellStore([
      [0, 0, { value: "Name", type: "string" }],
      [1, 0, { value: "Ada", type: "string", style: { font: { bold: true } } }],
      [1, 1, { value: 1, type: "formula", formula: "6*7", formulaResult: 42 }],
      [2, 0, { value: null, type: "empty" }],
    ]),
  }
}

describe("model projection uses effective cells", () => {
  it("text CSV pads ragged authored rows without changing their input", async () => {
    const rows = [["A", "B"], ["x"]]
    const text = new TextDecoder().decode(
      await write({ sheets: [{ name: "S", rows }] }, { format: "csv" }),
    )
    expect(text).toBe("A,B\r\nx,")
    expect(rows).toEqual([["A", "B"], ["x"]])
  })

  it("HTML checks merge layout before allocating hidden cells", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [["x"]],
      merges: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 2 }],
    }
    expect(() => toHtml(sheet, { maxTotalCells: 8 })).toThrow(ParseError)
    expect(toHtml(sheet, { maxTotalCells: 9 })).toContain('colspan="3" rowspan="3"')
  })
  it("CSV objects count maxRows after the header and bound padded output", () => {
    expect(parseCsvObjects("Name\nAda\nLin", { maxRows: 1 })).toEqual({
      headers: ["Name"],
      data: [{ Name: "Ada" }],
    })
    expect(() => parseCsvObjects("Name,Score\nAda,42", { maxTotalCells: 3 })).toThrow(ParseError)
  })
  it("rejects an oversized unfiltered table before invoking user transforms", () => {
    const headers: string[] = []
    const values: number[] = []
    expect(() =>
      sheetToObjects(overridden(), {
        skipEmptyRows: false,
        maxTotalCells: 5,
        transformHeader: (header) => {
          headers.push(header)
          return header
        },
        transformValue: (value, _header, row) => {
          values.push(row)
          return value
        },
      }),
    ).toThrow(ParseError)
    expect(headers).toEqual([])
    expect(values).toEqual([])
  })

  it("checks only the requested header and data rows before padding sparse gaps", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [],
      cells: createCellStore([
        [3, 0, { value: "Name", type: "string" }],
        [1_048_575, 1, { value: "last", type: "string" }],
      ]),
    }
    expect(sheetToArrays(sheet, { headerRow: 3, maxRows: 1, maxTotalCells: 4 })).toEqual({
      headers: ["Name", ""],
      data: [[null, null]],
    })
    expect(sheetToArrays(sheet, { headerRow: 3, maxRows: 0, maxTotalCells: 2 })).toEqual({
      headers: ["Name", ""],
      data: [],
    })
    expect(() => sheetToArrays(sheet, { headerRow: 3, maxRows: 1, maxTotalCells: 3 })).toThrow(
      ParseError,
    )
  })
  for (const [name, reader] of [
    ["readObjects", readObjects],
    ["readXlsxObjects", readXlsxObjects],
  ] as const) {
    it(`${name} projects sparse values and keeps physical transform indexes`, async () => {
      const indexes: number[] = []
      const result = await reader(await xlsxWithCells(tableXml), {
        sparse: true,
        maxRows: 1,
        transformHeader: (header) => header.toLowerCase(),
        transformValue: (value, _header, row) => {
          indexes.push(row)
          return value
        },
      })
      expect(result).toEqual({ headers: ["name", "score"], data: [{ name: "Ada", score: 42 }] })
      expect(indexes).toEqual([2, 2])
    })
    it(`${name} retains implicit blank rows on request`, async () => {
      const result = await reader(await xlsxWithCells(tableXml), {
        sparse: true,
        skipEmptyRows: false,
        maxRows: 2,
      })
      expect(result.data).toEqual([
        { Name: null, Score: null },
        { Name: "Ada", Score: 42 },
      ])
    })
    it(`${name} bounds the projected cells even when the read is sparse`, async () => {
      await expect(
        reader(await xlsxWithCells(tableXml), { sparse: true, maxTotalCells: 3 }),
      ).rejects.toBeInstanceOf(ParseError)
    })
  }

  it("in-memory object and array projections honor overrides and caches without mutating input", () => {
    const sheet = overridden()
    const before = structuredClone(sheet)
    expect(sheetToObjects(sheet)).toEqual({
      headers: ["Name", "Score"],
      data: [
        { Name: "Ada", Score: 42 },
        { Name: null, Score: 2 },
      ],
    })
    expect(sheetToArrays(sheet)).toEqual({
      headers: ["Name", "Score"],
      data: [
        ["Ada", 42],
        [null, 2],
      ],
    })
    expect(sheet).toEqual(before)
  })

  it("projects far-away rows in coordinate order without a dense bounding box", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [],
      cells: createCellStore([
        [1_048_575, 0, { value: "last", type: "string" }],
        [0, 0, { value: "Name", type: "string" }],
        [10, 0, { value: "first", type: "string" }],
      ]),
    }
    expect(sheetToObjects(sheet)).toEqual({
      headers: ["Name"],
      data: [{ Name: "first" }, { Name: "last" }],
    })
    expect(sheet.rows).toEqual([])
  })

  for (const format of ["objects", "arrays", "columns"] as const) {
    it(`JSON ${format} uses effective values`, () => {
      const expected =
        format === "objects"
          ? [
              { Name: "Ada", Score: 42 },
              { Name: null, Score: 2 },
            ]
          : format === "arrays"
            ? {
                headers: ["Name", "Score"],
                data: [
                  ["Ada", 42],
                  [null, 2],
                ],
              }
            : { Name: ["Ada", null], Score: [42, 2] }
      expect(JSON.parse(toJson(overridden(), { format }))).toEqual(expected)
    })
  }

  it("workbook JSON projects selected and keyed sparse sheets", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [],
      cells: createCellStore([
        [0, 0, { value: "Name", type: "string" }],
        [1, 0, { value: "Ada", type: "string" }],
      ]),
    }
    const wb: Workbook = { sheets: [sheet] }
    expect(JSON.parse(workbookToJson(wb, { sheet: "S" }))).toEqual([{ Name: "Ada" }])
    expect(JSON.parse(workbookToJson(wb, { shape: "sheets" }))).toEqual({ S: [{ Name: "Ada" }] })
  })

  it("HTML and Markdown show effective values with their original styles", () => {
    const sheet = overridden()
    expect(toHtml(sheet, { hasHeaderRow: true })).toContain(">Ada</td>")
    expect(toHtml(sheet)).toContain(">42</td>")
    expect(toHtml(sheet, { styles: true })).toContain('style="font-weight:bold">Ada</td>')
    expect(toMarkdown(sheet)).toContain("Ada")
    expect(toMarkdown(sheet)).not.toContain("stale")
  })

  it("accessibility reports a contiguous sparse gap once", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [],
      cells: createCellStore([
        [0, 0, { value: "Name", type: "string" }],
        [1_048_575, 0, { value: "Ada", type: "string" }],
      ]),
    }
    const gaps = audit({ sheets: [sheet] }).filter((issue) => issue.code === "blank-row-in-data")
    expect(gaps).toHaveLength(1)
    expect(gaps[0]?.location?.ref).toBe("2:1048575")
  })

  it("accessibility uses empty and non-empty formula caches", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [["stale"], [], ["end"]],
      cells: createCellStore([
        [0, 0, { value: "stale", type: "formula", formula: '""', formulaResult: null }],
        [
          2,
          0,
          {
            value: null,
            type: "formula",
            formula: '"end"',
            formulaResult: "end",
            style: {
              font: { color: { rgb: "AAAAAA" } },
              fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFFFF" } },
            },
          },
        ],
      ]),
    }
    const issues = audit({ sheets: [sheet] })
    expect(issues.filter((issue) => issue.code === "blank-row-in-data")).toEqual([])
    expect(issues.find((issue) => issue.code === "low-contrast")?.location?.ref).toBe("A3")
  })

  for (const [name, exporter] of [
    ["arrays", sheetToArrays],
    ["JSON", toJson],
    ["HTML", toHtml],
    ["Markdown", toMarkdown],
  ] as const) {
    it(`${name} rejects a sparse corner before expanding its rectangle`, () => {
      const sheet: Sheet = {
        name: "S",
        rows: [],
        cells: createCellStore([
          [0, 0, { value: "Name", type: "string" }],
          [1_048_575, 16_383, { value: "last", type: "string" }],
        ]),
      }
      expect(() => exporter(sheet)).toThrow(ParseError)
    })
  }

  it("preserves colliding and prototype header names as data in all record projections", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [
        ["__proto__", "Name", "Name"],
        ["safe", "Ada", "Lin"],
      ],
    }
    const record = JSON.parse('{"__proto__":"safe","Name":"Ada","Name_2":"Lin"}')
    expect(sheetToObjects(sheet)).toEqual({
      headers: ["__proto__", "Name", "Name_2"],
      data: [record],
    })
    expect(JSON.parse(toJson(sheet))).toEqual([record])
    expect(JSON.parse(toJson(sheet, { format: "columns" }))).toEqual(
      JSON.parse('{"__proto__":["safe"],"Name":["Ada"],"Name_2":["Lin"]}'),
    )
    expect(JSON.parse(workbookToJson({ sheets: [sheet] }))).toEqual([record])
    expect(Object.getPrototypeOf(sheetToObjects(sheet).data[0])).toBe(Object.prototype)
  })

  it("preserves typed values and treats explicit empty caches as empty rows", () => {
    const date = new Date("2026-10-05T00:00:00Z")
    const error = cellError("#N/A")
    const sheet: Sheet = {
      name: "S",
      rows: [
        ["When", "Error"],
        [1, 2],
        ["stale", null],
      ],
      cells: createCellStore([
        [1, 0, { value: date, type: "date" }],
        [1, 1, { value: error, type: "error" }],
        [2, 0, { value: "stale", type: "formula", formula: '""', formulaResult: null }],
      ]),
    }
    expect(sheetToObjects(sheet).data).toEqual([{ When: date, Error: error }])
    expect(JSON.parse(toJson(sheet))).toEqual([
      { When: date.toISOString(), Error: "#N/A" },
      { When: null, Error: null },
    ])
  })

  for (const format of ["csv", "tsv", "json", "ndjson", "xml", "html", "markdown"] as const) {
    it(`write(${format}) uses sparse values and inline formula caches`, async () => {
      const bytes = await write(
        {
          sheets: [
            {
              name: "S",
              rows: [
                ["Name", "Score"],
                ["stale", { formula: "6*7", formulaResult: 42 }],
              ],
              cells: createCellStore([[1, 0, { value: "Ada", type: "string" }]]),
            },
          ],
        },
        { format },
      )
      const text = new TextDecoder().decode(bytes)
      expect(text).toContain("Ada")
      expect(text).toContain("42")
      expect(text).not.toContain("stale")
    })
  }
})
