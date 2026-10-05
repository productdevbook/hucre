import { describe, expect, it } from "vitest"
import type { Cell, CellValue, Sheet, Workbook } from "../src/_types"
import { createCellStore, getCell } from "../src/cell-store"
import { cellError } from "../src/cell-error"
import { findCells, replaceCells } from "../src/sheet-ops"
import { fillTemplate } from "../src/template"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"

function sheet(value: CellValue = "old", metadata: Partial<Cell> = {}): Sheet {
  return {
    name: "S",
    rows: [[value]],
    cells: createCellStore([
      [
        0,
        0,
        {
          value,
          type: "string",
          style: { font: { bold: true } },
          ...metadata,
        },
      ],
    ]),
  }
}

describe("one value editing contract", () => {
  it.each(["xlsx", "ods"] as const)(
    "%s authoring respects an explicit null override",
    async (format) => {
      const s = sheet(42, { value: null, type: "empty" })
      const workbook = { sheets: [s] }
      const result =
        format === "xlsx"
          ? await readXlsx(await writeXlsx(workbook), { readStyles: true })
          : await readOds(await writeOds(workbook), { readStyles: true })
      expect(result.sheets[0].rows).toEqual([[null]])
    },
  )
  it.each(["xlsx", "ods"] as const)(
    "%s partial metadata retains an inline formula and its value",
    async (format) => {
      const workbook = {
        sheets: [
          {
            name: "S",
            rows: [[{ value: 7, formula: "1+6", formulaResult: 7 }]],
            cells: createCellStore<Partial<Cell>>([
              [0, 0, { style: { font: { bold: true } }, value: undefined }],
            ]),
          },
        ],
      }
      const result =
        format === "xlsx"
          ? await readXlsx(await writeXlsx(workbook), { readStyles: true })
          : await readOds(await writeOds(workbook), { readStyles: true })
      const s = result.sheets[0]
      expect(s.rows).toEqual([[7]])
      expect(getCell(s.cells, 0, 0)).toMatchObject({
        formula: "1+6",
        formulaResult: 7,
        style: { font: { bold: true } },
      })
    },
  )
  it.each([
    ["number", 7],
    ["boolean", false],
    ["date", new Date("2026-01-02T00:00:00Z")],
    ["error", cellError("#N/A")],
    ["empty", null],
  ] as const)("replacement keeps %s metadata consistent", (type, value) => {
    const s = sheet()
    expect(replaceCells(s, "old", value)).toBe(1)
    expect(getCell(s.cells, 0, 0)).toMatchObject({ value, type, style: { font: { bold: true } } })
    expect(s.rows).toEqual([[value]])
  })

  it.each(["replace", "template"] as const)(
    "%s updates the cached result without erasing its formula",
    (mode) => {
      const original = mode === "replace" ? "old" : "{{value}}"
      const s = sheet(original, { type: "formula", formula: '"old"', formulaResult: original })
      if (mode === "replace") replaceCells(s, "old", 7)
      else fillTemplate({ sheets: [s] }, { value: 7 })
      expect(getCell(s.cells, 0, 0)).toMatchObject({
        value: 7,
        type: "formula",
        formula: '"old"',
        formulaResult: 7,
      })
      expect(s.rows).toEqual([[7]])
    },
  )

  it("finds and replaces populated sparse cells without building the rectangle", () => {
    const s: Sheet = {
      name: "Sparse",
      rows: [],
      cells: createCellStore([
        [127, 3, { value: "old", type: "string" }],
        [1048575, 16383, { value: "old", type: "string" }],
      ]),
    }
    expect(findCells(s, "old")).toEqual([
      { row: 127, col: 3, value: "old" },
      { row: 1048575, col: 16383, value: "old" },
    ])
    expect(replaceCells(s, "old", false)).toBe(2)
    expect(getCell(s.cells, 1048575, 16383)).toEqual({ value: false, type: "boolean" })
    expect(s.rows).toEqual([])
  })

  it("uses an explicit value override once and synchronizes a materialized row", () => {
    const s = sheet("dense", { value: "metadata" })
    expect(findCells(s, "metadata")).toEqual([{ row: 0, col: 0, value: "metadata" }])
    expect(replaceCells(s, "metadata", 7)).toBe(1)
    expect(s.rows).toEqual([[7]])
    expect(getCell(s.cells, 0, 0)?.value).toBe(7)
  })

  it("does not let the caller's global regexp state skip replacements", () => {
    const s: Sheet = { name: "S", rows: [["old", "old"]] }
    const regex = /old/g
    regex.lastIndex = 3
    expect(replaceCells(s, regex, "new")).toBe(2)
    expect(s.rows).toEqual([["new", "new"]])
  })

  it("a null template substitution becomes an empty cell", () => {
    const s = sheet("{{value}}")
    fillTemplate({ sheets: [s] }, { value: null })
    expect(getCell(s.cells, 0, 0)?.type).toBe("empty")
  })

  it("fills each coordinate once when inserted text contains another placeholder", () => {
    const s = sheet("{{first}}")
    fillTemplate({ sheets: [s] }, { first: "{{second}}", second: "incorrect second pass" })
    expect(s.rows).toEqual([["{{second}}"]])
    expect(getCell(s.cells, 0, 0)?.value).toBe("{{second}}")
  })

  it.each(["authoring", "preserved"] as const)(
    "%s save emits the edited rich-text value and formula cache",
    async (mode) => {
      const s = sheet("old", {
        type: "richText",
        richText: [{ text: "old", font: { italic: true } }],
      })
      s.rows[0].push("old")
      s.cells = createCellStore([
        [0, 0, getCell(s.cells, 0, 0)!],
        [0, 1, { value: "old", type: "formula", formula: '"old"', formulaResult: "old" }],
      ])
      const source: Workbook = { sheets: [s] }
      const opened = mode === "preserved" ? await openXlsx(await writeXlsx(source)) : undefined
      const workbook = opened ?? source
      expect(replaceCells(workbook.sheets[0], "old", "new")).toBe(2)
      expect(getCell(workbook.sheets[0].cells, 0, 0)?.richText).toBeUndefined()
      const bytes = opened ? await saveXlsx(opened) : await writeXlsx(workbook)
      const again = (await readXlsx(bytes, { readStyles: true })).sheets[0]
      expect(again.rows).toEqual([["new", "new"]])
      expect(getCell(again.cells, 0, 0)?.style?.font?.bold).toBe(true)
      expect(getCell(again.cells, 0, 1)).toMatchObject({ formula: '"old"', formulaResult: "new" })
    },
  )
})
