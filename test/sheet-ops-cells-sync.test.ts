import { createCellStore, setCell, getCell } from "../src/cell-store"
import { describe, expect, it } from "vitest"
import type { Cell, Sheet } from "../src/_types"
import { replaceCells, sortRows } from "../src/sheet-ops"

describe("replaceCells — keeps cells override Map in sync", () => {
  it("updates the override value when the row value is replaced", () => {
    const cells = createCellStore<Cell>()
    setCell(cells, 0, 0, { type: "string", value: "old", style: { font: { bold: true } } })
    const sheet: Sheet = { name: "S", rows: [["old"]], cells }

    const n = replaceCells(sheet, "old", "new")
    expect(n).toBe(1)
    expect(sheet.rows[0][0]).toBe("new")
    // The styled override must not keep the stale value.
    expect(getCell(sheet.cells!, 0, 0)!.value).toBe("new")
    expect(getCell(sheet.cells!, 0, 0)!.style?.font?.bold).toBe(true)
  })
})

describe("sortRows — remaps cells override Map to new row positions", () => {
  it("moves a styled cell with its row", () => {
    const cells = createCellStore<Cell>()
    // Style attached to the row that currently holds value 3 (row 0).
    setCell(cells, 0, 0, { type: "number", value: 3, style: { font: { bold: true } } })
    const sheet: Sheet = { name: "S", rows: [[3], [1], [2]], cells }

    sortRows(sheet, 0, "asc")

    // Rows are now 1,2,3 — the styled cell (value 3) is at row index 2.
    expect(sheet.rows.map((r) => r[0])).toEqual([1, 2, 3])
    expect(getCell(sheet.cells!, 2, 0)?.style?.font?.bold).toBe(true)
    expect(getCell(sheet.cells!, 0, 0)?.style?.font?.bold).toBeUndefined()
  })
})
