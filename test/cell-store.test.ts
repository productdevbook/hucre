import { describe, expect, it } from "vitest"
import { createCellStore, getCell, setCell, hasCell, deleteCell, cellEntries } from "../src/cell"
import { InvalidArgumentError } from "../src/errors"
import type { Cell } from "../src/_types"

const value = (n: number): Cell => ({ value: n, type: "number" })

describe("portable sparse cell metadata", () => {
  it("creates, overwrites and removes metadata with an accurate cell count", () => {
    const cells = createCellStore([
      [0, 0, value(1)],
      [0, 0, value(2)],
    ])
    expect(cells.size).toBe(1)
    expect(getCell(cells, 0, 0)?.value).toBe(2)
    setCell(cells, 129, 2, value(3))
    expect(cells.size).toBe(2)
    expect(hasCell(cells, 129, 2)).toBe(true)
    expect(deleteCell(cells, 0, 0)).toBe(true)
    expect(deleteCell(cells, 0, 0)).toBe(false)
    expect(cells.size).toBe(1)
    expect(deleteCell(cells, 129, 2)).toBe(true)
    expect(cells.size).toBe(0)
    expect(cells.blocks.size).toBe(0)
  })

  it("addresses block transitions and every Excel corner without aliasing", () => {
    const entries: Array<[number, number, Cell]> = [
      [0, 16_383, value(1)],
      [1, 0, value(2)],
      [127, 16_383, value(3)],
      [128, 0, value(4)],
      [1_048_575, 16_383, value(5)],
    ]
    const cells = createCellStore(entries)
    expect([...cellEntries(cells)]).toEqual(entries)
    for (const [row, col, cell] of entries) expect(getCell(cells, row, col)).toBe(cell)
  })

  it("keeps partial authoring overrides, including empty ones", () => {
    const cells = createCellStore([[0, 2, {}]])
    expect(hasCell(cells, 0, 2)).toBe(true)
    expect(getCell(cells, 0, 1)).toBeUndefined()
    expect(hasCell(undefined, 0, 1)).toBe(false)
    expect([...cellEntries(undefined)]).toEqual([])
  })

  it("survives structuredClone with dates and independent nested styles", () => {
    const cells = createCellStore<Cell>([
      [
        256,
        3,
        {
          value: new Date("2026-10-04T00:00:00.000Z"),
          type: "date",
          style: { font: { bold: true } },
        },
      ],
    ])
    const cloned = structuredClone(cells)
    expect(cloned).toEqual(cells)
    getCell(cloned, 256, 3)!.style!.font!.bold = false
    expect(getCell(cells, 256, 3)!.style!.font!.bold).toBe(true)
    expect(getCell(cloned, 256, 3)!.value).toBeInstanceOf(Date)
    setCell(cloned, 257, 3, value(4))
    expect(cells.size).toBe(1)
    expect(cloned.size).toBe(2)
  })

  it.each([
    [-1, 0],
    [0, -1],
    [1_048_576, 0],
    [0, 16_384],
    [0.5, 0],
    [0, 0.5],
    [NaN, 0],
    [0, Infinity],
  ])("refuses invalid writes and never aliases invalid lookups (%s,%s)", (row, col) => {
    const cells = createCellStore([
      [0, 16_383, value(1)],
      [1, 0, value(2)],
    ])
    expect(() => setCell(cells, row, col, value(3))).toThrow(InvalidArgumentError)
    expect(getCell(cells, row, col)).toBeUndefined()
    expect(deleteCell(cells, row, col)).toBe(false)
    expect(cells.size).toBe(2)
    expect(getCell(cells, 0, 16_383)?.value).toBe(1)
    expect(getCell(cells, 1, 0)?.value).toBe(2)
  })
})
