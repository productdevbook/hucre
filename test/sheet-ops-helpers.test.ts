import { describe, expect, it } from "vitest"
import { deleteColumns, deleteRows, findCells, replaceCells, sortRows } from "../src/sheet-ops"
import { InvalidArgumentError } from "../src/errors"
import type { Cell, Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §AD, §AE, §AK — three small disagreements in the sheet helpers.
// ═══════════════════════════════════════════════════════════════════════

describe("sortRows moves everything that is keyed by row", () => {
  function sheet(): Sheet {
    const cells = new Map<string, Cell>([
      ["0,0", { value: 3, type: "number", style: { font: { bold: true } } }],
    ])
    return {
      name: "S",
      rows: [[3], [1], [2]],
      cells,
      rowDefs: new Map([
        [0, { height: 99 }],
        [2, { hidden: true }],
      ]),
    }
  }

  it("takes the row definitions with their rows", () => {
    const s = sheet()

    sortRows(s, 0, "asc")

    // 3 was row 0 and is now row 2; 2 was row 2 and is now row 1.
    expect(s.rows).toEqual([[1], [2], [3]])
    expect(s.rowDefs!.get(2)).toEqual({ height: 99 })
    expect(s.rowDefs!.get(1)).toEqual({ hidden: true })
    expect(s.rowDefs!.get(0)).toBeUndefined()
  })

  it("takes the cell overrides too, as it always did", () => {
    const s = sheet()

    sortRows(s, 0, "asc")

    expect([...s.cells!.keys()]).toEqual(["2,0"])
    expect(s.cells!.get("2,0")!.style!.font!.bold).toBe(true)
  })

  it("carries a single-row merge with its row", () => {
    const s: Sheet = {
      name: "S",
      rows: [
        [3, "a"],
        [1, "b"],
      ],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
    }

    sortRows(s, 0, "asc")

    expect(s.rows).toEqual([
      [1, "b"],
      [3, "a"],
    ])
    expect(s.merges).toEqual([{ startRow: 1, startCol: 0, endRow: 1, endCol: 1 }])
  })

  it("refuses to sort past a merge that spans rows", () => {
    const s: Sheet = {
      name: "S",
      rows: [[3], [1], [2]],
      merges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 0 }],
    }

    // Excel refuses the same operation for the same reason: no ordering
    // keeps both the sort and the merge.
    expect(() => sortRows(s, 0, "asc")).toThrow(InvalidArgumentError)
    expect(s.rows).toEqual([[3], [1], [2]])
  })

  it("still sorts a sheet with no overrides at all", () => {
    const s: Sheet = { name: "S", rows: [[3], [1], [2]] }

    sortRows(s, 0, "desc")

    expect(s.rows).toEqual([[3], [2], [1]])
  })
})

describe("findCells takes the same forms as replaceCells", () => {
  const sheet = (): Sheet => ({
    name: "S",
    rows: [
      ["b1", "x"],
      ["B2", 7],
    ],
  })

  it("accepts a RegExp", () => {
    const found = findCells(sheet(), /^b/)

    expect(found).toEqual([{ row: 0, col: 0, value: "b1" }])
  })

  it("finds exactly what replaceCells would replace", () => {
    const pattern = /b/i
    const s = sheet()

    const found = findCells(s, pattern)
    const replaced = replaceCells(s, pattern, "X")

    expect(found).toHaveLength(replaced)
  })

  it("does not let a /g pattern's lastIndex depend on call order", () => {
    const found = findCells(sheet(), /b/gi)

    expect(found.map((f) => f.value)).toEqual(["b1", "B2"])
  })

  it("still accepts a plain value and a predicate", () => {
    expect(findCells(sheet(), 7)).toEqual([{ row: 1, col: 1, value: 7 }])
    expect(findCells(sheet(), (v) => typeof v === "number")).toEqual([{ row: 1, col: 1, value: 7 }])
  })

  it("does not match a RegExp against a non-string cell", () => {
    expect(findCells(sheet(), /7/)).toEqual([])
  })
})

describe("a merge shrunk to one cell is no longer a merge", () => {
  it("drops it after deleteColumns", () => {
    const s: Sheet = {
      name: "S",
      rows: [["a", "b", "c"]],
      merges: [{ startRow: 0, startCol: 1, endRow: 0, endCol: 2 }],
    }

    deleteColumns(s, 1, 1)

    expect(s.merges).toEqual([])
  })

  it("drops it after deleteRows", () => {
    const s: Sheet = {
      name: "S",
      rows: [["a"], ["b"], ["c"]],
      merges: [{ startRow: 1, startCol: 0, endRow: 2, endCol: 0 }],
    }

    deleteRows(s, 1, 1)

    expect(s.merges).toEqual([])
  })

  it("keeps a merge that still spans more than one cell", () => {
    const s: Sheet = {
      name: "S",
      rows: [["a", "b", "c", "d"]],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
    }

    deleteColumns(s, 0, 1)

    expect(s.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }])
  })
})
