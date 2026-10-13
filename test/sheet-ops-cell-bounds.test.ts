import { expect, it } from "vitest"
import { createCellStore } from "../src/cell"
import { insertColumns, insertRows } from "../src/sheet-ops"
import { InvalidArgumentError } from "../src/errors"
import type { Sheet } from "../src/_types"

it.each([
  ["row", 1_048_575, 0, insertRows],
  ["column", 0, 16_383, insertColumns],
] as const)(
  "rejects moving %s metadata off the grid before changing the sheet",
  (_, row, col, insert) => {
    const sheet: Sheet = {
      name: "Bounds",
      rows: [[1]],
      cells: createCellStore([[row, col, { value: 1, type: "number" }]]),
      columns: [{ width: 12 }],
      rowDefs: new Map([[0, { height: 20 }]]),
      merges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
    }
    const before = structuredClone(sheet)
    expect(() => insert(sheet, 0, 1)).toThrow(InvalidArgumentError)
    expect(sheet).toEqual(before)
  },
)
