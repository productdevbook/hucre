import { describe, expect, it } from "vitest"
import { createCellStore, cellEntries } from "../src/cell"
import { insertRows, deleteRows, insertColumns, deleteColumns } from "../src/sheet-ops"
import { InvalidArgumentError } from "../src/errors"
import type { Sheet } from "../src/_types"

const operations = [
  ["insert rows", insertRows, "B2", "B3"],
  ["delete rows", deleteRows, "B3", "B2"],
  ["insert columns", insertColumns, "B2", "C2"],
  ["delete columns", deleteColumns, "C2", "B2"],
] as const

describe.each(operations)("%s", (_, edit, before, after) => {
  it("uses the writer's case-insensitive sheet-name identity", () => {
    const sheet: Sheet = {
      name: "Orders",
      rows: [[1]],
      cells: createCellStore([[0, 0, { value: 1, type: "formula", formula: `oRdErS!${before}` }]]),
    }
    edit(sheet, 1, 1)
    expect([...cellEntries(sheet.cells)][0][2].formula).toBe(`oRdErS!${after}`)
  })
  it.each(["Orders", "Bob's Data"])(
    "moves references qualified with its own sheet name (%s)",
    (name) => {
      const qualifier = name === "Orders" ? name : "'Bob''s Data'"
      const sheet: Sheet = {
        name,
        rows: [
          [1, 2, 3],
          [4, 5, 6],
          [7, 8, 9],
        ],
        cells: createCellStore([
          [2, 2, { value: 9, type: "formula", formula: `${qualifier}!${before}+Other!${before}` }],
        ]),
        dataValidations: [{ type: "custom", range: "A1:C3", formula1: `${qualifier}!${before}>0` }],
      }
      edit(sheet, 1, 1)
      expect([...cellEntries(sheet.cells)][0][2].formula).toBe(
        `${qualifier}!${after}+Other!${before}`,
      )
      expect(sheet.dataValidations?.[0].formula1).toBe(`${qualifier}!${after}>0`)
    },
  )

  it.each([
    [-1, 1],
    [0.5, 1],
    [0, 0.5],
    [NaN, 1],
    [0, NaN],
  ])("rejects invalid geometry (%s,%s) without changing the sheet", (index, count) => {
    const sheet: Sheet = {
      name: "Orders",
      rows: [
        [1, 2],
        [3, 4],
      ],
      merges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }],
      dataValidations: [{ type: "custom", range: "A1:B2", formula1: "A1>0" }],
    }
    const snapshot = structuredClone(sheet)
    expect(() => edit(sheet, index, count)).toThrow(InvalidArgumentError)
    expect(sheet).toEqual(snapshot)
  })
  it.each([
    [Infinity, 1],
    [0, Infinity],
    [1_048_576, 1],
    [0, 1_048_577],
  ])("rejects unbounded geometry (%s,%s)", (index, count) => {
    const sheet: Sheet = { name: "Bounds", rows: [[1]] }
    expect(() => edit(sheet, index, count)).toThrow(InvalidArgumentError)
    expect(sheet.rows).toEqual([[1]])
  })
})

it("inserts a valid large row batch without a function argument ceiling", () => {
  const sheet: Sheet = { name: "Bulk", rows: [[1], [2]] }
  insertRows(sheet, 1, 150_000)
  expect(sheet.rows.length).toBe(150_002)
  expect(sheet.rows[0]).toEqual([1])
  expect(sheet.rows[1]).toEqual([null])
  expect(sheet.rows[150_000]).toEqual([null])
  expect(sheet.rows[150_001]).toEqual([2])
  expect(sheet.rows[1]).not.toBe(sheet.rows[2])
})

it("keeps surviving areas when only part of a validation or rule is deleted", () => {
  const sheet: Sheet = {
    name: "Areas",
    rows: [
      [1, 2, 3],
      [4, 5, 6],
    ],
    dataValidations: [{ type: "list", range: "A1:A2 C1:C2", values: ["x"] }],
    conditionalRules: [{ type: "expression", range: "A1:A2 C1:C2", priority: 1, formula: "C1>0" }],
  }
  deleteColumns(sheet, 0, 1)
  expect(sheet.dataValidations?.[0].range).toBe("B1:B2")
  expect(sheet.conditionalRules?.[0].range).toBe("B1:B2")
  expect(sheet.conditionalRules?.[0].formula).toBe("B1>0")
})

it("does not mistake a sheet name containing #REF! for a deleted area", () => {
  const sheet: Sheet = {
    name: "Source #REF!",
    rows: [[1], [2]],
    dataValidations: [{ type: "custom", range: "'Source #REF!'!A1:A2", formula1: "A1>0" }],
  }
  insertRows(sheet, 0, 1)
  expect(sheet.dataValidations?.[0].range).toBe("'Source #REF!'!A2:A3")
})

it("preflights column capacity before replacing cell metadata", () => {
  const cells = createCellStore([[0, 0, { value: 1, type: "number" }]])
  const sheet: Sheet = { name: "Full", rows: [Array(16_384).fill(1)], cells }
  expect(() => insertColumns(sheet, 0, 1)).toThrow(InvalidArgumentError)
  expect(sheet.rows[0].length).toBe(16_384)
  expect(sheet.cells).toBe(cells)
})

it("preflights row capacity before replacing cell metadata", () => {
  const cells = createCellStore([[0, 0, { value: 1, type: "number" }]])
  const sheet: Sheet = { name: "Full", rows: Array(1_048_576).fill([]), cells }
  expect(() => insertRows(sheet, 0, 1)).toThrow(InvalidArgumentError)
  expect(sheet.rows.length).toBe(1_048_576)
  expect(sheet.cells).toBe(cells)
})

it("preflights row definitions outside the dense grid", () => {
  const sheet: Sheet = {
    name: "Styled",
    rows: [],
    rowDefs: new Map([[1_048_575, { height: 20 }]]),
  }
  const snapshot = structuredClone(sheet)
  expect(() => insertRows(sheet, 0, 1)).toThrow(InvalidArgumentError)
  expect(sheet).toEqual(snapshot)
})
