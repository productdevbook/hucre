import { describe, expect, it } from "vitest"
import { deleteColumns, deleteRows, insertColumns, insertRows } from "../src/sheet-ops"
import { shiftFormula } from "../src/_refs"
import type { Cell, Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §D — adjusting references is the *defining* behaviour of "insert a
// row", and sheet-ops did none of it. The cells moved and everything that
// named a position stayed:
//
//   insertRows(sheet, 0, 2)
//   cell key:   "4,0" -> "6,0"          correct
//   formula:    "SUM(A1:A4)" unchanged  WRONG — should be SUM(A3:A6)
//   rowBreaks:  [3] unchanged           WRONG — should be [5]
//   sparkline:  "A1:A4" unchanged       WRONG
//
// The rewriter was already in cell-utils and wired to nothing but the ODS
// writer.
// ═══════════════════════════════════════════════════════════════════════

describe("shiftFormula", () => {
  it("moves references at or after the insertion point", () => {
    expect(shiftFormula("SUM(A1:A4)", { axis: "row", at: 0, delta: 2 })).toBe("SUM(A3:A6)")
  })

  it("leaves references before it alone", () => {
    expect(shiftFormula("SUM(A5:A9)", { axis: "row", at: 10, delta: 2 })).toBe("SUM(A5:A9)")
  })

  it("moves an absolute reference too", () => {
    // `$` governs what happens when a formula is *copied*, not when the
    // grid moves underneath it. Excel shifts $A$5 on an insert as well.
    expect(shiftFormula("$A$5+B5", { axis: "row", at: 0, delta: 2 })).toBe("$A$7+B7")
  })

  it("turns a reference into a deleted row into #REF!", () => {
    expect(shiftFormula("A3", { axis: "row", at: 2, delta: -2 })).toBe("#REF!")
  })

  it("collapses a range the deletion swallowed whole", () => {
    // Once, not as `#REF!:#REF!`.
    expect(shiftFormula("SUM(A3:A4)", { axis: "row", at: 2, delta: -2 })).toBe("SUM(#REF!)")
  })

  it("shrinks a range the deletion clips", () => {
    expect(shiftFormula("SUM(A1:A10)", { axis: "row", at: 2, delta: -2 })).toBe("SUM(A1:A8)")
  })

  it("moves columns on the column axis", () => {
    expect(shiftFormula("SUM(B1:D1)", { axis: "col", at: 0, delta: 1 })).toBe("SUM(C1:E1)")
  })

  it("leaves another sheet's references where they are", () => {
    const shift = { axis: "row", at: 0, delta: 2, sheetName: "Sheet1" } as const

    expect(shiftFormula("Sheet2!A5", shift)).toBe("Sheet2!A5")
    expect(shiftFormula("Sheet1!A5", shift)).toBe("Sheet1!A7")
  })

  it("does not touch string literals or function names", () => {
    const shift = { axis: "row", at: 0, delta: 2 } as const

    expect(shiftFormula('IF(A5="A5","A5",A5)', shift)).toBe('IF(A7="A5","A5",A7)')
    expect(shiftFormula("LOG10(A5)", shift)).toBe("LOG10(A7)")
  })
})

function sheet(): Sheet {
  const cells = new Map<string, Cell>([
    ["4,0", { value: 10, type: "number", formula: "SUM(A1:A4)" }],
    ["4,1", { value: 1, type: "number", formula: "Other!B1+A5" }],
  ])
  return {
    name: "S",
    rows: [[1], [2], [3], [4], [10]],
    cells,
    rowBreaks: [3],
    colBreaks: [1],
    sparklines: [{ dataRange: "A1:A4", location: "B5", type: "line" }],
    dataValidations: [{ type: "list", range: "C1:C9", formula1: "$A$1:$A$4" }],
    conditionalRules: [
      { type: "cellIs", priority: 1, range: "A1:A9", operator: "greaterThan", formula: ["$A$5"] },
    ],
    textBoxes: [{ text: "note", anchor: { from: { row: 4, col: 0 } } }],
  }
}

describe("insertRows moves what names a position", () => {
  it("rewrites the formula whose arguments moved", () => {
    const s = sheet()

    insertRows(s, 0, 2)

    expect(s.cells!.get("6,0")!.formula).toBe("SUM(A3:A6)")
  })

  it("leaves another sheet's part of a formula alone", () => {
    const s = sheet()

    insertRows(s, 0, 2)

    expect(s.cells!.get("6,1")!.formula).toBe("Other!B1+A7")
  })

  it("moves the page breaks", () => {
    const s = sheet()

    insertRows(s, 0, 2)

    expect(s.rowBreaks).toEqual([5])
    expect(s.colBreaks).toEqual([1])
  })

  it("moves the sparkline's source and location", () => {
    const s = sheet()

    insertRows(s, 0, 2)

    expect(s.sparklines![0]).toMatchObject({ dataRange: "A3:A6", location: "B7" })
  })

  it("moves the formulas inside validations and conditional rules", () => {
    const s = sheet()

    insertRows(s, 0, 2)

    expect(s.dataValidations![0]!.formula1).toBe("$A$3:$A$6")
    expect(s.conditionalRules![0]!.formula).toEqual(["$A$7"])
  })

  it("moves a text box anchored below the insertion", () => {
    const s = sheet()

    insertRows(s, 0, 2)

    expect(s.textBoxes![0]!.anchor.from.row).toBe(6)
  })
})

describe("deleteRows leaves #REF! where a reference pointed", () => {
  it("marks a formula that pointed into the deleted rows", () => {
    const s: Sheet = {
      name: "S",
      rows: [[1], [2], [3]],
      cells: new Map<string, Cell>([["2,0", { value: 1, type: "number", formula: "A2" }]]),
    }

    deleteRows(s, 1, 1)

    expect(s.cells!.get("1,0")!.formula).toBe("#REF!")
  })

  it("shrinks a range the deletion clipped", () => {
    const s = sheet()

    deleteRows(s, 1, 2)

    expect(s.cells!.get("2,0")!.formula).toBe("SUM(A1:A2)")
  })

  it("moves the page breaks up", () => {
    const s = sheet()

    deleteRows(s, 0, 2)

    expect(s.rowBreaks).toEqual([1])
  })

  it("drops a sparkline whose whole source was deleted", () => {
    const s = sheet()

    deleteRows(s, 0, 4)

    expect(s.sparklines).toEqual([])
  })
})

describe("the column operations do the same on the other axis", () => {
  const withFormula = (): Sheet => ({
    name: "S",
    rows: [[1, 2, 3]],
    cells: new Map<string, Cell>([["0,2", { value: 3, type: "number", formula: "SUM(A1:B1)" }]]),
    colBreaks: [2],
  })

  it("moves references right on insert", () => {
    const s = withFormula()

    insertColumns(s, 0, 1)

    expect(s.cells!.get("0,3")!.formula).toBe("SUM(B1:C1)")
    expect(s.colBreaks).toEqual([3])
  })

  it("moves references left on delete, shrinking a clipped range", () => {
    const s = withFormula()

    deleteColumns(s, 0, 1)

    // A is gone and B slides into its place, so the range shrinks to the
    // one column that survived rather than going #REF!. Only a range the
    // deletion swallows *whole* has nothing left to point at.
    expect(s.cells!.get("0,1")!.formula).toBe("SUM(A1:A1)")
    expect(s.colBreaks).toEqual([1])
  })
})
