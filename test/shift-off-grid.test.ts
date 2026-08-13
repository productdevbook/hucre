import { describe, expect, it } from "vitest"
import { shiftFormula } from "../src/_refs"
import { insertColumns, insertRows } from "../src/sheet-ops"
import { MAX_COL_INDEX, MAX_ROW_INDEX } from "../src/limits"
import type { Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// An insert that pushes a reference past the last row or column had two
// different wrong answers, one per axis:
//
//   A1048576  + insert row   ->  A1048577   a row that cannot exist
//   XFD1      + insert col   ->  throws InvalidArgumentError
//
// The throw is the worse of the two. It comes out of `colToLetter`, which
// is right to refuse column 16384, but it comes out through
// `insertColumns` — so inserting a column anywhere in a sheet that has
// one formula mentioning column XFD failed the whole operation with a
// message about column indexes.
//
// Excel's answer for both is `#REF!`: the reference had somewhere to
// point and no longer does, which is the same thing a deletion does to
// it, and `shiftCoordinate` already returned `null` for that.
// ═══════════════════════════════════════════════════════════════════════

const LAST_ROW = MAX_ROW_INDEX + 1 // 1-based, as a formula writes it
const LAST_COL = "XFD"

const ROW_INSERT = { axis: "row", at: 0, delta: 1 } as const
const COL_INSERT = { axis: "col", at: 0, delta: 1 } as const

describe("a reference pushed off the grid becomes #REF!", () => {
  it("on the row axis", () => {
    expect(shiftFormula(`A${LAST_ROW}`, ROW_INSERT)).toBe("#REF!")
  })

  it("on the column axis, where it used to throw", () => {
    expect(shiftFormula(`${LAST_COL}1`, COL_INSERT)).toBe("#REF!")
  })

  it("and the row below the last one still moves normally", () => {
    expect(shiftFormula(`A${LAST_ROW - 1}`, ROW_INSERT)).toBe(`A${LAST_ROW}`)
  })

  it("as does the column before the last", () => {
    expect(shiftFormula("XFC1", COL_INSERT)).toBe(`${LAST_COL}1`)
  })

  it("by more than one, too", () => {
    expect(shiftFormula(`A${LAST_ROW - 1}`, { axis: "row", at: 0, delta: 5 })).toBe("#REF!")
  })
})

describe("a range clipped by the edge keeps what fits", () => {
  it("when only its end falls off", () => {
    // The range still has somewhere to start, so it stops at the edge
    // rather than vanishing.
    expect(shiftFormula(`SUM(A1:A${LAST_ROW})`, ROW_INSERT)).toBe(`SUM(A2:A${LAST_ROW})`)
  })

  it("and becomes #REF! when its start does too", () => {
    // The range token is what becomes `#REF!`; the call around it stays,
    // which is both what Excel shows and what a deletion already did.
    expect(shiftFormula(`SUM(A${LAST_ROW}:A${LAST_ROW})`, ROW_INSERT)).toBe("SUM(#REF!)")
  })

  it("the same on columns", () => {
    expect(shiftFormula(`SUM(A1:${LAST_COL}1)`, COL_INSERT)).toBe(`SUM(B1:${LAST_COL}1)`)
    expect(shiftFormula(`SUM(${LAST_COL}1:${LAST_COL}1)`, COL_INSERT)).toBe("SUM(#REF!)")
  })

  it("which is exactly what a deletion swallowing a range gives", () => {
    // The pair, so the two paths cannot drift.
    expect(shiftFormula("SUM(A3:A3)", { axis: "row", at: 2, delta: -1 })).toBe("SUM(#REF!)")
  })
})

describe("through the operations a caller runs", () => {
  it("insertColumns no longer fails the whole sheet", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [["a"]],
      cells: new Map([["0,0", { value: 1, type: "formula", formula: `${LAST_COL}1` }]]),
    }

    expect(() => insertColumns(sheet, 0, 1)).not.toThrow()
    expect(sheet.cells!.get("0,1")?.formula).toBe("#REF!")
  })

  it("insertRows leaves a valid reference behind, not an impossible one", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [["a"]],
      cells: new Map([["0,0", { value: 1, type: "formula", formula: `A${LAST_ROW}` }]]),
    }

    insertRows(sheet, 0, 1)

    expect(sheet.cells!.get("1,0")?.formula).toBe("#REF!")
  })
})

describe("deletion is unchanged", () => {
  it("still gives #REF! for what it removed", () => {
    expect(shiftFormula("A3", { axis: "row", at: 2, delta: -1 })).toBe("#REF!")
  })

  it("and still clips a range rather than dropping it", () => {
    expect(shiftFormula("SUM(A1:A5)", { axis: "row", at: 2, delta: -1 })).toBe("SUM(A1:A4)")
  })

  it("the bounds themselves are what they always were", () => {
    expect(MAX_ROW_INDEX).toBe(1_048_575)
    expect(MAX_COL_INDEX).toBe(16_383)
  })
})
