import { describe, expect, it } from "vitest"
import { shiftFormula } from "../src/_refs"
import { insertRows } from "../src/sheet-ops"
import type { Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// Inserting a row corrupted every formula that named a sheet in quotes:
//
//   'My Sheet'!A3   ->   '''My Sheet'''!A3
//
// `replaceA1Ranges` hands back the qualifier as it appeared in the
// formula — quotes included, since that is what makes `My Sheet` a legal
// qualifier — and `_refs.ts` then ran it through a `quoteSheet` that
// assumed a bare name and escaped the quotes it found. Excel rejects the
// result.
//
// The same mismatch caused a second failure, quieter than the first. The
// shift only applies to references on the sheet that moved, decided by
// `match.sheet1 !== shift.sheetName` — and `'My Sheet'` never equals
// `My Sheet`, so a formula pointing at its *own* sheet by a quoted name
// was left unshifted while `S!A3` on a sheet named `S` shifted correctly.
// The rows moved underneath it and the reference did not follow.
//
// The ODS formula writer reads the same qualifier and wants the quotes —
// `$'My Sheet'.A1` is correct ODF and round-trips today — so the fix is
// here rather than in the shared matcher.
// ═══════════════════════════════════════════════════════════════════════

const INSERT_AT_2 = { axis: "row", at: 2, delta: 1 } as const

describe("a quoted sheet name survives a shift", () => {
  it("unchanged when it names another sheet", () => {
    expect(shiftFormula("'My Sheet'!A3", { ...INSERT_AT_2, sheetName: "S" })).toBe("'My Sheet'!A3")
  })

  it("in a range, and in a function call", () => {
    expect(shiftFormula("SUM('My Sheet'!A3:A5)", { ...INSERT_AT_2, sheetName: "S" })).toBe(
      "SUM('My Sheet'!A3:A5)",
    )
  })

  it("with an escaped apostrophe in the name", () => {
    // Excel writes a literal `'` in a sheet name as `''`. Re-escaping it
    // turned `'Bob''s Data'` into `'''Bob''''s Data'''`.
    expect(shiftFormula("'Bob''s Data'!A3", { ...INSERT_AT_2, sheetName: "S" })).toBe(
      "'Bob''s Data'!A3",
    )
  })

  it("and an unquoted one is still left alone", () => {
    expect(shiftFormula("Sheet2!A3", { ...INSERT_AT_2, sheetName: "S" })).toBe("Sheet2!A3")
  })
})

describe("a reference to the moving sheet shifts, quoted or not", () => {
  it("unquoted, as it always did", () => {
    expect(shiftFormula("S!A3", { ...INSERT_AT_2, sheetName: "S" })).toBe("S!A4")
  })

  it("quoted, which it did not", () => {
    // The quiet half: `'My Sheet'` never equalled `My Sheet`, so this was
    // treated as a foreign sheet and left where it was.
    expect(shiftFormula("'My Sheet'!A3", { ...INSERT_AT_2, sheetName: "My Sheet" })).toBe(
      "'My Sheet'!A4",
    )
  })

  it("quoted, with an escaped apostrophe", () => {
    expect(shiftFormula("'Bob''s Data'!A3", { ...INSERT_AT_2, sheetName: "Bob's Data" })).toBe(
      "'Bob''s Data'!A4",
    )
  })

  it("quoted, in a range", () => {
    expect(shiftFormula("SUM('My Sheet'!A3:A5)", { ...INSERT_AT_2, sheetName: "My Sheet" })).toBe(
      "SUM('My Sheet'!A4:A6)",
    )
  })

  it("and a local reference is unaffected by any of this", () => {
    expect(shiftFormula("A3", { ...INSERT_AT_2, sheetName: "My Sheet" })).toBe("A4")
    expect(shiftFormula("SUM(A1:A5)", { ...INSERT_AT_2, sheetName: "My Sheet" })).toBe("SUM(A1:A6)")
  })
})

describe("through the operation a caller actually runs", () => {
  it("insertRows leaves a foreign quoted reference intact", () => {
    const sheet: Sheet = {
      name: "S",
      rows: [["a"], ["b"], ["c"]],
      cells: new Map([["2,0", { value: 1, type: "formula", formula: "'My Sheet'!A3" }]]),
    }

    insertRows(sheet, 0, 1)

    expect(sheet.cells!.get("3,0")?.formula).toBe("'My Sheet'!A3")
  })
})
