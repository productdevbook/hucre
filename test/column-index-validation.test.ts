import { describe, expect, it } from "vitest"
import { cellRef, colToLetter, rangeRef } from "../src/xlsx/worksheet-writer"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { InvalidArgumentError } from "../src/errors"
import { MAX_CELL_TEXT_LENGTH } from "../src/_validate"
import { MAX_COL_INDEX } from "../src/limits"

// ═══════════════════════════════════════════════════════════════════════
// #364 — colToLetter is pure arithmetic on an unchecked number, so an
// out-of-range or non-integer column produced a cell reference no reader
// can parse, inside a file that otherwise looked fine.
// ═══════════════════════════════════════════════════════════════════════

describe("colToLetter bounds", () => {
  it("still converts every legal column", () => {
    expect(colToLetter(0)).toBe("A")
    expect(colToLetter(25)).toBe("Z")
    expect(colToLetter(26)).toBe("AA")
    expect(colToLetter(MAX_COL_INDEX)).toBe("XFD")
  })

  const bad: Array<[string, number, string]> = [
    ["negative", -1, 'used to give "@"'],
    ["NaN", Number.NaN, "used to give a NUL character, illegal in XML"],
    ["fractional", 1.5, 'used to truncate silently to "B"'],
    ["past the last column", MAX_COL_INDEX + 1, 'used to give "XFE"'],
    ["absurdly large", 1_000_000, 'used to give "BDWGO"'],
  ]

  for (const [label, col, why] of bad) {
    it(`rejects a ${label} column — ${why}`, () => {
      expect(() => colToLetter(col)).toThrow(InvalidArgumentError)
    })
  }

  it("names the offending value", () => {
    expect(() => colToLetter(-1)).toThrow(/Column index -1/)
  })

  it("guards the helpers built on it", () => {
    expect(() => cellRef(0, -1)).toThrow(InvalidArgumentError)
    expect(() => rangeRef(0, 0, 0, MAX_COL_INDEX + 1)).toThrow(InvalidArgumentError)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// The length limits are deliberately NOT enforced. Pinned so the choice
// stays a decision rather than an omission someone "fixes" later.
// ═══════════════════════════════════════════════════════════════════════

describe("cell text beyond Excel's limit", () => {
  it("is written rather than rejected", async () => {
    // Excel's 32,767-character cap is an *application* limit, not a
    // format one: ECMA-376 imposes no such cap, the file stays valid
    // OOXML, and LibreOffice, pandas and hucre's own reader handle it.
    // Excel truncates the display rather than refusing the file, so
    // throwing here would make hucre stricter than the format it writes.
    const big = "A".repeat(MAX_CELL_TEXT_LENGTH + 1000)
    const workbook = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [[big]] }] }))
    expect(workbook.sheets[0].rows[0][0]).toBe(big)
  })

  it("exports the limit so a caller targeting Excel can check", () => {
    expect(MAX_CELL_TEXT_LENGTH).toBe(32_767)
  })
})
