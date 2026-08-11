import { describe, expect, it } from "vitest"
import { writeCsv } from "../src/csv/writer"
import { parseCsv } from "../src/csv/reader"
import { writeNdjson } from "../src/json/writer"
import { parseNdjson } from "../src/json/reader"
import {
  cellRef,
  parseCellRef,
  colToLetter,
  letterToCol,
  rangeRef,
  parseRange,
} from "../src/cell-utils"
import { MAX_COL_INDEX, MAX_ROW_INDEX } from "../src/limits"
import { seeded } from "./_fuzz"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #473 part 3 — property tests for the parsers that have an inverse.
//
// The issue named `fast-check` as the usual choice and noted it would be
// the repo's first test dependency beyond vitest. It is not added here:
// these properties are about the *inverse*, not about shrinking a
// counterexample out of a rich generator space, and a seeded loop over
// hand-built generators covers them without spending the zero-dependency
// claim. `fast-check` remains the right answer if the properties ever get
// harder than this — the door is not closed, it just is not open yet.
//
// Seeded for the same reason as the fuzzer: a failing property that
// cannot be reproduced has told you nothing.
// ═══════════════════════════════════════════════════════════════════════

const SEED = 0x5eed
const RUNS = 200

/** Strings chosen to sit on the boundaries a CSV writer has to handle. */
const AWKWARD = [
  "",
  " ",
  "  leading and trailing  ",
  "with,comma",
  'with"quote',
  '"fully quoted"',
  "with\nnewline",
  "with\r\ncrlf",
  "with\ttab",
  "şehir",
  "😀 astral",
  "-",
  "=SUM(A1)",
  "0007",
  "null",
  "undefined",
  "NaN",
  "true",
]

function randomCell(rnd: () => number): CellValue {
  const kind = Math.floor(rnd() * 10)
  if (kind < 4) return AWKWARD[Math.floor(rnd() * AWKWARD.length)]!
  if (kind < 6) return Math.floor(rnd() * 2_000_000) - 1_000_000
  if (kind < 7) return (rnd() - 0.5) * 1e6
  if (kind < 8) return rnd() < 0.5
  if (kind < 9) return null
  return `text ${Math.floor(rnd() * 1000)}`
}

function randomGrid(rnd: () => number): CellValue[][] {
  const cols = 1 + Math.floor(rnd() * 6)
  const rows = 1 + Math.floor(rnd() * 8)
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => randomCell(rnd)))
}

/** What CSV can carry: every value comes back as a string. */
function asText(value: CellValue): string {
  if (value === null || value === undefined) return ""
  if (value instanceof Date) return value.toISOString()
  return String(value)
}

/**
 * The one row CSV cannot round-trip, dropped from the expectation.
 *
 * A final row holding a single empty cell renders as nothing after the
 * preceding line's terminator, and a file that ends in a terminator is
 * universally read as having no record after it. RFC 4180 leaves the
 * trailing CRLF optional and says nothing about this, so there is no
 * spelling that distinguishes the two — every mainstream parser reads it
 * the way hucre does.
 *
 * A trailing row of *two* empty cells survives, because it renders as a
 * bare delimiter. So does an empty row anywhere but the end. The property
 * found this; it is a fact about the format, not a defect, and it is
 * narrowed here rather than papered over.
 */
function withoutUnrepresentableTail(grid: string[][]): string[][] {
  const last = grid[grid.length - 1]
  return last?.length === 1 && last[0] === "" ? grid.slice(0, -1) : grid
}

describe("parseCsv(writeCsv(rows)) preserves the text of every cell", () => {
  // The delimiter is named on both sides. A reader that has to *guess* it
  // is answering a different question, and gets its own test below —
  // where the property found a case worth knowing about.
  for (const delimiter of [",", "\t", ";", "|"]) {
    it(`over ${RUNS} generated grids, delimiter ${JSON.stringify(delimiter)}`, () => {
      const rnd = seeded(SEED)

      for (let run = 0; run < RUNS; run++) {
        const grid = randomGrid(rnd)
        const back = parseCsv(writeCsv(grid, { delimiter }), { delimiter })

        // CSV is a text format with no types: the property is that the
        // *text* survives, not that the values do. Anything less would be
        // asserting an accident of type inference.
        const want = withoutUnrepresentableTail(grid.map((row) => row.map(asText)))
        expect(back, `run ${run} seed ${SEED} delimiter ${JSON.stringify(delimiter)}`).toEqual(want)
      }
    })
  }
})

describe("auto-detection is a guess, and the property says where it misses", () => {
  it("a file whose only unquoted separator is a tab reads as tab-separated", () => {
    // Found by the property, at seed 0x5eed run 15. Written with the
    // default comma, this grid quotes its one comma — so the only
    // unquoted separator character left in the file is a tab, and the
    // sniffer is right to call it tab-separated. There is no reading of
    // those bytes that recovers the intent.
    const grid = [["with,comma"], ["with\ttab"]]
    const csv = writeCsv(grid)

    expect(csv).toContain('"with,comma"')
    expect(parseCsv(csv)).toEqual([["with,comma"], ["with", "tab"]])

    // Naming the delimiter is the answer, and it is exact.
    expect(parseCsv(csv, { delimiter: "," })).toEqual(grid)
  })

  it("detection is right whenever the file has unquoted delimiters to count", () => {
    const grid = [
      ["a", "b"],
      ["c\td", "e"],
    ]

    expect(parseCsv(writeCsv(grid))).toEqual(grid)
  })
})

describe("the one row CSV cannot carry", () => {
  it("a final row of one empty cell is indistinguishable from the terminator", () => {
    expect(writeCsv([["a"], [""]])).toBe("a\r\n")
    expect(parseCsv("a\r\n")).toEqual([["a"]])
  })

  it("but two empty cells survive, because they render as a delimiter", () => {
    expect(parseCsv(writeCsv([["a"], ["", ""]]))).toEqual([["a"], ["", ""]])
  })

  it("and an empty row anywhere but the end survives", () => {
    expect(parseCsv(writeCsv([["a"], [""], ["b"]]))).toEqual([["a"], [""], ["b"]])
  })
})

describe("parseNdjson(writeNdjson(records)) preserves the values", () => {
  it("over 200 generated record sets", () => {
    // JSON does carry types, so this is the stronger property: the values
    // themselves come back, not just their text.
    const rnd = seeded(SEED)

    for (let run = 0; run < RUNS; run++) {
      const keys = ["a", "b", "c"].slice(0, 1 + Math.floor(rnd() * 3))
      const records = Array.from({ length: 1 + Math.floor(rnd() * 6) }, () => {
        const record: Record<string, CellValue> = {}
        for (const key of keys) {
          const value = randomCell(rnd)
          // A Date serialises to a string and comes back as one; that is
          // JSON, not a defect, and is out of this property's scope.
          record[key] = value instanceof Date ? value.toISOString() : value
        }
        return record
      })

      expect(parseNdjson(writeNdjson(records)).data, `run ${run} seed ${SEED}`).toEqual(records)
    }
  })
})

describe("the A1 reference helpers are inverses", () => {
  it("parseCellRef(cellRef(row, col)) === { row, col }", () => {
    const rnd = seeded(SEED)

    for (let run = 0; run < RUNS * 5; run++) {
      const row = Math.floor(rnd() * (MAX_ROW_INDEX + 1))
      const col = Math.floor(rnd() * (MAX_COL_INDEX + 1))

      expect(parseCellRef(cellRef(row, col)), `row ${row} col ${col}`).toEqual({ row, col })
    }
  })

  it("letterToCol(colToLetter(col)) === col, across the whole grid", () => {
    // Every column, not a sample: 16,384 is small enough to be exhaustive,
    // and exhaustive beats random when it is affordable.
    for (let col = 0; col <= MAX_COL_INDEX; col++) {
      expect(letterToCol(colToLetter(col)), `col ${col}`).toBe(col)
    }
  })

  it("the boundaries specifically", () => {
    expect(cellRef(0, 0)).toBe("A1")
    expect(cellRef(MAX_ROW_INDEX, MAX_COL_INDEX)).toBe("XFD1048576")
    expect(parseCellRef("XFD1048576")).toEqual({ row: MAX_ROW_INDEX, col: MAX_COL_INDEX })
  })

  it("parseRange(rangeRef(...)) is the identity", () => {
    const rnd = seeded(SEED)

    for (let run = 0; run < RUNS; run++) {
      const startRow = Math.floor(rnd() * 10_000)
      const startCol = Math.floor(rnd() * 1000)
      const endRow = startRow + Math.floor(rnd() * 100)
      const endCol = startCol + Math.floor(rnd() * 50)

      expect(parseRange(rangeRef(startRow, startCol, endRow, endCol))).toEqual({
        startRow,
        startCol,
        endRow,
        endCol,
      })
    }
  })
})
