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
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
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

// ═══════════════════════════════════════════════════════════════════════
// The binary formats had no property test. Everything above is a parser
// with a *string* inverse; `writeXlsx`/`readXlsx` and `writeOds`/
// `readOds` are inverses too, and nobody had pointed random values at
// them.
//
// Doing so found two defects in the ODS writer on the first run — a
// carriage return written as Excel's `_x000D_`, and, chasing that, a
// sheet name escaped with the text escaper inside an attribute. Both are
// pinned properly in test/ods-escaping.test.ts; this is the thing that
// noticed, kept so it can notice the next one.
// ═══════════════════════════════════════════════════════════════════════

/** Values chosen to sit where a spreadsheet writer has to make a choice. */
const AWKWARD_CELLS: CellValue[] = [
  ...AWKWARD,
  "with\rbare cr",
  "trailing\r",
  "<xml>&amp;</xml>",
  "]]>",
  "a".repeat(300),
  0,
  -12.5,
  1e21,
  1e-7,
  Number.MAX_SAFE_INTEGER,
  0.1 + 0.2,
  true,
  false,
  null,
]

function randomBinaryCell(rnd: () => number): CellValue {
  if (rnd() < 0.12) {
    return new Date(
      Date.UTC(1900 + Math.floor(rnd() * 200), Math.floor(rnd() * 12), 1 + Math.floor(rnd() * 28)),
    )
  }
  return AWKWARD_CELLS[Math.floor(rnd() * AWKWARD_CELLS.length)]!
}

/**
 * Trailing empties are trimmed on read, so the comparison is on the
 * trimmed form — the readers do not promise to hand back a row's shape,
 * only its values. `docs/PARITY.md` says so.
 */
function trimTrailing(row: CellValue[]): unknown[] {
  const out = [...row]
  while (out.length > 0 && (out[out.length - 1] === null || out[out.length - 1] === "")) out.pop()
  return out.map((v) => (v instanceof Date ? `D:${v.toISOString()}` : v))
}

describe("write then read is the identity, for the binary formats", () => {
  const cases = [
    ["xlsx", writeXlsx, readXlsx],
    ["ods", writeOds, readOds],
  ] as const

  for (const [label, write, read] of cases) {
    it(`${label} carries every value it was given`, async () => {
      const rnd = seeded(SEED)

      for (let run = 0; run < 60; run++) {
        const rows: CellValue[][] = Array.from({ length: 1 + Math.floor(rnd() * 5) }, () =>
          Array.from({ length: 1 + Math.floor(rnd() * 5) }, () => randomBinaryCell(rnd)),
        )

        const bytes = await (write as typeof writeXlsx)({ sheets: [{ name: "S", rows }] })
        const back = (await (read as typeof readXlsx)(bytes)).sheets[0]!.rows

        for (let i = 0; i < rows.length; i++) {
          expect(trimTrailing(back[i] ?? []), `${label} run ${run} row ${i}`).toEqual(
            trimTrailing(rows[i]!),
          )
        }
      }
    })
  }
})

describe("the one value the generator above deliberately leaves out", () => {
  // `docs/PARITY.md` records that a cell whose text is literally
  // `_x0041_` reads back as `A`: OOXML uses `_xHHHH_` for characters XML
  // cannot hold, hucre decodes it on read, and it does not re-escape a
  // leading underscore on write because that would mangle the far more
  // common ordinary text containing one. The ambiguity is accepted.
  //
  // It is accepted **for XLSX**. ODS never had the convention, and since
  // the CR fix does not write it either, the same string survives there.
  // Feeding it to the property test above would have looked like a bug
  // in one format and a pass in the other, so it is stated here instead.
  const LITERAL = "_x000D_ already"

  it("xlsx decodes it, as documented", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[LITERAL]] }] })

    expect((await readXlsx(bytes)).sheets[0]!.rows[0]![0]).toBe("\r already")
  })

  it("ods carries it through unchanged", async () => {
    const bytes = await writeOds({ sheets: [{ name: "S", rows: [[LITERAL]] }] })

    expect((await readOds(bytes)).sheets[0]!.rows[0]![0]).toBe(LITERAL)
  })
})
