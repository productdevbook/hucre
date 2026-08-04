import { describe, expect, it } from "vitest"
import { parseCsv } from "../src/csv/reader"
import { streamCsvRows } from "../src/csv/stream"
import type { CellValue, CsvReadOptions } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

function stream(input: string, options?: CsvReadOptions): CellValue[][] {
  return Array.from(streamCsvRows(input, options))
}

const SIMPLE = "name,qty\r\nfoo,1\r\nbar,2\r\n"
const QUOTED = 'a,"b,c"\r\n"x""y",z\r\n'
// As writeCsv({ escapeFormulae: true }) emits it.
const ESCAPED = "'-5,'=SUM(A1)\r\n'@x,plain\r\n"
const MESSY = [
  "# leading comment",
  "name;qty;when",
  "foo;1;2024-01-15",
  "",
  '"semi;colon";02;true',
  "bar;3;x",
].join("\n")

// ═══════════════════════════════════════════════════════════════════════
// streamCsvRows must agree with parseCsv on the shared option set.
// Regression guard for the divergence documented in #353.
// ═══════════════════════════════════════════════════════════════════════

describe("streamCsvRows / parseCsv parity", () => {
  const inputs: Array<[string, string]> = [
    ["simple", SIMPLE],
    ["quoted", QUOTED],
    ["messy", MESSY],
    ["escaped", ESCAPED],
  ]

  const optionCases: Array<[string, CsvReadOptions]> = [
    ["defaults", {}],
    ["header", { header: true }],
    ["typeInference", { typeInference: true }],
    ["typeInference + header", { typeInference: true, header: true }],
    ["preserveLeadingZeros off", { typeInference: true, preserveLeadingZeros: false }],
    ["fastMode", { fastMode: true }],
    ["fastMode + typeInference", { fastMode: true, typeInference: true }],
    ["maxRows 1", { maxRows: 1 }],
    ["maxRows 0", { maxRows: 0 }],
    ["skipLines 1", { skipLines: 1 }],
    ["skipEmptyRows", { skipEmptyRows: true }],
    ["comment", { comment: "#" }],
    ["delimiter ;", { delimiter: ";" }],
    ["skipBom off", { skipBom: false }],
    // Both were on CsvReadOptions but implemented in one reader only —
    // skipHeaderRow in streamCsvRows, and neither honoured unescapeFormulae
    // before it existed (#408).
    ["header + skipHeaderRow", { header: true, skipHeaderRow: true }],
    ["header + skipHeaderRow + maxRows 1", { header: true, skipHeaderRow: true, maxRows: 1 }],
    ["unescapeFormulae", { unescapeFormulae: true }],
    ["unescapeFormulae + typeInference", { unescapeFormulae: true, typeInference: true }],
  ]

  for (const [inputLabel, input] of inputs) {
    for (const [optionLabel, options] of optionCases) {
      it(`${inputLabel} — ${optionLabel}`, () => {
        expect(stream(input, options)).toEqual(parseCsv(input, options))
      })
    }
  }
})

describe("streamCsvRows — fastMode", () => {
  it("splits without quote handling, like parseCsv", () => {
    expect(stream(QUOTED, { fastMode: true })).toEqual(parseCsv(QUOTED, { fastMode: true }))
    // And that really is different from the RFC 4180 parse.
    expect(stream(QUOTED, { fastMode: true })).not.toEqual(stream(QUOTED))
  })

  it("treats the quote char as ordinary content", () => {
    expect(stream('a,"b,c"\r\n', { fastMode: true })).toEqual([["a", '"b', 'c"']])
  })
})

describe("streamCsvRows — transformValue", () => {
  it("applies to every cell, like parseCsv", () => {
    const upper = (v: CellValue) => (typeof v === "string" ? v.toUpperCase() : v)
    expect(stream(SIMPLE, { transformValue: upper })).toEqual(
      parseCsv(SIMPLE, { transformValue: upper }),
    )
  })

  it("names columns from the header row when header is set", () => {
    const seen: string[] = []
    stream(SIMPLE, {
      header: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["name", "qty", "name", "qty", "name", "qty"])
  })

  it("falls back to the column index without a header", () => {
    const seen: string[] = []
    stream(SIMPLE, {
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(new Set(seen)).toEqual(new Set(["0", "1"]))
  })

  it("runs after type inference", () => {
    const types: string[] = []
    stream("1\r\n", {
      typeInference: true,
      transformValue: (v) => {
        types.push(typeof v)
        return v
      },
    })
    expect(types).toEqual(["number"])
  })

  it("reports row and column indices", () => {
    const coords: Array<[number, number]> = []
    stream(SIMPLE, {
      transformValue: (v, _h, row, col) => {
        coords.push([row, col])
        return v
      },
    })
    expect(coords).toEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [2, 0],
      [2, 1],
    ])
  })
})

describe("streamCsvRows — onRow", () => {
  it("fires once per yielded row, like parseCsv", () => {
    const streamed: Array<[number, CellValue[]]> = []
    stream(SIMPLE, { onRow: (row, i) => streamed.push([i, row]) })

    const parsed: Array<[number, CellValue[]]> = []
    parseCsv(SIMPLE, { onRow: (row, i) => parsed.push([i, row]) })

    expect(streamed).toEqual(parsed)
  })

  it("sees the transformed row", () => {
    const seen: CellValue[][] = []
    stream("a\r\n", {
      transformValue: () => "REPLACED",
      onRow: (row) => seen.push(row),
    })
    expect(seen).toEqual([["REPLACED"]])
  })

  it("does not fire for rows past maxRows", () => {
    let count = 0
    stream(SIMPLE, { maxRows: 1, onRow: () => count++ })
    expect(count).toBe(1)
  })
})

describe("streamCsvRows — skipHeaderRow", () => {
  it("drops the header row in parseCsv too", () => {
    // It was honoured here and ignored there — one option, two behaviours,
    // which is the divergence this whole file exists to prevent (#408).
    expect(parseCsv(SIMPLE, { header: true, skipHeaderRow: true })).toEqual(
      stream(SIMPLE, { header: true, skipHeaderRow: true }),
    )
  })

  it("drops the header row when asked", () => {
    expect(stream(SIMPLE, { header: true, skipHeaderRow: true })).toEqual([
      ["foo", "1"],
      ["bar", "2"],
    ])
  })

  it("keeps the header row by default, matching parseCsv", () => {
    expect(stream(SIMPLE, { header: true })).toEqual(parseCsv(SIMPLE, { header: true }))
    expect(stream(SIMPLE, { header: true })[0]).toEqual(["name", "qty"])
  })

  it("is inert without header: true", () => {
    expect(stream(SIMPLE, { skipHeaderRow: true })).toEqual(stream(SIMPLE))
  })

  it("still names transformValue columns from the consumed header", () => {
    const seen: string[] = []
    stream(SIMPLE, {
      header: true,
      skipHeaderRow: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["name", "qty", "name", "qty"])
  })

  it("counts maxRows against emitted rows only", () => {
    expect(stream(SIMPLE, { header: true, skipHeaderRow: true, maxRows: 1 })).toEqual([
      ["foo", "1"],
    ])
  })
})
