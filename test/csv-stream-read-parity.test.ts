import { describe, expect, it } from "vitest"
import { parseCsv } from "../src/csv/reader"
import { streamCsvRows } from "../src/csv/stream"
import type { CellValue, CsvReadOptions } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

async function stream(input: string, options?: CsvReadOptions): Promise<CellValue[][]> {
  const rows: CellValue[][] = []
  for await (const row of streamCsvRows(input, options)) rows.push(row.values)
  return rows
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
    ["header", { hasHeaderRow: true }],
    ["typeInference", { typeInference: true }],
    ["typeInference + header", { typeInference: true, hasHeaderRow: true }],
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
    ["header + skipHeaderRow", { hasHeaderRow: true, skipHeaderRow: true }],
    ["header + skipHeaderRow + maxRows 1", { hasHeaderRow: true, skipHeaderRow: true, maxRows: 1 }],
    ["unescapeFormulae", { unescapeFormulae: true }],
    ["unescapeFormulae + typeInference", { unescapeFormulae: true, typeInference: true }],
  ]

  for (const [inputLabel, input] of inputs) {
    for (const [optionLabel, options] of optionCases) {
      it(`${inputLabel} — ${optionLabel}`, async () => {
        expect(await stream(input, options)).toEqual(parseCsv(input, options))
      })
    }
  }
})

describe("streamCsvRows — fastMode", () => {
  it("splits without quote handling, like parseCsv", async () => {
    expect(await stream(QUOTED, { fastMode: true })).toEqual(parseCsv(QUOTED, { fastMode: true }))
    // And that really is different from the RFC 4180 parse.
    expect(await stream(QUOTED, { fastMode: true })).not.toEqual(await stream(QUOTED))
  })

  it("treats the quote char as ordinary content", async () => {
    expect(await stream('a,"b,c"\r\n', { fastMode: true })).toEqual([["a", '"b', 'c"']])
  })
})

describe("streamCsvRows — transformValue", () => {
  it("applies to every cell, like parseCsv", async () => {
    const upper = (v: CellValue) => (typeof v === "string" ? v.toUpperCase() : v)
    expect(await stream(SIMPLE, { transformValue: upper })).toEqual(
      parseCsv(SIMPLE, { transformValue: upper }),
    )
  })

  it("names columns from the header row when header is set", async () => {
    const seen: string[] = []
    await stream(SIMPLE, {
      hasHeaderRow: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["name", "qty", "name", "qty", "name", "qty"])
  })

  it("falls back to the column index without a header", async () => {
    const seen: string[] = []
    await stream(SIMPLE, {
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(new Set(seen)).toEqual(new Set(["0", "1"]))
  })

  it("runs after type inference", async () => {
    const types: string[] = []
    await stream("1\r\n", {
      typeInference: true,
      transformValue: (v) => {
        types.push(typeof v)
        return v
      },
    })
    expect(types).toEqual(["number"])
  })

  it("reports row and column indices", async () => {
    const coords: Array<[number, number]> = []
    await stream(SIMPLE, {
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
  it("fires once per yielded row, like parseCsv", async () => {
    const streamed: Array<[number, CellValue[]]> = []
    await stream(SIMPLE, { onRow: (row, i) => streamed.push([i, row]) })

    const parsed: Array<[number, CellValue[]]> = []
    parseCsv(SIMPLE, { onRow: (row, i) => parsed.push([i, row]) })

    expect(streamed).toEqual(parsed)
  })

  it("sees the transformed row", async () => {
    const seen: CellValue[][] = []
    await stream("a\r\n", {
      transformValue: () => "REPLACED",
      onRow: (row) => seen.push(row),
    })
    expect(seen).toEqual([["REPLACED"]])
  })

  it("does not fire for rows past maxRows", async () => {
    let count = 0
    await stream(SIMPLE, { maxRows: 1, onRow: () => count++ })
    expect(count).toBe(1)
  })
})

describe("streamCsvRows — skipHeaderRow", () => {
  it("drops the header row in parseCsv too", async () => {
    // It was honoured here and ignored there — one option, two behaviours,
    // which is the divergence this whole file exists to prevent (#408).
    expect(parseCsv(SIMPLE, { hasHeaderRow: true, skipHeaderRow: true })).toEqual(
      await stream(SIMPLE, { hasHeaderRow: true, skipHeaderRow: true }),
    )
  })

  it("drops the header row when asked", async () => {
    expect(await stream(SIMPLE, { hasHeaderRow: true, skipHeaderRow: true })).toEqual([
      ["foo", "1"],
      ["bar", "2"],
    ])
  })

  it("keeps the header row by default, matching parseCsv", async () => {
    expect(await stream(SIMPLE, { hasHeaderRow: true })).toEqual(
      parseCsv(SIMPLE, { hasHeaderRow: true }),
    )
    expect((await stream(SIMPLE, { hasHeaderRow: true }))[0]).toEqual(["name", "qty"])
  })

  it("is inert without hasHeaderRow: true", async () => {
    expect(await stream(SIMPLE, { skipHeaderRow: true })).toEqual(await stream(SIMPLE))
  })

  it("still names transformValue columns from the consumed header", async () => {
    const seen: string[] = []
    await stream(SIMPLE, {
      hasHeaderRow: true,
      skipHeaderRow: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["name", "qty", "name", "qty"])
  })

  it("counts maxRows against emitted rows only", async () => {
    expect(await stream(SIMPLE, { hasHeaderRow: true, skipHeaderRow: true, maxRows: 1 })).toEqual([
      ["foo", "1"],
    ])
  })
})
