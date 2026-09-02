import { valuesOf } from "./_stream"
import { describe, expect, it } from "vitest"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { parseCsv, parseCsvObjects } from "../src/csv/reader"
import { streamCsvRows } from "../src/csv/stream"

// ═══════════════════════════════════════════════════════════════════════
// #439 §U and §V — options a shared type promises and the code ignored.
//
// The project already decided this pattern is unacceptable, twice:
// `CsvReadOptions.schema` was removed before v1 because no CSV reader
// honoured it, and `WriteSheet.threadedComments` was removed in #404
// because "a typed field that is silently discarded is worse than no
// field at all". The rule was right; it just was not applied here.
// ═══════════════════════════════════════════════════════════════════════

const GRID = [
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9],
]

describe("readOds honours maxRows and range", () => {
  it("bounds the rows it returns", async () => {
    const bytes = await writeOds({ sheets: [{ name: "S", rows: GRID }] })

    const wb = await readOds(bytes, { maxRows: 2 })

    expect(wb.sheets[0]!.rows).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ])
  })

  it("masks cells outside the range, keeping column indexes stable", async () => {
    const bytes = await writeOds({ sheets: [{ name: "S", rows: GRID }] })

    const wb = await readOds(bytes, { range: "B2:C3" })

    // The same shape readXlsx returns: a row outside the span is present
    // and empty, a column outside it is null.
    expect(wb.sheets[0]!.rows).toEqual([
      [null, null, null],
      [null, 5, 6],
      [null, 8, 9],
    ])
  })

  it("returns the same shape readXlsx does for the same options", async () => {
    const ods = await readOds(await writeOds({ sheets: [{ name: "S", rows: GRID }] }), {
      range: "B2:C3",
    })
    const xlsx = await readXlsx(await writeXlsx({ sheets: [{ name: "S", rows: GRID }] }), {
      range: "B2:C3",
    })

    expect(ods.sheets[0]!.rows).toEqual(xlsx.sheets[0]!.rows)
  })

  it("drops the cell overrides it masked", async () => {
    const bytes = await writeOds({
      sheets: [{ name: "S", rows: GRID, cells: new Map([["0,0", { value: 1, formula: "1+0" }]]) }],
    })

    const wb = await readOds(bytes, { range: "B2:C3" })

    expect(wb.sheets[0]!.cells?.get("0,0")).toBeUndefined()
  })

  it("leaves everything alone when neither option is set", async () => {
    const bytes = await writeOds({ sheets: [{ name: "S", rows: GRID }] })

    expect((await readOds(bytes)).sheets[0]!.rows).toEqual(GRID)
  })
})

describe("transformHeader means the same thing in all three CSV readers", () => {
  const CSV = "name,age\nAda,36"
  const upper = (h: string) => h.toUpperCase()

  it("rewrites the header row in parseCsv", () => {
    expect(parseCsv(CSV, { header: true, transformHeader: upper })).toEqual([
      ["NAME", "AGE"],
      ["Ada", "36"],
    ])
  })

  it("rewrites the header row in streamCsvRows", async () => {
    expect(await valuesOf(streamCsvRows(CSV, { header: true, transformHeader: upper }))).toEqual([
      ["NAME", "AGE"],
      ["Ada", "36"],
    ])
  })

  it("still renames the object keys in parseCsvObjects", () => {
    const { data, headers } = parseCsvObjects(CSV, { header: true, transformHeader: upper })

    expect(headers).toEqual(["NAME", "AGE"])
    expect(data).toEqual([{ NAME: "Ada", AGE: "36" }])
  })

  it("names transformValue's columns by the transformed header", () => {
    const seen: string[] = []
    parseCsv(CSV, {
      header: true,
      transformHeader: upper,
      transformValue: (value, header) => {
        seen.push(header)
        return value
      },
    })

    expect(seen).toContain("NAME")
    expect(seen).not.toContain("name")
  })

  it("gets the same header names in the streaming reader", async () => {
    const seen: string[] = []
    for await (const _row of streamCsvRows(CSV, {
      header: true,
      transformHeader: upper,
      transformValue: (value, header) => {
        seen.push(header)
        return value
      },
    })) {
      // drained for the side effect
    }

    expect(seen).toContain("NAME")
  })

  it("does nothing without header: true", () => {
    expect(parseCsv(CSV, { transformHeader: upper })).toEqual([
      ["name", "age"],
      ["Ada", "36"],
    ])
  })

  it("still drops the header row when skipHeaderRow asks", () => {
    expect(parseCsv(CSV, { header: true, skipHeaderRow: true, transformHeader: upper })).toEqual([
      ["Ada", "36"],
    ])
  })
})
