import { describe, expect, it } from "vitest"
import { readFileSync, readdirSync } from "node:fs"
import { readXlsx } from "../src/xlsx/reader"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import type { CellValue, Sheet } from "../src/_types"
import { bufferedRowIndexes } from "./_xlsx-rows"

// ═══════════════════════════════════════════════════════════════════════
// XLSX has two readers and they had never been compared. The ODS pair
// were, and that comparison found #525; the XLSX pair had only spot
// checks — a chartsheet here, a row count there — and no sweep.
//
// It found the same class of bug. #496 taught the *buffered* reader
// openpyxl's `t="d"` cells, the ISO 8601 date form of ST_CellType, and
// left the streaming reader on the `n` path where `Number("2024-03-17")`
// is NaN and the value falls through as a string:
//
//     readXlsx         Date 2024-03-17T00:00:00.000Z
//     streamXlsxRows   "2024-03-17"
//
// One fix, two implementations, and only one of them got it.
//
// Two documented differences the comparison has to allow for, both in
// `docs/PARITY.md`:
//
//   - `readXlsx` pads every row to the sheet's bounding box; the stream
//     yields the row as the file had it. Compared on the trimmed form.
//   - the stream skips a row with no cell `readXlsx` would count — a
//     `<row>` of style-only cells, or of no cells at all — and keeps the
//     true index, so rows are matched by `index`, not by position.
//
// And one that was *not* documented until this test found it: this reads
// one sheet, where `streamOdsRows` reads them all. See below.
// ═══════════════════════════════════════════════════════════════════════

const norm = (v: CellValue): unknown => (v instanceof Date ? `D:${v.toISOString()}` : v)

/** `readXlsx` pads to the bounding box; the stream does not. */
function trim(row: CellValue[]): unknown[] {
  const out = [...row]
  while (out.length > 0 && out[out.length - 1] === null) out.pop()
  return out.map(norm)
}

function corpus(): string[] {
  const out: string[] = []
  for (const dir of ["test/fixtures", "test/fixtures/third-party"]) {
    for (const file of readdirSync(dir)) {
      if (file.endsWith(".xlsx")) out.push(`${dir}/${file}`)
    }
  }
  return out.sort()
}

/**
 * Each sheet's rows as `readXlsx` has them, by row index.
 *
 * Dense where it can be. `excel-sparse.xlsx` is over the dense read's
 * documented limit — `test/sparse-read.test.ts` owns that case — so it
 * is read sparse and its cells laid out by row instead of being left
 * out of the comparison.
 */
async function bufferedSheets(
  bytes: Uint8Array,
  options: { dates?: "serial" } | undefined,
): Promise<Array<Map<number, CellValue[]>>> {
  let sheets: Sheet[]
  let sparse = false
  try {
    sheets = (await readXlsx(bytes, options)).sheets
  } catch {
    sheets = (await readXlsx(bytes, { ...options, sparse: true })).sheets
    sparse = true
  }
  return sheets.map((sheet) => {
    const rows = new Map<number, CellValue[]>()
    // Not `sheet.cells` as the test: a dense read fills it too, for the
    // cells that carry more than a value.
    if (!sparse) {
      sheet.rows.forEach((values, index) => rows.set(index, values))
      return rows
    }
    for (const [key, cell] of sheet.cells ?? []) {
      const [r, c] = key.split(",").map(Number) as [number, number]
      const values = rows.get(r) ?? []
      values[c] = cell.value
      rows.set(r, values)
    }
    for (const [r, values] of rows)
      rows.set(
        r,
        Array.from(values, (v) => v ?? null),
      )
    return rows
  })
}

describe("the two XLSX readers agree, on files hucre did not write", () => {
  // `dates: "serial"` too: each reader has its own date branch, which is
  // how #496 reached one of them and not the other. See #600.
  it.each([
    ["by default", undefined],
    ['with dates: "serial"', { dates: "serial" }],
  ] as const)("row for row, across the whole corpus, %s", async (_label, options) => {
    const diffs: string[] = []
    let sheetsCompared = 0

    for (const path of corpus()) {
      const bytes = new Uint8Array(readFileSync(path))
      const sheets = await bufferedSheets(bytes, options)

      // Every sheet, not only the first. This compared sheet 0 alone, and
      // the row-set sweep below was all that looked at the rest — so a
      // value the stream got wrong on a second sheet had nothing to fail.
      for (let s = 0; s < sheets.length; s++) {
        const rows = sheets[s]!
        sheetsCompared++

        const seen = new Set<number>()
        for await (const row of streamXlsxRows(bytes, { ...options, sheet: s })) {
          const streamed = JSON.stringify(trim(row.values ?? []))
          const buffered = JSON.stringify(trim(rows.get(row.index) ?? []))

          if (streamed !== buffered && diffs.length < 8) {
            diffs.push(
              `${path} sheet ${s} row ${row.index}\n  stream ${streamed}\n  buffer ${buffered}`,
            )
          }
          seen.add(row.index)
        }

        // Anything the stream left out has to have been empty.
        for (const [index, values] of rows) {
          if (seen.has(index)) continue
          if (values.some((v) => v !== null && v !== "") && diffs.length < 8) {
            diffs.push(
              `${path} sheet ${s} row ${index} skipped but not empty: ${JSON.stringify(values.map(norm))}`,
            )
          }
        }
      }
    }

    // A floor, so a corpus that stopped loading cannot pass by comparing nothing.
    expect(sheetsCompared).toBeGreaterThan(20)
    expect(diffs).toEqual([])
  })
})

describe("the two XLSX readers agree on which rows exist", () => {
  // The sweep above could not see this. A row the stream yields is
  // compared against the buffered row at its index, and an all-null row
  // trims to `[]` on both sides — so `excel-styleonly.xlsx` streamed four
  // rows against `readXlsx`'s two and passed. Row *sets* are compared here.
  //
  // The buffered side is `bufferedRowIndexes`, shared with
  // `test/xlsx-stream-empty-rows.test.ts` so the two cannot read
  // "has content" differently. Sparse mode lets `excel-sparse.xlsx`
  // take part, which the dense read rejects.

  it("sheet by sheet, across the whole corpus", async () => {
    const diffs: string[] = []
    let sheetsCompared = 0

    for (const path of corpus()) {
      const bytes = new Uint8Array(readFileSync(path))
      const { sheets } = await readXlsx(bytes, { sparse: true })

      for (let s = 0; s < sheets.length; s++) {
        const streamed: number[] = []
        for await (const row of streamXlsxRows(bytes, { sheet: s })) streamed.push(row.index)
        const buffered = bufferedRowIndexes(sheets[s]!)
        sheetsCompared++

        if (JSON.stringify(streamed) !== JSON.stringify(buffered)) {
          diffs.push(`${path} sheet ${s}\n  stream ${streamed}\n  buffer ${buffered}`)
        }
      }
    }

    expect(sheetsCompared).toBeGreaterThan(20)
    expect(diffs).toEqual([])
  })
})

describe("an ISO date cell is a Date in both readers", () => {
  // openpyxl writes `t="d"` whenever `iso_dates=True`. It is the one
  // member of ST_CellType that carries a date rather than a serial, and
  // the streaming reader had no case for it.
  const FIXTURE = "test/fixtures/openpyxl-isodates.xlsx"

  async function streamed(): Promise<CellValue[][]> {
    const rows: CellValue[][] = []
    for await (const row of streamXlsxRows(new Uint8Array(readFileSync(FIXTURE)))) {
      rows[row.index] = row.values
    }
    return rows
  }

  it("a plain date", async () => {
    const rows = await streamed()

    expect(rows[0]![1]).toBeInstanceOf(Date)
    expect((rows[0]![1] as Date).toISOString()).toBe("2024-03-17T00:00:00.000Z")
  })

  it("a date-time, read as UTC when it names no zone", async () => {
    const rows = await streamed()

    expect((rows[1]![1] as Date).toISOString()).toBe("2024-03-17T13:45:30.000Z")
  })

  it("matching the buffered reader exactly", async () => {
    const buffered = (await readXlsx(new Uint8Array(readFileSync(FIXTURE)))).sheets[0]!.rows
    const rows = await streamed()

    for (let i = 0; i < buffered.length; i++) {
      expect(trim(rows[i] ?? []), `row ${i}`).toEqual(trim(buffered[i]!))
    }
  })
})

describe("the sheet each streaming reader walks", () => {
  it("streamXlsxRows reads one sheet, and `sheet` chooses it", async () => {
    // Not a defect, but not obvious either, and it differs from
    // `streamOdsRows`, which walks every sheet in the document. A caller
    // moving between the two formats loses sheets silently otherwise.
    const bytes = new Uint8Array(readFileSync("test/fixtures/third-party/multi-sheet.xlsx"))

    const fromDefault: CellValue[][] = []
    for await (const row of streamXlsxRows(bytes)) fromDefault.push(row.values)

    const fromSecond: CellValue[][] = []
    for await (const row of streamXlsxRows(bytes, { sheet: 1 })) fromSecond.push(row.values)

    const wb = await readXlsx(bytes)
    expect(wb.sheets.length).toBeGreaterThan(1)
    expect(trim(fromDefault[0]!)).toEqual(trim(wb.sheets[0]!.rows[0]!))
    expect(trim(fromSecond[0]!)).toEqual(trim(wb.sheets[1]!.rows[0]!))
    expect(trim(fromDefault[0]!)).not.toEqual(trim(fromSecond[0]!))
  })
})
