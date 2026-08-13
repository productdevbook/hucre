import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { ParseError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// #501 — `Sheet.rows` is a dense rectangle, so a read costs the bounding
// box rather than the cell count. That is right for almost every sheet
// and wrong for a sparse one: a real workbook with 82,000 values
// scattered over a 305,612,208-slot box — 0.03% filled — could not be
// read at all, and none of the three options the error named would help.
//
// Raising `maxTotalCells` trades a clean error for a multi-gigabyte
// allocation; `range` needs the caller to already know where the data
// is; `maxRows` bounds rows when the problem is columns.
//
// Two things fix that. `streamXlsxRows` already read the file and the
// message never said so, and `sparse: true` returns the cells with no
// grid at all.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

function colName(index: number): string {
  let n = index + 1
  let out = ""
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - rem) / 26)
  }
  return out
}

/**
 * A sheet shaped like the one in the issue: values scattered across
 * `cols` columns spread over a box `width` wide.
 *
 * The real file had 507 used columns; these use fewer, because what
 * makes it unreadable is the *box* — 2,000 x 15,312 is 30.6M slots
 * whatever fills it — and a million-value fixture costs the suite ten
 * seconds to prove the same thing.
 */
async function sparseSheet(rowCount: number, cols: number, width: number): Promise<Uint8Array> {
  const used = Array.from({ length: cols }, (_, i) => Math.round((i * (width - 1)) / (cols - 1)))
  const rows: string[] = []
  for (let r = 1; r <= rowCount; r++) {
    rows.push(
      `<row r="${r}">` +
        used.map((c) => `<c r="${colName(c)}${r}"><v>${(r * c) % 97}</v></c>`).join("") +
        "</row>",
    )
  }
  const base = await writeXlsx({ sheets: [{ name: "Sparse", rows: [["seed"]] }] })
  const all = await new ZipReader(base).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === "xl/worksheets/sheet1.xml"
        ? enc.encode(
            dec
              .decode(data)
              .replace(
                /<dimension ref="[^"]*"\/>/,
                `<dimension ref="A1:${colName(width - 1)}${rowCount}"/>`,
              )
              .replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${rows.join("")}</sheetData>`),
          )
        : data,
    )
  }
  return zw.build()
}

describe("the error names the options that actually work", () => {
  it("points at streamXlsxRows and sparse, not only at the three that do not help", async () => {
    const bytes = await sparseSheet(2000, 24, 15312)

    await expect(readXlsx(bytes)).rejects.toThrow(ParseError)
    await expect(readXlsx(bytes)).rejects.toThrow(/streamXlsxRows/)
    await expect(readXlsx(bytes)).rejects.toThrow(/sparse: true/)
  }, 120_000)

  it("says how empty the box actually is", async () => {
    // The number that turns "your sheet is too large" into "your sheet
    // is mostly nothing", which is a different problem with a different
    // answer.
    await expect(readXlsx(await sparseSheet(2000, 24, 15312))).rejects.toThrow(/% of them filled/)
  }, 120_000)
})

describe("sparse: true reads what the grid cannot hold", () => {
  it("returns every value, keyed by position", async () => {
    const wb = await readXlsx(await sparseSheet(2000, 24, 15312), { sparse: true })
    const sheet = wb.sheets[0]!

    expect(sheet.cells?.size).toBe(2000 * 24)
    expect(sheet.cells?.get("0,0")?.value).toBe(0)
    expect(sheet.cells?.get("1999,15311")?.value).toBe(70)
  }, 120_000)

  it("leaves rows empty, because the grid is what it is avoiding", async () => {
    const wb = await readXlsx(await sparseSheet(2000, 24, 15312), { sparse: true })

    expect(wb.sheets[0]!.rows).toEqual([])
  }, 120_000)

  it("streamXlsxRows reads the same file too, a row at a time", async () => {
    // The option that already worked and was never mentioned.
    let count = 0
    for await (const _row of streamXlsxRows(await sparseSheet(200, 24, 15312))) count++

    expect(count).toBe(200)
  }, 120_000)
})

describe("sparse agrees with dense on a sheet both can read", () => {
  const GRID = [
    ["name", "qty", "when"],
    ["Widget", 3, null],
    [null, null, "x"],
  ]

  it("same values at the same positions", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: GRID }] })

    const dense = (await readXlsx(bytes)).sheets[0]!
    const sparse = (await readXlsx(bytes, { sparse: true })).sheets[0]!

    for (let r = 0; r < GRID.length; r++) {
      for (let c = 0; c < 3; c++) {
        const fromDense = dense.rows[r]?.[c] ?? null
        const fromSparse = sparse.cells?.get(`${r},${c}`)?.value ?? null
        expect(fromSparse, `${r},${c}`).toEqual(fromDense)
      }
    }
  })

  it("carries styles when asked, the same as dense does", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          cells: new Map([["0,0", { value: "a", style: { font: { bold: true } } }]]),
        },
      ],
    })

    const sparse = (await readXlsx(bytes, { sparse: true, readStyles: true })).sheets[0]!

    expect(sparse.cells?.get("0,0")?.style?.font?.bold).toBe(true)
  })

  it("keeps formulas and their cached results", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[1, 2, null]],
          cells: new Map([["0,2", { value: 3, formula: "A1+B1", formulaResult: 3 }]]),
        },
      ],
    })

    const cell = (await readXlsx(bytes, { sparse: true })).sheets[0]!.cells?.get("0,2")

    expect(cell?.formula).toBe("A1+B1")
    expect(cell?.formulaResult).toBe(3)
  })

  it("is off by default, so nothing changes for anyone", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: GRID }] })
    const wb = await readXlsx(bytes)

    expect(wb.sheets[0]!.rows).toHaveLength(3)
    expect(wb.sheets[0]!.rows[0]).toEqual(["name", "qty", "when"])
  })
})
