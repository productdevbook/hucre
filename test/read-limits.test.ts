import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { readOds } from "../src/ods/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { writeOds } from "../src/ods/writer"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { MAX_TOTAL_CELLS } from "../src/limits"
import { ParseError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// #471 — every bound in limits.ts says what it defends against and why
// the number is what it is, and five of the six could not be moved. The
// defence is right; a defence with no escape hatch is also a ceiling, and
// a legitimate 25-million-cell sheet was simply unreadable.
//
// The defaults do not change. What changes is that a caller who knows
// their input can say so, and can now name the number when they do.
// ═══════════════════════════════════════════════════════════════════════

/** Rebuild an archive with one part rewritten. */
async function patch(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string,
): Promise<Uint8Array> {
  const all = await new ZipReader(bytes).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === path ? new TextEncoder().encode(edit(new TextDecoder().decode(data))) : data,
    )
  }
  return zw.build()
}

/**
 * An XLSX holding two corner cells at `A1` and `<lastCol><lastRow>` — a
 * few hundred bytes of XML describing a bounding box of any size you
 * like. This is the exact shape MAX_TOTAL_CELLS exists to refuse when it
 * is hostile, and the exact shape a real sparse export has when it is not.
 *
 * Both corners stay inside Excel's own row/column bounds, or the
 * coordinate check fires first and this proves nothing about the product.
 */
async function spanning(rows: number, lastCol = "T"): Promise<Uint8Array> {
  const base = await writeXlsx({ sheets: [{ name: "S", rows: [["a", "b"]] }] })
  return patch(base, "xl/worksheets/sheet1.xml", (xml) =>
    xml.replace(
      /<sheetData>.*<\/sheetData>/,
      `<sheetData><row r="1"><c r="A1" t="str"><v>a</v></c></row>` +
        `<row r="${rows}"><c r="${lastCol}${rows}" t="str"><v>b</v></c></row></sheetData>`,
    ),
  )
}

/** 1,048,576 rows x 20 columns = 20,971,520 — just over the 20M default. */
const OVER_DEFAULT = 1_048_576

describe("maxTotalCells", () => {
  it("still refuses the sparse-corner case by default", async () => {
    // 20,971,520 cells from two cells of XML. The whole point of the bound.
    await expect(readXlsx(await spanning(OVER_DEFAULT))).rejects.toThrow(ParseError)
  })

  it("names the option in the error, not just the number", async () => {
    // An error that states a ceiling without saying it can move sends
    // people to `range`/`maxRows`, which change what you get rather than
    // what is allowed.
    await expect(readXlsx(await spanning(OVER_DEFAULT))).rejects.toThrow(/maxTotalCells/)
  })

  it("moves in both directions, and the sheet is read either way", async () => {
    // A 1000 x 20 sparse span: 20,000 cells. Refused under a bound below
    // it, read in full under one above it — same file, same reader, the
    // only difference being that the caller said so.
    //
    // The 25-million-cell case from the issue is this same branch
    // (`ctx.maxTotalCells ?? MAX_TOTAL_CELLS`) at a size no test should
    // allocate: 21M slots is a few hundred MB of array before any value
    // lands in it.
    const bytes = await spanning(1000)

    await expect(readXlsx(bytes, { maxTotalCells: 5000 })).rejects.toThrow(/over the 5000 limit/)

    const wb = await readXlsx(bytes, { maxTotalCells: 25_000 })
    expect(wb.sheets[0]!.rows.length).toBe(1000)
    expect(wb.sheets[0]!.rows[0]![0]).toBe("a")
    expect(wb.sheets[0]!.rows[999]![19]).toBe("b")
  })

  it("tightens below the default, which was equally impossible", async () => {
    // A service reading files it did not choose wants a bound well under
    // 20M, and had no way to set one.
    const small = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            [1, 2, 3],
            [4, 5, 6],
          ],
        },
      ],
    })

    await expect(readXlsx(small, { maxTotalCells: 4 })).rejects.toThrow(/over the 4 limit/)
    expect((await readXlsx(small, { maxTotalCells: 6 })).sheets[0]!.rows.length).toBe(2)
  })

  it("is honoured by the ODS reader too", async () => {
    const small = await writeOds({
      sheets: [
        {
          name: "S",
          rows: [
            [1, 2, 3],
            [4, 5, 6],
          ],
        },
      ],
    })

    await expect(readOds(small, { maxTotalCells: 2 })).rejects.toThrow(ParseError)
    expect((await readOds(small, { maxTotalCells: 100 })).sheets[0]!.rows.length).toBe(2)
  })
})

describe("maxDecompressedBytes", () => {
  it("refuses an entry that expands past the bound", async () => {
    // Deflate compresses a run of one byte to almost nothing, which is
    // the whole mechanism of a zip bomb.
    const zw = new ZipWriter()
    zw.add("[Content_Types].xml", new Uint8Array(0))
    zw.add("big.txt", new TextEncoder().encode("A".repeat(200_000)))
    const bomb = await zw.build()

    const tight = new ZipReader(bomb, 1000)
    await expect(tight.extract("big.txt")).rejects.toThrow()

    // The same archive, same reader class, a bound that fits.
    const loose = new ZipReader(bomb, 500_000)
    expect((await loose.extract("big.txt")).length).toBe(200_000)
  })

  it("reaches the ZIP layer from a read option", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [["x".repeat(5000)]] }],
    })

    // 100 bytes cannot hold any part of a real workbook.
    await expect(readXlsx(bytes, { maxDecompressedBytes: 100 })).rejects.toThrow()
    // And the default path is unaffected.
    expect((await readXlsx(bytes)).sheets[0]!.rows[0]![0]).toBe("x".repeat(5000))
  })
})

describe("the constants themselves", () => {
  it("are the numbers the readers actually use", async () => {
    // Exported so they can be quoted; worth proving the export is not a
    // copy that has drifted from the enforcement.
    const bytes = await spanning(OVER_DEFAULT)

    await expect(readXlsx(bytes)).rejects.toThrow(new RegExp(`over the ${MAX_TOTAL_CELLS} limit`))
  })
})
