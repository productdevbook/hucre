import { describe, expect, it } from "vitest"
import { assertCellMapCapacity, oversizeSheetMessage } from "../src/xlsx/worksheet"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { ParseError } from "../src/errors"
import { MAX_CELL_MAP_ENTRIES } from "../src/limits"
import type { Cell } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #527 — `sparse: true` is one of the escapes the `maxTotalCells` error
// offers, and `Sheet.cells` is the only place a sparse read puts
// anything. A `Map` in V8 caps at 2^24 entries, so the sparse path has a
// ceiling of its own: a raw `RangeError: Map maximum size exceeded`,
// naming no sheet and saying nothing about spreadsheets.
//
// The two cases the old message could not tell apart:
//
//   large box, mostly empty   82k values over 305M slots   sparse works
//   genuinely dense           28.4M filled of 30.2M        sparse cannot
//
// For the second, the cell count that blew the bounding-box limit is the
// same count that blows the Map — so the advice pointed at the one path
// guaranteed not to work. Four workbooks in a ~600-file corpus are that
// shape, and `streamXlsxRows` reads all four.
//
// Raising the ceiling means `Sheet.cells` stops being a `Map`, which is
// a breaking change to a public field. This does the other thing: stops
// recommending a path already too small, and fails as a `ParseError`.
//
// Both pieces are exported and tested directly. Reaching either branch
// through a real workbook needs 16.7 million filled cells — a couple of
// gigabytes to build for one string — which is not a thing to put in a
// suite that runs on every push.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

describe("the advice depends on which kind of large the sheet is", () => {
  it("offers sparse for a big, mostly empty box", () => {
    // #501's file: 82,000 values over 305,612,208 slots.
    const message = oversizeSheetMessage("Data", 18_654, 16_384, 305_612_208, 82_000, 20_000_000)

    expect(message).toContain("readXlsx(input, { sparse: true })")
    expect(message).toContain("streamXlsxRows")
    expect(message).toContain("0.03% of them filled")
  })

  it("does not offer it for a dense sheet that would not fit", () => {
    // #527's fourth file: 916,449 x 33, 93.95% filled.
    const message = oversizeSheetMessage("Log", 916_449, 33, 30_242_817, 28_413_226, 20_000_000)

    // Not the *recommendation* — the message still names the option, to
    // say why it is not one.
    expect(message).not.toContain("readXlsx(input, { sparse: true })")
    expect(message).toContain("cannot help here")
    expect(message).toContain("28413226 filled cells")
    expect(message).toContain(String(MAX_CELL_MAP_ENTRIES))
    // The escape that does work is still named, and named first.
    expect(message).toContain("streamXlsxRows")
  })

  it("switches exactly at the Map bound, not near it", () => {
    const fits = oversizeSheetMessage("S", 2, 2, 30_000_000, MAX_CELL_MAP_ENTRIES, 20_000_000)
    const over = oversizeSheetMessage("S", 2, 2, 30_000_000, MAX_CELL_MAP_ENTRIES + 1, 20_000_000)

    expect(fits).toContain("readXlsx(input, { sparse: true })")
    expect(over).not.toContain("readXlsx(input, { sparse: true })")
  })

  it("keeps naming range, maxRows and maxTotalCells either way", () => {
    for (const cellCount of [1, MAX_CELL_MAP_ENTRIES + 1]) {
      const message = oversizeSheetMessage("S", 2, 2, 30_000_000, cellCount, 20_000_000)

      expect(message).toContain("`range` or `maxRows`")
      expect(message).toContain("`maxTotalCells`")
      expect(message).toContain('Sheet "S"')
    }
  })
})

describe("the overflow itself is a typed error", () => {
  /** A Map that reports itself full without holding 16.7M entries. */
  function fullMap(size: number, present?: string): Map<string, Cell> {
    const map = new Map<string, Cell>()
    Object.defineProperty(map, "size", { value: size })
    if (present) map.set(present, { value: null, type: "string" })
    return map
  }

  it("throws ParseError, not RangeError, at the bound", () => {
    // V8's own is a real `RangeError` with message "Map maximum size
    // exceeded" and no `code` — checked directly, because the last time
    // a ceiling was guarded here the error type was assumed and the
    // guard never fired. See #516.
    const cells = fullMap(MAX_CELL_MAP_ENTRIES)

    expect(() => assertCellMapCapacity(cells, "0,0", "Log")).toThrow(ParseError)
    expect(() => assertCellMapCapacity(cells, "0,0", "Log")).toThrow(/Sheet "Log"/)
    expect(() => assertCellMapCapacity(cells, "0,0", "Log")).toThrow(/streamXlsxRows/)
    expect(() => assertCellMapCapacity(cells, "0,0", "Log")).toThrow(/not damaged/)
  })

  it("allows the last entry that fits", () => {
    const cells = fullMap(MAX_CELL_MAP_ENTRIES - 1)

    expect(() => assertCellMapCapacity(cells, "0,0", "Log")).not.toThrow()
  })

  it("allows overwriting a key already present when full", () => {
    // `set` on an existing key does not grow the Map, so refusing it
    // would reject a cell the Map can hold.
    const cells = fullMap(MAX_CELL_MAP_ENTRIES, "5,5")

    expect(() => assertCellMapCapacity(cells, "5,5", "Log")).not.toThrow()
  })

  it("names the sheet even when there is no name", () => {
    const cells = fullMap(MAX_CELL_MAP_ENTRIES)

    expect(() => assertCellMapCapacity(cells, "0,0", undefined)).toThrow(ParseError)
  })
})

describe("ordinary sheets are untouched", () => {
  it("a sparse read still returns its cells", async () => {
    const base = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            ["a", "b"],
            [1, 2],
          ],
        },
      ],
    })
    const wb = await readXlsx(base, { sparse: true })

    expect(wb.sheets[0]!.cells?.get("0,0")?.value).toBe("a")
    expect(wb.sheets[0]!.cells?.get("1,1")?.value).toBe(2)
  })

  it("and the oversize error still fires on a big box", async () => {
    // Two cells at opposite corners: a large box, two filled cells, so
    // the message should recommend `sparse` — the case it is for.
    const base = await writeXlsx({ sheets: [{ name: "S", rows: [["seed"]] }] })
    const all = await new ZipReader(base).extractAll()
    const zw = new ZipWriter()
    for (const [path, data] of all) {
      zw.add(
        path,
        path === "xl/worksheets/sheet1.xml"
          ? enc.encode(
              dec
                .decode(data)
                .replace(
                  /<sheetData>.*<\/sheetData>/,
                  '<sheetData><row r="1"><c r="A1" t="str"><v>a</v></c></row>' +
                    '<row r="1048576"><c r="XFD1048576" t="str"><v>z</v></c></row></sheetData>',
                ),
            )
          : data,
      )
    }

    await expect(readXlsx(await zw.build())).rejects.toThrow(/sparse: true/)
  })
})
