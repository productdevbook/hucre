import { describe, expect, it } from "vitest"
import * as root from "../src/index"
import * as xlsxEntry from "../src/xlsx"
import { writeXlsx } from "../src/xlsx/writer"
import { writeOds } from "../src/ods/writer"
import { readXlsx } from "../src/xlsx/reader"
import { readOds } from "../src/ods/reader"
import { XlsxStreamWriter } from "../src/xlsx/stream-writer"
import { copyRange } from "../src/sheet-ops"
import { toRange, toRanges } from "../src/cell-utils"
import type { MergeRange, Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #474, the two API-shape items.
//
// 1. Cell utilities were split across two entry points, so someone on
//    `hucre/xlsx` who wanted `letterToCol` — a pure string helper with
//    nothing XLSX-specific about it — had to pull the root as well. The
//    JSON surface had exactly this disagreement and it was settled before
//    v1; this one was missed.
//
// 2. Ranges were A1 strings in half the API and coordinate objects in the
//    other half, with no rule to hold in your head about which a field
//    wanted. The authoring surfaces now take either.
// ═══════════════════════════════════════════════════════════════════════

const CELL_UTILS = [
  "parseCellRef",
  "colToLetter",
  "cellRef",
  "rangeRef",
  "letterToCol",
  "parseRange",
  "isInRange",
  "r1c1ToA1",
  "a1ToR1C1",
] as const

describe("the cell utilities are on both entry points", () => {
  it("hucre/xlsx carries all nine, not four", () => {
    for (const name of CELL_UTILS) {
      expect(xlsxEntry, `hucre/xlsx is missing ${name}`).toHaveProperty(name)
    }
  })

  it("and they are the same functions, not copies", () => {
    for (const name of CELL_UTILS) {
      expect(xlsxEntry[name], name).toBe(root[name])
    }
  })

  it("the one that prompted this actually works from here", () => {
    expect(xlsxEntry.letterToCol("AA")).toBe(26)
    expect(xlsxEntry.colToLetter(26)).toBe("AA")
  })
})

describe("toRange normalises either spelling", () => {
  const coords: MergeRange = { startRow: 0, startCol: 0, endRow: 2, endCol: 3 }

  it("parses a string and passes coordinates through", () => {
    expect(toRange("A1:D3")).toEqual(coords)
    expect(toRange(coords)).toBe(coords)
  })

  it("treats a single cell as a one-cell range", () => {
    expect(toRange("B2")).toEqual({ startRow: 1, startCol: 1, endRow: 1, endCol: 1 })
  })

  it("does not reallocate a list that needs no work", () => {
    // The common case is coordinates already; walking and rebuilding the
    // array for nothing would be a cost on every sheet written.
    const list = [coords]
    expect(toRanges(list)).toBe(list)
    expect(toRanges(undefined)).toBeUndefined()
  })

  it("normalises a mixed list", () => {
    expect(toRanges(["A1:B1", coords])).toEqual([
      { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      coords,
    ])
  })
})

describe("merges accept an A1 string", () => {
  it("writeXlsx writes the same file either way", async () => {
    const asString = await writeXlsx({
      sheets: [{ name: "S", rows: [["a", "b", "c"]], merges: ["A1:C1"] }],
    })
    const asCoords = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a", "b", "c"]],
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
        },
      ],
    })

    const fromString = await readXlsx(asString)
    const fromCoords = await readXlsx(asCoords)

    expect(fromString.sheets[0]!.merges).toEqual([
      { startRow: 0, startCol: 0, endRow: 0, endCol: 2 },
    ])
    expect(fromString.sheets[0]!.merges).toEqual(fromCoords.sheets[0]!.merges)
  })

  it("the read model still hands back coordinates", async () => {
    // Only the authoring side widened. The reader produces one form,
    // which is what makes it usable without a normalise call.
    const wb = await readXlsx(
      await writeXlsx({ sheets: [{ name: "S", rows: [["a", "b"]], merges: ["A1:B1"] }] }),
    )

    expect(typeof wb.sheets[0]!.merges![0]).toBe("object")
  })

  it("writeOds too", async () => {
    const wb = await readOds(
      await writeOds({ sheets: [{ name: "S", rows: [["a", "b", "c"]], merges: ["A1:C1"] }] }),
    )

    expect(wb.sheets[0]!.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }])
  })

  it("XlsxStreamWriter too", async () => {
    const w = new XlsxStreamWriter({ name: "S", merges: ["A1:C1"] })
    w.addRow(["a", "b", "c"])

    const wb = await readXlsx(await w.finish())
    expect(wb.sheets[0]!.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }])
  })

  it("a mixed list works, because there is no reason it should not", async () => {
    const wb = await readXlsx(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [
              ["a", "b", "c"],
              ["d", "e", "f"],
            ],
            merges: ["A1:B1", { startRow: 1, startCol: 0, endRow: 1, endCol: 1 }],
          },
        ],
      }),
    )

    expect(wb.sheets[0]!.merges).toHaveLength(2)
  })
})

describe("copyRange accepts A1 too", () => {
  function grid(): Sheet {
    return {
      name: "S",
      rows: [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
      ],
    }
  }

  it("moves the same block whichever way it is asked", () => {
    const a = grid()
    copyRange(a, "A1:B2", "D1")

    const b = grid()
    copyRange(b, { startRow: 0, startCol: 0, endRow: 1, endCol: 1 }, { startRow: 0, startCol: 3 })

    expect(a.rows).toEqual(b.rows)
    expect(a.rows[0]![3]).toBe(1)
    expect(a.rows[1]![4]).toBe(5)
  })

  it("takes one form for the source and the other for the target", () => {
    const s = grid()
    copyRange(s, "A1:A1", { startRow: 2, startCol: 2 })

    expect(s.rows[2]![2]).toBe(1)
  })
})
