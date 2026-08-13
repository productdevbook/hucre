import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { densify } from "../src/xls/reader"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #494 — `Sheet.rows` is a dense rectangle. `readXlsx` enforced that with
// an explicit pass at the end of its parse; `readXls` and `readXlsb` did
// not, so one authored sheet saved three ways came back three shapes.
//
// The sharper half is not the shape. Both binary readers created
// `rows[n]` only when a cell landed on row n, so a row Excel left empty
// came back as `undefined` — and `CellValue` is
// `string | number | boolean | Date | null`, with no `undefined` member.
// A hole there is a type violation, not an inconsistency.
//
// The XLS reader's own bounding-box guard is sized on the density it did
// not enforce: "their product is not [bounded], and `rows` is a dense
// rectangle". The invariant was assumed by the guard and not implemented
// by the code writing the rows.
// ═══════════════════════════════════════════════════════════════════════

describe("densify", () => {
  it("pads every row to the width", () => {
    const rows: CellValue[][] = [["a"], ["b", "c", "d"], ["e", "f"]]
    densify(rows, 3)

    expect(rows).toEqual([
      ["a", null, null],
      ["b", "c", "d"],
      ["e", "f", null],
    ])
  })

  it("fills a hole rather than leaving undefined", () => {
    // The type violation. A sparse array's missing index is not `null`.
    const rows: CellValue[][] = []
    rows[0] = ["a"]
    rows[2] = ["c"]
    densify(rows, 1)

    expect(rows).toHaveLength(3)
    expect(rows[1]).toEqual([null])
    expect(rows.every((r) => Array.isArray(r))).toBe(true)
    expect(Object.hasOwn(rows, 1)).toBe(true)
  })

  it("leaves an already-dense array alone", () => {
    const rows: CellValue[][] = [
      ["a", "b"],
      ["c", "d"],
    ]
    densify(rows, 2)

    expect(rows).toEqual([
      ["a", "b"],
      ["c", "d"],
    ])
  })

  it("does nothing to an empty sheet", () => {
    const rows: CellValue[][] = []
    densify(rows, 0)

    expect(rows).toEqual([])
  })

  it("never truncates a row that is already wider", () => {
    // Padding only. A width that is somehow short must not lose data.
    const rows: CellValue[][] = [["a", "b", "c"]]
    densify(rows, 2)

    expect(rows[0]).toEqual(["a", "b", "c"])
  })
})

describe("readXlsx still holds the contract it always did", () => {
  it("pads a short row to the sheet width", async () => {
    const wb = await readXlsx(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [
              ["a", "b", "c"],
              ["d", null, null],
            ],
          },
        ],
      }),
    )

    expect(wb.sheets[0]!.rows[1]).toHaveLength(3)
  })

  it("has no undefined holes", async () => {
    const wb = await readXlsx(
      await writeXlsx({
        sheets: [{ name: "S", rows: [["a"], [null], ["c"]] }],
      }),
    )

    for (const row of wb.sheets[0]!.rows) {
      expect(Array.isArray(row)).toBe(true)
      for (const cell of row) expect(cell).not.toBeUndefined()
    }
  })
})
