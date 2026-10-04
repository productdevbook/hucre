import { getCell } from "../src/cell-store"
import { describe, expect, it } from "vitest"
import { cellError } from "../src/cell-error"
import { readXlsx, writeXlsx, openXlsx, saveXlsx } from "../src/xlsx"
import { xlsxWithCells } from "./support/xlsx"
import type { CellValue } from "../src/_types"

const results: Array<{ type: string; xml: string; value: CellValue }> = [
  { type: "n", xml: "42", value: 42 },
  { type: "str", xml: "xy", value: "xy" },
  { type: "b", xml: "1", value: true },
  { type: "e", xml: "#DIV/0!", value: cellError("#DIV/0!") },
  { type: "d", xml: "2024-02-29T00:00:00Z", value: new Date("2024-02-29T00:00:00Z") },
  { type: "s", xml: "0", value: "shared text" },
]

describe("empty formula presence is independent of cached result type", () => {
  it("retains an empty formula without a cached value", async () => {
    const workbook = await readXlsx(await xlsxWithCells('<c r="A1"><f/></c>'))
    expect(getCell(workbook.sheets[0].cells, 0, 0)).toMatchObject({
      formula: "",
      type: "formula",
      value: null,
    })
    const again = await readXlsx(await writeXlsx(workbook))
    expect(getCell(again.sheets[0].cells, 0, 0)?.formula).toBe("")
  })

  for (const formula of ["<f/>", "<f></f>", '<f t="shared" si="0"/>']) {
    for (const result of results) {
      it(`${formula} preserves ${result.type} results through both saves`, async () => {
        const bytes = await xlsxWithCells(
          `<c r="A1" t="${result.type}">${formula}<v>${result.xml}</v></c>`,
          "shared text",
        )
        const workbook = await readXlsx(bytes)
        const cell = getCell(workbook.sheets[0].cells, 0, 0)
        expect(cell?.type).toBe("formula")
        expect(cell?.formula).toBe("")
        expect(cell?.formulaResult).toEqual(result.value)
        for (const output of [await writeXlsx(workbook), await saveXlsx(await openXlsx(bytes))]) {
          const again = await readXlsx(output)
          expect(again.sheets[0].rows[0][0]).toEqual(result.value)
          expect(getCell(again.sheets[0].cells, 0, 0)?.formulaResult).toEqual(result.value)
        }
      })
    }
  }
})
