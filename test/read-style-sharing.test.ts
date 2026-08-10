import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { cloneCellStyle } from "../src/_style"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §P asked whether a resolved style should be copied per cell. It
// should not: xl/styles.xml holds one font/fill/border record per
// distinct format, and copying it per cell nearly doubles peak memory on
// a styled read — measured at 407 MB against 787 MB over 720,000 styled
// cells, for a guarantee most callers never need.
//
// So sharing stays, and these pin the two halves of the contract: the
// sharing itself, and `cloneCellStyle` as the supported way out for a
// caller who does intend to edit one cell's format.
//
// 6-character RGB throughout: the writer prefixes the opaque alpha and
// the reader strips it again, so the values round-trip exactly.
// ═══════════════════════════════════════════════════════════════════════

const SHARED: CellStyle = {
  font: { name: "Arial", size: 9, bold: true, color: { rgb: "FF0000" } },
  fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "00FF00" } },
  border: { top: { style: "thin", color: { rgb: "0000FF" } } },
  alignment: { horizontal: "center" },
  protection: { locked: false },
}

async function twoStyledCells() {
  const bytes = await writeXlsx({
    sheets: [{ name: "S", rows: [["a"], ["b"]], columns: [{ style: SHARED }] }],
  })
  const wb = await readXlsx(bytes, { readStyles: true })
  const cells = wb.sheets[0]!.cells!
  return { first: cells.get("0,0")!.style!, second: cells.get("1,0")!.style! }
}

describe("a resolved style's parts are shared between cells of the same format", () => {
  it("hands both cells the same font, fill, border, alignment and protection", async () => {
    const { first, second } = await twoStyledCells()

    expect(first.font).toBe(second.font)
    expect(first.fill).toBe(second.fill)
    expect(first.border).toBe(second.border)
    expect(first.alignment).toBe(second.alignment)
    expect(first.protection).toBe(second.protection)
  })

  it("gives each cell its own CellStyle wrapper", async () => {
    const { first, second } = await twoStyledCells()

    // The wrapper differs even though its contents are shared, which is
    // what makes the aliasing easy to miss.
    expect(first).not.toBe(second)
    expect(first).toEqual(second)
  })
})

describe("cloneCellStyle is the way to edit one cell's format", () => {
  it("detaches every nested object", async () => {
    const { first, second } = await twoStyledCells()

    const mine = cloneCellStyle(first)

    expect(mine.font).not.toBe(first.font)
    expect(mine.font!.color).not.toBe(first.font!.color)
    expect(mine.fill).not.toBe(first.fill)
    expect(mine.border).not.toBe(first.border)
    expect(mine.border!.top).not.toBe(first.border!.top)
    expect(mine.alignment).not.toBe(first.alignment)
    expect(mine.protection).not.toBe(first.protection)

    mine.font!.name = "Manrope"
    mine.font!.color!.rgb = "123456"
    mine.border!.top!.style = "thick"

    expect(second.font!.name).toBe("Arial")
    expect(second.font!.color!.rgb).toBe("FF0000")
    expect(second.border!.top!.style).toBe("thin")
  })

  it("copies the same values, not just the same shape", async () => {
    const { first } = await twoStyledCells()

    expect(cloneCellStyle(first)).toEqual(first)
  })

  it("detaches a gradient fill's stops", () => {
    const gradient: CellStyle = {
      fill: {
        type: "gradient",
        degree: 90,
        stops: [
          { position: 0, color: { rgb: "FFFFFF" } },
          { position: 1, color: { rgb: "000000" } },
        ],
      },
    }

    const copy = cloneCellStyle(gradient)

    if (copy.fill?.type !== "gradient" || gradient.fill?.type !== "gradient") {
      throw new Error("not a gradient")
    }
    expect(copy.fill.stops[0]).not.toBe(gradient.fill.stops[0])
    expect(copy.fill.stops[0]!.color).not.toBe(gradient.fill.stops[0]!.color)

    copy.fill.stops[0]!.color.rgb = "123456"

    expect(gradient.fill.stops[0]!.color.rgb).toBe("FFFFFF")
  })

  it("leaves absent facets absent", () => {
    expect(cloneCellStyle({})).toEqual({})
    expect(cloneCellStyle({ numFmt: "0.00" })).toEqual({ numFmt: "0.00" })
  })
})
