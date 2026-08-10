import { describe, expect, it } from "vitest"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { writeXlsx } from "../src/xlsx/writer"
import { ZipReader } from "../src/zip/reader"

// ═══════════════════════════════════════════════════════════════════════
// #439 §AY asked for a guard: `ZipWriter.add` accepted two entries with
// one path and `extract` returned whichever came last. I filed it as a
// guard-rail request, since ZipWriter is internal.
//
// Adding the guard found a shipping bug. `saveXlsx` preserved the opened
// package's `xl/featurePropertyBag/featurePropertyBag.xml` *and* emitted
// its own, so a round-tripped workbook with Excel 2024 checkboxes carried
// the part twice — which Excel treats as damaged. The metadata part had
// an explicit "skip the opened copy" rule; the feature bag never got one.
// ═══════════════════════════════════════════════════════════════════════

const CHECKBOX_CELLS = new Map([["0,0", { value: true, checkbox: true }]])

function pathsIn(bytes: Uint8Array): string[] {
  return new ZipReader(bytes).entries()
}

function duplicates(paths: string[]): string[] {
  const seen = new Set<string>()
  const twice = new Set<string>()
  for (const path of paths) {
    if (seen.has(path)) twice.add(path)
    seen.add(path)
  }
  return [...twice]
}

describe("a saved package has each part once", () => {
  it("round-trips a workbook with checkboxes without duplicating the feature bag", async () => {
    const original = await writeXlsx({
      sheets: [{ name: "S", rows: [[true]], cells: CHECKBOX_CELLS as never }],
    })

    // The part has to be there in the first place, or the test proves nothing.
    expect(pathsIn(original)).toContain("xl/featurePropertyBag/featurePropertyBag.xml")

    const saved = await saveXlsx(await openXlsx(original))

    expect(duplicates(pathsIn(saved))).toEqual([])
    expect(pathsIn(saved)).toContain("xl/featurePropertyBag/featurePropertyBag.xml")
  })

  it("survives a second round trip", async () => {
    const original = await writeXlsx({
      sheets: [{ name: "S", rows: [[true]], cells: CHECKBOX_CELLS as never }],
    })

    const once = await saveXlsx(await openXlsx(original))
    const twice = await saveXlsx(await openXlsx(once))

    expect(duplicates(pathsIn(twice))).toEqual([])
  })

  it("keeps the checkbox through the round trip", async () => {
    const original = await writeXlsx({
      sheets: [{ name: "S", rows: [[true]], cells: CHECKBOX_CELLS as never }],
    })

    const saved = await saveXlsx(await openXlsx(original))
    const xml = new TextDecoder().decode(
      await new ZipReader(saved).extract("xl/featurePropertyBag/featurePropertyBag.xml"),
    )

    expect(xml).toContain("Checkbox")
  })

  it("has no duplicates on an ordinary workbook either", async () => {
    const original = await writeXlsx({
      sheets: [
        { name: "One", rows: [["a", 1]] },
        { name: "Two", rows: [["b", 2]] },
      ],
      properties: { title: "T" },
    })

    const saved = await saveXlsx(await openXlsx(original))

    expect(duplicates(pathsIn(saved))).toEqual([])
  })
})
