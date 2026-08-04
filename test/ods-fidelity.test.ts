import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #406 — the README now states what ODS carries. A fidelity table nobody
// checks rots the moment someone widens the model, and a *narrower*
// contract than the code delivers is just as misleading as a wider one.
//
// So this suite pins the table in both directions: the six documented
// facets must survive, and the undocumented ones must not silently start
// surviving without the README catching up. If you widen the ODS model —
// please do, it is deliberately narrow — move the entry here and in the
// README together.
// ═══════════════════════════════════════════════════════════════════════

const readme = (): string => readFileSync(new URL("../README.md", import.meta.url), "utf-8")

/** Every facet the README claims, plus every one it says is absent. */
const style: CellStyle = {
  numFmt: "0.00%",
  font: {
    bold: true,
    italic: true,
    size: 14,
    color: { rgb: "FF0000" },
    // documented as not carried:
    name: "Courier New",
    underline: true,
    strikethrough: true,
  },
  fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
  // documented as not carried:
  border: { top: { style: "thick" }, bottom: { style: "thin" } },
  alignment: { horizontal: "center", vertical: "top", wrapText: true, indent: 2 },
  protection: { locked: false },
}

const roundTrip = async (cell: unknown): Promise<CellStyle | undefined> => {
  const buf = await writeOds({
    sheets: [{ name: "S", rows: [[1]], cells: new Map([["0,0", cell]]) }] as never,
  })
  return (await readOds(buf, { readStyles: true })).sheets[0].cells?.get("0,0")?.style
}

describe("the six facets the README documents", () => {
  it("all survive a write and read", async () => {
    const back = await roundTrip({ value: 1, style })
    expect(back?.font?.bold).toBe(true)
    expect(back?.font?.italic).toBe(true)
    expect(back?.font?.size).toBe(14)
    expect(back?.font?.color?.rgb).toBe("FF0000")
    expect(back?.fill && "fgColor" in back.fill ? back.fill.fgColor?.rgb : undefined).toBe("FFFF00")
    expect(back?.numFmt).toBe("0.00%")
  })
})

describe("the facets the README says are not carried", () => {
  it("are still not carried — update the README before this test", async () => {
    const back = await roundTrip({ value: 1, style })
    expect(back?.font?.name).toBeUndefined()
    expect(back?.font?.underline).toBeUndefined()
    expect(back?.font?.strikethrough).toBeUndefined()
    expect(back?.border).toBeUndefined()
    expect(back?.alignment).toBeUndefined()
    expect(back?.protection).toBeUndefined()
  })

  it("leave a wholly-unsupported style with no reference at all", async () => {
    // Not an empty style — no `table:style-name` is emitted, so the cell
    // comes back with nothing rather than with a blank definition. Worth
    // pinning because "the style is there but empty" and "there is no
    // style" behave differently for a caller checking `style !== undefined`.
    const back = await roundTrip({ value: 1, style: { border: { top: { style: "thick" } } } })
    expect(back).toBeUndefined()
  })

  it("collapse two styles that differ only in an unsupported facet", async () => {
    // The dedupe key hashes the six supported facets only, so these two
    // share one definition. Documented in the README as a consequence;
    // pinned here so it stays a known trade-off rather than a surprise.
    const buf = await writeOds({
      sheets: [
        {
          name: "S",
          rows: [[1, 2]],
          cells: new Map([
            [
              "0,0",
              { value: 1, style: { font: { bold: true }, border: { top: { style: "thick" } } } },
            ],
            [
              "0,1",
              { value: 2, style: { font: { bold: true }, border: { top: { style: "thin" } } } },
            ],
          ]),
        },
      ] as never,
    })
    const cells = (await readOds(buf, { readStyles: true })).sheets[0].cells
    expect(cells?.get("0,0")?.style?.font?.bold).toBe(true)
    expect(cells?.get("0,1")?.style?.font?.bold).toBe(true)
    expect(cells?.get("0,0")?.style?.border).toBeUndefined()
    expect(cells?.get("0,1")?.style?.border).toBeUndefined()
  })
})

describe("ODS → ODS is lossless, which is the point of the table", () => {
  it("returns everything it wrote", async () => {
    // Dates are back in as of #415. They used to drift by the local UTC
    // offset on every round trip, cumulatively — the writer emitted UTC
    // components and the reader parsed the unqualified string as local
    // time — which is silent on a UTC machine and so invisible to CI.
    // Keep them here: this assertion is the one that would notice.
    const buf = await writeOds({
      sheets: [{ name: "S", rows: [["a", 1, true, new Date(Date.UTC(2024, 0, 15))]] }],
    })
    const first = await readOds(buf)
    const second = await readOds(await writeOds({ sheets: first.sheets as never }))
    expect(second.sheets[0].rows[0]).toEqual(first.sheets[0].rows[0])
  })
})

describe("the README says so", () => {
  it("carries the fidelity table", () => {
    const text = readme()
    expect(text).toContain("#### What ODS carries")
    expect(text).toContain("**six facets only**")
  })

  it("qualifies the comparison-table row rather than claiming a bare Yes", () => {
    const text = readme()
    expect(text).toMatch(/\*\*ODS support\*\*\s*\|\s*Yes<sup>§<\/sup>/)
    expect(text).toContain("ODS → ODS is\nlossless while XLSX → ODS drops them")
  })
})
