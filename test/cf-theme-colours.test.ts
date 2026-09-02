import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"

// A colour-scale stop, a data-bar fill and a sparkline series are colours
// like any other in the model. v1 typed them as RGB strings, so a theme
// colour — what Excel writes when a user picks one from the palette —
// read back as "" and was written back as rgb="". See v2 migration guide.

describe("conditional-format and sparkline colours are Color objects", () => {
  it("round-trips a theme colour on a colour scale and a data bar", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[1], [2], [3]],
          conditionalRules: [
            {
              type: "colorScale",
              priority: 1,
              range: "A1:A3",
              colorScale: {
                cfvo: [{ type: "min" }, { type: "max" }],
                colors: [{ theme: 5, tint: -0.25 }, { rgb: "63BE7B" }],
              },
            },
            {
              type: "dataBar",
              priority: 2,
              range: "A1:A3",
              dataBar: { cfvo: [{ type: "min" }, { type: "max" }], color: { theme: 4 } },
            },
          ],
        },
      ],
    })
    const wb = await readXlsx(bytes)
    const [scale, bar] = wb.sheets[0]!.conditionalRules!
    expect(scale!.colorScale!.colors).toEqual([{ theme: 5, tint: -0.25 }, { rgb: "63BE7B" }])
    expect(bar!.dataBar!.color).toEqual({ theme: 4 })
  })

  it("round-trips a theme colour on a sparkline series", async () => {
    const bytes = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[1, 2, 3]],
          sparklines: [{ location: "D1", dataRange: "S!A1:C1", color: { theme: 6, tint: 0.4 } }],
        },
      ],
    })
    const wb = await readXlsx(bytes)
    expect(wb.sheets[0]!.sparklines![0]!.color).toEqual({ theme: 6, tint: 0.4 })
  })
})
