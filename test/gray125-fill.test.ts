import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"

// #643: a cell whose xf applies fill index 1 (gray125) read back with no fill.
describe("gray125 fill (#643)", () => {
  it("round-trips a cell styled with the gray125 pattern", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [[1, 2]],
          cells: new Map([
            ["0,0", { value: 1, style: { fill: { type: "pattern", pattern: "gray125" } } }],
          ]),
        },
      ],
    })
    const wb = await readXlsx(buf, { readStyles: true })
    expect(wb.sheets[0].cells?.get("0,0")?.style?.fill).toEqual({
      type: "pattern",
      pattern: "gray125",
    })
    expect(wb.sheets[0].cells?.get("0,1")?.style?.fill).toBeUndefined()
  })
})
