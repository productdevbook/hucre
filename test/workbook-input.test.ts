import { createCellStore, getCell } from "../src/cell-store"
import { describe, expect, it } from "vitest"
import { readXlsx, writeXlsx } from "../src/xlsx"
import { readOds, writeOds } from "../src/ods"
import type { Workbook } from "../src/_types"

const workbook = (): Workbook => ({
  sheets: [
    {
      name: "Report",
      rows: [
        ["Amount", "Approved"],
        [125, true],
      ],
      freezePane: { rows: 1 },
      cells: createCellStore([[1, 0, { value: 125, type: "number", style: { numFmt: "0.00" } }]]),
      charts: [
        { kinds: ["bar"], series: [], seriesCount: 0, anchor: { from: { row: 5, col: 0 } } },
      ],
    },
  ],
  properties: { title: "Report" },
})

describe("one workbook model at the writer boundary", () => {
  it("writes a read workbook directly, preserving writable metadata", async () => {
    const input = workbook()
    const output = await readXlsx(await writeXlsx(input), { readStyles: true })
    expect(output.sheets[0].rows).toEqual(input.sheets[0].rows)
    expect(output.sheets[0].freezePane).toEqual({ rows: 1 })
    expect(getCell(output.sheets[0].cells, 1, 0)?.style?.numFmt).toBe("0.00")
    expect(output.properties?.title).toBe("Report")
  })

  it("keeps output encoding separate from the model", async () => {
    const input = { sheets: [{ name: "S", rows: [["text"]] }] }
    const { ZipReader } = await import("../src/zip")
    const zip = new ZipReader(await writeXlsx(input, { stringMode: "inline" }))
    expect(zip.has("xl/sharedStrings.xml")).toBe(false)
    expect(input).toEqual({ sheets: [{ name: "S", rows: [["text"]] }] })
  })

  it("reports authoring losses without mutating the read workbook", async () => {
    const input = workbook()
    const before = structuredClone(input)
    const drops: string[] = []
    await writeXlsx(input, { onDrop: (d) => drops.push(`${d.sheet}:${d.field}`) })
    expect(drops).toEqual(["Report:charts"])
    expect(input).toEqual(before)
  })

  it("accepts the same read workbook in the ODS writer", async () => {
    const input = workbook()
    const drops: string[] = []
    const output = await readOds(await writeOds(input, { onDrop: (d) => drops.push(d.field) }))
    expect(output.sheets[0].rows).toEqual(input.sheets[0].rows)
    expect(drops.sort()).toEqual(["charts", "freezePane"])
  })

  it("reports XLSX features the ODS authoring writer cannot carry", async () => {
    const drops: string[] = []
    await writeOds(
      {
        sheets: [
          {
            name: "S",
            rows: [[1]],
            images: [
              { data: new Uint8Array([1]), type: "png", anchor: { from: { row: 0, col: 0 } } },
            ],
            autoFilter: { range: "A1:A1" },
          },
        ],
        workbookProtection: { lockStructure: true },
      },
      { onDrop: (d) => drops.push(d.field) },
    )
    expect(drops.sort()).toEqual(["autoFilter", "images", "workbookProtection"])
  })
})
