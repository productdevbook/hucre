import { describe, expect, it } from "vitest"
import { parseJson, parseNdjson, writeJson, writeNdjson, workbookToJson } from "../src/json"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"

describe("writeJson", () => {
  it("emits a JSON array of objects", () => {
    const out = writeJson([
      { a: 1, b: 2 },
      { a: 3, b: 4 },
    ])
    expect(JSON.parse(out)).toEqual([
      { a: 1, b: 2 },
      { a: 3, b: 4 },
    ])
  })

  it("pretty-prints with 2-space indent", () => {
    const out = writeJson([{ a: 1 }], { pretty: true })
    expect(out).toContain("\n  ")
  })

  it("converts Date cells to ISO strings", () => {
    const d = new Date("2025-01-15T10:00:00Z")
    const out = writeJson([{ at: d }])
    expect(JSON.parse(out)).toEqual([{ at: d.toISOString() }])
  })

  it("round-trips via parseJson", () => {
    const data = [
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 },
    ]
    const out = writeJson(data)
    const back = parseJson(out)
    expect(back.data).toEqual(data)
  })
})

describe("writeNdjson", () => {
  it("emits one JSON object per line", () => {
    const out = writeNdjson([{ a: 1 }, { a: 2 }])
    expect(out).toBe('{"a":1}\n{"a":2}\n')
  })

  it("returns empty string for empty input", () => {
    expect(writeNdjson([])).toBe("")
  })

  it("round-trips via parseNdjson", () => {
    const data = [
      { a: 1, b: "x" },
      { a: 2, b: "y" },
    ]
    const back = parseNdjson(writeNdjson(data))
    expect(back.data).toEqual(data)
  })
})

describe("workbookToJson", () => {
  it("emits a single-sheet workbook as a flat array", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Name", "Age"],
            ["Alice", 30],
            ["Bob", 25],
          ],
        },
      ],
    })
    const wb = await readXlsx(buf)
    const json = workbookToJson(wb)
    expect(JSON.parse(json)).toEqual([
      { Name: "Alice", Age: 30 },
      { Name: "Bob", Age: 25 },
    ])
  })

  it("emits multi-sheet workbook as { sheetName: [...] }", async () => {
    const buf = await writeXlsx({
      sheets: [
        { name: "Users", rows: [["id"], [1], [2]] },
        { name: "Items", rows: [["sku"], ["A"], ["B"]] },
      ],
    })
    const wb = await readXlsx(buf)
    const obj = JSON.parse(workbookToJson(wb)) as Record<string, unknown[]>
    expect(Object.keys(obj)).toEqual(["Users", "Items"])
    expect(obj.Users).toEqual([{ id: 1 }, { id: 2 }])
    expect(obj.Items).toEqual([{ sku: "A" }, { sku: "B" }])
  })

  it("selects a specific sheet by name", async () => {
    const buf = await writeXlsx({
      sheets: [
        { name: "A", rows: [["x"], [1]] },
        { name: "B", rows: [["y"], [2]] },
      ],
    })
    const wb = await readXlsx(buf)
    const json = workbookToJson(wb, { sheet: "B" })
    expect(JSON.parse(json)).toEqual([{ y: 2 }])
  })
})
