import { describe, expect, it } from "vitest"
import { validateWithSchema } from "../src/_schema"
import { toHtml } from "../src/export/html"
import { toMarkdown } from "../src/export/markdown"
import type { CellValue, SchemaDefinition, Sheet } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const sheetOf = (rows: CellValue[][]): Sheet => ({ name: "S", rows })

const namedSchema: SchemaDefinition = {
  name: { column: "Name", type: "string" },
  price: { column: "Price", type: "number" },
}

// ═══════════════════════════════════════════════════════════════════════
// #365 — `headerRow` meant four different things across the public
// options types: 1-based here, 0-based on the *Objects readers and the
// JSON exporters, and a boolean on the HTML and Markdown exporters. An
// option name that means two things is an off-by-one that only shows up
// in someone's data.
// ═══════════════════════════════════════════════════════════════════════

describe("validateWithSchema headerRow is 0-based", () => {
  const rows: CellValue[][] = [["This is a title row"], ["Name", "Price"], ["Widget", 9.99]]

  it("treats headerRow: 1 as the second row", () => {
    const { data, errors } = validateWithSchema(rows, namedSchema, { headerRow: 1 })
    expect(errors).toHaveLength(0)
    expect(data).toEqual([{ name: "Widget", price: 9.99 }])
  })

  it("defaults to the first row", () => {
    const withHeaderFirst: CellValue[][] = [
      ["Name", "Price"],
      ["Widget", 9.99],
    ]
    // The old default was 1 under 1-based numbering and the new default
    // is 0 under 0-based — both mean the first row, so callers relying
    // on the default are unaffected.
    const { data } = validateWithSchema(withHeaderFirst, namedSchema)
    expect(data).toEqual([{ name: "Widget", price: 9.99 }])
  })

  it("uses -1 for rows that carry no header at all", () => {
    // `0` used to be the "no header row" sentinel, because 1-based
    // numbering had no other way to say it. The two concepts are
    // separate now.
    const byIndex: SchemaDefinition = { val: { columnIndex: 0, type: "string" } }
    const { data } = validateWithSchema([[42], [7]], byIndex, { headerRow: -1 })
    expect(data).toEqual([{ val: "42" }, { val: "7" }])
  })

  it("consumes the first row as a header when headerRow is 0", () => {
    const byIndex: SchemaDefinition = { val: { columnIndex: 0, type: "string" } }
    const { data } = validateWithSchema([["Header"], [42]], byIndex, { headerRow: 0 })
    expect(data).toEqual([{ val: "42" }])
  })
})

describe("hasHeaderRow on the exporters", () => {
  const rows: CellValue[][] = [
    ["Name", "Qty"],
    ["Widget", 2],
  ]

  it("toHtml emits a thead when asked", () => {
    expect(toHtml(sheetOf(rows), { hasHeaderRow: true })).toContain("<thead>")
    expect(toHtml(sheetOf(rows), { hasHeaderRow: false })).not.toContain("<thead>")
  })

  it("toMarkdown uses the first row as the header when asked", () => {
    // Markdown tables always carry a separator row, so the flag shows up
    // in *what* the header is: the real first row, or a generated 1/2/3.
    expect(toMarkdown(sheetOf(rows), { hasHeaderRow: true })).toContain("| Name")
    expect(toMarkdown(sheetOf(rows), { hasHeaderRow: false })).toContain("| 1 ")
  })

  it("keeps its own default when neither is given", () => {
    // HTML defaults to false, Markdown to true — unchanged.
    expect(toHtml(sheetOf(rows))).not.toContain("<thead>")
    expect(toMarkdown(sheetOf(rows))).toContain("| Name")
  })
})
