import { describe, expect, it } from "vitest"
import { fillTemplate } from "../src/template"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { readXlsx } from "../src/xlsx/reader"
import type { Cell, CellValue, Sheet, Workbook } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

/**
 * A workbook shaped the way the XLSX reader hands it back: `rows` carries
 * the plain values and `cells` carries the rich record for the same
 * coordinate. `fillTemplate` walks both, so both must be exercised.
 */
function workbookWithCells(entries: Array<[string, Cell]>, rows: CellValue[][] = [[]]): Workbook {
  const sheet: Sheet = { name: "Sheet1", rows, cells: new Map(entries) }
  return { sheets: [sheet] }
}

function cell(value: CellValue, type: Cell["type"] = "string"): Cell {
  return { value, type }
}

// ═══════════════════════════════════════════════════════════════════════
// The `cells` Map — the rich-cell half of the template engine.
// `rows` alone is what most callers see, but every workbook produced by
// `openXlsx({ readStyles: true })` also carries `cells`, and a placeholder
// that is only substituted in `rows` would be written back out unfilled.
// ═══════════════════════════════════════════════════════════════════════

describe("fillTemplate — cells Map substitution", () => {
  it("replaces a whole-cell placeholder in the cells Map", () => {
    const wb = workbookWithCells([["0,0", cell("{{company}}")]])
    fillTemplate(wb, { company: "Acme Corp" })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("Acme Corp")
  })

  it("leaves a whole-cell placeholder untouched when the key is absent", () => {
    const wb = workbookWithCells([["0,0", cell("{{missing}}")]])
    fillTemplate(wb, { present: "x" })
    const c = wb.sheets[0].cells!.get("0,0")!
    // Both value *and* type must survive: an unfilled template should be
    // re-savable and still look like the original template.
    expect(c.value).toBe("{{missing}}")
    expect(c.type).toBe("string")
  })

  it("skips cells whose value is not a string", () => {
    const wb = workbookWithCells([
      ["0,0", cell(42, "number")],
      ["0,1", cell(null, "empty")],
      ["0,2", cell(true, "boolean")],
    ])
    fillTemplate(wb, { anything: "x" })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe(42)
    expect(wb.sheets[0].cells!.get("0,1")!.value).toBe(null)
    expect(wb.sheets[0].cells!.get("0,2")!.value).toBe(true)
  })

  it("leaves a whole-cell placeholder in `rows` untouched when the key is absent", () => {
    // The rows path has its own single-placeholder shortcut; an unknown key
    // there must fall through to "leave as-is" and not become `undefined`.
    const wb: Workbook = { sheets: [{ name: "S", rows: [["{{missing}}", "{{present}}"]] }] }
    fillTemplate(wb, { present: "ok" })
    expect(wb.sheets[0].rows[0][0]).toBe("{{missing}}")
    expect(wb.sheets[0].rows[0][1]).toBe("ok")
  })

  it("skips string cells that contain no opening braces", () => {
    const wb = workbookWithCells([["0,0", cell("plain text }} with a stray close")]])
    fillTemplate(wb, { anything: "x" })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("plain text }} with a stray close")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Type synchronisation. A Cell carries both `value` and `type`; writing a
// number into `value` while leaving `type: "string"` would make the writer
// emit `<c t="s">` for a numeric value and corrupt the file.
// ═══════════════════════════════════════════════════════════════════════

describe("fillTemplate — cells Map keeps `type` in sync with `value`", () => {
  it("retypes to number for a numeric replacement", () => {
    const wb = workbookWithCells([["0,0", cell("{{total}}")]])
    fillTemplate(wb, { total: 12500 })
    const c = wb.sheets[0].cells!.get("0,0")!
    expect(c.value).toBe(12500)
    expect(c.type).toBe("number")
  })

  it("retypes to boolean for a boolean replacement", () => {
    const wb = workbookWithCells([["0,0", cell("{{active}}")]])
    fillTemplate(wb, { active: false })
    const c = wb.sheets[0].cells!.get("0,0")!
    expect(c.value).toBe(false)
    expect(c.type).toBe("boolean")
  })

  it("retypes to date for a Date replacement", () => {
    const due = new Date("2025-03-01T00:00:00Z")
    const wb = workbookWithCells([["0,0", cell("{{due}}", "number")]])
    fillTemplate(wb, { due })
    const c = wb.sheets[0].cells!.get("0,0")!
    expect(c.value).toBeInstanceOf(Date)
    expect(c.type).toBe("date")
  })

  it("retypes back to string when the replacement is a string", () => {
    const wb = workbookWithCells([["0,0", cell("{{name}}", "number")]])
    fillTemplate(wb, { name: "Acme" })
    const c = wb.sheets[0].cells!.get("0,0")!
    expect(c.value).toBe("Acme")
    expect(c.type).toBe("string")
  })

  it("stores a null replacement as a null value", () => {
    // The `type` for a null replacement falls through to "string"; the
    // writer keys off the null value, so the cell still serialises blank.
    const wb = workbookWithCells([["0,0", cell("{{blank}}")]])
    fillTemplate(wb, { blank: null })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe(null)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Mixed text in the cells Map goes through the same stringification rules
// as `rows`: null collapses to "", Date renders ISO-8601, everything else
// goes through String().
// ═══════════════════════════════════════════════════════════════════════

describe("fillTemplate — cells Map mixed text", () => {
  it("stringifies numbers embedded in surrounding text", () => {
    const wb = workbookWithCells([["0,0", cell("Total: {{amount}} USD")]])
    fillTemplate(wb, { amount: 500 })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("Total: 500 USD")
  })

  it("renders an embedded Date as an ISO-8601 string", () => {
    const wb = workbookWithCells([["0,0", cell("Due {{due}}.")]])
    fillTemplate(wb, { due: new Date("2025-03-01T00:00:00Z") })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("Due 2025-03-01T00:00:00.000Z.")
  })

  it("renders an embedded null as an empty string", () => {
    const wb = workbookWithCells([["0,0", cell("[{{nothing}}]")]])
    fillTemplate(wb, { nothing: null })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("[]")
  })

  it("keeps unmatched placeholders while substituting matched ones", () => {
    const wb = workbookWithCells([["0,0", cell("{{known}} / {{unknown}}")]])
    fillTemplate(wb, { known: "yes" })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("yes / {{unknown}}")
  })

  it("substitutes every occurrence of a repeated placeholder", () => {
    // PLACEHOLDER_RE is a module-level /g regex reused across calls — a
    // leaked lastIndex would make the second occurrence (or the second
    // call) silently skip. Two cells in one pass catch that.
    const wb = workbookWithCells([
      ["0,0", cell("{{x}}-{{x}}-{{x}}")],
      ["0,1", cell("{{x}}!")],
    ])
    fillTemplate(wb, { x: "A" })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("A-A-A")
    expect(wb.sheets[0].cells!.get("0,1")!.value).toBe("A!")
  })

  it("does not leak regex state between successive fillTemplate calls", () => {
    const first = workbookWithCells([["0,0", cell("{{a}} {{a}}")]])
    const second = workbookWithCells([["0,0", cell("{{a}} {{a}}")]])
    fillTemplate(first, { a: "1" })
    fillTemplate(second, { a: "2" })
    expect(second.sheets[0].cells!.get("0,0")!.value).toBe("2 2")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Structural edge cases — a template workbook is user-supplied data, so
// unusual shapes must not throw.
// ═══════════════════════════════════════════════════════════════════════

describe("fillTemplate — structural edge cases", () => {
  it("handles a workbook with no sheets", () => {
    const wb: Workbook = { sheets: [] }
    expect(fillTemplate(wb, { a: 1 })).toBe(wb)
  })

  it("handles a sheet with no rows and no cells Map", () => {
    const wb: Workbook = { sheets: [{ name: "Empty", rows: [] }] }
    expect(() => fillTemplate(wb, { a: 1 })).not.toThrow()
  })

  it("handles an empty cells Map", () => {
    const wb = workbookWithCells([])
    expect(() => fillTemplate(wb, { a: 1 })).not.toThrow()
  })

  it("skips holes in a sparse row without throwing", () => {
    const sparse: CellValue[] = ["{{a}}"]
    sparse[3] = "{{a}}"
    const wb: Workbook = { sheets: [{ name: "S", rows: [sparse] }] }
    fillTemplate(wb, { a: "filled" })
    expect(wb.sheets[0].rows[0][0]).toBe("filled")
    expect(wb.sheets[0].rows[0][3]).toBe("filled")
    expect(wb.sheets[0].rows[0][1]).toBeUndefined()
  })

  it("mutates and returns the same workbook instance", () => {
    const wb = workbookWithCells([["0,0", cell("{{a}}")]], [["{{a}}"]])
    expect(fillTemplate(wb, { a: "x" })).toBe(wb)
  })

  it("fills the cells Map on every sheet, not just the first", () => {
    const wb: Workbook = {
      sheets: [
        { name: "One", rows: [[]], cells: new Map([["0,0", cell("{{a}}")]]) },
        { name: "Two", rows: [[]], cells: new Map([["0,0", cell("Hi {{a}}")]]) },
      ],
    }
    fillTemplate(wb, { a: "there" })
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("there")
    expect(wb.sheets[1].cells!.get("0,0")!.value).toBe("Hi there")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Prototype-chain keys.
//
// KNOWN BUG — src/template.ts:48, :57, :79, :91 use `key in data`, which
// walks Object.prototype. A template containing `{{toString}}` therefore
// resolves to `Object.prototype.toString` and the *function object* is
// written into the cell (in the cells-Map path `type` is even set to
// "string" alongside it). See the report accompanying these tests.
// Fix: `Object.hasOwn(data, key)`.
// ═══════════════════════════════════════════════════════════════════════

describe("fillTemplate — inherited Object.prototype keys", () => {
  it.skip("leaves {{toString}} alone when `data` has no own `toString`", () => {
    const wb = workbookWithCells([["0,0", cell("{{toString}}")]], [["{{toString}}"]])
    fillTemplate(wb, { name: "Acme" })
    expect(wb.sheets[0].rows[0][0]).toBe("{{toString}}")
    expect(wb.sheets[0].cells!.get("0,0")!.value).toBe("{{toString}}")
  })

  it.skip("leaves an inherited key alone inside mixed text", () => {
    const wb: Workbook = { sheets: [{ name: "S", rows: [["Hi {{constructor}}"]] }] }
    fillTemplate(wb, { name: "Acme" })
    expect(wb.sheets[0].rows[0][0]).toBe("Hi {{constructor}}")
  })

  it.skip("leaves {{__proto__}} alone rather than injecting Object.prototype", () => {
    const wb: Workbook = { sheets: [{ name: "S", rows: [["{{__proto__}}"]] }] }
    fillTemplate(wb, {})
    expect(wb.sheets[0].rows[0][0]).toBe("{{__proto__}}")
  })

  it("honours an own property that shadows a prototype key", () => {
    const wb: Workbook = { sheets: [{ name: "S", rows: [["{{toString}}"]] }] }
    fillTemplate(wb, { toString: "explicitly provided" })
    expect(wb.sheets[0].rows[0][0]).toBe("explicitly provided")
  })

  it("fills from a null-prototype data object", () => {
    // The documented escape hatch for the bug above, and the shape a caller
    // gets from `JSON.parse` reviver / `Object.create(null)` pipelines.
    const data = Object.create(null) as Record<string, CellValue>
    data["name"] = "Acme"
    const wb: Workbook = { sheets: [{ name: "S", rows: [["{{name}} / {{toString}}"]] }] }
    fillTemplate(wb, data)
    expect(wb.sheets[0].rows[0][0]).toBe("Acme / {{toString}}")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// End-to-end: a real XLSX template opened with `openXlsx`, filled, saved,
// and re-read. This is the documented usage in the fillTemplate JSDoc and
// the only path that exercises `rows` and `cells` on the same workbook.
// ═══════════════════════════════════════════════════════════════════════

describe("fillTemplate — XLSX round-trip", () => {
  it("fills a styled template and survives openXlsx → fill → saveXlsx", async () => {
    const templateBytes = await writeXlsx({
      sheets: [
        {
          name: "Invoice",
          rows: [
            ["Customer", "{{customer}}"],
            ["Total", "{{total}}"],
            ["Note", "Due {{due}} — thanks {{customer}}"],
          ],
          // Styling the placeholder cells is what makes the reader emit a
          // `cells` Map for them, so this covers both substitution paths.
          cells: new Map([
            ["0,1", { style: { font: { bold: true } } }],
            ["1,1", { style: { font: { italic: true } } }],
          ]),
        },
      ],
    })

    const wb = await openXlsx(templateBytes, { readStyles: true })
    expect(wb.sheets[0].cells).toBeDefined()
    expect(wb.sheets[0].cells!.get("0,1")!.value).toBe("{{customer}}")

    fillTemplate(wb, {
      customer: "Acme Corp",
      total: 1250.5,
      due: new Date("2025-03-01T00:00:00Z"),
    })

    // The cells Map entry is substituted *and* retyped, not just `rows`.
    const filledCell = wb.sheets[0].cells!.get("1,1")!
    expect(filledCell.value).toBe(1250.5)
    expect(filledCell.type).toBe("number")
    expect(filledCell.style?.font?.italic).toBe(true)

    const out = await saveXlsx(wb)
    const reread = await readXlsx(out, { readStyles: true })
    expect(reread.sheets[0].rows[0]).toEqual(["Customer", "Acme Corp"])
    expect(reread.sheets[0].rows[1]).toEqual(["Total", 1250.5])
    expect(reread.sheets[0].rows[2][1]).toBe("Due 2025-03-01T00:00:00.000Z — thanks Acme Corp")
    // Styling from the template must not be lost by the fill.
    expect(reread.sheets[0].cells!.get("0,1")!.style?.font?.bold).toBe(true)
  })
})
