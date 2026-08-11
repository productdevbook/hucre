import { describe, expect, it } from "vitest"
import { fieldsOf, ownFieldsOf } from "./_reflect"

// ═══════════════════════════════════════════════════════════════════════
// #474 — four tests derive their expectations from `src/_types.ts` by
// regex, and a derive-from-source test that reads *less* than it should
// passes rather than fails. That is the dangerous direction, so the
// reader itself is worth testing.
// ═══════════════════════════════════════════════════════════════════════

describe("reading an interface's own fields", () => {
  it("finds the ones a hand-check confirms", () => {
    const fields = ownFieldsOf("MergeRange")

    expect(fields).toEqual(["startRow", "startCol", "endRow", "endCol"])
  })

  it("does not stop at a nested type's closing brace", () => {
    // `PageMargins` is small and flat; `PageSetup` holds `margins?:
    // PageMargins` plus a dozen fields after it, and several of its doc
    // comments carry braces. Truncation here would be silent.
    const fields = ownFieldsOf("PageSetup")

    expect(fields).toContain("paperSize") // first
    expect(fields).toContain("margins") // middle
    expect(fields).toContain("usePrinterDefaults") // last
  })

  it("does not collect a nested object type's fields as the parent's", () => {
    // `SheetImage.anchor` is an inline `{ from: ...; to?: ... }`. Its
    // `from` and `to` belong to that type, not to SheetImage.
    const fields = ownFieldsOf("SheetImage")

    expect(fields).toContain("anchor")
    expect(fields).not.toContain("from")
    expect(fields).not.toContain("to")
  })

  it("does not collect a commented-out field", () => {
    // Comments are stripped before matching, so a field someone parked
    // behind `//` is not counted as shipped.
    expect(ownFieldsOf("ReadOptions")).not.toContain("headerRow")
  })
})

describe("following extends", () => {
  it("includes an inherited field", () => {
    // The audit behind #439 filed a finding against `JsonReadOptions` for
    // a missing `typeInference` that was there all along, on the
    // interface it extends. A reader that cannot see `extends` is how
    // that mistake gets made twice.
    expect(ownFieldsOf("JsonReadOptions")).not.toContain("typeInference")
    expect(fieldsOf("JsonReadOptions")).toContain("typeInference")
  })

  it("keeps the interface's own fields as well", () => {
    const own = ownFieldsOf("JsonReadOptions")
    const all = fieldsOf("JsonReadOptions")

    expect(own.length).toBeGreaterThan(0)
    for (const field of own) expect(all, field).toContain(field)
    expect(all.length).toBeGreaterThan(own.length)
  })

  it("reports each field once even when a base repeats it", () => {
    const all = fieldsOf("JsonReadOptions")

    expect(new Set(all).size).toBe(all.length)
  })
})

describe("when the model moves under it", () => {
  it("says so rather than returning nothing", () => {
    expect(() => ownFieldsOf("NoSuchInterfaceAnywhere")).toThrow(/not found/)
  })
})
