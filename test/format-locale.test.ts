import { describe, expect, it } from "vitest"
import { formatValue } from "../src/_format"
import { parseJson, writeJson } from "../src/json"
import { InvalidArgumentError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// #439 §R — `formatValue`'s locale was a four-entry table, and every
// other tag resolved to `undefined`, which the formatter read as "use the
// defaults". So a caller who explicitly asked for es-ES got an en-US
// rendering, silently.
//
// `Intl.NumberFormat` knows every tag and separators are all the
// formatter needs, so the table is gone.
// ═══════════════════════════════════════════════════════════════════════

describe("locale separators come from Intl", () => {
  it("still formats the four that were hard-coded", () => {
    expect(formatValue(1234567.5, "#,##0.00", { locale: "en-US" })).toBe("1,234,567.50")
    expect(formatValue(1234567.5, "#,##0.00", { locale: "de-DE" })).toBe("1.234.567,50")
    expect(formatValue(1234567.5, "#,##0.00", { locale: "tr-TR" })).toBe("1.234.567,50")
    // fr-FR groups with a space of some kind — the old table said U+00A0,
    // CLDR says U+202F. Asserting the shape rather than the codepoint
    // keeps this from depending on the runtime's ICU version.
    expect(formatValue(1234567.5, "#,##0.00", { locale: "fr-FR" })).toMatch(/^1\s234\s567,50$/u)
  })

  it("formats the ones it used to answer wrongly", () => {
    // Each of these returned "1,234,567.50" before — an en-US rendering
    // for a locale the caller named.
    expect(formatValue(1234567.5, "#,##0.00", { locale: "es-ES" })).toBe("1.234.567,50")
    expect(formatValue(1234567.5, "#,##0.00", { locale: "it-IT" })).toBe("1.234.567,50")
    expect(formatValue(1234567.5, "#,##0.00", { locale: "pt-BR" })).toBe("1.234.567,50")
    expect(formatValue(1234567.5, "#,##0.00", { locale: "nb-NO" })).not.toBe("1,234,567.50")
  })

  it("keeps the default separators when no locale is given", () => {
    expect(formatValue(1234567.5, "#,##0.00")).toBe("1,234,567.50")
  })

  it("localises the decimal separator in scientific notation too", () => {
    expect(formatValue(1234.5, "0.00E+00", { locale: "de-DE" })).toBe("1,23E+03")
  })

  it("refuses a tag Intl cannot use, rather than answering wrongly", () => {
    expect(() => formatValue(1, "#,##0", { locale: "not a tag" })).toThrow(InvalidArgumentError)
    expect(() => formatValue(1, "#,##0", { locale: "en_US" })).toThrow(InvalidArgumentError)
  })

  it("caches, so repeated formatting does not rebuild the formatter", () => {
    // Behavioural proxy for the cache: the answer stays stable and cheap.
    const first = formatValue(1234.5, "#,##0.00", { locale: "de-DE" })
    for (let i = 0; i < 1000; i++) formatValue(i, "#,##0.00", { locale: "de-DE" })

    expect(formatValue(1234.5, "#,##0.00", { locale: "de-DE" })).toBe(first)
  })
})

describe("JSON dates round-trip under typeInference", () => {
  // Not a change — `JsonReadOptions extends FlattenOptions`, which has
  // carried this option all along. docs/PARITY.md claimed ISO dates
  // round-tripped without saying the option was needed, and these pin
  // what it now says.
  const WHEN = new Date(Date.UTC(2024, 0, 15))

  it("comes back as a string by default", () => {
    const back = parseJson(writeJson([{ when: WHEN }]))

    expect(back.data[0]!.when).toBe("2024-01-15T00:00:00.000Z")
  })

  it("comes back as a Date when asked", () => {
    const back = parseJson(writeJson([{ when: WHEN }]), { typeInference: true })

    expect(back.data[0]!.when).toBeInstanceOf(Date)
    expect((back.data[0]!.when as Date).toISOString()).toBe(WHEN.toISOString())
  })

  it("leaves numbers and booleans alone, which JSON already carries", () => {
    const back = parseJson(writeJson([{ n: 7, b: true, s: "007" }]), { typeInference: true })

    expect(back.data[0]).toEqual({ n: 7, b: true, s: "007" })
  })
})
