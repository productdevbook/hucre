import { describe, expect, it, vi } from "vitest"
import { dateToSerial, serialToDate } from "../src/_date"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import type { ConditionalRuleType } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §C2, §C3, §C4 and §L — three value-level losses and one wrong
// count, all of which were decisions written into a source comment and
// never into anything a user reads.
//
// These pin the documented behaviour, so docs/PARITY.md and the README
// cannot drift away from what the code does.
// ═══════════════════════════════════════════════════════════════════════

describe("dates convert as instants, in UTC", () => {
  it("gives the same answer whatever the machine's timezone", () => {
    const noon = new Date(Date.UTC(2024, 0, 15, 12, 0, 0))

    vi.stubEnv("TZ", "America/New_York")
    const inNewYork = dateToSerial(noon)
    vi.stubEnv("TZ", "Asia/Tokyo")
    const inTokyo = dateToSerial(noon)
    vi.unstubAllEnvs()

    expect(inNewYork).toBe(inTokyo)
  })

  it("converts the instant, not the wall-clock day", () => {
    // The README and PARITY.md both say this in as many words, because it
    // is the one thing that surprises people.
    expect(dateToSerial(new Date(Date.UTC(2024, 0, 15)))).toBe(45306)
    expect(dateToSerial(new Date("2024-01-15"))).toBe(45306)
  })

  it("round-trips a UTC-built date exactly", () => {
    const d = new Date(Date.UTC(2024, 0, 15, 9, 30, 0))

    expect(serialToDate(dateToSerial(d)).toISOString()).toBe(d.toISOString())
  })
})

describe("the Lotus phantom day is a fixed point, and says so", () => {
  it("maps serial 60 onto the same instant as 59", () => {
    // 29 February 1900 does not exist. There is no instant to give it, so
    // it collapses onto 28 February — see docs/PARITY.md.
    expect(serialToDate(60).toISOString()).toBe(serialToDate(59).toISOString())
  })

  it("sends it back as 59, which is the documented loss", () => {
    expect(dateToSerial(serialToDate(60))).toBe(59)
  })

  it("leaves every serial either side of it exact", () => {
    for (const serial of [1, 58, 59, 61, 62, 1000, 45306]) {
      expect(dateToSerial(serialToDate(serial)), String(serial)).toBe(serial)
    }
  })
})

describe("a literal _xHHHH_ in cell text is the documented ambiguity", () => {
  it("reads back as the character it encodes", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["_x0041_"]] }] })

    // Excel would escape the leading underscore as _x005F_; hucre does not,
    // because doing so would mangle ordinary text containing underscores.
    // Accepted, and now written into docs/PARITY.md.
    expect((await readXlsx(bytes)).sheets[0]!.rows[0]![0]).toBe("A")
  })

  it("leaves ordinary underscored text alone", async () => {
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [["snake_case_name", "_x", "x_0041_"]] }],
    })

    expect((await readXlsx(bytes)).sheets[0]!.rows[0]).toEqual(["snake_case_name", "_x", "x_0041_"])
  })
})

describe("the conditional rule count the docs quote", () => {
  it("is 15, which is what the type has", () => {
    // A Record over the union, so a new member fails `tsc` here rather
    // than quietly making the number in the docs wrong again.
    const every: Record<ConditionalRuleType, true> = {
      cellIs: true,
      expression: true,
      colorScale: true,
      dataBar: true,
      iconSet: true,
      top10: true,
      aboveAverage: true,
      duplicateValues: true,
      uniqueValues: true,
      containsText: true,
      notContainsText: true,
      beginsWith: true,
      endsWith: true,
      containsBlanks: true,
      notContainsBlanks: true,
    }

    const types: ConditionalRuleType[] = [
      "cellIs",
      "expression",
      "colorScale",
      "dataBar",
      "iconSet",
      "top10",
      "aboveAverage",
      "duplicateValues",
      "uniqueValues",
      "containsText",
      "notContainsText",
      "beginsWith",
      "endsWith",
      "containsBlanks",
      "notContainsBlanks",
    ]

    // The README said 13 while PARITY.md said 15. The type is the arbiter.
    expect(Object.keys(every)).toHaveLength(15)
    expect(types.sort()).toEqual(Object.keys(every).sort())
  })
})
