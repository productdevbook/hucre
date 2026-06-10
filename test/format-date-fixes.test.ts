import { describe, expect, it } from "vitest"
import { formatValue } from "../src/_format"

describe("formatValue: minutes vs months", () => {
  it("renders mm:ss as minutes:seconds, not month:seconds", () => {
    // 0.5 of a day = 12:00:00 → 00 minutes, 00 seconds
    expect(formatValue(0.5, "mm:ss")).toBe("00:00")
  })

  it("renders m before s as minutes", () => {
    // 0.5 day at 12:00 → minutes = 0
    expect(formatValue(0.5, "m:ss")).toBe("0:00")
  })

  it("still renders standalone month with day token", () => {
    // serial 60 ≈ 1900-02-28 in the 1900 system; mm/dd is a date
    expect(formatValue(60, "mm/dd")).toMatch(/^\d{2}\/\d{2}$/)
  })
})

describe("formatValue: elapsed time", () => {
  it("renders [h]:mm with total hours exceeding 24", () => {
    // 1.5 days = 36 hours, 0 minutes
    expect(formatValue(1.5, "[h]:mm")).toBe("36:00")
  })

  it("renders [m] as total minutes", () => {
    // 1.5 days = 2160 minutes
    expect(formatValue(1.5, "[m]")).toBe("2160")
  })

  it("renders [s] as total seconds", () => {
    // 0.5 day = 43200 seconds
    expect(formatValue(0.5, "[s]")).toBe("43200")
  })
})

describe("formatValue: 1904 date system", () => {
  it("formats serial 0 as 1904-01-01 under the 1904 system", () => {
    expect(formatValue(0, "yyyy-mm-dd", { is1904: true })).toBe("1904-01-01")
  })

  it("differs from the 1900 system for the same serial", () => {
    const s1900 = formatValue(1, "yyyy-mm-dd")
    const s1904 = formatValue(1, "yyyy-mm-dd", { is1904: true })
    expect(s1900).not.toBe(s1904)
  })
})
