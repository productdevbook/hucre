import { describe, expect, it, vi } from "vitest"
import { formatValue } from "../src/_format"
import { InvalidArgumentError } from "../src/errors"

function withArithmeticBudget<T>(fn: () => T): T {
  // A synchronous denominator scan cannot be interrupted by a test
  // timeout. Fail after bounded arithmetic instead of hanging on 16 '?'.
  let calls = 0
  const floor = Math.floor
  const round = Math.round
  const spend = () => {
    if (++calls > 256) throw new Error("Fraction formatting exceeded its arithmetic budget")
  }
  const floorSpy = vi.spyOn(Math, "floor").mockImplementation((value) => {
    spend()
    return floor(value)
  })
  const roundSpy = vi.spyOn(Math, "round").mockImplementation((value) => {
    spend()
    return round(value)
  })
  try {
    return fn()
  } finally {
    floorSpy.mockRestore()
    roundSpy.mockRestore()
  }
}

describe("formatValue resource limits", () => {
  it.each([9, 16, 250])("bounds work with %i denominator placeholders", (digits) => {
    expect(withArithmeticBudget(() => formatValue(1e-20, "# ?/" + "?".repeat(digits)))).toBe(
      "0      ",
    )
  })

  it.each([NaN, Infinity, -Infinity])("does not search for a non-finite value %s", (value) => {
    expect(withArithmeticBudget(() => formatValue(value, "# ?/?????????"))).toBe(String(value))
  })

  it("handles a subnormal without overflowing the reciprocal search", () => {
    expect(withArithmeticBudget(() => formatValue(Number.MIN_VALUE, "# ?/????????????????"))).toBe(
      "0      ",
    )
  })

  it("rejects a format longer than 255 characters before scanning it", () => {
    expect(() => formatValue(1.25, "0".repeat(256))).toThrow(InvalidArgumentError)
    expect(() => formatValue(1.25, "0".repeat(256))).toThrow(/255/)
  })

  it("admits the format-length boundary, including literals", () => {
    expect(formatValue(1, '"' + "x".repeat(253) + '"')).toBe("x".repeat(253))
  })

  it.each(["", "E+00", "E+?", "%"])("rejects excessive precision in a %s format", (suffix) => {
    expect(() => formatValue(1.25, "0." + "0".repeat(101) + suffix)).toThrow(InvalidArgumentError)
  })

  it.each(["", "E+00"])("accepts 100 decimal places in a %s format", (suffix) => {
    expect(formatValue(0.5, "0." + "0".repeat(100) + suffix)).toContain("5")
  })

  it("rejects fixed denominators that cannot be represented as safe integers", () => {
    expect(() => formatValue(0.5, "# ?/9007199254740992")).toThrow(InvalidArgumentError)
  })
})

describe("bounded fraction approximation", () => {
  it("prefers zero when it ties the closest nonzero fraction", () => {
    expect(formatValue(1 / 18, "# ?/?")).toBe("0      ")
  })

  it("keeps placeholder padding when the denominator range is capped", () => {
    expect(withArithmeticBudget(() => formatValue(0.5, "# ?/" + "0".repeat(16)))).toBe(
      "1/" + "2".padStart(16, " "),
    )
  })

  it("matches an exhaustive search for ordinary improper fractions", () => {
    // An independent small-denominator oracle protects rounding, ties,
    // and semiconvergents while the implementation avoids this scan.
    for (const digits of [1, 2, 3]) {
      const maxDen = 10 ** digits - 1
      for (let i = 1; i <= 200; i++) {
        const value = (i * Math.SQRT2) / 100
        let bestNum = 0
        let bestDen = 1
        let bestError = value
        for (let den = 1; den <= maxDen; den++) {
          const num = Math.round(value * den)
          const error = Math.abs(value - num / den)
          if (error < bestError) {
            bestNum = num
            bestDen = den
            bestError = error
          }
        }
        const expected =
          bestNum === 0 ? "0      " : `${bestNum}/${String(bestDen).padStart(digits, " ")}`
        expect(formatValue(value, "?/" + "?".repeat(digits))).toBe(expected)
      }
    }
  })
})
