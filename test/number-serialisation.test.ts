import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeCsv } from "../src/csv/writer"
import { parseCsv } from "../src/csv/reader"
import { ZipReader } from "../src/zip/reader"

// ═══════════════════════════════════════════════════════════════════════
// #474 asked whether numbers should keep being written with
// `String(value)` — so `1e21` becomes `1e+21` and `0.1 + 0.2` becomes
// `0.30000000000000004` (17 significant digits), where Excel writes
// `1E+21` and caps its *display* at 15.
//
// The decision is to keep it, and this file is why: `String(value)` is
// the only spelling that round-trips a double exactly. Truncating to 15
// significant digits would lose data that a caller put in the cell on
// purpose, to make the file look more like one Excel wrote. A library
// whose job is moving data faithfully should not make that trade.
//
// Both spellings are valid `xsd:double` — the lexical space is
// `[Ee](\+|-)?[0-9]+`, so lowercase `e` is conformant, not tolerated.
//
// These tests exist so the 15-digit "fix" cannot land quietly later.
// ═══════════════════════════════════════════════════════════════════════

/** The values where a naive round-trip goes wrong. */
const EXTREMES = [
  1e21, // where JS switches to exponent notation
  1e-7, // where it switches at the small end
  0.1 + 0.2, // 17 significant digits, and famously not 0.3
  1 / 3,
  1e300,
  5e-324, // Number.MIN_VALUE — one subnormal bit
  -1e21,
  Number.MAX_SAFE_INTEGER,
  Number.MIN_SAFE_INTEGER,
  Number.EPSILON,
]

const decoder = new TextDecoder("utf-8")

describe("XLSX numbers survive exactly", () => {
  it("every extreme comes back bit-identical", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [EXTREMES] }] })
    const back = (await readXlsx(bytes)).sheets[0]!.rows[0]!

    for (let i = 0; i < EXTREMES.length; i++) {
      // Object.is, not toBe: -0 and NaN would otherwise pass wrongly.
      expect(Object.is(back[i], EXTREMES[i]), `${EXTREMES[i]} came back as ${back[i]}`).toBe(true)
    }
  })

  it("writes the digits JS produces, exponent and all", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[1e21, 0.1 + 0.2]] }] })
    const xml = decoder.decode(await new ZipReader(bytes).extract("xl/worksheets/sheet1.xml"))

    // Both are conformant xsd:double. `1E+21` would be equally valid and
    // no more correct; `1000000000000000000000` and `0.3` would be
    // neither, the second because it is a different number.
    expect(xml).toContain("<v>1e+21</v>")
    expect(xml).toContain("<v>0.30000000000000004</v>")
  })

  it("does not round to 15 significant digits", async () => {
    // The change #474 raised. 0.1 + 0.2 at 15 digits is 0.3 — a value
    // the caller did not write.
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[0.1 + 0.2]] }] })

    expect((await readXlsx(bytes)).sheets[0]!.rows[0]![0]).not.toBe(0.3)
  })

  it("does not keep the sign of negative zero, and that is the one loss", async () => {
    // `String(-0)` is `"0"`, so the sign goes. Excel has no signed zero
    // either — writing `<v>-0</v>` would be hucre inventing a distinction
    // the format does not carry. Recorded here so it is a known answer
    // rather than a surprise.
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[-0]] }] })
    const back = (await readXlsx(bytes)).sheets[0]!.rows[0]!

    expect(back[0]).toBe(0)
    expect(Object.is(back[0], -0)).toBe(false)
  })
})

describe("CSV numbers survive exactly too", () => {
  it("every extreme comes back bit-identical", () => {
    const rows = parseCsv(writeCsv([EXTREMES]), { typeInference: true })

    for (let i = 0; i < EXTREMES.length; i++) {
      expect(
        Object.is(rows[0]![i], EXTREMES[i]),
        `${EXTREMES[i]} came back as ${rows[0]![i]}`,
      ).toBe(true)
    }
  })

  it("still prefers the plain form where it is the same number", () => {
    // The point of expanding at all: Excel shows a value written as
    // `1E-07` in exponent notation, and `0.0000001` is both prettier and
    // exactly equal.
    expect(writeCsv([[1e-7]]).trim()).toBe("0.0000001")
    expect(writeCsv([[1e16]]).trim()).toBe("10000000000000000")
  })

  it("stops expanding when the plain form is a different number", () => {
    // `toFixed(20)` caps at twenty decimal places. Number.MIN_VALUE came
    // out as "0.0" and Number.EPSILON kept five digits of seventeen.
    expect(writeCsv([[Number.MIN_VALUE]]).trim()).toBe("5e-324")
    expect(writeCsv([[Number.EPSILON]]).trim()).toBe("2.220446049250313e-16")
  })
})
