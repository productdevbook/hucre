import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import type { CellStyle } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// A number format whose *literal text* happened to contain h, s, d or y
// was read as a date or time format, and the digits were destroyed with
// it:
//
//   "Sales"#,##0   ->   "Sales""#","#""#""0"
//
// Every `#` and `0` came back as quoted literal text, so the number had
// no placeholder left to land in.
//
// The classifiers looked for those letters in the whole format code. In
// `#,##0" days"` the `d` and `y` are letters in a word, not tokens —
// `isPercentageFormat` right next to them already stripped quoted
// literals for exactly this reason, and the comment in `isDateFormat`
// says colour tags are stripped because `[White]` was mistaken for a time
// format by the `h` in "White". Quoted literals are the same problem and
// were the case nobody had.
//
// XLSX was never affected: it writes the format code verbatim.
// ═══════════════════════════════════════════════════════════════════════

async function odsNumFmt(numFmt: string): Promise<string | undefined> {
  const cells = new Map([["0,0", { value: 1234.5, style: { numFmt } as CellStyle }]])
  const bytes = await writeOds({ sheets: [{ name: "S", rows: [[1234.5]], cells }] })
  return (await readOds(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

async function xlsxNumFmt(numFmt: string): Promise<string | undefined> {
  const cells = new Map([["0,0", { value: 1234.5, style: { numFmt } as CellStyle }]])
  const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [[1234.5]], cells }] })
  return (await readXlsx(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

describe("letters inside a quoted literal are text, not tokens", () => {
  const CASES = [
    '"Sales"#,##0',
    '"Hours"0.0',
    '"Yes"0',
    '"Days"0',
    '"Month"0',
    '"Sum: "#,##0.00',
    '#,##0" hrs"',
    '#,##0" days"',
    '0" units"',
    '"Total"#,##0;"Loss"#,##0',
  ]

  it("round-trip through ODS unchanged", async () => {
    for (const numFmt of CASES) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("and XLSX carries them too, as it always did", async () => {
    for (const numFmt of CASES) {
      expect(await xlsxNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("a colour tag in front of one does not bring the problem back", async () => {
    // `[Red]` is dropped, which is what stopped `[White]` being read as a
    // time format. The literal after it still has to survive.
    expect(await odsNumFmt('[Red]"Sales"#,##0')).toBe('"Sales"#,##0')
  })
})

describe("real date and time formats still are ones", () => {
  it("the plain ones", async () => {
    for (const numFmt of ["yyyy-mm-dd", "dd/mm/yyyy", "hh:mm:ss", "hh:mm", "mmm-yy"]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })

  it("and one with a literal in the middle of it", async () => {
    // The tokens outside the quotes are what makes this a date, and they
    // still are. This is the case a fix that stripped too much breaks.
    const back = await odsNumFmt('dd" of "mmmm')

    expect(back).toContain("dd")
    expect(back).toContain("mmmm")
    expect(back).toContain(" of ")
  })

  it("an elapsed-time format is still elapsed", async () => {
    expect(await odsNumFmt("[hh]:mm")).toBe("[hh]:mm")
  })
})

describe("the plain formats around them are unchanged", () => {
  it("still round-trip", async () => {
    for (const numFmt of ["0.00", "#,##0", "#,##0.00", "0%", "0.00%", '"$"#,##0.00', "0.00E+00"]) {
      expect(await odsNumFmt(numFmt), numFmt).toBe(numFmt)
    }
  })
})
