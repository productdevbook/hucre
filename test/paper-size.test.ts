import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { PAPER_SIZE_MAP } from "../src/xlsx/worksheet-writer"
import type { PaperSize } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §Q — `PaperSize` was a closed union of nine names, and the reader
// dropped anything else outright:
//
//   const name = PAPER_SIZE_REVERSE[num]
//   if (name) ps.paperSize = name
//
// So a workbook set to A6, or to any envelope, lost its page size on read
// with no error and nothing in the parity statement — and a caller who
// knew the OOXML code could not pass it either.
// ═══════════════════════════════════════════════════════════════════════

async function roundTrip(paperSize: PaperSize): Promise<PaperSize | undefined> {
  const bytes = await writeXlsx({
    sheets: [{ name: "S", rows: [[1]], pageSetup: { paperSize } }],
  })
  return (await readXlsx(bytes)).sheets[0]!.pageSetup?.paperSize
}

describe("named paper sizes", () => {
  it("round-trips every name in the table", async () => {
    for (const name of Object.keys(PAPER_SIZE_MAP) as Array<keyof typeof PAPER_SIZE_MAP>) {
      expect(await roundTrip(name), name).toBe(name)
    }
  })

  it("still round-trips the nine that were there before", async () => {
    for (const name of ["letter", "legal", "a3", "a4", "a5", "b4", "b5", "executive", "tabloid"]) {
      expect(await roundTrip(name as PaperSize), name).toBe(name)
    }
  })

  it("covers the sizes people actually hit", async () => {
    expect(await roundTrip("a6")).toBe("a6")
    expect(await roundTrip("envelopeDL")).toBe("envelopeDL")
    expect(await roundTrip("japanesePostcard")).toBe("japanesePostcard")
    expect(await roundTrip("ledger")).toBe("ledger")
  })
})

describe("a raw OOXML code is the escape hatch", () => {
  it("writes and reads back a code with no name", async () => {
    // 256 is in the printer-defined range, which Excel leaves to the driver.
    expect(await roundTrip(256)).toBe(256)
  })

  it("normalises a code that does have a name", async () => {
    // 9 is A4. Round-tripping it as the name is more useful than as 9, and
    // both mean the same thing to Excel.
    expect(await roundTrip(9)).toBe("a4")
  })

  it("ignores a code that is not usable", async () => {
    expect(await roundTrip(0)).toBeUndefined()
    expect(await roundTrip(-1)).toBeUndefined()
    expect(await roundTrip(1.5)).toBeUndefined()
  })
})

describe("the name↔code table has one home", () => {
  it("maps each name to a distinct code", () => {
    const codes = Object.values(PAPER_SIZE_MAP)

    expect(new Set(codes).size).toBe(codes.length)
  })

  it("uses the codes ECMA-376 assigns", () => {
    expect(PAPER_SIZE_MAP.letter).toBe(1)
    expect(PAPER_SIZE_MAP.a4).toBe(9)
    expect(PAPER_SIZE_MAP.a6).toBe(70)
    expect(PAPER_SIZE_MAP.envelopeDL).toBe(27)
  })
})
