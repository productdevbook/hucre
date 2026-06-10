import { describe, expect, it } from "vitest"
import { a1ToR1C1 } from "../src/index"
import { writeOds } from "../src/ods/writer"
import { ZipReader } from "../src/zip/reader"

describe("a1ToR1C1 — does not corrupt function names or string literals", () => {
  it("rewrites a bare cell reference", () => {
    expect(a1ToR1C1("A1")).toBe("R1C1")
    expect(a1ToR1C1("$C$2")).toBe("R2C3")
  })

  it("leaves a function name like LOG10 untouched", () => {
    // "G10" inside "LOG10" must not be treated as a cell ref, and the
    // argument A1 must still be converted.
    expect(a1ToR1C1("LOG10(A1)")).toBe("LOG10(R1C1)")
  })

  it("leaves a quoted string literal untouched", () => {
    expect(a1ToR1C1('IF(A1>1,"AB1",0)')).toBe('IF(R1C1>1,"AB1",0)')
  })

  it("does not treat ATAN2( as the column AT", () => {
    expect(a1ToR1C1("ATAN2(B2,C3)")).toBe("ATAN2(R2C2,R3C3)")
  })
})

describe("ODS formula conversion — function names and literals preserved", () => {
  async function odsContent(formula: string): Promise<string> {
    const cells = new Map<string, { value: number; formula: string }>()
    cells.set("0,0", { value: 0, formula })
    const data = await writeOds({
      sheets: [{ name: "S", rows: [[0]], cells: cells as never }],
    })
    const raw = await new ZipReader(data).extract("content.xml")
    return new TextDecoder().decode(raw)
  }

  it("converts cell refs but not LOG10 or quoted literals", async () => {
    const xml = await odsContent('LOG10(A1)+IF(A1>1,"AB1",0)')
    // Refs A1 become [.A1]; LOG10 and the "AB1" literal stay intact.
    expect(xml).toContain("[.A1]")
    expect(xml).toContain("LOG10(")
    expect(xml).toContain("&quot;AB1&quot;")
    // The bug produced [.LOG10] and [.AB1] — assert those never appear.
    expect(xml).not.toContain("[.LOG10]")
    expect(xml).not.toContain("[.AB1]")
  })

  it("handles ranges", async () => {
    const xml = await odsContent("SUM(A1:B2)")
    expect(xml).toContain("[.A1:.B2]")
    expect(xml).not.toContain("[.SUM]")
  })
})
