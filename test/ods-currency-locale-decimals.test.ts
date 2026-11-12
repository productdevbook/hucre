import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import type { CellStyle } from "../src/_types"

// The `0` in a locale id like `[$€-407]` was counted as the number's
// only digit, so the ODS style had no decimals.

async function odsNumFmt(numFmt: string): Promise<string | undefined> {
  const cells = new Map([["0,0", { value: 1234.5, style: { numFmt } as CellStyle }]])
  const bytes = await writeOds({ sheets: [{ name: "S", rows: [[1234.5]], cells }] })
  return (await readOds(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

describe("ODS currency formats with a locale tag", () => {
  it("keeps the decimals when the tag comes first", async () => {
    expect(await odsNumFmt("[$€-407] #,##0.00")).toBe('"€"#,##0.00')
    expect(await odsNumFmt("[$£-809]#,##0.00")).toBe('"£"#,##0.00')
  })
})
