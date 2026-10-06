import { getCell } from "../src/cell-store"
import { describe, expect, it } from "vitest"
import { odsFromContent } from "./support/ods"
import { readOds } from "../src/ods/reader"

/** One raw cell; shared scaffolding keeps the relevant markup visible. */
function odsWithCell(cellInnerXml: string): Promise<Uint8Array> {
  return odsFromContent(`<table:table table:name="S"><table:table-row>
    <table:table-cell office:value-type="string">${cellInnerXml}</table:table-cell>
    </table:table-row></table:table>`)
}

describe("ODS reader — multi-paragraph and surrounded hyperlinks", () => {
  it("joins multiple <text:p> paragraphs with a newline", async () => {
    const data = await odsWithCell("<text:p>line1</text:p><text:p>line2</text:p>")
    const wb = await readOds(data)
    expect(wb.sheets[0].rows[0][0]).toBe("line1\nline2")
  })

  it("keeps text surrounding a hyperlink", async () => {
    const data = await odsWithCell(
      '<text:p>before <text:a xlink:href="https://example.com">link</text:a> after</text:p>',
    )
    const wb = await readOds(data)
    const cell = getCell(wb.sheets[0].cells, 0, 0)
    expect(wb.sheets[0].rows[0][0]).toBe("before link after")
    expect(cell?.hyperlink?.target).toBe("https://example.com")
    expect(cell?.hyperlink?.display).toBe("link")
  })
})
