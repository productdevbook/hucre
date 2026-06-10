import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readOds } from "../src/ods/reader"

const enc = new TextEncoder()

/** Build a minimal ODS containing one sheet whose single cell holds the
 *  given raw <table:table-cell> inner XML. */
async function odsWithCell(cellInnerXml: string): Promise<Uint8Array> {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body><office:spreadsheet>
    <table:table table:name="S">
      <table:table-row>
        <table:table-cell office:value-type="string">${cellInnerXml}</table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>`
  const zip = new ZipWriter()
  zip.add("mimetype", enc.encode("application/vnd.oasis.opendocument.spreadsheet"), {
    compress: false,
  })
  zip.add("content.xml", enc.encode(content))
  return zip.build()
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
    const cell = wb.sheets[0].cells?.get("0,0")
    expect(wb.sheets[0].rows[0][0]).toBe("before link after")
    expect(cell?.hyperlink?.target).toBe("https://example.com")
    expect(cell?.hyperlink?.display).toBe("link")
  })
})
