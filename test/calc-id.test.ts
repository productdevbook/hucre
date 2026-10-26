import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"

// #621: Excel for the web ignores fullCalcOnLoad when calcId is 0, so formulas
// never compute. Every writer path must emit a non-zero calcId.
const CALC = /<calcPr calcId="191029" fullCalcOnLoad="1"\/>/

async function workbookXml(data: Uint8Array): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(data).extract("xl/workbook.xml"))
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

describe("calcPr calcId (#621)", () => {
  const sheets = [{ name: "S", rows: [[{ formula: "1+1" }]] }]

  it("writeXlsx writes a non-zero calcId", async () => {
    const xml = await workbookXml(await writeXlsx({ sheets }))
    expect(xml).toMatch(CALC)
    expect(xml).not.toContain('calcId="0"')
  })

  it("writeXlsxStream never writes calcId=0 (it emits no calcPr at all)", async () => {
    const data = await collect(writeXlsxStream([[{ formula: "1+1" }]], { name: "S" }))
    expect(await workbookXml(data)).not.toContain(`calcId="0"`)
  })

  it("saveXlsx after openXlsx writes a non-zero calcId", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets }))
    expect(await workbookXml(await saveXlsx(wb))).toMatch(CALC)
  })
})
