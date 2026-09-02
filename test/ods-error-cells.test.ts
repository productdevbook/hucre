import { describe, expect, it } from "vitest"
import { cellError } from "../src/cell-error"
import { OdsStreamWriter } from "../src/ods/incremental-writer"
import { readOds } from "../src/ods/reader"
import { streamOdsRows } from "../src/ods/stream"
import { writeOdsStream } from "../src/ods/stream-writer"
import { writeOds } from "../src/ods/writer"

const collect = async (stream: ReadableStream<Uint8Array>): Promise<Uint8Array> => {
  const parts: Uint8Array[] = []
  for await (const chunk of stream as unknown as AsyncIterable<Uint8Array>) parts.push(chunk)
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

// ODF has no error value type. LibreOffice writes a string cell marked
// calcext:value-type="error"; hucre writes and reads the same mark, so an
// error stays an error through ODS and the text "#N/A" stays text.
describe("error cells through ODS", () => {
  const rows = [[cellError("#DIV/0!"), "#DIV/0!", 1]]

  it("round-trip through writeOds", async () => {
    const wb = await readOds(await writeOds({ sheets: [{ name: "S", rows }] }))
    expect(wb.sheets[0]!.rows[0]).toEqual([cellError("#DIV/0!"), "#DIV/0!", 1])
  })

  it("round-trip through writeOdsStream", async () => {
    const bytes = await collect(writeOdsStream(rows, { name: "S" }))
    expect((await readOds(bytes)).sheets[0]!.rows[0]).toEqual([cellError("#DIV/0!"), "#DIV/0!", 1])
  })

  it("round-trip through OdsStreamWriter", async () => {
    const w = new OdsStreamWriter({ name: "S" })
    w.addRow(rows[0]!)
    expect((await readOds(await w.finish())).sheets[0]!.rows[0]).toEqual([
      cellError("#DIV/0!"),
      "#DIV/0!",
      1,
    ])
  })

  it("the streaming reader agrees", async () => {
    const bytes = await writeOds({ sheets: [{ name: "S", rows }] })
    const seen = []
    for await (const row of streamOdsRows(bytes)) seen.push(row.values)
    expect(seen).toEqual([[cellError("#DIV/0!"), "#DIV/0!", 1]])
  })
})
