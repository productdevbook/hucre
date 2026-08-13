import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { writeOdsStream } from "../src/ods/stream-writer"
import { OdsStreamWriter } from "../src/ods/incremental-writer"
import { streamOdsRows } from "../src/ods/stream"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// A string containing a carriage return, written to ODS, came back with
// a literal `_x000D_` in it:
//
//   in   "with\r\ncrlf"
//   out  "with_x000D_\ncrlf"
//
// `_xHHHH_` is **Excel's** convention, from OOXML. XLSX needs it (a
// literal CR in an XML text node is normalised to LF by XML 1.0 §2.11,
// so an escape is the only way to carry one) and the XLSX reader decodes
// it back, so that pair is self-consistent.
//
// ODF has no such convention. The ODS writer shared `xmlEscape` with the
// XLSX writer and inherited the spelling; nothing on the ODS side decodes
// it. The damage is not confined to hucre either — LibreOffice shows the
// literal `_x000D_` too, because to it that is just seven characters.
//
// The fix is the spelling XML gives you for exactly this: `&#13;`. A
// character reference is not subject to end-of-line normalisation
// (§2.11 covers literal CR in the source), so it survives — and the ODS
// readers already handled it, which is what #493 sorted out when it put
// `normalizeEol` before `decodeEntities`.
//
// Found by a property test over random grids, comparing what went into
// each writer with what came back out of its reader.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

const VALUES = ["with\r\ncrlf", "bare\rcr", "lf\nonly", "trailing\r", "\rleading"]
const ROWS: CellValue[][] = VALUES.map((v) => [v])

async function contentXml(bytes: Uint8Array): Promise<string> {
  return dec.decode(await new ZipReader(bytes).extract("content.xml"))
}

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}

describe("a carriage return survives the ODS round trip", () => {
  it("through writeOds and readOds", async () => {
    const back = (await readOds(await writeOds({ sheets: [{ name: "S", rows: ROWS }] }))).sheets[0]!

    expect(back.rows.map((r) => r[0])).toEqual(VALUES)
  })

  it("and through the streaming reader, which shares nothing with it", async () => {
    const bytes = await writeOds({ sheets: [{ name: "S", rows: ROWS }] })
    const streamed: CellValue[] = []
    for await (const row of streamOdsRows(bytes)) streamed.push(row.values[0]!)

    expect(streamed).toEqual(VALUES)
  })

  it("out of writeOdsStream", async () => {
    const bytes = await drain(
      writeOdsStream(
        ROWS.map((r) => r),
        { name: "S" },
      ),
    )
    const back = (await readOds(bytes)).sheets[0]!

    expect(back.rows.map((r) => r[0])).toEqual(VALUES)
  })

  it("out of OdsStreamWriter", async () => {
    const writer = new OdsStreamWriter({ name: "S" })
    for (const row of ROWS) writer.addRow(row)
    const back = (await readOds(await writer.finish())).sheets[0]!

    expect(back.rows.map((r) => r[0])).toEqual(VALUES)
  })
})

describe("what is actually in the file", () => {
  it("is a character reference, not Excel's escape", async () => {
    // The assertion that matters to every *other* consumer. A reader
    // that decoded `_x000D_` would make hucre's round trip pass while
    // LibreOffice still showed seven literal characters.
    const xml = await contentXml(await writeOds({ sheets: [{ name: "S", rows: ROWS }] }))

    expect(xml).toContain("&#13;")
    expect(xml).not.toContain("_x000D_")
  })

  it("from every ODS writer", async () => {
    const streamed = await contentXml(
      await drain(
        writeOdsStream(
          ROWS.map((r) => r),
          { name: "S" },
        ),
      ),
    )
    const writer = new OdsStreamWriter({ name: "S" })
    for (const row of ROWS) writer.addRow(row)
    const incremental = await contentXml(await writer.finish())

    for (const xml of [streamed, incremental]) {
      expect(xml).toContain("&#13;")
      expect(xml).not.toContain("_x000D_")
    }
  })
})

describe("an attribute is escaped as an attribute", () => {
  // The second half of the same mistake. `writeOdsStream` builds its
  // opening tag by hand and reached for the *text* escaper for values
  // that sit inside quotes. `xmlEscape` does not escape `"` — it has no
  // reason to — so a sheet name containing one closed the attribute
  // early:
  //
  //   <table:table table:name="say "hi"">
  //
  // which is not well-formed. The name came back as `say ` — truncated
  // silently, with no error anywhere. `writeOds` was never affected,
  // because it goes through `xmlElement`, which escapes attributes
  // properly; only the hand-built tag was wrong.
  const AWKWARD_NAME = 'say "hi" & <that>'

  it("survives a sheet name full of markup characters", async () => {
    const bytes = await drain(writeOdsStream([["a"]], { name: AWKWARD_NAME }))
    const wb = await readOds(bytes)

    expect(wb.sheets.map((s) => s.name)).toEqual([AWKWARD_NAME])
  })

  it("the same name the buffered writer already handled", async () => {
    // Pinning the agreement, since the two took different routes to it.
    const buffered = await writeOds({ sheets: [{ name: AWKWARD_NAME, rows: [["a"]] }] })

    expect((await readOds(buffered)).sheets[0]!.name).toBe(AWKWARD_NAME)
  })

  it("and a formula with a quote in a string literal", async () => {
    // `table:formula` is the other hand-built attribute in that file.
    const bytes = await drain(
      writeOdsStream([[{ value: "x", formula: 'IF(A1="q",1,2)' }]], { name: "S" }),
    )

    expect(await contentXml(bytes)).toContain("&quot;")
  })
})

describe("XLSX keeps the convention that is right for it", () => {
  it("still writes _x000D_ and still reads it back", async () => {
    // OOXML's escape is correct here and Excel expects it. This is the
    // line that notices if the ODS fix is applied too widely.
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] })
    const sheet = dec.decode(await new ZipReader(bytes).extract("xl/sharedStrings.xml"))

    expect(sheet).toContain("_x000D_")
    expect((await readXlsx(bytes)).sheets[0]!.rows.map((r) => r[0])).toEqual(VALUES)
  })
})
