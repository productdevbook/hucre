import { describe, expect, it } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { ZipReader } from "../src/zip/reader"

// ═══════════════════════════════════════════════════════════════════════
// Named ranges were dropped on the way into ODS and never read back out.
// `PARITY.md` lists them among the things ODS does not model in either
// direction, and `SPEC-COVERAGE.md` had `table:named-expressions` as an
// open item — in the ODF grammar, in a LibreOffice document, and nowhere
// in `src/`.
//
// The reason this waited is worth recording: the LibreOffice fixture
// writes `<table:named-expressions/>` *empty*, so the corpus had no
// example of the populated form to work from, and implementing against
// the spec alone would have left the round trip checking itself. What
// changed is #552 — the output can now be validated against the OASIS
// grammar, which is the authority the corpus was standing in for.
//
// ODF puts them in the epilogue of `<office:spreadsheet>`, after every
// `<table:table>`, and spells an address `$Sheet1.$A$1:$Sheet1.$B$5`
// where Excel writes `Sheet1!$A$1:$B$5`.
// ═══════════════════════════════════════════════════════════════════════

const dec = new TextDecoder()

async function contentXml(bytes: Uint8Array): Promise<string> {
  return dec.decode(await new ZipReader(bytes).extract("content.xml"))
}

const SHEETS = [
  {
    name: "Data",
    rows: [
      ["a", "b"],
      [1, 2],
    ],
  },
  { name: "My Sheet", rows: [["x"]] },
]

describe("a named range survives ODS", () => {
  it("round-trips a workbook-level name", async () => {
    const bytes = await writeOds({
      sheets: SHEETS,
      namedRanges: [{ name: "Region", range: "Data!$A$1:$B$2" }],
    })
    const wb = await readOds(bytes)

    expect(wb.namedRanges).toEqual([{ name: "Region", range: "Data!$A$1:$B$2" }])
  })

  it("several of them, in order", async () => {
    const bytes = await writeOds({
      sheets: SHEETS,
      namedRanges: [
        { name: "One", range: "Data!$A$1:$A$1" },
        { name: "Two", range: "Data!$B$1:$B$2" },
      ],
    })
    const wb = await readOds(bytes)

    expect(wb.namedRanges?.map((r) => r.name)).toEqual(["One", "Two"])
    expect(wb.namedRanges?.[1]?.range).toBe("Data!$B$1:$B$2")
  })

  it("with a sheet name that needs quoting", async () => {
    const bytes = await writeOds({
      sheets: SHEETS,
      namedRanges: [{ name: "Q", range: "'My Sheet'!$A$1:$A$1" }],
    })
    const wb = await readOds(bytes)

    expect(wb.namedRanges?.[0]?.range).toBe("'My Sheet'!$A$1:$A$1")

    // And quoted in the file, which the round trip above cannot tell
    // you: this reader accepts `$My Sheet.$A$1` and the grammar does
    // not. Its regular expression for `cell-range-address` allows a run
    // with no dot, space or apostrophe, or a quoted string — nothing
    // else, and `jing` said so while the round trip was green.
    //
    // The apostrophes arrive XML-escaped, because they are in an
    // attribute: the value a parser sees is `$'My Sheet'.$A$1`.
    expect(await contentXml(bytes)).toContain("$&apos;My Sheet&apos;.$A$1")
  })

  it("in the place ODF puts them", async () => {
    // The epilogue: after every table, inside office:spreadsheet. The
    // grammar is a sequence, so anywhere else is invalid.
    const xml = await contentXml(
      await writeOds({ sheets: SHEETS, namedRanges: [{ name: "R", range: "Data!$A$1:$B$2" }] }),
    )

    expect(xml).toContain("<table:named-expressions>")
    expect(xml).toContain('table:name="R"')
    expect(xml).toContain('table:cell-range-address="$Data.$A$1:$Data.$B$2"')
    expect(xml.indexOf("</table:table>")).toBeLessThan(xml.indexOf("<table:named-expressions>"))
  })
})

describe("nothing is written when there is nothing to write", () => {
  it("no namedRanges means no element", async () => {
    const xml = await contentXml(await writeOds({ sheets: SHEETS }))

    expect(xml).not.toContain("named-expressions")
  })

  it("an empty array means no element either", async () => {
    const xml = await contentXml(await writeOds({ sheets: SHEETS, namedRanges: [] }))

    expect(xml).not.toContain("named-expressions")
  })

  it("and a document without them reads back undefined", async () => {
    const wb = await readOds(await writeOds({ sheets: SHEETS }))

    expect(wb.namedRanges).toBeUndefined()
  })
})
