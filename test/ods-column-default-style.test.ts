import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"

// ═══════════════════════════════════════════════════════════════════════
// LibreOffice puts a column's format on the column, not on its cells.
// `<table:table-column table:default-cell-style-name="ce1"/>` names an
// automatic style, and that style names the data style — so a date
// column's `yyyy-mm-dd` lives there and nowhere else. The cells
// themselves say `table:style-name="Default"`, a *named* style in
// `styles.xml`, which this reader does not open.
//
// hucre read neither, so a LibreOffice document came back with its values
// and none of its formats.
//
// The last open item in `SPEC-COVERAGE.md`: `table:default-cell-style-name`
// was in the grammar, in `libreoffice-basic.ods`, and nowhere in `src/`.
//
// Only the `content.xml` half is reachable. A column pointing at a named
// style — `"Default"`, which LibreOffice writes on the unformatted
// columns — still resolves to nothing, because that style is in
// `styles.xml`. PARITY records that, and this does not change it.
// ═══════════════════════════════════════════════════════════════════════

const LIBREOFFICE = "test/fixtures/third-party/libreoffice-basic.ods"

describe("a column's default style reaches its cells", () => {
  it("gives the LibreOffice date column its format back", async () => {
    const bytes = new Uint8Array(readFileSync(LIBREOFFICE))
    const wb = await readOds(bytes, { readStyles: true })
    const sheet = wb.sheets[0]!

    // Column C (index 2) is the date column; its style is on the column.
    const dateCell = sheet.cells?.get("1,2")

    expect(dateCell?.style?.numFmt).toBe("yyyy-mm-dd")
  })

  it("without readStyles, nothing changes", async () => {
    // Styles are not being read, so a column default is not information
    // the caller asked for — the same rule the rest of the reader follows.
    const bytes = new Uint8Array(readFileSync(LIBREOFFICE))
    const wb = await readOds(bytes)

    expect(wb.sheets[0]!.cells?.get("1,2")?.style).toBeUndefined()
  })

  it("and the values are untouched either way", async () => {
    const bytes = new Uint8Array(readFileSync(LIBREOFFICE))
    const wb = await readOds(bytes, { readStyles: true })

    expect(wb.sheets[0]!.rows[0]).toEqual(["Name", "Qty", "Date", "Active", "Total"])
    expect(wb.sheets[0]!.rows[1]![2]).toBeInstanceOf(Date)
  })
})

describe("a cell's own style still wins", () => {
  it("over the column it sits in", async () => {
    // hucre's own writer puts the style on the cell. That must not be
    // overridden by a column default, whichever order they are read in.
    const bytes = await writeOds({
      sheets: [
        {
          name: "S",
          rows: [[1, 2]],
          cells: new Map([["0,0", { value: 1, style: { numFmt: "0.000" } }]]),
        },
      ],
    })
    const cell = (await readOds(bytes, { readStyles: true })).sheets[0]!.cells?.get("0,0")

    expect(cell?.style?.numFmt).toBe("0.000")
  })
})

describe("what is still not reachable", () => {
  it("a column naming a style from styles.xml resolves to nothing", async () => {
    // LibreOffice writes `table:default-cell-style-name="Default"` on the
    // columns it did not format. `Default` lives in `styles.xml`, which
    // this reader does not open, so those cells stay unstyled — which is
    // the correct answer here, since the default style is the absence of
    // formatting.
    const bytes = new Uint8Array(readFileSync(LIBREOFFICE))
    const wb = await readOds(bytes, { readStyles: true })

    expect(wb.sheets[0]!.cells?.get("1,0")?.style?.numFmt).toBeUndefined()
  })
})
