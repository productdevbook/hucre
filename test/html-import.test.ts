import { describe, expect, it } from "vitest"
import { fromHtml } from "../src/export/html-import"
import { toHtml } from "../src/export/html"
import { ParseError } from "../src/errors"
import { MAX_TOTAL_CELLS } from "../src/limits"
import type { Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// fromHtml reads markup hucre did not write. Everything below is about
// input a scraper hands it: third-party tables, nested tables, cells that
// were never closed, and counts chosen by whoever served the page. See
// #410.
// ═══════════════════════════════════════════════════════════════════════

const table = (cells: string) => `<table><tr>${cells}</tr></table>`

// ── Type inference ──────────────────────────────────────────────────

describe("fromHtml type inference", () => {
  it("keeps leading zeros — a ZIP code is not the number 7", () => {
    // `Number("007")` is 7, and a bare Number() call is what this used to be.
    const sheet = fromHtml(table("<td>007</td><td>0123</td><td>00</td>"))
    expect(sheet.rows[0]).toEqual(["007", "0123", "00"])
  })

  it("does not read hex, binary, octal or Infinity as numbers", () => {
    // All four are things `Number()` accepts and no spreadsheet cell means.
    const sheet = fromHtml(table("<td>0x1A</td><td>0b101</td><td>0o17</td><td>Infinity</td>"))
    expect(sheet.rows[0]).toEqual(["0x1A", "0b101", "0o17", "Infinity"])
  })

  it("infers numbers, booleans and ISO dates like parseCsv does", () => {
    const sheet = fromHtml(
      table("<td>42</td><td>3.14</td><td>1,234</td><td>true</td><td>2024-01-15</td><td>hello</td>"),
    )
    const row = sheet.rows[0]!
    expect(row.slice(0, 4)).toEqual([42, 3.14, 1234, true])
    expect(row[4]).toBeInstanceOf(Date)
    expect((row[4] as Date).toISOString()).toBe("2024-01-15T00:00:00.000Z")
    expect(row[5]).toBe("hello")
  })

  it("returns cell text verbatim under typeInference: false", () => {
    const sheet = fromHtml(table("<td>42</td><td>true</td><td>2024-01-15</td>"), {
      typeInference: false,
    })
    expect(sheet.rows[0]).toEqual(["42", "true", "2024-01-15"])
  })

  it("coerces leading zeros when preserveLeadingZeros is off", () => {
    const sheet = fromHtml(table("<td>007</td>"), { preserveLeadingZeros: false })
    expect(sheet.rows[0]).toEqual([7])
  })
})

// ── Nested tables ───────────────────────────────────────────────────

describe("fromHtml nested tables", () => {
  it("keeps reading the outer table after an inner </table>", () => {
    // The inner close used to clear a boolean `inTable`, silently dropping
    // every remaining row of the table the caller actually asked for.
    const sheet = fromHtml(`<table>
      <tr><td>outer-1</td></tr>
      <tr><td><table><tr><td>inner</td></tr></table></td></tr>
      <tr><td>outer-3</td></tr>
      <tr><td>outer-4</td></tr>
    </table>`)
    expect(sheet.rows).toEqual([["outer-1"], ["inner"], ["outer-3"], ["outer-4"]])
  })

  it("folds a nested table's text into the cell that contains it", () => {
    const sheet = fromHtml(
      "<table><tr><td>a<table><tr><td>x</td><td>y</td></tr></table></td><td>b</td></tr></table>",
    )
    expect(sheet.rows).toEqual([["axy", "b"]])
  })

  it("survives three levels of nesting", () => {
    const sheet = fromHtml(
      "<table><tr><td><table><tr><td><table><tr><td>deep</td></tr></table></td></tr></table></td></tr>" +
        "<tr><td>after</td></tr></table>",
    )
    expect(sheet.rows).toEqual([["deep"], ["after"]])
  })

  it("keeps the outer table's merges intact across a nested table", () => {
    const sheet = fromHtml(
      "<table><tr><td><table><tr><td>i</td></tr></table></td></tr>" +
        '<tr><td colspan="2">wide</td></tr></table>',
    )
    expect(sheet.merges).toEqual([{ startRow: 1, startCol: 0, endRow: 1, endCol: 1 }])
  })
})

// ── Unclosed markup ─────────────────────────────────────────────────

describe("fromHtml unclosed markup", () => {
  it("keeps a cell that was never closed", () => {
    // `<td>a<td>b` is legal HTML; the second open tag ends the first cell.
    const sheet = fromHtml("<table><tr><td>a<td>b</tr></table>")
    expect(sheet.rows).toEqual([["a", "b"]])
  })

  it("keeps the last cell and row when only </table> closes them", () => {
    const sheet = fromHtml("<table><tr><td>a</td><td>b</table>")
    expect(sheet.rows).toEqual([["a", "b"]])
  })

  it("keeps rows when a <tr> is never closed", () => {
    const sheet = fromHtml("<table><tr><td>a</td><tr><td>b</td></table>")
    expect(sheet.rows).toEqual([["a"], ["b"]])
  })

  it("emits what it read when the document ends mid-table", () => {
    const sheet = fromHtml("<table><tr><td>a</td></tr><tr><td>b</td>")
    expect(sheet.rows).toEqual([["a"], ["b"]])
  })
})

// ── Malformed markup ────────────────────────────────────────────────

describe("fromHtml malformed markup", () => {
  it("returns the rows read so far instead of throwing on an unterminated comment", () => {
    // "Best-effort" was in the doc comment while parseSax threw straight
    // through the caller.
    const sheet = fromHtml("<table><tr><td>a</td></tr></table><!-- never closed")
    expect(sheet.rows).toEqual([["a"]])
  })

  it("survives a truncated opening tag", () => {
    const sheet = fromHtml("<table><tr><td>a</td></tr></table><table")
    expect(sheet.rows).toEqual([["a"]])
  })

  it("survives a truncated closing tag", () => {
    const sheet = fromHtml("<table><tr><td>x</td></tr></table></tr")
    expect(sheet.rows).toEqual([["x"]])
  })
})

// ── Structure that toHtml writes ────────────────────────────────────

describe("fromHtml structure", () => {
  it("reports a <thead> row as the header row", () => {
    const sheet = fromHtml(
      "<table><thead><tr><th>Name</th><th>Age</th></tr></thead>" +
        "<tbody><tr><td>Ada</td><td>36</td></tr></tbody></table>",
    )
    expect(sheet.a11y?.headerRow).toBe(0)
  })

  it("reports an all-<th> row as the header row without a <thead>", () => {
    const sheet = fromHtml(
      "<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>",
    )
    expect(sheet.a11y?.headerRow).toBe(0)
  })

  it("does not call a row of mixed <th>/<td> a header row", () => {
    const sheet = fromHtml("<table><tr><th>A</th><td>1</td></tr></table>")
    expect(sheet.a11y).toBeUndefined()
  })

  it("surfaces <caption> as the sheet summary", () => {
    const sheet = fromHtml("<table><caption>Q3 revenue</caption><tr><td>1</td></tr></table>")
    expect(sheet.a11y?.summary).toBe("Q3 revenue")
    expect(sheet.rows).toEqual([[1]])
  })

  it("leaves a11y undefined when there is no caption and no header", () => {
    const sheet = fromHtml("<table><tr><td>1</td></tr></table>")
    expect(sheet.a11y).toBeUndefined()
  })
})

// ── Type classes ────────────────────────────────────────────────────

describe("fromHtml type classes", () => {
  const source: Sheet = {
    name: "S",
    rows: [
      ["Name", "Flag"],
      ["hucre", true],
      [new Date(Date.UTC(2024, 0, 15)), null],
      ["42", 42],
    ],
  }

  it("recovers booleans, dates, nulls and numbers from toHtml output", () => {
    const sheet = fromHtml(toHtml(source, { hasHeaderRow: true }))
    expect(sheet.rows[1]).toEqual(["hucre", true])
    expect(sheet.rows[2]?.[0]).toBeInstanceOf(Date)
    expect(sheet.rows[2]?.[1]).toBeNull()
    expect(sheet.rows[3]).toEqual([42, 42])
  })

  it("honours declared types even with typeInference off", () => {
    // A class is the writer stating a type, not the reader guessing one.
    const sheet = fromHtml(toHtml(source), { typeInference: false })
    expect(sheet.rows[1]).toEqual(["hucre", true])
    expect(sheet.rows[3]?.[0]).toBe("42")
    expect(sheet.rows[3]?.[1]).toBe(42)
  })

  it("reads a custom classPrefix", () => {
    const html = toHtml(source, { classPrefix: "sp" })
    expect(fromHtml(html, { classPrefix: "sp" }).rows[1]?.[1]).toBe(true)
    // Wrong prefix: the class is not recognised, inference sees "true".
    expect(fromHtml(html, { classPrefix: "zz" }).rows[1]?.[1]).toBe(true)
    expect(fromHtml(html, { classPrefix: "zz", typeInference: false }).rows[1]?.[1]).toBe("true")
  })

  it("ignores type classes under classes: false", () => {
    const sheet = fromHtml(toHtml(source), { classes: false, typeInference: false })
    expect(sheet.rows[1]?.[1]).toBe("true")
    expect(sheet.rows[2]?.[1]).toBeNull()
  })

  it("prefers the text when a class and its content disagree", () => {
    const sheet = fromHtml(table('<td class="hucre-num">not a number</td>'))
    expect(sheet.rows[0]).toEqual(["not a number"])
  })
})

// ── Resource bounds ─────────────────────────────────────────────────

describe("fromHtml resource bounds", () => {
  it("refuses a table whose colspans describe more cells than a sheet holds", () => {
    // 52 KB of markup, 24 million array slots. Each `<td>` is 30 bytes and
    // costs 16,384 entries.
    const html = `<table>${'<tr><td colspan="16384">x</td></tr>'.repeat(1500)}</table>`
    expect(html.length).toBeLessThan(60_000)
    expect(() => fromHtml(html)).toThrow(ParseError)
    expect(() => fromHtml(html)).toThrow(String(MAX_TOTAL_CELLS))
  }, 20_000)

  it("refuses a page of hostile rowspans, not just one", () => {
    // The per-cell bound is paid once per cell; two cells pay it twice.
    const html = `<table>${'<tr><td rowspan="1000000">x</td></tr>'.repeat(2)}</table>`
    expect(() => fromHtml(html)).toThrow(ParseError)
  }, 20_000)

  it("still reads an ordinary table with spans", () => {
    const sheet = fromHtml(
      '<table><tr><td rowspan="2" colspan="3">x</td><td>y</td></tr><tr><td>z</td></tr></table>',
    )
    expect(sheet.rows).toEqual([
      ["x", null, null, "y"],
      [null, null, null, "z"],
    ])
  })
})

// ── Whitespace ──────────────────────────────────────────────────────

describe("fromHtml whitespace", () => {
  it("trims cell text, because indentation is markup and not data", () => {
    const sheet = fromHtml("<table>\n  <tr>\n    <td>\n      42\n    </td>\n  </tr>\n</table>")
    expect(sheet.rows).toEqual([[42]])
  })

  it("does not round-trip a padded string, and toHtml is the one that keeps it", () => {
    const html = toHtml({ name: "S", rows: [["  padded  "]] })
    expect(html).toContain("<td>  padded  </td>")
    expect(fromHtml(html).rows).toEqual([["padded"]])
  })
})
