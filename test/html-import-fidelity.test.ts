import { describe, expect, it } from "vitest"
import { fromHtml } from "../src/export/html-import"
import { decodeHtmlEntities } from "../src/export/html-entities"

// ═══════════════════════════════════════════════════════════════════════
// #439 Part 7 — `fromHtml` is documented as parsing "anyone's table", but
// it runs on an XML parser, and HTML is not XML in four ways that all
// show up in real scraped markup.
// ═══════════════════════════════════════════════════════════════════════

describe("named HTML entities are decoded", () => {
  it("expands the ones real markup is full of", () => {
    const sheet = fromHtml(
      "<table><tr><td>1&nbsp;234&nbsp;&euro;</td><td>a&mdash;b</td><td>20&deg;C</td></tr></table>",
      { typeInference: false },
    )

    expect(sheet.rows[0]).toEqual(["1 234 €", "a—b", "20°C"])
  })

  it("still expands the XML five and numeric references", () => {
    const sheet = fromHtml("<table><tr><td>&amp;&lt;&gt;&#65;&#x42;</td></tr></table>")

    expect(sheet.rows[0]![0]).toBe("&<>AB")
  })

  it("leaves an unknown reference alone rather than guessing", () => {
    const sheet = fromHtml("<table><tr><td>&notarealentity;</td></tr></table>")

    expect(sheet.rows[0]![0]).toBe("&notarealentity;")
  })

  it("decodes a caption too", () => {
    const sheet = fromHtml("<table><caption>Q1&nbsp;2024</caption><tr><td>1</td></tr></table>")

    expect(sheet.a11y!.summary).toBe("Q1 2024")
  })

  it("decodeHtmlEntities is a no-op on text with no ampersand", () => {
    expect(decodeHtmlEntities("plain")).toBe("plain")
  })
})

describe("<br> is a line break, not nothing", () => {
  it("keeps the two lines apart", () => {
    const sheet = fromHtml("<table><tr><td>Line one<br>Line two</td></tr></table>", {
      typeInference: false,
    })

    expect(sheet.rows[0]![0]).toBe("Line one\nLine two")
  })

  it("handles the self-closing spelling", () => {
    const sheet = fromHtml("<table><tr><td>a<br/>b</td></tr></table>", { typeInference: false })

    expect(sheet.rows[0]![0]).toBe("a\nb")
  })
})

describe("<script> and <style> are raw text", () => {
  it("does not let script source become a cell value, or end the cell", () => {
    const sheet = fromHtml(
      "<table><tr><td><script>var x = '</td>'</script>real</td><td>next</td></tr></table>",
      { typeInference: false },
    )

    expect(sheet.rows[0]).toEqual(["real", "next"])
  })

  it("ignores a <style> block inside a cell", () => {
    const sheet = fromHtml(
      "<table><tr><td><style>td { color: red }</style>value</td></tr></table>",
      { typeInference: false },
    )

    expect(sheet.rows[0]![0]).toBe("value")
  })

  it("ignores a script between rows", () => {
    const sheet = fromHtml(
      "<table><tr><td>a</td></tr><script>document.write('<tr><td>ghost</td></tr>')</script><tr><td>b</td></tr></table>",
      { typeInference: false },
    )

    expect(sheet.rows).toEqual([["a"], ["b"]])
  })
})

describe("a document with several tables", () => {
  const TWO =
    "<table><caption>one</caption><tr><td>1</td><td>2</td></tr></table>" +
    "<table><caption>two</caption><tr><td>a</td></tr></table>"

  it("reads the first by default rather than concatenating them", () => {
    const sheet = fromHtml(TWO)

    expect(sheet.rows).toEqual([[1, 2]])
    expect(sheet.a11y!.summary).toBe("one")
  })

  it("reads the one asked for", () => {
    const sheet = fromHtml(TWO, { tableIndex: 1, typeInference: false })

    expect(sheet.rows).toEqual([["a"]])
    expect(sheet.a11y!.summary).toBe("two")
  })

  it("keeps merges to the chosen table, numbered from its own first row", () => {
    const sheet = fromHtml(
      '<table><tr><td colspan="2">x</td></tr></table>' +
        '<table><tr><td colspan="2">y</td></tr></table>',
      { tableIndex: 1, typeInference: false },
    )

    expect(sheet.rows).toEqual([["y", null]])
    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }])
  })

  it("yields an empty sheet for an index past the last table", () => {
    const sheet = fromHtml(TWO, { tableIndex: 5 })

    expect(sheet.rows).toEqual([])
    expect(sheet.a11y).toBeUndefined()
  })

  it("still reads a nested table as the text of its containing cell", () => {
    const sheet = fromHtml(
      "<table><tr><td>outer<table><tr><td>inner</td></tr></table></td><td>b</td></tr></table>",
      { typeInference: false },
    )

    expect(sheet.rows).toEqual([["outerinner", "b"]])
  })
})

describe("<tfoot> lands where it renders, not where it is declared", () => {
  it("moves a footer declared before the body to the end", () => {
    const sheet = fromHtml(
      "<table>" +
        "<thead><tr><th>h</th></tr></thead>" +
        "<tfoot><tr><td>total</td></tr></tfoot>" +
        "<tbody><tr><td>one</td></tr><tr><td>two</td></tr></tbody>" +
        "</table>",
      { typeInference: false },
    )

    expect(sheet.rows).toEqual([["h"], ["one"], ["two"], ["total"]])
  })

  it("keeps the header index pointing at the header", () => {
    const sheet = fromHtml(
      "<table><tfoot><tr><td>total</td></tr></tfoot>" +
        "<thead><tr><th>h</th></tr></thead>" +
        "<tbody><tr><td>one</td></tr></tbody></table>",
      { typeInference: false },
    )

    expect(sheet.rows).toEqual([["h"], ["one"], ["total"]])
    expect(sheet.a11y!.headerRow).toBe(0)
  })

  it("moves the footer's merges with it", () => {
    const sheet = fromHtml(
      '<table><tfoot><tr><td colspan="2">total</td></tr></tfoot>' +
        "<tbody><tr><td>a</td><td>b</td></tr></tbody></table>",
      { typeInference: false },
    )

    expect(sheet.rows).toEqual([
      ["a", "b"],
      ["total", null],
    ])
    expect(sheet.merges).toEqual([{ startRow: 1, startCol: 0, endRow: 1, endCol: 1 }])
  })

  it("leaves a footer already at the end alone", () => {
    const sheet = fromHtml(
      "<table><tbody><tr><td>one</td></tr></tbody><tfoot><tr><td>total</td></tr></tfoot></table>",
      { typeInference: false },
    )

    expect(sheet.rows).toEqual([["one"], ["total"]])
  })
})
