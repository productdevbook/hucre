import { describe, it, expect } from "vitest"
import { parseXml } from "../src/xml/parser"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import {
  createSharedStrings,
  writeSharedStringsXml,
  writeWorksheetXml,
} from "../src/xlsx/worksheet-writer"
import { xmlEscapeAttr } from "../src/xml/writer"
import { sortRows } from "../src/sheet-ops"
import { parseCsv, writeCsvObjects } from "../src/csv/index"
import type { WriteSheet, Sheet } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector()
  const ss = createSharedStrings()
  const result = writeWorksheetXml(sheet, styles, ss)
  return result.xml
}

function parseSheet(xml: string) {
  return parseXml(xml)
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("")
}

// ── #101: dimension element ──────────────────────────────────────────

describe("#101: <dimension> element", () => {
  it("emits <dimension> with correct ref after sheetFormatPr", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [
        ["A", "B", "C"],
        [1, 2, 3],
        [4, 5, 6],
      ],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const dimension = findChild(doc, "dimension")
    expect(dimension).toBeDefined()
    expect(dimension.attrs["ref"]).toBe("A1:C3")
  })

  it("calculates dimension from column defs when wider than data", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      columns: [{ width: 10 }, { width: 20 }, { width: 30 }, { width: 40 }],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const dimension = findChild(doc, "dimension")
    expect(dimension).toBeDefined()
    expect(dimension.attrs["ref"]).toBe("A1:D1")
  })
})

// ── #102: printOptions element ───────────────────────────────────────

describe("#102: <printOptions> element", () => {
  it("emits <printOptions> when pageSetup exists", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      pageSetup: { orientation: "landscape" },
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const printOptions = findChild(doc, "printOptions")
    expect(printOptions).toBeDefined()
    expect(printOptions.attrs["headings"]).toBe("0")
    expect(printOptions.attrs["gridLines"]).toBe("0")
  })

  it("does not emit <printOptions> when pageSetup is absent", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const printOptions = findChild(doc, "printOptions")
    expect(printOptions).toBeUndefined()
  })
})

// ── #107: SST count and uniqueCount ──────────────────────────────────

describe("#107: SST count and uniqueCount", () => {
  it("has both count and uniqueCount attributes on <sst>", () => {
    const ss = createSharedStrings()
    ss.add("hello")
    ss.add("world")
    ss.add("hello") // duplicate — should not increase uniqueCount

    const xml = writeSharedStringsXml(ss)
    const doc = parseXml(xml)

    expect(doc.attrs["count"]).toBe("2")
    expect(doc.attrs["uniqueCount"]).toBe("2")
  })

  it("handles empty SST", () => {
    const ss = createSharedStrings()
    const xml = writeSharedStringsXml(ss)
    const doc = parseXml(xml)

    expect(doc.attrs["count"]).toBe("0")
    expect(doc.attrs["uniqueCount"]).toBe("0")
  })
})

// ── #109: Font child element ordering ────────────────────────────────

describe("#109: font child element ordering", () => {
  it("emits font children in ECMA-376 order: b, i, u, strike, charset, family, scheme, color, sz, name, vertAlign", () => {
    const styles = createStylesCollector()
    styles.addStyle({
      font: {
        bold: true,
        italic: true,
        underline: "single",
        strikethrough: true,
        charset: 1,
        family: 2,
        scheme: "minor",
        color: { rgb: "FF0000" },
        size: 12,
        name: "Arial",
        vertAlign: "superscript",
      },
    })

    const xml = styles.toXml()
    const doc = parseXml(xml)

    const fontsEl = findChild(doc, "fonts")
    // The second font (index 1) is our custom one
    const fontEls = findChildren(fontsEl, "font")
    const customFont = fontEls[1]
    expect(customFont).toBeDefined()

    // Extract child element tag names in order
    const childTags = customFont.children
      .filter((c: any) => typeof c !== "string")
      .map((c: any) => c.local || c.tag)

    const expectedOrder = [
      "b",
      "i",
      "u",
      "strike",
      "charset",
      "family",
      "scheme",
      "color",
      "sz",
      "name",
      "vertAlign",
    ]

    expect(childTags).toEqual(expectedOrder)
  })
})

// ── #115: xmlEscapeAttr encodes apostrophe ───────────────────────────

describe("#115: xmlEscapeAttr apostrophe", () => {
  it("encodes apostrophe as &apos;", () => {
    expect(xmlEscapeAttr("it's")).toBe("it&apos;s")
  })

  it("encodes double quote as &quot;", () => {
    expect(xmlEscapeAttr('say "hi"')).toBe("say &quot;hi&quot;")
  })

  it("encodes both in same string", () => {
    expect(xmlEscapeAttr(`it's "good"`)).toBe("it&apos;s &quot;good&quot;")
  })
})

// ── #92: Error cell values ───────────────────────────────────────────

describe("#92: error cell values", () => {
  it("writes #VALUE! as error cell with t='e'", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["#VALUE!"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const sheetData = findChild(doc, "sheetData")
    const row = findChild(sheetData, "row")
    const cell = findChild(row, "c")

    expect(cell.attrs["t"]).toBe("e")
    const v = findChild(cell, "v")
    expect(getElementText(v)).toBe("#VALUE!")
  })

  it("writes #N/A as error cell", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["#N/A"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const sheetData = findChild(doc, "sheetData")
    const row = findChild(sheetData, "row")
    const cell = findChild(row, "c")

    expect(cell.attrs["t"]).toBe("e")
    const v = findChild(cell, "v")
    expect(getElementText(v)).toBe("#N/A")
  })

  it("writes #DIV/0! as error cell", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["#DIV/0!"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const sheetData = findChild(doc, "sheetData")
    const row = findChild(sheetData, "row")
    const cell = findChild(row, "c")

    expect(cell.attrs["t"]).toBe("e")
  })

  it("does NOT treat arbitrary # strings as errors", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["#hashtag"]],
    }

    const xml = writeXml(sheet)
    const doc = parseSheet(xml)

    const sheetData = findChild(doc, "sheetData")
    const row = findChild(sheetData, "row")
    const cell = findChild(row, "c")

    // Should be a shared string, not an error
    expect(cell.attrs["t"]).toBe("s")
  })
})

// ── #91: sortRows ────────────────────────────────────────────────────

describe("#91: sortRows", () => {
  it("sorts ascending by default", () => {
    const sheet: Sheet = {
      name: "Test",
      rows: [
        ["Charlie", 3],
        ["Alice", 1],
        ["Bob", 2],
      ],
    }

    sortRows(sheet, 0)

    expect(sheet.rows[0]![0]).toBe("Alice")
    expect(sheet.rows[1]![0]).toBe("Bob")
    expect(sheet.rows[2]![0]).toBe("Charlie")
  })

  it("sorts descending", () => {
    const sheet: Sheet = {
      name: "Test",
      rows: [
        ["A", 1],
        ["C", 3],
        ["B", 2],
      ],
    }

    sortRows(sheet, 1, "desc")

    expect(sheet.rows[0]![1]).toBe(3)
    expect(sheet.rows[1]![1]).toBe(2)
    expect(sheet.rows[2]![1]).toBe(1)
  })

  it("handles nulls last", () => {
    const sheet: Sheet = {
      name: "Test",
      rows: [
        [null, "X"],
        [1, "A"],
        [null, "Y"],
        [2, "B"],
      ],
    }

    sortRows(sheet, 0, "asc")

    expect(sheet.rows[0]![0]).toBe(1)
    expect(sheet.rows[1]![0]).toBe(2)
    expect(sheet.rows[2]![0]).toBe(null)
    expect(sheet.rows[3]![0]).toBe(null)
  })

  it("sorts mixed types: numbers < strings < booleans", () => {
    const sheet: Sheet = {
      name: "Test",
      rows: [[true], ["hello"], [42], [false], ["abc"], [1]],
    }

    sortRows(sheet, 0, "asc")

    expect(sheet.rows[0]![0]).toBe(1)
    expect(sheet.rows[1]![0]).toBe(42)
    expect(sheet.rows[2]![0]).toBe("abc")
    expect(sheet.rows[3]![0]).toBe("hello")
    expect(sheet.rows[4]![0]).toBe(false)
    expect(sheet.rows[5]![0]).toBe(true)
  })
})

// ── #100: CSV columns option ─────────────────────────────────────────

describe("#100: CSV columns option for writeCsvObjects", () => {
  it("outputs only specified columns in given order", () => {
    const data = [
      { name: "Alice", age: 30, city: "NYC" },
      { name: "Bob", age: 25, city: "LA" },
    ]

    const csv = writeCsvObjects(data, { columns: ["city", "name"] })

    const lines = csv.split("\r\n")
    expect(lines[0]).toBe("city,name")
    expect(lines[1]).toBe("NYC,Alice")
    expect(lines[2]).toBe("LA,Bob")
  })

  it("handles missing keys gracefully", () => {
    const data = [
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 },
    ]

    const csv = writeCsvObjects(data, { columns: ["name", "nonexistent"] })

    const lines = csv.split("\r\n")
    expect(lines[0]).toBe("name,nonexistent")
    expect(lines[1]).toBe("Alice,")
  })
})

// ── #96: CSV skipLines ───────────────────────────────────────────────

describe("#96: CSV skipLines", () => {
  it("skips the first N lines before parsing", () => {
    const input = "metadata line 1\nmetadata line 2\nname,age\nAlice,30\nBob,25"

    const rows = parseCsv(input, { skipLines: 2 })

    expect(rows.length).toBe(3)
    expect(rows[0]).toEqual(["name", "age"])
    expect(rows[1]).toEqual(["Alice", "30"])
    expect(rows[2]).toEqual(["Bob", "25"])
  })

  it("handles CRLF line endings", () => {
    const input = "skip me\r\nskip me too\r\na,b\r\n1,2"

    const rows = parseCsv(input, { skipLines: 2 })

    expect(rows.length).toBe(2)
    expect(rows[0]).toEqual(["a", "b"])
    expect(rows[1]).toEqual(["1", "2"])
  })

  it("returns empty when all lines are skipped", () => {
    const input = "a,b\n1,2"

    const rows = parseCsv(input, { skipLines: 5 })

    expect(rows.length).toBe(0)
  })

  it("works with skipLines: 0 (no effect)", () => {
    const input = "a,b\n1,2"

    const rows = parseCsv(input, { skipLines: 0 })

    expect(rows.length).toBe(2)
  })
})
