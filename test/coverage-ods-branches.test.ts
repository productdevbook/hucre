import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"
import { readOds } from "../src/ods/reader"
import { writeOds } from "../src/ods/writer"
import { streamOdsRows } from "../src/ods/stream"
import { readOdsObjects, writeOdsObjects } from "../src/ods/objects"
import { ParseError, ZipError } from "../src/errors"
import type { Cell, PatternFill, StreamRow, Workbook, WriteSheet } from "../src/_types"

const enc = new TextEncoder()
const dec = new TextDecoder("utf-8")

// ── Building ODF fragments ──────────────────────────────────────────
//
// Every ODF namespace a real content.xml declares on its root element.
// Declaring them all keeps the fragments below valid documents rather
// than XML the parser only tolerates by accident.

const NS = [
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"`,
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"`,
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`,
  `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"`,
  `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"`,
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`,
  `xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"`,
  `xmlns:dc="http://purl.org/dc/elements/1.1/"`,
  `xmlns:xlink="http://www.w3.org/1999/xlink"`,
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"`,
].join(" ")

/** A complete `content.xml` around a `<office:spreadsheet>` body. */
function contentXml(body: string, automaticStyles = ""): string {
  const styles = automaticStyles
    ? `<office:automatic-styles>${automaticStyles}</office:automatic-styles>`
    : ""
  return (
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content ${NS} office:version="1.3">` +
    styles +
    `<office:body><office:spreadsheet>${body}</office:spreadsheet></office:body>` +
    `</office:document-content>`
  )
}

/** A complete `meta.xml` around the `<office:meta>` children. */
function metaXml(inner: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-meta ${NS} office:version="1.3">` +
    `<office:meta>${inner}</office:meta>` +
    `</office:document-meta>`
  )
}

interface OdsParts {
  /** Full content.xml. Omit to leave the entry out of the archive. */
  content?: string
  meta?: string
  /** Defaults to the spreadsheet media type. */
  mimetype?: string | null
}

async function odsFile(parts: OdsParts): Promise<Uint8Array> {
  const zip = new ZipWriter()
  if (parts.mimetype !== null) {
    zip.add(
      "mimetype",
      enc.encode(parts.mimetype ?? "application/vnd.oasis.opendocument.spreadsheet"),
      { compress: false },
    )
  }
  if (parts.content !== undefined) zip.add("content.xml", enc.encode(parts.content))
  if (parts.meta !== undefined) zip.add("meta.xml", enc.encode(parts.meta))
  return await zip.build()
}

/** One table containing exactly the given `<table:table-row>` markup. */
function table(rowsXml: string, name = "S"): string {
  return `<table:table table:name="${name}">${rowsXml}</table:table>`
}

async function readBody(body: string, automaticStyles = ""): Promise<Workbook> {
  return await readOds(await odsFile({ content: contentXml(body, automaticStyles) }), {
    readStyles: automaticStyles !== "",
  })
}

/**
 * Reconstruct the Excel-style `numFmt` the reader derives from a
 * `<number:*-style>` data style, by pointing a cell style at it and
 * reading the resulting `CellStyle`.
 */
async function numFmtFor(dataStyleXml: string, dataStyleName: string): Promise<string | undefined> {
  const styles =
    dataStyleXml +
    `<style:style style:name="ce1" style:family="table-cell" style:data-style-name="${dataStyleName}"/>`
  const body = table(
    `<table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>`,
  )
  const wb = await readBody(body, styles)
  return wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt
}

async function collectStream(data: Uint8Array): Promise<StreamRow[]> {
  const out: StreamRow[] = []
  for await (const row of streamOdsRows(data)) out.push(row)
  return out
}

// ── Reading back what writeOds produced ─────────────────────────────

async function contentOf(data: Uint8Array, path = "content.xml"): Promise<string> {
  const zip = new ZipReader(data)
  return dec.decode(await zip.extract(path))
}

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — data styles (number formats)
// ═══════════════════════════════════════════════════════════════════════
//
// LibreOffice stores a cell's number format as a `<number:*-style>`
// element referenced through `style:data-style-name`. The reader
// re-serializes that element tree back into an Excel format code, which
// is the only representation `CellStyle.numFmt` can carry.

describe("ODS reader — number data styles", () => {
  it("rebuilds a grouped two-decimal code from number:number attributes", async () => {
    const code = await numFmtFor(
      `<number:number-style style:name="N100"><number:number number:decimal-places="2" number:grouping="true" number:min-integer-digits="1"/></number:number-style>`,
      "N100",
    )
    expect(code).toBe("#,##0.00")
  })

  it("treats a missing number:decimal-places as zero decimals", async () => {
    // ODF makes the attribute optional; omitting it means an integer format.
    const code = await numFmtFor(
      `<number:number-style style:name="N100"><number:number number:min-integer-digits="1"/></number:number-style>`,
      "N100",
    )
    expect(code).toBe("0")
  })

  it("falls back to zero decimals when decimal-places is not a number", async () => {
    const code = await numFmtFor(
      `<number:number-style style:name="N100"><number:number number:decimal-places="lots"/></number:number-style>`,
      "N100",
    )
    expect(code).toBe("0")
  })

  it("caps an absurd decimal-places so the format code cannot be a memory bomb", async () => {
    // A hand-crafted file can ask for a billion decimals; the reader clamps
    // the repeat the same way it clamps text:c. See #363.
    const code = await numFmtFor(
      `<number:number-style style:name="N100"><number:number number:decimal-places="900000000"/></number:number-style>`,
      "N100",
    )
    expect(code!.length).toBe("0.".length + 100_000)
  })

  it("appends the percent sign when the style has no literal % child", async () => {
    // Some producers rely on the element name alone to mean "percent".
    const code = await numFmtFor(
      `<number:percentage-style style:name="N101"><number:number number:decimal-places="0"/></number:percentage-style>`,
      "N101",
    )
    expect(code).toBe("0%")
  })

  it("keeps a single-character separator bare rather than quoting it", async () => {
    const code = await numFmtFor(
      `<number:percentage-style style:name="N101"><number:number number:decimal-places="1"/><number:text>%</number:text></number:percentage-style>`,
      "N101",
    )
    expect(code).toBe("0.0%")
  })

  it("quotes a multi-character literal from number:text", async () => {
    const code = await numFmtFor(
      `<number:number-style style:name="N100"><number:number number:decimal-places="0"/><number:text> kg</number:text></number:number-style>`,
      "N100",
    )
    expect(code).toBe(`0" kg"`)
  })

  it("places a leading currency symbol before the number", async () => {
    const code = await numFmtFor(
      `<number:currency-style style:name="N102"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places="2" number:grouping="true"/></number:currency-style>`,
      "N102",
    )
    expect(code).toBe(`"$"#,##0.00`)
  })

  it("keeps a trailing currency symbol after the number", async () => {
    // European layouts write "1 234,50 €" — the symbol element comes last.
    const code = await numFmtFor(
      `<number:currency-style style:name="N102"><number:number number:decimal-places="2" number:grouping="true"/><number:text> </number:text><number:currency-symbol>€</number:currency-symbol></number:currency-style>`,
      "N102",
    )
    expect(code).toBe(`#,##0.00 "€"`)
  })
})

describe("ODS reader — date and time data styles", () => {
  it("maps number:style=long onto the two-letter date tokens", async () => {
    const code = await numFmtFor(
      `<number:date-style style:name="N103">` +
        `<number:year number:style="long"/><number:text>-</number:text>` +
        `<number:month number:style="long"/><number:text>-</number:text>` +
        `<number:day number:style="long"/>` +
        `</number:date-style>`,
      "N103",
    )
    expect(code).toBe("yyyy-mm-dd")
  })

  it("maps the short forms onto single-letter tokens", async () => {
    const code = await numFmtFor(
      `<number:date-style style:name="N103">` +
        `<number:day/><number:text>.</number:text>` +
        `<number:month/><number:text>.</number:text>` +
        `<number:year/>` +
        `</number:date-style>`,
      "N103",
    )
    expect(code).toBe("d.m.yy")
  })

  it("turns number:textual months into mmm / mmmm", async () => {
    const long = await numFmtFor(
      `<number:date-style style:name="N103"><number:month number:textual="true" number:style="long"/></number:date-style>`,
      "N103",
    )
    const short = await numFmtFor(
      `<number:date-style style:name="N103"><number:month number:textual="true"/></number:date-style>`,
      "N103",
    )
    expect([long, short]).toEqual(["mmmm", "mmm"])
  })

  it("turns number:day-of-week into ddd / dddd", async () => {
    const long = await numFmtFor(
      `<number:date-style style:name="N103"><number:day-of-week number:style="long"/></number:date-style>`,
      "N103",
    )
    const short = await numFmtFor(
      `<number:date-style style:name="N103"><number:day-of-week/></number:date-style>`,
      "N103",
    )
    expect([long, short]).toEqual(["dddd", "ddd"])
  })

  it("rebuilds a 12-hour clock with the AM/PM marker", async () => {
    const code = await numFmtFor(
      `<number:time-style style:name="N104">` +
        `<number:hours/><number:text>:</number:text>` +
        `<number:minutes number:style="long"/><number:text>:</number:text>` +
        `<number:seconds number:style="long"/><number:text> </number:text>` +
        `<number:am-pm/>` +
        `</number:time-style>`,
      "N104",
    )
    expect(code).toBe("h:mm:ss AM/PM")
  })

  it("brackets the hour token when truncate-on-overflow is false", async () => {
    // `number:truncate-on-overflow="false"` is ODF's way of saying "elapsed
    // time" — Excel spells the same thing `[hh]:mm`.
    const code = await numFmtFor(
      `<number:time-style style:name="N104" number:truncate-on-overflow="false">` +
        `<number:hours number:style="long"/><number:text>:</number:text>` +
        `<number:minutes number:style="long"/>` +
        `</number:time-style>`,
      "N104",
    )
    expect(code).toBe("[hh]:mm")
  })

  it("leaves a clock-style time unbracketed when truncate-on-overflow is true", async () => {
    const code = await numFmtFor(
      `<number:time-style style:name="N104" number:truncate-on-overflow="true">` +
        `<number:hours number:style="long"/><number:text>:</number:text>` +
        `<number:minutes number:style="long"/>` +
        `</number:time-style>`,
      "N104",
    )
    expect(code).toBe("hh:mm")
  })

  it("ignores data-style children it has no Excel token for", async () => {
    // ODF also defines number:era, number:quarter, number:week-of-year …;
    // none has a format-code equivalent, so they are dropped rather than
    // guessed at.
    const code = await numFmtFor(
      `<number:date-style style:name="N103">` +
        `<number:era number:style="long"/><number:quarter/>` +
        `<number:year number:style="long"/>` +
        `</number:date-style>`,
      "N103",
    )
    expect(code).toBe("yyyy")
  })

  it("ignores a data style that carries no style:name", async () => {
    // Nameless styles cannot be referenced, so there is nothing to resolve.
    const styles =
      `<number:number-style><number:number number:decimal-places="3"/></number:number-style>` +
      `<style:style style:name="ce1" style:family="table-cell" style:data-style-name="N100"/>`
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="1"/></table:table-row>`,
      ),
      styles,
    )
    expect(wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt).toBeUndefined()
  })

  it("tolerates whitespace between the automatic-style elements", async () => {
    // Pretty-printed ODF (hand-edited files, some non-LibreOffice writers)
    // puts text nodes between every element.
    const styles = `
      <number:number-style style:name="N100">
        <number:number number:decimal-places="1"/>
      </number:number-style>
      <style:style style:name="ce1" style:family="table-cell" style:data-style-name="N100"/>
    `
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="1"/></table:table-row>`,
      ),
      styles,
    )
    expect(wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt).toBe("0.0")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — cell styles
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — automatic cell styles", () => {
  const cellRow = `<table:table-row><table:table-cell table:style-name="ce1" office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row>`

  it("only reads styles from the table-cell family", async () => {
    // content.xml also holds column, row and table styles under the same
    // parent; a `co1` column style must not become a cell style.
    const styles = `<style:style style:name="ce1" style:family="table-column"><style:text-properties fo:font-weight="bold"/></style:style>`
    const wb = await readBody(table(cellRow), styles)
    expect(wb.sheets[0]!.cells).toBeUndefined()
  })

  it("skips a style element with no style:name", async () => {
    const styles = `<style:style style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>`
    const wb = await readBody(table(cellRow), styles)
    expect(wb.sheets[0]!.cells).toBeUndefined()
  })

  it("accepts colours written without the leading # as well as with it", async () => {
    // ODF requires `#rrggbb`, but files in the wild omit the hash; the
    // reader normalizes both to the bare uppercase hex an XLSX Color wants.
    const styles =
      `<style:style style:name="ce1" style:family="table-cell">` +
      `<style:text-properties fo:color="ff0000" fo:font-size="10.5pt"/>` +
      `<style:table-cell-properties fo:background-color="00ff00"/>` +
      `</style:style>`
    const wb = await readBody(table(cellRow), styles)
    const style = wb.sheets[0]!.cells?.get("0,0")?.style
    expect(style?.font?.color?.rgb).toBe("FF0000")
    expect(style?.font?.size).toBe(10.5)
    expect((style!.fill as PatternFill).fgColor?.rgb).toBe("00FF00")
  })

  it("ignores a font size it cannot parse", async () => {
    const styles =
      `<style:style style:name="ce1" style:family="table-cell">` +
      `<style:text-properties fo:font-size="medium" fo:font-weight="bold"/>` +
      `</style:style>`
    const wb = await readBody(table(cellRow), styles)
    const font = wb.sheets[0]!.cells?.get("0,0")?.style?.font
    expect(font).toEqual({ bold: true })
  })

  it("drops a transparent background instead of emitting a white fill", async () => {
    const styles =
      `<style:style style:name="ce1" style:family="table-cell">` +
      `<style:text-properties fo:font-style="italic"/>` +
      `<style:table-cell-properties fo:background-color="transparent"/>` +
      `</style:style>`
    const wb = await readBody(table(cellRow), styles)
    const style = wb.sheets[0]!.cells?.get("0,0")?.style
    expect(style?.font?.italic).toBe(true)
    expect(style?.fill).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — office:value-type variants
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — value types", () => {
  async function firstCell(cellXml: string): Promise<unknown> {
    // The trailing text cell anchors the row — a cell that reads as null
    // would otherwise be trimmed off the end before the row is built.
    const wb = await readBody(
      table(
        `<table:table-row>${cellXml}` +
          `<table:table-cell office:value-type="string"><text:p>end</text:p></table:table-cell>` +
          `</table:table-row>`,
      ),
    )
    return wb.sheets[0]!.rows[0]?.[0]
  }

  it("returns an ISO 8601 duration for office:value-type=time", async () => {
    // ODF stores times as durations (`PT12H30M`); there is no Date that can
    // represent "12:30 of no particular day", so the string is preserved.
    expect(
      await firstCell(
        `<table:table-cell office:value-type="time" office:time-value="PT12H30M00S"><text:p>12:30:00</text:p></table:table-cell>`,
      ),
    ).toBe("PT12H30M00S")
  })

  it("yields null for a time cell with no office:time-value", async () => {
    expect(
      await firstCell(`<table:table-cell office:value-type="time"><text:p/></table:table-cell>`),
    ).toBeNull()
  })

  it("reads both boolean literals and rejects anything else", async () => {
    expect(
      await firstCell(
        `<table:table-cell office:value-type="boolean" office:boolean-value="true"/>`,
      ),
    ).toBe(true)
    expect(
      await firstCell(
        `<table:table-cell office:value-type="boolean" office:boolean-value="false"/>`,
      ),
    ).toBe(false)
    expect(
      await firstCell(`<table:table-cell office:value-type="boolean" office:boolean-value="1"/>`),
    ).toBeNull()
  })

  it("yields null for an unparseable office:date-value", async () => {
    expect(
      await firstCell(
        `<table:table-cell office:value-type="date" office:date-value="not-a-date"/>`,
      ),
    ).toBeNull()
    expect(await firstCell(`<table:table-cell office:value-type="date"/>`)).toBeNull()
  })

  it("falls back to office:string-value when the cell has no text:p", async () => {
    expect(
      await firstCell(
        `<table:table-cell office:value-type="string" office:string-value="from attribute"/>`,
      ),
    ).toBe("from attribute")
  })

  it("returns an empty string for a string cell with neither text nor attribute", async () => {
    // Distinguishable from an untyped empty cell, which reads as null.
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell office:value-type="string"/><table:table-cell office:value-type="float" office:value="1"/></table:table-row>`,
      ),
    )
    expect(wb.sheets[0]!.rows[0]).toEqual(["", 1])
  })

  it("yields null for a float cell with no office:value", async () => {
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell office:value-type="float"/><table:table-cell office:value-type="string"><text:p>end</text:p></table:table-cell></table:table-row>`,
      ),
    )
    expect(wb.sheets[0]!.rows[0]).toEqual([null, "end"])
  })

  it("honours calcext:value-type when office:value-type is absent", async () => {
    // LibreOffice writes the calcext mirror for error/boolean cells.
    expect(
      await firstCell(
        `<table:table-cell calcext:value-type="boolean" office:boolean-value="true"/>`,
      ),
    ).toBe(true)
  })

  it("reads an untyped cell as its plain text", async () => {
    expect(await firstCell(`<table:table-cell><text:p>bare</text:p></table:table-cell>`)).toBe(
      "bare",
    )
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — text content and hyperlinks
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — hyperlinks", () => {
  it("finds a link nested inside a text:span", async () => {
    // LibreOffice wraps a formatted link in a span, so the anchor is not a
    // direct child of the paragraph.
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell office:value-type="string">` +
          `<text:p>see <text:span text:style-name="T1"><text:a xlink:href="https://example.com/">the docs</text:a></text:span></text:p>` +
          `</table:table-cell></table:table-row>`,
      ),
    )
    const cell = wb.sheets[0]!.cells?.get("0,0")
    expect(cell?.hyperlink).toEqual({ target: "https://example.com/", display: "the docs" })
    expect(cell?.value).toBe("see the docs")
  })

  it("ignores a text:a with no xlink:href and keeps looking", async () => {
    // An anchor without a target is a bookmark, not a hyperlink.
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell office:value-type="string">` +
          `<text:p><text:a office:name="bookmark">anchor</text:a><text:a xlink:href="https://real.example/">real</text:a></text:p>` +
          `</table:table-cell></table:table-row>`,
      ),
    )
    expect(wb.sheets[0]!.cells?.get("0,0")?.hyperlink?.target).toBe("https://real.example/")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — formulas
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — formula namespace prefixes", () => {
  async function formulaOf(attr: string): Promise<string | undefined> {
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell table:formula="${attr}" office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row>`,
      ),
    )
    return wb.sheets[0]!.cells?.get("0,0")?.formula
  }

  it("strips the legacy oooc: prefix OpenOffice.org 1.x wrote", async () => {
    expect(await formulaOf("oooc:=SUM([.A1:.A10])")).toBe("SUM(A1:A10)")
  })

  it("strips a bare = prefix", async () => {
    // Gnumeric and several exporters omit the namespace prefix entirely.
    expect(await formulaOf("=SUM([.A1])")).toBe("SUM(A1)")
  })

  it("leaves an unprefixed formula untouched", async () => {
    expect(await formulaOf("SUM([.A1])")).toBe("SUM(A1)")
  })

  it("keeps a formula on a cell whose style name is not a cell style", async () => {
    // `table:style-name` may point at a column or row style that
    // parseStyles never collected; the formula must survive anyway.
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell table:style-name="co1" table:formula="of:=1" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>`,
      ),
      `<style:style style:name="co1" style:family="table-column"/>`,
    )
    const cell = wb.sheets[0]!.cells?.get("0,0")
    expect(cell?.formula).toBe("1")
    expect(cell?.style).toBeUndefined()
  })

  it("marks a formula cell as type formula", async () => {
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row>`,
      ),
    )
    expect(wb.sheets[0]!.cells?.get("0,0")?.type).toBe("formula")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — table structure
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — repeated cells, covered cells and merges", () => {
  it("expands table:number-columns-repeated in place", async () => {
    const wb = await readBody(
      table(
        `<table:table-row>` +
          `<table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="3"><text:p>7</text:p></table:table-cell>` +
          `<table:table-cell office:value-type="string"><text:p>tail</text:p></table:table-cell>` +
          `</table:table-row>`,
      ),
    )
    expect(wb.sheets[0]!.rows[0]).toEqual([7, 7, 7, "tail"])
  })

  it("clamps a column repeat that would run past Excel's last column", async () => {
    // A non-trailing repeat is expanded, so an unbounded count is an
    // allocation of 2^31 slots from a few bytes of XML. See #363.
    const wb = await readBody(
      table(
        `<table:table-row>` +
          `<table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="99999999"><text:p>1</text:p></table:table-cell>` +
          `<table:table-cell office:value-type="string"><text:p>tail</text:p></table:table-cell>` +
          `</table:table-row>`,
      ),
    )
    expect(wb.sheets[0]!.rows[0]!.length).toBe(16_384 + 1)
  })

  it("fills covered cells of a merge with nulls", async () => {
    const wb = await readBody(
      table(
        `<table:table-row>` +
          `<table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="1" office:value-type="string"><text:p>wide</text:p></table:table-cell>` +
          `<table:covered-table-cell/>` +
          `<table:table-cell office:value-type="string"><text:p>after</text:p></table:table-cell>` +
          `</table:table-row>`,
      ),
    )
    const sheet = wb.sheets[0]!
    expect(sheet.rows[0]).toEqual(["wide", null, "after"])
    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }])
  })

  it("expands a repeated covered cell run", async () => {
    const wb = await readBody(
      table(
        `<table:table-row>` +
          `<table:table-cell table:number-columns-spanned="3" office:value-type="string"><text:p>wide</text:p></table:table-cell>` +
          `<table:covered-table-cell table:number-columns-repeated="2"/>` +
          `<table:table-cell office:value-type="string"><text:p>after</text:p></table:table-cell>` +
          `</table:table-row>`,
      ),
    )
    expect(wb.sheets[0]!.rows[0]).toEqual(["wide", null, null, "after"])
  })

  it("repeats a row that carries a merge without cloning the merge range", async () => {
    // Repeated rows containing merges are unusual; the reader duplicates the
    // row data but records the merge once, from the first occurrence.
    const wb = await readBody(
      table(
        `<table:table-row table:number-rows-repeated="2">` +
          `<table:table-cell table:number-columns-spanned="2" office:value-type="string"><text:p>m</text:p></table:table-cell>` +
          `<table:covered-table-cell/>` +
          `<table:table-cell office:value-type="string"><text:p>z</text:p></table:table-cell>` +
          `</table:table-row>`,
      ),
    )
    const sheet = wb.sheets[0]!
    expect(sheet.rows.length).toBe(2)
    expect(sheet.rows[1]).toEqual(["m", null, "z"])
    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }])
  })

  it("advances the row counter across empty repeated rows", async () => {
    // LibreOffice pads the tail of a sheet with a single empty row repeated
    // a million times; those must not become a million row arrays, but the
    // rows after them must still land at the right index.
    const wb = await readBody(
      table(
        `<table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>` +
          `<table:table-row table:number-rows-repeated="5"><table:table-cell/></table:table-row>` +
          `<table:table-row><table:table-cell table:formula="of:=1" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>`,
      ),
    )
    const sheet = wb.sheets[0]!
    expect(sheet.rows.length).toBe(2)
    expect([...sheet.cells!.keys()]).toEqual(["6,0"])
  })
})

describe("ODS reader — cell metadata types", () => {
  const styles = `<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>`

  async function typeOf(cellXml: string): Promise<string | undefined> {
    const wb = await readBody(table(`<table:table-row>${cellXml}</table:table-row>`), styles)
    return wb.sheets[0]!.cells?.get("0,0")?.type
  }

  it("tags a styled boolean cell as boolean", async () => {
    expect(
      await typeOf(
        `<table:table-cell table:style-name="ce1" office:value-type="boolean" office:boolean-value="true"/>`,
      ),
    ).toBe("boolean")
  })

  it("tags a styled date cell as date", async () => {
    expect(
      await typeOf(
        `<table:table-cell table:style-name="ce1" office:value-type="date" office:date-value="2024-03-01"/>`,
      ),
    ).toBe("date")
  })

  it("tags a styled valueless cell as empty", async () => {
    // A style-only cell is kept — the trailing-null trim skips entries that
    // carry a style name.
    expect(await typeOf(`<table:table-cell table:style-name="ce1"/>`)).toBe("empty")
  })

  it("tags a styled number cell as number", async () => {
    expect(
      await typeOf(
        `<table:table-cell table:style-name="ce1" office:value-type="float" office:value="2.5"/>`,
      ),
    ).toBe("number")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — sheet selection
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — the sheets filter", () => {
  const threeTables =
    table(
      `<table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>`,
      "Alpha",
    ) +
    table(
      `<table:table-row><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell></table:table-row>`,
      "Beta",
    ) +
    table(
      `<table:table-row><table:table-cell office:value-type="string"><text:p>c</text:p></table:table-cell></table:table-row>`,
      "Gamma",
    )

  async function names(sheets: Array<number | string>): Promise<string[]> {
    const wb = await readOds(await odsFile({ content: contentXml(threeTables) }), { sheets })
    return wb.sheets.map((s) => s.name)
  }

  it("treats an empty filter array as 'every sheet'", async () => {
    expect(await names([])).toEqual(["Alpha", "Beta", "Gamma"])
  })

  it("keeps genuinely empty sheets when the filter array is empty", async () => {
    const content = contentXml(threeTables + table("", "Blank"))
    const wb = await readOds(await odsFile({ content }), { sheets: [] })
    expect(wb.sheets.map((s) => s.name)).toEqual(["Alpha", "Beta", "Gamma", "Blank"])
  })

  it("selects by sheet name", async () => {
    expect(await names(["Beta"])).toEqual(["Beta"])
  })

  it("selects by 0-based index", async () => {
    expect(await names([2])).toEqual(["Gamma"])
  })

  it("ignores a filter entry that is neither a name nor an index", async () => {
    // JS callers can pass anything; an unrecognised spec matches nothing
    // rather than throwing or matching everything.
    expect(await names([null as unknown as string, "Alpha"])).toEqual(["Alpha"])
  })

  it("keeps a selected sheet that genuinely has no rows", async () => {
    // The skipped sheets are represented by empty placeholders, so the final
    // pass has to tell "empty because skipped" from "empty in the file".
    const content = contentXml(
      table(
        `<table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>`,
        "Alpha",
      ) + table("", "Empty"),
    )
    const byName = await readOds(await odsFile({ content }), { sheets: ["Empty"] })
    expect(byName.sheets.map((s) => s.name)).toEqual(["Empty"])

    const byPredicate = await readOds(await odsFile({ content }), {
      sheets: (info) => info.name === "Empty",
    })
    expect(byPredicate.sheets.map((s) => s.name)).toEqual(["Empty"])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — document structure and failure modes
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — malformed documents", () => {
  it("returns no sheets when content.xml has no office:body", async () => {
    const content =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-content ${NS}><office:scripts/></office:document-content>`
    const wb = await readOds(await odsFile({ content }))
    expect(wb.sheets).toEqual([])
  })

  it("returns no sheets when the body holds a text document instead of a spreadsheet", async () => {
    // An .odt renamed to .ods gets this far: valid ZIP, ODF mimetype prefix,
    // a body — but no <office:spreadsheet>.
    const content =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-content ${NS}><office:body><office:text/></office:body></office:document-content>`
    const wb = await readOds(
      await odsFile({ content, mimetype: "application/vnd.oasis.opendocument.text" }),
    )
    expect(wb.sheets).toEqual([])
  })

  it("names an unnamed table by its position", async () => {
    const content = contentXml(
      `<table:table><table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row></table:table>`,
    )
    const wb = await readOds(await odsFile({ content }))
    expect(wb.sheets[0]!.name).toBe("Sheet1")
  })

  it("rejects a ZIP whose mimetype is not an OpenDocument type", async () => {
    const data = await odsFile({ content: contentXml(""), mimetype: "application/zip" })
    await expect(readOds(data)).rejects.toThrow(/Invalid ODS mimetype/)
  })

  it("rejects an OpenDocument package with no content.xml", async () => {
    const data = await odsFile({})
    await expect(readOds(data)).rejects.toThrow(ParseError)
    await expect(readOds(data)).rejects.toThrow(/missing content\.xml/)
  })

  it("rejects a file that is not a ZIP archive at all", async () => {
    // A CSV renamed to .ods — the ZIP layer owns this failure, so the error
    // stays a ZipError instead of being re-wrapped as a ParseError.
    await expect(readOds(enc.encode("Name,Age\nAda,36\n"))).rejects.toThrow(ZipError)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS reader — meta.xml
// ═══════════════════════════════════════════════════════════════════════

describe("ODS reader — document properties from meta.xml", () => {
  it("maps every meta element the reader understands", async () => {
    const meta = metaXml(`
      <dc:title>Quarterly</dc:title>
      <dc:subject>Revenue</dc:subject>
      <meta:initial-creator>Ada Lovelace</meta:initial-creator>
      <dc:description>Numbers for Q1</dc:description>
      <meta:keyword>finance</meta:keyword>
      <meta:creation-date>2024-01-02T03:04:05</meta:creation-date>
      <dc:date>2024-05-06T07:08:09</dc:date>
    `)
    const wb = await readOds(await odsFile({ content: contentXml(""), meta }))
    const props = wb.properties!
    expect(props.title).toBe("Quarterly")
    expect(props.subject).toBe("Revenue")
    expect(props.creator).toBe("Ada Lovelace")
    expect(props.description).toBe("Numbers for Q1")
    expect(props.keywords).toBe("finance")
    expect(props.created?.getFullYear()).toBe(2024)
    expect(props.modified?.getMonth()).toBe(4)
  })

  it("ignores empty elements rather than writing empty-string properties", async () => {
    // LibreOffice emits the full set of meta elements even when the user
    // never filled them in.
    const meta = metaXml(
      `<dc:title></dc:title><dc:subject></dc:subject><meta:initial-creator></meta:initial-creator>` +
        `<dc:description></dc:description><meta:keyword></meta:keyword>` +
        `<meta:creation-date></meta:creation-date><dc:date></dc:date>` +
        `<meta:editing-cycles>3</meta:editing-cycles>`,
    )
    const wb = await readOds(await odsFile({ content: contentXml(""), meta }))
    expect(wb.properties).toBeUndefined()
  })

  it("ignores dates it cannot parse", async () => {
    const meta = metaXml(
      `<meta:creation-date>whenever</meta:creation-date><dc:date>soon</dc:date><dc:title>Kept</dc:title>`,
    )
    const wb = await readOds(await odsFile({ content: contentXml(""), meta }))
    expect(wb.properties).toEqual({ title: "Kept" })
  })

  it("ignores a meta.xml with no office:meta element", async () => {
    const meta = `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NS}></office:document-meta>`
    const wb = await readOds(await odsFile({ content: contentXml(""), meta }))
    expect(wb.properties).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// streamOdsRows
// ═══════════════════════════════════════════════════════════════════════

describe("streamOdsRows — cell shapes", () => {
  it("emits nulls for covered cells, including repeated runs", async () => {
    const data = await odsFile({
      content: contentXml(
        table(
          `<table:table-row>` +
            `<table:table-cell table:number-columns-spanned="3" office:value-type="string"><text:p>wide</text:p></table:table-cell>` +
            `<table:covered-table-cell table:number-columns-repeated="2"/>` +
            `<table:table-cell office:value-type="string"><text:p>after</text:p></table:table-cell>` +
            `</table:table-row>`,
        ),
      ),
    })
    const rows = await collectStream(data)
    expect(rows[0]!.values).toEqual(["wide", null, null, "after"])
  })

  it("expands a bare text:s as a single space", async () => {
    // `<text:s/>` without text:c means exactly one space per ODF §6.1.3.
    const data = await odsFile({
      content: contentXml(
        table(
          `<table:table-row><table:table-cell office:value-type="string"><text:p>a<text:s/>b</text:p></table:table-cell></table:table-row>`,
        ),
      ),
    })
    expect((await collectStream(data))[0]!.values[0]).toBe("a b")
  })

  it("treats a non-numeric or zero text:c as one space", async () => {
    const data = await odsFile({
      content: contentXml(
        table(
          `<table:table-row><table:table-cell office:value-type="string"><text:p>a<text:s text:c="many"/>b<text:s text:c="0"/>c</text:p></table:table-cell></table:table-row>`,
        ),
      ),
    })
    expect((await collectStream(data))[0]!.values[0]).toBe("a b c")
  })

  it("resolves date, boolean and empty-string cells the way the batch reader does", async () => {
    const body = table(
      `<table:table-row>` +
        `<table:table-cell office:value-type="date" office:date-value="2024-03-01"/>` +
        `<table:table-cell office:value-type="date" office:date-value="nope"/>` +
        `<table:table-cell office:value-type="date"/>` +
        `<table:table-cell office:value-type="boolean" office:boolean-value="false"/>` +
        `<table:table-cell office:value-type="boolean" office:boolean-value="maybe"/>` +
        `<table:table-cell office:value-type="float"/>` +
        `<table:table-cell office:value-type="string"/>` +
        `<table:table-cell><text:p>untyped</text:p></table:table-cell>` +
        `</table:table-row>`,
    )
    const data = await odsFile({ content: contentXml(body) })
    const values = (await collectStream(data))[0]!.values
    expect(values[0]).toBeInstanceOf(Date)
    expect(values.slice(1)).toEqual([null, null, false, null, null, "", "untyped"])

    // …and identically through readOds.
    const wb = await readOds(data)
    expect(wb.sheets[0]!.rows[0]!.slice(1)).toEqual(values.slice(1))
  })

  // ── Known defect ─────────────────────────────────────────────────
  it("does not fold a cell annotation into the cell value", async () => {
    // src/ods/stream.ts:131 sets `inP` for *any* <text:p> seen while inside
    // a <table:table-cell>, including the paragraphs of the
    // <office:annotation> LibreOffice writes as the cell's first child for
    // a comment. The batch reader takes only direct <text:p> children
    // (src/ods/reader.ts:283), so the two disagree:
    //   readOds        → "value"
    //   streamOdsRows  → "a notevalue"
    // Any ODS with cell comments streams back with the comment text glued
    // to the front of the cell's own text.
    const body = table(
      `<table:table-row><table:table-cell office:value-type="string">` +
        `<office:annotation><dc:creator>Ada</dc:creator><text:p>a note</text:p></office:annotation>` +
        `<text:p>value</text:p>` +
        `</table:table-cell></table:table-row>`,
    )
    const data = await odsFile({ content: contentXml(body) })
    const streamed = (await collectStream(data))[0]!.values
    const batched = (await readOds(data)).sheets[0]!.rows[0]
    expect(batched).toEqual(["value"])
    expect(streamed).toEqual(batched)
  })

  it("skips empty repeated rows but keeps the row index in step", async () => {
    const body = table(
      `<table:table-row><table:table-cell office:value-type="string"><text:p>a</text:p></table:table-cell></table:table-row>` +
        `<table:table-row table:number-rows-repeated="4"><table:table-cell/></table:table-row>` +
        `<table:table-row><table:table-cell office:value-type="string"><text:p>b</text:p></table:table-cell></table:table-row>`,
    )
    const rows = await collectStream(await odsFile({ content: contentXml(body) }))
    expect(rows.map((r) => [r.index, r.values[0]])).toEqual([
      [0, "a"],
      [5, "b"],
    ])
  })
})

describe("streamOdsRows — options and failure modes", () => {
  const twoSheets = contentXml(
    table(
      `<table:table-row><table:table-cell office:value-type="string"><text:p>a1</text:p></table:table-cell></table:table-row>` +
        `<table:table-row><table:table-cell office:value-type="string"><text:p>a2</text:p></table:table-cell></table:table-row>`,
      "Alpha",
    ) +
      table(
        `<table:table-row><table:table-cell office:value-type="string"><text:p>b1</text:p></table:table-cell></table:table-row>`,
        "Beta",
      ),
  )

  it("streams every sheet when the filter names sheets it cannot resolve", async () => {
    // The SAX pass never sees table names, so a name-only filter degrades to
    // "stream everything" rather than silently yielding nothing.
    const data = await odsFile({ content: twoSheets })
    const rows: StreamRow[] = []
    for await (const row of streamOdsRows(data, { sheets: ["Beta"] })) rows.push(row)
    expect(rows.length).toBe(3)
  })

  it("restricts streaming to the requested sheet index", async () => {
    const data = await odsFile({ content: twoSheets })
    const rows: StreamRow[] = []
    for await (const row of streamOdsRows(data, { sheets: [1] })) rows.push(row)
    expect(rows.map((r) => r.values[0])).toEqual(["b1"])
  })

  it("stops after maxRows", async () => {
    const data = await odsFile({ content: twoSheets })
    const rows: StreamRow[] = []
    for await (const row of streamOdsRows(data, { maxRows: 2 })) rows.push(row)
    expect(rows.map((r) => r.values[0])).toEqual(["a1", "a2"])
  })

  it("yields nothing for a text document's content.xml", async () => {
    // An .odt renamed to .ods clears the mimetype check (both start with
    // `application/vnd.oasis.opendocument`) and reaches the SAX pass. Its
    // body holds paragraphs and a text table built from the very same
    // table:/text: elements a spreadsheet uses — none of which may be
    // mistaken for sheet rows.
    const content =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-content ${NS}><office:body><office:text>` +
      `<text:p>a paragraph<text:s text:c="2"/><text:tab/><text:line-break/>more</text:p>` +
      `<table:table table:name="T1"><table:table-row>` +
      `<table:table-cell office:value-type="string"><text:p>in a text table</text:p></table:table-cell>` +
      `<table:covered-table-cell/>` +
      `</table:table-row></table:table>` +
      `</office:text></office:body></office:document-content>`
    const data = await odsFile({
      content,
      mimetype: "application/vnd.oasis.opendocument.text",
    })
    expect(await collectStream(data)).toEqual([])
    expect((await readOds(data)).sheets).toEqual([])
  })

  it("rejects a package with no mimetype entry", async () => {
    const zip = new ZipWriter()
    zip.add("content.xml", enc.encode(contentXml("")))
    const data = await zip.build()
    await expect(collectStream(data)).rejects.toThrow(/missing 'mimetype'/)
  })

  it("rejects a package with no content.xml", async () => {
    await expect(collectStream(await odsFile({}))).rejects.toThrow(/missing content\.xml/)
  })

  it("rejects input that is not a ZIP archive", async () => {
    await expect(collectStream(enc.encode("just some bytes"))).rejects.toThrow(ZipError)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS writer — number-format translation
// ═══════════════════════════════════════════════════════════════════════
//
// Writing a numFmt and reading it back with `readStyles` exercises both
// halves of the Excel ⇄ ODF format-code translation in one pass.

describe("ODS writer — date format code round-trips", () => {
  async function roundTrip(numFmt: string): Promise<string | undefined> {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: 1, style: { numFmt } })
    const data = await writeOds({
      sheets: [{ name: "S", rows: [[1]], cells }],
    })
    const wb = await readOds(data, { readStyles: true })
    return wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt
  }

  it.each([
    ["yy-m-d", "yy-m-d"],
    ["dd.mm.yyyy", "dd.mm.yyyy"],
    ["mmm-yy", "mmm-yy"],
    ["dddd, mmmm d, yyyy", "dddd, mmmm d, yyyy"],
    ["ddd d mmm", "ddd d mmm"],
    ["h:m:s AM/PM", "h:m:s AM/PM"],
    ["hh:mm:ss", "hh:mm:ss"],
  ])("preserves %s", async (input, expected) => {
    expect(await roundTrip(input)).toBe(expected)
  })

  it("keeps a quoted literal in a date code", async () => {
    expect(await roundTrip(`yyyy" (fiscal)"`)).toBe(`yyyy" (fiscal)"`)
  })

  it("normalizes a backslash-escaped separator to the bare character", async () => {
    // Excel's `\-` and a literal `-` mean the same thing in ODF, which has
    // only <number:text>.
    expect(await roundTrip(`yyyy\\-mm`)).toBe("yyyy-mm")
  })

  it("drops formats it cannot translate rather than emitting an empty style", async () => {
    // "General" and "@" have no ODF data-style equivalent; the cell style
    // would carry nothing else, so no <style:style> is emitted at all.
    expect(await roundTrip("General")).toBeUndefined()
    expect(await roundTrip("@")).toBeUndefined()
  })

  it("shares one data style between two different cell styles", async () => {
    // The cell-style cache is keyed on the whole style, so bold+0.000 and
    // italic+0.000 are two <style:style> elements — but they must point at
    // the same <number:number-style>.
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: 1, style: { numFmt: "0.000", font: { bold: true } } })
    cells.set("0,1", { value: 2, style: { numFmt: "0.000", font: { italic: true } } })
    const xml = await contentOf(await writeOds({ sheets: [{ name: "S", rows: [[1, 2]], cells }] }))
    expect(xml.match(/<number:number-style/g)!.length).toBe(1)
    expect(xml.match(/<style:style/g)!.length).toBe(2)
  })

  it("emits nothing for a style object with no properties at all", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: 1, style: {} })
    const xml = await contentOf(await writeOds({ sheets: [{ name: "S", rows: [[1]], cells }] }))
    expect(xml).not.toContain("<style:style")
    expect(xml).not.toContain("table:style-name")
  })

  it("reads a bracketed currency symbol out of an Excel locale tag", async () => {
    // Excel writes `[$€-2]#,##0` for a Euro format.
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: 1, style: { numFmt: "[$€-2]#,##0" } })
    const data = await writeOds({ sheets: [{ name: "S", rows: [[1]], cells }] })
    expect(await contentOf(data)).toContain("<number:currency-symbol>€</number:currency-symbol>")
    const wb = await readOds(data, { readStyles: true })
    expect(wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt).toBe(`"€"#,##0`)
  })

  it("marks bracketed minute and second durations as non-truncating", async () => {
    // `[m]:ss` is elapsed minutes. ODF expresses that with
    // number:truncate-on-overflow="false" on the time style; the bracket
    // itself has no ODF spelling, so this is checked on the written XML.
    const codes = ["[m]:ss", "[ss]", "[h]:mm", "[mm]:[s]"]
    const cells = new Map<string, Partial<Cell>>()
    codes.forEach((numFmt, i) => cells.set(`${i},0`, { value: i, style: { numFmt } }))
    const xml = await contentOf(
      await writeOds({ sheets: [{ name: "S", rows: codes.map((_, i) => [i]), cells }] }),
    )
    expect(xml.match(/number:truncate-on-overflow="false"/g)!.length).toBe(codes.length)
    expect(xml).toContain(`<number:minutes/>`)
    expect(xml).toContain(`<number:minutes number:style="long"/>`)
    expect(xml).toContain(`<number:seconds/>`)
    expect(xml).toContain(`<number:seconds number:style="long"/>`)
    expect(xml).toContain(`<number:hours/>`)
  })

  it("puts a trailing currency symbol after the number element", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: 1, style: { numFmt: `#,##0.00 "€"` } })
    const data = await writeOds({ sheets: [{ name: "S", rows: [[1]], cells }] })
    const doc = parseXml(await contentOf(data))
    const autoStyles = doc.children.find(
      (c): c is Exclude<typeof c, string> =>
        typeof c !== "string" && (c.local || c.tag) === "automatic-styles",
    )!
    const currency = autoStyles.children.find(
      (c): c is Exclude<typeof c, string> =>
        typeof c !== "string" && (c.local || c.tag) === "currency-style",
    )!
    const kinds = currency.children
      .filter((c): c is Exclude<typeof c, string> => typeof c !== "string")
      .map((c) => c.local || c.tag)
    expect(kinds[kinds.length - 1]).toBe("currency-symbol")
  })

  // ── Known defect ─────────────────────────────────────────────────
  it("treats [$-409] as a locale tag, not a currency symbol", async () => {
    // src/ods/writer.ts:120 — detectCurrencySymbol() falls through to a bare
    // /[$€£¥₺₽₹]/ scan of the whole code. Excel's locale prefix `[$-409]`
    // (and `[$-F800]`, used by the built-in long-date format) contains a "$",
    // so a *date* code is classified as currency: the writer emits
    // <number:currency-style> with a "$" symbol and drops every date token.
    // Input:    numFmt "[$-409]mmm-yy"
    // Expected: a <number:date-style> round-tripping to "mmm-yy"
    // Actual:   a <number:currency-style> round-tripping to `"$"0`
    // The guard at writer.ts:115 already recognises that `[$-409]` has an
    // empty symbol group; the fallback at :120 undoes that decision.
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: 1, style: { numFmt: "[$-409]mmm-yy" } })
    const data = await writeOds({ sheets: [{ name: "S", rows: [[1]], cells }] })
    const wb = await readOds(data, { readStyles: true })
    expect(wb.sheets[0]!.cells?.get("0,0")?.style?.numFmt).toBe("mmm-yy")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS writer — object data and sparse rows
// ═══════════════════════════════════════════════════════════════════════

describe("ODS writer — columns + data", () => {
  it("omits the header row when no column declares a header", async () => {
    const sheet: WriteSheet = {
      name: "S",
      columns: [{ key: "a" }, { key: "b" }],
      data: [{ a: 1, b: 2 }],
    }
    const wb = await readOds(await writeOds({ sheets: [sheet] }))
    expect(wb.sheets[0]!.rows).toEqual([[1, 2]])
  })

  it("uses the header as the lookup key when no key is given", async () => {
    const sheet: WriteSheet = {
      name: "S",
      columns: [{ header: "Name" }, { header: "Age" }],
      data: [{ Name: "Ada", Age: 36 }],
    }
    const wb = await readOds(await writeOds({ sheets: [sheet] }))
    expect(wb.sheets[0]!.rows).toEqual([
      ["Name", "Age"],
      ["Ada", 36],
    ])
  })

  it("writes null for a key the object does not have", async () => {
    const sheet: WriteSheet = {
      name: "S",
      columns: [
        { key: "a", header: "A" },
        { key: "missing", header: "M" },
      ],
      data: [{ a: 1 }],
    }
    const wb = await readOds(await writeOds({ sheets: [sheet] }))
    expect(wb.sheets[0]!.rows[1]).toEqual([1])
  })

  it("falls back to an empty header for a column with neither key nor header", async () => {
    // A column declared only for its width still occupies a position.
    const sheet: WriteSheet = {
      name: "S",
      columns: [{ header: "A", key: "a" }, { key: "b" }, { width: 10 }],
      data: [{ a: 1, b: 2 }],
    }
    const wb = await readOds(await writeOds({ sheets: [sheet] }))
    expect(wb.sheets[0]!.rows).toEqual([
      ["A", "b", ""],
      [1, 2],
    ])
  })

  it("writes an empty sheet when neither rows nor data are given", async () => {
    const wb = await readOds(await writeOds({ sheets: [{ name: "S" }] }))
    expect(wb.sheets[0]!.rows).toEqual([])
  })

  // ── Known defect ─────────────────────────────────────────────────
  it("writes a cells override that sits on a trailing empty cell", async () => {
    // src/ods/writer.ts:629-642 computes `lastMeaningful` from the row's own
    // values and then extends it only for merge starts and covered cells —
    // never for `sheet.cells`. Any override whose column is at or past the
    // row's last non-null value is therefore never emitted: the row stops
    // short of it. The same holds for an override below the last row, since
    // writeContentXml iterates `rows` (src/ods/writer.ts:736) and sizes the
    // table from `rows` + `merges` only (src/ods/writer.ts:762-772).
    // The XLSX writer grows the grid for exactly this case
    // (src/xlsx/worksheet-writer.ts:679-692), so one WriteSheet produces two
    // different documents depending on the output format.
    // Input:    rows [[1, null, null]], cells { "0,2": { formula: "NOW()" } }
    // Expected: <table:table-cell table:formula="of:=NOW()"> in column C
    // Actual:   the row ends after column A; the override is dropped
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,2", { value: null, formula: "NOW()" })
    cells.set("2,0", { value: "z" })
    const data = await writeOds({ sheets: [{ name: "S", rows: [[1, null, null]], cells }] })

    // The column-wise override round-trips.
    const wb = await readOds(data)
    expect(wb.sheets[0]!.cells?.get("0,2")?.formula).toBe("NOW()")

    // The row-wise one is asserted on the emitted XML rather than the
    // round-trip, because the reader collapses an *interior* empty row
    // instead of preserving its position — a separate defect, filed
    // separately. The writer's half is what this fix is about.
    const xml = new TextDecoder().decode(await new ZipReader(data).extract("content.xml"))
    const emittedRows = xml.match(/<table:table-row[\s\S]*?(?:<\/table:table-row>|\/>)/g) ?? []
    expect(emittedRows).toHaveLength(3)
    expect(emittedRows[2]).toContain("<text:p>z</text:p>")
  })

  it("writes a self-closing cell for a null override with nothing else on it", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: null })
    const data = await writeOds({ sheets: [{ name: "S", rows: [[null, "b"]], cells }] })
    expect(await contentOf(data)).toContain("<table:table-cell/>")
    expect((await readOds(data)).sheets[0]!.rows[0]).toEqual([null, "b"])
  })

  it("keeps an empty cell that carries a formula", async () => {
    // The empty-run collapse must not swallow a cell that has an override.
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,1", { formula: "SUM(A1:A1)" })
    const data = await writeOds({ sheets: [{ name: "S", rows: [[1, null, 3]], cells }] })
    const xml = await contentOf(data)
    expect(xml).toContain(`table:formula="of:=SUM([.A1:.A1])"`)
    const wb = await readOds(data)
    expect(wb.sheets[0]!.cells?.get("0,1")?.formula).toBe("SUM(A1:A1)")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ODS object shorthand
// ═══════════════════════════════════════════════════════════════════════

describe("readOdsObjects / writeOdsObjects", () => {
  it("writes only the data rows when writeHeaders is false", async () => {
    const data = await writeOdsObjects([{ a: 1, b: 2 }], { writeHeaders: false })
    const wb = await readOds(data as Uint8Array)
    expect(wb.sheets[0]!.rows).toEqual([[1, 2]])
  })

  it("uses an explicit header order and fills absent keys with null", async () => {
    const data = await writeOdsObjects([{ b: 2 }], { headers: ["a", "b"], sheetName: "Data" })
    const wb = await readOds(data as Uint8Array)
    expect(wb.sheets[0]!.name).toBe("Data")
    expect(wb.sheets[0]!.rows).toEqual([
      ["a", "b"],
      [null, 2],
    ])
  })

  it("writes a header-only sheet for an empty data array", async () => {
    const data = await writeOdsObjects([])
    const wb = await readOds(data as Uint8Array)
    expect(wb.sheets[0]!.rows).toEqual([])
  })

  it("round-trips through readOdsObjects", async () => {
    const data = await writeOdsObjects([
      { Name: "Ada", Age: 36 },
      { Name: "Grace", Age: 45 },
    ])
    const result = await readOdsObjects(data as Uint8Array)
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual([
      { Name: "Ada", Age: 36 },
      { Name: "Grace", Age: 45 },
    ])
  })
})
