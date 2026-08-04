import { describe, expect, it } from "vitest"
import type { CellValue } from "../src/_types"
import { calculateColumnWidth, measureValueWidth } from "../src/xlsx/auto-width"
import { calculateRowHeight } from "../src/xlsx/auto-size"
import {
  parseAppProperties,
  parseCoreProperties,
  parseCustomProperties,
} from "../src/xlsx/doc-props-reader"
import { parseExternalLink } from "../src/xlsx/external-link-reader"
import {
  parseSlicerCache,
  parseSlicers,
  parseTimelineCache,
  parseTimelines,
} from "../src/xlsx/slicer-reader"
import { parseComments } from "../src/xlsx/comments-reader"
import { parsePersons, parseThreadedComments } from "../src/xlsx/threaded-comments-reader"
import { parseCsv, parseCsvObjects } from "../src/csv/reader"

// ── Helpers ──────────────────────────────────────────────────────────

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/** Wrap external-link body XML in the element Excel actually writes. */
function externalLinkXml(body: string): string {
  return `${XML_DECL}<externalLink xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">${body}</externalLink>`
}

/** A single-relationship `_rels/externalLink1.xml.rels` part. */
function externalLinkRels(target: string, targetMode?: string): string {
  const mode = targetMode === undefined ? "" : ` TargetMode="${targetMode}"`
  return `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_REL}/externalLinkPath" Target="${target}"${mode}/></Relationships>`
}

// ═══════════════════════════════════════════════════════════════════════
// auto-width — measuring what a cell will actually display
// ═══════════════════════════════════════════════════════════════════════

describe("measureValueWidth — double-width scripts", () => {
  // Every CJK block the width table knows about, one character each.
  // A workbook with Japanese, Korean and fullwidth Latin headers hits all
  // of these; each must count as two character units.
  it.each([
    ["CJK Unified Ideographs", "\u6F22"],
    ["CJK Extension A", "\u3400"],
    ["CJK Extension B", "\u{20000}"],
    ["CJK Compatibility Ideographs", "\uF900"],
    ["Hangul Syllables", "\uAC00"],
    ["Katakana", "\u30AB"],
    ["Hiragana", "\u3072"],
    ["CJK Symbols and Punctuation", "\u3002"],
    ["Fullwidth Forms", "\uFF21"],
    ["Halfwidth/Fullwidth currency signs", "\uFFE0"],
  ])("counts %s as two units", (_name, char) => {
    expect(measureValueWidth(char)).toBe(2)
  })

  // A code point just outside every block stays single-width, which is
  // what keeps Latin and Cyrillic columns from being over-widened.
  it("counts a non-CJK character as one unit", () => {
    expect(measureValueWidth("\u2FFF")).toBe(1)
    expect(measureValueWidth("\uFFEF")).toBe(1)
  })
})

describe("measureValueWidth — number formats", () => {
  // Negative values pick the second section, and its parentheses/minus
  // do not change the digit count the column has to hold.
  it("measures a negative value through the negative section", () => {
    expect(measureValueWidth(-1234.5, "#,##0.00;(#,##0.00)")).toBe(8)
  })

  // With no negative section the minus sign is rendered by the default
  // path and does widen the column.
  it("includes the minus sign when the format has no negative section", () => {
    expect(measureValueWidth(-1234.5, "#,##0.00")).toBe(9)
  })

  it("measures zero through the third section", () => {
    expect(measureValueWidth(0, '#,##0.00;(#,##0.00);"-"')).toBe(1)
  })

  // Decimal counting stops at the first character that is not a digit
  // placeholder, so trailing literal text does not inflate the precision.
  it("stops counting decimals at the first non-placeholder", () => {
    // "#,##0.0 TL" → one decimal, so "1,234.5" (7) plus the 3 literal
    // characters of " TL".
    expect(measureValueWidth(1234.5, '#,##0.0" TL"')).toBe(10)
  })

  it("treats a format without a decimal point as zero decimals", () => {
    expect(measureValueWidth(1234.5, "#,##0")).toBe(5)
  })

  // Literal characters that widen the cell: quoted text, escaped
  // characters and the common currency symbols.
  it.each([
    ['$#,##0.00" USD"', 13],
    ["\\$#,##0.00", 9],
    ["\u20AC#,##0.00", 9],
    ["\u00A3#,##0.00", 9],
    ["\u00A5#,##0.00", 9],
  ])("adds the literal characters of %s to the measured width", (fmt, expected) => {
    expect(measureValueWidth(1234.5, fmt)).toBe(expected)
  })

  // Percentages are measured after the ×100, including grouping.
  it("measures a grouped percentage after scaling", () => {
    // 12.3456 → "1,234.6%"
    expect(measureValueWidth(12.3456, "#,##0.0%")).toBe(8)
  })

  // Colour and locale directives are stripped before measuring.
  it("ignores bracketed directives", () => {
    expect(measureValueWidth(1234.5, "[Red][$-409]#,##0.00")).toBe(8)
  })
})

describe("calculateColumnWidth", () => {
  // Excel snaps auto-fit widths to half-character increments and clamps
  // to the 255-character maximum.
  it("snaps to the next half character", () => {
    // 9 chars × 1.1 = 9.9, + 2 padding = 11.9 → 12
    expect(calculateColumnWidth(["123456789"], { minWidth: 0 })).toBe(12)
  })

  it("widens bold columns by the bold multiplier", () => {
    const plain = calculateColumnWidth(["123456789"], { minWidth: 0 })
    const bold = calculateColumnWidth(["123456789"], { minWidth: 0, font: { bold: true } })
    expect(bold).toBeGreaterThan(plain)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// auto-size — row heights
// ═══════════════════════════════════════════════════════════════════════

describe("calculateRowHeight", () => {
  it("returns the Calibri 11pt default for a single line", () => {
    expect(calculateRowHeight(["hello"])).toBe(15)
  })

  // Line height scales with the font, rounded up to Excel's quarter-point
  // granularity.
  it("scales the line height with the font size", () => {
    expect(calculateRowHeight(["hello"], { fontSize: 22 })).toBe(30)
    expect(calculateRowHeight(["hello"], { fontSize: 10 })).toBe(13.75)
  })

  // A sparse row hands `undefined` to the calculator even though it is
  // outside the CellValue union; both it and null are skipped.
  it("ignores empty and absent cells", () => {
    const sparse = [null, undefined, ""] as unknown as CellValue[]
    expect(calculateRowHeight(sparse)).toBe(15)
  })

  // Without wrapText only explicit newlines add height.
  it("counts explicit newlines when wrapping is off", () => {
    expect(calculateRowHeight(["a\nb\nc"])).toBe(45)
  })

  // With wrapText the text is divided by the column width.
  it("wraps text against the matching column width", () => {
    expect(calculateRowHeight(["abcdefghij"], { wrapText: true, columnWidths: [4] })).toBe(45)
  })

  // A column with no measured width falls back to Excel's 8.43 default,
  // including for cells past the end of the columnWidths array.
  it("falls back to the default column width when none is given", () => {
    expect(calculateRowHeight(["x".repeat(30)], { wrapText: true })).toBe(60)
    expect(calculateRowHeight(["a", "x".repeat(30)], { wrapText: true, columnWidths: [10] })).toBe(
      60,
    )
  })

  // A zero or negative width would divide by zero; the wrap maths clamps
  // it to one character.
  it("clamps a zero column width to one character", () => {
    expect(calculateRowHeight(["abc"], { wrapText: true, columnWidths: [0] })).toBe(45)
  })

  // A blank paragraph between two lines is still a line.
  it("counts a blank line between paragraphs", () => {
    expect(calculateRowHeight(["a\n\nb"], { wrapText: true, columnWidths: [10] })).toBe(45)
  })

  // CJK text occupies two units per character when deciding where it
  // wraps, so a five-character Japanese label needs more lines than a
  // five-character Latin one.
  it("wraps double-width text at half the character count", () => {
    expect(
      calculateRowHeight(["\u6F22\u5B57\u6F22\u5B57\u6F22"], {
        wrapText: true,
        columnWidths: [4],
      }),
    ).toBe(45)
    expect(calculateRowHeight(["abcde"], { wrapText: true, columnWidths: [4] })).toBe(30)
  })

  // The wrap calculation has its own copy of the double-width table; a
  // label mixing every CJK block must be measured at two units each.
  it.each([
    ["CJK Unified Ideographs", "\u6F22"],
    ["CJK Extension A", "\u3400"],
    ["CJK Extension B", "\u{20000}"],
    ["CJK Compatibility Ideographs", "\uF900"],
    ["Hangul Syllables", "\uAC00"],
    ["Katakana", "\u30AB"],
    ["Hiragana", "\u3072"],
    ["CJK Symbols and Punctuation", "\u3002"],
    ["Fullwidth Forms", "\uFF21"],
    ["Halfwidth/Fullwidth currency signs", "\uFFE0"],
  ])("wraps %s at two units per character", (_name, char) => {
    // Two double-width characters fill a 2-unit column across 2 lines,
    // where two single-width characters would fit on one.
    expect(calculateRowHeight([char + char], { wrapText: true, columnWidths: [2] })).toBe(30)
    expect(calculateRowHeight(["ab"], { wrapText: true, columnWidths: [2] })).toBe(15)
  })

  it("takes the tallest cell in the row", () => {
    expect(calculateRowHeight(["a", "b\nc\nd", "e"])).toBe(45)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// docProps — core.xml, custom.xml, app.xml
// ═══════════════════════════════════════════════════════════════════════

describe("parseCoreProperties", () => {
  // Empty elements are what Excel writes for properties the user cleared;
  // they must not turn into empty-string properties.
  it("skips elements with no text", () => {
    const xml = `${XML_DECL}<cp:coreProperties xmlns:cp="x" xmlns:dc="y" xmlns:dcterms="z"><dc:title></dc:title><dc:subject/><dc:creator/><cp:keywords/><dc:description/><cp:lastModifiedBy/><cp:category/><dcterms:created/><dcterms:modified/></cp:coreProperties>`
    expect(parseCoreProperties(xml)).toEqual({})
  })

  // A date Excel cannot round-trip (or a third-party tool's garbage) is
  // dropped rather than surfacing an Invalid Date.
  it("drops a created/modified stamp that is not a valid date", () => {
    const xml = `${XML_DECL}<cp:coreProperties xmlns:cp="x" xmlns:dcterms="z"><dcterms:created>not-a-date</dcterms:created><dcterms:modified>whenever</dcterms:modified></cp:coreProperties>`
    const props = parseCoreProperties(xml)
    expect(props.created).toBeUndefined()
    expect(props.modified).toBeUndefined()
  })

  it("ignores unknown child elements", () => {
    const xml = `${XML_DECL}<cp:coreProperties xmlns:cp="x" xmlns:dc="y"><cp:contentStatus>Draft</cp:contentStatus><dc:title>T</dc:title></cp:coreProperties>`
    expect(parseCoreProperties(xml)).toEqual({ title: "T" })
  })
})

describe("parseCustomProperties", () => {
  // A real docProps/custom.xml: one property per supported variant type
  // from the docPropsVTypes schema (ECMA-376 Part 4, §7.4).
  const CUSTOM_XML = `${XML_DECL}
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Department"><vt:lpwstr>Finance</vt:lpwstr></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Revision"><vt:i4>7</vt:i4></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Rows"><vt:i8>9007199254740</vt:i8></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Retries"><vt:int>3</vt:int></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="Ratio"><vt:r8>1.5</vt:r8></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Total"><vt:decimal>12.75</vt:decimal></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="8" name="Approved"><vt:bool>true</vt:bool></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="9" name="Archived"><vt:bool>1</vt:bool></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="10" name="Draft"><vt:bool>0</vt:bool></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="11" name="Deadline"><vt:filetime>2026-01-15T10:00:00Z</vt:filetime></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="12" name="Signed"><vt:date>2026-02-01T00:00:00Z</vt:date></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="13" name="Empty"><vt:lpwstr/></property>
</Properties>`

  it("reads every supported variant type", () => {
    expect(parseCustomProperties(CUSTOM_XML)).toEqual({
      Department: "Finance",
      Revision: 7,
      Rows: 9007199254740,
      Retries: 3,
      Ratio: 1.5,
      Total: 12.75,
      Approved: true,
      Archived: true,
      Draft: false,
      Deadline: new Date("2026-01-15T10:00:00Z"),
      Signed: new Date("2026-02-01T00:00:00Z"),
      Empty: "",
    })
  })

  // Properties without a name have nothing to key on; elements that are
  // not <property> at all (schema extensions) are skipped.
  it("skips unnamed properties and foreign elements", () => {
    const xml = `${XML_DECL}<Properties xmlns="c" xmlns:vt="v"><property pid="2"><vt:lpwstr>x</vt:lpwstr></property><ext/><property pid="3" name="Kept"><vt:lpwstr>y</vt:lpwstr></property></Properties>`
    expect(parseCustomProperties(xml)).toEqual({ Kept: "y" })
  })

  // Numeric and date variants with unusable text are dropped rather than
  // stored as NaN or Invalid Date.
  it("drops numeric and date variants that cannot be parsed", () => {
    const xml = `${XML_DECL}<Properties xmlns="c" xmlns:vt="v">
      <property pid="2" name="NoNumber"><vt:i4/></property>
      <property pid="3" name="NoFloat"><vt:r8/></property>
      <property pid="4" name="NoDate"><vt:filetime/></property>
      <property pid="5" name="BadDate"><vt:filetime>whenever</vt:filetime></property>
      <property pid="6" name="Unsupported"><vt:cy>1.0000</vt:cy></property>
    </Properties>`
    expect(parseCustomProperties(xml)).toEqual({})
  })

  // Only the first variant child counts — a malformed property carrying
  // two values must not have the second overwrite the first.
  // Pretty-printed custom.xml puts the variant element on its own line,
  // so the whitespace text nodes around it must be skipped.
  it("skips the whitespace of a pretty-printed property", () => {
    const xml = `${XML_DECL}<Properties xmlns="c" xmlns:vt="v">
  <property pid="2" name="Department">
    <vt:lpwstr>Finance</vt:lpwstr>
  </property>
</Properties>`
    expect(parseCustomProperties(xml)).toEqual({ Department: "Finance" })
  })

  it("takes only the first value element of a property", () => {
    const xml = `${XML_DECL}<Properties xmlns="c" xmlns:vt="v"><property pid="2" name="P"><vt:lpwstr>first</vt:lpwstr><vt:lpwstr>second</vt:lpwstr></property></Properties>`
    expect(parseCustomProperties(xml)).toEqual({ P: "first" })
  })
})

describe("parseAppProperties", () => {
  // app.xml always carries Application/DocSecurity/TitlesOfParts; only
  // Company and Manager map onto workbook properties.
  it("returns nothing when neither Company nor Manager is present", () => {
    const xml = `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity></Properties>`
    expect(parseAppProperties(xml)).toEqual({})
  })

  it("ignores Company and Manager elements that are empty", () => {
    const xml = `${XML_DECL}<Properties xmlns="e"><Company/><Manager></Manager></Properties>`
    expect(parseAppProperties(xml)).toEqual({})
  })
})

// ═══════════════════════════════════════════════════════════════════════
// externalLink1.xml — linked workbooks and their cached values
// ═══════════════════════════════════════════════════════════════════════

describe("parseExternalLink", () => {
  const FULL = externalLinkXml(`
    <externalBook r:id="rId1">
      <sheetNames><sheetName val="Sheet1"/><sheetName val="Budget"/><sheetName/></sheetNames>
      <definedNames>
        <definedName name="Rate" refersTo="'[1]Sheet1'!$A$1" sheetId="0"/>
        <definedName name="Bare"/>
        <definedName refersTo="'[1]Sheet1'!$B$1"/>
        <definedName name="BadSheet" sheetId="zero"/>
      </definedNames>
      <sheetDataSet>
        <sheetData sheetId="0">
          <row r="1">
            <cell r="A1"><v>42</v></cell>
            <cell r="B1" t="str"><v>Hello</v></cell>
            <cell r="C1" t="b"><v>1</v></cell>
            <cell r="D1" t="e"><v>#REF!</v></cell>
            <cell r="E1" t="s"><v>3</v></cell>
          </row>
          <row r="2">
            <cell r="A2"><v>not-a-number</v></cell>
            <cell r="B2" t="b"><v>true</v></cell>
            <cell r="C2" t="b"><v>0</v></cell>
            <cell r="D2"/>
          </row>
        </sheetData>
        <sheetData sheetId="notanumber"><row r="1"><cell r="A1"><v>1</v></cell></row></sheetData>
        <sheetData><row r="1"><cell r="A1"><v>1</v></cell></row></sheetData>
      </sheetDataSet>
    </externalBook>`)

  it("resolves the linked workbook path through the rels part", () => {
    const link = parseExternalLink(FULL, externalLinkRels("C:\\Budgets\\2026.xlsx", "External"))
    expect(link.target).toBe("C:\\Budgets\\2026.xlsx")
    expect(link.targetMode).toBe("External")
  })

  // A link with no rels part (or one whose r:id does not resolve) still
  // parses — only the target is unknown.
  it("leaves the target empty when no rels part is supplied", () => {
    const link = parseExternalLink(FULL)
    expect(link.target).toBe("")
    expect(link.targetMode).toBeUndefined()
    expect(link.sheetNames).toEqual(["Sheet1", "Budget", ""])
  })

  it("leaves the target empty when the relationship id does not resolve", () => {
    const link = parseExternalLink(FULL, externalLinkRels("other.xlsx", "External"))
    const missing = parseExternalLink(
      externalLinkXml('<externalBook r:id="rIdMissing"><sheetNames/></externalBook>'),
      externalLinkRels("other.xlsx", "External"),
    )
    expect(link.target).toBe("other.xlsx")
    expect(missing.target).toBe("")
  })

  // The strict-transitional variants of the schema use an unprefixed
  // `id`, so the reader accepts both spellings.
  it("accepts an unprefixed relationship id", () => {
    const xml = externalLinkXml('<externalBook id="rId1"><sheetNames/></externalBook>')
    expect(parseExternalLink(xml, externalLinkRels("book.xlsx")).target).toBe("book.xlsx")
  })

  // An externalBook with no r:id at all cannot be resolved, even when a
  // rels part is present — the reader never guesses at the sole entry.
  it("skips resolution when the externalBook carries no relationship id", () => {
    const xml = externalLinkXml("<externalBook><sheetNames/></externalBook>")
    expect(parseExternalLink(xml, externalLinkRels("book.xlsx")).target).toBe("")
  })

  // TargetMode is only reported for the two values the schema defines.
  it.each([
    ["External", "External"],
    ["Internal", "Internal"],
  ])("reports the %s target mode", (mode, expected) => {
    const xml = externalLinkXml('<externalBook r:id="rId1"><sheetNames/></externalBook>')
    expect(parseExternalLink(xml, externalLinkRels("b.xlsx", mode)).targetMode).toBe(expected)
  })

  it("omits an unrecognised or absent target mode", () => {
    const xml = externalLinkXml('<externalBook r:id="rId1"><sheetNames/></externalBook>')
    expect(parseExternalLink(xml, externalLinkRels("b.xlsx")).targetMode).toBeUndefined()
    expect(parseExternalLink(xml, externalLinkRels("b.xlsx", "Odd")).targetMode).toBeUndefined()
  })

  // Cached values keep the type the external workbook recorded; `s`
  // (shared-string index) stays numeric because the string table lives in
  // the *other* workbook.
  it("coerces cached cell values by their recorded type", () => {
    const link = parseExternalLink(FULL)
    expect(link.sheetData[0].cells).toEqual([
      { ref: "A1", type: "n", value: 42 },
      { ref: "B1", type: "str", value: "Hello" },
      { ref: "C1", type: "b", value: true },
      { ref: "D1", type: "e", value: "#REF!" },
      { ref: "E1", type: "s", value: 3 },
      { ref: "A2", type: "n", value: 0 },
      { ref: "B2", type: "b", value: true },
      { ref: "C2", type: "b", value: false },
      { ref: "D2", type: "n", value: 0 },
    ])
  })

  it("falls back to sheetId 0 when the attribute is missing or unparseable", () => {
    const link = parseExternalLink(FULL)
    expect(link.sheetData.map((s) => s.sheetId)).toEqual([0, 0, 0])
  })

  // Defined names need a name; the sheetId is optional and is dropped
  // when it is not a number.
  it("keeps only named defined names and numeric sheet ids", () => {
    const link = parseExternalLink(FULL)
    expect(link.definedNames).toEqual([
      { name: "Rate", refersTo: "'[1]Sheet1'!$A$1", sheetId: 0 },
      { name: "Bare" },
      { name: "BadSheet" },
    ])
  })

  // A link file that carries no externalBook (or an unexpected root)
  // degrades to an empty link rather than throwing.
  it("returns empty collections when the externalBook is missing", () => {
    const link = parseExternalLink(externalLinkXml("<extLst/>"))
    expect(link).toEqual({ target: "", sheetNames: [], sheetData: [] })
  })

  it("returns empty collections when the sections are absent", () => {
    const link = parseExternalLink(externalLinkXml("<externalBook/>"))
    expect(link.sheetNames).toEqual([])
    expect(link.sheetData).toEqual([])
    expect(link.definedNames).toBeUndefined()
  })

  // Unexpected element names inside each section are ignored so a file
  // carrying extension elements still parses.
  it("ignores foreign elements inside each section", () => {
    const link = parseExternalLink(
      externalLinkXml(`<externalBook r:id="rId1">
        <sheetNames><ext/><sheetName val="S"/></sheetNames>
        <sheetDataSet><ext/><sheetData sheetId="1"><ext/><row r="1"><ext/><cell r="A1"><v>1</v></cell><cell><v>2</v></cell></row></sheetData></sheetDataSet>
        <definedNames><ext/><definedName name="N"/></definedNames>
      </externalBook>`),
    )
    expect(link.sheetNames).toEqual(["S"])
    expect(link.sheetData).toEqual([{ sheetId: 1, cells: [{ ref: "A1", type: "n", value: 1 }] }])
    expect(link.definedNames).toEqual([{ name: "N" }])
  })

  // An unknown `t` value falls back to the numeric default rather than
  // producing a cell with a type the ExternalCellType union forbids.
  it("falls back to the numeric type for an unrecognised cell type", () => {
    const link = parseExternalLink(
      externalLinkXml(
        '<externalBook><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1" t="inlineStr"><v>7</v></cell></row></sheetData></sheetDataSet></externalBook>',
      ),
    )
    expect(link.sheetData[0].cells).toEqual([{ ref: "A1", type: "n", value: 7 }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Slicers and timelines
// ═══════════════════════════════════════════════════════════════════════

describe("parseSlicers", () => {
  it("reads every optional slicer attribute", () => {
    const xml = `${XML_DECL}<slicers xmlns="x14"><slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="2" style="SlicerStyleLight1" sortOrder="descending" rowHeight="234950"/></slicers>`
    expect(parseSlicers(xml)).toEqual([
      {
        name: "Region",
        cache: "Slicer_Region",
        caption: "Region",
        columnCount: 2,
        style: "SlicerStyleLight1",
        sortOrder: "descending",
        rowHeight: 234950,
      },
    ])
  })

  // name and cache are the only required attributes; a slicer missing
  // either cannot be tied back to its cache, so it is skipped.
  it("skips slicers without a name or cache and foreign elements", () => {
    const xml = `${XML_DECL}<slicers xmlns="x14"><extLst/><slicer cache="c"/><slicer name="n"/><slicer name="ok" cache="c"/></slicers>`
    expect(parseSlicers(xml)).toEqual([{ name: "ok", cache: "c" }])
  })

  // Non-numeric counts are dropped rather than stored as NaN.
  it("drops unparseable numeric attributes", () => {
    const xml = `${XML_DECL}<slicers xmlns="x14"><slicer name="n" cache="c" columnCount="many" rowHeight="tall"/></slicers>`
    expect(parseSlicers(xml)).toEqual([{ name: "n", cache: "c" }])
  })
})

describe("parseSlicerCache", () => {
  it("reads the pivot tables a cache drives", () => {
    const xml = `${XML_DECL}<slicerCacheDefinition xmlns="x14" name="Slicer_Region" sourceName="Region"><pivotTables><pivotTable tabId="1" name="PivotTable1"/><pivotTable name="NoTab"/><pivotTable tabId="2"/><ext/></pivotTables></slicerCacheDefinition>`
    expect(parseSlicerCache(xml)).toEqual({
      name: "Slicer_Region",
      sourceName: "Region",
      pivotTables: [{ tabId: 1, name: "PivotTable1" }],
    })
  })

  // Table slicers (Excel 2013+) record their source in an x15 extension
  // rather than in <pivotTables>.
  it("reads a table slicer's source from the x15 extension", () => {
    const xml = `${XML_DECL}<slicerCacheDefinition xmlns="x14" xmlns:x15="x15" name="Slicer_Name"><extLst><ext uri="{2F2917AC}"><x15:tableSlicerCache tableId="1" column="3"/></ext></extLst></slicerCacheDefinition>`
    expect(parseSlicerCache(xml)).toEqual({
      name: "Slicer_Name",
      tableSource: { name: "1", column: "3" },
    })
  })

  it("prefers an explicit table name over the table id", () => {
    const xml = `${XML_DECL}<slicerCacheDefinition xmlns="x14" xmlns:x15="x15" name="S"><extLst><ext><x15:tableSlicerCache name="Table1" tableId="1"/></ext></extLst></slicerCacheDefinition>`
    expect(parseSlicerCache(xml)?.tableSource).toEqual({ name: "Table1" })
  })

  it("ignores extensions that carry no table slicer cache", () => {
    const xml = `${XML_DECL}<slicerCacheDefinition xmlns="x14" xmlns:x15="x15" name="S"><extLst><notExt/><ext uri="{other}"><x15:somethingElse/></ext><ext><x15:tableSlicerCache/></ext></extLst></slicerCacheDefinition>`
    expect(parseSlicerCache(xml)?.tableSource).toBeUndefined()
  })

  // A cache without a name cannot be referenced by a slicer.
  it("returns undefined for a cache with no name or no definition", () => {
    expect(parseSlicerCache(`${XML_DECL}<slicerCacheDefinition xmlns="x14"/>`)).toBeUndefined()
    expect(parseSlicerCache(`${XML_DECL}<somethingElse/>`)).toBeUndefined()
  })

  // Excel wraps the definition in a container element in some files, so
  // the reader looks one level down as well.
  it("finds the definition nested inside a wrapper element", () => {
    const xml = `${XML_DECL}<wrapper><slicerCacheDefinition name="S"/></wrapper>`
    expect(parseSlicerCache(xml)).toEqual({ name: "S" })
  })
})

describe("parseTimelines", () => {
  it("reads the boolean display flags", () => {
    const xml = `${XML_DECL}<timelines xmlns="x15"><timeline name="OrderDate" cache="NativeTimeline_OrderDate" caption="Order Date" style="TimeSlicerStyleLight1" level="2" showHeader="1" showSelectionLabel="0" showTimeLevel="true" showHorizontalScrollbar="false"/></timelines>`
    expect(parseTimelines(xml)).toEqual([
      {
        name: "OrderDate",
        cache: "NativeTimeline_OrderDate",
        caption: "Order Date",
        style: "TimeSlicerStyleLight1",
        level: "2",
        showHeader: true,
        showSelectionLabel: false,
        showTimeLevel: true,
        showHorizontalScrollbar: false,
      },
    ])
  })

  it("skips timelines without a name or cache and foreign elements", () => {
    const xml = `${XML_DECL}<timelines xmlns="x15"><extLst/><timeline cache="c"/><timeline name="n"/><timeline name="ok" cache="c"/></timelines>`
    expect(parseTimelines(xml)).toEqual([{ name: "ok", cache: "c" }])
  })
})

describe("parseTimelineCache", () => {
  it("reads the cache name, source and pivot tables", () => {
    const xml = `${XML_DECL}<timelineCacheDefinition xmlns="x15" name="NativeTimeline_OrderDate" sourceName="OrderDate"><pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables></timelineCacheDefinition>`
    expect(parseTimelineCache(xml)).toEqual({
      name: "NativeTimeline_OrderDate",
      sourceName: "OrderDate",
      pivotTables: [{ tabId: 1, name: "PivotTable1" }],
    })
  })

  it("returns undefined without a definition or a name", () => {
    expect(parseTimelineCache(`${XML_DECL}<somethingElse/>`)).toBeUndefined()
    expect(parseTimelineCache(`${XML_DECL}<timelineCacheDefinition xmlns="x15"/>`)).toBeUndefined()
  })

  it("finds the definition nested inside a wrapper element", () => {
    const xml = `${XML_DECL}<wrapper><timelineCacheDefinition name="T"/></wrapper>`
    expect(parseTimelineCache(xml)).toEqual({ name: "T" })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// commentsN.xml — the classic (VML-anchored) comment system
// ═══════════════════════════════════════════════════════════════════════

describe("parseComments", () => {
  // Excel splits comment text into runs so the author name can be bold;
  // the reader concatenates every <t> in the <text> element.
  it("concatenates the text runs of a comment", () => {
    const xml = `${XML_DECL}<comments xmlns="${NS_MAIN}"><authors><author>Ada Lovelace</author></authors><commentList><comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t>Ada Lovelace:</t></r><r><t xml:space="preserve"> check this</t></r></text></comment></commentList></comments>`
    expect(parseComments(xml).get("A1")).toEqual({
      author: "Ada Lovelace",
      text: "Ada Lovelace: check this",
    })
  })

  // Some writers emit the whole spreadsheetml package with an `x:` prefix.
  it("handles prefixed element names", () => {
    const xml = `${XML_DECL}<x:comments xmlns:x="${NS_MAIN}"><x:authors><x:author>Grace</x:author></x:authors><x:commentList><x:comment ref="B2" authorId="0"><x:text><x:t>Hi</x:t></x:text></x:comment></x:commentList></x:comments>`
    expect(parseComments(xml).get("B2")).toEqual({ author: "Grace", text: "Hi" })
  })

  // An authorId that is out of range, absent, or points at an empty
  // author entry leaves the comment unattributed rather than crashing.
  it("leaves a comment unattributed when the author cannot be resolved", () => {
    const xml = `${XML_DECL}<comments xmlns="${NS_MAIN}"><authors><author>Ada</author><author></author></authors><commentList><comment ref="A1" authorId="9"><text><t>out of range</t></text></comment><comment ref="A2"><text><t>no id</t></text></comment><comment ref="A3" authorId="1"><text><t>blank author</t></text></comment></commentList></comments>`
    const comments = parseComments(xml)
    expect(comments.get("A1")).toEqual({ text: "out of range" })
    expect(comments.get("A2")).toEqual({ text: "no id" })
    expect(comments.get("A3")).toEqual({ text: "blank author" })
  })

  // A comment with no ref has no cell to attach to.
  it("skips a comment with no cell reference", () => {
    const xml = `${XML_DECL}<comments xmlns="${NS_MAIN}"><authors><author>Ada</author></authors><commentList><comment authorId="0"><text><t>orphan</t></text></comment></commentList></comments>`
    expect(parseComments(xml).size).toBe(0)
  })

  // Every piece of state is scoped to its container: an <author>,
  // <comment>, <text>, <r> or <t> that appears outside the element it
  // belongs to (a malformed file, or an element of the same name from
  // another schema) must be ignored rather than polluting the result.
  it("ignores comment-shaped elements outside their container", () => {
    const xml = `${XML_DECL}<comments xmlns="${NS_MAIN}"><author>stray author</author><comment ref="Z9"><text><t>stray comment</t></text></comment><r><t>stray run</t></r><t>stray</t><authors><author>Ada</author></authors><commentList><comment ref="A1" authorId="0"><text><t>real</t></text></comment></commentList></comments>`
    const comments = parseComments(xml)
    expect(comments.size).toBe(1)
    expect(comments.get("A1")).toEqual({ author: "Ada", text: "real" })
  })

  it("returns an empty map for a file with no comments", () => {
    const xml = `${XML_DECL}<comments xmlns="${NS_MAIN}"><authors/><commentList/></comments>`
    expect(parseComments(xml).size).toBe(0)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Threaded comments (Excel 365) — persons + threads
// ═══════════════════════════════════════════════════════════════════════

describe("parsePersons", () => {
  it("reads the optional identity attributes", () => {
    const xml = `${XML_DECL}<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="Ada Lovelace" id="{6EE7...}" userId="ada@example.com" providerId="AD"/></personList>`
    expect(parsePersons(xml)).toEqual([
      {
        id: "{6EE7...}",
        displayName: "Ada Lovelace",
        userId: "ada@example.com",
        providerId: "AD",
      },
    ])
  })

  // A person without an id cannot be referenced by a comment; a person
  // without a displayName is not a person entry at all.
  it("skips entries missing an id or display name, and foreign elements", () => {
    const xml = `${XML_DECL}<personList xmlns="tc"><ext/><person displayName="No Id"/><person id="{1}"/><person id="{2}" displayName=""/></personList>`
    expect(parsePersons(xml)).toEqual([{ id: "{2}", displayName: "" }])
  })
})

describe("parseThreadedComments", () => {
  const THREAD = `${XML_DECL}
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <ext/>
  <threadedComment ref="A1" dT="2026-01-15T10:00:00.00" personId="{P1}" id="{C1}" done="1">
    <text>Can you check @Ada Lovelace ?</text>
    <mentions>
      <ext/>
      <mention mentionpersonId="{P2}" mentionId="{M1}" startIndex="13" length="12"/>
      <mention mentionId="{M2}" startIndex="0" length="1"/>
      <mention mentionpersonId="{P3}" startIndex="0" length="1"/>
      <mention mentionpersonId="{P4}" mentionId="{M3}"/>
      <mention mentionpersonId="{P5}" mentionId="{M4}" startIndex="first" length="lots"/>
    </mentions>
  </threadedComment>
  <threadedComment personId="{P2}" id="{C2}" parentId="{C1}" done="true"><text>Checked</text></threadedComment>
  <threadedComment personId="{P2}" id="{C3}"/>
  <threadedComment id="{C4}"><text>no person</text></threadedComment>
  <threadedComment personId="{P2}"><text>no id</text></threadedComment>
</ThreadedComments>`

  it("reads a thread with its mentions", () => {
    expect(parseThreadedComments(THREAD)[0]).toEqual({
      id: "{C1}",
      personId: "{P1}",
      ref: "A1",
      date: "2026-01-15T10:00:00.00",
      done: true,
      text: "Can you check @Ada Lovelace ?",
      // A mention needs both ids; its offsets default to 0 when they are
      // absent or unparseable.
      mentions: [
        { mentionPersonId: "{P2}", mentionId: "{M1}", startIndex: 13, length: 12 },
        { mentionPersonId: "{P4}", mentionId: "{M3}", startIndex: 0, length: 0 },
        { mentionPersonId: "{P5}", mentionId: "{M4}", startIndex: 0, length: 0 },
      ],
    })
  })

  // A reply carries parentId and no ref; `done` accepts both spellings
  // Excel writes.
  it("reads a reply and accepts done written as 'true'", () => {
    expect(parseThreadedComments(THREAD)[1]).toEqual({
      id: "{C2}",
      personId: "{P2}",
      parentId: "{C1}",
      done: true,
      text: "Checked",
    })
  })

  // A comment element with no <text> child yields an empty body rather
  // than undefined.
  it("yields empty text when the text element is missing", () => {
    expect(parseThreadedComments(THREAD)[2]).toEqual({ id: "{C3}", personId: "{P2}", text: "" })
  })

  // id and personId are both required to place a comment in a thread.
  it("skips comments missing an id or a person id", () => {
    expect(parseThreadedComments(THREAD)).toHaveLength(3)
  })

  it("omits mentions entirely when none survive validation", () => {
    const xml = `${XML_DECL}<ThreadedComments xmlns="tc"><threadedComment id="{C}" personId="{P}"><text>x</text><mentions><mention/></mentions></threadedComment></ThreadedComments>`
    expect(parseThreadedComments(xml)[0].mentions).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// CSV reader — option interactions the other suites do not reach
// ═══════════════════════════════════════════════════════════════════════

describe("parseCsv — skipLines", () => {
  // Files exported from reporting tools often carry a title block above
  // the header; skipLines must count a CRLF pair as one line, not two.
  it("counts a CRLF pair as a single line", () => {
    expect(parseCsv("title\r\nsub\r\na,b\r\n1,2", { skipLines: 2 })).toEqual([
      ["a", "b"],
      ["1", "2"],
    ])
  })

  it("counts a bare CR as a line", () => {
    expect(parseCsv("title\rsub\ra,b", { skipLines: 2 })).toEqual([["a", "b"]])
  })

  it("returns nothing when every line is skipped", () => {
    expect(parseCsv("a,b\nc,d", { skipLines: 10 })).toEqual([])
  })
})

describe("parseCsv — transformValue", () => {
  // With `header: true` the callback receives the header cell's text as
  // the column name so a caller can switch on the column.
  it("passes the header text as the column name", () => {
    const seen: string[] = []
    parseCsv("name,age\nAda,36", {
      header: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["name", "age", "name", "age"])
  })

  // Without headers the column index stands in for the name.
  it("passes the column index when there is no header row", () => {
    const seen: string[] = []
    parseCsv("Ada,36", {
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["0", "1"])
  })

  // A ragged row reaching past the header falls back to the index.
  it("falls back to the column index past the end of the header row", () => {
    const seen: string[] = []
    parseCsv("name\nAda,36", {
      header: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["name", "name", "1"])
  })
})

describe("parseCsvObjects", () => {
  // The options argument is optional even though the type suggests a
  // `header: true` flag — the first row is always the header here.
  it("works with no options at all", () => {
    expect(parseCsvObjects("name,age\nAda,36")).toEqual({
      headers: ["name", "age"],
      data: [{ name: "Ada", age: "36" }],
    })
  })
})

describe("parseCsv — delimiter detection", () => {
  // Two candidates appearing equally often on every line tie; the first
  // candidate in the probe order (comma) wins, so a file mixing commas
  // and semicolons is still read as comma-separated.
  it("keeps the earlier candidate when two delimiters tie", () => {
    expect(parseCsv("a,b;c\n1,2;3")).toEqual([
      ["a", "b;c"],
      ["1", "2;3"],
    ])
  })
})

describe("parseCsv — type inference", () => {
  // A field shaped like an ISO date but naming a day that does not exist
  // stays a string rather than becoming an Invalid Date.
  it("leaves an impossible ISO-shaped date as text", () => {
    expect(parseCsv("2021-13-45,2021-01-15", { typeInference: true })).toEqual([
      ["2021-13-45", new Date("2021-01-15")],
    ])
  })

  // A lone sign is not a number; neither is a value that overflows to
  // Infinity.
  it("leaves non-numeric sign characters and overflows as text", () => {
    expect(parseCsv("-,+,1e999,1e308", { typeInference: true })).toEqual([
      ["-", "+", "1e999", 1e308],
    ])
  })

  it("leaves an all-whitespace field as text", () => {
    expect(parseCsv('" ",1', { typeInference: true })).toEqual([[" ", 1]])
  })

  // Thousands separators are stripped only when they group three digits.
  it("parses grouped numbers but not arbitrary comma text", () => {
    expect(parseCsv('"1,234.56";"1,23"', { typeInference: true, delimiter: ";" })).toEqual([
      [1234.56, "1,23"],
    ])
  })
})

describe("parseCsv — comments", () => {
  // A quoted leading # is data, not a comment — this is what keeps a
  // column of hashtags intact.
  it("keeps a quoted field that starts with the comment character", () => {
    expect(parseCsv('#skip me\n"#hashtag",1', { comment: "#" })).toEqual([["#hashtag", "1"]])
  })
})
