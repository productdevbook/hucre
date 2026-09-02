import { describe, expect, it } from "vitest"
import { parseWorksheet } from "../src/xlsx/worksheet"
import type { WorksheetContext } from "../src/xlsx/worksheet"
import { parseStyles } from "../src/xlsx/styles"
import type { SharedString } from "../src/xlsx/shared-strings"

// ── Helpers ──────────────────────────────────────────────────────────

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

function ctx(over: Partial<WorksheetContext> = {}): WorksheetContext {
  return { sharedStrings: [], styles: null, readStyles: false, dateSystem: "1900", ...over }
}

/** Wrap a worksheet body in the `<worksheet>` root Excel writes. */
function sheet(body: string, over?: Partial<WorksheetContext>) {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${NS} ${R}>${body}</worksheet>`
  return parseWorksheet(xml, "Sheet1", ctx(over))
}

/** Wrap rows in `<sheetData>`. */
function data(rowsXml: string, over?: Partial<WorksheetContext>) {
  return sheet(`<sheetData>${rowsXml}</sheetData>`, over)
}

const ss = (texts: string[]): SharedString[] => texts.map((text) => ({ text }))

// ═══════════════════════════════════════════════════════════════════════
// Cell references
//
// `parseCellRef` accepts more spellings than Excel itself writes, because
// the reader is fed by every generator on the internet, not just Excel.
// ═══════════════════════════════════════════════════════════════════════

describe("cell reference spellings", () => {
  it("accepts lower-case column letters", () => {
    // Excel always upper-cases the column part, but hand-rolled writers
    // and some XSLT pipelines emit `<c r="b2">`. The reader treats a-z
    // as the same column alphabet rather than dropping the cell.
    const s = data(`<row r="2"><c r="b2" t="inlineStr"><is><t>lower</t></is></c></row>`)
    expect(s.rows[1][1]).toBe("lower")
  })

  it("accepts a mixed-case multi-letter column", () => {
    const s = data(`<row r="1"><c r="aA1" t="inlineStr"><is><t>x</t></is></c></row>`)
    // aA == 26 + 1 - 1 == 26 (0-based)
    expect(s.rows[0][26]).toBe("x")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Implicit positioning
//
// The `r` attribute on `<row>` and `<c>` is optional in the schema.
// Files that omit it rely purely on document order.
// ═══════════════════════════════════════════════════════════════════════

describe("rows and cells without an r attribute", () => {
  it("numbers rows sequentially when <row> has no r", () => {
    const s = data(
      `<row><c t="inlineStr"><is><t>a</t></is></c></row>` +
        `<row><c t="inlineStr"><is><t>b</t></is></c></row>`,
    )
    expect(s.rows).toEqual([["a"], ["b"]])
  })

  it("numbers cells left to right within a row", () => {
    const s = data(`<row r="1"><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>`)
    expect(s.rows[0]).toEqual([1, 2, 3])
  })

  it("resumes implicit numbering after an explicitly placed cell", () => {
    // A mixed file: the first cell pins the column, the next continues
    // from there rather than restarting at A.
    const s = data(`<row r="1"><c r="C1"><v>1</v></c><c><v>2</v></c></row>`)
    expect(s.rows[0]).toEqual([null, null, 1, 2])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <cols>
// ═══════════════════════════════════════════════════════════════════════

describe("column definitions", () => {
  it("drops a <col> with no min/max instead of guessing a span", () => {
    // `min`/`max` are required; without them there is no range to apply
    // the width to, so the element carries no information.
    const s = sheet(`<cols><col width="12"/></cols><sheetData/>`)
    expect(s.columns).toBeUndefined()
  })

  it("drops a <col> whose max is not a number", () => {
    const s = sheet(`<cols><col min="1" max="nonsense" width="12"/></cols><sheetData/>`)
    expect(s.columns).toBeUndefined()
  })

  it("records hidden and outline levels for width-less columns", () => {
    // Grouping columns produces `<col>` entries with no width at all.
    const s = sheet(
      `<cols><col min="2" max="3" hidden="1" outlineLevel="2" collapsed="true"/></cols><sheetData/>`,
    )
    expect(s.columns![1]).toEqual({ hidden: true, outlineLevel: 2, collapsed: true })
    expect(s.columns![2]).toEqual({ hidden: true, outlineLevel: 2, collapsed: true })
    expect(s.columns![0]).toEqual({})
  })

  it("ignores outlineLevel 0, which means no grouping", () => {
    const s = sheet(`<cols><col min="1" max="1" outlineLevel="0" width="9"/></cols><sheetData/>`)
    expect(s.columns![0]).toEqual({ width: 9 })
  })

  it("ignores <col> outside a <cols> wrapper", () => {
    const s = sheet(`<col min="1" max="4" width="30"/><sheetData/>`)
    expect(s.columns).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Row properties
// ═══════════════════════════════════════════════════════════════════════

describe("row definitions", () => {
  it("accepts the boolean spelling Excel uses for row flags", () => {
    // OOXML booleans are `1`/`0` or `true`/`false`; both spellings occur
    // in real files depending on the producer.
    const s = data(
      `<row r="1" ht="30" customHeight="true" hidden="true" collapsed="true" outlineLevel="1">` +
        `<c r="A1"><v>1</v></c></row>`,
    )
    expect(s.rowDefs!.get(0)).toEqual({
      height: 30,
      hidden: true,
      outlineLevel: 1,
      collapsed: true,
    })
  })

  it("ignores ht when customHeight is not set", () => {
    // Excel writes the computed height on every row; only `customHeight`
    // marks it as user-chosen and worth round-tripping.
    const s = data(`<row r="1" ht="15"><c r="A1"><v>1</v></c></row>`)
    expect(s.rowDefs).toBeUndefined()
  })

  it("ignores outlineLevel 0", () => {
    const s = data(`<row r="1" outlineLevel="0"><c r="A1"><v>1</v></c></row>`)
    expect(s.rowDefs).toBeUndefined()
  })

  it("drops row flags on a row that carries no row number", () => {
    // Row properties are keyed by row index; with no `r` there is
    // nothing to key them to, so the flags are dropped even though the
    // row's cells are still read.
    const s = data(
      `<row hidden="1" collapsed="1" outlineLevel="2" ht="40" customHeight="1">` +
        `<c><v>1</v></c></row>`,
    )
    expect(s.rows).toEqual([[1]])
    expect(s.rowDefs).toBeUndefined()
  })

  it("stops parsing rows once maxRows is reached", () => {
    const s = data(
      `<row r="1"><c r="A1"><v>1</v></c></row>` +
        `<row r="2"><c r="A2"><v>2</v></c></row>` +
        `<row r="3"><c r="A3"><v>3</v></c></row>`,
      { maxRows: 2 },
    )
    expect(s.rows).toEqual([[1], [2]])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Cell types
// ═══════════════════════════════════════════════════════════════════════

describe("cell types", () => {
  it('reads t="str" without a formula as a plain string', () => {
    // `str` means "the value is text"; Excel only emits it for formula
    // results, but LibreOffice writes it for literal text too.
    const s = data(`<row r="1"><c r="A1" t="str"><v>plain_x000A_text</v></c></row>`)
    expect(s.rows[0][0]).toBe("plain\ntext")
    expect(s.cells).toBeUndefined()
  })

  it('reads t="str" with a formula as a formula cell', () => {
    const s = data(`<row r="1"><c r="A1" t="str"><f>UPPER(B1)</f><v>HI</v></c></row>`)
    const cell = s.cells!.get("0,0")!
    expect(cell.type).toBe("formula")
    expect(cell.formula).toBe("UPPER(B1)")
  })

  it('reads t="e" error values', () => {
    const s = data(`<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>`)
    expect(s.rows[0][0]).toBe("#DIV/0!")
    expect(s.cells!.get("0,0")!.type).toBe("error")
  })

  it('accepts both boolean spellings for t="b"', () => {
    const s = data(
      `<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c>` +
        `<c r="C1" t="b"><v>TRUE</v></c></row>`,
    )
    expect(s.rows[0]).toEqual([true, false, true])
  })

  it("returns null for a shared-string index past the end of the table", () => {
    // A truncated or mismatched sharedStrings.xml would otherwise leak
    // the raw index into the data as a string.
    const s = data(`<row r="1"><c r="A1" t="s"><v>7</v></c></row>`, { sharedStrings: ss(["only"]) })
    expect(s.rows[0][0]).toBeNull()
  })

  it("keeps non-numeric text in an untyped cell as a string", () => {
    // `<v>` under the default numeric type should be a number; when it
    // isn't, the text is preserved rather than turned into NaN.
    const s = data(`<row r="1"><c r="A1"><v>N/A</v></c></row>`)
    expect(s.rows[0][0]).toBe("N/A")
  })

  it("treats an untyped cell with no value as empty", () => {
    const s = data(`<row r="1"><c r="A1"/><c r="B1"><v>2</v></c></row>`)
    expect(s.rows[0]).toEqual([null, 2])
  })

  it("reads a shared string carrying rich-text runs as a richText cell", () => {
    const rich: SharedString[] = [
      { text: "ab", richText: [{ text: "a", font: { bold: true } }, { text: "b" }] },
    ]
    const s = data(`<row r="1"><c r="A1" t="s"><v>0</v></c></row>`, { sharedStrings: rich })
    const cell = s.cells!.get("0,0")!
    expect(cell.type).toBe("richText")
    expect(cell.richText).toHaveLength(2)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Inline strings and their rich-text runs
// ═══════════════════════════════════════════════════════════════════════

describe("inline strings", () => {
  it("concatenates rich-text runs into the cell value", () => {
    const s = data(
      `<row r="1"><c r="A1" t="inlineStr"><is>` +
        `<r><rPr><b/></rPr><t>Bold</t></r><r><t> tail</t></r>` +
        `</is></c></row>`,
    )
    expect(s.rows[0][0]).toBe("Bold tail")
    const cell = s.cells!.get("0,0")!
    expect(cell.type).toBe("richText")
    expect(cell.richText![0].font).toEqual({ bold: true })
    expect(cell.richText![1].font).toBeUndefined()
  })

  it("maps every rPr child Excel can write onto the run font", () => {
    // One run carrying the full DrawingML-free font vocabulary from
    // §18.4.7 (CT_RPrElt), so each property has a parse path.
    const s = data(
      `<row r="1"><c r="A1" t="inlineStr"><is><r><rPr>` +
        `<b/><i/><u val="double"/><strike/><sz val="14"/><rFont val="Arial"/>` +
        `<color rgb="FFFF0000"/><vertAlign val="superscript"/>` +
        `<family val="2"/><charset val="204"/><scheme val="minor"/>` +
        `</rPr><t>styled</t></r></is></c></row>`,
    )
    expect(s.cells!.get("0,0")!.richText![0].font).toEqual({
      bold: true,
      italic: true,
      underline: "double",
      strikethrough: true,
      size: 14,
      name: "Arial",
      color: { rgb: "FF0000" },
      vertAlign: "superscript",
      family: 2,
      charset: 204,
      scheme: "minor",
    })
  })

  it('honours val="0" as the off state for toggle properties', () => {
    // `<b/>` and `<b val="0"/>` are opposites — the bare element means on.
    const s = data(
      `<row r="1"><c r="A1" t="inlineStr"><is><r><rPr>` +
        `<b val="0"/><i val="false"/><strike val="0"/><u/>` +
        // Six-digit RGB: no alpha prefix to strip, unlike Excel's ARGB.
        `<color rgb="00FF00" theme="4" tint="-0.25" indexed="9"/>` +
        `</rPr><t>off</t></r></is></c></row>`,
    )
    expect(s.cells!.get("0,0")!.richText![0].font).toEqual({
      bold: false,
      italic: false,
      strikethrough: false,
      underline: true,
      color: { rgb: "00FF00", theme: 4, tint: -0.25, indexed: 9 },
    })
  })

  it("ignores unknown rPr children and out-of-range enum values", () => {
    const s = data(
      `<row r="1"><c r="A1" t="inlineStr"><is><r><rPr>` +
        `<condense/><sz/><rFont/><vertAlign val="sideways"/>` +
        `<family/><charset/><scheme val="bogus"/>` +
        `</rPr><t>t</t></r></is></c></row>`,
    )
    expect(s.cells!.get("0,0")!.richText![0].font).toEqual({})
  })

  it("decodes OOXML escapes in a plain inline string", () => {
    const s = data(`<row r="1"><c r="A1" t="inlineStr"><is><t>a_x000A_b</t></is></c></row>`)
    expect(s.rows[0][0]).toBe("a\nb")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Merges, hyperlinks
// ═══════════════════════════════════════════════════════════════════════

describe("merged cells", () => {
  it("accepts a degenerate single-cell merge ref", () => {
    // `<mergeCell ref="B2"/>` (no colon) is malformed but appears in
    // files produced by report generators; treat start == end.
    const s = sheet(`<sheetData/><mergeCells count="1"><mergeCell ref="B2"/></mergeCells>`)
    expect(s.merges).toEqual([{ startRow: 1, startCol: 1, endRow: 1, endCol: 1 }])
  })

  it("ignores a mergeCell with no ref", () => {
    const s = sheet(`<sheetData/><mergeCells count="1"><mergeCell/></mergeCells>`)
    expect(s.merges).toBeUndefined()
  })
})

describe("hyperlinks", () => {
  it("creates a cell for a link that points at an empty area of the sheet", () => {
    // Excel keeps hyperlinks after the text is deleted, so the target
    // row may not exist in <sheetData> at all.
    const s = sheet(
      `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
        `<hyperlinks><hyperlink ref="C10" r:id="rId1" tooltip="tip" display="shown"/></hyperlinks>`,
      { worksheetRels: [{ id: "rId1", type: "hyperlink", target: "https://example.com" }] },
    )
    const cell = s.cells!.get("9,2")!
    expect(cell.value).toBeNull()
    expect(cell.hyperlink).toEqual({
      target: "https://example.com",
      tooltip: "tip",
      display: "shown",
    })
  })

  it("leaves the target empty when the rId resolves to nothing", () => {
    // A dangling r:id — the .rels file lost the entry. Better an empty
    // target than a thrown read.
    const s = sheet(
      `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
        `<hyperlinks><hyperlink ref="A1" r:id="rId99"/></hyperlinks>`,
      { worksheetRels: [] },
    )
    expect(s.cells!.get("0,0")!.hyperlink).toEqual({ target: "" })
  })

  it("uses the location for an internal link", () => {
    const s = sheet(
      `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
        `<hyperlinks><hyperlink ref="A1" location="Sheet2!A1"/></hyperlinks>`,
    )
    expect(s.cells!.get("0,0")!.hyperlink).toEqual({
      target: "Sheet2!A1",
      location: "Sheet2!A1",
    })
  })

  it("ignores a hyperlink with no ref", () => {
    const s = sheet(`<sheetData/><hyperlinks><hyperlink r:id="rId1"/></hyperlinks>`)
    expect(s.cells).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Range filter
// ═══════════════════════════════════════════════════════════════════════

describe("range read option", () => {
  it("keeps only cells inside the window", () => {
    const s = data(
      `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>` +
        `<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>`,
      { range: "B2:B2" },
    )
    expect(s.rows[1][1]).toBe(4)
    // Rows are still padded to the sheet's bounding box, so the skipped
    // cells read back as null rather than disappearing.
    expect(s.rows[0]).toEqual([null, null])
  })

  it("accepts a single-cell range with no colon", () => {
    const s = data(`<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>`, {
      range: "A1",
    })
    expect(s.rows[0]).toEqual([1])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Sheet view, panes, protection
// ═══════════════════════════════════════════════════════════════════════

describe("sheet view", () => {
  it("reads a six-digit tab colour verbatim", () => {
    // Excel writes ARGB (8 chars); the alpha prefix is stripped. A
    // six-char value has no alpha to strip.
    const s = sheet(`<sheetPr><tabColor rgb="FF0000" theme="2" tint="0.4" indexed="7"/></sheetPr>`)
    expect(s.view!.tabColor).toEqual({ rgb: "FF0000", theme: 2, tint: 0.4, indexed: 7 })
  })

  it("ignores a tabColor outside <sheetPr>", () => {
    const s = sheet(`<tabColor rgb="FFFF0000"/><sheetData/>`)
    expect(s.view).toBeUndefined()
  })

  it("accepts the false/0 spellings for view toggles", () => {
    const s = sheet(
      `<sheetViews><sheetView showGridLines="false" showRowColHeaders="false" ` +
        `zoomScale="150" rightToLeft="true"/></sheetViews><sheetData/>`,
    )
    expect(s.view).toEqual({
      showGridLines: false,
      showRowColHeaders: false,
      zoomScale: 150,
      rightToLeft: true,
    })
  })

  it("ignores a non-numeric zoomScale", () => {
    const s = sheet(`<sheetViews><sheetView zoomScale="big"/></sheetViews><sheetData/>`)
    expect(s.view).toBeUndefined()
  })
})

describe("panes", () => {
  it("reads a split pane", () => {
    const s = sheet(
      `<sheetViews><sheetView><pane xSplit="1200" ySplit="600" state="split"/></sheetView>` +
        `</sheetViews><sheetData/>`,
    )
    expect(s.splitPane).toEqual({ xSplit: 1200, ySplit: 600 })
    expect(s.freezePane).toBeUndefined()
  })

  it("reads frozenSplit as a freeze", () => {
    const s = sheet(
      `<sheetViews><sheetView><pane xSplit="2" state="frozenSplit"/></sheetView>` +
        `</sheetViews><sheetData/>`,
    )
    expect(s.freezePane).toEqual({ columns: 2 })
  })

  it("ignores a pane that splits nothing", () => {
    const s = sheet(
      `<sheetViews><sheetView><pane xSplit="0" ySplit="0" state="split"/></sheetView>` +
        `</sheetViews><sheetData/>`,
    )
    expect(s.splitPane).toBeUndefined()
  })

  it("ignores a pane with no state (a plain scroll position)", () => {
    const s = sheet(
      `<sheetViews><sheetView><pane topLeftCell="B2"/></sheetView></sheetViews><sheetData/>`,
    )
    expect(s.freezePane).toBeUndefined()
    expect(s.splitPane).toBeUndefined()
  })
})

describe("sheet protection", () => {
  it("accepts the true/false spellings", () => {
    const s = sheet(
      `<sheetData/><sheetProtection sheet="true" objects="true" scenarios="true" ` +
        `formatCells="false" sort="1"/>`,
    )
    expect(s.protection).toEqual({
      sheet: true,
      objects: true,
      scenarios: true,
      formatCells: true,
      sort: false,
    })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Auto filter
// ═══════════════════════════════════════════════════════════════════════

describe("auto filter", () => {
  it("ignores an autoFilter with no ref", () => {
    const s = sheet(`<sheetData/><autoFilter/>`)
    expect(s.autoFilter).toBeUndefined()
  })

  it("drops a filterColumn that lists no values", () => {
    // A `<filterColumn>` holding only `<customFilters>` or `<top10>` has
    // no plain value list for us to surface yet.
    const s = sheet(
      `<sheetData/><autoFilter ref="A1:B9"><filterColumn colId="0"><top10 val="5"/>` +
        `</filterColumn></autoFilter>`,
    )
    expect(s.autoFilter).toEqual({ range: "A1:B9" })
  })

  it("collects the value list of a filterColumn", () => {
    const s = sheet(
      `<sheetData/><autoFilter ref="A1:B9"><filterColumn colId="1"><filters>` +
        `<filter val="x"/><filter val="y"/></filters></filterColumn></autoFilter>`,
    )
    expect(s.autoFilter!.columns).toEqual([{ colIndex: 1, filters: ["x", "y"] }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Data validation
// ═══════════════════════════════════════════════════════════════════════

describe("data validation", () => {
  it("drops a validation whose type is not in the schema enum", () => {
    const s = sheet(
      `<sheetData/><dataValidations count="1">` +
        `<dataValidation type="madeUp" sqref="A1"><formula1>1</formula1></dataValidation>` +
        `</dataValidations>`,
    )
    expect(s.dataValidations).toBeUndefined()
  })

  it("drops a validation with no sqref to apply it to", () => {
    const s = sheet(
      `<sheetData/><dataValidations count="1">` +
        `<dataValidation type="whole"><formula1>1</formula1></dataValidation></dataValidations>`,
    )
    expect(s.dataValidations).toBeUndefined()
  })

  it("splits a quoted inline list into values", () => {
    const s = sheet(
      `<sheetData/><dataValidations count="1">` +
        `<dataValidation type="list" sqref="A1:A5" allowBlank="true" showInputMessage="true" ` +
        `showErrorMessage="true" errorStyle="warning" promptTitle="Pick" prompt="One of these" ` +
        `errorTitle="Nope" error="Not in list">` +
        `<formula1>"red,green,blue"</formula1></dataValidation></dataValidations>`,
    )
    expect(s.dataValidations![0]).toEqual({
      type: "list",
      range: "A1:A5",
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      errorStyle: "warning",
      inputTitle: "Pick",
      inputMessage: "One of these",
      errorTitle: "Nope",
      errorMessage: "Not in list",
      values: ["red", "green", "blue"],
    })
  })

  it("keeps an unquoted list formula as a reference", () => {
    const s = sheet(
      `<sheetData/><dataValidations count="1">` +
        `<dataValidation type="list" sqref="A1"><formula1>Sheet2!$A$1:$A$9</formula1>` +
        `</dataValidation></dataValidations>`,
    )
    expect(s.dataValidations![0].formula1).toBe("Sheet2!$A$1:$A$9")
    expect(s.dataValidations![0].values).toBeUndefined()
  })

  it("reads both formulas of a between rule and ignores an unknown operator", () => {
    const s = sheet(
      `<sheetData/><dataValidations count="1">` +
        `<dataValidation type="decimal" operator="somethingElse" sqref="B1">` +
        `<formula1>0</formula1><formula2>100</formula2></dataValidation></dataValidations>`,
    )
    expect(s.dataValidations![0].operator).toBeUndefined()
    expect(s.dataValidations![0].formula1).toBe("0")
    expect(s.dataValidations![0].formula2).toBe("100")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Conditional formatting
// ═══════════════════════════════════════════════════════════════════════

describe("conditional formatting", () => {
  it("drops a rule whose block has no sqref", () => {
    const s = sheet(
      `<sheetData/><conditionalFormatting><cfRule type="expression" priority="1">` +
        `<formula>TRUE</formula></cfRule></conditionalFormatting>`,
    )
    expect(s.conditionalRules).toBeUndefined()
  })

  it("drops a rule of an unknown type", () => {
    const s = sheet(
      `<sheetData/><conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="notARule" priority="1"/></conditionalFormatting>`,
    )
    expect(s.conditionalRules).toBeUndefined()
  })

  it("defaults priority to 1 when the attribute is absent", () => {
    const s = sheet(
      `<sheetData/><conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="containsBlanks"/></conditionalFormatting>`,
    )
    expect(s.conditionalRules![0].priority).toBe(1)
  })

  it("keeps both formulas of a two-formula rule as an array", () => {
    // `cellIs` with operator `between` carries two `<formula>` children.
    const s = sheet(
      `<sheetData/><conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="cellIs" operator="between" priority="3" stopIfTrue="true">` +
        `<formula>1</formula><formula>10</formula></cfRule></conditionalFormatting>`,
    )
    expect(s.conditionalRules![0]).toMatchObject({
      operator: "between",
      priority: 3,
      stopIfTrue: true,
      formula: ["1", "10"],
    })
  })

  it("keeps the text attribute of a containsText rule", () => {
    const s = sheet(
      `<sheetData/><conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="containsText" operator="containsText" text="err" priority="1">` +
        `<formula>NOT(ISERROR(SEARCH("err",A1)))</formula></cfRule></conditionalFormatting>`,
    )
    expect(s.conditionalRules![0].text).toBe("err")
    expect(s.conditionalRules![0].formula).toBe('NOT(ISERROR(SEARCH("err",A1)))')
  })

  it("defaults a cfvo with no type to min, and keeps a theme colour on a scale stop", () => {
    const s = sheet(
      `<sheetData/><conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo/><cfvo type="max"/><color/><color theme="4"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(s.conditionalRules![0].colorScale).toEqual({
      cfvo: [
        { type: "min", value: undefined },
        { type: "max", value: undefined },
      ],
      colors: [{}, { theme: 4 }],
    })
  })

  it("reads a data bar, defaulting its cfvo type and colour", () => {
    const s = sheet(
      `<sheetData/><conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="dataBar" priority="1"><dataBar>` +
        `<cfvo/><cfvo type="max"/><color/></dataBar></cfRule></conditionalFormatting>`,
    )
    expect(s.conditionalRules![0].dataBar).toEqual({
      cfvo: [
        { type: "min", value: undefined },
        { type: "max", value: undefined },
      ],
      color: {},
    })
  })

  it("falls back to 3TrafficLights1 when the icon set is unnamed", () => {
    const s = sheet(
      `<sheetData/><conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="iconSet" priority="1"><iconSet reverse="true" showValue="false">` +
        `<cfvo/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>` +
        `</iconSet></cfRule></conditionalFormatting>`,
    )
    expect(s.conditionalRules![0].iconSet).toEqual({
      iconSet: "3TrafficLights1",
      reverse: true,
      showValue: false,
      cfvo: [
        { type: "min", value: undefined },
        { type: "percent", value: "33" },
        { type: "percent", value: "67" },
      ],
    })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Print settings
// ═══════════════════════════════════════════════════════════════════════

describe("page setup", () => {
  it("keeps a paper size hucre has no name for, as the code it is", () => {
    // Dropping it lost the page size with no error and nothing in the
    // parity statement. `PaperSize` admits a raw code, so an unnamed one
    // round-trips as the number. See #439 §Q.
    const s = sheet(`<sheetData/><pageSetup paperSize="256" orientation="sideways" scale="80"/>`)
    expect(s.pageSetup).toEqual({ paperSize: 256, scale: 80 })
  })

  it("still ignores a paper size that is not a usable code", () => {
    const s = sheet(`<sheetData/><pageSetup paperSize="0" scale="80"/>`)
    expect(s.pageSetup).toEqual({ scale: 80 })
  })

  it("turns on fitToPage when only fitToHeight is present", () => {
    // Either half of the fit-to pair implies the fitToPage mode; the
    // missing half stays unset rather than defaulting to 1.
    const s = sheet(`<sheetData/><pageSetup fitToHeight="2"/>`)
    expect(s.pageSetup).toEqual({ fitToPage: true, fitToHeight: 2 })
  })

  it("still accepts centering flags written on pageSetup", () => {
    const s = sheet(
      `<sheetData/><pageSetup horizontalCentered="true" verticalCentered="1" paperSize="9"/>`,
    )
    expect(s.pageSetup).toEqual({
      paperSize: "a4",
      horizontalCentered: true,
      verticalCentered: true,
    })
  })

  it("merges printOptions written before pageSetup", () => {
    const s = sheet(
      `<sheetData/><printOptions gridLines="true" headings="1" horizontalCentered="1" ` +
        `verticalCentered="true"/><pageSetup orientation="landscape"/>`,
    )
    expect(s.pageSetup).toEqual({
      showGridLines: true,
      showRowColHeaders: true,
      horizontalCentered: true,
      verticalCentered: true,
      orientation: "landscape",
    })
  })

  it("attaches margins even with no pageSetup element", () => {
    const s = sheet(
      `<sheetData/><pageMargins left="0.5" right="0.5" top="1" bottom="1" ` +
        `header="0.3" footer="0.3"/>`,
    )
    expect(s.pageSetup).toEqual({
      margins: { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.3, footer: 0.3 },
    })
  })
})

describe("headers, footers and page breaks", () => {
  it("reads every header/footer slot", () => {
    const s = sheet(
      `<sheetData/><headerFooter differentOddEven="true" differentFirst="1">` +
        `<oddHeader>&amp;Codd h</oddHeader><oddFooter>odd f</oddFooter>` +
        `<evenHeader>even h</evenHeader><evenFooter>even f</evenFooter>` +
        `<firstHeader>first h</firstHeader><firstFooter>first f</firstFooter></headerFooter>`,
    )
    expect(s.headerFooter).toEqual({
      differentOddEven: true,
      differentFirst: true,
      oddHeader: "&Codd h",
      oddFooter: "odd f",
      evenHeader: "even h",
      evenFooter: "even f",
      firstHeader: "first h",
      firstFooter: "first f",
    })
  })

  it("sorts page breaks and converts them to 0-based indices", () => {
    const s = sheet(
      `<sheetData/><rowBreaks count="2"><brk id="20" max="16383" man="1"/>` +
        `<brk id="10"/><brk/></rowBreaks>` +
        `<colBreaks count="1"><brk id="4"/></colBreaks>`,
    )
    expect(s.rowBreaks).toEqual([9, 19])
    expect(s.colBreaks).toEqual([3])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Sparklines
// ═══════════════════════════════════════════════════════════════════════

describe("sparklines", () => {
  it("reads a column sparkline group with markers and a six-digit colour", () => {
    const s = sheet(
      `<sheetData/><extLst><ext><x14:sparklineGroups ` +
        `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
        `<x14:sparklineGroup type="column" markers="true">` +
        `<x14:colorSeries rgb="336699"/>` +
        `<x14:sparklines><x14:sparkline>` +
        `<xm:f>Sheet1!B1:E1</xm:f><xm:sqref>A1</xm:sqref>` +
        `</x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`,
    )
    expect(s.sparklines).toEqual([
      {
        location: "A1",
        dataRange: "Sheet1!B1:E1",
        type: "column",
        color: { rgb: "336699" },
        markers: true,
      },
    ])
  })

  it("drops a sparkline that names no cell to draw in", () => {
    const s = sheet(
      `<sheetData/><extLst><ext><x14:sparklineGroups ` +
        `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
        `<x14:sparklineGroup><x14:colorSeries/><x14:sparklines><x14:sparkline>` +
        `<xm:f>Sheet1!B1:E1</xm:f></x14:sparkline></x14:sparklines>` +
        `</x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`,
    )
    expect(s.sparklines).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Formulas and styles
// ═══════════════════════════════════════════════════════════════════════

describe("formulas", () => {
  it("marks the slave cells of a shared formula", () => {
    // The master carries the text plus `ref`; the slaves carry only `si`,
    // so their formula text is empty but the link must survive.
    const s = data(
      `<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c></row>` +
        `<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row>`,
    )
    expect(s.cells!.get("0,0")).toMatchObject({
      formula: "B1*2",
      formulaType: "shared",
      formulaSharedIndex: 0,
      formulaRef: "A1:A3",
      formulaResult: 2,
    })
    expect(s.cells!.get("1,0")).toMatchObject({
      formula: "",
      formulaType: "shared",
      formulaSharedIndex: 0,
    })
  })

  it("flags a dynamic array formula from cm on the cell", () => {
    const s = data(
      `<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">SEQUENCE(3)</f>` + `<v>1</v></c></row>`,
      { dynamicArrayCm: new Set([1]) },
    )
    expect(s.cells!.get("0,0")).toMatchObject({
      formulaType: "array",
      formulaRef: "A1:A3",
      formulaDynamic: true,
    })
  })

  // hucre wrote `cm` here up to 0.6 (#423). The attribute means nothing
  // on `<f>` per the schema, so those files can only have come from
  // hucre — keep reading them rather than silently dropping the flag.
  it("still flags a dynamic array formula from cm on <f> (pre-0.7 hucre files)", () => {
    const s = data(
      `<row r="1"><c r="A1"><f t="array" ref="A1:A3" cm="1">SEQUENCE(3)</f>` + `<v>1</v></c></row>`,
    )
    expect(s.cells!.get("0,0")).toMatchObject({ formulaDynamic: true })
  })

  it("ignores a cm that the metadata part does not map to a dynamic array", () => {
    const s = data(
      `<row r="1"><c r="A1" cm="2"><f>SUM(B1:B3)</f><v>1</v></c></row>`,
      // Only index 1 is XLDAPR in this package; 2 is some other kind of
      // cell metadata and says nothing about spilling.
      { dynamicArrayCm: new Set([1]) },
    )
    expect(s.cells!.get("0,0")!.formulaDynamic).toBeUndefined()
  })
})

describe("styles", () => {
  const stylesXml =
    `<?xml version="1.0"?><styleSheet ${NS}>` +
    `<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>` +
    `<fonts count="1"><font><sz val="11"/></font></fonts>` +
    `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
    `<borders count="1"><border/></borders>` +
    `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
    `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>` +
    `</styleSheet>`

  it("turns a serial into a Date when the cell's format is a date format", () => {
    const s = data(`<row r="1"><c r="A1" s="1"><v>45000</v></c></row>`, {
      styles: parseStyles(stylesXml),
    })
    expect(s.rows[0][0]).toBeInstanceOf(Date)
  })

  it("uses the 1904 epoch when the workbook asks for it", () => {
    const y1900 = data(`<row r="1"><c r="A1" s="1"><v>45000</v></c></row>`, {
      styles: parseStyles(stylesXml),
    }).rows[0][0] as Date
    const y1904 = data(`<row r="1"><c r="A1" s="1"><v>45000</v></c></row>`, {
      styles: parseStyles(stylesXml),
      dateSystem: "1904",
    }).rows[0][0] as Date
    expect(y1904.getTime()).toBeGreaterThan(y1900.getTime())
  })

  it("emits no Cell object for a styled cell unless readStyles is on", () => {
    const withoutStyles = data(`<row r="1"><c r="A1" s="1"><v>1</v></c></row>`, {
      styles: parseStyles(stylesXml),
    })
    expect(withoutStyles.cells).toBeUndefined()

    const withStyles = data(`<row r="1"><c r="A1" s="1"><v>1</v></c></row>`, {
      styles: parseStyles(stylesXml),
      readStyles: true,
    })
    expect(withStyles.cells!.get("0,0")!.style).toBeDefined()
  })

  it("skips style resolution for a cell with no s attribute", () => {
    const s = data(`<row r="1"><c r="A1"><v>1</v></c></row>`, {
      styles: parseStyles(stylesXml),
      readStyles: true,
    })
    expect(s.cells).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Empty sheets
// ═══════════════════════════════════════════════════════════════════════

describe("empty worksheets", () => {
  it("returns no rows for a sheet with an empty sheetData", () => {
    const s = data("")
    expect(s.rows).toEqual([])
    expect(s.cells).toBeUndefined()
    expect(s.merges).toBeUndefined()
  })

  it("ignores <c> and <v> outside a row", () => {
    // Guards the SAX state machine: stray elements must not open cell
    // state and corrupt the next real row.
    const s = sheet(`<c r="A1"><v>9</v></c><sheetData/>`)
    expect(s.rows).toEqual([])
  })
})
