import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { XlsxStreamWriter, writeXlsxStream } from "../src/xlsx/stream-writer"
import { writeTable } from "../src/xlsx/table-writer"
import { writeContentTypes } from "../src/xlsx/content-types-writer"
import { writeCustomProperties } from "../src/xlsx/doc-props-writer"
import { writeXml } from "../src/xml/data-writer"
import { deserializeWorkbook, serializeWorkbook } from "../src/worker"
import { ZipReader } from "../src/zip/reader"
import type {
  Cell,
  CellStyle,
  ConditionalRule,
  RowDef,
  Sheet,
  Workbook,
  WriteOptions,
  WriteSheet,
} from "../src/_types"

const decoder = new TextDecoder("utf-8")

// ── Helpers ──────────────────────────────────────────────────────────

async function part(buf: Uint8Array, path: string): Promise<string> {
  return decoder.decode(await new ZipReader(buf).extract(path))
}

/** Write one sheet and hand back its `xl/worksheets/sheet1.xml`. */
async function sheetXml(sheet: WriteSheet, options?: Omit<WriteOptions, "sheets">) {
  const buf = await writeXlsx({ sheets: [sheet], ...options })
  return part(buf, "xl/worksheets/sheet1.xml")
}

/** Write one sheet and hand back its `xl/styles.xml`. */
async function stylesXml(sheet: WriteSheet) {
  const buf = await writeXlsx({ sheets: [sheet] })
  return part(buf, "xl/styles.xml")
}

const cellMap = (entries: Array<[string, Partial<Cell>]>): Map<string, Partial<Cell>> =>
  new Map(entries)

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    total += value.length
  }
  const out = new Uint8Array(total)
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}

// ═══════════════════════════════════════════════════════════════════════
// worksheet-writer — row emission
//
// A `<row>` can exist for two independent reasons: it holds cells, or it
// holds row-level properties (height / hidden / outline). The second one
// only ever appears in files that came from Excel, so the write path for
// it went unverified.
// ═══════════════════════════════════════════════════════════════════════

describe("worksheet rows without cells", () => {
  it("emits a self-closing <row> for a row that carries only row properties", async () => {
    // Row 3 is an empty spacer inside the data range, given a custom
    // height. Excel writes `<row r="3" ht="40" customHeight="1"/>`;
    // dropping it would lose the height the user set.
    const rowDefs = new Map<number, RowDef>([[2, { height: 40 }]])
    const xml = await sheetXml({ name: "S", rows: [["a"], ["b"], [], ["d"]], rowDefs })

    expect(xml).toContain('<row r="3" ht="40" customHeight="1"/>')
  })

  it("carries hidden, outlineLevel and collapsed onto a property-only row", async () => {
    const rowDefs = new Map<number, RowDef>([
      [1, { hidden: true }],
      [2, { outlineLevel: 2 }],
      [3, { collapsed: true }],
    ])
    const xml = await sheetXml({ name: "S", rows: [["a"], [], [], [], ["e"]], rowDefs })

    expect(xml).toContain('<row r="2" hidden="1"/>')
    expect(xml).toContain('<row r="3" outlineLevel="2"/>')
    expect(xml).toContain('<row r="4" collapsed="1"/>')
  })

  it("skips the placeholder cells a far-right cell override pads a row with", async () => {
    // Setting F1 on a sheet with no `rows` grows the row with nulls at
    // A1..E1. Those placeholders must not become `<c/>` elements.
    const xml = await sheetXml({
      name: "S",
      cells: cellMap([["0,5", { value: "far right" }]]),
    })

    expect(xml).toContain('<c r="F1"')
    expect(xml).not.toContain('r="A1"')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// worksheet-writer — cell serialization
// ═══════════════════════════════════════════════════════════════════════

describe("styled cells of every value shape", () => {
  const styled: CellStyle = { font: { bold: true } }

  it("keeps the style index on a rich-text cell", async () => {
    const xml = await sheetXml({
      name: "S",
      cells: cellMap([
        [
          "0,0",
          {
            style: styled,
            richText: [{ text: "bold-ish " }, { text: "run", font: { italic: true } }],
          },
        ],
      ]),
    })

    expect(xml).toMatch(/<c r="A1" t="inlineStr" s="\d+">/)
  })

  it("keeps the style index on an error cell", async () => {
    const xml = await sheetXml({
      name: "S",
      cells: cellMap([["0,0", { value: "#DIV/0!", style: styled }]]),
    })

    expect(xml).toMatch(/<c r="A1" t="e" s="\d+"><v>#DIV\/0!<\/v><\/c>/)
  })

  it("keeps the style index on an inline string when stringMode is inline", async () => {
    const xml = await sheetXml(
      { name: "S", cells: cellMap([["0,0", { value: "hello", style: styled }]]) },
      { stringMode: "inline" },
    )

    expect(xml).toMatch(/<c r="A1" t="inlineStr" s="\d+"><is><t>hello<\/t><\/is><\/c>/)
  })

  it("keeps the style on a cell whose number is not finite", async () => {
    // NaN / Infinity have no OOXML representation, but the formatting the
    // user applied to the cell still has to survive.
    const xml = await sheetXml({
      name: "S",
      cells: cellMap([
        ["0,0", { value: Number.POSITIVE_INFINITY, style: styled }],
        ["1,0", { value: Number.NaN }],
      ]),
    })

    expect(xml).toMatch(/<c r="A1" s="\d+"\/>/)
    // Unstyled, so there is nothing left to write at all.
    expect(xml).not.toContain('r="A2"')
  })

  it("drops a value the cell model cannot represent", async () => {
    // Objects reach here from untyped JSON input. There is no `t=` for
    // them, so the cell is skipped rather than written as "[object Object]".
    const xml = await sheetXml({
      name: "S",
      cells: cellMap([
        ["0,0", { value: { nested: true } as never }],
        ["0,1", { value: "kept" }],
      ]),
    })

    expect(xml).not.toContain('r="A1"')
    expect(xml).toContain('<c r="B1" t="s"')
  })

  it('writes a boolean cached formula result as t="b"', async () => {
    const xml = await sheetXml({
      name: "S",
      cells: cellMap([
        ["0,0", { formula: "ISBLANK(B1)", formulaResult: true }],
        ["1,0", { formula: "ISBLANK(B2)", formulaResult: false }],
      ]),
    })

    expect(xml).toContain('<c r="A1" t="b"><f>ISBLANK(B1)</f><v>1</v></c>')
    expect(xml).toContain('<c r="A2" t="b"><f>ISBLANK(B2)</f><v>0</v></c>')
  })

  it("leaves an explicit number format on a date cell alone", async () => {
    // Dates only get the yyyy-mm-dd default when the style says nothing;
    // a caller-supplied format must not be overwritten.
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          cells: cellMap([
            ["0,0", { value: new Date(Date.UTC(2024, 0, 15)), style: { numFmt: "mmm yyyy" } }],
          ]),
        },
      ],
    })
    const styles = await part(buf, "xl/styles.xml")

    expect(styles).toContain('formatCode="mmm yyyy"')
    expect(styles).not.toContain('formatCode="yyyy-mm-dd"')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// worksheet-writer — sheet-level blocks
// ═══════════════════════════════════════════════════════════════════════

describe("sheet-level OOXML blocks", () => {
  it("emits a dimension for a sheet that has columns but no rows", async () => {
    // A column-definitions-only sheet is what a template looks like before
    // any data is appended; `<dimension>` must still cover the columns.
    const xml = await sheetXml({
      name: "S",
      rows: [],
      columns: [{ key: "a", width: 20 }, { key: "b" }, { key: "c" }],
    })

    expect(xml).toContain('<dimension ref="A1:C1"/>')
  })

  it("writes summaryBelow / summaryRight as 0 when outlines summarise above and left", async () => {
    const xml = await sheetXml({
      name: "S",
      rows: [["a"]],
      outlineProperties: { summaryBelow: false, summaryRight: false },
    })

    expect(xml).toContain('<outlinePr summaryBelow="0" summaryRight="0"/>')
  })

  it("writes summaryBelow / summaryRight as 1 for Excel's default outline placement", async () => {
    const xml = await sheetXml({
      name: "S",
      rows: [["a"]],
      outlineProperties: { summaryBelow: true, summaryRight: true },
    })

    expect(xml).toContain('<outlinePr summaryBelow="1" summaryRight="1"/>')
  })

  it("passes an 8-character ARGB tab colour through without a second alpha prefix", async () => {
    const xml = await sheetXml({
      name: "S",
      rows: [["a"]],
      view: { tabColor: { rgb: "80FF0000" } },
    })

    expect(xml).toContain('<tabColor rgb="80FF0000"/>')
  })

  it("emits a self-closing <dataValidation> when the rule carries no formula", async () => {
    // `custom` without a formula is degenerate but legal — Excel keeps the
    // prompt text and drops the constraint.
    const xml = await sheetXml({
      name: "S",
      rows: [["a"]],
      dataValidations: [
        {
          type: "custom",
          range: "A1:A10",
          showInputMessage: true,
          inputTitle: "Heads up",
          inputMessage: "Anything goes",
        },
      ],
    })

    expect(xml).toContain('<dataValidation type="custom" sqref="A1:A10"')
    expect(xml).toContain('prompt="Anything goes"/>')
  })

  it("omits <headerFooter> entirely when only the flags are set", async () => {
    // `differentOddEven` with no odd/even strings describes nothing; an
    // empty `<headerFooter differentOddEven="1"/>` would just be noise.
    const xml = await sheetXml({
      name: "S",
      rows: [["a"]],
      headerFooter: { differentOddEven: true, differentFirst: true },
    })

    expect(xml).not.toContain("headerFooter")
  })

  it("serializes theme, tint and indexed colours on a tab colour", async () => {
    const xml = await sheetXml({
      name: "S",
      rows: [["a"]],
      view: { tabColor: { theme: 4, tint: -0.25, indexed: 12 } },
    })

    expect(xml).toContain('<tabColor theme="4" tint="-0.25" indexed="12"/>')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// worksheet-writer — rich text run properties
//
// `<rPr>` accepts the same font children as `<font>` in styles.xml. The
// less common ones (vertAlign / family / charset / scheme) only appear on
// text pasted out of Word or a non-Latin locale.
// ═══════════════════════════════════════════════════════════════════════

describe("rich text run properties", () => {
  it("writes vertAlign, family, charset and scheme into <rPr>", async () => {
    const xml = await sheetXml({
      name: "S",
      cells: cellMap([
        [
          "0,0",
          {
            richText: [
              { text: "E=mc" },
              {
                text: "2",
                font: {
                  vertAlign: "superscript",
                  family: 2,
                  charset: 204,
                  scheme: "minor",
                  color: { indexed: 10 },
                },
              },
            ],
          },
        ],
      ]),
    })

    expect(xml).toContain('<vertAlign val="superscript"/>')
    expect(xml).toContain('<family val="2"/>')
    expect(xml).toContain('<charset val="204"/>')
    expect(xml).toContain('<scheme val="minor"/>')
    expect(xml).toContain('<color indexed="10"/>')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// worksheet-writer — conditional formatting and sparklines
// ═══════════════════════════════════════════════════════════════════════

describe("conditional formatting rule bodies", () => {
  it("emits one <formula> per entry when the rule carries a formula pair", async () => {
    // `between` needs two operands, which the API accepts as an array.
    const rule: ConditionalRule = {
      type: "cellIs",
      range: "A1:A10",
      operator: "between",
      priority: 1,
      formula: ["10", "20"],
      style: { fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } } },
    }
    const xml = await sheetXml({ name: "S", rows: [[5]], conditionalRules: [rule] })

    expect(xml).toContain("<formula>10</formula><formula>20</formula>")
  })

  it("writes the cfvo values of a data bar", async () => {
    const rule: ConditionalRule = {
      type: "dataBar",
      range: "A1:A10",
      priority: 1,
      dataBar: {
        cfvo: [
          { type: "num", value: "0" },
          { type: "num", value: "100" },
        ],
        color: "638EC6",
      },
    }
    const xml = await sheetXml({ name: "S", rows: [[5]], conditionalRules: [rule] })

    expect(xml).toContain('<cfvo type="num" val="0"/>')
    expect(xml).toContain('<cfvo type="num" val="100"/>')
    expect(xml).toContain('<color rgb="638EC6"/>')
  })

  it("honours reverse and showValue:false on an icon set", async () => {
    const rule: ConditionalRule = {
      type: "iconSet",
      range: "A1:A10",
      priority: 1,
      iconSet: {
        iconSet: "3TrafficLights1",
        reverse: true,
        showValue: false,
        cfvo: [
          { type: "percent", value: "0" },
          { type: "percent", value: "33" },
          { type: "percent", value: "67" },
        ],
      },
    }
    const xml = await sheetXml({ name: "S", rows: [[5]], conditionalRules: [rule] })

    expect(xml).toContain('reverse="true"')
    expect(xml).toContain('showValue="false"')
  })
})

describe("sparkline colours", () => {
  it("passes an 8-character ARGB colour through untouched", async () => {
    // 6-char colours get an FF alpha prefix; a value that already carries
    // alpha must not be prefixed twice.
    const xml = await sheetXml({
      name: "S",
      rows: [[1, 2, 3]],
      sparklines: [{ location: "D1", dataRange: "Sheet1!A1:C1", color: "80FF0000" }],
    })

    expect(xml).toContain('<x14:colorSeries rgb="80FF0000"/>')
    expect(xml).not.toContain("FF80FF0000")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// worksheet-writer — object data (`data` + `columns`)
// ═══════════════════════════════════════════════════════════════════════

describe("object rows resolved through column keys", () => {
  it("applies a column numFmt to every cell in the column", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          columns: [
            { key: "name", header: "Name" },
            { key: "amount", header: "Amount", numFmt: "#,##0.00" },
          ],
          data: [{ name: "Widget", amount: 12.5 }],
        },
      ],
    })

    expect(await part(buf, "xl/styles.xml")).toContain('formatCode="#,##0.00"')
  })

  it("leaves a column's own style numFmt in place instead of overriding it", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          columns: [{ key: "amount", header: "Amount", numFmt: "0.00", style: { numFmt: "0%" } }],
          data: [{ amount: 0.5 }],
        },
      ],
    })
    const styles = await part(buf, "xl/styles.xml")

    expect(styles).toContain('formatCode="0%"')
    expect(styles).not.toContain('formatCode="0.00"')
  })

  it("falls back to the column key, then to nothing, for a missing header", async () => {
    // `columns` is often written key-only; the header row still has to be
    // emitted (one column here declares neither, and contributes nothing).
    const xml = await sheetXml({
      name: "S",
      columns: [{ key: "a", header: "A" }, { key: "b" }, {}],
      data: [{ a: 1, b: 2 }],
    })

    expect(xml).toContain('<row r="1">')
    // A and the key-named column are shared strings 0 and 1.
    expect(xml).toContain('<c r="B1" t="s"><v>1</v></c>')
    expect(xml).not.toContain('r="C1"')
  })

  it("writes an empty cell for a key the object does not have", async () => {
    const xml = await sheetXml({
      name: "S",
      columns: [
        { key: "a", header: "A" },
        { key: "missing", header: "B" },
      ],
      data: [{ a: 1 }],
    })

    // Row 2 has A2 only; the missing key contributes no `<c>`.
    expect(xml).toContain('<row r="2"><c r="A2"><v>1</v></c></row>')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// styles-writer
// ═══════════════════════════════════════════════════════════════════════

describe("styles.xml colour and alignment coverage", () => {
  it("writes theme, tint and indexed on a font colour", async () => {
    const xml = await stylesXml({
      name: "S",
      cells: cellMap([
        ["0,0", { value: "x", style: { font: { color: { theme: 1, tint: 0.5 } } } }],
      ]),
    })

    expect(xml).toContain('theme="1"')
    expect(xml).toContain('tint="0.5"')
  })

  it("distinguishes two fonts that differ only by indexed colour", async () => {
    // The dedup key has to include every colour field or the second cell
    // silently inherits the first cell's colour.
    const xml = await stylesXml({
      name: "S",
      cells: cellMap([
        ["0,0", { value: "a", style: { font: { color: { indexed: 10 } } } }],
        ["0,1", { value: "b", style: { font: { color: { indexed: 12 } } } }],
      ]),
    })

    expect(xml).toContain('<color indexed="10"/>')
    expect(xml).toContain('<color indexed="12"/>')
  })

  it("writes the full alignment attribute set", async () => {
    const xml = await stylesXml({
      name: "S",
      cells: cellMap([
        [
          "0,0",
          {
            value: "x",
            style: {
              alignment: {
                shrinkToFit: true,
                textRotation: 45,
                indent: 2,
                readingOrder: "rtl",
              },
            },
          },
        ],
      ]),
    })

    expect(xml).toContain('shrinkToFit="true"')
    expect(xml).toContain('textRotation="45"')
    expect(xml).toContain('indent="2"')
    expect(xml).toContain('readingOrder="2"')
  })

  it("keeps two alignments distinct when only the reading order differs", async () => {
    const xml = await stylesXml({
      name: "S",
      cells: cellMap([
        ["0,0", { value: "a", style: { alignment: { readingOrder: "ltr" } } }],
        ["0,1", { value: "b", style: { alignment: { readingOrder: "context" } } }],
      ]),
    })

    expect(xml).toContain('readingOrder="1"')
    expect(xml).toContain('readingOrder="0"')
  })

  it("writes protection hidden=0 when the formula is explicitly not hidden", async () => {
    const xml = await stylesXml({
      name: "S",
      cells: cellMap([
        ["0,0", { value: "x", style: { protection: { locked: true, hidden: false } } }],
      ]),
    })

    expect(xml).toContain('<protection locked="1" hidden="0"/>')
  })

  it("reuses one border definition across two different cell formats", async () => {
    // Same border, different fonts: the xf cache misses but the border
    // cache must hit, or styles.xml grows a duplicate `<border>`.
    const thin = { style: "thin" as const, color: { rgb: "FF000000" } }
    const xml = await stylesXml({
      name: "S",
      cells: cellMap([
        ["0,0", { value: "a", style: { border: { top: thin }, font: { bold: true } } }],
        ["0,1", { value: "b", style: { border: { top: thin }, font: { italic: true } } }],
      ]),
    })

    expect(xml).toContain('<borders count="2">')
  })

  it("gives a gradient fill without an explicit degree a stable key", async () => {
    const xml = await stylesXml({
      name: "S",
      cells: cellMap([
        [
          "0,0",
          {
            value: "a",
            style: {
              fill: {
                type: "gradient",
                stops: [
                  { position: 0, color: { rgb: "FFFFFF" } },
                  { position: 1, color: { rgb: "000000" } },
                ],
              },
            },
          },
        ],
      ]),
    })

    expect(xml).toContain("<gradientFill")
  })
})

describe("dxf (conditional formatting) styles", () => {
  it("serializes numFmt, border and alignment inside a <dxf>", async () => {
    const style: CellStyle = {
      numFmt: "0.00%",
      border: { bottom: { style: "double", color: { rgb: "FFFF0000" } } },
      alignment: { horizontal: "center" },
      fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
    }
    const xml = await stylesXml({
      name: "S",
      rows: [[1]],
      conditionalRules: [
        { type: "cellIs", range: "A1", operator: "greaterThan", formula: "0", priority: 1, style },
      ],
    })

    const dxf = xml.slice(xml.indexOf("<dxfs"))
    expect(dxf).toContain('formatCode="0.00%"')
    expect(dxf).toContain("<border")
    expect(dxf).toContain('<alignment horizontal="center"/>')
  })

  it("emits a single <dxf> when two rules share one style", async () => {
    const style: CellStyle = { font: { bold: true } }
    const xml = await stylesXml({
      name: "S",
      rows: [[1]],
      conditionalRules: [
        { type: "cellIs", range: "A1", operator: "greaterThan", formula: "0", priority: 1, style },
        {
          type: "cellIs",
          range: "A2",
          operator: "lessThan",
          formula: "0",
          priority: 2,
          style: { ...style },
        },
      ],
    })

    expect(xml).toContain('<dxfs count="1">')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// stream-writer (incremental XlsxStreamWriter + streaming writeXlsxStream)
// ═══════════════════════════════════════════════════════════════════════

describe("XlsxStreamWriter cell serialization", () => {
  const styledColumns = [{ key: "a", style: { font: { bold: true } } }]

  it("skips a row whose every cell is empty", async () => {
    const w = new XlsxStreamWriter({ name: "S" })
    w.addRow(["kept"])
    w.addRow([null, null])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toContain('<row r="1">')
    expect(xml).not.toContain('<row r="2"')
  })

  it("keeps a styled empty cell and a styled non-finite number", async () => {
    const w = new XlsxStreamWriter({ name: "S", columns: styledColumns })
    w.addRow([null])
    w.addRow([Number.POSITIVE_INFINITY])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toMatch(/<row r="1"><c r="A1" s="\d+"\/><\/row>/)
    expect(xml).toMatch(/<row r="2"><c r="A2" s="\d+"\/><\/row>/)
  })

  it("keeps the style index on shared-string and boolean cells", async () => {
    const w = new XlsxStreamWriter({ name: "S", columns: styledColumns })
    w.addRow(["text"])
    w.addRow([true])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toMatch(/<c r="A1" t="s" s="\d+">/)
    expect(xml).toMatch(/<c r="A2" t="b" s="\d+">/)
  })

  it("drops a value that is not a cell value at all", async () => {
    const w = new XlsxStreamWriter({ name: "S" })
    w.addRow([{ nested: true } as never, "kept"])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).not.toContain('r="A1"')
    expect(xml).toContain('<c r="B1"')
  })

  it("writes null for a column with no key, and for a key the object lacks", async () => {
    const w = new XlsxStreamWriter({
      name: "S",
      columns: [{ key: "a", header: "A" }, { header: "Spacer" }, { key: "missing", header: "M" }],
    })
    w.addObject({ a: 1 })
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toContain('<row r="2"><c r="A2"><v>1</v></c></row>')
  })

  it("takes the header row from the column keys when a header is missing", async () => {
    const w = new XlsxStreamWriter({
      name: "S",
      columns: [{ key: "a", header: "A" }, { key: "b" }, {}],
    })
    w.addRow([1, 2, 3])
    const buf = await w.finish()
    const shared = await part(buf, "xl/sharedStrings.xml")

    // Header cells: the declared header, then the bare key. The third
    // column declares neither and contributes no cell.
    expect(shared).toContain("<t>A</t>")
    expect(shared).toContain("<t>b</t>")
  })

  it("merges a column numFmt into a column style that has none", async () => {
    const w = new XlsxStreamWriter({
      name: "S",
      columns: [{ key: "amount", numFmt: "#,##0.00", style: { font: { bold: true } } }],
    })
    w.addRow([12.5])
    const styles = await part(await w.finish(), "xl/styles.xml")

    expect(styles).toContain('formatCode="#,##0.00"')
    expect(styles).toContain("<b/>")
  })

  it("leaves a column's own date format on a Date value", async () => {
    const w = new XlsxStreamWriter({
      name: "S",
      columns: [{ key: "when", style: { numFmt: "mmm yyyy" } }],
    })
    w.addRow([new Date(Date.UTC(2024, 0, 15))])
    const styles = await part(await w.finish(), "xl/styles.xml")

    expect(styles).toContain('formatCode="mmm yyyy"')
    expect(styles).not.toContain('formatCode="yyyy-mm-dd"')
  })

  it("truncates a rolled-over sheet name back to 31 characters", async () => {
    // "_2" pushes a legal 30-character base over Excel's limit. See #364.
    const base = "Quarterly Revenue By Region Co" // 30 chars
    const w = new XlsxStreamWriter({ name: base, maxRowsPerSheet: 2 })
    w.addRow(["h"])
    w.addRow([1])
    w.addRow([2])
    const wb = await part(await w.finish(), "xl/workbook.xml")

    expect(wb).toContain(`name="${base}"`)
    // The base is shortened, not the suffix — the sheet number survives.
    expect(wb).toContain('name="Quarterly Revenue By Region C_2"')
  })

  it("refuses addObject when no columns were declared", () => {
    const w = new XlsxStreamWriter({ name: "S" })
    expect(() => w.addObject({ a: 1 })).toThrow(/columns with key accessors/)
  })
})

describe("XlsxStreamWriter freeze panes and column properties", () => {
  it("freezes rows only, leaving the active pane at bottom-left", async () => {
    const w = new XlsxStreamWriter({ name: "S", freezePane: { rows: 1 } })
    w.addRow(["h"])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toContain('ySplit="1"')
    expect(xml).not.toContain("xSplit")
    expect(xml).toContain('activePane="bottomLeft"')
  })

  it("freezes columns only, putting the active pane at top-right", async () => {
    const w = new XlsxStreamWriter({ name: "S", freezePane: { columns: 2 } })
    w.addRow(["h"])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toContain('xSplit="2"')
    expect(xml).toContain('activePane="topRight"')
  })

  it("freezes both axes at once", async () => {
    const w = new XlsxStreamWriter({ name: "S", freezePane: { rows: 1, columns: 1 } })
    w.addRow(["h"])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toContain('activePane="bottomRight"')
    expect(xml).toContain('topLeftCell="B2"')
  })

  it("writes hidden and outlineLevel on <col>", async () => {
    const w = new XlsxStreamWriter({
      name: "S",
      columns: [
        { key: "a", hidden: true },
        { key: "b", outlineLevel: 2 },
      ],
    })
    w.addRow([1, 2])
    const xml = await part(await w.finish(), "xl/worksheets/sheet1.xml")

    expect(xml).toContain('<col min="1" max="1" hidden="true"/>')
    expect(xml).toContain('<col min="2" max="2" outlineLevel="2"/>')
  })
})

describe("writeXlsxStream", () => {
  it("emits a worksheet larger than one chunk in several pieces", async () => {
    // The chunker flushes at 64 KB; a few thousand wide-ish rows crosses it
    // repeatedly, which is the only way the mid-sheet flush path runs.
    const rows: Array<Array<string | number>> = []
    for (let i = 0; i < 4000; i++) {
      rows.push([`row-${i}`, i, `payload-${i}-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx`])
    }
    const buf = await drain(writeXlsxStream(rows, { name: "Big" }))
    const xml = await part(buf, "xl/worksheets/sheet1.xml")

    expect(xml).toContain("row-0")
    expect(xml).toContain("row-3999")
    expect(xml.endsWith("</sheetData></worksheet>")).toBe(true)
  })

  it("emits the sheet opening in its own chunk when the prelude is huge", async () => {
    // 1200 sized columns push the `<cols>` block past the 64 KB flush
    // threshold before the first row is ever serialized.
    const columns = Array.from({ length: 1200 }, (_, i) => ({
      key: `c${i}`,
      width: 12.5,
      hidden: true,
      outlineLevel: 1,
    }))
    const buf = await drain(writeXlsxStream([[1, 2, 3]], { name: "Wide", columns }))
    const xml = await part(buf, "xl/worksheets/sheet1.xml")

    expect(xml).toContain('<col min="1200" max="1200"')
    expect(xml).toContain("<sheetData>")
  })

  it("flushes a header row that alone crosses the chunk threshold", async () => {
    // Wide reports (a column per day, say) can carry a header row larger
    // than the 64 KB flush window all by itself.
    const columns = Array.from({ length: 300 }, (_, i) => ({
      key: `c${i}`,
      header: `Header ${i} `.padEnd(300, "-"),
    }))
    const buf = await drain(writeXlsxStream([[1]], { name: "Wide", columns }))
    const xml = await part(buf, "xl/worksheets/sheet1.xml")

    expect(xml).toContain("Header 0 ")
    expect(xml).toContain("Header 299 ")
    expect(xml.endsWith("</sheetData></worksheet>")).toBe(true)
  })

  it("fails the stream when object rows arrive without column definitions", async () => {
    const stream = writeXlsxStream([{ a: 1 }], { name: "S" })
    await expect(drain(stream)).rejects.toThrow(/columns with key accessors/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// writer.ts — pivot source collection and table range computation
// ═══════════════════════════════════════════════════════════════════════

describe("pivot source rows from object data", () => {
  it("names pivot fields from the header, then the key, and blanks the rest", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Data",
          columns: [
            { key: "region", header: "Region" },
            { key: "notes" },
            { header: "Spacer" },
            {},
            { key: "revenue", header: "Revenue" },
          ],
          data: [
            { region: "EU", revenue: 100 },
            { region: "US", revenue: 200 },
          ],
        },
        {
          name: "Pivot",
          pivotTables: [
            {
              name: "P",
              sourceSheet: "Data",
              rows: ["Region"],
              values: [{ field: "Revenue" }],
            },
          ],
        },
      ],
    })
    const cache = await part(buf, "xl/pivotCache/pivotCacheDefinition1.xml")
    const records = await part(buf, "xl/pivotCache/pivotCacheRecords1.xml")

    expect(cache).toContain('<cacheField name="notes"')
    expect(cache).toContain('<cacheField name=""')
    // The key-less columns contribute a blank `<m/>` in every record.
    expect(records).toContain("<m/>")
  })

  it("rejects a pivot whose source sheet holds no row-shaped data", async () => {
    await expect(
      writeXlsx({
        sheets: [
          { name: "Data", columns: [{ key: "a", header: "A" }] },
          {
            name: "Pivot",
            pivotTables: [{ name: "P", sourceSheet: "Data", values: [{ field: "A" }] }],
          },
        ],
      }),
    ).rejects.toThrow(/at least a header row plus one data row/)
  })
})

describe("auto-computed table ranges", () => {
  it("counts the header row when the sheet is object data", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          columns: [
            { key: "a", header: "A" },
            { key: "b", header: "B" },
          ],
          data: [
            { a: 1, b: 2 },
            { a: 3, b: 4 },
          ],
          tables: [{ name: "T", columns: [{ name: "A" }, { name: "B" }] }],
        },
      ],
    })

    // 2 data rows + 1 header row.
    expect(await part(buf, "xl/tables/table1.xml")).toContain('ref="A1:B3"')
  })

  it("adds a row for the totals row", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            ["A", "B"],
            [1, 2],
          ],
          tables: [
            {
              name: "T",
              showTotalRow: true,
              columns: [{ name: "A" }, { name: "B", totalFunction: "sum" }],
            },
          ],
        },
      ],
    })

    expect(await part(buf, "xl/tables/table1.xml")).toContain('ref="A1:B3"')
  })

  it("counts only the data rows when the columns declare no headers", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          columns: [{ key: "a" }, { key: "b" }],
          data: [
            { a: 1, b: 2 },
            { a: 3, b: 4 },
          ],
          tables: [{ name: "T", columns: [{ name: "A" }, { name: "B" }] }],
        },
      ],
    })

    expect(await part(buf, "xl/tables/table1.xml")).toContain('ref="A1:B2"')
  })

  it("never computes an empty range for a table on a sheet with no data", async () => {
    const buf = await writeXlsx({
      sheets: [{ name: "S", tables: [{ name: "T", columns: [{ name: "A" }, { name: "B" }] }] }],
    })

    expect(await part(buf, "xl/tables/table1.xml")).toContain('ref="A1:B1"')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// table-writer — degenerate ranges
// ═══════════════════════════════════════════════════════════════════════

describe("writeTable range handling", () => {
  it("writes an empty ref when the definition carries no range", () => {
    const { tableXml } = writeTable({ name: "T", columns: [{ name: "A" }] }, 1, 1)
    expect(tableXml).toContain('ref=""')
  })

  it("leaves a single-cell range alone when trimming the totals row", () => {
    const { tableXml } = writeTable(
      { name: "T", range: "A1", showTotalRow: true, columns: [{ name: "A" }] },
      1,
      1,
    )
    expect(tableXml).toContain('<autoFilter ref="A1"/>')
  })

  it("refuses to trim a totals row off a single-row table", () => {
    // "A1:B1" minus one row would be "A1:B0", which Excel rejects.
    const { tableXml } = writeTable(
      { name: "T", range: "A1:B1", showTotalRow: true, columns: [{ name: "A" }, { name: "B" }] },
      1,
      1,
    )
    expect(tableXml).toContain('<autoFilter ref="A1:B1"/>')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// content-types-writer / doc-props-writer
// ═══════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════
// drawing-writer
// ═══════════════════════════════════════════════════════════════════════

describe("chart graphic frame metadata", () => {
  it("puts frameTitle on the drawing shape, next to the alt text", async () => {
    // Screen readers announce `title`; `descr` carries the long
    // description. Both live on the drawing, not on the chart part.
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            ["Q", "Rev"],
            ["Q1", 100],
          ],
          charts: [
            {
              type: "column",
              frameTitle: "Revenue chart",
              altText: "Column chart of revenue by quarter",
              series: [{ name: "Rev", values: "B2:B2", categories: "A2:A2" }],
              anchor: { from: { row: 3, col: 0 } },
            },
          ],
        },
      ],
    })
    const drawing = await part(buf, "xl/drawings/drawing1.xml")

    expect(drawing).toContain('title="Revenue chart"')
    expect(drawing).toContain('descr="Column chart of revenue by quarter"')
  })
})

describe("writeContentTypes defaults", () => {
  it("omits the sharedStrings override when the flag is left out", () => {
    const xml = writeContentTypes({ sheetCount: 1, hasSharedStrings: false })
    expect(xml).toContain("/xl/worksheets/sheet1.xml")
    expect(xml).not.toContain("sharedStrings")
  })

  it("still accepts the positional (sheetCount, hasSharedStrings) signature", () => {
    // The older call shape is still in use inside the streaming writers.
    expect(writeContentTypes(2)).not.toContain("sharedStrings")
    expect(writeContentTypes(2, true)).toContain("/xl/sharedStrings.xml")
    expect(writeContentTypes(2)).toContain("/xl/worksheets/sheet2.xml")
  })
})

describe("writeCustomProperties", () => {
  it("returns null when the custom bag is present but empty", () => {
    expect(writeCustomProperties({ custom: {} })).toBeNull()
  })

  it("returns null when every custom value is of an unsupported type", () => {
    // Values arriving from JSON can be null or objects; those have no
    // `vt:` element, so the part must not be emitted at all.
    expect(writeCustomProperties({ custom: { junk: null as never } })).toBeNull()
  })

  it("skips the unsupported entries but keeps the rest", () => {
    const xml = writeCustomProperties({ custom: { junk: null as never, ok: "yes" } })
    expect(xml).toContain("<vt:lpwstr>yes</vt:lpwstr>")
    expect(xml).not.toContain("junk")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// xml/data-writer
// ═══════════════════════════════════════════════════════════════════════

describe("writeXml mixed content and holes", () => {
  it("skips a key whose value is undefined", () => {
    const out = writeXml([{ sku: "P1", note: undefined as never }])
    expect(out).toContain("<sku>P1</sku>")
    expect(out).not.toContain("note")
  })

  it("writes the #text key as the element's own text", () => {
    const out = writeXml([{ "@id": 7, "#text": "plain body" }], { rowTag: "item" })
    expect(out).toContain('<item id="7">plain body</item>')
  })

  it("keeps the text alongside child elements in mixed content", () => {
    const out = writeXml([{ "#text": "leading", child: "value" }], { rowTag: "item", pretty: true })
    expect(out).toContain("<child>value</child>")
    expect(out).toContain("leading")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// worker.ts — structured-clone-safe workbook transfer
// ═══════════════════════════════════════════════════════════════════════

describe("worker serialization round trip", () => {
  it("carries rich text, image metadata, conditional rules, a11y and external links", () => {
    const cells = new Map<string, Cell>([
      [
        "0,0",
        {
          value: "styled",
          type: "richText",
          richText: [{ text: "bold", font: { bold: true } }],
        },
      ],
    ])
    const sheet: Sheet = {
      name: "S",
      rows: [["styled"]],
      cells,
      conditionalRules: [
        { type: "cellIs", range: "A1", operator: "greaterThan", formula: "0", priority: 1 },
      ],
      a11y: { summary: "Revenue by region", headerRow: 0 },
      images: [
        {
          data: new Uint8Array([1, 2, 3]),
          type: "png",
          anchor: { from: { row: 0, col: 0 } },
          altText: "logo",
          title: "Logo",
        },
      ],
    }
    const wb: Workbook = {
      sheets: [sheet],
      externalLinks: [{ target: "../other.xlsx", sheetNames: ["Sheet1"], sheetData: [] }],
    }

    const round = deserializeWorkbook(serializeWorkbook(wb))

    expect(round.sheets[0].cells?.get("0,0")?.richText).toEqual([
      { text: "bold", font: { bold: true } },
    ])
    expect(round.sheets[0].images?.[0].altText).toBe("logo")
    expect(round.sheets[0].images?.[0].title).toBe("Logo")
    expect(round.sheets[0].conditionalRules).toHaveLength(1)
    expect(round.sheets[0].a11y).toEqual({ summary: "Revenue by region", headerRow: 0 })
    expect(round.externalLinks).toEqual([
      { target: "../other.xlsx", sheetNames: ["Sheet1"], sheetData: [] },
    ])
  })
})
