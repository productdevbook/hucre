import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ParseError } from "../src/errors"
import { REL_CELL_IMAGES } from "../src/xlsx/cell-images-reader"

// ── Package assembly helpers ─────────────────────────────────────────

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

/** A one-pixel PNG stand-in — the reader only ever copies these bytes. */
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

interface RelSpec {
  id: string
  /** Short OOXML type name, or a full URI for vendor namespaces. */
  type: string
  target: string
  mode?: string
}

function relsXml(entries: RelSpec[]): string {
  const items = entries
    .map((e) => {
      const type = e.type.includes("://") ? e.type : `${REL_BASE}/${e.type}`
      const mode = e.mode ? ` TargetMode="${e.mode}"` : ""
      return `<Relationship Id="${e.id}" Type="${type}" Target="${e.target}"${mode}/>`
    })
    .join("")
  return (
    `<?xml version="1.0"?><Relationships ` +
    `xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`
  )
}

// The reader only validates that this part parses; the per-part overrides
// are irrelevant to every branch under test, so one constant serves all.
const CONTENT_TYPES =
  `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/></Types>`

const workbookXml = (body: string): string =>
  `<?xml version="1.0"?><workbook ${NS} ${R}>${body}</workbook>`

const worksheetXml = (body: string): string =>
  `<?xml version="1.0"?><worksheet ${NS} ${R}>${body}</worksheet>`

const ONE_CELL = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>hi</t></is></c></row></sheetData>`

type Parts = Record<string, string | Uint8Array>

/** The five parts every valid single-sheet workbook needs. */
function defaultParts(): Parts {
  return {
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": relsXml([{ id: "rId1", type: "officeDocument", target: "xl/workbook.xml" }]),
    "xl/workbook.xml": workbookXml(
      `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>`,
    ),
    "xl/_rels/workbook.xml.rels": relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
    ]),
    "xl/worksheets/sheet1.xml": worksheetXml(ONE_CELL),
  }
}

async function build(parts: Parts): Promise<Uint8Array> {
  const zip = new ZipWriter()
  for (const [path, content] of Object.entries(parts)) {
    zip.add(path, typeof content === "string" ? enc.encode(content) : content)
  }
  return zip.build()
}

/** Read a package made of the default parts plus `extra` (which may replace them). */
async function read(extra: Parts = {}, options?: Parameters<typeof readXlsx>[1]) {
  return readXlsx(await build({ ...defaultParts(), ...extra }), options)
}

/** Read a package built only from the parts given. */
async function readExact(parts: Parts, options?: Parameters<typeof readXlsx>[1]) {
  return readXlsx(await build(parts), options)
}

// ═══════════════════════════════════════════════════════════════════════
// Package layout
//
// Nothing in OPC requires the workbook to live under `xl/`. Relationship
// targets may be package-absolute (`/workbook.xml`), and a producer that
// puts everything at the root exercises every "no directory prefix"
// branch in the path helpers at once.
// ═══════════════════════════════════════════════════════════════════════

describe("packages that do not use the xl/ layout", () => {
  it("reads a workbook addressed by an absolute target at the package root", async () => {
    const wb = await readExact({
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": relsXml([{ id: "rId1", type: "officeDocument", target: "/workbook.xml" }]),
      "workbook.xml": workbookXml(`<sheets><sheet name="Root" sheetId="1" r:id="rId1"/></sheets>`),
      "_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "sheet1.xml" },
        { id: "rId2", type: "externalLink", target: "link1.xml" },
      ]),
      // NB: no `_rels/sheet1.xml.rels` here — see the skipped test at the
      // bottom of this file for why a root-level part cannot find its own
      // relationships today.
      "sheet1.xml": worksheetXml(ONE_CELL),
      "theme/theme1.xml":
        `<?xml version="1.0"?><a:theme ${A}><a:themeElements><a:clrScheme name="x">` +
        `<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
        `</a:clrScheme></a:themeElements></a:theme>`,
      "link1.xml":
        `<?xml version="1.0"?><externalLink ${NS} ${R}><externalBook r:id="rId1">` +
        `<sheetNames><sheetName val="Budget"/></sheetNames></externalBook></externalLink>`,
      "_rels/link1.xml.rels": relsXml([
        { id: "rId1", type: "externalLinkPath", target: "C:/other.xlsx", mode: "External" },
      ]),
    })

    expect(wb.sheets[0].name).toBe("Root")
    expect(wb.sheets[0].rows[0][0]).toBe("hi")
    expect(wb.themeColors).toBeDefined()
    expect(wb.externalLinks![0].sheetNames).toEqual(["Budget"])
  })

  it("resolves a relative target that walks up out of its own directory", async () => {
    // `xl/_rels/workbook.xml.rels` pointing at `../sheets/sheet1.xml` —
    // legal OPC and produced by a few non-Excel writers.
    const wb = await read({
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "../sheets/./sheet1.xml" },
      ]),
      "sheets/sheet1.xml": worksheetXml(ONE_CELL),
    })
    expect(wb.sheets[0].rows[0][0]).toBe("hi")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Structural failures
// ═══════════════════════════════════════════════════════════════════════

describe("invalid packages", () => {
  it("rejects a package whose root rels name no officeDocument", async () => {
    const parts = defaultParts()
    parts["_rels/.rels"] = relsXml([
      { id: "rId1", type: "extended-properties", target: "docProps/app.xml" },
    ])
    await expect(readExact(parts)).rejects.toThrow(/cannot find workbook relationship/)
  })

  it("rejects a package whose workbook part is missing", async () => {
    const parts = defaultParts()
    delete parts["xl/workbook.xml"]
    await expect(readExact(parts)).rejects.toThrow(/missing workbook at xl\/workbook\.xml/)
  })

  it("rejects a sheet whose worksheet part is missing", async () => {
    const parts = defaultParts()
    delete parts["xl/worksheets/sheet1.xml"]
    await expect(readExact(parts)).rejects.toThrow(ParseError)
  })

  it("rejects a sheet whose rId has no worksheet relationship", async () => {
    await expect(
      read({
        "xl/workbook.xml": workbookXml(
          `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId404"/></sheets>`,
        ),
      }),
    ).rejects.toThrow(/missing worksheet file for sheet "Sheet1"/)
  })

  it("reads a workbook with no workbook.xml.rels at all", async () => {
    // Without the rels part there are no sheet targets, so no sheets —
    // but the file still opens rather than throwing.
    const parts = defaultParts()
    delete parts["xl/_rels/workbook.xml.rels"]
    parts["xl/workbook.xml"] = workbookXml(`<sheets/>`)
    const wb = await readExact(parts)
    expect(wb.sheets).toEqual([])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// workbook.xml
// ═══════════════════════════════════════════════════════════════════════

describe("sheet declarations", () => {
  it("marks hidden and veryHidden sheets", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `<sheets>` +
          `<sheet name="V" sheetId="1" r:id="rId1"/>` +
          `<sheet name="H" sheetId="2" r:id="rId1" state="hidden"/>` +
          `<sheet name="VH" sheetId="3" r:id="rId1" state="veryHidden"/>` +
          `<sheet name="Odd" sheetId="4" r:id="rId1" state="somethingElse"/>` +
          `</sheets>`,
      ),
    })
    expect(wb.sheets.map((s) => [s.hidden, s.veryHidden])).toEqual([
      [undefined, undefined],
      [true, undefined],
      [undefined, true],
      [undefined, undefined],
    ])
  })

  it("skips sheet entries with no name or no relationship id", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `<sheets><sheet sheetId="1" r:id="rId1"/><sheet name="NoRel" sheetId="2"/>` +
          `<sheet name="Good" r:id="rId1"/><notASheet name="X" r:id="rId1"/></sheets>`,
      ),
    })
    expect(wb.sheets.map((s) => s.name)).toEqual(["Good"])
  })

  it("finds the relationship id under any namespace prefix", async () => {
    // The `r` prefix is conventional, not required — the attribute is
    // identified by its namespace, which we approximate by shape.
    const wb = await read({
      "xl/workbook.xml":
        `<?xml version="1.0"?><workbook ${NS} ` +
        `xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<sheets><sheet name="Prefixed" sheetId="1" rel:id="rId1"/></sheets></workbook>`,
    })
    expect(wb.sheets[0].name).toBe("Prefixed")
  })
})

describe("date system", () => {
  const wb1904 = workbookXml(
    `<workbookPr date1904="true"/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`,
  )

  it("auto-detects the 1904 epoch from workbookPr", async () => {
    expect((await read({ "xl/workbook.xml": wb1904 })).dateSystem).toBe("1904")
    expect((await read({ "xl/workbook.xml": wb1904 }, { dateSystem: "auto" })).dateSystem).toBe(
      "1904",
    )
  })

  it("lets an explicit dateSystem option override the file", async () => {
    // Legacy Mac files sometimes carry the flag when the serials are in
    // fact 1900-based, so the option has to win.
    expect((await read({ "xl/workbook.xml": wb1904 }, { dateSystem: "1900" })).dateSystem).toBe(
      "1900",
    )
    expect((await read({}, { dateSystem: "1904" })).dateSystem).toBe("1904")
  })
})

describe("workbook protection", () => {
  it("reads the structure and window locks", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `<workbookProtection lockStructure="true" lockWindows="1"/>` +
          `<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`,
      ),
    })
    expect(wb.workbookProtection).toEqual({ lockStructure: true, lockWindows: true })
  })

  it("omits the block when nothing is actually locked", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `<workbookProtection workbookPassword="CC1A"/>` +
          `<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`,
      ),
    })
    expect(wb.workbookProtection).toBeUndefined()
  })
})

describe("defined names", () => {
  it("scopes a name to its sheet and keeps its comment", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>` +
          `<definedNames>` +
          `<definedName name="Global">S!$A$1</definedName>` +
          `<definedName name="Local" localSheetId="0" comment="notes">S!$B$1</definedName>` +
          `<definedName name="Dangling" localSheetId="9">S!$C$1</definedName>` +
          `<definedName name="Empty"></definedName>` +
          `<definedName localSheetId="0">S!$D$1</definedName>` +
          `</definedNames>`,
      ),
    })
    expect(wb.namedRanges).toEqual([
      { name: "Global", range: "S!$A$1" },
      { name: "Local", range: "S!$B$1", scope: "S", comment: "notes" },
      // A localSheetId past the end of the sheet list resolves to no
      // scope rather than crashing on the array lookup.
      { name: "Dangling", range: "S!$C$1" },
    ])
  })
})

describe("sheets read option", () => {
  const threeSheets = {
    "xl/workbook.xml": workbookXml(
      `<sheets><sheet name="A" sheetId="1" r:id="rId1"/>` +
        `<sheet name="B" sheetId="2" r:id="rId1" state="hidden"/>` +
        `<sheet name="C" sheetId="3" r:id="rId1" state="veryHidden"/></sheets>`,
    ),
  }

  it("treats an empty array as no filter at all", async () => {
    const wb = await read(threeSheets, { sheets: [] })
    expect(wb.sheets.map((s) => s.name)).toEqual(["A", "B", "C"])
  })

  it("ignores indices and names that match nothing", async () => {
    const wb = await read(threeSheets, { sheets: [2, 99, "B", "Nope", -1] })
    expect(wb.sheets.map((s) => s.name)).toEqual(["C", "B"])
  })

  it("passes hidden state to a predicate filter", async () => {
    const seen: Array<[string, number, boolean | undefined, boolean | undefined]> = []
    const wb = await read(threeSheets, {
      sheets: (info, i) => {
        seen.push([info.name, i, info.hidden, info.veryHidden])
        return !info.hidden && !info.veryHidden
      },
    })
    expect(seen).toEqual([
      ["A", 0, false, false],
      ["B", 1, true, false],
      ["C", 2, false, true],
    ])
    expect(wb.sheets.map((s) => s.name)).toEqual(["A"])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Optional workbook-level parts
// ═══════════════════════════════════════════════════════════════════════

describe("relationships that point at parts which are not in the package", () => {
  it("skips every dangling workbook-level relationship without failing", async () => {
    // Files edited by third-party tools routinely keep relationships to
    // parts that were dropped. None of these should abort the read.
    const wb = await read({
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
        { id: "rId2", type: "sharedStrings", target: "sharedStrings.xml" },
        { id: "rId3", type: "styles", target: "styles.xml" },
        { id: "rId4", type: "person", target: "persons/person.xml" },
        { id: "rId5", type: "externalLink", target: "externalLinks/externalLink1.xml" },
        { id: "rIdX", type: "externalLink", target: "externalLinks/externalLink2.xml" },
        { id: "rId6", type: "slicerCache", target: "slicerCaches/slicerCache1.xml" },
        { id: "rId7", type: "slicerCache", target: "slicerCaches/legacy.xml" },
        { id: "rId8", type: "timelineCache", target: "timelineCaches/timelineCache1.xml" },
        {
          id: "rId9",
          type: "pivotCacheDefinition",
          target: "pivotCache/pivotCacheDefinition1.xml",
        },
        { id: "rId10", type: REL_CELL_IMAGES, target: "cellimages.xml" },
      ]),
      "xl/workbook.xml": workbookXml(
        `<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>` +
          `<pivotCaches><pivotCache cacheId="0" r:id="rId9"/>` +
          `<pivotCache cacheId="1" r:id="rIdMissing"/>` +
          `<pivotCache r:id="rId9"/></pivotCaches>`,
      ),
    })
    expect(wb.sheets).toHaveLength(1)
    expect(wb.externalLinks).toBeUndefined()
    expect(wb.pivotCaches).toBeUndefined()
    expect(wb.slicerCaches).toBeUndefined()
    expect(wb.timelineCaches).toBeUndefined()
    expect(wb.cellImages).toBeUndefined()
    expect(wb.persons).toBeUndefined()
  })

  it("skips a worksheet-level relationship whose part is missing", async () => {
    const wb = await read({
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "drawing", target: "../drawings/drawing1.xml" },
        { id: "rId2", type: "comments", target: "../comments1.xml" },
        { id: "rId3", type: "threadedComment", target: "../threadedComments/tc1.xml" },
        { id: "rId4", type: "table", target: "../tables/table1.xml" },
        { id: "rId5", type: "image", target: "../media/bg.png" },
        { id: "rId6", type: "pivotTable", target: "../pivotTables/pivotTable1.xml" },
        { id: "rId7", type: "slicer", target: "../slicers/slicer1.xml" },
        { id: "rId8", type: "timeline", target: "../timelines/timeline1.xml" },
      ]),
    })
    const sheet = wb.sheets[0]
    expect(sheet.images).toBeUndefined()
    expect(sheet.tables).toBeUndefined()
    expect(sheet.backgroundImage).toBeUndefined()
    expect(sheet.pivotTables).toBeUndefined()
    expect(sheet.slicers).toBeUndefined()
    expect(sheet.timelines).toBeUndefined()
    expect(sheet.threadedComments).toBeUndefined()
  })
})

describe("shared strings, styles and background image", () => {
  it("resolves shared strings and a sheet background", async () => {
    const wb = await read({
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
        { id: "rId2", type: "sharedStrings", target: "sharedStrings.xml" },
      ]),
      "xl/sharedStrings.xml": `<?xml version="1.0"?><sst ${NS} count="1" uniqueCount="1"><si><t>Shared</t></si></sst>`,
      "xl/worksheets/sheet1.xml": worksheetXml(
        `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>`,
      ),
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "image", target: "../media/bg.png" },
      ]),
      "xl/media/bg.png": PNG,
    })
    expect(wb.sheets[0].rows[0][0]).toBe("Shared")
    expect(wb.sheets[0].backgroundImage).toEqual(PNG)
  })
})

describe("document properties", () => {
  it("attaches custom properties even with no core or app part", async () => {
    const wb = await read({
      "docProps/custom.xml":
        `<?xml version="1.0"?><Properties ` +
        `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ` +
        `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
        `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Dept">` +
        `<vt:lpwstr>Finance</vt:lpwstr></property></Properties>`,
    })
    expect(wb.properties).toEqual({ custom: { Dept: "Finance" } })
  })

  it("ignores empty docProps parts", async () => {
    const wb = await read({
      "docProps/core.xml":
        `<?xml version="1.0"?><cp:coreProperties ` +
        `xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`,
      "docProps/app.xml":
        `<?xml version="1.0"?><Properties ` +
        `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>`,
      "docProps/custom.xml":
        `<?xml version="1.0"?><Properties ` +
        `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>`,
    })
    expect(wb.properties).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Comments
// ═══════════════════════════════════════════════════════════════════════

describe("legacy comments", () => {
  it("creates a cell for a comment anchored on an empty cell", async () => {
    // Comments outlive their content: deleting the text leaves the note
    // attached to a cell that no longer appears in <sheetData>.
    const wb = await read({
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "comments", target: "../comments1.xml" },
      ]),
      "xl/comments1.xml":
        `<?xml version="1.0"?><comments ${NS}><authors><author>Ana</author></authors>` +
        `<commentList>` +
        `<comment ref="A1" authorId="0"><text><t>on data</t></text></comment>` +
        `<comment ref="D9" authorId="0"><text><t>orphan</t></text></comment>` +
        `</commentList></comments>`,
    })
    const cells = wb.sheets[0].cells!
    expect(cells.get("0,0")!.comment!.text).toBe("on data")
    const orphan = cells.get("8,3")!
    expect(orphan.value).toBeNull()
    expect(orphan.comment!.text).toBe("orphan")
  })

  it("ignores a comments part that lists no comments", async () => {
    const wb = await read({
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "comments", target: "../comments1.xml" },
      ]),
      "xl/comments1.xml": `<?xml version="1.0"?><comments ${NS}><commentList/></comments>`,
    })
    expect(wb.sheets[0].cells).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Tables
// ═══════════════════════════════════════════════════════════════════════

describe("table parts", () => {
  const withTable = (tableXml: string): Parts => ({
    "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
      { id: "rId1", type: "table", target: "../tables/table1.xml" },
    ]),
    "xl/tables/table1.xml": tableXml,
  })

  it("drops a table part with no name", async () => {
    const wb = await read(
      withTable(
        `<?xml version="1.0"?><table ${NS} id="1" ref="A1:B2"><tableColumns count="0"/></table>`,
      ),
    )
    expect(wb.sheets[0].tables).toBeUndefined()
  })

  it("reads columns, totals and style flags", async () => {
    const wb = await read(
      withTable(
        `<?xml version="1.0"?><table ${NS} id="1" name="Sales" displayName="Sales_2024" ` +
          `ref="A1:C4" totalsRowCount="1">` +
          `<autoFilter ref="A1:C3"/>` +
          `<tableColumns count="3">` +
          `<tableColumn id="1" name="Item"/>` +
          `<tableColumn id="2" name="Qty" totalsRowFunction="sum"/>` +
          `<tableColumn id="3" name="Note" totalsRowLabel="Total">` +
          `<totalsRowFormula>SUBTOTAL(109,Sales[Qty])</totalsRowFormula>` +
          `<calculatedColumnFormula>1</calculatedColumnFormula></tableColumn>` +
          `</tableColumns>` +
          `<tableStyleInfo name="TableStyleMedium2" showRowStripes="1" showColumnStripes="0"/>` +
          `</table>`,
      ),
    )
    expect(wb.sheets[0].tables![0]).toEqual({
      name: "Sales",
      displayName: "Sales_2024",
      range: "A1:C4",
      style: "TableStyleMedium2",
      showRowStripes: true,
      showColumnStripes: false,
      showAutoFilter: true,
      showTotalRow: true,
      columns: [
        { name: "Item" },
        { name: "Qty", totalFunction: "sum" },
        { name: "Note", totalLabel: "Total", totalFormula: "SUBTOTAL(109,Sales[Qty])" },
      ],
    })
  })

  it("omits displayName when it merely repeats the name, and ref when absent", async () => {
    const wb = await read(
      withTable(
        `<?xml version="1.0"?><table ${NS} id="1" name="T" displayName="T" totalsRowCount="0">` +
          `<tableColumns count="1"><tableColumn id="1" name="A">` +
          `<totalsRowFormula></totalsRowFormula></tableColumn></tableColumns></table>`,
      ),
    )
    const table = wb.sheets[0].tables![0]
    expect(table.displayName).toBeUndefined()
    expect(table.range).toBeUndefined()
    expect(table.style).toBeUndefined()
    expect(table.showTotalRow).toBeUndefined()
    expect(table.columns[0]).toEqual({ name: "A" })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Drawings: images, textboxes, charts
// ═══════════════════════════════════════════════════════════════════════

/** Wire a drawing part (plus its rels) onto sheet1. */
function withDrawing(drawing: string, drawingRels: RelSpec[], extra: Parts = {}): Parts {
  return {
    "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
      { id: "rId1", type: "drawing", target: "../drawings/drawing1.xml" },
    ]),
    "xl/drawings/drawing1.xml": `<?xml version="1.0"?><xdr:wsDr ${XDR} ${A} ${R}>${drawing}</xdr:wsDr>`,
    "xl/drawings/_rels/drawing1.xml.rels": relsXml(drawingRels),
    ...extra,
  }
}

const pic = (embed: string, meta = ""): string =>
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="p"${meta}/><xdr:cNvPicPr/></xdr:nvPicPr>` +
  `<xdr:blipFill><a:blip ${embed}/></xdr:blipFill></xdr:pic>`

const fromTo = (fc: number, fr: number, tc: number, tr: number): string =>
  `<xdr:from><xdr:col>${fc}</xdr:col><xdr:row>${fr}</xdr:row></xdr:from>` +
  `<xdr:to><xdr:col>${tc}</xdr:col><xdr:row>${tr}</xdr:row></xdr:to>`

describe("drawing images", () => {
  it("reads a oneCellAnchor image with its extent and accessibility metadata", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:oneCellAnchor>` +
          `<xdr:from><xdr:col>2</xdr:col><xdr:row>3</xdr:row></xdr:from>` +
          `<xdr:ext cx="952500" cy="476250"/>` +
          pic(`r:embed="rId1"`, ` descr="A logo" title="Logo"`) +
          `</xdr:oneCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/image1.png" }],
        { "xl/media/image1.png": PNG },
      ),
    )
    expect(wb.sheets[0].images![0]).toEqual({
      data: PNG,
      type: "png",
      anchor: { from: { row: 3, col: 2 } },
      width: 100,
      height: 50,
      altText: "A logo",
      title: "Logo",
    })
  })

  it("omits width and height when the extent is zero or unparseable", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from>` +
          `<xdr:ext cx="0" cy="not-a-number"/>` +
          pic(`r:embed="rId1"`) +
          `</xdr:oneCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/image1.png" }],
        { "xl/media/image1.png": PNG },
      ),
    )
    const img = wb.sheets[0].images![0]
    expect(img.width).toBeUndefined()
    expect(img.height).toBeUndefined()
  })

  it("falls back to png for a media file with an unknown extension", async () => {
    // Some producers store `image1.bin` or drop the extension entirely.
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}${pic(`r:embed="rId1"`)}</xdr:twoCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/image1.bin" }],
        { "xl/media/image1.bin": PNG },
      ),
    )
    expect(wb.sheets[0].images![0].type).toBe("png")
  })

  it("finds the embed id under a non-standard namespace prefix", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}` +
          pic(`rel:embed="rId1" xmlns:rel="${REL_BASE}"`) +
          `</xdr:twoCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/image1.jpeg" }],
        { "xl/media/image1.jpeg": PNG },
      ),
    )
    expect(wb.sheets[0].images![0].type).toBe("jpeg")
  })

  it("drops an anchor whose embed id resolves to no relationship", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}${pic(`r:embed="rId404"`)}</xdr:twoCellAnchor>` +
          `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from>` +
          `${pic(`r:embed="rId404"`)}</xdr:oneCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/image1.png" }],
        { "xl/media/image1.png": PNG },
      ),
    )
    expect(wb.sheets[0].images).toBeUndefined()
  })

  it("drops an anchor with no picture at all", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}<xdr:clientData/></xdr:twoCellAnchor>` +
          `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from>` +
          `<xdr:clientData/></xdr:oneCellAnchor>`,
        [],
      ),
    )
    expect(wb.sheets[0].images).toBeUndefined()
  })

  it("drops an image whose media file is missing from the package", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}${pic(`r:embed="rId1"`)}</xdr:twoCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/gone.png" }],
      ),
    )
    expect(wb.sheets[0].images).toBeUndefined()
  })

  it("ignores a drawing relationship whose part is not in the package", async () => {
    const wb = await read({
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "drawing", target: "../drawings/drawing1.xml" },
      ]),
    })
    expect(wb.sheets[0].images).toBeUndefined()
  })

  it("reads a drawing that has no rels file of its own", async () => {
    // No rels means no resolvable embeds, so the anchors yield nothing.
    const parts = withDrawing(
      `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}${pic(`r:embed="rId1"`)}</xdr:twoCellAnchor>`,
      [],
    )
    delete parts["xl/drawings/_rels/drawing1.xml.rels"]
    const wb = await read(parts)
    expect(wb.sheets[0].images).toBeUndefined()
  })
})

describe("drawing textboxes", () => {
  const txBody = (runs: string): string => `<xdr:txBody><a:bodyPr/><a:p>${runs}</a:p></xdr:txBody>`

  it("reads text, run styling and shape colours", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(1, 1, 4, 6)}<xdr:sp>` +
          `<xdr:nvSpPr><xdr:cNvPr id="2" name="tb" descr="Note" title="A note"/>` +
          `<xdr:cNvSpPr txBox="true"/></xdr:nvSpPr>` +
          `<xdr:spPr><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>` +
          `<a:ln><a:solidFill><a:srgbClr val="000080"/></a:solidFill></a:ln></xdr:spPr>` +
          txBody(
            `<a:r><a:rPr sz="1400" b="true"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
              `</a:rPr><a:t>Hello</a:t></a:r>`,
          ) +
          `</xdr:sp></xdr:twoCellAnchor>`,
        [],
      ),
    )
    expect(wb.sheets[0].textBoxes![0]).toEqual({
      text: "Hello",
      anchor: { from: { row: 1, col: 1 }, to: { row: 6, col: 4 } },
      altText: "Note",
      title: "A note",
      style: {
        fontSize: 14,
        bold: true,
        color: "FF0000",
        fillColor: "FFFF00",
        borderColor: "000080",
      },
    })
  })

  it("drops a textbox shape that holds no text", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}<xdr:sp>` +
          `<xdr:nvSpPr><xdr:cNvPr id="2" name="tb"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>` +
          `<xdr:txBody><a:bodyPr/><a:p/></xdr:txBody></xdr:sp></xdr:twoCellAnchor>`,
        [],
      ),
    )
    expect(wb.sheets[0].textBoxes).toBeUndefined()
  })

  it("ignores a shape that is not marked as a textbox", async () => {
    // A plain autoshape (`txBox` absent) is not a textbox, so neither
    // the textbox nor the image path claims it.
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}<xdr:sp>` +
          `<xdr:nvSpPr><xdr:cNvPr id="2" name="s"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
          txBody(`<a:r><a:t>x</a:t></a:r>`) +
          `</xdr:sp></xdr:twoCellAnchor>`,
        [],
      ),
    )
    expect(wb.sheets[0].textBoxes).toBeUndefined()
  })

  it("joins multiple paragraphs with newlines and keeps the first run's style", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 2, 2)}<xdr:sp>` +
          `<xdr:nvSpPr><xdr:cNvPr id="2" name="tb"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>` +
          `<xdr:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill>` +
          `<a:ln><a:noFill/></a:ln></xdr:spPr>` +
          `<xdr:txBody><a:bodyPr/>` +
          `<a:p><a:r><a:rPr sz="1200"/><a:t>one</a:t></a:r></a:p>` +
          `<a:p><a:r><a:rPr sz="9999"/><a:t>two</a:t></a:r></a:p>` +
          `<a:endParaRPr/></xdr:txBody></xdr:sp></xdr:twoCellAnchor>`,
        [],
      ),
    )
    const tb = wb.sheets[0].textBoxes![0]
    expect(tb.text).toBe("one\ntwo")
    // A theme fill has no srgbClr to read, so no fillColor is surfaced.
    expect(tb.style).toEqual({ fontSize: 12 })
  })
})

describe("drawing charts", () => {
  const chartFrame = (rid: string): string =>
    `<xdr:graphicFrame><a:graphic><a:graphicData ` +
    `uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ` +
    `r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`

  const barChart =
    `<?xml version="1.0"?><c:chartSpace ` +
    `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${A}><c:chart><c:plotArea>` +
    `<c:barChart><c:ser><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>` +
    `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val></c:ser></c:barChart>` +
    `</c:plotArea></c:chart></c:chartSpace>`

  it("pins a twoCellAnchor chart to its cell range", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 7, 15)}${chartFrame("rId1")}</xdr:twoCellAnchor>`,
        [{ id: "rId1", type: "chart", target: "../charts/chart1.xml" }],
        { "xl/charts/chart1.xml": barChart },
      ),
    )
    expect(wb.sheets[0].charts![0].anchor).toEqual({
      from: { row: 0, col: 0 },
      to: { row: 15, col: 7 },
    })
  })

  it("reads a oneCellAnchor chart with a from-only anchor", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:row>2</xdr:row></xdr:from>` +
          `<xdr:ext cx="1" cy="1"/>${chartFrame("rId1")}</xdr:oneCellAnchor>`,
        [{ id: "rId1", type: "chart", target: "../charts/chart1.xml" }],
        { "xl/charts/chart1.xml": barChart },
      ),
    )
    expect(wb.sheets[0].charts![0].anchor).toEqual({ from: { row: 2, col: 3 } })
  })

  it("leaves an absoluteAnchor chart unanchored rather than inventing A1", async () => {
    // absoluteAnchor positions in EMU, so there is no cell to report.
    const wb = await read(
      withDrawing(
        `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/>` +
          `${chartFrame("rId1")}</xdr:absoluteAnchor>`,
        [{ id: "rId1", type: "chart", target: "../charts/chart1.xml" }],
        { "xl/charts/chart1.xml": barChart },
      ),
    )
    expect(wb.sheets[0].charts![0].anchor).toBeUndefined()
  })

  it("reports no anchor for a twoCellAnchor missing its <from>", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor><xdr:to><xdr:col>3</xdr:col><xdr:row>3</xdr:row></xdr:to>` +
          `${chartFrame("rId1")}</xdr:twoCellAnchor>`,
        [{ id: "rId1", type: "chart", target: "../charts/chart1.xml" }],
        { "xl/charts/chart1.xml": barChart },
      ),
    )
    expect(wb.sheets[0].charts![0].anchor).toBeUndefined()
  })

  it("skips a chart whose part is missing or unparseable", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}${chartFrame("rId1")}</xdr:twoCellAnchor>` +
          `<xdr:twoCellAnchor>${fromTo(2, 2, 3, 3)}${chartFrame("rId2")}</xdr:twoCellAnchor>` +
          `<xdr:twoCellAnchor>${fromTo(4, 4, 5, 5)}${chartFrame("rId404")}</xdr:twoCellAnchor>`,
        [
          { id: "rId1", type: "chart", target: "../charts/gone.xml" },
          { id: "rId2", type: "chart", target: "../charts/chart2.xml" },
        ],
        { "xl/charts/chart2.xml": `<?xml version="1.0"?><notAChart/>` },
      ),
    )
    expect(wb.sheets[0].charts).toBeUndefined()
  })

  it("keeps only the first reference to a chart part shared by two anchors", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}${chartFrame("rId1")}</xdr:twoCellAnchor>` +
          `<xdr:twoCellAnchor>${fromTo(5, 5, 6, 6)}${chartFrame("rId1")}</xdr:twoCellAnchor>`,
        [{ id: "rId1", type: "chart", target: "../charts/chart1.xml" }],
        { "xl/charts/chart1.xml": barChart },
      ),
    )
    expect(wb.sheets[0].charts).toHaveLength(1)
    expect(wb.sheets[0].charts![0].anchor!.from).toEqual({ row: 0, col: 0 })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// WPS cell-embedded images (xl/cellimages.xml)
//
// The part backing `=_xlfn.DISPIMG("id", 1)`. Its rels file mixes image
// entries with whatever else the producer attached, so the media walk
// has to be selective.
// ═══════════════════════════════════════════════════════════════════════

describe("cell images", () => {
  const cellImage = (name: string, embed: string, descr = ""): string =>
    `<etc:cellImage><xdr:pic><xdr:nvPicPr>` +
    `<xdr:cNvPr id="1" name="${name}"${descr ? ` descr="${descr}"` : ""}/><xdr:cNvPicPr/>` +
    `</xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${embed}"/></xdr:blipFill></xdr:pic></etc:cellImage>`

  it("resolves each DISPIMG entry to its media bytes", async () => {
    const wb = await read({
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
        { id: "rId2", type: REL_CELL_IMAGES, target: "cellimages.xml" },
      ]),
      "xl/cellimages.xml":
        `<?xml version="1.0"?><etc:cellImages ` +
        `xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData" ${XDR} ${A} ${R}>` +
        cellImage("ID_OK", "rId1", "A photo") +
        // Points at a rel whose media file is not in the package.
        cellImage("ID_GONE", "rId2") +
        // Points at a rel whose extension is not a known image type.
        cellImage("ID_ODD", "rId3") +
        // Points at a rel that is not an image relationship at all.
        cellImage("ID_NOTIMAGE", "rId4") +
        `</etc:cellImages>`,
      "xl/_rels/cellimages.xml.rels": relsXml([
        { id: "rId1", type: "image", target: "media/ci1.png" },
        { id: "rId2", type: "image", target: "media/missing.png" },
        { id: "rId3", type: "image", target: "media/thing.dat" },
        { id: "rId4", type: "hyperlink", target: "https://example.com" },
      ]),
      "xl/media/ci1.png": PNG,
      "xl/media/thing.dat": PNG,
    })
    expect(wb.cellImages).toEqual([{ id: "ID_OK", data: PNG, type: "png", description: "A photo" }])
  })

  it("omits the workbook field when no entry resolves", async () => {
    const wb = await read({
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
        { id: "rId2", type: REL_CELL_IMAGES, target: "cellimages.xml" },
      ]),
      // No sibling rels file, so no embed id can be resolved to media.
      "xl/cellimages.xml":
        `<?xml version="1.0"?><etc:cellImages ` +
        `xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData" ${XDR} ${A} ${R}>` +
        cellImage("ID_A", "rId1") +
        `</etc:cellImages>`,
    })
    expect(wb.cellImages).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Pivot caches and pivot tables
// ═══════════════════════════════════════════════════════════════════════

describe("pivot wiring", () => {
  const CACHE_DEF =
    `<?xml version="1.0"?><pivotCacheDefinition ${NS} ${R} recordCount="3">` +
    `<cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Sheet1"/></cacheSource>` +
    `<cacheFields count="2"><cacheField name="Region"/><cacheField name="Amount"/></cacheFields>` +
    `</pivotCacheDefinition>`

  const PIVOT_TABLE =
    `<?xml version="1.0"?><pivotTableDefinition ${NS} name="PivotTable1" cacheId="5">` +
    `<location ref="A3:B6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>` +
    `<pivotFields count="2"><pivotField axis="axisRow"/><pivotField dataField="1"/></pivotFields>` +
    `</pivotTableDefinition>`

  it("joins a pivot table to its cache through the two rels files", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>` +
          // The first ref is dangling on purpose: the pivot-table lookup
          // walks every ref and must skip the ones with no relationship.
          `<pivotCaches><pivotCache cacheId="0" r:id="rIdGone"/>` +
          `<pivotCache cacheId="5" r:id="rId9"/></pivotCaches>`,
      ),
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
        {
          id: "rId9",
          type: "pivotCacheDefinition",
          target: "pivotCache/pivotCacheDefinition1.xml",
        },
      ]),
      "xl/pivotCache/pivotCacheDefinition1.xml": CACHE_DEF,
      "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels": relsXml([
        { id: "rId1", type: "pivotCacheRecords", target: "pivotCacheRecords1.xml" },
      ]),
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "pivotTable", target: "../pivotTables/pivotTable1.xml" },
      ]),
      "xl/pivotTables/pivotTable1.xml": PIVOT_TABLE,
      "xl/pivotTables/_rels/pivotTable1.xml.rels": relsXml([
        {
          id: "rId1",
          type: "pivotCacheDefinition",
          target: "../pivotCache/pivotCacheDefinition1.xml",
        },
      ]),
    })

    expect(wb.pivotCaches).toHaveLength(1)
    expect(wb.pivotCaches![0]).toMatchObject({ cacheId: 5, hasRecords: true, sourceRef: "A1:B4" })
    const pivot = wb.sheets[0].pivotTables![0]
    expect(pivot.cacheId).toBe(5)
    // The cache's real field names replace parsePivotTable's placeholders.
    expect(pivot.fields.map((f) => f.name)).toEqual(["Region", "Amount"])
  })

  it("skips cache and table parts whose XML is not what the relationship claims", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>` +
          `<pivotCaches><pivotCache cacheId="1" r:id="rId9"/></pivotCaches>`,
      ),
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
        {
          id: "rId9",
          type: "pivotCacheDefinition",
          target: "pivotCache/pivotCacheDefinition1.xml",
        },
      ]),
      "xl/pivotCache/pivotCacheDefinition1.xml": `<?xml version="1.0"?><somethingElse/>`,
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "pivotTable", target: "../pivotTables/pivotTable1.xml" },
      ]),
      "xl/pivotTables/pivotTable1.xml": `<?xml version="1.0"?><notAPivot/>`,
    })
    expect(wb.pivotCaches).toBeUndefined()
    expect(wb.sheets[0].pivotTables).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Whitespace tolerance
//
// Excel writes these parts with no whitespace between elements, but
// pretty-printed packages are common from other producers and from
// anyone who has opened the XML in an editor.
// ═══════════════════════════════════════════════════════════════════════

describe("pretty-printed parts", () => {
  it("ignores indentation between defined names and table columns", async () => {
    const wb = await read({
      "xl/workbook.xml": workbookXml(
        `\n  <sheets>\n    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>\n  </sheets>\n` +
          `  <definedNames>\n    <definedName name="Q1">Sheet1!$A$1</definedName>\n  </definedNames>\n`,
      ),
      "xl/worksheets/_rels/sheet1.xml.rels": relsXml([
        { id: "rId1", type: "table", target: "../tables/table1.xml" },
      ]),
      "xl/tables/table1.xml":
        `<?xml version="1.0"?>\n<table ${NS} id="1" name="T" ref="A1:A2">\n` +
        `  <tableColumns count="1">\n    <tableColumn id="1" name="A">\n` +
        `      <totalsRowFormula>SUM(T[A])</totalsRowFormula>\n    </tableColumn>\n` +
        `  </tableColumns>\n</table>`,
    })
    expect(wb.namedRanges).toEqual([{ name: "Q1", range: "Sheet1!$A$1" }])
    expect(wb.sheets[0].tables![0].columns[0].totalFormula).toBe("SUM(T[A])")
  })

  it("ignores indentation inside a drawing textbox", async () => {
    const wb = await read(
      withDrawing(
        `\n  <xdr:twoCellAnchor>\n    ${fromTo(0, 0, 2, 2)}\n    <xdr:sp>\n` +
          `      <xdr:nvSpPr><xdr:cNvPr id="2" name="tb"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>\n` +
          `      <xdr:txBody>\n        <a:bodyPr/>\n        <a:p>\n` +
          `          <a:r><a:t>spaced</a:t></a:r>\n        </a:p>\n      </xdr:txBody>\n` +
          `    </xdr:sp>\n  </xdr:twoCellAnchor>\n`,
        [],
      ),
    )
    expect(wb.sheets[0].textBoxes![0].text).toBe("spaced")
  })
})

describe("more relationship shapes", () => {
  it("resolves a package-absolute worksheet target", async () => {
    // `Target="/xl/worksheets/sheet1.xml"` is legal OPC and bypasses
    // the base-directory join entirely.
    const wb = await read({
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "/xl/worksheets/sheet1.xml" },
      ]),
    })
    expect(wb.sheets[0].rows[0][0]).toBe("hi")
  })

  it("reads an external link that has no sibling rels file", async () => {
    // Without the rels there is no path to the linked workbook, but the
    // cached sheet names and values are still usable.
    const wb = await read({
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
        { id: "rIdNoDigits", type: "externalLink", target: "externalLinks/externalLink1.xml" },
      ]),
      "xl/externalLinks/externalLink1.xml":
        `<?xml version="1.0"?><externalLink ${NS} ${R}><externalBook r:id="rId1">` +
        `<sheetNames><sheetName val="Prices"/></sheetNames></externalBook></externalLink>`,
    })
    expect(wb.externalLinks![0]).toMatchObject({ target: "", sheetNames: ["Prices"] })
  })

  // BUG (reported, not worked around): a part stored at the package root
  // never finds its own `_rels` file. `reader.ts:369` computes
  //     wsFileName = wsPath.slice(dirname(wsPath).length + 1)
  // and `reader.ts:747` does the same for drawings. With no directory
  // prefix `dirname()` returns "", so the `+ 1` eats the first character
  // of the file name: the reader looks for `_rels/heet1.xml.rels`
  // instead of `_rels/sheet1.xml.rels`. The sibling helper `relsPathFor()`
  // (reader.ts:686) special-cases the no-slash form correctly, so the two
  // code paths disagree. Consequence: for a root-level worksheet, all of
  // hyperlinks, comments, tables, drawings, pivot tables and the sheet
  // background silently vanish. Verified directly: the same package with
  // the rels part deliberately misnamed `_rels/heet1.xml.rels` DOES
  // resolve the hyperlink, while the correctly named one does not.
  it("reads a drawing stored at the package root", async () => {
    const wb = await readExact({
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": relsXml([{ id: "rId1", type: "officeDocument", target: "workbook.xml" }]),
      "workbook.xml": workbookXml(`<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`),
      "_rels/workbook.xml.rels": relsXml([{ id: "rId1", type: "worksheet", target: "sheet1.xml" }]),
      "sheet1.xml": worksheetXml(ONE_CELL),
      "_rels/sheet1.xml.rels": relsXml([{ id: "rId1", type: "drawing", target: "drawing1.xml" }]),
      "drawing1.xml":
        `<?xml version="1.0"?><xdr:wsDr ${XDR} ${A} ${R}>` +
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}${pic(`r:embed="rId1"`)}</xdr:twoCellAnchor>` +
        `</xdr:wsDr>`,
      "_rels/drawing1.xml.rels": relsXml([{ id: "rId1", type: "image", target: "media/i.gif" }]),
      "media/i.gif": PNG,
    })
    expect(wb.sheets[0].images![0].type).toBe("gif")
  })

  it("falls back to png for a oneCellAnchor with an unknown media extension", async () => {
    const wb = await read(
      withDrawing(
        `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from>` +
          `<xdr:ext cx="9525" cy="9525"/>${pic(`r:embed="rId1"`)}</xdr:oneCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/image1.tiff" }],
        { "xl/media/image1.tiff": PNG },
      ),
    )
    expect(wb.sheets[0].images![0]).toMatchObject({ type: "png", width: 1, height: 1 })
  })

  it("reads a picture whose shape metadata is absent", async () => {
    // `nvPicPr` / `cNvPr` are required by the schema but turn up missing
    // in files assembled by templating engines; alt text is simply
    // unavailable then.
    const wb = await read(
      withDrawing(
        `<xdr:twoCellAnchor>${fromTo(0, 0, 1, 1)}<xdr:pic>` +
          `<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>` +
          `<xdr:twoCellAnchor>${fromTo(2, 2, 3, 3)}<xdr:pic>` +
          `<xdr:nvPicPr><xdr:cNvPicPr/></xdr:nvPicPr>` +
          `<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>`,
        [{ id: "rId1", type: "image", target: "../media/image1.svg" }],
        { "xl/media/image1.svg": PNG },
      ),
    )
    const images = wb.sheets[0].images!
    expect(images).toHaveLength(2)
    expect(images[0].altText).toBeUndefined()
    expect(images[1].title).toBeUndefined()
    expect(images[0].type).toBe("svg")
  })
})
