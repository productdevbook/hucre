import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { resolvePivotSource, writePivotTable } from "../src/xlsx/pivot-writer"
import { cloneChart } from "../src/xlsx/chart-clone"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { EncryptedFileError } from "../src/errors"
import type { CellValue, Chart, SheetChart, WritePivotTable, WriteSheet } from "../src/_types"

const encoder = new TextEncoder()
const decoder = new TextDecoder("utf-8")

// ── Helpers ──────────────────────────────────────────────────────────

async function part(buf: Uint8Array, path: string): Promise<string> {
  return decoder.decode(await new ZipReader(buf).extract(path))
}

function entries(buf: Uint8Array): string[] {
  return new ZipReader(buf).entries()
}

/**
 * Re-package a workbook with extra ZIP entries, the way a real producer
 * (Excel, WPS) ships parts hucre never authors itself — chart style
 * sidecars, slicers, external links. Nothing else about the file changes.
 */
async function withParts(
  buf: Uint8Array,
  extra: Record<string, string | Uint8Array>,
): Promise<Uint8Array> {
  const all = await new ZipReader(buf).extractAll()
  for (const [path, body] of Object.entries(extra)) {
    all.set(path, typeof body === "string" ? encoder.encode(body) : body)
  }
  const out = new ZipWriter()
  for (const [path, data] of all) out.add(path, data, { compress: false })
  return out.build()
}

// Declared as its own grid rather than read back off `SALES` — a
// `WriteSheet` row may now hold a cell object as well as a value (#433),
// and `resolvePivotSource` takes values.
const SALES_ROWS: CellValue[][] = [
  ["Region", "Product", "Quarter", "Revenue"],
  ["EU", "Widget", "Q1", 100],
  ["US", "Widget", "Q1", 200],
  ["EU", "Gadget", "Q2", 50],
  ["US", "Gadget", "Q2", 75],
]

const SALES: WriteSheet = { name: "Data", rows: SALES_ROWS }

// ═══════════════════════════════════════════════════════════════════════
// pivot-writer — axis placement
//
// Phase 1 of #159 emits a structurally valid pivot Excel finishes on
// refresh. The page (filter) axis and multi-field axes were the parts
// nothing exercised, so a broken `<pageFields>` block would have shipped.
// ═══════════════════════════════════════════════════════════════════════

describe("pivot page fields", () => {
  it("emits <pageFields> and an axisPage pivotField for a filter field", async () => {
    const buf = await writeXlsx({
      sheets: [
        SALES,
        {
          name: "Pivot",
          pivotTables: [
            {
              name: "ByRegion",
              sourceSheet: "Data",
              pages: ["Quarter"],
              rows: ["Region"],
              values: [{ field: "Revenue" }],
            },
          ],
        },
      ],
    })
    const xml = await part(buf, "xl/pivotTables/pivotTable1.xml")

    expect(xml).toContain('<pageFields count="1">')
    expect(xml).toContain('<pageField fld="2" hier="-1"/>')
    expect(xml).toContain('axis="axisPage"')
  })

  it("keeps each axis in declaration order when several fields share it", () => {
    // The axis lists are sorted by `axisOrder`; with one field per axis
    // the comparator never runs, so the ordering contract was untested.
    const pivot: WritePivotTable = {
      name: "Wide",
      rows: ["Product", "Region"],
      columns: ["Year", "Quarter"],
      pages: ["Channel", "Currency"],
      values: [{ field: "Revenue" }],
    }
    const source = resolvePivotSource(pivot, "Data", [
      ["Region", "Product", "Quarter", "Year", "Channel", "Currency", "Revenue"],
      ["EU", "Widget", "Q1", "2024", "Web", "EUR", 100],
      ["US", "Gadget", "Q2", "2025", "Retail", "USD", 200],
    ])
    const { pivotTableXml } = writePivotTable(pivot, source, 0)

    // Product (index 1) is declared before Region (index 0) on the row axis.
    expect(pivotTableXml).toContain('<rowFields count="2"><field x="1"/><field x="0"/></rowFields>')
    expect(pivotTableXml).toContain('<colFields count="2"><field x="3"/><field x="2"/></colFields>')
    expect(pivotTableXml).toContain(
      '<pageFields count="2"><pageField fld="4" hier="-1"/><pageField fld="5" hier="-1"/></pageFields>',
    )
  })

  it("marks a source column that is on no axis as hidden", async () => {
    const buf = await writeXlsx({
      sheets: [
        SALES,
        {
          name: "Pivot",
          pivotTables: [
            {
              name: "Sparse",
              sourceSheet: "Data",
              rows: ["Region"],
              values: [{ field: "Revenue" }],
            },
          ],
        },
      ],
    })
    const xml = await part(buf, "xl/pivotTables/pivotTable1.xml")

    // Product and Quarter are cached but unplaced.
    expect(xml).toContain('<pivotField showAll="0"/>')
  })
})

describe("pivot data fields", () => {
  it("declares a numFmtId for a value field that asked for a number format", async () => {
    const buf = await writeXlsx({
      sheets: [
        SALES,
        {
          name: "Pivot",
          pivotTables: [
            {
              name: "Money",
              sourceSheet: "Data",
              rows: ["Region"],
              values: [{ field: "Revenue", numberFormat: "#,##0.00" }],
            },
          ],
        },
      ],
    })

    expect(await part(buf, "xl/pivotTables/pivotTable1.xml")).toContain('numFmtId="0"')
  })

  it("labels every aggregation Excel supports", () => {
    const fns = [
      ["product", "Product of Revenue"],
      ["countNums", "Count Nums of Revenue"],
      ["stdDev", "StdDev of Revenue"],
      ["stdDevp", "StdDevp of Revenue"],
      ["var", "Var of Revenue"],
      ["varp", "Varp of Revenue"],
    ] as const

    for (const [fn, label] of fns) {
      const pivot: WritePivotTable = {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue", function: fn }],
      }
      const source = resolvePivotSource(pivot, "Data", SALES_ROWS)
      const { pivotTableXml } = writePivotTable(pivot, source, 0)
      expect(pivotTableXml).toContain(`name="${label}"`)
      expect(pivotTableXml).toContain(`subtotal="${fn}"`)
    }
  })

  it("streams a text field placed on the data axis as inline strings", async () => {
    // `Count of Region` puts a string field on the data axis, where no
    // shared-items table is built — the records fall back to `<s v="…"/>`.
    const buf = await writeXlsx({
      sheets: [
        SALES,
        {
          name: "Pivot",
          pivotTables: [
            {
              name: "Counts",
              sourceSheet: "Data",
              rows: ["Product"],
              values: [{ field: "Region", function: "count" }],
            },
          ],
        },
      ],
    })
    const records = await part(buf, "xl/pivotCache/pivotCacheRecords1.xml")

    expect(records).toContain('<s v="EU"/>')
    expect(records).toContain('<s v="US"/>')
  })
})

describe("pivot cache fields", () => {
  it("names an unlabelled source column ColumnN", () => {
    const pivot: WritePivotTable = {
      name: "P",
      rows: ["Column2"],
      values: [{ field: "Revenue" }],
    }
    const source = resolvePivotSource(pivot, "Data", [
      ["Region", null, "Revenue"],
      ["EU", "x", 100],
      ["US", "y", 200],
    ])

    expect(source.fieldNames).toEqual(["Region", "Column2", "Revenue"])
  })

  it("accepts a pivot with no row axis at all", () => {
    // Columns-only pivots are legal; `rows` being absent must not throw
    // while collecting the field names to validate.
    const pivot: WritePivotTable = {
      name: "P",
      columns: ["Region"],
      values: [{ field: "Revenue" }],
    }
    const source = resolvePivotSource(pivot, "Data", SALES_ROWS)
    const { pivotTableXml } = writePivotTable(pivot, source, 0)

    expect(pivotTableXml).toContain("<colFields")
    expect(pivotTableXml).not.toContain("<rowFields")
  })

  it("registers the blank of a string axis field as an empty shared item", () => {
    const pivot: WritePivotTable = {
      name: "P",
      rows: ["Region"],
      values: [{ field: "Revenue" }],
    }
    const source = resolvePivotSource(pivot, "Data", [
      ["Region", "Revenue"],
      ["EU", 100],
      [null, 200],
    ])
    const { cacheDefinitionXml, cacheRecordsXml } = writePivotTable(pivot, source, 0)

    expect(cacheDefinitionXml).toContain('<s v=""/>')
    // The record side still writes `<m/>` for the blank — the empty
    // shared item exists so Excel can show a "(blank)" row label.
    expect(cacheRecordsXml).toContain("<m/>")
  })

  it("reports containsInteger=0 for a column with fractional values", () => {
    const pivot: WritePivotTable = {
      name: "P",
      rows: ["Region"],
      values: [{ field: "Revenue" }],
    }
    const source = resolvePivotSource(pivot, "Data", [
      ["Region", "Revenue"],
      ["EU", 10.5],
      ["US", 20.25],
    ])
    const { cacheDefinitionXml } = writePivotTable(pivot, source, 0)

    expect(cacheDefinitionXml).toContain('containsInteger="0"')
    expect(cacheDefinitionXml).toContain('minValue="10.5"')
  })
})

describe("pivot input validation", () => {
  it("rejects a targetCell that is not an A1 reference", () => {
    const pivot: WritePivotTable = {
      name: "P",
      targetCell: "top-left",
      rows: ["Region"],
      values: [{ field: "Revenue" }],
    }
    const source = resolvePivotSource(pivot, "Data", SALES_ROWS)

    expect(() => writePivotTable(pivot, source, 0)).toThrow(/A1-style reference/)
  })

  it("rejects a source sheet whose header row is empty", () => {
    // No header cells means no fields, so there is no range to describe.
    expect(() => resolvePivotSource({ name: "P", values: [] }, "Data", [[], [1]])).toThrow(
      /at least one column and row/,
    )
  })
})

// ═══════════════════════════════════════════════════════════════════════
// roundtrip — openXlsx
// ═══════════════════════════════════════════════════════════════════════

describe("openXlsx on an encrypted workbook", () => {
  it("refuses to open a password-protected file without the password", async () => {
    const enc = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]] }],
      encryption: { password: "pw", spinCount: 64 },
    })

    await expect(openXlsx(enc)).rejects.toBeInstanceOf(EncryptedFileError)
  })

  it("decrypts up front so the preserved raw entries are plaintext parts", async () => {
    const enc = await writeXlsx({
      sheets: [{ name: "S", rows: [["kept"]] }],
      encryption: { password: "pw", spinCount: 64 },
    })
    const wb = await openXlsx(enc, { password: "pw" })
    const saved = await saveXlsx(wb)

    expect(entries(saved)).toContain("xl/worksheets/sheet1.xml")
    expect(wb.sheets[0].rows[0]).toEqual(["kept"])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// roundtrip — tables
//
// Tables are regenerated (not preserved), so open → save has to rebuild
// the table part, its relationship, and its range.
// ═══════════════════════════════════════════════════════════════════════

describe("openXlsx → saveXlsx keeps Excel tables", () => {
  const withTable: WriteSheet = {
    name: "S",
    rows: [
      ["Name", "Price"],
      ["Bolt", 2],
      ["Nut", 1],
    ],
    tables: [
      {
        name: "Inventory",
        range: "A1:B3",
        columns: [{ name: "Name" }, { name: "Price" }],
      },
    ],
  }

  it("re-emits the table part and the sheet relationship that points at it", async () => {
    const saved = await saveXlsx(await openXlsx(await writeXlsx({ sheets: [withTable] })))

    expect(await part(saved, "xl/tables/table1.xml")).toContain('ref="A1:B3"')
    expect(await part(saved, "xl/worksheets/_rels/sheet1.xml.rels")).toContain(
      "../tables/table1.xml",
    )
    expect(await part(saved, "[Content_Types].xml")).toContain("/xl/tables/table1.xml")
  })

  it("numbers tables across sheets from a single global counter", async () => {
    const second: WriteSheet = { ...withTable, name: "S2" }
    second.tables = [{ name: "Second", range: "A1:B3", columns: [{ name: "Name" }] }]
    const saved = await saveXlsx(await openXlsx(await writeXlsx({ sheets: [withTable, second] })))

    expect(entries(saved)).toContain("xl/tables/table2.xml")
  })

  it("recomputes a range for a table added to an opened workbook", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["A"], [1], [2]] }] }))
    wb.sheets[0].tables = [
      { name: "Added", showTotalRow: true, columns: [{ name: "A" }, { name: "B" }] },
    ]
    const saved = await saveXlsx(wb)

    // 3 sheet rows + 1 totals row.
    expect(await part(saved, "xl/tables/table1.xml")).toContain('ref="A1:B4"')
  })

  it("never computes an empty range for a table on an empty sheet", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [] }] }))
    wb.sheets[0].tables = [{ name: "Added", columns: [{ name: "A" }] }]
    const saved = await saveXlsx(wb)

    expect(await part(saved, "xl/tables/table1.xml")).toContain('ref="A1:A1"')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// roundtrip — print settings
// ═══════════════════════════════════════════════════════════════════════

describe("openXlsx → saveXlsx rebuilds print defined names", () => {
  it("re-declares Print_Area and Print_Titles from the sheet page setup", async () => {
    const wb = await openXlsx(
      await writeXlsx({
        sheets: [
          {
            name: "Report",
            rows: [
              ["a", "b"],
              [1, 2],
            ],
          },
        ],
      }),
    )
    wb.sheets[0].pageSetup = {
      printArea: "$A$1:$B$2",
      printTitlesRow: "$1:$1",
      printTitlesColumn: "$A:$A",
    }
    const xml = await part(await saveXlsx(wb), "xl/workbook.xml")

    expect(xml).toContain("_xlnm.Print_Area")
    expect(xml).toContain("Report!$A$1:$B$2")
    expect(xml).toContain("_xlnm.Print_Titles")
    expect(xml).toContain("Report!$1:$1,Report!$A:$A")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// roundtrip — parts hucre does not model
// ═══════════════════════════════════════════════════════════════════════

describe("openXlsx → saveXlsx re-declares foreign workbook parts", () => {
  it("keeps two external links and gives each its own workbook relationship", async () => {
    const link = (n: number) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">
    <sheetNames><sheetName val="Sheet${n}"/></sheetNames>
  </externalBook>
</externalLink>`
    const base = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    const buf = await withParts(base, {
      "xl/externalLinks/externalLink1.xml": link(1),
      "xl/externalLinks/externalLink2.xml": link(2),
    })
    const saved = await saveXlsx(await openXlsx(buf))
    const rels = await part(saved, "xl/_rels/workbook.xml.rels")

    expect(rels).toContain("externalLinks/externalLink1.xml")
    expect(rels).toContain("externalLinks/externalLink2.xml")
    expect(await part(saved, "[Content_Types].xml")).toContain(
      "/xl/externalLinks/externalLink2.xml",
    )
  })

  it("keeps two slicers and two timelines wired to their caches", async () => {
    const slicers = (n: number) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="S${n}" cache="Slicer_S${n}" caption="S${n}"/>
</slicers>`
    const slicerCache = (n: number) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Slicer_S${n}" sourceName="S${n}"/>`
    const timelines = (n: number) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="T${n}" cache="NativeTimeline_T${n}" caption="T${n}" level="months"/>
</timelines>`
    const timelineCache = (n: number) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="NativeTimeline_T${n}" sourceName="T${n}"/>`

    const base = await writeXlsx({
      sheets: [
        { name: "Data", rows: [["a"]] },
        { name: "Filters", rows: [["b"]] },
      ],
    })
    const buf = await withParts(base, {
      "xl/slicers/slicer1.xml": slicers(1),
      "xl/slicers/slicer2.xml": slicers(2),
      "xl/slicerCaches/slicerCache1.xml": slicerCache(1),
      "xl/slicerCaches/slicerCache2.xml": slicerCache(2),
      "xl/timelines/timeline1.xml": timelines(1),
      "xl/timelines/timeline2.xml": timelines(2),
      "xl/timelineCaches/timelineCache1.xml": timelineCache(1),
      "xl/timelineCaches/timelineCache2.xml": timelineCache(2),
      "xl/worksheets/_rels/sheet2.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline2.xml"/>
</Relationships>`,
    })
    const saved = await saveXlsx(await openXlsx(buf))

    const wbRels = await part(saved, "xl/_rels/workbook.xml.rels")
    expect(wbRels).toContain("slicerCaches/slicerCache2.xml")
    expect(wbRels).toContain("timelineCaches/timelineCache2.xml")

    const sheetRels = await part(saved, "xl/worksheets/_rels/sheet2.xml.rels")
    expect(sheetRels).toContain("../slicers/slicer2.xml")
    expect(sheetRels).toContain("../timelines/timeline2.xml")
  })

  it("keeps two pivot caches and their pivot tables", async () => {
    const pivotSheet: WriteSheet = {
      name: "Pivots",
      pivotTables: [
        {
          name: "First",
          sourceSheet: "Data",
          rows: ["Region"],
          values: [{ field: "Revenue" }],
        },
        {
          name: "Second",
          sourceSheet: "Data",
          targetCell: "A20",
          rows: ["Product"],
          values: [{ field: "Revenue", function: "average" }],
        },
      ],
    }
    const saved = await saveXlsx(await openXlsx(await writeXlsx({ sheets: [SALES, pivotSheet] })))
    const names = entries(saved)

    expect(names).toContain("xl/pivotCache/pivotCacheDefinition2.xml")
    expect(names).toContain("xl/pivotCache/pivotCacheRecords2.xml")
    expect(names).toContain("xl/pivotTables/pivotTable2.xml")

    const sheetRels = await part(saved, "xl/worksheets/_rels/sheet2.xml.rels")
    expect(sheetRels).toContain("../pivotTables/pivotTable1.xml")
    expect(sheetRels).toContain("../pivotTables/pivotTable2.xml")
  })

  it("resolves cell-image media paths written as package-absolute or ../ targets", async () => {
    // WPS writes `media/imageN.png` relative to `xl/cellimages.xml`, but
    // the same registry is also seen with package-absolute targets
    // (`/xl/media/...`) and with `../` segments, which land at the
    // package root. Both forms have to be resolved, because a path that
    // resolves under `xl/media/` would otherwise be swept away with the
    // drawing images the writer regenerates.
    const base = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const buf = await withParts(base, {
      "xl/cellimages.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData"/>`,
      "xl/_rels/cellimages.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/xl/media/image90.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image91.png"/>
</Relationships>`,
      "xl/media/image90.png": png,
      "media/image91.png": png,
    })
    const saved = await saveXlsx(await openXlsx(buf))
    const names = entries(saved)

    // Without the absolute-target branch this one is filtered out as a
    // regenerated drawing image.
    expect(names).toContain("xl/media/image90.png")
    expect(names).toContain("media/image91.png")
    expect(names).toContain("xl/_rels/cellimages.xml.rels")
  })
})

describe("openXlsx → saveXlsx drops parts the workbook no longer needs", () => {
  it("leaves behind the worksheet of a sheet that was removed", async () => {
    const wb = await openXlsx(
      await writeXlsx({
        sheets: [
          { name: "Keep", rows: [["a"]] },
          { name: "Drop", rows: [["b"]] },
        ],
      }),
    )
    wb.sheets.splice(1, 1)
    const saved = await saveXlsx(wb)

    // sheet2.xml survives in the raw entries but is filtered by the
    // "xl/worksheets/" regenerated prefix, so it must not be re-added.
    expect(entries(saved)).not.toContain("xl/worksheets/sheet2.xml")
    expect(await part(saved, "xl/workbook.xml")).not.toContain('name="Drop"')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// roundtrip — chart preservation
//
// hucre's drawing writer cannot re-emit a chart graphicFrame, so a
// chart-bearing drawing is preserved byte-for-byte and re-anchored into
// the regenerated worksheet body.
// ═══════════════════════════════════════════════════════════════════════

function chartOn(row: number, title: string): SheetChart {
  return {
    type: "column",
    title,
    series: [{ name: "Revenue", values: "D2:D5", categories: "A2:A5" }],
    anchor: { from: { row, col: 0 }, to: { row: row + 12, col: 6 } },
  }
}

describe("chart parts survive the roundtrip", () => {
  it("declares preserved chart style and colour sidecars in [Content_Types].xml", async () => {
    // `style1.xml` / `colors1.xml` only come from Excel — hucre never
    // writes them — so they exercise the preserved-sidecar declaration.
    const base = await writeXlsx({
      sheets: [{ ...SALES, charts: [chartOn(7, "First"), chartOn(24, "Second")] }],
    })
    const sidecar = (tag: string) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${tag} xmlns="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>`
    const buf = await withParts(base, {
      "xl/charts/style1.xml": sidecar("cs:chartStyle"),
      "xl/charts/style2.xml": sidecar("cs:chartStyle"),
      "xl/charts/colors1.xml": sidecar("cs:colorStyle"),
      "xl/charts/colors2.xml": sidecar("cs:colorStyle"),
    })
    const saved = await saveXlsx(await openXlsx(buf))
    const ct = await part(saved, "[Content_Types].xml")

    expect(ct).toContain("/xl/charts/style2.xml")
    expect(ct).toContain("/xl/charts/colors2.xml")
    expect(entries(saved)).toContain("xl/charts/chart2.xml")
  })

  it("re-anchors the preserved drawing into the regenerated worksheet body", async () => {
    const base = await writeXlsx({ sheets: [{ ...SALES, charts: [chartOn(7, "Sales")] }] })
    const saved = await saveXlsx(await openXlsx(base))
    const ws = await part(saved, "xl/worksheets/sheet1.xml")

    expect(ws).toContain("<drawing r:id=")
    expect(await part(saved, "xl/worksheets/_rels/sheet1.xml.rels")).toContain(
      "../drawings/drawing1.xml",
    )
  })

  it("inserts the re-anchored <drawing> before <tableParts>", async () => {
    // CT_Worksheet fixes the child order: `drawing` precedes `tableParts`.
    // Appending at the end of the body would make Excel reject the sheet.
    const sheet: WriteSheet = {
      ...SALES,
      charts: [chartOn(7, "Sales")],
      tables: [
        {
          name: "T",
          range: "A1:D5",
          columns: [
            { name: "Region" },
            { name: "Product" },
            { name: "Quarter" },
            { name: "Revenue" },
          ],
        },
      ],
    }
    const saved = await saveXlsx(await openXlsx(await writeXlsx({ sheets: [sheet] })))
    const ws = await part(saved, "xl/worksheets/sheet1.xml")

    expect(ws.indexOf("<drawing ")).toBeGreaterThan(-1)
    expect(ws.indexOf("<drawing ")).toBeLessThan(ws.indexOf("<tableParts"))
  })

  it("does not preserve a drawing that holds no chart", async () => {
    // Sheet 2 keeps a hucre-managed image (its drawing is regenerated);
    // sheet 3's image is removed before saving, so its original drawing
    // is neither regenerated nor chart-bearing and is simply dropped.
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const base = await writeXlsx({
      sheets: [
        { ...SALES, charts: [chartOn(7, "Sales")] },
        {
          name: "Pics",
          rows: [["b"]],
          images: [{ data: png, type: "png", anchor: { from: { row: 0, col: 2 } } }],
        },
        {
          name: "WasPics",
          rows: [["c"]],
          images: [{ data: png, type: "png", anchor: { from: { row: 0, col: 2 } } }],
        },
      ],
    })
    const wb = await openXlsx(base)
    wb.sheets[2].images = []
    const saved = await saveXlsx(wb)
    const names = entries(saved)

    expect(names).toContain("xl/drawings/drawing1.xml") // chart drawing, preserved
    expect(names).toContain("xl/drawings/drawing2.xml") // image drawing, regenerated
    expect(names).not.toContain("xl/drawings/drawing3.xml")
  })

  it("finds the chart drawing past unrelated and unnumbered drawing relationships", async () => {
    // Sheet rels list hyperlinks first, and some producers ship an
    // unnumbered `drawing.xml`. Neither may stop the scan that maps the
    // sheet onto its chart-bearing drawing.
    const sheet: WriteSheet = {
      ...SALES,
      charts: [chartOn(7, "Sales")],
      cells: new Map([["0,0", { value: "Region", hyperlink: { target: "https://example.com" } }]]),
    }
    const base = await writeXlsx({ sheets: [sheet] })
    const relsPath = "xl/worksheets/_rels/sheet1.xml.rels"
    const original = await part(base, relsPath)
    const buf = await withParts(base, {
      "xl/drawings/drawing.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`,
      // The unnumbered drawing is listed first, so the scan has to skip
      // past it rather than stop at the first drawing relationship.
      [relsPath]: original.replace(
        "<Relationship ",
        `<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing.xml"/><Relationship `,
      ),
    })
    const saved = await saveXlsx(await openXlsx(buf))

    expect(await part(saved, "xl/worksheets/sheet1.xml")).toContain("<drawing r:id=")
    expect(entries(saved)).toContain("xl/drawings/drawing1.xml")
  })

  it("ignores a drawing that merely mentions :chartSpace in an attribute", async () => {
    // The detector matches the `chart` local name only — `:chartSpace`
    // and `:chartstyle` must not make an ordinary drawing look
    // chart-bearing, or the roundtrip preserves an orphan part.
    const base = await writeXlsx({ sheets: [{ ...SALES, charts: [chartOn(7, "Sales")] }] })
    const buf = await withParts(base, {
      "xl/drawings/drawing7.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:oneCellAnchor>
    <xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Legacy" descr="exported from c:chartSpace, see c:chartstyle"/></xdr:nvPicPr></xdr:pic>
  </xdr:oneCellAnchor>
</xdr:wsDr>`,
    })
    const saved = await saveXlsx(await openXlsx(buf))

    expect(entries(saved)).not.toContain("xl/drawings/drawing7.xml")
  })
})

describe("model charts added to an opened workbook", () => {
  it("folds into the drawing a sheet already owns for its images", async () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const wb = await openXlsx(
      await writeXlsx({
        sheets: [
          {
            ...SALES,
            images: [{ data: png, type: "png", anchor: { from: { row: 0, col: 6 } } }],
          },
        ],
      }),
    )
    // @ts-expect-error Sheet.charts is the read model; the roundtrip
    // bridge accepts write-model entries here (issue #136).
    wb.sheets[0].charts = [chartOn(7, "In time")]
    const saved = await saveXlsx(wb)

    // hucre authored that image drawing this run, so it is hucre's to
    // extend — the chart goes into it rather than being dropped. Skipping
    // was the old behaviour and it lost the chart outright. See #465.
    expect(entries(saved).some((n) => n.startsWith("xl/charts/"))).toBe(true)
    expect(entries(saved).filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))).toEqual([
      "xl/drawings/drawing1.xml",
    ])
  })

  it("skips a sheet whose original drawing hucre is no longer rebuilding", async () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    const wb = await openXlsx(
      await writeXlsx({
        sheets: [
          {
            ...SALES,
            images: [{ data: png, type: "png", anchor: { from: { row: 0, col: 6 } } }],
          },
        ],
      }),
    )
    wb.sheets[0].images = []
    // @ts-expect-error see above — write-model chart on a read-model sheet.
    wb.sheets[0].charts = [chartOn(7, "Still too late")]
    const saved = await saveXlsx(wb)

    expect(entries(saved).some((n) => n.startsWith("xl/charts/"))).toBe(false)
  })

  it("drops charts the write model cannot express instead of failing the save", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] }))
    const radar: Chart = {
      kinds: ["radar"],
      seriesCount: 1,
      series: [{ kind: "radar", index: 0, name: "R", valuesRef: "S!$A$1:$A$2" }],
    }
    wb.sheets[0].charts = [radar]
    const saved = await saveXlsx(wb)

    expect(entries(saved).some((n) => n.startsWith("xl/charts/"))).toBe(false)
    expect(await part(saved, "xl/worksheets/sheet1.xml")).not.toContain("<drawing ")
  })

  it("anchors a chart that arrives without one at the top-left cell", async () => {
    const wb = await openXlsx(await writeXlsx({ sheets: [{ ...SALES }] }))
    // A write-model chart with no anchor — the shape JSON or a template
    // helper hands over. Sheet.charts is the read model; the roundtrip
    // bridge accepts write-model entries here (issue #136).
    const unanchored = {
      type: "column",
      title: "Unanchored",
      series: [{ name: "Revenue", values: "D2:D5", categories: "A2:A5" }],
    }
    wb.sheets[0].charts = [unanchored as unknown as Chart]
    const saved = await saveXlsx(wb)
    const drawing = await part(saved, "xl/drawings/drawing1.xml")

    expect(drawing).toContain("<xdr:col>0</xdr:col>")
    expect(drawing).toContain("<xdr:row>0</xdr:row>")
  })

  it("ignores a non-object entry in a sheet's charts array", async () => {
    // `charts` can arrive from JSON with holes in it.
    const wb = await openXlsx(await writeXlsx({ sheets: [{ ...SALES }] }))
    // @ts-expect-error deliberately malformed input, as JSON would deliver it.
    wb.sheets[0].charts = [null, chartOn(7, "Real")]
    const saved = await saveXlsx(wb)

    expect(entries(saved)).toContain("xl/charts/chart1.xml")
    expect(await part(saved, "xl/charts/chart1.xml")).toContain("Real")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// chart-clone — line cap / compound overrides
//
// `<a:ln cap= cmpd=>` is the last pair of stroke knobs; the clone
// resolvers are shared with the colour / width / dash ones but each
// target lives on a different OOXML host.
// ═══════════════════════════════════════════════════════════════════════

describe("cloneChart border cap and compound", () => {
  const source: Chart = {
    kinds: ["bar"],
    seriesCount: 1,
    title: "Template",
    series: [
      {
        kind: "bar",
        index: 0,
        name: "Revenue",
        valuesRef: "Sheet1!$B$2:$B$5",
        categoriesRef: "Sheet1!$A$2:$A$5",
      },
    ],
  }

  it("applies cap and compound overrides to the legend, plot area and title borders", () => {
    const clone = cloneChart(source, {
      anchor: { from: { row: 0, col: 0 } },
      title: "Clone",
      legendBorderCap: "rnd",
      legendBorderCompound: "dbl",
      plotAreaBorderCap: "sq",
      plotAreaBorderCompound: "thickThin",
      titleBorderCap: "rnd",
      titleBorderCompound: "tri",
      // "flat" is the OOXML default cap and collapses to absence.
      plotAreaBorderDash: "dash",
    })

    expect(clone.legendBorderCap).toBe("rnd")
    expect(clone.legendBorderCompound).toBe("dbl")
    expect(clone.plotAreaBorderCap).toBe("sq")
    expect(clone.plotAreaBorderCompound).toBe("thickThin")
    expect(clone.titleBorderCap).toBe("rnd")
    expect(clone.titleBorderCompound).toBe("tri")
  })

  it("drops a series 3-D shape when the clone is coerced off the bar family", () => {
    // `<c:shape>` only exists on bar3D series; carrying it onto a line
    // clone would leak metadata the writer must ignore anyway.
    const bar3d: Chart = {
      ...source,
      kinds: ["bar3D"],
      series: [{ ...source.series![0], kind: "bar3D", shape3D: "cylinder" }],
    }

    const anchor = { from: { row: 0, col: 0 } }
    expect(cloneChart(bar3d, { anchor }).series?.[0].shape3D).toBe("cylinder")
    expect(cloneChart(bar3d, { anchor, type: "line" }).series?.[0].shape3D).toBeUndefined()
  })
})
