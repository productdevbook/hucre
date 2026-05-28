import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { parseXml } from "../src/xml/parser"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeTable } from "../src/xlsx/table-writer"
import { writeContentTypes } from "../src/xlsx/content-types-writer"
import type { WriteSheet, TableDefinition } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8")

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data)
  const raw = await zip.extract(path)
  return decoder.decode(raw)
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function zipHas(data: Uint8Array, path: string): boolean {
  const zip = new ZipReader(data)
  return zip.has(path)
}

// ── writeTable unit tests ────────────────────────────────────────────

describe("writeTable", () => {
  it("generates table XML with correct structure", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:D5",
      columns: [{ name: "Name" }, { name: "Price" }, { name: "Stock" }, { name: "Active" }],
    }

    const result = writeTable(table, 1, 1)

    expect(result.tableXml).toContain('<?xml version="1.0"')
    expect(result.tableXml).toContain("<table")
    expect(result.tableXml).toContain('id="1"')
    expect(result.tableXml).toContain('name="Table1"')
    expect(result.tableXml).toContain('displayName="Table1"')
    expect(result.tableXml).toContain('ref="A1:D5"')
    expect(result.tableXml).toContain('totalsRowShown="0"')
    expect(result.tableXml).toContain("<autoFilter")
    expect(result.tableXml).toContain('ref="A1:D5"')
    expect(result.tableXml).toContain('<tableColumns count="4"')
    expect(result.tableXml).toContain('name="Name"')
    expect(result.tableXml).toContain('name="Price"')
    expect(result.tableXml).toContain('name="Stock"')
    expect(result.tableXml).toContain('name="Active"')
    expect(result.tableXml).toContain("<tableStyleInfo")
    expect(result.tableId).toBe(1)
  })

  it("uses custom display name", () => {
    const table: TableDefinition = {
      name: "SalesTable",
      displayName: "Sales Data",
      range: "A1:B3",
      columns: [{ name: "Item" }, { name: "Amount" }],
    }

    const result = writeTable(table, 2, 2)
    expect(result.tableXml).toContain('name="SalesTable"')
    expect(result.tableXml).toContain('displayName="Sales Data"')
  })

  it("applies custom style name", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:B3",
      columns: [{ name: "A" }, { name: "B" }],
      style: "TableStyleLight1",
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).toContain('name="TableStyleLight1"')
    expect(result.tableXml).not.toContain("TableStyleMedium2")
  })

  it("defaults to TableStyleMedium2 when no style specified", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:B3",
      columns: [{ name: "A" }, { name: "B" }],
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).toContain('name="TableStyleMedium2"')
  })

  it("generates auto-filter by default", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:C5",
      columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).toContain("<autoFilter")
    expect(result.tableXml).toContain('ref="A1:C5"')
  })

  it("suppresses auto-filter when showAutoFilter is false", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:C5",
      columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
      showAutoFilter: false,
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).not.toContain("<autoFilter")
  })

  it("generates total row with functions", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:C6",
      columns: [
        { name: "Item", totalLabel: "Total" },
        { name: "Price", totalFunction: "sum" },
        { name: "Quantity", totalFunction: "count" },
      ],
      showTotalRow: true,
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).toContain('totalsRowCount="1"')
    expect(result.tableXml).not.toContain('totalsRowShown="0"')
    expect(result.tableXml).toContain('totalsRowLabel="Total"')
    expect(result.tableXml).toContain('totalsRowFunction="sum"')
    expect(result.tableXml).toContain('totalsRowFunction="count"')
  })

  it("generates total row with custom formula", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:B5",
      columns: [
        { name: "Value" },
        { name: "Computed", totalFunction: "custom", totalFormula: "SUM([Value])*1.1" },
      ],
      showTotalRow: true,
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).toContain("<totalsRowFormula")
    expect(result.tableXml).toContain("SUM([Value])*1.1")
  })

  it("auto-filter ref excludes total row when showTotalRow is true", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:C6",
      columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
      showTotalRow: true,
    }

    const result = writeTable(table, 1, 1)
    // Table ref is A1:C6 (includes total row), autoFilter ref should be A1:C5
    const doc = parseXml(result.tableXml)
    const autoFilter = findChild(doc, "autoFilter")
    expect(autoFilter).toBeTruthy()
    expect(autoFilter.attrs["ref"]).toBe("A1:C5")
  })

  it("configures showRowStripes and showColumnStripes", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:B3",
      columns: [{ name: "A" }, { name: "B" }],
      showRowStripes: false,
      showColumnStripes: true,
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).toContain('showRowStripes="0"')
    expect(result.tableXml).toContain('showColumnStripes="1"')
  })

  it("defaults to showRowStripes=true and showColumnStripes=false", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:B3",
      columns: [{ name: "A" }, { name: "B" }],
    }

    const result = writeTable(table, 1, 1)
    expect(result.tableXml).toContain('showRowStripes="1"')
    expect(result.tableXml).toContain('showColumnStripes="0"')
  })

  it("generates correct column IDs (1-based, sequential)", () => {
    const table: TableDefinition = {
      name: "Table1",
      range: "A1:C5",
      columns: [{ name: "First" }, { name: "Second" }, { name: "Third" }],
    }

    const result = writeTable(table, 1, 1)
    const doc = parseXml(result.tableXml)
    const tableColumns = findChild(doc, "tableColumns")
    const cols = findChildren(tableColumns, "tableColumn")

    expect(cols).toHaveLength(3)
    expect(cols[0].attrs["id"]).toBe("1")
    expect(cols[0].attrs["name"]).toBe("First")
    expect(cols[1].attrs["id"]).toBe("2")
    expect(cols[1].attrs["name"]).toBe("Second")
    expect(cols[2].attrs["id"]).toBe("3")
    expect(cols[2].attrs["name"]).toBe("Third")
  })
})

// ── Full XLSX write integration ──────────────────────────────────────

describe("XLSX table writing", () => {
  it("writes a single table to the ZIP", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Name", "Price", "Stock"],
        ["Widget", 9.99, 100],
        ["Gadget", 19.99, 50],
      ],
      tables: [
        {
          name: "Table1",
          range: "A1:C3",
          columns: [{ name: "Name" }, { name: "Price" }, { name: "Stock" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    // Table XML file should exist
    expect(zipHas(xlsx, "xl/tables/table1.xml")).toBe(true)

    // Table XML should have correct content
    const tableXml = await extractXml(xlsx, "xl/tables/table1.xml")
    expect(tableXml).toContain('name="Table1"')
    expect(tableXml).toContain('ref="A1:C3"')
    expect(tableXml).toContain('name="Name"')
    expect(tableXml).toContain('name="Price"')
    expect(tableXml).toContain('name="Stock"')
  })

  it("includes tableParts element in worksheet XML", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Name", "Value"],
        ["A", 1],
      ],
      tables: [
        {
          name: "Table1",
          range: "A1:B2",
          columns: [{ name: "Name" }, { name: "Value" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml")

    expect(wsXml).toContain("<tableParts")
    expect(wsXml).toContain('count="1"')
    expect(wsXml).toContain("<tablePart")
    expect(wsXml).toContain("r:id=")
  })

  it("includes content type override for table", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A"], [1]],
      tables: [
        {
          name: "Table1",
          range: "A1:A2",
          columns: [{ name: "A" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const ctXml = await extractXml(xlsx, "[Content_Types].xml")

    expect(ctXml).toContain("/xl/tables/table1.xml")
    expect(ctXml).toContain("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
  })

  it("includes table relationship in sheet rels", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A"], [1]],
      tables: [
        {
          name: "Table1",
          range: "A1:A2",
          columns: [{ name: "A" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    expect(zipHas(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(true)
    const relsXml = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")

    expect(relsXml).toContain("../tables/table1.xml")
    expect(relsXml).toContain(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
    )
  })

  it("writes table with style name", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["X"], [1]],
      tables: [
        {
          name: "StyledTable",
          range: "A1:A2",
          columns: [{ name: "X" }],
          style: "TableStyleDark1",
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const tableXml = await extractXml(xlsx, "xl/tables/table1.xml")

    expect(tableXml).toContain('name="TableStyleDark1"')
  })

  it("writes table with total row", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Item", "Price"],
        ["A", 10],
        ["B", 20],
        ["Total", 30],
      ],
      tables: [
        {
          name: "Table1",
          range: "A1:B4",
          columns: [
            { name: "Item", totalLabel: "Total" },
            { name: "Price", totalFunction: "sum" },
          ],
          showTotalRow: true,
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const tableXml = await extractXml(xlsx, "xl/tables/table1.xml")

    expect(tableXml).toContain('totalsRowCount="1"')
    expect(tableXml).toContain('totalsRowLabel="Total"')
    expect(tableXml).toContain('totalsRowFunction="sum"')
  })

  it("auto-calculates range from row data", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Name", "Value"],
        ["A", 1],
        ["B", 2],
        ["C", 3],
      ],
      tables: [
        {
          name: "Table1",
          // No range — should be auto-calculated
          columns: [{ name: "Name" }, { name: "Value" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const tableXml = await extractXml(xlsx, "xl/tables/table1.xml")

    // 4 rows, 2 columns → A1:B4
    expect(tableXml).toContain('ref="A1:B4"')
  })

  it("writes multiple tables on same sheet", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["A", "B", null, "C", "D"],
        [1, 2, null, 3, 4],
      ],
      tables: [
        {
          name: "Table1",
          range: "A1:B2",
          columns: [{ name: "A" }, { name: "B" }],
        },
        {
          name: "Table2",
          range: "D1:E2",
          columns: [{ name: "C" }, { name: "D" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    // Both table files should exist
    expect(zipHas(xlsx, "xl/tables/table1.xml")).toBe(true)
    expect(zipHas(xlsx, "xl/tables/table2.xml")).toBe(true)

    // Table 1
    const table1Xml = await extractXml(xlsx, "xl/tables/table1.xml")
    expect(table1Xml).toContain('name="Table1"')
    expect(table1Xml).toContain('ref="A1:B2"')

    // Table 2
    const table2Xml = await extractXml(xlsx, "xl/tables/table2.xml")
    expect(table2Xml).toContain('name="Table2"')
    expect(table2Xml).toContain('ref="D1:E2"')

    // Worksheet should have tableParts count="2"
    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml")
    expect(wsXml).toContain('count="2"')

    // Content types should have both tables
    const ctXml = await extractXml(xlsx, "[Content_Types].xml")
    expect(ctXml).toContain("/xl/tables/table1.xml")
    expect(ctXml).toContain("/xl/tables/table2.xml")

    // Rels should have both table relationships
    const relsXml = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")
    expect(relsXml).toContain("../tables/table1.xml")
    expect(relsXml).toContain("../tables/table2.xml")
  })

  it("writes tables on different sheets with global indexing", async () => {
    const sheet1: WriteSheet = {
      name: "Sheet1",
      rows: [["A"], [1]],
      tables: [
        {
          name: "Table1",
          range: "A1:A2",
          columns: [{ name: "A" }],
        },
      ],
    }
    const sheet2: WriteSheet = {
      name: "Sheet2",
      rows: [["B"], [2]],
      tables: [
        {
          name: "Table2",
          range: "A1:A2",
          columns: [{ name: "B" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })

    // Each sheet has its own table with global indices
    expect(zipHas(xlsx, "xl/tables/table1.xml")).toBe(true)
    expect(zipHas(xlsx, "xl/tables/table2.xml")).toBe(true)

    const table1Xml = await extractXml(xlsx, "xl/tables/table1.xml")
    expect(table1Xml).toContain('name="Table1"')

    const table2Xml = await extractXml(xlsx, "xl/tables/table2.xml")
    expect(table2Xml).toContain('name="Table2"')

    // Each sheet rels should reference its own table
    const rels1 = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")
    expect(rels1).toContain("../tables/table1.xml")

    const rels2 = await extractXml(xlsx, "xl/worksheets/_rels/sheet2.xml.rels")
    expect(rels2).toContain("../tables/table2.xml")

    // Content types should have both
    const ctXml = await extractXml(xlsx, "[Content_Types].xml")
    expect(ctXml).toContain("/xl/tables/table1.xml")
    expect(ctXml).toContain("/xl/tables/table2.xml")
  })

  it("does not create table parts for sheets without tables", async () => {
    const sheet1: WriteSheet = {
      name: "WithTable",
      rows: [["A"], [1]],
      tables: [
        {
          name: "Table1",
          range: "A1:A2",
          columns: [{ name: "A" }],
        },
      ],
    }
    const sheet2: WriteSheet = {
      name: "NoTable",
      rows: [["B"], [2]],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })

    // Sheet1 has table
    expect(zipHas(xlsx, "xl/tables/table1.xml")).toBe(true)

    // Sheet2 worksheet XML should not have tableParts
    const ws2Xml = await extractXml(xlsx, "xl/worksheets/sheet2.xml")
    expect(ws2Xml).not.toContain("tableParts")
    expect(ws2Xml).not.toContain("tablePart")
  })

  it("empty tables array does not create table parts", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      tables: [],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    expect(zipHas(xlsx, "xl/tables/table1.xml")).toBe(false)

    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml")
    expect(wsXml).not.toContain("tableParts")
  })
})

// ── Content Types ────────────────────────────────────────────────────

describe("content types with tables", () => {
  it("writeContentTypes function with table indices", () => {
    const xml = writeContentTypes({
      sheetCount: 1,
      hasSharedStrings: true,
      tableIndices: [1, 2],
    })

    expect(xml).toContain("/xl/tables/table1.xml")
    expect(xml).toContain("/xl/tables/table2.xml")
    expect(xml).toContain("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
  })

  it("does not include table content types when no tables", () => {
    const xml = writeContentTypes({
      sheetCount: 1,
      hasSharedStrings: false,
    })

    expect(xml).not.toContain("table")
  })
})

// ── Coexistence with other features ─────────────────────────────────

describe("table coexistence", () => {
  it("coexists with hyperlinks in sheet rels", async () => {
    const cells = new Map<string, Partial<import("../src/_types").Cell>>()
    cells.set("0,0", {
      value: "Link",
      hyperlink: { target: "https://example.com" },
    })

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Link", "Value"],
        ["data", 42],
      ],
      cells,
      tables: [
        {
          name: "Table1",
          range: "A1:B2",
          columns: [{ name: "Link" }, { name: "Value" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const relsXml = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")

    // Should contain both hyperlink and table relationships
    expect(relsXml).toContain("hyperlink")
    expect(relsXml).toContain("table")
    expect(relsXml).toContain("https://example.com")
    expect(relsXml).toContain("../tables/table1.xml")

    // rIds should be unique
    const doc = parseXml(relsXml)
    const rels = findChildren(doc, "Relationship")
    const ids = rels.map((r: any) => r.attrs["Id"])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it("coexists with images and comments", async () => {
    // Fake PNG
    const imageData = new Uint8Array(64)
    imageData[0] = 0x89
    imageData[1] = 0x50
    imageData[2] = 0x4e
    imageData[3] = 0x47

    const cells = new Map<string, Partial<import("../src/_types").Cell>>()
    cells.set("0,0", {
      value: "Data",
      comment: { text: "A note" },
    })

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Data", "Value"],
        ["A", 1],
      ],
      cells,
      images: [
        {
          data: imageData,
          type: "png",
          anchor: { from: { row: 3, col: 0 }, to: { row: 8, col: 3 } },
        },
      ],
      tables: [
        {
          name: "Table1",
          range: "A1:B2",
          columns: [{ name: "Data" }, { name: "Value" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })

    // All features should be present
    expect(zipHas(xlsx, "xl/tables/table1.xml")).toBe(true)
    expect(zipHas(xlsx, "xl/drawings/drawing1.xml")).toBe(true)
    expect(zipHas(xlsx, "xl/comments1.xml")).toBe(true)

    // Sheet rels should have all relationship types
    const relsXml = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels")
    expect(relsXml).toContain("drawing")
    expect(relsXml).toContain("vmlDrawing")
    expect(relsXml).toContain("comments")
    expect(relsXml).toContain("table")

    // All rIds should be unique
    const doc = parseXml(relsXml)
    const rels = findChildren(doc, "Relationship")
    const ids = rels.map((r: any) => r.attrs["Id"])
    expect(new Set(ids).size).toBe(ids.length)
  })
})

// ── Round-trip (write + read) ────────────────────────────────────────

describe("table round-trip", () => {
  it("round-trips a single table", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Name", "Price", "Stock"],
        ["Widget", 9.99, 100],
        ["Gadget", 19.99, 50],
      ],
      tables: [
        {
          name: "Products",
          range: "A1:C3",
          columns: [{ name: "Name" }, { name: "Price" }, { name: "Stock" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets).toHaveLength(1)
    expect(workbook.sheets[0].tables).toBeDefined()
    expect(workbook.sheets[0].tables).toHaveLength(1)

    const table = workbook.sheets[0].tables![0]
    expect(table.name).toBe("Products")
    expect(table.range).toBe("A1:C3")
    expect(table.columns).toHaveLength(3)
    expect(table.columns[0].name).toBe("Name")
    expect(table.columns[1].name).toBe("Price")
    expect(table.columns[2].name).toBe("Stock")
  })

  it("round-trips table with display name", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A"], [1]],
      tables: [
        {
          name: "T1",
          displayName: "My Table",
          range: "A1:A2",
          columns: [{ name: "A" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    const table = workbook.sheets[0].tables![0]
    expect(table.name).toBe("T1")
    expect(table.displayName).toBe("My Table")
  })

  it("round-trips table columns", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["First", "Second", "Third"],
        [1, 2, 3],
      ],
      tables: [
        {
          name: "Table1",
          range: "A1:C2",
          columns: [{ name: "First" }, { name: "Second" }, { name: "Third" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    const table = workbook.sheets[0].tables![0]
    expect(table.columns).toHaveLength(3)
    expect(table.columns.map((c) => c.name)).toEqual(["First", "Second", "Third"])
  })

  it("round-trips table range", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["X", "Y"],
        [1, 2],
        [3, 4],
        [5, 6],
      ],
      tables: [
        {
          name: "BigTable",
          range: "A1:B4",
          columns: [{ name: "X" }, { name: "Y" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    const table = workbook.sheets[0].tables![0]
    expect(table.range).toBe("A1:B4")
  })

  it("round-trips table style", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A"], [1]],
      tables: [
        {
          name: "Table1",
          range: "A1:A2",
          columns: [{ name: "A" }],
          style: "TableStyleLight9",
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    const table = workbook.sheets[0].tables![0]
    expect(table.style).toBe("TableStyleLight9")
  })

  it("round-trips showRowStripes and showColumnStripes", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A"], [1]],
      tables: [
        {
          name: "Table1",
          range: "A1:A2",
          columns: [{ name: "A" }],
          showRowStripes: false,
          showColumnStripes: true,
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    const table = workbook.sheets[0].tables![0]
    expect(table.showRowStripes).toBe(false)
    expect(table.showColumnStripes).toBe(true)
  })

  it("round-trips table with total row", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Item", "Price"],
        ["A", 10],
        ["B", 20],
        ["Total", 30],
      ],
      tables: [
        {
          name: "Table1",
          range: "A1:B4",
          columns: [
            { name: "Item", totalLabel: "Total" },
            { name: "Price", totalFunction: "sum" },
          ],
          showTotalRow: true,
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    const table = workbook.sheets[0].tables![0]
    expect(table.showTotalRow).toBe(true)
    expect(table.columns[0].totalLabel).toBe("Total")
    expect(table.columns[1].totalFunction).toBe("sum")
  })

  it("round-trips multiple tables on different sheets", async () => {
    const sheet1: WriteSheet = {
      name: "Sheet1",
      rows: [["A"], [1]],
      tables: [
        {
          name: "T1",
          range: "A1:A2",
          columns: [{ name: "A" }],
        },
      ],
    }
    const sheet2: WriteSheet = {
      name: "Sheet2",
      rows: [["B"], [2]],
      tables: [
        {
          name: "T2",
          range: "A1:A2",
          columns: [{ name: "B" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].tables).toHaveLength(1)
    expect(workbook.sheets[0].tables![0].name).toBe("T1")
    expect(workbook.sheets[0].tables![0].columns[0].name).toBe("A")

    expect(workbook.sheets[1].tables).toHaveLength(1)
    expect(workbook.sheets[1].tables![0].name).toBe("T2")
    expect(workbook.sheets[1].tables![0].columns[0].name).toBe("B")
  })

  it("sheet without tables has no tables array after read", async () => {
    const sheet1: WriteSheet = {
      name: "WithTable",
      rows: [["A"], [1]],
      tables: [
        {
          name: "Table1",
          range: "A1:A2",
          columns: [{ name: "A" }],
        },
      ],
    }
    const sheet2: WriteSheet = {
      name: "NoTable",
      rows: [["B"], [2]],
    }

    const xlsx = await writeXlsx({ sheets: [sheet1, sheet2] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].tables).toHaveLength(1)
    expect(workbook.sheets[1].tables).toBeUndefined()
  })

  it("round-trips multiple tables on the same sheet", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["A", "B", null, "C", "D"],
        [1, 2, null, 3, 4],
      ],
      tables: [
        {
          name: "Left",
          range: "A1:B2",
          columns: [{ name: "A" }, { name: "B" }],
        },
        {
          name: "Right",
          range: "D1:E2",
          columns: [{ name: "C" }, { name: "D" }],
        },
      ],
    }

    const xlsx = await writeXlsx({ sheets: [sheet] })
    const workbook = await readXlsx(xlsx)

    expect(workbook.sheets[0].tables).toHaveLength(2)
    expect(workbook.sheets[0].tables![0].name).toBe("Left")
    expect(workbook.sheets[0].tables![1].name).toBe("Right")
  })
})
