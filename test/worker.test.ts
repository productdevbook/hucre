import { describe, it, expect } from "vitest"
import { serializeWorkbook, deserializeWorkbook, WORKER_SAFE_FUNCTIONS } from "../src/worker"
import type { Workbook, Sheet, Cell } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

function makeSheet(overrides: Partial<Sheet> = {}): Sheet {
  return {
    name: "Sheet1",
    rows: [],
    ...overrides,
  }
}

function makeWorkbook(overrides: Partial<Workbook> = {}): Workbook {
  return {
    sheets: [makeSheet()],
    ...overrides,
  }
}

// ── serializeWorkbook ───────────────────────────────────────────────

describe("serializeWorkbook", () => {
  it("converts cells Map to plain object", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", { value: "hello", type: "string" })
    cells.set("1,2", { value: 42, type: "number" })

    const wb = makeWorkbook({
      sheets: [makeSheet({ cells })],
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.sheets[0]!.cells).toBeDefined()
    expect(serialized.sheets[0]!.cells).not.toBeInstanceOf(Map)
    expect(serialized.sheets[0]!.cells!["0,0"]).toEqual({
      value: "hello",
      type: "string",
    })
    expect(serialized.sheets[0]!.cells!["1,2"]).toEqual({
      value: 42,
      type: "number",
    })
  })

  it("converts Date values in rows to ISO-string markers", () => {
    const date = new Date("2025-06-15T10:30:00.000Z")
    const wb = makeWorkbook({
      sheets: [makeSheet({ rows: [["label", date, 42]] })],
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.sheets[0]!.rows[0]![0]).toBe("label")
    expect(serialized.sheets[0]!.rows[0]![1]).toEqual({
      __date: "2025-06-15T10:30:00.000Z",
    })
    expect(serialized.sheets[0]!.rows[0]![2]).toBe(42)
  })

  it("converts Date values in cell objects to ISO-string markers", () => {
    const date = new Date("2024-01-01T00:00:00.000Z")
    const cells = new Map<string, Cell>()
    cells.set("0,0", { value: date, type: "date" })

    const wb = makeWorkbook({
      sheets: [makeSheet({ cells })],
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.sheets[0]!.cells!["0,0"]!.value).toEqual({
      __date: "2024-01-01T00:00:00.000Z",
    })
  })

  it("converts rowDefs Map to array of entries", () => {
    const rowDefs = new Map<number, { height?: number; hidden?: boolean }>()
    rowDefs.set(0, { height: 20 })
    rowDefs.set(3, { hidden: true })

    const wb = makeWorkbook({
      sheets: [makeSheet({ rowDefs })],
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.sheets[0]!.rowDefs).toEqual([
      [0, { height: 20 }],
      [3, { hidden: true }],
    ])
  })

  it("converts WorkbookProperties Date fields to markers", () => {
    const created = new Date("2025-01-01T00:00:00.000Z")
    const modified = new Date("2025-06-15T12:00:00.000Z")

    const wb = makeWorkbook({
      properties: {
        title: "Test",
        creator: "Author",
        created,
        modified,
      },
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.properties!.title).toBe("Test")
    expect(serialized.properties!.creator).toBe("Author")
    expect(serialized.properties!.created).toEqual({
      __date: "2025-01-01T00:00:00.000Z",
    })
    expect(serialized.properties!.modified).toEqual({
      __date: "2025-06-15T12:00:00.000Z",
    })
  })

  it("converts custom property Date values to markers", () => {
    const wb = makeWorkbook({
      properties: {
        custom: {
          deadline: new Date("2026-12-31T23:59:59.000Z"),
          count: 5,
          active: true,
          label: "test",
        },
      },
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.properties!.custom!["deadline"]).toEqual({
      __date: "2026-12-31T23:59:59.000Z",
    })
    expect(serialized.properties!.custom!["count"]).toBe(5)
    expect(serialized.properties!.custom!["active"]).toBe(true)
    expect(serialized.properties!.custom!["label"]).toBe("test")
  })

  it("converts SheetImage data Uint8Array to plain array", () => {
    const imageData = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])
    const wb = makeWorkbook({
      sheets: [
        makeSheet({
          images: [
            {
              data: imageData,
              type: "png",
              anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
            },
          ],
        }),
      ],
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.sheets[0]!.images).toHaveLength(1)
    expect(serialized.sheets[0]!.images![0]!.data).toEqual([137, 80, 78, 71, 13, 10, 26, 10])
    expect(Array.isArray(serialized.sheets[0]!.images![0]!.data)).toBe(true)
  })

  it("preserves null values", () => {
    const wb = makeWorkbook({
      sheets: [makeSheet({ rows: [[null, "text", null]] })],
    })

    const serialized = serializeWorkbook(wb)
    expect(serialized.sheets[0]!.rows[0]).toEqual([null, "text", null])
  })

  it("preserves boolean values", () => {
    const wb = makeWorkbook({
      sheets: [makeSheet({ rows: [[true, false]] })],
    })

    const serialized = serializeWorkbook(wb)
    expect(serialized.sheets[0]!.rows[0]).toEqual([true, false])
  })

  it("serializes formulaResult Date values", () => {
    const date = new Date("2025-03-25T00:00:00.000Z")
    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: "=TODAY()",
      type: "formula",
      formula: "TODAY()",
      formulaResult: date,
    })

    const wb = makeWorkbook({
      sheets: [makeSheet({ cells })],
    })

    const serialized = serializeWorkbook(wb)

    expect(serialized.sheets[0]!.cells!["0,0"]!.formulaResult).toEqual({
      __date: "2025-03-25T00:00:00.000Z",
    })
  })
})

// ── deserializeWorkbook ─────────────────────────────────────────────

describe("deserializeWorkbook", () => {
  it("restores cells from plain object to Map", () => {
    const serialized = {
      sheets: [
        {
          name: "Sheet1",
          rows: [],
          cells: {
            "0,0": { value: "hello", type: "string" as const },
            "1,2": { value: 42, type: "number" as const },
          },
        },
      ],
    }

    const wb = deserializeWorkbook(serialized)

    expect(wb.sheets[0]!.cells).toBeInstanceOf(Map)
    expect(wb.sheets[0]!.cells!.get("0,0")).toEqual({
      value: "hello",
      type: "string",
    })
    expect(wb.sheets[0]!.cells!.get("1,2")).toEqual({
      value: 42,
      type: "number",
    })
  })

  it("restores Date values from ISO-string markers in rows", () => {
    const serialized = {
      sheets: [
        {
          name: "Sheet1",
          rows: [["label", { __date: "2025-06-15T10:30:00.000Z" }, 42]],
        },
      ],
    }

    const wb = deserializeWorkbook(serialized)

    expect(wb.sheets[0]!.rows[0]![0]).toBe("label")
    expect(wb.sheets[0]!.rows[0]![1]).toBeInstanceOf(Date)
    expect((wb.sheets[0]!.rows[0]![1] as Date).toISOString()).toBe("2025-06-15T10:30:00.000Z")
    expect(wb.sheets[0]!.rows[0]![2]).toBe(42)
  })

  it("restores Date values in cell objects", () => {
    const serialized = {
      sheets: [
        {
          name: "Sheet1",
          rows: [],
          cells: {
            "0,0": {
              value: { __date: "2024-01-01T00:00:00.000Z" },
              type: "date" as const,
            },
          },
        },
      ],
    }

    const wb = deserializeWorkbook(serialized)

    const cell = wb.sheets[0]!.cells!.get("0,0")!
    expect(cell.value).toBeInstanceOf(Date)
    expect((cell.value as Date).toISOString()).toBe("2024-01-01T00:00:00.000Z")
  })

  it("restores rowDefs from array of entries to Map", () => {
    const serialized = {
      sheets: [
        {
          name: "Sheet1",
          rows: [],
          rowDefs: [
            [0, { height: 20 }],
            [3, { hidden: true }],
          ] as Array<[number, { height?: number; hidden?: boolean }]>,
        },
      ],
    }

    const wb = deserializeWorkbook(serialized)

    expect(wb.sheets[0]!.rowDefs).toBeInstanceOf(Map)
    expect(wb.sheets[0]!.rowDefs!.get(0)).toEqual({ height: 20 })
    expect(wb.sheets[0]!.rowDefs!.get(3)).toEqual({ hidden: true })
  })

  it("restores WorkbookProperties Date fields", () => {
    const serialized = {
      sheets: [{ name: "Sheet1", rows: [] }],
      properties: {
        title: "Test",
        created: { __date: "2025-01-01T00:00:00.000Z" },
        modified: { __date: "2025-06-15T12:00:00.000Z" },
      },
    }

    const wb = deserializeWorkbook(serialized)

    expect(wb.properties!.title).toBe("Test")
    expect(wb.properties!.created).toBeInstanceOf(Date)
    expect(wb.properties!.created!.toISOString()).toBe("2025-01-01T00:00:00.000Z")
    expect(wb.properties!.modified).toBeInstanceOf(Date)
    expect(wb.properties!.modified!.toISOString()).toBe("2025-06-15T12:00:00.000Z")
  })

  it("restores custom property Date values", () => {
    const serialized = {
      sheets: [{ name: "Sheet1", rows: [] }],
      properties: {
        custom: {
          deadline: { __date: "2026-12-31T23:59:59.000Z" },
          count: 5,
          active: true,
          label: "test",
        },
      },
    }

    const wb = deserializeWorkbook(serialized)

    expect(wb.properties!.custom!["deadline"]).toBeInstanceOf(Date)
    expect((wb.properties!.custom!["deadline"] as Date).toISOString()).toBe(
      "2026-12-31T23:59:59.000Z",
    )
    expect(wb.properties!.custom!["count"]).toBe(5)
    expect(wb.properties!.custom!["active"]).toBe(true)
    expect(wb.properties!.custom!["label"]).toBe("test")
  })

  it("restores SheetImage data from plain array to Uint8Array", () => {
    const serialized = {
      sheets: [
        {
          name: "Sheet1",
          rows: [],
          images: [
            {
              data: [137, 80, 78, 71, 13, 10, 26, 10],
              type: "png" as const,
              anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
            },
          ],
        },
      ],
    }

    const wb = deserializeWorkbook(serialized)

    expect(wb.sheets[0]!.images).toHaveLength(1)
    expect(wb.sheets[0]!.images![0]!.data).toBeInstanceOf(Uint8Array)
    expect(wb.sheets[0]!.images![0]!.data).toEqual(
      new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    )
  })

  it("restores formulaResult Date values", () => {
    const serialized = {
      sheets: [
        {
          name: "Sheet1",
          rows: [],
          cells: {
            "0,0": {
              value: "=TODAY()",
              type: "formula" as const,
              formula: "TODAY()",
              formulaResult: { __date: "2025-03-25T00:00:00.000Z" },
            },
          },
        },
      ],
    }

    const wb = deserializeWorkbook(serialized)

    const cell = wb.sheets[0]!.cells!.get("0,0")!
    expect(cell.formulaResult).toBeInstanceOf(Date)
    expect((cell.formulaResult as Date).toISOString()).toBe("2025-03-25T00:00:00.000Z")
  })
})

// ── Round-trip: serialize then deserialize ───────────────────────────

describe("round-trip: serialize -> deserialize", () => {
  it("preserves all cell types through round-trip", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", { value: "text", type: "string" })
    cells.set("0,1", { value: 123.45, type: "number" })
    cells.set("0,2", { value: true, type: "boolean" })
    cells.set("0,3", { value: new Date("2025-06-15T10:30:00.000Z"), type: "date" })
    cells.set("0,4", { value: null, type: "empty" })
    cells.set("0,5", { value: "#REF!", type: "error" })

    const wb = makeWorkbook({
      sheets: [makeSheet({ rows: [], cells })],
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))

    expect(result.sheets[0]!.cells).toBeInstanceOf(Map)
    expect(result.sheets[0]!.cells!.size).toBe(6)

    // String
    expect(result.sheets[0]!.cells!.get("0,0")!.value).toBe("text")
    expect(result.sheets[0]!.cells!.get("0,0")!.type).toBe("string")

    // Number
    expect(result.sheets[0]!.cells!.get("0,1")!.value).toBe(123.45)

    // Boolean
    expect(result.sheets[0]!.cells!.get("0,2")!.value).toBe(true)

    // Date
    const dateCell = result.sheets[0]!.cells!.get("0,3")!
    expect(dateCell.value).toBeInstanceOf(Date)
    expect((dateCell.value as Date).toISOString()).toBe("2025-06-15T10:30:00.000Z")

    // Null / empty
    expect(result.sheets[0]!.cells!.get("0,4")!.value).toBeNull()

    // Error string
    expect(result.sheets[0]!.cells!.get("0,5")!.value).toBe("#REF!")
  })

  it("preserves an empty workbook", () => {
    const wb: Workbook = { sheets: [] }

    const result = deserializeWorkbook(serializeWorkbook(wb))

    expect(result.sheets).toEqual([])
    expect(result.properties).toBeUndefined()
    expect(result.namedRanges).toBeUndefined()
  })

  it("preserves multiple sheets", () => {
    const wb = makeWorkbook({
      sheets: [
        makeSheet({ name: "Data", rows: [["A", 1]] }),
        makeSheet({ name: "Summary", rows: [["Total", 100]] }),
        makeSheet({ name: "Hidden", rows: [], hidden: true }),
      ],
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))

    expect(result.sheets).toHaveLength(3)
    expect(result.sheets[0]!.name).toBe("Data")
    expect(result.sheets[0]!.rows).toEqual([["A", 1]])
    expect(result.sheets[1]!.name).toBe("Summary")
    expect(result.sheets[1]!.rows).toEqual([["Total", 100]])
    expect(result.sheets[2]!.name).toBe("Hidden")
    expect(result.sheets[2]!.hidden).toBe(true)
  })

  it("preserves styles, hyperlinks, and comments on cells", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: "styled",
      type: "string",
      style: {
        font: { bold: true, size: 14, color: { rgb: "FF0000" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "double" },
        },
        alignment: { horizontal: "center", wrapText: true },
        numFmt: "#,##0.00",
      },
    })
    cells.set("1,0", {
      value: "link",
      type: "string",
      hyperlink: {
        target: "https://example.com",
        tooltip: "Example",
        display: "Click here",
      },
    })
    cells.set("2,0", {
      value: "noted",
      type: "string",
      comment: {
        author: "Alice",
        text: "This is a comment",
        richText: [{ text: "Bold part", font: { bold: true } }, { text: " normal part" }],
      },
    })

    const wb = makeWorkbook({
      sheets: [makeSheet({ cells })],
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))
    const resultCells = result.sheets[0]!.cells!

    // Style
    const styledCell = resultCells.get("0,0")!
    expect(styledCell.style!.font!.bold).toBe(true)
    expect(styledCell.style!.font!.size).toBe(14)
    expect(styledCell.style!.font!.color!.rgb).toBe("FF0000")
    expect(styledCell.style!.fill!.type).toBe("pattern")
    expect(styledCell.style!.border!.top!.style).toBe("thin")
    expect(styledCell.style!.alignment!.horizontal).toBe("center")
    expect(styledCell.style!.numFmt).toBe("#,##0.00")

    // Hyperlink
    const linkCell = resultCells.get("1,0")!
    expect(linkCell.hyperlink!.target).toBe("https://example.com")
    expect(linkCell.hyperlink!.tooltip).toBe("Example")
    expect(linkCell.hyperlink!.display).toBe("Click here")

    // Comment
    const commentCell = resultCells.get("2,0")!
    expect(commentCell.comment!.author).toBe("Alice")
    expect(commentCell.comment!.text).toBe("This is a comment")
    expect(commentCell.comment!.richText).toHaveLength(2)
    expect(commentCell.comment!.richText![0]!.font!.bold).toBe(true)
  })

  it("preserves rowDefs through round-trip", () => {
    const rowDefs = new Map<number, { height?: number; hidden?: boolean }>()
    rowDefs.set(0, { height: 30 })
    rowDefs.set(5, { hidden: true })
    rowDefs.set(10, { height: 15, hidden: false })

    const wb = makeWorkbook({
      sheets: [makeSheet({ rowDefs })],
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))

    expect(result.sheets[0]!.rowDefs).toBeInstanceOf(Map)
    expect(result.sheets[0]!.rowDefs!.size).toBe(3)
    expect(result.sheets[0]!.rowDefs!.get(0)).toEqual({ height: 30 })
    expect(result.sheets[0]!.rowDefs!.get(5)).toEqual({ hidden: true })
    expect(result.sheets[0]!.rowDefs!.get(10)).toEqual({
      height: 15,
      hidden: false,
    })
  })

  it("preserves images through round-trip", () => {
    const imageData = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

    const wb = makeWorkbook({
      sheets: [
        makeSheet({
          images: [
            {
              data: imageData,
              type: "png",
              anchor: { from: { row: 0, col: 0 }, to: { row: 10, col: 5 } },
              width: 200,
              height: 150,
            },
            {
              data: new Uint8Array([0xff, 0xd8, 0xff, 0xe0]),
              type: "jpeg",
              anchor: { from: { row: 3, col: 2 } },
            },
          ],
        }),
      ],
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))

    expect(result.sheets[0]!.images).toHaveLength(2)

    const img1 = result.sheets[0]!.images![0]!
    expect(img1.data).toBeInstanceOf(Uint8Array)
    expect(img1.data).toEqual(imageData)
    expect(img1.type).toBe("png")
    expect(img1.anchor.from).toEqual({ row: 0, col: 0 })
    expect(img1.anchor.to).toEqual({ row: 10, col: 5 })
    expect(img1.width).toBe(200)
    expect(img1.height).toBe(150)

    const img2 = result.sheets[0]!.images![1]!
    expect(img2.data).toBeInstanceOf(Uint8Array)
    expect(img2.type).toBe("jpeg")
    expect(img2.anchor.from).toEqual({ row: 3, col: 2 })
    expect(img2.anchor.to).toBeUndefined()
  })

  it("preserves namedRanges and dateSystem", () => {
    const wb = makeWorkbook({
      namedRanges: [
        { name: "MyRange", range: "Sheet1!$A$1:$D$10" },
        { name: "LocalRange", range: "Sheet1!$B$2", scope: "Sheet1", comment: "test" },
      ],
      dateSystem: "1904",
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))

    expect(result.namedRanges).toHaveLength(2)
    expect(result.namedRanges![0]).toEqual({
      name: "MyRange",
      range: "Sheet1!$A$1:$D$10",
    })
    expect(result.namedRanges![1]).toEqual({
      name: "LocalRange",
      range: "Sheet1!$B$2",
      scope: "Sheet1",
      comment: "test",
    })
    expect(result.dateSystem).toBe("1904")
  })

  it("preserves defaultFont and activeSheet", () => {
    const wb = makeWorkbook({
      defaultFont: { name: "Calibri", size: 11, bold: false },
      activeSheet: 2,
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))

    expect(result.defaultFont).toEqual({
      name: "Calibri",
      size: 11,
      bold: false,
    })
    expect(result.activeSheet).toBe(2)
  })

  it("preserves sheet-level features (merges, validations, freeze, filter, protection)", () => {
    const wb = makeWorkbook({
      sheets: [
        makeSheet({
          merges: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 3 }],
          dataValidations: [
            {
              type: "list",
              range: "A1:A10",
              values: ["Yes", "No", "Maybe"],
              allowBlank: true,
            },
          ],
          autoFilter: { range: "A1:D100" },
          freezePane: { rows: 1, columns: 2 },
          protection: { sheet: true, password: "secret" },
          pageSetup: {
            orientation: "landscape",
            paperSize: "a4",
          },
          headerFooter: {
            oddHeader: "&CPage &P",
            oddFooter: "&LLeft&RRight",
          },
          view: { showGridLines: false, zoomScale: 150 },
          veryHidden: true,
          tables: [
            {
              name: "Table1",
              columns: [{ name: "Col1" }, { name: "Col2" }],
              range: "A1:B10",
              style: "TableStyleMedium2",
            },
          ],
        }),
      ],
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))
    const sheet = result.sheets[0]!

    expect(sheet.merges).toEqual([{ startRow: 0, startCol: 0, endRow: 2, endCol: 3 }])
    expect(sheet.dataValidations).toHaveLength(1)
    expect(sheet.dataValidations![0]!.type).toBe("list")
    expect(sheet.dataValidations![0]!.values).toEqual(["Yes", "No", "Maybe"])
    expect(sheet.autoFilter).toEqual({ range: "A1:D100" })
    expect(sheet.freezePane).toEqual({ rows: 1, columns: 2 })
    expect(sheet.protection!.sheet).toBe(true)
    expect(sheet.pageSetup!.orientation).toBe("landscape")
    expect(sheet.headerFooter!.oddHeader).toBe("&CPage &P")
    expect(sheet.view!.showGridLines).toBe(false)
    expect(sheet.view!.zoomScale).toBe(150)
    expect(sheet.veryHidden).toBe(true)
    expect(sheet.tables).toHaveLength(1)
    expect(sheet.tables![0]!.name).toBe("Table1")
  })

  it("preserves WorkbookProperties through round-trip", () => {
    const created = new Date("2025-01-01T00:00:00.000Z")
    const modified = new Date("2025-06-15T12:00:00.000Z")
    const deadline = new Date("2026-12-31T23:59:59.000Z")

    const wb = makeWorkbook({
      properties: {
        title: "My Workbook",
        subject: "Testing",
        creator: "Alice",
        keywords: "test, worker",
        description: "A test workbook",
        lastModifiedBy: "Bob",
        created,
        modified,
        company: "Acme",
        manager: "Carol",
        category: "Reports",
        custom: {
          deadline,
          version: 3,
          approved: true,
          label: "final",
        },
      },
    })

    const result = deserializeWorkbook(serializeWorkbook(wb))
    const props = result.properties!

    expect(props.title).toBe("My Workbook")
    expect(props.subject).toBe("Testing")
    expect(props.creator).toBe("Alice")
    expect(props.keywords).toBe("test, worker")
    expect(props.description).toBe("A test workbook")
    expect(props.lastModifiedBy).toBe("Bob")
    expect(props.company).toBe("Acme")
    expect(props.manager).toBe("Carol")
    expect(props.category).toBe("Reports")

    expect(props.created).toBeInstanceOf(Date)
    expect(props.created!.toISOString()).toBe("2025-01-01T00:00:00.000Z")
    expect(props.modified).toBeInstanceOf(Date)
    expect(props.modified!.toISOString()).toBe("2025-06-15T12:00:00.000Z")

    expect(props.custom!["deadline"]).toBeInstanceOf(Date)
    expect((props.custom!["deadline"] as Date).toISOString()).toBe("2026-12-31T23:59:59.000Z")
    expect(props.custom!["version"]).toBe(3)
    expect(props.custom!["approved"]).toBe(true)
    expect(props.custom!["label"]).toBe("final")
  })

  it("handles a complex workbook with all features combined", () => {
    const date1 = new Date("2025-01-15T08:00:00.000Z")
    const date2 = new Date("2025-06-20T16:30:00.000Z")

    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: "Header",
      type: "string",
      style: { font: { bold: true } },
    })
    cells.set("1,0", {
      value: date1,
      type: "date",
      comment: { text: "Start date", author: "System" },
    })
    cells.set("1,1", {
      value: 42.5,
      type: "number",
      hyperlink: { target: "https://example.com" },
    })

    const rowDefs = new Map<number, { height?: number }>()
    rowDefs.set(0, { height: 25 })

    const wb: Workbook = {
      sheets: [
        {
          name: "Report",
          rows: [
            ["Header", "Value"],
            [date1, 42.5],
            [null, date2],
          ],
          cells,
          rowDefs,
          columns: [
            { header: "A", width: 20 },
            { header: "B", width: 15 },
          ],
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
          freezePane: { rows: 1 },
        },
        {
          name: "Empty",
          rows: [],
          hidden: true,
        },
      ],
      properties: {
        title: "Complex Report",
        created: date1,
        custom: { revision: 5, dueDate: date2 },
      },
      namedRanges: [{ name: "Data", range: "Report!$A$2:$B$3" }],
      dateSystem: "1900",
      activeSheet: 0,
    }

    const result = deserializeWorkbook(serializeWorkbook(wb))

    // Verify sheets
    expect(result.sheets).toHaveLength(2)
    expect(result.sheets[0]!.name).toBe("Report")
    expect(result.sheets[1]!.name).toBe("Empty")
    expect(result.sheets[1]!.hidden).toBe(true)

    // Verify rows with dates
    expect(result.sheets[0]!.rows[0]).toEqual(["Header", "Value"])
    expect(result.sheets[0]!.rows[1]![0]).toBeInstanceOf(Date)
    expect((result.sheets[0]!.rows[1]![0] as Date).toISOString()).toBe(date1.toISOString())
    expect(result.sheets[0]!.rows[2]![1]).toBeInstanceOf(Date)

    // Verify cells Map
    expect(result.sheets[0]!.cells).toBeInstanceOf(Map)
    expect(result.sheets[0]!.cells!.size).toBe(3)
    expect(result.sheets[0]!.cells!.get("1,0")!.value).toBeInstanceOf(Date)
    expect(result.sheets[0]!.cells!.get("1,0")!.comment!.text).toBe("Start date")

    // Verify rowDefs Map
    expect(result.sheets[0]!.rowDefs).toBeInstanceOf(Map)
    expect(result.sheets[0]!.rowDefs!.get(0)).toEqual({ height: 25 })

    // Verify properties
    expect(result.properties!.title).toBe("Complex Report")
    expect(result.properties!.created).toBeInstanceOf(Date)
    expect(result.properties!.custom!["dueDate"]).toBeInstanceOf(Date)
    expect(result.properties!.custom!["revision"]).toBe(5)

    // Verify other fields
    expect(result.namedRanges).toHaveLength(1)
    expect(result.dateSystem).toBe("1900")
    expect(result.activeSheet).toBe(0)
  })

  it("is safe through JSON.parse(JSON.stringify()) pipeline", () => {
    const cells = new Map<string, Cell>()
    cells.set("0,0", {
      value: new Date("2025-03-25T00:00:00.000Z"),
      type: "date",
    })

    const rowDefs = new Map<number, { height?: number }>()
    rowDefs.set(0, { height: 20 })

    const wb = makeWorkbook({
      sheets: [
        makeSheet({
          rows: [[new Date("2025-01-01T00:00:00.000Z"), "text"]],
          cells,
          rowDefs,
        }),
      ],
      properties: {
        created: new Date("2025-06-01T00:00:00.000Z"),
      },
    })

    // Simulate JSON serialization (e.g., through localStorage or network)
    const serialized = serializeWorkbook(wb)
    const jsonString = JSON.stringify(serialized)
    const parsed = JSON.parse(jsonString)
    const result = deserializeWorkbook(parsed)

    // Dates should be restored
    expect(result.sheets[0]!.rows[0]![0]).toBeInstanceOf(Date)
    expect((result.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2025-01-01T00:00:00.000Z")

    // Cells Map should be restored
    expect(result.sheets[0]!.cells).toBeInstanceOf(Map)
    expect(result.sheets[0]!.cells!.get("0,0")!.value).toBeInstanceOf(Date)

    // RowDefs Map should be restored
    expect(result.sheets[0]!.rowDefs).toBeInstanceOf(Map)
    expect(result.sheets[0]!.rowDefs!.get(0)).toEqual({ height: 20 })

    // Properties dates should be restored
    expect(result.properties!.created).toBeInstanceOf(Date)
  })
})

// ── WORKER_SAFE_FUNCTIONS ───────────────────────────────────────────

describe("WORKER_SAFE_FUNCTIONS", () => {
  it("is a non-empty array of strings", () => {
    expect(Array.isArray(WORKER_SAFE_FUNCTIONS)).toBe(true)
    expect(WORKER_SAFE_FUNCTIONS.length).toBeGreaterThan(0)
    for (const fn of WORKER_SAFE_FUNCTIONS) {
      expect(typeof fn).toBe("string")
    }
  })

  it("includes core API functions", () => {
    expect(WORKER_SAFE_FUNCTIONS).toContain("read")
    expect(WORKER_SAFE_FUNCTIONS).toContain("write")
    expect(WORKER_SAFE_FUNCTIONS).toContain("readXlsx")
    expect(WORKER_SAFE_FUNCTIONS).toContain("writeXlsx")
    expect(WORKER_SAFE_FUNCTIONS).toContain("readOds")
    expect(WORKER_SAFE_FUNCTIONS).toContain("writeOds")
    expect(WORKER_SAFE_FUNCTIONS).toContain("parseCsv")
  })

  it("includes utility functions", () => {
    expect(WORKER_SAFE_FUNCTIONS).toContain("serialToDate")
    expect(WORKER_SAFE_FUNCTIONS).toContain("dateToSerial")
    expect(WORKER_SAFE_FUNCTIONS).toContain("insertRows")
    expect(WORKER_SAFE_FUNCTIONS).toContain("deleteRows")
    expect(WORKER_SAFE_FUNCTIONS).toContain("validateWithSchema")
  })
})
