import { describe, it, expect } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { parseXml } from "../src/xml/parser"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeComments } from "../src/xlsx/comments-writer"
import { parseComments } from "../src/xlsx/comments-reader"
import type { Cell } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8")
const encoder = new TextEncoder()

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

// ── writeComments unit tests ─────────────────────────────────────────

describe("writeComments", () => {
  it("returns null when no cells have comments", () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", { value: "Hello" })
    const result = writeComments(cells, 0)
    expect(result).toBeNull()
  })

  it("writes single comment on a cell", () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "This is a comment" },
    })
    const result = writeComments(cells, 0)
    expect(result).not.toBeNull()
    expect(result!.comments).toHaveLength(1)
    expect(result!.comments[0].ref).toBe("A1")
    expect(result!.comments[0].text).toBe("This is a comment")

    // Verify comments XML structure
    expect(result!.commentsXml).toContain("<comments")
    expect(result!.commentsXml).toContain("<authors>")
    expect(result!.commentsXml).toContain("<author")
    expect(result!.commentsXml).toContain("<commentList>")
    expect(result!.commentsXml).toContain('ref="A1"')
    expect(result!.commentsXml).toContain("This is a comment")
  })

  it("writes comment with author", () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "Review this", author: "John Doe" },
    })
    const result = writeComments(cells, 0)
    expect(result).not.toBeNull()
    expect(result!.comments[0].author).toBe("John Doe")
    expect(result!.commentsXml).toContain("John Doe")
    expect(result!.commentsXml).toContain('authorId="0"')
  })

  it("writes multiple comments on different cells", () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Cell A1",
      comment: { text: "Comment on A1" },
    })
    cells.set("2,1", {
      value: "Cell B3",
      comment: { text: "Comment on B3", author: "Alice" },
    })
    cells.set("1,0", {
      value: "Cell A2",
      comment: { text: "Comment on A2", author: "Bob" },
    })
    const result = writeComments(cells, 0)
    expect(result).not.toBeNull()
    expect(result!.comments).toHaveLength(3)

    // Comments should be sorted by row then column
    expect(result!.comments[0].ref).toBe("A1")
    expect(result!.comments[1].ref).toBe("A2")
    expect(result!.comments[2].ref).toBe("B3")
  })

  it("generates VML drawing with correct row and column", () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "A1",
      comment: { text: "Comment" },
    })
    const result = writeComments(cells, 0)
    expect(result).not.toBeNull()

    // VML should contain the correct row and column
    expect(result!.vmlXml).toContain("<x:Row>0</x:Row>")
    expect(result!.vmlXml).toContain("<x:Column>0</x:Column>")
    expect(result!.vmlXml).toContain("shapetype")
    expect(result!.vmlXml).toContain("ClientData")
    expect(result!.vmlXml).toContain('ObjectType="Note"')
  })

  it("generates VML with correct positions for multiple comments", () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "A1",
      comment: { text: "First" },
    })
    cells.set("4,2", {
      value: "C5",
      comment: { text: "Second" },
    })
    const result = writeComments(cells, 0)
    expect(result).not.toBeNull()

    expect(result!.vmlXml).toContain("<x:Row>0</x:Row>")
    expect(result!.vmlXml).toContain("<x:Column>0</x:Column>")
    expect(result!.vmlXml).toContain("<x:Row>4</x:Row>")
    expect(result!.vmlXml).toContain("<x:Column>2</x:Column>")
  })

  it("deduplicates authors", () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "A1",
      comment: { text: "First", author: "Alice" },
    })
    cells.set("1,0", {
      value: "A2",
      comment: { text: "Second", author: "Alice" },
    })
    cells.set("2,0", {
      value: "A3",
      comment: { text: "Third", author: "Bob" },
    })
    const result = writeComments(cells, 0)
    expect(result).not.toBeNull()

    // Parse the XML to verify author count
    const doc = parseXml(result!.commentsXml)
    const authors = findChild(doc, "authors")
    const authorElements = findChildren(authors, "author")
    expect(authorElements).toHaveLength(2) // Alice and Bob
  })
})

// ── parseComments unit tests ─────────────────────────────────────────

describe("parseComments", () => {
  it("parses basic comment", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Test User</author></authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t>Hello</t></r></text>
    </comment>
  </commentList>
</comments>`
    const comments = parseComments(xml)
    expect(comments.size).toBe(1)
    expect(comments.get("A1")).toBeDefined()
    expect(comments.get("A1")!.text).toBe("Hello")
    expect(comments.get("A1")!.author).toBe("Test User")
  })

  it("parses multiple comments with multiple authors", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Alice</author><author>Bob</author></authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t>From Alice</t></r></text>
    </comment>
    <comment ref="B2" authorId="1">
      <text><r><t>From Bob</t></r></text>
    </comment>
  </commentList>
</comments>`
    const comments = parseComments(xml)
    expect(comments.size).toBe(2)
    expect(comments.get("A1")!.author).toBe("Alice")
    expect(comments.get("A1")!.text).toBe("From Alice")
    expect(comments.get("B2")!.author).toBe("Bob")
    expect(comments.get("B2")!.text).toBe("From Bob")
  })

  it("parses comment without author", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author></author></authors>
  <commentList>
    <comment ref="C3" authorId="0">
      <text><r><t>No author</t></r></text>
    </comment>
  </commentList>
</comments>`
    const comments = parseComments(xml)
    expect(comments.size).toBe(1)
    expect(comments.get("C3")!.text).toBe("No author")
    // Empty string author should not be set
    expect(comments.get("C3")!.author).toBeUndefined()
  })
})

// ── XLSX Writing Tests ──────────────────────────────────────────────

describe("XLSX comment writing", () => {
  it("writes comments.xml and VML to ZIP", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "A comment", author: "Author" },
    })

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]], cells }],
    })

    // Verify comments file exists
    expect(zipHas(data, "xl/comments1.xml")).toBe(true)

    // Verify VML drawing file exists
    expect(zipHas(data, "xl/drawings/vmlDrawing1.vml")).toBe(true)

    // Verify comments XML content
    const commentsXml = await extractXml(data, "xl/comments1.xml")
    expect(commentsXml).toContain("A comment")
    expect(commentsXml).toContain("Author")
    expect(commentsXml).toContain('ref="A1"')
  })

  it("does not generate comments files when no comments exist", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello", "World"]] }],
    })

    expect(zipHas(data, "xl/comments1.xml")).toBe(false)
    expect(zipHas(data, "xl/drawings/vmlDrawing1.vml")).toBe(false)
  })

  it("includes legacyDrawing in worksheet XML when comments exist", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "Comment" },
    })

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]], cells }],
    })

    const wsXml = await extractXml(data, "xl/worksheets/sheet1.xml")
    expect(wsXml).toContain("legacyDrawing")
  })

  it("does not include legacyDrawing when no comments exist", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]] }],
    })

    const wsXml = await extractXml(data, "xl/worksheets/sheet1.xml")
    expect(wsXml).not.toContain("legacyDrawing")
  })

  it("updates content types when comments exist", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "Comment" },
    })

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]], cells }],
    })

    const ctXml = await extractXml(data, "[Content_Types].xml")
    expect(ctXml).toContain("comments")
    expect(ctXml).toContain("vml")
  })

  it("does not include comment content types when no comments exist", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]] }],
    })

    const ctXml = await extractXml(data, "[Content_Types].xml")
    expect(ctXml).not.toContain("comments")
    expect(ctXml).not.toContain("vml")
  })

  it("includes comment and vmlDrawing relationships in sheet rels", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "Comment" },
    })

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]], cells }],
    })

    // Sheet rels should exist
    expect(zipHas(data, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(true)

    const relsXml = await extractXml(data, "xl/worksheets/_rels/sheet1.xml.rels")
    expect(relsXml).toContain("vmlDrawing")
    expect(relsXml).toContain("comments")
  })

  it("writes comments on multiple sheets", async () => {
    const cells1 = new Map<string, Partial<Cell>>()
    cells1.set("0,0", {
      value: "Sheet1",
      comment: { text: "Comment on Sheet1" },
    })

    const cells2 = new Map<string, Partial<Cell>>()
    cells2.set("0,0", {
      value: "Sheet2",
      comment: { text: "Comment on Sheet2" },
    })

    const data = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["Sheet1"]], cells: cells1 },
        { name: "Sheet2", rows: [["Sheet2"]], cells: cells2 },
      ],
    })

    expect(zipHas(data, "xl/comments1.xml")).toBe(true)
    expect(zipHas(data, "xl/comments2.xml")).toBe(true)
    expect(zipHas(data, "xl/drawings/vmlDrawing1.vml")).toBe(true)
    expect(zipHas(data, "xl/drawings/vmlDrawing2.vml")).toBe(true)

    const comments1 = await extractXml(data, "xl/comments1.xml")
    expect(comments1).toContain("Comment on Sheet1")

    const comments2 = await extractXml(data, "xl/comments2.xml")
    expect(comments2).toContain("Comment on Sheet2")
  })
})

// ── XLSX Reading Tests ──────────────────────────────────────────────

describe("XLSX comment reading", () => {
  /**
   * Helper to create an XLSX with comments in the raw XML
   * for testing the reader path independently.
   */
  async function createXlsxWithComments(options: {
    comments: Array<{
      ref: string
      author: string
      text: string
    }>
    sharedStrings?: string[]
    cellData?: Array<{
      ref: string
      value: string
      type?: string
      ssIndex?: number
    }>
  }): Promise<Uint8Array> {
    const writer = new ZipWriter()
    const enc = encoder

    const hasSharedStrings = (options.sharedStrings?.length ?? 0) > 0

    // [Content_Types].xml
    let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`
    if (hasSharedStrings) {
      contentTypes += `\n  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`
    }
    contentTypes += `\n</Types>`
    writer.add("[Content_Types].xml", enc.encode(contentTypes))

    // _rels/.rels
    writer.add(
      "_rels/.rels",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    )

    // xl/workbook.xml
    writer.add(
      "xl/workbook.xml",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    )

    // xl/_rels/workbook.xml.rels
    let wbRels = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
    if (hasSharedStrings) {
      wbRels += `\n  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`
    }
    writer.add(
      "xl/_rels/workbook.xml.rels",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${wbRels}
</Relationships>`),
    )

    // xl/styles.xml
    writer.add(
      "xl/styles.xml",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`),
    )

    // xl/sharedStrings.xml
    if (hasSharedStrings) {
      const siElements = options.sharedStrings!.map((s) => `<si><t>${s}</t></si>`).join("")
      writer.add(
        "xl/sharedStrings.xml",
        enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${options.sharedStrings!.length}" uniqueCount="${options.sharedStrings!.length}">
  ${siElements}
</sst>`),
      )
    }

    // Build cell data for the worksheet
    let sheetDataXml = ""
    if (options.cellData && options.cellData.length > 0) {
      const cellParts = options.cellData.map((cd) => {
        const typeAttr = cd.type ? ` t="${cd.type}"` : ""
        if (cd.type === "s" && cd.ssIndex !== undefined) {
          return `<c r="${cd.ref}"${typeAttr}><v>${cd.ssIndex}</v></c>`
        }
        return `<c r="${cd.ref}"${typeAttr}><v>${cd.value}</v></c>`
      })
      sheetDataXml = `<row r="1">${cellParts.join("")}</row>`
    }

    // xl/worksheets/sheet1.xml
    writer.add(
      "xl/worksheets/sheet1.xml",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>${sheetDataXml}</sheetData>
  <legacyDrawing r:id="rId1"/>
</worksheet>`),
    )

    // Build comments XML
    const uniqueAuthors = [...new Set(options.comments.map((c) => c.author))]
    const authorElements = uniqueAuthors.map((a) => `<author>${a}</author>`).join("")
    const commentElements = options.comments
      .map((c) => {
        const authorId = uniqueAuthors.indexOf(c.author)
        return `<comment ref="${c.ref}" authorId="${authorId}"><text><r><t>${c.text}</t></r></text></comment>`
      })
      .join("")

    writer.add(
      "xl/comments1.xml",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>${authorElements}</authors>
  <commentList>${commentElements}</commentList>
</comments>`),
    )

    // xl/worksheets/_rels/sheet1.xml.rels
    writer.add(
      "xl/worksheets/_rels/sheet1.xml.rels",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>`),
    )

    // Minimal VML drawing (not parsed by reader, but included for completeness)
    writer.add(
      "xl/drawings/vmlDrawing1.vml",
      enc.encode(
        `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"></xml>`,
      ),
    )

    return writer.build()
  }

  it("reads comment from comments XML", async () => {
    const xlsxData = await createXlsxWithComments({
      comments: [{ ref: "A1", author: "Test User", text: "Hello comment" }],
      sharedStrings: ["Hello"],
      cellData: [{ ref: "A1", value: "Hello", type: "s", ssIndex: 0 }],
    })

    const workbook = await readXlsx(xlsxData)
    const sheet = workbook.sheets[0]

    expect(sheet.cells).toBeDefined()
    const cell = sheet.cells!.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.comment).toBeDefined()
    expect(cell!.comment!.text).toBe("Hello comment")
    expect(cell!.comment!.author).toBe("Test User")
  })

  it("reads multiple comments", async () => {
    const xlsxData = await createXlsxWithComments({
      comments: [
        { ref: "A1", author: "Alice", text: "First comment" },
        { ref: "B1", author: "Bob", text: "Second comment" },
      ],
      sharedStrings: ["Cell A1", "Cell B1"],
      cellData: [
        { ref: "A1", value: "Cell A1", type: "s", ssIndex: 0 },
        { ref: "B1", value: "Cell B1", type: "s", ssIndex: 1 },
      ],
    })

    const workbook = await readXlsx(xlsxData)
    const sheet = workbook.sheets[0]

    const cellA1 = sheet.cells!.get("0,0")
    expect(cellA1!.comment!.text).toBe("First comment")
    expect(cellA1!.comment!.author).toBe("Alice")

    const cellB1 = sheet.cells!.get("0,1")
    expect(cellB1!.comment!.text).toBe("Second comment")
    expect(cellB1!.comment!.author).toBe("Bob")
  })

  it("reads comment on cell with no prior cell detail (creates Cell entry)", async () => {
    const xlsxData = await createXlsxWithComments({
      comments: [{ ref: "A1", author: "User", text: "Comment on number" }],
      cellData: [{ ref: "A1", value: "42" }],
    })

    const workbook = await readXlsx(xlsxData)
    const sheet = workbook.sheets[0]

    const cell = sheet.cells!.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.comment!.text).toBe("Comment on number")
    expect(cell!.value).toBe(42)
  })
})

// ── Round-trip Tests ─────────────────────────────────────────────────

describe("XLSX comment round-trip", () => {
  it("round-trips single comment", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "This is a comment", author: "Test Author" },
    })

    const written = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]], cells }],
    })

    const workbook = await readXlsx(written)
    const sheet = workbook.sheets[0]

    expect(sheet.cells).toBeDefined()
    const cell = sheet.cells!.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.comment).toBeDefined()
    expect(cell!.comment!.text).toBe("This is a comment")
    expect(cell!.comment!.author).toBe("Test Author")
  })

  it("round-trips comment without author", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Hello",
      comment: { text: "No author comment" },
    })

    const written = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello"]], cells }],
    })

    const workbook = await readXlsx(written)
    const sheet = workbook.sheets[0]

    const cell = sheet.cells!.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.comment).toBeDefined()
    expect(cell!.comment!.text).toBe("No author comment")
  })

  it("round-trips multiple comments on different cells", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "A1",
      comment: { text: "Comment A1", author: "Alice" },
    })
    cells.set("1,0", {
      value: "A2",
      comment: { text: "Comment A2", author: "Bob" },
    })
    cells.set("0,1", {
      value: "B1",
      comment: { text: "Comment B1", author: "Alice" },
    })

    const written = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A1", "B1"], ["A2"]],
          cells,
        },
      ],
    })

    const workbook = await readXlsx(written)
    const sheet = workbook.sheets[0]

    const cellA1 = sheet.cells!.get("0,0")
    expect(cellA1!.comment!.text).toBe("Comment A1")
    expect(cellA1!.comment!.author).toBe("Alice")

    const cellA2 = sheet.cells!.get("1,0")
    expect(cellA2!.comment!.text).toBe("Comment A2")
    expect(cellA2!.comment!.author).toBe("Bob")

    const cellB1 = sheet.cells!.get("0,1")
    expect(cellB1!.comment!.text).toBe("Comment B1")
    expect(cellB1!.comment!.author).toBe("Alice")
  })

  it("round-trips comments on multiple sheets", async () => {
    const cells1 = new Map<string, Partial<Cell>>()
    cells1.set("0,0", {
      value: "Sheet1",
      comment: { text: "Comment on Sheet1" },
    })

    const cells2 = new Map<string, Partial<Cell>>()
    cells2.set("0,0", {
      value: "Sheet2",
      comment: { text: "Comment on Sheet2", author: "User" },
    })

    const written = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["Sheet1"]], cells: cells1 },
        { name: "Sheet2", rows: [["Sheet2"]], cells: cells2 },
      ],
    })

    const workbook = await readXlsx(written)

    const cell1 = workbook.sheets[0].cells!.get("0,0")
    expect(cell1!.comment!.text).toBe("Comment on Sheet1")

    const cell2 = workbook.sheets[1].cells!.get("0,0")
    expect(cell2!.comment!.text).toBe("Comment on Sheet2")
    expect(cell2!.comment!.author).toBe("User")
  })

  it("preserves cell value alongside comment", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Important data",
      comment: { text: "Review this value" },
    })

    const written = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Important data"]], cells }],
    })

    const workbook = await readXlsx(written)
    const sheet = workbook.sheets[0]

    // Value should be preserved
    expect(sheet.rows[0][0]).toBe("Important data")

    // Comment should also be present
    const cell = sheet.cells!.get("0,0")
    expect(cell!.comment!.text).toBe("Review this value")
    expect(cell!.value).toBe("Important data")
  })

  it("round-trips comments alongside hyperlinks", async () => {
    const cells = new Map<string, Partial<Cell>>()
    cells.set("0,0", {
      value: "Link with comment",
      hyperlink: { target: "https://example.com" },
      comment: { text: "This cell has both" },
    })

    const written = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Link with comment"]], cells }],
    })

    const workbook = await readXlsx(written)
    const sheet = workbook.sheets[0]

    const cell = sheet.cells!.get("0,0")
    expect(cell).toBeDefined()
    expect(cell!.hyperlink).toBeDefined()
    expect(cell!.hyperlink!.target).toBe("https://example.com")
    expect(cell!.comment).toBeDefined()
    expect(cell!.comment!.text).toBe("This cell has both")
  })
})
