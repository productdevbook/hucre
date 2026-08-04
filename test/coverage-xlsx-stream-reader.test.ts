import { describe, expect, it } from "vitest"
import { ZipWriter } from "../src/zip/writer"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import type { StreamRow } from "../src/xlsx/stream-reader"
import { ParseError, ZipError } from "../src/errors"
import type { CellValue, ReadOptions } from "../src/_types"

// ── Package assembly helpers ─────────────────────────────────────────

const enc = new TextEncoder()
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

function relsXml(entries: Array<{ id: string; type: string; target: string }>): string {
  const items = entries
    .map((e) => `<Relationship Id="${e.id}" Type="${REL_BASE}/${e.type}" Target="${e.target}"/>`)
    .join("")
  return (
    `<?xml version="1.0"?><Relationships ` +
    `xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`
  )
}

const CONTENT_TYPES =
  `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/></Types>`

const workbookXml = (body: string): string =>
  `<?xml version="1.0"?><workbook ${NS} ${R}>${body}</workbook>`

const worksheetXml = (body: string): string =>
  `<?xml version="1.0"?><worksheet ${NS} ${R}><sheetData>${body}</sheetData></worksheet>`

type Parts = Record<string, string | Uint8Array>

/**
 * Build a ZIP from an ordered part list. Order matters here: the true
 * streaming path walks entries in archive order and can only resolve the
 * target worksheet from metadata it has already passed.
 */
async function build(parts: Parts): Promise<Uint8Array> {
  const zip = new ZipWriter()
  for (const [path, content] of Object.entries(parts)) {
    zip.add(path, typeof content === "string" ? enc.encode(content) : content)
  }
  return zip.build()
}

function defaultParts(sheetBody: string): Parts {
  return {
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": relsXml([{ id: "rId1", type: "officeDocument", target: "xl/workbook.xml" }]),
    "xl/workbook.xml": workbookXml(
      `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>`,
    ),
    "xl/_rels/workbook.xml.rels": relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
    ]),
    "xl/worksheets/sheet1.xml": worksheetXml(sheetBody),
  }
}

function toStream(data: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(data)
      controller.close()
    },
  })
}

async function collect(
  gen: AsyncGenerator<StreamRow, void, undefined>,
): Promise<Array<CellValue[]>> {
  const out: Array<CellValue[]> = []
  for await (const row of gen) out.push(row.values)
  return out
}

/** Stream the default package's rows from a buffer. */
async function rowsOf(sheetBody: string, options?: ReadOptions & { sheet?: number | string }) {
  return collect(streamXlsxRows(await build(defaultParts(sheetBody)), options))
}

// ═══════════════════════════════════════════════════════════════════════
// Cell value resolution
//
// The streaming reader resolves values without building Cell objects, so
// it has its own copy of the type switch. It must agree with the batch
// reader on every `t` the spec defines (§18.18.11, ST_CellType).
// ═══════════════════════════════════════════════════════════════════════

describe("cell types", () => {
  it("resolves every ST_CellType the same way the batch reader does", async () => {
    const rows = await rowsOf(
      `<row r="1">` +
        `<c r="A1" t="str"><f>UPPER(B9)</f><v>line1_x000A_line2</v></c>` +
        `<c r="B1" t="b"><v>1</v></c>` +
        `<c r="C1" t="b"><v>TRUE</v></c>` +
        `<c r="D1" t="b"><v>0</v></c>` +
        `<c r="E1" t="e"><v>#REF!</v></c>` +
        `<c r="F1" t="n"><v>12.5</v></c>` +
        `<c r="G1"><v>7</v></c>` +
        `</row>`,
    )
    expect(rows[0]).toEqual(["line1\nline2", true, true, false, "#REF!", 12.5, 7])
  })

  it("returns null for an out-of-range shared string index", async () => {
    // No sharedStrings part at all, so every `t="s"` index is out of
    // range — the raw index must not leak through as text.
    const rows = await rowsOf(`<row r="1"><c r="A1" t="s"><v>3</v></c></row>`)
    expect(rows[0]).toEqual([null])
  })

  it("keeps unparseable numeric text as a string and empty text as null", async () => {
    const rows = await rowsOf(
      `<row r="1"><c r="A1"><v>N/A</v></c><c r="B1"><v></v></c><c r="C1"/></row>`,
    )
    expect(rows[0]).toEqual(["N/A", null, null])
  })

  it("joins the runs of an inline rich-text string", async () => {
    const rows = await rowsOf(
      `<row r="1">` +
        `<c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>Bo</t></r><r><t>ld</t></r></is></c>` +
        `<c r="B1" t="inlineStr"><is><t>plain_x000A_text</t></is></c>` +
        `</row>`,
    )
    expect(rows[0]).toEqual(["Bold", "plain\ntext"])
  })

  it("resolves shared strings declared in sharedStrings.xml", async () => {
    const parts = defaultParts(
      `<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>`,
    )
    parts["xl/_rels/workbook.xml.rels"] = relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
      { id: "rId2", type: "sharedStrings", target: "sharedStrings.xml" },
    ])
    parts["xl/sharedStrings.xml"] =
      `<?xml version="1.0"?><sst ${NS} count="2" uniqueCount="2">` +
      `<si><t>zero</t></si><si><t>one</t></si></sst>`
    expect(await collect(streamXlsxRows(await build(parts)))).toEqual([["one", "zero"]])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Row and column positioning
// ═══════════════════════════════════════════════════════════════════════

describe("implicit positioning", () => {
  it("numbers rows and cells by document order when r is absent", async () => {
    const rows = await collect(
      streamXlsxRows(
        await build(
          defaultParts(
            `<row><c><v>1</v></c><c><v>2</v></c></row>` +
              `<row><c r="C2"><v>3</v></c><c><v>4</v></c></row>`,
          ),
        ),
      ),
    )
    expect(rows).toEqual([
      [1, 2],
      [null, null, 3, 4],
    ])
  })

  it("places cells written out of column order at their stated column", async () => {
    // Nothing requires `<c>` elements to be sorted; the row width comes
    // from the largest column seen, not the last one.
    const rows = await rowsOf(`<row r="1"><c r="C1"><v>3</v></c><c r="A1"><v>1</v></c></row>`)
    expect(rows).toEqual([[1, null, 3]])
  })

  it("yields an empty value array for a row with no cells", async () => {
    const rows = await rowsOf(`<row r="1"/><row r="2"><c r="A2"><v>1</v></c></row>`)
    expect(rows).toEqual([[], [1]])
  })

  it("reads a worksheet written with a namespace prefix on every tag", async () => {
    // Files from some Java toolchains prefix the main namespace
    // (`<x:sheetData>`) instead of defaulting it.
    const parts = defaultParts("")
    parts["xl/worksheets/sheet1.xml"] =
      `<?xml version="1.0"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>pre</x:t></x:is></x:c>` +
      `<x:c r="B1"><x:v>2</x:v></x:c></x:row></x:sheetData></x:worksheet>`
    expect(await collect(streamXlsxRows(await build(parts)))).toEqual([["pre", 2]])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// range and maxRows
// ═══════════════════════════════════════════════════════════════════════

const GRID = Array.from(
  { length: 6 },
  (_, r) =>
    `<row r="${r + 1}">` +
    Array.from(
      { length: 4 },
      (_, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${r * 10 + c}</v></c>`,
    ).join("") +
    `</row>`,
).join("")

describe("range read option", () => {
  it("masks columns outside the window and keeps values at their own index", async () => {
    const rows = await rowsOf(GRID, { range: "B2:C4" })
    expect(rows).toEqual([
      [null, 11, 12, null],
      [null, 21, 22, null],
      [null, 31, 32, null],
    ])
  })

  it("stops reading once a row past the end of the range is seen", async () => {
    // Worksheet rows are written in ascending order, so the first row
    // past the window means no further row can qualify.
    const rows = await rowsOf(GRID, { range: "A1:B2" })
    expect(rows).toHaveLength(2)
  })

  it("pads the row out to the end column even when the sheet is narrower", async () => {
    const rows = await rowsOf(`<row r="1"><c r="A1"><v>1</v></c></row>`, { range: "A1:D1" })
    expect(rows[0]).toEqual([1, null, null, null])
  })

  it("accepts a single-cell range", async () => {
    const rows = await rowsOf(GRID, { range: "C3" })
    expect(rows).toEqual([[null, null, 22, null]])
  })

  it("normalises a reversed range", async () => {
    // `C3:A1` describes the same rectangle as `A1:C3`.
    const rows = await rowsOf(GRID, { range: "C3:A1" })
    expect(rows).toEqual([
      [0, 1, 2, null],
      [10, 11, 12, null],
      [20, 21, 22, null],
    ])
  })

  it("rejects a reference with more than two endpoints", async () => {
    await expect(
      collect(streamXlsxRows(await build(defaultParts(GRID)), { range: "A1:B2:C3" })),
    ).rejects.toThrow(ParseError)
  })
})

describe("maxRows read option", () => {
  it("stops after the requested number of rows", async () => {
    expect(await rowsOf(GRID, { maxRows: 2 })).toHaveLength(2)
  })

  it("treats zero and negative caps as unlimited", async () => {
    expect(await rowsOf(GRID, { maxRows: 0 })).toHaveLength(6)
    expect(await rowsOf(GRID, { maxRows: -5 })).toHaveLength(6)
  })

  it("counts only rows that survive the range filter", async () => {
    const rows = await rowsOf(GRID, { range: "A3:D6", maxRows: 2 })
    expect(rows).toHaveLength(2)
    expect(rows[0][0]).toBe(20)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Sheet selection
// ═══════════════════════════════════════════════════════════════════════

describe("sheet selection", () => {
  const twoSheets = (): Parts => {
    const parts = defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`)
    parts["xl/workbook.xml"] = workbookXml(
      `<sheets><sheet name="First" sheetId="1" r:id="rId1"/>` +
        `<sheet name="Second" sheetId="2" r:id="rId2"/></sheets>`,
    )
    parts["xl/_rels/workbook.xml.rels"] = relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
      { id: "rId2", type: "worksheet", target: "worksheets/sheet2.xml" },
    ])
    parts["xl/worksheets/sheet2.xml"] = worksheetXml(`<row r="1"><c r="A1"><v>2</v></c></row>`)
    return parts
  }

  it("selects by index and by name", async () => {
    const parts = twoSheets()
    expect(await collect(streamXlsxRows(await build(parts), { sheet: 1 }))).toEqual([[2]])
    expect(await collect(streamXlsxRows(await build(parts), { sheet: "Second" }))).toEqual([[2]])
  })

  it("yields nothing for an index past the end of the sheet list", async () => {
    expect(await collect(streamXlsxRows(await build(twoSheets()), { sheet: 9 }))).toEqual([])
    expect(await collect(streamXlsxRows(await build(twoSheets()), { sheet: -1 }))).toEqual([])
  })

  it("yields nothing when the workbook declares no sheets at all", async () => {
    const parts = defaultParts("")
    parts["xl/workbook.xml"] = workbookXml(`<sheets/>`)
    expect(await collect(streamXlsxRows(await build(parts)))).toEqual([])
  })

  it("finds the relationship id under any namespace prefix", async () => {
    const parts = defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`)
    parts["xl/workbook.xml"] =
      `<?xml version="1.0"?><workbook ${NS} xmlns:rel="${REL_BASE}">` +
      `<sheets><sheet name="S" sheetId="1" rel:id="rId1"/><sheet name="NoId" sheetId="2"/>` +
      `</sheets></workbook>`
    expect(await collect(streamXlsxRows(await build(parts)))).toEqual([[1]])
  })

  it("ignores entries in <sheets> that are not usable sheet declarations", async () => {
    // A sheet with no name, a non-sheet child, and a sheet with no
    // sheetId all appear in files stitched together by other tools.
    const parts = defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`)
    parts["xl/workbook.xml"] = workbookXml(
      `<sheets><sheet sheetId="1" r:id="rId1"/><sheetGroup name="x"/>` +
        `<sheet name="Real" r:id="rId1"/></sheets>`,
    )
    expect(await collect(streamXlsxRows(await build(parts)))).toEqual([[1]])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Date system
// ═══════════════════════════════════════════════════════════════════════

describe("date system", () => {
  const dated = (workbookPr: string): Parts => {
    const parts = defaultParts(`<row r="1"><c r="A1" s="1"><v>40000</v></c></row>`)
    parts["xl/workbook.xml"] = workbookXml(
      `${workbookPr}<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`,
    )
    parts["xl/_rels/workbook.xml.rels"] = relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
      { id: "rId2", type: "styles", target: "styles.xml" },
    ])
    parts["xl/styles.xml"] =
      `<?xml version="1.0"?><styleSheet ${NS}>` +
      `<fonts count="1"><font/></fonts><fills count="1"><fill/></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>` +
      `</styleSheet>`
    return parts
  }

  const serialOf = async (parts: Parts, options?: ReadOptions): Promise<Date> => {
    const rows = await collect(streamXlsxRows(await build(parts), options))
    return rows[0][0] as Date
  }

  it("auto-detects the 1904 epoch from workbookPr", async () => {
    const auto = await serialOf(dated(`<workbookPr date1904="true"/>`))
    const explicit = await serialOf(dated(`<workbookPr date1904="1"/>`), { dateSystem: "auto" })
    expect(auto).toBeInstanceOf(Date)
    expect(auto.getTime()).toBe(explicit.getTime())
  })

  it("lets an explicit dateSystem option override the file's flag", async () => {
    const forced1900 = await serialOf(dated(`<workbookPr date1904="1"/>`), {
      dateSystem: "1900",
    })
    const forced1904 = await serialOf(dated(``), { dateSystem: "1904" })
    expect(forced1904.getTime()).toBeGreaterThan(forced1900.getTime())
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Package layout and structural failures
// ═══════════════════════════════════════════════════════════════════════

describe("package layout", () => {
  it("reads a workbook addressed by an absolute target at the package root", async () => {
    const parts: Parts = {
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": relsXml([{ id: "rId1", type: "officeDocument", target: "/workbook.xml" }]),
      "workbook.xml": workbookXml(`<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`),
      "_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "../sheets/./sheet1.xml" },
      ]),
      "sheets/sheet1.xml": worksheetXml(`<row r="1"><c r="A1"><v>42</v></c></row>`),
    }
    expect(await collect(streamXlsxRows(await build(parts)))).toEqual([[42]])
  })

  it("streams the same root-layout package straight from a stream", async () => {
    // The stream resolver has its own copy of the path helpers, so the
    // root layout has to be exercised on that side too.
    const parts: Parts = {
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": relsXml([{ id: "rId1", type: "officeDocument", target: "/workbook.xml" }]),
      "workbook.xml": workbookXml(`<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`),
      "_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "/xl/worksheets/sheet1.xml" },
      ]),
      "xl/worksheets/sheet1.xml": worksheetXml(`<row r="1"><c r="A1"><v>42</v></c></row>`),
    }
    expect(await collect(streamXlsxRows(toStream(await build(parts))))).toEqual([[42]])
  })
})

describe("invalid packages", () => {
  const withoutPart = async (path: string): Promise<Uint8Array> => {
    const parts = defaultParts("")
    delete parts[path]
    return build(parts)
  }

  it("rejects input that is not a ZIP archive at all", async () => {
    const notAZip = enc.encode("This is a CSV,not,an,xlsx\n")
    await expect(collect(streamXlsxRows(notAZip))).rejects.toThrow(ZipError)
  })

  it("rejects a package with no [Content_Types].xml", async () => {
    await expect(collect(streamXlsxRows(await withoutPart("[Content_Types].xml")))).rejects.toThrow(
      /missing \[Content_Types\]\.xml/,
    )
  })

  it("rejects a package with no root relationships", async () => {
    await expect(collect(streamXlsxRows(await withoutPart("_rels/.rels")))).rejects.toThrow(
      /missing _rels\/\.rels/,
    )
  })

  it("rejects root rels that name no officeDocument", async () => {
    const parts = defaultParts("")
    parts["_rels/.rels"] = relsXml([
      { id: "rId1", type: "extended-properties", target: "docProps/app.xml" },
    ])
    await expect(collect(streamXlsxRows(await build(parts)))).rejects.toThrow(
      /cannot find workbook relationship/,
    )
  })

  it("rejects a package whose workbook part is missing", async () => {
    await expect(collect(streamXlsxRows(await withoutPart("xl/workbook.xml")))).rejects.toThrow(
      /missing workbook at xl\/workbook\.xml/,
    )
  })

  it("rejects a sheet whose worksheet part is missing", async () => {
    await expect(
      collect(streamXlsxRows(await withoutPart("xl/worksheets/sheet1.xml"))),
    ).rejects.toThrow(/missing worksheet file for sheet "Sheet1"/)
  })

  it("rejects a sheet with no worksheet relationship to resolve", async () => {
    // Dropping workbook.xml.rels leaves the declared sheet unreachable.
    await expect(
      collect(streamXlsxRows(await withoutPart("xl/_rels/workbook.xml.rels"))),
    ).rejects.toThrow(ParseError)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// True streaming path (ReadableStream input)
//
// `prepareStreaming` walks ZIP local headers in archive order. It can
// only stream the worksheet if every part it needs was already passed,
// so several realistic archive layouts must fall back to buffering —
// silently, and with identical results.
// ═══════════════════════════════════════════════════════════════════════

describe("ReadableStream input", () => {
  it("streams the target worksheet when metadata comes first", async () => {
    const rows = await collect(streamXlsxRows(toStream(await build(defaultParts(GRID)))))
    expect(rows).toHaveLength(6)
    expect(rows[5]).toEqual([50, 51, 52, 53])
  })

  it("applies range and maxRows while streaming", async () => {
    expect(
      await collect(
        streamXlsxRows(toStream(await build(defaultParts(GRID))), {
          range: "B2:C3",
        }),
      ),
    ).toEqual([
      [null, 11, 12, null],
      [null, 21, 22, null],
    ])
    expect(
      await collect(streamXlsxRows(toStream(await build(defaultParts(GRID))), { maxRows: 3 })),
    ).toHaveLength(3)
  })

  it("skips non-target worksheets it passes on the way", async () => {
    const parts = defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`)
    parts["xl/workbook.xml"] = workbookXml(
      `<sheets><sheet name="First" sheetId="1" r:id="rId1"/>` +
        `<sheet name="Second" sheetId="2" r:id="rId2"/></sheets>`,
    )
    parts["xl/_rels/workbook.xml.rels"] = relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
      { id: "rId2", type: "worksheet", target: "worksheets/sheet2.xml" },
    ])
    parts["xl/worksheets/sheet2.xml"] = worksheetXml(`<row r="1"><c r="A1"><v>2</v></c></row>`)
    expect(await collect(streamXlsxRows(toStream(await build(parts)), { sheet: 1 }))).toEqual([[2]])
  })

  it("falls back to buffering when the worksheet precedes the workbook metadata", async () => {
    // Some writers emit sheets first. The resolver has nothing to match
    // against at that point, so it buffers and re-reads by index.
    const src = defaultParts(`<row r="1"><c r="A1"><v>9</v></c></row>`)
    const reordered: Parts = { "xl/worksheets/sheet1.xml": src["xl/worksheets/sheet1.xml"] }
    for (const [k, v] of Object.entries(src)) {
      if (k !== "xl/worksheets/sheet1.xml") reordered[k] = v
    }
    expect(await collect(streamXlsxRows(toStream(await build(reordered))))).toEqual([[9]])
  })

  it("falls back when sharedStrings.xml is stored after the worksheet", async () => {
    // String cells cannot be resolved without the table, so streaming
    // must not start until it has been seen.
    const parts = defaultParts(`<row r="1"><c r="A1" t="s"><v>0</v></c></row>`)
    parts["xl/_rels/workbook.xml.rels"] = relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
      { id: "rId2", type: "sharedStrings", target: "sharedStrings.xml" },
    ])
    // Added last, so it lands after xl/worksheets/sheet1.xml.
    parts["xl/sharedStrings.xml"] =
      `<?xml version="1.0"?><sst ${NS} count="1" uniqueCount="1"><si><t>late</t></si></sst>`
    expect(await collect(streamXlsxRows(toStream(await build(parts))))).toEqual([["late"]])
  })

  it("falls back when styles.xml is stored after the worksheet", async () => {
    const parts = defaultParts(`<row r="1"><c r="A1"><v>5</v></c></row>`)
    parts["xl/_rels/workbook.xml.rels"] = relsXml([
      { id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" },
      { id: "rId2", type: "styles", target: "styles.xml" },
    ])
    parts["xl/styles.xml"] =
      `<?xml version="1.0"?><styleSheet ${NS}><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>`
    expect(await collect(streamXlsxRows(toStream(await build(parts))))).toEqual([[5]])
  })

  it("falls back when the worksheet is not stored under xl/worksheets/", async () => {
    // The stream resolver only recognises the canonical worksheet path;
    // anything else is resolved by the random-access reader.
    const parts: Parts = {
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": relsXml([{ id: "rId1", type: "officeDocument", target: "xl/workbook.xml" }]),
      "xl/workbook.xml": workbookXml(`<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`),
      "xl/_rels/workbook.xml.rels": relsXml([
        { id: "rId1", type: "worksheet", target: "sheets/sheet1.xml" },
      ]),
      "xl/sheets/sheet1.xml": worksheetXml(`<row r="1"><c r="A1"><v>3</v></c></row>`),
    }
    expect(await collect(streamXlsxRows(toStream(await build(parts))))).toEqual([[3]])
  })

  it("falls back and then yields nothing when the named sheet does not exist", async () => {
    const data = await build(defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`))
    expect(await collect(streamXlsxRows(toStream(data), { sheet: "Nope" }))).toEqual([])
  })

  it("falls back when the entries cannot be streamed from their local headers", async () => {
    // Bit 3 of the general-purpose flag says the sizes live in a trailing
    // data descriptor, so the local header alone cannot bound the entry.
    // The central directory is untouched, so the buffered reader copes.
    const data = await build(defaultParts(`<row r="1"><c r="A1"><v>8</v></c></row>`))
    const patched = withDataDescriptorFlag(data)
    expect(await collect(streamXlsxRows(toStream(patched)))).toEqual([[8]])
  })

  it("hands a broken package back to the buffered path for its error", async () => {
    // The stream resolver bails out silently on anything it cannot
    // resolve; the diagnosis is the buffered reader's job, so the same
    // messages must come out either way.
    const cases: Array<[string, RegExp]> = [
      ["_rels/.rels", /missing _rels\/\.rels/],
      ["xl/workbook.xml", /missing workbook at xl\/workbook\.xml/],
      ["xl/_rels/workbook.xml.rels", /missing worksheet file/],
    ]
    for (const [drop, message] of cases) {
      const parts = defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`)
      delete parts[drop]
      await expect(collect(streamXlsxRows(toStream(await build(parts))))).rejects.toThrow(message)
    }

    const noOfficeDoc = defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`)
    noOfficeDoc["_rels/.rels"] = relsXml([
      { id: "rId1", type: "extended-properties", target: "docProps/app.xml" },
    ])
    await expect(collect(streamXlsxRows(toStream(await build(noOfficeDoc))))).rejects.toThrow(
      /cannot find workbook relationship/,
    )

    const danglingSheet = defaultParts(`<row r="1"><c r="A1"><v>1</v></c></row>`)
    danglingSheet["xl/workbook.xml"] = workbookXml(
      `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId404"/></sheets>`,
    )
    await expect(collect(streamXlsxRows(toStream(await build(danglingSheet))))).rejects.toThrow(
      /missing worksheet file for sheet "Sheet1"/,
    )
  })

  it("propagates a row-handler failure instead of hanging on the stream", async () => {
    // The bound check in the row builder throws from inside the SAX
    // handler. A rejected parse used to be swallowed here, so the read
    // looked short-but-successful. See #363.
    const parts = defaultParts(`<row r="1"><c r="AAAAAA1"><v>1</v></c></row>`)
    await expect(collect(streamXlsxRows(toStream(await build(parts))))).rejects.toThrow(
      /outside the supported sheet bounds/,
    )
  })

  it("releases the stream when the consumer abandons the generator early", async () => {
    const gen = streamXlsxRows(toStream(await build(defaultParts(GRID))))
    const first = await gen.next()
    expect(first.value!.values).toEqual([0, 1, 2, 3])
    await gen.return()
    // A second pull after return() must simply report completion.
    expect(await gen.next()).toEqual({ done: true, value: undefined })
  })
})

/**
 * Set the "sizes follow in a data descriptor" flag on every *local*
 * header, leaving the central directory intact. Marks the archive
 * un-streamable without making it unreadable.
 */
function withDataDescriptorFlag(zip: Uint8Array): Uint8Array {
  const out = zip.slice()
  for (let i = 0; i + 3 < out.length; i++) {
    if (out[i] === 0x50 && out[i + 1] === 0x4b && out[i + 2] === 0x03 && out[i + 3] === 0x04) {
      out[i + 6] |= 0x08
    }
  }
  return out
}
