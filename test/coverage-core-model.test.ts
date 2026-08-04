import { describe, expect, it } from "vitest"
import type { Cell, CellValue, SchemaFieldType, Sheet, SheetTextBox, Workbook } from "../src/_types"
import { audit } from "../src/a11y"
import { selectSheet } from "../src/_objects"
import { validateWithSchema } from "../src/_schema"
import { WorkbookBuilder } from "../src/builder"
import { letterToCol } from "../src/cell-utils"
import { read, writeObjects } from "../src/defter"
import { ParseError, UnsupportedFormatError, ValidationError } from "../src/errors"
import { imageFromBase64 } from "../src/image"
import { sheetToArrays } from "../src/sheet-utils"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"

// ── Helpers ──────────────────────────────────────────────────────────

function sheet(overrides: Partial<Sheet> = {}): Sheet {
  return { name: "Sheet1", rows: [], ...overrides }
}

function workbook(...sheets: Sheet[]): Workbook {
  return { sheets, properties: { title: "T", description: "D" } }
}

/** A styled cell override, the shape the contrast audit walks. */
function styled(value: CellValue, fontRgb?: string, fillRgb?: string): Cell {
  return {
    value,
    type: typeof value === "number" ? "number" : "string",
    style: {
      font: fontRgb ? { color: { rgb: fontRgb } } : {},
      fill: { type: "pattern", pattern: "solid", fgColor: fillRgb ? { rgb: fillRgb } : undefined },
    },
  }
}

const decoder = new TextDecoder("utf-8")

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  return decoder.decode(await new ZipReader(data).extract(path))
}

// ═══════════════════════════════════════════════════════════════════════
// _schema — errorMode: "throw"
// ═══════════════════════════════════════════════════════════════════════

// Every validation stage collects into `errors` by default and throws a
// `ValidationError` carrying that one issue under `errorMode: "throw"`.
// Only the required-field stage had a test for the throwing path.
describe('validateWithSchema — errorMode "throw"', () => {
  const rows: CellValue[][] = [["value"], ["x"]]

  function expectThrow(field: Parameters<typeof validateWithSchema>[1]["f"], message: RegExp) {
    let caught: unknown
    try {
      validateWithSchema(rows, { f: { column: "value", ...field } }, { errorMode: "throw" })
    } catch (e) {
      caught = e
    }
    expect(caught).toBeInstanceOf(ValidationError)
    expect((caught as ValidationError).message).toMatch(message)
    expect((caught as ValidationError).errors).toHaveLength(1)
  }

  it("throws on a type coercion failure", () => {
    expectThrow({ type: "number" }, /Expected number/)
  })

  it("throws on a pattern mismatch", () => {
    expectThrow({ type: "string", pattern: /^\d+$/ }, /does not match pattern/)
  })

  it("throws on a min/max violation", () => {
    expectThrow({ type: "string", min: 5 }, /below minimum 5/)
  })

  it("throws on a value outside the enum", () => {
    expectThrow({ type: "string", enum: ["a", "b"] }, /must be one of: a, b/)
  })

  it("throws on a custom validate() rejection", () => {
    expectThrow({ validate: () => "nope" }, /nope/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// _schema — coercion edge cases
// ═══════════════════════════════════════════════════════════════════════

describe("validateWithSchema — coercion edges", () => {
  // A cell holding only a thousands separator is non-empty (so the
  // required/empty check lets it through) but strips to nothing.
  it("rejects a value that is nothing but thousands separators", () => {
    const num = validateWithSchema([["n"], [","]], { n: { type: "number" } })
    expect(num.errors[0]!.message).toBe("Expected number for 'n', got ''")
    expect(num.data[0]!.n).toBeNull()

    const int = validateWithSchema([["n"], [",,"]], { n: { type: "integer" } })
    expect(int.errors[0]!.message).toBe("Expected integer for 'n', got ''")
  })

  // A completely empty row (`[]`, which is what a trailing `\n` or a blank
  // `<row/>` yields) counts as empty for `skipEmptyRows`.
  it("skips zero-length rows when skipEmptyRows is set", () => {
    const rows: CellValue[][] = [["n"], [1], [], [2]]
    const kept = validateWithSchema(rows, { n: { type: "number" } }, { skipEmptyRows: true })
    expect(kept.data.map((d) => d.n)).toEqual([1, 2])

    const all = validateWithSchema(rows, { n: { type: "number" } })
    expect(all.data.map((d) => d.n)).toEqual([1, null, 2])
  })

  // JS callers are not held to the `SchemaFieldType` union; an unrecognised
  // type name passes the raw cell value through instead of throwing.
  it("passes the value through for an unrecognised field type", () => {
    const result = validateWithSchema([["v"], ["kept"]], {
      v: { type: "uuid" as SchemaFieldType },
    })
    expect(result.errors).toEqual([])
    expect(result.data[0]!.v).toBe("kept")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// _objects — selectSheet
// ═══════════════════════════════════════════════════════════════════════

describe("selectSheet", () => {
  it("reports an empty workbook before blaming the selector", () => {
    expect(() => selectSheet({ sheets: [] }, 0)).toThrow(ParseError)
    expect(() => selectSheet({ sheets: [] }, "Data")).toThrow(/no sheets/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// a11y.audit
// ═══════════════════════════════════════════════════════════════════════

describe("a11y.audit — sheet emptiness", () => {
  // A sheet can carry all of its content in the per-cell override Map with
  // an empty `rows` array (that is what the streaming writers accept), and
  // it is not an empty sheet.
  it("does not call a sheet empty when its content lives in the cells Map", () => {
    const wb = workbook(
      sheet({ rows: [], cells: new Map([["0,0", { value: "Header", type: "string" } as Cell]]) }),
    )

    const codes = audit(wb, { skipContrast: true }).map((i) => i.code)

    expect(codes).not.toContain("empty-sheet")
    expect(codes).toContain("no-header-row")
  })

  it("does not report blank rows when every row is blank", () => {
    const wb = workbook(
      sheet({
        rows: [
          [null, null],
          ["", ""],
        ],
        cells: new Map([["0,0", { value: "x", type: "string" } as Cell]]),
        a11y: { headerRow: 0 },
      }),
    )

    expect(audit(wb, { skipContrast: true }).filter((i) => i.code === "blank-row-in-data")).toEqual(
      [],
    )
  })

  // `rows[5] = [...]` on a fresh array leaves holes; a hole is a blank row,
  // not a crash.
  it("treats a hole in the rows array as a blank row", () => {
    const rows: CellValue[][] = []
    rows[0] = ["a"]
    rows[2] = ["b"]
    const wb = workbook(sheet({ rows, a11y: { headerRow: 0 } }))

    const issue = audit(wb, { skipContrast: true }).find((i) => i.code === "blank-row-in-data")

    expect(issue?.location?.ref).toBe("2:2")
  })
})

describe("a11y.audit — text boxes", () => {
  const anchor = { from: { row: 1, col: 1 } }

  it("warns about a text box with no alt text", () => {
    const textBoxes: SheetTextBox[] = [
      { text: "Q3 target", anchor },
      { text: "blank alt", anchor, altText: "   " },
    ]
    const wb = workbook(sheet({ rows: [["a"]], a11y: { headerRow: 0 }, textBoxes }))

    const issues = audit(wb, { skipContrast: true }).filter((i) => i.code === "missing-alt-text")

    expect(issues).toHaveLength(2)
    // A text box is advisory (warning); a missing image alt text is an error.
    expect(issues[0]!.type).toBe("warning")
    expect(issues[0]!.location).toEqual({ sheet: "Sheet1", ref: "B2", textBox: 0 })
    expect(issues[1]!.location?.textBox).toBe(1)
  })

  it("stays quiet when every text box has alt text", () => {
    const textBoxes: SheetTextBox[] = [{ text: "Q3 target", anchor, altText: "Q3 target callout" }]
    const wb = workbook(sheet({ rows: [["a"]], a11y: { headerRow: 0 }, textBoxes }))

    expect(audit(wb, { skipContrast: true }).filter((i) => i.code === "missing-alt-text")).toEqual(
      [],
    )
  })
})

describe("a11y.audit — contrast sampling", () => {
  function lowContrastSheet(count: number): Sheet {
    const cells = new Map<string, Cell>()
    for (let i = 0; i < count; i++) cells.set(`${i},0`, styled("text", "FFCCCCCC", "FFFFFFFF"))
    return sheet({ rows: [["text"]], a11y: { headerRow: 0 }, cells })
  }

  it("stops inspecting once contrastSampleLimit cells have been walked", () => {
    const wb = workbook(lowContrastSheet(5))

    const limited = audit(wb, { contrastSampleLimit: 2 }).filter((i) => i.code === "low-contrast")
    const full = audit(wb).filter((i) => i.code === "low-contrast")

    expect(limited).toHaveLength(2)
    expect(full).toHaveLength(5)
  })

  it("ignores cells with no user-visible text", () => {
    const cells = new Map<string, Cell>([
      ["0,0", styled("", "FFCCCCCC", "FFFFFFFF")],
      ["1,0", styled(null, "FFCCCCCC", "FFFFFFFF")],
    ])
    const wb = workbook(sheet({ rows: [[""]], a11y: { headerRow: 0 }, cells }))

    expect(audit(wb).filter((i) => i.code === "low-contrast")).toEqual([])
  })

  // Gradient fills have no single background colour to measure against, and
  // an unfilled cell inherits the (unknown) theme background.
  it("skips cells without a resolvable pattern background", () => {
    const cells = new Map<string, Cell>([
      ["0,0", { value: "x", type: "string", style: { font: { color: { rgb: "FFCCCCCC" } } } }],
      [
        "1,0",
        {
          value: "y",
          type: "string",
          style: {
            font: { color: { rgb: "FFCCCCCC" } },
            fill: { type: "gradient", stops: [{ position: 0, color: { rgb: "FFFFFFFF" } }] },
          },
        },
      ],
    ])
    const wb = workbook(sheet({ rows: [["x"]], a11y: { headerRow: 0 }, cells }))

    expect(audit(wb).filter((i) => i.code === "low-contrast")).toEqual([])
  })

  // Theme-driven fonts have no `rgb`, and indexed/named colours are not
  // hex — neither can be measured, so the check backs off silently.
  it("skips cells whose font or fill colour cannot be resolved to a hex triple", () => {
    const cells = new Map<string, Cell>([
      ["0,0", styled("no font colour", undefined, "FFFFFFFF")],
      ["1,0", styled("no fill colour", "FFCCCCCC", undefined)],
      ["2,0", styled("not hex", "theme:accent1", "FFFFFFFF")],
    ])
    const wb = workbook(sheet({ rows: [["x"]], a11y: { headerRow: 0 }, cells }))

    expect(audit(wb).filter((i) => i.code === "low-contrast")).toEqual([])
  })

  // ODS and hand-authored XLSX write plain 6-digit RRGGBB where Excel
  // writes 8-digit AARRGGBB; both have to measure the same.
  it("accepts 6-digit RRGGBB as well as 8-digit AARRGGBB", () => {
    const cells = new Map<string, Cell>([["0,0", styled("text", "CCCCCC", "FFFFFF")]])
    const wb = workbook(sheet({ rows: [["text"]], a11y: { headerRow: 0 }, cells }))

    const issue = audit(wb).find((i) => i.code === "low-contrast")

    expect(issue?.location?.ref).toBe("A1")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// defter — format detection
// ═══════════════════════════════════════════════════════════════════════

describe("read() — malformed ZIP envelopes", () => {
  it("rejects a file that is only a ZIP signature", async () => {
    const data = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0])

    await expect(read(data)).rejects.toThrow(UnsupportedFormatError)
    await expect(read(data)).rejects.toThrow(/ZIP too short/)
  })

  it("rejects a local file header whose name runs past the end of the file", async () => {
    const data = new Uint8Array(34)
    data.set([0x50, 0x4b, 0x03, 0x04], 0)
    data[26] = 200 // filename length far beyond the 34 bytes we have
    data[27] = 0

    await expect(read(data)).rejects.toThrow(/ZIP truncated/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// defter — writeObjects
// ═══════════════════════════════════════════════════════════════════════

describe("writeObjects()", () => {
  // Columns come from the first object's keys; later objects that are
  // missing one of them get a blank rather than an `undefined` cell.
  it("writes a blank for keys absent from a later object", async () => {
    const output = await writeObjects([
      { Name: "Alice", Age: 30 },
      { Name: "Bob" } as Record<string, CellValue>,
    ])

    const rows = (await readXlsx(output)).sheets[0]!.rows
    expect(rows[2]).toEqual(["Bob", null])
  })

  it("wraps the rows in a native table sized to the data", async () => {
    const output = await writeObjects(
      [
        { Region: "EMEA", Revenue: 10 },
        { Region: "APAC", Revenue: 20 },
      ],
      { table: { name: "Sales", style: "TableStyleMedium2", showRowStripes: false } },
    )

    const table = (await readXlsx(output)).sheets[0]!.tables![0]!
    // 2 data rows + header, two columns → A1:B3.
    expect(table.range).toBe("A1:B3")
    expect(table.columns.map((c) => c.name)).toEqual(["Region", "Revenue"])
    expect(table.columns.every((c) => c.totalFunction === undefined)).toBe(true)
  })

  // A totals row occupies one extra row beyond the data, and only the
  // columns named in `totals` get a function.
  it("extends the table range by one row and tags columns when totals are requested", async () => {
    const output = await writeObjects(
      [
        { Region: "EMEA", Revenue: 10 },
        { Region: "APAC", Revenue: 20 },
      ],
      {
        table: {
          name: "Sales",
          showTotalRow: true,
          showAutoFilter: false,
          totals: { Revenue: "sum" },
        },
      },
    )

    const xml = await extractXml(output, "xl/tables/table1.xml")
    expect(xml).toContain('ref="A1:B4"')
    expect(xml).toContain('totalsRowCount="1"')
    expect(xml).toContain('totalsRowFunction="sum"')
  })

  // `colToLetterSimple` has to roll over past Z for wide object shapes.
  it("computes the table range past column Z", async () => {
    const wide: Record<string, CellValue> = {}
    for (let i = 0; i < 28; i++) wide[`c${i}`] = i
    const output = await writeObjects([wide], { table: { name: "Wide" } })

    expect((await readXlsx(output)).sheets[0]!.tables![0]!.range).toBe("A1:AB2")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// builder
// ═══════════════════════════════════════════════════════════════════════

describe("WorkbookBuilder", () => {
  // The override Map is created lazily on the first cell() call; the
  // second call has to reuse it rather than replace it.
  it("accumulates repeated cell() overrides into one map", async () => {
    const data = await WorkbookBuilder.create()
      .addSheet("Sales")
      .row(["Widget", 100])
      .cell(0, 0, { style: { font: { bold: true } } })
      .cell(0, 1, { formula: "1+1" })
      .build()

    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml")
    const styles = await extractXml(data, "xl/styles.xml")

    expect(sheetXml).toContain("<f>1+1</f>")
    expect(styles).toContain("<b/>")
    // A1 carries a non-default style index, so both overrides survived.
    expect(sheetXml).toMatch(/<c r="A1"[^>]*\ss="[1-9]/)
  })

  // `hidden()` / `veryHidden()` default to true so the common call is
  // argument-free; passing false is the explicit un-hide.
  it("defaults hidden() and veryHidden() to true", async () => {
    const data = await WorkbookBuilder.create()
      .addSheet("Visible")
      .row(["a"])
      .done()
      .addSheet("Hidden")
      .row(["b"])
      .hidden()
      .done()
      .addSheet("Gone")
      .row(["c"])
      .veryHidden()
      .build()

    const xml = await extractXml(data, "xl/workbook.xml")
    expect(xml).toContain('state="hidden"')
    expect(xml).toContain('state="veryHidden"')
  })
})

// ═══════════════════════════════════════════════════════════════════════
// cell-utils / sheet-utils / image
// ═══════════════════════════════════════════════════════════════════════

describe("letterToCol", () => {
  // The inverse of `colToLetter` only consumes the letter prefix, so an
  // A1-style reference resolves to its column and stops at the digits.
  it("stops at the first non-letter character", () => {
    expect(letterToCol("A1")).toBe(0)
    expect(letterToCol("AB12")).toBe(27)
    expect(letterToCol("A:A")).toBe(0)
  })

  it("returns -1 for an empty string", () => {
    expect(letterToCol("")).toBe(-1)
  })
})

describe("sheetToArrays", () => {
  // Header cells are stringified and trimmed; a blank leading column reads
  // back as "" rather than "null".
  it("renders blank header cells as empty strings", () => {
    const result = sheetToArrays(
      sheet({
        rows: [
          [null, "  Name  ", 2025],
          [1, "Alice", 30],
        ],
      }),
    )

    expect(result.headers).toEqual(["", "Name", "2025"])
    expect(result.data).toEqual([[1, "Alice", 30]])
  })
})

describe("imageFromBase64", () => {
  const anchor = { from: { row: 0, col: 0 } }
  // 1×1 transparent GIF — the smallest real image that survives a writer.
  const gif = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"

  it("accepts a bare base64 payload", () => {
    const image = imageFromBase64(gif, "gif", anchor)

    expect(image.type).toBe("gif")
    expect(image.data[0]).toBe(0x47) // "G" of GIF89a
    expect(image.data).toHaveLength(42)
  })

  it("strips a data URI prefix before decoding", () => {
    const bare = imageFromBase64(gif, "gif", anchor)
    const uri = imageFromBase64(`data:image/gif;base64,${gif}`, "gif", anchor)

    expect(uri.data).toEqual(bare.data)
  })
})
