import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { writeOds } from "../src/ods/writer"
import { writeCsv } from "../src/csv/writer"
import { writeXml } from "../src/xml/data-writer"
import { toHtml } from "../src/export/html"
import { toMarkdown } from "../src/export/markdown"
import { XlsxStreamWriter, writeXlsxStream } from "../src/xlsx/stream-writer"
import { ZipReader } from "../src/zip/reader"
import { InvalidArgumentError } from "../src/errors"
import { MAX_SHEET_NAME_LENGTH } from "../src/_validate"
import type { Sheet } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

async function odsContent(rows: unknown[][]): Promise<string> {
  const buf = await writeOds({ sheets: [{ name: "S", rows: rows as never }] })
  return new TextDecoder().decode(await new ZipReader(buf).extract("content.xml"))
}

const sheetOf = (rows: unknown[][]): Sheet => ({ name: "S", rows: rows as never })

// ═══════════════════════════════════════════════════════════════════════
// #364 — every check hucre performed was on the read side. The write
// path trusted the caller, so ordinary mistakes produced files Excel
// calls corrupt, with no warning.
// ═══════════════════════════════════════════════════════════════════════

describe("sheet name validation", () => {
  const illegal: Array<[string, string]> = [
    ["Sales: Q1/Q2", "a name copied from a report title"],
    ["Data[2024]", "square brackets"],
    ["What?", "a question mark"],
    ["a*b", "an asterisk"],
    ["a\\b", "a backslash"],
  ]

  for (const [name, why] of illegal) {
    it(`rejects ${why}: ${JSON.stringify(name)}`, async () => {
      await expect(writeXlsx({ sheets: [{ name, rows: [["a"]] }] })).rejects.toThrow(
        InvalidArgumentError,
      )
    })
  }

  it("rejects a name past Excel's 31-character limit", async () => {
    const name = "x".repeat(MAX_SHEET_NAME_LENGTH + 1)
    await expect(writeXlsx({ sheets: [{ name, rows: [["a"]] }] })).rejects.toThrow(
      /32 characters.*at most 31/,
    )
  })

  it("accepts a name exactly at the limit", async () => {
    const name = "x".repeat(MAX_SHEET_NAME_LENGTH)
    await expect(writeXlsx({ sheets: [{ name, rows: [["a"]] }] })).resolves.toBeInstanceOf(
      Uint8Array,
    )
  })

  it("rejects an empty name", async () => {
    await expect(writeXlsx({ sheets: [{ name: "", rows: [["a"]] }] })).rejects.toThrow(/is empty/)
  })

  it("rejects leading and trailing apostrophes", async () => {
    // They break quoted range references: 'My Sheet'!A1.
    for (const name of ["'Data", "Data'"]) {
      await expect(writeXlsx({ sheets: [{ name, rows: [["a"]] }] })).rejects.toThrow(/apostrophe/)
    }
  })

  it("rejects the reserved History name, in any case", async () => {
    for (const name of ["History", "history", "HISTORY"]) {
      await expect(writeXlsx({ sheets: [{ name, rows: [["a"]] }] })).rejects.toThrow(/reserved/)
    }
  })

  it("still accepts the punctuation Excel allows", async () => {
    const name = "Sheet (1) - Test & 2"
    await expect(writeXlsx({ sheets: [{ name, rows: [["a"]] }] })).resolves.toBeInstanceOf(
      Uint8Array,
    )
  })

  it("names the sheet position in the message", async () => {
    await expect(
      writeXlsx({
        sheets: [
          { name: "Fine", rows: [["a"]] },
          { name: "Bad:Name", rows: [["b"]] },
        ],
      }),
    ).rejects.toThrow(/sheet 2/)
  })
})

describe("duplicate sheet names", () => {
  it("rejects exact duplicates", async () => {
    await expect(
      writeXlsx({
        sheets: [
          { name: "Data", rows: [["a"]] },
          { name: "Data", rows: [["b"]] },
        ],
      }),
    ).rejects.toThrow(/Duplicate sheet name/)
  })

  it("rejects names differing only in case", async () => {
    // Excel compares case-insensitively; ExcelJS refuses the second with
    // "Worksheet name already exists".
    await expect(
      writeXlsx({
        sheets: [
          { name: "Data", rows: [["a"]] },
          { name: "DATA", rows: [["b"]] },
        ],
      }),
    ).rejects.toThrow(/Duplicate sheet name/)
  })

  it("points at both offending positions", async () => {
    await expect(
      writeXlsx({
        sheets: [
          { name: "A", rows: [["a"]] },
          { name: "B", rows: [["b"]] },
          { name: "a", rows: [["c"]] },
        ],
      }),
    ).rejects.toThrow(/sheets 1 and 3/)
  })
})

describe("sheet name validation applies to every writer", () => {
  it("writeOds", async () => {
    await expect(writeOds({ sheets: [{ name: "Bad:Name", rows: [["a"]] }] })).rejects.toThrow(
      InvalidArgumentError,
    )
  })

  it("XlsxStreamWriter", () => {
    expect(() => new XlsxStreamWriter({ name: "Bad:Name" })).toThrow(InvalidArgumentError)
  })

  it("writeXlsxStream", () => {
    expect(() => writeXlsxStream([], { name: "x".repeat(40) })).toThrow(InvalidArgumentError)
  })

  it("still allows the streaming writer to truncate its own generated names", async () => {
    // A legal 31-character base plus a "_2" suffix exceeds the limit, so
    // rollover names are still truncated — that truncation is hucre's,
    // not the caller's.
    const base = "x".repeat(MAX_SHEET_NAME_LENGTH)
    const writer = new XlsxStreamWriter({ name: base, maxRowsPerSheet: 2 })
    for (let i = 0; i < 5; i++) writer.addRow([i])

    const zip = new ZipReader(await writer.finish())
    const workbookXml = new TextDecoder().decode(await zip.extract("xl/workbook.xml"))
    for (const match of workbookXml.matchAll(/<sheet name="([^"]+)"/g)) {
      expect(match[1]!.length).toBeLessThanOrEqual(MAX_SHEET_NAME_LENGTH)
    }
  })
})

describe("values the format cannot represent", () => {
  it("ODS emits an empty cell for NaN and Infinity, like XLSX", async () => {
    // office:value="NaN" is not a valid ODF float — LibreOffice reads
    // garbage. The XLSX writer already guarded this; ODS did not.
    const xml = await odsContent([[Number.NaN, Number.POSITIVE_INFINITY, 42]])
    expect(xml).not.toContain("NaN")
    expect(xml).not.toContain("Infinity")
    expect(xml).toContain('office:value="42"')
  })

  it("ODS emits an empty cell for an unparseable Date", async () => {
    // It used to write office:date-value="NaN-NaN-NaNTNaN:NaN:NaN".
    const xml = await odsContent([[new Date("nope")]])
    expect(xml).not.toContain("NaN")
  })

  it("the text writers no longer throw a raw RangeError", () => {
    // Four writers called toISOString() unguarded, so an unparseable Date
    // aborted the write with an untyped error — in the streaming writer,
    // after bytes had already gone out.
    const bad = new Date("nope")

    expect(() => writeCsv([[bad]])).not.toThrow()
    expect(() => toHtml(sheetOf([[bad]]))).not.toThrow()
    expect(() => toMarkdown(sheetOf([[bad]]))).not.toThrow()
    expect(() => writeXml([{ d: bad }] as never)).not.toThrow()
  })

  it("keeps writing valid dates", () => {
    const good = new Date(Date.UTC(2024, 0, 15))
    expect(writeCsv([[good]])).toContain("2024-01-15")
    expect(toMarkdown(sheetOf([[good]]))).toContain("2024-01-15")
  })

  it("XLSX keeps its existing empty-cell behaviour for non-finite numbers", async () => {
    const buf = await writeXlsx({ sheets: [{ name: "S", rows: [[Number.NaN, 7]] }] })
    const xml = new TextDecoder().decode(
      await new ZipReader(buf).extract("xl/worksheets/sheet1.xml"),
    )
    expect(xml).not.toContain("NaN")
    expect(xml).toContain(">7<")
  })
})
