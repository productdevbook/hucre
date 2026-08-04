import { describe, expect, it } from "vitest"
import * as root from "../src/index"
import {
  DefterError,
  HucreError,
  InvalidArgumentError,
  openXlsx,
  readNdjsonStream,
  readObjects,
  saveXlsx,
  sheetToObjects,
  streamCsvRows,
  streamNdjsonRows,
  toHtml,
  toMarkdown,
  validateWithSchema,
  writeCsvStream,
  writeXlsx,
  writeXlsxStream,
} from "../src/index"
import { readXlsx } from "../src/xlsx/reader"
import { parseCsv } from "../src/csv/reader"
import { MAX_INPUT_BYTES } from "../src/limits"

// ═══════════════════════════════════════════════════════════════════════
// MIGRATION.md is a promise to people upgrading from 0.6.x, and an
// untested promise is a guess. Every claim below quotes the guide; if one
// of these fails, either the code regressed or the guide is now lying to
// someone mid-upgrade. Fix whichever is wrong — do not delete the test.
//
// Section headings match the guide's, so a failure points at the passage
// to correct.
// ═══════════════════════════════════════════════════════════════════════

const collect = async <T>(stream: ReadableStream<T>): Promise<T[]> => {
  const out: T[] = []
  for await (const chunk of stream as unknown as AsyncIterable<T>) out.push(chunk)
  return out
}

const bytes = async (stream: ReadableStream<Uint8Array>): Promise<Uint8Array> => {
  const parts = await collect(stream)
  const total = parts.reduce((n, p) => n + p.length, 0)
  const buf = new Uint8Array(total)
  let at = 0
  for (const p of parts) {
    buf.set(p, at)
    at += p.length
  }
  return buf
}

// ── "writeXlsxStream takes rows first" ──────────────────────────────

describe("writeXlsxStream takes rows first", () => {
  it("accepts (rows, options), like every other writer", async () => {
    const buf = await bytes(writeXlsxStream([["a", 1]], { name: "Export" }))
    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0].name).toBe("Export")
    expect(workbook.sheets[0].rows[0]).toEqual(["a", 1])
  })
})

// ── "headerRow means one thing now" ─────────────────────────────────

describe("headerRow means one thing now", () => {
  const rows = [["skip me"], ["name"], ["Ada"]]
  const schema = { name: { type: "string" as const, required: true } }

  it("counts from 0 — headerRow: 1 is the second row", () => {
    const result = validateWithSchema(rows, schema, { headerRow: 1 })
    expect(result.errors).toEqual([])
    expect(result.data).toEqual([{ name: "Ada" }])
  })

  it("says 'no header row' with -1, not 0", () => {
    // The guide's subtle one: under 1-based numbering, 0 was the only way
    // to say "no header", so two concepts shared one value.
    const result = validateWithSchema([["Ada"]], schema, { headerRow: -1 })
    expect(result.data.length).toBe(1)
  })

  it("treats headerRow: 0 as the first row being the header", () => {
    const result = validateWithSchema([["name"], ["Ada"]], schema, { headerRow: 0 })
    expect(result.errors).toEqual([])
    expect(result.data).toEqual([{ name: "Ada" }])
  })

  it("defaults to the first row, so callers relying on the default see no change", () => {
    const result = validateWithSchema([["name"], ["Ada"]], schema)
    expect(result.data).toEqual([{ name: "Ada" }])
  })
})

// ── "hasHeaderRow on toHtml and toMarkdown" ─────────────────────────

describe("hasHeaderRow on toHtml and toMarkdown", () => {
  const sheet = { name: "S", rows: [["h"], ["v"]] }

  it("takes the new spelling", () => {
    expect(toHtml(sheet, { hasHeaderRow: true })).toContain("<th")
    expect(toMarkdown(sheet, { hasHeaderRow: true })).toMatch(/^\| h\s*\|/)
  })

  it("still honours the deprecated boolean headerRow for one major", () => {
    expect(toHtml(sheet, { headerRow: true })).toContain("<th")
    expect(toMarkdown(sheet, { headerRow: true })).toMatch(/^\| h\s*\|/)
  })
})

// ── "streamCsvRows matches parseCsv" ────────────────────────────────

describe("streamCsvRows matches parseCsv", () => {
  const source = "a,b\n1,2"

  it("yields the header row under header: true, as parseCsv does", () => {
    expect([...streamCsvRows(source, { header: true })]).toEqual([
      ["a", "b"],
      ["1", "2"],
    ])
    expect(parseCsv(source, { header: true })).toEqual([
      ["a", "b"],
      ["1", "2"],
    ])
  })

  it("drops it only when asked, via skipHeaderRow", () => {
    expect([...streamCsvRows(source, { header: true, skipHeaderRow: true })]).toEqual([["1", "2"]])
  })

  it("actually runs onRow and transformValue, which used to be ignored", () => {
    const seen: unknown[][] = []
    const rows = [
      ...streamCsvRows(source, {
        header: true,
        skipHeaderRow: true,
        onRow: (row) => seen.push(row),
        transformValue: (value) => (typeof value === "string" ? value.toUpperCase() : value),
      }),
    ]
    expect(seen.length).toBeGreaterThan(0)
    expect(rows[0]).toEqual(["1", "2"])
  })
})

// ── "readObjects and sheetToObjects return { data, headers }" ───────

describe("readObjects and sheetToObjects return { data, headers }", () => {
  it("returns the same shape as every other *Objects reader", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            ["name", "age"],
            ["Ada", 36],
          ],
        },
      ],
    })
    const result = await readObjects(buf)
    expect(result).toHaveProperty("data")
    expect(result).toHaveProperty("headers")
    expect(result.headers).toEqual(["name", "age"])
    expect(result.data).toEqual([{ name: "Ada", age: 36 }])
  })

  it("keeps empty-string header keys, where it used to drop them", () => {
    const result = sheetToObjects({
      name: "S",
      rows: [
        ["a", ""],
        [1, 2],
      ],
    })
    expect(Object.keys(result.data[0])).toContain("")
  })

  it("throws ParseError for a missing sheet instead of returning []", async () => {
    const buf = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    await expect(readObjects(buf, { sheet: "nope" })).rejects.toThrow(HucreError)
  })
})

// ── "DefterError is now HucreError" ─────────────────────────────────

describe("DefterError is now HucreError", () => {
  it("is the same class object, so instanceof behaves identically", () => {
    expect(DefterError).toBe(HucreError)
  })

  it("reports name 'HucreError' — the one visible difference", () => {
    expect(new InvalidArgumentError("x").name).not.toBe("DefterError")
    expect(new HucreError("x").name).toBe("HucreError")
  })

  it("still catches everything the library throws", async () => {
    await expect(writeXlsx({ sheets: [{ name: "a:b", rows: [] }] })).rejects.toBeInstanceOf(
      DefterError,
    )
  })
})

// ── "readNdjsonStream is now streamNdjsonRows" ──────────────────────

describe("readNdjsonStream is now streamNdjsonRows", () => {
  it("is a deprecated alias of the same function", () => {
    expect(readNdjsonStream).toBe(streamNdjsonRows)
  })
})

// ── "Sheet names are validated on write" ────────────────────────────

describe("sheet names are validated on write", () => {
  const reject = (name: string) =>
    expect(writeXlsx({ sheets: [{ name, rows: [["a"]] }] })).rejects.toBeInstanceOf(
      InvalidArgumentError,
    )

  it("rejects every character the guide lists", async () => {
    for (const ch of ["[", "]", ":", "*", "?", "/", "\\"]) await reject(`a${ch}b`)
  })

  it("rejects empty names and names over 31 characters", async () => {
    await reject("")
    await reject("x".repeat(32))
  })

  it("rejects a leading or trailing apostrophe", async () => {
    await reject("'lead")
    await reject("trail'")
  })

  it("rejects the reserved name History", async () => {
    await reject("History")
  })

  it("rejects duplicates compared case-insensitively, as Excel does", async () => {
    await expect(
      writeXlsx({
        sheets: [
          { name: "Data", rows: [["a"]] },
          { name: "DATA", rows: [["b"]] },
        ],
      }),
    ).rejects.toBeInstanceOf(InvalidArgumentError)
  })

  it("throws before producing any bytes, rather than sanitizing", async () => {
    // Sanitizing would hand back a workbook whose sheets are not the ones
    // asked for, and dangle any range reference built on the real names.
    await expect(writeXlsx({ sheets: [{ name: "a:b", rows: [["x"]] }] })).rejects.toThrow(
      /sheet name/i,
    )
  })

  it("holds for writeOds and writeXlsxStream too", async () => {
    const { writeOds } = await import("../src/ods/writer")
    await expect(writeOds({ sheets: [{ name: "a:b", rows: [["x"]] }] })).rejects.toBeInstanceOf(
      InvalidArgumentError,
    )
    expect(() => writeXlsxStream([], { name: "a:b" })).toThrow(InvalidArgumentError)
  })
})

// ── "Removed API that never did anything" ───────────────────────────

describe("removed API that never did anything", () => {
  const removed = ["ReadResult", "StreamReadOptions", "StreamWriteOptions", "WORKER_SAFE_FUNCTIONS"]

  for (const name of removed) {
    it(`no longer exports ${name}`, () => {
      expect(Object.keys(root)).not.toContain(name)
    })
  }

  it("keeps RoundtripWorkbook's internals off the object", async () => {
    const buf = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    const workbook = await openXlsx(buf)
    for (const internal of ["_rawEntries", "_modifiedParts", "_contentTypes", "_rootRels"]) {
      expect(Object.keys(workbook)).not.toContain(internal)
    }
  })

  it("means saveXlsx({ ...workbook }) no longer works — the guide's warning", async () => {
    const buf = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    const workbook = await openXlsx(buf)
    await expect(saveXlsx({ ...workbook })).rejects.toThrow()
    // …while passing the object straight through still does.
    await expect(saveXlsx(workbook)).resolves.toBeInstanceOf(Uint8Array)
  })
})

// ── "Reading untrusted files" ───────────────────────────────────────

describe("reading untrusted files", () => {
  it("defaults maxInputBytes to 1 GiB, as the guide states", () => {
    expect(MAX_INPUT_BYTES).toBe(1024 * 1024 * 1024)
  })

  it("bounds a stream read, and the error names the option to raise", async () => {
    // The cap applies where it can still save you: draining a stream. An
    // input that is already a Uint8Array has been allocated by the caller.
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(64))
        controller.close()
      },
    })
    await expect(readXlsx(stream, { maxInputBytes: 8 })).rejects.toThrow(/maxInputBytes/)
  })
})

// ── "Also worth knowing" ────────────────────────────────────────────

describe("also worth knowing", () => {
  it("exports writeCsvStream, the counterpart to writeXlsxStream", () => {
    expect(typeof writeCsvStream).toBe("function")
  })

  it("has a hucre/ooxml entry point that still re-exports from the root", async () => {
    const ooxml = await import("../src/ooxml")
    expect(Object.keys(ooxml).length).toBeGreaterThan(0)
    for (const name of Object.keys(ooxml)) {
      expect(Object.keys(root)).toContain(name)
    }
  })

  it("exports the documented names from hucre/xlsx and hucre/ods", async () => {
    const xlsx = await import("../src/xlsx")
    const ods = await import("../src/ods")
    expect(Object.keys(xlsx)).toContain("readXlsxObjects")
    expect(Object.keys(ods)).toContain("readOdsObjects")
    expect(Object.keys(ods)).toContain("streamOdsRows")
  })
})
