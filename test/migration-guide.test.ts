import { describe, expect, it } from "vitest"
import * as root from "../src/index"
import {
  DefterError,
  fromHtml,
  HucreError,
  InvalidArgumentError,
  jsonToWorkbook,
  NdjsonStreamWriter,
  openXlsx,
  parseJson,
  parseNdjson,
  readNdjsonStream,
  readObjects,
  saveXlsx,
  sheetToObjects,
  streamCsvRows,
  streamNdjsonRows,
  toHtml,
  toMarkdown,
  unflattenRow,
  validateWithSchema,
  writeCsv,
  workbookToJson,
  writeCsvStream,
  writeJson,
  writeNdjson,
  writeXlsx,
  writeXlsxStream,
} from "../src/index"
import { readXlsx } from "../src/xlsx/reader"
import { parseCsv } from "../src/csv/reader"
import type { CsvReadOptions } from "../src/_types"
import { MAX_INPUT_BYTES } from "../src/limits"
import type { Workbook } from "../src/_types"

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

  it("drops it only when asked, via skipHeaderRow — in both readers", () => {
    expect([...streamCsvRows(source, { header: true, skipHeaderRow: true })]).toEqual([["1", "2"]])
    expect(parseCsv(source, { header: true, skipHeaderRow: true })).toEqual([["1", "2"]])
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

// ── "fromHtml reads cell text the way parseCsv does" ────────────────

describe("fromHtml reads cell text the way parseCsv does", () => {
  const cell = (text: string) => fromHtml(`<table><tr><td>${text}</td></tr></table>`).rows[0][0]

  it("keeps leading zeros instead of turning 007 into 7", () => {
    expect(cell("007")).toBe("007")
  })

  it("matches the guide's before/after table", () => {
    expect(cell("0x1A")).toBe("0x1A")
    expect(cell("Infinity")).toBe("Infinity")
    expect(cell("1,234")).toBe(1234)
    expect(cell("true")).toBe(true)
    expect(cell("2024-01-15")).toBeInstanceOf(Date)
  })

  it("defaults typeInference to true, where parseCsv defaults it to false", () => {
    expect(cell("42")).toBe(42)
    expect(parseCsv("42")[0][0]).toBe("42")
  })

  it("returns cell text exactly as written under typeInference: false", () => {
    const rows = fromHtml("<table><tr><td>42</td></tr></table>", { typeInference: false }).rows
    expect(rows[0][0]).toBe("42")
  })

  it("surfaces a header row and a caption on sheet.a11y", () => {
    const sheet = fromHtml(
      "<table><caption>Sales</caption><thead><tr><th>A</th></tr></thead>" +
        "<tbody><tr><td>1</td></tr></tbody></table>",
    )
    expect(sheet.a11y).toEqual({ summary: "Sales", headerRow: 0 })
  })

  it("returns the rows it read instead of throwing on malformed markup", () => {
    expect(fromHtml("<table><tr><td>a</td></tr></table><!-- unterminated").rows).toEqual([["a"]])
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

// ── "parseJson rejects a workbook instead of mangling it" ───────────

describe("parseJson rejects a workbook instead of mangling it", () => {
  const wb: Workbook = {
    sheets: [
      { name: "S1", rows: [["a"], [1]] },
      { name: "S2", rows: [["b"], [2]] },
    ],
  }

  it("throws ParseError where it used to return one row of nonsense", () => {
    expect(() => parseJson(workbookToJson(wb))).toThrow(HucreError)
  })

  it("offers all three ways forward the guide's diff lists", () => {
    const json = workbookToJson(wb)
    expect(jsonToWorkbook(json)).toEqual(wb)
    expect(parseJson(json, { rowsAt: "S1" }).data).toEqual([{ a: 1 }])
    expect(parseJson(json, { rowsAt: "" }).data).toEqual([{ S1: '[{"a":1}]', S2: '[{"b":2}]' }])
  })

  it("still reads { a: [1,2], b: [3,4] } as one row — the guard is that narrow", () => {
    expect(parseJson('{"a":[1,2],"b":[3,4]}').data).toEqual([{ a: "1, 2", b: "3, 4" }])
  })

  it("keeps today's count-dependent shape under the default, and pins it under shape: sheets", () => {
    expect(workbookToJson({ sheets: [wb.sheets[0]!] })).toBe('[{"a":1}]')
    expect(workbookToJson({ sheets: [wb.sheets[0]!] }, { shape: "sheets" })).toBe(
      '{"S1":[{"a":1}]}',
    )
  })
})

// ── "Nested JSON can be rebuilt: unflattenRow" ──────────────────────

describe("nested JSON can be rebuilt", () => {
  it("is opt-in — the default still destroys the nesting, as it always did", () => {
    expect(writeJson(parseJson('[{"user":{"name":"Ada"}}]').data)).toBe('[{"user.name":"Ada"}]')
  })

  it("restores it under unflatten: true", () => {
    const out = writeJson(parseJson('[{"user":{"name":"Ada"}}]').data, { unflatten: true })
    expect(JSON.parse(out)).toEqual([{ user: { name: "Ada" } }])
  })

  it("takes the same option on writeNdjson and NdjsonStreamWriter", () => {
    expect(writeNdjson([{ "user.name": "Ada" }], { unflatten: true })).toBe(
      '{"user":{"name":"Ada"}}\n',
    )
    const w = new NdjsonStreamWriter({ unflatten: true })
    w.addObject({ "user.name": "Ada" })
    expect(w.finish()).toBe('{"user":{"name":"Ada"}}\n')
  })

  it("does not undo the two losses the guide says it cannot", () => {
    // A joined primitive array is a string by then; a literal dot is a path.
    expect(JSON.parse(writeJson(parseJson('[{"a":[1,2]}]').data, { unflatten: true }))).toEqual([
      { a: "1, 2" },
    ])
    expect(JSON.stringify(unflattenRow({ "a.b": 1 }))).toBe('{"a":{"b":1}}')
  })
})

// ── "Dates come back from JSON" ─────────────────────────────────────

describe("dates come back from JSON", () => {
  const at = new Date("2024-01-15T10:30:00.000Z")

  it("is off by default in both readers", () => {
    expect(parseJson(writeJson([{ at }])).data[0]!.at).toBe(at.toISOString())
    expect(parseCsv(`at\n${at.toISOString()}`, { header: true })[1]![0]).toBe(at.toISOString())
  })

  it("revives the Date under typeInference: true", () => {
    expect(parseJson(writeJson([{ at }]), { typeInference: true }).data[0]!.at).toEqual(at)
  })

  it("accepts the same instants everywhere the option exists", () => {
    for (const raw of ["2024-01-15", "2024-01-15T10:30:00Z", "2024-13-45", "3/4/2021", "2024"]) {
      const viaCsv = parseCsv(`v\n"${raw}"`, { typeInference: true, header: true })[1]![0]
      const viaStream = [
        ...streamCsvRows(`v\n"${raw}"`, { typeInference: true, header: true }),
      ][1]![0]
      const viaJson = parseJson(`[{"v":${JSON.stringify(raw)}}]`, { typeInference: true }).data[0]!
        .v
      const viaNdjson = parseNdjson(`{"v":${JSON.stringify(raw)}}`, { typeInference: true })
        .data[0]!.v
      const asDate = (v: unknown) => v instanceof Date
      expect([raw, asDate(viaJson), asDate(viaNdjson), asDate(viaStream)]).toEqual([
        raw,
        asDate(viaCsv),
        asDate(viaCsv),
        asDate(viaCsv),
      ])
    }
  })

  it("infers only dates for JSON — a string stays a string", () => {
    expect(parseJson('[{"n":"007"}]', { typeInference: true }).data[0]!.n).toBe("007")
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

  it("no longer accepts WriteSheet.threadedComments", async () => {
    // Typed and accepted, written by nothing — see #404. Removing it is
    // the only option that tells the caller anything.
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          threadedComments: [{ id: "{c1}", ref: "A1", personId: "{p1}", text: "hi" }],
        } as never,
      ],
    })
    const workbook = await openXlsx(buf)
    expect(workbook.sheets[0].threadedComments).toBeUndefined()
  })

  it("no longer accepts CsvReadOptions.schema", () => {
    // A removed *type member* leaves no runtime trace, so the claim is a
    // compile-time one: if `schema` came back, the @ts-expect-error below
    // would itself be the error. The row of data proves what the option
    // never did — parseCsv returned it unvalidated.
    const options: CsvReadOptions = {
      header: true,
      // @ts-expect-error — removed in v1; no CSV reader ever validated with it
      schema: { a: { type: "number", required: true } },
    }
    expect(parseCsv("a\r\nnot-a-number", options)).toEqual([["a"], ["not-a-number"]])
  })

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
  it("has no reader for toJson's write-only arrays and columns formats", async () => {
    // The guide's claim is a negative, so what it can check is that the
    // readable format really is readable and the other two really are not.
    const { toJson } = await import("../src/export/json")
    const sheet = {
      name: "S",
      rows: [
        ["a", "b"],
        [1, 2],
      ],
    }
    expect(parseJson(toJson(sheet)).data).toEqual([{ a: 1, b: 2 }])
    expect(parseJson(toJson(sheet, { format: "arrays" })).data).not.toEqual([{ a: 1, b: 2 }])
    expect(parseJson(toJson(sheet, { format: "columns" })).data).not.toEqual([{ a: 1, b: 2 }])
  })

  it("exports writeCsvStream, the counterpart to writeXlsxStream", () => {
    expect(typeof writeCsvStream).toBe("function")
  })

  it("gives escapeFormulae a way back in, and only where it applies", () => {
    const written = writeCsv([["-5", "'quoted'"]], { escapeFormulae: true })
    expect(written).toBe("'-5,'quoted'")
    expect(parseCsv(written, { unescapeFormulae: true })).toEqual([["-5", "'quoted'"]])
  })

  it("escapes formulae in the streaming CSV writers too", async () => {
    const escaped = writeCsv([["=SUM(A1)"]], { escapeFormulae: true })
    const streamed = await bytes(writeCsvStream([["=SUM(A1)"]], { escapeFormulae: true }))
    expect(new TextDecoder().decode(streamed)).toBe(escaped)
  })

  it("quotes values a comment-configured reader would delete", () => {
    const written = writeCsv([["#1", "a"]], { comment: "#" })
    expect(parseCsv(written, { comment: "#" })).toEqual([["#1", "a"]])
    // …and without the option, the guide's warning holds.
    expect(parseCsv(writeCsv([["#1", "a"]]), { comment: "#" })).toEqual([])
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
