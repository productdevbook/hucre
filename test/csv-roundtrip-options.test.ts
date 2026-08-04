import { describe, expect, it } from "vitest"
import { parseCsv } from "../src/csv/reader"
import { writeCsv } from "../src/csv/writer"
import { CsvStreamWriter, streamCsvRows, writeCsvStream } from "../src/csv/stream"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #408 — write options that transform values on the way out need a way
// back in. Everything here is a writeCsv → parseCsv round trip: what goes
// in must come out, or the one-way-ness must be a documented decision.
// ═══════════════════════════════════════════════════════════════════════

const collect = async (stream: ReadableStream<Uint8Array>): Promise<string> => {
  const decoder = new TextDecoder()
  const reader = stream.getReader()
  let out = ""
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    out += decoder.decode(value, { stream: true })
  }
  return out + decoder.decode()
}

// ── skipHeaderRow ────────────────────────────────────────────────────

describe("parseCsv — skipHeaderRow", () => {
  const SIMPLE = "name,qty\r\nfoo,1\r\nbar,2\r\n"

  it("consumes the header row when asked", () => {
    expect(parseCsv(SIMPLE, { header: true, skipHeaderRow: true })).toEqual([
      ["foo", "1"],
      ["bar", "2"],
    ])
  })

  it("keeps the header row by default", () => {
    expect(parseCsv(SIMPLE, { header: true })[0]).toEqual(["name", "qty"])
  })

  it("is inert without header: true", () => {
    expect(parseCsv(SIMPLE, { skipHeaderRow: true })).toEqual(parseCsv(SIMPLE))
  })

  it("still names transformValue columns from the consumed header", () => {
    const seen: string[] = []
    parseCsv(SIMPLE, {
      header: true,
      skipHeaderRow: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["name", "qty", "name", "qty"])
  })

  it("counts maxRows against emitted rows only", () => {
    expect(parseCsv(SIMPLE, { header: true, skipHeaderRow: true, maxRows: 1 })).toEqual([
      ["foo", "1"],
    ])
  })

  it("agrees with streamCsvRows", () => {
    const options = { header: true, skipHeaderRow: true }
    expect(parseCsv(SIMPLE, options)).toEqual(Array.from(streamCsvRows(SIMPLE, options)))
  })
})

// ── escapeFormulae / unescapeFormulae ────────────────────────────────

describe("escapeFormulae round trip", () => {
  const TRIGGERS = ["=SUM(A1)", "+1", "-5", "@handle", "|pipe", "\ttab", "\rcr", "\nlf", "\0nul"]

  it("returns the original values with unescapeFormulae", () => {
    const rows: CellValue[][] = [TRIGGERS]
    const written = writeCsv(rows, { escapeFormulae: true })
    expect(parseCsv(written, { unescapeFormulae: true })).toEqual(rows)
  })

  it("still corrupts without it — the option is opt-in", () => {
    const written = writeCsv([["-5"]], { escapeFormulae: true })
    expect(written).toBe("'-5")
    expect(parseCsv(written)).toEqual([["'-5"]])
  })

  it("strips only what the writer would have added", () => {
    // A value that genuinely starts with an apostrophe is never escaped
    // (the writer only fires on formula triggers), so it must survive.
    expect(parseCsv("'foo", { unescapeFormulae: true })).toEqual([["'foo"]])
    expect(parseCsv("'", { unescapeFormulae: true })).toEqual([["'"]])
  })

  it("runs before type inference, so an escaped number comes back a number", () => {
    const written = writeCsv([["-5"]], { escapeFormulae: true })
    expect(parseCsv(written, { unescapeFormulae: true, typeInference: true })).toEqual([[-5]])
  })

  it("un-escapes header names too", () => {
    const written = writeCsv([["-h"], ["v"]], { escapeFormulae: true })
    const seen: string[] = []
    parseCsv(written, {
      header: true,
      unescapeFormulae: true,
      transformValue: (v, header) => {
        seen.push(header)
        return v
      },
    })
    expect(seen).toEqual(["-h", "-h"])
  })

  it("works the same in streamCsvRows", () => {
    const written = writeCsv([TRIGGERS], { escapeFormulae: true })
    const options = { unescapeFormulae: true }
    expect(Array.from(streamCsvRows(written, options))).toEqual(parseCsv(written, options))
  })
})

describe("escapeFormulae in the streaming writers", () => {
  it("CsvStreamWriter escapes like writeCsv", () => {
    const writer = new CsvStreamWriter({ escapeFormulae: true })
    writer.addRow(["=SUM(A1)", "-5"])
    expect(writer.finish()).toBe(writeCsv([["=SUM(A1)", "-5"]], { escapeFormulae: true }))
  })

  it("writeCsvStream escapes like writeCsv", async () => {
    const out = await collect(writeCsvStream([["=SUM(A1)", "-5"]], { escapeFormulae: true }))
    expect(out).toBe(writeCsv([["=SUM(A1)", "-5"]], { escapeFormulae: true }))
  })
})

// ── comment character ────────────────────────────────────────────────

describe("comment character on write", () => {
  it("quotes a value that would otherwise be read as a comment", () => {
    const rows = [
      ["#1", "a"],
      ["b", "c"],
    ]
    const written = writeCsv(rows, { comment: "#" })
    expect(written).toBe('"#1",a\r\nb,c')
    expect(parseCsv(written, { comment: "#" })).toEqual(rows)
  })

  it("deletes the row without the option — hence the option", () => {
    const written = writeCsv([["#1", "a"]])
    expect(parseCsv(written, { comment: "#" })).toEqual([])
  })

  it("quotes header cells too", () => {
    const written = writeCsv([["v"]], { headers: ["#h"], comment: "#" })
    expect(parseCsv(written, { comment: "#" })).toEqual([["#h"], ["v"]])
  })

  it("leaves values that do not start with the comment character alone", () => {
    expect(writeCsv([["a#b", "c"]], { comment: "#" })).toBe("a#b,c")
  })

  it("is honoured by the streaming writers", async () => {
    const writer = new CsvStreamWriter({ comment: "#" })
    writer.addRow(["#1", "a"])
    expect(writer.finish()).toBe('"#1",a')
    const out = await collect(writeCsvStream([["#1", "a"]], { comment: "#" }))
    expect(out).toBe('"#1",a')
  })
})

// ── Documented one-way options ───────────────────────────────────────

describe("write options documented as one-way", () => {
  // These have no inverse by decision, not by oversight (#408). The tests
  // pin the documented behaviour so that adding an inverse later is a
  // deliberate change to a stated contract rather than a silent one.

  it("nullValue does not survive the round trip", () => {
    expect(parseCsv(writeCsv([[null, "a"]], { nullValue: "NULL" }))).toEqual([["NULL", "a"]])
    // Even the default is one-way: CSV has no null, so it reads back as "".
    expect(parseCsv(writeCsv([[null, "a"]]))).toEqual([["", "a"]])
  })

  it("a custom dateFormat does not survive, while the ISO default does", () => {
    const date = new Date(Date.UTC(2024, 0, 15, 12, 0, 0))
    const custom = writeCsv([[date]], { dateFormat: "DD/MM/YYYY" })
    expect(parseCsv(custom, { typeInference: true })).toEqual([["15/01/2024"]])
    expect(parseCsv(writeCsv([[date]]), { typeInference: true })).toEqual([[date]])
  })
})
