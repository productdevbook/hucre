import { describe, expect, it, vi } from "vitest"
import { writeCsv, writeCsvObjects } from "../src/csv/writer"
import { CsvStreamWriter, writeCsvStream } from "../src/csv/stream"

// ═══════════════════════════════════════════════════════════════════════
// #439 — `CsvWriteOptions.dateFormat` used to go through a private
// formatter that had nothing to do with the library's own. It accepted a
// different token vocabulary (`YYYY MM DD HH mm ss`), read local-time
// components while the no-format path read UTC, and substituted with a
// non-global `.replace()`.
//
// These pin all four consequences. The CSV writers now delegate to
// `formatDate`, so a format string means one thing everywhere.
// ═══════════════════════════════════════════════════════════════════════

const AT = new Date(Date.UTC(2024, 0, 15, 2, 30, 45))

async function drain(stream: ReadableStream<Uint8Array>): Promise<string> {
  const chunks: string[] = []
  const reader = stream.getReader()
  const decoder = new TextDecoder()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(decoder.decode(value, { stream: true }))
  }
  chunks.push(decoder.decode())
  return chunks.join("")
}

describe("CSV dateFormat", () => {
  it("accepts the library's own Excel-style tokens", () => {
    // Against the old formatter this produced the literal text
    // "yyyy-30-dd": `mm` matched *minutes* and nothing matched yyyy or dd.
    expect(writeCsv([[AT]], { dateFormat: "yyyy-mm-dd" })).toBe("2024-01-15")
  })

  it("still accepts the uppercase spelling", () => {
    expect(writeCsv([[AT]], { dateFormat: "YYYY-MM-DD" })).toBe("2024-01-15")
    expect(writeCsv([[AT]], { dateFormat: "DD/MM/YYYY" })).toBe("15/01/2024")
  })

  it("substitutes every occurrence of a token, not only the first", () => {
    // The old `.replace()` had no /g, so the second MM stayed literal.
    expect(writeCsv([[AT]], { dateFormat: "MM/DD/YYYY MM" })).toBe("01/15/2024 01")
  })

  it("reads UTC components, so the output does not depend on the machine", () => {
    // 02:30 UTC is the previous day in New York. The old formatter used
    // getFullYear()/getMonth()/getDate() and produced 2024-01-14 there,
    // while the no-format path produced an ISO string in UTC — so passing
    // a format silently changed the timezone basis.
    vi.stubEnv("TZ", "America/New_York")
    expect(writeCsv([[AT]], { dateFormat: "YYYY-MM-DD HH:mm:ss" })).toBe("2024-01-15 02:30:45")
    vi.stubEnv("TZ", "Asia/Tokyo")
    expect(writeCsv([[AT]], { dateFormat: "YYYY-MM-DD HH:mm:ss" })).toBe("2024-01-15 02:30:45")
    vi.unstubAllEnvs()
  })

  it("keeps the ISO default and the invalid-date guard", () => {
    expect(writeCsv([[AT]])).toBe("2024-01-15T02:30:45.000Z")
    expect(writeCsv([[new Date("nonsense")]])).toBe("")
    expect(writeCsv([[new Date("nonsense")]], { dateFormat: "YYYY-MM-DD" })).toBe("")
  })

  it("supports the tokens the private formatter never had", () => {
    expect(writeCsv([[AT]], { dateFormat: "d mmm yyyy" })).toBe("15 Jan 2024")
    expect(writeCsv([[AT]], { dateFormat: "h:mm AM/PM" })).toBe("2:30 AM")
  })

  it("means the same thing in every CSV writer", async () => {
    const format = "yyyy-mm-dd"
    const expected = "2024-01-15"

    expect(writeCsv([[AT]], { dateFormat: format })).toBe(expected)
    expect(writeCsvObjects([{ d: AT }], { dateFormat: format })).toContain(expected)

    const incremental = new CsvStreamWriter({ dateFormat: format })
    incremental.addRow([AT])
    expect(incremental.finishText()).toBe(expected)

    expect(await drain(writeCsvStream([[AT]], { dateFormat: format }))).toBe(expected)
  })
})
