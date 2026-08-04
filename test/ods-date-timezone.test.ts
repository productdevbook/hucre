// Regression tests for #415 — ODS dates drifted by the reader's UTC offset
// on every round trip, and the drift accumulated.
//
// These tests only mean anything away from Greenwich: with TZ=UTC the buggy
// and the correct reading coincide, which is why CI never noticed. Each
// test therefore pins a non-UTC zone; Asia/Tokyo (+09:00, no DST) makes the
// shift large enough to change the calendar day within three round trips.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { writeOds } from "../src/ods/writer"
import { readOds } from "../src/ods/reader"
import { streamOdsRows } from "../src/ods/stream"
import { ZipWriter } from "../src/zip/writer"
import type { CellValue, WriteSheet } from "../src/_types"

const encoder = new TextEncoder()

/** A minimal .ods whose first row holds the given date-typed cells. */
async function odsWithDates(...dateValues: string[]): Promise<Uint8Array> {
  const cells = dateValues
    .map(
      (v) =>
        `<table:table-cell office:value-type="date" office:date-value="${v}">` +
        `<text:p>${v}</text:p></table:table-cell>`,
    )
    .join("")
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content ` +
    `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
    `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
    `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">` +
    `<office:body><office:spreadsheet><table:table table:name="Sheet1">` +
    `<table:table-row>${cells}</table:table-row>` +
    `</table:table></office:spreadsheet></office:body></office:document-content>`

  const zip = new ZipWriter()
  zip.add("mimetype", encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), {
    compress: false,
  })
  zip.add("content.xml", encoder.encode(content))
  return await zip.build()
}

describe("ODS #415 — dates do not drift with the reader's time zone", () => {
  beforeEach(() => {
    // Node re-reads TZ on the next Date operation, so this takes effect for
    // the whole test; vi.stubEnv restores the real zone afterwards.
    vi.stubEnv("TZ", "Asia/Tokyo")
  })

  afterEach(() => {
    vi.unstubAllEnvs()
  })

  it("survives repeated write → read round trips unchanged", async () => {
    // The zone has to be in effect while the test runs, not just when the
    // module loaded — assert it, so the guard cannot quietly go decorative.
    expect(new Date("2024-01-15T00:00:00Z").getTimezoneOffset()).toBe(-540)

    const original = new Date("2024-01-15T00:00:00Z")
    let value: CellValue = original

    for (let pass = 0; pass < 4; pass++) {
      const sheets: WriteSheet[] = [{ name: "Sheet1", rows: [[value]] }]
      const wb = await readOds(await writeOds({ sheets }))
      value = wb.sheets[0]!.rows[0]![0]!
      expect(value).toBeInstanceOf(Date)
      // Before the fix this lost nine hours per pass: 2024-01-13 by pass 3.
      expect((value as Date).toISOString()).toBe(original.toISOString())
    }
  })

  it("reads an unqualified office:date-value as UTC", async () => {
    const wb = await readOds(await odsWithDates("2024-01-15T00:00:00"))
    expect((wb.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2024-01-15T00:00:00.000Z")
  })

  it("honours an explicit offset rather than overriding it", async () => {
    // LibreOffice may write one, and it means what it says — appending `Z`
    // unconditionally would move the value by two hours.
    const wb = await readOds(
      await odsWithDates("2024-01-15T00:00:00+02:00", "2024-01-15T05:00:00Z"),
    )
    expect((wb.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2024-01-14T22:00:00.000Z")
    expect((wb.sheets[0]!.rows[0]![1] as Date).toISOString()).toBe("2024-01-15T05:00:00.000Z")
  })

  it("reads a date-only value as UTC, as ECMAScript already did", async () => {
    const wb = await readOds(await odsWithDates("2024-01-15"))
    expect((wb.sheets[0]!.rows[0]![0] as Date).toISOString()).toBe("2024-01-15T00:00:00.000Z")
  })

  it("gives the streaming reader the same instant as readOds", async () => {
    const data = await odsWithDates("2024-01-15T00:00:00")
    const rows = []
    for await (const row of streamOdsRows(data)) rows.push(row)
    expect((rows[0]!.values[0] as Date).toISOString()).toBe("2024-01-15T00:00:00.000Z")
  })

  it("round-trips the meta.xml document dates", async () => {
    const created = new Date("2024-01-15T00:00:00Z")
    const modified = new Date("2024-06-30T12:34:56Z")
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[1]] }],
      properties: { created, modified },
    })
    const wb = await readOds(data)
    expect(wb.properties!.created!.toISOString()).toBe(created.toISOString())
    expect(wb.properties!.modified!.toISOString()).toBe(modified.toISOString())
  })

  it("reads a zone-less meta:creation-date as UTC", async () => {
    // What LibreOffice writes — no zone designator, sub-second precision.
    const meta =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-meta ` +
      `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
      `xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3"><office:meta>` +
      `<meta:creation-date>2024-01-15T00:00:00.123</meta:creation-date>` +
      `</office:meta></office:document-meta>`

    const zip = new ZipWriter()
    zip.add("mimetype", encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), {
      compress: false,
    })
    zip.add("meta.xml", encoder.encode(meta))
    zip.add(
      "content.xml",
      encoder.encode(
        `<?xml version="1.0" encoding="UTF-8"?>` +
          `<office:document-content ` +
          `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
          `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3">` +
          `<office:body><office:spreadsheet><table:table table:name="Sheet1"/>` +
          `</office:spreadsheet></office:body></office:document-content>`,
      ),
    )
    const wb = await readOds(await zip.build())
    expect(wb.properties!.created!.toISOString()).toBe("2024-01-15T00:00:00.123Z")
  })
})
