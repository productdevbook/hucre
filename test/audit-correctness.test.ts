import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { toMarkdown } from "../src/export/markdown"
import { formatValue } from "../src/_format"
import { parseCoreProperties } from "../src/xlsx/doc-props-reader"
import { parseUtcDefaultDateTime } from "../src/_date"
import type { Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #474 — three of the smaller items from the audit, each real and each
// too small for its own issue.
// ═══════════════════════════════════════════════════════════════════════

// ── docProps timestamps ──────────────────────────────────────────────
//
// `parseW3CDTF` was `new Date(value)`, which under ECMA-262 reads an
// unqualified date-time as *local*. W3CDTF requires a zone designator and
// hucre's own writer always emits `Z`, so compliant files were fine — a
// non-compliant producer shifted `created` and `modified` by the reader's
// own offset. This is the #415 shape in a third place, so the three now
// share one function instead of drifting apart again.

const CORE = (created: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
  ` xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/">` +
  `<dcterms:created>${created}</dcterms:created></cp:coreProperties>`

describe("docProps timestamps without a zone designator", () => {
  it("reads a bare date-time as UTC, not as the reader's local time", () => {
    const props = parseCoreProperties(CORE("2024-01-15T10:30:00"))

    expect(props.created?.toISOString()).toBe("2024-01-15T10:30:00.000Z")
  })

  it("leaves an explicit zone exactly as the file states it", () => {
    // The file said +02:00 and means it. Only silence is filled in.
    expect(parseCoreProperties(CORE("2024-01-15T10:30:00+02:00")).created?.toISOString()).toBe(
      "2024-01-15T08:30:00.000Z",
    )
    expect(parseCoreProperties(CORE("2024-01-15T10:30:00Z")).created?.toISOString()).toBe(
      "2024-01-15T10:30:00.000Z",
    )
  })

  it("still rejects what is not a date at all", () => {
    expect(parseCoreProperties(CORE("not a date")).created).toBeUndefined()
    expect(parseUtcDefaultDateTime("")).toBeUndefined()
  })

  it("round-trips a workbook hucre wrote, which always carries Z", async () => {
    const created = new Date("2024-03-01T09:00:00Z")
    const bytes = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]] }],
      properties: { created },
    })

    expect((await readXlsx(bytes)).properties?.created?.toISOString()).toBe(created.toISOString())
  })
})

// ── Markdown inline escaping ─────────────────────────────────────────
//
// `escapePipe` protected the table's structure and nothing else, so a
// cell reading `*not emphasis*` rendered as emphasis.

function sheet(rows: Sheet["rows"]): Sheet {
  return { name: "S", rows }
}

describe("toMarkdown escapes what would change the rendering", () => {
  it("a cell of *not emphasis* stays those words", () => {
    const md = toMarkdown(sheet([["text"], ["*not emphasis*"]]))

    expect(md).toContain("\\*not emphasis\\*")
    expect(md).not.toContain("| *not emphasis* |")
  })

  it("covers the rest of the inline syntax", () => {
    const md = toMarkdown(sheet([["v"], ["_a_ `b` [c] <d> e\\f"]]))

    // `>` is deliberately not escaped: a blockquote needs the start of a
    // line and a table cell is never one, so escaping it would only add
    // noise. `<` is escaped because raw HTML does work inside a cell.
    expect(md).toContain("\\_a\\_ \\`b\\` \\[c\\] \\<d> e\\\\f")
  })

  it("still escapes the structural characters, which is not optional", () => {
    // Losing a pipe loses the table, so this happens either way.
    const off = toMarkdown(sheet([["v"], ["a|b\nc"]]), { escapeInline: false })

    expect(off).toContain("a\\|b<br>c")
  })

  it("leaves inline syntax alone when the caller wants it rendered", () => {
    const md = toMarkdown(sheet([["v"], ["**bold**"]]), { escapeInline: false })

    expect(md).toContain("| **bold** |")
  })

  it("does not escape numbers or dates into nonsense", () => {
    const md = toMarkdown(
      sheet([
        ["n", "d"],
        [1234.5, new Date("2024-01-15T00:00:00Z")],
      ]),
    )

    expect(md).toContain("1234.5")
    expect(md).toContain("2024-01-15")
    expect(md).not.toContain("\\")
  })
})

// ── Locale grouping ──────────────────────────────────────────────────
//
// Since #456 the separator comes from Intl for any locale, but the
// grouping *pattern* was fixed at threes — so hi-IN got the right
// separator in the wrong places.

describe("formatValue groups the way the locale does", () => {
  it("matches Intl for the Indian system", () => {
    // 12,345,678.50 was the old answer. The lakh/crore system puts the
    // first separator after three digits and every one after that after
    // two.
    expect(formatValue(12345678.5, "#,##0.00", { locale: "hi-IN" })).toBe("1,23,45,678.50")
    expect(formatValue(12345678.5, "#,##0.00", { locale: "en-IN" })).toBe("1,23,45,678.50")
  })

  it("leaves the three-digit locales exactly where they were", () => {
    expect(formatValue(12345678.5, "#,##0.00", { locale: "en-US" })).toBe("12,345,678.50")
    expect(formatValue(12345678.5, "#,##0.00", { locale: "de-DE" })).toBe("12.345.678,50")
    expect(formatValue(12345678.5, "#,##0.00", { locale: "tr-TR" })).toBe("12.345.678,50")
  })

  it("agrees with Intl.NumberFormat across all of them", () => {
    // The check that does not need updating when a locale's data moves.
    for (const locale of ["en-US", "de-DE", "fr-FR", "tr-TR", "hi-IN", "en-IN", "es-ES"]) {
      const want = new Intl.NumberFormat(locale, {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(12345678.5)

      expect(formatValue(12345678.5, "#,##0.00", { locale }), locale).toBe(want)
    }
  })

  it("groups short numbers the same as long ones", () => {
    // The Indian system's first group is three wide, so anything under a
    // lakh looks like every other locale.
    expect(formatValue(1234.5, "#,##0.00", { locale: "hi-IN" })).toBe("1,234.50")
    expect(formatValue(123456.5, "#,##0.00", { locale: "hi-IN" })).toBe("1,23,456.50")
  })

  it("does not group when the format did not ask", () => {
    expect(formatValue(12345678, "0", { locale: "hi-IN" })).toBe("12345678")
  })

  it("leaves a Special format's literals alone", () => {
    // `000-00-0000` interleaves literals with placeholders; re-grouping
    // those digits would be nonsense.
    expect(formatValue(123456789, "000-00-0000", { locale: "hi-IN" })).toBe("123-45-6789")
  })

  it("still refuses a locale Intl cannot use", () => {
    expect(() => formatValue(1, "#,##0", { locale: "not a tag" })).toThrow()
  })
})
