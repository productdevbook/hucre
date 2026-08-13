import { describe, expect, it } from "vitest"
import { parseXml } from "../src/xml/parser"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// ═══════════════════════════════════════════════════════════════════════
// #493 — XML 1.0 §2.11 requires a processor to turn a literal CRLF, and
// a literal lone CR, into a single LF before the application sees the
// content. hucre's writer knew this and the parser did not, so the two
// disagreed with each other.
//
// Excel writes a multi-line cell with a literal CRLF inside `<t>`, so
// `readXlsx` returned "line one\r\nline two" where the same authored
// workbook saved as XLSB — which stores a bare LF — gave
// "line one\nline two". One cell, two containers, two strings.
//
// Nothing caught it because every multi-line-string test in the suite
// builds its own XML with `\n`, so the reader had never been shown what
// Excel actually emits.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()
const dec = new TextDecoder()

const textOf = (xml: string): string => {
  const el = parseXml(xml)
  return el.children.filter((c): c is string => typeof c === "string").join("")
}

describe("literal line endings in text", () => {
  it("CRLF becomes LF", () => {
    expect(textOf("<t>line one\r\nline two</t>")).toBe("line one\nline two")
  })

  it("a lone CR becomes LF", () => {
    expect(textOf("<t>line one\rline two</t>")).toBe("line one\nline two")
  })

  it("an LF is left alone", () => {
    expect(textOf("<t>line one\nline two</t>")).toBe("line one\nline two")
  })

  it("a run of them normalizes to a run of LFs", () => {
    expect(textOf("<t>a\r\n\r\nb\r\rc</t>")).toBe("a\n\nb\n\nc")
  })
})

describe("a character reference is not a literal line ending", () => {
  it("&#13; survives as CR", () => {
    // This is why normalization runs on the raw source *before* entity
    // decoding, and why it is a separate pass rather than a step inside
    // the decoder. hucre's own writer escapes CR as `&#13;` precisely so
    // a deliberate one can be told apart from a line break.
    expect(textOf("<t>a&#13;b</t>")).toBe("a\rb")
  })

  it("&#13;&#10; survives as CRLF", () => {
    expect(textOf("<t>a&#13;&#10;b</t>")).toBe("a\r\nb")
  })

  it("&#xD; survives too", () => {
    expect(textOf("<t>a&#xD;b</t>")).toBe("a\rb")
  })
})

describe("attribute values get the same treatment", () => {
  it("a literal CRLF in an attribute normalizes", () => {
    expect(parseXml('<c v="a\r\nb"/>').attrs.v).toBe("a\nb")
  })

  it("and a character reference in one does not", () => {
    expect(parseXml('<c v="a&#13;b"/>').attrs.v).toBe("a\rb")
  })
})

describe("through the readers", () => {
  /** Put a raw shared string into a workbook, bytes and all. */
  async function withSharedString(inner: string): Promise<Uint8Array> {
    const base = await writeXlsx({ sheets: [{ name: "S", rows: [["placeholder"]] }] })
    const all = await new ZipReader(base).extractAll()
    const zw = new ZipWriter()
    for (const [name, data] of all) {
      zw.add(
        name,
        name === "xl/sharedStrings.xml"
          ? enc.encode(dec.decode(data).replace(/<si>.*<\/si>/, `<si><t>${inner}</t></si>`))
          : data,
      )
    }
    return zw.build()
  }

  it("a cell Excel wrote with a literal CRLF reads as LF", async () => {
    const wb = await readXlsx(await withSharedString("line one\r\nline two"))

    expect(wb.sheets[0]!.rows[0]![0]).toBe("line one\nline two")
  })

  it("a deliberate CR still round-trips through hucre's own writer", async () => {
    // The writer escapes it; the parser must not then eat it.
    const wb = await readXlsx(
      await writeXlsx({ sheets: [{ name: "S", rows: [["a\rb", "c\r\nd", "e\nf"]] }] }),
    )

    expect(wb.sheets[0]!.rows[0]).toEqual(["a\rb", "c\r\nd", "e\nf"])
  })
})
