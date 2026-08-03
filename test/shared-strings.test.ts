import { describe, expect, it } from "vitest"
import { parseSharedStrings } from "../src/xlsx/shared-strings"
import { parseStyles } from "../src/xlsx/styles"

// ── Helpers ──────────────────────────────────────────────────────────

const SST_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

/** Wrap `<si>` fragments in a real `<sst>` document, as Excel writes it. */
function sst(inner: string): ReturnType<typeof parseSharedStrings> {
  return parseSharedStrings(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${SST_NS}">${inner}</sst>`)
}

/**
 * Parse one rich-text run built from the given `<rPr>` body. The fragment is
 * indented on purpose: pretty-printed sharedStrings.xml puts text nodes
 * between every element, so the run and run-property walkers must step over
 * them to find the elements they care about.
 */
function runFont(rPrBody: string) {
  const entries = sst(
    `<si>\n  <r>\n    <rPr>\n      ${rPrBody}\n    </rPr>\n    <t>text</t>\n  </r>\n</si>`,
  )
  return entries[0].richText![0].font
}

// ═══════════════════════════════════════════════════════════════════════
// Document level: only <si> children are entries. Everything else in an
// <sst> (the count/uniqueCount attributes, a trailing <extLst>, and the
// indentation text nodes between elements) must be ignored rather than
// shifting every shared-string index by one.
// ═══════════════════════════════════════════════════════════════════════

describe("parseSharedStrings — document structure", () => {
  it("returns an empty table for an <sst> with no entries", () => {
    expect(sst("")).toEqual([])
    expect(parseSharedStrings(`<sst xmlns="${SST_NS}" count="0" uniqueCount="0"/>`)).toEqual([])
  })

  it("ignores non-<si> siblings so indices stay aligned", () => {
    const strings = sst(`
      <si><t>first</t></si>
      <extLst><ext uri="{ABC}"><whatever/></ext></extLst>
      <si><t>second</t></si>
    `)
    expect(strings.map((s) => s.text)).toEqual(["first", "second"])
  })

  it("ignores whitespace text nodes between entries", () => {
    // The indentation between <si> elements arrives as string children of
    // the root; treating one as an entry would corrupt every index.
    const strings = sst("\n  <si><t>a</t></si>\n  <si><t>b</t></si>\n")
    expect(strings).toHaveLength(2)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <si> shapes. CT_Rst allows a bare <t>, a sequence of <r> runs, and
// phonetic hints (<rPh>, <phoneticPr>) that must never leak into text.
// ═══════════════════════════════════════════════════════════════════════

describe("parseSharedStrings — <si> shapes", () => {
  it("reads an empty <si/> as an empty string", () => {
    // Excel emits a bare <si/> for a shared blank; dropping it would shift
    // every later index.
    const strings = sst("<si/>")
    expect(strings).toHaveLength(1)
    expect(strings[0]).toEqual({ text: "" })
  })

  it("reads a self-closing <t/> as an empty string", () => {
    expect(sst("<si><t/></si>")[0]).toEqual({ text: "" })
  })

  it('preserves whitespace under xml:space="preserve"', () => {
    const strings = sst(`<si><t xml:space="preserve">  padded  </t></si>`)
    expect(strings[0].text).toBe("  padded  ")
  })

  it("ignores phonetic runs so only the base text is returned", () => {
    // Japanese workbooks carry furigana in <rPh>; its <t> must not be
    // concatenated onto the visible value.
    const strings = sst(
      `<si><t>東京</t><rPh sb="0" eb="2"><t>とうきょう</t></rPh><phoneticPr fontId="1"/></si>`,
    )
    expect(strings[0].text).toBe("東京")
    expect(strings[0].richText).toBeUndefined()
  })

  it("decodes OOXML _xHHHH_ escapes in a simple string", () => {
    expect(sst("<si><t>a_x000D__x000A_b</t></si>")[0].text).toBe("a\r\nb")
  })

  it("decodes XML entities in a simple string", () => {
    expect(sst("<si><t>Tom &amp; &quot;Jerry&quot; &lt;x&gt;</t></si>")[0].text).toBe(
      'Tom & "Jerry" <x>',
    )
  })

  it("leaves richText undefined for a plain string", () => {
    // Consumers branch on `richText` to decide between <t> and <r> output;
    // an empty array here would round-trip a plain string as rich text.
    expect(sst("<si><t>plain</t></si>")[0].richText).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Rich text: <r> runs. `text` is the concatenation of the runs and must
// stay consistent with the per-run text.
// ═══════════════════════════════════════════════════════════════════════

describe("parseSharedStrings — rich text runs", () => {
  it("concatenates run text into the flat `text` field", () => {
    const strings = sst("<si><r><t>Hello</t></r><r><t>, </t></r><r><t>world</t></r></si>")
    expect(strings[0].text).toBe("Hello, world")
    expect(strings[0].richText!.map((r) => r.text)).toEqual(["Hello", ", ", "world"])
  })

  it("omits `font` on runs without <rPr>", () => {
    const strings = sst("<si><r><t>bare</t></r></si>")
    expect(strings[0].richText![0]).toEqual({ text: "bare" })
    expect("font" in strings[0].richText![0]).toBe(false)
  })

  it("keeps an empty run rather than collapsing it", () => {
    // A run with formatting but no text is where Excel parks a cursor
    // format; dropping it changes the run count on round-trip.
    const strings = sst("<si><r><rPr><b/></rPr></r><r><t>x</t></r></si>")
    expect(strings[0].richText).toHaveLength(2)
    expect(strings[0].richText![0].text).toBe("")
    expect(strings[0].richText![0].font?.bold).toBe(true)
    expect(strings[0].text).toBe("x")
  })

  it("decodes escapes per run, not across the joined text", () => {
    const strings = sst("<si><r><t>a_x000A_</t></r><r><t>_x0009_b</t></r></si>")
    expect(strings[0].richText!.map((r) => r.text)).toEqual(["a\n", "\tb"])
    expect(strings[0].text).toBe("a\n\tb")
  })

  it("ignores unknown children inside a run", () => {
    const strings = sst('<si><r><rPr><b/></rPr><unknown val="1"/><t>ok</t></r></si>')
    expect(strings[0].richText![0].text).toBe("ok")
    expect(strings[0].richText![0].font?.bold).toBe(true)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <rPr> run properties. Each toggle in ECMA-376 §18.8.x defaults to *on*
// when the element is present without a `val`, and off only for
// val="0"/"false" — the accept-or-drop grammar the parser implements.
// ═══════════════════════════════════════════════════════════════════════

describe("parseSharedStrings — <rPr> boolean toggles", () => {
  it("treats a bare <b/>, <i/>, <strike/> as on", () => {
    const font = runFont("<b/><i/><strike/>")
    expect(font).toMatchObject({ bold: true, italic: true, strikethrough: true })
  })

  it('treats val="1" and val="true" as on', () => {
    expect(runFont('<b val="1"/><i val="true"/><strike val="1"/>')).toMatchObject({
      bold: true,
      italic: true,
      strikethrough: true,
    })
  })

  it('treats val="0" and val="false" as off', () => {
    // These appear when a run explicitly cancels inherited cell formatting;
    // reading them as `true` would bold text Excel renders plain.
    expect(runFont('<b val="0"/><i val="false"/><strike val="0"/>')).toMatchObject({
      bold: false,
      italic: false,
      strikethrough: false,
    })
  })
})

describe("parseSharedStrings — <rPr> underline", () => {
  it("maps a bare <u/> to true", () => {
    expect(runFont("<u/>")?.underline).toBe(true)
  })

  it('maps val="single" to true', () => {
    expect(runFont('<u val="single"/>')?.underline).toBe(true)
  })

  it("maps the three named accounting/double variants", () => {
    expect(runFont('<u val="double"/>')?.underline).toBe("double")
    expect(runFont('<u val="singleAccounting"/>')?.underline).toBe("singleAccounting")
    expect(runFont('<u val="doubleAccounting"/>')?.underline).toBe("doubleAccounting")
  })

  it('reads val="none" as underlined — a known lossy fallback', () => {
    // <u val="none"/> is legal and means "not underlined", but the parser
    // funnels every unrecognised token into `true`. Documented here so the
    // behaviour is a deliberate choice rather than an accident.
    expect(runFont('<u val="none"/>')?.underline).toBe(true)
  })
})

describe("parseSharedStrings — <rPr> scalar properties", () => {
  it("reads sz as a number", () => {
    expect(runFont('<sz val="11.5"/>')?.size).toBe(11.5)
  })

  it("ignores sz with no val", () => {
    expect(runFont("<sz/>")?.size).toBeUndefined()
  })

  it("reads the font name from <rFont>, not <name>", () => {
    // Runs spell the typeface `<rFont>`; `<name>` is the styles.xml form
    // and must not be honoured here.
    expect(runFont('<rFont val="Cambria"/>')?.name).toBe("Cambria")
    expect(runFont('<name val="Cambria"/>')?.name).toBeUndefined()
  })

  it("ignores rFont with no val", () => {
    expect(runFont("<rFont/>")?.name).toBeUndefined()
  })

  it("reads family and charset as numbers", () => {
    const font = runFont('<family val="2"/><charset val="204"/>')
    expect(font?.family).toBe(2)
    expect(font?.charset).toBe(204)
  })

  it("ignores family and charset with no val", () => {
    const font = runFont("<family/><charset/>")
    expect(font?.family).toBeUndefined()
    expect(font?.charset).toBeUndefined()
  })

  it("accepts only the enumerated vertAlign tokens", () => {
    expect(runFont('<vertAlign val="superscript"/>')?.vertAlign).toBe("superscript")
    expect(runFont('<vertAlign val="subscript"/>')?.vertAlign).toBe("subscript")
    // "baseline" is the default and carries no information; anything else
    // is invalid. Both are dropped rather than widening the union type.
    expect(runFont('<vertAlign val="baseline"/>')?.vertAlign).toBeUndefined()
    expect(runFont("<vertAlign/>")?.vertAlign).toBeUndefined()
  })

  it("accepts only the enumerated scheme tokens", () => {
    expect(runFont('<scheme val="major"/>')?.scheme).toBe("major")
    expect(runFont('<scheme val="minor"/>')?.scheme).toBe("minor")
    expect(runFont('<scheme val="none"/>')?.scheme).toBe("none")
    expect(runFont('<scheme val="bogus"/>')?.scheme).toBeUndefined()
  })

  it("ignores unknown <rPr> children", () => {
    expect(runFont('<outline/><shadow/><condense val="1"/>')).toEqual({})
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Run colours. `Color` is documented as "hex RGB without '#'", and the
// same `<color>` grammar appears in styles.xml, in inline strings, and
// here — all three must agree.
// ═══════════════════════════════════════════════════════════════════════

describe("parseSharedStrings — <rPr> colour", () => {
  it("strips the opaque FF alpha from an 8-digit ARGB value", () => {
    expect(runFont('<color rgb="FFFF0000"/>')?.color).toEqual({ rgb: "FF0000" })
  })

  it("reads theme, tint and indexed colours", () => {
    expect(runFont('<color theme="4" tint="-0.25"/>')?.color).toEqual({ theme: 4, tint: -0.25 })
    expect(runFont('<color indexed="64"/>')?.color).toEqual({ indexed: 64 })
  })

  it("reads theme 0 and a zero tint", () => {
    // theme="0" / tint="0" are falsy-looking strings; an `if (attr)` guard
    // that coerced to number first would drop them.
    expect(runFont('<color theme="0" tint="0"/>')?.color).toEqual({ theme: 0, tint: 0 })
  })

  it("emits an empty colour object for a valueless <color/>", () => {
    expect(runFont("<color/>")?.color).toEqual({})
  })

  // ── KNOWN BUG ──────────────────────────────────────────────────────
  // src/xlsx/shared-strings.ts:124 uses `rgb.replace(/^FF/, "")`, which
  // strips a *literal* leading "FF" instead of the ARGB alpha byte. Both
  // sibling implementations — src/xlsx/styles.ts:232 and the inline-string
  // path at src/xlsx/worksheet.ts:1686 — use
  // `rgb.length === 8 ? rgb.slice(2) : rgb`. Consequences below.
  it("must not mutate a 6-digit RGB value that has no alpha byte", () => {
    // rgb="FF0000" (plain red, written by several non-Excel producers)
    // currently becomes "0000".
    expect(runFont('<color rgb="FF0000"/>')?.color).toEqual({ rgb: "FF0000" })
  })

  it("must strip a non-opaque alpha byte like every other colour parser", () => {
    // rgb="80FF0000" currently stays 8 characters long, so `Color.rgb`
    // stops being a hex RGB triplet.
    expect(runFont('<color rgb="80FF0000"/>')?.color).toEqual({ rgb: "FF0000" })
  })

  it("must agree with the styles.xml colour parser on identical input", () => {
    const cases = ["FFFF0000", "80FF0000", "00FF00FF", "FF0000"]
    const styles = parseStyles(
      `<styleSheet xmlns="${SST_NS}"><fonts>${cases
        .map((rgb) => `<font><color rgb="${rgb}"/></font>`)
        .join("")}</fonts></styleSheet>`,
    )
    const fromRuns = cases.map((rgb) => runFont(`<color rgb="${rgb}"/>`)?.color?.rgb)
    expect(fromRuns).toEqual(styles.fonts.map((f) => f.color?.rgb))
  })
})
