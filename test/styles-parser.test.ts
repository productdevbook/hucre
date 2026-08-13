import { describe, expect, it } from "vitest"
import { isDateStyle, parseStyles, resolveStyle } from "../src/xlsx/styles"
import type { GradientFill, PatternFill } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

/** Wrap styles.xml fragments in a real `<styleSheet>` document. */
function styleSheet(inner: string): ReturnType<typeof parseStyles> {
  return parseStyles(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN_NS}">${inner}</styleSheet>`)
}

// The fragment helpers below indent their bodies on purpose: styles.xml
// from LibreOffice and most XML pretty-printers is whitespace-formatted, so
// every container parser has to walk past text children to find elements.

/** Parse a single `<font>` from its child elements. */
function font(body: string) {
  return styleSheet(`<fonts count="1">\n  <font>\n    ${body}\n  </font>\n</fonts>`).fonts[0]
}

/** Parse a single `<fill>` from its child elements. */
function fill(body: string) {
  return styleSheet(`<fills count="1">\n  <fill>\n    ${body}\n  </fill>\n</fills>`).fills[0]
}

/** Parse a single `<border>` from an attribute string plus children. */
function border(attrs: string, body: string) {
  return styleSheet(
    `<borders count="1">\n  <border ${attrs}>\n    ${body}\n  </border>\n</borders>`,
  ).borders[0]
}

/** Parse a single `<xf>` from an attribute string plus children. */
function xf(attrs: string, body = "") {
  return styleSheet(`<cellXfs count="1">\n  <xf ${attrs}>\n    ${body}\n  </xf>\n</cellXfs>`)
    .cellXfs[0]
}

// ═══════════════════════════════════════════════════════════════════════
// Document level. styles.xml carries a dozen sibling sections
// (cellStyleXfs, dxfs, tableStyles, colors, extLst …) that this parser
// deliberately does not model; they must be skipped, not mistaken for the
// sections it does read.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — document structure", () => {
  it("returns empty collections for a bare <styleSheet/>", () => {
    const styles = parseStyles(`<styleSheet xmlns="${MAIN_NS}"/>`)
    expect(styles.numFmts.size).toBe(0)
    expect(styles.fonts).toEqual([])
    expect(styles.fills).toEqual([])
    expect(styles.borders).toEqual([])
    expect(styles.cellXfs).toEqual([])
  })

  it("ignores sections it does not model and reads cellXfs, not cellStyleXfs", () => {
    // cellStyleXfs has an identical <xf> grammar. Reading it as well would
    // shift every `s="n"` cell reference by the named-style count.
    const styles = styleSheet(`
      <cellStyleXfs count="1"><xf numFmtId="14" fontId="9"/></cellStyleXfs>
      <cellXfs count="1"><xf numFmtId="3" fontId="0"/></cellXfs>
      <cellStyles count="1"><cellStyle name="Normal" xfId="0"/></cellStyles>
      <dxfs count="1"><dxf><font><b/></font></dxf></dxfs>
      <tableStyles count="0"/>
      <extLst><ext uri="{X}"/></extLst>
    `)
    expect(styles.cellXfs).toHaveLength(1)
    expect(styles.cellXfs[0].numFmtId).toBe(3)
    // The <font> nested in <dxfs> must not join the workbook font table.
    expect(styles.fonts).toEqual([])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <numFmts>. Entries are keyed by numeric id; a malformed id is worse than
// useless because it would collide with every other malformed entry.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — number formats", () => {
  it("indexes custom formats by numFmtId", () => {
    const styles = styleSheet(`<numFmts count="2">
      <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
      <numFmt numFmtId="165" formatCode="0.000&quot; kg&quot;"/>
    </numFmts>`)
    expect(styles.numFmts.get(164)).toBe("yyyy-mm-dd")
    expect(styles.numFmts.get(165)).toBe('0.000" kg"')
  })

  it("drops entries whose numFmtId is missing or not a number", () => {
    const styles = styleSheet(`<numFmts count="2">
      <numFmt formatCode="0.00"/>
      <numFmt numFmtId="abc" formatCode="0.00"/>
      <numFmt numFmtId="164" formatCode="0.00"/>
    </numFmts>`)
    expect([...styles.numFmts.keys()]).toEqual([164])
  })

  it("stores an empty formatCode when the attribute is absent", () => {
    const styles = styleSheet(`<numFmts count="1"><numFmt numFmtId="165"/></numFmts>`)
    expect(styles.numFmts.get(165)).toBe("")
  })

  it("lets a later duplicate id win", () => {
    const styles = styleSheet(`<numFmts count="2">
      <numFmt numFmtId="166" formatCode="first"/>
      <numFmt numFmtId="166" formatCode="second"/>
    </numFmts>`)
    expect(styles.numFmts.get(166)).toBe("second")
  })

  it("ignores non-<numFmt> children of <numFmts>", () => {
    const styles = styleSheet(`<numFmts count="1"><junk numFmtId="1"/></numFmts>`)
    expect(styles.numFmts.size).toBe(0)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <font>. Same accept-or-drop grammar as run properties, but the typeface
// element is <name> here rather than <rFont>.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — fonts", () => {
  it('treats bare toggles as on and val="0"/"false" as off', () => {
    expect(font("<b/><i/><strike/>")).toMatchObject({
      bold: true,
      italic: true,
      strikethrough: true,
    })
    expect(font('<b val="0"/><i val="false"/><strike val="0"/>')).toMatchObject({
      bold: false,
      italic: false,
      strikethrough: false,
    })
  })

  it("maps every underline token", () => {
    expect(font("<u/>").underline).toBe(true)
    expect(font('<u val="single"/>').underline).toBe(true)
    expect(font('<u val="double"/>').underline).toBe("double")
    expect(font('<u val="singleAccounting"/>').underline).toBe("singleAccounting")
    expect(font('<u val="doubleAccounting"/>').underline).toBe("doubleAccounting")
  })

  it("reads the typeface from <name>, not <rFont>", () => {
    expect(font('<name val="Calibri"/>').name).toBe("Calibri")
    expect(font('<rFont val="Calibri"/>').name).toBeUndefined()
  })

  it("ignores valueless scalar elements", () => {
    expect(font("<sz/><name/><family/><charset/>")).toEqual({})
  })

  it("reads sz, family and charset as numbers", () => {
    expect(font('<sz val="9.5"/><family val="2"/><charset val="238"/>')).toMatchObject({
      size: 9.5,
      family: 2,
      charset: 238,
    })
  })

  it("accepts only the enumerated vertAlign and scheme tokens", () => {
    expect(font('<vertAlign val="superscript"/>').vertAlign).toBe("superscript")
    expect(font('<vertAlign val="subscript"/>').vertAlign).toBe("subscript")
    expect(font('<vertAlign val="baseline"/>').vertAlign).toBeUndefined()
    expect(font('<scheme val="major"/>').scheme).toBe("major")
    expect(font('<scheme val="minor"/>').scheme).toBe("minor")
    expect(font('<scheme val="none"/>').scheme).toBe("none")
    expect(font('<scheme val="unexpected"/>').scheme).toBeUndefined()
  })

  it("ignores unmodelled font children", () => {
    expect(font("<outline/><shadow/><extend/><condense/>")).toEqual({})
  })

  it("ignores non-<font> children of <fonts>", () => {
    expect(styleSheet('<fonts count="1"><notAFont/></fonts>').fonts).toEqual([])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <color>. `Color.rgb` is documented as hex RGB with no '#', so the ARGB
// alpha byte is dropped by length, not by matching a particular alpha.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — colors", () => {
  it("strips the alpha byte from any 8-digit ARGB value", () => {
    expect(font('<color rgb="FFFF0000"/>').color).toEqual({ rgb: "FF0000" })
    // A non-opaque alpha must be stripped just the same — the rule is
    // positional, not "starts with FF".
    expect(font('<color rgb="80FF0000"/>').color).toEqual({ rgb: "FF0000" })
    expect(font('<color rgb="00FF00FF"/>').color).toEqual({ rgb: "FF00FF" })
  })

  it("passes a 6-digit RGB value through unchanged", () => {
    expect(font('<color rgb="FF0000"/>').color).toEqual({ rgb: "FF0000" })
  })

  it("reads theme, tint and indexed", () => {
    expect(font('<color theme="4" tint="-0.499984740745262"/>').color).toEqual({
      theme: 4,
      tint: -0.499984740745262,
    })
    expect(font('<color indexed="64"/>').color).toEqual({ indexed: 64 })
  })

  it("reads theme 0 and tint 0 rather than treating them as absent", () => {
    // The point is the zeros: none of them may be dropped as falsy.
    // `rgb` joins them because index 0 is black in the palette — see
    // test/indexed-color-palette.test.ts. The index itself still stands.
    expect(font('<color theme="0" tint="0" indexed="0"/>').color).toEqual({
      theme: 0,
      tint: 0,
      indexed: 0,
      rgb: "000000",
    })
  })

  it("returns an empty Color for a valueless <color/>", () => {
    expect(font("<color/>").color).toEqual({})
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <fill>. Index 0 (none) and 1 (gray125) are mandated by the spec, so real
// files always start with them; indices 2+ are the interesting ones.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — fills", () => {
  it("reads a solid pattern fill with fg and bg colors", () => {
    const f = fill(
      '<patternFill patternType="solid"><fgColor rgb="FFFFC000"/><bgColor indexed="64"/></patternFill>',
    ) as PatternFill
    expect(f).toEqual({
      type: "pattern",
      pattern: "solid",
      fgColor: { rgb: "FFC000" },
      bgColor: { indexed: 64 },
    })
  })

  it('defaults a patternFill with no patternType to "none"', () => {
    expect(fill("<patternFill/>")).toEqual({ type: "pattern", pattern: "none" })
  })

  it("falls back to a none pattern for a <fill> with no recognised child", () => {
    // Malformed or future fill types must still occupy their index so the
    // fillId → fills[] mapping stays correct.
    expect(fill("")).toEqual({ type: "pattern", pattern: "none" })
    expect(fill("<somethingElse/>")).toEqual({ type: "pattern", pattern: "none" })
  })

  it("reads a gradient fill with degree and stops", () => {
    const f = fill(`<gradientFill degree="90">
      <stop position="0"><color rgb="FFFFFFFF"/></stop>
      <stop position="1"><color theme="4"/></stop>
    </gradientFill>`) as GradientFill
    expect(f.type).toBe("gradient")
    expect(f.degree).toBe(90)
    expect(f.stops).toEqual([
      { position: 0, color: { rgb: "FFFFFF" } },
      { position: 1, color: { theme: 4 } },
    ])
  })

  it("leaves degree undefined when the attribute is absent", () => {
    const f = fill(
      '<gradientFill><stop position="0.5"><color indexed="1"/></stop></gradientFill>',
    ) as GradientFill
    expect(f.degree).toBeUndefined()
    expect(f.stops[0].position).toBe(0.5)
  })

  it("defaults a stop with no position to 0 and skips stops with no <color>", () => {
    const f = fill(
      '<gradientFill><stop/><stop><color rgb="FF00FF00"/></stop></gradientFill>',
    ) as GradientFill
    expect(f.stops).toEqual([{ position: 0, color: { rgb: "00FF00" } }])
  })

  it("takes the first recognised child when both fill types are present", () => {
    // Invalid per CT_Fill, but a reader should pick one deterministically.
    expect(fill('<patternFill patternType="solid"/><gradientFill degree="45"/>')).toEqual({
      type: "pattern",
      pattern: "solid",
    })
  })

  it("ignores unknown children inside patternFill and inside a stop", () => {
    expect(
      fill('<patternFill patternType="solid">\n  <unknownColor rgb="FF000000"/>\n</patternFill>'),
    ).toEqual({ type: "pattern", pattern: "solid" })
    const gradient = fill(
      '<gradientFill>\n  <stop position="0">\n    <notAColor/>\n    <color rgb="FF112233"/>\n  </stop>\n</gradientFill>',
    ) as GradientFill
    expect(gradient.stops).toEqual([{ position: 0, color: { rgb: "112233" } }])
  })

  it("ignores non-<fill> children of <fills>", () => {
    expect(styleSheet('<fills count="1"><notAFill/></fills>').fills).toEqual([])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <border>. A side element with no `style` attribute means "no border on
// that edge" — the single most common shape in real files, since Excel
// writes <left/><right/><top/><bottom/><diagonal/> for border 0.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — borders", () => {
  it("returns an empty BorderStyle for the default all-empty border", () => {
    expect(border("", "<left/><right/><top/><bottom/><diagonal/>")).toEqual({})
  })

  it("reads each side's style and color", () => {
    const b = border(
      "",
      `<left style="thin"><color indexed="64"/></left>
       <right style="medium"><color rgb="FF0000FF"/></right>
       <top style="dashed"/>
       <bottom style="double"><color theme="1"/></bottom>`,
    )
    expect(b.left).toEqual({ style: "thin", color: { indexed: 64 } })
    expect(b.right).toEqual({ style: "medium", color: { rgb: "0000FF" } })
    expect(b.top).toEqual({ style: "dashed" })
    expect(b.bottom).toEqual({ style: "double", color: { theme: 1 } })
  })

  it("reads the diagonal side together with its direction flags", () => {
    const b = border('diagonalUp="1" diagonalDown="1"', '<diagonal style="thin"/>')
    expect(b.diagonal).toEqual({ style: "thin" })
    expect(b.diagonalUp).toBe(true)
    expect(b.diagonalDown).toBe(true)
  })

  it('accepts "true" as well as "1" for the diagonal flags', () => {
    const b = border('diagonalUp="true" diagonalDown="true"', "")
    expect(b.diagonalUp).toBe(true)
    expect(b.diagonalDown).toBe(true)
  })

  it('omits the diagonal flags for "0"/"false" instead of storing false', () => {
    // Absent-means-false keeps the parsed object minimal so round-tripped
    // borders compare equal to freshly built ones.
    expect(border('diagonalUp="0" diagonalDown="false"', "")).toEqual({})
  })

  it("ignores side elements it does not model", () => {
    // <vertical>/<horizontal> only apply to dxf borders, not cell borders.
    expect(border("", '<vertical style="thin"/><horizontal style="thin"/>')).toEqual({})
  })

  it("ignores unknown children inside a side element", () => {
    expect(border("", '<top style="thin">\n  <notAColor val="1"/>\n</top>')).toEqual({
      top: { style: "thin" },
    })
  })

  it("ignores non-<border> children of <borders>", () => {
    expect(styleSheet('<borders count="1"><notABorder/></borders>').borders).toEqual([])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// <xf>. The apply* flags are booleans in "1"/"true" form; the id attributes
// all default to 0 when absent.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — cell XFs", () => {
  it("defaults every id to 0 when the attributes are absent", () => {
    expect(xf("")).toEqual({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })
  })

  it("reads all four id attributes", () => {
    expect(xf('numFmtId="164" fontId="2" fillId="3" borderId="4"')).toMatchObject({
      numFmtId: 164,
      fontId: 2,
      fillId: 3,
      borderId: 4,
    })
  })

  it('accepts "1" and "true" for every apply* flag', () => {
    const one = xf(
      'applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"',
    )
    const word = xf(
      'applyNumberFormat="true" applyFont="true" applyFill="true" applyBorder="true" applyAlignment="true" applyProtection="true"',
    )
    for (const parsed of [one, word]) {
      expect(parsed).toMatchObject({
        applyNumberFormat: true,
        applyFont: true,
        applyFill: true,
        applyBorder: true,
        applyAlignment: true,
        applyProtection: true,
      })
    }
  })

  it('leaves apply* flags undefined for "0"', () => {
    expect(xf('applyFont="0" applyFill="false"')).toEqual({
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: 0,
    })
  })

  it("reads every alignment attribute", () => {
    const parsed = xf(
      "",
      '<alignment horizontal="centerContinuous" vertical="distributed" wrapText="1" shrinkToFit="true" textRotation="45" indent="2"/>',
    )
    expect(parsed.alignment).toEqual({
      horizontal: "centerContinuous",
      vertical: "distributed",
      wrapText: true,
      shrinkToFit: true,
      textRotation: 45,
      indent: 2,
    })
  })

  it("maps readingOrder 1/2 to ltr/rtl and anything else to context", () => {
    expect(xf("", '<alignment readingOrder="1"/>').alignment?.readingOrder).toBe("ltr")
    expect(xf("", '<alignment readingOrder="2"/>').alignment?.readingOrder).toBe("rtl")
    // 0 is the spec's "context dependent" value.
    expect(xf("", '<alignment readingOrder="0"/>').alignment?.readingOrder).toBe("context")
  })

  it("omits alignment flags that are off", () => {
    expect(xf("", '<alignment wrapText="0" shrinkToFit="false"/>').alignment).toEqual({})
  })

  it("reads locked and hidden protection in both true and false form", () => {
    expect(xf("", '<protection locked="1" hidden="true"/>').protection).toEqual({
      locked: true,
      hidden: true,
    })
    // Explicit `false` must be preserved: unlocking a cell is the whole
    // point of the element on a protected sheet.
    expect(xf("", '<protection locked="0" hidden="false"/>').protection).toEqual({
      locked: false,
      hidden: false,
    })
  })

  it("omits protection keys that are not present", () => {
    expect(xf("", '<protection locked="1"/>').protection).toEqual({ locked: true })
    expect(xf("", "<protection/>").protection).toEqual({})
  })

  it("detects the Excel 2024 checkbox extension by its feature URI", () => {
    const parsed = xf(
      "",
      '<extLst><ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><xfpb:xfComplement i="0"/></ext></extLst>',
    )
    expect(parsed.hasCheckboxFeature).toBe(true)
  })

  it("ignores an extLst carrying some other extension", () => {
    const parsed = xf("", '<extLst><ext uri="{SOMETHING-ELSE}"><x/></ext></extLst>')
    expect(parsed.hasCheckboxFeature).toBeUndefined()
  })

  it("ignores an empty extLst and non-<ext> children", () => {
    expect(xf("", "<extLst/>").hasCheckboxFeature).toBeUndefined()
    expect(xf("", "<extLst>\n  <notAnExt/>\n</extLst>").hasCheckboxFeature).toBeUndefined()
  })

  it("finds the checkbox ext among several indented siblings", () => {
    const parsed = xf(
      "",
      `<extLst>
        <ext uri="{OTHER-EXTENSION}"><x/></ext>
        <ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}"><y/></ext>
      </extLst>`,
    )
    expect(parsed.hasCheckboxFeature).toBe(true)
  })

  it("ignores unmodelled <xf> children", () => {
    expect(xf("", "<somethingNew/>")).toEqual({
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: 0,
    })
  })

  it("ignores non-<xf> children of <cellXfs>", () => {
    expect(styleSheet('<cellXfs count="1"><notAnXf/></cellXfs>').cellXfs).toEqual([])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// resolveStyle — turning an `s="n"` index into a CellStyle. It emits only
// properties that actually deviate from the workbook default, which is what
// keeps a plain round-trip from decorating every cell.
// ═══════════════════════════════════════════════════════════════════════

describe("resolveStyle", () => {
  const styles = styleSheet(`
    <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
    <fonts count="2"><font><sz val="11"/></font><font><b/></font></fonts>
    <fills count="3">
      <fill><patternFill patternType="none"/></fill>
      <fill><patternFill patternType="gray125"/></fill>
      <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
    </fills>
    <borders count="2">
      <border><left/><right/><top/><bottom/><diagonal/></border>
      <border><left style="thin"/></border>
    </borders>
    <cellXfs count="9">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
      <xf numFmtId="164" fontId="1" fillId="2" borderId="1"/>
      <xf numFmtId="9" fontId="0" fillId="1" borderId="0"/>
      <xf numFmtId="999" fontId="0" fillId="0" borderId="0"/>
      <xf numFmtId="0" fontId="99" fillId="0" borderId="0"/>
      <xf numFmtId="0" fontId="0" fillId="99" borderId="0"/>
      <xf numFmtId="0" fontId="0" fillId="0" borderId="99"/>
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment horizontal="right"/><protection hidden="1"/></xf>
      <xf numFmtId="0" fontId="1" fillId="2" borderId="1"/>
    </cellXfs>
  `)

  it("returns an empty style for the default xf", () => {
    // numFmt 0 (General), font 0, fill 0 and border 0 are all the workbook
    // defaults, so a plain cell carries no style at all.
    expect(resolveStyle(styles, 0)).toEqual({})
  })

  it("resolves custom format, font, fill and border together", () => {
    const style = resolveStyle(styles, 1)
    expect(style.numFmt).toBe("0.000")
    expect(style.font).toEqual({ bold: true })
    expect(style.fill).toMatchObject({ type: "pattern", pattern: "solid" })
    expect(style.border).toEqual({ left: { style: "thin" } })
  })

  it("resolves a built-in numFmtId through the built-in table", () => {
    expect(resolveStyle(styles, 2).numFmt).toBe("0%")
  })

  it("skips fill index 1 — the mandatory gray125 placeholder", () => {
    // Every styles.xml declares gray125 at index 1 whether or not anything
    // uses it; surfacing it would paint cells with a grey hatch.
    expect(resolveStyle(styles, 2).fill).toBeUndefined()
  })

  it("omits numFmt for an id that exists in neither table", () => {
    expect(resolveStyle(styles, 3).numFmt).toBeUndefined()
  })

  it("ignores font, fill and border ids past the end of their tables", () => {
    // Truncated or hand-edited files point past the table; reading
    // undefined into CellStyle would crash consumers instead.
    expect(resolveStyle(styles, 4).font).toBeUndefined()
    expect(resolveStyle(styles, 5).fill).toBeUndefined()
    expect(resolveStyle(styles, 6).border).toBeUndefined()
  })

  it("passes alignment and protection through unchanged", () => {
    const style = resolveStyle(styles, 7)
    expect(style.alignment).toEqual({ horizontal: "right" })
    expect(style.protection).toEqual({ hidden: true })
  })

  it("returns an empty style for an out-of-range or negative index", () => {
    expect(resolveStyle(styles, 99)).toEqual({})
    expect(resolveStyle(styles, -1)).toEqual({})
  })

  it("shares the parsed font/fill/border instances between cells", () => {
    // Deliberate: the tables are parsed once and referenced, not cloned.
    // Callers that mutate a resolved style would affect every other cell
    // using the same xf, so this identity is part of the contract.
    const a = resolveStyle(styles, 1)
    const b = resolveStyle(styles, 8)
    expect(a.font).toBe(b.font)
    expect(a.fill).toBe(b.fill)
    expect(a.border).toBe(b.border)
    expect(a.font).toBe(styles.fonts[1])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// isDateStyle — decides whether a numeric cell value is a serial date.
// A false positive turns a quantity into a 1900-era timestamp, so the
// boundaries matter more than most.
// ═══════════════════════════════════════════════════════════════════════

describe("isDateStyle", () => {
  it("returns false for an unknown style index", () => {
    const styles = styleSheet('<cellXfs count="1"><xf numFmtId="14"/></cellXfs>')
    expect(isDateStyle(styles, 5)).toBe(false)
    expect(isDateStyle(styles, -1)).toBe(false)
  })

  it("recognises the locale-specific built-in date ids 27-36 and 50-58", () => {
    // These ids have no entry in the built-in format table — only the id
    // set identifies them, and East-Asian locale files rely on it.
    const ids = [27, 30, 36, 50, 55, 58]
    const styles = styleSheet(
      `<cellXfs count="${ids.length}">${ids
        .map((id) => `<xf numFmtId="${id}"/>`)
        .join("")}</cellXfs>`,
    )
    ids.forEach((_id, i) => expect(isDateStyle(styles, i)).toBe(true))
  })

  it("recognises the elapsed-time built-ins 45, 46 and 47", () => {
    const ids = [45, 46, 47]
    const styles = styleSheet(
      `<cellXfs count="3">${ids.map((id) => `<xf numFmtId="${id}"/>`).join("")}</cellXfs>`,
    )
    ids.forEach((_id, i) => expect(isDateStyle(styles, i)).toBe(true))
  })

  it("rejects the numeric built-ins that sit between the date ranges", () => {
    const ids = [23, 24, 25, 26, 37, 38, 39, 40, 48, 49]
    const styles = styleSheet(
      `<cellXfs count="${ids.length}">${ids
        .map((id) => `<xf numFmtId="${id}"/>`)
        .join("")}</cellXfs>`,
    )
    ids.forEach((_id, i) => expect(isDateStyle(styles, i)).toBe(false))
  })

  it("analyses a custom format string when the id is not a known date id", () => {
    const styles = styleSheet(`
      <numFmts count="4">
        <numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>
        <numFmt numFmtId="165" formatCode="[$-409]h:mm:ss AM/PM"/>
        <numFmt numFmtId="166" formatCode="#,##0.00&quot; kg&quot;"/>
        <numFmt numFmtId="167" formatCode="0.00%"/>
      </numFmts>
      <cellXfs count="4">
        <xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="166"/><xf numFmtId="167"/>
      </cellXfs>
    `)
    expect(isDateStyle(styles, 0)).toBe(true)
    expect(isDateStyle(styles, 1)).toBe(true)
    expect(isDateStyle(styles, 2)).toBe(false)
    expect(isDateStyle(styles, 3)).toBe(false)
  })

  it("returns false for a custom id with no format code at all", () => {
    const styles = styleSheet(`
      <numFmts count="1"><numFmt numFmtId="170"/></numFmts>
      <cellXfs count="2"><xf numFmtId="170"/><xf numFmtId="171"/></cellXfs>
    `)
    expect(isDateStyle(styles, 0)).toBe(false)
    expect(isDateStyle(styles, 1)).toBe(false)
  })

  // ── KNOWN BUG ──────────────────────────────────────────────────────
  // ECMA-376 §18.8.30 allows a <numFmt> entry to redefine a built-in id.
  // resolveStyle (src/xlsx/styles.ts:502) gives the custom formatCode
  // priority, but isDateStyle (src/xlsx/styles.ts:547) checks DATE_FMT_IDS
  // *before* consulting numFmts, so the two disagree: a cell redefining id
  // 14 as "#,##0" is formatted as a number yet still converted to a Date by
  // the reader. Fix: look up styles.numFmts first in isDateStyle too.
  it("must let a redefined built-in id override the date-id table", () => {
    const styles = styleSheet(`
      <numFmts count="2">
        <numFmt numFmtId="14" formatCode="#,##0"/>
        <numFmt numFmtId="3" formatCode="yyyy-mm-dd"/>
      </numFmts>
      <cellXfs count="2"><xf numFmtId="14"/><xf numFmtId="3"/></cellXfs>
    `)
    // resolveStyle already honours the redefinition …
    expect(resolveStyle(styles, 0).numFmt).toBe("#,##0")
    expect(resolveStyle(styles, 1).numFmt).toBe("yyyy-mm-dd")
    // … isDateStyle must agree.
    expect(isDateStyle(styles, 0)).toBe(false)
    expect(isDateStyle(styles, 1)).toBe(true)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// A complete styles.xml as Excel writes it, parsed end to end. Guards the
// section ordering and the index arithmetic that the fragment tests above
// deliberately isolate away.
// ═══════════════════════════════════════════════════════════════════════

describe("parseStyles — realistic workbook stylesheet", () => {
  const XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN_NS}">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts>
  <fonts count="3">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
    <font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
    <font><i/><u/><sz val="10"/><color rgb="FF7F7F7F"/><name val="Arial"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="14" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`

  it("parses every section with the right cardinality", () => {
    const styles = parseStyles(XML)
    expect(styles.numFmts.size).toBe(1)
    expect(styles.fonts).toHaveLength(3)
    expect(styles.fills).toHaveLength(3)
    expect(styles.borders).toHaveLength(2)
    expect(styles.cellXfs).toHaveLength(4)
  })

  it("resolves the header style (xf 1) end to end", () => {
    const style = resolveStyle(parseStyles(XML), 1)
    expect(style.font).toEqual({
      bold: true,
      size: 11,
      color: { theme: 0 },
      name: "Calibri",
      family: 2,
      scheme: "minor",
    })
    expect(style.fill).toEqual({
      type: "pattern",
      pattern: "solid",
      fgColor: { rgb: "4472C4" },
      bgColor: { indexed: 64 },
    })
    expect(style.border?.top).toEqual({ style: "thin", color: { indexed: 64 } })
    expect(style.alignment).toEqual({ horizontal: "center", vertical: "center", wrapText: true })
    expect(style.numFmt).toBeUndefined()
  })

  it("resolves the currency style (xf 2) without treating it as a date", () => {
    const styles = parseStyles(XML)
    expect(resolveStyle(styles, 2).numFmt).toBe('"$"#,##0.00')
    expect(isDateStyle(styles, 2)).toBe(false)
  })

  it("resolves the date style (xf 3) from the built-in table", () => {
    const styles = parseStyles(XML)
    expect(resolveStyle(styles, 3).numFmt).toBe("m/d/yyyy")
    expect(isDateStyle(styles, 3)).toBe(true)
  })
})
