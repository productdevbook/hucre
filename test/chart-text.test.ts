import { describe, expect, it } from "vitest"
import { parseXml } from "../src/xml/parser"
import {
  FONT_SIZE_MAX_PT,
  FONT_SIZE_MIN_PT,
  FONT_SZ_PER_POINT,
  ROTATION_MAX_DEG,
  ROTATION_MIN_DEG,
  TXPR_ROT_PER_DEGREE,
  makeFontSizeNormalizer,
  normalizeBoolFlag,
  normalizeFontFamily,
  normalizeFontSizePt,
  parseFontFamilyFromDefRPr,
  parseFontSizeFromDefRPr,
  readBoolAttrValue,
  readStrikeToken,
  readUnderlineToken,
} from "../src/xlsx/chart/text"

// ── Helpers ──────────────────────────────────────────────────────────

/** Build a real `<a:defRPr>` through the project's parser, not a literal. */
function defRPr(inner: string): ReturnType<typeof parseXml> {
  return parseXml(
    `<a:defRPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"${inner}`,
  )
}

/** `<a:defRPr sz="..."/>` with the attribute spelled exactly as given. */
function withSz(raw: string): ReturnType<typeof parseXml> {
  return defRPr(` sz="${raw}"/>`)
}

// ═══════════════════════════════════════════════════════════════════════
// Constants
// ═══════════════════════════════════════════════════════════════════════

describe("chart/text constants", () => {
  it("matches the OOXML units the JSDoc documents", () => {
    expect(FONT_SZ_PER_POINT).toBe(100)
    expect(TXPR_ROT_PER_DEGREE).toBe(60000)
    expect(FONT_SIZE_MIN_PT).toBe(1)
    expect(FONT_SIZE_MAX_PT).toBe(400)
    expect(ROTATION_MIN_DEG).toBe(-90)
    expect(ROTATION_MAX_DEG).toBe(90)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// parseFontSizeFromDefRPr
// ═══════════════════════════════════════════════════════════════════════

describe("parseFontSizeFromDefRPr", () => {
  it("converts 100ths of a point to points", () => {
    expect(parseFontSizeFromDefRPr(withSz("1200"))).toBe(12)
    expect(parseFontSizeFromDefRPr(withSz("1800"))).toBe(18)
  })

  it("returns undefined when sz is absent", () => {
    expect(parseFontSizeFromDefRPr(defRPr("/>"))).toBeUndefined()
  })

  it("returns undefined for an empty or whitespace-only sz", () => {
    expect(parseFontSizeFromDefRPr(withSz(""))).toBeUndefined()
    expect(parseFontSizeFromDefRPr(withSz("   "))).toBeUndefined()
  })

  it("returns undefined for a non-numeric sz", () => {
    expect(parseFontSizeFromDefRPr(withSz("abc"))).toBeUndefined()
    expect(parseFontSizeFromDefRPr(withSz("+"))).toBeUndefined()
  })

  it("snaps to the half-point grid Excel's UI exposes", () => {
    // 12.25pt is not on the grid — rounds to 12.5.
    expect(parseFontSizeFromDefRPr(withSz("1225"))).toBe(12.5)
    // 12.24pt rounds down to 12.
    expect(parseFontSizeFromDefRPr(withSz("1224"))).toBe(12)
    expect(parseFontSizeFromDefRPr(withSz("1250"))).toBe(12.5)
  })

  it("accepts the inclusive band boundaries", () => {
    expect(parseFontSizeFromDefRPr(withSz("100"))).toBe(FONT_SIZE_MIN_PT)
    expect(parseFontSizeFromDefRPr(withSz("40000"))).toBe(FONT_SIZE_MAX_PT)
  })

  it("drops values outside the band rather than clamping", () => {
    // A reader that clamped here would surface a size the writer never emits.
    expect(parseFontSizeFromDefRPr(withSz("50"))).toBeUndefined()
    expect(parseFontSizeFromDefRPr(withSz("40100"))).toBeUndefined()
    expect(parseFontSizeFromDefRPr(withSz("-1200"))).toBeUndefined()
    expect(parseFontSizeFromDefRPr(withSz("0"))).toBeUndefined()
  })

  it("tolerates trailing junk the way parseInt does", () => {
    expect(parseFontSizeFromDefRPr(withSz("1200pt"))).toBe(12)
    // Fractional input truncates before the unit conversion.
    expect(parseFontSizeFromDefRPr(withSz("1200.9"))).toBe(12)
  })

  it("tolerates surrounding whitespace", () => {
    expect(parseFontSizeFromDefRPr(withSz("  1400  "))).toBe(14)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// makeFontSizeNormalizer / normalizeFontSizePt
// ═══════════════════════════════════════════════════════════════════════

describe("makeFontSizeNormalizer", () => {
  const normalize = makeFontSizeNormalizer(8, 20)

  it("snaps to the half-point grid", () => {
    expect(normalize(12)).toBe(12)
    expect(normalize(12.25)).toBe(12.5)
    expect(normalize(12.24)).toBe(12)
    expect(normalize(12.75)).toBe(13)
  })

  it("clamps to the supplied band instead of dropping", () => {
    // Unlike the parser, the writer-side normalizer clamps — an
    // out-of-band authored value still produces a valid file.
    expect(normalize(1)).toBe(8)
    expect(normalize(999)).toBe(20)
  })

  it("passes the band boundaries through", () => {
    expect(normalize(8)).toBe(8)
    expect(normalize(20)).toBe(20)
  })

  it("returns undefined for non-finite and non-numeric input", () => {
    expect(normalize(undefined)).toBeUndefined()
    expect(normalize(Number.NaN)).toBeUndefined()
    expect(normalize(Number.POSITIVE_INFINITY)).toBeUndefined()
    expect(normalize(Number.NEGATIVE_INFINITY)).toBeUndefined()
    expect(normalize("12" as unknown as number)).toBeUndefined()
    expect(normalize(null as unknown as number)).toBeUndefined()
  })

  it("builds independent normalizers", () => {
    const wide = makeFontSizeNormalizer(1, 400)
    expect(wide(1)).toBe(1)
    expect(normalize(1)).toBe(8)
  })
})

describe("normalizeFontSizePt", () => {
  it("uses the standard 1..400 pt band", () => {
    expect(normalizeFontSizePt(0.4)).toBe(FONT_SIZE_MIN_PT)
    expect(normalizeFontSizePt(1000)).toBe(FONT_SIZE_MAX_PT)
    expect(normalizeFontSizePt(11)).toBe(11)
    expect(normalizeFontSizePt(undefined)).toBeUndefined()
  })

  it("agrees with the parser on values that round-trip", () => {
    for (const raw of ["100", "1200", "1250", "40000"]) {
      const parsed = parseFontSizeFromDefRPr(withSz(raw))
      expect(normalizeFontSizePt(parsed)).toBe(parsed)
    }
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Font family
// ═══════════════════════════════════════════════════════════════════════

describe("parseFontFamilyFromDefRPr", () => {
  it("reads the typeface off <a:latin>", () => {
    expect(parseFontFamilyFromDefRPr(defRPr('><a:latin typeface="Calibri"/></a:defRPr>'))).toBe(
      "Calibri",
    )
  })

  it("trims surrounding whitespace", () => {
    expect(parseFontFamilyFromDefRPr(defRPr('><a:latin typeface="  Arial  "/></a:defRPr>'))).toBe(
      "Arial",
    )
  })

  it("returns undefined when <a:latin> is absent", () => {
    expect(parseFontFamilyFromDefRPr(defRPr("/>"))).toBeUndefined()
    expect(
      parseFontFamilyFromDefRPr(
        defRPr('><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr>'),
      ),
    ).toBeUndefined()
  })

  it("returns undefined when typeface is missing or blank", () => {
    expect(parseFontFamilyFromDefRPr(defRPr("><a:latin/></a:defRPr>"))).toBeUndefined()
    expect(parseFontFamilyFromDefRPr(defRPr('><a:latin typeface=""/></a:defRPr>'))).toBeUndefined()
    expect(
      parseFontFamilyFromDefRPr(defRPr('><a:latin typeface="   "/></a:defRPr>')),
    ).toBeUndefined()
  })

  it("ignores a non-latin child that carries a typeface", () => {
    expect(
      parseFontFamilyFromDefRPr(defRPr('><a:ea typeface="MS Gothic"/></a:defRPr>')),
    ).toBeUndefined()
  })
})

describe("normalizeFontFamily", () => {
  it("trims and passes through", () => {
    expect(normalizeFontFamily("Calibri")).toBe("Calibri")
    expect(normalizeFontFamily("  Arial  ")).toBe("Arial")
  })

  it("drops blank and non-string tokens", () => {
    expect(normalizeFontFamily(undefined)).toBeUndefined()
    expect(normalizeFontFamily("")).toBeUndefined()
    expect(normalizeFontFamily("   ")).toBeUndefined()
    expect(normalizeFontFamily(12 as unknown as string)).toBeUndefined()
    expect(normalizeFontFamily(null as unknown as string)).toBeUndefined()
  })

  it("agrees with the parser, so a parsed family survives a rewrite", () => {
    const parsed = parseFontFamilyFromDefRPr(defRPr('><a:latin typeface=" Verdana "/></a:defRPr>'))
    expect(normalizeFontFamily(parsed)).toBe("Verdana")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Boolean-style attributes
// ═══════════════════════════════════════════════════════════════════════

describe("readBoolAttrValue", () => {
  it("accepts both OOXML spellings", () => {
    expect(readBoolAttrValue("1")).toBe(true)
    expect(readBoolAttrValue("true")).toBe(true)
    expect(readBoolAttrValue("0")).toBe(false)
    expect(readBoolAttrValue("false")).toBe(false)
  })

  it("returns undefined for anything else", () => {
    expect(readBoolAttrValue(undefined)).toBeUndefined()
    expect(readBoolAttrValue("")).toBeUndefined()
    expect(readBoolAttrValue("TRUE")).toBeUndefined()
    expect(readBoolAttrValue("False")).toBeUndefined()
    expect(readBoolAttrValue(" 1 ")).toBeUndefined()
    expect(readBoolAttrValue("yes")).toBeUndefined()
    expect(readBoolAttrValue("2")).toBeUndefined()
  })
})

describe("normalizeBoolFlag", () => {
  it("passes literal booleans through", () => {
    expect(normalizeBoolFlag(true)).toBe(true)
    expect(normalizeBoolFlag(false)).toBe(false)
  })

  it("drops everything else, including truthy escapes from untyped callers", () => {
    expect(normalizeBoolFlag(undefined)).toBeUndefined()
    expect(normalizeBoolFlag(null as unknown as boolean)).toBeUndefined()
    expect(normalizeBoolFlag(1 as unknown as boolean)).toBeUndefined()
    expect(normalizeBoolFlag(0 as unknown as boolean)).toBeUndefined()
    expect(normalizeBoolFlag("true" as unknown as boolean)).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Underline / strike
// ═══════════════════════════════════════════════════════════════════════

describe("readUnderlineToken", () => {
  it("maps the two tokens Excel's UI authors", () => {
    expect(readUnderlineToken("sng")).toBe(true)
    expect(readUnderlineToken("none")).toBe(false)
  })

  it("drops the vocabulary Excel never authors", () => {
    // Dropping rather than coercing is what makes absence and the OOXML
    // default round-trip identically through cloneChart.
    for (const token of ["dbl", "words", "heavy", "dotted", "wavy", "sngStrike", ""]) {
      expect(readUnderlineToken(token), token).toBeUndefined()
    }
    expect(readUnderlineToken(undefined)).toBeUndefined()
  })
})

describe("readStrikeToken", () => {
  it("maps the two tokens Excel's UI authors", () => {
    expect(readStrikeToken("sngStrike")).toBe(true)
    expect(readStrikeToken("noStrike")).toBe(false)
  })

  it("drops dblStrike and every other token", () => {
    for (const token of ["dblStrike", "sng", "none", "", "strike"]) {
      expect(readStrikeToken(token), token).toBeUndefined()
    }
    expect(readStrikeToken(undefined)).toBeUndefined()
  })

  it("does not share a vocabulary with the underline reader", () => {
    expect(readUnderlineToken("sngStrike")).toBeUndefined()
    expect(readStrikeToken("sng")).toBeUndefined()
  })
})
