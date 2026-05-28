// ── Chart Text ─────────────────────────────────────────────────────
// Generic helpers for typography — font size, font family, bold,
// italic, underline, strikethrough, color, rotation — that every
// chart-frame text host (chart-title, axis-title, axis tick labels,
// legend, data-labels, data-table) shares.
//
// Excel and OOXML store typography in two equivalent forms:
//
//   - The chart-title and axis-title hosts wrap the text in
//     `<c:tx><c:rich><a:p><a:pPr><a:defRPr ...>` (CT_Title /
//     CT_TextBody, ECMA-376 Part 1, §21.2.2.210 / §20.1.2.2.34) — the
//     "rich-text body".
//   - The legend, axis tick labels, data-labels and data-table hosts
//     wrap typography in `<c:txPr><a:p><a:pPr><a:defRPr ...>` (the
//     "text-properties body").
//
// The two paths differ only in their outermost wrapper (`<c:tx><c:rich>`
// vs `<c:txPr>`). Once you reach `<a:p><a:pPr><a:defRPr>` the schema
// is identical — same `sz` / `b` / `i` / `u` / `strike` / `<a:solidFill>`
// / `<a:latin>` children. The helpers in this module take the *outermost*
// host element and the kind ("rich" or "txPr") and walk the canonical
// chain, surfacing the parsed value or `undefined` for the standard
// accept-or-drop grammar.
//
// JSDoc preserved verbatim from the per-host parsers when each parser
// was lifted out — the per-host commentary stays attached to its
// caller in chart-reader / chart-writer / chart-clone since the
// host-specific schema notes (where the host sits inside `<c:chart>`,
// what siblings can leak in, what JSDoc points back through `@link`)
// remain meaningful at the call site.

import type { XmlElement } from "../../xml/parser"

/** See `chart/shape.ts` for the equivalent helper. */
function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c
  }
  return undefined
}

// ── Rotation constants ────────────────────────────────────────────

/**
 * Conversion factor between OOXML's `rot` attribute (60000ths of a
 * degree, the integer Excel writes inside `<a:bodyPr rot="N"/>`) and
 * whole degrees. Excel's UI exposes the -90..90 degree band — the
 * reader clamps anything outside that band so a corrupt template
 * cannot surface a value the writer would never emit.
 */
export const TXPR_ROT_PER_DEGREE = 60000
/** Lower bound of the rotation band Excel's UI exposes. */
export const ROTATION_MIN_DEG = -90
/** Upper bound of the rotation band Excel's UI exposes. */
export const ROTATION_MAX_DEG = 90

// ── Font size constants ───────────────────────────────────────────

/**
 * Conversion factor between OOXML's `sz` attribute (100ths of a point,
 * the integer Excel writes inside `<a:defRPr sz="N"/>` /
 * `<a:rPr sz="N"/>`) and whole / half points. The OOXML
 * `ST_TextFontSize` schema restricts `sz` to the inclusive
 * `100..400000` band — the writer's clamp uses the same range
 * converted to points (`1..400`), so any out-of-range value collapses
 * to `undefined` rather than surface a token Excel would never emit.
 */
export const FONT_SZ_PER_POINT = 100
/** Lower bound of the font-size band Excel's UI exposes. */
export const FONT_SIZE_MIN_PT = 1
/** Upper bound of the font-size band Excel's UI exposes. */
export const FONT_SIZE_MAX_PT = 400

// ── Font size normalize / parse / clamp ───────────────────────────

/**
 * Parse the `sz` attribute on a `<a:defRPr>` (or `<a:rPr>`) element
 * and convert from OOXML's 100ths-of-a-point to half-point precision.
 *
 * Returns the value in whole / half points (range `FONT_SIZE_MIN_PT..
 * FONT_SIZE_MAX_PT`). Returns `undefined` when the attribute is
 * missing, when the value cannot be parsed as a finite integer, or
 * when the computed point value falls outside the supported band.
 *
 * Mirrors the writer's `clampFontSizePt` — a parsed-then-written value
 * does not drift across round-trips because the writer snaps to the
 * same half-point grid Excel's UI exposes.
 */
export function parseFontSizeFromDefRPr(defRPr: XmlElement): number | undefined {
  const raw = defRPr.attrs.sz
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  const parsed = Number.parseInt(trimmed, 10)
  if (!Number.isFinite(parsed)) return undefined
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes.
  const halfSteps = Math.round((parsed / FONT_SZ_PER_POINT) * 2)
  const points = halfSteps / 2
  if (points < FONT_SIZE_MIN_PT || points > FONT_SIZE_MAX_PT) return undefined
  return points
}

/**
 * Build a font-size normalizer that snaps an input value to the
 * 0.5 pt grid Excel's UI exposes and clamps to the supplied min / max
 * band. Non-finite / non-numeric tokens collapse to `undefined` so
 * the writer skips the `sz` attribute entirely.
 *
 * Used by both the per-host writer normalizers (`normalizeTitleFontSize`,
 * `normalizeLegendFontSize`, `normalizeAxisLabelFontSize`, ...) and the
 * cloner's resolvers — every typography slot shares one snap / clamp
 * grammar.
 */
export function makeFontSizeNormalizer(
  minPt: number,
  maxPt: number,
): (value: number | undefined) => number | undefined {
  return (value) => {
    if (typeof value !== "number" || !Number.isFinite(value)) return undefined
    const halfSteps = Math.round(value * 2)
    const snapped = halfSteps / 2
    if (snapped < minPt) return minPt
    if (snapped > maxPt) return maxPt
    return snapped
  }
}

/** Normalize a font-size value against the standard `1..400` pt band. */
export const normalizeFontSizePt: (value: number | undefined) => number | undefined =
  makeFontSizeNormalizer(FONT_SIZE_MIN_PT, FONT_SIZE_MAX_PT)

// ── Font family ───────────────────────────────────────────────────

/**
 * Pull `<a:latin typeface=".."/>` off a `<a:defRPr>` element and return
 * the typeface name as a trimmed string. Returns `undefined` when the
 * `<a:latin>` child is absent, when the `typeface` attribute is missing,
 * or when it is empty / whitespace-only — so a chart that pinned an
 * empty typeface drops the field rather than carry a value the writer
 * would silently elide back to absence.
 */
export function parseFontFamilyFromDefRPr(defRPr: XmlElement): string | undefined {
  const latin = findChild(defRPr, "latin")
  if (!latin) return undefined
  const typeface = latin.attrs.typeface
  if (typeof typeface !== "string") return undefined
  const trimmed = typeface.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

/**
 * Normalize a font-family value for the writer's `<a:latin typeface="">`
 * slot. Returns the trimmed string, or `undefined` for non-string
 * tokens and empty / whitespace-only values so the writer skips the
 * `<a:latin>` element entirely.
 */
export function normalizeFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

// ── Bold / italic ─────────────────────────────────────────────────

/**
 * Read a boolean-style OOXML attribute (`"1"` / `"0"` / `"true"` /
 * `"false"`). Returns `true` for the truthy tokens, `false` for
 * `"0"` / `"false"`, and `undefined` for any other string and for
 * missing / non-string values.
 *
 * Used to parse the `b` / `i` flags on `<a:defRPr>` and the canonical
 * `val` attribute on numeric / scale child elements (`<c:smooth>`,
 * `<c:overlay>`, `<c:auto>`, etc.).
 */
export function readBoolAttrValue(raw: string | undefined): boolean | undefined {
  if (raw === undefined) return undefined
  if (raw === "1" || raw === "true") return true
  if (raw === "0" || raw === "false") return false
  return undefined
}

/**
 * Normalize a literal boolean for a writer's flag attribute. Returns
 * the input when it is exactly `true` or `false`, or `undefined` for
 * any other token (including `null`-shaped escapes from an untyped
 * caller).
 */
export function normalizeBoolFlag(value: boolean | undefined): boolean | undefined {
  if (value === true) return true
  if (value === false) return false
  return undefined
}

// ── Underline / strike helpers ────────────────────────────────────
//
// The OOXML `u` (underline) attribute on `<a:defRPr>` accepts a wide
// vocabulary (`"none"`, `"sng"`, `"dbl"`, `"words"`, etc.) but Excel's
// UI only authors `"sng"` (single underline) and the OOXML default
// `"none"` collapses to absence on the wire. Same story for `strike`:
// `"sngStrike"` is the only value Excel authors; `"noStrike"` (the
// OOXML default) and `"dblStrike"` (non-UI) both collapse to absence.

/**
 * Map the OOXML `u` attribute value to the boolean Excel's UI exposes.
 * Returns `true` for `"sng"`, `false` for `"none"`, and `undefined` for
 * every other token — including `"dbl"`, `"words"`, `"heavy"`,
 * `"dotted"`, etc. — so absence and the OOXML default round-trip
 * identically through `cloneChart`.
 */
export function readUnderlineToken(raw: string | undefined): boolean | undefined {
  if (raw === undefined) return undefined
  if (raw === "sng") return true
  if (raw === "none") return false
  return undefined
}

/**
 * Map the OOXML `strike` attribute value to the boolean Excel's UI
 * exposes. Returns `true` for `"sngStrike"`, `false` for `"noStrike"`,
 * and `undefined` for every other token — including the non-UI
 * `"dblStrike"` — so absence and the OOXML default round-trip
 * identically through `cloneChart`.
 */
export function readStrikeToken(raw: string | undefined): boolean | undefined {
  if (raw === undefined) return undefined
  if (raw === "sngStrike") return true
  if (raw === "noStrike") return false
  return undefined
}
