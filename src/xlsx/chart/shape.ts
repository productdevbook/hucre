// ── Chart Shape ────────────────────────────────────────────────────
// Generic helpers for the OOXML `<c:spPr>` (CT_ShapeProperties,
// ECMA-376 Part 1, §20.1.2.3.13) block that every chart-frame element
// (chart-space, plot-area, legend, title, axis title, data-table,
// data-labels) carries to declare its fill, line color, line width
// and line dash.
//
// The reader uses these to lift the shared `<c:spPr>` grammar off any
// host element; the writer uses them to clamp + normalize values into
// the same uppercase-hex / point-snapped shapes the reader emits.
//
// All four readers (`parseSpPrFill`, `parseSpPrBorderColor`,
// `parseBorderWidthFromSpPr`, `parseBorderDashFromSpPr`) take the
// `<c:spPr>`'s **parent** element rather than the `<c:spPr>` itself
// so the per-host scope check stays inside the helper — the per-host
// "find legend / title / axis-title first" lookup happens at the
// caller, the helper takes care of the inner walk.

import type {
  ChartBorderDash,
  ChartColor,
  ChartLineCap,
  ChartLineCompound,
  ChartLineDashStyle,
  ChartThemeColor,
  ChartThemeColorName,
} from "./types"
import type { XmlElement } from "../../xml/parser"
import { xmlElement, xmlSelfClose } from "../../xml/writer"

/**
 * Local copy of `findChild`. The xml/parser module does not export the
 * helper, and the chart-reader / chart-clone files each carry their
 * own copy. Keeping a local definition here avoids a cross-module
 * refactor of the parser surface while still letting these primitives
 * walk a chart subtree.
 */
function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c
  }
  return undefined
}

// ── Stroke width ──────────────────────────────────────────────────
//
// Excel's UI exposes stroke widths on a `0.25..13.5` pt band, snapped
// to the 0.25 pt grid. The OOXML wire encodes the same value in
// English Metric Units (1 pt = 12 700 EMU) per CT_LineProperties
// (ECMA-376 Part 1, §20.1.2.3.24).

/** Smallest stroke width Excel's UI exposes, in points. */
export const STROKE_WIDTH_MIN_PT = 0.25
/** Largest stroke width Excel's UI exposes, in points. */
export const STROKE_WIDTH_MAX_PT = 13.5
/** Conversion factor between OOXML EMU and points. */
export const EMU_PER_PT = 12700

// ── Line dash ──────────────────────────────────────────────────────

/**
 * Recognized values of {@link ChartLineDashStyle} — the per-series
 * preset dash enum. Mirrors the OOXML `ST_PresetLineDashVal` set.
 */
export const VALID_DASH_STYLES: ReadonlySet<ChartLineDashStyle> = new Set([
  "solid",
  "dot",
  "dash",
  "lgDash",
  "dashDot",
  "lgDashDot",
  "lgDashDotDot",
  "sysDash",
  "sysDot",
  "sysDashDot",
  "sysDashDotDot",
])

/**
 * Recognized values of {@link ChartBorderDash} — the chart-frame
 * preset dash enum. Mirrors the OOXML `ST_PresetLineDashVal` set
 * exactly (see {@link VALID_DASH_STYLES} on the per-series side); the
 * reader collapses `"solid"` (the OOXML default) to `undefined` for
 * round-trip symmetry with the writer.
 */
export const VALID_BORDER_DASHES: ReadonlySet<ChartBorderDash> = new Set([
  "solid",
  "dash",
  "dashDot",
  "dot",
  "lgDash",
  "lgDashDot",
  "lgDashDotDot",
  "sysDash",
  "sysDashDot",
  "sysDashDotDot",
  "sysDot",
])

// ── Line cap / compound ───────────────────────────────────────────

/**
 * Recognized values of {@link ChartLineCap} — the OOXML `ST_LineCap`
 * enum on `<a:ln cap="..."/>`. The default token `"flat"` round-trips
 * as absence so the writer skips emitting the attribute.
 */
export const VALID_LINE_CAPS: ReadonlySet<ChartLineCap> = new Set(["rnd", "sq", "flat"])

/**
 * Recognized values of {@link ChartLineCompound} — the OOXML
 * `ST_CompoundLine` enum on `<a:ln cmpd="..."/>`. The default token
 * `"sng"` round-trips as absence so the writer skips emitting the
 * attribute.
 */
export const VALID_LINE_COMPOUNDS: ReadonlySet<ChartLineCompound> = new Set([
  "sng",
  "dbl",
  "thickThin",
  "thinThick",
  "tri",
])

/**
 * Normalize a {@link ChartLineCap} value for any chart-frame `<a:ln>`
 * slot the writer authors. Returns the recognized token when the input
 * is a valid `ST_LineCap` other than the OOXML default `"flat"`;
 * returns `undefined` for `"flat"` and for every unrecognized token so
 * absence and the default round-trip identically.
 */
export function normalizeLineCap(value: ChartLineCap | undefined): ChartLineCap | undefined {
  if (typeof value !== "string") return undefined
  if (!VALID_LINE_CAPS.has(value)) return undefined
  if (value === "flat") return undefined
  return value
}

/**
 * Normalize a {@link ChartLineCompound} value for any chart-frame
 * `<a:ln>` slot the writer authors. Returns the recognized token when
 * the input is a valid `ST_CompoundLine` other than the OOXML default
 * `"sng"`; returns `undefined` for `"sng"` and for every unrecognized
 * token so absence and the default round-trip identically.
 */
export function normalizeLineCompound(
  value: ChartLineCompound | undefined,
): ChartLineCompound | undefined {
  if (typeof value !== "string") return undefined
  if (!VALID_LINE_COMPOUNDS.has(value)) return undefined
  if (value === "sng") return undefined
  return value
}

/**
 * Resolve a chart-frame border-cap override.
 *
 * `undefined` → inherit the source value (after running it through
 *               {@link normalizeLineCap}).
 * `null`      → drop the inherited cap.
 * value       → replace with the normalized cap. Unrecognized tokens
 *               and the OOXML default `"flat"` collapse to `undefined`.
 */
export function resolveLineCap(
  sourceValue: ChartLineCap | undefined,
  override: ChartLineCap | null | undefined,
): ChartLineCap | undefined {
  if (override === undefined) return normalizeLineCap(sourceValue)
  if (override === null) return undefined
  return normalizeLineCap(override)
}

/**
 * Resolve a chart-frame border-compound override.
 *
 * `undefined` → inherit the source value (after running it through
 *               {@link normalizeLineCompound}).
 * `null`      → drop the inherited compound style.
 * value       → replace with the normalized compound. Unrecognized
 *               tokens and the OOXML default `"sng"` collapse to
 *               `undefined`.
 */
export function resolveLineCompound(
  sourceValue: ChartLineCompound | undefined,
  override: ChartLineCompound | null | undefined,
): ChartLineCompound | undefined {
  if (override === undefined) return normalizeLineCompound(sourceValue)
  if (override === null) return undefined
  return normalizeLineCompound(override)
}

/**
 * Pull the `cap` attribute off a `<c:spPr><a:ln cap="..."/>` block
 * scoped to the supplied parent. Returns the {@link ChartLineCap} when
 * the value is a recognized `ST_LineCap` token other than the OOXML
 * default `"flat"`. Returns `undefined` when the chain is missing,
 * when the attribute is absent, when the value is unrecognized, or
 * when it matches `"flat"`.
 */
export function parseBorderCapFromSpPr(parent: XmlElement): ChartLineCap | undefined {
  const spPr = findChild(parent, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const raw = ln.attrs.cap
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim() as ChartLineCap
  if (!VALID_LINE_CAPS.has(trimmed)) return undefined
  if (trimmed === "flat") return undefined
  return trimmed
}

/**
 * Pull the `cmpd` attribute off a `<c:spPr><a:ln cmpd="..."/>` block
 * scoped to the supplied parent. Returns the {@link ChartLineCompound}
 * when the value is a recognized `ST_CompoundLine` token other than
 * the OOXML default `"sng"`. Returns `undefined` when the chain is
 * missing, when the attribute is absent, when the value is
 * unrecognized, or when it matches `"sng"`.
 */
export function parseBorderCompoundFromSpPr(parent: XmlElement): ChartLineCompound | undefined {
  const spPr = findChild(parent, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const raw = ln.attrs.cmpd
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim() as ChartLineCompound
  if (!VALID_LINE_COMPOUNDS.has(trimmed)) return undefined
  if (trimmed === "sng") return undefined
  return trimmed
}

// ── Theme color refs ──────────────────────────────────────────────

/**
 * Recognized values of {@link ChartThemeColorName} — the OOXML
 * `ST_SchemeColorVal` enum on `<a:schemeClr val="..."/>`.
 */
export const VALID_THEME_COLOR_NAMES: ReadonlySet<ChartThemeColorName> = new Set([
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
  "phClr",
  "dk1",
  "lt1",
  "dk2",
  "lt2",
])

const POSITIVE_PERCENT_MIN = 0
const POSITIVE_PERCENT_MAX = 100000
const FIXED_PERCENT_MIN = -100000
const FIXED_PERCENT_MAX = 100000

/**
 * Pull a single integer modifier off a `<a:schemeClr>` child, clamping
 * to the supplied range. Mod children carry a `val` attribute that
 * encodes the percentage value as an integer per OOXML
 * `ST_PositivePercentage` / `ST_FixedPercentage`. Returns `undefined`
 * when the child is absent, the attribute is missing, the value is not
 * a finite integer, or the value falls outside the allowed band.
 */
function parseSchemeClrMod(
  schemeClr: XmlElement,
  localName: string,
  min: number,
  max: number,
): number | undefined {
  const child = findChild(schemeClr, localName)
  if (!child) return undefined
  const raw = child.attrs.val
  if (typeof raw !== "string") return undefined
  const n = Number.parseInt(raw.trim(), 10)
  if (!Number.isFinite(n)) return undefined
  if (n < min || n > max) return undefined
  return n
}

/**
 * Lift a `<a:schemeClr>` element into a {@link ChartThemeColor}.
 * Returns `undefined` when the `val` attribute is missing or not a
 * recognized theme color name. Walks every supported modifier child
 * (`<a:lumMod>`, `<a:lumOff>`, `<a:tint>`, `<a:shade>`, `<a:alpha>`)
 * and surfaces only the ones that carry a parseable integer in the
 * relevant percentage band; malformed or out-of-range mods drop
 * silently so a parsed-then-written theme color still survives a
 * round-trip even when the source carried garbage modifiers.
 */
export function parseSchemeClr(schemeClr: XmlElement): ChartThemeColor | undefined {
  const raw = schemeClr.attrs.val
  if (typeof raw !== "string") return undefined
  const name = raw.trim() as ChartThemeColorName
  if (!VALID_THEME_COLOR_NAMES.has(name)) return undefined
  const out: ChartThemeColor = { theme: name }
  const lumMod = parseSchemeClrMod(schemeClr, "lumMod", POSITIVE_PERCENT_MIN, POSITIVE_PERCENT_MAX)
  if (lumMod !== undefined) out.lumMod = lumMod
  const lumOff = parseSchemeClrMod(schemeClr, "lumOff", POSITIVE_PERCENT_MIN, POSITIVE_PERCENT_MAX)
  if (lumOff !== undefined) out.lumOff = lumOff
  const tint = parseSchemeClrMod(schemeClr, "tint", FIXED_PERCENT_MIN, FIXED_PERCENT_MAX)
  if (tint !== undefined) out.tint = tint
  const shade = parseSchemeClrMod(schemeClr, "shade", FIXED_PERCENT_MIN, FIXED_PERCENT_MAX)
  if (shade !== undefined) out.shade = shade
  const alpha = parseSchemeClrMod(schemeClr, "alpha", POSITIVE_PERCENT_MIN, POSITIVE_PERCENT_MAX)
  if (alpha !== undefined) out.alpha = alpha
  return out
}

/**
 * Normalize a {@link ChartColor} input for emit. Returns the
 * normalized form when the input is a valid sRGB triple (string) or a
 * recognized theme-color reference (object); returns `undefined`
 * otherwise. Dropping the value lets the caller skip emitting the
 * surrounding `<a:solidFill>` / `<a:ln>` block entirely so absence
 * round-trips identically through the writer.
 */
export function normalizeChartColor(value: ChartColor | undefined): ChartColor | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value === "string") {
    const hex = normalizeRgbHex(value)
    return hex
  }
  if (typeof value !== "object") return undefined
  const name = value.theme
  if (typeof name !== "string" || !VALID_THEME_COLOR_NAMES.has(name)) return undefined
  const out: ChartThemeColor = { theme: name }
  const validateMod = (raw: number | undefined, min: number, max: number): number | undefined => {
    if (typeof raw !== "number" || !Number.isFinite(raw)) return undefined
    const n = Math.round(raw)
    if (n < min || n > max) return undefined
    return n
  }
  const lumMod = validateMod(value.lumMod, POSITIVE_PERCENT_MIN, POSITIVE_PERCENT_MAX)
  if (lumMod !== undefined) out.lumMod = lumMod
  const lumOff = validateMod(value.lumOff, POSITIVE_PERCENT_MIN, POSITIVE_PERCENT_MAX)
  if (lumOff !== undefined) out.lumOff = lumOff
  const tint = validateMod(value.tint, FIXED_PERCENT_MIN, FIXED_PERCENT_MAX)
  if (tint !== undefined) out.tint = tint
  const shade = validateMod(value.shade, FIXED_PERCENT_MIN, FIXED_PERCENT_MAX)
  if (shade !== undefined) out.shade = shade
  const alpha = validateMod(value.alpha, POSITIVE_PERCENT_MIN, POSITIVE_PERCENT_MAX)
  if (alpha !== undefined) out.alpha = alpha
  return out
}

/**
 * Build the inner color element for a `<a:solidFill>` / `<a:ln>` slot:
 * `<a:srgbClr val="RRGGBB"/>` for a literal hex string, or
 * `<a:schemeClr val="..."><mods/></a:schemeClr>` for a
 * {@link ChartThemeColor}. The caller normalizes before emit; this
 * builder assumes the input is already valid (it does not re-validate).
 */
export function buildColorElement(value: ChartColor): string {
  if (typeof value === "string") {
    return xmlSelfClose("a:srgbClr", { val: value })
  }
  // ChartThemeColor — emit the modifiers in the OOXML schema order
  // documented on `CT_SchemeColor` (ECMA-376 Part 1, §20.1.2.3.29):
  // tint, shade, alpha, lumMod, lumOff. Excel tolerates other orders
  // but the schema sequence is canonical.
  const children: string[] = []
  if (value.tint !== undefined) {
    children.push(xmlSelfClose("a:tint", { val: value.tint }))
  }
  if (value.shade !== undefined) {
    children.push(xmlSelfClose("a:shade", { val: value.shade }))
  }
  if (value.alpha !== undefined) {
    children.push(xmlSelfClose("a:alpha", { val: value.alpha }))
  }
  if (value.lumMod !== undefined) {
    children.push(xmlSelfClose("a:lumMod", { val: value.lumMod }))
  }
  if (value.lumOff !== undefined) {
    children.push(xmlSelfClose("a:lumOff", { val: value.lumOff }))
  }
  if (children.length === 0) {
    return xmlSelfClose("a:schemeClr", { val: value.theme })
  }
  return xmlElement("a:schemeClr", { val: value.theme }, children)
}

/**
 * Build a `<a:solidFill>` block wrapping the supplied color reference.
 */
export function buildSolidFill(value: ChartColor): string {
  return xmlElement("a:solidFill", undefined, [buildColorElement(value)])
}

/**
 * Resolve a chart-frame fill / line color override carrying the
 * widened {@link ChartColor} type.
 *
 * `undefined` → inherit the source value (after running it through
 *               {@link normalizeChartColor}).
 * `null`      → drop the inherited color.
 * value       → replace with the normalized color. Malformed inputs
 *               (wrong-length hex, unrecognized theme names,
 *               out-of-range mods) collapse to `undefined`.
 */
export function resolveChartColor(
  sourceValue: ChartColor | undefined,
  override: ChartColor | null | undefined,
): ChartColor | undefined {
  if (override === undefined) return normalizeChartColor(sourceValue)
  if (override === null) return undefined
  return normalizeChartColor(override)
}

// ── Hex normalization ─────────────────────────────────────────────

/**
 * Normalize an `<a:srgbClr val=".."/>` token into a 6-character
 * uppercase hex string. Accepts the input with or without a leading
 * `#`, and tolerates leading / trailing whitespace. Returns
 * `undefined` for any malformed input — wrong length, non-hex
 * characters, alpha-channel forms, or non-string tokens.
 */
export function normalizeRgbHex(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  const hex = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed
  if (hex.length !== 6) return undefined
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return undefined
  return hex.toUpperCase()
}

// ── Generic spPr readers ──────────────────────────────────────────

/**
 * Pull `<c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>` (or
 * `<a:schemeClr val=".."/>`) off the supplied parent element (one of
 * `<c:title>` / `<c:legend>` / `<c:plotArea>` / `<c:chartSpace>` /
 * `<c:dTable>` / `<c:dLbls>` / a series — anywhere a
 * `CT_ShapeProperties` block can live). Returns the fill color as a
 * {@link ChartColor} — a 6-character uppercase hex string for
 * `<a:srgbClr>`, or a {@link ChartThemeColor} object for
 * `<a:schemeClr>`.
 *
 * The OOXML `<a:solidFill>` choice (`CT_SolidColorFillProperties`,
 * ECMA-376 Part 1, §20.1.8.54) accepts both `<a:srgbClr>` and
 * `<a:schemeClr>` (plus a handful of less-common forms — `<a:sysClr>`,
 * `<a:scrgbClr>`, `<a:hslClr>`, `<a:prstClr>` — which still collapse
 * to `undefined` to keep the writer's emit grammar tractable). Theme
 * color references carry the named slot in `val` and may layer
 * luminance / tint / shade / alpha modifiers; the reader walks each
 * mod child and drops malformed / out-of-range values silently.
 *
 * Non-solid fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
 * `<a:blipFill>`) collapse to `undefined`. Malformed `val` tokens
 * likewise drop to `undefined`.
 */
export function parseSpPrFill(parent: XmlElement): ChartColor | undefined {
  const spPr = findChild(parent, "spPr")
  if (!spPr) return undefined
  const solidFill = findChild(spPr, "solidFill")
  if (!solidFill) return undefined
  const srgbClr = findChild(solidFill, "srgbClr")
  if (srgbClr) {
    return normalizeRgbHex(srgbClr.attrs.val)
  }
  const schemeClr = findChild(solidFill, "schemeClr")
  if (schemeClr) {
    return parseSchemeClr(schemeClr)
  }
  return undefined
}

/**
 * Pull `<c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>` (or
 * `<a:schemeClr val=".."/>`) off the supplied parent element. Returns
 * the line stroke color as a {@link ChartColor}.
 *
 * Same accept-or-drop grammar as {@link parseSpPrFill}; lands on
 * the line-color child rather than the fill child of the same
 * `<c:spPr>` block.
 */
export function parseSpPrBorderColor(parent: XmlElement): ChartColor | undefined {
  const spPr = findChild(parent, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const solidFill = findChild(ln, "solidFill")
  if (!solidFill) return undefined
  const srgbClr = findChild(solidFill, "srgbClr")
  if (srgbClr) {
    return normalizeRgbHex(srgbClr.attrs.val)
  }
  const schemeClr = findChild(solidFill, "schemeClr")
  if (schemeClr) {
    return parseSchemeClr(schemeClr)
  }
  return undefined
}

/**
 * Pull the `w` attribute off a `<c:spPr><a:ln w="EMU"/>` block scoped
 * to the supplied parent (`<c:plotArea>` / `<c:legend>` / `<c:title>` /
 * `<c:chartSpace>` / `<c:dTable>` / `<c:dLbls>`). Returns the stroke
 * width in points after clamping to the `0.25..13.5` pt band Excel's
 * UI exposes; the OOXML `w` attribute carries the value in EMU
 * (1 pt = 12 700 EMU) per CT_LineProperties (ECMA-376 Part 1,
 * §20.1.2.3.24). Snaps to the 0.25 pt grid Excel's UI exposes so a
 * parsed-then-written width does not drift across round-trips.
 *
 * Returns `undefined` when there is no `<c:spPr><a:ln>` block, when
 * the attribute is missing, when the value cannot be parsed as a
 * finite positive number, or when it parses to zero (Excel's "no
 * border" marker — the writer-side knob does not model that state).
 */
export function parseBorderWidthFromSpPr(parent: XmlElement): number | undefined {
  const spPr = findChild(parent, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const wAttr = ln.attrs.w
  if (typeof wAttr !== "string") return undefined
  const emu = Number.parseFloat(wAttr)
  if (!Number.isFinite(emu) || emu <= 0) return undefined
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const pt = Math.round((emu / EMU_PER_PT) * 4) / 4
  if (pt < STROKE_WIDTH_MIN_PT) return STROKE_WIDTH_MIN_PT
  if (pt > STROKE_WIDTH_MAX_PT) return STROKE_WIDTH_MAX_PT
  return pt
}

/**
 * Pull the `val` attribute off a `<c:spPr><a:ln><a:prstDash val=".."/>`
 * chain scoped to the supplied parent. Returns the {@link ChartBorderDash}
 * value when the chain is present and the value is a recognized
 * `ST_PresetLineDashVal` token other than the OOXML default `"solid"`.
 *
 * Returns `undefined` when the chain is missing at any link, when the
 * attribute is absent, when the value is unrecognized, or when it
 * matches the OOXML default `"solid"` (so absence and the default
 * round-trip identically through `cloneChart`). Mirrors the
 * writer-side `normalizeBorderDash` so the accept-or-drop grammar
 * matches every chart-frame border-dash slot the writer authors.
 */
export function parseBorderDashFromSpPr(parent: XmlElement): ChartBorderDash | undefined {
  const spPr = findChild(parent, "spPr")
  if (!spPr) return undefined
  const ln = findChild(spPr, "ln")
  if (!ln) return undefined
  const prstDash = findChild(ln, "prstDash")
  if (!prstDash) return undefined
  const raw = prstDash.attrs.val
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim() as ChartBorderDash
  if (!VALID_BORDER_DASHES.has(trimmed)) return undefined
  if (trimmed === "solid") return undefined
  return trimmed
}

// ── Stroke width clamp ────────────────────────────────────────────

/**
 * Convert a stroke width in points to the integer EMU value the OOXML
 * `w` attribute requires. Excel's UI exposes 0.25..13.5 pt — values
 * outside that band are clamped to keep round-trips inside the range
 * Excel will render. Non-finite values collapse to `undefined` so the
 * writer can omit the attribute entirely. The point value is also
 * snapped to the nearest quarter-point so a parsed-then-written stroke
 * does not drift across round-trips (Excel rounds in the UI anyway).
 */
export function clampStrokeWidthPt(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const snapped = Math.round(value * 4) / 4
  if (snapped < STROKE_WIDTH_MIN_PT) return STROKE_WIDTH_MIN_PT
  if (snapped > STROKE_WIDTH_MAX_PT) return STROKE_WIDTH_MAX_PT
  return snapped
}

/**
 * Normalize a {@link ChartBorderDash} value for any chart-frame
 * `<a:prstDash>` slot the writer authors. Returns the recognized
 * token when the input is a valid `ST_PresetLineDashVal` other than
 * the OOXML default `"solid"`; returns `undefined` for `"solid"` and
 * for every unrecognized token so absence and the default round-trip
 * identically through `cloneChart`.
 */
export function normalizeBorderDash(
  value: ChartBorderDash | undefined,
): ChartBorderDash | undefined {
  if (typeof value !== "string") return undefined
  if (!VALID_BORDER_DASHES.has(value)) return undefined
  if (value === "solid") return undefined
  return value
}

/**
 * Resolve a chart-frame border-width override.
 *
 * `undefined` → inherit the source value (after running it through
 *               the writer-side clamp `clampStrokeWidthPt`, so a
 *               malformed source value drops cleanly to `undefined`).
 * `null`      → drop the inherited width (the writer falls back to
 *               Excel's auto-stroke thickness — no `w` attribute on
 *               `<a:ln>`).
 * `number`    → replace with the clamped value. Out-of-range values
 *               clamp to the `0.25..13.5` pt band Excel's UI exposes.
 *               Non-finite / non-numeric overrides collapse to
 *               `undefined` via the normalizer.
 *
 * Used by every chart-frame border-width slot the clone surface
 * exposes — chart-space, axis-title, data table, data labels — and
 * mirrors the existing host-specific resolvers for plot-area / legend /
 * title border widths.
 */
export function resolveBorderWidthPt(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return clampStrokeWidthPt(sourceValue)
  if (override === null) return undefined
  return clampStrokeWidthPt(override)
}

/**
 * Resolve a chart-frame border-dash override.
 *
 * `undefined` → inherit the source value (after running it through
 *               {@link normalizeBorderDash}).
 * `null`      → drop the inherited dash (the writer renders solid).
 * value       → replace with the normalized dash. Unrecognized tokens
 *               and the OOXML default `"solid"` collapse to
 *               `undefined`.
 */
export function resolveBorderDash(
  sourceValue: ChartBorderDash | undefined,
  override: ChartBorderDash | null | undefined,
): ChartBorderDash | undefined {
  if (override === undefined) return normalizeBorderDash(sourceValue)
  if (override === null) return undefined
  return normalizeBorderDash(override)
}
