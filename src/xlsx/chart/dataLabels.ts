// ── Chart Data Labels ─────────────────────────────────────────────
// Per-host module for `<c:dLbls>` (CT_DLbls, ECMA-376 Part 1, §21.2.2.49)
// — both the chart-level block (sibling of `<c:ser>` inside the
// chart-type element) and the per-series block (child of `<c:ser>`).
// Holds the reader / writer helpers — every `parse*` / `build*` /
// `resolve*` function for the data-labels block, including its
// `<c:txPr>` typography, `<c:numFmt>`, `<c:spPr>` fill / border slots,
// and the per-position visibility flags.
//
// The clone-side `resolveChartDataLabels` / `resolveSeriesDataLabels`
// override functions live in `chart-clone.ts` because their three-arg
// `(source, override)` signature differs from the writer-side single-arg
// resolvers.

import type {
  ChartAxisNumberFormat,
  ChartBorderDash,
  ChartColor,
  ChartDataLabels,
  ChartDataLabelPosition,
  ChartDataLabelsInfo,
  ChartLineCap,
  ChartLineCompound,
  WriteChartKind,
} from "../../_types"
import type { XmlElement } from "../../xml/parser"
import { xmlElement, xmlEscape, xmlSelfClose } from "../../xml/writer"
import {
  EMU_PER_PT,
  buildColorElement,
  clampStrokeWidthPt,
  normalizeBorderDash,
  normalizeLineCap,
  normalizeLineCompound,
  normalizeRgbHex,
  parseBorderCapFromSpPr,
  parseBorderCompoundFromSpPr,
  parseBorderDashFromSpPr,
  parseBorderWidthFromSpPr,
  parseSchemeClr,
  parseSpPrBorderColor,
  parseSpPrFill,
} from "./shape"
import { elementText, findChild, readBoolAttr } from "./util"
import { FONT_SIZE_MAX_PT, FONT_SIZE_MIN_PT, FONT_SZ_PER_POINT } from "./text"
import { normalizeTitleColor, normalizeTitleFontSize } from "./title"
import type { SheetChart } from "../../_types"

const TITLE_FONT_SZ_PER_POINT = FONT_SZ_PER_POINT
const TITLE_FONT_SIZE_MIN_PT = FONT_SIZE_MIN_PT
const TITLE_FONT_SIZE_MAX_PT = FONT_SIZE_MAX_PT

/** Recognised OOXML data-label positions (`ChartDataLabelPosition`). */
const VALID_DLBL_POSITIONS: ReadonlySet<ChartDataLabelPosition> = new Set([
  "t",
  "b",
  "l",
  "r",
  "ctr",
  "inEnd",
  "inBase",
  "outEnd",
  "bestFit",
])

// ── Reader ────────────────────────────────────────────────────────

/**
 * Read a `<c:dLbls>` block. Returns `undefined` when the block is
 * empty or only contains a `<c:delete val="1">` (which suppresses
 * labels rather than describing them). All toggles default to `false`
 * when the matching `<c:show*>` element is absent.
 */
export function parseDataLabels(el: XmlElement): ChartDataLabelsInfo | undefined {
  // <c:delete val="1"> at the root of <c:dLbls> means "suppress for
  // this scope". We don't surface a dataLabels record for that case —
  // it's the absence of labels, not a configuration.
  const deleteEl = findChild(el, "delete")
  if (deleteEl && readBoolAttr(deleteEl) === true) return undefined

  const out: ChartDataLabelsInfo = {}

  const pos = findChild(el, "dLblPos")
  if (pos) {
    const val = pos.attrs.val
    if (typeof val === "string" && VALID_DLBL_POSITIONS.has(val as ChartDataLabelPosition)) {
      out.position = val as ChartDataLabelPosition
    }
  }

  // `<c:showLegendKey val=".."/>` mirrors Excel's "Format Data Labels
  // -> Legend Key" checkbox. The OOXML default is `false`, so absence
  // and an explicit `val="0"` collapse to `undefined` — only an
  // explicit `val="1"` (or `"true"`) surfaces `true`. Same shape as the
  // other `show*` toggles so the parsed record can be fed straight back
  // into {@link cloneChart}.
  const showLeg = findChild(el, "showLegendKey")
  if (showLeg && readBoolAttr(showLeg) === true) out.showLegendKey = true

  const showVal = findChild(el, "showVal")
  if (showVal && readBoolAttr(showVal) === true) out.showValue = true

  const showCat = findChild(el, "showCatName")
  if (showCat && readBoolAttr(showCat) === true) out.showCategoryName = true

  const showSer = findChild(el, "showSerName")
  if (showSer && readBoolAttr(showSer) === true) out.showSeriesName = true

  const showPct = findChild(el, "showPercent")
  if (showPct && readBoolAttr(showPct) === true) out.showPercent = true

  const sep = findChild(el, "separator")
  if (sep) {
    const text = elementText(sep)
    if (text.length > 0) out.separator = text
  }

  // `<c:numFmt formatCode=".." sourceLinked=".."/>` mirrors Excel's
  // "Format Data Labels -> Number" panel — pinning a custom number
  // format on the rendered label values. Same shape as the axis-side
  // `<c:numFmt>` so the parsed value can be fed straight back into
  // {@link cloneChart}. The element sits early in the CT_DLbls
  // sequence (after the optional `<c:dLbl>` instances), so the lookup
  // is scoped to direct `<c:dLbls>` children — a `<c:numFmt>` nested
  // inside a per-point `<c:dLbl>` does not leak into the block-level
  // record.
  const numFmt = parseDataLabelsNumberFormat(el)
  if (numFmt) out.numberFormat = numFmt

  // `<c:showLeaderLines val=".."/>` mirrors Excel's "Format Data
  // Labels -> Show Leader Lines" checkbox. The OOXML default is
  // `true` (Excel paints leader lines on every label that gets pushed
  // outside its slice), so absence and `val="1"` collapse to
  // `undefined` — only an explicit `val="0"` (or `"false"`) surfaces
  // `false`. Mirrors how the writer-side scope guard treats the
  // element: only meaningful on pie / doughnut, but the parser is
  // permissive (the OOXML schema scopes the element to `EG_DLbls` for
  // `CT_PieChart` / `CT_DoughnutChart`, but a templated chart whose
  // type element ends up coerced should still surface the source's
  // intent so the cloned model stays accurate).
  // Only the literal OOXML falsy spellings (`"0"` / `"false"`) flip the
  // toggle — unknown / missing val tokens collapse to `undefined` rather
  // than surface a `false` the writer would round-trip into a non-default
  // `<c:showLeaderLines val="0"/>` Excel never authored.
  const showLeader = findChild(el, "showLeaderLines")
  if (showLeader) {
    const v = showLeader.attrs.val
    if (typeof v === "string" && (v === "0" || v.toLowerCase() === "false")) {
      out.showLeaderLines = false
    }
  }

  // `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
  // — data-label font size pinned via Excel's "Format Data Labels ->
  // Font -> Size" knob. The OOXML attribute is in 100ths of a point on
  // `CT_TextCharacterProperties`' `sz` slot; the reader divides by
  // 100 at parse time so the surfaced value matches what the user
  // sees in the UI. Out-of-range / malformed tokens collapse to
  // `undefined` so absence and a malformed source value round-trip
  // identically through `cloneChart`.
  const fontSize = parseDataLabelsFontSize(el)
  if (fontSize !== undefined) out.fontSize = fontSize

  // `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
  // </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` — data-label
  // font color pinned via Excel's "Format Data Labels -> Font -> Font
  // color" picker. Theme references (`<a:schemeClr>`) and malformed
  // `val` tokens collapse to `undefined` since only the literal RGB
  // triple round-trips losslessly through `writeChart`.
  const fontColor = parseDataLabelsFontColor(el)
  if (fontColor !== undefined) out.fontColor = fontColor

  // `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>` —
  // data-label bold flag pinned via Excel's "Format Data Labels ->
  // Font -> Bold" toggle. Only an explicit `b="1"` surfaces `true`;
  // the OOXML default `b="0"` collapses to `undefined` so absence
  // and `b="0"` round-trip identically through `cloneChart`.
  const bold = parseDataLabelsBold(el)
  if (bold !== undefined) out.bold = bold

  // `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>` —
  // data-label italic flag pinned via Excel's "Format Data Labels ->
  // Font -> Italic" toggle. Only an explicit `i="1"` surfaces `true`;
  // the OOXML default `i="0"` collapses to `undefined` so absence
  // and `i="0"` round-trip identically through `cloneChart`.
  const italic = parseDataLabelsItalic(el)
  if (italic !== undefined) out.italic = italic

  // `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>` —
  // data-label underline flag pinned via Excel's "Format Data Labels
  // -> Font -> Underline" toggle. Only `u="sng"` (Excel's UI variant
  // — single underline) surfaces `true`; the OOXML default `"none"`
  // (and every other ST_TextUnderlineType variant) collapse to
  // `undefined` so absence and `u="none"` round-trip identically
  // through `cloneChart`.
  const underline = parseDataLabelsUnderline(el)
  if (underline !== undefined) out.underline = underline

  // `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p>
  // </c:txPr>` — data-label strikethrough flag pinned via Excel's
  // "Format Data Labels -> Font -> Strikethrough" toggle. Only
  // `strike="sngStrike"` (Excel's UI variant — single line) surfaces
  // `true`; the OOXML default `"noStrike"` and the non-UI variant
  // `"dblStrike"` (and any malformed token) collapse to `undefined`
  // so absence and `"noStrike"` round-trip identically through
  // `cloneChart`.
  const strikethrough = parseDataLabelsStrikethrough(el)
  if (strikethrough !== undefined) out.strikethrough = strikethrough

  // `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
  // </a:pPr></a:p></c:txPr>` — data-label font family pinned via
  // Excel's "Format Data Labels -> Font -> Font" picker. Empty /
  // whitespace-only `typeface` attributes and missing `<a:latin>`
  // elements both collapse to `undefined` so absence and the empty
  // form round-trip identically through the writer.
  const fontFamily = parseDataLabelsFontFamily(el)
  if (fontFamily !== undefined) out.fontFamily = fontFamily

  // `<c:spPr><a:solidFill><a:srgbClr val=".."/></a:solidFill></c:spPr>`
  // — data-labels background fill pinned via Excel's "Format Data Labels
  // -> Fill -> Solid fill -> Color" picker. Theme references
  // (`<a:schemeClr>`), non-solid fills (`<a:noFill>` / `<a:gradFill>` /
  // `<a:pattFill>` / `<a:blipFill>`), and malformed `val` tokens
  // collapse to `undefined` since only the literal RGB triple
  // round-trips losslessly through `writeChart`. Distinct from
  // `fontColor` (which lives on `<c:txPr><a:p><a:pPr><a:defRPr>
  // <a:solidFill>`); the two knobs target different children of
  // `<c:dLbls>` so a caller can pin both without conflict.
  const fillColor = parseDataLabelsFillColor(el)
  if (fillColor !== undefined) out.fillColor = fillColor

  // `<c:spPr><a:ln><a:solidFill><a:srgbClr val=".."/></a:solidFill>
  // </a:ln></c:spPr>` — data-labels border (line) color pinned via
  // Excel's "Format Data Labels -> Border -> Solid line -> Color"
  // picker. Theme references (`<a:schemeClr>`), non-solid line fills
  // (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>`), and malformed
  // `val` tokens collapse to `undefined` since only the literal RGB
  // triple round-trips losslessly through `writeChart`. Distinct from
  // `fillColor` (which lives on `<c:spPr><a:solidFill>`); the two
  // knobs target different children of the same `<c:spPr>` so a
  // caller can pin both without conflict.
  const borderColor = parseDataLabelsBorderColor(el)
  if (borderColor !== undefined) out.borderColor = borderColor

  // `<c:dLbls><c:spPr><a:ln w="EMU">` carries Excel's "Format Data
  // Labels -> Border -> Width" pin. Delegates to the shared
  // {@link parseBorderWidthFromSpPr} so the EMU encoding and snap /
  // clamp grammar match every other chart-frame border-width slot.
  const borderWidth = parseBorderWidthFromSpPr(el)
  if (borderWidth !== undefined) out.borderWidth = borderWidth

  // `<c:dLbls><c:spPr><a:ln><a:prstDash val=".."/>` carries Excel's
  // "Format Data Labels -> Border -> Dash type" pin. Delegates to the
  // shared {@link parseBorderDashFromSpPr} helper.
  const borderDash = parseBorderDashFromSpPr(el)
  if (borderDash !== undefined) out.borderDash = borderDash

  // `<c:dLbls><c:spPr><a:ln cap=".."/>` carries Excel's per-line cap
  // attribute. Delegates to the shared {@link parseBorderCapFromSpPr}
  // helper so the accept-or-drop grammar matches every other chart-
  // frame `<a:ln>` slot the reader surfaces.
  const borderCap = parseBorderCapFromSpPr(el)
  if (borderCap !== undefined) out.borderCap = borderCap

  // `<c:dLbls><c:spPr><a:ln cmpd=".."/>` carries Excel's per-line
  // compound attribute. Delegates to the shared
  // {@link parseBorderCompoundFromSpPr} helper.
  const borderCompound = parseBorderCompoundFromSpPr(el)
  if (borderCompound !== undefined) out.borderCompound = borderCompound

  // Empty record is meaningless to a consumer — collapse to undefined.
  if (
    out.position === undefined &&
    !out.showValue &&
    !out.showCategoryName &&
    !out.showSeriesName &&
    !out.showPercent &&
    !out.showLegendKey &&
    out.separator === undefined &&
    out.numberFormat === undefined &&
    out.showLeaderLines === undefined &&
    out.fontSize === undefined &&
    out.fontColor === undefined &&
    out.bold === undefined &&
    out.italic === undefined &&
    out.underline === undefined &&
    out.strikethrough === undefined &&
    out.fontFamily === undefined &&
    out.fillColor === undefined &&
    out.borderColor === undefined &&
    out.borderWidth === undefined &&
    out.borderDash === undefined &&
    out.borderCap === undefined &&
    out.borderCompound === undefined
  ) {
    return undefined
  }
  return out
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the font size
 * in points (`1..400`), or `undefined` when the element is absent /
 * the chain is malformed at any link / the surfaced value is out of
 * the supported range.
 *
 * The OOXML attribute is in 100ths of a point on
 * `CT_TextCharacterProperties`' `sz` slot (ECMA-376 Part 1,
 * §21.1.2.3.7); the reader divides by 100 at parse time so the
 * surfaced value matches what the user sees in Excel's UI. Mirrors
 * the chart-title / axis-title / axis tick-label / legend font-size
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
export function parseDataLabelsFontSize(dLbls: XmlElement): number | undefined {
  const txPr = findChild(dLbls, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.sz
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  const parsed = Number.parseInt(trimmed, 10)
  if (!Number.isFinite(parsed)) return undefined
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. Mirrors the
  // chart-title / axis-title / tick-label / legend sibling parsers
  // exactly so a parsed value flows through every typography slot
  // without bookkeeping the units.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2)
  const points = halfSteps / 2
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined
  return points
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the
 * 6-character uppercase hex string when the parser walks the full
 * chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme references
 * (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>`
 * all collapse to `undefined` — only the literal RGB triple
 * round-trips losslessly through {@link writeChart}. Malformed `val`
 * tokens (wrong length, non-hex characters) likewise drop to
 * `undefined` rather than fabricate a value the writer would
 * round-trip into a malformed `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the data-labels block omits `<c:txPr>`
 * entirely or the canonical chain is malformed at any link. Mirrors
 * the chart-title / axis-title / axis tick-label / legend color
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
export function parseDataLabelsFontColor(dLbls: XmlElement): ChartColor | undefined {
  const txPr = findChild(dLbls, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const solidFill = findChild(defRPr, "solidFill")
  if (!solidFill) return undefined
  const srgbClr = findChild(solidFill, "srgbClr")
  if (srgbClr) return normalizeRgbHex(srgbClr.attrs.val)
  const schemeClr = findChild(solidFill, "schemeClr")
  if (schemeClr) return parseSchemeClr(schemeClr)
  return undefined
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the bold
 * flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties`. Only an explicit `b="1"` (or
 * `"true"`) surfaces `true`; the OOXML default `0` (and absence /
 * malformed tokens) collapses to `undefined` so absence and the
 * default round-trip identically through `cloneChart`. Mirrors the
 * chart-title / axis-title / axis tick-label / legend bold readers
 * exactly so a parsed value slots straight back into the writer's
 * emit path.
 */
export function parseDataLabelsBold(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.b
  if (raw === "1" || raw === "true") return true
  return undefined
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the italic
 * flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties`. Only an explicit `i="1"` (or
 * `"true"`) surfaces `true`; the OOXML default `0` (and absence /
 * malformed tokens) collapses to `undefined` so absence and the
 * default round-trip identically through `cloneChart`. Mirrors the
 * chart-title / axis-title / axis tick-label / legend italic readers
 * exactly so a parsed value slots straight back into the writer's
 * emit path.
 */
export function parseDataLabelsItalic(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.i
  if (raw === "1" || raw === "true") return true
  return undefined
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p>
 * </c:txPr></c:dLbls>` off a data-labels block. Returns the underline
 * flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enumeration
 * on `CT_TextCharacterProperties`. Only `u="sng"` (Excel's UI variant
 * — single underline) surfaces `true`; the OOXML default `"none"`
 * (and every other variant the schema allows — `"dbl"`, `"heavy"`,
 * `"dotted"`, `"dotDash"`, `"wavy"`, etc.) collapse to `undefined` so
 * absence and `u="none"` round-trip identically through `cloneChart`.
 * Reporting any non-`"sng"` underline as `true` would silently
 * downgrade the choice to a single line on round-trip; the writer
 * emits only `u="sng"` / `u="none"`, matching the boolean shape the
 * UI exposes. Mirrors the chart-title / axis-title / axis tick-label
 * / legend underline readers exactly so a parsed value slots straight
 * back into the writer's emit path.
 */
export function parseDataLabelsUnderline(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.u
  if (raw === "sng") return true
  return undefined
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:txPr></c:dLbls>` off a data-labels block. Returns the
 * strikethrough flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType`
 * enumeration on `CT_TextCharacterProperties`. Only
 * `strike="sngStrike"` (Excel's UI variant — single line) surfaces
 * `true`; the OOXML default `"noStrike"` and the non-UI variant
 * `"dblStrike"` (and any malformed token) collapse to `undefined` so
 * absence and `"noStrike"` round-trip identically through
 * `cloneChart`. Reporting `"dblStrike"` as `true` would silently
 * downgrade the choice to a single line on round-trip; the writer
 * emits only `"sngStrike"`, matching the boolean shape the UI
 * exposes. Mirrors the chart-title / axis-title / axis tick-label /
 * legend strikethrough readers exactly so a parsed value slots
 * straight back into the writer's emit path.
 */
export function parseDataLabelsStrikethrough(dLbls: XmlElement): boolean | undefined {
  const txPr = findChild(dLbls, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.strike
  if (raw === "sngStrike") return true
  return undefined
}

/**
 * Pull `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbls>` off a
 * data-labels block. Returns the typeface string the labels were
 * authored with.
 *
 * The OOXML `<a:latin>` element carries the typeface name on
 * `CT_TextFont` (ECMA-376 Part 1, §21.1.2.3.7). The reader trims
 * surrounding whitespace and reports the trimmed typeface; empty /
 * whitespace-only `typeface` attributes and missing `<a:latin>`
 * elements both collapse to `undefined` so absence and the empty
 * form round-trip identically through the writer. Non-string
 * `typeface` tokens (defensive — the XML parser only ever surfaces
 * strings) likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the data-labels block omits `<c:txPr>`
 * entirely or the canonical `<a:p><a:pPr><a:defRPr><a:latin>` chain
 * is malformed at any link. Mirrors the chart-title / axis-title /
 * axis tick-label / legend font family readers exactly so a parsed
 * value slots straight back into the writer's emit path.
 */
export function parseDataLabelsFontFamily(dLbls: XmlElement): string | undefined {
  const txPr = findChild(dLbls, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const latin = findChild(defRPr, "latin")
  if (!latin) return undefined
  const raw = latin.attrs.typeface
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

/**
 * Pull `<c:dLbls><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:dLbls>` off a data-labels block. Returns
 * the 6-character uppercase hex string when the parser walks the full
 * chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme references
 * (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>`
 * all collapse to `undefined` — only the literal RGB triple
 * round-trips losslessly through {@link writeChart}. Non-solid fills
 * (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`, `<a:blipFill>`)
 * likewise drop to `undefined` so a round-trip never fabricates a fill
 * the writer cannot reproduce on emit. Malformed `val` tokens (wrong
 * length, non-hex characters) drop to `undefined` rather than fabricate
 * a value the writer would round-trip into a malformed `<a:srgbClr>`.
 *
 * The OOXML `<c:spPr>` block sits on `CT_DLbls` between `<c:numFmt>`
 * and `<c:txPr>` per the schema sequence (ECMA-376 Part 1,
 * §21.2.2.50). Distinct from {@link parseDataLabelsFontColor}: the
 * fill lives on `<c:dLbls><c:spPr>`, the font color lives on
 * `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>` — the two
 * readers walk disjoint paths so a caller can pin both knobs without
 * conflict. Mirrors the chart-title fill {@link parseTitleFillColor} /
 * axis-title fill {@link parseAxisTitleFillColor} / plot-area fill
 * {@link parsePlotAreaFillColor} / legend fill
 * {@link parseLegendFillColor} so a parsed value slots straight into
 * the writer-side {@link ChartDataLabels.fillColor}.
 */
export function parseDataLabelsFillColor(dLbls: XmlElement): ChartColor | undefined {
  return parseSpPrFill(dLbls)
}

/**
 * Pull `<c:dLbls><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:dLbls>` off a data-labels block.
 * Returns the data-labels border (line) color as a 6-character
 * uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the line's
 * solid fill choice (`CT_LineProperties`'s solid fill — §20.1.2.3.24).
 * The `<a:ln>` slot sits inside the `<c:spPr>` block on `<c:dLbls>`
 * alongside the optional `<a:solidFill>` fill child, in
 * `CT_ShapeProperties` schema order (fill before stroke). The
 * `<c:spPr>` itself sits between `<c:numFmt>` and `<c:txPr>` per
 * CT_DLbls (ECMA-376 Part 1, §21.2.2.50).
 *
 * The reader surfaces only the literal `<a:srgbClr>` form — absence,
 * non-solid line fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>`),
 * and theme-color references (`<a:schemeClr>`) all collapse to
 * `undefined` so a chart that pinned a stroke the writer cannot
 * reproduce on emit drops the field rather than fabricate one Excel
 * would render differently. Malformed `val` tokens (wrong length,
 * non-hex characters, alpha-channel forms, non-string escapes)
 * likewise drop to `undefined`.
 *
 * Mirrors the writer-side {@link ChartDataLabels.borderColor} so a
 * parsed value slots straight into {@link cloneChart} without
 * conversion. The lookup is scoped to direct children of `<c:dLbls>`
 * so a stray `<c:spPr>` elsewhere (e.g. on `<c:plotArea>` /
 * `<c:legend>` / `<c:title>` / a series) cannot leak in. Mirrors
 * {@link parseDataLabelsFillColor} — same `<c:spPr>` host element on
 * the same `<c:dLbls>` parent — but lands on the line
 * (`<a:ln><a:solidFill>`) child rather than the fill (`<a:solidFill>`)
 * child.
 */
export function parseDataLabelsBorderColor(dLbls: XmlElement): ChartColor | undefined {
  return parseSpPrBorderColor(dLbls)
}

/**
 * Pull `<c:numFmt formatCode=".." sourceLinked=".."/>` off a
 * `<c:dLbls>` block. Returns `undefined` when the element is absent or
 * when `formatCode` is missing / empty (the OOXML schema requires the
 * attribute on every emitted `<c:numFmt>` so a fabricated empty record
 * cannot round-trip cleanly).
 *
 * `sourceLinked` accepts the same OOXML truthy / falsy spellings the
 * other boolean attributes do (`"1"` / `"true"` / `"0"` / `"false"`);
 * absence and the OOXML default `"0"` collapse to `undefined` so the
 * parsed shape stays minimal — only an explicit `"1"` / `"true"`
 * surfaces `true`. Mirrors how the axis-side numFmt parser shapes its
 * output.
 */
export function parseDataLabelsNumberFormat(el: XmlElement): ChartAxisNumberFormat | undefined {
  const numFmt = findChild(el, "numFmt")
  if (!numFmt) return undefined
  const formatCode = numFmt.attrs.formatCode
  if (typeof formatCode !== "string" || formatCode.length === 0) return undefined
  const out: ChartAxisNumberFormat = { formatCode }
  const sourceLinked = numFmt.attrs.sourceLinked
  if (typeof sourceLinked === "string") {
    if (sourceLinked === "1" || sourceLinked.toLowerCase() === "true") {
      out.sourceLinked = true
    }
  }
  return out
}

// ── Writer ────────────────────────────────────────────────────────

/**
 * Resolve and emit the `<c:dLbls>` element for a single series.
 *
 * Series override semantics:
 *
 * - Series sets `dataLabels: false`  → emit a `delete=1` block to
 *   suppress this series even when the chart-level default enables labels.
 * - Series sets `dataLabels: <obj>`  → emit `<obj>`. Chart-level config is ignored.
 * - Series omits `dataLabels`        → no per-series `<c:dLbls>`. Excel
 *   inherits the chart-type-level `<c:dLbls>` block emitted by
 *   `buildChartLevelDataLabels` instead.
 *
 * Returns `undefined` when nothing should be emitted at the series level.
 */
export function buildSeriesDataLabels(
  seriesDLbls: ChartDataLabels | false | undefined,
  chartDLbls: ChartDataLabels | undefined,
  chartType: WriteChartKind,
): string | undefined {
  if (seriesDLbls === false) {
    // Suppress this series even when chart-level labels are on.
    return xmlElement("c:dLbls", undefined, [
      xmlElement("c:dLbl", undefined, [
        xmlSelfClose("c:idx", { val: 0 }),
        xmlSelfClose("c:delete", { val: 1 }),
      ]),
      xmlSelfClose("c:delete", { val: 1 }),
    ])
  }
  if (seriesDLbls) {
    return buildDataLabelsBody(seriesDLbls, chartType)
  }
  // Series doesn't override → fall through to chart-level. Returning
  // undefined here keeps the chart-level <c:dLbls> as the single source
  // of truth so we don't duplicate the same toggles N times.
  void chartDLbls
  return undefined
}

/**
 * Build the chart-type-level `<c:dLbls>` block from
 * {@link SheetChart.dataLabels}. Returns `undefined` when no chart-level
 * labels are configured.
 */
export function buildChartLevelDataLabels(chart: SheetChart): string | undefined {
  if (!chart.dataLabels) return undefined
  return buildDataLabelsBody(chart.dataLabels, chart.type)
}

/**
 * Render the OOXML `<c:dLbls>` body. Element order follows CT_DLbls:
 * delete? before numFmt? before spPr? before txPr? before dLblPos? before
 * showLegendKey, showVal, showCatName, showSerName, showPercent,
 * showBubbleSize, separator?, showLeaderLines? — toggles must appear
 * in that exact order or Excel ignores the block.
 *
 * `chartType` lets the builder gate the pie / doughnut-only
 * `<c:showLeaderLines>` element — the OOXML schema scopes that flag to
 * `EG_DLbls` for `CT_PieChart` / `CT_DoughnutChart` only, so the writer
 * silently drops `dl.showLeaderLines` on every other family rather than
 * emit a child Excel's strict validator would reject.
 */
export function buildDataLabelsBody(dl: ChartDataLabels, chartType: WriteChartKind): string {
  const children: string[] = []

  // `<c:numFmt>` sits at the head of the CT_DLbls sequence (before
  // `<c:spPr>` / `<c:txPr>` / `<c:dLblPos>` / the show* toggles). The
  // writer skips emission entirely when the caller leaves `numberFormat`
  // unset so a fresh chart matches Excel's reference shape (no number
  // override means Excel inherits from the source cells).
  const numFmt = resolveDataLabelsNumberFormat(dl.numberFormat)
  if (numFmt) {
    children.push(
      xmlSelfClose("c:numFmt", {
        formatCode: numFmt.formatCode,
        sourceLinked: numFmt.sourceLinked === true ? 1 : 0,
      }),
    )
  }

  // `<c:spPr>` sits between `<c:numFmt>` and `<c:txPr>` in the
  // CT_DLbls schema sequence (ECMA-376 Part 1, §21.2.2.50). The writer
  // authors `<a:solidFill>` here from {@link ChartDataLabels.fillColor}
  // and `<a:ln>` here from {@link ChartDataLabels.borderColor}; other
  // `CT_ShapeProperties` children (gradient / pattern fills, line dash
  // / width / compound styles, `<a:effectLst>` effects) are not
  // modelled at this layer. The builder skips emission entirely when
  // no fill / border is pinned so a fresh chart matches Excel's
  // reference shape (no `<c:spPr>` block on a theme-default data-
  // labels rendering — typically a transparent label background with
  // no border). When at least one knob lands on the wire, the children
  // are emitted in `CT_ShapeProperties` schema order: `<a:solidFill>`
  // (fill) then `<a:ln>` (line / stroke). Distinct from the
  // `<a:defRPr><a:solidFill>` font-color slot inside `<c:txPr>` that
  // {@link ChartDataLabels.fontColor} pins — the three knobs target
  // different children of `<c:dLbls>` so a caller can pin them all
  // without conflict.
  const spPrXml = buildDataLabelsSpPr(
    resolveDataLabelsFillColor(dl.fillColor),
    resolveDataLabelsBorderColor(dl.borderColor),
    clampStrokeWidthPt(dl.borderWidth),
    normalizeBorderDash(dl.borderDash),
    normalizeLineCap(dl.borderCap),
    normalizeLineCompound(dl.borderCompound),
  )
  if (spPrXml !== undefined) {
    children.push(spPrXml)
  }

  // CT_DLbls schema places `<c:txPr>` between `<c:spPr>` and
  // `<c:dLblPos>` (ECMA-376 Part 1, §21.2.2.50). The block currently
  // carries the data-label font size, font color, bold flag, italic
  // flag, underline flag, strikethrough flag, and font family — every
  // typography pin lands on the same `<a:defRPr>` slot. The writer
  // skips the entire block when no font knob is pinned so a fresh
  // chart matches Excel's reference shape.
  const txPrXml = buildDataLabelsTxPr(
    resolveDataLabelsFontSize(dl.fontSize),
    resolveDataLabelsFontColor(dl.fontColor),
    resolveDataLabelsBold(dl.bold),
    resolveDataLabelsItalic(dl.italic),
    resolveDataLabelsUnderline(dl.underline),
    resolveDataLabelsStrikethrough(dl.strikethrough),
    resolveDataLabelsFontFamily(dl.fontFamily),
  )
  if (txPrXml !== undefined) {
    children.push(txPrXml)
  }

  if (dl.position) {
    children.push(xmlSelfClose("c:dLblPos", { val: dl.position }))
  }

  // OOXML requires showLegendKey to appear first when any toggle is set.
  // Always emit it explicitly so the rendered XML is deterministic.
  // Non-boolean inputs collapse to `false` to keep the on-the-wire output
  // stable, mirroring how the other `show*` toggles treat their inputs.
  children.push(xmlSelfClose("c:showLegendKey", { val: dl.showLegendKey === true ? 1 : 0 }))
  children.push(xmlSelfClose("c:showVal", { val: dl.showValue ? 1 : 0 }))
  children.push(xmlSelfClose("c:showCatName", { val: dl.showCategoryName ? 1 : 0 }))
  children.push(xmlSelfClose("c:showSerName", { val: dl.showSeriesName ? 1 : 0 }))
  children.push(xmlSelfClose("c:showPercent", { val: dl.showPercent ? 1 : 0 }))
  children.push(xmlSelfClose("c:showBubbleSize", { val: 0 }))

  if (dl.separator !== undefined) {
    children.push(xmlElement("c:separator", undefined, xmlEscape(dl.separator)))
  }

  // `<c:showLeaderLines>` sits at the tail of the `EG_DLbls` group
  // (after `<c:separator>`, before `<c:extLst>`). The OOXML schema
  // scopes the element to pie / doughnut chart families exclusively
  // (`EG_DLbls` for `CT_PieChart` / `CT_DoughnutChart` only — bar /
  // column / line / area / scatter route through `EG_DLblsShared` which
  // omits it). The writer drops the field silently on every non-pie /
  // non-doughnut family to mirror Excel's reference serialization.
  //
  // The OOXML default is `true` (Excel paints leader lines on every
  // label that gets pushed outside its slice). Only an explicit
  // `false` flips the toggle, so absence and the default round-trip
  // identically through {@link parseChart}. Non-boolean inputs collapse
  // to the default, mirroring how the other `show*` toggles treat
  // their inputs.
  if ((chartType === "pie" || chartType === "doughnut") && dl.showLeaderLines === false) {
    children.push(xmlSelfClose("c:showLeaderLines", { val: 0 }))
  }

  return xmlElement("c:dLbls", undefined, children)
}

/**
 * Resolve the `<c:numFmt>` value emitted inside `<c:dLbls>`.
 *
 * Returns `undefined` when the caller leaves `numberFormat` unset or
 * when `formatCode` is missing / non-string / empty — the OOXML schema
 * requires `formatCode` on every emitted `<c:numFmt>` so a malformed
 * record is dropped rather than fabricate a placeholder Excel rejects.
 * Mirrors how the axis-side number-format pipeline shapes its output
 * so a parsed value can flow straight from the read side back into the
 * write side without transformation.
 *
 * `sourceLinked` is normalized to a literal boolean — only `true`
 * survives, every other shape (`undefined` / `false` / non-boolean)
 * collapses so the writer's `val` attribute defaults to `0`.
 */
export function resolveDataLabelsNumberFormat(
  value: ChartAxisNumberFormat | undefined,
): ChartAxisNumberFormat | undefined {
  if (!value) return undefined
  const formatCode = value.formatCode
  if (typeof formatCode !== "string" || formatCode.length === 0) return undefined
  const out: ChartAxisNumberFormat = { formatCode }
  if (value.sourceLinked === true) out.sourceLinked = true
  return out
}

/**
 * Resolve `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
 * </a:p></c:txPr></c:dLbls>` from {@link ChartDataLabels.fontSize}.
 *
 * Returns the size in points (`1..400`), or `undefined` when the
 * caller leaves the field unset / passed an out-of-range or
 * non-numeric token. Delegates to {@link normalizeTitleFontSize} so
 * the chart-title / axis-title / axis tick-label / legend / data-label
 * font-size resolvers share the same range, the same fractional
 * rounding (Excel's UI step is 0.5pt), and the same OOXML conversion
 * (100ths of a point at emit time).
 */
export function resolveDataLabelsFontSize(value: number | undefined): number | undefined {
  return normalizeTitleFontSize(value)
}

/**
 * Resolve `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:txPr></c:dLbls>` from {@link ChartDataLabels.fontColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches the chart-title /
 * axis-title / axis tick-label / legend color resolvers exactly.
 */
export function resolveDataLabelsFontColor(value: ChartColor | undefined): ChartColor | undefined {
  return normalizeTitleColor(value)
}

/**
 * Resolve `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
 * </a:p></c:txPr></c:dLbls>` from {@link ChartDataLabels.bold}.
 *
 * Returns the bold flag, or `undefined` when the caller leaves the
 * field unset / passed a non-boolean token. Mirrors the chart-title /
 * axis-title / axis tick-label / legend bold resolvers exactly —
 * only literal `true` / `false` pass through; non-boolean tokens
 * (typed escapes from an untyped caller) collapse to `undefined`.
 */
export function resolveDataLabelsBold(value: boolean | undefined): boolean | undefined {
  if (value === true) return true
  if (value === false) return false
  return undefined
}

/**
 * Resolve `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:txPr></c:dLbls>` from {@link ChartDataLabels.italic}.
 *
 * Returns the italic flag, or `undefined` when the caller leaves the
 * field unset / passed a non-boolean token. Mirrors the chart-title /
 * axis-title / axis tick-label / legend italic resolvers exactly —
 * only literal `true` / `false` pass through; non-boolean tokens
 * (typed escapes from an untyped caller) collapse to `undefined`.
 */
export function resolveDataLabelsItalic(value: boolean | undefined): boolean | undefined {
  if (value === true) return true
  if (value === false) return false
  return undefined
}

/**
 * Resolve `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
 * </a:p></c:txPr></c:dLbls>` from {@link ChartDataLabels.underline}.
 *
 * Returns the underline flag, or `undefined` when the caller leaves
 * the field unset / passed a non-boolean token. Mirrors the
 * chart-title / axis-title / axis tick-label / legend underline
 * resolvers exactly — only literal `true` / `false` pass through;
 * non-boolean tokens (typed escapes from an untyped caller) collapse
 * to `undefined`. The writer translates `true` into `u="sng"` (Excel's
 * UI variant — single underline) and `false` into `u="none"` at emit
 * time.
 */
export function resolveDataLabelsUnderline(value: boolean | undefined): boolean | undefined {
  if (value === true) return true
  if (value === false) return false
  return undefined
}

/**
 * Resolve `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr strike=".."/>
 * </a:pPr></a:p></c:txPr></c:dLbls>` from
 * {@link ChartDataLabels.strikethrough}.
 *
 * Returns `true` when the caller pins the strikethrough flag
 * literally; every other value (explicit `false`, absence, non-boolean
 * tokens leaking past the type guard) collapses to `undefined` so the
 * writer never emits a `strike` attribute below `"sngStrike"`. The
 * OOXML default `"noStrike"` is functionally identical to absence —
 * the writer keeps the surfaced shape consistent with what Excel's
 * UI authors (`"sngStrike"` only, never `"noStrike"` or
 * `"dblStrike"`), mirroring how `resolveTitleStrike` /
 * `resolveLegendStrikethrough` land on their `<a:defRPr>` slots.
 */
export function resolveDataLabelsStrikethrough(value: boolean | undefined): boolean | undefined {
  if (value === true) return true
  return undefined
}

/**
 * Resolve `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbls>` from
 * {@link ChartDataLabels.fontFamily}.
 *
 * Returns the trimmed typeface string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed an
 * empty / whitespace-only / non-string token. Mirrors the chart-title
 * / axis-title / axis tick-label / legend font family resolvers
 * exactly so a single configuration call threads cleanly through
 * every typography slot Excel exposes.
 */
export function resolveDataLabelsFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

/**
 * Resolve `<c:dLbls><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:dLbls>` from {@link ChartDataLabels.fillColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches the chart-title /
 * axis-title / plot-area / legend / data-label color resolvers
 * exactly. Distinct from {@link resolveDataLabelsFontColor}: the fill
 * lives on `<c:dLbls><c:spPr>`, the font color lives on
 * `<c:dLbls><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>` — the two
 * resolvers feed disjoint slots so a caller can pin both without
 * conflict.
 */
export function resolveDataLabelsFillColor(value: ChartColor | undefined): ChartColor | undefined {
  return normalizeTitleColor(value)
}

/**
 * Resolve `<c:dLbls><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:dLbls>` from
 * {@link ChartDataLabels.borderColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches every other `<a:srgbClr>`
 * fill / line slot exactly. Composes independently with
 * {@link ChartDataLabels.fillColor} — the two knobs share the same
 * `<c:spPr>` host on `<c:dLbls>` but land on different children
 * (`<a:solidFill>` for the fill, `<a:ln><a:solidFill>` for the
 * stroke). Mirrors the chart-title / axis-title / chart-space /
 * plot-area / legend / data-table `<c:spPr>` border slots — same hex
 * grammar, same `<a:ln>` slot on the `CT_ShapeProperties` schema —
 * but lands on `<c:dLbls>`'s own `<c:spPr>` block.
 */
export function resolveDataLabelsBorderColor(
  value: ChartColor | undefined,
): ChartColor | undefined {
  return normalizeTitleColor(value)
}

/**
 * Build the optional `<c:spPr>` block inside `<c:dLbls>`. Surfaces the
 * solid fill color knob ({@link ChartDataLabels.fillColor}) and the
 * border (line) color knob ({@link ChartDataLabels.borderColor}) —
 * every other `<c:spPr>` child (`<a:effectLst>` effects, gradient /
 * pattern / picture fills, line dash / width / compound styles) is
 * intentionally not modelled at this layer.
 *
 * Returns `undefined` when both fields are unset / malformed so the
 * writer skips the entire `<c:spPr>` block — an empty `<c:spPr/>`
 * collapses to the inherited theme fill / stroke Excel picks anyway,
 * and omitting it keeps untouched chart XML byte-clean. When at least
 * one knob lands on the wire, the children are emitted in
 * `CT_ShapeProperties` schema order: `<a:solidFill>` (fill) then
 * `<a:ln>` (line / stroke).
 *
 * The emitted block mirrors the minimal `<c:spPr>` shape Excel writes
 * when the user pins "Format Data Labels -> Fill -> Solid fill ->
 * Color" and / or "Format Data Labels -> Border -> Solid line ->
 * Color": `<c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr>`. The `val` attribute holds the
 * canonical 6-character uppercase hex form (the writer normalizes the
 * input ahead of this call so a malformed source value never reaches
 * emit).
 *
 * Mirrors {@link buildLegendSpPr} / {@link buildDataTableSpPr} but
 * on a distinct host element — the data-labels fill / stroke paint
 * the background and outline of each label box, while the legend /
 * data-table variants paint a different chart-frame element.
 */
export function buildDataLabelsSpPr(
  fillRgbHex: ChartColor | undefined,
  borderRgbHex: ChartColor | undefined,
  borderWidthPt: number | undefined,
  borderDash: ChartBorderDash | undefined,
  borderCap?: ChartLineCap | undefined,
  borderCompound?: ChartLineCompound | undefined,
): string | undefined {
  if (
    fillRgbHex === undefined &&
    borderRgbHex === undefined &&
    borderWidthPt === undefined &&
    borderDash === undefined &&
    borderCap === undefined &&
    borderCompound === undefined
  ) {
    return undefined
  }
  const children: string[] = []
  if (fillRgbHex !== undefined) {
    children.push(xmlElement("a:solidFill", undefined, [buildColorElement(fillRgbHex)]))
  }
  if (
    borderRgbHex !== undefined ||
    borderWidthPt !== undefined ||
    borderDash !== undefined ||
    borderCap !== undefined ||
    borderCompound !== undefined
  ) {
    const lnAttrs: Record<string, string | number> = {}
    if (borderWidthPt !== undefined) {
      // OOXML stores stroke width in EMU (1 pt = 12 700 EMU). Round to
      // the nearest integer because the schema types `w` as `xsd:int`.
      lnAttrs.w = Math.round(borderWidthPt * EMU_PER_PT)
    }
    if (borderCap !== undefined) lnAttrs.cap = borderCap
    if (borderCompound !== undefined) lnAttrs.cmpd = borderCompound
    const lnChildren: string[] = []
    if (borderRgbHex !== undefined) {
      lnChildren.push(xmlElement("a:solidFill", undefined, [buildColorElement(borderRgbHex)]))
    }
    // `<a:prstDash>` follows `<a:solidFill>` per CT_LineProperties
    // schema sequence (ECMA-376 Part 1, §20.1.2.3.24).
    if (borderDash !== undefined) {
      lnChildren.push(xmlSelfClose("a:prstDash", { val: borderDash }))
    }
    children.push(
      lnChildren.length === 0
        ? xmlSelfClose("a:ln", lnAttrs)
        : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren),
    )
  }
  return xmlElement("c:spPr", undefined, children)
}

/**
 * Build the `<c:txPr>` block that carries a data-label's typography
 * pins. Returns `undefined` when every input is unset so the caller
 * can elide the element entirely (Excel's reference serialization
 * omits `<c:txPr>` from `<c:dLbls>` when the labels render at the
 * theme-default style).
 *
 * The emitted block mirrors the minimal `<c:txPr>` shape Excel writes
 * when the user pins a data-label typography knob — `<a:bodyPr/>` (no
 * rotation because the data-label rotation is parked in a separate
 * extension element Excel emits at write time), `<a:lstStyle/>` is
 * the empty list-style placeholder the schema requires, and the
 * `<a:p><a:pPr><a:defRPr ...><a:solidFill>...</a:solidFill></a:defRPr>
 * </a:pPr><a:endParaRPr/></a:p>` paragraph stub Excel always emits
 * hosts the typography attributes on `<a:defRPr>`. The `<a:defRPr>`
 * element expands from self-closing to wrapping a single
 * `<a:solidFill>` child when a color is set; otherwise the writer
 * keeps the existing self-closing form so a fresh chart with no
 * custom color matches Excel's reference serialization byte-for-byte.
 * The bold and italic flags each emit a literal `b="1"` / `b="0"`
 * (or `i="1"` / `i="0"`) whenever the input is a boolean — `false`
 * pins the OOXML default explicitly, which is functionally identical
 * to absence but lets a clone target override an upstream `b="1"` /
 * `i="1"` from a templated chart. The underline flag emits `u="sng"`
 * (single underline — Excel's UI variant) for `true` and `u="none"`
 * (the OOXML default) for `false`, with the same override semantics
 * as bold and italic. The strikethrough flag rides as
 * `strike="sngStrike"` on the same `<a:defRPr>` slot when the input
 * is `true`; absence (and explicit `false`, which the resolver
 * collapses to `undefined`) skips the attribute entirely since the
 * OOXML default `"noStrike"` is functionally identical to absence.
 * Mirrors the chart-title / axis-title / axis tick-label / legend
 * `<c:txPr>` slots exactly so a re-parse picks the value off the
 * canonical default-paragraph slot every other typography reader
 * expects.
 */
export function buildDataLabelsTxPr(
  fontSizePt: number | undefined,
  rgbHex: ChartColor | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
  underline: boolean | undefined,
  strikethrough: boolean | undefined,
  fontFamily: string | undefined,
): string | undefined {
  if (
    fontSizePt === undefined &&
    rgbHex === undefined &&
    bold === undefined &&
    italic === undefined &&
    underline === undefined &&
    strikethrough === undefined &&
    fontFamily === undefined
  )
    return undefined
  const defRPrAttrs: Record<string, string | number> = {}
  if (fontSizePt !== undefined) defRPrAttrs.sz = fontSizePt * TITLE_FONT_SZ_PER_POINT
  if (bold !== undefined) defRPrAttrs.b = bold ? 1 : 0
  if (italic !== undefined) defRPrAttrs.i = italic ? 1 : 0
  if (underline !== undefined) defRPrAttrs.u = underline ? "sng" : "none"
  // Strikethrough rides as `strike="sngStrike"` on the same
  // `<a:defRPr>` slot. Absence collapses to omitting the attribute
  // entirely (the OOXML default `"noStrike"` is functionally
  // identical to absence — the reader collapses both to `undefined`).
  // The writer never emits `"noStrike"` or `"dblStrike"` so the
  // surfaced shape stays consistent with Excel's UI checkbox.
  if (strikethrough === true) defRPrAttrs.strike = "sngStrike"
  // OOXML's `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:defRPr>` carries the data-label font color.
  // Absence (`undefined`) collapses to skipping the `<a:solidFill>`
  // child entirely so the labels inherit the theme text color
  // (Excel's reference behavior for fresh data labels that have not
  // had a custom color picked).
  const solidFillChild =
    rgbHex !== undefined
      ? xmlElement("a:solidFill", undefined, [buildColorElement(rgbHex)])
      : undefined
  // OOXML's `<a:defRPr><a:latin typeface=".."/></a:defRPr>` carries
  // the data-label font family. The `<a:latin>` element follows
  // `<a:solidFill>` per the CT_TextCharacterProperties child sequence
  // (ECMA-376 Part 1, §21.1.2.3.7). Absence (`undefined`) collapses
  // to omitting the entire `<a:latin>` element so the labels inherit
  // the theme typeface (Excel's reference behavior for fresh data
  // labels that have not had a custom font picked).
  const latinChild = fontFamily ? xmlSelfClose("a:latin", { typeface: fontFamily }) : undefined
  // When a fill color or a typeface is set the `<a:defRPr>` slot
  // expands from self-closing to wrapping the children; otherwise the
  // writer keeps the existing self-closing form so a fresh chart with
  // no custom color or font matches Excel's reference serialization
  // byte-for-byte. Children are emitted in
  // CT_TextCharacterProperties' canonical schema order: solidFill
  // first, then latin.
  const defRPrChildren: string[] = []
  if (solidFillChild) defRPrChildren.push(solidFillChild)
  if (latinChild) defRPrChildren.push(latinChild)
  const defRPr =
    defRPrChildren.length > 0
      ? xmlElement("a:defRPr", defRPrAttrs, defRPrChildren)
      : xmlSelfClose("a:defRPr", defRPrAttrs)
  return xmlElement("c:txPr", undefined, [
    xmlSelfClose("a:bodyPr"),
    xmlSelfClose("a:lstStyle"),
    xmlElement("a:p", undefined, [
      xmlElement("a:pPr", undefined, [defRPr]),
      xmlSelfClose("a:endParaRPr", { lang: "en-US" }),
    ]),
  ])
}

// ── Clone resolvers (3-arg source/override) ───────────────────────

/**
 * Resolve a chart-level data-labels override.
 *
 * `undefined` → inherit the source's parsed `dataLabels` (downcast from
 * the read-side {@link ChartDataLabelsInfo} to the write-side
 * {@link ChartDataLabels} shape — they share field semantics).
 * `null`      → drop the inherited block.
 * object      → replace.
 */
export function resolveChartDataLabels(
  sourceLabels: ChartDataLabelsInfo | undefined,
  override: ChartDataLabels | null | undefined,
): ChartDataLabels | undefined {
  if (override === undefined) {
    return sourceLabels ? { ...sourceLabels } : undefined
  }
  if (override === null) return undefined
  return override
}

/**
 * Resolve a per-series data-labels override.
 *
 * `undefined` → inherit the source series' `dataLabels`.
 * `null`      → drop the inherited block (series will fall back to
 *               whatever the chart-level default is at write time).
 * `false`     → suppress labels on this series alone.
 * object      → replace.
 */
export function resolveSeriesDataLabels(
  sourceLabels: ChartDataLabelsInfo | undefined,
  override: ChartDataLabels | false | null | undefined,
): ChartDataLabels | false | undefined {
  if (override === undefined) {
    return sourceLabels ? { ...sourceLabels } : undefined
  }
  if (override === null) return undefined
  return override
}
