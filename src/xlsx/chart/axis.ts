// ── Chart Axis (reader) ───────────────────────────────────────────
// Per-host module for the reader-side parsers of the four axis flavours
// — `<c:catAx>`, `<c:valAx>`, `<c:dateAx>`, `<c:serAx>` (CT_CatAx /
// CT_ValAx / CT_DateAx / CT_SerAx, ECMA-376 Part 1, §21.2.2.x). Holds
// every `parseAxis*` function for the per-axis block — including its
// nested `<c:title>` (axis-title typography, fill / border, manual
// layout, overlay), `<c:txPr>` tick-label typography, `<c:scaling>`
// (min / max / log base / orient), `<c:numFmt>`, `<c:majorGridlines>` /
// `<c:minorGridlines>`, `<c:majorTickMark>` / `<c:minorTickMark>`,
// `<c:tickLblPos>`, `<c:dispUnits>`, `<c:crosses>` / `<c:crossesAt>`,
// `<c:crossBetween>`, `<c:lblOffset>`, `<c:lblAlgn>`, `<c:noMultiLvlLbl>`,
// `<c:auto>`, and `<c:delete>` (hidden) children.
//
// The writer-side normalize / build helpers and the cloner-side resolve
// / apply override helpers stay in chart-writer.ts and chart-clone.ts
// respectively because they are heavily entangled with chart-level
// constants (axis-id allocation, render options) that would require a
// further round of refactoring to extract cleanly.

import type {
  Chart,
  ChartAxisCrossBetween,
  ChartAxisCrosses,
  ChartAxisDispUnit,
  ChartAxisDispUnits,
  ChartAxisGridlines,
  ChartAxisInfo,
  ChartAxisLabelAlign,
  ChartAxisNumberFormat,
  ChartAxisScale,
  ChartAxisTickLabelPosition,
  ChartAxisTickMark,
  ChartBorderDash,
  ChartColor,
  ChartLineCap,
  ChartLineCompound,
  ChartManualLayout,
  WriteChartKind,
} from "../../_types"
import type { CloneChartOptions } from "../chart-clone"
import type { XmlElement } from "../../xml/parser"
import {
  buildColorElement,
  normalizeRgbHex,
  parseBorderCapFromSpPr,
  parseBorderCompoundFromSpPr,
  parseBorderDashFromSpPr,
  parseBorderWidthFromSpPr,
  parseSchemeClr,
  parseSpPrBorderColor,
  parseSpPrFill,
  resolveBorderDash,
  resolveBorderWidthPt,
} from "./shape"
import { type ResolvedManualLayout, buildManualLayout, parseManualLayout } from "./layout"
import {
  applyOverride,
  childElements,
  collectTextRuns,
  elementText,
  findChild,
  parseBoolAttr,
  parseNumericChildVal,
} from "./util"
import {
  FONT_SIZE_MAX_PT,
  FONT_SIZE_MIN_PT,
  FONT_SZ_PER_POINT,
  ROTATION_MAX_DEG,
  ROTATION_MIN_DEG,
  TXPR_ROT_PER_DEGREE,
} from "./text"
import { xmlElement, xmlEscape, xmlSelfClose } from "../../xml/writer"
import {
  buildTitleSpPr,
  normalizeTitleBold,
  normalizeTitleColor,
  normalizeTitleFontSize,
  normalizeTitleItalic,
  normalizeTitleRotation,
  normalizeTitleStrike,
  normalizeTitleUnderline,
} from "./title"
import type { SheetChart } from "../../_types"
import { normalizeLegendLayout } from "./legend"

const TITLE_FONT_SZ_PER_POINT = FONT_SZ_PER_POINT
const TITLE_FONT_SIZE_MIN_PT = FONT_SIZE_MIN_PT
const TITLE_FONT_SIZE_MAX_PT = FONT_SIZE_MAX_PT

// ── Axis-scope enumerations ───────────────────────────────────────

/**
 * Recognized values of `<c:majorTickMark>` / `<c:minorTickMark>` per
 * the OOXML `ST_TickMark` enumeration.
 */
const VALID_TICK_MARKS: ReadonlySet<ChartAxisTickMark> = new Set(["none", "in", "out", "cross"])

/**
 * Recognized values of `<c:tickLblPos>` per the OOXML
 * `ST_TickLblPos` enumeration.
 */
const VALID_TICK_LBL_POSITIONS: ReadonlySet<ChartAxisTickLabelPosition> = new Set([
  "nextTo",
  "low",
  "high",
  "none",
])

/**
 * Recognized values of `<c:lblAlgn>` per the OOXML `ST_LblAlgn`
 * enumeration.
 */
const VALID_LBL_ALIGNS: ReadonlySet<ChartAxisLabelAlign> = new Set(["ctr", "l", "r"])

/**
 * Conversion factor between OOXML's `rot` attribute (60000ths of a
 * degree, the integer Excel writes inside `<a:bodyPr rot="N"/>`) and
 * whole degrees. Excel's UI exposes the -90..90 degree band — the
 * reader clamps anything outside that band so a corrupt template
 * cannot surface a value the writer would never emit.
 *
 * `TXPR_ROT_PER_DEGREE` / `ROTATION_MIN_DEG` / `ROTATION_MAX_DEG` are
 * imported from `./text` so every typography host shares the same
 * conversion factor. Aliased onto `LABEL_ROTATION_MIN_DEG` /
 * `LABEL_ROTATION_MAX_DEG` for parity with the original axis-tick-
 * label call sites.
 */
const LABEL_ROTATION_MIN_DEG = ROTATION_MIN_DEG
const LABEL_ROTATION_MAX_DEG = ROTATION_MAX_DEG

/** Recognized values of `<c:crosses>` per the OOXML `ST_Crosses` enum. */
const VALID_CROSSES: ReadonlySet<ChartAxisCrosses> = new Set(["autoZero", "min", "max"])

/** Recognized values of `<c:builtInUnit>` per the OOXML `ST_BuiltInUnit` enum. */
const VALID_DISP_UNITS: ReadonlySet<ChartAxisDispUnit> = new Set([
  "hundreds",
  "thousands",
  "tenThousands",
  "hundredThousands",
  "millions",
  "tenMillions",
  "hundredMillions",
  "billions",
  "trillions",
])

/** Recognized values of `<c:crossBetween>` per the OOXML `ST_CrossBetween` enum. */
const VALID_CROSS_BETWEEN: ReadonlySet<ChartAxisCrossBetween> = new Set(["between", "midCat"])

// ── Reader ────────────────────────────────────────────────────────

export function parseAxisInfo(
  axis: XmlElement,
  familyDefaultCrossBetween: ChartAxisCrossBetween,
): ChartAxisInfo | undefined {
  const title = parseAxisTitle(axis)
  // `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>` —
  // axis-title rotation in 60000ths of a degree. Sits on every axis
  // flavour per the OOXML schema (CT_CatAx, CT_ValAx, CT_DateAx,
  // CT_SerAx all share the same `<c:title>` shape). The lookup is
  // scoped to the `<c:title>` body so a stray `<a:bodyPr>` elsewhere
  // on the axis (e.g. on the tick-label `<c:txPr>`) cannot leak in.
  // Out-of-range values clamp to the `-90..90` band Excel's UI
  // exposes; the OOXML default `0` and absence both collapse to
  // `undefined`. Returns `undefined` when the axis omits `<c:title>`
  // entirely or when the title is a `<c:strRef>` (formula reference)
  // with no `<c:rich>` body.
  const axisTitleRotation = parseAxisTitleRotation(axis)
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title font size in 100ths of a
  // point. Same `<c:title>` body scope as `axisTitleRotation` so a
  // stray `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. The value comes back in points (range
  // `1..400`); out-of-range / non-numeric inputs drop to `undefined`,
  // and absence of `<c:title>` / `<c:rich>` / `<a:p>` / `<a:pPr>` /
  // `<a:defRPr>` / the `sz` attribute likewise collapses to
  // `undefined` for symmetry with the writer-side
  // {@link SheetChart.axes.x.axisTitleFontSize}.
  const axisTitleFontSize = parseAxisTitleFontSize(axis)
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title bold flag. Same `<c:title>`
  // body scope as `axisTitleRotation` so a stray `<a:defRPr>` elsewhere
  // on the axis (e.g. on the tick-label `<c:txPr>`) cannot leak in.
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through {@link cloneChart} — only
  // an explicit `b="1"` surfaces `true`. Returns `undefined` when the
  // axis omits `<c:title>` entirely or when the title is a `<c:strRef>`
  // (formula reference) with no `<c:rich>` body.
  const axisTitleBold = parseAxisTitleBold(axis)
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title italic flag. Same
  // `<c:title>` body scope as `axisTitleFontSize`, so a stray
  // `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. The OOXML default `false` collapses to
  // `undefined` so absence and `i="0"` round-trip identically through
  // the writer — only an explicit `i="1"` surfaces `true`. Returns
  // `undefined` when the axis omits `<c:title>` entirely or when the
  // title is a `<c:strRef>` (formula reference) with no `<c:rich>`
  // body.
  const axisTitleItalic = parseAxisTitleItalic(axis)
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
  // <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title font color. Same
  // `<c:title>` body scope as `axisTitleItalic`, so a stray
  // `<a:solidFill>` elsewhere on the axis (e.g. on a tick-label
  // `<c:txPr>` block or a `<c:spPr>` series fill) cannot leak in.
  // Theme references (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`,
  // `<a:prstClr>`, and malformed `val` tokens all collapse to
  // `undefined` since only the literal RGB triple round-trips
  // losslessly through the writer. Returns `undefined` when the axis
  // omits `<c:title>` entirely or when the title is a `<c:strRef>`
  // (formula reference) with no `<c:rich>` body.
  const axisTitleColor = parseAxisTitleColor(axis)
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` — axis-title strikethrough flag.
  // Same `<c:title>` body scope as `axisTitleItalic`, so a stray
  // `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. Only the UI-default `"sngStrike"`
  // surfaces as `true`; `"noStrike"` (the OOXML application default)
  // and the non-UI `"dblStrike"` both collapse to `undefined` so absence
  // and the OOXML default round-trip identically through the writer
  // (which emits only `"sngStrike"`). Returns `undefined` when the axis
  // omits `<c:title>` entirely or when the title is a `<c:strRef>`
  // (formula reference) with no `<c:rich>` body.
  const axisTitleStrike = parseAxisTitleStrike(axis)
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` — axis-title underline flag.
  // Same `<c:title>` body scope as `axisTitleStrike`, so a stray
  // `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. Only the UI-default `"sng"` surfaces
  // as `true`; `"none"` (the OOXML application default), the non-UI
  // `"dbl"` variant, and the sixteen exotic tokens all collapse to
  // `undefined` so absence and the OOXML default round-trip
  // identically through the writer (which emits only `"sng"`). Returns
  // `undefined` when the axis omits `<c:title>` entirely or when the
  // title is a `<c:strRef>` (formula reference) with no `<c:rich>`
  // body.
  const axisTitleUnderline = parseAxisTitleUnderline(axis)
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
  // typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>` —
  // axis-title font family. Same `<c:title>` body scope as
  // `axisTitleColor` / `axisTitleStrike` / `axisTitleUnderline`, so a
  // stray `<a:latin>` elsewhere on the axis (e.g. on the tick-label
  // `<c:txPr>`) cannot leak in. Empty / whitespace-only `typeface`
  // attributes and missing `<a:latin>` elements both collapse to
  // `undefined` so absence and the empty form round-trip identically
  // through the writer. Returns `undefined` when the axis omits
  // `<c:title>` entirely or when the title is a `<c:strRef>` (formula
  // reference) with no `<c:rich>` body.
  const axisTitleFontFamily = parseAxisTitleFontFamily(axis)
  // `<c:title><c:overlay val=".."/></c:title>` — axis-title overlay
  // flag. Sits as a direct child of `<c:title>` per CT_Title schema,
  // so the lookup is scoped to direct title children. The OOXML
  // default `false` collapses to `undefined` so absence and
  // `<c:overlay val="0"/>` round-trip identically through
  // {@link cloneChart}. Returns `undefined` when the axis omits
  // `<c:title>` entirely.
  const axisTitleOverlay = parseAxisTitleOverlay(axis)
  // `<c:title><c:layout><c:manualLayout>...</c:manualLayout></c:layout>
  // </c:title>` — axis-title manual placement. Sits inside `<c:title>`
  // between `<c:tx>` and `<c:overlay>` per CT_Title schema (ECMA-376
  // Part 1, §21.2.2.210), so the lookup is scoped to the `<c:layout>`
  // child of `<c:title>` (a stray `<c:layout>` elsewhere on the axis
  // — e.g. on `<c:plotArea>` — cannot leak in). The OOXML
  // `<c:manualLayout>` block (`CT_ManualLayout`, §21.2.2.115) carries
  // the `(x, y)` anchor and `(w, h)` size as fractions of the chart
  // frame in the `0..1` band; out-of-range / non-finite / non-numeric
  // tokens drop on the matching axis so absence and a malformed token
  // round-trip identically through {@link cloneChart}. Returns
  // `undefined` whenever the axis omits the `<c:title>` /
  // `<c:layout>` / `<c:manualLayout>` chain at any link or when every
  // coordinate dropped on normalization. Mirrors the chart-level
  // `legendLayout` / `plotAreaLayout` parsers — same accept-or-drop
  // grammar, same `xMode="edge"` / `xMode="factor"` admission.
  const axisTitleLayout = parseAxisTitleLayout(axis)
  // `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></c:spPr></c:title>` — axis-title background fill.
  // Sits on the axis's `<c:title>` directly per CT_Title schema (the
  // `<c:spPr>` block follows `<c:overlay>` in the schema sequence,
  // ECMA-376 Part 1, §21.2.2.210). The reader surfaces only the
  // literal `<a:srgbClr val="RRGGBB"/>` form; theme references
  // (`<a:schemeClr>`), non-solid fills (`<a:noFill>` / `<a:gradFill>` /
  // `<a:pattFill>` / `<a:blipFill>`), and malformed `val` tokens all
  // collapse to `undefined` so a round-trip never fabricates a fill the
  // writer cannot reproduce on emit. Independent of `axisTitleColor`
  // (which lives on the inner `<a:defRPr><a:solidFill>` slot for the
  // font color) — the two readers walk disjoint paths so a caller can
  // pin both knobs without conflict. Returns `undefined` when the axis
  // omits `<c:title>` entirely; unlike `axisTitleColor`, the lookup is
  // not gated on `<c:rich>` so a title authored as a `<c:strRef>`
  // formula reference can still surface its background fill — Excel's
  // "Format Axis Title -> Fill" dialog is independent of whether the
  // text body is rich or a formula.
  const axisTitleFillColor = parseAxisTitleFillColor(axis)
  // `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:ln></c:spPr></c:title>` — axis-title border
  // (line stroke) color. Sits on the axis's `<c:title>` directly per
  // CT_Title schema; the `<a:ln>` block lives inside the same
  // `<c:spPr>` slot as the fill (`<a:solidFill>`), per
  // CT_ShapeProperties — the reader scopes the lookup to direct
  // children of the axis's `<c:title>` so a stray `<c:spPr>`
  // elsewhere (on the plot area, a series, on the legend, on the
  // chart-level title) cannot leak into this field. Theme references
  // (`<a:schemeClr>`) and non-solid line fills (`<a:noFill>` /
  // `<a:gradFill>` / `<a:pattFill>`) all collapse to `undefined` so
  // a round-trip never fabricates a stroke the writer cannot
  // reproduce on emit. Independent of `axisTitleFillColor` (which
  // lives on `<c:spPr><a:solidFill>` — the fill child of the same
  // `<c:spPr>` block) and `axisTitleColor` (which lives on the
  // inner `<a:defRPr><a:solidFill>` slot for the font color) — the
  // three readers walk disjoint paths so a caller can pin all three
  // knobs without conflict.
  const axisTitleBorderColor = parseAxisTitleBorderColor(axis)
  // `<c:catAx><c:title><c:spPr><a:ln w="EMU"/>` (or `<c:valAx>` /
  // `<c:dateAx>` / `<c:serAx>`) carries Excel's "Format Axis Title ->
  // Border -> Width" pin. Same EMU encoding and clamp / snap grammar
  // as every other chart-frame border-width slot. Independent of
  // `axisTitleBorderColor` (color child) and `axisTitleBorderDash`
  // (`<a:prstDash>` child) — the three readers walk disjoint slots of
  // the shared `<a:ln>` element so a caller can pin all three knobs
  // without conflict.
  const axisTitleBorderWidth = parseAxisTitleBorderWidth(axis)
  // `<c:catAx><c:title><c:spPr><a:ln><a:prstDash val=".."/>` (or
  // `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`) carries Excel's "Format
  // Axis Title -> Border -> Dash type" pin.
  const axisTitleBorderDash = parseAxisTitleBorderDash(axis)
  // `<c:title><c:spPr><a:ln cap=".."/>` / `<a:ln cmpd=".."/>` — Excel's
  // axis-title border line cap and compound styles. Same accept-or-drop
  // grammar as every other chart-frame `<a:ln>` slot.
  const axisTitleBorderCap = parseAxisTitleBorderCap(axis)
  const axisTitleBorderCompound = parseAxisTitleBorderCompound(axis)
  const gridlines = parseAxisGridlines(axis)
  const scale = parseAxisScale(axis)
  const numberFormat = parseAxisNumberFormat(axis)
  // Tick-mark and tick-label-position children sit alongside the
  // gridlines / numFmt on every CT_CatAx / CT_ValAx / CT_DateAx /
  // CT_SerAx — see CT_TickMark, ST_TickMark, ST_TickLblPos in
  // ECMA-376 Part 1, §21.2.2. The reader collapses each value to
  // `undefined` when it matches the OOXML default so absence and the
  // default round-trip identically through {@link cloneChart}.
  const majorTickMark = parseAxisTickMark(axis, "majorTickMark", "out")
  const minorTickMark = parseAxisTickMark(axis, "minorTickMark", "none")
  const tickLblPos = parseAxisTickLblPos(axis)
  // `<c:txPr><a:bodyPr rot="N"/></c:txPr>` — tick-label rotation in
  // 60000ths of a degree. The element sits on every axis flavour per
  // the OOXML schema (CT_CatAx, CT_ValAx, CT_DateAx, CT_SerAx all
  // carry an optional `<c:txPr>`), so the reader runs on every axis
  // flavour. Out-of-range values clamp to the `-90..90` band Excel's
  // UI exposes; the OOXML default `0` and absence both collapse to
  // `undefined`.
  const labelRotation = parseAxisLabelRotation(axis)
  // `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>` —
  // tick-label font size in 100ths of a point. Same `<c:txPr>` slot
  // as `labelRotation` above. Out-of-range / non-numeric values drop
  // to `undefined` so a corrupt template cannot surface a value the
  // writer would never emit. Surfaced on every axis flavour for
  // symmetry with the writer.
  const labelFontSize = parseAxisLabelFontSize(axis)
  // `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label bold flag. Same `<c:txPr>` slot scope as the rotation /
  // size readers above. The OOXML default `false` collapses to
  // `undefined` so absence and `b="0"` round-trip identically; only
  // an explicit `b="1"` surfaces `true`. Surfaced on every axis
  // flavour for symmetry with the writer.
  const labelBold = parseAxisLabelBold(axis)
  // `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label italic flag. Same `<c:txPr>` slot scope as the bold
  // reader above. The OOXML default `false` collapses to `undefined`
  // so absence and `i="0"` round-trip identically; only an explicit
  // `i="1"` surfaces `true`. Surfaced on every axis flavour for
  // symmetry with the writer.
  const labelItalic = parseAxisLabelItalic(axis)
  // `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
  // </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` — tick-label
  // font color. Same `<c:txPr>` slot scope as the rotation / size /
  // bold / italic readers above. Theme references (`<a:schemeClr>`),
  // `<a:hslClr>`, `<a:sysClr>`, `<a:prstClr>`, and malformed `val`
  // tokens all collapse to `undefined` since only the literal RGB
  // triple round-trips losslessly through the writer. Surfaced on
  // every axis flavour for symmetry with the writer.
  const labelColor = parseAxisLabelColor(axis)
  // `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label underline flag. Same `<c:txPr>` slot scope as the
  // rotation / size / bold / italic / color readers above. Only the
  // UI-default `"sng"` surfaces as `true`; the OOXML default `"none"`,
  // absence, the non-UI `"dbl"` variant, and the sixteen exotic
  // tokens all collapse to `undefined` so absence and the OOXML
  // default round-trip identically through the writer (which emits
  // only `"sng"`). Surfaced on every axis flavour for symmetry with
  // the writer.
  const labelUnderline = parseAxisLabelUnderline(axis)
  // `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>` —
  // tick-label strikethrough flag. Same `<c:txPr>` slot scope as the
  // rotation / size / bold / italic / color / underline readers above.
  // Only the UI-default `"sngStrike"` surfaces as `true`; the OOXML
  // default `"noStrike"`, absence, and the non-UI `"dblStrike"` variant
  // all collapse to `undefined` so absence and the OOXML default round-
  // trip identically through the writer (which emits only
  // `"sngStrike"`). Surfaced on every axis flavour for symmetry with
  // the writer.
  const labelStrike = parseAxisLabelStrike(axis)
  // `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
  // </a:pPr></a:p></c:txPr>` — axis tick-label font family. Same
  // axis-level `<c:txPr>` body scope as the rotation / size / bold /
  // italic / color / underline / strike readers above. Empty /
  // whitespace-only `typeface` attributes and missing `<a:latin>`
  // elements both collapse to `undefined` so absence and the empty
  // form round-trip identically through the writer. The reader is
  // scoped to the axis-level `<c:txPr>` so a stray `<a:latin>` inside
  // `<c:title>` (surfaced by `axisTitleFontFamily`) cannot leak in.
  // Surfaced on every axis flavour for symmetry with the writer.
  const labelFontFamily = parseAxisLabelFontFamily(axis)
  // <c:scaling><c:orientation val=".."/></c:scaling> — ST_Orientation
  // accepts "minMax" (default, low → high) and "maxMin" (reversed).
  // The default collapses to undefined so a fresh chart and a chart
  // that explicitly pins "minMax" round-trip identically.
  const reverse = parseAxisReverse(axis)
  // `<c:tickLblSkip>` / `<c:tickMarkSkip>` live exclusively on
  // `CT_CatAx` / `CT_DateAx` per ECMA-376 Part 1, §21.2.2 — the
  // `<c:valAx>` schema rejects them entirely. Skip the parse on
  // value axes so a corrupt template carrying a stray skip element
  // on a value axis does not surface a field the writer would never
  // emit anyway.
  const isCategoryAxis = axis.local === "catAx" || axis.local === "dateAx"
  const tickLblSkip = isCategoryAxis ? parseAxisSkip(axis, "tickLblSkip") : undefined
  const tickMarkSkip = isCategoryAxis ? parseAxisSkip(axis, "tickMarkSkip") : undefined
  // `<c:lblOffset>` lives exclusively on `CT_CatAx` / `CT_DateAx` per
  // ECMA-376 Part 1, §21.2.2 — the `<c:valAx>` and `<c:serAx>` schemas
  // reject it. Skip the parse on value axes for the same reason as
  // the skip elements above.
  const lblOffset = isCategoryAxis ? parseAxisLblOffset(axis) : undefined
  // `<c:lblAlgn>` is also category-axis-only per ECMA-376 Part 1,
  // §21.2.2 — the OOXML `ST_LblAlgn` schema places the element on
  // `CT_CatAx` / `CT_DateAx` only. Same scope rule as `lblOffset`.
  const lblAlgn = isCategoryAxis ? parseAxisLblAlgn(axis) : undefined
  // `<c:noMultiLvlLbl>` lives exclusively on `CT_CatAx` per ECMA-376
  // Part 1, §21.2.2 — even `<c:dateAx>`, `<c:valAx>`, and `<c:serAx>`
  // reject the element. Skip the parse on every other axis flavour so
  // a corrupt template carrying a stray flag does not surface a value
  // the writer would never emit anyway.
  const noMultiLvlLbl = axis.local === "catAx" ? parseAxisNoMultiLvlLbl(axis) : undefined
  // `<c:auto>` lives exclusively on `CT_CatAx` per ECMA-376 Part 1,
  // §21.2.2.7 — `<c:dateAx>`, `<c:valAx>`, and `<c:serAx>` reject the
  // element. Skip the parse on every other axis flavour for symmetry
  // with the writer's catAx-only emit path. Only `false` surfaces; the
  // OOXML default `true` (Excel inspects the data and decides whether
  // to treat the axis as a date axis) collapses to `undefined`.
  const auto = axis.local === "catAx" ? parseAxisAuto(axis) : undefined
  // `<c:delete>` sits on every axis flavour (CT_CatAx / CT_ValAx /
  // CT_DateAx / CT_SerAx) per ECMA-376 Part 1, §21.2.2. The OOXML
  // default `val="0"` (axis visible) collapses to `undefined` so
  // absence and the default round-trip identically.
  const hidden = parseAxisHidden(axis)
  // `<c:crosses>` and `<c:crossesAt>` sit on every axis flavour and live
  // in an XSD choice (CT_Crosses ⊕ CT_Double) — only one may legally
  // appear at a time per ECMA-376 Part 1, §21.2.2. The reader honours
  // the schema by preferring `crossesAt` when both elements show up
  // together (a malformed template); the writer mirrors that order so a
  // round-trip surfaces the numeric pin and drops the redundant
  // semantic toggle.
  const crossesPair = parseAxisCrosses(axis)
  const crosses = crossesPair.crosses
  const crossesAt = crossesPair.crossesAt
  // `<c:dispUnits>` lives exclusively on `<c:valAx>` per ECMA-376 Part 1,
  // §21.2.2.32 (CT_ValAx → CT_DispUnits). Skip the parse on every other
  // axis flavour so a corrupt template carrying a stray element does
  // not surface a value the writer would never emit anyway.
  const dispUnits = axis.local === "valAx" ? parseAxisDispUnits(axis) : undefined
  // `<c:crossBetween>` is also value-axis-only per ECMA-376 Part 1,
  // §21.2.2.10 (CT_ValAx → CT_CrossBetween). The OOXML schema rejects
  // the element on `<c:catAx>` / `<c:dateAx>` / `<c:serAx>`, so the
  // reader skips the parse on every other axis flavour to mirror the
  // writer's scope rule. The element is required on every `<c:valAx>`
  // and Excel always emits the family default — collapse the parsed
  // value when it matches the family default so absence and the
  // default round-trip identically through {@link cloneChart}.
  const parsedCrossBetween = axis.local === "valAx" ? parseAxisCrossBetween(axis) : undefined
  const crossBetween =
    parsedCrossBetween === familyDefaultCrossBetween ? undefined : parsedCrossBetween
  if (
    title === undefined &&
    axisTitleRotation === undefined &&
    axisTitleFontSize === undefined &&
    axisTitleBold === undefined &&
    axisTitleItalic === undefined &&
    axisTitleColor === undefined &&
    axisTitleStrike === undefined &&
    axisTitleUnderline === undefined &&
    axisTitleFontFamily === undefined &&
    axisTitleOverlay === undefined &&
    axisTitleLayout === undefined &&
    axisTitleFillColor === undefined &&
    axisTitleBorderColor === undefined &&
    axisTitleBorderWidth === undefined &&
    axisTitleBorderDash === undefined &&
    axisTitleBorderCap === undefined &&
    axisTitleBorderCompound === undefined &&
    gridlines === undefined &&
    scale === undefined &&
    numberFormat === undefined &&
    majorTickMark === undefined &&
    minorTickMark === undefined &&
    tickLblPos === undefined &&
    labelRotation === undefined &&
    labelFontSize === undefined &&
    labelBold === undefined &&
    labelItalic === undefined &&
    labelColor === undefined &&
    labelFontFamily === undefined &&
    labelUnderline === undefined &&
    labelStrike === undefined &&
    reverse === undefined &&
    tickLblSkip === undefined &&
    tickMarkSkip === undefined &&
    lblOffset === undefined &&
    lblAlgn === undefined &&
    noMultiLvlLbl === undefined &&
    auto === undefined &&
    hidden === undefined &&
    crosses === undefined &&
    crossesAt === undefined &&
    dispUnits === undefined &&
    crossBetween === undefined
  ) {
    return undefined
  }
  const out: ChartAxisInfo = {}
  if (title !== undefined) out.title = title
  if (axisTitleRotation !== undefined) out.axisTitleRotation = axisTitleRotation
  if (axisTitleFontSize !== undefined) out.axisTitleFontSize = axisTitleFontSize
  if (axisTitleBold !== undefined) out.axisTitleBold = axisTitleBold
  if (axisTitleItalic !== undefined) out.axisTitleItalic = axisTitleItalic
  if (axisTitleColor !== undefined) out.axisTitleColor = axisTitleColor
  if (axisTitleStrike !== undefined) out.axisTitleStrike = axisTitleStrike
  if (axisTitleUnderline !== undefined) out.axisTitleUnderline = axisTitleUnderline
  if (axisTitleFontFamily !== undefined) out.axisTitleFontFamily = axisTitleFontFamily
  if (axisTitleOverlay !== undefined) out.axisTitleOverlay = axisTitleOverlay
  if (axisTitleLayout !== undefined) out.axisTitleLayout = axisTitleLayout
  if (axisTitleFillColor !== undefined) out.axisTitleFillColor = axisTitleFillColor
  if (axisTitleBorderColor !== undefined) out.axisTitleBorderColor = axisTitleBorderColor
  if (axisTitleBorderWidth !== undefined) out.axisTitleBorderWidth = axisTitleBorderWidth
  if (axisTitleBorderDash !== undefined) out.axisTitleBorderDash = axisTitleBorderDash
  if (axisTitleBorderCap !== undefined) out.axisTitleBorderCap = axisTitleBorderCap
  if (axisTitleBorderCompound !== undefined) {
    out.axisTitleBorderCompound = axisTitleBorderCompound
  }
  if (gridlines !== undefined) out.gridlines = gridlines
  if (scale !== undefined) out.scale = scale
  if (numberFormat !== undefined) out.numberFormat = numberFormat
  if (majorTickMark !== undefined) out.majorTickMark = majorTickMark
  if (minorTickMark !== undefined) out.minorTickMark = minorTickMark
  if (tickLblPos !== undefined) out.tickLblPos = tickLblPos
  if (labelRotation !== undefined) out.labelRotation = labelRotation
  if (labelFontSize !== undefined) out.labelFontSize = labelFontSize
  if (labelBold !== undefined) out.labelBold = labelBold
  if (labelItalic !== undefined) out.labelItalic = labelItalic
  if (labelColor !== undefined) out.labelColor = labelColor
  if (labelFontFamily !== undefined) out.labelFontFamily = labelFontFamily
  if (labelUnderline !== undefined) out.labelUnderline = labelUnderline
  if (labelStrike !== undefined) out.labelStrike = labelStrike
  if (reverse !== undefined) out.reverse = reverse
  if (tickLblSkip !== undefined) out.tickLblSkip = tickLblSkip
  if (tickMarkSkip !== undefined) out.tickMarkSkip = tickMarkSkip
  if (lblOffset !== undefined) out.lblOffset = lblOffset
  if (lblAlgn !== undefined) out.lblAlgn = lblAlgn
  if (noMultiLvlLbl !== undefined) out.noMultiLvlLbl = noMultiLvlLbl
  if (auto !== undefined) out.auto = auto
  if (hidden !== undefined) out.hidden = hidden
  if (crosses !== undefined) out.crosses = crosses
  if (crossesAt !== undefined) out.crossesAt = crossesAt
  if (dispUnits !== undefined) out.dispUnits = dispUnits
  if (crossBetween !== undefined) out.crossBetween = crossBetween
  return out
}

/**
 * Pull `<c:majorTickMark val=".."/>` (or `<c:minorTickMark>`) off an
 * axis element. Returns `undefined` when the element is absent, the
 * `val` attribute is missing, the value is not in
 * {@link VALID_TICK_MARKS}, or the value matches the per-element
 * OOXML default — `"out"` for major, `"none"` for minor — so absence
 * and the default round-trip identically.
 */
export function parseAxisTickMark(
  axis: XmlElement,
  localName: "majorTickMark" | "minorTickMark",
  defaultValue: ChartAxisTickMark,
): ChartAxisTickMark | undefined {
  const el = findChild(axis, localName)
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  const value = raw.trim() as ChartAxisTickMark
  if (!VALID_TICK_MARKS.has(value)) return undefined
  return value === defaultValue ? undefined : value
}

/**
 * Pull `<c:tickLblPos val=".."/>` off an axis element. Returns
 * `undefined` when the element is absent, the `val` attribute is
 * missing or unrecognized, or the value matches the OOXML default
 * `"nextTo"` so absence and the default round-trip identically.
 */
export function parseAxisTickLblPos(axis: XmlElement): ChartAxisTickLabelPosition | undefined {
  const el = findChild(axis, "tickLblPos")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  const value = raw.trim() as ChartAxisTickLabelPosition
  if (!VALID_TICK_LBL_POSITIONS.has(value)) return undefined
  return value === "nextTo" ? undefined : value
}

/**
 * Pull the `ST_Orientation` value off `<c:scaling><c:orientation/></c:scaling>`.
 * Returns `true` only when the axis pinned `"maxMin"` (Excel's
 * "Categories / Values in reverse order" toggle); the OOXML default
 * `"minMax"` collapses to `undefined` so absence and the default
 * round-trip identically. Unknown tokens (e.g. typo'd templates) drop
 * to `undefined` rather than fabricate a flag.
 */
export function parseAxisReverse(axis: XmlElement): boolean | undefined {
  const scaling = findChild(axis, "scaling")
  if (!scaling) return undefined
  const orientation = findChild(scaling, "orientation")
  if (!orientation) return undefined
  const raw = orientation.attrs.val
  if (typeof raw !== "string") return undefined
  const value = raw.trim()
  if (value === "maxMin") return true
  // "minMax" and unknown tokens both fall through to undefined — only
  // an explicit reversed orientation surfaces.
  return undefined
}

/**
 * Pull `<c:tickLblSkip val=".."/>` or `<c:tickMarkSkip val=".."/>`
 * off a category axis element. Returns `undefined` when:
 *   - the element is absent,
 *   - the `val` attribute is missing or non-numeric,
 *   - the parsed value is `1` (the OOXML default — show every label /
 *     mark),
 *   - the parsed value falls outside the OOXML `ST_SkipIntervals`
 *     range (`1..32767`).
 *
 * Negative / zero / out-of-range inputs are dropped rather than
 * clamped so a corrupt template cannot leak a skip count Excel would
 * reject.
 */
export function parseAxisSkip(
  axis: XmlElement,
  localName: "tickLblSkip" | "tickMarkSkip",
): number | undefined {
  const el = findChild(axis, localName)
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  const parsed = Number.parseInt(trimmed, 10)
  if (!Number.isFinite(parsed)) return undefined
  if (parsed < 1 || parsed > 32767) return undefined
  if (parsed === 1) return undefined
  return parsed
}

/**
 * Pull `<c:lblOffset val=".."/>` off a category axis element. Returns
 * `undefined` when:
 *   - the element is absent,
 *   - the `val` attribute is missing or non-numeric,
 *   - the parsed value is `100` (the OOXML default — Excel's
 *     reference label spacing),
 *   - the parsed value falls outside the OOXML
 *     `ST_LblOffsetPercent` range (`0..1000`).
 *
 * Out-of-range / non-numeric inputs are dropped rather than clamped
 * so a corrupt template cannot leak an offset Excel would reject.
 */
export function parseAxisLblOffset(axis: XmlElement): number | undefined {
  const el = findChild(axis, "lblOffset")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  const parsed = Number.parseInt(trimmed, 10)
  if (!Number.isFinite(parsed)) return undefined
  if (parsed < 0 || parsed > 1000) return undefined
  if (parsed === 100) return undefined
  return parsed
}

/**
 * Pull `<c:lblAlgn val=".."/>` off a category axis element. Returns
 * `undefined` when:
 *   - the element is absent,
 *   - the `val` attribute is missing or blank,
 *   - the value is not in {@link VALID_LBL_ALIGNS},
 *   - the value is `"ctr"` (the OOXML default — Excel's reference
 *     centered alignment).
 *
 * Unknown tokens drop rather than fall through to the default so a
 * corrupt template cannot leak an alignment Excel would reject.
 */
export function parseAxisLblAlgn(axis: XmlElement): ChartAxisLabelAlign | undefined {
  const el = findChild(axis, "lblAlgn")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  const value = raw.trim() as ChartAxisLabelAlign
  if (!VALID_LBL_ALIGNS.has(value)) return undefined
  return value === "ctr" ? undefined : value
}

/**
 * Pull `<c:noMultiLvlLbl val=".."/>` off a category axis element.
 * Returns `true` only when the axis pinned `val="1"` / `val="true"`
 * (Excel's "Multi-level Category Labels" checkbox unchecked, i.e.
 * tiered category labels collapsed onto a single line). The OOXML
 * default `val="0"` / `val="false"`, absence, missing `val`, and
 * unknown tokens all collapse to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart}.
 *
 * Mirrors the truthy / falsy parsing in {@link parseAxisHidden} —
 * the OOXML schema (`xsd:boolean`) accepts `0` / `1` / `false` /
 * `true` for `<c:noMultiLvlLbl>` just as it does for every other
 * Boolean-valued chart attribute.
 */
export function parseAxisNoMultiLvlLbl(axis: XmlElement): boolean | undefined {
  const el = findChild(axis, "noMultiLvlLbl")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  switch (raw.trim()) {
    case "1":
    case "true":
      return true
    case "0":
    case "false":
      return undefined
    default:
      return undefined
  }
}

/**
 * Pull `<c:auto val=".."/>` off a category axis element. Returns
 * `false` only when the axis pinned `val="0"` / `val="false"` (Excel's
 * "Text axis" radio button under "Format Axis -> Axis Options -> Axis
 * Type" — Excel keeps every label as-is regardless of whether the
 * cells parse as dates / numerics). The OOXML default `val="1"` /
 * `val="true"` (Excel inspects the data and decides whether to treat
 * the axis as a discrete category axis or a chronological date axis),
 * absence, missing `val`, and unknown tokens all collapse to
 * `undefined` so absence and the default round-trip identically
 * through {@link cloneChart}.
 *
 * Mirrors the truthy / falsy parsing in {@link parseAxisNoMultiLvlLbl}
 * — the OOXML schema (`xsd:boolean`) accepts `0` / `1` / `false` /
 * `true` for `<c:auto>` just as it does for every other Boolean-valued
 * chart attribute. The element's default is the OOXML inverse of
 * `noMultiLvlLbl` (auto defaults to `true`, noMultiLvlLbl defaults to
 * `false`), so this parser collapses `true` rather than `false`.
 */
export function parseAxisAuto(axis: XmlElement): boolean | undefined {
  const el = findChild(axis, "auto")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  switch (raw.trim()) {
    case "0":
    case "false":
      return false
    case "1":
    case "true":
      // OOXML default — collapse to undefined so absence and the
      // default round-trip identically.
      return undefined
    default:
      return undefined
  }
}

/**
 * Pull `<c:delete val=".."/>` off an axis element. Returns `true`
 * only when the axis pinned `val="1"` / `val="true"` (Excel's "hide
 * axis" toggle). The OOXML default `val="0"` / `val="false"`,
 * absence, missing `val`, and unknown tokens all collapse to
 * `undefined` so absence and the default round-trip identically.
 *
 * Mirrors the truthy / falsy parsing in {@link parsePlotVisOnly} —
 * the OOXML schema (`xsd:boolean`) accepts `0` / `1` / `false` /
 * `true` for `<c:delete>` just as it does for every other Boolean-
 * valued chart attribute.
 */
export function parseAxisHidden(axis: XmlElement): boolean | undefined {
  const el = findChild(axis, "delete")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  switch (raw.trim()) {
    case "1":
    case "true":
      return true
    case "0":
    case "false":
      // OOXML default — collapse to undefined so absence and the
      // default round-trip identically.
      return undefined
    default:
      return undefined
  }
}

/**
 * Pull `<c:txPr><a:bodyPr rot="N"/></c:txPr>` off an axis element.
 * Returns the rotation in whole degrees (range `-90..90`).
 *
 * The OOXML default `0` (and absence of the element / attribute) all
 * collapse to `undefined` so absence and the default round-trip
 * identically through {@link cloneChart}. Non-integer / non-numeric /
 * out-of-range values clamp to the nearest endpoint of the
 * `-90..90` band Excel's UI exposes; non-finite (`NaN`, `Infinity`)
 * inputs drop to `undefined`.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the rotation
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelRotation.
 */
export function parseAxisLabelRotation(axis: XmlElement): number | undefined {
  const txPr = findChild(axis, "txPr")
  if (!txPr) return undefined
  const bodyPr = findChild(txPr, "bodyPr")
  if (!bodyPr) return undefined
  const raw = bodyPr.attrs.rot
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  const parsed = Number.parseInt(trimmed, 10)
  if (!Number.isFinite(parsed)) return undefined
  // Convert from 60000ths of a degree to whole degrees.
  const degrees = Math.round(parsed / TXPR_ROT_PER_DEGREE)
  if (degrees === 0) return undefined
  if (degrees < LABEL_ROTATION_MIN_DEG) return LABEL_ROTATION_MIN_DEG
  if (degrees > LABEL_ROTATION_MAX_DEG) return LABEL_ROTATION_MAX_DEG
  return degrees
}

/**
 * Pull `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
 * off an axis element. Returns the tick-label font size in points
 * (range `1..400`).
 *
 * Mirrors {@link parseAxisTitleFontSize} for tick labels — same
 * 0.5pt half-step granularity, same `1..400`pt band, same
 * drop-on-out-of-range / non-numeric / absence semantics. The lookup
 * is scoped to the axis-level `<c:txPr>` so a stray `<a:defRPr>`
 * inside `<c:title><c:tx><c:rich>` (surfaced by
 * {@link parseAxisTitleFontSize}) cannot leak in.
 *
 * The OOXML default — `<a:defRPr>` with no `sz` attribute — collapses
 * to `undefined` so absence and the default round-trip identically
 * through {@link cloneChart}. Out-of-range / non-numeric `sz` values
 * drop rather than fabricate a value the writer would never emit.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the size
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelFontSize.
 */
export function parseAxisLabelFontSize(axis: XmlElement): number | undefined {
  const txPr = findChild(axis, "txPr")
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
  // chart-level / axis-title `parseTitleFontSize` half-step
  // normalisation.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2)
  const points = halfSteps / 2
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined
  return points
}

/**
 * Pull `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>`
 * off an axis element. Returns the tick-label bold flag.
 *
 * Mirrors {@link parseAxisTitleBold} for tick labels — same
 * canonical-slot pair, same drop-on-default-`false` semantics. The
 * OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * default `false` collapses to `undefined` so absence and `b="0"`
 * round-trip identically — only an explicit `b="1"` surfaces `true`.
 * Unknown / malformed `b` tokens drop to `undefined` rather than
 * fabricate a value the writer would never emit.
 *
 * The lookup is scoped to the axis-level `<c:txPr>` so a stray
 * `<a:defRPr b=".."/>` inside `<c:title><c:tx><c:rich>` (surfaced by
 * {@link parseAxisTitleBold}) cannot leak in. Returns `undefined`
 * whenever the axis omits `<c:txPr>` entirely or the canonical
 * `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelBold.
 */
export function parseAxisLabelBold(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const parsed = parseBoolAttr(defRPr.attrs.b)
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through the writer — only an
  // explicit `b="1"` surfaces `true`.
  if (parsed === true) return true
  return undefined
}

/**
 * Pull the axis tick-label italic flag off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>` chain
 * Excel writes when the user pins italic on the axis tick labels.
 *
 * Returns `true` only when the parser walks the full chain and lands on
 * an `<a:defRPr i="1"/>` (or the OOXML truthy spelling `i="true"`); the
 * OOXML default `i="0"` collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart}. Returns
 * `undefined` whenever the axis omits `<c:txPr>` entirely or the
 * canonical `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelItalic.
 */
export function parseAxisLabelItalic(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const parsed = parseBoolAttr(defRPr.attrs.i)
  // The OOXML default `false` collapses to `undefined` so absence and
  // `i="0"` round-trip identically through the writer — only an
  // explicit `i="1"` surfaces `true`.
  if (parsed === true) return true
  return undefined
}

/**
 * Pull the axis tick-label font color off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` chain Excel writes
 * when the user pins a custom font color on the axis tick labels.
 *
 * Returns the 6-character uppercase hex string when the parser walks
 * the full chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme
 * references (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and
 * `<a:prstClr>` all collapse to `undefined` — only the literal RGB
 * triple round-trips losslessly through {@link writeChart}. Malformed
 * `val` tokens (wrong length, non-hex characters) likewise drop to
 * `undefined` rather than fabricate a value the writer would round-
 * trip into a malformed `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the axis omits `<c:txPr>` entirely or
 * the canonical `<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr>` chain
 * is malformed at any link.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the value
 * regardless of axis flavour so a parsed chart preserves the color
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelColor. The lookup is scoped to the
 * axis-level `<c:txPr>` so a stray `<a:solidFill>` inside `<c:title>`
 * (surfaced by {@link parseAxisTitleColor}) or on a `<c:spPr>` series
 * fill cannot leak in.
 */
export function parseAxisLabelColor(axis: XmlElement): ChartColor | undefined {
  const txPr = findChild(axis, "txPr")
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
 * Pull the axis tick-label underline flag off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>` chain
 * Excel writes when the user pins an underline on the axis tick labels.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * eighteen values; Excel's UI exposes only `"sng"` (single line — the
 * default underline checkbox) and `"dbl"` (double line). The reader
 * surfaces only the UI-default `"sng"` as `true`; `"none"` (the OOXML
 * application default), absence, the non-UI `"dbl"` variant, and the
 * sixteen exotic tokens (`"words"`, `"heavy"`, `"dotted"`,
 * `"dottedHeavy"`, `"dash"`, `"dashHeavy"`, `"dashLong"`,
 * `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
 * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`) all
 * collapse to `undefined` — the writer emits only `"sng"`, so
 * reporting any non-single underline as `true` would silently
 * downgrade the choice to a single line on round-trip. Unknown /
 * malformed `u` tokens likewise drop to `undefined`.
 *
 * Mirrors {@link parseAxisTitleUnderline} for tick labels — same
 * canonical-slot semantics but scoped to the axis-level `<c:txPr>` so
 * a stray `<a:defRPr u=".."/>` inside `<c:title><c:tx><c:rich>`
 * (surfaced by {@link parseAxisTitleUnderline}) cannot leak in.
 * Returns `undefined` whenever the axis omits `<c:txPr>` entirely or
 * the canonical `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelUnderline.
 */
export function parseAxisLabelUnderline(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.u
  // Only the UI-default `"sng"` surfaces as `true`. The OOXML
  // application default `"none"`, the non-UI `"dbl"` variant, and
  // every exotic token (`"words"`, `"heavy"`, `"dotted"`, etc.) all
  // collapse to `undefined` so absence and the OOXML default
  // round-trip identically through the writer; the writer emits only
  // `"sng"`, so reporting a non-single underline here would silently
  // downgrade the choice on round-trip.
  if (raw === "sng") return true
  return undefined
}

/**
 * Pull the axis tick-label strikethrough flag off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>`
 * chain Excel writes when the user pins a strikethrough on the axis
 * tick labels.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * three values: `"noStrike"` (the OOXML application default),
 * `"sngStrike"` (single line — the value Excel's UI checkbox
 * emits), and `"dblStrike"` (double line — a non-UI variant). The
 * reader surfaces only the UI-default `"sngStrike"` as `true`;
 * `"noStrike"` (the OOXML application default), absence, and the
 * non-UI `"dblStrike"` variant all collapse to `undefined` — the
 * writer emits only `"sngStrike"`, so reporting `"dblStrike"` as
 * `true` would silently downgrade the choice to a single line on
 * round-trip. Unknown / malformed `strike` tokens likewise drop to
 * `undefined`.
 *
 * Mirrors {@link parseAxisTitleStrike} for tick labels — same
 * canonical-slot semantics but scoped to the axis-level `<c:txPr>` so
 * a stray `<a:defRPr strike=".."/>` inside `<c:title><c:tx><c:rich>`
 * (surfaced by {@link parseAxisTitleStrike}) cannot leak in. Returns
 * `undefined` whenever the axis omits `<c:txPr>` entirely or the
 * canonical `<a:p><a:pPr><a:defRPr>` chain is malformed.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the flag
 * regardless of axis flavour so a parsed chart preserves the value
 * for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelStrike.
 */
export function parseAxisLabelStrike(axis: XmlElement): boolean | undefined {
  const txPr = findChild(axis, "txPr")
  if (!txPr) return undefined
  const p = findChild(txPr, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.strike
  // Only the UI-default `"sngStrike"` surfaces as `true`. The OOXML
  // application default `"noStrike"`, the non-UI `"dblStrike"` variant,
  // and unknown / malformed tokens all collapse to `undefined` so
  // absence and the OOXML default round-trip identically through the
  // writer; the writer emits only `"sngStrike"`, so reporting
  // `"dblStrike"` here would silently downgrade the choice on round-
  // trip.
  if (raw === "sngStrike") return true
  return undefined
}

/**
 * Pull the axis tick-label font family off the canonical
 * `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
 * </a:pPr></a:p></c:txPr>` chain Excel writes when the user pins a
 * typeface on the axis tick labels.
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
 * Returns `undefined` whenever the axis omits `<c:txPr>` entirely or
 * the canonical `<a:p><a:pPr><a:defRPr><a:latin>` chain is malformed
 * at any link.
 *
 * The `<c:txPr>` element sits on every axis flavour — `<c:catAx>` /
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>` all carry the optional
 * element per the OOXML schema. The reader surfaces the value
 * regardless of axis flavour so a parsed chart preserves the
 * typeface for symmetry with the writer-side
 * {@link SheetChart.axes}.x.labelFontFamily. The lookup is scoped to
 * the axis-level `<c:txPr>` so a stray `<a:latin>` inside the axis
 * `<c:title><c:tx><c:rich>` body cannot leak in.
 */
export function parseAxisLabelFontFamily(axis: XmlElement): string | undefined {
  const txPr = findChild(axis, "txPr")
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
 * Pull the axis crossing pin off `<c:crosses>` / `<c:crossesAt>`. The
 * OOXML schema (`CT_CatAx`, `CT_ValAx`, `CT_DateAx`, `CT_SerAx`) places
 * the two elements in an XSD choice — only one may legally appear at a
 * time per ECMA-376 Part 1, §21.2.2. The reader still handles both
 * appearing on the same axis (a malformed template) by preferring
 * `crossesAt` and dropping the redundant `crosses` value, mirroring the
 * writer's emit order.
 *
 * Returns:
 *   - `crosses`   — set when only `<c:crosses>` is present and the value
 *                   is a non-default token. The OOXML default `"autoZero"`
 *                   collapses to `undefined` so absence and the default
 *                   round-trip identically. Unknown tokens drop rather
 *                   than fabricate a value the writer would never emit.
 *   - `crossesAt` — set when `<c:crossesAt>` is present with a
 *                   parseable numeric `val`. Non-numeric / missing
 *                   `val` attributes drop to `undefined`. `0` is
 *                   preserved (it is a valid pin, distinct from the
 *                   `"autoZero"` default).
 */
export function parseAxisCrosses(axis: XmlElement): {
  crosses?: ChartAxisCrosses
  crossesAt?: number
} {
  const crossesAtEl = findChild(axis, "crossesAt")
  if (crossesAtEl) {
    const raw = crossesAtEl.attrs.val
    if (typeof raw === "string") {
      const trimmed = raw.trim()
      if (trimmed.length > 0) {
        const parsed = Number.parseFloat(trimmed)
        if (Number.isFinite(parsed)) {
          return { crossesAt: parsed }
        }
      }
    }
  }

  const crossesEl = findChild(axis, "crosses")
  if (!crossesEl) return {}
  const raw = crossesEl.attrs.val
  if (typeof raw !== "string") return {}
  const value = raw.trim() as ChartAxisCrosses
  if (!VALID_CROSSES.has(value)) return {}
  if (value === "autoZero") return {}
  return { crosses: value }
}

/**
 * Read a value axis's `<c:dispUnits>` block. The element holds an
 * `xsd:choice` between `<c:builtInUnit val=".."/>` and
 * `<c:custUnit val=".."/>`, optionally followed by `<c:dispUnitsLbl>`.
 * The reader surfaces both: a recognized `<c:builtInUnit>` token lands
 * in `unit`, and a finite positive `<c:custUnit>` value lands in
 * `custUnit`. When both children are present (a malformed template,
 * since the schema's choice forbids it), `custUnit` wins and `unit`
 * drops — the writer applies the same precedence on emit, so the parsed
 * shape round-trips identically through {@link cloneChart}.
 *
 * Returns `undefined` when:
 *   - the axis declares no `<c:dispUnits>` at all,
 *   - `<c:dispUnits>` is present but neither child resolves to a
 *     valid value (missing children, malformed `val`, unknown
 *     `<c:builtInUnit>` token, non-positive / non-finite `<c:custUnit>`).
 *
 * `showLabel` is set `true` only when `<c:dispUnitsLbl>` is present
 * inside `<c:dispUnits>` (Excel paints its automatic annotation in
 * that case). Absence collapses to absence on the surfaced object so
 * a round-trip stays minimal.
 */
export function parseAxisDispUnits(axis: XmlElement): ChartAxisDispUnits | undefined {
  const dispUnits = findChild(axis, "dispUnits")
  if (!dispUnits) return undefined
  const out: ChartAxisDispUnits = {}
  // `<c:custUnit>` wins when both children are pinned — the OOXML
  // schema's `xsd:choice` forbids both, but a corrupt template may
  // declare them simultaneously. The writer mirrors this preference so
  // the round-trip stays consistent.
  const custUnit = findChild(dispUnits, "custUnit")
  if (custUnit) {
    const raw = custUnit.attrs.val
    if (typeof raw === "string") {
      const parsed = Number.parseFloat(raw.trim())
      if (Number.isFinite(parsed) && parsed > 0) {
        out.custUnit = parsed
      }
    }
  }
  if (out.custUnit === undefined) {
    const builtInUnit = findChild(dispUnits, "builtInUnit")
    if (builtInUnit) {
      const raw = builtInUnit.attrs.val
      if (typeof raw === "string") {
        const trimmed = raw.trim() as ChartAxisDispUnit
        if (VALID_DISP_UNITS.has(trimmed)) {
          out.unit = trimmed
        }
      }
    }
  }
  if (out.unit === undefined && out.custUnit === undefined) return undefined
  const lbl = findChild(dispUnits, "dispUnitsLbl")
  if (lbl) {
    out.showLabel = true
    // Walk `<c:dispUnitsLbl><c:tx><c:rich><a:p><a:r><a:t>...</a:t>` for
    // an optional custom label. Multiple paragraphs / runs concatenate
    // with newlines so a richly-formatted label round-trips as plain
    // text. Empty / whitespace-only strings collapse to absence.
    const tx = findChild(lbl, "tx")
    if (tx) {
      const rich = findChild(tx, "rich")
      if (rich) {
        const buf: string[] = []
        for (const p of rich.children) {
          if (typeof p === "string") continue
          if (p.local !== "p") continue
          const paraBuf: string[] = []
          for (const r of p.children) {
            if (typeof r === "string") continue
            if (r.local !== "r") continue
            for (const t of r.children) {
              if (typeof t === "string") continue
              if (t.local !== "t") continue
              let text = ""
              for (const c of t.children) {
                if (typeof c === "string") text += c
              }
              paraBuf.push(text)
            }
          }
          if (paraBuf.length > 0) buf.push(paraBuf.join(""))
        }
        const joined = buf.join("\n").trim()
        if (joined.length > 0) out.customLabel = joined
      }
    }
  }
  return out
}

/**
 * Read a value axis's `<c:crossBetween val=".."/>`. The OOXML schema
 * places the element exclusively on `CT_ValAx` per ECMA-376 Part 1,
 * §21.2.2.10 — `<c:catAx>`, `<c:dateAx>`, and `<c:serAx>` reject it —
 * so the caller is expected to gate the parse on `axis.local === "valAx"`
 * before calling this helper.
 *
 * Returns `undefined` when:
 *   - the axis declares no `<c:crossBetween>` at all,
 *   - the `val` attribute is missing, empty, or not a string,
 *   - the `val` attribute is not one of the OOXML `ST_CrossBetween`
 *     tokens (`"between"` / `"midCat"`).
 *
 * Unknown tokens drop rather than fabricate a value the writer would
 * never emit — the caller cannot tell absence from a corrupt template
 * without the parser's help.
 */
export function parseAxisCrossBetween(axis: XmlElement): ChartAxisCrossBetween | undefined {
  const el = findChild(axis, "crossBetween")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim() as ChartAxisCrossBetween
  if (!VALID_CROSS_BETWEEN.has(trimmed)) return undefined
  return trimmed
}

/**
 * Read an axis's numeric scale block. The scale lives inside
 * `<c:scaling>`, with one optional child per pinned bound:
 *
 *   <c:scaling>
 *     <c:orientation val="minMax"/>
 *     <c:logBase val="10"/>
 *     <c:min val="0"/>
 *     <c:max val="100"/>
 *     <c:majorUnit val="20"/>
 *     <c:minorUnit val="5"/>
 *   </c:scaling>
 *
 * Returns `undefined` when none of the numeric children declare a
 * usable value — the orientation child alone (Excel's autoscale
 * baseline) does not surface a scale.
 */
export function parseAxisScale(axis: XmlElement): ChartAxisScale | undefined {
  const out: ChartAxisScale = {}

  // <c:min>, <c:max>, and <c:logBase> live inside <c:scaling>; the
  // tick-spacing children <c:majorUnit> / <c:minorUnit> sit directly
  // under <c:catAx>/<c:valAx> per CT_CatAx / CT_ValAx in ECMA-376.
  const scaling = findChild(axis, "scaling")
  if (scaling) {
    const min = parseNumericChildVal(scaling, "min")
    if (min !== undefined) out.min = min

    const max = parseNumericChildVal(scaling, "max")
    if (max !== undefined) out.max = max

    const logBase = parseNumericChildVal(scaling, "logBase")
    if (logBase !== undefined) out.logBase = logBase
  }

  const majorUnit = parseNumericChildVal(axis, "majorUnit")
  if (majorUnit !== undefined && majorUnit > 0) out.majorUnit = majorUnit

  const minorUnit = parseNumericChildVal(axis, "minorUnit")
  if (minorUnit !== undefined && minorUnit > 0) out.minorUnit = minorUnit

  return Object.keys(out).length > 0 ? out : undefined
}

/**
 * Read an axis's `<c:numFmt formatCode=".." sourceLinked=".."/>`.
 * Returns `undefined` when the element is absent or carries an empty
 * `formatCode`. `sourceLinked` is normalized to a boolean — `0`/`1`
 * and `"true"`/`"false"` are both accepted.
 */
export function parseAxisNumberFormat(axis: XmlElement): ChartAxisNumberFormat | undefined {
  const numFmt = findChild(axis, "numFmt")
  if (!numFmt) return undefined
  const formatCode = numFmt.attrs.formatCode
  if (typeof formatCode !== "string" || formatCode.length === 0) return undefined
  const out: ChartAxisNumberFormat = { formatCode }
  const sourceLinked = numFmt.attrs.sourceLinked
  if (sourceLinked !== undefined && parseBoolAttr(sourceLinked) === true) {
    out.sourceLinked = true
  }
  return out
}

/**
 * Detect `<c:majorGridlines>` / `<c:minorGridlines>` children on an
 * axis element. The mere presence of either child element flips the
 * corresponding flag on — Excel allows but does not require nested
 * `<c:spPr>` styling, and the toggle survives even when the body is
 * empty.
 *
 * Returns `undefined` when neither element is present so the consumer
 * never sees a "{ major: false, minor: false }" record that
 * round-trips into a redundant write.
 */
export function parseAxisGridlines(axis: XmlElement): ChartAxisGridlines | undefined {
  const major = findChild(axis, "majorGridlines") !== undefined
  const minor = findChild(axis, "minorGridlines") !== undefined
  if (!major && !minor) return undefined
  const out: ChartAxisGridlines = {}
  if (major) out.major = true
  if (minor) out.minor = true
  return out
}

/**
 * Read an axis's `<c:title>` text. Mirrors {@link parseTitle} but
 * scoped to a single axis element rather than the chart root.
 */
export function parseAxisTitle(axis: XmlElement): string | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (rich) {
    const parts: string[] = []
    collectTextRuns(rich, parts)
    const joined = parts.join("").trim()
    return joined.length > 0 ? joined : undefined
  }
  const strRef = findChild(tx, "strRef")
  if (strRef) {
    const cache = findChild(strRef, "strCache")
    if (cache) {
      for (const pt of childElements(cache)) {
        if (pt.local !== "pt") continue
        const v = findChild(pt, "v")
        if (v) {
          const text = elementText(v).trim()
          if (text.length > 0) return text
        }
      }
    }
  }
  return undefined
}

/**
 * Pull `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`
 * off an axis element. Returns the rotation in whole degrees (range
 * `-90..90`).
 *
 * The OOXML default `0` (and absence of the element / attribute) all
 * collapse to `undefined` so absence and the default round-trip
 * identically through {@link cloneChart}. Out-of-range values clamp to
 * the nearest endpoint of the `-90..90` band Excel's UI exposes;
 * non-finite (`NaN`, `Infinity`) inputs drop to `undefined`.
 *
 * The lookup is scoped to the title's `<c:rich>` body so a stray
 * `<a:bodyPr>` elsewhere on the axis (e.g. the tick-label `<c:txPr>`
 * surfaced by {@link parseAxisLabelRotation}) cannot leak in. Returns
 * `undefined` when the axis omits `<c:title>` entirely or when the
 * title is a `<c:strRef>` (formula reference) with no `<c:rich>` body.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title rotation
 * {@link parseTitleRotation} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleRotation.
 */
export function parseAxisTitleRotation(axis: XmlElement): number | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  const bodyPr = findChild(rich, "bodyPr")
  if (!bodyPr) return undefined
  const raw = bodyPr.attrs.rot
  if (typeof raw !== "string") return undefined
  const trimmed = raw.trim()
  if (trimmed.length === 0) return undefined
  const parsed = Number.parseInt(trimmed, 10)
  if (!Number.isFinite(parsed)) return undefined
  // Convert from 60000ths of a degree to whole degrees.
  const degrees = Math.round(parsed / TXPR_ROT_PER_DEGREE)
  if (degrees === 0) return undefined
  if (degrees < LABEL_ROTATION_MIN_DEG) return LABEL_ROTATION_MIN_DEG
  if (degrees > LABEL_ROTATION_MAX_DEG) return LABEL_ROTATION_MAX_DEG
  return degrees
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off an axis element. Returns the
 * font size in points (range `1..400`).
 *
 * Mirrors {@link parseTitleFontSize} for axis titles — same 0.5pt
 * half-step granularity, same `1..400`pt band, same drop-on-out-of-
 * range / non-numeric / absence semantics. The lookup is scoped to
 * the axis title's `<c:rich>` body so a stray `<a:defRPr>` elsewhere
 * on the axis (e.g. on the tick-label `<c:txPr>` surfaced by
 * {@link parseAxisLabelRotation}) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the size
 * from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title size
 * {@link parseTitleFontSize} so a parsed value slots straight into
 * the writer-side {@link SheetChart.axes}.x.axisTitleFontSize.
 */
export function parseAxisTitleFontSize(axis: XmlElement): number | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph font size. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p")
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
  // chart-level `parseTitleFontSize` half-step normalisation.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2)
  const points = halfSteps / 2
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined
  return points
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off an axis element. Returns the
 * bold flag.
 *
 * Mirrors {@link parseTitleBold} for axis titles — same
 * canonical-slot pair, same drop-on-default-`false` semantics. The
 * OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * default `false` collapses to `undefined` so absence and `b="0"`
 * round-trip identically — only an explicit `b="1"` surfaces `true`.
 * Unknown / malformed `b` tokens drop to `undefined` rather than
 * fabricate a value the writer would never emit.
 *
 * The lookup is scoped to the axis title's `<c:rich>` body so a stray
 * `<a:defRPr>` elsewhere on the axis (e.g. on the tick-label
 * `<c:txPr>` surfaced by {@link parseAxisLabelRotation}) cannot leak
 * in. Returns `undefined` when the axis omits `<c:title>` entirely or
 * when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to host the flag in
 * either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title bold
 * {@link parseTitleBold} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleBold.
 */
export function parseAxisTitleBold(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph bold flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const parsed = parseBoolAttr(defRPr.attrs.b)
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through the writer — only an
  // explicit `b="1"` surfaces `true`.
  if (parsed === true) return true
  return undefined
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off an axis element. Returns
 * `true` only when the OOXML attribute is the literal truthy spelling
 * (`"1"` / `"true"`); the OOXML default `false` collapses to
 * `undefined` so absence and `i="0"` round-trip identically through
 * the writer.
 *
 * Mirrors {@link parseTitleItalic} for axis titles — same canonical-
 * slot pair (`<a:defRPr>` carries the default-paragraph italic flag,
 * which the writer keeps in sync with the literal run's `<a:rPr>` so
 * the reader only needs to consult one of the two slots), same
 * drop-on-default semantics. The lookup is scoped to the axis title's
 * `<c:rich>` body so a stray `<a:defRPr>` elsewhere on the axis (e.g.
 * on the tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the flag
 * from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title italic
 * {@link parseTitleItalic} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleItalic.
 */
export function parseAxisTitleItalic(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph italic flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const parsed = parseBoolAttr(defRPr.attrs.i)
  // The OOXML default `false` collapses to `undefined` so absence and
  // `i="0"` round-trip identically through the writer — only an
  // explicit `i="1"` surfaces `true`.
  if (parsed === true) return true
  return undefined
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` off an axis element. Returns the axis
 * title's sRGB font color as a 6-character uppercase hex string.
 *
 * Mirrors {@link parseTitleColor} for axis titles — same canonical-
 * slot pair (`<a:defRPr>` carries the default-paragraph fill, which
 * the writer keeps in sync with the literal run's `<a:rPr>` so the
 * reader only needs to consult one of the two slots), same drop-on-
 * non-sRGB semantics. The lookup is scoped to the axis title's
 * `<c:rich>` body so a stray `<a:solidFill>` elsewhere on the axis
 * (e.g. on a tick-label `<c:txPr>` block or a `<c:spPr>` series
 * fill) cannot leak in.
 *
 * The OOXML `<a:srgbClr val=".."/>` is the literal sRGB triple Excel
 * lands on the axis title's default-paragraph properties when the
 * user picks a custom font color. Theme references (`<a:schemeClr>`),
 * `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>` all collapse to
 * `undefined` — only the literal RGB triple round-trips losslessly
 * through {@link writeChart}. Malformed `val` tokens (wrong length,
 * non-hex characters) likewise drop to `undefined` rather than
 * fabricate a value the writer would round-trip into a malformed
 * `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the fill
 * from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title color
 * {@link parseTitleColor} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleColor.
 */
export function parseAxisTitleColor(axis: XmlElement): ChartColor | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  // `<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr>` is the OOXML path
  // Excel writes for the default-paragraph font color. The reader walks
  // the canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p")
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
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off an axis element.
 * Returns the axis-title strikethrough flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * three values: `"noStrike"` (the OOXML application default),
 * `"sngStrike"` (single line, the value Excel's UI checkbox emits),
 * and `"dblStrike"` (double line, a non-UI variant). The reader
 * surfaces only the UI-default `"sngStrike"` as `true`; `"noStrike"`,
 * absence, and the non-UI `"dblStrike"` all collapse to `undefined` —
 * the writer emits only `"sngStrike"`, so reporting `"dblStrike"` as
 * `true` would silently downgrade the choice to a single line on
 * round-trip. Unknown / malformed `strike` tokens likewise drop to
 * `undefined`.
 *
 * Mirrors {@link parseTitleStrike} for axis titles — same canonical-
 * slot pair (`<a:defRPr>` carries the default-paragraph strike flag,
 * which the writer keeps in sync with the literal run's `<a:rPr>` so
 * the reader only needs to consult one of the two slots), same
 * drop-on-default semantics. The lookup is scoped to the axis title's
 * `<c:rich>` body so a stray `<a:defRPr>` elsewhere on the axis (e.g.
 * on the tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely or
 * when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the flag from
 * in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title strike
 * {@link parseTitleStrike} so a parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleStrike.
 */
export function parseAxisTitleStrike(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph strikethrough flag. The reader walks the
  // canonical chain and bails on the first missing link so a malformed
  // `<c:rich>` surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.strike
  // Only the UI-default `"sngStrike"` surfaces as `true`. The OOXML
  // application default `"noStrike"` and the non-UI `"dblStrike"` both
  // collapse to `undefined` so absence and the OOXML default round-trip
  // identically through the writer; the writer emits only `"sngStrike"`,
  // so reporting `"dblStrike"` here would silently downgrade the choice
  // on round-trip.
  if (raw === "sngStrike") return true
  return undefined
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off an axis element.
 * Returns the axis-title underline flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * eighteen values; Excel's UI exposes only `"sng"` (single line — the
 * default underline checkbox) and `"dbl"` (double line). The reader
 * surfaces only the UI-default `"sng"` as `true`; `"none"` (the OOXML
 * application default), absence, the non-UI `"dbl"` variant, and the
 * sixteen exotic tokens (`"words"`, `"heavy"`, `"dotted"`,
 * `"dottedHeavy"`, `"dash"`, `"dashHeavy"`, `"dashLong"`,
 * `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
 * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`) all
 * collapse to `undefined` — the writer emits only `"sng"`, so
 * reporting any non-single underline as `true` would silently
 * downgrade the choice to a single line on round-trip. Unknown /
 * malformed `u` tokens likewise drop to `undefined`.
 *
 * Mirrors {@link parseTitleUnderline} for axis titles — same
 * canonical-slot pair (`<a:defRPr>` carries the default-paragraph
 * underline flag, which the writer keeps in sync with the literal
 * run's `<a:rPr>` so the reader only needs to consult one of the two
 * slots), same drop-on-default semantics. The lookup is scoped to the
 * axis title's `<c:rich>` body so a stray `<a:defRPr>` elsewhere on
 * the axis (e.g. on the tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely or
 * when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the flag from
 * in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape per
 * the OOXML schema. Mirrors the chart-level title underline
 * {@link parseTitleUnderline} so a parsed value slots straight into
 * the writer-side {@link SheetChart.axes}.x.axisTitleUnderline.
 */
export function parseAxisTitleUnderline(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph underline flag. The reader walks the canonical
  // chain and bails on the first missing link so a malformed
  // `<c:rich>` surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p")
  if (!p) return undefined
  const pPr = findChild(p, "pPr")
  if (!pPr) return undefined
  const defRPr = findChild(pPr, "defRPr")
  if (!defRPr) return undefined
  const raw = defRPr.attrs.u
  // Only the UI-default `"sng"` surfaces as `true`. The OOXML
  // application default `"none"`, the non-UI `"dbl"` variant, and
  // every exotic token (`"words"`, `"heavy"`, `"dotted"`, etc.) all
  // collapse to `undefined` so absence and the OOXML default
  // round-trip identically through the writer; the writer emits only
  // `"sng"`, so reporting a non-single underline here would silently
  // downgrade the choice on round-trip.
  if (raw === "sng") return true
  return undefined
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>`
 * off the axis. Returns the typeface string the title was authored
 * with.
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
 * Mirrors the chart-level title typeface {@link parseTitleFontFamily} —
 * same canonical-slot pair (`<a:defRPr>` carries the default-paragraph
 * typeface, which the writer keeps in sync with the literal run's
 * `<a:rPr>` so the reader only needs to consult one of the two
 * slots). The lookup is scoped to the axis title's `<c:rich>` body
 * so a stray `<a:latin>` elsewhere on the axis (e.g. on the
 * tick-label `<c:txPr>`) cannot leak in.
 *
 * Returns `undefined` whenever the axis omits `<c:title>` entirely
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body — there is no `<a:p>` slot to surface the
 * typeface from in either case.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. The parsed value slots straight into the
 * writer-side {@link SheetChart.axes}.x.axisTitleFontFamily.
 */
export function parseAxisTitleFontFamily(axis: XmlElement): string | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const tx = findChild(title, "tx")
  if (!tx) return undefined
  const rich = findChild(tx, "rich")
  if (!rich) return undefined
  // `<a:p><a:pPr><a:defRPr><a:latin>` is the OOXML path Excel writes
  // for the default-paragraph typeface. The reader walks the
  // canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p")
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
 * Pull `<c:title><c:overlay val=".."/></c:title>` off an axis
 * element. Returns the axis-title overlay flag.
 *
 * The OOXML default `false` (the title reserves its own slot adjacent
 * to the axis, no overlap with the plot area) collapses to
 * `undefined` so absence and `<c:overlay val="0"/>` round-trip
 * identically through {@link cloneChart} — only an explicit
 * `<c:overlay val="1"/>` surfaces `true`.
 *
 * Returns `undefined` whenever the axis omits the `<c:title>`
 * element — there is no overlay slot to surface in that case. The
 * element is a sibling of `<c:tx>` inside `<c:title>` per the
 * CT_Title schema, so the lookup is scoped to direct title children
 * (a stray `<c:overlay>` elsewhere on the axis or chart cannot leak
 * in). Mirrors the chart-level title-overlay reader so a parsed
 * value flows straight back into the writer-side
 * {@link SheetChart.axes}.x.axisTitleOverlay.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` /
 * `"0"` / `"false"`); unknown values and missing `val` attributes
 * drop to `undefined` rather than fabricate a flag Excel would not
 * emit.
 */
export function parseAxisTitleOverlay(axis: XmlElement): boolean | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  const overlay = findChild(title, "overlay")
  if (!overlay) return undefined
  const raw = overlay.attrs.val
  if (typeof raw !== "string") return undefined
  switch (raw) {
    case "1":
    case "true":
      return true
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `axisTitleOverlay` field.
      return undefined
    default:
      return undefined
  }
}

/**
 * Pull `<c:title><c:layout><c:manualLayout>` off an axis element.
 * Reflects Excel's "Format Axis Title -> Title Options -> Position ->
 * Custom" placement — the `(x, y)` anchor and `(w, h)` size of the
 * axis-title block as fractions of the chart frame in the `0..1`
 * band.
 *
 * The OOXML schema (`CT_Title`, ECMA-376 Part 1, §21.2.2.210) places
 * `<c:layout>` inside `<c:title>` between `<c:tx>` and `<c:overlay>`,
 * and the `<c:manualLayout>` block (`CT_ManualLayout`, §21.2.2.115)
 * exposes optional `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` children
 * whose `val` attributes carry an `xsd:double`. The reader admits
 * the coordinate only when `val` parses to a finite number in the
 * `0..1` band; out-of-range / non-finite / non-numeric tokens drop
 * to `undefined` on the matching axis so absence and a malformed
 * token round-trip identically through {@link cloneChart}.
 *
 * Both `<c:xMode val="edge"/>` (absolute fraction of the chart frame)
 * and `<c:xMode val="factor"/>` (delta from auto-layout) are accepted
 * — the reader surfaces the same `ChartManualLayout` shape regardless,
 * since the writer always normalizes to `"edge"` on emit (Excel itself
 * emits the absolute form when the user drags an element to a custom
 * position).
 *
 * Returns `undefined` whenever the axis omits the `<c:title>` /
 * `<c:layout>` / `<c:manualLayout>` chain at any link, or when every
 * coordinate dropped on normalization — the field is omitted entirely
 * on a clean parse so absence and an empty layout round-trip
 * identically through the writer. Mirrors the chart-level
 * {@link parseLegendLayout} / {@link parsePlotAreaLayout} so the
 * three layout knobs share parsing semantics.
 */
export function parseAxisTitleLayout(axis: XmlElement): ChartManualLayout | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  return parseManualLayout(title)
}

/**
 * Pull `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:title>` off an axis element. Returns the
 * axis title's solid fill color as a 6-character uppercase hex string
 * the writer can round-trip via
 * {@link SheetChart.axes.x.axisTitleFillColor}. Mirrors
 * {@link parseTitleFillColor} for axis titles — same canonical
 * `<c:spPr><a:solidFill><a:srgbClr>` chain, same drop-on-non-sRGB
 * semantics.
 *
 * Returns the 6-character uppercase hex string when the parser walks
 * the full chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme
 * references (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and
 * `<a:prstClr>` all collapse to `undefined` — only the literal RGB
 * triple round-trips losslessly through {@link writeChart}. Non-solid
 * fills (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`, `<a:blipFill>`)
 * likewise drop to `undefined` so a round-trip never fabricates a
 * fill the writer cannot reproduce on emit. Malformed `val` tokens
 * (wrong length, non-hex characters) drop to `undefined` rather than
 * fabricate a value the writer would round-trip into a malformed
 * `<a:srgbClr>`.
 *
 * The lookup is scoped to direct children of the axis's `<c:title>`
 * so a stray `<c:spPr>` elsewhere on the axis (e.g. on a `<c:txPr>`
 * tick-label block, or on the axis itself) cannot leak in. Returns
 * `undefined` whenever the axis omits the `<c:title>` element or the
 * `<c:spPr><a:solidFill><a:srgbClr>` chain is malformed at any link.
 *
 * Independent of {@link parseAxisTitleColor}: the fill lives on
 * `<c:title><c:spPr>`, the font color lives on
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>` — the
 * two readers walk disjoint paths so a caller can pin both knobs
 * without conflict. Unlike {@link parseAxisTitleColor}, the lookup is
 * on `<c:title>` directly rather than gated on `<c:rich>` so a title
 * authored as a `<c:strRef>` formula reference can still surface its
 * background fill — Excel's "Format Axis Title -> Fill" dialog is
 * independent of whether the text body is rich or a formula.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title fill
 * {@link parseTitleFillColor} so a parsed value slots straight into
 * the writer-side {@link SheetChart.axes.x.axisTitleFillColor}.
 */
export function parseAxisTitleFillColor(axis: XmlElement): ChartColor | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  return parseSpPrFill(title)
}

/**
 * Pull `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:title>` off an axis element.
 * Returns the axis title's line stroke color as a 6-character
 * uppercase hex string the writer can round-trip via
 * {@link SheetChart.axes.x.axisTitleBorderColor}. Mirrors
 * {@link parseTitleBorderColor} for axis titles — same canonical
 * `<c:spPr><a:ln><a:solidFill><a:srgbClr>` chain, same drop-on-non-
 * sRGB semantics.
 *
 * Returns the 6-character uppercase hex string when the parser walks
 * the full chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme
 * references (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and
 * `<a:prstClr>` all collapse to `undefined` — only the literal RGB
 * triple round-trips losslessly through {@link writeChart}. Non-solid
 * line fills (`<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`) likewise
 * drop to `undefined` so a round-trip never fabricates a stroke the
 * writer cannot reproduce on emit. Malformed `val` tokens (wrong
 * length, non-hex characters) drop to `undefined` rather than
 * fabricate a value the writer would round-trip into a malformed
 * `<a:srgbClr>`.
 *
 * The lookup is scoped to direct children of the axis's `<c:title>`
 * so a stray `<c:spPr>` elsewhere on the axis (e.g. on a `<c:txPr>`
 * tick-label block, or on the axis itself), or on the chart-level
 * `<c:title>`, cannot leak in. Returns `undefined` whenever the axis
 * omits the `<c:title>` element or the `<c:spPr><a:ln><a:solidFill>
 * <a:srgbClr>` chain is malformed at any link.
 *
 * Independent of {@link parseAxisTitleFillColor} (the fill on the
 * same `<c:spPr>` block) and {@link parseAxisTitleColor} (the font
 * color on the inner `<a:defRPr><a:solidFill>` slot inside
 * `<c:tx><c:rich><a:p><a:pPr>`) — the three readers walk disjoint
 * children of the shared `<c:title>` block so a caller can pin all
 * three knobs without conflict. Unlike {@link parseAxisTitleColor},
 * the lookup is on `<c:title>` directly rather than gated on
 * `<c:rich>` so a title authored as a `<c:strRef>` formula reference
 * can still surface its border color — Excel's "Format Axis Title
 * -> Border" dialog is independent of whether the text body is rich
 * or a formula.
 *
 * Sits on every axis flavour — `<c:catAx>` / `<c:valAx>` /
 * `<c:dateAx>` / `<c:serAx>` all share the same `<c:title>` shape
 * per the OOXML schema. Mirrors the chart-level title border
 * {@link parseTitleBorderColor} so a parsed value slots straight
 * into the writer-side {@link SheetChart.axes.x.axisTitleBorderColor}.
 */
export function parseAxisTitleBorderColor(axis: XmlElement): ChartColor | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  return parseSpPrBorderColor(title)
}

/**
 * Pull the `w` attribute off `<c:catAx><c:title><c:spPr><a:ln w="EMU">`
 * (or `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`) and return the stroke
 * width in points after clamping to the `0.25..13.5` pt band Excel's
 * UI exposes. Same EMU encoding (1 pt = 12 700 EMU) and snap / clamp
 * grammar as every other chart-frame border-width slot. Delegates to
 * {@link parseBorderWidthFromSpPr} so the implementation stays
 * uniform across all hosts.
 *
 * Returns `undefined` when the axis omits `<c:title>`, when the title
 * has no `<c:spPr><a:ln w=..>` slot, when the attribute is absent,
 * when the value cannot be parsed as a finite positive number, or
 * when it parses to zero. Independent of {@link parseAxisTitleBorderColor}
 * and {@link parseAxisTitleBorderDash}: all three readers walk
 * disjoint slots of the shared `<a:ln>` element.
 */
export function parseAxisTitleBorderWidth(axis: XmlElement): number | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  return parseBorderWidthFromSpPr(title)
}

/**
 * Pull the `val` attribute off `<c:catAx><c:title><c:spPr><a:ln>
 * <a:prstDash val=".."/></a:ln></c:spPr></c:title></c:catAx>` (or
 * `<c:valAx>` / `<c:dateAx>` / `<c:serAx>`) and return the recognized
 * {@link ChartBorderDash} value. Returns `undefined` when the chain
 * is missing at any link, when the attribute is absent / unrecognized,
 * or when it matches the OOXML default `"solid"`.
 */
export function parseAxisTitleBorderDash(axis: XmlElement): ChartBorderDash | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  return parseBorderDashFromSpPr(title)
}

/**
 * Pull the `cap` attribute off `<c:title><c:spPr><a:ln cap=".."/>`
 * scoped to an axis element. Returns the {@link ChartLineCap} or
 * `undefined` for absence / OOXML default `"flat"`.
 */
export function parseAxisTitleBorderCap(axis: XmlElement): ChartLineCap | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  return parseBorderCapFromSpPr(title)
}

/**
 * Pull the `cmpd` attribute off `<c:title><c:spPr><a:ln cmpd=".."/>`
 * scoped to an axis element. Returns the {@link ChartLineCompound} or
 * `undefined` for absence / OOXML default `"sng"`.
 */
export function parseAxisTitleBorderCompound(axis: XmlElement): ChartLineCompound | undefined {
  const title = findChild(axis, "title")
  if (!title) return undefined
  return parseBorderCompoundFromSpPr(title)
}

/**
 * Pull `<c:autoTitleDeleted val=".."/>` off `<c:chart>`. Surfaces
 * `true` only when the chart pinned `<c:autoTitleDeleted val="1"/>`
 * (the non-default state — the user explicitly deleted the
 * auto-generated title that single-series charts synthesise from the
 * series name). The OOXML default `val="0"` and absence both collapse
 * to `undefined` so absence and the default round-trip identically
 * through {@link cloneChart}.
 *
 * The element is independent of `<c:title>` — it sits on `<c:chart>`
 * directly (between `<c:title>` and `<c:plotArea>` per CT_Chart,
 * ECMA-376 Part 1, §21.2.2.4), not nested inside `<c:title>`. A chart
 * with a literal `<c:title>` typically pins `val="0"` because the user
 * has not deleted the auto-title (they overrode it with a literal
 * one); a chart with no `<c:title>` may pin `val="1"` to suppress
 * Excel's auto-title synthesis or omit the element entirely (Excel
 * may still synthesise an auto-title in that case for a single-series
 * chart).
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"`
 * / `"false"`); unknown values and missing `val` attributes drop to
 * `undefined` rather than fabricate a flag Excel would not emit.
 */
export function parseAutoTitleDeleted(chartEl: XmlElement): boolean | undefined {
  const el = findChild(chartEl, "autoTitleDeleted")
  if (!el) return undefined
  const raw = el.attrs.val
  if (typeof raw !== "string") return undefined
  switch (raw) {
    case "1":
    case "true":
      return true
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `autoTitleDeleted` field.
      return undefined
    default:
      return undefined
  }
}

// ── Writer-side axis types and constants ──────────────────────────

export interface AxisRenderOptions {
  xAxisTitle: string | undefined
  yAxisTitle: string | undefined
  /**
   * Axis-title rotation in whole degrees emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`.
   * The OOXML `rot` attribute is in 60000ths of a degree; the writer
   * converts at emit time. Range: `-90..90` (Excel's UI band).
   * `undefined` collapses to the OOXML default `0` so a fresh chart
   * matches Excel's reference serialization byte-for-byte. Only
   * meaningful when the axis renders a title — the per-family axis
   * builders gate the value on the `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleRotation: number | undefined
  /**
   * Axis-title rotation in whole degrees emitted on the Y axis. Same
   * shape and conversion semantics as {@link xAxisTitleRotation}.
   */
  yAxisTitleRotation: number | undefined
  /**
   * Axis-title font size in points emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
   * <a:r><a:rPr sz="N"/></a:r></a:p></c:rich></c:tx></c:title>`. The
   * OOXML `sz` attribute is in 100ths of a point; the writer converts
   * at emit time. Range: `1..400`pt. `undefined` collapses to the
   * hardcoded `1000` (10pt) default Excel itself emits on a fresh
   * axis title. Only meaningful when the axis renders a title — the
   * per-family axis builders gate the value on the `xAxisTitle` /
   * `yAxisTitle` field.
   */
  xAxisTitleFontSize: number | undefined
  /**
   * Axis-title font size in points emitted on the Y axis. Same shape
   * and conversion semantics as {@link xAxisTitleFontSize}.
   */
  yAxisTitleFontSize: number | undefined
  /**
   * Axis-title bold flag emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
   * <a:r><a:rPr b=".."/></a:r></a:p></c:rich></c:tx></c:title>`. The
   * OOXML `b` attribute is the `xsd:boolean` bold flag on
   * `CT_TextCharacterProperties`; the writer emits `1` / `0` at the
   * canonical slots. `undefined` collapses to the OOXML default `0`
   * (non-bold) so a fresh chart matches Excel's reference
   * serialization byte-for-byte. Only meaningful when the axis
   * renders a title — the per-family axis builders gate the value on
   * the `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleBold: boolean | undefined
  /**
   * Axis-title bold flag emitted on the Y axis. Same shape and
   * semantics as {@link xAxisTitleBold}.
   */
  yAxisTitleBold: boolean | undefined
  /**
   * Axis-title italic flag emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
   * <a:r><a:rPr i=".."/></a:r></a:p></c:rich></c:tx></c:title>`. The
   * OOXML attribute is the `xsd:boolean` `i` on
   * `CT_TextCharacterProperties`. `undefined` and `false` both collapse
   * to omitting the attribute (Excel's reference serialization for a
   * non-italic axis title); only `true` emits `i="1"`. Only meaningful
   * when the axis renders a title — the per-family axis builders gate
   * the value on the `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleItalic: boolean | undefined
  /**
   * Axis-title italic flag emitted on the Y axis. Same shape and emit
   * semantics as {@link xAxisTitleItalic}.
   */
  yAxisTitleItalic: boolean | undefined
  /**
   * Axis-title font color emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
   * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr>
   * <a:r><a:rPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
   * </a:rPr></a:r></a:p></c:rich></c:tx></c:title>`. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color. `undefined` collapses to omitting the entire
   * `<a:solidFill>` block (Excel's reference serialization for an
   * axis title that inherits the theme text color); any non-`undefined`
   * value is the normalized uppercase hex string the writer lands on
   * both `<a:defRPr>` and `<a:rPr>`. Only meaningful when the axis
   * renders a title — the per-family axis builders gate the value on
   * the `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleColor: ChartColor | undefined
  /**
   * Axis-title font color emitted on the Y axis. Same shape and emit
   * semantics as {@link xAxisTitleColor}.
   */
  yAxisTitleColor: ChartColor | undefined
  /**
   * Axis-title strikethrough flag emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
   * <a:r><a:rPr strike=".."/></a:r></a:p></c:rich></c:tx></c:title>`.
   * The OOXML attribute is the `ST_TextStrikeType` enum on
   * `CT_TextCharacterProperties`. `undefined` and `false` both collapse
   * to omitting the attribute (Excel's reference serialization for a
   * non-strikethrough axis title); only `true` emits
   * `strike="sngStrike"` (Excel's UI checkbox — single line). Only
   * meaningful when the axis renders a title — the per-family axis
   * builders gate the value on the `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleStrike: boolean | undefined
  /**
   * Axis-title strikethrough flag emitted on the Y axis. Same shape
   * and emit semantics as {@link xAxisTitleStrike}.
   */
  yAxisTitleStrike: boolean | undefined
  /**
   * Axis-title underline flag emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
   * <a:r><a:rPr u=".."/></a:r></a:p></c:rich></c:tx></c:title>`. The
   * OOXML attribute is the `ST_TextUnderlineType` enum on
   * `CT_TextCharacterProperties`. `undefined` and `false` both collapse
   * to omitting the attribute (Excel's reference serialization for a
   * non-underlined axis title); only `true` emits `u="sng"` (Excel's UI
   * checkbox — single line). Only meaningful when the axis renders a
   * title — the per-family axis builders gate the value on the
   * `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleUnderline: boolean | undefined
  /**
   * Axis-title underline flag emitted on the Y axis. Same shape
   * and emit semantics as {@link xAxisTitleUnderline}.
   */
  yAxisTitleUnderline: boolean | undefined
  /**
   * Axis-title font family / typeface emitted on the X axis via
   * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
   * typeface=".."/></a:defRPr></a:pPr><a:r><a:rPr><a:latin
   * typeface=".."/></a:rPr></a:r></a:p></c:rich></c:tx></c:title>`.
   * The OOXML `<a:latin typeface=".."/>` element carries the
   * typeface name on `CT_TextFont`. `undefined` collapses to
   * omitting the element (Excel's reference serialization for an
   * axis title that inherits the theme typeface); a non-empty
   * trimmed string emits `<a:latin typeface=".."/>` on both the
   * default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>`.
   * Only meaningful when the axis renders a title — the per-family
   * axis builders gate the value on the `xAxisTitle` / `yAxisTitle`
   * field.
   */
  xAxisTitleFontFamily: string | undefined
  /**
   * Axis-title font family / typeface emitted on the Y axis. Same
   * shape and emit semantics as {@link xAxisTitleFontFamily}.
   */
  yAxisTitleFontFamily: string | undefined
  /**
   * Axis-title overlay flag emitted on the X axis via
   * `<c:catAx><c:title><c:overlay val=".."/></c:title></c:catAx>`.
   * The OOXML `<c:overlay val=".."/>` element carries the boolean.
   * `false` (the OOXML default) emits `val="0"` so the title
   * reserves its own slot adjacent to the axis; `true` emits
   * `val="1"` so the title overlaps the plot area. The writer
   * always emits `<c:overlay>` because Excel's reference
   * serialization includes it on every visible axis title — only
   * the `val` flips. Only meaningful when the axis renders a title
   * — the per-family axis builders gate the value on the
   * `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleOverlay: boolean
  /**
   * Axis-title overlay flag emitted on the Y axis. Same shape and
   * emit semantics as {@link xAxisTitleOverlay}.
   */
  yAxisTitleOverlay: boolean
  /**
   * Axis-title manual placement emitted on the X axis via
   * `<c:title><c:layout><c:manualLayout>...</c:manualLayout></c:layout>
   * </c:title>`. The OOXML `CT_ManualLayout` block (ECMA-376 Part 1,
   * §21.2.2.115) sits inside `CT_Title` between `<c:tx>` and
   * `<c:overlay>` and carries the title's `(x, y)` anchor and
   * `(w, h)` size as fractions of the chart frame in the `0..1` band.
   * Absence (`undefined`) collapses to omitting the entire `<c:layout>`
   * block so the axis title renders at Excel's auto-layout position.
   * Only meaningful when the axis renders a title — the per-family
   * axis builders gate the value on the `xAxisTitle` / `yAxisTitle`
   * field. Mirrors the chart-level `titleLayout` / `legendLayout` /
   * `plotAreaLayout` slots so the four manual-layout knobs share a
   * normalization grammar (`normalizeManualLayout`).
   */
  xAxisTitleLayout: ResolvedManualLayout | undefined
  /**
   * Axis-title manual placement emitted on the Y axis. Same shape
   * and emit semantics as {@link xAxisTitleLayout}.
   */
  yAxisTitleLayout: ResolvedManualLayout | undefined
  /**
   * Axis-title background fill emitted on the X axis via
   * `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
   * </a:solidFill></c:spPr></c:title>`. The OOXML `<a:srgbClr
   * val=".."/>` carries the 6-character uppercase hex sRGB color
   * (CT_SRgbColor inside CT_ShapeProperties' fill choice — ECMA-376
   * Part 1, §20.1.2.3.32 / §20.1.8.54). `undefined` collapses to
   * omitting the entire `<c:spPr>` block (Excel's reference
   * serialization for an axis title that inherits the theme default
   * fill — typically a transparent title background); any
   * non-`undefined` value is the normalized uppercase hex string the
   * writer lands on the title's `<c:spPr>` slot. Only meaningful
   * when the axis renders a title — the per-family axis builders
   * gate the value on the `xAxisTitle` / `yAxisTitle` field.
   */
  xAxisTitleFillColor: ChartColor | undefined
  /**
   * Axis-title background fill emitted on the Y axis. Same shape and
   * emit semantics as {@link xAxisTitleFillColor}.
   */
  yAxisTitleFillColor: ChartColor | undefined
  /**
   * Axis-title border (line stroke) color emitted on the X axis via
   * `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
   * </a:solidFill></a:ln></c:spPr></c:title>`. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (CT_SRgbColor inside the line's solid fill choice —
   * ECMA-376 Part 1, §20.1.2.3.32 / §20.1.2.3.24). `undefined`
   * collapses to omitting the `<a:ln>` block (Excel's reference
   * serialization for an axis title that inherits the auto-stroke —
   * typically no visible border); any non-`undefined` value is the
   * normalized uppercase hex string the writer lands on the title's
   * `<c:spPr><a:ln>` slot. Composes independently with
   * {@link xAxisTitleFillColor} — the fill lands on
   * `<c:spPr><a:solidFill>`, the stroke lands on
   * `<c:spPr><a:ln><a:solidFill>`; the writer authors a single
   * `<c:spPr>` whenever either knob is set with children in
   * CT_ShapeProperties schema order (fill before stroke). Only
   * meaningful when the axis renders a title — the per-family axis
   * builders gate the value on the `xAxisTitle` / `yAxisTitle`
   * field.
   */
  xAxisTitleBorderColor: ChartColor | undefined
  /**
   * Axis-title border emitted on the Y axis. Same shape and emit
   * semantics as {@link xAxisTitleBorderColor}.
   */
  yAxisTitleBorderColor: ChartColor | undefined
  /**
   * Axis-title border (line stroke) thickness emitted on the X axis
   * via `<c:title><c:spPr><a:ln w="EMU"/></c:spPr></c:title>`. The
   * OOXML `w` attribute is in EMU (1 pt = 12 700 EMU); the writer
   * converts at emit time. `undefined` collapses to omitting the
   * attribute (Excel's reference auto-thickness, typically 0.75 pt).
   * Composes independently with {@link xAxisTitleBorderColor} and
   * {@link xAxisTitleBorderDash} on the same `<a:ln>` element.
   * Only meaningful when the axis renders a title.
   */
  xAxisTitleBorderWidth: number | undefined
  /**
   * Axis-title border thickness emitted on the Y axis. Same shape and
   * emit semantics as {@link xAxisTitleBorderWidth}.
   */
  yAxisTitleBorderWidth: number | undefined
  /**
   * Axis-title border (line stroke) preset dash pattern emitted on the
   * X axis via `<c:title><c:spPr><a:ln><a:prstDash val=".."/></a:ln>
   * </c:spPr></c:title>`. The OOXML `<a:prstDash>` element follows
   * `<a:solidFill>` per CT_LineProperties schema sequence (ECMA-376
   * Part 1, §20.1.2.3.24). `undefined` (and the OOXML default
   * `"solid"`) collapses to omitting the element so a fresh axis
   * title renders solid. Composes independently with
   * {@link xAxisTitleBorderColor} and {@link xAxisTitleBorderWidth}.
   */
  xAxisTitleBorderDash: ChartBorderDash | undefined
  /**
   * Axis-title border dash emitted on the Y axis. Same shape and emit
   * semantics as {@link xAxisTitleBorderDash}.
   */
  yAxisTitleBorderDash: ChartBorderDash | undefined
  /**
   * Axis-title border line cap style emitted on the X axis via
   * `<c:title><c:spPr><a:ln cap=".."/>`. Mirrors
   * {@link SheetChart.titleBorderCap}.
   */
  xAxisTitleBorderCap: ChartLineCap | undefined
  /**
   * Axis-title border line cap style emitted on the Y axis. Same shape
   * and emit semantics as {@link xAxisTitleBorderCap}.
   */
  yAxisTitleBorderCap: ChartLineCap | undefined
  /**
   * Axis-title border compound line style emitted on the X axis via
   * `<c:title><c:spPr><a:ln cmpd=".."/>`. Mirrors
   * {@link SheetChart.titleBorderCompound}.
   */
  xAxisTitleBorderCompound: ChartLineCompound | undefined
  /**
   * Axis-title border compound emitted on the Y axis. Same shape and
   * emit semantics as {@link xAxisTitleBorderCompound}.
   */
  yAxisTitleBorderCompound: ChartLineCompound | undefined
  xGridlines: { major: boolean; minor: boolean } | undefined
  yGridlines: { major: boolean; minor: boolean } | undefined
  xScale: ChartAxisScale | undefined
  yScale: ChartAxisScale | undefined
  xNumFmt: ChartAxisNumberFormat | undefined
  yNumFmt: ChartAxisNumberFormat | undefined
  xMajorTickMark: ChartAxisTickMark | undefined
  yMajorTickMark: ChartAxisTickMark | undefined
  xMinorTickMark: ChartAxisTickMark | undefined
  yMinorTickMark: ChartAxisTickMark | undefined
  xTickLblPos: ChartAxisTickLabelPosition | undefined
  yTickLblPos: ChartAxisTickLabelPosition | undefined
  /**
   * Tick-label rotation in whole degrees emitted on the X axis via
   * `<c:txPr><a:bodyPr rot="N"/></c:txPr>`. The OOXML `rot` attribute
   * is in 60000ths of a degree; the writer converts at emit time.
   * Range: `-90..90` (Excel's UI band). `undefined` skips the
   * `<c:txPr>` block entirely so a fresh chart matches Excel's
   * minimal serialization. Surfaces on every axis flavour because the
   * OOXML schema places `<c:txPr>` on `<c:catAx>` / `<c:valAx>` /
   * `<c:dateAx>` / `<c:serAx>` alike.
   */
  xLabelRotation: number | undefined
  /**
   * Tick-label rotation in whole degrees emitted on the Y axis. Same
   * shape and conversion semantics as {@link xLabelRotation}.
   */
  yLabelRotation: number | undefined
  /**
   * Tick-label font size in points emitted on the X axis via
   * `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`.
   * The OOXML `sz` attribute is in 100ths of a point; the writer
   * converts at emit time. Range: `1..400`pt (the OOXML
   * `ST_TextFontSize` band). `undefined` skips the size attribute so
   * a fresh chart inherits Excel's reference 10pt tick-label size.
   * The block is emitted whenever `xLabelRotation` or
   * `xLabelFontSize` is set so the OOXML schema's `<c:txPr>` slot
   * carries every pinned typography knob.
   */
  xLabelFontSize: number | undefined
  /**
   * Tick-label font size in points emitted on the Y axis. Same shape
   * and conversion semantics as {@link xLabelFontSize}.
   */
  yLabelFontSize: number | undefined
  /**
   * Tick-label bold flag emitted on the X axis via
   * `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>`.
   * The OOXML `b` attribute is the `xsd:boolean` bold flag on
   * `CT_TextCharacterProperties`; the writer emits `1` / `0` at the
   * canonical slot. `undefined` skips the attribute so a fresh chart
   * inherits the theme-default tick-label weight. The block is
   * emitted whenever `xLabelRotation`, `xLabelFontSize`, or
   * `xLabelBold` is set so the OOXML schema's `<c:txPr>` slot
   * carries every pinned typography knob.
   */
  xLabelBold: boolean | undefined
  /**
   * Tick-label bold flag emitted on the Y axis. Same shape and
   * semantics as {@link xLabelBold}.
   */
  yLabelBold: boolean | undefined
  /**
   * Tick-label italic flag emitted on the X axis via
   * `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>`.
   * The OOXML `i` attribute is the `xsd:boolean` italic flag on
   * `CT_TextCharacterProperties`; the writer emits `1` / `0` at the
   * canonical slot. `undefined` skips the attribute so a fresh chart
   * inherits the theme-default tick-label slant. The block is
   * emitted whenever `xLabelRotation`, `xLabelFontSize`, `xLabelBold`,
   * or `xLabelItalic` is set so the OOXML schema's `<c:txPr>` slot
   * carries every pinned typography knob.
   */
  xLabelItalic: boolean | undefined
  /**
   * Tick-label italic flag emitted on the Y axis. Same shape and
   * semantics as {@link xLabelItalic}.
   */
  yLabelItalic: boolean | undefined
  /**
   * Tick-label font color emitted on the X axis via
   * `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
   * </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>`. The OOXML
   * `<a:srgbClr val=".."/>` carries the 6-character uppercase hex
   * sRGB color (`CT_SRgbColor` inside `CT_TextCharacterProperties`'
   * fill choice — ECMA-376 Part 1, §20.1.2.3.32 / §21.1.2.3.7).
   * `undefined` collapses to omitting the entire `<a:solidFill>`
   * block (Excel's reference serialization for tick labels that
   * inherit the theme text color); any non-`undefined` value is the
   * normalized uppercase hex string the writer lands on the
   * default-paragraph `<a:defRPr>` slot. The block is emitted
   * whenever `xLabelRotation`, `xLabelFontSize`, `xLabelBold`,
   * `xLabelItalic`, or `xLabelColor` is set so the OOXML schema's
   * `<c:txPr>` slot carries every pinned typography knob.
   */
  xLabelColor: ChartColor | undefined
  /**
   * Tick-label font color emitted on the Y axis. Same shape and
   * semantics as {@link xLabelColor}.
   */
  yLabelColor: ChartColor | undefined
  /**
   * Tick-label underline flag emitted on the X axis via
   * `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>`.
   * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer emits only the UI variant `"sng"` when the input is
   * `true`. `undefined` and `false` both collapse to omitting the
   * attribute so a fresh chart inherits Excel's reference
   * non-underlined tick labels (the OOXML default `"none"` collapses
   * to absence). The block is emitted whenever any tick-label
   * typography knob (`xLabelRotation`, `xLabelFontSize`,
   * `xLabelBold`, `xLabelItalic`, `xLabelColor`, or
   * `xLabelUnderline`) is set so the OOXML schema's `<c:txPr>` slot
   * carries every pinned typography knob.
   */
  xLabelUnderline: boolean | undefined
  /**
   * Tick-label underline flag emitted on the Y axis. Same shape and
   * semantics as {@link xLabelUnderline}.
   */
  yLabelUnderline: boolean | undefined
  /**
   * Tick-label strikethrough flag emitted on the X axis via
   * `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>`.
   * The OOXML `strike` attribute is the `ST_TextStrikeType` enum on
   * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7); the
   * writer emits only the UI variant `"sngStrike"` when the input is
   * `true`. `undefined` and `false` both collapse to omitting the
   * attribute so a fresh chart inherits Excel's reference non-
   * strikethrough tick labels (the OOXML default `"noStrike"`
   * collapses to absence). The block is emitted whenever any tick-
   * label typography knob (`xLabelRotation`, `xLabelFontSize`,
   * `xLabelBold`, `xLabelItalic`, `xLabelColor`, `xLabelUnderline`,
   * or `xLabelStrike`) is set so the OOXML schema's `<c:txPr>` slot
   * carries every pinned typography knob.
   */
  xLabelStrike: boolean | undefined
  /**
   * Tick-label strikethrough flag emitted on the Y axis. Same shape
   * and semantics as {@link xLabelStrike}.
   */
  yLabelStrike: boolean | undefined
  /**
   * Tick-label font family / typeface emitted on the X axis via
   * `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
   * </a:pPr></a:p></c:txPr>`. The OOXML `<a:latin>` element carries
   * the typeface name on `CT_TextFont`. `undefined` collapses to
   * omitting the element (Excel's reference serialization for tick
   * labels that inherit the theme typeface); a non-empty trimmed
   * string emits `<a:latin typeface=".."/>`. The block is emitted
   * whenever any tick-label typography knob is set so the OOXML
   * schema's `<c:txPr>` slot carries every pinned typography knob.
   */
  xLabelFontFamily: string | undefined
  /**
   * Tick-label font family / typeface emitted on the Y axis. Same
   * shape and semantics as {@link xLabelFontFamily}.
   */
  yLabelFontFamily: string | undefined
  xReverse: boolean
  yReverse: boolean
  /**
   * Tick-label skip interval emitted on the X axis only when the axis
   * is `<c:catAx>` (i.e. bar / column / line / area). Scatter charts
   * have no category axis, so the skip is dropped silently.
   */
  xTickLblSkip: number | undefined
  /**
   * Tick-mark skip interval emitted on the X axis only when the axis
   * is `<c:catAx>`. Same scope rule as {@link xTickLblSkip}.
   */
  xTickMarkSkip: number | undefined
  /**
   * Label offset percentage emitted on the X axis only when the axis
   * is `<c:catAx>` (i.e. bar / column / line / area). Scatter charts
   * have no category axis, so the value is dropped silently.
   */
  xLblOffset: number | undefined
  /**
   * Tick-label horizontal alignment emitted on the X axis only when
   * the axis is `<c:catAx>`. Scatter charts have no category axis, so
   * the value is dropped silently. `undefined` means absent (the
   * writer falls back to the OOXML default `"ctr"`).
   */
  xLblAlgn: ChartAxisLabelAlign | undefined
  /**
   * Whether the X axis should pin `<c:noMultiLvlLbl val="1"/>`
   * (multi-level category labels suppressed). Always defined — `false`
   * keeps Excel's reference `val="0"` while `true` collapses multi-tier
   * category labels onto a single line. Only meaningful for the catAx
   * builder; scatter has no category axis, so the value is silently
   * dropped at the per-chart-type branch.
   */
  xNoMultiLvlLbl: boolean
  /**
   * Whether the X axis should render its `<c:auto>` element with
   * `val="1"` (Excel's default — auto-detect whether the axis is a
   * date axis or category axis). Always defined — `true` keeps Excel's
   * reference `val="1"` while `false` pins the axis as a literal
   * category axis (Excel's "Text axis" radio under "Format Axis -> Axis
   * Options"). Only meaningful for the catAx builder; scatter has no
   * category axis, so the value is silently dropped at the per-chart-
   * type branch.
   */
  xAuto: boolean
  /**
   * Whether the X axis should render its `<c:delete>` element with
   * `val="1"` (axis hidden). Always defined — `false` keeps Excel's
   * reference `val="0"` while `true` collapses the axis line, ticks,
   * and labels off the rendered chart.
   */
  xHidden: boolean
  /** Whether the Y axis should render hidden. Same shape as {@link xHidden}. */
  yHidden: boolean
  /**
   * Resolved axis-crosses pin for the X axis. The XSD choice between
   * `<c:crosses>` and `<c:crossesAt>` is collapsed to a single tagged
   * union: `kind: "default"` emits the OOXML default `<c:crosses
   * val="autoZero"/>`, `kind: "semantic"` emits the resolved
   * {@link ChartAxisCrosses} token, and `kind: "numeric"` emits
   * `<c:crossesAt>` with the literal value the caller pinned.
   */
  xCrosses: ResolvedAxisCrosses
  /** Resolved axis-crosses pin for the Y axis. Same shape as {@link xCrosses}. */
  yCrosses: ResolvedAxisCrosses
  /**
   * Display-unit preset emitted on the X axis only when the axis is
   * `<c:valAx>` (i.e. scatter charts). Bar / column / line / area route
   * the X axis through `<c:catAx>` which rejects `<c:dispUnits>`, so
   * the catAx builder ignores this field.
   */
  xDispUnits: ChartAxisDispUnits | undefined
  /**
   * Display-unit preset emitted on the value axis. The catAx builder
   * (bar / column / line / area) routes the Y axis through `<c:valAx>`,
   * and the scatter builder routes both axes through `<c:valAx>` — so
   * this field surfaces on every chart family that has a value axis.
   * Pie / doughnut have no axes at all and the caller already
   * short-circuits those branches.
   */
  yDispUnits: ChartAxisDispUnits | undefined
  /**
   * Cross-between override for the X axis. Only honoured on scatter
   * (the X axis is a value axis there); the catAx builder ignores it
   * because `<c:crossBetween>` is value-axis-only per ECMA-376
   * §21.2.2.10. `undefined` falls back to the per-family default each
   * axis builder pins today.
   */
  xCrossBetween: ChartAxisCrossBetween | undefined
  /**
   * Cross-between override for the value axis. The catAx builder (bar
   * / column / line / area) routes the Y axis through `<c:valAx>`, and
   * the scatter builder routes both axes through `<c:valAx>` — so this
   * field surfaces on every chart family that has a value axis. Pie /
   * doughnut have no axes at all and the caller already short-circuits
   * those branches.
   */
  yCrossBetween: ChartAxisCrossBetween | undefined
}

/**
 * Resolved per-axis crossing pin. The OOXML schema places `<c:crosses>`
 * and `<c:crossesAt>` in an XSD choice — only one may appear at a time.
 * `normalizeAxisCrosses` collapses the writer's two input fields
 * (`crosses` and `crossesAt`) into this tagged union so the per-family
 * axis builders can emit the right element without re-implementing the
 * precedence rule.
 */
type ResolvedAxisCrosses =
  | { kind: "default" }
  | { kind: "semantic"; value: ChartAxisCrosses }
  | { kind: "numeric"; value: number }

/** Recognized values of `<c:crosses>` per the OOXML `ST_Crosses` enum. */
const VALID_AXIS_CROSSES: ReadonlySet<ChartAxisCrosses> = new Set(["autoZero", "min", "max"])

export const AXIS_ID_CAT = 111111111
export const AXIS_ID_VAL = 222222222
export const AXIS_ID_VAL_X = 333333333
export const AXIS_ID_VAL_Y = 444444444

/**
 * Application-default `sz` value for an axis title's `<a:defRPr>` /
 * `<a:rPr>` slots — Excel renders axis titles at 10pt (`sz="1000"`)
 * unless the user pins a custom size. Absence of
 * {@link SheetChart.axes.x.axisTitleFontSize} resolves to this default
 * so a fresh chart matches Excel's reference serialization byte-for-
 * byte, and round-trips of templates that never pinned the field stay
 * stable across the parse -> clone -> write loop.
 */
const AXIS_TITLE_DEFAULT_FONT_SIZE_SZ = 1000

const TITLE_ROT_PER_DEGREE = TXPR_ROT_PER_DEGREE

/** Recognized values of `<c:majorTickMark>` / `<c:minorTickMark>` (writer-side). */
const TICK_MARK_VALUES: ReadonlySet<ChartAxisTickMark> = new Set(["none", "in", "out", "cross"])

/** Recognized values of `<c:tickLblPos>` (writer-side). */
const TICK_LBL_POS_VALUES: ReadonlySet<ChartAxisTickLabelPosition> = new Set([
  "nextTo",
  "low",
  "high",
  "none",
])

// ── Writer ────────────────────────────────────────────────────────

/**
 * Resolve `<c:autoTitleDeleted val=".."/>` from
 * {@link SheetChart.autoTitleDeleted}.
 *
 * The element records whether the user explicitly deleted the
 * auto-generated title that single-series charts synthesise from the
 * series name. The flag is independent of whether a literal
 * `<c:title>` is emitted — a chart with no title may still pin
 * `val="0"` to let Excel synthesise the auto-title, or `val="1"` to
 * suppress it.
 *
 * When the caller pins {@link SheetChart.autoTitleDeleted} explicitly
 * the literal boolean value wins. When the field is omitted the writer
 * derives the value from the title presence so back-compat holds: a
 * chart with a literal title (and `showTitle !== false`) emits
 * `val="0"` so Excel keeps the literal visible; a chart with no
 * literal title emits `val="1"` so Excel does not silently grow an
 * auto-title from the series name.
 *
 * Anything other than literal `true` / `false` collapses to the
 * derived default so a stray non-boolean leaking through the type
 * guard (e.g. `0` / `1` / `"true"` / `null`) never inverts the
 * derivation. This matches how `titleOverlay` / `roundedCorners` /
 * `plotVisOnly` treat their inputs.
 */
export function resolveAutoTitleDeleted(chart: SheetChart): boolean {
  if (chart.autoTitleDeleted === true) return true
  if (chart.autoTitleDeleted === false) return false
  // Derive from title presence — preserves back-compat for callers
  // that never set the field.
  const showTitle = chart.showTitle ?? Boolean(chart.title)
  return !(showTitle && chart.title)
}

/**
 * Normalize an axis title input to either a non-empty trimmed string
 * or `undefined`. Empty strings are dropped so the writer never emits
 * an empty `<c:title>` element (Excel renders that as an unintended
 * blank label).
 */
export function normalizeAxisTitle(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : undefined
}

/**
 * Resolve the gridline toggles to a stable record (or `undefined` when
 * neither is on). Mirrors {@link normalizeAxisTitle} so the per-branch
 * code in `buildPlotArea` only needs a single null check.
 */
export function normalizeAxisGridlines(
  value: ChartAxisGridlines | undefined,
): { major: boolean; minor: boolean } | undefined {
  if (!value) return undefined
  const major = value.major === true
  const minor = value.minor === true
  if (!major && !minor) return undefined
  return { major, minor }
}

/**
 * Build the `<c:majorGridlines>` / `<c:minorGridlines>` block for an
 * axis. The returned XML fragments must be appended in spec order
 * (major before minor) and slot in immediately after `<c:axPos>`,
 * before the optional `<c:title>`. Excel's strict-validator rejects
 * any other position.
 */
export function buildAxisGridlines(
  gridlines: { major: boolean; minor: boolean } | undefined,
): string[] {
  if (!gridlines) return []
  const out: string[] = []
  if (gridlines.major) out.push(xmlElement("c:majorGridlines", undefined, []))
  if (gridlines.minor) out.push(xmlElement("c:minorGridlines", undefined, []))
  return out
}

/**
 * Drop fields that won't survive Excel's strict validator. Non-finite
 * numbers, `min >= max`, and zero/negative tick spacings all collapse
 * the corresponding entry to `undefined` so the writer never emits a
 * `<c:min>`/`<c:max>`/`<c:majorUnit>`/`<c:minorUnit>` Excel would
 * reject.
 *
 * Returns `undefined` when nothing usable remains so the writer can
 * skip the entire `<c:scaling>` augmentation.
 */
export function normalizeAxisScale(value: ChartAxisScale | undefined): ChartAxisScale | undefined {
  if (!value) return undefined
  const out: ChartAxisScale = {}
  if (typeof value.min === "number" && Number.isFinite(value.min)) out.min = value.min
  if (typeof value.max === "number" && Number.isFinite(value.max)) out.max = value.max
  if (out.min !== undefined && out.max !== undefined && out.min >= out.max) {
    // min >= max is meaningless; preserve the user-supplied min only
    // so validators don't choke on a flipped/empty axis range.
    delete out.max
  }
  if (
    typeof value.majorUnit === "number" &&
    Number.isFinite(value.majorUnit) &&
    value.majorUnit > 0
  ) {
    out.majorUnit = value.majorUnit
  }
  if (
    typeof value.minorUnit === "number" &&
    Number.isFinite(value.minorUnit) &&
    value.minorUnit > 0
  ) {
    out.minorUnit = value.minorUnit
  }
  if (
    typeof value.logBase === "number" &&
    Number.isFinite(value.logBase) &&
    value.logBase >= 2 &&
    value.logBase <= 1000
  ) {
    out.logBase = value.logBase
  }
  return Object.keys(out).length > 0 ? out : undefined
}

/**
 * Normalize a tick-label number format to a value the writer can emit.
 * An empty `formatCode` collapses the whole record — Excel rejects
 * `<c:numFmt formatCode=""/>`.
 */
export function normalizeAxisNumberFormat(
  value: ChartAxisNumberFormat | undefined,
): ChartAxisNumberFormat | undefined {
  if (!value) return undefined
  const formatCode = typeof value.formatCode === "string" ? value.formatCode : ""
  if (formatCode.length === 0) return undefined
  const out: ChartAxisNumberFormat = { formatCode }
  if (value.sourceLinked === true) out.sourceLinked = true
  return out
}

/**
 * Normalize a `tickLblSkip` / `tickMarkSkip` value to a positive
 * integer in the OOXML `ST_SkipIntervals` band (`1..32767`).
 *
 * Returns `undefined` when:
 *   - the input is missing or non-finite,
 *   - the rounded value is `1` (the OOXML default — show every label /
 *     mark — and what absence already means),
 *   - the rounded value falls outside the `1..32767` range.
 *
 * Out-of-range values drop rather than clamp because a skip count of
 * `100` and `32767` mean structurally different things to Excel — a
 * silent clamp would mask the configuration error rather than reveal
 * it.
 */
export function normalizeAxisSkip(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined
  const rounded = Math.round(value)
  if (rounded < 1 || rounded > 32767) return undefined
  if (rounded === 1) return undefined
  return rounded
}

/**
 * Normalize a category-axis `lblOffset` percentage to an integer in
 * the OOXML `ST_LblOffsetPercent` band (`0..1000`).
 *
 * Returns `undefined` when:
 *   - the input is missing or non-finite,
 *   - the rounded value is `100` (the OOXML default — Excel's
 *     reference label spacing — and what absence already means),
 *   - the rounded value falls outside the `0..1000` range.
 *
 * Out-of-range values drop rather than clamp so a malformed override
 * surfaces as "no offset emitted" instead of silently snapping to the
 * extreme — a clamp from `9999` to `1000` would mask a programming
 * error in the caller.
 */
export function normalizeAxisLblOffset(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined
  const rounded = Math.round(value)
  if (rounded < 0 || rounded > 1000) return undefined
  if (rounded === 100) return undefined
  return rounded
}

/**
 * Normalize a category-axis `lblAlgn` value to one of the three OOXML
 * `ST_LblAlgn` tokens (`"ctr"` / `"l"` / `"r"`).
 *
 * Returns `undefined` when:
 *   - the input is missing,
 *   - the value is not in the `ST_LblAlgn` enumeration,
 *   - the value is `"ctr"` (the OOXML default — Excel's reference
 *     centered alignment — and what absence already means).
 *
 * Unknown tokens drop rather than fall back to the default so a
 * malformed override surfaces as "no alignment emitted" instead of
 * silently snapping to `"ctr"` (which would mask the configuration
 * error in the caller).
 */
export function normalizeAxisLblAlgn(
  value: ChartAxisLabelAlign | undefined,
): ChartAxisLabelAlign | undefined {
  if (value === undefined) return undefined
  if (value !== "ctr" && value !== "l" && value !== "r") return undefined
  if (value === "ctr") return undefined
  return value
}

/**
 * Normalize an axis `hidden` flag to a strict boolean. Anything other
 * than literal `true` collapses to `false` so the writer never emits
 * `<c:delete val="1"/>` from a stray non-boolean leaking through the
 * type guard (e.g. `0` / `1` / `"true"` / `null`). This matches how
 * `roundedCorners` / `plotVisOnly` / `varyColors` treat their inputs:
 * a literal boolean is the only path to a non-default value.
 */
export function normalizeAxisHidden(value: boolean | undefined): boolean {
  return value === true
}

/**
 * Normalize an axis `labelRotation` value (whole degrees) for the
 * `<c:txPr><a:bodyPr rot="N"/></c:txPr>` writer slot. Returns
 * `undefined` when the input is unset, non-finite, non-numeric, or
 * resolves to `0` after rounding — every absence path collapses to the
 * same omit-the-element shape so absence and the OOXML default `0`
 * round-trip identically through {@link cloneChart}. Out-of-range
 * inputs clamp to the `-90..90` band Excel's UI exposes; non-integer
 * inputs round to the nearest whole degree (the OOXML attribute is an
 * integer in 60000ths of a degree, so a fractional whole-degree value
 * has no meaningful refinement at emit time).
 */
export function normalizeAxisLabelRotation(value: number | undefined): number | undefined {
  if (value === undefined || typeof value !== "number" || !Number.isFinite(value)) return undefined
  let degrees = Math.round(value)
  if (degrees < LABEL_ROTATION_MIN_DEG) degrees = LABEL_ROTATION_MIN_DEG
  else if (degrees > LABEL_ROTATION_MAX_DEG) degrees = LABEL_ROTATION_MAX_DEG
  if (degrees === 0) return undefined
  return degrees
}

/**
 * Normalize an axis `labelFontSize` value (whole / half points) for
 * the `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
 * writer slot. Returns `undefined` when the input is unset,
 * non-finite, non-numeric, or out of the `1..400`pt band the OOXML
 * `ST_TextFontSize` schema exposes — every absence path collapses to
 * the same omit-the-attribute shape so a fresh chart inherits Excel's
 * reference 10pt tick-label size.
 *
 * Fractional inputs round to the nearest 0.5pt (the OOXML attribute
 * is an integer in 100ths of a point and Excel's UI exposes the same
 * 0.5pt granularity, so finer fractions have no meaningful refinement
 * at emit time). Mirrors {@link normalizeTitleFontSize} so a value
 * threads cleanly through both the title and axis-label slots.
 */
export function normalizeAxisLabelFontSize(value: number | undefined): number | undefined {
  if (value === undefined || typeof value !== "number" || !Number.isFinite(value)) return undefined
  const halfSteps = Math.round(value * 2)
  const points = halfSteps / 2
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined
  return points
}

/**
 * Normalize an axis `labelBold` value for the
 * `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>`
 * writer slot. Delegates to the chart-level {@link normalizeTitleBold}
 * — `true` / `false` pass through literally, every other token (typed
 * escape from an untyped caller, including `null`-shaped values)
 * collapses to `undefined` so the writer omits the `b` attribute
 * entirely. Absence then collapses to the OOXML default Excel itself
 * emits on a fresh axis (the theme-default tick-label weight).
 */
export function normalizeAxisLabelBold(value: boolean | undefined): boolean | undefined {
  return normalizeTitleBold(value)
}

/**
 * Normalize an axis `labelItalic` value for the
 * `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>`
 * writer slot. Delegates to the chart-level {@link normalizeTitleItalic}
 * — `true` / `false` pass through literally, every other token (typed
 * escape from an untyped caller, including `null`-shaped values)
 * collapses to `undefined` so the writer omits the `i` attribute
 * entirely. Absence then collapses to the OOXML default Excel itself
 * emits on a fresh axis (the theme-default tick-label slant).
 */
export function normalizeAxisLabelItalic(value: boolean | undefined): boolean | undefined {
  return normalizeTitleItalic(value)
}

/**
 * Normalize an axis `labelColor` value for the
 * `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
 * </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` writer slot.
 * Delegates to the chart-level {@link normalizeTitleColor} so the two
 * share the same accept-with-or-without-`#` grammar — `"FF0000"`,
 * `"#FF0000"`, and `"ff0000"` all collapse to the OOXML uppercase
 * canonical form; malformed inputs (wrong length, non-hex characters,
 * alpha-channel forms, non-string escapes from an untyped caller)
 * collapse to `undefined` so the writer skips the entire
 * `<a:solidFill>` block and the tick labels inherit the theme text
 * color (Excel's reference behavior for fresh tick labels without a
 * custom color).
 */
export function normalizeAxisLabelColor(value: ChartColor | undefined): ChartColor | undefined {
  return normalizeTitleColor(value)
}

/**
 * Normalize an axis `labelUnderline` value for the
 * `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>`
 * writer slot. Delegates to the chart-level
 * {@link normalizeTitleUnderline} — `true` / `false` pass through
 * literally, every other token (typed escape from an untyped caller,
 * including `null`-shaped values) collapses to `undefined` so the
 * writer omits the `u` attribute entirely. Absence then collapses to
 * the OOXML default Excel itself emits on a fresh axis (the
 * theme-default non-underlined tick labels).
 */
export function normalizeAxisLabelUnderline(value: boolean | undefined): boolean | undefined {
  return normalizeTitleUnderline(value)
}

/**
 * Normalize an axis `labelStrike` value for the
 * `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>`
 * writer slot. Delegates to the chart-level
 * {@link normalizeTitleStrike} — `true` / `false` pass through
 * literally, every other token (typed escape from an untyped caller,
 * including `null`-shaped values) collapses to `undefined` so the
 * writer omits the `strike` attribute entirely. Absence then collapses
 * to the OOXML default Excel itself emits on a fresh axis (the theme-
 * default non-strikethrough tick labels).
 */
export function normalizeAxisLabelStrike(value: boolean | undefined): boolean | undefined {
  return normalizeTitleStrike(value)
}

/**
 * Normalize an axis `labelFontFamily` value for the
 * `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
 * </a:pPr></a:p></c:txPr>` writer slot. Returns the trimmed typeface
 * string when the input is a non-empty string, or `undefined` for any
 * malformed token — empty / whitespace-only strings, or non-string
 * escapes from an untyped caller (`null`, numbers, booleans, etc.).
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the entire `<a:latin>` element and the tick labels
 * inherit the theme typeface (Excel's reference behavior for a fresh
 * axis without a custom tick-label font picked).
 */
export function normalizeAxisLabelFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

/**
 * Build the `<c:txPr>` block that carries an axis tick-label rotation,
 * font size, bold flag, italic flag, font color, underline flag, and /
 * or strikethrough flag. Returns `undefined` when every input is unset
 * so the caller can elide the element entirely (Excel's reference
 * serialization on a fresh axis omits `<c:txPr>` when the labels render
 * at the default rotation and inherit the theme font / weight / slant /
 * color / underline / strike).
 *
 * The emitted block mirrors the minimal `<c:txPr>` shape Excel writes
 * when the user pins a custom typography knob — `<a:bodyPr rot="N"/>`
 * carries the rotation (when set), `<a:lstStyle/>` is the empty
 * list-style placeholder the schema requires, and the
 * `<a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p>` paragraph
 * stub Excel always emits hosts the optional `sz` / `b` / `i` / `u` /
 * `strike` attributes on `<a:defRPr>`. Additional `<a:bodyPr>` attributes Excel writes in
 * its full reference (`spcFirstLastPara` / `vertOverflow` / `wrap` /
 * `anchor` / `anchorCtr`) are intentionally omitted — the OOXML
 * schema marks them all optional, and dropping them keeps the
 * writer's footprint minimal while preserving the typography intent.
 *
 * When only a rotation is pinned, the `<a:defRPr>` slot self-closes
 * with no attributes (matching the legacy rotation-only emit). When a
 * font size, bold flag, italic flag, font color, underline flag, or
 * strikethrough flag is pinned, the slot carries `sz="N"` (in 100ths
 * of a point), `b="1"` / `b="0"`, `i="1"` / `i="0"`, `u="sng"`,
 * `strike="sngStrike"`, and / or wraps an `<a:solidFill>` child so a
 * re-parse picks the values up off the canonical default-paragraph
 * slot. When the rotation is absent but any other knob is pinned, the
 * writer omits `rot` from `<a:bodyPr>` so the OOXML default `0`
 * collapses to absence. The bold / italic flags each emit a literal
 * `1` / `0` whenever the input is a boolean — `false` pins the OOXML
 * default explicitly, which is functionally identical to absence but
 * lets a clone target override an upstream `1` from a templated chart.
 * The underline flag emits only `u="sng"` (Excel's UI variant — single
 * line) when the input is `true`; absence and explicit `false` both
 * collapse to omitting the attribute since the OOXML default `"none"`
 * collapses to absence. The strikethrough flag emits only
 * `strike="sngStrike"` (Excel's UI variant — single line) when the
 * input is `true`; absence and explicit `false` both collapse to
 * omitting the attribute since the OOXML default `"noStrike"`
 * collapses to absence.
 *
 * The `rgbHex` parameter pins the tick-label color via the
 * `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
 * </a:defRPr>` slot. When set, the `<a:defRPr>` element expands from
 * self-closing to wrapping a single `<a:solidFill>` child; otherwise
 * the writer keeps the existing self-closing form so a fresh axis
 * with no custom color matches Excel's reference serialization (the
 * tick labels inherit the theme text color in that case).
 */
export function buildAxisTxPr(
  rotationDeg: number | undefined,
  fontSizePt: number | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
  rgbHex: ChartColor | undefined,
  underline: boolean | undefined,
  strike: boolean | undefined,
  fontFamily: string | undefined,
): string | undefined {
  if (
    rotationDeg === undefined &&
    fontSizePt === undefined &&
    bold === undefined &&
    italic === undefined &&
    rgbHex === undefined &&
    underline === undefined &&
    strike === undefined &&
    fontFamily === undefined
  )
    return undefined
  const rot = rotationDeg === undefined ? undefined : rotationDeg * TXPR_ROT_PER_DEGREE
  const sz = fontSizePt === undefined ? undefined : fontSizePt * TITLE_FONT_SZ_PER_POINT
  const b = bold === undefined ? undefined : bold ? 1 : 0
  const i = italic === undefined ? undefined : italic ? 1 : 0
  // OOXML's `<a:defRPr u=".."/>` attribute is the
  // `ST_TextUnderlineType` enum on `CT_TextCharacterProperties` —
  // eighteen values total, with `"none"` as the OOXML default and
  // `"sng"` as the value Excel's UI authors for the "Underline"
  // checkbox (single line). The writer emits only the UI variant
  // `"sng"` to keep the surfaced shape consistent with what Excel's
  // reference UI authors. Absence (`undefined`) and explicit `false`
  // both collapse to omitting the attribute (the OOXML default
  // `"none"` collapses to absence; Excel itself omits `u` when the
  // tick labels are not underlined).
  const u = underline === true ? "sng" : undefined
  // OOXML's `<a:defRPr strike=".."/>` attribute is the
  // `ST_TextStrikeType` enum on `CT_TextCharacterProperties` — three
  // values total, with `"noStrike"` as the OOXML default and
  // `"sngStrike"` as the value Excel's UI authors for the
  // "Strikethrough" checkbox (single line). The writer emits only the
  // UI variant `"sngStrike"` to keep the surfaced shape consistent
  // with what Excel's reference UI authors. Absence (`undefined`) and
  // explicit `false` both collapse to omitting the attribute (the
  // OOXML default `"noStrike"` collapses to absence; Excel itself
  // omits `strike` when the tick labels are not strikethrough).
  const strikeAttr = strike === true ? "sngStrike" : undefined
  // OOXML's `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:defRPr>` carries the tick-label font color.
  // Absence (`undefined`) collapses to omitting the entire
  // `<a:solidFill>` block so the labels inherit the theme text color
  // (Excel's reference behavior for a fresh axis that has not had a
  // custom tick-label color picked).
  const solidFillChild =
    rgbHex !== undefined
      ? xmlElement("a:solidFill", undefined, [buildColorElement(rgbHex)])
      : undefined
  // OOXML's `<a:defRPr><a:latin typeface=".."/></a:defRPr>` carries
  // the tick-label font family. The `<a:latin>` element follows
  // `<a:solidFill>` per the CT_TextCharacterProperties child sequence
  // (ECMA-376 Part 1, §21.1.2.3.7). Absence (`undefined`) collapses to
  // omitting the entire `<a:latin>` element so the labels inherit the
  // theme typeface (Excel's reference behavior for a fresh axis that
  // has not had a custom tick-label font picked).
  const latinChild = fontFamily ? xmlSelfClose("a:latin", { typeface: fontFamily }) : undefined
  // When a fill color or a typeface is set the `<a:defRPr>` slot
  // expands from self-closing to wrapping the children; otherwise the
  // writer keeps the existing self-closing form so a fresh axis with
  // no custom color or font matches Excel's reference serialization
  // byte-for-byte. Children are emitted in
  // CT_TextCharacterProperties' canonical schema order: solidFill
  // first, then latin.
  const defRPrChildren: string[] = []
  if (solidFillChild) defRPrChildren.push(solidFillChild)
  if (latinChild) defRPrChildren.push(latinChild)
  const defRPr =
    defRPrChildren.length > 0
      ? xmlElement("a:defRPr", { sz, b, i, u, strike: strikeAttr }, defRPrChildren)
      : xmlSelfClose("a:defRPr", { sz, b, i, u, strike: strikeAttr })
  return xmlElement("c:txPr", undefined, [
    xmlSelfClose("a:bodyPr", { rot }),
    xmlSelfClose("a:lstStyle"),
    xmlElement("a:p", undefined, [
      xmlElement("a:pPr", undefined, [defRPr]),
      xmlSelfClose("a:endParaRPr", { lang: "en-US" }),
    ]),
  ])
}

/**
 * Resolve the writer's `axes.x.crosses` / `axes.x.crossesAt` pair into
 * the {@link ResolvedAxisCrosses} tagged union the per-family axis
 * builders emit. The OOXML schema places `<c:crosses>` and
 * `<c:crossesAt>` in an XSD choice — only one may legally appear at a
 * time per ECMA-376 Part 1, §21.2.2 — so the normalizer collapses the
 * caller's two fields to a single resolved shape:
 *
 *   - A finite numeric `crossesAt` always wins, mirroring how Excel
 *     treats the choice (the explicit numeric pin overrides the
 *     semantic default). Non-finite inputs (NaN / Infinity) drop so the
 *     writer never emits an attribute Excel would reject.
 *   - When only `crosses` is set, the resolved kind is `"semantic"` for
 *     `"min"` / `"max"`. The OOXML default `"autoZero"` collapses to
 *     `kind: "default"` so absence and the default emit the same
 *     `<c:crosses val="autoZero"/>` byte-for-byte. Unknown tokens drop
 *     to `kind: "default"` for the same reason.
 *   - When neither is set, the resolved kind is `"default"` (the writer
 *     still emits `<c:crosses val="autoZero"/>` to match Excel's
 *     reference serialization on every freshly-drawn axis).
 */
export function normalizeAxisCrosses(
  semantic: ChartAxisCrosses | undefined,
  numeric: number | undefined,
): ResolvedAxisCrosses {
  if (typeof numeric === "number" && Number.isFinite(numeric)) {
    return { kind: "numeric", value: numeric }
  }
  if (semantic !== undefined && VALID_AXIS_CROSSES.has(semantic) && semantic !== "autoZero") {
    return { kind: "semantic", value: semantic }
  }
  return { kind: "default" }
}

/**
 * Render the resolved axis crossing pin as the matching child element.
 * `kind: "numeric"` emits `<c:crossesAt val=".."/>`; every other kind
 * emits `<c:crosses val=".."/>` so Excel's reference serialization
 * (which always pins `<c:crosses val="autoZero"/>` on every axis) is
 * preserved on freshly-drawn charts.
 */
export function buildAxisCrosses(resolved: ResolvedAxisCrosses): string {
  switch (resolved.kind) {
    case "numeric":
      return xmlSelfClose("c:crossesAt", { val: resolved.value })
    case "semantic":
      return xmlSelfClose("c:crosses", { val: resolved.value })
    case "default":
      return xmlSelfClose("c:crosses", { val: "autoZero" })
  }
}

/**
 * Build the children that augment a `<c:scaling>` element. Order is
 * spec-enforced: `<c:logBase>` → `<c:orientation>` → `<c:max>` →
 * `<c:min>`. The orientation child is always emitted by the caller
 * (every axis declares `minMax`); this helper handles the rest.
 *
 * Returns the children to splice in after `<c:orientation>`.
 */
export function buildAxisScalingExtras(scale: ChartAxisScale | undefined): {
  before: string[]
  after: string[]
} {
  if (!scale) return { before: [], after: [] }
  const before: string[] = []
  const after: string[] = []
  // logBase comes before orientation per CT_Scaling.
  if (scale.logBase !== undefined) {
    before.push(xmlSelfClose("c:logBase", { val: scale.logBase }))
  }
  // max and min come after orientation, with max first (CT_Scaling).
  if (scale.max !== undefined) after.push(xmlSelfClose("c:max", { val: scale.max }))
  if (scale.min !== undefined) after.push(xmlSelfClose("c:min", { val: scale.min }))
  return { before, after }
}

/**
 * Build the `<c:scaling>` element. Always emits `<c:orientation>` so
 * the axis renders correctly even when no extra scale fields are set —
 * `"minMax"` (the OOXML default) for a forward axis, `"maxMin"` when
 * the caller pinned `reverse: true` to flip the plotting order.
 */
export function buildAxisScaling(
  scale: ChartAxisScale | undefined,
  reverse: boolean = false,
): string {
  const { before, after } = buildAxisScalingExtras(scale)
  const children: string[] = [
    ...before,
    xmlSelfClose("c:orientation", { val: reverse ? "maxMin" : "minMax" }),
    ...after,
  ]
  return xmlElement("c:scaling", undefined, children)
}

/**
 * Build the optional `<c:majorUnit>` / `<c:minorUnit>` siblings that
 * sit later in the axis-element child sequence (after `<c:numFmt>`,
 * before `<c:crossAx>` per CT_CatAx / CT_ValAx).
 */
export function buildAxisTickUnits(scale: ChartAxisScale | undefined): string[] {
  if (!scale) return []
  const out: string[] = []
  if (scale.majorUnit !== undefined) {
    out.push(xmlSelfClose("c:majorUnit", { val: scale.majorUnit }))
  }
  if (scale.minorUnit !== undefined) {
    out.push(xmlSelfClose("c:minorUnit", { val: scale.minorUnit }))
  }
  return out
}

/**
 * Normalize the {@link SheetChart.axes.x.dispUnits} /
 * {@link SheetChart.axes.y.dispUnits} input — accept either the
 * `ChartAxisDispUnit` shorthand (e.g. `"millions"`) or the full
 * {@link ChartAxisDispUnits} object — into a single canonical shape the
 * writer can hand off to {@link buildAxisDispUnits}. Unknown / typo'd
 * tokens collapse to `undefined` so the writer never emits a value the
 * OOXML `ST_BuiltInUnit` enum rejects. Non-object / non-string inputs
 * (e.g. `null`, numbers, arrays) also collapse to `undefined`.
 *
 * The OOXML schema places `<c:builtInUnit>` and `<c:custUnit>` in an
 * `xsd:choice` — at most one of the two may appear inside `<c:dispUnits>`.
 * The normalizer keeps both fields when the input declares both, leaving
 * the precedence pick to {@link buildAxisDispUnits} (where `custUnit`
 * wins because it is the more specific OOXML element). When neither
 * field is valid, the normalizer returns `undefined` so the writer skips
 * the `<c:dispUnits>` element entirely — a `<c:dispUnits>` shell with
 * no `<c:builtInUnit>` / `<c:custUnit>` child would fail Excel's strict
 * validator.
 */
export function normalizeAxisDispUnits(
  value: ChartAxisDispUnits | ChartAxisDispUnit | undefined,
): ChartAxisDispUnits | undefined {
  if (value === undefined) return undefined
  if (typeof value === "string") {
    return VALID_DISP_UNITS.has(value as ChartAxisDispUnit)
      ? { unit: value as ChartAxisDispUnit }
      : undefined
  }
  if (typeof value !== "object" || value === null) return undefined
  const out: ChartAxisDispUnits = {}
  const unit = value.unit
  if (typeof unit === "string" && VALID_DISP_UNITS.has(unit as ChartAxisDispUnit)) {
    out.unit = unit as ChartAxisDispUnit
  }
  const custUnit = value.custUnit
  if (typeof custUnit === "number" && Number.isFinite(custUnit) && custUnit > 0) {
    out.custUnit = custUnit
  }
  // Drop the entire object when neither child resolves — a bare
  // `<c:dispUnits/>` shell would fail Excel's strict validator (the
  // CT_DispUnits choice has `minOccurs="0"` on the choice itself, but
  // an empty element with the parent's `<c:extLst>` slot also empty
  // is rejected by Excel's reference renderer).
  if (out.unit === undefined && out.custUnit === undefined) return undefined
  if (value.showLabel === true) out.showLabel = true
  if (typeof value.customLabel === "string") {
    const trimmed = value.customLabel.trim()
    if (trimmed.length > 0) out.customLabel = trimmed
  }
  return out
}

/**
 * Build the optional `<c:dispUnits>` block that sits at the very end of
 * `<c:valAx>` per CT_ValAx (after `<c:minorUnit>`). The element itself
 * holds the choice between `<c:builtInUnit>` and `<c:custUnit>` — the
 * writer emits exactly one per the OOXML `xsd:choice`, preferring
 * `<c:custUnit>` when both fields are pinned because the more specific
 * numeric divisor takes precedence (a caller appending a custom unit to
 * a cloned source need not manually prune the inherited preset). When
 * `showLabel` is `true` the writer emits a bare `<c:dispUnitsLbl/>` so
 * Excel paints its default automatic annotation; the rich-text label
 * customization is intentionally not surfaced.
 *
 * Returns an empty array when the caller did not pin either child so
 * the writer leaves Excel's default "no display unit" state untouched.
 */
export function buildAxisDispUnits(dispUnits: ChartAxisDispUnits | undefined): string[] {
  if (!dispUnits) return []
  const children: string[] = []
  if (dispUnits.custUnit !== undefined) {
    children.push(xmlSelfClose("c:custUnit", { val: dispUnits.custUnit }))
  } else if (dispUnits.unit !== undefined) {
    children.push(xmlSelfClose("c:builtInUnit", { val: dispUnits.unit }))
  } else {
    // Neither child resolved — skip emission rather than ship a bare
    // `<c:dispUnits/>` Excel rejects. The normalizer should have
    // pre-filtered this case, but the guard here keeps the writer
    // robust against a stray runtime object slipping past the type
    // boundary.
    return []
  }
  if (
    dispUnits.showLabel === true ||
    (typeof dispUnits.customLabel === "string" && dispUnits.customLabel.trim().length > 0)
  ) {
    const customLabel =
      typeof dispUnits.customLabel === "string" ? dispUnits.customLabel.trim() : ""
    if (customLabel.length > 0) {
      // Build `<c:dispUnitsLbl><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      // <a:p><a:r><a:t>...</a:t></a:r></a:p></c:rich></c:tx></c:dispUnitsLbl>`.
      // Excel's reference serialization for a custom display-unit label
      // emits a bare `<a:bodyPr/>` and `<a:lstStyle/>` placeholder
      // before the paragraph — mirror that minimal shape so a re-parse
      // walks the canonical path.
      const escaped = customLabel.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      const richBlock = xmlElement("c:rich", undefined, [
        xmlSelfClose("a:bodyPr"),
        xmlSelfClose("a:lstStyle"),
        xmlElement("a:p", undefined, [
          xmlElement("a:r", undefined, [xmlElement("a:t", undefined, escaped)]),
        ]),
      ])
      const txBlock = xmlElement("c:tx", undefined, [richBlock])
      children.push(xmlElement("c:dispUnitsLbl", undefined, [txBlock]))
    } else {
      children.push(xmlSelfClose("c:dispUnitsLbl"))
    }
  }
  return [xmlElement("c:dispUnits", undefined, children)]
}

/**
 * Normalize the {@link SheetChart.axes.x.crossBetween} /
 * {@link SheetChart.axes.y.crossBetween} input. Unknown / typo'd tokens
 * collapse to `undefined` so the writer never emits a value the OOXML
 * `ST_CrossBetween` enum rejects — the caller's per-family default
 * (`"between"` on bar / column / line / area Y axes; `"midCat"` on
 * scatter axes) takes over instead.
 *
 * Non-string inputs (e.g. `null`, numbers, arrays) likewise collapse to
 * `undefined` so a stray runtime value leaking through the type guard
 * cannot poison the output.
 */
export function normalizeAxisCrossBetween(
  value: ChartAxisCrossBetween | undefined,
): ChartAxisCrossBetween | undefined {
  if (typeof value !== "string") return undefined
  return VALID_CROSS_BETWEEN.has(value as ChartAxisCrossBetween)
    ? (value as ChartAxisCrossBetween)
    : undefined
}

/**
 * Build the axis tick-label `<c:numFmt formatCode=".." sourceLinked=".."/>`.
 * Returns an empty array when the axis declares no number format — the
 * writer then leaves Excel's default linked behaviour untouched.
 */
export function buildAxisNumFmt(numFmt: ChartAxisNumberFormat | undefined): string[] {
  if (!numFmt) return []
  const sourceLinked = numFmt.sourceLinked === true ? 1 : 0
  return [xmlSelfClose("c:numFmt", { formatCode: numFmt.formatCode, sourceLinked })]
}

/**
 * Normalize a tick-mark value to a token Excel accepts. Unknown / typo'd
 * inputs collapse to `undefined` so the writer never emits a value the
 * OOXML `ST_TickMark` enum rejects.
 */
export function normalizeTickMark(
  value: ChartAxisTickMark | undefined,
): ChartAxisTickMark | undefined {
  if (value === undefined) return undefined
  return TICK_MARK_VALUES.has(value) ? value : undefined
}

/**
 * Normalize a tick-label-position value to a token Excel accepts.
 * Unknown / typo'd inputs collapse to `undefined` so the writer never
 * emits a value the OOXML `ST_TickLblPos` enum rejects.
 */
export function normalizeTickLblPos(
  value: ChartAxisTickLabelPosition | undefined,
): ChartAxisTickLabelPosition | undefined {
  if (value === undefined) return undefined
  return TICK_LBL_POS_VALUES.has(value) ? value : undefined
}

/**
 * Build the `<c:majorTickMark>` / `<c:minorTickMark>` / `<c:tickLblPos>`
 * children for an axis. The OOXML schema (CT_CatAx / CT_ValAx /
 * CT_DateAx / CT_SerAx) places the three elements together right after
 * `<c:numFmt>` and before `<c:crossAx>`. Excel's strict validator
 * rejects any other ordering — keep the tuple together.
 *
 * Each value is omitted when the caller did not pin it; the OOXML
 * defaults (`majorTickMark="out"`, `minorTickMark="none"`,
 * `tickLblPos="nextTo"`) match Excel's reference serialization, so
 * absence and the default round-trip identically through the reader.
 */
export function buildAxisTickRendering(
  majorTickMark: ChartAxisTickMark | undefined,
  minorTickMark: ChartAxisTickMark | undefined,
  tickLblPos: ChartAxisTickLabelPosition | undefined,
): string[] {
  const out: string[] = []
  if (majorTickMark !== undefined) {
    out.push(xmlSelfClose("c:majorTickMark", { val: majorTickMark }))
  }
  if (minorTickMark !== undefined) {
    out.push(xmlSelfClose("c:minorTickMark", { val: minorTickMark }))
  }
  if (tickLblPos !== undefined) {
    out.push(xmlSelfClose("c:tickLblPos", { val: tickLblPos }))
  }
  return out
}

/**
 * Build the `<c:tickLblSkip>` / `<c:tickMarkSkip>` siblings that sit
 * between `<c:lblOffset>` and `<c:noMultiLvlLbl>` inside `<c:catAx>`
 * (CT_CatAx). Order is `tickLblSkip` first, then `tickMarkSkip` per
 * the OOXML schema. Each element is emitted only when the caller
 * pinned a non-default value (the helper relies on
 * {@link normalizeAxisSkip} having already collapsed `1` and out-of-
 * range inputs to `undefined`).
 */
export function buildAxisSkips(
  tickLblSkip: number | undefined,
  tickMarkSkip: number | undefined,
): string[] {
  const out: string[] = []
  if (tickLblSkip !== undefined) {
    out.push(xmlSelfClose("c:tickLblSkip", { val: tickLblSkip }))
  }
  if (tickMarkSkip !== undefined) {
    out.push(xmlSelfClose("c:tickMarkSkip", { val: tickMarkSkip }))
  }
  return out
}

export function buildBarAxes(orientation: "bar" | "column", opts: AxisRenderOptions): string[] {
  // For a vertical column chart, categories sit on the bottom (catAx)
  // and values run vertically (valAx). For a horizontal bar chart the
  // axes swap orientation.
  const catPos = orientation === "column" ? "b" : "l"
  const valPos = orientation === "column" ? "l" : "b"

  // OOXML enforces a strict child order inside <c:catAx>/<c:valAx>:
  // axId → scaling → delete → axPos → majorGridlines → minorGridlines
  // → title → numFmt → majorTickMark → minorTickMark → tickLblPos →
  // crossAx → crosses → ... → majorUnit → minorUnit. Each block below
  // mirrors that order.
  // The category axis on bar/column rarely uses scaling, but Excel
  // tolerates the augmentation either way; surface it whenever the
  // caller pinned a value so write-side templates round-trip.
  const catAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_CAT }),
    buildAxisScaling(opts.xScale, opts.xReverse),
    xmlSelfClose("c:delete", { val: opts.xHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: catPos }),
    ...buildAxisGridlines(opts.xGridlines),
  ]
  if (opts.xAxisTitle)
    catAxChildren.push(
      buildAxisTitle(
        opts.xAxisTitle,
        opts.xAxisTitleRotation,
        opts.xAxisTitleFontSize,
        opts.xAxisTitleBold,
        opts.xAxisTitleItalic,
        opts.xAxisTitleColor,
        opts.xAxisTitleStrike,
        opts.xAxisTitleUnderline,
        opts.xAxisTitleFontFamily,
        opts.xAxisTitleOverlay,
        opts.xAxisTitleLayout,
        opts.xAxisTitleFillColor,
        opts.xAxisTitleBorderColor,
        opts.xAxisTitleBorderWidth,
        opts.xAxisTitleBorderDash,
        opts.xAxisTitleBorderCap,
        opts.xAxisTitleBorderCompound,
      ),
    )
  catAxChildren.push(
    ...buildAxisNumFmt(opts.xNumFmt),
    ...buildAxisTickRendering(opts.xMajorTickMark, opts.xMinorTickMark, opts.xTickLblPos),
  )
  // `<c:txPr>` sits between `<c:tickLblPos>` (the last child of
  // `buildAxisTickRendering`) and `<c:crossAx>` per CT_CatAx (ECMA-376
  // Part 1, §21.2.2.7). Skip the entire block when the caller did not
  // pin a rotation so a fresh chart matches Excel's minimal serialization.
  const xCatAxTxPr = buildAxisTxPr(
    opts.xLabelRotation,
    opts.xLabelFontSize,
    opts.xLabelBold,
    opts.xLabelItalic,
    opts.xLabelColor,
    opts.xLabelUnderline,
    opts.xLabelStrike,
    opts.xLabelFontFamily,
  )
  if (xCatAxTxPr) catAxChildren.push(xCatAxTxPr)
  catAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL }),
    buildAxisCrosses(opts.xCrosses),
    // `<c:auto>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when `false`; absence (and any non-
    // boolean) collapses to the OOXML default `true` so untouched
    // charts match Excel's output byte-for-byte.
    xmlSelfClose("c:auto", { val: opts.xAuto ? 1 : 0 }),
    // `<c:lblAlgn>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when set; absence (or the OOXML
    // default `"ctr"` collapsed by `normalizeAxisLblAlgn`) emits the
    // default so untouched charts match Excel's output byte-for-byte.
    xmlSelfClose("c:lblAlgn", { val: opts.xLblAlgn ?? "ctr" }),
    // `<c:lblOffset>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when set; absence (or the OOXML
    // default `100` collapsed by `normalizeAxisLblOffset`) emits the
    // default so untouched charts match Excel's output byte-for-byte.
    xmlSelfClose("c:lblOffset", { val: opts.xLblOffset ?? 100 }),
    // OOXML CT_CatAx places `<c:tickLblSkip>` / `<c:tickMarkSkip>`
    // after `<c:lblOffset>` and before `<c:noMultiLvlLbl>`. Only
    // emit each element when the caller pinned a non-default value
    // so a fresh chart matches Excel's reference serialization (the
    // default `1` is omitted and Excel renders every tick).
    ...buildAxisSkips(opts.xTickLblSkip, opts.xTickMarkSkip),
    // `<c:noMultiLvlLbl>` is always emitted because Excel's reference
    // serialization includes it on every category axis. The writer
    // pins the caller's override when `true`; absence and an explicit
    // `false` both produce `val="0"` so untouched charts match Excel's
    // output byte-for-byte.
    xmlSelfClose("c:noMultiLvlLbl", { val: opts.xNoMultiLvlLbl ? 1 : 0 }),
  )

  const valAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL }),
    buildAxisScaling(opts.yScale, opts.yReverse),
    xmlSelfClose("c:delete", { val: opts.yHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: valPos }),
    ...buildAxisGridlines(opts.yGridlines),
  ]
  if (opts.yAxisTitle)
    valAxChildren.push(
      buildAxisTitle(
        opts.yAxisTitle,
        opts.yAxisTitleRotation,
        opts.yAxisTitleFontSize,
        opts.yAxisTitleBold,
        opts.yAxisTitleItalic,
        opts.yAxisTitleColor,
        opts.yAxisTitleStrike,
        opts.yAxisTitleUnderline,
        opts.yAxisTitleFontFamily,
        opts.yAxisTitleOverlay,
        opts.yAxisTitleLayout,
        opts.yAxisTitleFillColor,
        opts.yAxisTitleBorderColor,
        opts.yAxisTitleBorderWidth,
        opts.yAxisTitleBorderDash,
        opts.yAxisTitleBorderCap,
        opts.yAxisTitleBorderCompound,
      ),
    )
  valAxChildren.push(
    ...buildAxisNumFmt(opts.yNumFmt),
    ...buildAxisTickRendering(opts.yMajorTickMark, opts.yMinorTickMark, opts.yTickLblPos),
  )
  // `<c:txPr>` sits between `<c:tickLblPos>` and `<c:crossAx>` per
  // CT_ValAx (ECMA-376 Part 1, §21.2.2.32). Same omit-by-default
  // contract as the catAx slot above — emit nothing when the caller
  // did not pin a rotation so the writer matches Excel's reference
  // serialization on a fresh value axis.
  const yValAxTxPr = buildAxisTxPr(
    opts.yLabelRotation,
    opts.yLabelFontSize,
    opts.yLabelBold,
    opts.yLabelItalic,
    opts.yLabelColor,
    opts.yLabelUnderline,
    opts.yLabelStrike,
    opts.yLabelFontFamily,
  )
  if (yValAxTxPr) valAxChildren.push(yValAxTxPr)
  valAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_CAT }),
    buildAxisCrosses(opts.yCrosses),
    // `<c:crossBetween>` sits between `<c:crosses>`/`<c:crossesAt>` and
    // `<c:majorUnit>` per CT_ValAx (ECMA-376 §21.2.2.32). The default
    // for bar / column / line / area is `"between"` — the writer
    // honours an override when the caller pinned `"midCat"` and falls
    // back to the family default otherwise.
    xmlSelfClose("c:crossBetween", { val: opts.yCrossBetween ?? "between" }),
    ...buildAxisTickUnits(opts.yScale),
    // `<c:dispUnits>` is the last child slot on `<c:valAx>` per
    // CT_ValAx (after `<c:minorUnit>`). Bar / column / line / area
    // charts route the X axis through `<c:catAx>` (which rejects the
    // element), so only the Y axis picks up the writer-side input.
    ...buildAxisDispUnits(opts.yDispUnits),
  )

  return [
    xmlElement("c:catAx", undefined, catAxChildren),
    xmlElement("c:valAx", undefined, valAxChildren),
  ]
}

export function buildScatterAxes(opts: AxisRenderOptions): string[] {
  const xAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_X }),
    buildAxisScaling(opts.xScale, opts.xReverse),
    xmlSelfClose("c:delete", { val: opts.xHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: "b" }),
    ...buildAxisGridlines(opts.xGridlines),
  ]
  if (opts.xAxisTitle)
    xAxChildren.push(
      buildAxisTitle(
        opts.xAxisTitle,
        opts.xAxisTitleRotation,
        opts.xAxisTitleFontSize,
        opts.xAxisTitleBold,
        opts.xAxisTitleItalic,
        opts.xAxisTitleColor,
        opts.xAxisTitleStrike,
        opts.xAxisTitleUnderline,
        opts.xAxisTitleFontFamily,
        opts.xAxisTitleOverlay,
        opts.xAxisTitleLayout,
        opts.xAxisTitleFillColor,
        opts.xAxisTitleBorderColor,
        opts.xAxisTitleBorderWidth,
        opts.xAxisTitleBorderDash,
        opts.xAxisTitleBorderCap,
        opts.xAxisTitleBorderCompound,
      ),
    )
  xAxChildren.push(
    ...buildAxisNumFmt(opts.xNumFmt),
    ...buildAxisTickRendering(opts.xMajorTickMark, opts.xMinorTickMark, opts.xTickLblPos),
  )
  // `<c:txPr>` slot — same CT_ValAx position as the bar / column
  // builder above. Scatter X is a value axis, so the rotation pins on
  // the X-axis just as it does on the Y-axis.
  const xValAxTxPr = buildAxisTxPr(
    opts.xLabelRotation,
    opts.xLabelFontSize,
    opts.xLabelBold,
    opts.xLabelItalic,
    opts.xLabelColor,
    opts.xLabelUnderline,
    opts.xLabelStrike,
    opts.xLabelFontFamily,
  )
  if (xValAxTxPr) xAxChildren.push(xValAxTxPr)
  xAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_Y }),
    buildAxisCrosses(opts.xCrosses),
    // Scatter charts default to `"midCat"` (data points sit ON the
    // perpendicular-axis ticks rather than between them). The writer
    // honours an override when the caller pinned `"between"` and falls
    // back to the family default otherwise.
    xmlSelfClose("c:crossBetween", { val: opts.xCrossBetween ?? "midCat" }),
    ...buildAxisTickUnits(opts.xScale),
    // `<c:dispUnits>` slots onto `<c:valAx>` per CT_ValAx (after
    // `<c:minorUnit>`). Scatter charts route both axes through
    // `<c:valAx>`, so the X-axis builder picks up `xDispUnits` here.
    ...buildAxisDispUnits(opts.xDispUnits),
  )

  const yAxChildren: string[] = [
    xmlSelfClose("c:axId", { val: AXIS_ID_VAL_Y }),
    buildAxisScaling(opts.yScale, opts.yReverse),
    xmlSelfClose("c:delete", { val: opts.yHidden ? 1 : 0 }),
    xmlSelfClose("c:axPos", { val: "l" }),
    ...buildAxisGridlines(opts.yGridlines),
  ]
  if (opts.yAxisTitle)
    yAxChildren.push(
      buildAxisTitle(
        opts.yAxisTitle,
        opts.yAxisTitleRotation,
        opts.yAxisTitleFontSize,
        opts.yAxisTitleBold,
        opts.yAxisTitleItalic,
        opts.yAxisTitleColor,
        opts.yAxisTitleStrike,
        opts.yAxisTitleUnderline,
        opts.yAxisTitleFontFamily,
        opts.yAxisTitleOverlay,
        opts.yAxisTitleLayout,
        opts.yAxisTitleFillColor,
        opts.yAxisTitleBorderColor,
        opts.yAxisTitleBorderWidth,
        opts.yAxisTitleBorderDash,
        opts.yAxisTitleBorderCap,
        opts.yAxisTitleBorderCompound,
      ),
    )
  yAxChildren.push(
    ...buildAxisNumFmt(opts.yNumFmt),
    ...buildAxisTickRendering(opts.yMajorTickMark, opts.yMinorTickMark, opts.yTickLblPos),
  )
  // `<c:txPr>` slot for the scatter Y axis — same CT_ValAx position
  // and omit-by-default contract as the catAx / valAx builders above.
  const yScatterTxPr = buildAxisTxPr(
    opts.yLabelRotation,
    opts.yLabelFontSize,
    opts.yLabelBold,
    opts.yLabelItalic,
    opts.yLabelColor,
    opts.yLabelUnderline,
    opts.yLabelStrike,
    opts.yLabelFontFamily,
  )
  if (yScatterTxPr) yAxChildren.push(yScatterTxPr)
  yAxChildren.push(
    xmlSelfClose("c:crossAx", { val: AXIS_ID_VAL_X }),
    buildAxisCrosses(opts.yCrosses),
    // Scatter Y axis defaults to `"midCat"`. Same override grammar as
    // the X axis above.
    xmlSelfClose("c:crossBetween", { val: opts.yCrossBetween ?? "midCat" }),
    ...buildAxisTickUnits(opts.yScale),
    // `<c:dispUnits>` on the Y axis. Scatter Y is also a value axis,
    // so the same builder applies. See `buildBarAxes` for the broader
    // scope notes.
    ...buildAxisDispUnits(opts.yDispUnits),
  )

  return [
    xmlElement("c:valAx", undefined, xAxChildren),
    xmlElement("c:valAx", undefined, yAxChildren),
  ]
}

/**
 * Build a `<c:title>` for an axis. The structure mirrors the chart-
 * level title but renders the label at a smaller default font (10pt vs
 * 14pt) to match Excel's axis-title style.
 *
 * The optional `rotationDeg` parameter pins the title's
 * `<a:bodyPr rot="N"/>` attribute. The OOXML attribute is in 60000ths
 * of a degree; the writer holds the rotation in whole degrees and
 * converts at emit time. Absence (`undefined`) collapses to the OOXML
 * default `0` so a fresh chart matches Excel's reference serialization
 * byte-for-byte. Mirrors the chart-title `buildTitle` slot exactly so
 * an axis title and the chart-level title carry the same shape.
 *
 * The optional `fontSizePt` parameter pins the title's
 * `<a:defRPr sz="N"/>` / `<a:rPr sz="N"/>` attributes. The OOXML
 * attribute is in 100ths of a point; the writer holds the size in
 * points and converts at emit time. Absence (`undefined`) collapses
 * to the hardcoded `1000` (10pt) default Excel itself emits on a
 * fresh axis title, keeping the rendered shape stable for callers
 * that never set the field. The size lands on both the
 * default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>` so
 * a re-parse picks the value up off either canonical slot.
 *
 * The optional `bold` parameter pins the title's `<a:defRPr b=".."/>`
 * / `<a:rPr b=".."/>` attributes. The OOXML `b` attribute is the
 * `xsd:boolean` bold flag on `CT_TextCharacterProperties`; the writer
 * emits `1` / `0` at the canonical slots. Absence (`undefined`)
 * collapses to the OOXML default `0` (non-bold) so a fresh chart
 * matches Excel's reference serialization byte-for-byte. The flag
 * lands on both `<a:defRPr>` and `<a:rPr>` so a re-parse picks the
 * value up off either canonical slot — Excel keeps the two attributes
 * in sync.
 *
 * The optional `italic` parameter pins the title's
 * `<a:defRPr i=".."/>` / `<a:rPr i=".."/>` attributes. The OOXML
 * attribute is the `xsd:boolean` `i` on `CT_TextCharacterProperties`.
 * Absence (`undefined`) and explicit `false` both collapse to omitting
 * the attribute (Excel's reference serialization for a non-italic
 * axis title — only the bold flag is always emitted); only `true`
 * emits `i="1"` on both slots so a re-parse picks the flag up off
 * either canonical slot.
 *
 * The optional `rgbHex` parameter pins the title's
 * `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
 * </a:defRPr>` / `<a:rPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:rPr>` font color. Mirrors the chart-level
 * `buildTitle` color emit: when set, the `<a:defRPr>` / `<a:rPr>`
 * slots expand from self-closing to wrapping a single
 * `<a:solidFill>` child; otherwise the writer keeps the existing
 * self-closing form so a fresh axis title with no custom color
 * matches Excel's reference serialization byte-for-byte (the title
 * inherits the theme text color in that case).
 */
export function buildAxisTitle(
  label: string,
  rotationDeg: number | undefined,
  fontSizePt: number | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
  rgbHex: ChartColor | undefined,
  strike: boolean | undefined,
  underline: boolean | undefined,
  fontFamily: string | undefined,
  overlay: boolean,
  layout: ResolvedManualLayout | undefined,
  fillRgbHex: ChartColor | undefined,
  borderRgbHex: ChartColor | undefined,
  borderWidthPt: number | undefined,
  borderDash: ChartBorderDash | undefined,
  borderCap?: ChartLineCap | undefined,
  borderCompound?: ChartLineCompound | undefined,
): string {
  const rot = rotationDeg === undefined ? 0 : rotationDeg * TITLE_ROT_PER_DEGREE
  // OOXML's `<a:defRPr sz="N"/>` / `<a:rPr sz="N"/>` attribute is in
  // 100ths of a point. The writer holds the size in points and
  // converts at emit time. Absence (`undefined`) collapses to the
  // hardcoded `1000` (10pt) default — Excel's reference axis-title
  // size — so callers that never pin the field still match Excel's
  // shape byte-for-byte. The size lands on both `<a:defRPr>` and
  // `<a:rPr>` so a re-parse picks the value up off either canonical
  // slot, mirroring the chart-level `buildTitle` writer.
  const sz =
    fontSizePt === undefined
      ? AXIS_TITLE_DEFAULT_FONT_SIZE_SZ
      : fontSizePt * TITLE_FONT_SZ_PER_POINT
  // OOXML's `<a:defRPr b=".."/>` / `<a:rPr b=".."/>` attribute is the
  // `xsd:boolean` bold flag on `CT_TextCharacterProperties`. The
  // writer holds `axisTitleBold` as a boolean and emits `1` / `0` at
  // the canonical slots. Absence (`undefined`) collapses to the OOXML
  // default `0` (non-bold) so a fresh chart matches Excel's reference
  // serialization byte-for-byte. The flag lands on both
  // `<a:defRPr>` and `<a:rPr>` so a re-parse picks the value up off
  // either canonical slot, mirroring the chart-level `buildTitle`
  // writer.
  const b = bold ? 1 : 0
  // OOXML's `<a:defRPr i=".."/>` / `<a:rPr i=".."/>` attribute is the
  // `xsd:boolean` italic flag on `CT_TextCharacterProperties`. Mirrors
  // the chart-level `buildTitle` italic emit: `axisTitleItalic` lands
  // on both the default-paragraph `<a:defRPr>` and the literal run's
  // `<a:rPr>` so a re-parse picks the value up off either canonical
  // slot — Excel keeps the two attributes in sync. Absence
  // (`undefined`) and explicit `false` both collapse to omitting the
  // attribute so a fresh axis title matches Excel's reference
  // serialization byte-for-byte (Excel itself omits `i` on a non-
  // italic axis title — only the bold flag is always emitted).
  const i = italic === true ? 1 : undefined
  // OOXML's `<a:defRPr strike=".."/>` / `<a:rPr strike=".."/>` attribute
  // is the `ST_TextStrikeType` enum on `CT_TextCharacterProperties` —
  // `"noStrike"` (default), `"sngStrike"` (single line, the value
  // Excel's UI emits), `"dblStrike"` (double line, non-UI). The writer
  // emits only the UI variant `"sngStrike"` to keep the surfaced shape
  // consistent with what Excel's reference UI authors. Absence
  // (`undefined`) and explicit `false` both collapse to omitting the
  // attribute (Excel itself omits `strike` when the title is not
  // strikethrough — the OOXML default `"noStrike"` collapses to
  // absence). Like bold / italic, the value lands on both the
  // default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>` so
  // a re-parse picks the value up off either canonical slot — Excel
  // keeps the two attributes in sync.
  const strikeAttr = strike === true ? "sngStrike" : undefined
  // OOXML's `<a:defRPr u=".."/>` / `<a:rPr u=".."/>` attribute is the
  // `ST_TextUnderlineType` enum on `CT_TextCharacterProperties` —
  // eighteen values total, with `"none"` as the OOXML default,
  // `"sng"` as the value Excel's UI authors for the "Underline"
  // checkbox (single line), `"dbl"` for the non-UI double-line
  // variant, and sixteen exotic types Excel does not surface. The
  // writer emits only the UI variant `"sng"` to keep the surfaced
  // shape consistent with what Excel's reference UI authors. Absence
  // (`undefined`) and explicit `false` both collapse to omitting the
  // attribute (Excel itself omits `u` when the title is not
  // underlined — the OOXML default `"none"` collapses to absence).
  // Like bold / italic / strike, the value lands on both the
  // default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>` so
  // a re-parse picks the value up off either canonical slot — Excel
  // keeps the two attributes in sync.
  const underlineAttr = underline === true ? "sng" : undefined
  // OOXML's `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:defRPr>` carries the title's font color. Mirrors
  // the chart-level `buildTitle` color emit: `axisTitleColor` lands on
  // both the default-paragraph `<a:defRPr>` and the literal run's
  // `<a:rPr>` so a re-parse picks the value up off either canonical
  // slot. Absence (`undefined`) collapses to omitting the entire
  // `<a:solidFill>` block so the title inherits the theme text color
  // (Excel's reference behavior for a fresh axis title that has not
  // had a custom color picked).
  const solidFillChild =
    rgbHex !== undefined
      ? xmlElement("a:solidFill", undefined, [buildColorElement(rgbHex)])
      : undefined
  // OOXML's `<a:defRPr><a:latin typeface=".."/></a:defRPr>` carries the
  // axis title's font family. Mirrors the chart-level `buildTitle`
  // typeface emit: `axisTitleFontFamily` lands on both the default-
  // paragraph `<a:defRPr>` and the literal run's `<a:rPr>` so a
  // re-parse picks the typeface up off either canonical slot. Absence
  // (`undefined`) collapses to omitting the entire `<a:latin>` element
  // so the title inherits the theme typeface (Excel's reference
  // behavior for a fresh axis title that has not had a custom font
  // picked). The `<a:latin>` element follows `<a:solidFill>` per the
  // CT_TextCharacterProperties child sequence (ECMA-376 Part 1,
  // §21.1.2.3.7).
  const latinChild = fontFamily ? xmlSelfClose("a:latin", { typeface: fontFamily }) : undefined
  // When a fill color or a typeface is set the `<a:defRPr>` /
  // `<a:rPr>` slots expand from self-closing to wrapping the
  // children; otherwise the writer keeps the existing self-closing
  // form so a fresh axis title with no custom color or font matches
  // Excel's reference serialization byte-for-byte. Children are
  // emitted in CT_TextCharacterProperties' canonical schema order:
  // solidFill first, then latin.
  const rPrChildren: string[] = []
  if (solidFillChild) rPrChildren.push(solidFillChild)
  if (latinChild) rPrChildren.push(latinChild)
  const defRPr =
    rPrChildren.length > 0
      ? xmlElement("a:defRPr", { sz, b, i, u: underlineAttr, strike: strikeAttr }, rPrChildren)
      : xmlSelfClose("a:defRPr", { sz, b, i, u: underlineAttr, strike: strikeAttr })
  const rPr =
    rPrChildren.length > 0
      ? xmlElement(
          "a:rPr",
          { lang: "en-US", sz, b, i, u: underlineAttr, strike: strikeAttr },
          rPrChildren,
        )
      : xmlSelfClose("a:rPr", {
          lang: "en-US",
          sz,
          b,
          i,
          u: underlineAttr,
          strike: strikeAttr,
        })
  // `<c:layout>` sits between `<c:tx>` and `<c:overlay>` per CT_Title
  // (ECMA-376 Part 1, §21.2.2.210) — the schema sequence is
  // `<c:tx>?` / `<c:layout>?` / `<c:overlay>?` / `<c:spPr>?` /
  // `<c:txPr>?`. Skip the entire block when `layout` is `undefined`
  // (every coordinate either unset or dropped on normalization) so a
  // fresh axis title matches Excel's reference shape byte-for-byte.
  const layoutXml = buildManualLayout(layout)
  const titleChildren: string[] = [
    xmlElement("c:tx", undefined, [
      xmlElement("c:rich", undefined, [
        xmlElement(
          "a:bodyPr",
          {
            rot,
            spcFirstLastPara: 1,
            vertOverflow: "ellipsis",
            wrap: "square",
            anchor: "ctr",
            anchorCtr: 1,
          },
          [],
        ),
        xmlSelfClose("a:lstStyle"),
        xmlElement("a:p", undefined, [
          xmlElement("a:pPr", undefined, [defRPr]),
          xmlElement("a:r", undefined, [rPr, xmlElement("a:t", undefined, xmlEscape(label))]),
        ]),
      ]),
    ]),
  ]
  if (layoutXml) titleChildren.push(layoutXml)
  titleChildren.push(xmlSelfClose("c:overlay", { val: overlay ? 1 : 0 }))
  // CT_Title (ECMA-376 Part 1, §21.2.2.210) places the optional
  // `<c:spPr>` between `<c:overlay>` and `<c:txPr>` / `<c:extLst>`.
  // Mirrors `buildTitle`: the writer skips emission entirely when
  // neither the fill nor the border color is pinned so a fresh axis
  // title matches Excel's reference shape byte-for-byte (Excel
  // itself omits the block whenever the title renders at the theme
  // defaults — typically a transparent title background with no
  // visible border, no `<c:spPr>` block). Authors `<a:solidFill>`
  // for the fill ({@link SheetChart.axes.x.axisTitleFillColor}) and
  // `<a:ln>` for the stroke
  // ({@link SheetChart.axes.x.axisTitleBorderColor}) in
  // CT_ShapeProperties schema order; other CT_ShapeProperties
  // children (effects, gradient / pattern / picture fills, line
  // dash / width / compound styles) are not modelled at this layer.
  // Distinct from the `<a:defRPr><a:solidFill>` font-color slot
  // inside `<c:tx><c:rich>` that
  // {@link SheetChart.axes.x.axisTitleColor} pins — the typography
  // knobs target different children of `<c:title>` so a caller can
  // pin all three without conflict.
  const titleSpPrXml = buildTitleSpPr(
    fillRgbHex,
    borderRgbHex,
    borderWidthPt,
    borderDash,
    borderCap,
    borderCompound,
  )
  if (titleSpPrXml !== undefined) titleChildren.push(titleSpPrXml)
  return xmlElement("c:title", undefined, titleChildren)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleRotation value (whole
 * degrees) for the `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`
 * writer slot inside an axis. Same conversion / clamping grammar as
 * the chart-level {@link normalizeTitleRotation} — non-finite,
 * non-numeric, and out-of-range inputs collapse to `undefined` (or
 * clamp to the `-90..90` band Excel's UI exposes), and the OOXML
 * default `0` collapses to `undefined` so absence and the default
 * round-trip identically through {@link cloneChart}.
 */
export function normalizeAxisTitleRotation(value: number | undefined): number | undefined {
  return normalizeTitleRotation(value)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleFontSize value
 * (whole / half points) for the `<c:title><c:tx><c:rich><a:p><a:pPr>
 * <a:defRPr sz="N"/></a:pPr></a:p></c:rich></c:tx></c:title>` writer
 * slot inside an axis. Delegates to the chart-level
 * {@link normalizeTitleFontSize} so the two share a clamping band
 * (`1..400`pt, the OOXML `ST_TextFontSize` schema range) and the
 * same 0.5pt half-step rounding. Out-of-range, non-finite, and
 * non-numeric inputs all collapse to `undefined` so the writer falls
 * back to the hardcoded 10pt axis-title default.
 */
export function normalizeAxisTitleFontSize(value: number | undefined): number | undefined {
  return normalizeTitleFontSize(value)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleBold value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` writer slot inside an axis. Delegates
 * to the chart-level {@link normalizeTitleBold} — `true` / `false`
 * pass through literally, every other token (typed escape from an
 * untyped caller, including `null`-shaped values) collapses to
 * `undefined` so the writer falls back to the OOXML default `b="0"`
 * (non-bold) Excel itself emits on a fresh axis title.
 */
export function normalizeAxisTitleBold(value: boolean | undefined): boolean | undefined {
  return normalizeTitleBold(value)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleItalic value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` writer slot inside an axis.
 * Delegates to the chart-level {@link normalizeTitleItalic} so the two
 * share the same drop-on-non-boolean grammar — `true` / `false` pass
 * through literally, every other token (typed escape from an untyped
 * caller) collapses to `undefined` and the writer omits the `i`
 * attribute (Excel's reference serialization for a non-italic axis
 * title).
 */
export function normalizeAxisTitleItalic(value: boolean | undefined): boolean | undefined {
  return normalizeTitleItalic(value)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleColor value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` writer slot inside an axis. Delegates
 * to the chart-level {@link normalizeTitleColor} so the two share the
 * same accept-with-or-without-`#` grammar — `"FF0000"`, `"#FF0000"`,
 * and `"ff0000"` all collapse to the OOXML uppercase canonical form;
 * malformed inputs (wrong length, non-hex characters, alpha-channel
 * forms, non-string escapes from an untyped caller) collapse to
 * `undefined` so the writer skips the entire `<a:solidFill>` block
 * and the axis title inherits the theme text color (Excel's
 * reference behavior for a fresh axis title without a custom color).
 */
export function normalizeAxisTitleColor(value: ChartColor | undefined): ChartColor | undefined {
  return normalizeTitleColor(value)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleStrike value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` writer slot inside an axis.
 * Delegates to the chart-level {@link normalizeTitleStrike} so the two
 * share the same drop-on-non-boolean grammar — `true` / `false` pass
 * through literally, every other token (typed escape from an untyped
 * caller) collapses to `undefined` and the writer omits the `strike`
 * attribute (Excel's reference serialization for a non-strikethrough
 * axis title).
 */
export function normalizeAxisTitleStrike(value: boolean | undefined): boolean | undefined {
  return normalizeTitleStrike(value)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleUnderline value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` writer slot inside an axis.
 * Delegates to the chart-level {@link normalizeTitleUnderline} so the
 * two share the same drop-on-non-boolean grammar — `true` / `false`
 * pass through literally, every other token (typed escape from an
 * untyped caller) collapses to `undefined` and the writer omits the
 * `u` attribute (Excel's reference serialization for a non-underlined
 * axis title).
 */
export function normalizeAxisTitleUnderline(value: boolean | undefined): boolean | undefined {
  return normalizeTitleUnderline(value)
}

/**
 * Normalize a {@link SheetChart.axes}.x.axisTitleFontFamily value for
 * the `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>`
 * writer slot inside an axis. Returns the trimmed typeface string
 * when the input is a non-empty string, or `undefined` for any
 * malformed token — empty / whitespace-only strings, or non-string
 * escapes from an untyped caller (`null`, numbers, booleans, etc.).
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the entire `<a:latin>` element and the axis title
 * inherits the theme typeface (Excel's reference behavior for a
 * fresh axis title without a custom font picked).
 */
export function normalizeAxisTitleFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

// ── Clone-side axis constants ─────────────────────────────────────

/** Recognized values of `<c:majorTickMark>` / `<c:minorTickMark>` (clone-side). */
const VALID_TICK_MARK_VALUES: ReadonlySet<ChartAxisTickMark> = new Set([
  "none",
  "in",
  "out",
  "cross",
])

/** Recognized values of `<c:tickLblPos>` (clone-side). */
const VALID_TICK_LBL_POS_VALUES: ReadonlySet<ChartAxisTickLabelPosition> = new Set([
  "nextTo",
  "low",
  "high",
  "none",
])

/** Recognized values of `<c:crosses>` per the OOXML `ST_Crosses` enum (clone-side). */
const VALID_CROSSES_VALUES: ReadonlySet<ChartAxisCrosses> = new Set(["autoZero", "min", "max"])

interface CrossesPairSource {
  crosses?: ChartAxisCrosses
  crossesAt?: number
}

interface CrossesPairOverride {
  crosses?: ChartAxisCrosses | null
  crossesAt?: number | null
}

interface CrossesPair {
  crosses?: ChartAxisCrosses
  crossesAt?: number
}

/** Recognized values of `<c:builtInUnit>` per the OOXML `ST_BuiltInUnit` enum (clone-side). */
const VALID_DISP_UNIT_VALUES: ReadonlySet<ChartAxisDispUnit> = new Set([
  "hundreds",
  "thousands",
  "tenThousands",
  "hundredThousands",
  "millions",
  "tenMillions",
  "hundredMillions",
  "billions",
  "trillions",
])

/** Recognized values of `<c:crossBetween>` per the OOXML `ST_CrossBetween` enum (clone-side). */
const VALID_CROSS_BETWEEN_VALUES: ReadonlySet<ChartAxisCrossBetween> = new Set([
  "between",
  "midCat",
])

// ── Clone ─────────────────────────────────────────────────────────

/**
 * Resolve an `autoTitleDeleted` override.
 *
 * `undefined` → inherit the source's parsed `autoTitleDeleted`.
 * `null`      → drop the inherited value (the writer falls back to its
 *               title-presence-derived default — `val="0"` when the
 *               cloned chart has a literal title, `val="1"` when it
 *               does not).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleOverlay` / `roundedCorners` /
 * `plotVisOnly` so the chart-level title flags compose the same way
 * at the call site. Independent of the resolved title presence —
 * `<c:autoTitleDeleted>` sits on `<c:chart>` directly (between
 * `<c:title>` and `<c:plotArea>` per CT_Chart, ECMA-376 Part 1,
 * §21.2.2.4), not nested inside `<c:title>`, so a clone with no
 * literal title can still pin `false` (let Excel synthesise the
 * series-name auto-title) and a clone with a literal title can pin
 * `true` (suppress the synthesis even though the literal renders).
 */
export function resolveCloneAutoTitleDeleted(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue
  if (override === null) return undefined
  return override
}

/**
 * Merge the source chart's `axes` block with per-axis overrides. The
 * result mirrors the writer's {@link SheetChart.axes} shape — missing
 * fields are dropped so the writer doesn't emit empty `<c:title>`
 * elements or redundant gridline blocks.
 */
export function resolveAxes(
  sourceAxes: Chart["axes"],
  overrides: CloneChartOptions["axes"],
  type: WriteChartKind,
): SheetChart["axes"] | undefined {
  const xTitle = applyOverride(sourceAxes?.x?.title, overrides?.x?.title)
  const yTitle = applyOverride(sourceAxes?.y?.title, overrides?.y?.title)
  // `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx></c:title>`
  // lives on every axis flavour per the OOXML schema (CT_CatAx,
  // CT_ValAx, CT_DateAx, CT_SerAx all share the same `<c:title>`
  // shape), so the resolver applies on every chart family that has
  // axes (pie / doughnut were short-circuited upstream). Out-of-range
  // / non-numeric values clamp to the `-90..90` band the writer
  // accepts; the OOXML default `0` collapses to `undefined` so absence
  // and the default round-trip identically. The writer drops the
  // rotation when the matching axis title is unset, so a stray pin on
  // an axis with no title silently disappears at emit time.
  const xAxisTitleRotation = applyAxisTitleRotationOverride(
    sourceAxes?.x?.axisTitleRotation,
    overrides?.x?.axisTitleRotation,
  )
  const yAxisTitleRotation = applyAxisTitleRotationOverride(
    sourceAxes?.y?.axisTitleRotation,
    overrides?.y?.axisTitleRotation,
  )
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis title font size in 100ths of a
  // point. Sits on the same `<c:title>` body as `axisTitleRotation`,
  // so the resolver applies on every chart family that has axes (pie /
  // doughnut were short-circuited upstream). Out-of-range / non-finite
  // / non-numeric values collapse to `undefined` so the writer falls
  // back to the hardcoded 10pt axis-title default. Like the rotation,
  // the writer drops the size when the matching axis title is unset,
  // so a stray pin on an axis with no title silently disappears at
  // emit time.
  const xAxisTitleFontSize = applyAxisTitleFontSizeOverride(
    sourceAxes?.x?.axisTitleFontSize,
    overrides?.x?.axisTitleFontSize,
  )
  const yAxisTitleFontSize = applyAxisTitleFontSizeOverride(
    sourceAxes?.y?.axisTitleFontSize,
    overrides?.y?.axisTitleFontSize,
  )
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis-title bold flag. Sits on the
  // same `<c:title>` body as `axisTitleRotation`, so the resolver
  // applies on every chart family that has axes (pie / doughnut were
  // short-circuited upstream). Non-boolean overrides collapse to a
  // drop so the cloned `SheetChart` always carries a value the writer
  // will accept. Like the rotation, the writer drops the flag when
  // the matching axis title is unset, so a stray pin on an axis with
  // no title silently disappears at emit time.
  const xAxisTitleBold = applyAxisTitleBoldOverride(
    sourceAxes?.x?.axisTitleBold,
    overrides?.x?.axisTitleBold,
  )
  const yAxisTitleBold = applyAxisTitleBoldOverride(
    sourceAxes?.y?.axisTitleBold,
    overrides?.y?.axisTitleBold,
  )
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis title italic flag. Sits on the
  // same `<c:title>` body as `axisTitleRotation` / `axisTitleFontSize` /
  // `axisTitleBold`, so the resolver applies on every chart family
  // that has axes (pie / doughnut were short-circuited upstream).
  // Non-boolean overrides collapse to `undefined` so the writer omits
  // the `i` attribute. Like the rotation and the size, the writer
  // drops the flag when the matching axis title is unset, so a stray
  // pin on an axis with no title silently disappears at emit time.
  const xAxisTitleItalic = applyAxisTitleItalicOverride(
    sourceAxes?.x?.axisTitleItalic,
    overrides?.x?.axisTitleItalic,
  )
  const yAxisTitleItalic = applyAxisTitleItalicOverride(
    sourceAxes?.y?.axisTitleItalic,
    overrides?.y?.axisTitleItalic,
  )
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
  // <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
  // </c:rich></c:tx></c:title>` — axis title font color. Sits on the
  // same `<c:title>` body as `axisTitleRotation` / `axisTitleFontSize` /
  // `axisTitleBold` / `axisTitleItalic`, so the resolver applies on
  // every chart family that has axes (pie / doughnut were short-
  // circuited upstream). Malformed overrides collapse to a drop via
  // the normalizer so the cloned `SheetChart` always carries a value
  // the writer will accept. Like the rotation, size, bold, and italic
  // knobs, the writer drops the fill when the matching axis title is
  // unset, so a stray pin on an axis with no title silently disappears
  // at emit time.
  const xAxisTitleColor = applyAxisTitleColorOverride(
    sourceAxes?.x?.axisTitleColor,
    overrides?.x?.axisTitleColor,
  )
  const yAxisTitleColor = applyAxisTitleColorOverride(
    sourceAxes?.y?.axisTitleColor,
    overrides?.y?.axisTitleColor,
  )
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` — axis title strikethrough flag.
  // Sits on the same `<c:title>` body as `axisTitleRotation` /
  // `axisTitleFontSize` / `axisTitleBold` / `axisTitleItalic` /
  // `axisTitleColor`, so the resolver applies on every chart family
  // that has axes (pie / doughnut were short-circuited upstream).
  // Non-boolean overrides collapse to `undefined` so the writer omits
  // the `strike` attribute. Like the rotation, size, bold, italic, and
  // color knobs, the writer drops the flag when the matching axis
  // title is unset, so a stray pin on an axis with no title silently
  // disappears at emit time.
  const xAxisTitleStrike = applyAxisTitleStrikeOverride(
    sourceAxes?.x?.axisTitleStrike,
    overrides?.x?.axisTitleStrike,
  )
  const yAxisTitleStrike = applyAxisTitleStrikeOverride(
    sourceAxes?.y?.axisTitleStrike,
    overrides?.y?.axisTitleStrike,
  )
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
  // </a:p></c:rich></c:tx></c:title>` — axis title underline flag.
  // Sits on the same `<c:title>` body as `axisTitleRotation` /
  // `axisTitleFontSize` / `axisTitleBold` / `axisTitleItalic` /
  // `axisTitleColor` / `axisTitleStrike`, so the resolver applies on
  // every chart family that has axes (pie / doughnut were short-
  // circuited upstream). Non-boolean overrides collapse to `undefined`
  // so the writer omits the `u` attribute. Like the rotation, size,
  // bold, italic, color, and strike knobs, the writer drops the flag
  // when the matching axis title is unset, so a stray pin on an axis
  // with no title silently disappears at emit time.
  const xAxisTitleUnderline = applyAxisTitleUnderlineOverride(
    sourceAxes?.x?.axisTitleUnderline,
    overrides?.x?.axisTitleUnderline,
  )
  const yAxisTitleUnderline = applyAxisTitleUnderlineOverride(
    sourceAxes?.y?.axisTitleUnderline,
    overrides?.y?.axisTitleUnderline,
  )
  // `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
  // typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>` —
  // axis title font family. Sits on the same `<c:title>` body as
  // `axisTitleRotation` / `axisTitleFontSize` / `axisTitleBold` /
  // `axisTitleItalic` / `axisTitleColor` / `axisTitleStrike` /
  // `axisTitleUnderline`, so the resolver applies on every chart
  // family that has axes (pie / doughnut were short-circuited
  // upstream). Empty / whitespace-only / non-string overrides collapse
  // to `undefined` so the writer omits the `<a:latin>` element. Like
  // the other axis-title knobs, the writer drops the typeface when
  // the matching axis title is unset, so a stray pin on an axis with
  // no title silently disappears at emit time.
  const xAxisTitleFontFamily = applyAxisTitleFontFamilyOverride(
    sourceAxes?.x?.axisTitleFontFamily,
    overrides?.x?.axisTitleFontFamily,
  )
  const yAxisTitleFontFamily = applyAxisTitleFontFamilyOverride(
    sourceAxes?.y?.axisTitleFontFamily,
    overrides?.y?.axisTitleFontFamily,
  )
  // `<c:title><c:overlay val=".."/></c:title>` — axis-title overlay
  // flag. Sits as a direct child of `<c:title>` per CT_Title schema,
  // so the resolver applies on every chart family that has axes (pie
  // / doughnut were short-circuited upstream). Non-boolean overrides
  // collapse to `undefined` via the resolver. Like the other axis-
  // title knobs, the writer drops the flag when the matching axis
  // title is unset, so a stray pin on an axis with no title silently
  // disappears at emit time.
  const xAxisTitleOverlay = applyAxisTitleOverlayOverride(
    sourceAxes?.x?.axisTitleOverlay,
    overrides?.x?.axisTitleOverlay,
  )
  const yAxisTitleOverlay = applyAxisTitleOverlayOverride(
    sourceAxes?.y?.axisTitleOverlay,
    overrides?.y?.axisTitleOverlay,
  )
  // `<c:title><c:layout><c:manualLayout>...</c:manualLayout></c:layout>
  // </c:title>` — axis-title manual placement. Sits inside `<c:title>`
  // between `<c:tx>` and `<c:overlay>` per CT_Title schema, so the
  // resolver applies on every chart family that has axes (pie /
  // doughnut were short-circuited upstream). Out-of-range / non-finite /
  // non-numeric coordinates collapse on the matching axis via the
  // resolver; an override whose every coordinate dropped collapses to
  // `undefined` so the cloned `SheetChart` skips the entire `<c:layout>`
  // block. Like the other axis-title knobs, the writer drops the layout
  // when the matching axis title is unset, so a stray pin on an axis
  // with no title silently disappears at emit time.
  const xAxisTitleLayout = resolveAxisTitleLayout(
    sourceAxes?.x?.axisTitleLayout,
    overrides?.x?.axisTitleLayout,
  )
  const yAxisTitleLayout = resolveAxisTitleLayout(
    sourceAxes?.y?.axisTitleLayout,
    overrides?.y?.axisTitleLayout,
  )
  // `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></c:spPr></c:title>` — axis-title background fill.
  // Lives on the axis's `<c:title>` directly per CT_Title schema (the
  // `<c:spPr>` block follows `<c:overlay>` in the schema sequence).
  // Independent of `axisTitleColor` (which lives on the inner
  // `<a:defRPr><a:solidFill>` slot for the font color) — the two
  // resolvers walk disjoint paths so a caller can pin both knobs
  // without conflict. Malformed overrides (wrong length, non-hex
  // characters, alpha-channel forms, empty / whitespace-only strings,
  // non-string escapes from an untyped caller) collapse to a drop via
  // the normalizer so the cloned `SheetChart` always carries a value
  // the writer will accept. Like the other axis-title knobs, the
  // writer drops the fill when the matching axis title is unset, so
  // a stray pin on an axis with no title silently disappears at emit
  // time.
  const xAxisTitleFillColor = applyAxisTitleFillColorOverride(
    sourceAxes?.x?.axisTitleFillColor,
    overrides?.x?.axisTitleFillColor,
  )
  const yAxisTitleFillColor = applyAxisTitleFillColorOverride(
    sourceAxes?.y?.axisTitleFillColor,
    overrides?.y?.axisTitleFillColor,
  )
  // `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:ln></c:spPr></c:title>` — axis-title border
  // (line stroke) color. Lives on the axis's `<c:title>` directly per
  // CT_Title schema, on the same `<c:spPr>` block as the background
  // fill but on a sibling child (`<a:ln>` for the stroke,
  // `<a:solidFill>` for the fill). Independent of `axisTitleFillColor`
  // and `axisTitleColor` (font color on the inner `<a:defRPr>
  // <a:solidFill>` slot) — the three resolvers walk disjoint paths so
  // a caller can pin all three knobs without conflict. Malformed
  // overrides (wrong length, non-hex characters, alpha-channel forms,
  // empty / whitespace-only strings, non-string escapes from an
  // untyped caller) collapse to a drop via the normalizer so the
  // cloned `SheetChart` always carries a value the writer will
  // accept. Like the other axis-title knobs, the writer drops the
  // stroke when the matching axis title is unset, so a stray pin on
  // an axis with no title silently disappears at emit time.
  const xAxisTitleBorderColor = applyAxisTitleBorderColorOverride(
    sourceAxes?.x?.axisTitleBorderColor,
    overrides?.x?.axisTitleBorderColor,
  )
  const yAxisTitleBorderColor = applyAxisTitleBorderColorOverride(
    sourceAxes?.y?.axisTitleBorderColor,
    overrides?.y?.axisTitleBorderColor,
  )
  // `<c:title><c:spPr><a:ln w="EMU"/></c:spPr></c:title>` —
  // axis-title border (line stroke) thickness. Reuse the shared
  // {@link resolveBorderWidthPt} helper so the snap / clamp grammar
  // matches every other chart-frame border-width slot. Like every
  // other axis-title knob, the writer drops the width when the
  // matching axis title is unset, so a stray pin on an axis with no
  // title silently disappears at emit time.
  const xAxisTitleBorderWidth = resolveBorderWidthPt(
    sourceAxes?.x?.axisTitleBorderWidth,
    overrides?.x?.axisTitleBorderWidth,
  )
  const yAxisTitleBorderWidth = resolveBorderWidthPt(
    sourceAxes?.y?.axisTitleBorderWidth,
    overrides?.y?.axisTitleBorderWidth,
  )
  // `<c:title><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
  // </c:title>` — axis-title border preset dash pattern. Same accept-
  // or-drop grammar as every other chart-frame border-dash slot.
  const xAxisTitleBorderDash = resolveBorderDash(
    sourceAxes?.x?.axisTitleBorderDash,
    overrides?.x?.axisTitleBorderDash,
  )
  const yAxisTitleBorderDash = resolveBorderDash(
    sourceAxes?.y?.axisTitleBorderDash,
    overrides?.y?.axisTitleBorderDash,
  )
  const xGridlines = applyGridlinesOverride(sourceAxes?.x?.gridlines, overrides?.x?.gridlines)
  const yGridlines = applyGridlinesOverride(sourceAxes?.y?.gridlines, overrides?.y?.gridlines)
  const xScale = applyScaleOverride(sourceAxes?.x?.scale, overrides?.x?.scale)
  const yScale = applyScaleOverride(sourceAxes?.y?.scale, overrides?.y?.scale)
  const xNumFmt = applyNumberFormatOverride(sourceAxes?.x?.numberFormat, overrides?.x?.numberFormat)
  const yNumFmt = applyNumberFormatOverride(sourceAxes?.y?.numberFormat, overrides?.y?.numberFormat)
  const xMajorTickMark = applyTickMarkOverride(
    sourceAxes?.x?.majorTickMark,
    overrides?.x?.majorTickMark,
  )
  const yMajorTickMark = applyTickMarkOverride(
    sourceAxes?.y?.majorTickMark,
    overrides?.y?.majorTickMark,
  )
  const xMinorTickMark = applyTickMarkOverride(
    sourceAxes?.x?.minorTickMark,
    overrides?.x?.minorTickMark,
  )
  const yMinorTickMark = applyTickMarkOverride(
    sourceAxes?.y?.minorTickMark,
    overrides?.y?.minorTickMark,
  )
  const xTickLblPos = applyTickLblPosOverride(sourceAxes?.x?.tickLblPos, overrides?.x?.tickLblPos)
  const yTickLblPos = applyTickLblPosOverride(sourceAxes?.y?.tickLblPos, overrides?.y?.tickLblPos)
  // `<c:txPr><a:bodyPr rot="N"/></c:txPr>` lives on every axis flavour
  // per the OOXML schema (CT_CatAx, CT_ValAx, CT_DateAx, CT_SerAx all
  // carry an optional `<c:txPr>`), so the resolver applies on every
  // chart family that has axes (pie / doughnut were short-circuited
  // upstream). Out-of-range / non-numeric values clamp to the
  // `-90..90` band the writer accepts; the OOXML default `0` collapses
  // to `undefined` so absence and the default round-trip identically.
  const xLabelRotation = applyLabelRotationOverride(
    sourceAxes?.x?.labelRotation,
    overrides?.x?.labelRotation,
  )
  const yLabelRotation = applyLabelRotationOverride(
    sourceAxes?.y?.labelRotation,
    overrides?.y?.labelRotation,
  )
  // `<c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p></c:txPr>`
  // shares the same `<c:txPr>` slot as the rotation resolver above,
  // and the same per-axis scope rule (every axis flavour carries
  // `<c:txPr>`; pie / doughnut already short-circuited upstream).
  // Out-of-range / non-finite / non-numeric inputs collapse to
  // `undefined`; fractional inputs round to the nearest 0.5pt.
  const xLabelFontSize = applyLabelFontSizeOverride(
    sourceAxes?.x?.labelFontSize,
    overrides?.x?.labelFontSize,
  )
  const yLabelFontSize = applyLabelFontSizeOverride(
    sourceAxes?.y?.labelFontSize,
    overrides?.y?.labelFontSize,
  )
  // `<c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p></c:txPr>`
  // shares the same `<c:txPr>` slot as the rotation / size resolvers
  // above, and the same per-axis scope rule (every axis flavour
  // carries `<c:txPr>`; pie / doughnut already short-circuited
  // upstream). Non-boolean overrides collapse to a drop so the cloned
  // `SheetChart` always carries a value the writer will accept.
  const xLabelBold = applyLabelBoldOverride(sourceAxes?.x?.labelBold, overrides?.x?.labelBold)
  const yLabelBold = applyLabelBoldOverride(sourceAxes?.y?.labelBold, overrides?.y?.labelBold)
  // `<c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p></c:txPr>`
  // shares the same `<c:txPr>` slot as the rotation / size / bold
  // resolvers above, and the same per-axis scope rule (every axis
  // flavour carries `<c:txPr>`; pie / doughnut already short-circuited
  // upstream). Non-boolean overrides collapse to a drop so the cloned
  // `SheetChart` always carries a value the writer will accept.
  const xLabelItalic = applyLabelItalicOverride(
    sourceAxes?.x?.labelItalic,
    overrides?.x?.labelItalic,
  )
  const yLabelItalic = applyLabelItalicOverride(
    sourceAxes?.y?.labelItalic,
    overrides?.y?.labelItalic,
  )
  // `<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=".."/>
  // </a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>` shares the same
  // `<c:txPr>` slot as the rotation / size / bold / italic resolvers
  // above, and the same per-axis scope rule (every axis flavour
  // carries `<c:txPr>`; pie / doughnut already short-circuited
  // upstream). Malformed overrides collapse to a drop via the
  // normalizer so the cloned `SheetChart` always carries a value the
  // writer will accept.
  const xLabelColor = applyLabelColorOverride(sourceAxes?.x?.labelColor, overrides?.x?.labelColor)
  const yLabelColor = applyLabelColorOverride(sourceAxes?.y?.labelColor, overrides?.y?.labelColor)
  // `<c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p></c:txPr>`
  // shares the same `<c:txPr>` slot as the rotation / size / bold /
  // italic / color resolvers above, and the same per-axis scope rule
  // (every axis flavour carries `<c:txPr>`; pie / doughnut already
  // short-circuited upstream). Non-boolean overrides collapse to a
  // drop so the cloned `SheetChart` always carries a value the writer
  // will accept.
  const xLabelUnderline = applyLabelUnderlineOverride(
    sourceAxes?.x?.labelUnderline,
    overrides?.x?.labelUnderline,
  )
  const yLabelUnderline = applyLabelUnderlineOverride(
    sourceAxes?.y?.labelUnderline,
    overrides?.y?.labelUnderline,
  )
  // `<c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr></a:p></c:txPr>`
  // shares the same `<c:txPr>` slot as the rotation / size / bold /
  // italic / color / underline resolvers above, and the same per-axis
  // scope rule (every axis flavour carries `<c:txPr>`; pie / doughnut
  // already short-circuited upstream). Non-boolean overrides collapse
  // to a drop so the cloned `SheetChart` always carries a value the
  // writer will accept.
  const xLabelStrike = applyLabelStrikeOverride(
    sourceAxes?.x?.labelStrike,
    overrides?.x?.labelStrike,
  )
  const yLabelStrike = applyLabelStrikeOverride(
    sourceAxes?.y?.labelStrike,
    overrides?.y?.labelStrike,
  )
  // `<c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/></a:defRPr>
  // </a:pPr></a:p></c:txPr>` shares the same `<c:txPr>` slot as the
  // rotation / size / bold / italic / color / underline / strike
  // resolvers above, and the same per-axis scope rule (every axis
  // flavour carries `<c:txPr>`; pie / doughnut already short-
  // circuited upstream). Empty / whitespace-only / non-string
  // overrides collapse to a drop so the cloned `SheetChart` always
  // carries a value the writer will accept.
  const xLabelFontFamily = applyLabelFontFamilyOverride(
    sourceAxes?.x?.labelFontFamily,
    overrides?.x?.labelFontFamily,
  )
  const yLabelFontFamily = applyLabelFontFamilyOverride(
    sourceAxes?.y?.labelFontFamily,
    overrides?.y?.labelFontFamily,
  )
  const xReverse = applyReverseOverride(sourceAxes?.x?.reverse, overrides?.x?.reverse)
  const yReverse = applyReverseOverride(sourceAxes?.y?.reverse, overrides?.y?.reverse)
  // `tickLblSkip` / `tickMarkSkip` only render on category axes
  // (`<c:catAx>`). Scatter charts use two value axes, so the X axis
  // skip would be silently dropped by the writer anyway — collapse it
  // to undefined here so the cloned `SheetChart` accurately reflects
  // what the chart will paint.
  const isCatAxisX = type !== "scatter"
  const xTickLblSkip = isCatAxisX
    ? applySkipOverride(sourceAxes?.x?.tickLblSkip, overrides?.x?.tickLblSkip)
    : undefined
  const xTickMarkSkip = isCatAxisX
    ? applySkipOverride(sourceAxes?.x?.tickMarkSkip, overrides?.x?.tickMarkSkip)
    : undefined
  // `lblOffset` is also category-axis-only (CT_CatAx / CT_DateAx) per
  // the OOXML schema. Same scope rule as the skip elements above.
  const xLblOffset = isCatAxisX
    ? applyLblOffsetOverride(sourceAxes?.x?.lblOffset, overrides?.x?.lblOffset)
    : undefined
  // `lblAlgn` is category-axis-only as well (CT_CatAx / CT_DateAx
  // per ECMA-376 §21.2.2). Same scope as `lblOffset`.
  const xLblAlgn = isCatAxisX
    ? applyLblAlgnOverride(sourceAxes?.x?.lblAlgn, overrides?.x?.lblAlgn)
    : undefined
  // `noMultiLvlLbl` is even tighter — `CT_CatAx` only (no `<c:dateAx>`
  // slot per ECMA-376 §21.2.2). Reuse the catAx scope rule above; the
  // resolved chart type still funnels through `<c:catAx>` for every
  // bar / column / line / area family the writer supports.
  const xNoMultiLvlLbl = isCatAxisX
    ? applyNoMultiLvlLblOverride(sourceAxes?.x?.noMultiLvlLbl, overrides?.x?.noMultiLvlLbl)
    : undefined
  // `<c:auto>` is also `CT_CatAx`-only per ECMA-376 §21.2.2.7 — same
  // scope rule as `noMultiLvlLbl`. The flag defaults to `true` in the
  // OOXML schema (Excel auto-detects whether to render the axis as a
  // date or category axis), so the resolver collapses `true` to
  // `undefined` and only surfaces an explicit `false`.
  const xAuto = isCatAxisX ? applyAutoOverride(sourceAxes?.x?.auto, overrides?.x?.auto) : undefined
  // `<c:delete>` lives on every axis flavour — both `<c:catAx>` and
  // `<c:valAx>` accept it — so the hidden flag carries through every
  // chart family that has axes. Pie / doughnut have no axes at all
  // and the caller already short-circuited those above.
  const xHidden = applyHiddenOverride(sourceAxes?.x?.hidden, overrides?.x?.hidden)
  const yHidden = applyHiddenOverride(sourceAxes?.y?.hidden, overrides?.y?.hidden)
  // `<c:crosses>` and `<c:crossesAt>` live in an XSD choice on every
  // axis flavour. Resolve the pair together so the precedence rule
  // (numeric pin wins over semantic token) survives the inherit / null
  // / replace grammar — a `crossesAt` override of `null` falls through
  // to the (possibly inherited) semantic `crosses`, and vice versa.
  const xCrossesPair = applyCrossesOverride(
    { crosses: sourceAxes?.x?.crosses, crossesAt: sourceAxes?.x?.crossesAt },
    { crosses: overrides?.x?.crosses, crossesAt: overrides?.x?.crossesAt },
  )
  const yCrossesPair = applyCrossesOverride(
    { crosses: sourceAxes?.y?.crosses, crossesAt: sourceAxes?.y?.crossesAt },
    { crosses: overrides?.y?.crosses, crossesAt: overrides?.y?.crossesAt },
  )
  // `<c:dispUnits>` lives exclusively on `<c:valAx>` per ECMA-376
  // §21.2.2.32 (CT_ValAx → CT_DispUnits). Bar / column / line / area
  // route the X axis through `<c:catAx>`, so the X-axis override is
  // only honoured when the resolved chart type is `scatter` (both axes
  // are value axes). Pie / doughnut were already short-circuited
  // upstream — they have no axes at all. The Y axis is a value axis on
  // every remaining family, so the Y override always carries through.
  const xDispUnits =
    type === "scatter"
      ? applyDispUnitsOverride(sourceAxes?.x?.dispUnits, overrides?.x?.dispUnits)
      : undefined
  const yDispUnits = applyDispUnitsOverride(sourceAxes?.y?.dispUnits, overrides?.y?.dispUnits)
  // `<c:crossBetween>` is also value-axis-only per ECMA-376 §21.2.2.10
  // (CT_ValAx → CT_CrossBetween). Same scope rule as `dispUnits` — the
  // X-axis override is only honoured on scatter (both axes are value
  // axes); bar / column / line / area route X through `<c:catAx>` which
  // rejects `<c:crossBetween>`. The Y axis is a value axis on every
  // family that has axes, so the Y override always carries through.
  const xCrossBetween =
    type === "scatter"
      ? applyCrossBetweenOverride(sourceAxes?.x?.crossBetween, overrides?.x?.crossBetween)
      : undefined
  const yCrossBetween = applyCrossBetweenOverride(
    sourceAxes?.y?.crossBetween,
    overrides?.y?.crossBetween,
  )

  // The axis-title rotation only renders when the axis carries a
  // title — drop a stray inherited rotation when the resolved axis
  // title is unset so the cloned `SheetChart` accurately reflects what
  // the chart will paint. Symmetric with the writer's title-presence
  // gate (the per-family axis builder only invokes `buildAxisTitle`
  // when `opts.xAxisTitle` / `opts.yAxisTitle` is set).
  const xAxisTitleRotationResolved = xTitle === undefined ? undefined : xAxisTitleRotation
  const yAxisTitleRotationResolved = yTitle === undefined ? undefined : yAxisTitleRotation
  // The axis-title font size only renders when the axis carries a
  // title — drop a stray inherited size when the resolved axis title
  // is unset so the cloned `SheetChart` accurately reflects what the
  // chart will paint. Symmetric with the writer's title-presence gate
  // (the per-family axis builder only invokes `buildAxisTitle` when
  // `opts.xAxisTitle` / `opts.yAxisTitle` is set).
  const xAxisTitleFontSizeResolved = xTitle === undefined ? undefined : xAxisTitleFontSize
  const yAxisTitleFontSizeResolved = yTitle === undefined ? undefined : yAxisTitleFontSize
  // Same title-presence gate for the axis-title bold flag — drop a
  // stray inherited flag when the resolved axis title is unset so the
  // cloned `SheetChart` accurately reflects what the chart will paint.
  const xAxisTitleBoldResolved = xTitle === undefined ? undefined : xAxisTitleBold
  const yAxisTitleBoldResolved = yTitle === undefined ? undefined : yAxisTitleBold
  // The axis-title italic flag only renders when the axis carries a
  // title — drop a stray inherited flag when the resolved axis title
  // is unset so the cloned `SheetChart` accurately reflects what the
  // chart will paint. Symmetric with the writer's title-presence gate
  // (the per-family axis builder only invokes `buildAxisTitle` when
  // `opts.xAxisTitle` / `opts.yAxisTitle` is set).
  const xAxisTitleItalicResolved = xTitle === undefined ? undefined : xAxisTitleItalic
  const yAxisTitleItalicResolved = yTitle === undefined ? undefined : yAxisTitleItalic
  // The axis-title font color only renders when the axis carries a
  // title — drop a stray inherited fill when the resolved axis title
  // is unset so the cloned `SheetChart` accurately reflects what the
  // chart will paint. Symmetric with the writer's title-presence gate
  // (the per-family axis builder only invokes `buildAxisTitle` when
  // `opts.xAxisTitle` / `opts.yAxisTitle` is set).
  const xAxisTitleColorResolved = xTitle === undefined ? undefined : xAxisTitleColor
  const yAxisTitleColorResolved = yTitle === undefined ? undefined : yAxisTitleColor
  // The axis-title strikethrough flag only renders when the axis
  // carries a title — drop a stray inherited flag when the resolved
  // axis title is unset so the cloned `SheetChart` accurately reflects
  // what the chart will paint. Symmetric with the writer's
  // title-presence gate (the per-family axis builder only invokes
  // `buildAxisTitle` when `opts.xAxisTitle` / `opts.yAxisTitle` is
  // set).
  const xAxisTitleStrikeResolved = xTitle === undefined ? undefined : xAxisTitleStrike
  const yAxisTitleStrikeResolved = yTitle === undefined ? undefined : yAxisTitleStrike
  // The axis-title underline flag only renders when the axis carries a
  // title — drop a stray inherited flag when the resolved axis title
  // is unset so the cloned `SheetChart` accurately reflects what the
  // chart will paint. Symmetric with the writer's title-presence gate
  // (the per-family axis builder only invokes `buildAxisTitle` when
  // `opts.xAxisTitle` / `opts.yAxisTitle` is set).
  const xAxisTitleUnderlineResolved = xTitle === undefined ? undefined : xAxisTitleUnderline
  const yAxisTitleUnderlineResolved = yTitle === undefined ? undefined : yAxisTitleUnderline
  // Same title-presence gate as the rotation / size / bold / italic /
  // color / strike / underline resolved values — the writer skips the
  // entire `<a:latin>` element when the matching axis title is unset,
  // so the cloned `SheetChart` accurately reflects what the chart
  // will paint.
  const xAxisTitleFontFamilyResolved = xTitle === undefined ? undefined : xAxisTitleFontFamily
  const yAxisTitleFontFamilyResolved = yTitle === undefined ? undefined : yAxisTitleFontFamily
  // Same title-presence gate as the rotation / size / bold / italic /
  // color / strike / underline / font-family resolved values — the
  // writer always emits `<c:overlay>` when it emits the axis title,
  // and the writer skips the entire axis-title block when the matching
  // axis title is unset, so the cloned `SheetChart` accurately
  // reflects what the chart will paint.
  const xAxisTitleOverlayResolved = xTitle === undefined ? undefined : xAxisTitleOverlay
  const yAxisTitleOverlayResolved = yTitle === undefined ? undefined : yAxisTitleOverlay
  // Same title-presence gate as the other axis-title knobs — the writer
  // skips the entire `<c:layout>` block (and the surrounding `<c:title>`
  // element) when the matching axis title is unset, so the cloned
  // `SheetChart` accurately reflects what the chart will paint.
  const xAxisTitleLayoutResolved = xTitle === undefined ? undefined : xAxisTitleLayout
  const yAxisTitleLayoutResolved = yTitle === undefined ? undefined : yAxisTitleLayout
  // The axis-title background fill only renders when the axis carries
  // a title — drop a stray inherited fill when the resolved axis title
  // is unset so the cloned `SheetChart` accurately reflects what the
  // chart will paint. Symmetric with the writer's title-presence gate
  // (the per-family axis builder only invokes `buildAxisTitle` when
  // `opts.xAxisTitle` / `opts.yAxisTitle` is set, and the writer skips
  // the `<c:spPr>` block entirely when no fill is pinned).
  const xAxisTitleFillColorResolved = xTitle === undefined ? undefined : xAxisTitleFillColor
  const yAxisTitleFillColorResolved = yTitle === undefined ? undefined : yAxisTitleFillColor
  // The axis-title border (line stroke) only renders when the axis
  // carries a title — drop a stray inherited stroke when the resolved
  // axis title is unset so the cloned `SheetChart` accurately reflects
  // what the chart will paint. Symmetric with the writer's title-
  // presence gate (the per-family axis builder only invokes
  // `buildAxisTitle` when `opts.xAxisTitle` / `opts.yAxisTitle` is
  // set, and the writer skips the `<a:ln>` block entirely when no
  // border is pinned — and skips the entire `<c:spPr>` block when
  // both the fill and border are absent).
  const xAxisTitleBorderColorResolved = xTitle === undefined ? undefined : xAxisTitleBorderColor
  const yAxisTitleBorderColorResolved = yTitle === undefined ? undefined : yAxisTitleBorderColor
  // Same hidden-axis-title scoping as the color knob — drop the
  // inherited width / dash on an axis whose title is unset so the
  // cloned `SheetChart` accurately reflects what the chart will paint.
  const xAxisTitleBorderWidthResolved = xTitle === undefined ? undefined : xAxisTitleBorderWidth
  const yAxisTitleBorderWidthResolved = yTitle === undefined ? undefined : yAxisTitleBorderWidth
  const xAxisTitleBorderDashResolved = xTitle === undefined ? undefined : xAxisTitleBorderDash
  const yAxisTitleBorderDashResolved = yTitle === undefined ? undefined : yAxisTitleBorderDash

  const out: NonNullable<SheetChart["axes"]> = {}
  if (
    xTitle !== undefined ||
    xAxisTitleRotationResolved !== undefined ||
    xAxisTitleFontSizeResolved !== undefined ||
    xAxisTitleBoldResolved !== undefined ||
    xAxisTitleItalicResolved !== undefined ||
    xAxisTitleColorResolved !== undefined ||
    xAxisTitleStrikeResolved !== undefined ||
    xAxisTitleUnderlineResolved !== undefined ||
    xAxisTitleFontFamilyResolved !== undefined ||
    xAxisTitleOverlayResolved !== undefined ||
    xAxisTitleLayoutResolved !== undefined ||
    xAxisTitleFillColorResolved !== undefined ||
    xAxisTitleBorderColorResolved !== undefined ||
    xAxisTitleBorderWidthResolved !== undefined ||
    xAxisTitleBorderDashResolved !== undefined ||
    xGridlines !== undefined ||
    xScale !== undefined ||
    xNumFmt !== undefined ||
    xMajorTickMark !== undefined ||
    xMinorTickMark !== undefined ||
    xTickLblPos !== undefined ||
    xLabelRotation !== undefined ||
    xLabelFontSize !== undefined ||
    xLabelBold !== undefined ||
    xLabelItalic !== undefined ||
    xLabelColor !== undefined ||
    xLabelUnderline !== undefined ||
    xLabelStrike !== undefined ||
    xLabelFontFamily !== undefined ||
    xReverse !== undefined ||
    xTickLblSkip !== undefined ||
    xTickMarkSkip !== undefined ||
    xLblOffset !== undefined ||
    xLblAlgn !== undefined ||
    xNoMultiLvlLbl !== undefined ||
    xAuto !== undefined ||
    xHidden !== undefined ||
    xCrossesPair.crosses !== undefined ||
    xCrossesPair.crossesAt !== undefined ||
    xDispUnits !== undefined ||
    xCrossBetween !== undefined
  ) {
    out.x = {}
    if (xTitle !== undefined) out.x.title = xTitle
    if (xAxisTitleRotationResolved !== undefined)
      out.x.axisTitleRotation = xAxisTitleRotationResolved
    if (xAxisTitleFontSizeResolved !== undefined)
      out.x.axisTitleFontSize = xAxisTitleFontSizeResolved
    if (xAxisTitleBoldResolved !== undefined) out.x.axisTitleBold = xAxisTitleBoldResolved
    if (xAxisTitleItalicResolved !== undefined) out.x.axisTitleItalic = xAxisTitleItalicResolved
    if (xAxisTitleColorResolved !== undefined) out.x.axisTitleColor = xAxisTitleColorResolved
    if (xAxisTitleStrikeResolved !== undefined) out.x.axisTitleStrike = xAxisTitleStrikeResolved
    if (xAxisTitleUnderlineResolved !== undefined)
      out.x.axisTitleUnderline = xAxisTitleUnderlineResolved
    if (xAxisTitleFontFamilyResolved !== undefined)
      out.x.axisTitleFontFamily = xAxisTitleFontFamilyResolved
    if (xAxisTitleOverlayResolved !== undefined) out.x.axisTitleOverlay = xAxisTitleOverlayResolved
    if (xAxisTitleLayoutResolved !== undefined) out.x.axisTitleLayout = xAxisTitleLayoutResolved
    if (xAxisTitleFillColorResolved !== undefined)
      out.x.axisTitleFillColor = xAxisTitleFillColorResolved
    if (xAxisTitleBorderColorResolved !== undefined)
      out.x.axisTitleBorderColor = xAxisTitleBorderColorResolved
    if (xAxisTitleBorderWidthResolved !== undefined)
      out.x.axisTitleBorderWidth = xAxisTitleBorderWidthResolved
    if (xAxisTitleBorderDashResolved !== undefined)
      out.x.axisTitleBorderDash = xAxisTitleBorderDashResolved
    if (xGridlines !== undefined) out.x.gridlines = xGridlines
    if (xScale !== undefined) out.x.scale = xScale
    if (xNumFmt !== undefined) out.x.numberFormat = xNumFmt
    if (xMajorTickMark !== undefined) out.x.majorTickMark = xMajorTickMark
    if (xMinorTickMark !== undefined) out.x.minorTickMark = xMinorTickMark
    if (xTickLblPos !== undefined) out.x.tickLblPos = xTickLblPos
    if (xLabelRotation !== undefined) out.x.labelRotation = xLabelRotation
    if (xLabelFontSize !== undefined) out.x.labelFontSize = xLabelFontSize
    if (xLabelBold !== undefined) out.x.labelBold = xLabelBold
    if (xLabelItalic !== undefined) out.x.labelItalic = xLabelItalic
    if (xLabelColor !== undefined) out.x.labelColor = xLabelColor
    if (xLabelUnderline !== undefined) out.x.labelUnderline = xLabelUnderline
    if (xLabelStrike !== undefined) out.x.labelStrike = xLabelStrike
    if (xLabelFontFamily !== undefined) out.x.labelFontFamily = xLabelFontFamily
    if (xReverse !== undefined) out.x.reverse = xReverse
    if (xTickLblSkip !== undefined) out.x.tickLblSkip = xTickLblSkip
    if (xTickMarkSkip !== undefined) out.x.tickMarkSkip = xTickMarkSkip
    if (xLblOffset !== undefined) out.x.lblOffset = xLblOffset
    if (xLblAlgn !== undefined) out.x.lblAlgn = xLblAlgn
    if (xNoMultiLvlLbl !== undefined) out.x.noMultiLvlLbl = xNoMultiLvlLbl
    if (xAuto !== undefined) out.x.auto = xAuto
    if (xHidden !== undefined) out.x.hidden = xHidden
    if (xCrossesPair.crosses !== undefined) out.x.crosses = xCrossesPair.crosses
    if (xCrossesPair.crossesAt !== undefined) out.x.crossesAt = xCrossesPair.crossesAt
    if (xDispUnits !== undefined) out.x.dispUnits = xDispUnits
    if (xCrossBetween !== undefined) out.x.crossBetween = xCrossBetween
  }
  if (
    yTitle !== undefined ||
    yAxisTitleRotationResolved !== undefined ||
    yAxisTitleFontSizeResolved !== undefined ||
    yAxisTitleBoldResolved !== undefined ||
    yAxisTitleItalicResolved !== undefined ||
    yAxisTitleColorResolved !== undefined ||
    yAxisTitleStrikeResolved !== undefined ||
    yAxisTitleUnderlineResolved !== undefined ||
    yAxisTitleFontFamilyResolved !== undefined ||
    yAxisTitleOverlayResolved !== undefined ||
    yAxisTitleLayoutResolved !== undefined ||
    yAxisTitleFillColorResolved !== undefined ||
    yAxisTitleBorderColorResolved !== undefined ||
    yAxisTitleBorderWidthResolved !== undefined ||
    yAxisTitleBorderDashResolved !== undefined ||
    yGridlines !== undefined ||
    yScale !== undefined ||
    yNumFmt !== undefined ||
    yMajorTickMark !== undefined ||
    yMinorTickMark !== undefined ||
    yTickLblPos !== undefined ||
    yLabelRotation !== undefined ||
    yLabelFontSize !== undefined ||
    yLabelBold !== undefined ||
    yLabelItalic !== undefined ||
    yLabelColor !== undefined ||
    yLabelUnderline !== undefined ||
    yLabelStrike !== undefined ||
    yLabelFontFamily !== undefined ||
    yHidden !== undefined ||
    yReverse !== undefined ||
    yCrossesPair.crosses !== undefined ||
    yCrossesPair.crossesAt !== undefined ||
    yDispUnits !== undefined ||
    yCrossBetween !== undefined
  ) {
    out.y = {}
    if (yTitle !== undefined) out.y.title = yTitle
    if (yAxisTitleRotationResolved !== undefined)
      out.y.axisTitleRotation = yAxisTitleRotationResolved
    if (yAxisTitleFontSizeResolved !== undefined)
      out.y.axisTitleFontSize = yAxisTitleFontSizeResolved
    if (yAxisTitleBoldResolved !== undefined) out.y.axisTitleBold = yAxisTitleBoldResolved
    if (yAxisTitleItalicResolved !== undefined) out.y.axisTitleItalic = yAxisTitleItalicResolved
    if (yAxisTitleColorResolved !== undefined) out.y.axisTitleColor = yAxisTitleColorResolved
    if (yAxisTitleStrikeResolved !== undefined) out.y.axisTitleStrike = yAxisTitleStrikeResolved
    if (yAxisTitleUnderlineResolved !== undefined)
      out.y.axisTitleUnderline = yAxisTitleUnderlineResolved
    if (yAxisTitleFontFamilyResolved !== undefined)
      out.y.axisTitleFontFamily = yAxisTitleFontFamilyResolved
    if (yAxisTitleOverlayResolved !== undefined) out.y.axisTitleOverlay = yAxisTitleOverlayResolved
    if (yAxisTitleLayoutResolved !== undefined) out.y.axisTitleLayout = yAxisTitleLayoutResolved
    if (yAxisTitleFillColorResolved !== undefined)
      out.y.axisTitleFillColor = yAxisTitleFillColorResolved
    if (yAxisTitleBorderColorResolved !== undefined)
      out.y.axisTitleBorderColor = yAxisTitleBorderColorResolved
    if (yAxisTitleBorderWidthResolved !== undefined)
      out.y.axisTitleBorderWidth = yAxisTitleBorderWidthResolved
    if (yAxisTitleBorderDashResolved !== undefined)
      out.y.axisTitleBorderDash = yAxisTitleBorderDashResolved
    if (yGridlines !== undefined) out.y.gridlines = yGridlines
    if (yScale !== undefined) out.y.scale = yScale
    if (yNumFmt !== undefined) out.y.numberFormat = yNumFmt
    if (yMajorTickMark !== undefined) out.y.majorTickMark = yMajorTickMark
    if (yMinorTickMark !== undefined) out.y.minorTickMark = yMinorTickMark
    if (yTickLblPos !== undefined) out.y.tickLblPos = yTickLblPos
    if (yLabelRotation !== undefined) out.y.labelRotation = yLabelRotation
    if (yLabelFontSize !== undefined) out.y.labelFontSize = yLabelFontSize
    if (yLabelBold !== undefined) out.y.labelBold = yLabelBold
    if (yLabelItalic !== undefined) out.y.labelItalic = yLabelItalic
    if (yLabelColor !== undefined) out.y.labelColor = yLabelColor
    if (yLabelUnderline !== undefined) out.y.labelUnderline = yLabelUnderline
    if (yLabelStrike !== undefined) out.y.labelStrike = yLabelStrike
    if (yLabelFontFamily !== undefined) out.y.labelFontFamily = yLabelFontFamily
    if (yHidden !== undefined) out.y.hidden = yHidden
    if (yReverse !== undefined) out.y.reverse = yReverse
    if (yCrossesPair.crosses !== undefined) out.y.crosses = yCrossesPair.crosses
    if (yCrossesPair.crossesAt !== undefined) out.y.crossesAt = yCrossesPair.crossesAt
    if (yDispUnits !== undefined) out.y.dispUnits = yDispUnits
    if (yCrossBetween !== undefined) out.y.crossBetween = yCrossBetween
  }

  return out.x || out.y ? out : undefined
}

/**
 * Resolve a `tickLblSkip` / `tickMarkSkip` override using the same
 * `undefined` (inherit) / `null` (drop) / value (replace) grammar as
 * the other axis helpers. Out-of-range / non-positive values collapse
 * to `undefined` so they cannot leak into the writer (which would
 * silently drop them anyway via {@link normalizeAxisSkip}).
 */
export function applySkipOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined
    const rounded = Math.round(source)
    if (rounded < 1 || rounded > 32767 || rounded === 1) return undefined
    return rounded
  }
  if (override === null) return undefined
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined
  const rounded = Math.round(override)
  if (rounded < 1 || rounded > 32767 || rounded === 1) return undefined
  return rounded
}

/**
 * Resolve an `lblOffset` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis helpers.
 * Out-of-range / non-numeric values collapse to `undefined` so they
 * cannot leak into the writer (which would silently drop them anyway
 * via {@link normalizeAxisLblOffset}). The OOXML default `100` also
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
export function applyLblOffsetOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined
    const rounded = Math.round(source)
    if (rounded < 0 || rounded > 1000 || rounded === 100) return undefined
    return rounded
  }
  if (override === null) return undefined
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined
  const rounded = Math.round(override)
  if (rounded < 0 || rounded > 1000 || rounded === 100) return undefined
  return rounded
}

/**
 * Resolve an `lblAlgn` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis helpers.
 * Unknown tokens collapse to `undefined` so they cannot leak into the
 * writer (which would silently drop them anyway via
 * {@link normalizeAxisLblAlgn}). The OOXML default `"ctr"` also
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
export function applyLblAlgnOverride(
  source: ChartAxisLabelAlign | undefined,
  override: ChartAxisLabelAlign | null | undefined,
): ChartAxisLabelAlign | undefined {
  if (override === undefined) {
    if (source !== "l" && source !== "r" && source !== "ctr") return undefined
    return source === "ctr" ? undefined : source
  }
  if (override === null) return undefined
  if (override !== "l" && override !== "r" && override !== "ctr") return undefined
  return override === "ctr" ? undefined : override
}

/**
 * Resolve a `noMultiLvlLbl` override using the same `undefined`
 * (inherit) / `null` (drop) / `boolean` (replace) grammar as the
 * other axis helpers. Only `true` surfaces (the writer treats `false`
 * and absence identically — both produce `<c:noMultiLvlLbl val="0"/>`),
 * so an override of `false` collapses to `undefined` to keep the
 * cloned `SheetChart` shape minimal. Non-boolean inputs fall through
 * the type guard to `undefined`.
 */
export function applyNoMultiLvlLblOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined
  }
  if (override === null) return undefined
  return override === true ? true : undefined
}

/**
 * Resolve an `auto` override using the same `undefined` (inherit) /
 * `null` (drop) / `boolean` (replace) grammar as the other axis
 * helpers. Only `false` surfaces (the writer treats `true` and absence
 * identically — both produce `<c:auto val="1"/>`), so an override of
 * `true` collapses to `undefined` to keep the cloned `SheetChart`
 * shape minimal. Non-boolean inputs fall through the type guard to
 * `undefined`.
 *
 * Inverse of {@link applyNoMultiLvlLblOverride}: `<c:auto>` defaults to
 * `true` in the OOXML schema, so the helper collapses `true` rather
 * than `false` — symmetric with the parser-side default-collapse on
 * the read layer.
 */
export function applyAutoOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === false ? false : undefined
  }
  if (override === null) return undefined
  return override === false ? false : undefined
}

/**
 * Resolve an axis `hidden` override using the same `undefined`
 * (inherit) / `null` (drop) / `boolean` (replace) grammar as the
 * other axis helpers. Only `true` surfaces (the writer treats `false`
 * and absence identically — both produce `<c:delete val="0"/>`), so
 * an override of `false` collapses to `undefined` to keep the cloned
 * `SheetChart` shape minimal. Non-boolean inputs fall through the
 * type guard to `undefined`.
 */
export function applyHiddenOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined
  }
  if (override === null) return undefined
  return override === true ? true : undefined
}

/**
 * Resolve gridlines using the same `undefined` (inherit) / `null`
 * (drop) / object (replace) grammar as the other axis overrides.
 * Returns `undefined` when neither source nor override declares a
 * non-empty gridline configuration.
 */
export function applyGridlinesOverride(
  source: ChartAxisGridlines | undefined,
  override: ChartAxisGridlines | null | undefined,
): ChartAxisGridlines | undefined {
  if (override === undefined) {
    if (!source) return undefined
    const out: ChartAxisGridlines = {}
    if (source.major) out.major = true
    if (source.minor) out.minor = true
    return out.major || out.minor ? out : undefined
  }
  if (override === null) return undefined
  const out: ChartAxisGridlines = {}
  if (override.major === true) out.major = true
  if (override.minor === true) out.minor = true
  return out.major || out.minor ? out : undefined
}

/**
 * Resolve a scale override using the same `undefined` / `null` /
 * object grammar as {@link applyGridlinesOverride}. The override
 * replaces the source wholesale rather than merging field-by-field —
 * a partial template scale `{ min: 0 }` plus an override
 * `{ max: 100 }` yields `{ max: 100 }`, not `{ min: 0, max: 100 }`.
 * Per-field merges proved confusing in the dashboard composition flow
 * (callers expected the override to fully describe the target scale),
 * so wholesale replacement is the simpler contract.
 */
export function applyScaleOverride(
  source: ChartAxisScale | undefined,
  override: ChartAxisScale | null | undefined,
): ChartAxisScale | undefined {
  if (override === undefined) {
    if (!source) return undefined
    return cloneScale(source)
  }
  if (override === null) return undefined
  return cloneScale(override)
}

export function cloneScale(source: ChartAxisScale): ChartAxisScale | undefined {
  const out: ChartAxisScale = {}
  if (typeof source.min === "number" && Number.isFinite(source.min)) out.min = source.min
  if (typeof source.max === "number" && Number.isFinite(source.max)) out.max = source.max
  if (
    typeof source.majorUnit === "number" &&
    Number.isFinite(source.majorUnit) &&
    source.majorUnit > 0
  ) {
    out.majorUnit = source.majorUnit
  }
  if (
    typeof source.minorUnit === "number" &&
    Number.isFinite(source.minorUnit) &&
    source.minorUnit > 0
  ) {
    out.minorUnit = source.minorUnit
  }
  if (typeof source.logBase === "number" && Number.isFinite(source.logBase)) {
    out.logBase = source.logBase
  }
  return Object.keys(out).length > 0 ? out : undefined
}

/**
 * Resolve a number-format override. Same grammar as the other
 * per-axis helpers: `undefined` inherits, `null` drops, an object
 * replaces.
 */
export function applyNumberFormatOverride(
  source: ChartAxisNumberFormat | undefined,
  override: ChartAxisNumberFormat | null | undefined,
): ChartAxisNumberFormat | undefined {
  if (override === undefined) {
    if (!source) return undefined
    if (typeof source.formatCode !== "string" || source.formatCode.length === 0) return undefined
    const out: ChartAxisNumberFormat = { formatCode: source.formatCode }
    if (source.sourceLinked === true) out.sourceLinked = true
    return out
  }
  if (override === null) return undefined
  if (typeof override.formatCode !== "string" || override.formatCode.length === 0) return undefined
  const out: ChartAxisNumberFormat = { formatCode: override.formatCode }
  if (override.sourceLinked === true) out.sourceLinked = true
  return out
}

/**
 * Resolve a tick-mark override using the same `undefined` (inherit) /
 * `null` (drop) / value (replace) grammar as the other axis helpers.
 * Unknown / typo'd inputs collapse to `undefined` so the writer never
 * emits a value the OOXML `ST_TickMark` enum rejects.
 */
export function applyTickMarkOverride(
  source: ChartAxisTickMark | undefined,
  override: ChartAxisTickMark | null | undefined,
): ChartAxisTickMark | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined
    return VALID_TICK_MARK_VALUES.has(source) ? source : undefined
  }
  if (override === null) return undefined
  return VALID_TICK_MARK_VALUES.has(override) ? override : undefined
}

/**
 * Resolve a tick-label-position override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Unknown / typo'd inputs collapse to `undefined` so
 * the writer never emits a value the OOXML `ST_TickLblPos` enum
 * rejects.
 */
export function applyTickLblPosOverride(
  source: ChartAxisTickLabelPosition | undefined,
  override: ChartAxisTickLabelPosition | null | undefined,
): ChartAxisTickLabelPosition | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined
    return VALID_TICK_LBL_POS_VALUES.has(source) ? source : undefined
  }
  if (override === null) return undefined
  return VALID_TICK_LBL_POS_VALUES.has(override) ? override : undefined
}

/**
 * Resolve a `labelRotation` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Out-of-range / non-numeric values clamp to the
 * `-90..90` band the writer accepts and the OOXML default `0`
 * collapses to `undefined` so absence and the default round-trip
 * identically — symmetric with the parser-side default-collapse.
 */
export function applyLabelRotationOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined
    return clampLabelRotationDeg(source)
  }
  if (override === null) return undefined
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined
  return clampLabelRotationDeg(override)
}

/**
 * Snap a `labelRotation` value (whole degrees) into the `-90..90` band
 * the writer accepts. Returns `undefined` for `0` so the OOXML default
 * collapses to absence — symmetric with the writer-side
 * {@link normalizeAxisLabelRotation} contract.
 */
export function clampLabelRotationDeg(value: number): number | undefined {
  let degrees = Math.round(value)
  if (degrees < -90) degrees = -90
  else if (degrees > 90) degrees = 90
  if (degrees === 0) return undefined
  return degrees
}

/**
 * Resolve a `labelFontSize` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. The conversion / clamping rules delegate to
 * {@link normalizeTitleFontSize} — out-of-range, non-finite, and
 * non-numeric inputs all collapse to `undefined`, fractional inputs
 * round to the nearest 0.5pt (Excel's UI granularity), and a `null`
 * override always drops the inherited size so the writer falls back
 * to Excel's reference 10pt tick-label size.
 *
 * The `<c:txPr>` block sits on every axis flavour per the OOXML
 * schema, so the override applies on every chart family that has
 * axes. The pie / doughnut short-circuit upstream collapses the
 * field on those families since neither has axes.
 */
export function applyLabelFontSizeOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeTitleFontSize(source)
  if (override === null) return undefined
  return normalizeTitleFontSize(override)
}

/**
 * Resolve a `labelBold` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis
 * helpers. Mirrors the chart-level `resolveTitleBold` — non-boolean
 * overrides (typed escapes from an untyped caller) collapse to
 * `undefined`, a `null` override always drops the inherited flag, and
 * a literal `true` / `false` replaces it.
 *
 * The `<c:txPr>` block sits on every axis flavour per the OOXML
 * schema, so the override applies on every chart family that has
 * axes. The pie / doughnut short-circuit upstream collapses the
 * field on those families since neither has axes.
 */
export function applyLabelBoldOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    if (source === true) return true
    if (source === false) return false
    return undefined
  }
  if (override === null) return undefined
  if (override === true) return true
  if (override === false) return false
  return undefined
}

/**
 * Resolve a `labelItalic` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis
 * helpers. Mirrors {@link applyLabelBoldOverride} — non-boolean
 * overrides (typed escapes from an untyped caller) collapse to
 * `undefined`, a `null` override always drops the inherited flag, and
 * a literal `true` / `false` replaces it.
 *
 * The `<c:txPr>` block sits on every axis flavour per the OOXML
 * schema, so the override applies on every chart family that has
 * axes. The pie / doughnut short-circuit upstream collapses the
 * field on those families since neither has axes.
 */
export function applyLabelItalicOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    if (source === true) return true
    if (source === false) return false
    return undefined
  }
  if (override === null) return undefined
  if (override === true) return true
  if (override === false) return false
  return undefined
}

/**
 * Resolve a `labelColor` override using the same `undefined` (inherit)
 * / `null` (drop) / value (replace) grammar as the other axis helpers.
 * Non-string overrides and malformed hex tokens (typed escapes from
 * an untyped caller, wrong length, non-hex characters, alpha-channel
 * forms) collapse to `undefined` via {@link normalizeTitleColor} so
 * the cloned `SheetChart` always carries a value the writer will
 * accept. A `null` override always drops the inherited fill (the
 * writer falls back to the theme text color — no `<a:solidFill>`
 * block on the axis tick-label `<c:txPr>` default-paragraph
 * properties).
 *
 * The `<c:txPr>` block sits on every axis flavour per the OOXML
 * schema, so the override applies on every chart family that has
 * axes. The pie / doughnut short-circuit upstream collapses the
 * field on those families since neither has axes.
 */
export function applyLabelColorOverride(
  source: ChartColor | undefined,
  override: ChartColor | null | undefined,
): ChartColor | undefined {
  if (override === undefined) return normalizeTitleColor(source)
  if (override === null) return undefined
  return normalizeTitleColor(override)
}

/**
 * Resolve a `labelUnderline` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Mirrors {@link applyLabelBoldOverride} /
 * {@link applyLabelItalicOverride} — non-boolean overrides (typed
 * escapes from an untyped caller) collapse to `undefined`, a `null`
 * override always drops the inherited flag, and a literal `true` /
 * `false` replaces it.
 *
 * The `<c:txPr>` block sits on every axis flavour per the OOXML
 * schema, so the override applies on every chart family that has
 * axes. The pie / doughnut short-circuit upstream collapses the
 * field on those families since neither has axes.
 */
export function applyLabelUnderlineOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    if (source === true) return true
    if (source === false) return false
    return undefined
  }
  if (override === null) return undefined
  if (override === true) return true
  if (override === false) return false
  return undefined
}

/**
 * Resolve a `labelStrike` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Mirrors {@link applyLabelBoldOverride} /
 * {@link applyLabelItalicOverride} / {@link applyLabelUnderlineOverride}
 * — non-boolean overrides (typed escapes from an untyped caller)
 * collapse to `undefined`, a `null` override always drops the
 * inherited flag, and a literal `true` / `false` replaces it.
 *
 * The `<c:txPr>` block sits on every axis flavour per the OOXML
 * schema, so the override applies on every chart family that has
 * axes. The pie / doughnut short-circuit upstream collapses the
 * field on those families since neither has axes.
 */
export function applyLabelStrikeOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    if (source === true) return true
    if (source === false) return false
    return undefined
  }
  if (override === null) return undefined
  if (override === true) return true
  if (override === false) return false
  return undefined
}

/**
 * Resolve a `labelFontFamily` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis tick-label typography knobs.
 *
 * Empty / whitespace-only strings and non-string overrides (typed
 * escapes from an untyped caller) collapse to `undefined` via
 * {@link normalizeLabelFontFamily} so the cloned `SheetChart` always
 * carries a value the writer will accept. A `null` override always
 * drops the inherited typeface (the writer falls back to the OOXML
 * default — no `<a:latin>` element, the labels inherit the theme
 * typeface).
 *
 * The `<c:txPr>` block sits on every axis flavour per the OOXML
 * schema, so the override applies on every chart family that has
 * axes. The pie / doughnut short-circuit upstream collapses the
 * field on those families since neither has axes.
 */
export function applyLabelFontFamilyOverride(
  source: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeLabelFontFamily(source)
  if (override === null) return undefined
  return normalizeLabelFontFamily(override)
}

/**
 * Normalize a `labelFontFamily` value for the cloned `SheetChart`.
 * Mirrors the writer's `normalizeAxisLabelFontFamily` — non-empty
 * strings pass through trimmed, every other token (empty /
 * whitespace-only strings, typed escapes from an untyped caller)
 * collapses to `undefined` so the cloned chart drops the field
 * rather than carry a value the writer would silently elide.
 */
export function normalizeLabelFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

/**
 * Resolve an `axisTitleRotation` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. The conversion / clamping rules mirror
 * {@link applyLabelRotationOverride} — out-of-range and non-numeric
 * inputs clamp to the `-90..90` band the writer accepts, the OOXML
 * default `0` collapses to `undefined`, and a `null` override always
 * drops the inherited rotation. The caller is expected to additionally
 * gate the resolved value on the matching axis title's presence so the
 * cloned shape never carries a rotation that the writer would silently
 * elide.
 */
export function applyAxisTitleRotationOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (typeof source !== "number" || !Number.isFinite(source)) return undefined
    return clampLabelRotationDeg(source)
  }
  if (override === null) return undefined
  if (typeof override !== "number" || !Number.isFinite(override)) return undefined
  return clampLabelRotationDeg(override)
}

/**
 * Resolve an `axisTitleFontSize` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. The conversion / clamping rules delegate to
 * {@link normalizeTitleFontSize} — out-of-range, non-finite, and
 * non-numeric inputs all collapse to `undefined`, fractional inputs
 * round to the nearest 0.5pt (Excel's UI granularity), and a `null`
 * override always drops the inherited size.
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a size that the writer would silently elide (the writer
 * scopes the size emission to `<c:title>`, which is omitted when the
 * axis renders no title).
 */
export function applyAxisTitleFontSizeOverride(
  source: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeTitleFontSize(source)
  if (override === null) return undefined
  return normalizeTitleFontSize(override)
}

/**
 * Resolve an `axisTitleBold` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Mirrors the chart-level `resolveTitleBold` —
 * non-boolean overrides (typed escapes from an untyped caller)
 * collapse to `undefined`, a `null` override always drops the
 * inherited flag, and a literal `true` / `false` replaces it.
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a flag the writer would silently elide (the writer scopes
 * the flag emission to `<c:title>`, which is omitted when the axis
 * renders no title).
 */
export function applyAxisTitleBoldOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    if (source === true) return true
    if (source === false) return false
    return undefined
  }
  if (override === null) return undefined
  if (override === true) return true
  if (override === false) return false
  return undefined
}

/**
 * Resolve an `axisTitleItalic` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Non-boolean overrides (typed escape from an untyped
 * caller) collapse to `undefined` via {@link normalizeTitleItalic} so
 * the cloned `SheetChart` always carries a value the writer will
 * accept. A `null` override always drops the inherited flag (the
 * writer falls back to the OOXML default — no `i` attribute,
 * equivalent to non-italic).
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a flag that the writer would silently elide (the writer
 * scopes the flag emission to `<c:title>`, which is omitted when the
 * axis renders no title).
 */
export function applyAxisTitleItalicOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeTitleItalic(source)
  if (override === null) return undefined
  return normalizeTitleItalic(override)
}

/**
 * Resolve an `axisTitleColor` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Non-string overrides and malformed hex tokens (typed
 * escapes from an untyped caller, wrong length, non-hex characters,
 * alpha-channel forms) collapse to `undefined` via
 * {@link normalizeTitleColor} so the cloned `SheetChart` always
 * carries a value the writer will accept. A `null` override always
 * drops the inherited fill (the writer falls back to the theme text
 * color — no `<a:solidFill>` block on the axis title's
 * default-paragraph properties).
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a fill that the writer would silently elide (the writer
 * scopes the fill emission to `<c:title>`, which is omitted when
 * the axis renders no title).
 */
export function applyAxisTitleColorOverride(
  source: ChartColor | undefined,
  override: ChartColor | null | undefined,
): ChartColor | undefined {
  if (override === undefined) return normalizeTitleColor(source)
  if (override === null) return undefined
  return normalizeTitleColor(override)
}

/**
 * Resolve an `axisTitleFillColor` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Non-string overrides and malformed hex tokens (typed
 * escapes from an untyped caller, wrong length, non-hex characters,
 * alpha-channel forms) collapse to `undefined` via
 * {@link normalizeTitleColor} so the cloned `SheetChart` always
 * carries a value the writer will accept. A `null` override always
 * drops the inherited fill (the writer falls back to the theme
 * default fill — no `<c:spPr>` block on the axis title, typically a
 * transparent title background matching Excel's reference shape).
 *
 * Independent of {@link applyAxisTitleColorOverride}: this resolver
 * targets the axis title's background `<c:spPr><a:solidFill>` slot,
 * while the font color targets the inner `<a:defRPr><a:solidFill>`
 * inside `<c:tx><c:rich><a:p><a:pPr>` — the two knobs compose
 * without conflict.
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a fill that the writer would silently elide (the writer
 * scopes the fill emission to `<c:title>`, which is omitted when
 * the axis renders no title).
 */
export function applyAxisTitleFillColorOverride(
  source: ChartColor | undefined,
  override: ChartColor | null | undefined,
): ChartColor | undefined {
  if (override === undefined) return normalizeTitleColor(source)
  if (override === null) return undefined
  return normalizeTitleColor(override)
}

/**
 * Resolve an `axisTitleBorderColor` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Non-string overrides and malformed hex tokens (typed
 * escapes from an untyped caller, wrong length, non-hex characters,
 * alpha-channel forms) collapse to `undefined` via
 * {@link normalizeTitleColor} so the cloned `SheetChart` always
 * carries a value the writer will accept. A `null` override always
 * drops the inherited stroke (the writer falls back to the theme-
 * default auto-stroke — no `<a:ln>` block on the axis title's
 * `<c:spPr>`, typically no visible border matching Excel's reference
 * shape).
 *
 * Independent of {@link applyAxisTitleFillColorOverride}: this
 * resolver targets the axis title's stroke
 * `<c:spPr><a:ln><a:solidFill>` slot, while the fill targets the
 * sibling `<c:spPr><a:solidFill>` slot — the two knobs land on
 * different children of the shared `<c:spPr>` block so a caller can
 * pin both without conflict. Independent of
 * {@link applyAxisTitleColorOverride}: the font color targets the
 * inner `<a:defRPr><a:solidFill>` slot inside `<c:tx><c:rich><a:p>
 * <a:pPr>` — disjoint from both the fill and the stroke. Mirrors
 * {@link applyTitleBorderColorOverride} for axis titles — same
 * accept-with-or-without-`#` hex grammar, distinct host element.
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a stroke that the writer would silently elide (the writer
 * scopes the stroke emission to `<c:title>`, which is omitted when
 * the axis renders no title).
 */
export function applyAxisTitleBorderColorOverride(
  source: ChartColor | undefined,
  override: ChartColor | null | undefined,
): ChartColor | undefined {
  if (override === undefined) return normalizeTitleColor(source)
  if (override === null) return undefined
  return normalizeTitleColor(override)
}

/**
 * Resolve an `axisTitleStrike` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Non-boolean overrides (typed escape from an untyped
 * caller) collapse to `undefined` via {@link normalizeTitleStrike} so
 * the cloned `SheetChart` always carries a value the writer will
 * accept. A `null` override always drops the inherited flag (the
 * writer falls back to the OOXML default — no `strike` attribute,
 * equivalent to no strikethrough).
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a flag that the writer would silently elide (the writer
 * scopes the flag emission to `<c:title>`, which is omitted when the
 * axis renders no title).
 */
export function applyAxisTitleStrikeOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeTitleStrike(source)
  if (override === null) return undefined
  return normalizeTitleStrike(override)
}

/**
 * Resolve an `axisTitleUnderline` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers. Non-boolean overrides (typed escape from an untyped
 * caller) collapse to `undefined` via {@link normalizeTitleUnderline}
 * so the cloned `SheetChart` always carries a value the writer will
 * accept. A `null` override always drops the inherited flag (the
 * writer falls back to the OOXML default — no `u` attribute,
 * equivalent to no underline).
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a flag that the writer would silently elide (the writer
 * scopes the flag emission to `<c:title>`, which is omitted when the
 * axis renders no title).
 */
export function applyAxisTitleUnderlineOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeTitleUnderline(source)
  if (override === null) return undefined
  return normalizeTitleUnderline(override)
}

/**
 * Resolve an `axisTitleFontFamily` override using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis-title typography knobs.
 *
 * Empty / whitespace-only strings and non-string overrides (typed
 * escapes from an untyped caller) collapse to `undefined` via
 * {@link normalizeTitleFontFamily} so the cloned `SheetChart` always
 * carries a value the writer will accept. A `null` override always
 * drops the inherited typeface (the writer falls back to the OOXML
 * default — no `<a:latin>` element, the title inherits the theme
 * typeface).
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a typeface that the writer would silently elide (the
 * writer scopes the element emission to `<c:title>`, which is omitted
 * when the axis renders no title).
 */
export function applyAxisTitleFontFamilyOverride(
  source: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeAxisTitleFontFamilyClone(source)
  if (override === null) return undefined
  return normalizeAxisTitleFontFamilyClone(override)
}

/**
 * Normalize an `axisTitleFontFamily` value for the cloned `SheetChart`.
 * Mirrors the writer's `normalizeAxisTitleFontFamily` — the cloned
 * shape is guaranteed to round-trip through the writer without
 * surprise: non-empty strings pass through trimmed, every other token
 * (empty / whitespace-only strings, typed escapes from an untyped
 * caller) collapses to `undefined` so the cloned chart drops the
 * field rather than carry a value the writer would silently elide.
 */
export function normalizeAxisTitleFontFamilyClone(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  if (trimmed.length === 0) return undefined
  return trimmed
}

/**
 * Resolve an `axisTitleOverlay` override.
 *
 * `undefined` → inherit the source axis's parsed `axisTitleOverlay`.
 * `null`      → drop the inherited value (the writer falls back to
 *               the OOXML `false` default — `<c:overlay val="0"/>`,
 *               the title reserves its own slot adjacent to the axis
 *               with no overlap).
 * `boolean`   → replace.
 *
 * Mirrors the chart-level `titleOverlay` resolver (PR #224) and the
 * other axis-title knobs (`axisTitleRotation` / `axisTitleFontSize` /
 * `axisTitleBold` / `axisTitleItalic` / `axisTitleColor` /
 * `axisTitleStrike` / `axisTitleUnderline`) so the axis-title knobs
 * compose the same way at the call site. Non-boolean overrides (typed
 * escape from an untyped caller) collapse to `undefined` so the
 * cloned `SheetChart` always carries a value the writer will accept.
 *
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never
 * carries a flag that the writer would silently elide (the writer
 * scopes the flag emission to `<c:title>`, which is omitted when the
 * axis renders no title).
 */
export function applyAxisTitleOverlayOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    if (source === true) return true
    if (source === false) return false
    return undefined
  }
  if (override === null) return undefined
  if (override === true) return true
  if (override === false) return false
  return undefined
}

/**
 * Resolve an `axisTitleLayout` override.
 *
 * `undefined` → inherit the source axis's parsed `axisTitleLayout`
 *               (after running through {@link normalizeLegendLayout}
 *               so a malformed source value drops cleanly — the
 *               normalizer is purely shape-based, no host-element
 *               awareness, so it applies identically to legend /
 *               plot-area / title / axis-title layouts).
 * `null`      → drop the inherited layout (the writer falls back to
 *               Excel's auto-layout position — no `<c:layout>` block
 *               on the axis title).
 * `ChartManualLayout` → replace, after running through
 *               {@link normalizeLegendLayout}. Coordinates outside the
 *               `0..1` band collapse on the matching axis so the
 *               cloned `SheetChart` always carries a value the writer
 *               will accept; an override whose every axis dropped
 *               collapses to `undefined` so the cloned shape skips
 *               the entire `<c:layout>` block.
 *
 * The grammar mirrors `resolveLegendLayout` / `resolvePlotAreaLayout`
 * so the manual-layout knobs compose the same way at the call site.
 * The caller is expected to additionally gate the resolved value on
 * the matching axis title's presence so the cloned shape never carries
 * a layout that the writer would silently elide (the writer scopes the
 * `<c:layout>` emission to `<c:title>`, which is omitted when the axis
 * renders no title).
 */
export function resolveAxisTitleLayout(
  sourceValue: ChartManualLayout | undefined,
  override: ChartManualLayout | null | undefined,
): ChartManualLayout | undefined {
  if (override === undefined) return normalizeLegendLayout(sourceValue)
  if (override === null) return undefined
  return normalizeLegendLayout(override)
}

/**
 * Resolve a reverse-axis override using the same `undefined` (inherit) /
 * `null` (drop) / value (replace) grammar as the other axis helpers.
 *
 * Only `true` round-trips meaningfully — `false` is the OOXML default
 * (`orientation="minMax"`) so it collapses to `undefined` to keep the
 * cloned shape minimal. A source carrying `false` (e.g. an over-eager
 * parser that surfaced the default) collapses to `undefined` on
 * inherit; an explicit `false` override likewise drops the field. The
 * writer's per-axis `reverse: false` default already produces a forward
 * orientation, so the dropped state is indistinguishable from a literal
 * `false`.
 */
export function applyReverseOverride(
  source: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return source === true ? true : undefined
  }
  if (override === null) return undefined
  return override === true ? true : undefined
}

/**
 * Resolve the `crosses` / `crossesAt` pair using the same `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar as the other
 * axis helpers — but applied to the XSD choice between `<c:crosses>`
 * and `<c:crossesAt>`. The two fields are resolved independently
 * (each follows the standard inherit / null / replace contract); the
 * writer's normalizer enforces the choice rule downstream by
 * preferring the numeric pin when both are set.
 *
 * The OOXML default `crosses: "autoZero"` collapses to `undefined` so
 * the cloned shape stays minimal. `crossesAt: 0` is preserved (it is
 * a valid pin, distinct from the `"autoZero"` default). Non-finite
 * inputs and unknown semantic tokens drop to `undefined` so they
 * cannot leak into the writer.
 */
export function applyCrossesOverride(
  source: CrossesPairSource,
  override: CrossesPairOverride,
): CrossesPair {
  const out: CrossesPair = {}

  if (override.crosses !== undefined) {
    if (override.crosses !== null) {
      const value = override.crosses
      if (VALID_CROSSES_VALUES.has(value) && value !== "autoZero") {
        out.crosses = value
      }
    }
    // override.crosses === null drops the field entirely.
  } else if (source.crosses !== undefined) {
    if (VALID_CROSSES_VALUES.has(source.crosses) && source.crosses !== "autoZero") {
      out.crosses = source.crosses
    }
  }

  if (override.crossesAt !== undefined) {
    if (
      override.crossesAt !== null &&
      typeof override.crossesAt === "number" &&
      Number.isFinite(override.crossesAt)
    ) {
      out.crossesAt = override.crossesAt
    }
    // override.crossesAt === null drops the field entirely.
  } else if (typeof source.crossesAt === "number" && Number.isFinite(source.crossesAt)) {
    out.crossesAt = source.crossesAt
  }

  return out
}

/**
 * Normalize a {@link ChartAxisDispUnit} shorthand or full
 * {@link ChartAxisDispUnits} object into a stable shape the resolver
 * can hand back to the writer-side `SheetChart.axes.{x,y}.dispUnits`
 * field. Unknown / typo'd tokens collapse to `undefined` so they cannot
 * leak past the clone layer.
 *
 * Both `unit` (built-in preset) and `custUnit` (custom numeric divisor)
 * pass through when valid. The OOXML schema's `xsd:choice` between
 * `<c:builtInUnit>` and `<c:custUnit>` is enforced at emit time by the
 * writer (which prefers `custUnit` when both are pinned); the
 * normalizer keeps both fields so the clone layer can append a
 * `custUnit` override to a source whose parsed value pinned `unit`
 * without manually pruning the inherited preset. Returns `undefined`
 * when neither field resolves to a valid value — a stray
 * `ChartAxisDispUnits` with no usable child has nothing to emit.
 */
export function normalizeDispUnits(
  value: ChartAxisDispUnits | ChartAxisDispUnit | undefined,
): ChartAxisDispUnits | undefined {
  if (value === undefined) return undefined
  if (typeof value === "string") {
    return VALID_DISP_UNIT_VALUES.has(value as ChartAxisDispUnit)
      ? { unit: value as ChartAxisDispUnit }
      : undefined
  }
  if (typeof value !== "object" || value === null) return undefined
  const out: ChartAxisDispUnits = {}
  const unit = value.unit
  if (typeof unit === "string" && VALID_DISP_UNIT_VALUES.has(unit as ChartAxisDispUnit)) {
    out.unit = unit as ChartAxisDispUnit
  }
  const custUnit = value.custUnit
  if (typeof custUnit === "number" && Number.isFinite(custUnit) && custUnit > 0) {
    out.custUnit = custUnit
  }
  if (out.unit === undefined && out.custUnit === undefined) return undefined
  if (value.showLabel === true) out.showLabel = true
  if (typeof value.customLabel === "string") {
    const trimmed = value.customLabel.trim()
    if (trimmed.length > 0) out.customLabel = trimmed
  }
  return out
}

/**
 * Resolve a `dispUnits` override using the standard `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar. Both inputs go
 * through {@link normalizeDispUnits} so unknown tokens collapse to
 * `undefined` rather than fabricate a value the writer would never
 * emit. The reader and writer mirror this normalizer so a parsed
 * source value slots straight back into a clone target without
 * transformation.
 */
export function applyDispUnitsOverride(
  source: ChartAxisDispUnits | undefined,
  override: ChartAxisDispUnits | ChartAxisDispUnit | null | undefined,
): ChartAxisDispUnits | undefined {
  if (override === undefined) return normalizeDispUnits(source)
  if (override === null) return undefined
  return normalizeDispUnits(override)
}

/**
 * Resolve a `crossBetween` override using the standard `undefined`
 * (inherit) / `null` (drop) / value (replace) grammar. Unknown / typo'd
 * tokens collapse to `undefined` rather than fabricate a value the
 * writer would never emit — the writer's per-family default
 * (`"between"` on bar / column / line / area Y axes; `"midCat"` on
 * scatter axes) takes over instead. The reader and writer mirror this
 * normalizer so a parsed source value slots straight back into a clone
 * target without transformation.
 */
export function applyCrossBetweenOverride(
  source: ChartAxisCrossBetween | undefined,
  override: ChartAxisCrossBetween | null | undefined,
): ChartAxisCrossBetween | undefined {
  if (override === undefined) {
    if (source === undefined) return undefined
    return VALID_CROSS_BETWEEN_VALUES.has(source) ? source : undefined
  }
  if (override === null) return undefined
  return VALID_CROSS_BETWEEN_VALUES.has(override) ? override : undefined
}
