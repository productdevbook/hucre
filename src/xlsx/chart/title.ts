// ── Chart Title ────────────────────────────────────────────────────
// Per-host module for `<c:chart><c:title>` (CT_Title, ECMA-376 Part 1,
// §21.2.2.210). Holds the reader / writer helpers — every `parse*` /
// `build*` / `resolve*` / `normalize*` function for the chart-level
// title block, including its `<c:tx><c:rich>` typography (font size /
// bold / italic / color / strike / underline / family), `<c:overlay>`
// flag, `<c:layout>` (manual layout), and `<c:spPr>` (fill / border)
// slots.
//
// The clone-side `resolveTitle*` overrides (3-arg `(source, override)`
// resolvers) live in `chart-clone.ts` because their signatures are
// shape-incompatible with the writer-side `resolveTitle*` resolvers
// (1-arg `(chart: SheetChart)`); they do delegate to the shared
// `normalizeTitle*` exports here so the per-field clamp / drop grammar
// stays in one place.
//
// JSDoc preserved verbatim from the call sites where each helper was
// originally inlined; the host-specific scope notes ("only meaningful
// when the chart actually emits a title", "the caller is expected to
// gate the call on showTitle && chart.title") stay attached because the
// per-host commentary remains meaningful at the call site even after
// relocation.

import type { ChartBorderDash, ChartManualLayout, SheetChart } from "../../_types";
import type { XmlElement } from "../../xml/parser";
import { xmlElement, xmlEscape, xmlSelfClose } from "../../xml/writer";
import {
  EMU_PER_PT,
  clampStrokeWidthPt,
  normalizeBorderDash,
  normalizeRgbHex,
  normalizeRgbHex as normalizeRgbHexShared,
  parseBorderDashFromSpPr,
  parseBorderWidthFromSpPr,
  parseSpPrBorderColor,
  parseSpPrFill,
} from "./shape";
import {
  type ResolvedManualLayout,
  buildManualLayout,
  normalizeChartManualLayout,
  normalizeManualLayout,
  parseManualLayout,
} from "./layout";
import { childElements, collectTextRuns, elementText, findChild, parseBoolAttr } from "./util";
import {
  FONT_SIZE_MAX_PT,
  FONT_SIZE_MIN_PT,
  FONT_SZ_PER_POINT,
  ROTATION_MAX_DEG,
  ROTATION_MIN_DEG,
  TXPR_ROT_PER_DEGREE,
} from "./text";

// ── Constants (chart-title scope) ──────────────────────────────────

/**
 * OOXML's `<a:bodyPr rot="N"/>` attribute is in 60000ths of a degree —
 * the writer holds `titleRotation` in whole degrees and converts at
 * emit time. Excel's UI exposes the `-90..90` band; out-of-band values
 * clamp to the nearest endpoint so a corrupt template cannot leak
 * through to the writer either.
 *
 * Aliased onto the shared {@link TXPR_ROT_PER_DEGREE} /
 * {@link ROTATION_MIN_DEG} / {@link ROTATION_MAX_DEG} constants so
 * every typography host (chart-title, axis-title, tick-label, legend,
 * data-label, data-table) shares the same conversion factor.
 */
const TITLE_ROT_PER_DEGREE = TXPR_ROT_PER_DEGREE;
const TITLE_ROTATION_MIN_DEG = ROTATION_MIN_DEG;
const TITLE_ROTATION_MAX_DEG = ROTATION_MAX_DEG;

/**
 * OOXML's `<a:defRPr sz="N"/>` / `<a:rPr sz="N"/>` attribute is in
 * 100ths of a point — the writer holds {@link SheetChart.titleFontSize}
 * in points and converts at emit time. The OOXML `ST_TextFontSize`
 * schema restricts `sz` to the inclusive `100..400000` band; the
 * writer's clamp uses the same range converted to points (`1..400`pt),
 * so any out-of-range value drops at emit time rather than surface a
 * token Excel would reject.
 */
const TITLE_FONT_SZ_PER_POINT = FONT_SZ_PER_POINT;
const TITLE_FONT_SIZE_MIN_PT = FONT_SIZE_MIN_PT;
const TITLE_FONT_SIZE_MAX_PT = FONT_SIZE_MAX_PT;

/**
 * Application-default `sz` value for the chart title's `<a:defRPr>` /
 * `<a:rPr>` slots — Excel renders the title at 14pt (`sz="1400"`)
 * unless the user pins a custom size. Absence of
 * {@link SheetChart.titleFontSize} resolves to this default so a fresh
 * chart matches Excel's reference serialization byte-for-byte.
 */
const TITLE_DEFAULT_FONT_SIZE_SZ = 1400;

// ── Reader ────────────────────────────────────────────────────────

/**
 * Pull `<c:title><c:layout><c:manualLayout>` off the chart. Reflects
 * Excel's "Format Chart Title -> Title Options -> Position -> Custom"
 * knob — the `(x, y)` anchor and `(w, h)` size of the title block as
 * fractions of the chart frame in the `0..1` band.
 *
 * The OOXML schema (`CT_ManualLayout`, ECMA-376 Part 1, §21.2.2.115)
 * exposes optional `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` children whose
 * `val` attributes carry an `xsd:double`. The reader admits the
 * coordinate only when `val` parses to a finite number in the `0..1`
 * band; out-of-range / non-finite / non-numeric tokens drop to
 * `undefined` on the matching axis so absence and a malformed token
 * round-trip identically through {@link cloneChart}.
 *
 * Both `<c:xMode val="edge"/>` (absolute fraction of the chart frame)
 * and `<c:xMode val="factor"/>` (delta from auto-layout) are accepted
 * — the reader surfaces the same `ChartManualLayout` shape regardless,
 * since the writer always normalizes to `"edge"` on emit (Excel itself
 * emits the absolute form when the user drags the title to a custom
 * position).
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` /
 * `<c:layout>` / `<c:manualLayout>` chain at any link, or when every
 * coordinate dropped on normalization — the field is omitted entirely
 * on a clean parse so absence and an empty layout round-trip identically
 * through the writer. Mirrors {@link parseLegendLayout} so a parsed
 * value flows through the same `ChartManualLayout` shape regardless of
 * which manual-layout slot the source chart pinned.
 */
export function parseTitleLayout(chartEl: XmlElement): ChartManualLayout | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  return parseManualLayout(title);
}

/**
 * Pull `<c:title><c:overlay val=".."/></c:title>` off the chart. The
 * OOXML default `false` (the title reserves its own slot above the plot
 * area, no overlap) collapses to `undefined` so absence and
 * `<c:overlay val="0"/>` round-trip identically through
 * {@link cloneChart} — only an explicit `<c:overlay val="1"/>` surfaces
 * `true`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no overlay slot to surface in that case. The element is a
 * sibling of `<c:tx>` inside `<c:title>` per the CT_Title schema, so the
 * lookup is scoped to direct title children.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"`
 * / `"false"`); unknown values and missing `val` attributes drop to
 * `undefined` rather than fabricate a flag Excel would not emit.
 */
export function parseTitleOverlay(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const overlay = findChild(title, "overlay");
  if (!overlay) return undefined;
  const raw = overlay.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      // OOXML default — collapse to undefined for symmetry with the
      // writer's `titleOverlay` field.
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx>
 * </c:title>` off the chart. Returns the rotation in whole degrees
 * (range `-90..90`).
 *
 * The OOXML default `0` (and absence of the `<a:bodyPr>` element /
 * `rot` attribute) all collapse to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart}.
 * Non-integer / non-numeric / out-of-range values clamp to the nearest
 * endpoint of the `-90..90` band Excel's UI exposes; non-finite
 * (`NaN`, `Infinity`) inputs drop to `undefined`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no rotation slot to surface in that case. The
 * `<a:bodyPr>` lives inside `<c:tx><c:rich>` per the CT_Title schema
 * (the rich-text body's body-properties); the lookup is scoped to that
 * path so a stray `<a:bodyPr>` elsewhere in the chart cannot leak in.
 */
export function parseTitleRotation(chartEl: XmlElement): number | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  const bodyPr = findChild(rich, "bodyPr");
  if (!bodyPr) return undefined;
  const raw = bodyPr.attrs.rot;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 60000ths of a degree to whole degrees.
  const degrees = Math.round(parsed / TITLE_ROT_PER_DEGREE);
  if (degrees === 0) return undefined;
  if (degrees < TITLE_ROTATION_MIN_DEG) return TITLE_ROTATION_MIN_DEG;
  if (degrees > TITLE_ROTATION_MAX_DEG) return TITLE_ROTATION_MAX_DEG;
  return degrees;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off the chart. Returns the font
 * size in points (range `1..400`).
 *
 * The OOXML `sz` attribute is in 100ths of a point — the reader
 * converts to points and rounds to the nearest 0.5pt (Excel's UI
 * exposes the same 0.5pt granularity). Absence of the element /
 * attribute and out-of-range / non-numeric / non-finite values all
 * collapse to `undefined` so a fresh chart and a chart that pinned an
 * out-of-range size both round-trip to the writer's "skip the size
 * attribute" path.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the size from in that case —
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the
 * default-paragraph properties on the rich-text body's first
 * paragraph); the lookup is scoped to that path so a stray
 * `<a:defRPr>` elsewhere in the chart (e.g. on an axis title or a
 * data-labels block) cannot leak in.
 */
export function parseTitleFontSize(chartEl: XmlElement): number | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph font size. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.sz;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return undefined;
  // Convert from 100ths of a point to points, rounding to the nearest
  // 0.5pt to match the granularity Excel's UI exposes. `Math.round`
  // on `2 * (parsed / 100)` and dividing by 2 gives a clean half-step
  // band that mirrors the writer's emit-time normalization.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off the chart. Returns the bold
 * flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `b="0"` round-trip identically — only an explicit `b="1"` surfaces
 * `true`. Unknown / malformed `b` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the flag from in that case —
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the
 * default-paragraph properties on the rich-text body's first
 * paragraph); the lookup is scoped to that path so a stray
 * `<a:defRPr>` elsewhere in the chart (e.g. on an axis title or a
 * data-labels block) cannot leak in.
 */
export function parseTitleBold(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph bold flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.b);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `b="0"` round-trip identically through the writer — only an
  // explicit `b="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` off the chart. Returns the italic
 * flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `i="0"` round-trip identically — only an explicit `i="1"` surfaces
 * `true`. Unknown / malformed `i` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the flag from in that case —
 * or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the
 * default-paragraph properties on the rich-text body's first
 * paragraph); the lookup is scoped to that path so a stray
 * `<a:defRPr>` elsewhere in the chart (e.g. on an axis title or a
 * data-labels block) cannot leak in.
 */
export function parseTitleItalic(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph italic flag. The reader walks the canonical chain
  // and bails on the first missing link so a malformed `<c:rich>`
  // surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const parsed = parseBoolAttr(defRPr.attrs.i);
  // The OOXML default `false` collapses to `undefined` so absence and
  // `i="0"` round-trip identically through the writer — only an
  // explicit `i="1"` surfaces `true`.
  if (parsed === true) return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` off the chart. Returns the title's
 * sRGB font color as a 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr val=".."/>` is the literal sRGB triple Excel
 * lands on the title's default-paragraph properties when the user
 * picks a custom font color. Theme references (`<a:schemeClr>`),
 * `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>` all collapse to
 * `undefined` — only the literal RGB triple round-trips losslessly
 * through {@link writeChart}. Malformed `val` tokens (wrong length,
 * non-hex characters) likewise drop to `undefined` rather than
 * fabricate a value the writer would round-trip into a malformed
 * `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>`
 * element — there is no `<a:p>` slot to surface the fill from in
 * that case — or when the title is a `<c:strRef>` (formula
 * reference) with no `<c:rich>` body. The `<a:solidFill>` lives
 * inside `<c:tx><c:rich><a:p><a:pPr><a:defRPr>` per the CT_Title
 * schema (the default-paragraph properties on the rich-text body's
 * first paragraph); the lookup is scoped to that path so a stray
 * `<a:solidFill>` elsewhere in the chart (e.g. on an axis title or
 * a data-labels block) cannot leak in.
 */
export function parseTitleColor(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr>` is the OOXML path
  // Excel writes for the default-paragraph font color. The reader
  // walks the canonical chain and bails on the first missing link so
  // a malformed `<c:rich>` surfaces as absence rather than a
  // fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  const raw = srgbClr.attrs.val;
  return normalizeRgbHex(raw);
}

/**
 * Pull the chart-title background fill color off the canonical
 * `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:title>` chain Excel writes when the
 * user pins "Format Chart Title -> Fill -> Solid fill -> Color".
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
 * The lookup is scoped to direct children of `<c:title>` so a stray
 * `<c:spPr>` elsewhere in the chart (e.g. on the plot area, a series,
 * or the legend) cannot leak in. Returns `undefined` whenever the
 * chart omits the `<c:title>` element or the
 * `<c:spPr><a:solidFill><a:srgbClr>` chain is malformed at any link.
 * Mirrors the legend / plot-area fill readers exactly so a parsed
 * value slots straight back into the writer's emit path.
 *
 * Independent of {@link parseTitleColor}: the fill lives on
 * `<c:title><c:spPr>`, the font color lives on
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>` — the
 * two readers walk disjoint paths so a caller can pin both knobs
 * without conflict. Unlike {@link parseTitleColor}, the lookup is on
 * `<c:title>` directly rather than gated on `<c:rich>` so a title
 * authored as a `<c:strRef>` formula reference can still surface its
 * background fill.
 */
export function parseTitleFillColor(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  return parseSpPrFill(title);
}

/**
 * Pull `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:title>` off the chart-level
 * `<c:title>` block. Returns the title border (line) stroke color as
 * a 6-character uppercase hex string the writer can round-trip via
 * {@link SheetChart.titleBorderColor}.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the line's
 * solid fill choice (`CT_LineProperties`, §20.1.2.3.24) which itself
 * sits inside `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13). The
 * `<c:spPr>` slot lives between `<c:overlay>` and `<c:txPr>` /
 * `<c:extLst>` per CT_Title (§21.2.2.210); `<a:ln>` follows the
 * optional `<a:solidFill>` (fill) child inside `<c:spPr>`.
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
 * The lookup is scoped to direct children of `<c:title>` so a stray
 * `<c:spPr>` elsewhere in the chart (e.g. on the plot area, a
 * series, or the legend) cannot leak in. Returns `undefined`
 * whenever the chart omits the `<c:title>` element or the
 * `<c:spPr><a:ln><a:solidFill><a:srgbClr>` chain is malformed at
 * any link. Mirrors {@link parsePlotAreaBorderColor} — same
 * `<a:ln>` chain on a different host element. Independent of
 * {@link parseTitleFillColor}: the two readers walk disjoint
 * children of the shared `<c:spPr>` block (`<a:solidFill>` for the
 * fill, `<a:ln>` for the stroke) so a caller can pin both knobs
 * without conflict. Unlike {@link parseTitleColor}, the lookup is
 * on `<c:title>` directly rather than gated on `<c:rich>` so a
 * title authored as a `<c:strRef>` formula reference can still
 * surface its border color.
 */
export function parseTitleBorderColor(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  return parseSpPrBorderColor(title);
}

/**
 * Pull the `w` attribute off `<c:title><c:spPr><a:ln w="EMU"/>` and
 * return the stroke width in points after clamping to the
 * `0.25..13.5` pt band Excel's UI exposes. The OOXML `w` attribute
 * carries the stroke width in English Metric Units (1 pt = 12 700 EMU)
 * per `CT_LineProperties` (ECMA-376 Part 1, §20.1.2.3.24); the reader
 * snaps the result to the 0.25 pt grid so a parsed-then-written width
 * does not drift across round-trips (Excel rounds in the UI anyway).
 *
 * Returns `undefined` when the chart omits `<c:title>`, when the
 * title has no `<c:spPr><a:ln w=..>` slot, when the attribute is
 * missing, when the value cannot be parsed as a finite positive
 * number, or when it parses to zero (Excel's "no border" marker — the
 * writer-side knob does not model that state). Mirrors the writer-side
 * {@link SheetChart.titleBorderWidth} so a parsed value slots
 * straight into {@link cloneChart} without conversion.
 *
 * The lookup is scoped to direct children of `<c:title>` so a stray
 * `<a:ln w=..>` elsewhere (on a series stroke, on an axis line, on the
 * plot-area / legend border) cannot leak in. Mirrors {@link parseTitleBorderColor} —
 * same `<c:spPr>` host on the same `<c:title>` parent — but lands on
 * the `w` attribute rather than the `<a:solidFill><a:srgbClr>` color
 * child.
 */
export function parseTitleBorderWidth(chartEl: XmlElement): number | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  return parseBorderWidthFromSpPr(title);
}

/**
 * Pull the `val` attribute off `<c:title><c:spPr><a:ln><a:prstDash
 * val=".."/></a:ln></c:spPr></c:title>` and return the recognized
 * {@link ChartBorderDash} value. Returns `undefined` when the chain
 * is missing at any link, when the attribute is absent / unrecognized,
 * or when it matches the OOXML default `"solid"` (so absence and the
 * default round-trip identically through {@link cloneChart}).
 *
 * Delegates to {@link parseBorderDashFromSpPr} so the accept-or-drop
 * grammar matches every chart-frame border-dash slot the reader
 * surfaces.
 */
export function parseTitleBorderDash(chartEl: XmlElement): ChartBorderDash | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  return parseBorderDashFromSpPr(title);
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off the chart. Returns
 * the strikethrough flag.
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
 * Returns `undefined` whenever the chart omits the `<c:title>`
 * element — there is no `<a:p>` slot to surface the flag from in
 * that case — or when the title is a `<c:strRef>` (formula
 * reference) with no `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the default-
 * paragraph properties on the rich-text body's first paragraph); the
 * lookup is scoped to that path so a stray `<a:defRPr>` elsewhere in
 * the chart (e.g. on an axis title or a data-labels block) cannot
 * leak in.
 */
export function parseTitleStrike(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph strikethrough flag. The reader walks the
  // canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  // Only the UI-default `"sngStrike"` surfaces as `true`. The OOXML
  // application default `"noStrike"` and the non-UI `"dblStrike"` both
  // collapse to `undefined` so absence and the OOXML default round-trip
  // identically through the writer; the writer emits only `"sngStrike"`,
  // so reporting `"dblStrike"` here would silently downgrade the choice
  // on round-trip.
  if (raw === "sngStrike") return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` off the chart. Returns
 * the underline flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enum on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7) with
 * eighteen values; Excel's UI exposes only `"sng"` (single line —
 * the default underline checkbox) and `"dbl"` (double line). The
 * reader surfaces only the UI-default `"sng"` as `true`; `"none"`
 * (the OOXML application default), absence, the non-UI `"dbl"`
 * variant, and the sixteen exotic tokens (`"words"`, `"heavy"`,
 * `"dotted"`, `"dottedHeavy"`, `"dash"`, `"dashHeavy"`, `"dashLong"`,
 * `"dashLongHeavy"`, `"dotDash"`, `"dotDashHeavy"`, `"dotDotDash"`,
 * `"dotDotDashHeavy"`, `"wavy"`, `"wavyHeavy"`, `"wavyDbl"`) all
 * collapse to `undefined` — the writer emits only `"sng"`, so
 * reporting any non-single underline as `true` would silently
 * downgrade the choice to a single line on round-trip. Unknown /
 * malformed `u` tokens likewise drop to `undefined`.
 *
 * Returns `undefined` whenever the chart omits the `<c:title>`
 * element — there is no `<a:p>` slot to surface the flag from in
 * that case — or when the title is a `<c:strRef>` (formula
 * reference) with no `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the default-
 * paragraph properties on the rich-text body's first paragraph); the
 * lookup is scoped to that path so a stray `<a:defRPr>` elsewhere in
 * the chart (e.g. on an axis title or a data-labels block) cannot
 * leak in.
 */
export function parseTitleUnderline(chartEl: XmlElement): boolean | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr>` is the OOXML path Excel writes for the
  // default-paragraph underline flag. The reader walks the canonical
  // chain and bails on the first missing link so a malformed
  // `<c:rich>` surfaces as absence rather than a fabricated value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  // Only the UI-default `"sng"` surfaces as `true`. The OOXML
  // application default `"none"`, the non-UI `"dbl"` variant, and
  // every exotic token (`"words"`, `"heavy"`, `"dotted"`, etc.) all
  // collapse to `undefined` so absence and the OOXML default
  // round-trip identically through the writer; the writer emits only
  // `"sng"`, so reporting a non-single underline here would silently
  // downgrade the choice on round-trip.
  if (raw === "sng") return true;
  return undefined;
}

/**
 * Pull `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>`
 * off the chart. Returns the typeface string the title was authored
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
 * Returns `undefined` whenever the chart omits the `<c:title>` element
 * — there is no `<a:p>` slot to surface the typeface from in that case
 * — or when the title is a `<c:strRef>` (formula reference) with no
 * `<c:rich>` body. The `<a:defRPr>` lives inside
 * `<c:tx><c:rich><a:p><a:pPr>` per the CT_Title schema (the default-
 * paragraph properties on the rich-text body's first paragraph); the
 * lookup is scoped to that path so a stray `<a:latin>` elsewhere in
 * the chart (e.g. on an axis title or a data-labels block) cannot
 * leak in.
 */
export function parseTitleFontFamily(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  const rich = findChild(tx, "rich");
  if (!rich) return undefined;
  // `<a:p><a:pPr><a:defRPr><a:latin>` is the OOXML path Excel writes
  // for the default-paragraph typeface. The reader walks the
  // canonical chain and bails on the first missing link so a
  // malformed `<c:rich>` surfaces as absence rather than a fabricated
  // value.
  const p = findChild(rich, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const latin = findChild(defRPr, "latin");
  if (!latin) return undefined;
  const raw = latin.attrs.typeface;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Read `<c:title>` text. The title may be a rich-text run tree or a
 * formula reference; we only surface plain text runs joined together.
 */
export function parseTitle(chartEl: XmlElement): string | undefined {
  const title = findChild(chartEl, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
  // tx can hold either <c:rich> (literal text) or <c:strRef> (formula).
  const rich = findChild(tx, "rich");
  if (rich) {
    const parts: string[] = [];
    collectTextRuns(rich, parts);
    const joined = parts.join("").trim();
    return joined.length > 0 ? joined : undefined;
  }
  const strRef = findChild(tx, "strRef");
  if (strRef) {
    const cache = findChild(strRef, "strCache");
    if (cache) {
      for (const pt of childElements(cache)) {
        if (pt.local !== "pt") continue;
        const v = findChild(pt, "v");
        if (v) {
          const text = elementText(v).trim();
          if (text.length > 0) return text;
        }
      }
    }
  }
  return undefined;
}

// ── Writer ────────────────────────────────────────────────────────

export function buildTitle(
  title: string,
  overlay: boolean,
  rotationDeg: number | undefined,
  fontSizePt: number | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
  rgbHex: string | undefined,
  strike: boolean | undefined,
  underline: boolean | undefined,
  fontFamily: string | undefined,
  layout: ResolvedManualLayout | undefined,
  fillRgbHex: string | undefined,
  borderRgbHex: string | undefined,
  borderWidthPt: number | undefined,
  borderDash: ChartBorderDash | undefined,
): string {
  // OOXML's `<a:bodyPr rot="N"/>` attribute is in 60000ths of a degree.
  // The writer holds `titleRotation` in whole degrees and converts at
  // emit time. Absence (`undefined`) collapses to the OOXML default
  // `0` so a fresh chart matches Excel's reference serialization
  // byte-for-byte.
  const rot = rotationDeg === undefined ? 0 : rotationDeg * TITLE_ROT_PER_DEGREE;
  // OOXML's `<a:defRPr sz="N"/>` / `<a:rPr sz="N"/>` attribute is in
  // 100ths of a point. The writer holds `titleFontSize` in points and
  // converts at emit time. Absence (`undefined`) collapses to the
  // application-default `1400` (14pt) so a fresh chart matches Excel's
  // reference serialization byte-for-byte. The size lands on both the
  // default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>` so
  // a re-parse picks the value up off either canonical slot.
  const sz =
    fontSizePt === undefined ? TITLE_DEFAULT_FONT_SIZE_SZ : fontSizePt * TITLE_FONT_SZ_PER_POINT;
  // OOXML's `<a:defRPr b=".."/>` / `<a:rPr b=".."/>` attribute is the
  // `xsd:boolean` bold flag on `CT_TextCharacterProperties`. The writer
  // holds `titleBold` as a boolean and emits `1` / `0` at the canonical
  // slots. Absence (`undefined`) collapses to the OOXML default `0`
  // (non-bold) so a fresh chart matches Excel's reference serialization
  // byte-for-byte. Like the size, the flag lands on both the
  // default-paragraph `<a:defRPr>` and the literal run's `<a:rPr>` so
  // a re-parse picks the value up off either canonical slot — Excel
  // keeps the two attributes in sync.
  const b = bold ? 1 : 0;
  // OOXML's `<a:defRPr i=".."/>` / `<a:rPr i=".."/>` attribute is the
  // `xsd:boolean` italic flag on `CT_TextCharacterProperties`. Mirrors
  // the bold pattern: `titleItalic` lands on both the default-paragraph
  // `<a:defRPr>` and the literal run's `<a:rPr>` so a re-parse picks
  // the value up off either canonical slot — Excel keeps the two
  // attributes in sync. Absence (`undefined`) and explicit `false` both
  // collapse to omitting the attribute so a fresh chart matches Excel's
  // reference serialization byte-for-byte (Excel itself omits `i` when
  // the title is non-italic — only the bold flag is always emitted).
  const i = italic === true ? 1 : undefined;
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
  const strikeAttr = strike === true ? "sngStrike" : undefined;
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
  const underlineAttr = underline === true ? "sng" : undefined;
  // OOXML's `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:defRPr>` carries the title's font color. The
  // writer holds `titleColor` as a 6-character uppercase hex string
  // and lands the `<a:solidFill>` block on both the default-paragraph
  // `<a:defRPr>` and the literal run's `<a:rPr>` so a re-parse picks
  // the value up off either canonical slot. Absence (`undefined`)
  // collapses to omitting the entire `<a:solidFill>` block so the
  // title inherits the theme text color (Excel's reference behavior
  // for a fresh chart title that has not had a custom color picked).
  const solidFillChild = rgbHex
    ? xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: rgbHex })])
    : undefined;
  // OOXML's `<a:defRPr><a:latin typeface=".."/></a:defRPr>` carries the
  // title's font family. The writer holds `titleFontFamily` as a non-
  // empty string and lands the `<a:latin>` element on both the default-
  // paragraph `<a:defRPr>` and the literal run's `<a:rPr>` so a re-
  // parse picks the typeface up off either canonical slot. Absence
  // (`undefined`) collapses to omitting the entire `<a:latin>` element
  // so the title inherits the theme typeface (Excel's reference
  // behavior for a fresh chart title that has not had a custom font
  // picked). The `<a:latin>` element follows `<a:solidFill>` per the
  // CT_TextCharacterProperties child sequence (ECMA-376 Part 1,
  // §21.1.2.3.7) so a fresh chart with both color and family matches
  // Excel's reference serialization byte-for-byte.
  const latinChild = fontFamily ? xmlSelfClose("a:latin", { typeface: fontFamily }) : undefined;
  // When a fill color or a typeface is set the `<a:defRPr>` /
  // `<a:rPr>` slots expand from self-closing to wrapping the children;
  // otherwise the writer keeps the existing self-closing form so a
  // fresh chart with no custom color or font matches Excel's reference
  // serialization byte-for-byte. Children are emitted in CT_TextChar
  // acterProperties' canonical schema order: solidFill first, then
  // latin.
  const rPrChildren: string[] = [];
  if (solidFillChild) rPrChildren.push(solidFillChild);
  if (latinChild) rPrChildren.push(latinChild);
  const defRPr =
    rPrChildren.length > 0
      ? xmlElement("a:defRPr", { sz, b, i, u: underlineAttr, strike: strikeAttr }, rPrChildren)
      : xmlSelfClose("a:defRPr", { sz, b, i, u: underlineAttr, strike: strikeAttr });
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
        });
  // CT_Title (ECMA-376 Part 1, §21.2.2.210) places the optional
  // `<c:layout>` between `<c:tx>` and `<c:overlay>`. The writer skips
  // emission entirely when the caller pinned no coordinates so a fresh
  // chart matches Excel's reference serialization byte-for-byte (Excel
  // itself omits the block when the title renders at the auto-layout
  // position above the plot area). Each axis is independently optional
  // so the helper drops `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` slots
  // whose value did not survive normalization.
  const layoutXml = buildManualLayout(layout);
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
          xmlElement("a:r", undefined, [rPr, xmlElement("a:t", undefined, xmlEscape(title))]),
        ]),
      ]),
    ]),
  ];
  if (layoutXml !== undefined) {
    titleChildren.push(layoutXml);
  }
  titleChildren.push(xmlSelfClose("c:overlay", { val: overlay ? 1 : 0 }));
  // CT_Title (ECMA-376 Part 1, §21.2.2.210) places the optional
  // `<c:spPr>` between `<c:overlay>` and `<c:txPr>` / `<c:extLst>`.
  // The writer skips emission entirely when the caller did not pin a
  // fill or border color so a fresh chart matches Excel's reference
  // serialization byte-for-byte — Excel itself omits the block
  // whenever the title renders at the theme defaults (typically a
  // transparent title background with no visible border, no
  // `<c:spPr>` block). Authors `<a:solidFill>` for the fill and
  // `<a:ln>` for the stroke in CT_ShapeProperties schema order;
  // other CT_ShapeProperties children (effects, gradient / pattern /
  // picture fills, line dash / width / compound styles) are not
  // modelled at this layer. Distinct from the `<a:defRPr><a:solidFill>`
  // font-color slot inside `<c:tx><c:rich>` that
  // {@link SheetChart.titleColor} pins — the typography knobs target
  // different children of `<c:title>` so a caller can pin both
  // without conflict.
  const titleSpPrXml = buildTitleSpPr(fillRgbHex, borderRgbHex, borderWidthPt, borderDash);
  if (titleSpPrXml !== undefined) {
    titleChildren.push(titleSpPrXml);
  }
  return xmlElement("c:title", undefined, titleChildren);
}

/**
 * Build the `<c:spPr>` element on `<c:title>` that carries the
 * title's background fill ({@link SheetChart.titleFillColor}),
 * border-stroke color ({@link SheetChart.titleBorderColor}), and
 * border width ({@link SheetChart.titleBorderWidth}). Returns
 * `undefined` when no knob is pinned so the caller can elide the
 * entire block — Excel's reference serialization omits `<c:spPr>`
 * from `<c:title>` whenever the title renders at the theme default
 * fill / stroke (typically a transparent title background with no
 * visible border).
 *
 * When at least one knob lands on the wire, the children are emitted
 * in `CT_ShapeProperties` (ECMA-376 Part 1, §20.1.2.3.13) schema
 * order: `<a:solidFill>` (fill) first, then `<a:ln>` (line / stroke).
 * The fill block has the form `<a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill>`; the stroke block has the form `<a:ln w="EMU">
 * <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill></a:ln>`. The
 * `val` attribute holds the canonical 6-character uppercase hex form
 * (the writer normalizes the inputs ahead of this call so malformed
 * source values never reach emit). The width attribute lands on
 * `<a:ln>` (EMU; 1 pt = 12 700 EMU) authored together with the
 * border-color child so a stroke-only or color-only title still emits
 * a single `<a:ln>` block.
 *
 * Mirrors the plot-area / legend `<c:spPr>` slots so a single hex
 * string threads cleanly through every fill / stroke knob the writer
 * authors.
 */
export function buildTitleSpPr(
  fillRgbHex: string | undefined,
  borderRgbHex: string | undefined,
  borderWidthPt: number | undefined,
  borderDash: ChartBorderDash | undefined,
): string | undefined {
  if (
    fillRgbHex === undefined &&
    borderRgbHex === undefined &&
    borderWidthPt === undefined &&
    borderDash === undefined
  ) {
    return undefined;
  }
  const children: string[] = [];
  if (fillRgbHex !== undefined) {
    children.push(
      xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillRgbHex })]),
    );
  }
  if (borderRgbHex !== undefined || borderWidthPt !== undefined || borderDash !== undefined) {
    const lnAttrs: Record<string, string | number> = {};
    if (borderWidthPt !== undefined) {
      // OOXML stores stroke width in EMU (1 pt = 12 700 EMU). Round to
      // the nearest integer because the schema types `w` as `xsd:int`.
      lnAttrs.w = Math.round(borderWidthPt * EMU_PER_PT);
    }
    const lnChildren: string[] = [];
    if (borderRgbHex !== undefined) {
      lnChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: borderRgbHex })]),
      );
    }
    // `<a:prstDash>` follows `<a:solidFill>` per CT_LineProperties
    // schema sequence (ECMA-376 Part 1, §20.1.2.3.24).
    if (borderDash !== undefined) {
      lnChildren.push(xmlSelfClose("a:prstDash", { val: borderDash }));
    }
    children.push(
      lnChildren.length === 0
        ? xmlSelfClose("a:ln", lnAttrs)
        : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren),
    );
  }
  return xmlElement("c:spPr", undefined, children);
}

/**
 * Normalize a {@link SheetChart.titleRotation} value (whole degrees)
 * for the `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx>
 * </c:title>` writer slot. Returns `undefined` when the input is unset,
 * non-finite, non-numeric, or resolves to `0` after rounding — every
 * absence path collapses to the same omit-the-attribute shape so
 * absence and the OOXML default `0` round-trip identically through
 * {@link cloneChart}. Out-of-range inputs clamp to the `-90..90` band
 * Excel's UI exposes; non-integer inputs round to the nearest whole
 * degree (the OOXML attribute is an integer in 60000ths of a degree,
 * so a fractional whole-degree value has no meaningful refinement at
 * emit time).
 */
export function normalizeTitleRotation(value: number | undefined): number | undefined {
  if (value === undefined || typeof value !== "number" || !Number.isFinite(value)) return undefined;
  let degrees = Math.round(value);
  if (degrees < TITLE_ROTATION_MIN_DEG) degrees = TITLE_ROTATION_MIN_DEG;
  else if (degrees > TITLE_ROTATION_MAX_DEG) degrees = TITLE_ROTATION_MAX_DEG;
  if (degrees === 0) return undefined;
  return degrees;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:bodyPr rot="N"/></c:rich></c:tx>
 * </c:title>` from {@link SheetChart.titleRotation}.
 *
 * Returns the rotation in whole degrees, or `undefined` when the chart
 * leaves the field unset / pinned the OOXML default `0` / passed a
 * non-numeric or non-finite token. The flag is only meaningful when
 * the chart actually emits a title — the caller is expected to gate
 * the call on `showTitle && chart.title`. A chart whose title is
 * suppressed has no `<c:title>` block to host the rotation in either
 * case.
 */
export function resolveTitleRotation(chart: SheetChart): number | undefined {
  return normalizeTitleRotation(chart.titleRotation);
}

/**
 * Normalize a {@link SheetChart.titleFontSize} value (whole / half
 * points) for the `<c:title><c:tx><c:rich><a:p><a:pPr>
 * <a:defRPr sz="N"/></a:pPr></a:p></c:rich></c:tx></c:title>` writer
 * slot. Returns `undefined` when the input is unset, non-finite,
 * non-numeric, or out of the `1..400`pt band the OOXML
 * `ST_TextFontSize` schema exposes — every absence path collapses to
 * the same default-the-attribute shape so absence and an out-of-range
 * input both fall back to Excel's reference 14pt.
 *
 * Fractional inputs round to the nearest 0.5pt (the OOXML attribute is
 * an integer in 100ths of a point and Excel's UI exposes the same
 * 0.5pt granularity, so finer fractions have no meaningful refinement
 * at emit time).
 */
export function normalizeTitleFontSize(value: number | undefined): number | undefined {
  if (value === undefined || typeof value !== "number" || !Number.isFinite(value)) return undefined;
  // Round to the nearest 0.5pt (Excel's UI granularity). `Math.round`
  // on `2 * value` and dividing by 2 gives a clean half-step band.
  const halfSteps = Math.round(value * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="N"/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` from
 * {@link SheetChart.titleFontSize}.
 *
 * Returns the size in points (`1..400`), or `undefined` when the chart
 * leaves the field unset / passed an out-of-range or non-numeric
 * token. The flag is only meaningful when the chart actually emits a
 * title — the caller is expected to gate the call on
 * `showTitle && chart.title`. A chart whose title is suppressed has
 * no `<c:title>` block to host the size in either case.
 */
export function resolveTitleFontSize(chart: SheetChart): number | undefined {
  return normalizeTitleFontSize(chart.titleFontSize);
}

/**
 * Normalize a {@link SheetChart.titleBold} value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` writer slot. Returns the literal
 * boolean when the input is `true` / `false`, or `undefined` for any
 * other token (including `null`-shaped escapes from an untyped
 * caller). Absence and non-boolean tokens both collapse to
 * `undefined` so the writer falls back to the OOXML default `b="0"`
 * (non-bold) Excel itself emits on a fresh chart title.
 */
export function normalizeTitleBold(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` from
 * {@link SheetChart.titleBold}.
 *
 * Returns the literal boolean, or `undefined` when the chart leaves
 * the field unset / passed a non-boolean token. The flag is only
 * meaningful when the chart actually emits a title — the caller is
 * expected to gate the call on `showTitle && chart.title`. A chart
 * whose title is suppressed has no `<c:title>` block to host the flag
 * in either case.
 */
export function resolveTitleBold(chart: SheetChart): boolean | undefined {
  return normalizeTitleBold(chart.titleBold);
}

/**
 * Normalize a {@link SheetChart.titleItalic} value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` writer slot. Returns the literal
 * boolean when the input is `true` / `false`, or `undefined` for any
 * other token (including `null`-shaped escapes from an untyped
 * caller). Absence and non-boolean tokens both collapse to
 * `undefined` so the writer omits the `i` attribute (Excel's reference
 * serialization for a non-italic title — the OOXML default `false`
 * collapses to absence).
 */
export function normalizeTitleItalic(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr i=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` from
 * {@link SheetChart.titleItalic}.
 *
 * Returns the literal boolean, or `undefined` when the chart leaves
 * the field unset / passed a non-boolean token. The flag is only
 * meaningful when the chart actually emits a title — the caller is
 * expected to gate the call on `showTitle && chart.title`. A chart
 * whose title is suppressed has no `<c:title>` block to host the flag
 * in either case.
 */
export function resolveTitleItalic(chart: SheetChart): boolean | undefined {
  return normalizeTitleItalic(chart.titleItalic);
}

/**
 * Normalize a {@link SheetChart.titleColor} value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` writer slot. Returns the 6-character
 * uppercase hex form when the input is a valid sRGB triple (with or
 * without a leading `#`), or `undefined` for any malformed token —
 * wrong length, non-hex characters, alpha-channel forms, or
 * non-string escapes from an untyped caller.
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the entire `<a:solidFill>` block and the title
 * inherits the theme text color (Excel's reference behavior for a
 * fresh chart title without a custom color).
 */
export function normalizeTitleColor(value: string | undefined): string | undefined {
  return normalizeRgbHexShared(value);
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:rich></c:tx></c:title>` from {@link SheetChart.titleColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed a
 * malformed token. The fill is only meaningful when the chart
 * actually emits a title — the caller is expected to gate the call
 * on `showTitle && chart.title`. A chart whose title is suppressed
 * has no `<c:title>` block to host the fill in either case.
 */
export function resolveTitleColor(chart: SheetChart): string | undefined {
  return normalizeTitleColor(chart.titleColor);
}

/**
 * Resolve `<c:title><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:title>` from
 * {@link SheetChart.titleFillColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches the chart-title font
 * color / plot-area fill / legend fill resolvers exactly. The fill
 * is only meaningful when the chart actually emits a title — the
 * caller is expected to gate the call on `showTitle && chart.title`.
 * A chart whose title is suppressed has no `<c:title>` block to host
 * the `<c:spPr>` slot in either case.
 *
 * Independent of {@link resolveTitleColor}: the fill lands on
 * `<c:title><c:spPr>`, the font color lands on the
 * `<a:defRPr><a:solidFill>` slot inside `<c:tx><c:rich><a:p><a:pPr>`
 * — the two resolvers target different children of `<c:title>` so a
 * single configuration call can pin both.
 */
export function resolveTitleFillColor(chart: SheetChart): string | undefined {
  return normalizeTitleColor(chart.titleFillColor);
}

/**
 * Resolve `<c:title><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:title>` from
 * {@link SheetChart.titleBorderColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches every other
 * `<a:srgbClr>` slot the writer authors. The stroke is only
 * meaningful when the chart actually emits a title — the caller is
 * expected to gate the call on `showTitle && chart.title`. A chart
 * whose title is suppressed has no `<c:title>` block to host the
 * `<c:spPr>` slot in either case.
 *
 * Independent of {@link resolveTitleFillColor}: the stroke lands on
 * `<c:title><c:spPr><a:ln>`, the fill lands on
 * `<c:title><c:spPr><a:solidFill>` — the two resolvers target
 * different children of the shared `<c:spPr>` block so a single
 * configuration call can pin both. Mirrors
 * {@link normalizePlotAreaBorderColor} — same hex grammar, distinct
 * host element (`<c:title>` vs `<c:plotArea>`).
 */
export function resolveTitleBorderColor(chart: SheetChart): string | undefined {
  return normalizeTitleColor(chart.titleBorderColor);
}

/**
 * Resolve `<c:title><c:spPr><a:ln w="EMU"/></c:spPr></c:title>` from
 * {@link SheetChart.titleBorderWidth}.
 *
 * Returns the point value clamped to the `0.25..13.5` pt band Excel's
 * UI exposes and snapped to the 0.25 pt grid, or `undefined` when the
 * chart leaves the field unset / passed a malformed token (`NaN`,
 * `Infinity`, non-finite). Delegates to {@link clampStrokeWidthPt} so
 * the snap / clamp grammar matches every other `<a:ln w=..>` slot the
 * writer authors (the series stroke knob `series[i].stroke.width`,
 * the plot-area border width knob {@link SheetChart.plotAreaBorderWidth},
 * and the legend border width knob {@link SheetChart.legendBorderWidth}).
 * The width is only meaningful when the chart actually emits a
 * title — the caller is expected to gate the call on
 * `showTitle && chart.title`. A chart whose title is suppressed has no
 * `<c:title>` block to host the `<c:spPr>` slot in either case.
 *
 * Independent of {@link resolveTitleBorderColor}: both knobs land on
 * the same `<a:ln>` element but on a different slot (the color child
 * `<a:solidFill>` versus the line's `w` attribute). Mirrors the
 * plot-area / legend `<c:spPr>` slots — same EMU encoding, same
 * `<a:ln>` host — but lands on `<c:title>`'s own `<c:spPr>` block.
 */
export function resolveTitleBorderWidth(chart: SheetChart): number | undefined {
  return clampStrokeWidthPt(chart.titleBorderWidth);
}

/**
 * Resolve `<c:title><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
 * </c:title>` from {@link SheetChart.titleBorderDash}.
 *
 * Returns the recognized {@link ChartBorderDash} value, or `undefined`
 * for the OOXML default `"solid"` and every unrecognized token —
 * delegates to {@link normalizeBorderDash} so the accept / drop grammar
 * matches every other `<a:prstDash>` slot the writer authors. The
 * caller is expected to gate the call on `showTitle && chart.title`
 * since a chart whose title is suppressed has no `<c:title>` block to
 * host the `<c:spPr>` slot.
 *
 * Independent of {@link resolveTitleBorderColor} and
 * {@link resolveTitleBorderWidth}: all three knobs land on the same
 * `<a:ln>` element but on different children / attributes — color is
 * `<a:solidFill>`, width is the `w` attribute, dash is `<a:prstDash>`.
 */
export function resolveTitleBorderDash(chart: SheetChart): ChartBorderDash | undefined {
  return normalizeBorderDash(chart.titleBorderDash);
}

/**
 * Normalize a {@link SheetChart.titleStrike} value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` writer slot. Returns the literal
 * boolean when the input is `true` / `false`, or `undefined` for any
 * other token (including `null`-shaped escapes from an untyped
 * caller). Absence and non-boolean tokens both collapse to
 * `undefined` so the writer omits the `strike` attribute entirely
 * (Excel's reference serialization for a non-strikethrough title —
 * the OOXML default `"noStrike"` collapses to absence; only an
 * explicit `true` emits `strike="sngStrike"`).
 */
export function normalizeTitleStrike(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr strike=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` from
 * {@link SheetChart.titleStrike}.
 *
 * Returns the literal boolean, or `undefined` when the chart leaves
 * the field unset / passed a non-boolean token. The flag is only
 * meaningful when the chart actually emits a title — the caller is
 * expected to gate the call on `showTitle && chart.title`. A chart
 * whose title is suppressed has no `<c:title>` block to host the flag
 * in either case.
 */
export function resolveTitleStrike(chart: SheetChart): boolean | undefined {
  return normalizeTitleStrike(chart.titleStrike);
}

/**
 * Normalize a {@link SheetChart.titleUnderline} value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
 * </a:p></c:rich></c:tx></c:title>` writer slot. Returns the literal
 * boolean when the input is `true` / `false`, or `undefined` for any
 * other token (including `null`-shaped escapes from an untyped
 * caller). Absence and non-boolean tokens both collapse to
 * `undefined` so the writer omits the `u` attribute entirely (Excel's
 * reference serialization for a non-underlined title — the OOXML
 * default `"none"` collapses to absence; only an explicit `true`
 * emits `u="sng"`).
 */
export function normalizeTitleUnderline(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr u=".."/>
 * </a:pPr></a:p></c:rich></c:tx></c:title>` from
 * {@link SheetChart.titleUnderline}.
 *
 * Returns the literal boolean, or `undefined` when the chart leaves
 * the field unset / passed a non-boolean token. The flag is only
 * meaningful when the chart actually emits a title — the caller is
 * expected to gate the call on `showTitle && chart.title`. A chart
 * whose title is suppressed has no `<c:title>` block to host the flag
 * in either case.
 */
export function resolveTitleUnderline(chart: SheetChart): boolean | undefined {
  return normalizeTitleUnderline(chart.titleUnderline);
}

/**
 * Normalize a {@link SheetChart.titleFontFamily} value for the
 * `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx>
 * </c:title>` writer slot. Returns the trimmed typeface string when
 * the input is a non-empty string, or `undefined` for any malformed
 * token — empty / whitespace-only strings, or non-string escapes from
 * an untyped caller (`null`, numbers, booleans, etc.).
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the entire `<a:latin>` element and the title inherits
 * the theme typeface (Excel's reference behavior for a fresh chart
 * title without a custom font picked).
 */
export function normalizeTitleFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Resolve `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:rich></c:tx></c:title>`
 * from {@link SheetChart.titleFontFamily}.
 *
 * Returns the trimmed typeface string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed an empty
 * or non-string token. The element is only meaningful when the chart
 * actually emits a title — the caller is expected to gate the call
 * on `showTitle && chart.title`. A chart whose title is suppressed
 * has no `<c:title>` block to host the typeface in either case.
 */
export function resolveTitleFontFamily(chart: SheetChart): string | undefined {
  return normalizeTitleFontFamily(chart.titleFontFamily);
}

/**
 * Resolve `<c:title><c:overlay val=".."/></c:title>` from
 * {@link SheetChart.titleOverlay}.
 *
 * Defaults to `false` (the OOXML default Excel itself emits — the title
 * reserves its own slot above the plot area and the plot area shrinks
 * to make room). Anything other than literal `true` collapses to `false`
 * so a stray non-boolean leaking through the type guard (e.g. `0` / `1` /
 * `"true"` / `null`) never produces `<c:overlay val="1"/>`. This matches
 * how `legendOverlay` / `roundedCorners` / `plotVisOnly` / axis `hidden`
 * treat their inputs: a literal boolean is the only path to a non-default
 * value.
 *
 * The writer always emits `<c:overlay>` inside `<c:title>` because Excel's
 * reference serialization includes the element on every visible title;
 * only the `val` flips when the caller pins `titleOverlay: true`.
 *
 * The flag is only meaningful when the chart actually emits a title — the
 * caller is expected to gate the call on `showTitle && chart.title`. A
 * chart whose title is suppressed has no `<c:title>` block to host the
 * overlay element.
 */
export function resolveTitleOverlay(chart: SheetChart): boolean {
  return chart.titleOverlay === true;
}

/**
 * Resolve `<c:title><c:layout><c:manualLayout>...</c:manualLayout>
 * </c:layout></c:title>` from {@link SheetChart.titleLayout}.
 *
 * Returns the normalized coordinate set, or `undefined` when every
 * axis the caller pinned dropped to `undefined` (so the writer can
 * elide the entire `<c:layout>` block — Excel's reference serialization
 * omits the element when the title renders at the auto-layout position
 * above the plot area). The element is only meaningful when the chart
 * actually emits a title — the caller is expected to gate the call on
 * the resolved title visibility (showTitle && chart.title).
 *
 * Coordinates outside the OOXML `0..1` band, `NaN`, `Infinity`, and
 * non-numeric inputs all collapse to `undefined` on the matching axis
 * so the writer drops the matching `<c:x>` / `<c:y>` / `<c:w>` /
 * `<c:h>` slot rather than emit a token Excel would reject. Mirrors
 * {@link resolveLegendLayout} — same accept-or-drop grammar, same
 * `ChartManualLayout` shape — so a caller can thread a single layout
 * value through both the chart title and the legend without
 * bookkeeping a second type.
 */
export function resolveTitleLayout(chart: SheetChart): ResolvedManualLayout | undefined {
  return normalizeManualLayout(chart.titleLayout);
}

// ── Clone-side title constants ────────────────────────────────────

const TITLE_BORDER_WIDTH_MIN_PT = 0.25;
const TITLE_BORDER_WIDTH_MAX_PT = 13.5;

// ── Clone resolvers (3-arg source/override) ───────────────────────

/**
 * Resolve a `titleLayout` override.
 *
 * `undefined` → inherit the source's parsed `titleLayout` (after
 *               running it through {@link normalizeLegendLayout} so a
 *               malformed source value drops cleanly — both manual-
 *               layout slots share the same normalizer).
 * `null`      → drop the inherited layout (the writer falls back to
 *               Excel's auto-layout position above the plot area —
 *               no `<c:layout>` block on `<c:title>`).
 * `ChartManualLayout` → replace, after running through
 *               {@link normalizeLegendLayout}. Coordinates outside the
 *               `0..1` band collapse on the matching axis so the
 *               cloned `SheetChart` always carries a value the writer
 *               will accept; an override whose every axis dropped
 *               collapses to `undefined` so the writer skips the
 *               `<c:layout>` block entirely.
 *
 * The grammar mirrors `legendLayout` — both manual-layout slots
 * compose the same way at the call site. Callers should gate the
 * result on the resolved title visibility — when no title is emitted,
 * the layout has no slot in the rendered chart.
 */
export function resolveCloneTitleLayout(
  sourceValue: ChartManualLayout | undefined,
  override: ChartManualLayout | null | undefined,
): ChartManualLayout | undefined {
  if (override === undefined) return normalizeChartManualLayout(sourceValue);
  if (override === null) return undefined;
  return normalizeChartManualLayout(override);
}

/**
 * Resolve a `titleOverlay` override.
 *
 * `undefined` → inherit the source's parsed `titleOverlay`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — the title reserves its own slot
 *               above the plot area, no overlap with it).
 * `boolean`   → replace.
 *
 * The grammar mirrors `legendOverlay` / `roundedCorners` so the chart-
 * level overlay toggles compose the same way at the call site. Callers
 * should gate the result on the resolved title visibility — when no
 * title is emitted, the overlay flag has no slot in the rendered chart.
 */
export function resolveCloneTitleOverlay(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `titleRotation` override.
 *
 * `undefined` → inherit the source's parsed `titleRotation`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `0` default — the title renders horizontally).
 * `number`    → replace, after clamping / rounding through
 *               {@link normalizeTitleRotation}.
 *
 * The grammar mirrors `titleOverlay` / `legendOverlay` so the chart-
 * level title knobs compose the same way at the call site. Callers
 * should gate the result on the resolved title visibility — when no
 * title is emitted, the rotation has no slot in the rendered chart.
 */
export function resolveCloneTitleRotation(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeTitleRotation(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleRotation(override);
}

/**
 * Resolve a `titleFontSize` override.
 *
 * `undefined` → inherit the source's parsed `titleFontSize`.
 * `null`      → drop the inherited value (the writer falls back to
 *               Excel's default 14pt).
 * `number`    → replace, after clamping / rounding through
 *               {@link normalizeTitleFontSize}.
 *
 * The grammar mirrors `titleRotation` / `titleOverlay` /
 * `legendOverlay` so the chart-level title knobs compose the same way
 * at the call site. Callers should gate the result on the resolved
 * title visibility — when no title is emitted, the size has no slot
 * in the rendered chart.
 */
export function resolveCloneTitleFontSize(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeTitleFontSize(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleFontSize(override);
}

/**
 * Resolve a `titleBold` override.
 *
 * `undefined` → inherit the source's parsed `titleBold`.
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default `b="0"`, non-bold).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleFontSize` / `titleRotation` /
 * `titleOverlay` so the chart-level title knobs compose the same way
 * at the call site. Callers should gate the result on the resolved
 * title visibility — when no title is emitted, the flag has no slot
 * in the rendered chart.
 */
export function resolveCloneTitleBold(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeTitleBold(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleBold(override);
}

/**
 * Resolve a `titleItalic` override.
 *
 * `undefined` → inherit the source's parsed `titleItalic`.
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default — no `i` attribute, equivalent to
 *               non-italic).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleBold` / `titleFontSize` / `titleRotation`
 * / `titleOverlay` so the chart-level title knobs compose the same way
 * at the call site. Callers should gate the result on the resolved
 * title visibility — when no title is emitted, the flag has no slot
 * in the rendered chart.
 */
export function resolveCloneTitleItalic(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeTitleItalic(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleItalic(override);
}

/**
 * Resolve a `titleColor` override.
 *
 * `undefined` → inherit the source's parsed `titleColor`.
 * `null`      → drop the inherited fill (the writer falls back to the
 *               theme text color — no `<a:solidFill>` block on the
 *               title's default-paragraph properties).
 * `string`    → replace with the normalized 6-character uppercase
 *               hex form.
 *
 * The grammar mirrors `titleBold` / `titleItalic` / `titleFontSize` /
 * `titleRotation` / `titleOverlay` so the chart-level title knobs
 * compose the same way at the call site. Callers should gate the
 * result on the resolved title visibility — when no title is
 * emitted, the fill has no slot in the rendered chart.
 */
export function resolveCloneTitleColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeTitleColor(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleColor(override);
}

/**
 * Resolve a `titleFillColor` override.
 *
 * `undefined` → inherit the source's parsed `titleFillColor` (after
 *               running it through {@link normalizeTitleColor} so a
 *               malformed source value drops cleanly — the hex
 *               normalizer is purely shape-based and applies
 *               identically to every `<a:srgbClr val="RRGGBB"/>`
 *               slot).
 * `null`      → drop the inherited fill (the writer emits no
 *               `<c:spPr>` block on `<c:title>`, falling back to the
 *               theme default — typically a transparent title
 *               background).
 * `string`    → replace, after running through
 *               {@link normalizeTitleColor} so the override accepts
 *               `"FF0000"` / `"#FF0000"` / `"ff0000"` and collapses
 *               malformed tokens to `undefined`.
 *
 * The grammar mirrors `plotAreaFillColor` / `legendFillColor` /
 * `titleColor` / `axisTitleColor` so the fill / color knobs compose
 * the same way at the call site. Callers should gate the result on
 * the resolved title visibility — when no title is emitted, the fill
 * has no slot in the rendered chart.
 *
 * Independent of `titleColor`: the two knobs target different
 * children of `<c:title>` (`<c:spPr>` for the background fill,
 * `<c:tx><c:rich><a:p><a:pPr><a:defRPr><a:solidFill>` for the font
 * color), so a caller can pin both without conflict.
 */
export function resolveCloneTitleFillColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeTitleColor(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleColor(override);
}

/**
 * Resolve a `titleBorderColor` override.
 *
 * `undefined` → inherit the source's parsed `titleBorderColor` (after
 *               running it through {@link normalizeTitleColor} so a
 *               malformed source value drops cleanly — the hex
 *               normalizer is purely shape-based and applies
 *               identically to every `<a:srgbClr val="RRGGBB"/>`
 *               slot).
 * `null`      → drop the inherited stroke (the writer emits no
 *               `<a:ln>` block on `<c:title><c:spPr>`, falling back
 *               to the theme default — typically no visible border).
 * `string`    → replace, after running through
 *               {@link normalizeTitleColor} so the override accepts
 *               `"1F77B4"` / `"#1F77B4"` / `"1f77b4"` and collapses
 *               malformed tokens to `undefined`.
 *
 * The grammar mirrors `plotAreaBorderColor` / `titleFillColor` so the
 * chart `<c:spPr>` knobs compose the same way at the call site.
 * Callers should gate the result on the resolved title visibility —
 * when no title is emitted, the stroke has no slot in the rendered
 * chart.
 *
 * Independent of `titleFillColor`: the two knobs target different
 * children of the shared `<c:spPr>` block on `<c:title>`
 * (`<a:solidFill>` for the fill, `<a:ln>` for the stroke), so a
 * caller can pin both without conflict.
 */
export function resolveCloneTitleBorderColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeTitleColor(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleColor(override);
}

/**
 * Normalize a `titleBorderWidth` value for the cloned `SheetChart`.
 * Mirrors the writer's `clampStrokeWidthPt` — values are clamped to the
 * `0.25..13.5` pt band Excel's UI exposes and snapped to the 0.25 pt
 * grid so a parsed-then-cloned-then-written width does not drift across
 * round-trips (Excel rounds in the UI anyway). Non-finite / non-numeric
 * tokens (`NaN`, `Infinity`, strings, `null` from an untyped caller)
 * collapse to `undefined` so the cloned chart drops the field rather
 * than carry a value the writer would silently elide back to absence.
 */
export function normalizeTitleBorderWidth(value: number | undefined): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const snapped = Math.round(value * 4) / 4;
  if (snapped < TITLE_BORDER_WIDTH_MIN_PT) return TITLE_BORDER_WIDTH_MIN_PT;
  if (snapped > TITLE_BORDER_WIDTH_MAX_PT) return TITLE_BORDER_WIDTH_MAX_PT;
  return snapped;
}

/**
 * Resolve a `titleBorderWidth` override.
 *
 * `undefined` → inherit the source's parsed `titleBorderWidth` (after
 *               running it through {@link normalizeTitleBorderWidth}
 *               so a malformed source value drops cleanly).
 * `null`      → drop the inherited width (the writer emits `<a:ln>`
 *               without a `w` attribute, the line keeps Excel's
 *               auto-thickness).
 * `number`    → replace with the clamped / snapped point value.
 *               Non-finite / non-numeric overrides collapse to
 *               `undefined` via the normalizer so the cloned
 *               `SheetChart` always carries a value the writer will
 *               accept.
 *
 * The grammar mirrors `plotAreaBorderWidth` / `legendBorderWidth` /
 * the series-line stroke width so the chart `<a:ln w=..>` knobs
 * compose the same way at the call site. Callers should gate the
 * result on the resolved title visibility — when no title is emitted,
 * the width has no slot in the rendered chart.
 *
 * Independent of `titleBorderColor`: both knobs land on the same
 * `<a:ln>` element but on a different slot (color is
 * `<a:solidFill><a:srgbClr>`, width is the `w` attribute on `<a:ln>`),
 * so a caller can pin both without conflict.
 */
export function resolveCloneTitleBorderWidth(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeTitleBorderWidth(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleBorderWidth(override);
}

/**
 * Resolve a `titleStrike` override.
 *
 * `undefined` → inherit the source's parsed `titleStrike`.
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default — no `strike` attribute, equivalent to
 *               no strikethrough).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleBold` / `titleItalic` / `titleColor` /
 * `titleFontSize` / `titleRotation` / `titleOverlay` so the chart-level
 * title knobs compose the same way at the call site. Callers should
 * gate the result on the resolved title visibility — when no title is
 * emitted, the flag has no slot in the rendered chart.
 */
export function resolveCloneTitleStrike(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeTitleStrike(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleStrike(override);
}

/**
 * Resolve a `titleUnderline` override.
 *
 * `undefined` → inherit the source's parsed `titleUnderline`.
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default — no `u` attribute, equivalent to no
 *               underline).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleBold` / `titleItalic` / `titleStrike` /
 * `titleColor` / `titleFontSize` / `titleRotation` / `titleOverlay`
 * so the chart-level title knobs compose the same way at the call
 * site. Callers should gate the result on the resolved title
 * visibility — when no title is emitted, the flag has no slot in the
 * rendered chart.
 */
export function resolveCloneTitleUnderline(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeTitleUnderline(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleUnderline(override);
}

/**
 * Resolve a `titleFontFamily` override.
 *
 * `undefined` → inherit the source's parsed `titleFontFamily`,
 *               running it through {@link normalizeTitleFontFamily}
 *               so a malformed source value cannot leak through to
 *               the cloned chart.
 * `null`      → drop the inherited typeface (the writer falls back to
 *               the OOXML default — no `<a:latin>` element, the title
 *               inherits the theme typeface).
 * `string`    → replace, running it through
 *               {@link normalizeTitleFontFamily} so the override
 *               accepts any caller spelling that the writer will
 *               accept (with surrounding whitespace trimmed; empty /
 *               whitespace-only strings collapse to a drop).
 *
 * The grammar mirrors `titleColor` (the other string-typed knob) /
 * `titleBold` / `titleItalic` / `titleStrike` / `titleUnderline` /
 * `titleFontSize` / `titleRotation` / `titleOverlay` so the chart-
 * level title knobs compose the same way at the call site. Callers
 * should gate the result on the resolved title visibility — when no
 * title is emitted, the typeface has no slot in the rendered chart.
 */
export function resolveCloneTitleFontFamily(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeTitleFontFamily(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleFontFamily(override);
}
