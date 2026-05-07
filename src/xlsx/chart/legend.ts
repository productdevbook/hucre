// ── Chart Legend ──────────────────────────────────────────────────
// Per-host module for `<c:chart><c:legend>` (CT_Legend, ECMA-376 Part 1,
// §21.2.2.114). Holds the reader / writer / clone helpers — every
// `parse*` / `build*` / `resolve*` / `normalize*` function for the
// chart-level legend block, including its `<c:legendPos>` /
// `<c:legendEntry>` / `<c:overlay>` / `<c:layout>` / `<c:txPr>` /
// `<c:spPr>` children.
//
// The clone-side `resolveLegend*` overrides (3-arg `(source, override)`
// resolvers) live in `chart-clone.ts` because their signatures are
// shape-incompatible with the writer-side single-arg resolvers; they
// delegate to the shared `normalizeLegend*` exports here so the
// per-field clamp / drop grammar stays in one place.

import type {
  ChartBorderDash,
  ChartLegendEntry,
  ChartLegendPosition,
  ChartManualLayout,
  SheetChart,
} from "../../_types";
import type { XmlElement } from "../../xml/parser";
import { xmlElement, xmlSelfClose } from "../../xml/writer";
import {
  EMU_PER_PT,
  clampStrokeWidthPt,
  normalizeBorderDash,
  normalizeRgbHex,
  parseBorderDashFromSpPr,
  parseBorderWidthFromSpPr,
  parseSpPrBorderColor,
  parseSpPrFill,
} from "./shape";
import {
  type ResolvedManualLayout,
  buildManualLayout,
  normalizeChartManualLayout,
  normalizeLayoutCoordinate,
  normalizeManualLayout,
  parseManualLayout,
} from "./layout";
import { childElements, findChild, parseBoolAttr, readBoolVal } from "./util";
import {
  FONT_SIZE_MAX_PT,
  FONT_SIZE_MIN_PT,
  FONT_SZ_PER_POINT,
  ROTATION_MAX_DEG,
  ROTATION_MIN_DEG,
} from "./text";
import { normalizeTitleColor, normalizeTitleFontSize } from "./title";

// ── Legend types (writer-side) ────────────────────────────────────

export type LegendPos = "t" | "b" | "l" | "r" | "tr";

export interface ResolvedLegendEntry {
  idx: number;
  delete: boolean;
}

// ── Constants (legend-scope aliases) ──────────────────────────────

const TITLE_FONT_SZ_PER_POINT = FONT_SZ_PER_POINT;
const TITLE_FONT_SIZE_MIN_PT = FONT_SIZE_MIN_PT;
const TITLE_FONT_SIZE_MAX_PT = FONT_SIZE_MAX_PT;

const LEGEND_BORDER_WIDTH_MIN_PT = 0.25;
const LEGEND_BORDER_WIDTH_MAX_PT = 13.5;

// ── Reader ────────────────────────────────────────────────────────

/**
 * Map `<c:legend><c:legendPos val=".."/></c:legend>` to the writer-side
 * {@link ChartLegendPosition}. Returns `false` when `<c:delete val="1"/>`
 * is present (Excel's "no legend" state); returns `undefined` when the
 * chart has no `<c:legend>` element at all.
 */
export function parseLegend(chartEl: XmlElement): false | ChartLegendPosition | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;

  // <c:delete val="1"/> means the chart explicitly suppresses the
  // legend. Some Excel versions emit just an empty `<c:legend/>`
  // followed by `<c:overlay/>` even when the legend is hidden, but
  // `<c:delete val="1">` is the canonical "no legend" marker.
  const del = findChild(legend, "delete");
  if (del && readBoolVal(del.attrs.val) === true) return false;

  const pos = findChild(legend, "legendPos");
  if (!pos) {
    // A legend element without legendPos is valid OOXML (Excel falls
    // back to "right"). Surface "right" so the cloned chart preserves
    // the visible-legend state.
    return "right";
  }
  const val = pos.attrs.val;
  if (typeof val !== "string") return "right";
  switch (val) {
    case "t":
      return "top";
    case "b":
      return "bottom";
    case "l":
      return "left";
    case "r":
      return "right";
    case "tr":
      return "topRight";
    default:
      // Unknown legendPos values are dropped rather than fabricated.
      return undefined;
  }
}

/**
 * Pull `<c:legend><c:overlay val=".."/></c:legend>` off the chart. The
 * OOXML default `false` (the legend reserves its own slot, no overlap
 * with the plot area) collapses to `undefined` so absence and
 * `<c:overlay val="0"/>` round-trip identically through
 * {@link cloneChart} — only an explicit `<c:overlay val="1"/>` surfaces
 * `true`.
 *
 * The caller is expected to confirm a visible legend exists before
 * invoking this — `<c:overlay>` only renders when the legend is part of
 * the chart, so reading it from a chart that hides or omits the legend
 * would surface a flag with no on-screen effect.
 *
 * Accepts the OOXML truthy / falsy spellings (`"1"` / `"true"` / `"0"`
 * / `"false"`); unknown values and missing `val` attributes drop to
 * `undefined` rather than fabricate a flag Excel would not emit.
 */
export function parseLegendOverlay(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const overlay = findChild(legend, "overlay");
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
      // writer's `legendOverlay` field.
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:legend><c:legendEntry>` overrides off the chart. Returns
 * `undefined` when the chart declares no entries so the field is
 * elided entirely on a clean parse — absence and an empty array
 * round-trip identically through {@link cloneChart} (the writer skips
 * emission when the resolved list is empty).
 *
 * Each entry is admitted only when its `<c:idx val=".."/>` selector
 * parses to a non-negative integer (matches the OOXML
 * `xsd:unsignedInt` schema). Entries without an `<c:idx>` child or with
 * a malformed `val` attribute are dropped rather than surface a
 * fabricated index. The `<c:delete>` flag accepts the OOXML truthy /
 * falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`); absence
 * collapses to `false` (the OOXML default — the entry renders).
 *
 * The caller is expected to confirm a visible legend exists before
 * invoking this — `<c:legendEntry>` only renders inside `<c:legend>`,
 * so reading from a chart that hides or omits the legend would surface
 * overrides with no on-screen effect.
 *
 * Duplicate `idx` values keep the first occurrence — Excel's renderer
 * treats later duplicates as overrides on the same series, but the
 * writer's `resolveLegendEntries` deduplicates with last-wins semantics
 * to give clone-through callers a way to override without manually
 * pruning. Reading "first wins" pairs naturally with that behaviour:
 * a parsed list re-emits cleanly, and an explicit clone override that
 * appends an entry still beats the parsed value.
 */
export function parseLegendEntries(chartEl: XmlElement): ChartLegendEntry[] | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;

  const seen = new Set<number>();
  const out: ChartLegendEntry[] = [];
  for (const child of childElements(legend)) {
    if (child.local !== "legendEntry") continue;
    const idxEl = findChild(child, "idx");
    if (!idxEl) continue;
    const raw = idxEl.attrs.val;
    if (typeof raw !== "string") continue;
    const idx = Number.parseInt(raw, 10);
    if (!Number.isFinite(idx) || idx < 0) continue;
    if (seen.has(idx)) continue;
    seen.add(idx);

    const deleteEl = findChild(child, "delete");
    const deleteFlag = deleteEl !== undefined ? readBoolVal(deleteEl.attrs.val) === true : false;
    out.push({ idx, delete: deleteFlag });
  }

  return out.length > 0 ? out : undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the font size in points
 * (range `1..400`).
 *
 * The OOXML `sz` attribute is in 100ths of a point — the reader
 * converts to points and rounds to the nearest 0.5pt (Excel's UI
 * exposes the same 0.5pt granularity). Absence of the element /
 * attribute and out-of-range / non-numeric / non-finite values all
 * collapse to `undefined` so a fresh chart and a chart that pinned an
 * out-of-range size both round-trip to the writer's "skip the size
 * attribute" path.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>` element
 * — there is no `<c:txPr>` slot to surface the size from in that case.
 * The `<a:defRPr>` lives inside `<c:txPr><a:p><a:pPr>` per the
 * CT_TextBody schema (the default-paragraph properties on the
 * legend's text-body's first paragraph); the lookup is scoped to that
 * path so a stray `<a:defRPr>` elsewhere in the chart (e.g. on the
 * chart title or an axis) cannot leak in. Mirrors the chart-title /
 * axis-title / axis tick-label readers exactly so a parsed value slots
 * straight back into the writer's emit path.
 */
export function parseLegendFontSize(chartEl: XmlElement): number | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
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
  // band that mirrors the writer's emit-time normalization. The
  // chart-title / axis-title / tick-label sibling parsers use the
  // identical conversion so a parsed value flows through every
  // typography slot without bookkeeping the units.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the bold flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `b="0"` round-trip identically — only an explicit `b="1"` surfaces
 * `true`. Unknown / malformed `b` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element — there is no `<c:txPr>` slot to surface the flag from in
 * that case. The `<a:defRPr>` lives inside `<c:txPr><a:p><a:pPr>` per
 * the CT_TextBody schema (the default-paragraph properties on the
 * legend's text-body's first paragraph); the lookup is scoped to that
 * path so a stray `<a:defRPr>` elsewhere in the chart (e.g. on the
 * chart title or an axis) cannot leak in. Mirrors the chart-title /
 * axis-title / axis tick-label readers exactly so a parsed value
 * slots straight back into the writer's emit path.
 */
export function parseLegendBold(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.b;
  // OOXML `xsd:boolean` accepts `"1"` / `"true"` (truthy) and `"0"` /
  // `"false"` (falsy). Truthy spellings surface `true`; falsy
  // spellings collapse to `undefined` so the OOXML default and an
  // explicit `b="0"` round-trip identically through `cloneChart`.
  // Unknown / missing `b` tokens drop to `undefined` for the same
  // reason — never fabricate a flag Excel would not emit.
  if (raw === "1" || raw === "true") return true;
  return undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the italic flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties` (ECMA-376 Part 1, §21.1.2.3.7). The
 * OOXML default `false` collapses to `undefined` so absence and
 * `i="0"` round-trip identically — only an explicit `i="1"` surfaces
 * `true`. Unknown / malformed `i` tokens drop to `undefined` rather
 * than fabricate a value the writer would never emit.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element — there is no `<c:txPr>` slot to surface the flag from in
 * that case. Mirrors the chart-title / axis-title / axis tick-label
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
export function parseLegendItalic(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.i;
  if (raw === "1" || raw === "true") return true;
  return undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p>
 * </c:txPr></c:legend>` off the chart. Returns the underline flag.
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
 * UI exposes.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element. Mirrors the chart-title / axis-title / axis tick-label
 * underline readers exactly so a parsed value slots straight back
 * into the writer's emit path.
 */
export function parseLegendUnderline(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  if (raw === "sng") return true;
  return undefined;
}

/**
 * Pull `<c:legend><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:txPr></c:legend>` off the chart. Returns the strikethrough
 * flag.
 *
 * The OOXML `strike` attribute is the `ST_TextStrikeType` enumeration
 * on `CT_TextCharacterProperties` — `"noStrike"` (default),
 * `"sngStrike"` (single line, Excel's UI variant), `"dblStrike"`
 * (double line, non-UI). Only `strike="sngStrike"` surfaces `true`;
 * the OOXML default `"noStrike"` and the double-line variant both
 * collapse to `undefined` so absence and `strike="noStrike"`
 * round-trip identically through `cloneChart`. Reporting `"dblStrike"`
 * as `true` would silently downgrade the choice to a single line on
 * re-emit; the writer emits only `"sngStrike"`, matching the boolean
 * shape the UI exposes.
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element. Mirrors the chart-title / axis-title / axis tick-label
 * strikethrough readers exactly so a parsed value slots straight back
 * into the writer's emit path.
 */
export function parseLegendStrikethrough(chartEl: XmlElement): boolean | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.strike;
  if (raw === "sngStrike") return true;
  return undefined;
}

/**
 * Pull the legend font color off the canonical
 * `<c:legend><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr
 * val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
 * </c:legend>` chain Excel writes when the user pins a custom font
 * color on the legend.
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
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element or the canonical `<c:txPr><a:p><a:pPr><a:defRPr>
 * <a:solidFill><a:srgbClr>` chain is malformed at any link. Mirrors
 * the chart-title / axis-title / axis tick-label color readers
 * exactly so a parsed value slots straight back into the writer's
 * emit path.
 */
export function parseLegendFontColor(chartEl: XmlElement): string | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const solidFill = findChild(defRPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  return normalizeRgbHex(srgbClr.attrs.val);
}

/**
 * Pull the legend font family off the canonical
 * `<c:legend><c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/>
 * </a:defRPr></a:pPr></a:p></c:txPr></c:legend>` chain Excel writes
 * when the user pins a typeface on the legend.
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
 * Returns `undefined` whenever the chart omits the `<c:legend>`
 * element or the canonical `<c:txPr><a:p><a:pPr><a:defRPr><a:latin>`
 * chain is malformed at any link. Mirrors the chart-title /
 * axis-title / axis tick-label font family readers exactly so a
 * parsed value slots straight back into the writer's emit path.
 */
export function parseLegendFontFamily(chartEl: XmlElement): string | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  const txPr = findChild(legend, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
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
 * Pull `<c:legend><c:layout><c:manualLayout>` off the chart. Reflects
 * Excel's "Format Legend -> Position -> Custom" knob — the `(x, y)`
 * anchor and `(w, h)` size of the legend block as fractions of the
 * chart frame in the `0..1` band.
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
 * emits the absolute form when the user drags an element to a custom
 * position).
 *
 * Returns `undefined` whenever the chart omits the `<c:legend>` /
 * `<c:layout>` / `<c:manualLayout>` chain at any link, or when every
 * coordinate dropped on normalization — the field is omitted entirely
 * on a clean parse so absence and an empty layout round-trip identically
 * through the writer.
 */
export function parseLegendLayout(chartEl: XmlElement): ChartManualLayout | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  return parseManualLayout(legend);
}

/**
 * Pull the legend background fill color off the canonical
 * `<c:legend><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:legend>` chain Excel writes when the
 * user pins "Format Legend -> Fill -> Solid fill -> Color".
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
 * The lookup is scoped to direct children of `<c:legend>` so a stray
 * `<c:spPr>` elsewhere (e.g. on a series or on the legend's
 * `<c:txPr>` block — the latter is purely text-character-properties,
 * but a malformed source might place a `<c:spPr>` there in error)
 * cannot leak in. Returns `undefined` whenever the chart omits the
 * `<c:legend>` element or the `<c:spPr><a:solidFill><a:srgbClr>`
 * chain is malformed at any link. Mirrors the chart-title /
 * axis-title / axis tick-label / legend font color readers exactly
 * so a parsed value slots straight back into the writer's emit path.
 *
 * Independent of {@link parseLegendFontColor}: the fill lives on
 * `<c:legend><c:spPr>`, the font color lives on
 * `<c:legend><c:txPr>` — the two readers walk disjoint paths so a
 * caller can pin both knobs without conflict.
 */
export function parseLegendFillColor(chartEl: XmlElement): string | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  return parseSpPrFill(legend);
}

/**
 * Pull `<c:legend><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:legend>` off the legend block.
 * Returns the line stroke color as a 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the line's
 * solid fill choice (`CT_LineProperties`, §20.1.2.3.24) which itself
 * sits inside `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13). The
 * `<c:spPr>` slot lives between `<c:overlay>` and `<c:txPr>` per
 * CT_Legend (ECMA-376 Part 1, §21.2.2.114).
 *
 * The reader surfaces only the literal `<a:srgbClr>` form — absence,
 * non-solid line fills (`<a:noFill>` / `<a:gradFill>` /
 * `<a:pattFill>`), and theme-color references (`<a:schemeClr>`) all
 * collapse to `undefined` so a chart that pinned a stroke the writer
 * cannot reproduce on emit drops the field rather than fabricate one
 * Excel would render differently. Malformed `val` tokens (wrong
 * length, non-hex characters, alpha-channel forms, non-string
 * escapes) likewise drop to `undefined`.
 *
 * The lookup is scoped to direct children of `<c:legend>` so a stray
 * `<c:spPr>` elsewhere (e.g. on a series, on an axis, on the legend's
 * `<c:txPr>` block) cannot leak in. Returns `undefined` whenever the
 * chart omits the `<c:legend>` element or the `<c:spPr><a:ln>
 * <a:solidFill><a:srgbClr>` chain is malformed at any link. Mirrors
 * the writer-side {@link SheetChart.legendBorderColor} so a parsed
 * value slots straight into {@link cloneChart} without conversion.
 *
 * Independent of {@link parseLegendFillColor}: the stroke lives on
 * `<c:legend><c:spPr><a:ln><a:solidFill>`, the fill lives on
 * `<c:legend><c:spPr><a:solidFill>` — the two readers walk disjoint
 * children of the same `<c:spPr>` block so a caller can pin both
 * knobs without conflict.
 */
export function parseLegendBorderColor(chartEl: XmlElement): string | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  return parseSpPrBorderColor(legend);
}

/**
 * Pull the `w` attribute off `<c:legend><c:spPr><a:ln w="EMU"/>` and
 * return the stroke width in points after clamping to the
 * `0.25..13.5` pt band Excel's UI exposes. The OOXML `w` attribute
 * carries the stroke width in English Metric Units (1 pt = 12 700 EMU)
 * per `CT_LineProperties` (ECMA-376 Part 1, §20.1.2.3.24); the reader
 * snaps the result to the 0.25 pt grid so a parsed-then-written width
 * does not drift across round-trips (Excel rounds in the UI anyway).
 *
 * Returns `undefined` when the chart omits `<c:legend>`, when the
 * legend has no `<c:spPr><a:ln w=..>` slot, when the attribute is
 * missing, when the value cannot be parsed as a finite positive
 * number, or when it parses to zero (Excel's "no border" marker — the
 * writer-side knob does not model that state). Mirrors the writer-side
 * {@link SheetChart.legendBorderWidth} so a parsed value slots
 * straight into {@link cloneChart} without conversion.
 *
 * The lookup is scoped to direct children of `<c:legend>` so a stray
 * `<a:ln w=..>` elsewhere (on a series stroke, on an axis line, on the
 * plot-area border) cannot leak in. Mirrors {@link parseLegendBorderColor} —
 * same `<c:spPr>` host on the same `<c:legend>` parent — but lands on
 * the `w` attribute rather than the `<a:solidFill><a:srgbClr>` color
 * child.
 */
export function parseLegendBorderWidth(chartEl: XmlElement): number | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  return parseBorderWidthFromSpPr(legend);
}

/**
 * Pull `<c:legend><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
 * </c:legend>` off a chart. Returns the recognized
 * {@link ChartBorderDash} value when the chain is present and the
 * value is a valid `ST_PresetLineDashVal` token other than the OOXML
 * default `"solid"`. Returns `undefined` when the chart omits
 * `<c:legend>`, when the chain is broken, or when the value matches
 * the default. Mirrors {@link parseTitleBorderDash} on a different
 * host element.
 */
export function parseLegendBorderDash(chartEl: XmlElement): ChartBorderDash | undefined {
  const legend = findChild(chartEl, "legend");
  if (!legend) return undefined;
  return parseBorderDashFromSpPr(legend);
}

// ── Writer ────────────────────────────────────────────────────────

export function resolveLegendPosition(chart: SheetChart): LegendPos | null {
  if (chart.legend === false) return null;
  if (chart.legend === undefined) {
    // Sensible defaults that match Excel's behaviour.
    return chart.type === "scatter" ? "b" : "r";
  }
  switch (chart.legend) {
    case "top":
      return "t";
    case "bottom":
      return "b";
    case "left":
      return "l";
    case "right":
      return "r";
    case "topRight":
      return "tr";
  }
}

export function buildLegend(
  pos: LegendPos,
  overlay: boolean,
  entries: readonly ResolvedLegendEntry[],
  fontSizePt: number | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
  underline: boolean | undefined,
  strikethrough: boolean | undefined,
  rgbHex: string | undefined,
  fontFamily: string | undefined,
  layout: ResolvedManualLayout | undefined,
  fillRgbHex: string | undefined,
  borderRgbHex: string | undefined,
  borderWidthPt: number | undefined,
  borderDash: ChartBorderDash | undefined,
): string {
  const children: string[] = [xmlSelfClose("c:legendPos", { val: pos })];

  // CT_Legend sequence places `<c:legendEntry>` after `<c:legendPos>`
  // and before `<c:layout>` / `<c:overlay>` (ECMA-376 Part 1,
  // §21.2.2.114). Each entry is emitted with both `<c:idx>` and
  // `<c:delete>` so a re-parse sees the canonical shape — Excel itself
  // emits `<c:delete val="1"/>` whenever the action is "Hide legend
  // entry", and the writer mirrors that even for the OOXML default
  // `false` value (an explicit `<c:delete val="0"/>` round-trips
  // cleanly through `parseChart`).
  for (const entry of entries) {
    children.push(
      xmlElement("c:legendEntry", undefined, [
        xmlSelfClose("c:idx", { val: entry.idx }),
        xmlSelfClose("c:delete", { val: entry.delete ? 1 : 0 }),
      ]),
    );
  }

  // CT_Legend sequence places `<c:layout>` between `<c:legendEntry>`
  // and `<c:overlay>` per ECMA-376 Part 1, §21.2.2.114. The writer
  // skips emission entirely when the caller pinned no coordinates so a
  // fresh chart matches Excel's reference serialization byte-for-byte
  // (Excel itself omits the block when the legend renders at the
  // auto-layout position). Each axis is independently optional so the
  // helper drops `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` slots whose
  // value did not survive normalization.
  const layoutXml = buildManualLayout(layout);
  if (layoutXml !== undefined) {
    children.push(layoutXml);
  }

  children.push(xmlSelfClose("c:overlay", { val: overlay ? 1 : 0 }));

  // CT_Legend sequence places `<c:spPr>` between `<c:overlay>` and
  // `<c:txPr>` (ECMA-376 Part 1, §21.2.2.114). The writer skips
  // emission entirely when the caller did not pin a fill / border
  // color so a fresh chart matches Excel's reference serialization
  // byte-for-byte — Excel itself omits the block whenever the legend
  // renders at the theme default fill / stroke (typically a transparent
  // legend background with no `<c:spPr>` block). The writer authors
  // `<a:solidFill>` (fill) and `<a:ln>` (stroke) here in
  // `CT_ShapeProperties` schema order; other `CT_ShapeProperties`
  // children (`<a:effectLst>` effects, gradient / pattern / picture
  // fills, line dash / width / compound styles) are not modelled at
  // this layer.
  const legendSpPrXml = buildLegendSpPr(fillRgbHex, borderRgbHex, borderWidthPt, borderDash);
  if (legendSpPrXml !== undefined) {
    children.push(legendSpPrXml);
  }

  // CT_Legend sequence places `<c:txPr>` after `<c:spPr>` (and before
  // the optional `<c:extLst>`). The writer skips emission entirely
  // when no typography knob is pinned so a fresh chart matches Excel's
  // reference serialization byte-for-byte (Excel itself omits the
  // block whenever the legend renders at the theme-default style).
  // The block currently carries the legend font size, bold, italic,
  // underline, strikethrough, font color, and font family knobs. The
  // `<a:bodyPr>` carries no rotation attribute — the legend is not
  // rotatable in Excel's UI, mirroring how the axis tick-label
  // `<c:txPr>` slot drops `rot` when only typography knobs are pinned.
  const txPrXml = buildLegendTxPr(
    fontSizePt,
    bold,
    italic,
    underline,
    strikethrough,
    rgbHex,
    fontFamily,
  );
  if (txPrXml !== undefined) {
    children.push(txPrXml);
  }

  return xmlElement("c:legend", undefined, children);
}

/**
 * Build the `<c:spPr>` element on `<c:legend>` that carries the
 * legend's background fill, border (line stroke) color, and border
 * width. Returns `undefined` when no knob is pinned so the caller can
 * elide the entire block — Excel's reference serialization omits
 * `<c:spPr>` from `<c:legend>` whenever the legend renders at the
 * theme default fill / stroke (typically a transparent legend
 * background with no `<c:spPr>` block).
 *
 * The emitted block mirrors the minimal `<c:spPr>` shape Excel writes
 * when the user pins "Format Legend -> Fill -> Solid fill -> Color"
 * and / or "Format Legend -> Border -> Solid line -> Color" / "Width":
 * `<c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
 * <a:ln w="EMU"><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
 * </a:ln></c:spPr>`. The `val` attribute holds the canonical
 * 6-character uppercase hex form (the writer normalizes the input
 * ahead of this call so a malformed source value never reaches emit).
 * The width attribute lands on `<a:ln>` (EMU; 1 pt = 12 700 EMU)
 * authored together with the border-color child so a stroke-only or
 * color-only legend still emits a single `<a:ln>` block. When at least
 * one knob lands on the wire, the children are emitted in
 * `CT_ShapeProperties` schema order: `<a:solidFill>` (fill) then
 * `<a:ln>` (line / stroke).
 *
 * Mirrors the chart-title / plot-area / axis-title `<c:spPr>` slots
 * so a single hex string threads cleanly through every fill / stroke
 * knob the writer authors.
 */
export function buildLegendSpPr(
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
 * Build the `<c:txPr>` block that carries the legend's typography pins
 * (currently the font size, bold, and italic flags). Returns
 * `undefined` when every input is unset so the caller can elide the
 * element entirely (Excel's reference serialization omits `<c:txPr>`
 * from `<c:legend>` when the legend renders at the theme-default
 * style).
 *
 * The emitted block mirrors the minimal `<c:txPr>` shape Excel writes
 * when the user pins a legend typography knob — `<a:bodyPr/>` (no
 * rotation because the legend is not rotatable), `<a:lstStyle/>` is
 * the empty list-style placeholder the schema requires, and the
 * `<a:p><a:pPr><a:defRPr sz="N" b=".." i=".."/></a:pPr><a:endParaRPr/>
 * </a:p>` paragraph stub Excel always emits hosts the typography
 * attributes on `<a:defRPr>`. Mirrors the chart-title / axis-title /
 * tick-label `<c:txPr>` slots exactly so a re-parse picks the values
 * off the canonical default-paragraph slot every other typography
 * reader expects.
 *
 * The bold / italic flags emit literal `b="1"` / `b="0"` / `i="1"` /
 * `i="0"` whenever the input is a boolean — `false` pins the OOXML
 * default explicitly, which is functionally identical to absence but
 * lets a clone target override an upstream `1` from a templated chart.
 */
export function buildLegendTxPr(
  fontSizePt: number | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined,
  underline: boolean | undefined,
  strikethrough: boolean | undefined,
  rgbHex: string | undefined,
  fontFamily: string | undefined,
): string | undefined {
  if (
    fontSizePt === undefined &&
    bold === undefined &&
    italic === undefined &&
    underline === undefined &&
    strikethrough === undefined &&
    rgbHex === undefined &&
    fontFamily === undefined
  )
    return undefined;
  const defRPrAttrs: Record<string, string | number> = {};
  if (fontSizePt !== undefined) defRPrAttrs.sz = fontSizePt * TITLE_FONT_SZ_PER_POINT;
  if (bold !== undefined) defRPrAttrs.b = bold ? 1 : 0;
  if (italic !== undefined) defRPrAttrs.i = italic ? 1 : 0;
  if (underline !== undefined) defRPrAttrs.u = underline ? "sng" : "none";
  // Strikethrough rides as `strike="sngStrike"` on the same
  // `<a:defRPr>` slot. Absence collapses to omitting the attribute
  // entirely (the OOXML default `"noStrike"` is functionally identical
  // to absence — the reader collapses both to `undefined`). The writer
  // never emits `"noStrike"` or `"dblStrike"` so the surfaced shape
  // stays consistent with Excel's UI checkbox.
  if (strikethrough === true) defRPrAttrs.strike = "sngStrike";
  // OOXML's `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:defRPr>` carries the legend font color.
  // Absence (`undefined`) collapses to omitting the entire
  // `<a:solidFill>` block so the legend inherits the theme text
  // color (Excel's reference behavior for a fresh legend that has
  // not had a custom font color picked).
  const solidFillChild = rgbHex
    ? xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: rgbHex })])
    : undefined;
  // OOXML's `<a:defRPr><a:latin typeface=".."/></a:defRPr>` carries
  // the legend font family. The `<a:latin>` element follows
  // `<a:solidFill>` per the CT_TextCharacterProperties child sequence
  // (ECMA-376 Part 1, §21.1.2.3.7). Absence (`undefined`) collapses
  // to omitting the entire `<a:latin>` element so the legend inherits
  // the theme typeface (Excel's reference behavior for a fresh legend
  // that has not had a custom font picked).
  const latinChild = fontFamily ? xmlSelfClose("a:latin", { typeface: fontFamily }) : undefined;
  // When a fill color or a typeface is set the `<a:defRPr>` slot
  // expands from self-closing to wrapping the children; otherwise the
  // writer keeps the existing self-closing form so a fresh legend
  // with no custom color or font matches Excel's reference
  // serialization byte-for-byte. Children are emitted in
  // CT_TextCharacterProperties' canonical schema order: solidFill
  // first, then latin.
  const defRPrChildren: string[] = [];
  if (solidFillChild) defRPrChildren.push(solidFillChild);
  if (latinChild) defRPrChildren.push(latinChild);
  const defRPr =
    defRPrChildren.length > 0
      ? xmlElement("a:defRPr", defRPrAttrs, defRPrChildren)
      : xmlSelfClose("a:defRPr", defRPrAttrs);
  return xmlElement("c:txPr", undefined, [
    xmlSelfClose("a:bodyPr"),
    xmlSelfClose("a:lstStyle"),
    xmlElement("a:p", undefined, [
      xmlElement("a:pPr", undefined, [defRPr]),
      xmlSelfClose("a:endParaRPr", { lang: "en-US" }),
    ]),
  ]);
}

/**
 * Resolve `<c:legend><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
 * </a:p></c:txPr></c:legend>` from {@link SheetChart.legendFontSize}.
 *
 * Returns the size in points (`1..400`), or `undefined` when the chart
 * leaves the field unset / passed an out-of-range or non-numeric token.
 * The flag is only meaningful when the chart actually emits a legend —
 * the caller is expected to gate the call on the resolved legend
 * visibility (`resolveLegendPosition` returning a non-null value), so a
 * chart that hides the legend silently drops the value rather than
 * emit a `<c:txPr>` block with no on-screen effect.
 *
 * Mirrors `resolveTitleFontSize` / `resolveAxisTitleFontSize` /
 * `resolveAxisLabelFontSize` exactly so a single configuration call
 * threads cleanly through every typography slot Excel exposes.
 */
export function resolveLegendFontSize(chart: SheetChart): number | undefined {
  return normalizeTitleFontSize(chart.legendFontSize);
}

/**
 * Resolve `<c:legend><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
 * </a:p></c:txPr></c:legend>` from {@link SheetChart.legendBold}.
 *
 * Returns the bold flag, or `undefined` when the chart leaves the
 * field unset / passed a non-boolean token. The flag is only
 * meaningful when the chart actually emits a legend — the caller is
 * expected to gate the call on the resolved legend visibility
 * (`resolveLegendPosition` returning a non-null value), so a chart
 * that hides the legend silently drops the value rather than emit a
 * `<c:txPr>` block with no on-screen effect.
 *
 * Mirrors `resolveTitleBold` / `resolveAxisTitleBold` /
 * `resolveAxisLabelBold` exactly — only literal `true` / `false` pass
 * through; non-boolean tokens collapse to `undefined` so the writer
 * drops the slot rather than emit a value Excel would reject.
 */
export function resolveLegendBold(chart: SheetChart): boolean | undefined {
  const value = chart.legendBold;
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:legend><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:txPr></c:legend>` from {@link SheetChart.legendItalic}.
 *
 * Returns the italic flag, or `undefined` when the chart leaves the
 * field unset / passed a non-boolean token. The flag is only
 * meaningful when the chart actually emits a legend — the caller is
 * expected to gate the call on the resolved legend visibility.
 *
 * Mirrors `resolveTitleItalic` / `resolveAxisTitleItalic` /
 * `resolveAxisLabelItalic` exactly — only literal `true` / `false`
 * pass through; non-boolean tokens collapse to `undefined`.
 */
export function resolveLegendItalic(chart: SheetChart): boolean | undefined {
  const value = chart.legendItalic;
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:legend><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
 * </a:p></c:txPr></c:legend>` from {@link SheetChart.legendUnderline}.
 *
 * Returns the underline flag, or `undefined` when the chart leaves
 * the field unset / passed a non-boolean token. The flag is only
 * meaningful when the chart actually emits a legend — the caller is
 * expected to gate the call on the resolved legend visibility.
 *
 * Mirrors `resolveTitleUnderline` / `resolveAxisTitleUnderline` /
 * `resolveAxisLabelUnderline` exactly — only literal `true` / `false`
 * pass through; non-boolean tokens collapse to `undefined`. The
 * writer translates `true` into `u="sng"` (Excel's UI variant —
 * single underline) and `false` into `u="none"` at emit time.
 */
export function resolveLegendUnderline(chart: SheetChart): boolean | undefined {
  const value = chart.legendUnderline;
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:legend><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:txPr></c:legend>` from
 * {@link SheetChart.legendStrikethrough}.
 *
 * Returns `true` when the chart pins the strikethrough flag literally;
 * every other value (explicit `false`, absence, non-boolean tokens
 * leaking past the type guard) collapses to `undefined` so the writer
 * never emits a `strike` attribute below `"sngStrike"`. The OOXML
 * default `"noStrike"` is functionally identical to absence — the
 * writer keeps the surfaced shape consistent with what Excel's UI
 * authors (`"sngStrike"` only, never `"noStrike"` or `"dblStrike"`),
 * mirroring how `resolveTitleStrike` lands on the title's `<a:defRPr>`.
 *
 * Collapsing `false` to `undefined` (instead of the `boolean`-pass-through
 * shape that `legendBold` / `legendItalic` / `legendUnderline` use)
 * mirrors how the chart-title / axis-title / axis tick-label
 * strikethrough writers also drop the attribute on `false`: a
 * standalone `legendStrikethrough: false` does not gratuitously
 * trigger the `<c:txPr>` block emission, so the legend stays at the
 * theme-default style and a fresh chart with no custom strikethrough
 * matches Excel's reference serialization byte-for-byte. The flag is
 * only meaningful when the chart actually emits a legend — the caller
 * is expected to gate the call on the resolved legend visibility.
 */
export function resolveLegendStrikethrough(chart: SheetChart): boolean | undefined {
  if (chart.legendStrikethrough === true) return true;
  return undefined;
}

/**
 * Resolve `<c:legend><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:txPr></c:legend>` from {@link SheetChart.legendFontColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches the chart-title /
 * axis-title / axis tick-label color resolvers exactly. The fill is
 * only meaningful when the chart actually emits a legend — the
 * caller is expected to gate the call on the resolved legend
 * visibility. A chart whose legend is suppressed has no `<c:txPr>`
 * slot to host the fill in either case.
 */
export function resolveLegendFontColor(chart: SheetChart): string | undefined {
  return normalizeTitleColor(chart.legendFontColor);
}

/**
 * Normalize a {@link SheetChart.legendFontFamily} value for the
 * `<c:legend><c:txPr><a:p><a:pPr><a:defRPr><a:latin typeface=".."/>
 * </a:defRPr></a:pPr></a:p></c:txPr></c:legend>` writer slot. Returns
 * the trimmed typeface string when the input is a non-empty string,
 * or `undefined` for any malformed token — empty / whitespace-only
 * strings, or non-string escapes from an untyped caller (`null`,
 * numbers, booleans, etc.).
 *
 * Absence and malformed tokens both collapse to `undefined` so the
 * writer skips the entire `<a:latin>` element and the legend inherits
 * the theme typeface (Excel's reference behavior for a fresh legend
 * without a custom font picked).
 */
export function normalizeLegendFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Resolve `<c:legend><c:txPr><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:txPr></c:legend>` from
 * {@link SheetChart.legendFontFamily}.
 *
 * Returns the trimmed typeface string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed an empty
 * or non-string token. Delegates to {@link normalizeLegendFontFamily}
 * so the accept-and-trim grammar matches the chart-title /
 * axis-title / axis tick-label font family resolvers exactly. The
 * element is only meaningful when the chart actually emits a legend —
 * the caller is expected to gate the call on the resolved legend
 * visibility. A chart whose legend is suppressed has no `<c:txPr>`
 * slot to host the typeface in either case.
 */
export function resolveLegendFontFamily(chart: SheetChart): string | undefined {
  return normalizeLegendFontFamily(chart.legendFontFamily);
}

/**
 * Normalize {@link SheetChart.legendEntries} into an emit-ready list.
 *
 * The OOXML schema (`CT_LegendEntry`) places `<c:idx val="N"/>` as the
 * required selector and `<c:delete val=".."/>` as the hide flag. Hucre
 * accepts a free-form `ChartLegendEntry[]` from callers; this helper
 * strips entries whose `idx` cannot land on a real series and
 * deduplicates duplicate `idx` values so the writer never emits the
 * same selector twice (the last entry wins so a clone-through that
 * appends an override naturally beats the source's parsed value).
 *
 * Validation rules:
 *   - `idx` must be a non-negative integer (matches `xsd:unsignedInt`
 *     on `<c:idx val=".."/>`); non-finite, negative, or non-integer
 *     values drop entirely rather than emit a token Excel rejects.
 *   - `delete` collapses to a strict boolean — anything other than
 *     literal `true` resolves to `false`. Mirrors how `legendOverlay`
 *     / `roundedCorners` / `plotVisOnly` treat their inputs.
 *
 * Entries are emitted in ascending `idx` order so the rendered chart
 * matches Excel's reference serialization (Excel sorts by `<c:idx>` on
 * write even when the in-memory list arrived unsorted). Returns an
 * empty array when the chart has no entries to emit so the caller can
 * avoid touching the legend block.
 */
export function resolveLegendEntries(chart: SheetChart): ResolvedLegendEntry[] {
  const raw = chart.legendEntries;
  if (!Array.isArray(raw) || raw.length === 0) return [];

  const byIdx = new Map<number, ResolvedLegendEntry>();
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") continue;
    const idx = entry.idx;
    if (typeof idx !== "number" || !Number.isFinite(idx)) continue;
    if (!Number.isInteger(idx) || idx < 0) continue;
    byIdx.set(idx, { idx, delete: entry.delete === true });
  }

  return Array.from(byIdx.values()).sort((a, b) => a.idx - b.idx);
}

/**
 * Resolve `<c:legend><c:overlay val=".."/></c:legend>` from
 * {@link SheetChart.legendOverlay}.
 *
 * Defaults to `false` (the OOXML default Excel itself emits — the
 * legend reserves its own slot and the plot area shrinks to make room).
 * Anything other than literal `true` collapses to `false` so a stray
 * non-boolean leaking through the type guard (e.g. `0` / `1` / `"true"`
 * / `null`) never produces `<c:overlay val="1"/>`. This matches how
 * `roundedCorners` / `plotVisOnly` / axis `hidden` treat their inputs:
 * a literal boolean is the only path to a non-default value.
 *
 * The writer always emits `<c:overlay>` because Excel's reference
 * serialization includes the element on every visible legend; only the
 * `val` flips when the caller pins `legendOverlay: true`.
 */
export function resolveLegendOverlay(chart: SheetChart): boolean {
  return chart.legendOverlay === true;
}

/**
 * Resolve `<c:legend><c:layout><c:manualLayout>...</c:manualLayout>
 * </c:layout></c:legend>` from {@link SheetChart.legendLayout}.
 *
 * Returns the normalized coordinate set, or `undefined` when every
 * axis the caller pinned dropped to `undefined` (so the writer can
 * elide the entire `<c:layout>` block — Excel's reference
 * serialization omits the element when the legend renders at the
 * auto-layout position). The element is only meaningful when the chart
 * actually emits a legend — the caller is expected to gate the call on
 * the resolved legend visibility.
 *
 * Coordinates outside the OOXML `0..1` band, `NaN`, `Infinity`, and
 * non-numeric inputs all collapse to `undefined` on the matching axis
 * so the writer drops the matching `<c:x>` / `<c:y>` / `<c:w>` /
 * `<c:h>` slot rather than emit a token Excel would reject.
 */
export function resolveLegendLayout(chart: SheetChart): ResolvedManualLayout | undefined {
  return normalizeManualLayout(chart.legendLayout);
}

/**
 * Resolve `<c:legend><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:legend>` from
 * {@link SheetChart.legendFillColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches the chart-title /
 * axis-title / axis tick-label / legend font color resolvers exactly.
 * The fill is only meaningful when the chart actually emits a
 * legend — the caller is expected to gate the call on the resolved
 * legend visibility (`resolveLegendPosition` returning a non-null
 * value). A chart whose legend is suppressed has no `<c:legend>`
 * block to host the `<c:spPr>` slot in either case.
 *
 * Independent of {@link resolveLegendFontColor}: the fill lands on
 * `<c:legend><c:spPr>`, the font color lands on
 * `<c:legend><c:txPr>` — the two resolvers target different children
 * of `<c:legend>` so a single configuration call can pin both.
 */
export function resolveLegendFillColor(chart: SheetChart): string | undefined {
  return normalizeTitleColor(chart.legendFillColor);
}

/**
 * Resolve `<c:legend><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:legend>` from
 * {@link SheetChart.legendBorderColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the chart leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches every other `<a:srgbClr>`
 * fill / line slot exactly. The stroke is only meaningful when the
 * chart actually emits a legend — the caller is expected to gate the
 * call on the resolved legend visibility (`resolveLegendPosition`
 * returning a non-null value). A chart whose legend is suppressed has
 * no `<c:legend>` block to host the `<c:spPr>` slot in either case.
 *
 * Independent of {@link resolveLegendFillColor}: the fill lands on
 * `<c:legend><c:spPr><a:solidFill>`, the stroke lands on
 * `<c:legend><c:spPr><a:ln><a:solidFill>` — the two resolvers feed
 * disjoint children of `<c:spPr>` so a single configuration call can
 * pin both. Mirrors the chart-title / axis-title / chart-space /
 * plot-area `<c:spPr>` slots — same hex grammar, same `<a:ln>` slot
 * on the `CT_ShapeProperties` schema — but lands on `<c:legend>`'s
 * own `<c:spPr>` block.
 */
export function resolveLegendBorderColor(chart: SheetChart): string | undefined {
  return normalizeTitleColor(chart.legendBorderColor);
}

/**
 * Resolve `<c:legend><c:spPr><a:ln w="EMU"/></c:spPr></c:legend>` from
 * {@link SheetChart.legendBorderWidth}.
 *
 * Returns the point value clamped to the `0.25..13.5` pt band Excel's
 * UI exposes and snapped to the 0.25 pt grid, or `undefined` when the
 * chart leaves the field unset / passed a malformed token (`NaN`,
 * `Infinity`, non-finite). Delegates to {@link clampStrokeWidthPt} so
 * the snap / clamp grammar matches every other `<a:ln w=..>` slot the
 * writer authors (the series stroke knob `series[i].stroke.width`,
 * the plot-area border width knob {@link SheetChart.plotAreaBorderWidth}).
 * The width is only meaningful when the chart actually emits a
 * legend — the caller is expected to gate the call on the resolved
 * legend visibility (`resolveLegendPosition` returning a non-null
 * value). A chart whose legend is suppressed has no `<c:legend>`
 * block to host the `<c:spPr>` slot in either case.
 *
 * Independent of {@link resolveLegendBorderColor}: both knobs land on
 * the same `<a:ln>` element but on a different slot (the color child
 * `<a:solidFill>` versus the line's `w` attribute). Mirrors the
 * chart-title / axis-title / chart-space / plot-area `<c:spPr>` slots —
 * same EMU encoding, same `<a:ln>` host — but lands on `<c:legend>`'s
 * own `<c:spPr>` block.
 */
export function resolveLegendBorderWidth(chart: SheetChart): number | undefined {
  return clampStrokeWidthPt(chart.legendBorderWidth);
}

/**
 * Resolve `<c:legend><c:spPr><a:ln><a:prstDash val=".."/></a:ln></c:spPr>
 * </c:legend>` from {@link SheetChart.legendBorderDash}.
 *
 * Returns the recognized {@link ChartBorderDash} value, or `undefined`
 * for the OOXML default `"solid"` and every unrecognized token —
 * delegates to {@link normalizeBorderDash} so absence and the OOXML
 * default round-trip identically. The dash is only meaningful when the
 * chart actually emits a legend; the caller gates this on the resolved
 * legend visibility.
 *
 * Independent of {@link resolveLegendBorderColor} and
 * {@link resolveLegendBorderWidth}: all three knobs land on the same
 * `<a:ln>` element but on different children / attributes.
 */
export function resolveLegendBorderDash(chart: SheetChart): ChartBorderDash | undefined {
  return normalizeBorderDash(chart.legendBorderDash);
}

// ── Clone normalizers ─────────────────────────────────────────────

/**
 * Normalize a `legendBold` value for the cloned `SheetChart`. Mirrors
 * the writer's `resolveLegendBold` — the cloned shape is guaranteed to
 * round-trip through the writer without surprise: `true` / `false`
 * pass through literally, every other token (typed escape from an
 * untyped caller) collapses to `undefined` so the cloned chart drops
 * the field rather than carry a value the writer would silently elide
 * back to absence.
 */
export function normalizeLegendBold(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Normalize a `legendItalic` value for the cloned `SheetChart`. Mirrors
 * the writer's `resolveLegendItalic` — `true` / `false` pass through
 * literally, every other token (typed escape from an untyped caller)
 * collapses to `undefined`.
 */
export function normalizeLegendItalic(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Normalize a `legendUnderline` value for the cloned `SheetChart`.
 * Mirrors the writer's `resolveLegendUnderline` — `true` / `false`
 * pass through literally, every other token collapses to `undefined`.
 */
export function normalizeLegendUnderline(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Normalize a `legendStrikethrough` value for the cloned `SheetChart`.
 * Mirrors the writer's `resolveLegendStrikethrough` — `true` / `false`
 * pass through literally, every other token collapses to `undefined`.
 *
 * The cloned `SheetChart` retains a literal `false` (the writer drops
 * `false` to absence at emit time, so pinning `false` on the cloned
 * chart is functionally identical to omission, but it lets a downstream
 * consumer that re-clones the chart distinguish "explicit no-strike
 * pin" from "field never set"). The chart-title / axis-title / axis
 * tick-label strike clone resolvers use the same shape — only at the
 * writer's `<a:defRPr>` slot does the `false` collapse to attribute
 * omission.
 */
export function normalizeLegendStrikethrough(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Normalize a {@link ChartManualLayout} for the cloned `SheetChart`.
 * Drops every axis whose input is non-numeric / non-finite / outside
 * the `0..1` band; returns `undefined` when every axis dropped so the
 * cloned shape elides the field entirely (mirrors the writer-side
 * normalization so a parsed value flows through {@link cloneChart}
 * without bookkeeping the units). Coordinates outside the `0..1` band
 * collapse rather than clamp — same accept-or-drop grammar as
 * `titleFontSize` / `axisTitleFontSize` / `legendFontSize`.
 */
export function normalizeLegendLayout(
  value: ChartManualLayout | undefined,
): ChartManualLayout | undefined {
  return normalizeChartManualLayout(value);
}

/**
 * Normalize a `legendBorderWidth` value for the cloned `SheetChart`.
 * Mirrors the writer's `clampStrokeWidthPt` — values are clamped to the
 * `0.25..13.5` pt band Excel's UI exposes and snapped to the 0.25 pt
 * grid so a parsed-then-cloned-then-written width does not drift across
 * round-trips (Excel rounds in the UI anyway). Non-finite / non-numeric
 * tokens (`NaN`, `Infinity`, strings, `null` from an untyped caller)
 * collapse to `undefined` so the cloned chart drops the field rather
 * than carry a value the writer would silently elide back to absence.
 */
export function normalizeLegendBorderWidth(value: number | undefined): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const snapped = Math.round(value * 4) / 4;
  if (snapped < LEGEND_BORDER_WIDTH_MIN_PT) return LEGEND_BORDER_WIDTH_MIN_PT;
  if (snapped > LEGEND_BORDER_WIDTH_MAX_PT) return LEGEND_BORDER_WIDTH_MAX_PT;
  return snapped;
}

// ── Clone resolvers (3-arg source/override) ───────────────────────

/**
 * Resolve a `legendOverlay` override.
 *
 * `undefined` → inherit the source's parsed `legendOverlay`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — the legend reserves its own
 *               slot, no overlap with the plot area).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `roundedCorners` so the chart-
 * level toggles compose the same way at the call site. Callers should
 * gate the result on the resolved legend visibility — when no legend
 * is emitted, the overlay flag has no slot in the rendered chart.
 */
export function resolveCloneLegendOverlay(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `legendEntries` override.
 *
 * `undefined` → inherit the source's parsed `legendEntries`.
 * `null`      → drop the inherited list (the writer emits no
 *               `<c:legendEntry>` children).
 * `array`     → replace the inherited list outright. Empty arrays
 *               collapse to `undefined` so the writer never emits an
 *               empty selector block — Excel's reference serialization
 *               omits the children entirely when no entry is hidden.
 *
 * Callers should gate the result on the resolved legend visibility —
 * when no legend is emitted, the entry list has no slot in the rendered
 * chart. Mirrors the `legendOverlay` grammar so the legend-scoped
 * fields compose the same way at the call site.
 *
 * The returned array is always a fresh copy of the source / override
 * (never a shared reference) so a downstream mutation to the cloned
 * `SheetChart` never leaks back into the parsed `Chart` the caller
 * passed in. Each entry is also copied to keep the writer's resolution
 * pass free to dedupe / sort without touching the inputs.
 */
export function resolveCloneLegendEntries(
  sourceValue: ChartLegendEntry[] | undefined,
  override: ChartLegendEntry[] | null | undefined,
): ChartLegendEntry[] | undefined {
  if (override === undefined) {
    if (!sourceValue || sourceValue.length === 0) return undefined;
    return sourceValue.map((entry) => ({ ...entry }));
  }
  if (override === null) return undefined;
  if (!Array.isArray(override) || override.length === 0) return undefined;
  return override.map((entry) => ({ ...entry }));
}

/**
 * Resolve a `legendFontSize` override.
 *
 * `undefined` → inherit the source's parsed `legendFontSize` (after
 *               running it through {@link normalizeTitleFontSize} so
 *               an out-of-range parsed value drops cleanly).
 * `null`      → drop the inherited value (the writer falls back to
 *               Excel's theme-default 9pt — no `<c:txPr>` block on
 *               the legend).
 * `number`    → replace, after clamping / rounding through
 *               {@link normalizeTitleFontSize}.
 *
 * The grammar mirrors `titleFontSize` / `axisTitleFontSize` /
 * `axes.x.labelFontSize` so the typography knobs compose the same way
 * at the call site. Callers should gate the result on the resolved
 * legend visibility — when no legend is emitted, the size has no slot
 * in the rendered chart.
 */
export function resolveCloneLegendFontSize(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeTitleFontSize(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleFontSize(override);
}

/**
 * Resolve a `legendBold` override.
 *
 * `undefined` → inherit the source's parsed `legendBold` (after
 *               running it through {@link normalizeLegendBold} so a
 *               typed escape on the source path drops cleanly).
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default — no `b` attribute, equivalent to
 *               non-bold).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleBold` / `axisTitleBold` /
 * `axes.x.labelBold` so the typography knobs compose the same way at
 * the call site. Callers should gate the result on the resolved legend
 * visibility — when no legend is emitted, the flag has no slot in the
 * rendered chart.
 */
export function resolveCloneLegendBold(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeLegendBold(sourceValue);
  if (override === null) return undefined;
  return normalizeLegendBold(override);
}

/**
 * Resolve a `legendItalic` override.
 *
 * `undefined` → inherit the source's parsed `legendItalic` (after
 *               running it through {@link normalizeLegendItalic} so a
 *               typed escape on the source path drops cleanly).
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default — no `i` attribute, equivalent to
 *               non-italic).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleItalic` / `axisTitleItalic` /
 * `axes.x.labelItalic` so the typography knobs compose the same way at
 * the call site. Callers should gate the result on the resolved legend
 * visibility — when no legend is emitted, the flag has no slot in the
 * rendered chart.
 */
export function resolveCloneLegendItalic(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeLegendItalic(sourceValue);
  if (override === null) return undefined;
  return normalizeLegendItalic(override);
}

/**
 * Resolve a `legendUnderline` override.
 *
 * `undefined` → inherit the source's parsed `legendUnderline` (after
 *               running it through {@link normalizeLegendUnderline}
 *               so a typed escape on the source path drops cleanly).
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default — no `u` attribute, equivalent to
 *               non-underlined).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleUnderline` / `axisTitleUnderline` /
 * `axes.x.labelUnderline` so the typography knobs compose the same way
 * at the call site. Callers should gate the result on the resolved
 * legend visibility — when no legend is emitted, the flag has no slot
 * in the rendered chart.
 */
export function resolveCloneLegendUnderline(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeLegendUnderline(sourceValue);
  if (override === null) return undefined;
  return normalizeLegendUnderline(override);
}

/**
 * Resolve a `legendStrikethrough` override.
 *
 * `undefined` → inherit the source's parsed `legendStrikethrough`
 *               (after running it through
 *               {@link normalizeLegendStrikethrough} so a typed escape
 *               on the source path drops cleanly).
 * `null`      → drop the inherited flag (the writer falls back to the
 *               OOXML default — no `strike` attribute, equivalent to
 *               non-strikethrough).
 * `boolean`   → replace.
 *
 * The grammar mirrors `titleStrikethrough` / `axisTitleStrike` /
 * `axes.x.labelStrikethrough` so the typography knobs compose the same
 * way at the call site. Callers should gate the result on the resolved
 * legend visibility — when no legend is emitted, the flag has no slot
 * in the rendered chart.
 */
export function resolveCloneLegendStrikethrough(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return normalizeLegendStrikethrough(sourceValue);
  if (override === null) return undefined;
  return normalizeLegendStrikethrough(override);
}

/**
 * Resolve a `legendFontColor` override.
 *
 * `undefined` → inherit the source's parsed `legendFontColor` (after
 *               running it through {@link normalizeTitleColor} so a
 *               malformed source value drops cleanly).
 * `null`      → drop the inherited fill (the writer falls back to the
 *               theme text color — no `<a:solidFill>` block on the
 *               legend's `<a:defRPr>`).
 * `string`    → replace, after running through
 *               {@link normalizeTitleColor} so the override accepts
 *               `"FF0000"` / `"#FF0000"` / `"ff0000"` and collapses
 *               malformed tokens to `undefined`.
 *
 * The grammar mirrors `titleColor` / `axisTitleColor` /
 * `axes.x.labelColor` so the typography knobs compose the same way at
 * the call site. Callers should gate the result on the resolved
 * legend visibility — when no legend is emitted, the fill has no slot
 * in the rendered chart.
 */
export function resolveCloneLegendFontColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeTitleColor(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleColor(override);
}

/**
 * Resolve a `legendFontFamily` override.
 *
 * `undefined` → inherit the source's parsed `legendFontFamily` (after
 *               running it through {@link normalizeLegendFontFamily}
 *               so a malformed source value drops cleanly).
 * `null`      → drop the inherited typeface (the writer falls back to
 *               the theme typeface — no `<a:latin>` element on the
 *               legend's `<a:defRPr>`).
 * `string`    → replace, after running through
 *               {@link normalizeLegendFontFamily} so the override
 *               accepts any caller spelling that the writer will
 *               accept (with surrounding whitespace trimmed; empty /
 *               whitespace-only strings collapse to a drop).
 *
 * The grammar mirrors `titleFontFamily` /
 * `axes.x.axisTitleFontFamily` / `axes.x.labelFontFamily` so the
 * typography knobs compose the same way at the call site. Callers
 * should gate the result on the resolved legend visibility — when no
 * legend is emitted, the typeface has no slot in the rendered chart.
 */
export function resolveCloneLegendFontFamily(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeLegendFontFamily(sourceValue);
  if (override === null) return undefined;
  return normalizeLegendFontFamily(override);
}

/**
 * Resolve a `legendLayout` override.
 *
 * `undefined` → inherit the source's parsed `legendLayout` (after
 *               running it through {@link normalizeLegendLayout} so a
 *               malformed source value drops cleanly).
 * `null`      → drop the inherited layout (the writer falls back to
 *               Excel's auto-layout position — no `<c:layout>` block
 *               on the legend).
 * `ChartManualLayout` → replace, after running through
 *               {@link normalizeLegendLayout}. Coordinates outside the
 *               `0..1` band collapse on the matching axis so the
 *               cloned `SheetChart` always carries a value the writer
 *               will accept; an override whose every axis dropped
 *               collapses to `undefined` so the writer skips the
 *               `<c:layout>` block entirely.
 *
 * The grammar mirrors `legendOverlay` / `legendEntries` /
 * `legendFontSize` so the legend knobs compose the same way at the
 * call site. Callers should gate the result on the resolved legend
 * visibility — when no legend is emitted, the layout has no slot in
 * the rendered chart.
 */
export function resolveCloneLegendLayout(
  sourceValue: ChartManualLayout | undefined,
  override: ChartManualLayout | null | undefined,
): ChartManualLayout | undefined {
  if (override === undefined) return normalizeLegendLayout(sourceValue);
  if (override === null) return undefined;
  return normalizeLegendLayout(override);
}

/**
 * Resolve a `legendFillColor` override.
 *
 * `undefined` → inherit the source's parsed `legendFillColor` (after
 *               running it through {@link normalizeTitleColor} so a
 *               malformed source value drops cleanly — the hex
 *               normalizer is purely shape-based and applies
 *               identically to every `<a:srgbClr val="RRGGBB"/>`
 *               slot).
 * `null`      → drop the inherited fill (the writer emits no
 *               `<c:spPr>` block on `<c:legend>`, falling back to the
 *               theme default — typically a transparent legend
 *               background).
 * `string`    → replace, after running through
 *               {@link normalizeTitleColor} so the override accepts
 *               `"FF0000"` / `"#FF0000"` / `"ff0000"` and collapses
 *               malformed tokens to `undefined`.
 *
 * The grammar mirrors `plotAreaFillColor` / `titleColor` /
 * `axisTitleColor` / `legendFontColor` so the fill / color knobs
 * compose the same way at the call site. Callers should gate the
 * result on the resolved legend visibility — when no legend is
 * emitted, the fill has no slot in the rendered chart.
 *
 * Independent of `legendFontColor`: the two knobs target different
 * children of `<c:legend>` (`<c:spPr>` for the background fill,
 * `<c:txPr>` for the font color), so a caller can pin both without
 * conflict.
 */
export function resolveCloneLegendFillColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeTitleColor(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleColor(override);
}

/**
 * Resolve a `legendBorderColor` override.
 *
 * `undefined` → inherit the source's parsed `legendBorderColor` (after
 *               running it through {@link normalizeTitleColor} so a
 *               malformed source value drops cleanly — the hex
 *               normalizer is purely shape-based and applies
 *               identically to every `<a:srgbClr val="RRGGBB"/>`
 *               slot).
 * `null`      → drop the inherited stroke (the writer emits no
 *               `<a:ln>` block on `<c:legend><c:spPr>`, the legend
 *               inherits the auto-stroke Excel picks from the chart's
 *               theme).
 * `string`    → replace with the normalized 6-character uppercase hex
 *               form. Malformed overrides collapse to `undefined` via
 *               the normalizer so the cloned `SheetChart` always
 *               carries a value the writer will accept.
 *
 * The grammar mirrors `legendFillColor` so the legend `<c:spPr>` knobs
 * compose the same way at the call site. Callers should gate the
 * result on the resolved legend visibility — when no legend is
 * emitted, the stroke has no slot in the rendered chart.
 *
 * Independent of `legendFillColor`: the two knobs target different
 * children of `<c:legend><c:spPr>` (`<a:solidFill>` for fill,
 * `<a:ln>` for stroke), so a caller can pin both without conflict.
 */
export function resolveCloneLegendBorderColor(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return normalizeTitleColor(sourceValue);
  if (override === null) return undefined;
  return normalizeTitleColor(override);
}

/**
 * Resolve a `legendBorderWidth` override.
 *
 * `undefined` → inherit the source's parsed `legendBorderWidth` (after
 *               running it through {@link normalizeLegendBorderWidth}
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
 * The grammar mirrors `plotAreaBorderWidth` / the series-line stroke
 * width so the chart `<a:ln w=..>` knobs compose the same way at the
 * call site. Callers should gate the result on the resolved legend
 * visibility — when no legend is emitted, the width has no slot in the
 * rendered chart.
 *
 * Independent of `legendBorderColor`: both knobs land on the same
 * `<a:ln>` element but on a different slot (color is
 * `<a:solidFill><a:srgbClr>`, width is the `w` attribute on `<a:ln>`),
 * so a caller can pin both without conflict.
 */
export function resolveCloneLegendBorderWidth(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return normalizeLegendBorderWidth(sourceValue);
  if (override === null) return undefined;
  return normalizeLegendBorderWidth(override);
}
