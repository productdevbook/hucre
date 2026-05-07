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

import type { ChartBorderDash, ChartLineDashStyle } from "./types";
import type { XmlElement } from "../../xml/parser";

/**
 * Local copy of `findChild`. The xml/parser module does not export the
 * helper, and the chart-reader / chart-clone files each carry their
 * own copy. Keeping a local definition here avoids a cross-module
 * refactor of the parser surface while still letting these primitives
 * walk a chart subtree.
 */
function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

// ── Stroke width ──────────────────────────────────────────────────
//
// Excel's UI exposes stroke widths on a `0.25..13.5` pt band, snapped
// to the 0.25 pt grid. The OOXML wire encodes the same value in
// English Metric Units (1 pt = 12 700 EMU) per CT_LineProperties
// (ECMA-376 Part 1, §20.1.2.3.24).

/** Smallest stroke width Excel's UI exposes, in points. */
export const STROKE_WIDTH_MIN_PT = 0.25;
/** Largest stroke width Excel's UI exposes, in points. */
export const STROKE_WIDTH_MAX_PT = 13.5;
/** Conversion factor between OOXML EMU and points. */
export const EMU_PER_PT = 12700;

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
]);

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
]);

// ── Hex normalization ─────────────────────────────────────────────

/**
 * Normalize an `<a:srgbClr val=".."/>` token into a 6-character
 * uppercase hex string. Accepts the input with or without a leading
 * `#`, and tolerates leading / trailing whitespace. Returns
 * `undefined` for any malformed input — wrong length, non-hex
 * characters, alpha-channel forms, or non-string tokens.
 */
export function normalizeRgbHex(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const hex = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed;
  if (hex.length !== 6) return undefined;
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  return hex.toUpperCase();
}

// ── Generic spPr readers ──────────────────────────────────────────

/**
 * Pull `<c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>` off the
 * supplied parent element (one of `<c:title>` / `<c:legend>` /
 * `<c:plotArea>` / `<c:chartSpace>` / `<c:dTable>` / `<c:dLbls>` /
 * a series — anywhere a `CT_ShapeProperties` block can live).
 * Returns the fill color as a 6-character uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the fill
 * choice of `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13).
 *
 * Surfaces only the literal `<a:srgbClr>` form — absence, non-solid
 * fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
 * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
 * collapse to `undefined` so a chart that pinned a fill the writer
 * cannot reproduce on emit drops the field rather than fabricate one
 * Excel would render differently. Malformed `val` tokens (wrong
 * length, non-hex characters, alpha-channel forms, non-string
 * escapes) likewise drop to `undefined`.
 */
export function parseSpPrFill(parent: XmlElement): string | undefined {
  const spPr = findChild(parent, "spPr");
  if (!spPr) return undefined;
  const solidFill = findChild(spPr, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  return normalizeRgbHex(srgbClr.attrs.val);
}

/**
 * Pull `<c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>` off
 * the supplied parent element. Returns the line stroke color as a
 * 6-character uppercase hex string.
 *
 * Same accept-or-drop grammar as {@link parseSpPrFill}; lands on
 * the line-color child rather than the fill child of the same
 * `<c:spPr>` block.
 */
export function parseSpPrBorderColor(parent: XmlElement): string | undefined {
  const spPr = findChild(parent, "spPr");
  if (!spPr) return undefined;
  const ln = findChild(spPr, "ln");
  if (!ln) return undefined;
  const solidFill = findChild(ln, "solidFill");
  if (!solidFill) return undefined;
  const srgbClr = findChild(solidFill, "srgbClr");
  if (!srgbClr) return undefined;
  return normalizeRgbHex(srgbClr.attrs.val);
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
  const spPr = findChild(parent, "spPr");
  if (!spPr) return undefined;
  const ln = findChild(spPr, "ln");
  if (!ln) return undefined;
  const wAttr = ln.attrs.w;
  if (typeof wAttr !== "string") return undefined;
  const emu = Number.parseFloat(wAttr);
  if (!Number.isFinite(emu) || emu <= 0) return undefined;
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const pt = Math.round((emu / EMU_PER_PT) * 4) / 4;
  if (pt < STROKE_WIDTH_MIN_PT) return STROKE_WIDTH_MIN_PT;
  if (pt > STROKE_WIDTH_MAX_PT) return STROKE_WIDTH_MAX_PT;
  return pt;
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
  const spPr = findChild(parent, "spPr");
  if (!spPr) return undefined;
  const ln = findChild(spPr, "ln");
  if (!ln) return undefined;
  const prstDash = findChild(ln, "prstDash");
  if (!prstDash) return undefined;
  const raw = prstDash.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim() as ChartBorderDash;
  if (!VALID_BORDER_DASHES.has(trimmed)) return undefined;
  if (trimmed === "solid") return undefined;
  return trimmed;
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
  if (value === undefined || !Number.isFinite(value)) return undefined;
  // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
  const snapped = Math.round(value * 4) / 4;
  if (snapped < STROKE_WIDTH_MIN_PT) return STROKE_WIDTH_MIN_PT;
  if (snapped > STROKE_WIDTH_MAX_PT) return STROKE_WIDTH_MAX_PT;
  return snapped;
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
  if (typeof value !== "string") return undefined;
  if (!VALID_BORDER_DASHES.has(value)) return undefined;
  if (value === "solid") return undefined;
  return value;
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
  if (override === undefined) return clampStrokeWidthPt(sourceValue);
  if (override === null) return undefined;
  return clampStrokeWidthPt(override);
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
  if (override === undefined) return normalizeBorderDash(sourceValue);
  if (override === null) return undefined;
  return normalizeBorderDash(override);
}
