// ── Chart Space ──────────────────────────────────────────────────
// Per-host module for the clone-side chart-space-level resolvers
// (`<c:chartSpace>` is the document root that wraps `<c:chart>`).
// These resolvers compose source values from `parseChart` with
// 3-arg `(source, override)` overrides from `cloneChart`.
//
// The writer-side single-arg `resolve*(chart: SheetChart)` resolvers
// for the same chart-space toggles (resolveDispBlanksAs,
// resolvePlotVisOnly, resolveShowDLblsOverMax, resolveRoundedCorners,
// resolveStyle, resolveLang, resolveDate1904, resolveProtection,
// resolveChartSpaceFillColor, resolveChartSpaceBorderColor) live in
// chart-writer.ts and stay there because their signatures and call
// patterns remain orchestration-shape rather than per-host.

import type { ChartColor, ChartDisplayBlanksAs, ChartProtection } from "../../_types"
import { normalizeChartColor } from "./shape"

/**
 * Resolve a `dispBlanksAs` override.
 *
 * `undefined` → inherit the source's parsed `dispBlanksAs`.
 * `null`      → drop the inherited value (the writer falls back to
 *               the OOXML `"gap"` default).
 * value       → replace.
 *
 * Unknown strings are ignored (treated as `undefined`); only the three
 * OOXML-defined tokens propagate through to the writer.
 */
export function resolveCloneDispBlanksAs(
  sourceValue: ChartDisplayBlanksAs | undefined,
  override: ChartDisplayBlanksAs | null | undefined,
): ChartDisplayBlanksAs | undefined {
  if (override === undefined) return sourceValue
  if (override === null) return undefined
  return override
}

/**
 * Resolve a `plotVisOnly` override.
 *
 * `undefined` → inherit the source's parsed `plotVisOnly`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `true` default — hidden cells drop out of the chart).
 * `boolean`   → replace.
 *
 * The grammar mirrors `dispBlanksAs` / `varyColors` so the chart-level
 * toggles compose the same way at the call site.
 */
export function resolveClonePlotVisOnly(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue
  if (override === null) return undefined
  return override
}

/**
 * Resolve a `showDLblsOverMax` override.
 *
 * `undefined` → inherit the source's parsed `showDLblsOverMax`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `true` default — labels render for every point
 *               regardless of the pinned axis ceiling).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `dispBlanksAs` so the chart-level
 * toggles compose the same way at the call site.
 */
export function resolveCloneShowDLblsOverMax(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue
  if (override === null) return undefined
  return override
}

/**
 * Resolve a `roundedCorners` override.
 *
 * `undefined` → inherit the source's parsed `roundedCorners`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               OOXML `false` default — square chart frame).
 * `boolean`   → replace.
 *
 * The grammar mirrors `plotVisOnly` / `varyColors` so the chart-frame
 * toggles compose the same way at the call site.
 */
export function resolveCloneRoundedCorners(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue
  if (override === null) return undefined
  return override
}

/**
 * Resolve a `style` (built-in chart preset) override.
 *
 * `undefined` → inherit the source's parsed `style`.
 * `null`      → drop the inherited value (the writer skips `<c:style>`
 *               so Excel falls back to its application default look).
 * `number`    → replace. Out-of-range / non-integer values are not
 *               filtered here — the writer's `resolveCloneStyle` performs
 *               the same shape check on emit, so a stray value never
 *               reaches the rendered XML regardless of the path it
 *               took through clone.
 *
 * The grammar mirrors `roundedCorners` / `plotVisOnly` so the chart-
 * frame toggles compose the same way at the call site.
 */
export function resolveCloneStyle(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return sourceValue
  if (override === null) return undefined
  return override
}

/**
 * Resolve a `lang` (chart-space editing-locale hint) override.
 *
 * `undefined` → inherit the source's parsed `lang`.
 * `null`      → drop the inherited value (the writer skips `<c:lang>`
 *               so Excel falls back to the host workbook's editing
 *               language).
 * `string`    → replace. Malformed culture names are not filtered
 *               here — the writer's `resolveCloneLang` performs the same
 *               BCP-47 shape check on emit, so a stray value never
 *               reaches the rendered XML regardless of the path it
 *               took through clone.
 *
 * The grammar mirrors `style` / `roundedCorners` / `plotVisOnly` so
 * the chart-space toggles compose the same way at the call site.
 */
export function resolveCloneLang(
  sourceValue: string | undefined,
  override: string | null | undefined,
): string | undefined {
  if (override === undefined) return sourceValue
  if (override === null) return undefined
  return override
}

/**
 * Resolve a `date1904` (chart-space date-system) override.
 *
 * `undefined` → inherit the source's parsed `date1904`.
 * `null`      → drop the inherited value (the writer skips
 *               `<c:date1904>` so Excel falls back to the host
 *               workbook's date system).
 * `boolean`   → replace. `false` collapses to absence on the writer
 *               side because `<c:date1904 val="0"/>` is the OOXML
 *               default and the writer follows the minimal-shape
 *               contract every other chart-space toggle uses.
 *
 * The grammar mirrors `roundedCorners` / `plotVisOnly` so the
 * chart-space toggles compose the same way at the call site. `false`
 * here means "explicitly pin the 1900 base" — but because absence
 * and `val="0"` round-trip identically the resolved value still
 * collapses to `undefined` (the writer would emit nothing either
 * way).
 */
export function resolveCloneDate1904(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  const merged = override === undefined ? sourceValue : override === null ? undefined : override
  if (merged === true) return true
  // `false` and `undefined` both collapse to `undefined` — absence
  // and the OOXML default `<c:date1904 val="0"/>` round-trip the
  // same way through parseChart -> cloneChart -> writeChart, so the
  // resolved chart drops the field rather than carry a value the
  // writer would skip on emit anyway.
  return undefined
}

/**
 * Resolve a `protection` (chart-space protection) override.
 *
 * `undefined` → inherit the source's parsed {@link Chart.protection}.
 * `null`      → drop the inherited block so the writer skips
 *               `<c:protection>` entirely (no chart-level lock).
 * `false`     → equivalent to `null` (suppression); kept distinct in
 *               the API surface so callers can write `protection:
 *               false` for symmetry with the writer's `boolean |
 *               object` shape.
 * `true`      → enable with the OOXML reference defaults (every flag
 *               `false` — the bare `<c:protection>` shell).
 * `object`    → replace the inherited block wholesale (no per-field
 *               merge with the source — pass every flag the cloned
 *               protection should pin). Each unspecified field falls
 *               back to `false` at the writer side because every
 *               `<c:protection>` boolean child is independently
 *               optional and Excel treats a missing child as
 *               `false`.
 *
 * The grammar mirrors {@link resolveCloneDataTable} so the chart-level
 * block toggles compose the same way at the call site. Unlike
 * `dataTable`, `<c:protection>` lives on `<c:chartSpace>` (not inside
 * `<c:plotArea>`) so the resolver applies to every chart family —
 * pie / doughnut included.
 */
export function resolveCloneProtection(
  sourceValue: ChartProtection | undefined,
  override: ChartProtection | boolean | null | undefined,
): ChartProtection | boolean | undefined {
  if (override === undefined) {
    // Inherit — pass the source through verbatim. The writer accepts
    // both the boolean and object shapes, so a parsed
    // {@link ChartProtection} round-trips directly.
    return sourceValue
  }
  if (override === null) {
    // Drop the inherited block. The writer treats `undefined` as
    // suppression and skips `<c:protection>` entirely.
    return undefined
  }
  if (override === false) {
    // Symmetric with `null` — kept distinct in the API surface for
    // ergonomic alignment with the writer's `boolean | object` shape,
    // but emits the same on-the-wire result (no `<c:protection>`).
    return undefined
  }
  // `true` or a {@link ChartProtection} object — replace the inherited
  // block wholesale. The writer accepts both forms and falls back to
  // the OOXML reference default `false` for any field the object
  // leaves unset.
  return override
}

/**
 * Normalize a `plotAreaFillColor` value for the cloned `SheetChart`.
 * Mirrors the writer's `normalizeClonePlotAreaFillColor` — the cloned shape
 * is guaranteed to round-trip through the writer without surprise: a
 * leading `#` and any case are accepted, then the value collapses to
 * the OOXML canonical uppercase form. Malformed inputs (wrong length,
 * non-hex characters, alpha-channel forms, non-string escapes from an
 * untyped caller) collapse to `undefined` so the cloned chart drops
 * the field rather than carry a value the writer would silently elide
 * back to absence.
 */
export function normalizeClonePlotAreaFillColor(
  value: ChartColor | undefined,
): ChartColor | undefined {
  return normalizeChartColor(value)
}

/**
 * Normalize a `plotAreaBorderColor` value for the cloned `SheetChart`.
 * Mirrors the writer's `normalizeClonePlotAreaBorderColor` — the cloned
 * shape is guaranteed to round-trip through the writer without
 * surprise: a leading `#` and any case are accepted, then the value
 * collapses to the OOXML canonical uppercase form. Malformed inputs
 * (wrong length, non-hex characters, alpha-channel forms, non-string
 * escapes from an untyped caller) collapse to `undefined` so the
 * cloned chart drops the field rather than carry a value the writer
 * would silently elide back to absence. Mirrors
 * {@link normalizeClonePlotAreaFillColor} — same hex grammar.
 */
export function normalizeClonePlotAreaBorderColor(
  value: ChartColor | undefined,
): ChartColor | undefined {
  return normalizeChartColor(value)
}

/**
 * Normalize a `chartSpaceFillColor` value for the cloned `SheetChart`.
 * Mirrors the writer's `normalizeCloneChartSpaceFillColor` (which itself
 * delegates to the chart-title / plot-area / legend hex normalizer) —
 * the cloned shape is guaranteed to round-trip through the writer
 * without surprise: a leading `#` and any case are accepted, then the
 * value collapses to the OOXML canonical uppercase form. Malformed
 * inputs (wrong length, non-hex characters, alpha-channel forms,
 * non-string escapes from an untyped caller) collapse to `undefined`
 * so the cloned chart drops the field rather than carry a value the
 * writer would silently elide back to absence.
 */
export function normalizeCloneChartSpaceFillColor(
  value: ChartColor | undefined,
): ChartColor | undefined {
  return normalizeChartColor(value)
}

/**
 * Resolve a `chartSpaceFillColor` override.
 *
 * `undefined` → inherit the source's parsed `chartSpaceFillColor`
 *               (after running it through
 *               {@link normalizeCloneChartSpaceFillColor} so a malformed
 *               source value drops cleanly).
 * `null`      → drop the inherited fill (the writer emits no
 *               `<c:spPr>` block on `<c:chartSpace>`, the chart
 *               inherits the auto-fill Excel picks from the workbook
 *               theme).
 * `string`    → replace with the normalized 6-character uppercase hex
 *               form. Malformed overrides collapse to `undefined` via
 *               the normalizer so the cloned `SheetChart` always
 *               carries a value the writer will accept.
 *
 * The grammar mirrors `plotAreaFillColor` / `legendFillColor` /
 * `titleColor` so the chart `<a:srgbClr>` fill / color knobs compose
 * the same way at the call site. Unlike the title / legend variants,
 * the chart-space fill is never gated on a visibility flag — every
 * chart has a `<c:chartSpace>` document root to host the fill.
 */
export function resolveCloneChartSpaceFillColor(
  sourceValue: ChartColor | undefined,
  override: ChartColor | null | undefined,
): ChartColor | undefined {
  if (override === undefined) return normalizeCloneChartSpaceFillColor(sourceValue)
  if (override === null) return undefined
  return normalizeCloneChartSpaceFillColor(override)
}

/**
 * Normalize a `chartSpaceBorderColor` value for the cloned `SheetChart`.
 * Mirrors the writer's `normalizeCloneChartSpaceBorderColor` (which itself
 * delegates to the chart-title / plot-area / legend hex normalizer) —
 * the cloned shape is guaranteed to round-trip through the writer
 * without surprise: a leading `#` and any case are accepted, then the
 * value collapses to the OOXML canonical uppercase form. Malformed
 * inputs (wrong length, non-hex characters, alpha-channel forms,
 * non-string escapes from an untyped caller) collapse to `undefined`
 * so the cloned chart drops the field rather than carry a value the
 * writer would silently elide back to absence.
 */
export function normalizeCloneChartSpaceBorderColor(
  value: ChartColor | undefined,
): ChartColor | undefined {
  return normalizeChartColor(value)
}

/**
 * Resolve a `chartSpaceBorderColor` override.
 *
 * `undefined` → inherit the source's parsed `chartSpaceBorderColor`
 *               (after running it through
 *               {@link normalizeCloneChartSpaceBorderColor} so a malformed
 *               source value drops cleanly).
 * `null`      → drop the inherited stroke (the writer emits no
 *               `<a:ln>` block on `<c:chartSpace>`'s `<c:spPr>`, the
 *               chart inherits the auto-stroke Excel picks from the
 *               workbook theme).
 * `string`    → replace with the normalized 6-character uppercase hex
 *               form. Malformed overrides collapse to `undefined` via
 *               the normalizer so the cloned `SheetChart` always
 *               carries a value the writer will accept.
 *
 * The grammar mirrors `plotAreaBorderColor` / `legendBorderColor` /
 * `titleBorderColor` so the chart `<a:ln>` stroke knobs compose the
 * same way at the call site. Unlike the title / legend variants, the
 * chart-space border is never gated on a visibility flag — every chart
 * has a `<c:chartSpace>` document root to host the stroke.
 */
export function resolveCloneChartSpaceBorderColor(
  sourceValue: ChartColor | undefined,
  override: ChartColor | null | undefined,
): ChartColor | undefined {
  if (override === undefined) return normalizeCloneChartSpaceBorderColor(sourceValue)
  if (override === null) return undefined
  return normalizeCloneChartSpaceBorderColor(override)
}
