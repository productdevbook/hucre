// ── Chart Layout ───────────────────────────────────────────────────
// Helpers for the OOXML `<c:layout><c:manualLayout>` block that every
// chart-frame slot (chart-title, axis-title, legend, plot-area) carries
// to declare a custom position / size.
//
// CT_ManualLayout (ECMA-376 Part 1, §21.2.2.176) places `<c:x>` /
// `<c:y>` / `<c:w>` / `<c:h>` children with `val` attributes in the
// 0..1 band (fractions of the chart frame). Anything outside that band
// collapses to `undefined` per the accept-or-drop grammar mirrored by
// the writer.
//
// JSDoc on the per-host parsers (parseTitleLayout, parseLegendLayout,
// parsePlotAreaLayout, parseAxisTitleLayout) stays attached to those
// callers in chart-reader.ts because each has host-specific scope
// commentary worth keeping.

import type { ChartManualLayout } from "./types";
import type { XmlElement } from "../../xml/parser";
import { findChild } from "./util";
import { xmlElement, xmlSelfClose } from "../../xml/writer";

/**
 * Parse a single `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` element off a
 * `<c:manualLayout>` block. Returns the `val` attribute as a finite
 * number in the `0..1` band; everything else (missing element, missing
 * attribute, non-numeric / non-finite / out-of-range token) collapses
 * to `undefined` so the matching axis on the parsed `ChartManualLayout`
 * is omitted.
 */
export function readLayoutCoordinate(el: XmlElement | undefined): number | undefined {
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed.length === 0) return undefined;
  const parsed = Number(trimmed);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 0 || parsed > 1) return undefined;
  return parsed;
}

/**
 * Walk the `<c:layout><c:manualLayout>` chain on the supplied parent
 * element (`<c:title>` / `<c:legend>` / `<c:plotArea>` / a `<c:title>`
 * nested under an axis) and surface its `<c:x>` / `<c:y>` / `<c:w>` /
 * `<c:h>` coordinates as a {@link ChartManualLayout} record.
 *
 * Returns `undefined` when neither `<c:layout>` nor `<c:manualLayout>`
 * is present, when none of the four coordinates surface a meaningful
 * value (each one runs through {@link readLayoutCoordinate}), or when
 * the chain is malformed at any link. Same accept-or-drop grammar
 * shared by every per-host wrapper that uses it.
 */
export function parseManualLayout(parent: XmlElement): ChartManualLayout | undefined {
  const layout = findChild(parent, "layout");
  if (!layout) return undefined;
  const manual = findChild(layout, "manualLayout");
  if (!manual) return undefined;
  const x = readLayoutCoordinate(findChild(manual, "x"));
  const y = readLayoutCoordinate(findChild(manual, "y"));
  const w = readLayoutCoordinate(findChild(manual, "w"));
  const h = readLayoutCoordinate(findChild(manual, "h"));
  if (x === undefined && y === undefined && w === undefined && h === undefined) {
    return undefined;
  }
  const out: ChartManualLayout = {};
  if (x !== undefined) out.x = x;
  if (y !== undefined) out.y = y;
  if (w !== undefined) out.w = w;
  if (h !== undefined) out.h = h;
  return out;
}

/**
 * Normalized `<c:manualLayout>` coordinate set after the writer runs
 * the caller's input through the `0..1` range filter. Each axis is
 * independently optional — a caller can pin only the position
 * (`x` / `y`) and let the element keep its automatic size, only the
 * size (`w` / `h`) and let it keep its automatic anchor, or any
 * combination.
 */
export interface ResolvedManualLayout {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

/**
 * Normalize a {@link ChartManualLayout} into the writer's emit-ready
 * shape. Drops every axis whose input is non-numeric / non-finite /
 * out of the `0..1` band; returns `undefined` when every axis dropped
 * so the caller can elide the entire `<c:layout>` block.
 *
 * The accept-and-clamp grammar matches the OOXML `CT_ManualLayout`
 * schema — `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` carry `xsd:double`
 * values in the `0..1` band per Excel's reference serialization. The
 * normalizer does not silently clamp out-of-range inputs to the
 * endpoints — it drops them outright, mirroring how `titleFontSize` /
 * `axisTitleFontSize` / `legendFontSize` collapse out-of-range numbers
 * rather than emit a token Excel would reject.
 */
export function normalizeManualLayout(
  raw: ChartManualLayout | undefined,
): ResolvedManualLayout | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const out: ResolvedManualLayout = {};
  const x = normalizeLayoutCoordinate(raw.x);
  if (x !== undefined) out.x = x;
  const y = normalizeLayoutCoordinate(raw.y);
  if (y !== undefined) out.y = y;
  const w = normalizeLayoutCoordinate(raw.w);
  if (w !== undefined) out.w = w;
  const h = normalizeLayoutCoordinate(raw.h);
  if (h !== undefined) out.h = h;
  if (out.x === undefined && out.y === undefined && out.w === undefined && out.h === undefined) {
    return undefined;
  }
  return out;
}

/**
 * Normalize a single `<c:x>` / `<c:y>` / `<c:w>` / `<c:h>` coordinate.
 * Accepts a finite number in the `0..1` band; everything else drops to
 * `undefined`.
 */
export function normalizeLayoutCoordinate(raw: unknown): number | undefined {
  if (typeof raw !== "number") return undefined;
  if (!Number.isFinite(raw)) return undefined;
  if (raw < 0 || raw > 1) return undefined;
  return raw;
}

/**
 * Normalize a {@link ChartManualLayout} for a cloned `SheetChart`.
 * Drops every axis whose input is non-numeric / non-finite / outside
 * the `0..1` band; returns `undefined` when every axis dropped so the
 * cloned shape elides the field entirely (mirrors the writer-side
 * normalization so a parsed value flows through `cloneChart` without
 * bookkeeping the units). Coordinates outside the `0..1` band collapse
 * rather than clamp — same accept-or-drop grammar as `titleFontSize` /
 * `axisTitleFontSize` / `legendFontSize`.
 *
 * Distinct from {@link normalizeManualLayout} which returns
 * {@link ResolvedManualLayout} for the writer; this variant preserves
 * the public {@link ChartManualLayout} surface so the cloned
 * `SheetChart` carries the same shape consumers passed in.
 */
export function normalizeChartManualLayout(
  value: ChartManualLayout | undefined,
): ChartManualLayout | undefined {
  if (!value || typeof value !== "object") return undefined;
  const out: ChartManualLayout = {};
  const x = normalizeLayoutCoordinate(value.x);
  if (x !== undefined) out.x = x;
  const y = normalizeLayoutCoordinate(value.y);
  if (y !== undefined) out.y = y;
  const w = normalizeLayoutCoordinate(value.w);
  if (w !== undefined) out.w = w;
  const h = normalizeLayoutCoordinate(value.h);
  if (h !== undefined) out.h = h;
  if (out.x === undefined && out.y === undefined && out.w === undefined && out.h === undefined) {
    return undefined;
  }
  return out;
}

/**
 * Build the `<c:layout><c:manualLayout>...</c:manualLayout></c:layout>`
 * block for a resolved layout. Returns `undefined` when the input is
 * `undefined` so the caller can elide the entire block.
 *
 * The writer always emits the `<c:xMode>` / `<c:yMode>` / `<c:wMode>` /
 * `<c:hMode>` children with `val="edge"` whenever the matching `<c:x>` /
 * `<c:y>` / `<c:w>` / `<c:h>` slot is present — `"edge"` is Excel's
 * reference shape when the user drags an element to a custom position
 * (the coordinates are absolute fractions of the chart frame, not
 * deltas from the auto-layout baseline). The `"factor"` form (delta
 * from auto-layout) is read on parse but normalized to `"edge"` on
 * emit so a re-parse after a clone-through stays canonical.
 *
 * The OOXML `CT_ManualLayout` sequence places the mode children before
 * the value children: `<c:layoutTarget>?` / `<c:xMode>?` / `<c:yMode>?`
 * / `<c:wMode>?` / `<c:hMode>?` / `<c:x>?` / `<c:y>?` / `<c:w>?` /
 * `<c:h>?` (ECMA-376 Part 1, §21.2.2.115). The writer emits in that
 * order so a re-parse sees the canonical shape.
 */
export function buildManualLayout(layout: ResolvedManualLayout | undefined): string | undefined {
  if (!layout) return undefined;
  const children: string[] = [];
  if (layout.x !== undefined) children.push(xmlSelfClose("c:xMode", { val: "edge" }));
  if (layout.y !== undefined) children.push(xmlSelfClose("c:yMode", { val: "edge" }));
  if (layout.w !== undefined) children.push(xmlSelfClose("c:wMode", { val: "edge" }));
  if (layout.h !== undefined) children.push(xmlSelfClose("c:hMode", { val: "edge" }));
  if (layout.x !== undefined) children.push(xmlSelfClose("c:x", { val: layout.x }));
  if (layout.y !== undefined) children.push(xmlSelfClose("c:y", { val: layout.y }));
  if (layout.w !== undefined) children.push(xmlSelfClose("c:w", { val: layout.w }));
  if (layout.h !== undefined) children.push(xmlSelfClose("c:h", { val: layout.h }));
  if (children.length === 0) return undefined;
  return xmlElement("c:layout", undefined, [xmlElement("c:manualLayout", undefined, children)]);
}
