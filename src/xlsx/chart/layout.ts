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

/** See `chart/shape.ts` for the equivalent helper. */
function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

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
