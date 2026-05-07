// ── Chart Walls ────────────────────────────────────────────────────
// Per-host module for the 3D-only chart-frame children — `<c:view3D>`,
// `<c:floor>`, `<c:sideWall>`, `<c:backWall>` (CT_View3D + three
// `CT_Surface` blocks per CT_Chart, ECMA-376 Part 1, §21.2.2.4 +
// §21.2.2.69 / §21.2.2.187 / §21.2.2.31). The OOXML schema accepts the
// blocks on every CT_Chart even though they are only meaningful on 3D
// chart families (`bar3D`, `line3D`, `pie3D`, `area3D`, `surface3D`);
// Excel silently ignores them on 2D families. The reader / writer /
// clone helpers in this module preserve that scope so a 3D template
// can round-trip through hucre cleanly.

import type { ChartView3D } from "./types";
import type { XmlElement } from "../../xml/parser";
import { findChild } from "./util";
import { xmlElement, xmlSelfClose } from "../../xml/writer";

// ── Reader ────────────────────────────────────────────────────────

/**
 * Pull `<c:view3D>` (CT_View3D) off `<c:chart>`. Surfaces a
 * {@link ChartView3D} object whenever the source chart declares the
 * element. Each of the six children (`<c:rotX>`, `<c:hPercent>`,
 * `<c:rotY>`, `<c:depthPercent>`, `<c:rAngAx>`, `<c:perspective>`)
 * is independently optional on CT_View3D, so the reader only surfaces
 * the fields the file actually pinned. A child that is missing or
 * carries an out-of-range / unparseable `val` attribute drops to
 * `undefined` for that field rather than fabricate a value the file
 * did not declare.
 *
 * The element itself is the gating signal — a `<c:view3D>` block with
 * no resolvable children surfaces as an empty `{}`, mirroring how
 * `dataTable` / `protection` handle a malformed inner block. This
 * keeps a chart that authors the bare element (Excel's "default 3D
 * view" preset) from silently disappearing through the parse loop.
 *
 * Note: `<c:view3D>` lives on `<c:chart>` (between `<c:autoTitleDeleted>`
 * / `<c:pivotFmts>` and `<c:floor>` / `<c:plotArea>` per CT_Chart
 * §21.2.2.4), not on `<c:chartSpace>` — the toggle governs the 3D
 * projection of the rendered chart, not the outer chart frame.
 */
export function parseView3D(chartEl: XmlElement): ChartView3D | undefined {
  const el = findChild(chartEl, "view3D");
  if (!el) return undefined;
  const out: ChartView3D = {};
  // `<c:rotX>` (CT_RotX, ST_RotX) is a signed byte in the range
  // -90..90. Out-of-range values drop rather than emit a token Excel
  // would clamp at parse time.
  const rotX = parseView3DInt(el, "rotX", -90, 90);
  if (rotX !== undefined) out.rotX = rotX;
  // `<c:hPercent>` (CT_HPercent, ST_HPercent) is a percent value in
  // the range 5..500. Same drop-on-out-of-range rule.
  const hPercent = parseView3DInt(el, "hPercent", 5, 500);
  if (hPercent !== undefined) out.hPercent = hPercent;
  // `<c:rotY>` (CT_RotY, ST_RotY) is an unsigned short in the range
  // 0..360.
  const rotY = parseView3DInt(el, "rotY", 0, 360);
  if (rotY !== undefined) out.rotY = rotY;
  // `<c:depthPercent>` (CT_DepthPercent, ST_DepthPercent) is a percent
  // value in the range 20..2000.
  const depthPercent = parseView3DInt(el, "depthPercent", 20, 2000);
  if (depthPercent !== undefined) out.depthPercent = depthPercent;
  // `<c:rAngAx>` (CT_Boolean) — accepts the OOXML truthy / falsy
  // spellings; unknown values and missing `val` attributes drop to
  // `undefined`. Mirrors the parsing semantics of the chartSpace-level
  // `<c:protection>` boolean children.
  const rAngAx = parseView3DBoolean(el, "rAngAx");
  if (rAngAx !== undefined) out.rAngAx = rAngAx;
  // `<c:perspective>` (CT_Perspective, ST_Perspective) is a percent
  // value in the range 0..240.
  const perspective = parseView3DInt(el, "perspective", 0, 240);
  if (perspective !== undefined) out.perspective = perspective;
  return out;
}

/**
 * Pull a single integer child off `<c:view3D>`. Surfaces the value
 * only when `val` parses as an integer inside the matching OOXML
 * simple-type range; absence and out-of-range / non-integer values
 * collapse to `undefined`.
 *
 * Accepts an optional leading `-` so signed types (`<c:rotX>`) round-
 * trip cleanly. The strict integer regex rejects fractional values
 * (`"15.5"`) and non-numeric tokens (`"15px"`) — `parseInt` would
 * coerce both into a number Excel never emits.
 */
function parseView3DInt(
  view3D: XmlElement,
  local: string,
  min: number,
  max: number,
): number | undefined {
  const el = findChild(view3D, local);
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  if (!/^-?\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  if (n < min || n > max) return undefined;
  return n;
}

/**
 * Pull a single boolean child off `<c:view3D>`. Accepts the OOXML
 * truthy / falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`);
 * unknown tokens, missing `val` attributes, and missing elements all
 * collapse to `undefined` rather than fabricate a flag the file did
 * not pin. Mirrors {@link parseProtectionFlag} — the same OOXML
 * `<xsd:boolean>` lexical-space rule.
 */
function parseView3DBoolean(view3D: XmlElement, local: string): boolean | undefined {
  const el = findChild(view3D, local);
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  switch (raw) {
    case "1":
    case "true":
      return true;
    case "0":
    case "false":
      return false;
    default:
      return undefined;
  }
}

/**
 * Pull `<c:chart><c:floor><c:thickness val=".."/></c:floor>` off
 * `<c:chart>`. The `<c:floor>` element (CT_Surface, ECMA-376 Part 1,
 * §21.2.2.69) sits on `<c:chart>` between `<c:view3D>` and
 * `<c:sideWall>` / `<c:backWall>` / `<c:plotArea>` per CT_Chart and
 * carries an optional `<c:thickness>` child whose `val` attribute is
 * an `xsd:unsignedInt` — Excel's "Format Floor -> Floor -> Thickness"
 * pin on 3D chart families.
 *
 * Returns the integer pinned by the source chart. The OOXML default
 * `0` (and absence of the `<c:thickness>` child or the parent
 * `<c:floor>` element) collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart} — only an
 * explicit positive thickness surfaces here. Out-of-range or
 * unparseable values also drop to `undefined` rather than fabricate a
 * value the file did not declare.
 *
 * The `<c:thickness>` element only carries the `val` attribute on
 * `CT_Thickness` — other floor styling (`<c:spPr>`, `<c:pictureOptions>`,
 * `<c:extLst>`) is not modelled at this layer, so a stray styling
 * block on the floor passes through the parse loop without surfacing.
 */
export function parseFloorThickness(chartEl: XmlElement): number | undefined {
  const floor = findChild(chartEl, "floor");
  if (!floor) return undefined;
  const thickness = findChild(floor, "thickness");
  if (!thickness) return undefined;
  const raw = thickness.attrs.val;
  if (typeof raw !== "string") return undefined;
  // ST_Thickness is `xsd:unsignedInt` — strict integer regex rejects
  // fractional / negative / non-numeric tokens.
  if (!/^\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  // Collapse the OOXML default `0` to undefined so absence and the
  // default round-trip identically through cloneChart — only an
  // explicit positive thickness surfaces here. Mirrors how the writer
  // skips emission entirely for `0` / undefined.
  if (n === 0) return undefined;
  // Cap at Excel's UI band ceiling (`100`) — the OOXML schema accepts
  // the full `xsd:unsignedInt` range but Excel's "Format Floor"
  // dialogue rejects values above 100 with a repair warning. Anything
  // larger drops here so a corrupt template does not silently rewrite
  // as an absurd thickness; absence keeps the round-trip stable.
  if (n > 100) return undefined;
  return n;
}

/**
 * Pull `<c:chart><c:sideWall><c:thickness val=".."/></c:sideWall>` off
 * `<c:chart>`. The `<c:sideWall>` element (CT_Surface, ECMA-376 Part 1,
 * §21.2.2.187) sits on `<c:chart>` between `<c:floor>` and
 * `<c:backWall>` / `<c:plotArea>` per CT_Chart and carries an optional
 * `<c:thickness>` child whose `val` attribute is an `xsd:unsignedInt` —
 * Excel's "Format Side Wall -> Side Wall -> Thickness" pin on 3D chart
 * families.
 *
 * Returns the integer pinned by the source chart. The OOXML default
 * `0` (and absence of the `<c:thickness>` child or the parent
 * `<c:sideWall>` element) collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart} — only an
 * explicit positive thickness surfaces here. Out-of-range or
 * unparseable values also drop to `undefined` rather than fabricate a
 * value the file did not declare.
 *
 * The `<c:thickness>` element only carries the `val` attribute on
 * `CT_Thickness` — other side-wall styling (`<c:spPr>`,
 * `<c:pictureOptions>`, `<c:extLst>`) is not modelled at this layer,
 * so a stray styling block on the wall passes through the parse loop
 * without surfacing.
 */
export function parseSideWallThickness(chartEl: XmlElement): number | undefined {
  const sideWall = findChild(chartEl, "sideWall");
  if (!sideWall) return undefined;
  const thickness = findChild(sideWall, "thickness");
  if (!thickness) return undefined;
  const raw = thickness.attrs.val;
  if (typeof raw !== "string") return undefined;
  // ST_Thickness is `xsd:unsignedInt` — strict integer regex rejects
  // fractional / negative / non-numeric tokens.
  if (!/^\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  // Collapse the OOXML default `0` to undefined so absence and the
  // default round-trip identically through cloneChart — only an
  // explicit positive thickness surfaces here. Mirrors how the writer
  // skips emission entirely for `0` / undefined.
  if (n === 0) return undefined;
  // Cap at Excel's UI band ceiling (`100`) — the OOXML schema accepts
  // the full `xsd:unsignedInt` range but Excel's "Format Side Wall"
  // dialogue rejects values above 100 with a repair warning. Anything
  // larger drops here so a corrupt template does not silently rewrite
  // as an absurd thickness; absence keeps the round-trip stable.
  if (n > 100) return undefined;
  return n;
}

/**
 * Pull `<c:chart><c:backWall><c:thickness val=".."/></c:backWall>` off
 * `<c:chart>`. The `<c:backWall>` element (CT_Surface, ECMA-376 Part 1,
 * §21.2.2.31) sits on `<c:chart>` between `<c:sideWall>` and
 * `<c:plotArea>` per CT_Chart and carries an optional `<c:thickness>`
 * child whose `val` attribute is an `xsd:unsignedInt` — Excel's
 * "Format Back Wall -> Back Wall -> Thickness" pin on 3D chart families.
 *
 * Returns the integer pinned by the source chart. The OOXML default
 * `0` (and absence of the `<c:thickness>` child or the parent
 * `<c:backWall>` element) collapses to `undefined` so absence and the
 * default round-trip identically through {@link cloneChart} — only an
 * explicit positive thickness surfaces here. Out-of-range or
 * unparseable values also drop to `undefined` rather than fabricate a
 * value the file did not declare.
 *
 * The `<c:thickness>` element only carries the `val` attribute on
 * `CT_Thickness` — other back-wall styling (`<c:spPr>`,
 * `<c:pictureOptions>`, `<c:extLst>`) is not modelled at this layer,
 * so a stray styling block on the back wall passes through the parse
 * loop without surfacing.
 */
export function parseBackWallThickness(chartEl: XmlElement): number | undefined {
  const backWall = findChild(chartEl, "backWall");
  if (!backWall) return undefined;
  const thickness = findChild(backWall, "thickness");
  if (!thickness) return undefined;
  const raw = thickness.attrs.val;
  if (typeof raw !== "string") return undefined;
  // ST_Thickness is `xsd:unsignedInt` — strict integer regex rejects
  // fractional / negative / non-numeric tokens.
  if (!/^\d+$/.test(raw)) return undefined;
  const n = Number(raw);
  if (!Number.isInteger(n)) return undefined;
  // Collapse the OOXML default `0` to undefined so absence and the
  // default round-trip identically through cloneChart — only an
  // explicit positive thickness surfaces here. Mirrors how the writer
  // skips emission entirely for `0` / undefined.
  if (n === 0) return undefined;
  // Cap at Excel's UI band ceiling (`100`) — the OOXML schema accepts
  // the full `xsd:unsignedInt` range but Excel's "Format Back Wall"
  // dialogue rejects values above 100 with a repair warning. Anything
  // larger drops here so a corrupt template does not silently rewrite
  // as an absurd thickness; absence keeps the round-trip stable.
  if (n > 100) return undefined;
  return n;
}

// ── Writer ────────────────────────────────────────────────────────

/**
 * Serialize a {@link ChartView3D} into `<c:view3D>` with one self-
 * closing child per pinned field, in the order CT_View3D mandates:
 * `<c:rotX>`, `<c:hPercent>`, `<c:rotY>`, `<c:depthPercent>`,
 * `<c:rAngAx>`, `<c:perspective>`. Returns `undefined` when the
 * caller did not opt in (`view3D` is `undefined`) so the writer can
 * skip the element entirely; returns the bare `<c:view3D/>` shell
 * when an empty object is passed (round-trips a template that
 * authored the element with no children pinned).
 *
 * Each numeric field is clamped against the matching OOXML simple-
 * type range — out-of-range and non-finite inputs drop silently
 * rather than emit a token Excel's strict validator would reject.
 * The boolean field surfaces only as a literal `0` / `1` `val`
 * attribute; non-boolean inputs collapse to `false` (the OOXML
 * default), mirroring how every other chart-level boolean writer
 * treats its input.
 */
export function buildView3D(view3D: ChartView3D | undefined): string | undefined {
  if (view3D === undefined) return undefined;
  const children: string[] = [];
  // CT_View3D children sequence per ECMA-376 §21.2.2.228:
  // rotX?, hPercent?, rotY?, depthPercent?, rAngAx?, perspective?,
  // extLst?
  const rotX = clampView3DInt(view3D.rotX, -90, 90);
  if (rotX !== undefined) children.push(xmlSelfClose("c:rotX", { val: rotX }));
  const hPercent = clampView3DInt(view3D.hPercent, 5, 500);
  if (hPercent !== undefined) {
    // `<c:hPercent>` accepts the bare integer per ST_HPercent — Excel
    // emits a plain percent value with no `%` suffix.
    children.push(xmlSelfClose("c:hPercent", { val: hPercent }));
  }
  const rotY = clampView3DInt(view3D.rotY, 0, 360);
  if (rotY !== undefined) children.push(xmlSelfClose("c:rotY", { val: rotY }));
  const depthPercent = clampView3DInt(view3D.depthPercent, 20, 2000);
  if (depthPercent !== undefined) {
    children.push(xmlSelfClose("c:depthPercent", { val: depthPercent }));
  }
  if (view3D.rAngAx === true) {
    children.push(xmlSelfClose("c:rAngAx", { val: 1 }));
  } else if (view3D.rAngAx === false) {
    // Explicit `false` round-trips as `<c:rAngAx val="0"/>` so the
    // caller can pin the OOXML default literally — useful for parity
    // with templates that author the explicit value.
    children.push(xmlSelfClose("c:rAngAx", { val: 0 }));
  }
  const perspective = clampView3DInt(view3D.perspective, 0, 240);
  if (perspective !== undefined) {
    children.push(xmlSelfClose("c:perspective", { val: perspective }));
  }
  // Empty object (`{}`) collapses to a bare `<c:view3D/>` shell —
  // `xmlElement` with an empty child array emits the self-closing form.
  return xmlElement("c:view3D", undefined, children);
}

/**
 * Clamp a `<c:view3D>` numeric field against the matching OOXML
 * simple-type range. Returns `undefined` when the input is non-finite,
 * non-integer, or out-of-range — the writer drops the matching child
 * rather than emit a token Excel's strict validator would reject.
 *
 * The strict integer check rejects fractional inputs (`15.5`) so the
 * round-trip stays lossless — `parseView3DInt` on the reader side
 * also rejects fractional `val` attributes, and a fractional input
 * here would silently mismatch on the next parse.
 */
function clampView3DInt(value: number | undefined, min: number, max: number): number | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  if (value < min || value > max) return undefined;
  return value;
}

/**
 * Serialize {@link SheetChart.floorThickness} into
 * `<c:floor><c:thickness val="N"/></c:floor>`. Returns `undefined`
 * when the caller did not opt in (`floorThickness` is `undefined`,
 * `0`, non-finite, non-integer, negative, or out of the Excel UI
 * band `1..100`) so the writer can skip the `<c:floor>` element
 * entirely — Excel renders no floor extrusion on a fresh chart and
 * absence matches the reference serialization byte-for-byte.
 *
 * The OOXML schema (`ST_Thickness`, `xsd:unsignedInt`) accepts any
 * non-negative integer, but Excel's "Format Floor -> Floor ->
 * Thickness" pane only exposes `0..100` — values above that band
 * render but trigger Excel's repair dialog. Drop out-of-range and
 * non-integer inputs rather than emit a token Excel rejects, mirroring
 * how every other chart-level numeric writer ({@link clampView3DInt}
 * / {@link clampHoleSize} / {@link clampFirstSliceAng}) treats its
 * input.
 *
 * The element only carries the `<c:thickness>` child — other
 * `CT_Surface` children (`<c:spPr>`, `<c:pictureOptions>`,
 * `<c:extLst>`) are not modelled at this layer, so the emitted
 * `<c:floor>` block is the minimal shape Excel itself emits when the
 * user pins a thickness with no other floor styling.
 */
export function buildFloorThickness(value: number | undefined): string | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  if (value <= 0) return undefined;
  if (value > 100) return undefined;
  return xmlElement("c:floor", undefined, [xmlSelfClose("c:thickness", { val: value })]);
}

/**
 * Serialize {@link SheetChart.sideWallThickness} into
 * `<c:sideWall><c:thickness val="N"/></c:sideWall>`. Returns
 * `undefined` when the caller did not opt in
 * (`sideWallThickness` is `undefined`, `0`, non-finite, non-integer,
 * negative, or out of the Excel UI band `1..100`) so the writer can
 * skip the `<c:sideWall>` element entirely — Excel renders no side-
 * wall extrusion on a fresh chart and absence matches the reference
 * serialization byte-for-byte.
 *
 * The OOXML schema (`ST_Thickness`, `xsd:unsignedInt`) accepts any
 * non-negative integer, but Excel's "Format Side Wall -> Side Wall ->
 * Thickness" pane only exposes `0..100` — values above that band
 * render but trigger Excel's repair dialog. Drop out-of-range and
 * non-integer inputs rather than emit a token Excel rejects, mirroring
 * how every other chart-level numeric writer ({@link clampView3DInt}
 * / {@link clampHoleSize} / {@link clampFirstSliceAng}) treats its
 * input.
 *
 * The element only carries the `<c:thickness>` child — other
 * `CT_Surface` children (`<c:spPr>`, `<c:pictureOptions>`,
 * `<c:extLst>`) are not modelled at this layer, so the emitted
 * `<c:sideWall>` block is the minimal shape Excel itself emits when
 * the user pins a thickness with no other side-wall styling.
 */
export function buildSideWallThickness(value: number | undefined): string | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  if (value <= 0) return undefined;
  if (value > 100) return undefined;
  return xmlElement("c:sideWall", undefined, [xmlSelfClose("c:thickness", { val: value })]);
}

/**
 * Serialize {@link SheetChart.backWallThickness} into
 * `<c:backWall><c:thickness val="N"/></c:backWall>`. Returns
 * `undefined` when the caller did not opt in (`backWallThickness` is
 * `undefined`, `0`, non-finite, non-integer, negative, or out of the
 * Excel UI band `1..100`) so the writer can skip the `<c:backWall>`
 * element entirely — Excel renders no back-wall extrusion on a fresh
 * chart and absence matches the reference serialization byte-for-byte.
 *
 * The OOXML schema (`ST_Thickness`, `xsd:unsignedInt`) accepts any
 * non-negative integer, but Excel's "Format Back Wall -> Back Wall ->
 * Thickness" pane only exposes `0..100` — values above that band
 * render but trigger Excel's repair dialog. Drop out-of-range and
 * non-integer inputs rather than emit a token Excel rejects, mirroring
 * how every other chart-level numeric writer ({@link buildFloorThickness}
 * / {@link clampView3DInt} / {@link clampHoleSize}) treats its input.
 *
 * The element only carries the `<c:thickness>` child — other
 * `CT_Surface` children (`<c:spPr>`, `<c:pictureOptions>`,
 * `<c:extLst>`) are not modelled at this layer, so the emitted
 * `<c:backWall>` block is the minimal shape Excel itself emits when
 * the user pins a thickness with no other back-wall styling.
 */
export function buildBackWallThickness(value: number | undefined): string | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  if (value <= 0) return undefined;
  if (value > 100) return undefined;
  return xmlElement("c:backWall", undefined, [xmlSelfClose("c:thickness", { val: value })]);
}

// ── Clone ─────────────────────────────────────────────────────────

/**
 * Resolve a `view3D` override.
 *
 * `undefined` → inherit the source's parsed {@link Chart.view3D}. The
 *               source's parsed object is defensively shallow-copied
 *               so a downstream mutation to the cloned SheetChart
 *               never leaks back into the parsed Chart.
 * `null`      → drop the inherited block so the writer skips
 *               `<c:view3D>` entirely (no chart-level 3D pin).
 * `object`    → replace the inherited block wholesale (no per-field
 *               merge with the source — pass every field the cloned
 *               view3D should pin). An empty object emits a bare
 *               `<c:view3D/>` shell at the writer side.
 *
 * The grammar mirrors {@link resolveProtection} / {@link resolveDataTable}
 * so the chart-level block toggles compose the same way at the call
 * site. Unlike `dataTable`, `<c:view3D>` lives on `<c:chart>` (not
 * inside `<c:plotArea>`) so the resolver applies to every chart family
 * — pie / doughnut included.
 */
export function resolveView3D(
  sourceValue: ChartView3D | undefined,
  override: ChartView3D | null | undefined,
): ChartView3D | undefined {
  if (override === undefined) {
    // Inherit — defensively shallow-copy the source so a downstream
    // mutation to the cloned SheetChart never leaks back into the
    // parsed Chart. The CT_View3D children are all scalars (numbers
    // and a boolean), so a single-level spread is enough.
    if (sourceValue === undefined) return undefined;
    return { ...sourceValue };
  }
  if (override === null) {
    // Drop the inherited block. The writer treats `undefined` as
    // suppression and skips `<c:view3D>` entirely.
    return undefined;
  }
  // Replace the inherited block wholesale. The writer accepts the
  // empty-object shape and emits a bare `<c:view3D/>` shell, mirroring
  // how `resolveProtection` handles the `true` / `{}` forms.
  return { ...override };
}

/**
 * Resolve a `floorThickness` override.
 *
 * `undefined` → inherit the source's parsed `floorThickness`.
 * `null`      → drop the inherited value (the writer skips `<c:floor>`
 *               entirely — Excel falls back to no extrusion).
 * `number`    → replace. Out-of-range, `0`, or non-finite values still
 *               surface in the cloned `SheetChart` for symmetry with
 *               the other override helpers; the writer's
 *               `buildFloorThickness` then drops them at emit time so
 *               a fresh chart matches Excel's reference serialization.
 *
 * The grammar mirrors `upDownBarsGapWidth` / `gapWidth` / `holeSize` /
 * `firstSliceAng` so the numeric chart-level knobs compose the same
 * way at the call site.
 */
export function resolveFloorThickness(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `sideWallThickness` override.
 *
 * `undefined` → inherit the source's parsed `sideWallThickness`.
 * `null`      → drop the inherited value (the writer skips
 *               `<c:sideWall>` entirely — Excel falls back to no
 *               extrusion).
 * `number`    → replace. Out-of-range, `0`, or non-finite values still
 *               surface in the cloned `SheetChart` for symmetry with
 *               the other override helpers; the writer's
 *               `buildSideWallThickness` then drops them at emit time
 *               so a fresh chart matches Excel's reference
 *               serialization.
 *
 * The grammar mirrors `upDownBarsGapWidth` / `gapWidth` / `holeSize` /
 * `firstSliceAng` so the numeric chart-level knobs compose the same
 * way at the call site.
 */
export function resolveSideWallThickness(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

/**
 * Resolve a `backWallThickness` override.
 *
 * `undefined` → inherit the source's parsed `backWallThickness`.
 * `null`      → drop the inherited value (the writer skips
 *               `<c:backWall>` entirely — Excel falls back to no
 *               extrusion).
 * `number`    → replace. Out-of-range, `0`, or non-finite values still
 *               surface in the cloned `SheetChart` for symmetry with
 *               the other override helpers; the writer's
 *               `buildBackWallThickness` then drops them at emit time
 *               so a fresh chart matches Excel's reference
 *               serialization.
 *
 * The grammar mirrors `floorThickness` / `upDownBarsGapWidth` /
 * `gapWidth` / `holeSize` / `firstSliceAng` so the numeric chart-level
 * knobs compose the same way at the call site.
 */
export function resolveBackWallThickness(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}
