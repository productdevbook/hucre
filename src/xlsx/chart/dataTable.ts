// ── Chart Data Table ──────────────────────────────────────────────
// Per-host module for `<c:plotArea><c:dTable>` (CT_DTable, ECMA-376
// Part 1, §21.2.2.54). Holds the reader / writer helpers — every
// `parse*` / `build*` / `resolve*` function for the data-table block,
// including its four boolean children, `<c:txPr>` typography, and
// `<c:spPr>` fill / border slots.
//
// The clone-side `resolveDataTable` override (takes `(source, override)`
// and returns `ChartDataTable | boolean | undefined`) lives in
// `chart-clone.ts` because its signature is shape-incompatible with the
// writer-side single-arg resolver.

import type { ChartBorderDash, ChartDataTable, SheetChart } from "../../_types";
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
import { findChild, parseBoolAttr } from "./util";
import { FONT_SIZE_MAX_PT, FONT_SIZE_MIN_PT, FONT_SZ_PER_POINT } from "./text";
import { normalizeTitleColor, normalizeTitleFontSize } from "./title";

const TITLE_FONT_SZ_PER_POINT = FONT_SZ_PER_POINT;
const TITLE_FONT_SIZE_MIN_PT = FONT_SIZE_MIN_PT;
const TITLE_FONT_SIZE_MAX_PT = FONT_SIZE_MAX_PT;

// ── Reader ────────────────────────────────────────────────────────

/**
 * Pull `<c:dTable>...</c:dTable>` off `<c:plotArea>`. Surfaces a
 * {@link ChartDataTable} whenever the source chart declares the
 * element; absence collapses to `undefined`.
 *
 * Each of the four boolean children (`<c:showHorzBorder>`,
 * `<c:showVertBorder>`, `<c:showOutline>`, `<c:showKeys>`) round-trips
 * literally — the reader does not collapse any per-field default
 * because all four are required on `CT_DTable` and Excel always emits
 * every one. Children that are missing or carry an unknown `val`
 * attribute drop to `undefined` rather than fabricate a flag the file
 * did not pin; the writer falls back to the OOXML reference defaults
 * (`true` for every child) on round-trip.
 */
export function parseDataTable(plotArea: XmlElement): ChartDataTable | undefined {
  const el = findChild(plotArea, "dTable");
  if (!el) return undefined;
  const out: ChartDataTable = {};
  const showHorzBorder = parseDataTableFlag(el, "showHorzBorder");
  if (showHorzBorder !== undefined) out.showHorzBorder = showHorzBorder;
  const showVertBorder = parseDataTableFlag(el, "showVertBorder");
  if (showVertBorder !== undefined) out.showVertBorder = showVertBorder;
  const showOutline = parseDataTableFlag(el, "showOutline");
  if (showOutline !== undefined) out.showOutline = showOutline;
  const showKeys = parseDataTableFlag(el, "showKeys");
  if (showKeys !== undefined) out.showKeys = showKeys;
  const fontSize = parseDataTableFontSize(el);
  if (fontSize !== undefined) out.fontSize = fontSize;
  const fontColor = parseDataTableFontColor(el);
  if (fontColor !== undefined) out.fontColor = fontColor;
  const bold = parseDataTableBold(el);
  if (bold !== undefined) out.bold = bold;
  const italic = parseDataTableItalic(el);
  if (italic !== undefined) out.italic = italic;
  const underline = parseDataTableUnderline(el);
  if (underline !== undefined) out.underline = underline;
  const strikethrough = parseDataTableStrikethrough(el);
  if (strikethrough !== undefined) out.strikethrough = strikethrough;
  const fontFamily = parseDataTableFontFamily(el);
  if (fontFamily !== undefined) out.fontFamily = fontFamily;
  const fillColor = parseDataTableFillColor(el);
  if (fillColor !== undefined) out.fillColor = fillColor;
  const borderColor = parseDataTableBorderColor(el);
  if (borderColor !== undefined) out.borderColor = borderColor;
  // `<c:dTable><c:spPr><a:ln w="EMU">` carries Excel's "Format Data
  // Table -> Border -> Width" pin. Delegates to the shared
  // {@link parseBorderWidthFromSpPr} so the EMU encoding and snap /
  // clamp grammar match every other chart-frame border-width slot.
  const borderWidth = parseBorderWidthFromSpPr(el);
  if (borderWidth !== undefined) out.borderWidth = borderWidth;
  // `<c:dTable><c:spPr><a:ln><a:prstDash val=".."/>` carries Excel's
  // "Format Data Table -> Border -> Dash type" pin. Delegates to the
  // shared {@link parseBorderDashFromSpPr} so the accept-or-drop
  // grammar matches every chart-frame border-dash slot.
  const borderDash = parseBorderDashFromSpPr(el);
  if (borderDash !== undefined) out.borderDash = borderDash;
  return out;
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the bold flag.
 *
 * The OOXML `b` attribute is the `xsd:boolean` bold flag on
 * `CT_TextCharacterProperties`. Only an explicit `b="1"` (or `"true"`)
 * surfaces `true`; the OOXML default `0` (and absence / malformed
 * tokens) collapses to `undefined` so absence and the default
 * round-trip identically through `cloneChart`. Mirrors the
 * chart-title / axis-title / axis tick-label / legend / data-label
 * bold readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
export function parseDataTableBold(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.b;
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
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the italic
 * flag.
 *
 * The OOXML `i` attribute is the `xsd:boolean` italic flag on
 * `CT_TextCharacterProperties`. An explicit `i="1"` / `"true"`
 * surfaces `true`; an explicit `i="0"` / `"false"` surfaces `false`
 * so a templated chart that pinned `i="0"` round-trips literally and
 * a clone target can override an upstream `i="1"` cleanly. Absence /
 * malformed tokens collapse to `undefined` so a fresh data table that
 * never pinned the flag round-trips identically through `cloneChart`.
 * Mirrors the data-table bold reader (and the chart-title /
 * axis-title / axis tick-label / legend / data-label italic readers
 * for the OOXML attribute layout) so a parsed value slots straight
 * back into the writer's emit path.
 */
export function parseDataTableItalic(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.i;
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
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the underline
 * flag.
 *
 * The OOXML `u` attribute is the `ST_TextUnderlineType` enumeration
 * on `CT_TextCharacterProperties`. `u="sng"` (Excel's UI variant —
 * single underline) surfaces `true`; `u="none"` (the OOXML default)
 * surfaces `false` so a templated chart that pinned `u="none"`
 * round-trips literally and a clone target can override an upstream
 * `u="sng"` cleanly. The schema's other variants (`"dbl"`, `"heavy"`,
 * `"dotted"`, `"dotDash"`, `"wavy"`, etc.) and absence / malformed
 * tokens collapse to `undefined` so a fresh data table that never
 * pinned the flag round-trips identically through `cloneChart`. The
 * reader never silently downgrades a non-`"sng"` underline to a
 * single line on round-trip; only the boolean shape Excel's UI
 * exposes survives the parse.
 *
 * Mirrors the data-table bold reader pattern — `true` and `false`
 * round-trip explicitly so a clone can override either direction —
 * while the OOXML attribute layout matches the chart-title /
 * axis-title / axis tick-label / legend / data-label underline
 * readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
export function parseDataTableUnderline(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
  if (!txPr) return undefined;
  const p = findChild(txPr, "p");
  if (!p) return undefined;
  const pPr = findChild(p, "pPr");
  if (!pPr) return undefined;
  const defRPr = findChild(pPr, "defRPr");
  if (!defRPr) return undefined;
  const raw = defRPr.attrs.u;
  if (raw === "sng") return true;
  if (raw === "none") return false;
  return undefined;
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr strike=".."/></a:pPr>
 * </a:p></c:txPr></c:dTable>` off a data-table block. Returns the
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
 * legend / data-label strikethrough readers exactly so a parsed value
 * slots straight back into the writer's emit path.
 */
export function parseDataTableStrikethrough(dTable: XmlElement): boolean | undefined {
  const txPr = findChild(dTable, "txPr");
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
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:txPr></c:dTable>` off
 * a data-table block. Returns the typeface string the table was
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
 * Returns `undefined` whenever the data-table block omits `<c:txPr>`
 * entirely or the canonical `<a:p><a:pPr><a:defRPr><a:latin>` chain
 * is malformed at any link. Mirrors the chart-title / axis-title /
 * axis tick-label / legend / data-label font family readers exactly
 * so a parsed value slots straight back into the writer's emit path.
 */
export function parseDataTableFontFamily(dTable: XmlElement): string | undefined {
  const txPr = findChild(dTable, "txPr");
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
 * Pull `<c:dTable><c:spPr><a:solidFill><a:srgbClr val=".."/>
 * </a:solidFill></c:spPr></c:dTable>` off a data-table block.
 * Returns the data-table background fill color as a 6-character
 * uppercase hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the fill
 * choice of `<c:spPr>` (`CT_ShapeProperties`, §20.1.2.3.13). The
 * `<c:spPr>` slot sits inside `<c:dTable>` after the four required
 * boolean children (`<c:showHorzBorder>`, `<c:showVertBorder>`,
 * `<c:showOutline>`, `<c:showKeys>`) and before the optional
 * `<c:txPr>` per CT_DTable (ECMA-376 Part 1, §21.2.2.54).
 *
 * The reader surfaces only the literal `<a:srgbClr>` form — absence,
 * non-solid fills (`<a:noFill>` / `<a:gradFill>` / `<a:pattFill>` /
 * `<a:blipFill>`), and theme-color references (`<a:schemeClr>`) all
 * collapse to `undefined` so a chart that pinned a fill the writer
 * cannot reproduce on emit drops the field rather than fabricate one
 * Excel would render differently. Malformed `val` tokens (wrong
 * length, non-hex characters, alpha-channel forms, non-string escapes)
 * likewise drop to `undefined`.
 *
 * Mirrors the writer-side {@link ChartDataTable.fillColor} so a parsed
 * value slots straight into {@link cloneChart} without conversion.
 * The lookup is scoped to direct children of `<c:dTable>` so a stray
 * `<c:spPr>` elsewhere (e.g. on `<c:plotArea>` / `<c:legend>` /
 * `<c:title>` / a series) cannot leak into this field. Mirrors
 * {@link parsePlotAreaFillColor} / {@link parseLegendFillColor} —
 * same `<c:spPr><a:solidFill><a:srgbClr>` chain on a different host
 * element. Independent of {@link parseDataTableBorderColor}: the fill
 * lives on `<c:dTable><c:spPr><a:solidFill>`, the stroke lives on
 * `<c:dTable><c:spPr><a:ln><a:solidFill>` — the two readers walk
 * disjoint children of the same `<c:spPr>` block so a caller can pin
 * both knobs without conflict.
 */
export function parseDataTableFillColor(dTable: XmlElement): string | undefined {
  return parseSpPrFill(dTable);
}

/**
 * Pull `<c:dTable><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:dTable>` off a data-table block.
 * Returns the data-table border (line) color as a 6-character uppercase
 * hex string.
 *
 * The OOXML `<a:srgbClr>` element carries the literal sRGB color
 * (`CT_SRgbColor`, ECMA-376 Part 1, §20.1.2.3.32) inside the line's
 * solid fill choice (`CT_LineProperties`'s solid fill — §20.1.2.3.24).
 * The `<a:ln>` slot sits inside the `<c:spPr>` block on `<c:dTable>`
 * alongside the optional `<a:solidFill>` fill child, in
 * `CT_ShapeProperties` schema order (fill before stroke). The
 * `<c:spPr>` itself sits between the four required boolean children
 * (`<c:showHorzBorder>`, `<c:showVertBorder>`, `<c:showOutline>`,
 * `<c:showKeys>`) and the optional `<c:txPr>` per CT_DTable
 * (ECMA-376 Part 1, §21.2.2.54).
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
 * Mirrors the writer-side {@link ChartDataTable.borderColor} so a
 * parsed value slots straight into {@link cloneChart} without
 * conversion. The lookup is scoped to direct children of `<c:dTable>`
 * so a stray `<c:spPr>` elsewhere (e.g. on `<c:plotArea>` /
 * `<c:legend>` / `<c:title>` / a series) cannot leak in. Mirrors
 * {@link parseDataTableFillColor} — same `<c:spPr>` host element on
 * the same `<c:dTable>` parent — but lands on the line
 * (`<a:ln><a:solidFill>`) child rather than the fill (`<a:solidFill>`)
 * child.
 */
export function parseDataTableBorderColor(dTable: XmlElement): string | undefined {
  return parseSpPrBorderColor(dTable);
}

/**
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the
 * 6-character uppercase hex string when the parser walks the full
 * chain and lands on an `<a:srgbClr val="RRGGBB"/>`. Theme references
 * (`<a:schemeClr>`), `<a:hslClr>`, `<a:sysClr>`, and `<a:prstClr>`
 * all collapse to `undefined` — only the literal RGB triple round-
 * trips losslessly through {@link writeChart}. Malformed `val` tokens
 * (wrong length, non-hex characters) likewise drop to `undefined`
 * rather than fabricate a value the writer would round-trip into a
 * malformed `<a:srgbClr>`.
 *
 * Returns `undefined` whenever the data-table block omits `<c:txPr>`
 * entirely or the canonical chain is malformed at any link. Mirrors
 * the chart-title / axis-title / axis tick-label / legend / data-label
 * color readers exactly so a parsed value slots straight back into the
 * writer's emit path.
 */
export function parseDataTableFontColor(dTable: XmlElement): string | undefined {
  const txPr = findChild(dTable, "txPr");
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
 * Pull `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
 * </c:txPr></c:dTable>` off a data-table block. Returns the font size
 * in points (`1..400`), or `undefined` when the element is absent /
 * the chain is malformed at any link / the surfaced value is out of
 * the supported range.
 *
 * The OOXML attribute is in 100ths of a point on
 * `CT_TextCharacterProperties`' `sz` slot (ECMA-376 Part 1,
 * §21.1.2.3.7); the reader divides by 100 at parse time so the
 * surfaced value matches what the user sees in Excel's UI. Mirrors
 * the chart-title / axis-title / axis tick-label / legend / data-label
 * font-size readers exactly so a parsed value slots straight back into
 * the writer's emit path.
 */
export function parseDataTableFontSize(dTable: XmlElement): number | undefined {
  const txPr = findChild(dTable, "txPr");
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
  // 0.5pt to match the granularity Excel's UI exposes. Mirrors the
  // chart-title / axis-title / tick-label / legend / data-label
  // sibling parsers exactly so a parsed value flows through every
  // typography slot without bookkeeping the units.
  const halfSteps = Math.round((parsed / TITLE_FONT_SZ_PER_POINT) * 2);
  const points = halfSteps / 2;
  if (points < TITLE_FONT_SIZE_MIN_PT || points > TITLE_FONT_SIZE_MAX_PT) return undefined;
  return points;
}

/**
 * Pull a single boolean child off `<c:dTable>`. Accepts the OOXML
 * truthy / falsy spellings (`"1"` / `"true"` / `"0"` / `"false"`);
 * unknown tokens, missing `val` attributes, and missing elements all
 * collapse to `undefined` rather than fabricate a flag the file did
 * not pin.
 */
export function parseDataTableFlag(dTable: XmlElement, local: string): boolean | undefined {
  const el = findChild(dTable, local);
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

// ── Writer ────────────────────────────────────────────────────────

/**
 * Resolve the {@link SheetChart.dataTable} field into the four boolean
 * children `<c:dTable>` requires, or `undefined` to signal that the
 * writer should skip emission of the element.
 *
 * Returns `undefined` when:
 *  - The chart's family has no axes (pie / doughnut). The OOXML schema
 *    places `<c:dTable>` inside `<c:plotArea>` alongside the axes, so
 *    no axes means no slot to host the element.
 *  - The caller did not opt in (`dataTable` is `undefined` or `false`).
 *
 * Returns the four resolved booleans when the caller passed `true`
 * (every default `true`) or an object (per-field overrides on top of the
 * `true` defaults). Stray non-boolean inputs collapse to the matching
 * default rather than emit a token Excel rejects, mirroring how every
 * other chart-level boolean writer treats its input.
 */
export function resolveDataTable(chart: SheetChart):
  | {
      showHorzBorder: boolean;
      showVertBorder: boolean;
      showOutline: boolean;
      showKeys: boolean;
      fontSize: number | undefined;
      fontColor: string | undefined;
      bold: boolean | undefined;
      italic: boolean | undefined;
      underline: boolean | undefined;
      strikethrough: boolean | undefined;
      fontFamily: string | undefined;
      fillColor: string | undefined;
      borderColor: string | undefined;
      borderWidth: number | undefined;
      borderDash: ChartBorderDash | undefined;
    }
  | undefined {
  // Pie / doughnut have no axes — the OOXML schema places `<c:dTable>`
  // alongside `<c:catAx>` / `<c:valAx>`, so there is no slot for it on
  // those families. Drop the field silently rather than emit an element
  // Excel's strict validator would reject.
  if (chart.type === "pie" || chart.type === "doughnut") return undefined;

  const raw = chart.dataTable;
  if (raw === undefined || raw === false) return undefined;

  if (raw === true) {
    return {
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: true,
      fontSize: undefined,
      fontColor: undefined,
      bold: undefined,
      italic: undefined,
      underline: undefined,
      strikethrough: undefined,
      fontFamily: undefined,
      fillColor: undefined,
      borderColor: undefined,
      borderWidth: undefined,
      borderDash: undefined,
    };
  }

  // Per-field overrides on top of the `true` defaults. Only literal
  // `false` flips a flag — anything else (including stray `undefined`,
  // `null`, or a non-boolean) falls back to the default `true` so the
  // writer never emits a value the OOXML schema would refuse.
  return {
    showHorzBorder: raw.showHorzBorder !== false,
    showVertBorder: raw.showVertBorder !== false,
    showOutline: raw.showOutline !== false,
    showKeys: raw.showKeys !== false,
    fontSize: resolveDataTableFontSize(raw.fontSize),
    fontColor: resolveDataTableFontColor(raw.fontColor),
    bold: resolveDataTableBold(raw.bold),
    italic: resolveDataTableItalic(raw.italic),
    underline: resolveDataTableUnderline(raw.underline),
    strikethrough: resolveDataTableStrikethrough(raw.strikethrough),
    fontFamily: resolveDataTableFontFamily(raw.fontFamily),
    fillColor: resolveDataTableFillColor(raw.fillColor),
    borderColor: resolveDataTableBorderColor(raw.borderColor),
    borderWidth: clampStrokeWidthPt(raw.borderWidth),
    borderDash: normalizeBorderDash(raw.borderDash),
  };
}

/**
 * Resolve `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr sz="N"/></a:pPr></a:p>
 * </c:txPr></c:dTable>` from {@link ChartDataTable.fontSize}.
 *
 * Returns the size in points (`1..400`), or `undefined` when the caller
 * leaves the field unset / passed an out-of-range or non-numeric token.
 * Delegates to {@link normalizeTitleFontSize} so the chart-title /
 * axis-title / axis tick-label / legend / data-label / data-table
 * font-size resolvers share the same range, the same fractional rounding
 * (Excel's UI step is 0.5pt), and the same OOXML conversion (100ths of
 * a point at emit time).
 */
export function resolveDataTableFontSize(value: number | undefined): number | undefined {
  return normalizeTitleFontSize(value);
}

/**
 * Resolve `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>
 * <a:srgbClr val="RRGGBB"/></a:solidFill></a:defRPr></a:pPr></a:p>
 * </c:txPr></c:dTable>` from {@link ChartDataTable.fontColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches the chart-title /
 * axis-title / axis tick-label / legend / data-label color resolvers
 * exactly.
 */
export function resolveDataTableFontColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

/**
 * Resolve `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr b=".."/></a:pPr>
 * </a:p></c:txPr></c:dTable>` from {@link ChartDataTable.bold}.
 *
 * Returns the bold flag, or `undefined` when the caller leaves the
 * field unset / passed a non-boolean token. Mirrors the chart-title /
 * axis-title / axis tick-label / legend / data-label bold resolvers
 * exactly — only literal `true` / `false` pass through; non-boolean
 * tokens (typed escapes from an untyped caller) collapse to
 * `undefined`.
 */
export function resolveDataTableBold(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr i=".."/></a:pPr>
 * </a:p></c:txPr></c:dTable>` from {@link ChartDataTable.italic}.
 *
 * Returns the italic flag, or `undefined` when the caller leaves the
 * field unset / passed a non-boolean token. Mirrors the chart-title /
 * axis-title / axis tick-label / legend / data-label italic resolvers
 * (and the data-table bold resolver) exactly — only literal `true` /
 * `false` pass through; non-boolean tokens (typed escapes from an
 * untyped caller) collapse to `undefined`.
 */
export function resolveDataTableItalic(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr u=".."/></a:pPr>
 * </a:p></c:txPr></c:dTable>` from {@link ChartDataTable.underline}.
 *
 * Returns the underline flag, or `undefined` when the caller leaves
 * the field unset / passed a non-boolean token. Mirrors the
 * chart-title / axis-title / axis tick-label / legend / data-label
 * underline resolvers exactly — only literal `true` / `false` pass
 * through; non-boolean tokens (typed escapes from an untyped caller)
 * collapse to `undefined`. The writer translates `true` into
 * `u="sng"` (Excel's UI variant — single underline) and `false` into
 * `u="none"` at emit time.
 */
export function resolveDataTableUnderline(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  if (value === false) return false;
  return undefined;
}

/**
 * Resolve `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr strike=".."/>
 * </a:pPr></a:p></c:txPr></c:dTable>` from
 * {@link ChartDataTable.strikethrough}.
 *
 * Returns `true` when the caller pins the strikethrough flag literally;
 * every other value (explicit `false`, absence, non-boolean tokens
 * leaking past the type guard) collapses to `undefined` so the writer
 * never emits a `strike` attribute below `"sngStrike"`. The OOXML
 * default `"noStrike"` is functionally identical to absence — the
 * writer keeps the surfaced shape consistent with what Excel's UI
 * authors (`"sngStrike"` only, never `"noStrike"` or `"dblStrike"`),
 * mirroring how `resolveTitleStrike` / `resolveLegendStrikethrough` /
 * `resolveDataLabelsStrikethrough` land on their `<a:defRPr>` slots.
 */
export function resolveDataTableStrikethrough(value: boolean | undefined): boolean | undefined {
  if (value === true) return true;
  return undefined;
}

/**
 * Resolve `<c:dTable><c:txPr><a:p><a:pPr><a:defRPr><a:latin
 * typeface=".."/></a:defRPr></a:pPr></a:p></c:txPr></c:dTable>` from
 * {@link ChartDataTable.fontFamily}.
 *
 * Returns the trimmed typeface string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed an
 * empty / whitespace-only / non-string token. Mirrors the chart-title
 * / axis-title / axis tick-label / legend / data-label font family
 * resolvers exactly so a single configuration call threads cleanly
 * through every typography slot Excel exposes.
 */
export function resolveDataTableFontFamily(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (trimmed.length === 0) return undefined;
  return trimmed;
}

/**
 * Resolve `<c:dTable><c:spPr><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></c:spPr></c:dTable>` from
 * {@link ChartDataTable.fillColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches the chart-title /
 * plot-area / legend / chart-space / axis-title fill color resolvers
 * exactly. The `<c:spPr>` slot lives between the four required
 * boolean children and the optional `<c:txPr>` per CT_DTable
 * (ECMA-376 Part 1, §21.2.2.54), distinct from the `<c:txPr>` block
 * that carries {@link ChartDataTable.fontColor}.
 */
export function resolveDataTableFillColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

/**
 * Resolve `<c:dTable><c:spPr><a:ln><a:solidFill><a:srgbClr val="RRGGBB"/>
 * </a:solidFill></a:ln></c:spPr></c:dTable>` from
 * {@link ChartDataTable.borderColor}.
 *
 * Returns the 6-character uppercase hex string the writer emits, or
 * `undefined` when the caller leaves the field unset / passed a
 * malformed token. Delegates to {@link normalizeTitleColor} so the
 * accept-with-or-without-`#` grammar matches every other `<a:srgbClr>`
 * fill / line slot exactly. Composes independently with
 * {@link ChartDataTable.fillColor} — the two knobs share the same
 * `<c:spPr>` host on `<c:dTable>` but land on different children
 * (`<a:solidFill>` for the fill, `<a:ln><a:solidFill>` for the stroke).
 * Mirrors the chart-title / axis-title / chart-space / plot-area /
 * legend `<c:spPr>` border slots — same hex grammar, same `<a:ln>`
 * slot on the `CT_ShapeProperties` schema — but lands on `<c:dTable>`'s
 * own `<c:spPr>` block.
 */
export function resolveDataTableBorderColor(value: string | undefined): string | undefined {
  return normalizeTitleColor(value);
}

/**
 * Serialize a resolved data-table into `<c:dTable>` with its four
 * required boolean children, in the order CT_DTable mandates:
 * `showHorzBorder`, `showVertBorder`, `showOutline`, `showKeys`. When
 * a typography knob is pinned the writer emits an additional `<c:txPr>`
 * block after the four booleans; absence collapses to the existing
 * four-child shape so a fresh chart with no typography pin matches
 * Excel's reference serialization byte-for-byte.
 *
 * The writer always emits all four boolean children — the OOXML schema
 * marks them required on `CT_DTable`, and Excel's reference
 * serialization includes every one even when the caller leaves it at
 * the default. The optional `<c:spPr>` / `<c:extLst>` children are
 * skipped because hucre's data-table model does not surface fill /
 * extension styling yet.
 */
export function buildDataTable(table: {
  showHorzBorder: boolean;
  showVertBorder: boolean;
  showOutline: boolean;
  showKeys: boolean;
  fontSize: number | undefined;
  fontColor: string | undefined;
  bold: boolean | undefined;
  italic: boolean | undefined;
  underline: boolean | undefined;
  strikethrough: boolean | undefined;
  fontFamily: string | undefined;
  fillColor: string | undefined;
  borderColor: string | undefined;
  borderWidth: number | undefined;
  borderDash: ChartBorderDash | undefined;
}): string {
  const children: string[] = [
    xmlSelfClose("c:showHorzBorder", { val: table.showHorzBorder ? 1 : 0 }),
    xmlSelfClose("c:showVertBorder", { val: table.showVertBorder ? 1 : 0 }),
    xmlSelfClose("c:showOutline", { val: table.showOutline ? 1 : 0 }),
    xmlSelfClose("c:showKeys", { val: table.showKeys ? 1 : 0 }),
  ];
  // CT_DTable schema places `<c:spPr>` after the four required
  // boolean children, before `<c:txPr>` and the optional `<c:extLst>`
  // (ECMA-376 Part 1, §21.2.2.54). The writer skips emission entirely
  // when no fill / border knob is pinned so a fresh chart matches
  // Excel's reference serialization byte-for-byte.
  const spPrXml = buildDataTableSpPr(
    table.fillColor,
    table.borderColor,
    table.borderWidth,
    table.borderDash,
  );
  if (spPrXml !== undefined) children.push(spPrXml);
  // CT_DTable schema places `<c:txPr>` after `<c:spPr>` and before
  // the optional `<c:extLst>` (ECMA-376 Part 1, §21.2.2.54). The writer
  // skips emission entirely when no typography knob is pinned so a
  // fresh chart matches Excel's reference serialization byte-for-byte.
  const txPrXml = buildDataTableTxPr(
    table.fontSize,
    table.fontColor,
    table.bold,
    table.italic,
    table.underline,
    table.strikethrough,
    table.fontFamily,
  );
  if (txPrXml !== undefined) children.push(txPrXml);
  return xmlElement("c:dTable", undefined, children);
}

/**
 * Build the optional `<c:spPr>` block inside `<c:dTable>`. Surfaces
 * the solid fill color knob ({@link ChartDataTable.fillColor}) and
 * the border (line) color knob ({@link ChartDataTable.borderColor}) —
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
 * Mirrors {@link buildPlotAreaSpPr} / {@link buildChartSpaceSpPr}
 * but on a distinct host element — the data-table fill / stroke
 * paint the background and outline of the table grid, while the
 * plot-area / chart-space variants paint the inner band / entire
 * chart frame.
 */
export function buildDataTableSpPr(
  fillColor: string | undefined,
  borderColor: string | undefined,
  borderWidthPt: number | undefined,
  borderDash: ChartBorderDash | undefined,
): string | undefined {
  if (
    fillColor === undefined &&
    borderColor === undefined &&
    borderWidthPt === undefined &&
    borderDash === undefined
  ) {
    return undefined;
  }
  const children: string[] = [];
  if (fillColor !== undefined) {
    children.push(
      xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillColor })]),
    );
  }
  if (borderColor !== undefined || borderWidthPt !== undefined || borderDash !== undefined) {
    const lnAttrs: Record<string, string | number> = {};
    if (borderWidthPt !== undefined) {
      // OOXML stores stroke width in EMU (1 pt = 12 700 EMU). Round to
      // the nearest integer because the schema types `w` as `xsd:int`.
      lnAttrs.w = Math.round(borderWidthPt * EMU_PER_PT);
    }
    const lnChildren: string[] = [];
    if (borderColor !== undefined) {
      lnChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: borderColor })]),
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
 * Build the `<c:txPr>` block that carries a data-table's typography
 * pins (currently font size and font color). Returns `undefined` when
 * every input is unset so the caller can elide the element entirely
 * (Excel's reference serialization omits `<c:txPr>` from `<c:dTable>`
 * when the table renders at the theme-default style).
 *
 * The emitted block mirrors the minimal `<c:txPr>` shape Excel writes
 * when the user pins a data-table typography knob — `<a:bodyPr/>`,
 * `<a:lstStyle/>`, and the `<a:p><a:pPr><a:defRPr sz="N"/></a:pPr>
 * <a:endParaRPr/></a:p>` paragraph stub Excel always emits hosts the
 * typography attributes on `<a:defRPr>`. The `<a:defRPr>` element
 * expands from self-closing to wrapping a single `<a:solidFill>` child
 * when a color is set; otherwise the writer keeps the existing
 * self-closing form so a fresh chart with no custom color matches
 * Excel's reference serialization byte-for-byte. Mirrors the
 * chart-title / axis-title / tick-label / legend / data-label
 * `<c:txPr>` slots exactly so a re-parse picks the values off the
 * canonical default-paragraph slot every other typography reader
 * expects.
 */
export function buildDataTableTxPr(
  fontSizePt: number | undefined,
  rgbHex: string | undefined,
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
    return undefined;
  const defRPrAttrs: Record<string, string | number> = {};
  if (fontSizePt !== undefined) defRPrAttrs.sz = fontSizePt * TITLE_FONT_SZ_PER_POINT;
  if (bold !== undefined) defRPrAttrs.b = bold ? 1 : 0;
  if (italic !== undefined) defRPrAttrs.i = italic ? 1 : 0;
  if (underline !== undefined) defRPrAttrs.u = underline ? "sng" : "none";
  // Strikethrough rides as `strike="sngStrike"` on the same
  // `<a:defRPr>` slot. Absence collapses to omitting the attribute
  // entirely (the OOXML default `"noStrike"` is functionally
  // identical to absence — the reader collapses both to `undefined`).
  // The writer never emits `"noStrike"` or `"dblStrike"` so the
  // surfaced shape stays consistent with Excel's UI checkbox.
  if (strikethrough === true) defRPrAttrs.strike = "sngStrike";
  // OOXML's `<a:defRPr><a:solidFill><a:srgbClr val="RRGGBB"/>
  // </a:solidFill></a:defRPr>` carries the data-table font color.
  // Absence (`undefined`) collapses to skipping the `<a:solidFill>`
  // child entirely so the data table inherits the theme text color
  // (Excel's reference behavior for fresh data tables that have not
  // had a custom color picked).
  const solidFillChild = rgbHex
    ? xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: rgbHex })])
    : undefined;
  // OOXML's `<a:defRPr><a:latin typeface=".."/></a:defRPr>` carries
  // the data-table font family. The `<a:latin>` element follows
  // `<a:solidFill>` per the CT_TextCharacterProperties child sequence
  // (ECMA-376 Part 1, §21.1.2.3.7). Absence (`undefined`) collapses
  // to omitting the entire `<a:latin>` element so the data table
  // inherits the theme typeface (Excel's reference behavior for
  // fresh data tables that have not had a custom font picked).
  const latinChild = fontFamily ? xmlSelfClose("a:latin", { typeface: fontFamily }) : undefined;
  // When a fill color or a typeface is set the `<a:defRPr>` slot
  // expands from self-closing to wrapping the children; otherwise the
  // writer keeps the existing self-closing form so a fresh chart with
  // no custom color or font matches Excel's reference serialization
  // byte-for-byte. Children are emitted in
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

// ── Clone resolvers (3-arg source/override) ───────────────────────

/**
 * Resolve a `dataTable` (plot-area data-table) override.
 *
 * `undefined` → inherit the source's parsed {@link Chart.dataTable}.
 * `null`      → drop the inherited block so the writer skips
 *               `<c:dTable>` entirely (no data table rendered).
 * `false`     → equivalent to `null` (suppression); kept distinct in
 *               the API surface so callers can write `dataTable: false`
 *               for symmetry with the writer's `boolean | object` shape.
 * `true`      → enable with the OOXML reference defaults (every flag
 *               `true`).
 * `object`    → replace the inherited block wholesale (no per-field
 *               merge with the source — pass every flag the cloned
 *               table should render). Each unspecified field falls back
 *               to `true` at the writer side because every `<c:dTable>`
 *               boolean child is required on `CT_DTable` and Excel
 *               always emits all four.
 *
 * The grammar mirrors {@link CloneChartSeriesOverride.marker} (and the
 * other `object | null` / wholesale-replace patterns) so the
 * chart-level block toggles compose the same way at the call site.
 *
 * The caller already short-circuits this for pie / doughnut clones
 * because the OOXML schema places `<c:dTable>` inside `<c:plotArea>`
 * alongside the axes, and pie / doughnut have no axes at all.
 */
export function resolveCloneDataTable(
  sourceValue: ChartDataTable | undefined,
  override: ChartDataTable | boolean | null | undefined,
): ChartDataTable | boolean | undefined {
  if (override === undefined) {
    // Inherit — pass the source through verbatim. The writer accepts
    // both the boolean and object shapes, so a parsed `ChartDataTable`
    // round-trips directly.
    return sourceValue;
  }
  if (override === null) {
    // Drop the inherited block. The writer treats `undefined` as
    // suppression and skips `<c:dTable>` entirely.
    return undefined;
  }
  if (override === false) {
    // Symmetric with `null` — kept distinct in the API surface for
    // ergonomic alignment with the writer's `boolean | object` shape,
    // but emits the same on-the-wire result (no `<c:dTable>`).
    return undefined;
  }
  // `true` or a {@link ChartDataTable} object — replace the inherited
  // block wholesale. The writer accepts both forms and falls back to
  // the OOXML reference defaults for any field the object leaves unset.
  return override;
}
