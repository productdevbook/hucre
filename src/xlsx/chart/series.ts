// ── Chart Series ──────────────────────────────────────────────────
// Per-host module for `<c:ser>` (CT_LineSer / CT_BarSer / CT_PieSer /
// CT_ScatterSer / CT_AreaSer / CT_BubbleSer, ECMA-376 Part 1, §21.2.2.x).
// Holds the reader / writer / clone helpers — every `parse*` /
// `build*` / `resolve*` / `normalize*` function for the per-series
// block, including its name, values / categories refs, color
// (`<c:spPr>`), stroke (`<a:ln>`), marker (`<c:marker>`), smooth flag,
// invert-if-negative flag, explosion (pie / doughnut), and the per-
// series data-labels block.
//
// `buildSeriesFromSource` / `mergeSeries` (the clone-side merge that
// composes per-series overrides on top of the source's series array)
// stay in chart-clone.ts because they reference the clone-only
// `CloneChartSeriesOverride` type and the local `applyOverride`
// helper. They delegate to the per-field resolvers / cloners exported
// here.

import type {
  Chart,
  ChartDataLabels,
  ChartDataLabelsInfo,
  ChartKind,
  ChartLineDashStyle,
  ChartLineStroke,
  ChartMarker,
  ChartMarkerSymbol,
  ChartSeries,
  ChartSeriesInfo,
  WriteChartKind,
} from "../../_types";
import type { XmlElement } from "../../xml/parser";
import { xmlElement, xmlEscape, xmlSelfClose } from "../../xml/writer";
import {
  EMU_PER_PT,
  STROKE_WIDTH_MAX_PT,
  STROKE_WIDTH_MIN_PT,
  VALID_DASH_STYLES,
  clampStrokeWidthPt,
  normalizeRgbHex,
} from "./shape";
import {
  applyOverride,
  childElements,
  elementText,
  findChild,
  formulaText,
  parseBoolAttr,
  readBoolAttr,
} from "./util";
import { buildSeriesDataLabels, parseDataLabels, resolveSeriesDataLabels } from "./dataLabels";
import type { CloneChartSeriesOverride } from "../chart-clone";

// ── Marker / explosion constants ──────────────────────────────────

const EXPLOSION_MAX = 400;

/** Recognized marker symbols — mirrors OOXML `ST_MarkerStyle`. */
const VALID_MARKER_SYMBOLS: ReadonlySet<ChartMarkerSymbol> = new Set([
  "none",
  "auto",
  "circle",
  "square",
  "diamond",
  "triangle",
  "x",
  "star",
  "dot",
  "dash",
  "plus",
]);

const MARKER_SIZE_MIN = 2;
const MARKER_SIZE_MAX = 72;

// ── Series options (writer) ───────────────────────────────────────

export interface SeriesOptions {
  smooth?: boolean;
  /**
   * Owning chart's family. Used to scope-guard schema-restricted
   * data-label flags such as `<c:showLeaderLines>` (pie / doughnut
   * only) so a templated chart's pin never leaks onto a chart family
   * Excel's strict validator would reject. Required so the per-family
   * caller cannot forget to thread the type through; every chart
   * builder passes `chart.type` directly.
   */
  chartType: WriteChartKind;
  /**
   * Chart-level data label defaults from {@link SheetChart.dataLabels}.
   * Used when the series itself does not specify `dataLabels`. Series
   * passing `dataLabels: false` always wins over this default.
   */
  dataLabels?: ChartDataLabels;
  /**
   * Per-series line stroke (dash pattern + width). Only meaningful for
   * line / scatter series — every other family ignores the field. The
   * OOXML schema places stroke styling inside `<c:spPr><a:ln>` which is
   * shared with the series fill color, so the writer threads the
   * stroke into the same `<c:spPr>` block whether or not a fill color
   * is set.
   */
  stroke?: ChartLineStroke;
  /**
   * Per-series marker styling. Only meaningful for line / scatter
   * series — every other family ignores the field. The OOXML schema
   * places `<c:marker>` between `<c:spPr>` and `<c:dLbls>` on
   * `CT_LineSer` / `CT_ScatterSer`, so the writer slots it there
   * regardless of which fields are populated.
   */
  marker?: ChartMarker;
  /**
   * Per-series invert-if-negative flag. Only meaningful for bar /
   * column series — every other family ignores the field. The OOXML
   * schema places `<c:invertIfNegative>` between `<c:spPr>` and
   * `<c:dLbls>` on `CT_BarSer` / `CT_Bar3DSer`, so the writer slots
   * it there. The element is only emitted when the field resolves to
   * `true` — `false` is the OOXML default and absence round-trips
   * identically.
   */
  invertIfNegative?: boolean;
  /**
   * Per-series slice explosion (percentage of the radius). Only
   * meaningful for pie / doughnut series — every other family ignores
   * the field. The OOXML schema places `<c:explosion>` between
   * `<c:spPr>` and `<c:dPt>` / `<c:dLbls>` on `CT_PieSer`. The element
   * is only emitted when the resolved value is `> 0` — `0` is the OOXML
   * default and absence round-trips identically.
   */
  explosion?: number;
}

// ── Reference qualification (writer) ──────────────────────────────

/**
 * Ensure a range reference is sheet-qualified. Excel chart `<c:f>`
 * elements accept either `Sheet1!$A$2:$A$10` or the unquoted form
 * `Sheet1!A2:A10`; the input is preserved when a sheet is already
 * present. Bare ranges like `B2:B10` are auto-qualified with the
 * owning sheet's name.
 */
export function qualifyRef(ref: string, sheetName: string): string {
  if (ref.includes("!")) return ref;
  return `${quoteSheetName(sheetName)}!${ref}`;
}

/**
 * Quote a sheet name when it contains characters Excel considers
 * unsafe in a 3D reference (whitespace, punctuation, etc.). Single
 * quotes inside the name are doubled per the OOXML spec.
 */
export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

// ── Reader ────────────────────────────────────────────────────────

/**
 * Pull the metadata fields {@link ChartSeriesInfo} surfaces out of a
 * single `<c:ser>` element. Missing pieces (no name, no categories,
 * literal numbers instead of a range) are simply omitted.
 */
export function parseSeries(ser: XmlElement, kind: ChartKind, index: number): ChartSeriesInfo {
  const out: ChartSeriesInfo = { kind, index };

  const name = parseSeriesName(ser);
  if (name !== undefined) out.name = name;

  // Numeric values land in <c:val> for most chart types; scatter and
  // bubble use <c:yVal> instead.
  const valuesRef = formulaText(findChild(ser, "val")) ?? formulaText(findChild(ser, "yVal"));
  if (valuesRef !== undefined) out.valuesRef = valuesRef;

  // Categories live in <c:cat> for category-axis charts and in
  // <c:xVal> for scatter/bubble.
  const catRef = formulaText(findChild(ser, "cat")) ?? formulaText(findChild(ser, "xVal"));
  if (catRef !== undefined) out.categoriesRef = catRef;

  const color = parseSeriesColor(ser);
  if (color !== undefined) out.color = color;

  const dLbls = findChild(ser, "dLbls");
  if (dLbls) {
    const parsed = parseDataLabels(dLbls);
    if (parsed) out.dataLabels = parsed;
  }

  // `<c:smooth>` lives on `CT_LineSer` and `CT_ScatterSer` only — every
  // other chart family rejects the element. Surface it just for those
  // two kinds so a corrupt template carrying `<c:smooth>` on a bar/pie
  // series does not silently flip a flag that the writer would never
  // emit anyway.
  if (kind === "line" || kind === "line3D" || kind === "scatter") {
    const smooth = parseSmooth(ser);
    if (smooth !== undefined) out.smooth = smooth;

    // Stroke (dash + width) lives in `<c:spPr><a:ln>`. The same
    // schema-only-on-line/scatter rule applies — bar / pie / area
    // never paint a connecting line, so surfacing a stroke field
    // there would mislead a clone consumer about what the chart
    // actually renders.
    const stroke = parseSeriesStroke(ser);
    if (stroke !== undefined) out.stroke = stroke;

    // `<c:marker>` mirrors the same scope — CT_LineSer / CT_ScatterSer
    // only. Skip the element on every other family so a stray
    // `<c:marker>` on a bar / pie / area template does not surface a
    // setting that the writer would never emit anyway.
    const marker = parseMarker(ser);
    if (marker !== undefined) out.marker = marker;
  }

  // `<c:invertIfNegative>` lives on `CT_BarSer` / `CT_Bar3DSer` only —
  // every other chart family rejects the element. Surface the flag
  // just for those two kinds so a corrupt template carrying
  // `<c:invertIfNegative>` on a line/pie/area/scatter series does not
  // silently flip a flag that the writer would never emit anyway.
  if (kind === "bar" || kind === "bar3D") {
    const invertIfNegative = parseInvertIfNegative(ser);
    if (invertIfNegative !== undefined) out.invertIfNegative = invertIfNegative;
  }

  // `<c:explosion>` lives on `CT_PieSer` only — the OOXML schema
  // shares the type across every pie-family chart (`<c:pieChart>`,
  // `<c:pie3DChart>`, `<c:doughnutChart>`, `<c:ofPieChart>`) so
  // surface the value for any of those kinds. A stray element on a
  // bar / line / area / scatter template is dropped rather than
  // surfaced — the writer would never emit it back anyway.
  if (kind === "pie" || kind === "pie3D" || kind === "doughnut" || kind === "ofPie") {
    const explosion = parseExplosion(ser);
    if (explosion !== undefined) out.explosion = explosion;
  }

  return out;
}

/**
 * Pull `<c:smooth val=".."/>` off a series element. Returns `undefined`
 * when the attribute is absent, malformed, or carries the OOXML default
 * `false` — absence and `false` round-trip identically through the
 * writer's elision logic, so collapsing them keeps the parsed shape
 * minimal.
 */
export function parseSmooth(ser: XmlElement): boolean | undefined {
  const el = findChild(ser, "smooth");
  if (!el) return undefined;
  const v = readBoolAttr(el);
  if (v !== true) return undefined;
  return true;
}

/**
 * Pull `<c:invertIfNegative val=".."/>` off a bar/column series
 * element. Returns `undefined` when the attribute is absent,
 * malformed, or carries the OOXML default `false` — absence and
 * `false` round-trip identically through the writer's elision logic,
 * so collapsing them keeps the parsed shape minimal.
 */
export function parseInvertIfNegative(ser: XmlElement): boolean | undefined {
  const el = findChild(ser, "invertIfNegative");
  if (!el) return undefined;
  const v = readBoolAttr(el);
  if (v !== true) return undefined;
  return true;
}

/**
 * Pull `<c:explosion val=".."/>` off a pie / doughnut series element.
 * The element's `val` attribute is `xsd:unsignedInt` per the OOXML
 * schema (CT_UnsignedInt) — the slice is pulled away from the chart
 * center by `val` percent of the radius. Returns `undefined` when the
 * attribute is absent, malformed, negative, or carries the OOXML
 * default `0` — absence and `0` round-trip identically through the
 * writer's elision logic, so collapsing them keeps the parsed shape
 * minimal. Non-integer input rounds to the nearest integer for parity
 * with the writer (Excel's UI accepts integer percentages only).
 */
export function parseExplosion(ser: XmlElement): number | undefined {
  const el = findChild(ser, "explosion");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const n = Number.parseFloat(raw);
  if (!Number.isFinite(n) || n < 0) return undefined;
  const rounded = Math.round(n);
  if (rounded === 0) return undefined;
  return rounded;
}

/**
 * Pull `<c:marker>` off a line / scatter series. Returns `undefined`
 * when the marker block is absent or carries no meaningful settings —
 * an empty `<c:marker/>` element collapses identically to absence
 * through the writer's elision logic, so omitting it keeps the parsed
 * shape minimal.
 *
 * Field semantics mirror {@link ChartMarker}: an unknown `<c:symbol>`
 * value is dropped (rather than surfaced), `<c:size>` outside the
 * 2..72 band is clamped, and the fill / outline colors come from
 * `<c:spPr><a:solidFill>` and `<c:spPr><a:ln><a:solidFill>`
 * respectively.
 */
export function parseMarker(ser: XmlElement): ChartMarker | undefined {
  const el = findChild(ser, "marker");
  if (!el) return undefined;

  const out: ChartMarker = {};

  const sym = findChild(el, "symbol");
  if (sym) {
    const v = sym.attrs.val;
    if (typeof v === "string" && VALID_MARKER_SYMBOLS.has(v as ChartMarkerSymbol)) {
      out.symbol = v as ChartMarkerSymbol;
    }
  }

  const sizeEl = findChild(el, "size");
  if (sizeEl) {
    const v = sizeEl.attrs.val;
    if (typeof v === "string") {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) {
        // OOXML ST_MarkerSize is `xsd:unsignedByte` constrained to
        // 2..72; clamp anything outside that band on the way in so a
        // template with an out-of-range value still round-trips.
        if (n < 2) out.size = 2;
        else if (n > 72) out.size = 72;
        else out.size = n;
      }
    }
  }

  const spPr = findChild(el, "spPr");
  if (spPr) {
    const fill = findChild(spPr, "solidFill");
    if (fill) {
      const srgb = findChild(fill, "srgbClr");
      const v = srgb?.attrs.val;
      if (typeof v === "string") {
        const hex = v.replace(/^#/, "").toUpperCase();
        if (/^[0-9A-F]{6}$/.test(hex)) out.fill = hex;
      }
    }
    const ln = findChild(spPr, "ln");
    if (ln) {
      const lnFill = findChild(ln, "solidFill");
      if (lnFill) {
        const srgb = findChild(lnFill, "srgbClr");
        const v = srgb?.attrs.val;
        if (typeof v === "string") {
          const hex = v.replace(/^#/, "").toUpperCase();
          if (/^[0-9A-F]{6}$/.test(hex)) out.line = hex;
        }
      }
    }
  }

  if (
    out.symbol === undefined &&
    out.size === undefined &&
    out.fill === undefined &&
    out.line === undefined
  ) {
    return undefined;
  }
  return out;
}

export function parseSeriesName(ser: XmlElement): string | undefined {
  const tx = findChild(ser, "tx");
  if (!tx) return undefined;
  const literal = findChild(tx, "v");
  if (literal) {
    const text = elementText(literal).trim();
    if (text.length > 0) return text;
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
    // Fall back to the formula reference itself when no cached value.
    const f = formulaText(strRef);
    if (f) return f;
  }
  return undefined;
}

export function parseSeriesColor(ser: XmlElement): string | undefined {
  const spPr = findChild(ser, "spPr");
  if (!spPr) return undefined;
  const fill = findChild(spPr, "solidFill");
  if (!fill) return undefined;
  const srgb = findChild(fill, "srgbClr");
  if (!srgb) return undefined;
  const val = srgb.attrs.val;
  if (typeof val !== "string") return undefined;
  const normalized = val.replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? normalized : undefined;
}

/**
 * Pull `<c:spPr><a:ln>` off a series and surface its dash + width as
 * a {@link ChartLineStroke}. Returns `undefined` when the block is
 * absent or carries no meaningful settings — an empty `<a:ln/>`
 * collapses identically to absence through the writer's elision
 * logic, so omitting it keeps the parsed shape minimal.
 *
 * `<a:ln>` also nests the line color (`<a:solidFill>`) which mirrors
 * the series fill — parseSeriesColor already surfaces that as
 * {@link ChartSeriesInfo.color}, so the stroke object intentionally
 * does not duplicate the field.
 */
export function parseSeriesStroke(ser: XmlElement): ChartLineStroke | undefined {
  const spPr = findChild(ser, "spPr");
  if (!spPr) return undefined;
  const ln = findChild(spPr, "ln");
  if (!ln) return undefined;

  const out: ChartLineStroke = {};

  // Stroke width is on the `w` attribute of `<a:ln>` (EMU). Convert
  // back to points and clamp to the band Excel's UI exposes so a
  // template carrying an exotic width still round-trips through the
  // writer's clamp.
  const wAttr = ln.attrs.w;
  if (typeof wAttr === "string") {
    const emu = Number.parseFloat(wAttr);
    if (Number.isFinite(emu) && emu > 0) {
      // Snap to the 0.25 pt grid Excel's UI exposes (Math.round(x * 4) / 4).
      const pt = Math.round((emu / EMU_PER_PT) * 4) / 4;
      if (pt < STROKE_WIDTH_MIN_PT) out.width = STROKE_WIDTH_MIN_PT;
      else if (pt > STROKE_WIDTH_MAX_PT) out.width = STROKE_WIDTH_MAX_PT;
      else out.width = pt;
    }
  }

  // Dash style is `<a:prstDash val="..."/>` inside `<a:ln>`.
  const dashEl = findChild(ln, "prstDash");
  if (dashEl) {
    const v = dashEl.attrs.val;
    if (typeof v === "string" && VALID_DASH_STYLES.has(v as ChartLineDashStyle)) {
      out.dash = v as ChartLineDashStyle;
    }
  }

  if (out.dash === undefined && out.width === undefined) return undefined;
  return out;
}

// ── Writer ────────────────────────────────────────────────────────

/**
 * Normalize {@link ChartSeries.explosion} for emission inside
 * `<c:explosion val=".."/>` on a pie / doughnut series.
 *
 * The OOXML schema (`CT_UnsignedInt`) accepts any non-negative integer,
 * but Excel's UI only exposes 0..400% — values outside that band render
 * but trigger Excel's repair dialog. Clamp to the UI band on the way
 * out so a round-trip stays inside the range Excel will render.
 *
 * Returns `undefined` for the default `0` (and any negative / non-finite
 * input) so the writer can elide the element entirely; absence and `0`
 * round-trip identically through the parser's collapse logic.
 */
export function clampExplosion(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded <= 0) return undefined;
  if (rounded > EXPLOSION_MAX) return EXPLOSION_MAX;
  return rounded;
}

export function buildSeries(
  series: ChartSeries,
  index: number,
  sheetName: string,
  numericCategories: boolean,
  options: SeriesOptions,
): string {
  const children: string[] = [
    xmlSelfClose("c:idx", { val: index }),
    xmlSelfClose("c:order", { val: index }),
  ];

  if (series.name) {
    // Literal series names go inside <c:tx><c:v>…</c:v></c:tx>. Excel
    // also accepts <c:strRef> for cell-bound names; literals are the
    // simpler shape and round-trip just as well.
    children.push(
      xmlElement("c:tx", undefined, [xmlElement("c:v", undefined, xmlEscape(series.name))]),
    );
  }

  // Optional fill color and / or line stroke (line / scatter only emit
  // `<a:ln>` width / dash from `options.stroke`; non-line callers leave
  // `options.stroke` undefined so the field is silently dropped on
  // every other chart family).
  const spPr = buildSeriesSpPr(series.color, options?.stroke);
  if (spPr) children.push(spPr);

  // Marker — only line/scatter series honor `<c:marker>` per the OOXML
  // schema (CT_LineSer / CT_ScatterSer). The element sits between
  // `<c:spPr>` and `<c:dLbls>`; non-line/non-scatter callers leave
  // `options.marker` undefined so the field is silently dropped on
  // every other chart family.
  const markerXml = buildSeriesMarker(options?.marker);
  if (markerXml) children.push(markerXml);

  // `<c:invertIfNegative>` — only bar / column (CT_BarSer /
  // CT_Bar3DSer) series carry the element per the OOXML schema. It
  // sits between `<c:spPr>` (and the bar-irrelevant `<c:marker>`
  // slot, which is never populated for bar/column callers anyway)
  // and `<c:dLbls>`. Non-bar callers leave `options.invertIfNegative`
  // undefined so the field is silently dropped on every other chart
  // family. Emit only when the resolved value is `true` — `false`
  // matches the OOXML default and absence round-trips identically.
  if (options?.invertIfNegative === true) {
    children.push(xmlSelfClose("c:invertIfNegative", { val: 1 }));
  }

  // `<c:explosion>` — only pie / doughnut (CT_PieSer, shared across
  // the pie family via `EG_PieSer`) series carry the element per the
  // OOXML schema. It sits between `<c:spPr>` and `<c:dPt>` / `<c:dLbls>`.
  // Non-pie callers leave `options.explosion` undefined so the field
  // is silently dropped on every other chart family. Emit only when
  // the resolved value is non-zero — `0` matches the OOXML default and
  // absence round-trips identically.
  const explosion = clampExplosion(options?.explosion);
  if (explosion !== undefined) {
    children.push(xmlSelfClose("c:explosion", { val: explosion }));
  }

  // Data labels — series-level override always wins over the chart-level
  // default. `<c:dLbls>` sits between <c:spPr> and <c:cat>/<c:val> per
  // the OOXML series schema (CT_BarSer, CT_LineSer, ...). The chart
  // type threads through so the dLbls body can scope-guard the pie /
  // doughnut-only `<c:showLeaderLines>` flag.
  const seriesDLblsXml = buildSeriesDataLabels(
    series.dataLabels,
    options.dataLabels,
    options.chartType,
  );
  if (seriesDLblsXml) children.push(seriesDLblsXml);

  // Categories (skipped for pie when omitted; allowed for all)
  if (series.categories) {
    const ref = qualifyRef(series.categories, sheetName);
    if (numericCategories) {
      children.push(
        xmlElement("c:xVal", undefined, [
          xmlElement("c:numRef", undefined, [xmlElement("c:f", undefined, xmlEscape(ref))]),
        ]),
      );
    } else {
      children.push(
        xmlElement("c:cat", undefined, [
          xmlElement("c:strRef", undefined, [xmlElement("c:f", undefined, xmlEscape(ref))]),
        ]),
      );
    }
  }

  // Values
  const valuesRef = qualifyRef(series.values, sheetName);
  if (numericCategories) {
    children.push(
      xmlElement("c:yVal", undefined, [
        xmlElement("c:numRef", undefined, [xmlElement("c:f", undefined, xmlEscape(valuesRef))]),
      ]),
    );
  } else {
    children.push(
      xmlElement("c:val", undefined, [
        xmlElement("c:numRef", undefined, [xmlElement("c:f", undefined, xmlEscape(valuesRef))]),
      ]),
    );
  }

  if (options?.smooth !== undefined) {
    children.push(xmlSelfClose("c:smooth", { val: options.smooth ? 1 : 0 }));
  }

  return xmlElement("c:ser", undefined, children);
}

/**
 * Validate a dash style against `ST_PresetLineDashVal`. Returns
 * `undefined` for unrecognized values so the writer can elide
 * `<a:prstDash>` rather than emit a token Excel will reject.
 */
export function normalizeDashStyle(
  value: ChartLineDashStyle | undefined,
): ChartLineDashStyle | undefined {
  if (value === undefined) return undefined;
  return VALID_DASH_STYLES.has(value) ? value : undefined;
}

/**
 * Build the `<c:spPr>` element shared by series fill color and series
 * line stroke. Returns `undefined` when neither field carries any
 * meaningful settings — an empty `<c:spPr/>` collapses to the
 * inherited series-rotation default Excel picks anyway, so omitting it
 * keeps untouched chart XML byte-clean.
 *
 * The OOXML `<a:ln>` element accepts both a `w` attribute (stroke
 * width in EMU) and child elements `<a:solidFill>` / `<a:prstDash>` in
 * a fixed order. When a fill color is set, the stroke also renders the
 * same color (matching Excel's "Format Data Series → Fill" default
 * which paints the line in the fill color). Stroke metadata (dash and
 * width) layers on top without overriding the line color so a `color +
 * stroke` combo behaves like Excel's UI: the line picks up the fill
 * color and the dash / width override visibility-only attributes.
 */
export function buildSeriesSpPr(
  rgbHex: string | undefined,
  stroke: ChartLineStroke | undefined,
): string | undefined {
  const fillHex = rgbHex ? rgbHex.replace(/^#/, "").toUpperCase() : undefined;
  const dash = normalizeDashStyle(stroke?.dash);
  const widthPt = clampStrokeWidthPt(stroke?.width);

  if (!fillHex && dash === undefined && widthPt === undefined) {
    return undefined;
  }

  const spPrChildren: string[] = [];
  if (fillHex) {
    spPrChildren.push(
      xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillHex })]),
    );
  }

  // `<a:ln>` carries stroke metadata. Emit it whenever a fill color is
  // set (so the connecting line picks up the same color, matching the
  // legacy behavior) or whenever stroke width / dash is configured.
  if (fillHex || dash !== undefined || widthPt !== undefined) {
    const lnAttrs: Record<string, string | number> = {};
    if (widthPt !== undefined) {
      // OOXML stores stroke width in EMU (1 pt = 12 700 EMU). Round to
      // the nearest integer because the schema types `w` as `xsd:int`.
      lnAttrs.w = Math.round(widthPt * EMU_PER_PT);
    }
    const lnChildren: string[] = [];
    if (fillHex) {
      lnChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillHex })]),
      );
    }
    if (dash !== undefined) {
      lnChildren.push(xmlSelfClose("a:prstDash", { val: dash }));
    }
    spPrChildren.push(
      lnChildren.length === 0
        ? xmlSelfClose("a:ln", lnAttrs)
        : xmlElement("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined, lnChildren),
    );
  }

  return xmlElement("c:spPr", undefined, spPrChildren);
}

/**
 * Normalize a marker size to the OOXML 2..72 band (`ST_MarkerSize`).
 * Excel's UI clamps anything outside this range; we mirror that on the
 * write side so an out-of-range hint never reaches the chart XML.
 *
 * Returns `undefined` for non-finite values so the writer can elide
 * `<c:size>` (Excel falls back to its series-rotation default).
 */
export function clampMarkerSize(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value)) return undefined;
  const rounded = Math.round(value);
  if (rounded < MARKER_SIZE_MIN) return MARKER_SIZE_MIN;
  if (rounded > MARKER_SIZE_MAX) return MARKER_SIZE_MAX;
  return rounded;
}

/**
 * Validate a marker symbol against the OOXML `ST_MarkerStyle` enum.
 * Returns `undefined` for unrecognized values so the writer can elide
 * `<c:symbol>` rather than emit a token Excel will reject.
 */
export function normalizeMarkerSymbol(
  value: ChartMarkerSymbol | undefined,
): ChartMarkerSymbol | undefined {
  if (value === undefined) return undefined;
  return VALID_MARKER_SYMBOLS.has(value) ? value : undefined;
}

/**
 * Build a `<c:marker>` element for a series. Returns `undefined` when
 * the marker block carries no meaningful settings — an empty marker
 * element collapses to the inherited series-rotation default Excel
 * picks anyway, so omitting it keeps untouched XML byte-clean.
 */
export function buildSeriesMarker(marker: ChartMarker | undefined): string | undefined {
  if (!marker) return undefined;
  const symbol = normalizeMarkerSymbol(marker.symbol);
  const size = clampMarkerSize(marker.size);
  const fill = normalizeRgbHex(marker.fill);
  const line = normalizeRgbHex(marker.line);

  if (symbol === undefined && size === undefined && !fill && !line) return undefined;

  const children: string[] = [];
  if (symbol !== undefined) children.push(xmlSelfClose("c:symbol", { val: symbol }));
  if (size !== undefined) children.push(xmlSelfClose("c:size", { val: size }));

  if (fill || line) {
    const spPrChildren: string[] = [];
    if (fill) {
      spPrChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fill })]),
      );
    }
    if (line) {
      spPrChildren.push(
        xmlElement("a:ln", undefined, [
          xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: line })]),
        ]),
      );
    }
    children.push(xmlElement("c:spPr", undefined, spPrChildren));
  }

  return xmlElement("c:marker", undefined, children);
}

// ── Clone ─────────────────────────────────────────────────────────

/**
 * Resolve a per-series line-stroke override.
 *
 * `undefined` → inherit the source series' `stroke` (a fresh shallow
 *               copy so the caller cannot mutate the parsed source).
 * `null`      → drop the inherited block.
 * object      → replace the inherited block wholesale (no per-field
 *               merge; pass the full shape you want).
 *
 * An empty stroke block (no dash, no width) collapses to `undefined`
 * so the writer can elide the element rather than emit a bare
 * `<a:ln/>` that Excel paints with the inherited default.
 */
export function resolveStroke(
  sourceStroke: ChartLineStroke | undefined,
  override: ChartLineStroke | null | undefined,
): ChartLineStroke | undefined {
  if (override === undefined) {
    if (!sourceStroke) return undefined;
    return cloneStroke(sourceStroke);
  }
  if (override === null) return undefined;
  return cloneStroke(override);
}

export function cloneStroke(source: ChartLineStroke): ChartLineStroke | undefined {
  const out: ChartLineStroke = {};
  if (source.dash !== undefined) out.dash = source.dash;
  if (typeof source.width === "number" && Number.isFinite(source.width)) out.width = source.width;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a per-series smooth-line override.
 *
 * `undefined` → inherit the source series' `smooth`.
 * `null`      → drop the inherited flag (the cloned series renders straight).
 * `boolean`   → replace.
 *
 * Only the `true` outcome materializes on the result — `false` collapses
 * to `undefined` so absence and the OOXML default round-trip identically
 * (the writer emits straight segments either way).
 */
export function resolveSmooth(
  sourceSmooth: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceSmooth === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a per-series invert-if-negative override.
 *
 * `undefined` → inherit the source series' `invertIfNegative`.
 * `null`      → drop the inherited flag (the cloned series renders
 *               negatives in the series fill color).
 * `boolean`   → replace.
 *
 * Only the `true` outcome materializes on the result — `false` collapses
 * to `undefined` so absence and the OOXML default round-trip identically
 * (the writer omits `<c:invertIfNegative>` either way).
 */
export function resolveInvertIfNegative(
  sourceFlag: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) {
    return sourceFlag === true ? true : undefined;
  }
  if (override === null) return undefined;
  return override === true ? true : undefined;
}

/**
 * Resolve a per-series slice-explosion override.
 *
 * `undefined` → inherit the source series' `explosion`.
 * `null`      → drop the inherited value (the cloned series renders
 *               flush against its neighbors).
 * `number`    → replace.
 *
 * Non-finite or non-positive numbers (and the OOXML default `0`)
 * collapse to `undefined` so absence and the default round-trip
 * identically through the writer's elision logic. Out-of-band values
 * (the writer also clamps) are passed through here — the writer
 * applies the final `0..400` clamp at emit time so a parsed-then-cloned
 * value remains visible on the resulting `SheetChart` object.
 */
export function resolveExplosion(
  sourceValue: number | undefined,
  override: number | null | undefined,
): number | undefined {
  if (override === undefined) {
    if (sourceValue === undefined || !Number.isFinite(sourceValue) || sourceValue <= 0) {
      return undefined;
    }
    return sourceValue;
  }
  if (override === null) return undefined;
  if (!Number.isFinite(override) || override <= 0) return undefined;
  return override;
}

/**
 * Resolve a per-series marker override.
 *
 * `undefined` → inherit the source series' `marker` (a fresh shallow
 * copy so the caller cannot mutate the parsed source).
 * `null`      → drop the inherited block (the cloned series falls back
 *               to Excel's series-rotation default).
 * object      → replace the inherited block wholesale.
 *
 * An empty marker block (no symbol, size, or color) collapses to
 * `undefined` so the writer can elide the element rather than emit a
 * bare `<c:marker/>` that Excel paints with the inherited default.
 */
export function resolveMarker(
  sourceMarker: ChartMarker | undefined,
  override: ChartMarker | null | undefined,
): ChartMarker | undefined {
  if (override === undefined) {
    if (!sourceMarker) return undefined;
    return cloneMarker(sourceMarker);
  }
  if (override === null) return undefined;
  return cloneMarker(override);
}

export function cloneMarker(source: ChartMarker): ChartMarker | undefined {
  const out: ChartMarker = {};
  if (source.symbol !== undefined) out.symbol = source.symbol;
  if (typeof source.size === "number" && Number.isFinite(source.size)) out.size = source.size;
  if (typeof source.fill === "string" && source.fill.length > 0) out.fill = source.fill;
  if (typeof source.line === "string" && source.line.length > 0) out.line = source.line;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve a `showLineMarkers` override.
 *
 * `undefined` → inherit the source's parsed `showLineMarkers`.
 * `null`      → drop the inherited value (the writer falls back to the
 *               Excel default — `<c:marker val="1"/>`, markers shown).
 * `boolean`   → replace.
 *
 * The grammar mirrors `upDownBars` / `dropLines` / `hiLowLines` so the
 * chart-level line-only toggles compose the same way at the call site.
 * `true` collapses to `undefined` on the writer side because the writer
 * already emits `val="1"` by default; the `true` value still surfaces
 * in the cloned `SheetChart` for symmetry with other resolve helpers,
 * leaving the renderer to fold it back into the default during emit.
 */
export function resolveShowLineMarkers(
  sourceValue: boolean | undefined,
  override: boolean | null | undefined,
): boolean | undefined {
  if (override === undefined) return sourceValue;
  if (override === null) return undefined;
  return override;
}

// ── Clone-side series merge ────────────────────────────────────────

export function buildSeriesFromSource(
  source: Chart,
  overrides: ReadonlyArray<CloneChartSeriesOverride | undefined> | undefined,
): ChartSeries[] {
  const sourceSeries = source.series ?? [];
  // The override array can be longer than the source (caller wants to
  // append a fully-specified series). Walk the union of both lengths.
  const length = Math.max(sourceSeries.length, overrides?.length ?? 0);
  const out: ChartSeries[] = [];

  for (let i = 0; i < length; i++) {
    const src: ChartSeriesInfo | undefined = sourceSeries[i];
    const ov = overrides?.[i];
    const merged = mergeSeries(src, ov, i);
    out.push(merged);
  }

  return out;
}

export function mergeSeries(
  src: ChartSeriesInfo | undefined,
  ov: CloneChartSeriesOverride | undefined,
  index: number,
): ChartSeries {
  // Resolve `values` first — it's the only mandatory field.
  const values = ov?.values ?? src?.valuesRef;
  if (!values) {
    throw new Error(
      `cloneChart: series #${index} has no values reference; provide \`seriesOverrides[${index}].values\``,
    );
  }

  const out: ChartSeries = { values };

  const name = applyOverride(src?.name, ov?.name);
  if (name !== undefined) out.name = name;

  const categories = applyOverride(src?.categoriesRef, ov?.categories);
  if (categories !== undefined) out.categories = categories;

  const color = applyOverride(src?.color, ov?.color);
  if (color !== undefined) out.color = color;

  const dataLabels = resolveSeriesDataLabels(src?.dataLabels, ov?.dataLabels);
  if (dataLabels !== undefined) out.dataLabels = dataLabels;

  const smooth = resolveSmooth(src?.smooth, ov?.smooth);
  if (smooth !== undefined) out.smooth = smooth;

  const stroke = resolveStroke(src?.stroke, ov?.stroke);
  if (stroke !== undefined) out.stroke = stroke;

  const marker = resolveMarker(src?.marker, ov?.marker);
  if (marker !== undefined) out.marker = marker;

  const invertIfNegative = resolveInvertIfNegative(src?.invertIfNegative, ov?.invertIfNegative);
  if (invertIfNegative !== undefined) out.invertIfNegative = invertIfNegative;

  const explosion = resolveExplosion(src?.explosion, ov?.explosion);
  if (explosion !== undefined) out.explosion = explosion;

  return out;
}
