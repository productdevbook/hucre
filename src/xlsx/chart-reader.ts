// ── Chart Reader ──────────────────────────────────────────────────
// Parses xl/charts/chartN.xml into a minimal structured record.
//
// Charts in OOXML live under the `c` (DrawingML chart) namespace. The
// root element is `<c:chartSpace>` and the visible chart sits inside
// `<c:chart>`. The `<c:plotArea>` child contains one or more chart-type
// elements (`<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`, ...) — each
// chart-type element holds the series and axis bindings.
//
// This reader only extracts metadata that's cheap to surface: the chart
// kind(s), title, and series count. It does not decode series bindings.
//
// OOXML reference: ECMA-376 Part 1, §21.2 (DrawingML — Charts).

import type {
  Chart,
  ChartAxisGridlines,
  ChartAxisInfo,
  ChartBarGrouping,
  ChartDataLabelPosition,
  ChartDataLabelsInfo,
  ChartKind,
  ChartLegendPosition,
  ChartLineAreaGrouping,
  ChartSeriesInfo,
} from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";

/** All chart-type element local names recognized by Excel. */
const CHART_KIND_TAGS: ReadonlyMap<string, ChartKind> = new Map([
  ["barChart", "bar"],
  ["bar3DChart", "bar3D"],
  ["lineChart", "line"],
  ["line3DChart", "line3D"],
  ["pieChart", "pie"],
  ["pie3DChart", "pie3D"],
  ["doughnutChart", "doughnut"],
  ["areaChart", "area"],
  ["area3DChart", "area3D"],
  ["scatterChart", "scatter"],
  ["bubbleChart", "bubble"],
  ["radarChart", "radar"],
  ["surfaceChart", "surface"],
  ["surface3DChart", "surface3D"],
  ["stockChart", "stock"],
  ["ofPieChart", "ofPie"],
]);

/**
 * Parse a chart file (`xl/charts/chartN.xml`) into a {@link Chart}.
 *
 * Returns `undefined` when the document is not recognizable as a
 * `c:chartSpace`. Returns a record with `kinds: []` when the chart has
 * no chart-type element (extremely rare, but possible for empty charts).
 */
export function parseChart(xml: string): Chart | undefined {
  const root = parseXml(xml);
  // chartSpace can be the root, or wrapped if the file has been
  // pre-processed; tolerate both shapes.
  const chartSpace = root.local === "chartSpace" ? root : findDescendant(root, "chartSpace");
  if (!chartSpace) return undefined;

  const chartEl = findChild(chartSpace, "chart");
  if (!chartEl) return { kinds: [], seriesCount: 0 };

  const out: Chart = { kinds: [], seriesCount: 0 };

  const title = parseTitle(chartEl);
  if (title !== undefined) out.title = title;

  const plotArea = findChild(chartEl, "plotArea");
  if (plotArea) {
    let seriesCount = 0;
    const series: ChartSeriesInfo[] = [];
    let barGrouping: ChartBarGrouping | undefined;
    let lineGrouping: ChartLineAreaGrouping | undefined;
    let areaGrouping: ChartLineAreaGrouping | undefined;
    let chartLevelLabels: ChartDataLabelsInfo | undefined;
    let holeSize: number | undefined;
    let firstSliceAng: number | undefined;
    for (const child of childElements(plotArea)) {
      const kind = CHART_KIND_TAGS.get(child.local);
      if (!kind) continue;
      if (!out.kinds.includes(kind)) out.kinds.push(kind);
      // Pull grouping off the first bar/column-flavored chart-type
      // element. Combo charts that mix bar with line/area would
      // otherwise need a per-series field; for the common case of a
      // single `<c:barChart>` body this is the value Excel applies.
      if (barGrouping === undefined && (kind === "bar" || kind === "bar3D")) {
        barGrouping = parseBarGrouping(child);
      }
      // Same shape for line/area: surface the first stacked variant
      // we encounter. `"standard"` collapses to undefined for symmetry
      // with the writer's default.
      if (lineGrouping === undefined && (kind === "line" || kind === "line3D")) {
        lineGrouping = parseLineAreaGrouping(child);
      }
      if (areaGrouping === undefined && (kind === "area" || kind === "area3D")) {
        areaGrouping = parseLineAreaGrouping(child);
      }
      // Pull `<c:holeSize>` off a doughnut chart so a parsed template
      // can round-trip its hole back through {@link cloneChart}.
      if (holeSize === undefined && kind === "doughnut") {
        holeSize = parseHoleSize(child);
      }
      // `<c:firstSliceAng>` lives on `<c:pieChart>` and
      // `<c:doughnutChart>` (also pie3D / ofPie which we lump in here
      // for symmetry — the writer never emits those, but a parsed
      // template carrying one round-trips cleanly into a pie/doughnut
      // clone). `0` collapses to undefined because it is the OOXML
      // default that the writer also treats as absence of the field.
      if (
        firstSliceAng === undefined &&
        (kind === "pie" || kind === "pie3D" || kind === "doughnut" || kind === "ofPie")
      ) {
        firstSliceAng = parseFirstSliceAng(child);
      }
      let localIndex = 0;
      for (const ser of childElements(child)) {
        if (ser.local !== "ser") continue;
        seriesCount++;
        series.push(parseSeries(ser, kind, localIndex));
        localIndex++;
      }
      // Chart-type-level <c:dLbls> sits as a sibling of <c:ser> inside
      // the chart-type element. Surface the first one we find — combo
      // charts can carry one per kind, but the common case is a single
      // chart-type element so we keep the model flat.
      if (chartLevelLabels === undefined) {
        const dLbls = findChild(child, "dLbls");
        if (dLbls) {
          const parsed = parseDataLabels(dLbls);
          if (parsed) chartLevelLabels = parsed;
        }
      }
    }
    out.seriesCount = seriesCount;
    if (series.length > 0) out.series = series;
    if (barGrouping !== undefined) out.barGrouping = barGrouping;
    if (lineGrouping !== undefined) out.lineGrouping = lineGrouping;
    if (areaGrouping !== undefined) out.areaGrouping = areaGrouping;
    if (chartLevelLabels) out.dataLabels = chartLevelLabels;
    if (holeSize !== undefined) out.holeSize = holeSize;
    if (firstSliceAng !== undefined) out.firstSliceAng = firstSliceAng;

    const axes = parseAxes(plotArea);
    if (axes !== undefined) out.axes = axes;
  }

  const legend = parseLegend(chartEl);
  if (legend !== undefined) out.legend = legend;

  return out;
}

// ── Axes ──────────────────────────────────────────────────────────

/**
 * Pull per-axis metadata from the plot area's `<c:catAx>` / `<c:valAx>`
 * children.
 *
 * The mapping mirrors the writer side:
 *   - bar / column / line / area: `x` = `<c:catAx>`, `y` = first `<c:valAx>`.
 *   - scatter / bubble:           `x` = first `<c:valAx>`, `y` = second `<c:valAx>`.
 *
 * Returns `undefined` when neither axis surfaces a title — keeps the
 * default `Chart` shape lean.
 */
function parseAxes(plotArea: XmlElement): { x?: ChartAxisInfo; y?: ChartAxisInfo } | undefined {
  let catAx: XmlElement | undefined;
  const valAxes: XmlElement[] = [];
  for (const child of childElements(plotArea)) {
    if (child.local === "catAx") {
      catAx ??= child;
    } else if (child.local === "valAx") {
      valAxes.push(child);
    }
  }

  let xAxis: XmlElement | undefined;
  let yAxis: XmlElement | undefined;
  if (catAx) {
    xAxis = catAx;
    yAxis = valAxes[0];
  } else {
    // Scatter / bubble: both axes are valAx. The first declared one is
    // the X axis (`axPos="b"`), the second is the Y axis (`axPos="l"`).
    xAxis = valAxes[0];
    yAxis = valAxes[1];
  }

  const x = xAxis ? parseAxisInfo(xAxis) : undefined;
  const y = yAxis ? parseAxisInfo(yAxis) : undefined;

  if (!x && !y) return undefined;
  const out: { x?: ChartAxisInfo; y?: ChartAxisInfo } = {};
  if (x) out.x = x;
  if (y) out.y = y;
  return out;
}

function parseAxisInfo(axis: XmlElement): ChartAxisInfo | undefined {
  const title = parseAxisTitle(axis);
  const gridlines = parseAxisGridlines(axis);
  if (title === undefined && gridlines === undefined) return undefined;
  const out: ChartAxisInfo = {};
  if (title !== undefined) out.title = title;
  if (gridlines !== undefined) out.gridlines = gridlines;
  return out;
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
function parseAxisGridlines(axis: XmlElement): ChartAxisGridlines | undefined {
  const major = findChild(axis, "majorGridlines") !== undefined;
  const minor = findChild(axis, "minorGridlines") !== undefined;
  if (!major && !minor) return undefined;
  const out: ChartAxisGridlines = {};
  if (major) out.major = true;
  if (minor) out.minor = true;
  return out;
}

/**
 * Read an axis's `<c:title>` text. Mirrors {@link parseTitle} but
 * scoped to a single axis element rather than the chart root.
 */
function parseAxisTitle(axis: XmlElement): string | undefined {
  const title = findChild(axis, "title");
  if (!title) return undefined;
  const tx = findChild(title, "tx");
  if (!tx) return undefined;
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

// ── Series ────────────────────────────────────────────────────────

/**
 * Pull the metadata fields {@link ChartSeriesInfo} surfaces out of a
 * single `<c:ser>` element. Missing pieces (no name, no categories,
 * literal numbers instead of a range) are simply omitted.
 */
function parseSeries(ser: XmlElement, kind: ChartKind, index: number): ChartSeriesInfo {
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

  return out;
}

// ── Data Labels ───────────────────────────────────────────────────

const VALID_DLBL_POSITIONS: ReadonlySet<ChartDataLabelPosition> = new Set([
  "t",
  "b",
  "l",
  "r",
  "ctr",
  "inEnd",
  "inBase",
  "outEnd",
  "bestFit",
]);

/**
 * Read a `<c:dLbls>` block. Returns `undefined` when the block is
 * empty or only contains a `<c:delete val="1">` (which suppresses
 * labels rather than describing them). All toggles default to `false`
 * when the matching `<c:show*>` element is absent.
 */
function parseDataLabels(el: XmlElement): ChartDataLabelsInfo | undefined {
  // <c:delete val="1"> at the root of <c:dLbls> means "suppress for
  // this scope". We don't surface a dataLabels record for that case —
  // it's the absence of labels, not a configuration.
  const deleteEl = findChild(el, "delete");
  if (deleteEl && readBoolAttr(deleteEl) === true) return undefined;

  const out: ChartDataLabelsInfo = {};

  const pos = findChild(el, "dLblPos");
  if (pos) {
    const val = pos.attrs.val;
    if (typeof val === "string" && VALID_DLBL_POSITIONS.has(val as ChartDataLabelPosition)) {
      out.position = val as ChartDataLabelPosition;
    }
  }

  const showVal = findChild(el, "showVal");
  if (showVal && readBoolAttr(showVal) === true) out.showValue = true;

  const showCat = findChild(el, "showCatName");
  if (showCat && readBoolAttr(showCat) === true) out.showCategoryName = true;

  const showSer = findChild(el, "showSerName");
  if (showSer && readBoolAttr(showSer) === true) out.showSeriesName = true;

  const showPct = findChild(el, "showPercent");
  if (showPct && readBoolAttr(showPct) === true) out.showPercent = true;

  const sep = findChild(el, "separator");
  if (sep) {
    const text = elementText(sep);
    if (text.length > 0) out.separator = text;
  }

  // Empty record is meaningless to a consumer — collapse to undefined.
  if (
    out.position === undefined &&
    !out.showValue &&
    !out.showCategoryName &&
    !out.showSeriesName &&
    !out.showPercent &&
    out.separator === undefined
  ) {
    return undefined;
  }
  return out;
}

/**
 * Read a boolean-style `val` attribute. Excel emits `"1"` / `"0"` but
 * the OOXML spec also blesses `"true"` / `"false"`. Returns `undefined`
 * when the attribute is missing.
 */
function readBoolAttr(el: XmlElement): boolean | undefined {
  const v = el.attrs.val;
  if (typeof v !== "string") return undefined;
  return v === "1" || v.toLowerCase() === "true";
}

/** Read the `<c:tx>` series-name element (literal or strRef cache). */
function parseSeriesName(ser: XmlElement): string | undefined {
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

/**
 * Walk `<c:val>` / `<c:cat>` / `<c:xVal>` / `<c:yVal>` to its inner
 * `<c:f>` formula text. Returns `undefined` for embedded `<c:numLit>`
 * literal data (no formula) or when the element is absent.
 */
function formulaText(wrapper: XmlElement | undefined): string | undefined {
  if (!wrapper) return undefined;
  const numRef = findChild(wrapper, "numRef") ?? findChild(wrapper, "strRef");
  if (numRef) {
    const f = findChild(numRef, "f");
    if (f) {
      const text = elementText(f).trim();
      if (text.length > 0) return text;
    }
  }
  // Some writers put <c:f> directly under <c:strRef> (already handled
  // above via numRef fallback) or under the wrapper itself.
  const direct = findChild(wrapper, "f");
  if (direct) {
    const text = elementText(direct).trim();
    if (text.length > 0) return text;
  }
  return undefined;
}

/** Pull the first `<a:srgbClr val="RRGGBB">` under `<c:spPr>`. */
function parseSeriesColor(ser: XmlElement): string | undefined {
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

// ── Legend ────────────────────────────────────────────────────────

/**
 * Map `<c:legend><c:legendPos val=".."/></c:legend>` to the writer-side
 * {@link ChartLegendPosition}. Returns `false` when `<c:delete val="1"/>`
 * is present (Excel's "no legend" state); returns `undefined` when the
 * chart has no `<c:legend>` element at all.
 */
function parseLegend(chartEl: XmlElement): false | ChartLegendPosition | undefined {
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

// ── Bar Grouping ──────────────────────────────────────────────────

/**
 * Pull `<c:grouping val=".."/>` off a `<c:barChart>` element. Returns
 * `undefined` when the grouping element is missing or carries the
 * default `"standard"` / `"clustered"` value — the writer's
 * {@link SheetChart.barGrouping} treats both as the unspecified
 * default, so omitting them keeps the parsed shape minimal.
 */
function parseBarGrouping(barChart: XmlElement): ChartBarGrouping | undefined {
  const grouping = findChild(barChart, "grouping");
  if (!grouping) return undefined;
  const val = grouping.attrs.val;
  if (typeof val !== "string") return undefined;
  switch (val) {
    case "stacked":
      return "stacked";
    case "percentStacked":
      return "percentStacked";
    case "clustered":
      return "clustered";
    case "standard":
      // OOXML's `standard` for barChart is functionally equivalent to
      // `clustered` (Excel renders side-by-side). Surface neither so
      // the cloned chart inherits the writer's default.
      return undefined;
    default:
      return undefined;
  }
}

// ── Doughnut Hole ─────────────────────────────────────────────────

/**
 * Pull `<c:holeSize val=".."/>` off a `<c:doughnutChart>` element.
 * Returns `undefined` when the attribute is missing, malformed, or
 * outside the 1–99 range OOXML allows. Excel itself only writes values
 * in 10–90 (the UI clamps to that band) but the spec is wider, so we
 * accept the full schema range and let the writer re-clamp on the way
 * back out.
 */
function parseHoleSize(doughnut: XmlElement): number | undefined {
  const el = findChild(doughnut, "holeSize");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 1 || parsed > 99) return undefined;
  return parsed;
}

// ── First Slice Angle ─────────────────────────────────────────────

/**
 * Pull `<c:firstSliceAng val=".."/>` off a `<c:pieChart>` /
 * `<c:doughnutChart>` element. Returns `undefined` when the attribute
 * is missing, malformed, or carries the OOXML default of `0` — the
 * writer's {@link SheetChart.firstSliceAng} treats absence and `0`
 * identically, so collapsing here keeps the round-trip stable.
 *
 * The OOXML schema (CT_FirstSliceAng) restricts the value to the
 * inclusive range `0..360`; out-of-range values are dropped rather
 * than clamped so a corrupt template does not silently rewrite as a
 * different angle.
 */
function parseFirstSliceAng(chartType: XmlElement): number | undefined {
  const el = findChild(chartType, "firstSliceAng");
  if (!el) return undefined;
  const raw = el.attrs.val;
  if (typeof raw !== "string") return undefined;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed)) return undefined;
  if (parsed < 0 || parsed > 360) return undefined;
  // Collapse `0` and the schema-equivalent `360` to undefined — both
  // mean "first slice at 12 o'clock", which is the writer's default.
  if (parsed === 0 || parsed === 360) return undefined;
  return parsed;
}

/**
 * Pull `<c:grouping val=".."/>` off a `<c:lineChart>` or `<c:areaChart>`
 * element. Returns `undefined` when the grouping element is missing or
 * carries the default `"standard"` value — the writer's
 * {@link SheetChart.lineGrouping} / {@link SheetChart.areaGrouping}
 * treat that as the absence of the field.
 */
function parseLineAreaGrouping(chartType: XmlElement): ChartLineAreaGrouping | undefined {
  const grouping = findChild(chartType, "grouping");
  if (!grouping) return undefined;
  const val = grouping.attrs.val;
  if (typeof val !== "string") return undefined;
  switch (val) {
    case "stacked":
      return "stacked";
    case "percentStacked":
      return "percentStacked";
    case "standard":
      return undefined;
    default:
      return undefined;
  }
}

/**
 * Parse an OOXML boolean attribute. The spec allows `"1"` / `"0"` /
 * `"true"` / `"false"`.
 */
function readBoolVal(raw: string | undefined): boolean | undefined {
  if (raw === undefined) return undefined;
  if (raw === "1" || raw === "true") return true;
  if (raw === "0" || raw === "false") return false;
  return undefined;
}

// ── Internals ─────────────────────────────────────────────────────

/**
 * Read `<c:title>` text. The title may be a rich-text run tree or a
 * formula reference; we only surface plain text runs joined together.
 */
function parseTitle(chartEl: XmlElement): string | undefined {
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

function collectTextRuns(el: XmlElement, out: string[]): void {
  for (const child of el.children) {
    if (typeof child === "string") continue;
    if (child.local === "t") {
      out.push(elementText(child));
    } else {
      collectTextRuns(child, out);
    }
  }
}

function elementText(el: XmlElement): string {
  let buf = "";
  for (const child of el.children) {
    if (typeof child === "string") buf += child;
    else buf += elementText(child);
  }
  return buf;
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

function findDescendant(el: XmlElement, localName: string): XmlElement | undefined {
  if (el.local === localName) return el;
  for (const c of el.children) {
    if (typeof c === "string") continue;
    const hit = findDescendant(c, localName);
    if (hit) return hit;
  }
  return undefined;
}

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}
