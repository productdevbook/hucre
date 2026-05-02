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

import type { Chart, ChartAxisInfo, ChartKind, ChartSeriesInfo } from "../_types";
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
    for (const child of childElements(plotArea)) {
      const kind = CHART_KIND_TAGS.get(child.local);
      if (!kind) continue;
      if (!out.kinds.includes(kind)) out.kinds.push(kind);
      let localIndex = 0;
      for (const ser of childElements(child)) {
        if (ser.local !== "ser") continue;
        seriesCount++;
        series.push(parseSeries(ser, kind, localIndex));
        localIndex++;
      }
    }
    out.seriesCount = seriesCount;
    if (series.length > 0) out.series = series;

    const axes = parseAxes(plotArea);
    if (axes !== undefined) out.axes = axes;
  }

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
  if (title === undefined) return undefined;
  return { title };
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

  return out;
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
