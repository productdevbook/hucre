import { describe, expect, it } from "vitest"
import type {
  ChartDataLabels,
  ChartDataPoint,
  ChartErrorBars,
  ChartSeries,
  ChartTrendline,
} from "../src/_types"
import { type XmlElement, parseXml } from "../src/xml/parser"
import {
  buildAllErrorBars,
  buildDataPoints,
  buildErrorBars,
  buildShape3D,
  buildTrendline,
  buildTrendlines,
  cloneAllErrorBars,
  cloneDataPoint,
  cloneDataPoints,
  cloneErrorBars,
  cloneTrendline,
  cloneTrendlines,
  normalizeShape3D,
  parseBubbleSizeRef,
  parseDataPoints,
  parseErrorBars,
  parseShape3D,
  parseTrendlines,
  resolveDataPoints,
  resolveErrorBars,
  resolveTrendlines,
} from "../src/xlsx/chart/seriesExtras"
import {
  buildSeries,
  buildSeriesSpPr,
  mergeSeries,
  parseMarker,
  parseSeries,
  parseSeriesColor,
  parseSeriesName,
  parseSeriesStroke,
} from "../src/xlsx/chart/series"
import {
  buildDataLabelsBody,
  parseDataLabels,
  parseDataLabelsBold,
  parseDataLabelsFontColor,
  parseDataLabelsFontFamily,
  parseDataLabelsFontSize,
  parseDataLabelsItalic,
  parseDataLabelsStrikethrough,
  parseDataLabelsUnderline,
} from "../src/xlsx/chart/dataLabels"
import {
  parseDataTable,
  parseDataTableBold,
  parseDataTableFlag,
  parseDataTableFontColor,
  parseDataTableFontFamily,
  parseDataTableFontSize,
  parseDataTableItalic,
  parseDataTableStrikethrough,
  parseDataTableUnderline,
} from "../src/xlsx/chart/dataTable"

// ── Helpers ──────────────────────────────────────────────────────────

const NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

/** Build a real `<c:ser>` through the project's parser, not a literal. */
function ser(inner: string): XmlElement {
  return parseXml(`<c:ser ${NS}>${inner}</c:ser>`)
}

/** Build an arbitrary chart-namespace element, e.g. `<c:dLbls>`. */
function el(tag: string, inner: string): XmlElement {
  return parseXml(`<c:${tag} ${NS}>${inner}</c:${tag}>`)
}

/**
 * The five-link `<c:txPr><a:p><a:pPr><a:defRPr>` chain every chart
 * typography reader walks. `depth` truncates it so a test can assert
 * the reader bails on the first missing link.
 */
function txPrChain(depth: 0 | 1 | 2 | 3, defRPrAttrs = ""): string {
  if (depth === 0) return "<c:txPr/>"
  if (depth === 1) return "<c:txPr><a:p/></c:txPr>"
  if (depth === 2) return "<c:txPr><a:p><a:pPr/></a:p></c:txPr>"
  return `<c:txPr><a:p><a:pPr><a:defRPr${defRPrAttrs}/></a:pPr></a:p></c:txPr>`
}

// ═══════════════════════════════════════════════════════════════════════
// seriesExtras — <c:shape> (ST_Shape, §21.2.3.34)
// ═══════════════════════════════════════════════════════════════════════

describe("parseShape3D", () => {
  it("reads every ST_Shape token Excel's 3-D bar UI exposes", () => {
    for (const token of ["cone", "coneToMax", "box", "cylinder", "pyramid", "pyramidToMax"]) {
      expect(parseShape3D(ser(`<c:shape val="${token}"/>`)), token).toBe(token)
    }
  })

  it("tolerates surrounding whitespace on the val token", () => {
    expect(parseShape3D(ser('<c:shape val="  cylinder  "/>'))).toBe("cylinder")
  })

  it("drops absence, a missing val, and tokens outside ST_Shape", () => {
    // Excel falls back to "box" for all three, so surfacing nothing keeps
    // absence and the default indistinguishable on re-emit.
    expect(parseShape3D(ser(""))).toBeUndefined()
    expect(parseShape3D(ser("<c:shape/>"))).toBeUndefined()
    expect(parseShape3D(ser('<c:shape val="sphere"/>'))).toBeUndefined()
    expect(parseShape3D(ser('<c:shape val=""/>'))).toBeUndefined()
  })
})

describe("normalizeShape3D / buildShape3D", () => {
  it("emits only recognized tokens so the writer can elide the element", () => {
    expect(buildShape3D("cone")).toBe('<c:shape val="cone"/>')
    expect(buildShape3D(undefined)).toBeUndefined()
    expect(buildShape3D("sphere" as never)).toBeUndefined()
    expect(normalizeShape3D(42 as never)).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// seriesExtras — <c:bubbleSize> (CT_NumDataSource, §21.2.2.30)
// ═══════════════════════════════════════════════════════════════════════

describe("parseBubbleSizeRef", () => {
  it("walks <c:bubbleSize><c:numRef><c:f> to the formula text", () => {
    expect(
      parseBubbleSizeRef(
        ser("<c:bubbleSize><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:bubbleSize>"),
      ),
    ).toBe("Sheet1!$C$2:$C$5")
  })

  it("returns undefined for an embedded <c:numLit> with no formula", () => {
    expect(
      parseBubbleSizeRef(
        ser('<c:bubbleSize><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:bubbleSize>'),
      ),
    ).toBeUndefined()
    expect(parseBubbleSizeRef(ser(""))).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// seriesExtras — <c:dPt> (CT_DPt, §21.2.2.52)
// ═══════════════════════════════════════════════════════════════════════

describe("parseDataPoints", () => {
  it("requires a well-formed <c:idx> selector on every point", () => {
    // `<c:idx>` is the required CT_DPt selector — without a usable index
    // the override cannot land on a slice, so the whole `<c:dPt>` drops
    // rather than surface a fabricated index 0.
    expect(parseDataPoints(ser('<c:dPt><c:bubble3D val="1"/></c:dPt>'))).toBeUndefined()
    expect(parseDataPoints(ser("<c:dPt><c:idx/></c:dPt>"))).toBeUndefined()
    expect(parseDataPoints(ser('<c:dPt><c:idx val="-1"/></c:dPt>'))).toBeUndefined()
    expect(parseDataPoints(ser('<c:dPt><c:idx val="abc"/></c:dPt>'))).toBeUndefined()
  })

  it("returns undefined when the series declares no <c:dPt> at all", () => {
    expect(parseDataPoints(ser('<c:idx val="0"/>'))).toBeUndefined()
  })

  it("rounds <c:explosion> and drops the OOXML default 0", () => {
    const points = parseDataPoints(
      ser(
        '<c:dPt><c:idx val="0"/><c:explosion val="24.6"/></c:dPt>' +
          '<c:dPt><c:idx val="1"/><c:explosion val="0"/></c:dPt>' +
          '<c:dPt><c:idx val="2"/><c:explosion val="-5"/></c:dPt>' +
          '<c:dPt><c:idx val="3"/><c:explosion val="0.4"/></c:dPt>' +
          '<c:dPt><c:idx val="4"/><c:explosion/></c:dPt>',
      ),
    )
    expect(points?.map((p) => p.explosion)).toEqual([
      25,
      undefined,
      undefined,
      undefined,
      undefined,
    ])
  })

  it("surfaces <c:bubble3D> only for the explicit truthy spelling", () => {
    const points = parseDataPoints(
      ser(
        '<c:dPt><c:idx val="0"/><c:bubble3D val="1"/></c:dPt>' +
          '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/></c:dPt>',
      ),
    )
    expect(points?.map((p) => p.bubble3D)).toEqual([true, undefined])
  })

  it("bails at each broken link of the <c:spPr><a:solidFill><a:srgbClr> fill chain", () => {
    const cases = [
      "",
      "<c:spPr/>",
      "<c:spPr><a:noFill/></c:spPr>",
      "<c:spPr><a:solidFill/></c:spPr>",
      '<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>',
    ]
    for (const body of cases) {
      const points = parseDataPoints(ser(`<c:dPt><c:idx val="0"/>${body}</c:dPt>`))
      expect(points?.[0].fillColor, body).toBeUndefined()
    }
    const ok = parseDataPoints(
      ser(
        '<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="ff0000"/>' +
          "</a:solidFill></c:spPr></c:dPt>",
      ),
    )
    expect(ok?.[0].fillColor).toBe("FF0000")
  })

  it("bails at each broken link of the <a:ln><a:solidFill><a:srgbClr> border chain", () => {
    const cases = [
      "<c:spPr/>",
      "<c:spPr><a:ln/></c:spPr>",
      "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>",
      "<c:spPr><a:ln><a:solidFill/></a:ln></c:spPr>",
    ]
    for (const body of cases) {
      const points = parseDataPoints(ser(`<c:dPt><c:idx val="0"/>${body}</c:dPt>`))
      expect(points?.[0].borderColor, body).toBeUndefined()
    }
  })

  it("reads the per-point border colour off <c:spPr><a:ln><a:solidFill>", () => {
    const points = parseDataPoints(
      ser(
        '<c:dPt><c:idx val="4"/><c:spPr><a:ln><a:solidFill>' +
          '<a:srgbClr val="#00FF7F"/></a:solidFill><a:prstDash val="dashDot"/></a:ln>' +
          "</c:spPr></c:dPt>",
      ),
    )
    expect(points).toEqual([{ idx: 4, borderColor: "00FF7F", borderDash: "dashDot" }])
  })

  it("converts <a:ln w=..> from EMU to points and clamps to Excel's 0.25..13.5pt band", () => {
    const read = (w: string): number | undefined =>
      parseDataPoints(ser(`<c:dPt><c:idx val="0"/><c:spPr><a:ln w="${w}"/></c:spPr></c:dPt>`))?.[0]
        .borderWidth
    // 1 pt = 12 700 EMU per CT_LineProperties (§20.1.2.3.24).
    expect(read("25400")).toBe(2)
    // Below the UI minimum clamps up; above the maximum clamps down.
    expect(read("1000")).toBe(0.25)
    expect(read("1000000")).toBe(13.5)
    // Zero is Excel's "no border" marker, which this field does not model.
    expect(read("0")).toBeUndefined()
    expect(read("-25400")).toBeUndefined()
    expect(read("wide")).toBeUndefined()
    expect(
      parseDataPoints(ser('<c:dPt><c:idx val="0"/><c:spPr><a:ln/></c:spPr></c:dPt>'))?.[0]
        .borderWidth,
    ).toBeUndefined()
  })

  it("drops the OOXML default dash `solid` and every unrecognized token", () => {
    const read = (body: string): string | undefined =>
      parseDataPoints(
        ser(`<c:dPt><c:idx val="0"/><c:spPr><a:ln>${body}</a:ln></c:spPr></c:dPt>`),
      )?.[0].borderDash
    expect(read('<a:prstDash val="dash"/>')).toBe("dash")
    expect(read('<a:prstDash val="solid"/>')).toBeUndefined()
    expect(read('<a:prstDash val="zigzag"/>')).toBeUndefined()
    expect(read("<a:prstDash/>")).toBeUndefined()
    expect(read("")).toBeUndefined()
  })

  it("lifts a per-point <c:marker> through the shared series-marker reader", () => {
    const points = parseDataPoints(
      ser('<c:dPt><c:idx val="2"/><c:marker><c:symbol val="diamond"/></c:marker></c:dPt>'),
    )
    expect(points?.[0]).toEqual({ idx: 2, marker: { symbol: "diamond" } })
  })
})

describe("buildDataPoints", () => {
  it("emits <c:explosion> only on the pie family", () => {
    const dp: ChartDataPoint[] = [{ idx: 0, explosion: 25 }]
    expect(buildDataPoints(dp, "pie")[0]).toContain('<c:explosion val="25"/>')
    expect(buildDataPoints(dp, "doughnut")[0]).toContain('<c:explosion val="25"/>')
    expect(buildDataPoints(dp, "pie3D")[0]).toContain('<c:explosion val="25"/>')
    // Bar / column / line have no `<c:explosion>` slot on CT_DPt's
    // per-family sequence, so a pinned value is dropped rather than
    // emitted where Excel's validator would reject it.
    expect(buildDataPoints(dp, "bar")[0]).not.toContain("c:explosion")
  })

  it("clamps explosion to the 0..400 band and drops the default 0", () => {
    expect(buildDataPoints([{ idx: 0, explosion: 900 }], "pie")[0]).toContain(
      '<c:explosion val="400"/>',
    )
    expect(buildDataPoints([{ idx: 0, explosion: 0 }], "pie")[0]).not.toContain("c:explosion")
    expect(buildDataPoints([{ idx: 0, explosion: Number.NaN }], "pie")[0]).not.toContain(
      "c:explosion",
    )
  })

  it("always emits <c:bubble3D> because CT_DPt lists it as required", () => {
    expect(buildDataPoints([{ idx: 0 }], "bar")[0]).toContain('<c:bubble3D val="0"/>')
    expect(buildDataPoints([{ idx: 0, bubble3D: true }], "bar")[0]).toContain(
      '<c:bubble3D val="1"/>',
    )
  })

  it("self-closes <a:ln> when only the width lands on the wire", () => {
    // No `<a:solidFill>` / `<a:prstDash>` children means the element has
    // nothing to wrap, so the writer emits the attribute-only form.
    const xml = buildDataPoints([{ idx: 0, borderWidth: 2 }], "bar")[0]
    expect(xml).toContain('<a:ln w="25400"/>')
  })

  it("wraps <a:ln> with no width attribute at all when only a colour is pinned", () => {
    const xml = buildDataPoints([{ idx: 0, borderColor: "112233", borderDash: "dot" }], "bar")[0]
    expect(xml).toContain('<a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill>')
    expect(xml).toContain('<a:prstDash val="dot"/>')
    expect(xml).not.toContain("<a:ln w=")
  })

  it("skips points whose idx cannot land on a real data point", () => {
    expect(buildDataPoints([{ idx: -1 }, { idx: Number.NaN }], "bar")).toEqual([])
    expect(buildDataPoints(undefined, "bar")).toEqual([])
    expect(buildDataPoints([], "bar")).toEqual([])
    // A fractional index floors rather than drops — Excel indexes points
    // with `xsd:unsignedInt`.
    expect(buildDataPoints([{ idx: 2.9 }], "bar")[0]).toContain('<c:idx val="2"/>')
  })
})

describe("cloneDataPoint / resolveDataPoints", () => {
  it("keeps only the fields the writer can re-emit", () => {
    expect(
      cloneDataPoint({
        idx: 1,
        explosion: Number.NaN,
        bubble3D: false,
        borderWidth: Number.POSITIVE_INFINITY,
        marker: {},
      } as ChartDataPoint),
    ).toEqual({ idx: 1 })
  })

  it("follows the inherit / drop / replace override grammar", () => {
    const source: ChartDataPoint[] = [{ idx: 0, fillColor: "FF0000" }]
    expect(resolveDataPoints(source, undefined)).toEqual(source)
    expect(resolveDataPoints(source, null)).toBeUndefined()
    expect(resolveDataPoints(source, [{ idx: 3 }])).toEqual([{ idx: 3 }])
    expect(resolveDataPoints(source, [])).toBeUndefined()
    expect(cloneDataPoints(undefined)).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// seriesExtras — <c:trendline> (CT_Trendline, §21.2.2.211)
// ═══════════════════════════════════════════════════════════════════════

describe("parseTrendlines", () => {
  const one = (inner: string): ChartTrendline | undefined =>
    parseTrendlines(ser(`<c:trendline>${inner}</c:trendline>`))?.[0]

  it("requires a recognized ST_TrendlineType before admitting the block", () => {
    expect(parseTrendlines(ser("<c:trendline/>"))).toBeUndefined()
    expect(parseTrendlines(ser("<c:trendline><c:trendlineType/></c:trendline>"))).toBeUndefined()
    expect(
      parseTrendlines(ser('<c:trendline><c:trendlineType val="spline"/></c:trendline>')),
    ).toBeUndefined()
    for (const t of ["linear", "log", "exp", "power", "poly", "movingAvg"]) {
      expect(one(`<c:trendlineType val="${t}"/>`)?.type, t).toBe(t)
    }
  })

  it("reads the display name off <c:name>'s text content", () => {
    expect(one('<c:trendlineType val="linear"/><c:name>  Trend A  </c:name>')?.name).toBe("Trend A")
    // `<c:name>` is plain text — an element child contributes nothing.
    expect(one('<c:trendlineType val="linear"/><c:name>Trend<c:extLst/></c:name>')?.name).toBe(
      "Trend",
    )
    // A whitespace-only name is indistinguishable from absence to Excel.
    expect(one('<c:trendlineType val="linear"/><c:name>   </c:name>')?.name).toBeUndefined()
    expect(one('<c:trendlineType val="linear"/><c:name/>')?.name).toBeUndefined()
  })

  it("admits <c:order> only inside the 2..6 polynomial band Excel's UI exposes", () => {
    const order = (v: string): number | undefined =>
      one(`<c:trendlineType val="poly"/><c:order val="${v}"/>`)?.order
    expect(order("2")).toBe(2)
    expect(order("6")).toBe(6)
    expect(order("3.4")).toBe(3)
    expect(order("1")).toBeUndefined()
    expect(order("7")).toBeUndefined()
    expect(order("")).toBeUndefined()
    expect(one('<c:trendlineType val="poly"/><c:order/>')?.order).toBeUndefined()
  })

  it("admits <c:period> only inside the 2..100 moving-average band", () => {
    const period = (v: string): number | undefined =>
      one(`<c:trendlineType val="movingAvg"/><c:period val="${v}"/>`)?.period
    expect(period("2")).toBe(2)
    expect(period("100")).toBe(100)
    expect(period("1")).toBeUndefined()
    expect(period("101")).toBeUndefined()
    expect(one('<c:trendlineType val="movingAvg"/><c:period/>')?.period).toBeUndefined()
  })

  it("accepts any finite forecast / intercept, including negatives", () => {
    const t = one(
      '<c:trendlineType val="linear"/><c:forward val="2.5"/><c:backward val="-1"/>' +
        '<c:intercept val="0"/>',
    )
    expect(t).toMatchObject({ forward: 2.5, backward: -1, intercept: 0 })
    const empty = one('<c:trendlineType val="linear"/><c:forward/><c:backward/><c:intercept/>')
    expect(empty).toEqual({ type: "linear" })
  })

  it("surfaces <c:dispEq> / <c:dispRSqr> only on the explicit truthy spelling", () => {
    expect(
      one('<c:trendlineType val="linear"/><c:dispEq val="1"/><c:dispRSqr val="true"/>'),
    ).toMatchObject({ dispEquation: true, dispRSquared: true })
    expect(one('<c:trendlineType val="linear"/><c:dispEq val="0"/><c:dispRSqr val="0"/>')).toEqual({
      type: "linear",
    })
  })

  it("bails at each broken link of the <c:spPr><a:ln> stroke chain", () => {
    for (const body of ["<c:spPr/>", "<c:spPr><a:ln/></c:spPr>", ""]) {
      expect(one(`<c:trendlineType val="linear"/>${body}`), body).toEqual({ type: "linear" })
    }
    expect(
      one('<c:trendlineType val="linear"/><c:spPr><a:ln><a:solidFill/></a:ln></c:spPr>'),
    ).toEqual({ type: "linear" })
  })

  it("clamps the stroke width into the 0.25..13.5pt band on both ends", () => {
    const width = (w: string): number | undefined =>
      one(`<c:trendlineType val="linear"/><c:spPr><a:ln w="${w}"/></c:spPr>`)?.lineWidth
    expect(width("25400")).toBe(2)
    expect(width("100")).toBe(0.25)
    expect(width("9999999")).toBe(13.5)
    expect(width("0")).toBeUndefined()
    expect(width("thick")).toBeUndefined()
  })

  it("reads the line color and dash, dropping the OOXML default `solid`", () => {
    const t = one(
      '<c:trendlineType val="linear"/><c:spPr><a:ln>' +
        '<a:solidFill><a:srgbClr val="#00ff00"/></a:solidFill>' +
        '<a:prstDash val="sysDot"/></a:ln></c:spPr>',
    )
    expect(t).toMatchObject({ lineColor: "00FF00", lineDash: "sysDot" })
    expect(
      one(
        '<c:trendlineType val="linear"/><c:spPr><a:ln>' +
          '<a:solidFill><a:srgbClr val="nothex"/></a:solidFill>' +
          '<a:prstDash val="solid"/></a:ln></c:spPr>',
      ),
    ).toEqual({ type: "linear" })
    expect(
      one('<c:trendlineType val="linear"/><c:spPr><a:ln><a:prstDash/></a:ln></c:spPr>'),
    ).toEqual({ type: "linear" })
  })
})

describe("buildTrendline", () => {
  it("scopes <c:order> to poly and <c:period> to movingAvg", () => {
    // Excel rejects `<c:order>` on a linear trendline — the writer drops
    // the field rather than emit an element outside its schema slot.
    expect(buildTrendline({ type: "linear", order: 3, period: 5 })).not.toContain("c:order")
    expect(buildTrendline({ type: "poly", order: 3 })).toContain('<c:order val="3"/>')
    expect(buildTrendline({ type: "movingAvg", period: 5 })).toContain('<c:period val="5"/>')
  })

  it("clamps rather than drops an out-of-band order / period on the write side", () => {
    // A reader drops out-of-band values; the writer clamps so an
    // authored-but-invalid value still produces a file Excel opens.
    expect(buildTrendline({ type: "poly", order: 99 })).toContain('<c:order val="6"/>')
    expect(buildTrendline({ type: "poly", order: 0 })).toContain('<c:order val="2"/>')
    expect(buildTrendline({ type: "movingAvg", period: 9999 })).toContain('<c:period val="100"/>')
    expect(buildTrendline({ type: "movingAvg", period: 1 })).toContain('<c:period val="2"/>')
  })

  it("escapes the display name so an ampersand cannot break the document", () => {
    expect(buildTrendline({ type: "linear", name: "R&D <2024>" })).toContain(
      "<c:name>R&amp;D &lt;2024&gt;</c:name>",
    )
    expect(buildTrendline({ type: "linear", name: "   " })).not.toContain("c:name")
  })

  it("self-closes <a:ln> when only the width is pinned", () => {
    expect(buildTrendline({ type: "linear", lineWidth: 1 })).toContain('<a:ln w="12700"/>')
    // With a color child the element must wrap, keeping `w` as an attribute.
    expect(buildTrendline({ type: "linear", lineWidth: 1, lineColor: "FF0000" })).toContain(
      '<a:ln w="12700">',
    )
    // A colour-only stroke wraps with no `w` attribute at all.
    expect(buildTrendline({ type: "linear", lineColor: "FF0000", lineDash: "dash" })).toContain(
      "<a:ln><a:solidFill>",
    )
    // No stroke knob at all elides the whole `<c:spPr>`.
    expect(buildTrendline({ type: "linear" })).not.toContain("c:spPr")
  })

  it("refuses to emit a block whose type is outside ST_TrendlineType", () => {
    expect(buildTrendline({ type: "spline" as never })).toBeUndefined()
    expect(buildTrendlines([{ type: "spline" as never }, { type: "linear" }])).toHaveLength(1)
    expect(buildTrendlines(undefined)).toEqual([])
    expect(buildTrendlines([])).toEqual([])
  })
})

describe("cloneTrendline / resolveTrendlines", () => {
  it("drops a trendline whose type no longer validates", () => {
    expect(cloneTrendline({ type: "spline" as never })).toBeUndefined()
    expect(cloneTrendlines([{ type: "spline" as never }])).toBeUndefined()
    expect(cloneTrendlines(undefined)).toBeUndefined()
    expect(cloneTrendlines([])).toBeUndefined()
  })

  it("keeps only finite numerics and non-empty strings", () => {
    expect(
      cloneTrendline({
        type: "poly",
        name: "",
        order: Number.NaN,
        period: Number.POSITIVE_INFINITY,
        forward: Number.NaN,
        backward: Number.NaN,
        intercept: Number.NaN,
        lineWidth: Number.NaN,
      }),
    ).toEqual({ type: "poly" })
  })

  it("follows the inherit / drop / replace override grammar", () => {
    const source: ChartTrendline[] = [{ type: "linear" }]
    expect(resolveTrendlines(source, undefined)).toEqual(source)
    expect(resolveTrendlines(source, null)).toBeUndefined()
    expect(resolveTrendlines(source, [{ type: "exp" }])).toEqual([{ type: "exp" }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// seriesExtras — <c:errBars> (CT_ErrBars, §21.2.2.55)
// ═══════════════════════════════════════════════════════════════════════

describe("parseErrorBars", () => {
  const HEAD = '<c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/>'
  const one = (inner: string): ChartErrorBars | undefined =>
    parseErrorBars(ser(`<c:errBars>${inner}</c:errBars>`))?.[0]

  it("requires all three ST_Err* selectors before admitting the block", () => {
    // CT_ErrBars makes errDir / errBarType / errValType the identity of
    // the block; a partial header cannot be rendered, so it drops.
    const partials = [
      "",
      '<c:errDir val="y"/>',
      "<c:errDir/>",
      '<c:errDir val="z"/>',
      '<c:errDir val="y"/><c:errBarType/>',
      '<c:errDir val="y"/><c:errBarType val="sideways"/>',
      '<c:errDir val="y"/><c:errBarType val="both"/>',
      '<c:errDir val="y"/><c:errBarType val="both"/><c:errValType/>',
      '<c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="guess"/>',
    ]
    for (const p of partials) {
      expect(parseErrorBars(ser(`<c:errBars>${p}</c:errBars>`)), p).toBeUndefined()
    }
    expect(one(HEAD)).toEqual({ direction: "y", type: "both", valType: "fixedVal" })
  })

  it("reads <c:val> only when it parses to a finite number", () => {
    expect(one(`${HEAD}<c:val val="1.5"/>`)?.value).toBe(1.5)
    expect(one(`${HEAD}<c:val/>`)?.value).toBeUndefined()
    expect(one(`${HEAD}<c:val val="x"/>`)?.value).toBeUndefined()
  })

  it("surfaces <c:noEndCap> only on the explicit truthy spelling", () => {
    expect(one(`${HEAD}<c:noEndCap val="1"/>`)?.noEndCap).toBe(true)
    expect(one(`${HEAD}<c:noEndCap val="0"/>`)?.noEndCap).toBeUndefined()
  })

  it("bails at each broken link of the <c:spPr><a:ln> stroke chain", () => {
    for (const body of ["<c:spPr/>", "<c:spPr><a:ln/></c:spPr>"]) {
      expect(one(`${HEAD}${body}`), body).toEqual({
        direction: "y",
        type: "both",
        valType: "fixedVal",
      })
    }
    expect(one(`${HEAD}<c:spPr><a:ln><a:solidFill/></a:ln></c:spPr>`)?.lineColor).toBeUndefined()
    expect(one(`${HEAD}<c:spPr><a:ln><a:prstDash/></a:ln></c:spPr>`)?.lineDash).toBeUndefined()
  })

  it("clamps the stroke width into the 0.25..13.5pt band on both ends", () => {
    const width = (w: string): number | undefined =>
      one(`${HEAD}<c:spPr><a:ln w="${w}"/></c:spPr>`)?.lineWidth
    expect(width("12700")).toBe(1)
    expect(width("100")).toBe(0.25)
    expect(width("9999999")).toBe(13.5)
    expect(width("0")).toBeUndefined()
    expect(width("thin")).toBeUndefined()
  })

  it("drops the OOXML default dash `solid` and unrecognized colors", () => {
    expect(
      one(
        `${HEAD}<c:spPr><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill>` +
          '<a:prstDash val="dot"/></a:ln></c:spPr>',
      ),
    ).toMatchObject({ lineColor: "123456", lineDash: "dot" })
    expect(
      one(
        `${HEAD}<c:spPr><a:ln><a:solidFill><a:srgbClr val="zz"/></a:solidFill>` +
          '<a:prstDash val="solid"/></a:ln></c:spPr>',
      ),
    ).toEqual({ direction: "y", type: "both", valType: "fixedVal" })
  })
})

describe("buildErrorBars", () => {
  const base: ChartErrorBars = { direction: "y", type: "both", valType: "fixedVal" }

  it("refuses a block whose direction / type / valType is outside its enum", () => {
    expect(buildErrorBars({ ...base, direction: "z" as never })).toBeUndefined()
    expect(buildErrorBars({ ...base, type: "sideways" as never })).toBeUndefined()
    expect(buildErrorBars({ ...base, valType: "guess" as never })).toBeUndefined()
  })

  it("omits <c:val> for the value types that derive their own magnitude", () => {
    // `stdErr` computes from the data and `cust` reads `<c:plus>` /
    // `<c:minus>` ranges — a `<c:val>` alongside either is meaningless.
    expect(buildErrorBars({ ...base, valType: "stdErr", value: 3 })).not.toContain("c:val ")
    expect(buildErrorBars({ ...base, valType: "cust", value: 3 })).not.toContain("c:val ")
    expect(buildErrorBars({ ...base, value: 3 })).toContain('<c:val val="3"/>')
    expect(buildErrorBars({ ...base, value: Number.NaN })).not.toContain("c:val ")
  })

  it("emits <c:noEndCap> only when explicitly pinned", () => {
    expect(buildErrorBars({ ...base, noEndCap: true })).toContain('<c:noEndCap val="1"/>')
    expect(buildErrorBars({ ...base, noEndCap: false })).not.toContain("c:noEndCap")
  })

  it("self-closes <a:ln> when only the width is pinned", () => {
    expect(buildErrorBars({ ...base, lineWidth: 1 })).toContain('<a:ln w="12700"/>')
    expect(buildErrorBars({ ...base, lineDash: "dash" })).toContain("<a:ln>")
    expect(buildErrorBars(base)).not.toContain("c:spPr")
  })

  it("skips invalid entries when building the whole list", () => {
    expect(buildAllErrorBars([{ ...base, direction: "z" as never }, base])).toHaveLength(1)
    expect(buildAllErrorBars(undefined)).toEqual([])
    expect(buildAllErrorBars([])).toEqual([])
  })
})

describe("cloneErrorBars / resolveErrorBars", () => {
  const base: ChartErrorBars = { direction: "x", type: "plus", valType: "percentage" }

  it("drops a record whose enum fields no longer validate", () => {
    expect(cloneErrorBars({ ...base, direction: "z" as never })).toBeUndefined()
    expect(cloneErrorBars({ ...base, type: "sideways" as never })).toBeUndefined()
    expect(cloneErrorBars({ ...base, valType: "guess" as never })).toBeUndefined()
    expect(cloneAllErrorBars([{ ...base, direction: "z" as never }])).toBeUndefined()
    expect(cloneAllErrorBars(undefined)).toBeUndefined()
    expect(cloneAllErrorBars([])).toBeUndefined()
  })

  it("keeps only finite numerics and literal booleans", () => {
    expect(
      cloneErrorBars({
        ...base,
        value: Number.NaN,
        lineWidth: Number.POSITIVE_INFINITY,
        noEndCap: false,
      }),
    ).toEqual(base)
    expect(
      cloneErrorBars({
        ...base,
        value: 5,
        lineWidth: 2,
        noEndCap: true,
        lineColor: "112233",
        lineDash: "dash",
      }),
    ).toEqual({
      ...base,
      value: 5,
      lineWidth: 2,
      noEndCap: true,
      lineColor: "112233",
      lineDash: "dash",
    })
  })

  it("follows the inherit / drop / replace override grammar", () => {
    expect(resolveErrorBars([base], undefined)).toEqual([base])
    expect(resolveErrorBars([base], null)).toBeUndefined()
    expect(resolveErrorBars([base], [{ ...base, direction: "y" }])).toEqual([
      { ...base, direction: "y" },
    ])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// series — per-family gating on <c:ser>
// ═══════════════════════════════════════════════════════════════════════

describe("parseSeries family gating", () => {
  const EXTRAS =
    '<c:trendline><c:trendlineType val="linear"/></c:trendline>' +
    '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/>' +
    '<c:errValType val="fixedVal"/></c:errBars>' +
    "<c:bubbleSize><c:numRef><c:f>Sheet1!$D$2:$D$4</c:f></c:numRef></c:bubbleSize>" +
    '<c:shape val="cylinder"/>'

  it("reads <c:bubbleSize> only on a bubble series", () => {
    expect(parseSeries(ser(EXTRAS), "bubble", 0).bubbleSizeRef).toBe("Sheet1!$D$2:$D$4")
    // A stray `<c:bubbleSize>` on a bar template has no CT_BarSer slot.
    expect(parseSeries(ser(EXTRAS), "bar", 0).bubbleSizeRef).toBeUndefined()
  })

  it("reads <c:shape> only on a bar3D series", () => {
    expect(parseSeries(ser(EXTRAS), "bar3D", 0).shape3D).toBe("cylinder")
    expect(parseSeries(ser(EXTRAS), "bar", 0).shape3D).toBeUndefined()
  })

  it("leaves bubbleSizeRef and shape3D unset when the series declares neither", () => {
    expect(parseSeries(ser('<c:idx val="0"/>'), "bubble", 0).bubbleSizeRef).toBeUndefined()
    expect(parseSeries(ser('<c:idx val="0"/>'), "bar3D", 0).shape3D).toBeUndefined()
  })

  it("drops a <c:dLbls> block that carries nothing the writer models", () => {
    // `<c:dLbls>` with only OOXML defaults collapses in the reader, so
    // the parsed series must not carry an empty `dataLabels` record.
    expect(
      parseSeries(ser('<c:dLbls><c:showVal val="0"/></c:dLbls>'), "bar", 0).dataLabels,
    ).toBeUndefined()
  })

  it("drops trendlines and error bars on pie, which has no slot for either", () => {
    const pie = parseSeries(ser(EXTRAS), "pie", 0)
    expect(pie.trendlines).toBeUndefined()
    expect(pie.errorBars).toBeUndefined()
    const bar = parseSeries(ser(EXTRAS), "bar", 0)
    expect(bar.trendlines).toHaveLength(1)
    expect(bar.errorBars).toHaveLength(1)
  })
})

describe("parseSeriesColor", () => {
  it("lifts an <a:schemeClr> theme reference when no literal sRGB is present", () => {
    expect(
      parseSeriesColor(
        ser(
          '<c:spPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/>' +
            "</a:schemeClr></a:solidFill></c:spPr>",
        ),
      ),
    ).toEqual({ theme: "accent2", lumMod: 75000 })
  })

  it("drops an <a:srgbClr> with no val attribute", () => {
    expect(
      parseSeriesColor(ser("<c:spPr><a:solidFill><a:srgbClr/></a:solidFill></c:spPr>")),
    ).toBeUndefined()
  })

  it("drops a fill the writer cannot reproduce", () => {
    expect(
      parseSeriesColor(
        ser('<c:spPr><a:solidFill><a:sysClr val="windowText"/></a:solidFill></c:spPr>'),
      ),
    ).toBeUndefined()
    expect(parseSeriesColor(ser("<c:spPr><a:noFill/></c:spPr>"))).toBeUndefined()
    expect(parseSeriesColor(ser(""))).toBeUndefined()
  })
})

describe("parseSeriesStroke", () => {
  it("clamps the stroke width up to the 0.25pt UI minimum", () => {
    expect(parseSeriesStroke(ser('<c:spPr><a:ln w="500"/></c:spPr>'))).toEqual({ width: 0.25 })
    expect(parseSeriesStroke(ser('<c:spPr><a:ln w="9999999"/></c:spPr>'))).toEqual({ width: 13.5 })
  })

  it("collapses the OOXML defaults cap=flat and cmpd=sng to absence", () => {
    expect(
      parseSeriesStroke(ser('<c:spPr><a:ln w="12700" cap="flat" cmpd="sng"/></c:spPr>')),
    ).toEqual({ width: 1 })
    expect(parseSeriesStroke(ser('<c:spPr><a:ln cap="rnd" cmpd="dbl"/></c:spPr>'))).toEqual({
      cap: "rnd",
      compound: "dbl",
    })
  })

  it("returns undefined when the <a:ln> carries nothing the writer models", () => {
    expect(parseSeriesStroke(ser("<c:spPr><a:ln/></c:spPr>"))).toBeUndefined()
    // Zero is Excel's "no line" marker and is not modelled as a width.
    expect(parseSeriesStroke(ser('<c:spPr><a:ln w="0"/></c:spPr>'))).toBeUndefined()
    expect(parseSeriesStroke(ser('<c:spPr><a:ln w="thick"/></c:spPr>'))).toBeUndefined()
    expect(
      parseSeriesStroke(ser('<c:spPr><a:ln><a:prstDash val="zigzag"/></a:ln></c:spPr>')),
    ).toBeUndefined()
    expect(
      parseSeriesStroke(ser('<c:spPr><a:ln cap="bevel" cmpd="quad"/></c:spPr>')),
    ).toBeUndefined()
    expect(parseSeriesStroke(ser("<c:spPr/>"))).toBeUndefined()
    expect(parseSeriesStroke(ser(""))).toBeUndefined()
  })
})

describe("parseMarker", () => {
  it("clamps <c:size> into the ST_MarkerSize 2..72 band", () => {
    expect(parseMarker(ser('<c:marker><c:size val="1"/></c:marker>'))).toEqual({ size: 2 })
    expect(parseMarker(ser('<c:marker><c:size val="99"/></c:marker>'))).toEqual({ size: 72 })
    expect(parseMarker(ser('<c:marker><c:size val="7"/></c:marker>'))).toEqual({ size: 7 })
    expect(parseMarker(ser('<c:marker><c:size val="big"/></c:marker>'))).toBeUndefined()
    expect(parseMarker(ser("<c:marker><c:size/></c:marker>"))).toBeUndefined()
  })

  it("drops a symbol outside ST_MarkerStyle", () => {
    expect(parseMarker(ser('<c:marker><c:symbol val="blob"/></c:marker>'))).toBeUndefined()
    expect(parseMarker(ser("<c:marker><c:symbol/></c:marker>"))).toBeUndefined()
  })

  it("collapses an empty <c:marker/> to absence", () => {
    expect(parseMarker(ser("<c:marker/>"))).toBeUndefined()
    expect(parseMarker(ser(""))).toBeUndefined()
  })

  it("drops a marker fill / outline the writer cannot reproduce", () => {
    const read = (spPr: string): unknown => parseMarker(ser(`<c:marker>${spPr}</c:marker>`))
    // `<c:spPr>` with no fill and an `<a:ln>` with no fill of its own.
    expect(read("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>")).toBeUndefined()
    // An `<a:srgbClr>` with no val, and one whose val is not 6 hex digits.
    expect(read("<c:spPr><a:solidFill><a:srgbClr/></a:solidFill></c:spPr>")).toBeUndefined()
    expect(
      read('<c:spPr><a:solidFill><a:srgbClr val="12345"/></a:solidFill></c:spPr>'),
    ).toBeUndefined()
    expect(
      read('<c:spPr><a:ln><a:solidFill><a:srgbClr val="nothex"/></a:solidFill></a:ln></c:spPr>'),
    ).toBeUndefined()
    expect(
      read("<c:spPr><a:ln><a:solidFill><a:srgbClr/></a:solidFill></a:ln></c:spPr>"),
    ).toBeUndefined()
  })

  it("reads the fill and outline off <c:spPr> and <c:spPr><a:ln>", () => {
    expect(
      parseMarker(
        ser(
          '<c:marker><c:spPr><a:solidFill><a:srgbClr val="#abcdef"/></a:solidFill>' +
            '<a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>' +
            "</c:spPr></c:marker>",
        ),
      ),
    ).toEqual({ fill: "ABCDEF", line: "112233" })
    // A non-sRGB marker fill has no lossless writer form.
    expect(
      parseMarker(
        ser(
          '<c:marker><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr></c:marker>',
        ),
      ),
    ).toBeUndefined()
  })
})

describe("parseSeriesName", () => {
  it("falls back to the formula when <c:strCache> holds no usable value", () => {
    expect(
      parseSeriesName(ser("<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache/></c:strRef></c:tx>")),
    ).toBe("Sheet1!$B$1")
    expect(
      parseSeriesName(
        ser(
          "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>" +
            '<c:pt idx="0"><c:v>  </c:v></c:pt></c:strCache></c:strRef></c:tx>',
        ),
      ),
    ).toBe("Sheet1!$B$1")
  })

  it("keeps scanning the cache past entries with no usable <c:v>", () => {
    expect(
      parseSeriesName(
        ser(
          "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>" +
            '<c:ptCount val="2"/><c:pt idx="0"/><c:pt idx="1"><c:v>Units</c:v></c:pt>' +
            "</c:strCache></c:strRef></c:tx>",
        ),
      ),
    ).toBe("Units")
  })

  it("prefers the cached literal over the formula", () => {
    expect(
      parseSeriesName(
        ser(
          "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>" +
            '<c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>',
        ),
      ),
    ).toBe("Revenue")
  })

  it("returns undefined when <c:tx> holds nothing readable", () => {
    expect(parseSeriesName(ser(""))).toBeUndefined()
    expect(parseSeriesName(ser("<c:tx/>"))).toBeUndefined()
    expect(parseSeriesName(ser("<c:tx><c:v>  </c:v></c:tx>"))).toBeUndefined()
    expect(parseSeriesName(ser("<c:tx><c:strRef><c:f>  </c:f></c:strRef></c:tx>"))).toBeUndefined()
  })
})

describe("buildSeriesSpPr", () => {
  it("emits <a:schemeClr> for a theme-color reference rather than a hex triple", () => {
    const xml = buildSeriesSpPr({ theme: "accent3", lumMod: 60000 }, undefined)
    expect(xml).toContain('<a:schemeClr val="accent3">')
    expect(xml).toContain('<a:lumMod val="60000"/>')
  })

  it("drops a malformed hex and an unrecognized theme name alike", () => {
    expect(buildSeriesSpPr("nothex", undefined)).toBeUndefined()
    expect(buildSeriesSpPr({ theme: "accent9" as never }, undefined)).toBeUndefined()
  })

  it("omits cap / cmpd when they carry the OOXML defaults", () => {
    // `cap="flat"` and `cmpd="sng"` are the schema defaults, so emitting
    // them would make absence and the default differ byte-for-byte.
    const xml = buildSeriesSpPr(undefined, { width: 1, cap: "flat", compound: "sng" })
    expect(xml).toContain('<a:ln w="12700"/>')
    expect(xml).not.toContain("cap=")
    expect(xml).not.toContain("cmpd=")
  })

  it("emits an attribute-only <a:ln> when a cap or compound stands alone", () => {
    expect(buildSeriesSpPr(undefined, { cap: "sq" })).toContain('<a:ln cap="sq"/>')
    expect(buildSeriesSpPr(undefined, { compound: "thickThin" })).toContain(
      '<a:ln cmpd="thickThin"/>',
    )
  })

  it("returns undefined when neither fill nor stroke carries anything", () => {
    expect(buildSeriesSpPr(undefined, undefined)).toBeUndefined()
    expect(buildSeriesSpPr(undefined, {})).toBeUndefined()
  })
})

describe("buildSeries", () => {
  const values = "Sheet1!$B$2:$B$5"

  it("emits <c:shape> at the tail of a bar3D series", () => {
    const xml = buildSeries({ values } as ChartSeries, 0, "Data", false, {
      chartType: "bar",
      shape3D: "pyramid",
    })
    expect(xml).toContain('<c:shape val="pyramid"/>')
    // An unrecognized token elides the element entirely.
    expect(
      buildSeries({ values } as ChartSeries, 0, "Data", false, {
        chartType: "bar",
        shape3D: "sphere" as never,
      }),
    ).not.toContain("c:shape")
  })

  it("drops trendlines and error bars on the pie family, which has no slot", () => {
    const extras = {
      trendlines: [{ type: "linear" as const }],
      errorBars: [{ direction: "y" as const, type: "both" as const, valType: "fixedVal" as const }],
    }
    const pie = buildSeries({ values } as ChartSeries, 0, "Data", false, {
      chartType: "pie",
      ...extras,
    })
    expect(pie).not.toContain("c:trendline")
    expect(pie).not.toContain("c:errBars")
    expect(
      buildSeries({ values } as ChartSeries, 0, "Data", false, {
        chartType: "doughnut",
        ...extras,
      }),
    ).not.toContain("c:trendline")
    const bar = buildSeries({ values } as ChartSeries, 0, "Data", false, {
      chartType: "bar",
      ...extras,
    })
    expect(bar).toContain("<c:trendline>")
    expect(bar).toContain("<c:errBars>")
  })

  it("qualifies a bare <c:bubbleSize> range with the owning sheet", () => {
    const xml = buildSeries({ values } as ChartSeries, 0, "My Data", false, {
      chartType: "scatter",
      bubbleSize: "D2:D5",
    })
    expect(xml).toContain("<c:bubbleSize><c:numRef><c:f>'My Data'!D2:D5</c:f>")
    // An empty range is indistinguishable from absence.
    expect(
      buildSeries({ values } as ChartSeries, 0, "Data", false, {
        chartType: "scatter",
        bubbleSize: "",
      }),
    ).not.toContain("c:bubbleSize")
  })
})

describe("mergeSeries", () => {
  it("treats an explicit bubbleSize / shape3D override as replace-or-drop", () => {
    const src = { valuesRef: "Sheet1!$B$2:$B$5", bubbleSizeRef: "Sheet1!$D$2:$D$5" } as never
    // Absent key inherits the source reference.
    expect(mergeSeries(src, undefined, 0).bubbleSize).toBe("Sheet1!$D$2:$D$5")
    // Present-but-null drops it.
    expect(mergeSeries(src, { bubbleSize: null } as never, 0).bubbleSize).toBeUndefined()
    // Present-and-string replaces it.
    expect(mergeSeries(src, { bubbleSize: "Sheet1!$E$2:$E$5" } as never, 0).bubbleSize).toBe(
      "Sheet1!$E$2:$E$5",
    )
  })

  it("applies the same grammar to shape3D", () => {
    const src = { valuesRef: "Sheet1!$B$2:$B$5", shape3D: "cone" } as never
    expect(mergeSeries(src, undefined, 0).shape3D).toBe("cone")
    expect(mergeSeries(src, { shape3D: null } as never, 0).shape3D).toBeUndefined()
    expect(mergeSeries(src, { shape3D: "box" } as never, 0).shape3D).toBe("box")
  })

  it("inherits the source's parsed per-point overrides", () => {
    const src = {
      valuesRef: "Sheet1!$B$2:$B$5",
      dataPoints: [{ idx: 0, fillColor: "FF0000" }],
    } as never
    expect(mergeSeries(src, undefined, 0).dataPoints).toEqual([{ idx: 0, fillColor: "FF0000" }])
  })

  it("throws when neither the source nor the override supplies a values range", () => {
    expect(() => mergeSeries(undefined, undefined, 2)).toThrow(/series #2 has no values reference/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// dataLabels — <c:dLbls> (CT_DLbls, §21.2.2.50)
// ═══════════════════════════════════════════════════════════════════════

describe("data-label typography readers", () => {
  // Every reader walks `<c:txPr><a:p><a:pPr><a:defRPr>` and must bail on
  // the first missing link rather than fabricate a value.
  const readers: Array<[string, (e: XmlElement) => unknown]> = [
    ["fontSize", parseDataLabelsFontSize],
    ["fontColor", parseDataLabelsFontColor],
    ["bold", parseDataLabelsBold],
    ["italic", parseDataLabelsItalic],
    ["underline", parseDataLabelsUnderline],
    ["strikethrough", parseDataLabelsStrikethrough],
    ["fontFamily", parseDataLabelsFontFamily],
  ]

  it("bails at every broken link of the <c:txPr><a:p><a:pPr><a:defRPr> chain", () => {
    for (const [name, read] of readers) {
      for (const depth of [0, 1, 2, 3] as const) {
        expect(read(el("dLbls", txPrChain(depth))), `${name}@${depth}`).toBeUndefined()
      }
      expect(read(el("dLbls", "")), name).toBeUndefined()
    }
  })

  it("reads the whole typography block off a single <a:defRPr>", () => {
    const dLbls = el(
      "dLbls",
      '<c:txPr><a:p><a:pPr><a:defRPr sz="1050" b="1" i="1" u="sng" strike="sngStrike">' +
        '<a:solidFill><a:srgbClr val="336699"/></a:solidFill>' +
        '<a:latin typeface="  Consolas  "/></a:defRPr></a:pPr></a:p></c:txPr>',
    )
    expect(parseDataLabelsFontSize(dLbls)).toBe(10.5)
    expect(parseDataLabelsFontColor(dLbls)).toBe("336699")
    expect(parseDataLabelsBold(dLbls)).toBe(true)
    expect(parseDataLabelsItalic(dLbls)).toBe(true)
    expect(parseDataLabelsUnderline(dLbls)).toBe(true)
    expect(parseDataLabelsStrikethrough(dLbls)).toBe(true)
    expect(parseDataLabelsFontFamily(dLbls)).toBe("Consolas")
  })

  it("lifts an <a:schemeClr> font color when no literal sRGB is present", () => {
    expect(
      parseDataLabelsFontColor(
        el(
          "dLbls",
          "<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>" +
            '<a:schemeClr val="tx1"><a:alpha val="50000"/></a:schemeClr>' +
            "</a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>",
        ),
      ),
    ).toEqual({ theme: "tx1", alpha: 50000 })
  })

  it("drops an <a:latin> with no typeface attribute", () => {
    expect(
      parseDataLabelsFontFamily(
        el("dLbls", "<c:txPr><a:p><a:pPr><a:defRPr><a:latin/></a:defRPr></a:pPr></a:p></c:txPr>"),
      ),
    ).toBeUndefined()
  })

  it("drops an <a:solidFill> that carries neither an sRGB nor a theme colour", () => {
    expect(
      parseDataLabelsFontColor(
        el(
          "dLbls",
          "<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:hslClr/></a:solidFill>" +
            "</a:defRPr></a:pPr></a:p></c:txPr>",
        ),
      ),
    ).toBeUndefined()
  })

  it("drops an out-of-band or non-numeric sz", () => {
    const size = (sz: string): number | undefined =>
      parseDataLabelsFontSize(el("dLbls", txPrChain(3, ` sz="${sz}"`)))
    expect(size("1200")).toBe(12)
    // The band is 1..400 pt = 100..40000 in OOXML hundredths.
    expect(size("50")).toBeUndefined()
    expect(size("40100")).toBeUndefined()
    expect(size("   ")).toBeUndefined()
    expect(size("abc")).toBeUndefined()
  })
})

describe("parseDataLabels", () => {
  it("surfaces the <a:ln> cap and compound off the <c:spPr> block", () => {
    const parsed = parseDataLabels(
      el("dLbls", '<c:spPr><a:ln cap="rnd" cmpd="thickThin"/></c:spPr>'),
    )
    expect(parsed).toMatchObject({ borderCap: "rnd", borderCompound: "thickThin" })
  })

  it("drops an empty <c:separator> and a <c:numFmt> with no sourceLinked", () => {
    const parsed = parseDataLabels(
      el("dLbls", '<c:numFmt formatCode="0.0%"/><c:separator></c:separator>'),
    )
    expect(parsed).toEqual({ numberFormat: { formatCode: "0.0%" } })
  })

  it("collapses a block carrying nothing the writer models", () => {
    // A `<c:dLbls>` with only the OOXML defaults has no visible effect,
    // so it must not round-trip into a redundant write.
    expect(
      parseDataLabels(el("dLbls", '<c:spPr><a:ln cap="flat" cmpd="sng"/></c:spPr>')),
    ).toBeUndefined()
    expect(parseDataLabels(el("dLbls", ""))).toBeUndefined()
  })
})

describe("buildDataLabelsBody", () => {
  it("emits every show* toggle in CT_DLbls order, defaulting each to 0", () => {
    const xml = buildDataLabelsBody({} as ChartDataLabels, "bar")
    const order = [
      "c:showLegendKey",
      "c:showVal",
      "c:showCatName",
      "c:showSerName",
      "c:showPercent",
      "c:showBubbleSize",
    ]
    let cursor = -1
    for (const tag of order) {
      const at = xml.indexOf(tag)
      expect(at, tag).toBeGreaterThan(cursor)
      cursor = at
    }
    expect(xml).toContain('<c:showSerName val="0"/>')
  })

  it("flips <c:showSerName> only for a literal true", () => {
    expect(buildDataLabelsBody({ showSeriesName: true } as ChartDataLabels, "bar")).toContain(
      '<c:showSerName val="1"/>',
    )
    expect(buildDataLabelsBody({ showSeriesName: false } as ChartDataLabels, "bar")).toContain(
      '<c:showSerName val="0"/>',
    )
  })

  it("scopes <c:showLeaderLines> to the pie family and to an explicit false", () => {
    // The OOXML default is `true` (Excel paints leader lines), so only
    // an explicit `false` needs to reach the wire.
    const off = { showLeaderLines: false } as ChartDataLabels
    expect(buildDataLabelsBody(off, "pie")).toContain('<c:showLeaderLines val="0"/>')
    expect(buildDataLabelsBody(off, "doughnut")).toContain('<c:showLeaderLines val="0"/>')
    // Bar / column route through EG_DLblsShared, which has no slot for it.
    expect(buildDataLabelsBody(off, "bar")).not.toContain("c:showLeaderLines")
    expect(buildDataLabelsBody({ showLeaderLines: true } as ChartDataLabels, "pie")).not.toContain(
      "c:showLeaderLines",
    )
  })
})

// ═══════════════════════════════════════════════════════════════════════
// dataTable — <c:dTable> (CT_DTable, §21.2.2.54)
// ═══════════════════════════════════════════════════════════════════════

describe("data-table typography readers", () => {
  const readers: Array<[string, (e: XmlElement) => unknown]> = [
    ["bold", parseDataTableBold],
    ["italic", parseDataTableItalic],
    ["underline", parseDataTableUnderline],
    ["strikethrough", parseDataTableStrikethrough],
    ["fontFamily", parseDataTableFontFamily],
    ["fontColor", parseDataTableFontColor],
    ["fontSize", parseDataTableFontSize],
  ]

  it("bails at every broken link of the <c:txPr><a:p><a:pPr><a:defRPr> chain", () => {
    for (const [name, read] of readers) {
      for (const depth of [0, 1, 2, 3] as const) {
        expect(read(el("dTable", txPrChain(depth))), `${name}@${depth}`).toBeUndefined()
      }
    }
  })

  it('round-trips an explicit b="0" / i="0" rather than collapsing it', () => {
    // Unlike the title / legend readers, the data-table readers surface
    // the literal `false` so a clone target can override an upstream
    // `b="1"` without the flag silently reverting to inherit.
    const dTable = el("dTable", txPrChain(3, ' b="0" i="false"'))
    expect(parseDataTableBold(dTable)).toBe(false)
    expect(parseDataTableItalic(dTable)).toBe(false)
    expect(parseDataTableBold(el("dTable", txPrChain(3, ' b="yes"')))).toBeUndefined()
  })

  it("drops a data-table <a:solidFill> with neither an sRGB nor a theme colour", () => {
    expect(
      parseDataTableFontColor(
        el(
          "dTable",
          "<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:prstClr/></a:solidFill>" +
            "</a:defRPr></a:pPr></a:p></c:txPr>",
        ),
      ),
    ).toBeUndefined()
    // A blank `sz` is indistinguishable from absence.
    expect(parseDataTableFontSize(el("dTable", txPrChain(3, ' sz="   "')))).toBeUndefined()
  })

  it("lifts an <a:schemeClr> font color when no literal sRGB is present", () => {
    expect(
      parseDataTableFontColor(
        el(
          "dTable",
          "<c:txPr><a:p><a:pPr><a:defRPr><a:solidFill>" +
            '<a:schemeClr val="dk2"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>',
        ),
      ),
    ).toEqual({ theme: "dk2" })
  })
})

describe("parseDataTableFlag", () => {
  it("accepts both OOXML boolean spellings and drops everything else", () => {
    const dTable = el(
      "dTable",
      '<c:showHorzBorder val="1"/><c:showVertBorder val="false"/>' +
        '<c:showOutline val="maybe"/><c:showKeys/>',
    )
    expect(parseDataTableFlag(dTable, "showHorzBorder")).toBe(true)
    expect(parseDataTableFlag(dTable, "showVertBorder")).toBe(false)
    expect(parseDataTableFlag(dTable, "showOutline")).toBeUndefined()
    expect(parseDataTableFlag(dTable, "showKeys")).toBeUndefined()
    // A missing element is absence, not a fabricated default.
    expect(parseDataTableFlag(dTable, "showLegendKey")).toBeUndefined()
  })
})

describe("parseDataTable", () => {
  it("surfaces the <a:ln> cap and compound off the <c:spPr> block", () => {
    const parsed = parseDataTable(
      el("plotArea", '<c:dTable><c:spPr><a:ln cap="sq" cmpd="dbl"/></c:spPr></c:dTable>'),
    )
    expect(parsed).toMatchObject({ borderCap: "sq", borderCompound: "dbl" })
  })

  it("returns undefined when the plot area declares no <c:dTable>", () => {
    expect(parseDataTable(el("plotArea", ""))).toBeUndefined()
  })
})
