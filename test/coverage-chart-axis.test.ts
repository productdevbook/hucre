import { describe, expect, it } from "vitest"
import type {
  ChartAxisDispUnits,
  ChartAxisGridlines,
  ChartAxisNumberFormat,
  ChartAxisScale,
  ChartLegendEntry,
  SheetChart,
} from "../src/_types"
import { type XmlElement, parseXml } from "../src/xml/parser"
import {
  applyAutoOverride,
  applyAxisTitleBoldOverride,
  applyAxisTitleItalicOverride,
  applyAxisTitleOverlayOverride,
  applyAxisTitleStrikeOverride,
  applyAxisTitleUnderlineOverride,
  applyCrossBetweenOverride,
  applyCrossesOverride,
  applyDispUnitsOverride,
  applyGridlinesOverride,
  applyHiddenOverride,
  applyLabelBoldOverride,
  applyLabelColorOverride,
  applyLabelFontFamilyOverride,
  applyLabelFontSizeOverride,
  applyLabelItalicOverride,
  applyLabelStrikeOverride,
  applyLabelUnderlineOverride,
  applyLblAlgnOverride,
  applyLblOffsetOverride,
  applyNoMultiLvlLblOverride,
  applyNumberFormatOverride,
  applyScaleOverride,
  applySkipOverride,
  applyTickLblPosOverride,
  applyTickMarkOverride,
  buildAxisDispUnits,
  cloneScale,
  normalizeAxisDispUnits,
  normalizeAxisNumberFormat,
  normalizeAxisSkip,
  normalizeDispUnits,
  parseAxisAuto,
  parseAxisCrossBetween,
  parseAxisCrosses,
  parseAxisDispUnits,
  parseAxisGridlines,
  parseAxisHidden,
  parseAxisInfo,
  parseAxisLabelBold,
  parseAxisLabelColor,
  parseAxisLabelFontSize,
  parseAxisLabelItalic,
  parseAxisLabelFontFamily,
  parseAxisLabelRotation,
  parseAxisLabelStrike,
  parseAxisLabelUnderline,
  parseAxisLblAlgn,
  parseAxisLblOffset,
  parseAxisNoMultiLvlLbl,
  parseAxisNumberFormat,
  parseAxisReverse,
  parseAxisScale,
  parseAxisSkip,
  parseAxisTickLblPos,
  parseAxisTickMark,
  parseAxisTitle,
  parseAxisTitleBold,
  parseAxisTitleColor,
  parseAxisTitleFontFamily,
  parseAxisTitleFontSize,
  parseAxisTitleItalic,
  parseAxisTitleOverlay,
  parseAxisTitleRotation,
  parseAxisTitleStrike,
  parseAxisTitleUnderline,
  resolveAxes,
} from "../src/xlsx/chart/axis"
import {
  parseLegend,
  parseLegendBold,
  parseLegendBorderCap,
  parseLegendBorderColor,
  parseLegendBorderCompound,
  parseLegendBorderDash,
  parseLegendBorderWidth,
  parseLegendEntries,
  parseLegendFillColor,
  parseLegendFontColor,
  parseLegendFontFamily,
  parseLegendFontSize,
  parseLegendItalic,
  parseLegendLayout,
  parseLegendOverlay,
  parseLegendStrikethrough,
  parseLegendUnderline,
  resolveLegendEntries,
  resolveLegendPosition,
} from "../src/xlsx/chart/legend"
import {
  parseTitle,
  parseTitleBold,
  parseTitleColor,
  parseTitleFontFamily,
  parseTitleFontSize,
  parseTitleItalic,
  parseTitleRotation,
  parseTitleStrike,
  parseTitleUnderline,
} from "../src/xlsx/chart/title"
import {
  applyOverride,
  childElements,
  elementText,
  findChild,
  findDescendant,
  formulaText,
  parseBoolAttr,
  parseNumericChildVal,
  readBoolAttr,
  readBoolVal,
} from "../src/xlsx/chart/util"
import {
  normalizeChartColor,
  parseSchemeClr,
  resolveChartColor,
  resolveLineCap,
  resolveLineCompound,
} from "../src/xlsx/chart/shape"
import {
  buildPieChart,
  buildPlotArea,
  parseBarGrouping,
  parseFirstSliceAng,
  parseGapWidth,
  parseHoleSize,
  parseLineAreaGrouping,
  parseOverlap,
  parseUpDownBarsGapWidth,
} from "../src/xlsx/chart/plotArea"
import {
  buildManualLayout,
  normalizeLayoutCoordinate,
  parseManualLayout,
  readLayoutCoordinate,
} from "../src/xlsx/chart/layout"
import {
  buildFloorThickness,
  parseBackWallThickness,
  parseFloorThickness,
  parseSideWallThickness,
  parseView3D,
} from "../src/xlsx/chart/walls"

// ── Helpers ──────────────────────────────────────────────────────────

const NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

/** Build a chart-namespace element through the project's parser. */
function el(tag: string, inner = ""): XmlElement {
  return parseXml(`<c:${tag} ${NS}>${inner}</c:${tag}>`)
}

/** A `<c:valAx>` carrying the supplied children. */
function axis(inner: string): XmlElement {
  return el("valAx", inner)
}

/** A `<c:chart>` carrying the supplied children. */
function chartEl(inner: string): XmlElement {
  return el("chart", inner)
}

/**
 * The `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr>` chain every
 * title reader walks. `depth` truncates it link by link.
 */
function titleChain(depth: 0 | 1 | 2 | 3 | 4 | 5, defRPrAttrs = "", defRPrBody = ""): string {
  const inner = [
    "<c:title/>",
    "<c:title><c:tx/></c:title>",
    "<c:title><c:tx><c:rich/></c:tx></c:title>",
    "<c:title><c:tx><c:rich><a:p/></c:rich></c:tx></c:title>",
    "<c:title><c:tx><c:rich><a:p><a:pPr/></a:p></c:rich></c:tx></c:title>",
    `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr${defRPrAttrs}>${defRPrBody}</a:defRPr>` +
      "</a:pPr></a:p></c:rich></c:tx></c:title>",
  ]
  return inner[depth]
}

/** The `<c:txPr><a:p><a:pPr><a:defRPr>` chain, truncated at `depth`. */
function txPrChain(depth: 0 | 1 | 2 | 3, defRPrAttrs = "", defRPrBody = ""): string {
  const inner = [
    "<c:txPr/>",
    "<c:txPr><a:p/></c:txPr>",
    "<c:txPr><a:p><a:pPr/></a:p></c:txPr>",
    `<c:txPr><a:p><a:pPr><a:defRPr${defRPrAttrs}>${defRPrBody}</a:defRPr></a:pPr></a:p></c:txPr>`,
  ]
  return inner[depth]
}

// ═══════════════════════════════════════════════════════════════════════
// chart/util — the XML walk every per-host reader is built on
// ═══════════════════════════════════════════════════════════════════════

describe("findChild / findDescendant / childElements", () => {
  const tree = el("plotArea", 'text<c:valAx><c:scaling><c:min val="0"/></c:scaling></c:valAx>')

  it("matches on the local name, ignoring the namespace prefix", () => {
    expect(findChild(tree, "valAx")?.local).toBe("valAx")
    expect(findChild(tree, "scaling")).toBeUndefined()
  })

  it("returns the element itself when it already matches", () => {
    expect(findDescendant(tree, "plotArea")).toBe(tree)
  })

  it("descends through text nodes to reach a nested match", () => {
    expect(findDescendant(tree, "min")?.attrs.val).toBe("0")
    expect(findDescendant(tree, "logBase")).toBeUndefined()
  })

  it("skips text nodes when listing children", () => {
    expect(childElements(tree).map((c) => c.local)).toEqual(["valAx"])
  })
})

describe("elementText", () => {
  it("concatenates nested run text, not just the direct children", () => {
    // `<a:t>` bodies can be split across entity boundaries, so the walk
    // has to recurse rather than read a single text node.
    expect(elementText(el("v", "a<c:x>b</c:x>c"))).toBe("abc")
    expect(elementText(el("v", ""))).toBe("")
  })
})

describe("formulaText", () => {
  it("prefers <c:numRef>/<c:strRef> but falls back to a direct <c:f>", () => {
    expect(formulaText(el("val", "<c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef>"))).toBe(
      "Sheet1!$A$1:$A$3",
    )
    expect(formulaText(el("cat", "<c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef>"))).toBe(
      "Sheet1!$B$1",
    )
    // Some writers hoist `<c:f>` straight onto the wrapper.
    expect(formulaText(el("val", "<c:f>Sheet1!$C$1</c:f>"))).toBe("Sheet1!$C$1")
  })

  it("returns undefined for literal data and for blank formulas", () => {
    expect(formulaText(undefined)).toBeUndefined()
    expect(
      formulaText(el("val", '<c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit>')),
    ).toBeUndefined()
    expect(formulaText(el("val", "<c:numRef><c:f>  </c:f></c:numRef>"))).toBeUndefined()
    expect(formulaText(el("val", "<c:numRef/>"))).toBeUndefined()
    expect(formulaText(el("val", "<c:f>   </c:f>"))).toBeUndefined()
  })
})

describe("parseNumericChildVal", () => {
  it("admits any finite xsd:double, including exponent notation", () => {
    const parent = el("scaling", '<c:min val="-2.5"/><c:max val="1e3"/><c:logBase val=" 10 "/>')
    expect(parseNumericChildVal(parent, "min")).toBe(-2.5)
    expect(parseNumericChildVal(parent, "max")).toBe(1000)
    expect(parseNumericChildVal(parent, "logBase")).toBe(10)
  })

  it("drops absence, a missing val, a blank val, and non-numeric tokens", () => {
    const parent = el("scaling", '<c:min/><c:max val="  "/><c:logBase val="ten"/>')
    expect(parseNumericChildVal(parent, "orientation")).toBeUndefined()
    expect(parseNumericChildVal(parent, "min")).toBeUndefined()
    expect(parseNumericChildVal(parent, "max")).toBeUndefined()
    expect(parseNumericChildVal(parent, "logBase")).toBeUndefined()
  })
})

describe("parseBoolAttr / readBoolAttr / readBoolVal", () => {
  it("parseBoolAttr accepts the OOXML spellings case-insensitively after trimming", () => {
    expect(parseBoolAttr(" TRUE ")).toBe(true)
    expect(parseBoolAttr("1")).toBe(true)
    expect(parseBoolAttr("False")).toBe(false)
    expect(parseBoolAttr("0")).toBe(false)
    expect(parseBoolAttr("yes")).toBeUndefined()
    expect(parseBoolAttr(1)).toBeUndefined()
    expect(parseBoolAttr(undefined)).toBeUndefined()
  })

  it("readBoolAttr treats anything but a truthy spelling as false, never undefined", () => {
    // `<c:smooth val="0"/>` and `<c:smooth val="junk"/>` both mean "off"
    // to Excel, so the element-level reader collapses them together.
    expect(readBoolAttr(parseXml(`<c:smooth ${NS} val="1"/>`))).toBe(true)
    expect(readBoolAttr(parseXml(`<c:smooth ${NS} val="TRUE"/>`))).toBe(true)
    expect(readBoolAttr(parseXml(`<c:smooth ${NS} val="0"/>`))).toBe(false)
    expect(readBoolAttr(parseXml(`<c:smooth ${NS} val="junk"/>`))).toBe(false)
    expect(readBoolAttr(parseXml(`<c:smooth ${NS}/>`))).toBeUndefined()
  })

  it("readBoolVal is exact-match only — no trimming, no case folding", () => {
    expect(readBoolVal("1")).toBe(true)
    expect(readBoolVal("true")).toBe(true)
    expect(readBoolVal("0")).toBe(false)
    expect(readBoolVal("false")).toBe(false)
    expect(readBoolVal("True")).toBeUndefined()
    expect(readBoolVal(" 1")).toBeUndefined()
    expect(readBoolVal(undefined)).toBeUndefined()
  })
})

describe("applyOverride", () => {
  it("implements the inherit / suppress / replace grammar every clone resolver shares", () => {
    expect(applyOverride("src", undefined)).toBe("src")
    expect(applyOverride("src", null)).toBeUndefined()
    expect(applyOverride("src", "ov")).toBe("ov")
    // A falsy replacement still replaces — only `null` suppresses.
    expect(applyOverride(1, 0)).toBe(0)
    expect(applyOverride(undefined, undefined)).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// chart/shape — <a:schemeClr> and the cap / compound resolvers
// ═══════════════════════════════════════════════════════════════════════

describe("parseSchemeClr", () => {
  it("reads every supported modifier that lands in its percentage band", () => {
    const clr = parseXml(
      `<a:schemeClr ${NS} val="accent1"><a:lumMod val="75000"/><a:lumOff val="25000"/>` +
        '<a:tint val="40000"/><a:shade val="60000"/><a:alpha val="50000"/></a:schemeClr>',
    )
    expect(parseSchemeClr(clr)).toEqual({
      theme: "accent1",
      lumMod: 75000,
      lumOff: 25000,
      tint: 40000,
      shade: 60000,
      alpha: 50000,
    })
  })

  it("drops a modifier whose val is missing, non-numeric, or out of band", () => {
    // A garbage modifier must not sink the whole colour — the theme
    // reference still round-trips, just without the bad mod.
    const clr = parseXml(
      `<a:schemeClr ${NS} val="tx1"><a:lumMod/><a:lumOff val="nope"/>` +
        '<a:alpha val="500000"/></a:schemeClr>',
    )
    expect(parseSchemeClr(clr)).toEqual({ theme: "tx1" })
  })

  it("drops the element when val is absent or not an ST_SchemeColorVal name", () => {
    expect(parseSchemeClr(parseXml(`<a:schemeClr ${NS}/>`))).toBeUndefined()
    expect(parseSchemeClr(parseXml(`<a:schemeClr ${NS} val="accent9"/>`))).toBeUndefined()
  })
})

describe("resolveChartColor / resolveLineCap / resolveLineCompound", () => {
  it("follows the inherit / drop / replace grammar and normalizes both sides", () => {
    expect(resolveChartColor("#ff0000", undefined)).toBe("FF0000")
    expect(resolveChartColor("FF0000", null)).toBeUndefined()
    expect(resolveChartColor("FF0000", "#00ff00")).toBe("00FF00")
    expect(resolveChartColor("FF0000", "nothex")).toBeUndefined()
  })

  it("collapses the OOXML defaults cap=flat and cmpd=sng on both source and override", () => {
    expect(resolveLineCap("rnd", undefined)).toBe("rnd")
    expect(resolveLineCap("rnd", null)).toBeUndefined()
    expect(resolveLineCap("rnd", "flat")).toBeUndefined()
    expect(resolveLineCap("rnd", "bevel" as never)).toBeUndefined()
    expect(resolveLineCompound("dbl", undefined)).toBe("dbl")
    expect(resolveLineCompound("dbl", null)).toBeUndefined()
    expect(resolveLineCompound("dbl", "sng")).toBeUndefined()
    expect(resolveLineCompound("dbl", "quad" as never)).toBeUndefined()
  })

  it("normalizeChartColor keeps a theme reference but drops an unknown theme name", () => {
    expect(normalizeChartColor({ theme: "bg1" })).toEqual({ theme: "bg1" })
    expect(normalizeChartColor({ theme: "accent9" as never })).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// chart/layout — <c:manualLayout> (CT_ManualLayout, §21.2.2.115)
// ═══════════════════════════════════════════════════════════════════════

describe("readLayoutCoordinate", () => {
  it("admits only a finite xsd:double inside the 0..1 chart-frame band", () => {
    const read = (attr: string): number | undefined =>
      readLayoutCoordinate(parseXml(`<c:x ${NS} ${attr}/>`))
    expect(read('val="0.25"')).toBe(0.25)
    expect(read('val="0"')).toBe(0)
    expect(read('val="1"')).toBe(1)
    expect(read('val="1.5"')).toBeUndefined()
    expect(read('val="-0.1"')).toBeUndefined()
    expect(read('val="  "')).toBeUndefined()
    expect(read('val="half"')).toBeUndefined()
    expect(read("")).toBeUndefined()
    expect(readLayoutCoordinate(undefined)).toBeUndefined()
  })
})

describe("parseManualLayout / buildManualLayout", () => {
  it("collapses to undefined when every coordinate dropped on normalization", () => {
    expect(parseManualLayout(el("legend", ""))).toBeUndefined()
    expect(parseManualLayout(el("legend", "<c:layout/>"))).toBeUndefined()
    expect(
      parseManualLayout(
        el("legend", '<c:layout><c:manualLayout><c:x val="9"/></c:manualLayout></c:layout>'),
      ),
    ).toBeUndefined()
  })

  it("emits an xMode/yMode edge pair before the coordinates, per CT_ManualLayout order", () => {
    const xml = buildManualLayout({ x: 0.1, h: 0.4 })
    expect(xml).toBe(
      "<c:layout><c:manualLayout>" +
        '<c:xMode val="edge"/><c:hMode val="edge"/><c:x val="0.1"/><c:h val="0.4"/>' +
        "</c:manualLayout></c:layout>",
    )
  })

  it("returns undefined rather than a bare <c:layout> shell", () => {
    expect(buildManualLayout(undefined)).toBeUndefined()
    expect(buildManualLayout({})).toBeUndefined()
    expect(normalizeLayoutCoordinate("0.5")).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// chart/walls — <c:view3D> / <c:floor> / <c:sideWall> / <c:backWall>
// ═══════════════════════════════════════════════════════════════════════

describe("parseView3D", () => {
  it('surfaces <c:rAngAx val="0"/> as an explicit false', () => {
    // Excel writes `rAngAx="0"` when the user turns off "Right Angle
    // Axes" on a 3-D chart, so the literal false has to survive.
    expect(parseView3D(chartEl('<c:view3D><c:rAngAx val="0"/></c:view3D>'))).toEqual({
      rAngAx: false,
    })
    expect(parseView3D(chartEl('<c:view3D><c:rAngAx val="false"/></c:view3D>'))).toEqual({
      rAngAx: false,
    })
    expect(parseView3D(chartEl('<c:view3D><c:rAngAx val="1"/></c:view3D>'))).toEqual({
      rAngAx: true,
    })
    expect(parseView3D(chartEl('<c:view3D><c:rAngAx val="maybe"/></c:view3D>'))).toEqual({})
    expect(parseView3D(chartEl("<c:view3D><c:rAngAx/></c:view3D>"))).toEqual({})
  })

  it("preserves a bare <c:view3D/> shell as an empty record", () => {
    // `buildView3D` re-emits `{}` as `<c:view3D/>`, so the reader keeps
    // the empty object rather than collapsing the element away.
    expect(parseView3D(chartEl("<c:view3D/>"))).toEqual({})
  })

  it("drops a <c:rotX> digit string so long it overflows to Infinity", () => {
    // `Number("9".repeat(400))` is `Infinity` — it passes the signed-
    // integer regex but is not an integer, so the guard behind the
    // regex is what keeps it out.
    const huge = "9".repeat(400)
    expect(parseView3D(chartEl(`<c:view3D><c:rotX val="${huge}"/></c:view3D>`))).toEqual({})
  })

  it("returns undefined when the chart declares no <c:view3D>", () => {
    expect(parseView3D(chartEl(""))).toBeUndefined()
  })
})

describe("wall / floor thickness readers", () => {
  const readers: Array<[string, string, (e: XmlElement) => number | undefined]> = [
    ["floor", "floor", parseFloorThickness],
    ["sideWall", "sideWall", parseSideWallThickness],
    ["backWall", "backWall", parseBackWallThickness],
  ]

  it("reads ST_Thickness as a strict unsigned integer, collapsing the default 0", () => {
    for (const [name, tag, read] of readers) {
      expect(read(chartEl(`<c:${tag}><c:thickness val="25"/></c:${tag}>`)), name).toBe(25)
      expect(read(chartEl(`<c:${tag}><c:thickness val="0"/></c:${tag}>`)), name).toBeUndefined()
      expect(read(chartEl(`<c:${tag}><c:thickness val="-5"/></c:${tag}>`)), name).toBeUndefined()
      expect(read(chartEl(`<c:${tag}><c:thickness val="2.5"/></c:${tag}>`)), name).toBeUndefined()
      expect(read(chartEl(`<c:${tag}><c:thickness/></c:${tag}>`)), name).toBeUndefined()
      expect(read(chartEl(`<c:${tag}/>`)), name).toBeUndefined()
      expect(read(chartEl("")), name).toBeUndefined()
    }
  })

  it("drops a digit string so long it overflows to Infinity", () => {
    // `Number("9".repeat(400))` is `Infinity`, which passes the digits-
    // only regex but is not an integer — the guard after the regex is
    // what stops it reaching the range check.
    const huge = "9".repeat(400)
    for (const [name, tag, read] of readers) {
      expect(
        read(chartEl(`<c:${tag}><c:thickness val="${huge}"/></c:${tag}>`)),
        name,
      ).toBeUndefined()
    }
  })

  it("clamps nothing on the write side — out-of-band values elide the element", () => {
    expect(buildFloorThickness(25)).toBe('<c:floor><c:thickness val="25"/></c:floor>')
    expect(buildFloorThickness(0)).toBeUndefined()
    expect(buildFloorThickness(101)).toBeUndefined()
    expect(buildFloorThickness(2.5)).toBeUndefined()
    expect(buildFloorThickness(Number.NaN)).toBeUndefined()
    expect(buildFloorThickness("25" as never)).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// chart/plotArea — per-family numeric and grouping readers
// ═══════════════════════════════════════════════════════════════════════

describe("parseBarGrouping / parseLineAreaGrouping", () => {
  it("maps the ST_BarGrouping tokens Excel's UI exposes", () => {
    const g = (v: string): unknown => parseBarGrouping(el("barChart", `<c:grouping val="${v}"/>`))
    expect(g("stacked")).toBe("stacked")
    expect(g("percentStacked")).toBe("percentStacked")
    expect(g("clustered")).toBe("clustered")
    // `standard` renders side-by-side exactly like `clustered`, so it
    // collapses and the cloned chart inherits the writer's default.
    expect(g("standard")).toBeUndefined()
    expect(g("pyramid")).toBeUndefined()
    expect(parseBarGrouping(el("barChart", "<c:grouping/>"))).toBeUndefined()
    expect(parseBarGrouping(el("barChart", ""))).toBeUndefined()
  })

  it("maps the ST_Grouping tokens for line and area, dropping the default", () => {
    const g = (v: string): unknown =>
      parseLineAreaGrouping(el("lineChart", `<c:grouping val="${v}"/>`))
    expect(g("stacked")).toBe("stacked")
    expect(g("percentStacked")).toBe("percentStacked")
    expect(g("standard")).toBeUndefined()
    expect(g("clustered")).toBeUndefined()
    expect(parseLineAreaGrouping(el("lineChart", "<c:grouping/>"))).toBeUndefined()
    expect(parseLineAreaGrouping(el("lineChart", ""))).toBeUndefined()
  })
})

describe("plot-area numeric readers", () => {
  it("accepts the full 1..99 <c:holeSize> schema band, wider than Excel's UI", () => {
    const h = (v: string): number | undefined =>
      parseHoleSize(el("doughnutChart", `<c:holeSize val="${v}"/>`))
    expect(h("1")).toBe(1)
    expect(h("99")).toBe(99)
    expect(h("0")).toBeUndefined()
    expect(h("100")).toBeUndefined()
    expect(h("half")).toBeUndefined()
    expect(parseHoleSize(el("doughnutChart", "<c:holeSize/>"))).toBeUndefined()
    expect(parseHoleSize(el("doughnutChart", ""))).toBeUndefined()
  })

  it("drops an out-of-band <c:gapWidth> rather than clamp it", () => {
    // A clamp would silently rewrite a corrupt template as a different
    // gap; dropping lets the writer's default take over instead.
    const g = (v: string): number | undefined =>
      parseGapWidth(el("barChart", `<c:gapWidth val="${v}"/>`))
    expect(g("0")).toBe(0)
    expect(g("500")).toBe(500)
    expect(g("150")).toBeUndefined() // ST_GapAmount default
    expect(g("501")).toBeUndefined()
    expect(g("-1")).toBeUndefined()
    expect(g("wide")).toBeUndefined()
    expect(parseGapWidth(el("barChart", "<c:gapWidth/>"))).toBeUndefined()
    expect(parseGapWidth(el("barChart", ""))).toBeUndefined()
  })

  it("applies the same band to the up/down-bars <c:gapWidth>", () => {
    const g = (v: string): number | undefined =>
      parseUpDownBarsGapWidth(el("upDownBars", `<c:gapWidth val="${v}"/>`))
    expect(g("100")).toBe(100)
    expect(g("150")).toBeUndefined()
    expect(g("600")).toBeUndefined()
    expect(g("x")).toBeUndefined()
    expect(parseUpDownBarsGapWidth(el("upDownBars", "<c:gapWidth/>"))).toBeUndefined()
    expect(parseUpDownBarsGapWidth(el("upDownBars", ""))).toBeUndefined()
  })

  it("surfaces the literal <c:overlap>, including Excel's stacked-chart 100", () => {
    const o = (v: string): number | undefined =>
      parseOverlap(el("barChart", `<c:overlap val="${v}"/>`))
    expect(o("100")).toBe(100)
    expect(o("-100")).toBe(-100)
    expect(o("0")).toBeUndefined() // ST_Overlap default
    expect(o("101")).toBeUndefined()
    expect(o("-101")).toBeUndefined()
    expect(o("none")).toBeUndefined()
    expect(parseOverlap(el("barChart", "<c:overlap/>"))).toBeUndefined()
    expect(parseOverlap(el("barChart", ""))).toBeUndefined()
  })

  it("treats <c:firstSliceAng> 0 and 360 as the same 12-o'clock default", () => {
    const a = (v: string): number | undefined =>
      parseFirstSliceAng(el("pieChart", `<c:firstSliceAng val="${v}"/>`))
    expect(a("90")).toBe(90)
    expect(a("0")).toBeUndefined()
    expect(a("360")).toBeUndefined()
    expect(a("361")).toBeUndefined()
    expect(a("-1")).toBeUndefined()
    expect(a("north")).toBeUndefined()
    expect(parseFirstSliceAng(el("pieChart", "<c:firstSliceAng/>"))).toBeUndefined()
    expect(parseFirstSliceAng(el("pieChart", ""))).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// axis — scaling, tick marks, skips and number formats
// ═══════════════════════════════════════════════════════════════════════

describe("parseAxisTickMark", () => {
  it("drops the per-element OOXML default so absence and the default match", () => {
    // CT_CatAx defaults majorTickMark to "out" and minorTickMark to "none".
    expect(
      parseAxisTickMark(axis('<c:majorTickMark val="out"/>'), "majorTickMark", "out"),
    ).toBeUndefined()
    expect(parseAxisTickMark(axis('<c:majorTickMark val="cross"/>'), "majorTickMark", "out")).toBe(
      "cross",
    )
    expect(
      parseAxisTickMark(axis('<c:minorTickMark val="none"/>'), "minorTickMark", "none"),
    ).toBeUndefined()
    expect(parseAxisTickMark(axis('<c:minorTickMark val=" in "/>'), "minorTickMark", "none")).toBe(
      "in",
    )
  })

  it("drops absence, a missing val, and tokens outside ST_TickMark", () => {
    expect(parseAxisTickMark(axis(""), "majorTickMark", "out")).toBeUndefined()
    expect(parseAxisTickMark(axis("<c:majorTickMark/>"), "majorTickMark", "out")).toBeUndefined()
    expect(
      parseAxisTickMark(axis('<c:majorTickMark val="both"/>'), "majorTickMark", "out"),
    ).toBeUndefined()
  })
})

describe("parseAxisTickLblPos", () => {
  it("drops the ST_TickLblPos default `nextTo` and unknown tokens", () => {
    expect(parseAxisTickLblPos(axis('<c:tickLblPos val="low"/>'))).toBe("low")
    expect(parseAxisTickLblPos(axis('<c:tickLblPos val="nextTo"/>'))).toBeUndefined()
    expect(parseAxisTickLblPos(axis('<c:tickLblPos val="outside"/>'))).toBeUndefined()
    expect(parseAxisTickLblPos(axis("<c:tickLblPos/>"))).toBeUndefined()
    expect(parseAxisTickLblPos(axis(""))).toBeUndefined()
  })
})

describe("parseAxisReverse", () => {
  it("surfaces true only for the ST_Orientation value maxMin", () => {
    expect(parseAxisReverse(axis('<c:scaling><c:orientation val="maxMin"/></c:scaling>'))).toBe(
      true,
    )
    // `minMax` is the default — collapsing it keeps absence and the
    // default indistinguishable on re-emit.
    expect(
      parseAxisReverse(axis('<c:scaling><c:orientation val="minMax"/></c:scaling>')),
    ).toBeUndefined()
    expect(
      parseAxisReverse(axis('<c:scaling><c:orientation val="upDown"/></c:scaling>')),
    ).toBeUndefined()
    expect(parseAxisReverse(axis("<c:scaling><c:orientation/></c:scaling>"))).toBeUndefined()
    expect(parseAxisReverse(axis("<c:scaling/>"))).toBeUndefined()
    expect(parseAxisReverse(axis(""))).toBeUndefined()
  })
})

describe("parseAxisSkip", () => {
  it("admits only the ST_SkipIntervals band 1..32767, minus the default 1", () => {
    const s = (v: string): number | undefined =>
      parseAxisSkip(axis(`<c:tickLblSkip val="${v}"/>`), "tickLblSkip")
    expect(s("2")).toBe(2)
    expect(s("32767")).toBe(32767)
    expect(s("1")).toBeUndefined()
    expect(s("0")).toBeUndefined()
    expect(s("32768")).toBeUndefined()
    expect(s("  ")).toBeUndefined()
    expect(s("every")).toBeUndefined()
    expect(parseAxisSkip(axis("<c:tickLblSkip/>"), "tickLblSkip")).toBeUndefined()
    expect(parseAxisSkip(axis(""), "tickMarkSkip")).toBeUndefined()
  })
})

describe("parseAxisLblOffset", () => {
  it("admits only the ST_LblOffsetPercent band 0..1000, minus the default 100", () => {
    const o = (v: string): number | undefined =>
      parseAxisLblOffset(axis(`<c:lblOffset val="${v}"/>`))
    expect(o("0")).toBe(0)
    expect(o("1000")).toBe(1000)
    expect(o("100")).toBeUndefined()
    expect(o("1001")).toBeUndefined()
    expect(o("-1")).toBeUndefined()
    expect(o("   ")).toBeUndefined()
    expect(o("far")).toBeUndefined()
    expect(parseAxisLblOffset(axis("<c:lblOffset/>"))).toBeUndefined()
    expect(parseAxisLblOffset(axis(""))).toBeUndefined()
  })
})

describe("parseAxisLblAlgn", () => {
  it("drops the ST_LblAlgn default `ctr` and unknown tokens", () => {
    expect(parseAxisLblAlgn(axis('<c:lblAlgn val="l"/>'))).toBe("l")
    expect(parseAxisLblAlgn(axis('<c:lblAlgn val=" r "/>'))).toBe("r")
    expect(parseAxisLblAlgn(axis('<c:lblAlgn val="ctr"/>'))).toBeUndefined()
    expect(parseAxisLblAlgn(axis('<c:lblAlgn val="just"/>'))).toBeUndefined()
    expect(parseAxisLblAlgn(axis("<c:lblAlgn/>"))).toBeUndefined()
    expect(parseAxisLblAlgn(axis(""))).toBeUndefined()
  })
})

describe("parseAxisNoMultiLvlLbl / parseAxisAuto / parseAxisHidden", () => {
  it("each collapses its own OOXML default and surfaces only the opposite state", () => {
    // `<c:noMultiLvlLbl>` and `<c:delete>` default to false; `<c:auto>`
    // defaults to true — so `auto` is the one that collapses `true`.
    expect(parseAxisNoMultiLvlLbl(axis('<c:noMultiLvlLbl val="1"/>'))).toBe(true)
    expect(parseAxisNoMultiLvlLbl(axis('<c:noMultiLvlLbl val="0"/>'))).toBeUndefined()
    expect(parseAxisAuto(axis('<c:auto val="0"/>'))).toBe(false)
    expect(parseAxisAuto(axis('<c:auto val="1"/>'))).toBeUndefined()
    expect(parseAxisHidden(axis('<c:delete val="true"/>'))).toBe(true)
    expect(parseAxisHidden(axis('<c:delete val="false"/>'))).toBeUndefined()
  })

  it("drops absence, a missing val, and unknown tokens on all three", () => {
    expect(parseAxisNoMultiLvlLbl(axis("<c:noMultiLvlLbl/>"))).toBeUndefined()
    expect(parseAxisNoMultiLvlLbl(axis('<c:noMultiLvlLbl val="on"/>'))).toBeUndefined()
    expect(parseAxisNoMultiLvlLbl(axis(""))).toBeUndefined()
    expect(parseAxisAuto(axis("<c:auto/>"))).toBeUndefined()
    expect(parseAxisAuto(axis('<c:auto val="on"/>'))).toBeUndefined()
    expect(parseAxisAuto(axis(""))).toBeUndefined()
    expect(parseAxisHidden(axis("<c:delete/>"))).toBeUndefined()
    expect(parseAxisHidden(axis('<c:delete val="on"/>'))).toBeUndefined()
    expect(parseAxisHidden(axis(""))).toBeUndefined()
  })
})

describe("parseAxisScale", () => {
  it("reads min/max/logBase from <c:scaling> but the units from the axis itself", () => {
    // CT_ValAx places `<c:majorUnit>` / `<c:minorUnit>` as siblings of
    // `<c:scaling>`, not inside it.
    const scale = parseAxisScale(
      axis(
        '<c:scaling><c:min val="0"/><c:max val="100"/><c:logBase val="10"/></c:scaling>' +
          '<c:majorUnit val="20"/><c:minorUnit val="5"/>',
      ),
    )
    expect(scale).toEqual({ min: 0, max: 100, logBase: 10, majorUnit: 20, minorUnit: 5 })
  })

  it("drops non-positive tick units and surfaces nothing for an orientation-only scaling", () => {
    expect(parseAxisScale(axis('<c:majorUnit val="0"/><c:minorUnit val="-1"/>'))).toBeUndefined()
    expect(
      parseAxisScale(axis('<c:scaling><c:orientation val="minMax"/></c:scaling>')),
    ).toBeUndefined()
    expect(parseAxisScale(axis(""))).toBeUndefined()
  })
})

describe("parseAxisNumberFormat", () => {
  it("requires a non-empty formatCode and reads sourceLinked as xsd:boolean", () => {
    expect(parseAxisNumberFormat(axis('<c:numFmt formatCode="0.00%" sourceLinked="1"/>'))).toEqual({
      formatCode: "0.00%",
      sourceLinked: true,
    })
    expect(
      parseAxisNumberFormat(axis('<c:numFmt formatCode="General" sourceLinked="0"/>')),
    ).toEqual({ formatCode: "General" })
    expect(parseAxisNumberFormat(axis('<c:numFmt formatCode=""/>'))).toBeUndefined()
    expect(parseAxisNumberFormat(axis("<c:numFmt/>"))).toBeUndefined()
    expect(parseAxisNumberFormat(axis(""))).toBeUndefined()
  })
})

describe("parseAxisGridlines", () => {
  it("flips a flag on the mere presence of the element, empty body included", () => {
    expect(parseAxisGridlines(axis("<c:majorGridlines/>"))).toEqual({ major: true })
    expect(parseAxisGridlines(axis("<c:minorGridlines><c:spPr/></c:minorGridlines>"))).toEqual({
      minor: true,
    })
    expect(parseAxisGridlines(axis("<c:majorGridlines/><c:minorGridlines/>"))).toEqual({
      major: true,
      minor: true,
    })
    // Never surface `{ major: false, minor: false }` — it would
    // round-trip into a redundant write.
    expect(parseAxisGridlines(axis(""))).toBeUndefined()
  })
})

describe("parseAxisCrossBetween", () => {
  it("admits only the two ST_CrossBetween tokens", () => {
    expect(parseAxisCrossBetween(axis('<c:crossBetween val="midCat"/>'))).toBe("midCat")
    expect(parseAxisCrossBetween(axis('<c:crossBetween val=" between "/>'))).toBe("between")
    expect(parseAxisCrossBetween(axis('<c:crossBetween val="onTick"/>'))).toBeUndefined()
    expect(parseAxisCrossBetween(axis("<c:crossBetween/>"))).toBeUndefined()
    expect(parseAxisCrossBetween(axis(""))).toBeUndefined()
  })
})

describe("parseAxisCrosses", () => {
  it("prefers an explicit <c:crossesAt> over the <c:crosses> enum", () => {
    // CT_ValAx puts the two in an `xsd:choice`; a template that pins
    // both resolves to the literal coordinate.
    expect(parseAxisCrosses(axis('<c:crossesAt val="-5.5"/><c:crosses val="max"/>'))).toEqual({
      crossesAt: -5.5,
    })
  })

  it("falls back to <c:crosses> when the coordinate is missing or unusable", () => {
    expect(parseAxisCrosses(axis('<c:crossesAt/><c:crosses val="max"/>'))).toEqual({
      crosses: "max",
    })
    expect(parseAxisCrosses(axis('<c:crossesAt val="  "/><c:crosses val="min"/>'))).toEqual({
      crosses: "min",
    })
    expect(parseAxisCrosses(axis('<c:crossesAt val="edge"/><c:crosses val="min"/>'))).toEqual({
      crosses: "min",
    })
  })

  it("surfaces nothing when neither child is usable", () => {
    expect(parseAxisCrosses(axis(""))).toEqual({})
    expect(parseAxisCrosses(axis("<c:crosses/>"))).toEqual({})
    expect(parseAxisCrosses(axis('<c:crosses val="sideways"/>'))).toEqual({})
    // `autoZero` is the OOXML default, so it collapses.
    expect(parseAxisCrosses(axis('<c:crosses val="autoZero"/>'))).toEqual({})
  })
})

describe("parseAxisDispUnits", () => {
  it("prefers <c:custUnit> over <c:builtInUnit> when a template declares both", () => {
    // The OOXML choice forbids both; a corrupt template that pins both
    // resolves the same way in the reader and the writer.
    expect(
      parseAxisDispUnits(
        axis('<c:dispUnits><c:custUnit val="500"/><c:builtInUnit val="millions"/></c:dispUnits>'),
      ),
    ).toEqual({ custUnit: 500 })
  })

  it("falls back to <c:builtInUnit> when the custom divisor is unusable", () => {
    expect(
      parseAxisDispUnits(
        axis('<c:dispUnits><c:custUnit val="0"/><c:builtInUnit val="thousands"/></c:dispUnits>'),
      ),
    ).toEqual({ unit: "thousands" })
    expect(
      parseAxisDispUnits(
        axis('<c:dispUnits><c:custUnit/><c:builtInUnit val=" billions "/></c:dispUnits>'),
      ),
    ).toEqual({ unit: "billions" })
  })

  it("drops the block when neither child resolves", () => {
    expect(parseAxisDispUnits(axis("<c:dispUnits/>"))).toBeUndefined()
    expect(
      parseAxisDispUnits(axis('<c:dispUnits><c:builtInUnit val="gazillions"/></c:dispUnits>')),
    ).toBeUndefined()
    expect(parseAxisDispUnits(axis("<c:dispUnits><c:builtInUnit/></c:dispUnits>"))).toBeUndefined()
    expect(parseAxisDispUnits(axis(""))).toBeUndefined()
  })

  it("joins a rich <c:dispUnitsLbl> label across runs and paragraphs", () => {
    const parsed = parseAxisDispUnits(
      axis(
        '<c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl><c:tx><c:rich>' +
          "<a:p><a:r><a:t>in </a:t></a:r><a:r><a:t>millions</a:t></a:r></a:p>" +
          "<a:p><a:r><a:t>(USD)</a:t></a:r></a:p>" +
          "</c:rich></c:tx></c:dispUnitsLbl></c:dispUnits>",
      ),
    )
    expect(parsed).toEqual({ unit: "millions", showLabel: true, customLabel: "in millions\n(USD)" })
  })

  it("ignores stray text and non-run children while walking the label body", () => {
    // A rich body can legitimately carry whitespace text nodes plus
    // `<a:pPr>` / `<a:rPr>` / `<a:endParaRPr>` siblings; only
    // `<a:p><a:r><a:t>` contributes to the label text.
    const parsed = parseAxisDispUnits(
      axis(
        '<c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl><c:tx><c:rich>' +
          "  <a:bodyPr/>" +
          "<a:p><a:pPr/>  <a:r><a:rPr/>x<a:t>M</a:t></a:r><a:endParaRPr/></a:p>" +
          "</c:rich></c:tx></c:dispUnitsLbl></c:dispUnits>",
      ),
    )
    expect(parsed).toEqual({ unit: "millions", showLabel: true, customLabel: "M" })
  })

  it("ignores element children nested inside an <a:t> run", () => {
    // `<a:t>` is plain text per CT_TextBody; a stray element child
    // contributes nothing to the label rather than stringifying.
    const parsed = parseAxisDispUnits(
      axis(
        '<c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl><c:tx><c:rich>' +
          "<a:p><a:r><a:t>M<a:br/></a:t></a:r></a:p>" +
          "</c:rich></c:tx></c:dispUnitsLbl></c:dispUnits>",
      ),
    )
    expect(parsed).toEqual({ unit: "millions", showLabel: true, customLabel: "M" })
  })

  it("surfaces showLabel with no customLabel for a bare or blank label element", () => {
    const bare = axis('<c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl/></c:dispUnits>')
    expect(parseAxisDispUnits(bare)).toEqual({ unit: "millions", showLabel: true })
    const blank = axis(
      '<c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl><c:tx><c:rich>' +
        "<a:p><a:r><a:t>   </a:t></a:r></a:p><a:p/></c:rich></c:tx></c:dispUnitsLbl></c:dispUnits>",
    )
    expect(parseAxisDispUnits(blank)).toEqual({ unit: "millions", showLabel: true })
    const noRich = axis(
      '<c:dispUnits><c:builtInUnit val="millions"/><c:dispUnitsLbl><c:tx/></c:dispUnitsLbl></c:dispUnits>',
    )
    expect(parseAxisDispUnits(noRich)).toEqual({ unit: "millions", showLabel: true })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// axis — tick-label typography (<c:txPr>)
// ═══════════════════════════════════════════════════════════════════════

describe("axis tick-label typography readers", () => {
  const readers: Array<[string, (e: XmlElement) => unknown]> = [
    ["fontSize", parseAxisLabelFontSize],
    ["bold", parseAxisLabelBold],
    ["italic", parseAxisLabelItalic],
    ["color", parseAxisLabelColor],
    ["underline", parseAxisLabelUnderline],
    ["strike", parseAxisLabelStrike],
    ["fontFamily", parseAxisLabelFontFamily],
  ]

  it("bails at every broken link of the <c:txPr><a:p><a:pPr><a:defRPr> chain", () => {
    for (const [name, read] of readers) {
      for (const depth of [0, 1, 2, 3] as const) {
        expect(read(axis(txPrChain(depth))), `${name}@${depth}`).toBeUndefined()
      }
      expect(read(axis("")), name).toBeUndefined()
    }
  })

  it("converts <a:defRPr sz> from OOXML hundredths onto the 0.5pt grid", () => {
    const size = (sz: string): number | undefined =>
      parseAxisLabelFontSize(axis(txPrChain(3, ` sz="${sz}"`)))
    expect(size("900")).toBe(9)
    expect(size("925")).toBe(9.5)
    expect(size("100")).toBe(1)
    expect(size("40000")).toBe(400)
    expect(size("50")).toBeUndefined()
    expect(size("40100")).toBeUndefined()
    expect(size("  ")).toBeUndefined()
    expect(size("big")).toBeUndefined()
  })

  it("surfaces bold / italic only on the truthy spelling", () => {
    expect(parseAxisLabelBold(axis(txPrChain(3, ' b="true"')))).toBe(true)
    expect(parseAxisLabelBold(axis(txPrChain(3, ' b="0"')))).toBeUndefined()
    expect(parseAxisLabelItalic(axis(txPrChain(3, ' i="1"')))).toBe(true)
    expect(parseAxisLabelItalic(axis(txPrChain(3, ' i="false"')))).toBeUndefined()
  })

  it("surfaces underline / strike only for the single-line UI variants", () => {
    // The writer emits `u="sng"` / `strike="sngStrike"` only, so
    // reporting `dbl` / `dblStrike` as true would silently downgrade
    // the choice to a single line on re-emit.
    expect(parseAxisLabelUnderline(axis(txPrChain(3, ' u="sng"')))).toBe(true)
    expect(parseAxisLabelUnderline(axis(txPrChain(3, ' u="dbl"')))).toBeUndefined()
    expect(parseAxisLabelUnderline(axis(txPrChain(3, ' u="none"')))).toBeUndefined()
    expect(parseAxisLabelStrike(axis(txPrChain(3, ' strike="sngStrike"')))).toBe(true)
    expect(parseAxisLabelStrike(axis(txPrChain(3, ' strike="dblStrike"')))).toBeUndefined()
    expect(parseAxisLabelStrike(axis(txPrChain(3, ' strike="noStrike"')))).toBeUndefined()
  })

  it("trims the tick-label <a:latin typeface> and drops a blank one", () => {
    expect(
      parseAxisLabelFontFamily(axis(txPrChain(3, "", '<a:latin typeface=" Consolas "/>'))),
    ).toBe("Consolas")
    expect(
      parseAxisLabelFontFamily(axis(txPrChain(3, "", '<a:latin typeface="  "/>'))),
    ).toBeUndefined()
    expect(parseAxisLabelFontFamily(axis(txPrChain(3, "", "<a:latin/>")))).toBeUndefined()
    expect(parseAxisLabelFontFamily(axis(txPrChain(3)))).toBeUndefined()
  })

  it("lifts an <a:schemeClr> tick-label color when no literal sRGB is present", () => {
    expect(
      parseAxisLabelColor(
        axis(txPrChain(3, "", '<a:solidFill><a:schemeClr val="lt2"/></a:solidFill>')),
      ),
    ).toEqual({ theme: "lt2" })
    expect(
      parseAxisLabelColor(axis(txPrChain(3, "", "<a:solidFill><a:hslClr/></a:solidFill>"))),
    ).toBeUndefined()
    expect(parseAxisLabelColor(axis(txPrChain(3, "", "<a:noFill/>")))).toBeUndefined()
  })

  it("converts <a:bodyPr rot> from 60000ths of a degree and clamps to -90..90", () => {
    const rot = (r: string): number | undefined =>
      parseAxisLabelRotation(axis(`<c:txPr><a:bodyPr rot="${r}"/></c:txPr>`))
    expect(rot("-2700000")).toBe(-45)
    expect(rot("2700000")).toBe(45)
    // Beyond the UI band Excel snaps to the endpoint rather than wrap.
    expect(rot("9000000")).toBe(90)
    expect(rot("-9000000")).toBe(-90)
    expect(rot("0")).toBeUndefined()
    expect(rot("  ")).toBeUndefined()
    expect(rot("tilted")).toBeUndefined()
    expect(parseAxisLabelRotation(axis("<c:txPr><a:bodyPr/></c:txPr>"))).toBeUndefined()
    expect(parseAxisLabelRotation(axis("<c:txPr/>"))).toBeUndefined()
    expect(parseAxisLabelRotation(axis(""))).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// axis — <c:title> readers
// ═══════════════════════════════════════════════════════════════════════

describe("axis title readers", () => {
  const readers: Array<[string, (e: XmlElement) => unknown, 3 | 5]> = [
    ["rotation", parseAxisTitleRotation, 3],
    ["fontSize", parseAxisTitleFontSize, 5],
    ["bold", parseAxisTitleBold, 5],
    ["italic", parseAxisTitleItalic, 5],
    ["color", parseAxisTitleColor, 5],
    ["strike", parseAxisTitleStrike, 5],
    ["underline", parseAxisTitleUnderline, 5],
    ["fontFamily", parseAxisTitleFontFamily, 5],
  ]

  it("bails at every broken link of the <c:title><c:tx><c:rich>... chain", () => {
    for (const [name, read, deepest] of readers) {
      for (let depth = 0; depth < deepest; depth++) {
        expect(read(axis(titleChain(depth as 0))), `${name}@${depth}`).toBeUndefined()
      }
      expect(read(axis("")), name).toBeUndefined()
    }
  })

  it("returns undefined for a <c:strRef> axis title with no <c:rich> body", () => {
    // A formula-bound axis title has no `<a:p>` slot for typography.
    const strRef = axis(
      "<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>" +
        '<c:pt idx="0"><c:v>Quarter</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>',
    )
    expect(parseAxisTitle(strRef)).toBe("Quarter")
    expect(parseAxisTitleFontSize(strRef)).toBeUndefined()
    expect(parseAxisTitleBold(strRef)).toBeUndefined()
  })

  it("reads the title text out of a rich body and out of a string cache", () => {
    expect(
      parseAxisTitle(
        axis(
          "<c:title><c:tx><c:rich><a:p><a:r><a:t> Revenue </a:t></a:r></a:p></c:rich></c:tx></c:title>",
        ),
      ),
    ).toBe("Revenue")
    // A blank rich body is indistinguishable from no title at all.
    expect(
      parseAxisTitle(
        axis(
          "<c:title><c:tx><c:rich><a:p><a:r><a:t>  </a:t></a:r></a:p></c:rich></c:tx></c:title>",
        ),
      ),
    ).toBeUndefined()
    expect(
      parseAxisTitle(axis("<c:title><c:tx><c:strRef><c:strCache/></c:strRef></c:tx></c:title>")),
    ).toBeUndefined()
    // A cache entry with no `<c:v>` (or a blank one) keeps scanning
    // rather than reporting the empty string as the title.
    expect(
      parseAxisTitle(
        axis(
          "<c:title><c:tx><c:strRef><c:strCache>" +
            '<c:ptCount val="2"/><c:pt idx="0"/><c:pt idx="1"><c:v>  </c:v></c:pt>' +
            "</c:strCache></c:strRef></c:tx></c:title>",
        ),
      ),
    ).toBeUndefined()
    expect(parseAxisTitle(axis("<c:title><c:tx/></c:title>"))).toBeUndefined()
    // A `<c:title>` with no `<c:tx>` at all (Excel writes this when the
    // title is styled but its text is auto-derived) has nothing to read.
    expect(parseAxisTitle(axis("<c:title/>"))).toBeUndefined()
    expect(parseAxisTitle(axis(""))).toBeUndefined()
  })

  it("clamps the axis-title rotation to the -90..90 band", () => {
    const rot = (r: string): number | undefined =>
      parseAxisTitleRotation(
        axis(`<c:title><c:tx><c:rich><a:bodyPr rot="${r}"/></c:rich></c:tx></c:title>`),
      )
    expect(rot("-5400000")).toBe(-90)
    expect(rot("-16200000")).toBe(-90)
    expect(rot("16200000")).toBe(90)
    expect(rot("0")).toBeUndefined()
    expect(rot("   ")).toBeUndefined()
    expect(rot("sideways")).toBeUndefined()
    expect(
      parseAxisTitleRotation(axis("<c:title><c:tx><c:rich><a:bodyPr/></c:rich></c:tx></c:title>")),
    ).toBeUndefined()
  })

  it("surfaces the axis-title underline / strike only for the single-line variants", () => {
    expect(parseAxisTitleUnderline(axis(titleChain(5, ' u="sng"')))).toBe(true)
    expect(parseAxisTitleUnderline(axis(titleChain(5, ' u="dbl"')))).toBeUndefined()
    expect(parseAxisTitleStrike(axis(titleChain(5, ' strike="sngStrike"')))).toBe(true)
    expect(parseAxisTitleStrike(axis(titleChain(5, ' strike="dblStrike"')))).toBeUndefined()
  })

  it("lifts an <a:schemeClr> axis-title color when no literal sRGB is present", () => {
    expect(
      parseAxisTitleColor(
        axis(titleChain(5, "", '<a:solidFill><a:schemeClr val="accent5"/></a:solidFill>')),
      ),
    ).toEqual({ theme: "accent5" })
    expect(
      parseAxisTitleColor(
        axis(titleChain(5, "", '<a:solidFill><a:srgbClr val="0f0f0f"/></a:solidFill>')),
      ),
    ).toBe("0F0F0F")
    expect(
      parseAxisTitleColor(axis(titleChain(5, "", "<a:solidFill><a:sysClr/></a:solidFill>"))),
    ).toBeUndefined()
    expect(parseAxisTitleColor(axis(titleChain(5)))).toBeUndefined()
  })

  it("trims the <a:latin typeface> and drops a blank one", () => {
    expect(
      parseAxisTitleFontFamily(axis(titleChain(5, "", '<a:latin typeface=" Cambria "/>'))),
    ).toBe("Cambria")
    expect(
      parseAxisTitleFontFamily(axis(titleChain(5, "", '<a:latin typeface="  "/>'))),
    ).toBeUndefined()
    expect(parseAxisTitleFontFamily(axis(titleChain(5, "", "<a:latin/>")))).toBeUndefined()
    expect(parseAxisTitleFontFamily(axis(titleChain(5)))).toBeUndefined()
  })

  it("reads <c:title><c:overlay> as an xsd:boolean, collapsing the default", () => {
    const ov = (attr: string): boolean | undefined =>
      parseAxisTitleOverlay(axis(`<c:title><c:overlay ${attr}/></c:title>`))
    expect(ov('val="1"')).toBe(true)
    expect(ov('val="true"')).toBe(true)
    expect(ov('val="0"')).toBeUndefined()
    expect(ov('val="false"')).toBeUndefined()
    expect(ov('val="maybe"')).toBeUndefined()
    expect(ov("")).toBeUndefined()
    expect(parseAxisTitleOverlay(axis("<c:title/>"))).toBeUndefined()
    expect(parseAxisTitleOverlay(axis(""))).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// axis — writer-side normalizers
// ═══════════════════════════════════════════════════════════════════════

describe("normalizeAxisSkip / normalizeAxisNumberFormat", () => {
  it("rounds a skip then drops it outside 1..32767 and at the default 1", () => {
    expect(normalizeAxisSkip(2.4)).toBe(2)
    expect(normalizeAxisSkip(32767)).toBe(32767)
    expect(normalizeAxisSkip(1)).toBeUndefined()
    expect(normalizeAxisSkip(0)).toBeUndefined()
    expect(normalizeAxisSkip(40000)).toBeUndefined()
    expect(normalizeAxisSkip(Number.NaN)).toBeUndefined()
    expect(normalizeAxisSkip(undefined)).toBeUndefined()
  })

  it("collapses a number format whose formatCode is empty or non-string", () => {
    // Excel rejects `<c:numFmt formatCode=""/>` outright.
    expect(normalizeAxisNumberFormat({ formatCode: "0.0", sourceLinked: true })).toEqual({
      formatCode: "0.0",
      sourceLinked: true,
    })
    expect(normalizeAxisNumberFormat({ formatCode: "" })).toBeUndefined()
    expect(normalizeAxisNumberFormat({ formatCode: 5 as never })).toBeUndefined()
    expect(normalizeAxisNumberFormat(undefined)).toBeUndefined()
    expect(normalizeAxisNumberFormat({ formatCode: "0", sourceLinked: false })).toEqual({
      formatCode: "0",
    })
  })
})

describe("normalizeAxisDispUnits / normalizeDispUnits / buildAxisDispUnits", () => {
  for (const [label, normalize] of [
    ["writer", normalizeAxisDispUnits],
    ["clone", normalizeDispUnits],
  ] as const) {
    it(`accepts the bare ST_BuiltInUnit string form (${label})`, () => {
      expect(normalize("millions")).toEqual({ unit: "millions" })
      expect(normalize("gazillions" as never)).toBeUndefined()
    })

    it(`drops a record with no usable child (${label})`, () => {
      // A bare `<c:dispUnits/>` shell fails Excel's strict validator.
      expect(normalize({})).toBeUndefined()
      expect(normalize({ unit: "zillions" as never, custUnit: 0 })).toBeUndefined()
      expect(normalize({ custUnit: Number.NaN })).toBeUndefined()
      expect(normalize(undefined)).toBeUndefined()
      expect(normalize(null as never)).toBeUndefined()
    })

    it(`keeps both children so a clone can append a custUnit override (${label})`, () => {
      expect(
        normalize({ unit: "thousands", custUnit: 250, showLabel: true, customLabel: "  k  " }),
      ).toEqual({ unit: "thousands", custUnit: 250, showLabel: true, customLabel: "k" })
      expect(normalize({ unit: "thousands", customLabel: "   " })).toEqual({ unit: "thousands" })
    })
  }

  it("emits <c:custUnit> in preference to <c:builtInUnit> when both survive", () => {
    const xml = buildAxisDispUnits({ unit: "millions", custUnit: 500 })
    expect(xml).toEqual(['<c:dispUnits><c:custUnit val="500"/></c:dispUnits>'])
    expect(buildAxisDispUnits({ unit: "millions" })).toEqual([
      '<c:dispUnits><c:builtInUnit val="millions"/></c:dispUnits>',
    ])
  })

  it("refuses to ship a bare <c:dispUnits> shell", () => {
    expect(buildAxisDispUnits(undefined)).toEqual([])
    expect(buildAxisDispUnits({} as ChartAxisDispUnits)).toEqual([])
  })

  it("emits a rich <c:dispUnitsLbl> only when a custom label survives trimming", () => {
    const rich = buildAxisDispUnits({ unit: "millions", customLabel: "R&D <m>" })
    expect(rich.join("")).toContain("<a:t>R&amp;D &lt;m&gt;</a:t>")
    // showLabel alone gets Excel's automatic annotation, no rich body.
    expect(buildAxisDispUnits({ unit: "millions", showLabel: true }).join("")).toContain(
      "<c:dispUnitsLbl/>",
    )
    expect(
      buildAxisDispUnits({ unit: "millions", showLabel: true, customLabel: "  " }).join(""),
    ).toContain("<c:dispUnitsLbl/>")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// axis — clone-side apply*Override resolvers
// ═══════════════════════════════════════════════════════════════════════

describe("axis clone overrides", () => {
  it("applySkipOverride validates both the inherited and the replacing value", () => {
    expect(applySkipOverride(3, undefined)).toBe(3)
    expect(applySkipOverride(2.6, undefined)).toBe(3)
    expect(applySkipOverride(1, undefined)).toBeUndefined()
    expect(applySkipOverride(99999, undefined)).toBeUndefined()
    expect(applySkipOverride(Number.NaN, undefined)).toBeUndefined()
    expect(applySkipOverride(3, null)).toBeUndefined()
    expect(applySkipOverride(3, 5)).toBe(5)
    expect(applySkipOverride(3, 1)).toBeUndefined()
    expect(applySkipOverride(3, "5" as never)).toBeUndefined()
  })

  it("applyLblOffsetOverride collapses the OOXML default 100 on both sides", () => {
    expect(applyLblOffsetOverride(250, undefined)).toBe(250)
    expect(applyLblOffsetOverride(100, undefined)).toBeUndefined()
    expect(applyLblOffsetOverride(2000, undefined)).toBeUndefined()
    expect(applyLblOffsetOverride(Number.NaN, undefined)).toBeUndefined()
    expect(applyLblOffsetOverride(250, null)).toBeUndefined()
    expect(applyLblOffsetOverride(250, 100)).toBeUndefined()
    expect(applyLblOffsetOverride(250, -5)).toBeUndefined()
    expect(applyLblOffsetOverride(250, 0)).toBe(0)
    expect(applyLblOffsetOverride(250, "0" as never)).toBeUndefined()
  })

  it("applyLblAlgnOverride collapses the OOXML default `ctr` on both sides", () => {
    expect(applyLblAlgnOverride("l", undefined)).toBe("l")
    expect(applyLblAlgnOverride("ctr", undefined)).toBeUndefined()
    expect(applyLblAlgnOverride("just" as never, undefined)).toBeUndefined()
    expect(applyLblAlgnOverride("l", null)).toBeUndefined()
    expect(applyLblAlgnOverride("l", "r")).toBe("r")
    expect(applyLblAlgnOverride("l", "ctr")).toBeUndefined()
    expect(applyLblAlgnOverride("l", "just" as never)).toBeUndefined()
  })

  it("applyNoMultiLvlLbl / applyAuto / applyHidden each keep only their non-default state", () => {
    expect(applyNoMultiLvlLblOverride(true, undefined)).toBe(true)
    expect(applyNoMultiLvlLblOverride(false, undefined)).toBeUndefined()
    expect(applyNoMultiLvlLblOverride(true, null)).toBeUndefined()
    expect(applyNoMultiLvlLblOverride(false, true)).toBe(true)
    expect(applyNoMultiLvlLblOverride(true, false)).toBeUndefined()
    // `<c:auto>` defaults to true, so it is the inverse of the other two.
    expect(applyAutoOverride(false, undefined)).toBe(false)
    expect(applyAutoOverride(true, undefined)).toBeUndefined()
    expect(applyAutoOverride(false, null)).toBeUndefined()
    expect(applyAutoOverride(true, false)).toBe(false)
    expect(applyAutoOverride(false, true)).toBeUndefined()
    expect(applyHiddenOverride(true, undefined)).toBe(true)
    expect(applyHiddenOverride(false, undefined)).toBeUndefined()
    expect(applyHiddenOverride(true, null)).toBeUndefined()
    expect(applyHiddenOverride(false, true)).toBe(true)
    expect(applyHiddenOverride(true, false)).toBeUndefined()
  })

  it("applyGridlinesOverride drops an all-false record on both sides", () => {
    const both: ChartAxisGridlines = { major: true, minor: true }
    expect(applyGridlinesOverride(both, undefined)).toEqual(both)
    expect(applyGridlinesOverride({ major: false, minor: false }, undefined)).toBeUndefined()
    expect(applyGridlinesOverride(undefined, undefined)).toBeUndefined()
    expect(applyGridlinesOverride(both, null)).toBeUndefined()
    expect(applyGridlinesOverride(both, { minor: true })).toEqual({ minor: true })
    expect(applyGridlinesOverride(both, {})).toBeUndefined()
  })

  it("applyScaleOverride replaces wholesale rather than merging field by field", () => {
    const source: ChartAxisScale = { min: 0, majorUnit: 20 }
    expect(applyScaleOverride(source, undefined)).toEqual(source)
    expect(applyScaleOverride(undefined, undefined)).toBeUndefined()
    expect(applyScaleOverride(source, null)).toBeUndefined()
    // `{ min: 0 }` + `{ max: 100 }` is `{ max: 100 }`, not the union.
    expect(applyScaleOverride(source, { max: 100 })).toEqual({ max: 100 })
  })

  it("cloneScale keeps only finite bounds and positive tick units", () => {
    expect(
      cloneScale({
        min: Number.NaN,
        max: Number.POSITIVE_INFINITY,
        majorUnit: 0,
        minorUnit: -1,
        logBase: Number.NaN,
      }),
    ).toBeUndefined()
    expect(cloneScale({ min: -5, max: 5, majorUnit: 1, minorUnit: 0.5, logBase: 2 })).toEqual({
      min: -5,
      max: 5,
      majorUnit: 1,
      minorUnit: 0.5,
      logBase: 2,
    })
  })

  it("applyNumberFormatOverride requires a usable formatCode on both sides", () => {
    const fmt: ChartAxisNumberFormat = { formatCode: "0.0%", sourceLinked: true }
    expect(applyNumberFormatOverride(fmt, undefined)).toEqual(fmt)
    expect(applyNumberFormatOverride({ formatCode: "" }, undefined)).toBeUndefined()
    expect(applyNumberFormatOverride(undefined, undefined)).toBeUndefined()
    expect(applyNumberFormatOverride(fmt, null)).toBeUndefined()
    expect(applyNumberFormatOverride(fmt, { formatCode: "" })).toBeUndefined()
    expect(applyNumberFormatOverride(fmt, { formatCode: "General" })).toEqual({
      formatCode: "General",
    })
    expect(applyNumberFormatOverride(fmt, { formatCode: "0", sourceLinked: true })).toEqual({
      formatCode: "0",
      sourceLinked: true,
    })
  })

  it("applyTickMark / applyTickLblPos reject tokens outside their enums", () => {
    expect(applyTickMarkOverride("cross", undefined)).toBe("cross")
    expect(applyTickMarkOverride("both" as never, undefined)).toBeUndefined()
    expect(applyTickMarkOverride(undefined, undefined)).toBeUndefined()
    expect(applyTickMarkOverride("cross", null)).toBeUndefined()
    expect(applyTickMarkOverride("cross", "in")).toBe("in")
    expect(applyTickMarkOverride("cross", "both" as never)).toBeUndefined()
    expect(applyTickLblPosOverride("high", undefined)).toBe("high")
    expect(applyTickLblPosOverride("outside" as never, undefined)).toBeUndefined()
    expect(applyTickLblPosOverride(undefined, undefined)).toBeUndefined()
    expect(applyTickLblPosOverride("high", null)).toBeUndefined()
    expect(applyTickLblPosOverride("high", "low")).toBe("low")
    expect(applyTickLblPosOverride("high", "outside" as never)).toBeUndefined()
  })

  it("the tick-label typography overrides share one inherit / drop / replace grammar", () => {
    const flags = [
      applyLabelBoldOverride,
      applyLabelItalicOverride,
      applyLabelUnderlineOverride,
      applyLabelStrikeOverride,
    ]
    for (const apply of flags) {
      expect(apply(true, undefined)).toBe(true)
      expect(apply(false, undefined)).toBe(false)
      expect(apply(undefined, undefined)).toBeUndefined()
      expect(apply(true, null)).toBeUndefined()
      expect(apply(true, false)).toBe(false)
      // A typed escape from an untyped caller must not reach the writer.
      expect(apply(true, 1 as never)).toBeUndefined()
      expect(apply("yes" as never, undefined)).toBeUndefined()
    }
  })

  it("the axis-title flag overrides share the same grammar as the tick-label ones", () => {
    const flags = [
      applyAxisTitleBoldOverride,
      applyAxisTitleItalicOverride,
      applyAxisTitleStrikeOverride,
      applyAxisTitleUnderlineOverride,
      applyAxisTitleOverlayOverride,
    ]
    for (const apply of flags) {
      expect(apply(true, undefined)).toBe(true)
      expect(apply(false, undefined)).toBe(false)
      expect(apply(undefined, undefined)).toBeUndefined()
      expect(apply(true, null)).toBeUndefined()
      expect(apply(true, false)).toBe(false)
      expect(apply(false, true)).toBe(true)
      expect(apply(true, 1 as never)).toBeUndefined()
      expect(apply("yes" as never, undefined)).toBeUndefined()
    }
  })

  it("applyLabelFontSize / Color / FontFamily normalize both sides", () => {
    expect(applyLabelFontSizeOverride(12.25, undefined)).toBe(12.5)
    expect(applyLabelFontSizeOverride(12, null)).toBeUndefined()
    // Out-of-band sizes drop rather than clamp: the writer would emit a
    // token Excel rejects, and a silent clamp would mask the mistake.
    expect(applyLabelFontSizeOverride(12, 9999)).toBeUndefined()
    expect(applyLabelFontSizeOverride(12, 400)).toBe(400)
    expect(applyLabelColorOverride("#abcdef", undefined)).toBe("ABCDEF")
    expect(applyLabelColorOverride("ABCDEF", null)).toBeUndefined()
    expect(applyLabelColorOverride("ABCDEF", "nothex")).toBeUndefined()
    expect(applyLabelFontFamilyOverride("  Arial  ", undefined)).toBe("Arial")
    expect(applyLabelFontFamilyOverride("Arial", null)).toBeUndefined()
    expect(applyLabelFontFamilyOverride("Arial", "   ")).toBeUndefined()
    expect(applyLabelFontFamilyOverride("Arial", "Verdana")).toBe("Verdana")
  })

  it("applyCrossesOverride keeps the two crossing knobs independent", () => {
    // `<c:crosses>` and `<c:crossesAt>` are separate CT_ValAx children,
    // so an override may replace one and inherit the other.
    expect(applyCrossesOverride({ crosses: "max" }, {})).toEqual({ crosses: "max" })
    expect(applyCrossesOverride({ crosses: "autoZero" }, {})).toEqual({})
    expect(applyCrossesOverride({ crosses: "sideways" as never }, {})).toEqual({})
    expect(applyCrossesOverride({ crosses: "max" }, { crosses: null })).toEqual({})
    expect(applyCrossesOverride({ crosses: "max" }, { crosses: "min" })).toEqual({
      crosses: "min",
    })
    expect(applyCrossesOverride({ crosses: "max" }, { crosses: "autoZero" })).toEqual({})
    expect(applyCrossesOverride({ crossesAt: 5 }, {})).toEqual({ crossesAt: 5 })
    expect(applyCrossesOverride({ crossesAt: Number.NaN }, {})).toEqual({})
    expect(applyCrossesOverride({ crossesAt: 5 }, { crossesAt: null })).toEqual({})
    expect(applyCrossesOverride({ crossesAt: 5 }, { crossesAt: 0 })).toEqual({ crossesAt: 0 })
    expect(applyCrossesOverride({ crossesAt: 5 }, { crossesAt: Number.NaN })).toEqual({})
  })

  it("applyDispUnits / applyCrossBetween normalize both sides", () => {
    expect(applyDispUnitsOverride({ unit: "millions" }, undefined)).toEqual({ unit: "millions" })
    expect(applyDispUnitsOverride({ unit: "millions" }, null)).toBeUndefined()
    expect(applyDispUnitsOverride({ unit: "millions" }, "thousands")).toEqual({ unit: "thousands" })
    expect(applyDispUnitsOverride({ unit: "millions" }, {})).toBeUndefined()
    expect(applyCrossBetweenOverride("midCat", undefined)).toBe("midCat")
    expect(applyCrossBetweenOverride("onTick" as never, undefined)).toBeUndefined()
    expect(applyCrossBetweenOverride(undefined, undefined)).toBeUndefined()
    expect(applyCrossBetweenOverride("midCat", null)).toBeUndefined()
    expect(applyCrossBetweenOverride("midCat", "between")).toBe("between")
    expect(applyCrossBetweenOverride("midCat", "onTick" as never)).toBeUndefined()
  })
})

describe("resolveAxes", () => {
  it("carries every parsed X-axis knob through an override-free clone", () => {
    // `cloneChart` with no `axes` override must reproduce the source
    // axis verbatim — this is the path that silently drops a field
    // when a new knob is added to the reader but not to the resolver.
    const resolved = resolveAxes(
      {
        x: {
          scale: { min: 0, max: 100, majorUnit: 25 },
          numberFormat: { formatCode: "0.0", sourceLinked: true },
          majorTickMark: "cross",
          minorTickMark: "in",
          crosses: "max",
        },
      },
      undefined,
      "bar",
    )
    expect(resolved?.x).toMatchObject({
      scale: { min: 0, max: 100, majorUnit: 25 },
      numberFormat: { formatCode: "0.0", sourceLinked: true },
      majorTickMark: "cross",
      minorTickMark: "in",
      crosses: "max",
    })
  })

  it("returns undefined when neither axis carries anything to emit", () => {
    expect(resolveAxes(undefined, undefined, "bar")).toBeUndefined()
  })
})

describe("parseAxisInfo", () => {
  it("threads the axis-title border cap and compound onto the parsed record", () => {
    // `<c:catAx><c:title><c:spPr><a:ln cap=".." cmpd=".."/>` is the only
    // route these two knobs reach `ChartAxisInfo`.
    const info = parseAxisInfo(
      el(
        "catAx",
        '<c:title><c:spPr><a:ln cap="rnd" cmpd="thickThin"/></c:spPr>' +
          "<c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>",
      ),
      "between",
    )
    expect(info).toMatchObject({
      title: "Quarter",
      axisTitleBorderCap: "rnd",
      axisTitleBorderCompound: "thickThin",
    })
  })

  it("returns undefined for an axis that pins nothing the writer models", () => {
    expect(parseAxisInfo(el("catAx", ""), "between")).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// legend — <c:legend> (CT_Legend, §21.2.2.114)
// ═══════════════════════════════════════════════════════════════════════

describe("legend readers on a chart with no <c:legend>", () => {
  // Every legend reader is called unconditionally by `parseChart`, so
  // each one has to survive a chart that never declared a legend.
  const readers: Array<[string, (e: XmlElement) => unknown]> = [
    ["position", parseLegend],
    ["overlay", parseLegendOverlay],
    ["entries", parseLegendEntries],
    ["fontSize", parseLegendFontSize],
    ["bold", parseLegendBold],
    ["italic", parseLegendItalic],
    ["underline", parseLegendUnderline],
    ["strikethrough", parseLegendStrikethrough],
    ["fontColor", parseLegendFontColor],
    ["fontFamily", parseLegendFontFamily],
    ["layout", parseLegendLayout],
    ["fillColor", parseLegendFillColor],
    ["borderColor", parseLegendBorderColor],
    ["borderWidth", parseLegendBorderWidth],
    ["borderDash", parseLegendBorderDash],
    ["borderCap", parseLegendBorderCap],
    ["borderCompound", parseLegendBorderCompound],
  ]

  it("returns undefined from every reader", () => {
    const bare = chartEl("<c:plotArea/>")
    for (const [name, read] of readers) {
      expect(read(bare), name).toBeUndefined()
    }
  })
})

describe("parseLegend", () => {
  it("maps every ST_LegendPos token onto the writer-side vocabulary", () => {
    const pos = (v: string): unknown =>
      parseLegend(chartEl(`<c:legend><c:legendPos val="${v}"/></c:legend>`))
    expect(pos("t")).toBe("top")
    expect(pos("b")).toBe("bottom")
    expect(pos("l")).toBe("left")
    expect(pos("r")).toBe("right")
    expect(pos("tr")).toBe("topRight")
    expect(pos("middle")).toBeUndefined()
  })

  it("falls back to Excel's default `right` when <c:legendPos> is absent or bare", () => {
    expect(parseLegend(chartEl("<c:legend/>"))).toBe("right")
    expect(parseLegend(chartEl("<c:legend><c:legendPos/></c:legend>"))).toBe("right")
  })

  it('reports false for the canonical <c:delete val="1"/> hidden marker', () => {
    expect(parseLegend(chartEl('<c:legend><c:delete val="1"/></c:legend>'))).toBe(false)
    expect(parseLegend(chartEl('<c:legend><c:delete val="true"/></c:legend>'))).toBe(false)
    // `<c:delete val="0"/>` means the legend renders — fall through.
    expect(
      parseLegend(chartEl('<c:legend><c:delete val="0"/><c:legendPos val="b"/></c:legend>')),
    ).toBe("bottom")
  })
})

describe("parseLegendOverlay", () => {
  it("surfaces only the explicit truthy spelling", () => {
    const ov = (attr: string): boolean | undefined =>
      parseLegendOverlay(chartEl(`<c:legend><c:overlay ${attr}/></c:legend>`))
    expect(ov('val="1"')).toBe(true)
    expect(ov('val="true"')).toBe(true)
    // The OOXML default collapses so absence and `val="0"` round-trip
    // identically through cloneChart.
    expect(ov('val="0"')).toBeUndefined()
    expect(ov('val="false"')).toBeUndefined()
    expect(ov('val="on"')).toBeUndefined()
    expect(ov("")).toBeUndefined()
    expect(parseLegendOverlay(chartEl("<c:legend/>"))).toBeUndefined()
  })
})

describe("parseLegendEntries", () => {
  const legend = (inner: string): XmlElement => chartEl(`<c:legend>${inner}</c:legend>`)

  it("requires a non-negative <c:idx> selector on every entry", () => {
    expect(parseLegendEntries(legend("<c:legendEntry/>"))).toBeUndefined()
    expect(parseLegendEntries(legend("<c:legendEntry><c:idx/></c:legendEntry>"))).toBeUndefined()
    expect(
      parseLegendEntries(legend('<c:legendEntry><c:idx val="-1"/></c:legendEntry>')),
    ).toBeUndefined()
    expect(
      parseLegendEntries(legend('<c:legendEntry><c:idx val="x"/></c:legendEntry>')),
    ).toBeUndefined()
    expect(parseLegendEntries(legend(""))).toBeUndefined()
  })

  it("keeps the first of duplicate idx values", () => {
    // Reading first-wins pairs with the writer's last-wins dedupe so a
    // clone override that appends an entry still beats the parsed one.
    const entries = parseLegendEntries(
      legend(
        '<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>' +
          '<c:legendEntry><c:idx val="0"/><c:delete val="0"/></c:legendEntry>',
      ),
    )
    expect(entries).toEqual([{ idx: 0, delete: true }])
  })

  it("defaults <c:delete> to false when absent or malformed", () => {
    expect(
      parseLegendEntries(
        legend(
          '<c:legendEntry><c:idx val="1"/></c:legendEntry>' +
            '<c:legendEntry><c:idx val="2"/><c:delete val="junk"/></c:legendEntry>',
        ),
      ),
    ).toEqual([
      { idx: 1, delete: false },
      { idx: 2, delete: false },
    ])
  })

  it("bails at every broken link of the per-entry <c:txPr> chain", () => {
    for (const chain of [
      "",
      "<c:txPr/>",
      "<c:txPr><a:p/></c:txPr>",
      "<c:txPr><a:p><a:pPr/></a:p></c:txPr>",
    ]) {
      const entries = parseLegendEntries(
        legend(`<c:legendEntry><c:idx val="0"/>${chain}</c:legendEntry>`),
      )
      expect(entries, chain).toEqual([{ idx: 0, delete: false }])
    }
  })

  it("reads the per-entry typography off <a:defRPr>", () => {
    const entries = parseLegendEntries(
      legend(
        '<c:legendEntry><c:idx val="3"/><c:txPr><a:p><a:pPr>' +
          '<a:defRPr sz="1100" b="1" i="true" u="sng" strike="sngStrike">' +
          '<a:solidFill><a:srgbClr val="#aabbcc"/></a:solidFill>' +
          '<a:latin typeface="  Georgia  "/></a:defRPr>' +
          "</a:pPr></a:p></c:txPr></c:legendEntry>",
      ),
    )
    expect(entries).toEqual([
      {
        idx: 3,
        delete: false,
        fontSize: 11,
        bold: true,
        italic: true,
        underline: true,
        strikethrough: true,
        color: "AABBCC",
        fontFamily: "Georgia",
      },
    ])
  })

  it("drops per-entry typography the writer cannot re-emit", () => {
    const entries = parseLegendEntries(
      legend(
        '<c:legendEntry><c:idx val="0"/><c:txPr><a:p><a:pPr>' +
          '<a:defRPr sz="50" b="0" i="0" u="dbl" strike="dblStrike">' +
          '<a:solidFill><a:srgbClr val="nothex"/></a:solidFill>' +
          '<a:latin typeface="   "/></a:defRPr>' +
          "</a:pPr></a:p></c:txPr></c:legendEntry>",
      ),
    )
    expect(entries).toEqual([{ idx: 0, delete: false }])
  })

  it("drops a per-entry <a:srgbClr> with no val and an <a:latin> with no typeface", () => {
    const entries = parseLegendEntries(
      legend(
        '<c:legendEntry><c:idx val="0"/><c:txPr><a:p><a:pPr><a:defRPr>' +
          "<a:solidFill><a:srgbClr/></a:solidFill><a:latin/></a:defRPr>" +
          "</a:pPr></a:p></c:txPr></c:legendEntry>",
      ),
    )
    expect(entries).toEqual([{ idx: 0, delete: false }])
  })

  it("drops a per-entry size that is blank or non-numeric", () => {
    for (const sz of ["", "   ", "abc"]) {
      const entries = parseLegendEntries(
        legend(
          `<c:legendEntry><c:idx val="0"/><c:txPr><a:p><a:pPr><a:defRPr sz="${sz}"/>` +
            "</a:pPr></a:p></c:txPr></c:legendEntry>",
        ),
      )
      expect(entries?.[0].fontSize, sz).toBeUndefined()
    }
  })

  it("skips an <a:solidFill> that carries no <a:srgbClr>", () => {
    const entries = parseLegendEntries(
      legend(
        '<c:legendEntry><c:idx val="0"/><c:txPr><a:p><a:pPr><a:defRPr>' +
          '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr>' +
          "</a:pPr></a:p></c:txPr></c:legendEntry>",
      ),
    )
    expect(entries?.[0].color).toBeUndefined()
  })
})

describe("legend typography readers", () => {
  const legendWith = (inner: string): XmlElement => chartEl(`<c:legend>${inner}</c:legend>`)
  const readers: Array<[string, (e: XmlElement) => unknown]> = [
    ["fontSize", parseLegendFontSize],
    ["bold", parseLegendBold],
    ["italic", parseLegendItalic],
    ["underline", parseLegendUnderline],
    ["strikethrough", parseLegendStrikethrough],
    ["fontColor", parseLegendFontColor],
    ["fontFamily", parseLegendFontFamily],
  ]

  it("bails at every broken link of the <c:txPr><a:p><a:pPr><a:defRPr> chain", () => {
    for (const [name, read] of readers) {
      for (const depth of [0, 1, 2, 3] as const) {
        expect(read(legendWith(txPrChain(depth))), `${name}@${depth}`).toBeUndefined()
      }
      expect(read(legendWith("")), name).toBeUndefined()
    }
  })

  it("reads the whole typography block off a single <a:defRPr>", () => {
    const legend = legendWith(
      txPrChain(
        3,
        ' sz="900" b="1" i="true" u="sng" strike="sngStrike"',
        '<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface=" Calibri "/>',
      ),
    )
    expect(parseLegendFontSize(legend)).toBe(9)
    expect(parseLegendBold(legend)).toBe(true)
    expect(parseLegendItalic(legend)).toBe(true)
    expect(parseLegendUnderline(legend)).toBe(true)
    expect(parseLegendStrikethrough(legend)).toBe(true)
    expect(parseLegendFontColor(legend)).toBe("112233")
    expect(parseLegendFontFamily(legend)).toBe("Calibri")
  })

  it("collapses the OOXML defaults and the non-UI variants", () => {
    const legend = legendWith(
      txPrChain(3, ' sz="40100" b="0" i="false" u="dbl" strike="dblStrike"'),
    )
    expect(parseLegendFontSize(legend)).toBeUndefined()
    expect(parseLegendBold(legend)).toBeUndefined()
    expect(parseLegendItalic(legend)).toBeUndefined()
    // Reporting `dbl` as `true` would silently downgrade to a single
    // line on re-emit, since the writer only ever emits `sng`.
    expect(parseLegendUnderline(legend)).toBeUndefined()
    expect(parseLegendStrikethrough(legend)).toBeUndefined()
  })

  it("lifts an <a:schemeClr> legend font color when no literal sRGB is present", () => {
    expect(
      parseLegendFontColor(
        legendWith(txPrChain(3, "", '<a:solidFill><a:schemeClr val="accent4"/></a:solidFill>')),
      ),
    ).toEqual({ theme: "accent4" })
    expect(
      parseLegendFontColor(legendWith(txPrChain(3, "", "<a:solidFill><a:sysClr/></a:solidFill>"))),
    ).toBeUndefined()
    expect(parseLegendFontColor(legendWith(txPrChain(3)))).toBeUndefined()
  })

  it("drops a blank <a:latin typeface>", () => {
    expect(
      parseLegendFontFamily(legendWith(txPrChain(3, "", '<a:latin typeface="  "/>'))),
    ).toBeUndefined()
    expect(parseLegendFontFamily(legendWith(txPrChain(3, "", "<a:latin/>")))).toBeUndefined()
  })
})

describe("legend <c:spPr> readers", () => {
  const legendWith = (inner: string): XmlElement => chartEl(`<c:legend>${inner}</c:legend>`)

  it("reads the fill, stroke, width, dash, cap and compound off one <c:spPr>", () => {
    const legend = legendWith(
      '<c:spPr><a:solidFill><a:srgbClr val="ffffff"/></a:solidFill>' +
        '<a:ln w="19050" cap="rnd" cmpd="dbl"><a:solidFill><a:srgbClr val="333333"/></a:solidFill>' +
        '<a:prstDash val="sysDash"/></a:ln></c:spPr>',
    )
    expect(parseLegendFillColor(legend)).toBe("FFFFFF")
    expect(parseLegendBorderColor(legend)).toBe("333333")
    // 19050 EMU = 1.5 pt at 12 700 EMU per point.
    expect(parseLegendBorderWidth(legend)).toBe(1.5)
    expect(parseLegendBorderDash(legend)).toBe("sysDash")
    expect(parseLegendBorderCap(legend)).toBe("rnd")
    expect(parseLegendBorderCompound(legend)).toBe("dbl")
  })

  it("lifts an <a:schemeClr> legend fill and stroke", () => {
    const legend = legendWith(
      '<c:spPr><a:solidFill><a:schemeClr val="bg2"/></a:solidFill>' +
        '<a:ln><a:solidFill><a:schemeClr val="accent6"><a:shade val="50000"/>' +
        "</a:schemeClr></a:solidFill></a:ln></c:spPr>",
    )
    expect(parseLegendFillColor(legend)).toEqual({ theme: "bg2" })
    expect(parseLegendBorderColor(legend)).toEqual({ theme: "accent6", shade: 50000 })
    // A fill the writer cannot reproduce drops rather than approximate.
    expect(
      parseLegendFillColor(legendWith("<c:spPr><a:solidFill><a:hslClr/></a:solidFill></c:spPr>")),
    ).toBeUndefined()
    expect(
      parseLegendBorderColor(
        legendWith("<c:spPr><a:ln><a:solidFill><a:sysClr/></a:solidFill></a:ln></c:spPr>"),
      ),
    ).toBeUndefined()
  })

  it("drops the OOXML defaults cap=flat, cmpd=sng and dash=solid", () => {
    const legend = legendWith(
      '<c:spPr><a:ln cap="flat" cmpd="sng"><a:prstDash val="solid"/></a:ln></c:spPr>',
    )
    expect(parseLegendBorderCap(legend)).toBeUndefined()
    expect(parseLegendBorderCompound(legend)).toBeUndefined()
    expect(parseLegendBorderDash(legend)).toBeUndefined()
  })

  it("reads <c:legend><c:layout><c:manualLayout> as a 0..1 fraction set", () => {
    expect(
      parseLegendLayout(
        legendWith(
          '<c:layout><c:manualLayout><c:xMode val="edge"/><c:x val="0.7"/>' +
            '<c:y val="0.1"/></c:manualLayout></c:layout>',
        ),
      ),
    ).toEqual({ x: 0.7, y: 0.1 })
    expect(parseLegendLayout(legendWith("<c:layout/>"))).toBeUndefined()
  })
})

describe("resolveLegendPosition", () => {
  it("maps every writer-side position onto its ST_LegendPos token", () => {
    const pos = (legend: SheetChart["legend"]): unknown =>
      resolveLegendPosition({ legend } as SheetChart)
    expect(pos("top")).toBe("t")
    expect(pos("bottom")).toBe("b")
    expect(pos("left")).toBe("l")
    expect(pos("right")).toBe("r")
    expect(pos("topRight")).toBe("tr")
    expect(pos(false)).toBeNull()
  })

  it("defaults scatter to the bottom and every other family to the right", () => {
    // Excel's own new-chart defaults — a scatter legend sits below the
    // plot so the square plot area keeps its aspect ratio.
    expect(resolveLegendPosition({ type: "scatter" } as SheetChart)).toBe("b")
    expect(resolveLegendPosition({ type: "bar" } as SheetChart)).toBe("r")
  })
})

describe("resolveLegendEntries", () => {
  it("dedupes on idx with last-wins and emits in ascending order", () => {
    const entries = resolveLegendEntries({
      legendEntries: [
        { idx: 2, delete: true },
        { idx: 0, delete: false },
        { idx: 2, delete: false },
      ],
    } as SheetChart)
    expect(entries).toEqual([
      { idx: 0, delete: false },
      { idx: 2, delete: false },
    ])
  })

  it("drops entries whose idx cannot land on a real series", () => {
    expect(
      resolveLegendEntries({
        legendEntries: [
          { idx: -1 },
          { idx: 1.5 },
          { idx: Number.NaN },
          { idx: "0" as never },
          null as never,
        ] as ChartLegendEntry[],
      } as SheetChart),
    ).toEqual([])
    expect(resolveLegendEntries({} as SheetChart)).toEqual([])
    expect(resolveLegendEntries({ legendEntries: [] } as unknown as SheetChart)).toEqual([])
  })

  it("keeps a literal false on each per-entry flag so a clone can turn one off", () => {
    expect(
      resolveLegendEntries({
        legendEntries: [
          {
            idx: 0,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            fontSize: 10.25,
            color: "#aabbcc",
            fontFamily: "  Arial  ",
          },
        ],
      } as SheetChart),
    ).toEqual([
      {
        idx: 0,
        delete: false,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10.5,
        color: "AABBCC",
        fontFamily: "Arial",
      },
    ])
  })

  it("drops out-of-band sizes, malformed colours and blank typefaces", () => {
    expect(
      resolveLegendEntries({
        legendEntries: [
          { idx: 0, fontSize: 999, color: "nothex", fontFamily: "   " },
          { idx: 1, fontSize: Number.NaN },
        ],
      } as SheetChart),
    ).toEqual([
      { idx: 0, delete: false },
      { idx: 1, delete: false },
    ])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// title — chart-level <c:title> readers
// ═══════════════════════════════════════════════════════════════════════

describe("chart title readers", () => {
  const readers: Array<[string, (e: XmlElement) => unknown, number]> = [
    ["rotation", parseTitleRotation, 3],
    ["fontSize", parseTitleFontSize, 5],
    ["bold", parseTitleBold, 5],
    ["italic", parseTitleItalic, 5],
    ["color", parseTitleColor, 5],
    ["strike", parseTitleStrike, 5],
    ["underline", parseTitleUnderline, 5],
    ["fontFamily", parseTitleFontFamily, 5],
  ]

  it("bails at every broken link of the <c:title><c:tx><c:rich>... chain", () => {
    for (const [name, read, deepest] of readers) {
      for (let depth = 0; depth < deepest; depth++) {
        expect(read(chartEl(titleChain(depth as 0))), `${name}@${depth}`).toBeUndefined()
      }
      expect(read(chartEl("")), name).toBeUndefined()
    }
  })

  it("returns undefined for a formula-bound title with no <c:rich> body", () => {
    const strRef = chartEl(
      "<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>" +
        '<c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>',
    )
    expect(parseTitle(strRef)).toBe("Sales")
    expect(parseTitleFontSize(strRef)).toBeUndefined()
    expect(parseTitleBold(strRef)).toBeUndefined()
    expect(parseTitleColor(strRef)).toBeUndefined()
  })

  it("keeps scanning the <c:strCache> past entries with no usable <c:v>", () => {
    expect(
      parseTitle(
        chartEl(
          "<c:title><c:tx><c:strRef><c:strCache>" +
            '<c:ptCount val="3"/><c:pt idx="0"/><c:pt idx="1"><c:v> </c:v></c:pt>' +
            '<c:pt idx="2"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>',
        ),
      ),
    ).toBe("Q1")
  })

  it("returns undefined for a title with no <c:tx>, a blank body, or an empty cache", () => {
    expect(parseTitle(chartEl("<c:title/>"))).toBeUndefined()
    expect(
      parseTitle(
        chartEl(
          "<c:title><c:tx><c:rich><a:p><a:r><a:t>  </a:t></a:r></a:p></c:rich></c:tx></c:title>",
        ),
      ),
    ).toBeUndefined()
    expect(
      parseTitle(chartEl("<c:title><c:tx><c:strRef><c:strCache/></c:strRef></c:tx></c:title>")),
    ).toBeUndefined()
    expect(parseTitle(chartEl("<c:title><c:tx><c:strRef/></c:tx></c:title>"))).toBeUndefined()
    // `<c:tx>` must hold one of the two choices; neither means no title.
    expect(parseTitle(chartEl("<c:title><c:tx/></c:title>"))).toBeUndefined()
    expect(parseTitle(chartEl(""))).toBeUndefined()
  })

  it("reads the whole typography block off a single <a:defRPr>", () => {
    const title = chartEl(
      titleChain(
        5,
        ' sz="1600" b="1" i="1" u="sng" strike="sngStrike"',
        '<a:solidFill><a:srgbClr val="ff8800"/></a:solidFill><a:latin typeface=" Cambria "/>',
      ),
    )
    expect(parseTitleFontSize(title)).toBe(16)
    expect(parseTitleBold(title)).toBe(true)
    expect(parseTitleItalic(title)).toBe(true)
    expect(parseTitleUnderline(title)).toBe(true)
    expect(parseTitleStrike(title)).toBe(true)
    expect(parseTitleColor(title)).toBe("FF8800")
    expect(parseTitleFontFamily(title)).toBe("Cambria")
  })

  it("collapses the OOXML defaults and the non-UI dbl variants", () => {
    const title = chartEl(titleChain(5, ' b="0" i="0" u="dbl" strike="dblStrike"'))
    expect(parseTitleBold(title)).toBeUndefined()
    expect(parseTitleItalic(title)).toBeUndefined()
    expect(parseTitleUnderline(title)).toBeUndefined()
    expect(parseTitleStrike(title)).toBeUndefined()
  })

  it("lifts an <a:schemeClr> title color when no literal sRGB is present", () => {
    expect(
      parseTitleColor(
        chartEl(titleChain(5, "", '<a:solidFill><a:schemeClr val="dk1"/></a:solidFill>')),
      ),
    ).toEqual({ theme: "dk1" })
    expect(
      parseTitleColor(chartEl(titleChain(5, "", "<a:solidFill><a:prstClr/></a:solidFill>"))),
    ).toBeUndefined()
    expect(parseTitleColor(chartEl(titleChain(5)))).toBeUndefined()
  })

  it("trims the typeface and drops a blank one", () => {
    expect(
      parseTitleFontFamily(chartEl(titleChain(5, "", '<a:latin typeface="   "/>'))),
    ).toBeUndefined()
    expect(parseTitleFontFamily(chartEl(titleChain(5, "", "<a:latin/>")))).toBeUndefined()
  })

  it("clamps the title rotation to the -90..90 band", () => {
    const rot = (r: string): number | undefined =>
      parseTitleRotation(
        chartEl(`<c:title><c:tx><c:rich><a:bodyPr rot="${r}"/></c:rich></c:tx></c:title>`),
      )
    expect(rot("-2700000")).toBe(-45)
    expect(rot("20000000")).toBe(90)
    expect(rot("-20000000")).toBe(-90)
    expect(rot("0")).toBeUndefined()
    expect(rot("   ")).toBeUndefined()
    expect(rot("up")).toBeUndefined()
    expect(
      parseTitleRotation(chartEl("<c:title><c:tx><c:rich><a:bodyPr/></c:rich></c:tx></c:title>")),
    ).toBeUndefined()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// plotArea — writer entry points
// ═══════════════════════════════════════════════════════════════════════

describe("buildPlotArea", () => {
  it('emits <c:delete val="1"/> on a scatter value axis the caller hid', () => {
    // `axes.y.hidden` is Excel's "Format Axis -> Labels -> None" plus
    // the axis-line removal; the writer flips `<c:delete>` on the
    // matching `<c:valAx>` rather than omitting the element (Excel
    // still needs the axis id to bind the series to an axis pair).
    const xml = buildPlotArea(
      {
        type: "scatter",
        series: [{ values: "Sheet1!$B$2:$B$5" }],
        axes: { y: { hidden: true } },
      } as SheetChart,
      "Sheet1",
    )
    expect(xml).toContain('<c:delete val="1"/>')
  })
})

describe("buildPieChart", () => {
  it("emits only <c:varyColors> for a chart with no series at all", () => {
    // A pie chart paints its first series only; with none there is
    // nothing to emit beyond the required colour-variation toggle.
    expect(buildPieChart({ type: "pie", series: [] } as unknown as SheetChart, "Sheet1")).toBe(
      '<c:pieChart><c:varyColors val="1"/></c:pieChart>',
    )
  })
})
