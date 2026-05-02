import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { writeChart, chartKindElement } from "../src/xlsx/chart-writer";
import { parseChart } from "../src/xlsx/chart-reader";
import { writeDrawing } from "../src/xlsx/drawing-writer";
import type { WriteChartKind, SheetChart, WriteSheet } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

function zipHas(data: Uint8Array, path: string): boolean {
  const zip = new ZipReader(data);
  return zip.has(path);
}

function makeChart(overrides: Partial<SheetChart> = {}): SheetChart {
  return {
    type: "column",
    title: "Test Chart",
    series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
    anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
    ...overrides,
  };
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

// ── writeChart unit tests ────────────────────────────────────────────

describe("writeChart", () => {
  it("produces a valid c:chartSpace document with expected namespaces", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain("<?xml");
    expect(result.chartXml).toContain("c:chartSpace");
    expect(result.chartXml).toContain(
      'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
    );
    expect(result.chartXml).toContain(
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    );
    expect(result.chartXml).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    );
  });

  it("emits an empty rels file alongside each chart", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartRels).toContain("Relationships");
    // No outgoing relationships in Phase 1.
    expect(result.chartRels).not.toContain("<Relationship ");
  });

  it("renders the title when showTitle is unset but a title is provided", () => {
    const result = writeChart(makeChart({ title: "Q1 Revenue" }), "Sheet1");
    expect(result.chartXml).toContain("c:title");
    expect(result.chartXml).toContain("Q1 Revenue");
    expect(result.chartXml).toContain('c:autoTitleDeleted val="0"');
  });

  it("hides the title when showTitle is explicitly false", () => {
    const result = writeChart(makeChart({ title: "X", showTitle: false }), "Sheet1");
    expect(result.chartXml).not.toContain("Q1 Revenue");
    expect(result.chartXml).not.toContain("<c:title>");
    expect(result.chartXml).toContain('c:autoTitleDeleted val="1"');
  });

  it("escapes XML-special characters in the title and series name", () => {
    const result = writeChart(
      makeChart({
        title: '<Sales> & "profits"',
        series: [{ name: "A & B", values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('&lt;Sales&gt; &amp; "profits"');
    expect(result.chartXml).toContain("A &amp; B");
    expect(result.chartXml).not.toContain("<Sales>");
  });

  it("auto-qualifies bare ranges with the owning sheet name", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain("Sheet1!B2:B4");
    expect(result.chartXml).toContain("Sheet1!A2:A4");
  });

  it("preserves ranges already qualified with a sheet", () => {
    const result = writeChart(
      makeChart({
        series: [{ values: "Other!$B$2:$B$4", categories: "Other!$A$2:$A$4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("Other!$B$2:$B$4");
    expect(result.chartXml).not.toContain("Sheet1!Other!");
  });

  it("quotes sheet names containing whitespace or punctuation", () => {
    const result = writeChart(makeChart(), "Q1 Sales");
    expect(result.chartXml).toContain("'Q1 Sales'!B2:B4");
  });

  it("doubles single quotes inside quoted sheet names", () => {
    const result = writeChart(makeChart(), "Bob's Sheet");
    expect(result.chartXml).toContain("'Bob''s Sheet'!B2:B4");
  });

  it("renders bar direction as horizontal for type=bar", () => {
    const result = writeChart(makeChart({ type: "bar" }), "Sheet1");
    expect(result.chartXml).toContain("c:barChart");
    expect(result.chartXml).toContain('c:barDir val="bar"');
  });

  it("renders bar direction as vertical for type=column", () => {
    const result = writeChart(makeChart({ type: "column" }), "Sheet1");
    expect(result.chartXml).toContain("c:barChart");
    expect(result.chartXml).toContain('c:barDir val="col"');
  });

  it("uses overlap=100 for stacked bar charts", () => {
    const result = writeChart(makeChart({ type: "column", barGrouping: "stacked" }), "Sheet1");
    expect(result.chartXml).toContain('c:grouping val="stacked"');
    expect(result.chartXml).toContain('c:overlap val="100"');
    expect(result.chartXml).not.toContain("c:gapWidth");
  });

  it("emits c:lineChart for type=line with smooth=false marker", () => {
    const result = writeChart(makeChart({ type: "line" }), "Sheet1");
    expect(result.chartXml).toContain("c:lineChart");
    expect(result.chartXml).toContain('c:grouping val="standard"');
    expect(result.chartXml).toContain('c:smooth val="0"');
  });

  it("emits c:pieChart with varyColors=1 for type=pie", () => {
    const result = writeChart(makeChart({ type: "pie" }), "Sheet1");
    expect(result.chartXml).toContain("c:pieChart");
    expect(result.chartXml).toContain('c:varyColors val="1"');
    // Pie has no axes
    expect(result.chartXml).not.toContain("c:catAx");
    expect(result.chartXml).not.toContain("c:valAx");
  });

  it("emits c:scatterChart with xVal/yVal references", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:scatterChart");
    expect(result.chartXml).toContain('c:scatterStyle val="lineMarker"');
    expect(result.chartXml).toContain("c:xVal");
    expect(result.chartXml).toContain("c:yVal");
    // Scatter uses two value axes, not a category axis
    expect(result.chartXml).not.toContain("c:catAx");
  });

  it("emits c:areaChart for type=area", () => {
    const result = writeChart(makeChart({ type: "area" }), "Sheet1");
    expect(result.chartXml).toContain("c:areaChart");
  });

  it("emits stacked grouping for type=line with lineGrouping=stacked", () => {
    const result = writeChart(makeChart({ type: "line", lineGrouping: "stacked" }), "Sheet1");
    expect(result.chartXml).toContain("c:lineChart");
    expect(result.chartXml).toContain('c:grouping val="stacked"');
  });

  it("emits percentStacked grouping for type=line with lineGrouping=percentStacked", () => {
    const result = writeChart(
      makeChart({ type: "line", lineGrouping: "percentStacked" }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:grouping val="percentStacked"');
  });

  it("falls back to standard grouping when lineGrouping is unset", () => {
    const result = writeChart(makeChart({ type: "line" }), "Sheet1");
    expect(result.chartXml).toContain('c:grouping val="standard"');
  });

  it("ignores lineGrouping on non-line chart kinds", () => {
    // Setting lineGrouping on a column chart should not affect its
    // grouping element — the column writer reads barGrouping, not
    // lineGrouping.
    const result = writeChart(
      makeChart({ type: "column", lineGrouping: "stacked" } as SheetChart),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:grouping val="clustered"');
    expect(result.chartXml).not.toContain('c:grouping val="stacked"');
  });

  it("emits stacked grouping for type=area with areaGrouping=stacked", () => {
    const result = writeChart(makeChart({ type: "area", areaGrouping: "stacked" }), "Sheet1");
    expect(result.chartXml).toContain("c:areaChart");
    expect(result.chartXml).toContain('c:grouping val="stacked"');
  });

  it("emits percentStacked grouping for type=area with areaGrouping=percentStacked", () => {
    const result = writeChart(
      makeChart({ type: "area", areaGrouping: "percentStacked" }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:grouping val="percentStacked"');
  });

  it("falls back to standard grouping when areaGrouping is unset", () => {
    const result = writeChart(makeChart({ type: "area" }), "Sheet1");
    expect(result.chartXml).toContain('c:grouping val="standard"');
  });

  it("ignores areaGrouping on non-area chart kinds", () => {
    const result = writeChart(
      makeChart({ type: "line", areaGrouping: "stacked" } as SheetChart),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:lineChart");
    // The line writer reads lineGrouping, so the stacked areaGrouping
    // should be ignored and the line falls back to its standard default.
    expect(result.chartXml).toContain('c:grouping val="standard"');
    expect(result.chartXml).not.toContain('c:grouping val="stacked"');
  });

  it("emits a series fill spPr when color is set", () => {
    const result = writeChart(
      makeChart({
        series: [{ name: "S1", values: "B2:B4", color: "1F77B4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:spPr");
    expect(result.chartXml).toContain('a:srgbClr val="1F77B4"');
  });

  it("normalizes hex colors with leading # to uppercase no-#", () => {
    const result = writeChart(
      makeChart({
        series: [{ values: "B2:B4", color: "#abcdef" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('a:srgbClr val="ABCDEF"');
  });

  it("places the legend on the right by default", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain("c:legend");
    expect(result.chartXml).toContain('c:legendPos val="r"');
  });

  it("hides the legend when legend=false", () => {
    const result = writeChart(makeChart({ legend: false }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:legend>");
  });

  it("places the legend at the top when legend='top'", () => {
    const result = writeChart(makeChart({ legend: "top" }), "Sheet1");
    expect(result.chartXml).toContain('c:legendPos val="t"');
  });

  it("renders multiple series with sequential idx/order", () => {
    const result = writeChart(
      makeChart({
        series: [
          { name: "A", values: "B2:B4" },
          { name: "B", values: "C2:C4" },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toMatch(/c:idx val="0"[\s\S]*c:order val="0"/);
    expect(result.chartXml).toMatch(/c:idx val="1"[\s\S]*c:order val="1"/);
  });

  it.each<WriteChartKind>(["bar", "column", "line", "pie", "doughnut", "scatter", "area"])(
    "kind %s parses as well-formed XML",
    (kind) => {
      const result = writeChart(makeChart({ type: kind }), "Sheet1");
      const doc = parseXml(result.chartXml);
      // Document parses without throwing
      expect(doc).toBeTruthy();
    },
  );
});

// ── Doughnut ─────────────────────────────────────────────────────────

describe("writeChart — doughnut", () => {
  it("emits c:doughnutChart with varyColors=1 and no axes", () => {
    const result = writeChart(makeChart({ type: "doughnut" }), "Sheet1");
    expect(result.chartXml).toContain("c:doughnutChart");
    expect(result.chartXml).toContain('c:varyColors val="1"');
    // Doughnut, like pie, has no axes
    expect(result.chartXml).not.toContain("c:catAx");
    expect(result.chartXml).not.toContain("c:valAx");
  });

  it("declares the schema-required holeSize element with the Excel default of 50", () => {
    const result = writeChart(makeChart({ type: "doughnut" }), "Sheet1");
    expect(result.chartXml).toContain('c:holeSize val="50"');
    expect(result.chartXml).toContain('c:firstSliceAng val="0"');
  });

  it("threads an explicit holeSize through to the XML", () => {
    const result = writeChart(makeChart({ type: "doughnut", holeSize: 75 }), "Sheet1");
    expect(result.chartXml).toContain('c:holeSize val="75"');
  });

  it("clamps holeSize to the 10–90 band Excel's UI enforces", () => {
    const lo = writeChart(makeChart({ type: "doughnut", holeSize: 5 }), "Sheet1");
    expect(lo.chartXml).toContain('c:holeSize val="10"');
    const hi = writeChart(makeChart({ type: "doughnut", holeSize: 120 }), "Sheet1");
    expect(hi.chartXml).toContain('c:holeSize val="90"');
  });

  it("rounds non-integer holeSize values", () => {
    const result = writeChart(makeChart({ type: "doughnut", holeSize: 42.7 }), "Sheet1");
    expect(result.chartXml).toContain('c:holeSize val="43"');
  });

  it("falls back to the default when holeSize is NaN or Infinity", () => {
    const nan = writeChart(makeChart({ type: "doughnut", holeSize: NaN }), "Sheet1");
    expect(nan.chartXml).toContain('c:holeSize val="50"');
    const inf = writeChart(
      makeChart({ type: "doughnut", holeSize: Number.POSITIVE_INFINITY }),
      "Sheet1",
    );
    expect(inf.chartXml).toContain('c:holeSize val="50"');
  });

  it("paints every series declared on a doughnut chart (concentric rings)", () => {
    const result = writeChart(
      makeChart({
        type: "doughnut",
        series: [
          { name: "Inner", values: "B2:B4", categories: "A2:A4" },
          { name: "Outer", values: "C2:C4", categories: "A2:A4" },
        ],
      }),
      "Sheet1",
    );
    // Two <c:ser> entries with sequential idx/order.
    expect(result.chartXml).toMatch(/c:idx val="0"[\s\S]*c:order val="0"/);
    expect(result.chartXml).toMatch(/c:idx val="1"[\s\S]*c:order val="1"/);
    expect(result.chartXml).toContain("Inner");
    expect(result.chartXml).toContain("Outer");
  });

  it("omits holeSize on non-doughnut kinds even when the field is set", () => {
    // SheetChart.holeSize is silently ignored for pie / column / line / etc.
    const pie = writeChart(makeChart({ type: "pie", holeSize: 75 }), "Sheet1");
    expect(pie.chartXml).not.toContain("c:holeSize");
    const col = writeChart(makeChart({ type: "column", holeSize: 75 }), "Sheet1");
    expect(col.chartXml).not.toContain("c:holeSize");
  });

  it("places the legend on the right by default for doughnut, matching pie", () => {
    const result = writeChart(makeChart({ type: "doughnut" }), "Sheet1");
    expect(result.chartXml).toContain('c:legendPos val="r"');
  });

  it("ignores the axes block on doughnut charts", () => {
    const result = writeChart(
      makeChart({
        type: "doughnut",
        axes: { x: { title: "Should not render" }, y: { title: "Same" } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:catAx");
    expect(result.chartXml).not.toContain("c:valAx");
    expect(result.chartXml).not.toContain("Should not render");
  });
});

// ── First slice angle ────────────────────────────────────────────────

describe("writeChart — firstSliceAng", () => {
  it("omits <c:firstSliceAng> on a pie chart with no rotation set", () => {
    // The pie writer treats the OOXML default `0` as the absence of
    // the element so untouched pie charts stay byte-clean.
    const result = writeChart(makeChart({ type: "pie" }), "Sheet1");
    expect(result.chartXml).not.toContain("c:firstSliceAng");
  });

  it("emits <c:firstSliceAng> on a pie chart when firstSliceAng is set", () => {
    const result = writeChart(makeChart({ type: "pie", firstSliceAng: 90 }), "Sheet1");
    expect(result.chartXml).toContain('c:firstSliceAng val="90"');
  });

  it("threads an explicit firstSliceAng through to a doughnut chart", () => {
    const result = writeChart(makeChart({ type: "doughnut", firstSliceAng: 270 }), "Sheet1");
    expect(result.chartXml).toContain('c:firstSliceAng val="270"');
  });

  it("falls back to the default 0 on doughnut when firstSliceAng is unset", () => {
    // Doughnut always emits <c:firstSliceAng> — Excel's reference
    // serialization includes it even at the default. Pie elides it.
    const result = writeChart(makeChart({ type: "doughnut" }), "Sheet1");
    expect(result.chartXml).toContain('c:firstSliceAng val="0"');
  });

  it("wraps angles into the 0..360 band by modulo (stays inside CT_FirstSliceAng)", () => {
    // Excel itself normalizes wrap-arounds the same way when the user
    // types e.g. 380 into the chart-formatting pane.
    const wrap = writeChart(makeChart({ type: "pie", firstSliceAng: 380 }), "Sheet1");
    expect(wrap.chartXml).toContain('c:firstSliceAng val="20"');
    const neg = writeChart(makeChart({ type: "pie", firstSliceAng: -90 }), "Sheet1");
    expect(neg.chartXml).toContain('c:firstSliceAng val="270"');
  });

  it("rounds non-integer firstSliceAng values", () => {
    const result = writeChart(makeChart({ type: "pie", firstSliceAng: 47.6 }), "Sheet1");
    expect(result.chartXml).toContain('c:firstSliceAng val="48"');
  });

  it("falls back to the default 0 when firstSliceAng is NaN or Infinity", () => {
    // Pie elides on the default; doughnut still emits 0.
    const pieNan = writeChart(makeChart({ type: "pie", firstSliceAng: NaN }), "Sheet1");
    expect(pieNan.chartXml).not.toContain("c:firstSliceAng");
    const ringNan = writeChart(makeChart({ type: "doughnut", firstSliceAng: NaN }), "Sheet1");
    expect(ringNan.chartXml).toContain('c:firstSliceAng val="0"');
    const ringInf = writeChart(
      makeChart({ type: "doughnut", firstSliceAng: Number.POSITIVE_INFINITY }),
      "Sheet1",
    );
    expect(ringInf.chartXml).toContain('c:firstSliceAng val="0"');
  });

  it("wraps the schema-equivalent 360 down to 0 (omitted on pie)", () => {
    const result = writeChart(makeChart({ type: "pie", firstSliceAng: 360 }), "Sheet1");
    expect(result.chartXml).not.toContain("c:firstSliceAng");
  });

  it("omits firstSliceAng on non-pie / non-doughnut kinds even when the field is set", () => {
    const col = writeChart(makeChart({ type: "column", firstSliceAng: 90 }), "Sheet1");
    expect(col.chartXml).not.toContain("c:firstSliceAng");
    const line = writeChart(makeChart({ type: "line", firstSliceAng: 90 }), "Sheet1");
    expect(line.chartXml).not.toContain("c:firstSliceAng");
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        firstSliceAng: 90,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).not.toContain("c:firstSliceAng");
  });

  it("places <c:firstSliceAng> inside <c:pieChart> (not at chart level)", () => {
    const result = writeChart(makeChart({ type: "pie", firstSliceAng: 90 }), "Sheet1");
    const pieBlock = result.chartXml.match(/<c:pieChart>[\s\S]*?<\/c:pieChart>/);
    expect(pieBlock).not.toBeNull();
    expect(pieBlock![0]).toContain('c:firstSliceAng val="90"');
  });

  it("places <c:firstSliceAng> before <c:holeSize> inside <c:doughnutChart> (OOXML order)", () => {
    const result = writeChart(
      makeChart({ type: "doughnut", firstSliceAng: 90, holeSize: 60 }),
      "Sheet1",
    );
    // CT_DoughnutChart: varyColors, ser*, dLbls?, firstSliceAng?, holeSize?, extLst?
    expect(result.chartXml.indexOf("c:firstSliceAng")).toBeLessThan(
      result.chartXml.indexOf("c:holeSize"),
    );
  });
});

// ── Smooth lines ─────────────────────────────────────────────────────

describe("writeChart — series smooth flag", () => {
  it('emits <c:smooth val="0"/> on a line series with smooth left unset (default)', () => {
    // <c:smooth> is required on CT_LineSer per the schema, so the line
    // writer always emits the element — straight by default.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:smooth val="0"');
  });

  it('emits <c:smooth val="1"/> on a line series when smooth=true', () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", smooth: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:smooth val="1"');
  });

  it("renders smooth per-series independently on a multi-series line chart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          { name: "Curved", values: "B2:B4", smooth: true },
          { name: "Straight", values: "C2:C4" },
          { name: "ExplicitFalse", values: "D2:D4", smooth: false },
        ],
      }),
      "Sheet1",
    );
    // Three <c:smooth> elements, in series order.
    const matches = result.chartXml.match(/c:smooth val="[01]"/g) ?? [];
    expect(matches).toEqual(['c:smooth val="1"', 'c:smooth val="0"', 'c:smooth val="0"']);
  });

  it("omits <c:smooth> on a scatter series with smooth left unset", () => {
    // Scatter's <c:smooth> is optional (CT_ScatterSer). Untouched
    // scatter series stay byte-clean.
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:smooth");
  });

  it('emits <c:smooth val="1"/> on a scatter series when smooth=true', () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4", smooth: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:smooth val="1"');
  });

  it("ignores smooth on chart kinds whose schema rejects <c:smooth>", () => {
    // The OOXML schema places <c:smooth> only on CT_LineSer and
    // CT_ScatterSer. Setting smooth on a bar / column / pie / doughnut
    // / area series must not leak the element into the output.
    const cases: Array<["column" | "bar" | "pie" | "doughnut" | "area"]> = [
      ["column"],
      ["bar"],
      ["pie"],
      ["doughnut"],
      ["area"],
    ];
    for (const [type] of cases) {
      const result = writeChart(
        makeChart({
          type,
          series: [{ values: "B2:B4", categories: "A2:A4", smooth: true }],
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:smooth");
    }
  });

  it("places <c:smooth> as the last child of <c:ser> (OOXML order)", () => {
    // CT_LineSer puts <c:smooth> after <c:val>, which is itself after
    // <c:cat>. The element must land at the tail of the series block so
    // Excel's strict validator does not reject the file.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ name: "Curved", values: "B2:B4", categories: "A2:A4", smooth: true }],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock.indexOf("c:val")).toBeLessThan(serBlock.indexOf("c:smooth"));
    expect(serBlock.indexOf("c:cat")).toBeLessThan(serBlock.indexOf("c:smooth"));
  });
});

// ── Series markers ───────────────────────────────────────────────────

describe("writeChart — series marker", () => {
  it("omits <c:marker> on a line series when marker is not set", () => {
    // The line writer keeps the chart-type-level `<c:marker val="1"/>`
    // toggle (Excel's per-series default) but does not emit a per-series
    // marker block until the caller pins one.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    // The chart-type-level toggle is still present.
    expect(result.chartXml).toContain('c:marker val="1"');
    // No per-series marker element though.
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock).not.toContain("<c:marker>");
  });

  it("emits <c:marker> with <c:symbol> on a line series", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", marker: { symbol: "diamond" } }],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock).toContain("<c:marker>");
    expect(serBlock).toContain('c:symbol val="diamond"');
  });

  it("emits <c:size> inside <c:marker>", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", marker: { symbol: "circle", size: 12 } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:size val="12"');
  });

  it("clamps size into the OOXML 2..72 band", () => {
    const lo = writeChart(
      makeChart({ type: "line", series: [{ values: "B2:B4", marker: { size: 0 } }] }),
      "Sheet1",
    );
    expect(lo.chartXml).toContain('c:size val="2"');
    const hi = writeChart(
      makeChart({ type: "line", series: [{ values: "B2:B4", marker: { size: 999 } }] }),
      "Sheet1",
    );
    expect(hi.chartXml).toContain('c:size val="72"');
  });

  it("rounds non-integer size values", () => {
    const result = writeChart(
      makeChart({ type: "line", series: [{ values: "B2:B4", marker: { size: 7.6 } }] }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:size val="8"');
  });

  it("emits <c:spPr> with <a:solidFill> when marker.fill is set", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", marker: { symbol: "circle", fill: "1F77B4" } }],
      }),
      "Sheet1",
    );
    const markerBlock = result.chartXml.match(/<c:marker>[\s\S]*?<\/c:marker>/)![0];
    expect(markerBlock).toContain("<c:spPr>");
    expect(markerBlock).toContain("<a:solidFill>");
    expect(markerBlock).toContain('a:srgbClr val="1F77B4"');
  });

  it("emits <a:ln> with a solidFill when marker.line is set", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", marker: { symbol: "circle", line: "FF0000" } }],
      }),
      "Sheet1",
    );
    const markerBlock = result.chartXml.match(/<c:marker>[\s\S]*?<\/c:marker>/)![0];
    expect(markerBlock).toContain("<a:ln>");
    expect(markerBlock).toMatch(/<a:ln>[\s\S]*a:srgbClr val="FF0000"/);
  });

  it("strips a leading '#' and uppercases hex color values", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", marker: { fill: "#1f77b4", line: "#aabbcc" } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('a:srgbClr val="1F77B4"');
    expect(result.chartXml).toContain('a:srgbClr val="AABBCC"');
  });

  it("drops malformed hex color values", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", marker: { fill: "not-a-color", symbol: "circle" } }],
      }),
      "Sheet1",
    );
    // Symbol still surfaces, but fill is dropped — no <a:solidFill>
    // for the marker, since the hex was invalid.
    const markerBlock = result.chartXml.match(/<c:marker>[\s\S]*?<\/c:marker>/)![0];
    expect(markerBlock).toContain('c:symbol val="circle"');
    expect(markerBlock).not.toContain("<a:solidFill>");
  });

  it("drops unknown marker symbols rather than emit invalid XML", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        // @ts-expect-error: deliberately pass an out-of-enum symbol.
        series: [{ values: "B2:B4", marker: { symbol: "pentagon", size: 5 } }],
      }),
      "Sheet1",
    );
    // The size still surfaces but the bogus symbol is dropped.
    expect(result.chartXml).toContain('c:size val="5"');
    expect(result.chartXml).not.toContain("c:symbol");
  });

  it("collapses an empty marker block to no <c:marker> at all", () => {
    // No symbol, size, or color → nothing meaningful to write, so the
    // writer omits the element entirely (same shape as if marker was
    // never set on the series).
    const result = writeChart(
      makeChart({ type: "line", series: [{ values: "B2:B4", marker: {} }] }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock).not.toContain("<c:marker>");
  });

  it("emits <c:marker> on a scatter series", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4", marker: { symbol: "x", size: 8 } }],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock).toContain('c:symbol val="x"');
    expect(serBlock).toContain('c:size val="8"');
  });

  it("renders markers per-series independently on a multi-series line chart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          { name: "A", values: "B2:B4", marker: { symbol: "circle", size: 6 } },
          { name: "B", values: "C2:C4", marker: { symbol: "square" } },
          { name: "C", values: "D2:D4" },
        ],
      }),
      "Sheet1",
    );
    const markers = result.chartXml.match(/<c:marker>[\s\S]*?<\/c:marker>/g) ?? [];
    expect(markers).toHaveLength(2);
    expect(markers[0]).toContain('c:symbol val="circle"');
    expect(markers[1]).toContain('c:symbol val="square"');
  });

  it("ignores marker on chart families whose schema rejects <c:marker>", () => {
    // The OOXML schema places <c:marker> on the series only on
    // CT_LineSer and CT_ScatterSer. Setting marker on a bar / column /
    // pie / doughnut / area series must not leak the element into the
    // output.
    for (const type of ["column", "bar", "pie", "doughnut", "area"] as const) {
      const result = writeChart(
        makeChart({
          type,
          series: [{ values: "B2:B4", categories: "A2:A4", marker: { symbol: "circle" } }],
        }),
        "Sheet1",
      );
      // No per-series marker block on these chart families.
      const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
      expect(serBlock).not.toContain("<c:marker>");
    }
  });

  it("places <c:marker> between <c:spPr> and <c:dLbls> inside <c:ser> (OOXML order)", () => {
    // CT_LineSer / CT_ScatterSer order: idx, order, tx, spPr, marker,
    // dPt*, dLbls?, ..., cat?, val?, smooth?. Excel's strict validator
    // rejects markers placed elsewhere.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            values: "B2:B4",
            categories: "A2:A4",
            color: "1F77B4",
            marker: { symbol: "circle" },
            dataLabels: { showValue: true },
          },
        ],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock.indexOf("<c:spPr>")).toBeLessThan(serBlock.indexOf("<c:marker>"));
    expect(serBlock.indexOf("<c:marker>")).toBeLessThan(serBlock.indexOf("<c:dLbls>"));
    expect(serBlock.indexOf("<c:dLbls>")).toBeLessThan(serBlock.indexOf("<c:cat>"));
    expect(serBlock.indexOf("<c:cat>")).toBeLessThan(serBlock.indexOf("<c:val>"));
    expect(serBlock.indexOf("<c:val>")).toBeLessThan(serBlock.indexOf("<c:smooth"));
  });

  it("survives a writeXlsx → parseChart round-trip", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [
            ["Q", "Rev"],
            ["Q1", 100],
            ["Q2", 150],
            ["Q3", 175],
          ],
          charts: [
            {
              type: "line",
              series: [
                {
                  name: "Rev",
                  values: "B2:B4",
                  categories: "A2:A4",
                  marker: { symbol: "diamond", size: 10, fill: "1F77B4", line: "0F3F60" },
                },
              ],
              anchor: { from: { row: 5, col: 0 } },
            },
          ],
        },
      ],
    });
    const chartXml = await extractXml(xlsx, "xl/charts/chart1.xml");
    const reparsed = parseChart(chartXml);
    expect(reparsed?.series?.[0].marker).toEqual({
      symbol: "diamond",
      size: 10,
      fill: "1F77B4",
      line: "0F3F60",
    });
  });
});

// ── Axis titles ──────────────────────────────────────────────────────

describe("writeChart — axis titles", () => {
  it("emits a <c:title> inside <c:catAx> when axes.x.title is set", () => {
    const result = writeChart(
      makeChart({
        axes: { x: { title: "Quarter" } },
      }),
      "Sheet1",
    );
    // The axis title lives inside c:catAx, not at chart level.
    expect(result.chartXml).toContain("<c:catAx>");
    // Either form is fine, but the literal label must be present.
    expect(result.chartXml).toContain("Quarter");
    // catAx must contain the title (between catAx open and its close).
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/);
    expect(catAxBlock).not.toBeNull();
    expect(catAxBlock![0]).toContain("c:title");
    expect(catAxBlock![0]).toContain("Quarter");
  });

  it("emits a <c:title> inside <c:valAx> when axes.y.title is set", () => {
    const result = writeChart(
      makeChart({
        axes: { y: { title: "Revenue (USD)" } },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/);
    expect(valAxBlock).not.toBeNull();
    expect(valAxBlock![0]).toContain("c:title");
    expect(valAxBlock![0]).toContain("Revenue (USD)");
  });

  it("places axis titles after axPos but before crossAx (OOXML order)", () => {
    const result = writeChart(
      makeChart({
        axes: { x: { title: "X" }, y: { title: "Y" } },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const axPosIdx = catAxBlock.indexOf("c:axPos");
    const titleIdx = catAxBlock.indexOf("<c:title>");
    const crossAxIdx = catAxBlock.indexOf("c:crossAx");
    expect(axPosIdx).toBeGreaterThanOrEqual(0);
    expect(titleIdx).toBeGreaterThan(axPosIdx);
    expect(crossAxIdx).toBeGreaterThan(titleIdx);
  });

  it("escapes XML-special characters in axis titles", () => {
    const result = writeChart(
      makeChart({
        axes: { x: { title: 'A & "B" <C>' } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('A &amp; "B" &lt;C&gt;');
    expect(result.chartXml).not.toContain("<C>");
  });

  it("drops empty / whitespace-only axis titles", () => {
    const result = writeChart(
      makeChart({
        axes: { x: { title: "   " }, y: { title: "" } },
      }),
      "Sheet1",
    );
    // No title element should be emitted inside either axis.
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).not.toContain("c:title");
    expect(valAxBlock).not.toContain("c:title");
  });

  it("works for line and area charts (which share the bar axis builder)", () => {
    for (const type of ["line", "area"] as const) {
      const result = writeChart(
        makeChart({ type, axes: { x: { title: "Date" }, y: { title: "Score" } } }),
        "Sheet1",
      );
      expect(result.chartXml).toContain("Date");
      expect(result.chartXml).toContain("Score");
      const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
      const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
      expect(catAxBlock).toContain("c:title");
      expect(valAxBlock).toContain("c:title");
    }
  });

  it("emits scatter axis titles on the X (b) and Y (l) value axes respectively", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { title: "X-Time" }, y: { title: "Y-Mag" } },
      }),
      "Sheet1",
    );
    const valAxBlocks = [...result.chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map(
      (m) => m[0],
    );
    expect(valAxBlocks).toHaveLength(2);
    // First valAx is the X axis (axPos="b"), second is Y (axPos="l").
    expect(valAxBlocks[0]).toContain('c:axPos val="b"');
    expect(valAxBlocks[0]).toContain("X-Time");
    expect(valAxBlocks[1]).toContain('c:axPos val="l"');
    expect(valAxBlocks[1]).toContain("Y-Mag");
  });

  it("skips axes for pie charts even when axes.x is set", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        axes: { x: { title: "Ignored" } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:catAx");
    expect(result.chartXml).not.toContain("c:valAx");
    expect(result.chartXml).not.toContain("Ignored");
  });

  it("renders well-formed XML when both axes are titled", () => {
    const result = writeChart(
      makeChart({ axes: { x: { title: "X" }, y: { title: "Y" } } }),
      "Sheet1",
    );
    const doc = parseXml(result.chartXml);
    expect(doc).toBeTruthy();
  });
});

describe("writeChart — axis gridlines", () => {
  it("emits <c:majorGridlines> inside the value axis when y.gridlines.major is true", () => {
    const result = writeChart(makeChart({ axes: { y: { gridlines: { major: true } } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/);
    expect(valAxBlock).not.toBeNull();
    expect(valAxBlock![0]).toContain("c:majorGridlines");
    // No minor gridlines should slip in.
    expect(valAxBlock![0]).not.toContain("c:minorGridlines");
  });

  it("emits both <c:majorGridlines> and <c:minorGridlines> when both are true", () => {
    const result = writeChart(
      makeChart({ axes: { y: { gridlines: { major: true, minor: true } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain("c:majorGridlines");
    expect(valAxBlock).toContain("c:minorGridlines");
    // Major must precede minor per OOXML schema.
    expect(valAxBlock.indexOf("c:majorGridlines")).toBeLessThan(
      valAxBlock.indexOf("c:minorGridlines"),
    );
  });

  it("places gridlines after axPos but before any axis title (OOXML order)", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: { title: "Revenue", gridlines: { major: true } },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const axPosIdx = valAxBlock.indexOf("c:axPos");
    const gridlinesIdx = valAxBlock.indexOf("c:majorGridlines");
    const titleIdx = valAxBlock.indexOf("<c:title>");
    const crossAxIdx = valAxBlock.indexOf("c:crossAx");
    expect(axPosIdx).toBeGreaterThanOrEqual(0);
    expect(gridlinesIdx).toBeGreaterThan(axPosIdx);
    expect(titleIdx).toBeGreaterThan(gridlinesIdx);
    expect(crossAxIdx).toBeGreaterThan(titleIdx);
  });

  it("emits gridlines on the category axis when x.gridlines is set", () => {
    const result = writeChart(makeChart({ axes: { x: { gridlines: { major: true } } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain("c:majorGridlines");
  });

  it("emits no gridlines when both flags are false or omitted", () => {
    const result = writeChart(
      makeChart({ axes: { y: { gridlines: { major: false, minor: false } } } }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:majorGridlines");
    expect(result.chartXml).not.toContain("c:minorGridlines");
  });

  it("emits gridlines for line and area charts (sharing the bar axis builder)", () => {
    for (const type of ["line", "area"] as const) {
      const result = writeChart(
        makeChart({
          type,
          axes: { y: { gridlines: { major: true } } },
        }),
        "Sheet1",
      );
      const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
      expect(valAxBlock).toContain("c:majorGridlines");
    }
  });

  it("emits scatter gridlines on the X (b) and Y (l) value axes respectively", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: {
          x: { gridlines: { major: true } },
          y: { gridlines: { minor: true } },
        },
      }),
      "Sheet1",
    );
    const valAxBlocks = [...result.chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map(
      (m) => m[0],
    );
    expect(valAxBlocks).toHaveLength(2);
    expect(valAxBlocks[0]).toContain('c:axPos val="b"');
    expect(valAxBlocks[0]).toContain("c:majorGridlines");
    expect(valAxBlocks[0]).not.toContain("c:minorGridlines");
    expect(valAxBlocks[1]).toContain('c:axPos val="l"');
    expect(valAxBlocks[1]).toContain("c:minorGridlines");
    expect(valAxBlocks[1]).not.toContain("c:majorGridlines");
  });

  it("skips gridlines on pie charts (pie has no axes)", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        axes: { y: { gridlines: { major: true, minor: true } } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:majorGridlines");
    expect(result.chartXml).not.toContain("c:minorGridlines");
  });

  it("renders well-formed XML when titles and gridlines coexist", () => {
    const result = writeChart(
      makeChart({
        axes: {
          x: { title: "Quarter", gridlines: { major: true } },
          y: { title: "Revenue", gridlines: { major: true, minor: true } },
        },
      }),
      "Sheet1",
    );
    const doc = parseXml(result.chartXml);
    expect(doc).toBeTruthy();
  });
});

// ── writeChart — axis scale (min/max/majorUnit/minorUnit/logBase) ────

describe("writeChart — axis scale", () => {
  it("emits <c:min> and <c:max> inside <c:scaling> on the value axis", () => {
    const result = writeChart(
      makeChart({ axes: { y: { scale: { min: 0, max: 100 } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const scalingBlock = valAxBlock.match(/<c:scaling>[\s\S]*?<\/c:scaling>/)![0];
    expect(scalingBlock).toContain('<c:max val="100"/>');
    expect(scalingBlock).toContain('<c:min val="0"/>');
    // Spec order: orientation must precede max which precedes min.
    const orientationIdx = scalingBlock.indexOf("c:orientation");
    const maxIdx = scalingBlock.indexOf("c:max");
    const minIdx = scalingBlock.indexOf("c:min");
    expect(orientationIdx).toBeLessThan(maxIdx);
    expect(maxIdx).toBeLessThan(minIdx);
  });

  it("does not pollute the category axis scaling when only y.scale is set", () => {
    const result = writeChart(
      makeChart({ axes: { y: { scale: { min: 0, max: 100 } } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).not.toContain("<c:min");
    expect(catAxBlock).not.toContain("<c:max");
  });

  it("emits <c:majorUnit> and <c:minorUnit> as siblings of crossBetween (after)", () => {
    const result = writeChart(
      makeChart({ axes: { y: { scale: { majorUnit: 25, minorUnit: 5 } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:majorUnit val="25"/>');
    expect(valAxBlock).toContain('<c:minorUnit val="5"/>');
    // Tick units come AFTER crossBetween per CT_ValAx.
    const crossBetweenIdx = valAxBlock.indexOf("c:crossBetween");
    const majorUnitIdx = valAxBlock.indexOf("c:majorUnit");
    const minorUnitIdx = valAxBlock.indexOf("c:minorUnit");
    expect(crossBetweenIdx).toBeGreaterThan(0);
    expect(majorUnitIdx).toBeGreaterThan(crossBetweenIdx);
    expect(minorUnitIdx).toBeGreaterThan(majorUnitIdx);
  });

  it("emits <c:logBase> before <c:orientation> per CT_Scaling order", () => {
    const result = writeChart(makeChart({ axes: { y: { scale: { logBase: 10 } } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const scalingBlock = valAxBlock.match(/<c:scaling>[\s\S]*?<\/c:scaling>/)![0];
    expect(scalingBlock).toContain('<c:logBase val="10"/>');
    const logBaseIdx = scalingBlock.indexOf("c:logBase");
    const orientationIdx = scalingBlock.indexOf("c:orientation");
    expect(logBaseIdx).toBeLessThan(orientationIdx);
  });

  it("drops max when min >= max (degenerate range)", () => {
    const result = writeChart(
      makeChart({ axes: { y: { scale: { min: 10, max: 10 } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const scalingBlock = valAxBlock.match(/<c:scaling>[\s\S]*?<\/c:scaling>/)![0];
    expect(scalingBlock).toContain('<c:min val="10"/>');
    expect(scalingBlock).not.toContain("<c:max");
  });

  it("ignores non-finite, zero, and negative tick spacings", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            scale: {
              majorUnit: Number.NaN,
              minorUnit: 0,
            },
          },
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:majorUnit");
    expect(result.chartXml).not.toContain("c:minorUnit");
  });

  it("ignores log bases outside the spec-allowed 2..1000 band", () => {
    const result = writeChart(makeChart({ axes: { y: { scale: { logBase: 1 } } } }), "Sheet1");
    expect(result.chartXml).not.toContain("c:logBase");
    const result2 = writeChart(makeChart({ axes: { y: { scale: { logBase: 5000 } } } }), "Sheet1");
    expect(result2.chartXml).not.toContain("c:logBase");
  });

  it("emits scale on the scatter X axis when xScale is set", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { scale: { min: 0, max: 50 } } },
      }),
      "Sheet1",
    );
    const valAxBlocks = [...result.chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map(
      (m) => m[0],
    );
    expect(valAxBlocks).toHaveLength(2);
    // First valAx is the X axis (axPos="b").
    expect(valAxBlocks[0]).toContain('c:axPos val="b"');
    expect(valAxBlocks[0]).toContain('<c:max val="50"/>');
    expect(valAxBlocks[0]).toContain('<c:min val="0"/>');
    expect(valAxBlocks[1]).not.toContain('<c:max val="50"/>');
  });

  it("skips scaling extras on pie charts (pie has no axes)", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        axes: { y: { scale: { min: 0, max: 100 } } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:max");
    expect(result.chartXml).not.toContain("<c:majorUnit");
  });

  it("renders well-formed XML when scale extras coexist with title and gridlines", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            title: "Revenue",
            gridlines: { major: true },
            scale: { min: 0, max: 100, majorUnit: 25 },
          },
        },
      }),
      "Sheet1",
    );
    const doc = parseXml(result.chartXml);
    expect(doc).toBeTruthy();
  });
});

// ── writeChart — axis number format ──────────────────────────────────

describe("writeChart — axis number format", () => {
  it("emits <c:numFmt> with the formatCode and sourceLinked=0 by default", () => {
    const result = writeChart(
      makeChart({ axes: { y: { numberFormat: { formatCode: "#,##0" } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('formatCode="#,##0"');
    expect(valAxBlock).toContain('sourceLinked="0"');
  });

  it("emits sourceLinked=1 when explicitly set", () => {
    const result = writeChart(
      makeChart({
        axes: { y: { numberFormat: { formatCode: "0.00%", sourceLinked: true } } },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('formatCode="0.00%"');
    expect(valAxBlock).toContain('sourceLinked="1"');
  });

  it("places <c:numFmt> after the optional <c:title> and before <c:crossAx>", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: { title: "Revenue", numberFormat: { formatCode: "$#,##0" } },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const titleIdx = valAxBlock.indexOf("<c:title>");
    const numFmtIdx = valAxBlock.indexOf("<c:numFmt");
    const crossAxIdx = valAxBlock.indexOf("c:crossAx");
    expect(titleIdx).toBeGreaterThan(0);
    expect(numFmtIdx).toBeGreaterThan(titleIdx);
    expect(crossAxIdx).toBeGreaterThan(numFmtIdx);
  });

  it("omits <c:numFmt> when formatCode is empty", () => {
    const result = writeChart(
      makeChart({ axes: { y: { numberFormat: { formatCode: "" } } } }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:numFmt");
  });

  it("escapes XML-special characters in the formatCode", () => {
    const result = writeChart(
      makeChart({ axes: { y: { numberFormat: { formatCode: '"<x>"&"y"' } } } }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('formatCode="&quot;&lt;x&gt;&quot;&amp;&quot;y&quot;"');
  });

  it("emits a number format on the scatter Y axis", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { y: { numberFormat: { formatCode: "0.00%" } } },
      }),
      "Sheet1",
    );
    const valAxBlocks = [...result.chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map(
      (m) => m[0],
    );
    // Y axis is the second valAx (axPos="l").
    expect(valAxBlocks[1]).toContain('c:axPos val="l"');
    expect(valAxBlocks[1]).toContain('formatCode="0.00%"');
    expect(valAxBlocks[0]).not.toContain('formatCode="0.00%"');
  });

  it("skips number format on pie charts (pie has no axes)", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        axes: { y: { numberFormat: { formatCode: "#,##0" } } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:numFmt");
  });
});

describe("chartKindElement", () => {
  it("maps each chart kind to the matching DrawingML element", () => {
    expect(chartKindElement("bar")).toBe("c:barChart");
    expect(chartKindElement("column")).toBe("c:barChart");
    expect(chartKindElement("line")).toBe("c:lineChart");
    expect(chartKindElement("pie")).toBe("c:pieChart");
    expect(chartKindElement("doughnut")).toBe("c:doughnutChart");
    expect(chartKindElement("scatter")).toBe("c:scatterChart");
    expect(chartKindElement("area")).toBe("c:areaChart");
  });
});

// ── writeDrawing chart anchor tests ──────────────────────────────────

describe("writeDrawing with charts", () => {
  it("emits an xdr:graphicFrame anchor referencing the chart relationship", () => {
    const result = writeDrawing([], 1, undefined, [makeChart()], 1);

    expect(result.drawingXml).toContain("xdr:graphicFrame");
    expect(result.drawingXml).toContain("a:graphicData");
    expect(result.drawingXml).toContain(
      'uri="http://schemas.openxmlformats.org/drawingml/2006/chart"',
    );
    expect(result.drawingXml).toContain("c:chart");
    expect(result.drawingXml).toContain('r:id="rId1"');
  });

  it("registers a chart relationship pointing to ../charts/chart{N}.xml", () => {
    const result = writeDrawing([], 1, undefined, [makeChart()], 7);
    expect(result.drawingRels).toContain('Target="../charts/chart7.xml"');
    expect(result.drawingRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"',
    );
  });

  it("returns one entry per chart with its global index", () => {
    const result = writeDrawing([], 1, undefined, [makeChart(), makeChart()], 5);
    expect(result.charts).toHaveLength(2);
    expect(result.charts[0].globalChartIndex).toBe(5);
    expect(result.charts[1].globalChartIndex).toBe(6);
  });

  it("propagates chart anchor coordinates into xdr:from/xdr:to", () => {
    const result = writeDrawing(
      [],
      1,
      undefined,
      [makeChart({ anchor: { from: { row: 4, col: 2 }, to: { row: 12, col: 9 } } })],
      1,
    );
    expect(result.drawingXml).toContain("<xdr:col>2</xdr:col>");
    expect(result.drawingXml).toContain("<xdr:row>4</xdr:row>");
    expect(result.drawingXml).toContain("<xdr:col>9</xdr:col>");
    expect(result.drawingXml).toContain("<xdr:row>12</xdr:row>");
  });

  it("falls back to a sensible default when 'to' is omitted", () => {
    const result = writeDrawing(
      [],
      1,
      undefined,
      [makeChart({ anchor: { from: { row: 0, col: 0 } } })],
      1,
    );
    // Default footprint is from + (8, 15)
    expect(result.drawingXml).toContain("<xdr:col>8</xdr:col>");
    expect(result.drawingXml).toContain("<xdr:row>15</xdr:row>");
  });

  it("places chart rIds after image rIds in the drawing rels", () => {
    const fakePng = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const result = writeDrawing(
      [
        {
          data: fakePng,
          type: "png",
          anchor: { from: { row: 0, col: 0 } },
        },
      ],
      1,
      undefined,
      [makeChart()],
      1,
    );
    // Image is rId1, chart should be rId2
    expect(result.drawingRels).toContain('Id="rId1"');
    expect(result.drawingRels).toContain('Id="rId2"');
    expect(result.drawingRels).toContain('Target="../media/image1.png"');
    expect(result.drawingRels).toContain('Target="../charts/chart1.xml"');
  });

  it("writes alt text into xdr:cNvPr/@descr when set", () => {
    const result = writeDrawing(
      [],
      1,
      undefined,
      [makeChart({ altText: "Quarterly revenue chart" })],
      1,
    );
    expect(result.drawingXml).toContain('descr="Quarterly revenue chart"');
  });
});

// ── End-to-end writeXlsx tests ───────────────────────────────────────

describe("writeXlsx with charts", () => {
  it("emits xl/charts/chart1.xml for a single bar chart", async () => {
    const sheet: WriteSheet = {
      name: "Sales",
      rows: [
        ["Quarter", "Revenue"],
        ["Q1", 12000],
        ["Q2", 15500],
        ["Q3", 14000],
      ],
      charts: [makeChart({ type: "bar" })],
    };

    const data = await writeXlsx({ sheets: [sheet] });

    expect(zipHas(data, "xl/charts/chart1.xml")).toBe(true);
    expect(zipHas(data, "xl/charts/_rels/chart1.xml.rels")).toBe(true);
    expect(zipHas(data, "xl/drawings/drawing1.xml")).toBe(true);
    expect(zipHas(data, "xl/drawings/_rels/drawing1.xml.rels")).toBe(true);

    const chartXml = await extractXml(data, "xl/charts/chart1.xml");
    expect(chartXml).toContain("c:barChart");
    expect(chartXml).toContain('c:barDir val="bar"');
  });

  it("registers chart parts in [Content_Types].xml", async () => {
    const sheet: WriteSheet = {
      name: "Sales",
      rows: [
        ["A", "B"],
        [1, 2],
      ],
      charts: [makeChart()],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const contentTypes = await extractXml(data, "[Content_Types].xml");

    expect(contentTypes).toContain('PartName="/xl/charts/chart1.xml"');
    expect(contentTypes).toContain(
      'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"',
    );
    expect(contentTypes).toContain('PartName="/xl/drawings/drawing1.xml"');
  });

  it("wires the worksheet to the drawing via <drawing r:id>", async () => {
    const sheet: WriteSheet = {
      name: "Data",
      rows: [
        ["x", "y"],
        [1, 10],
      ],
      charts: [makeChart()],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml");
    expect(sheetXml).toMatch(/<drawing r:id="rId\d+"\s*\/>/);

    const sheetRels = await extractXml(data, "xl/worksheets/_rels/sheet1.xml.rels");
    expect(sheetRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"',
    );
    expect(sheetRels).toContain('Target="../drawings/drawing1.xml"');
  });

  it("supports multiple charts on the same sheet", async () => {
    const sheet: WriteSheet = {
      name: "Dashboard",
      rows: [
        ["Month", "Revenue", "Cost"],
        ["Jan", 100, 60],
        ["Feb", 150, 90],
      ],
      charts: [
        makeChart({
          type: "column",
          title: "Revenue",
          series: [{ values: "B2:B3", categories: "A2:A3" }],
          anchor: { from: { row: 5, col: 0 } },
        }),
        makeChart({
          type: "pie",
          title: "Costs",
          series: [{ values: "C2:C3", categories: "A2:A3" }],
          anchor: { from: { row: 5, col: 8 } },
        }),
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });

    expect(zipHas(data, "xl/charts/chart1.xml")).toBe(true);
    expect(zipHas(data, "xl/charts/chart2.xml")).toBe(true);

    const chart1 = await extractXml(data, "xl/charts/chart1.xml");
    const chart2 = await extractXml(data, "xl/charts/chart2.xml");
    expect(chart1).toContain("c:barChart");
    expect(chart2).toContain("c:pieChart");
  });

  it("assigns unique global chart indices across sheets", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Q1",
        rows: [
          ["A", "B"],
          [1, 2],
        ],
        charts: [makeChart()],
      },
      {
        name: "Q2",
        rows: [
          ["A", "B"],
          [1, 2],
        ],
        charts: [makeChart(), makeChart()],
      },
    ];

    const data = await writeXlsx({ sheets });

    expect(zipHas(data, "xl/charts/chart1.xml")).toBe(true);
    expect(zipHas(data, "xl/charts/chart2.xml")).toBe(true);
    expect(zipHas(data, "xl/charts/chart3.xml")).toBe(true);

    // Sheet 2's drawing rels should point to chart2 and chart3
    const drawing2Rels = await extractXml(data, "xl/drawings/_rels/drawing2.xml.rels");
    expect(drawing2Rels).toContain("../charts/chart2.xml");
    expect(drawing2Rels).toContain("../charts/chart3.xml");
  });

  it("co-exists with images on the same drawing", async () => {
    const fakePng = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]);
    const sheet: WriteSheet = {
      name: "Mixed",
      rows: [
        ["x", "y"],
        [1, 10],
      ],
      images: [
        {
          data: fakePng,
          type: "png",
          anchor: { from: { row: 0, col: 0 } },
        },
      ],
      charts: [makeChart()],
    };

    const data = await writeXlsx({ sheets: [sheet] });

    const drawingRels = await extractXml(data, "xl/drawings/_rels/drawing1.xml.rels");
    expect(drawingRels).toContain("../media/image1.png");
    expect(drawingRels).toContain("../charts/chart1.xml");

    const drawingXml = await extractXml(data, "xl/drawings/drawing1.xml");
    expect(drawingXml).toContain("xdr:pic");
    expect(drawingXml).toContain("xdr:graphicFrame");
  });

  it("auto-qualifies bare ranges with the owning sheet's name", async () => {
    const sheet: WriteSheet = {
      name: "My Sheet",
      rows: [
        ["A", "B"],
        ["x", 1],
      ],
      charts: [
        makeChart({
          series: [{ values: "B2:B2", categories: "A2:A2" }],
        }),
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const chartXml = await extractXml(data, "xl/charts/chart1.xml");
    // Sheet name has a space → must be quoted
    expect(chartXml).toContain("'My Sheet'!B2:B2");
    expect(chartXml).toContain("'My Sheet'!A2:A2");
  });

  it("does not emit chart parts when no charts are declared", async () => {
    const sheet: WriteSheet = {
      name: "NoCharts",
      rows: [["A"], [1]],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    expect(zipHas(data, "xl/charts/chart1.xml")).toBe(false);

    const contentTypes = await extractXml(data, "[Content_Types].xml");
    expect(contentTypes).not.toContain("/xl/charts/");
  });

  it("produces parseable chart XML that round-trips through the SAX parser", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["A", "B", "C"],
        [1, 2, 3],
        [4, 5, 6],
      ],
      charts: [
        makeChart({ type: "line" }),
        makeChart({
          type: "bar",
          barGrouping: "stacked",
          legend: "top",
        }),
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });

    for (const path of ["xl/charts/chart1.xml", "xl/charts/chart2.xml"]) {
      const xml = await extractXml(data, path);
      const doc = parseXml(xml);
      expect(doc).toBeTruthy();
      const chartSpace = findChild(doc, "chartSpace");
      const root = chartSpace ?? doc;
      expect(root).toBeTruthy();
    }
  });

  it("packages a doughnut chart that parseChart can re-read end-to-end", async () => {
    const sheet: WriteSheet = {
      name: "Distribution",
      rows: [
        ["Category", "Share"],
        ["Cloud", 42],
        ["On-prem", 28],
        ["Hybrid", 30],
      ],
      charts: [
        makeChart({
          type: "doughnut",
          title: "Workload Mix",
          holeSize: 60,
          series: [{ name: "Share", values: "B2:B4", categories: "A2:A4", color: "1070CA" }],
        }),
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    expect(zipHas(data, "xl/charts/chart1.xml")).toBe(true);

    const chartXml = await extractXml(data, "xl/charts/chart1.xml");
    expect(chartXml).toContain("c:doughnutChart");
    expect(chartXml).toContain('c:holeSize val="60"');

    const parsed = parseChart(chartXml);
    expect(parsed?.kinds).toEqual(["doughnut"]);
    expect(parsed?.title).toBe("Workload Mix");
    expect(parsed?.holeSize).toBe(60);
    expect(parsed?.seriesCount).toBe(1);
    expect(parsed?.series?.[0]?.name).toBe("Share");
    expect(parsed?.series?.[0]?.color).toBe("1070CA");
  });
});

// ── Data labels ──────────────────────────────────────────────────────

describe("writeChart — data labels", () => {
  it("emits no <c:dLbls> when neither chart nor series declare labels", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:dLbls>");
  });

  it("emits a chart-level <c:dLbls> with showVal=1 when configured", () => {
    const result = writeChart(makeChart({ dataLabels: { showValue: true } }), "Sheet1");
    expect(result.chartXml).toContain("<c:dLbls>");
    expect(result.chartXml).toContain('c:showVal val="1"');
    expect(result.chartXml).toContain('c:showCatName val="0"');
    expect(result.chartXml).toContain('c:showSerName val="0"');
    expect(result.chartXml).toContain('c:showPercent val="0"');
  });

  it("places the chart-level <c:dLbls> after series and before <c:axId>", () => {
    const result = writeChart(makeChart({ dataLabels: { showValue: true } }), "Sheet1");
    const xml = result.chartXml;
    const lastSer = xml.lastIndexOf("</c:ser>");
    const dLbls = xml.indexOf("<c:dLbls>");
    const firstAxId = xml.indexOf("<c:axId");
    expect(lastSer).toBeGreaterThan(0);
    expect(dLbls).toBeGreaterThan(lastSer);
    expect(firstAxId).toBeGreaterThan(dLbls);
  });

  it("emits the position element before the show* toggles", () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, position: "outEnd" } }),
      "Sheet1",
    );
    const xml = result.chartXml;
    const pos = xml.indexOf("<c:dLblPos");
    const showVal = xml.indexOf("<c:showVal");
    expect(pos).toBeGreaterThan(0);
    expect(showVal).toBeGreaterThan(pos);
    expect(xml).toContain('c:dLblPos val="outEnd"');
  });

  it("emits the separator when set", () => {
    const result = writeChart(
      makeChart({
        dataLabels: { showValue: true, showCategoryName: true, separator: " | " },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:separator> | </c:separator>");
  });

  it("escapes XML-special characters in the separator", () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, separator: " <> & " } }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:separator> &lt;&gt; &amp; </c:separator>");
  });

  it("emits a series-level <c:dLbls> when set on a single series", () => {
    const result = writeChart(
      makeChart({
        series: [
          {
            name: "S1",
            values: "B2:B4",
            dataLabels: { showValue: true, position: "outEnd" },
          },
        ],
      }),
      "Sheet1",
    );
    // The series-level block lives inside <c:ser>.
    const xml = result.chartXml;
    const serStart = xml.indexOf("<c:ser>");
    const serEnd = xml.indexOf("</c:ser>");
    const inner = xml.slice(serStart, serEnd);
    expect(inner).toContain("<c:dLbls>");
    expect(inner).toContain('c:showVal val="1"');
  });

  it("places the series <c:dLbls> after <c:spPr> and before <c:cat>/<c:val>", () => {
    const result = writeChart(
      makeChart({
        series: [
          {
            name: "S1",
            values: "B2:B4",
            categories: "A2:A4",
            color: "1F77B4",
            dataLabels: { showValue: true },
          },
        ],
      }),
      "Sheet1",
    );
    const xml = result.chartXml;
    const spPr = xml.indexOf("<c:spPr>");
    const dLbls = xml.indexOf("<c:dLbls>");
    const cat = xml.indexOf("<c:cat>");
    const val = xml.indexOf("<c:val>");
    expect(spPr).toBeGreaterThan(0);
    expect(dLbls).toBeGreaterThan(spPr);
    expect(cat).toBeGreaterThan(dLbls);
    expect(val).toBeGreaterThan(cat);
  });

  it("suppresses a single series with dataLabels=false even when chart-level is on", () => {
    const result = writeChart(
      makeChart({
        dataLabels: { showValue: true },
        series: [
          { name: "Visible", values: "B2:B4" },
          { name: "Hidden", values: "C2:C4", dataLabels: false },
        ],
      }),
      "Sheet1",
    );
    const xml = result.chartXml;
    // Hidden series block is the second <c:ser>...</c:ser>.
    const firstSerEnd = xml.indexOf("</c:ser>");
    const secondSerStart = xml.indexOf("<c:ser>", firstSerEnd);
    const secondSerEnd = xml.indexOf("</c:ser>", secondSerStart);
    const hiddenInner = xml.slice(secondSerStart, secondSerEnd);
    // Excel's "delete this series' labels" idiom: a <c:dLbls> with delete=1.
    expect(hiddenInner).toContain('<c:delete val="1"/>');
  });

  it("supports pie chart with showPercent=true and bestFit position", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        dataLabels: { showPercent: true, position: "bestFit" },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:pieChart>");
    expect(result.chartXml).toContain('c:dLblPos val="bestFit"');
    expect(result.chartXml).toContain('c:showPercent val="1"');
  });

  it("supports line chart with chart-level data labels", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        dataLabels: { showValue: true, position: "t" },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:lineChart>");
    expect(result.chartXml).toContain('c:dLblPos val="t"');
    // Sanity: line chart's marker tag should still come after dLbls.
    const xml = result.chartXml;
    expect(xml.indexOf("<c:dLbls>")).toBeLessThan(xml.indexOf("<c:marker"));
  });
});

// ── Bar/column gapWidth & overlap ────────────────────────────────────

describe("writeChart — gapWidth", () => {
  it("emits <c:gapWidth val='150'/> on a clustered column chart by default", () => {
    // Excel's reference serialization for an unstacked bar/column chart
    // pins gapWidth at 150% of the bar width, so untouched charts stay
    // byte-identical with what Excel writes.
    const result = writeChart(makeChart({ type: "column" }), "Sheet1");
    expect(result.chartXml).toContain('c:gapWidth val="150"');
  });

  it("threads an explicit gapWidth through to the XML", () => {
    const result = writeChart(makeChart({ type: "column", gapWidth: 50 }), "Sheet1");
    expect(result.chartXml).toContain('c:gapWidth val="50"');
    expect(result.chartXml).not.toContain('c:gapWidth val="150"');
  });

  it("clamps gapWidth into the 0..500 band the OOXML schema allows", () => {
    const lo = writeChart(makeChart({ type: "column", gapWidth: -25 }), "Sheet1");
    expect(lo.chartXml).toContain('c:gapWidth val="0"');
    const hi = writeChart(makeChart({ type: "column", gapWidth: 999 }), "Sheet1");
    expect(hi.chartXml).toContain('c:gapWidth val="500"');
  });

  it("rounds non-integer gapWidth values", () => {
    const result = writeChart(makeChart({ type: "column", gapWidth: 175.6 }), "Sheet1");
    expect(result.chartXml).toContain('c:gapWidth val="176"');
  });

  it("falls back to the default when gapWidth is NaN or Infinity", () => {
    const nan = writeChart(makeChart({ type: "column", gapWidth: NaN }), "Sheet1");
    expect(nan.chartXml).toContain('c:gapWidth val="150"');
    const inf = writeChart(
      makeChart({ type: "column", gapWidth: Number.POSITIVE_INFINITY }),
      "Sheet1",
    );
    expect(inf.chartXml).toContain('c:gapWidth val="150"');
  });

  it("emits <c:gapWidth> on a stacked chart only when explicitly set", () => {
    // Stacked charts default to gapWidth omitted (Excel's reference),
    // but pinning a value forces emission.
    const def = writeChart(makeChart({ type: "column", barGrouping: "stacked" }), "Sheet1");
    expect(def.chartXml).not.toContain("c:gapWidth");
    const explicit = writeChart(
      makeChart({ type: "column", barGrouping: "stacked", gapWidth: 75 }),
      "Sheet1",
    );
    expect(explicit.chartXml).toContain('c:gapWidth val="75"');
  });

  it("emits <c:gapWidth> on a horizontal bar chart too", () => {
    const result = writeChart(makeChart({ type: "bar", gapWidth: 200 }), "Sheet1");
    expect(result.chartXml).toContain('c:barDir val="bar"');
    expect(result.chartXml).toContain('c:gapWidth val="200"');
  });

  it("omits gapWidth on non-bar chart kinds even when the field is set", () => {
    // SheetChart.gapWidth is silently ignored for line / pie / area / scatter / doughnut.
    const line = writeChart(makeChart({ type: "line", gapWidth: 75 }), "Sheet1");
    expect(line.chartXml).not.toContain("c:gapWidth");
    const pie = writeChart(makeChart({ type: "pie", gapWidth: 75 }), "Sheet1");
    expect(pie.chartXml).not.toContain("c:gapWidth");
    const area = writeChart(makeChart({ type: "area", gapWidth: 75 }), "Sheet1");
    expect(area.chartXml).not.toContain("c:gapWidth");
  });
});

describe("writeChart — overlap", () => {
  it("emits <c:overlap val='100'/> on a stacked bar chart by default", () => {
    // Stacked bar charts pin overlap at 100% (series fully overlapped)
    // so series stack on top of each other rather than render
    // side-by-side.
    const result = writeChart(makeChart({ type: "column", barGrouping: "stacked" }), "Sheet1");
    expect(result.chartXml).toContain('c:overlap val="100"');
  });

  it("emits <c:overlap val='100'/> on a percentStacked bar chart by default", () => {
    const result = writeChart(
      makeChart({ type: "column", barGrouping: "percentStacked" }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:overlap val="100"');
  });

  it("threads an explicit overlap through to the XML", () => {
    const result = writeChart(makeChart({ type: "column", overlap: -25 }), "Sheet1");
    expect(result.chartXml).toContain('c:overlap val="-25"');
  });

  it("clamps overlap into the -100..100 band the OOXML schema allows", () => {
    const lo = writeChart(makeChart({ type: "column", overlap: -250 }), "Sheet1");
    expect(lo.chartXml).toContain('c:overlap val="-100"');
    const hi = writeChart(makeChart({ type: "column", overlap: 200 }), "Sheet1");
    expect(hi.chartXml).toContain('c:overlap val="100"');
  });

  it("rounds non-integer overlap values", () => {
    const result = writeChart(makeChart({ type: "column", overlap: -33.4 }), "Sheet1");
    expect(result.chartXml).toContain('c:overlap val="-33"');
  });

  it("falls back to the per-grouping default when overlap is NaN or Infinity", () => {
    // Clustered: omitted; stacked: 100.
    const nanClustered = writeChart(makeChart({ type: "column", overlap: NaN }), "Sheet1");
    expect(nanClustered.chartXml).not.toContain("c:overlap");
    const nanStacked = writeChart(
      makeChart({ type: "column", barGrouping: "stacked", overlap: NaN }),
      "Sheet1",
    );
    expect(nanStacked.chartXml).toContain('c:overlap val="100"');
  });

  it("forces overlap emission on a clustered chart when explicitly set", () => {
    // Clustered defaults to no <c:overlap> element; an explicit value
    // overrides that and ships the element through.
    const def = writeChart(makeChart({ type: "column" }), "Sheet1");
    expect(def.chartXml).not.toContain("c:overlap");
    const explicit = writeChart(makeChart({ type: "column", overlap: -50 }), "Sheet1");
    expect(explicit.chartXml).toContain('c:overlap val="-50"');
  });

  it("omits overlap on non-bar chart kinds even when the field is set", () => {
    const line = writeChart(makeChart({ type: "line", overlap: 50 }), "Sheet1");
    expect(line.chartXml).not.toContain("c:overlap");
    const pie = writeChart(makeChart({ type: "pie", overlap: 50 }), "Sheet1");
    expect(pie.chartXml).not.toContain("c:overlap");
  });

  it("places <c:gapWidth> before <c:overlap> inside <c:barChart> (OOXML order)", () => {
    // CT_BarChart sequence: ... dLbls? → gapWidth? → overlap? → serLines* → axId+
    const result = writeChart(makeChart({ type: "column", gapWidth: 50, overlap: -25 }), "Sheet1");
    expect(result.chartXml.indexOf("c:gapWidth")).toBeLessThan(
      result.chartXml.indexOf("c:overlap"),
    );
  });

  it("places <c:gapWidth> / <c:overlap> before <c:axId> inside <c:barChart>", () => {
    const result = writeChart(makeChart({ type: "column", gapWidth: 50, overlap: 25 }), "Sheet1");
    const barBlock = result.chartXml.match(/<c:barChart>[\s\S]*?<\/c:barChart>/);
    expect(barBlock).not.toBeNull();
    expect(barBlock![0].indexOf("c:overlap")).toBeLessThan(barBlock![0].indexOf("c:axId"));
  });
});

// ── Display blanks as ────────────────────────────────────────────────

describe("writeChart — dispBlanksAs", () => {
  it('emits <c:dispBlanksAs val="gap"/> when the field is unset (OOXML default)', () => {
    // The writer always emits the element so the rendered intent is
    // explicit on roundtrip — Excel itself includes it in every
    // reference serialization.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain('c:dispBlanksAs val="gap"');
  });

  it('threads dispBlanksAs="zero" through to <c:chart>', () => {
    const result = writeChart(makeChart({ dispBlanksAs: "zero" }), "Sheet1");
    expect(result.chartXml).toContain('c:dispBlanksAs val="zero"');
  });

  it('threads dispBlanksAs="span" through to <c:chart>', () => {
    const result = writeChart(makeChart({ type: "line", dispBlanksAs: "span" }), "Sheet1");
    expect(result.chartXml).toContain('c:dispBlanksAs val="span"');
  });

  it("falls back to gap on unknown dispBlanksAs values rather than emit one Excel rejects", () => {
    const result = writeChart(
      // @ts-expect-error — testing runtime guard for malformed input
      makeChart({ dispBlanksAs: "bogus" }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:dispBlanksAs val="gap"');
    expect(result.chartXml).not.toContain('c:dispBlanksAs val="bogus"');
  });

  it("places <c:dispBlanksAs> after <c:plotVisOnly> inside <c:chart> (OOXML order)", () => {
    // CT_Chart sequence: ... plotArea, legend?, plotVisOnly?, dispBlanksAs?, ...
    const result = writeChart(makeChart({ dispBlanksAs: "zero" }), "Sheet1");
    expect(result.chartXml.indexOf("c:plotVisOnly")).toBeLessThan(
      result.chartXml.indexOf("c:dispBlanksAs"),
    );
  });

  it("only emits <c:dispBlanksAs> once even on a chart that overrides it", () => {
    // Earlier writers emitted a hardcoded `gap` even when the chart
    // requested a different value. Guard against any regression that
    // would double-emit the element.
    const result = writeChart(makeChart({ dispBlanksAs: "span" }), "Sheet1");
    const occurrences = result.chartXml.match(/c:dispBlanksAs/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads dispBlanksAs through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, dispBlanksAs: "zero" }), "Sheet1");
      expect(result.chartXml).toContain('c:dispBlanksAs val="zero"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dispBlanksAs: "span",
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:dispBlanksAs val="span"');
  });
});
