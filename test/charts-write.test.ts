import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { writeChart, chartKindElement } from "../src/xlsx/chart-writer";
import { parseChart } from "../src/xlsx/chart-reader";
import { writeDrawing } from "../src/xlsx/drawing-writer";
import type { ChartScatterStyle, WriteChartKind, SheetChart, WriteSheet } from "../src/_types";

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

// ── Line stroke (dash + width) ───────────────────────────────────────

describe("writeChart — series line stroke", () => {
  it("emits <a:prstDash> on a line series stroke.dash", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          { name: "Forecast", values: "B2:B4", categories: "A2:A4", stroke: { dash: "dash" } },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<a:prstDash val="dash"/>');
  });

  it("emits <a:ln w=...> in EMU for a line series stroke.width (1 pt = 12 700 EMU)", () => {
    // 2.5 pt → 31 750 EMU.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { width: 2.5 } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('w="31750"');
  });

  it("snaps a stroke.width to the 0.25 pt grid before converting to EMU", () => {
    // 1.13 pt should snap to 1.25 pt → 15 875 EMU (matching what Excel
    // rounds to in its UI).
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { width: 1.13 } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('w="15875"');
  });

  it("clamps stroke.width below 0.25 pt to 0.25 pt", () => {
    // 0.1 pt clamps to 0.25 pt → 3 175 EMU.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { width: 0.1 } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('w="3175"');
  });

  it("clamps stroke.width above 13.5 pt to 13.5 pt", () => {
    // 50 pt clamps to 13.5 pt → 171 450 EMU.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { width: 50 } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('w="171450"');
  });

  it("drops an unknown dash value and emits no <a:prstDash>", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          // @ts-expect-error – exercising the runtime guard
          { values: "B2:B4", categories: "A2:A4", stroke: { dash: "wiggle" } },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("a:prstDash");
  });

  it("drops a non-finite stroke.width and emits no <a:ln w=...>", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { width: Number.NaN } }],
      }),
      "Sheet1",
    );
    // Line writer always emits the chart-type-level <c:marker val="1"/>;
    // the regex below specifically targets the per-series <a:ln> attr.
    expect(result.chartXml).not.toMatch(/<a:ln\s+w="/);
  });

  it("collapses an empty stroke {} to no <c:spPr> for a series without color", () => {
    // Empty stroke + no fill color must not introduce a `<c:spPr>` block —
    // an empty wrapper would override Excel's series-rotation default
    // with no actual styling.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: {} }],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock).not.toContain("<c:spPr>");
  });

  it("layers stroke.dash onto a series with an existing fill color", () => {
    // Series.color emits both <a:solidFill> and a colored <a:ln>; adding
    // stroke.dash should append <a:prstDash> inside the same <a:ln>.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            values: "B2:B4",
            categories: "A2:A4",
            color: "1F77B4",
            stroke: { dash: "dashDot" },
          },
        ],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock).toMatch(/<a:ln[\s\S]*?<a:srgbClr val="1F77B4"\/>[\s\S]*?<a:prstDash/);
  });

  it("emits <a:ln> on a colorless line series when stroke.dash is set", () => {
    // No fill color, but a dash style — the writer must still emit
    // `<c:spPr><a:ln>` to carry the prstDash, otherwise the dash
    // setting silently drops at write time.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { dash: "dot" } }],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    expect(serBlock).toContain("<c:spPr>");
    expect(serBlock).toContain("<a:ln");
    expect(serBlock).toContain('<a:prstDash val="dot"/>');
    // No accidental fill block when only stroke is requested.
    expect(serBlock).not.toContain("<a:solidFill>");
  });

  it("renders stroke per-series independently across a multi-series line chart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          { name: "S1", values: "B2:B4", stroke: { dash: "dash" } },
          { name: "S2", values: "C2:C4" },
          { name: "S3", values: "D2:D4", stroke: { dash: "dot", width: 1.5 } },
        ],
      }),
      "Sheet1",
    );
    const matches = result.chartXml.match(/<a:prstDash val="[^"]+"\/>/g) ?? [];
    expect(matches).toEqual(['<a:prstDash val="dash"/>', '<a:prstDash val="dot"/>']);
  });

  it("emits stroke on a scatter series", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { dash: "lgDash", width: 0.75 } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<a:prstDash val="lgDash"/>');
    // 0.75 pt → 9 525 EMU.
    expect(result.chartXml).toContain('w="9525"');
  });

  it("ignores stroke on chart kinds whose schema does not paint a connecting line", () => {
    // Bar / column / pie / doughnut / area never render a per-series
    // line stroke (each has its own per-data-point border instead). A
    // stroke field on those series must drop at write time.
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
          series: [{ values: "B2:B4", categories: "A2:A4", stroke: { dash: "dash", width: 2 } }],
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("a:prstDash");
      // The stroke width must not leak into a non-line family either.
      expect(result.chartXml).not.toMatch(/<a:ln\s+w="/);
    }
  });

  it("snaps a 9 525-EMU width back to 9 525 EMU on round-trip (idempotent)", () => {
    // Half-EMU drift is the most common round-trip bug — ensure 0.75 pt
    // (Excel default) is byte-stable.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4", stroke: { width: 0.75 } }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('w="9525"');
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

// ── invertIfNegative (per-series flag, bar / column only) ────────────

describe("writeChart — series invertIfNegative flag", () => {
  it("omits <c:invertIfNegative> on a bar series with the flag left unset", () => {
    // Absence of <c:invertIfNegative> matches the OOXML default
    // (`val="0"`); the writer keeps untouched series byte-clean.
    const result = writeChart(
      makeChart({
        type: "column",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:invertIfNegative");
  });

  it('emits <c:invertIfNegative val="1"/> on a column series when invertIfNegative=true', () => {
    const result = writeChart(
      makeChart({
        type: "column",
        series: [{ values: "B2:B4", categories: "A2:A4", invertIfNegative: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:invertIfNegative val="1"');
  });

  it('emits <c:invertIfNegative val="1"/> on a horizontal bar series when invertIfNegative=true', () => {
    const result = writeChart(
      makeChart({
        type: "bar",
        series: [{ values: "B2:B4", categories: "A2:A4", invertIfNegative: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:invertIfNegative val="1"');
  });

  it("renders invertIfNegative per-series independently on a multi-series column chart", () => {
    const result = writeChart(
      makeChart({
        type: "column",
        series: [
          { name: "Inverted", values: "B2:B4", invertIfNegative: true },
          { name: "Default", values: "C2:C4" },
          { name: "ExplicitFalse", values: "D2:D4", invertIfNegative: false },
        ],
      }),
      "Sheet1",
    );
    // Only the first series carries <c:invertIfNegative>. Series with
    // the flag explicitly false collapse to absence (the OOXML default).
    const matches = result.chartXml.match(/c:invertIfNegative val="[01]"/g) ?? [];
    expect(matches).toEqual(['c:invertIfNegative val="1"']);
  });

  it("ignores invertIfNegative on chart kinds whose schema rejects <c:invertIfNegative>", () => {
    // The OOXML schema places <c:invertIfNegative> only on CT_BarSer
    // and CT_Bar3DSer. Setting the flag on a line / pie / doughnut /
    // area / scatter series must not leak the element into the output.
    const cases: Array<["line" | "pie" | "doughnut" | "area" | "scatter"]> = [
      ["line"],
      ["pie"],
      ["doughnut"],
      ["area"],
      ["scatter"],
    ];
    for (const [type] of cases) {
      const result = writeChart(
        makeChart({
          type,
          series: [{ values: "B2:B4", categories: "A2:A4", invertIfNegative: true }],
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:invertIfNegative");
    }
  });

  it("places <c:invertIfNegative> between <c:spPr> and <c:cat>/<c:val> (OOXML order)", () => {
    // CT_BarSer puts <c:invertIfNegative> after <c:spPr> and before
    // <c:dLbls> / <c:cat> / <c:val>. The element must land between the
    // styling block and the data references so Excel's strict validator
    // does not reject the file.
    const result = writeChart(
      makeChart({
        type: "column",
        series: [
          {
            name: "Inverted",
            values: "B2:B4",
            categories: "A2:A4",
            color: "FF0000",
            invertIfNegative: true,
          },
        ],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    const spPrIdx = serBlock.indexOf("c:spPr");
    const invertIdx = serBlock.indexOf("c:invertIfNegative");
    const catIdx = serBlock.indexOf("c:cat");
    const valIdx = serBlock.indexOf("c:val");
    expect(spPrIdx).toBeLessThan(invertIdx);
    expect(invertIdx).toBeLessThan(catIdx);
    expect(invertIdx).toBeLessThan(valIdx);
  });

  it("emits <c:invertIfNegative> alongside other bar-only fields without disturbing them", () => {
    // The barChart-level fields (<c:gapWidth>, <c:overlap>) are
    // independent of the per-series invertIfNegative flag. Both must
    // emit cleanly without interfering.
    const result = writeChart(
      makeChart({
        type: "column",
        gapWidth: 50,
        overlap: -10,
        series: [{ values: "B2:B4", categories: "A2:A4", invertIfNegative: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:invertIfNegative val="1"');
    expect(result.chartXml).toContain('c:gapWidth val="50"');
    expect(result.chartXml).toContain('c:overlap val="-10"');
  });

  it("survives a parseChart round-trip with invertIfNegative preserved", async () => {
    // writeChart → parseChart pulls the flag straight back. Confirms the
    // reader and writer agree on the element and that the value is
    // surfaced on the resulting ChartSeriesInfo.
    const written = writeChart(
      makeChart({
        type: "column",
        series: [
          { name: "Inverted", values: "B2:B4", invertIfNegative: true },
          { name: "Default", values: "C2:C4" },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(written.chartXml);
    expect(parsed?.series).toHaveLength(2);
    expect(parsed?.series?.[0].invertIfNegative).toBe(true);
    expect(parsed?.series?.[1].invertIfNegative).toBeUndefined();
  });
});

// ── explosion (per-series, pie / doughnut only) ─────────────────────

describe("writeChart — series explosion", () => {
  it("omits <c:explosion> on a pie series with the field left unset", () => {
    // Absence of <c:explosion> matches the OOXML default
    // (`val="0"`); the writer keeps untouched series byte-clean.
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:explosion");
  });

  it('emits <c:explosion val="25"/> on a pie series when explosion is set', () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [{ values: "B2:B4", categories: "A2:A4", explosion: 25 }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:explosion val="25"');
  });

  it('emits <c:explosion val="50"/> on a doughnut series when explosion is set', () => {
    const result = writeChart(
      makeChart({
        type: "doughnut",
        series: [{ values: "B2:B4", categories: "A2:A4", explosion: 50 }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:explosion val="50"');
  });

  it("renders explosion per-series independently across a multi-series doughnut chart", () => {
    const result = writeChart(
      makeChart({
        type: "doughnut",
        series: [
          { name: "Exploded", values: "B2:B4", explosion: 30 },
          { name: "Default", values: "C2:C4" },
          { name: "Zero", values: "D2:D4", explosion: 0 },
        ],
      }),
      "Sheet1",
    );
    // Only the first series carries <c:explosion>. Series with the
    // value explicitly 0 collapse to absence (the OOXML default).
    const matches = result.chartXml.match(/c:explosion val="\d+"/g) ?? [];
    expect(matches).toEqual(['c:explosion val="30"']);
  });

  it("ignores explosion on chart kinds whose schema rejects <c:explosion>", () => {
    // The OOXML schema places <c:explosion> only on CT_PieSer (and
    // its EG_PieSer-sharing siblings). Setting the field on a bar /
    // column / line / area / scatter series must not leak the element.
    const cases: Array<["bar" | "column" | "line" | "area" | "scatter"]> = [
      ["bar"],
      ["column"],
      ["line"],
      ["area"],
      ["scatter"],
    ];
    for (const [type] of cases) {
      const result = writeChart(
        makeChart({
          type,
          series: [{ values: "B2:B4", categories: "A2:A4", explosion: 25 }],
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:explosion");
    }
  });

  it("places <c:explosion> between <c:spPr> and <c:cat>/<c:val> (OOXML order)", () => {
    // CT_PieSer puts <c:explosion> after <c:spPr> and before
    // <c:dPt> / <c:dLbls> / <c:cat> / <c:val>. The element must land
    // between the styling block and the data references so Excel's
    // strict validator does not reject the file.
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [
          {
            name: "Exploded",
            values: "B2:B4",
            categories: "A2:A4",
            color: "FF0000",
            explosion: 30,
          },
        ],
      }),
      "Sheet1",
    );
    const serBlock = result.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/)![0];
    const spPrIdx = serBlock.indexOf("c:spPr");
    const explosionIdx = serBlock.indexOf("c:explosion");
    const catIdx = serBlock.indexOf("c:cat");
    const valIdx = serBlock.indexOf("c:val");
    expect(spPrIdx).toBeLessThan(explosionIdx);
    expect(explosionIdx).toBeLessThan(catIdx);
    expect(explosionIdx).toBeLessThan(valIdx);
  });

  it("clamps an explosion value above 400 down to 400", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [{ values: "B2:B4", categories: "A2:A4", explosion: 9999 }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:explosion val="400"');
  });

  it("rounds non-integer explosion values to the nearest integer", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [{ values: "B2:B4", categories: "A2:A4", explosion: 33.6 }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:explosion val="34"');
  });

  it("collapses negative or non-finite explosion values to absence (OOXML default)", () => {
    const cases = [-50, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY];
    for (const value of cases) {
      const result = writeChart(
        makeChart({
          type: "pie",
          series: [{ values: "B2:B4", categories: "A2:A4", explosion: value }],
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:explosion");
    }
  });

  it("emits <c:explosion> alongside the pie-only <c:firstSliceAng> without disturbing it", () => {
    // The pieChart-level <c:firstSliceAng> is independent of the
    // per-series <c:explosion>. Both must emit cleanly without
    // interfering and the chart must still parse back.
    const result = writeChart(
      makeChart({
        type: "pie",
        firstSliceAng: 90,
        series: [{ values: "B2:B4", categories: "A2:A4", explosion: 25 }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:explosion val="25"');
    expect(result.chartXml).toContain('c:firstSliceAng val="90"');
  });

  it("survives a parseChart round-trip with explosion preserved", async () => {
    // writeChart → parseChart pulls the value straight back. Confirms
    // the reader and writer agree on the element and that the value is
    // surfaced on the resulting ChartSeriesInfo.
    const written = writeChart(
      makeChart({
        type: "doughnut",
        series: [
          { name: "Exploded", values: "B2:B4", explosion: 30 },
          { name: "Default", values: "C2:C4" },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(written.chartXml);
    expect(parsed?.series).toHaveLength(2);
    expect(parsed?.series?.[0].explosion).toBe(30);
    expect(parsed?.series?.[1].explosion).toBeUndefined();
  });

  it("threads explosion through an end-to-end writeXlsx round-trip", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [
            ["Region", "Revenue"],
            ["North", 100],
            ["South", 200],
            ["East", 150],
            ["West", 175],
          ],
          charts: [
            {
              type: "pie",
              series: [{ values: "B2:B5", categories: "A2:A5", explosion: 40 }],
              anchor: { from: { row: 6, col: 0 } },
            },
          ],
        },
      ],
    });
    const written = await extractXml(xlsx, "xl/charts/chart1.xml");
    expect(written).toContain('c:explosion val="40"');
    const reparsed = parseChart(written);
    expect(reparsed?.series?.[0].explosion).toBe(40);
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

// ── Vary colors ──────────────────────────────────────────────────────

describe("writeChart — varyColors", () => {
  it('emits <c:varyColors val="0"/> on a column chart by default', () => {
    // Column / bar / line / area / scatter all default to false — each
    // series renders in a single color.
    const result = writeChart(makeChart({ type: "column" }), "Sheet1");
    expect(result.chartXml).toContain('c:varyColors val="0"');
    expect(result.chartXml).not.toContain('c:varyColors val="1"');
  });

  it('emits <c:varyColors val="1"/> on a pie chart by default', () => {
    // Pie / doughnut default to true — each slice paints in its own color.
    const result = writeChart(makeChart({ type: "pie" }), "Sheet1");
    expect(result.chartXml).toContain('c:varyColors val="1"');
    expect(result.chartXml).not.toContain('c:varyColors val="0"');
  });

  it('emits <c:varyColors val="1"/> on a doughnut chart by default', () => {
    const result = writeChart(makeChart({ type: "doughnut" }), "Sheet1");
    expect(result.chartXml).toContain('c:varyColors val="1"');
    expect(result.chartXml).not.toContain('c:varyColors val="0"');
  });

  it("lets varyColors=true flip a column chart to per-point colors", () => {
    const result = writeChart(makeChart({ type: "column", varyColors: true }), "Sheet1");
    expect(result.chartXml).toContain('c:varyColors val="1"');
    expect(result.chartXml).not.toContain('c:varyColors val="0"');
  });

  it("lets varyColors=false collapse a doughnut chart to a single color", () => {
    const result = writeChart(makeChart({ type: "doughnut", varyColors: false }), "Sheet1");
    expect(result.chartXml).toContain('c:varyColors val="0"');
    expect(result.chartXml).not.toContain('c:varyColors val="1"');
  });

  it("lets varyColors=false on a pie chart override the per-family default", () => {
    const result = writeChart(makeChart({ type: "pie", varyColors: false }), "Sheet1");
    expect(result.chartXml).toContain('c:varyColors val="0"');
    expect(result.chartXml).not.toContain('c:varyColors val="1"');
  });

  it("threads varyColors through every authored chart family", () => {
    // Authoring true on every family flips the bar / column / line /
    // area / scatter defaults from 0 to 1 and leaves pie / doughnut at
    // 1. The element appears exactly once on each rendered chart.
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, varyColors: true }), "Sheet1");
      expect(result.chartXml).toContain('c:varyColors val="1"');
      const occurrences = result.chartXml.match(/c:varyColors/g) ?? [];
      expect(occurrences).toHaveLength(1);
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        varyColors: true,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:varyColors val="1"');
  });

  it("places <c:varyColors> after <c:grouping> inside <c:barChart> (OOXML order)", () => {
    // CT_BarChart sequence: barDir → grouping → varyColors → ser*
    const result = writeChart(makeChart({ type: "column" }), "Sheet1");
    expect(result.chartXml.indexOf("c:grouping")).toBeLessThan(
      result.chartXml.indexOf("c:varyColors"),
    );
    expect(result.chartXml.indexOf("c:varyColors")).toBeLessThan(result.chartXml.indexOf("c:ser"));
  });

  it("places <c:varyColors> after <c:scatterStyle> inside <c:scatterChart>", () => {
    // CT_ScatterChart sequence: scatterStyle → varyColors → ser*
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml.indexOf("c:scatterStyle")).toBeLessThan(
      result.chartXml.indexOf("c:varyColors"),
    );
  });

  it("only emits <c:varyColors> once even when the chart pins the field", () => {
    // Guard against any regression that would double-emit the element
    // — both the default emission and a future explicit pass.
    const result = writeChart(makeChart({ type: "column", varyColors: true }), "Sheet1");
    const occurrences = result.chartXml.match(/c:varyColors/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("round-trips a non-default varyColors value through parseChart", () => {
    // A column chart with varyColors=true should re-parse into a Chart
    // whose `varyColors` field is `true` (not collapsed to undefined,
    // since true is not the column-family default).
    const written = writeChart(makeChart({ type: "column", varyColors: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.varyColors).toBe(true);
  });

  it("collapses a defaulted varyColors round-trip back to undefined", () => {
    // A fresh column chart (varyColors omitted) writes `0` and re-parses
    // to undefined — absence and the per-family default round-trip
    // identically through parseChart.
    const written = writeChart(makeChart({ type: "column" }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.varyColors).toBeUndefined();
  });

  it("collapses a defaulted varyColors on pie / doughnut back to undefined", () => {
    const pie = writeChart(makeChart({ type: "pie" }), "Sheet1").chartXml;
    expect(parseChart(pie)?.varyColors).toBeUndefined();
    const dough = writeChart(makeChart({ type: "doughnut" }), "Sheet1").chartXml;
    expect(parseChart(dough)?.varyColors).toBeUndefined();
  });
});

// ── Scatter style ────────────────────────────────────────────────────

describe("writeChart — scatterStyle", () => {
  function makeScatter(overrides: Partial<SheetChart> = {}): SheetChart {
    return makeChart({
      type: "scatter",
      series: [{ values: "B2:B4", categories: "A2:A4" }],
      ...overrides,
    });
  }

  it('emits <c:scatterStyle val="lineMarker"/> on a fresh scatter chart', () => {
    // The writer's default mirrors Excel's chart-picker default —
    // straight lines with markers — even though the OOXML schema
    // default is `"marker"`. Matching Excel's UI default keeps fresh
    // charts visually identical to what the user would draw by hand.
    const result = writeChart(makeScatter(), "Sheet1");
    expect(result.chartXml).toContain('c:scatterStyle val="lineMarker"');
  });

  it("threads an explicit scatterStyle through to the rendered chart", () => {
    const result = writeChart(makeScatter({ scatterStyle: "smooth" }), "Sheet1");
    expect(result.chartXml).toContain('c:scatterStyle val="smooth"');
    expect(result.chartXml).not.toContain('c:scatterStyle val="lineMarker"');
  });

  it("emits every ST_ScatterStyle preset literally when pinned", () => {
    for (const preset of [
      "none",
      "line",
      "lineMarker",
      "marker",
      "smooth",
      "smoothMarker",
    ] as const) {
      const result = writeChart(makeScatter({ scatterStyle: preset }), "Sheet1");
      expect(result.chartXml).toContain(`c:scatterStyle val="${preset}"`);
      // Element appears exactly once on the rendered chart.
      const occurrences = result.chartXml.match(/c:scatterStyle/g) ?? [];
      expect(occurrences).toHaveLength(1);
    }
  });

  it("falls back to the default lineMarker on an unrecognized scatterStyle", () => {
    // Type-cheat with an enum-violating string to exercise the
    // validate-or-default branch — the writer never emits a token
    // Excel's strict validator would reject.
    const result = writeChart(
      makeScatter({ scatterStyle: "bogus" as ChartScatterStyle }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:scatterStyle val="lineMarker"');
    expect(result.chartXml).not.toContain('c:scatterStyle val="bogus"');
  });

  it("ignores scatterStyle on non-scatter chart families", () => {
    // The OOXML schema places <c:scatterStyle> exclusively on
    // <c:scatterChart>; the writer drops the field on every other
    // family rather than emit an element Excel would refuse.
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, scatterStyle: "smooth" }), "Sheet1");
      expect(result.chartXml).not.toContain("c:scatterStyle");
    }
  });

  it("places <c:scatterStyle> as the first child of <c:scatterChart>", () => {
    // CT_ScatterChart sequence: scatterStyle → varyColors → ser*
    const result = writeChart(makeScatter({ scatterStyle: "smoothMarker" }), "Sheet1");
    const styleIdx = result.chartXml.indexOf("c:scatterStyle");
    const varyIdx = result.chartXml.indexOf("c:varyColors");
    const serIdx = result.chartXml.indexOf("c:ser>");
    expect(styleIdx).toBeGreaterThan(-1);
    expect(varyIdx).toBeGreaterThan(styleIdx);
    expect(serIdx).toBeGreaterThan(varyIdx);
  });

  it("round-trips a non-default scatterStyle through parseChart", () => {
    const written = writeChart(makeScatter({ scatterStyle: "smooth" }), "Sheet1").chartXml;
    expect(parseChart(written)?.scatterStyle).toBe("smooth");
  });

  it("round-trips the lineMarker default through parseChart", () => {
    // The writer always emits `lineMarker` by default — re-parsing
    // surfaces it literally because the reader does not collapse the
    // writer's chosen default (only the OOXML schema default `marker`
    // would be a candidate for collapse, but the reader keeps every
    // token literal so a clone preserves the exact preset).
    const written = writeChart(makeScatter(), "Sheet1").chartXml;
    expect(parseChart(written)?.scatterStyle).toBe("lineMarker");
  });
});

// ── writeChart — axis tick marks and tick label position ─────────────

describe("writeChart — axis tick marks and tick label position", () => {
  it("emits <c:majorTickMark> on the value axis when y.majorTickMark is set", () => {
    const result = writeChart(makeChart({ axes: { y: { majorTickMark: "cross" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:majorTickMark val="cross"/>');
  });

  it("emits <c:minorTickMark> on the value axis when y.minorTickMark is set", () => {
    const result = writeChart(makeChart({ axes: { y: { minorTickMark: "out" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:minorTickMark val="out"/>');
  });

  it("emits <c:tickLblPos> on the value axis when y.tickLblPos is set", () => {
    const result = writeChart(makeChart({ axes: { y: { tickLblPos: "low" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:tickLblPos val="low"/>');
  });

  it("omits all three elements when none of the fields are set", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("c:majorTickMark");
    expect(result.chartXml).not.toContain("c:minorTickMark");
    expect(result.chartXml).not.toContain("c:tickLblPos");
  });

  it("places tick rendering after <c:numFmt> but before <c:crossAx> (OOXML order)", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            numberFormat: { formatCode: "$#,##0" },
            majorTickMark: "cross",
            minorTickMark: "in",
            tickLblPos: "low",
          },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const numFmtIdx = valAxBlock.indexOf("<c:numFmt");
    const majorIdx = valAxBlock.indexOf("<c:majorTickMark");
    const minorIdx = valAxBlock.indexOf("<c:minorTickMark");
    const tickLblIdx = valAxBlock.indexOf("<c:tickLblPos");
    const crossAxIdx = valAxBlock.indexOf("c:crossAx");
    expect(numFmtIdx).toBeGreaterThan(0);
    expect(majorIdx).toBeGreaterThan(numFmtIdx);
    expect(minorIdx).toBeGreaterThan(majorIdx);
    expect(tickLblIdx).toBeGreaterThan(minorIdx);
    expect(crossAxIdx).toBeGreaterThan(tickLblIdx);
  });

  it("emits tick rendering on the category axis when x.* is set", () => {
    const result = writeChart(
      makeChart({
        axes: { x: { majorTickMark: "in", tickLblPos: "high" } },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('<c:majorTickMark val="in"/>');
    expect(catAxBlock).toContain('<c:tickLblPos val="high"/>');
    // The value axis should not pick up the X-axis settings.
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).not.toContain("c:majorTickMark");
    expect(valAxBlock).not.toContain("c:tickLblPos");
  });

  it("works for line and area charts (which share the bar axis builder)", () => {
    for (const type of ["line", "area"] as const) {
      const result = writeChart(
        makeChart({ type, axes: { y: { majorTickMark: "cross" } } }),
        "Sheet1",
      );
      const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
      expect(valAxBlock).toContain('<c:majorTickMark val="cross"/>');
    }
  });

  it("emits tick rendering on scatter X (axPos=b) and Y (axPos=l) value axes", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: {
          x: { majorTickMark: "cross" },
          y: { tickLblPos: "high" },
        },
      }),
      "Sheet1",
    );
    const valAxBlocks = [...result.chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map(
      (m) => m[0],
    );
    // First valAx is the X axis (axPos="b"), second is Y (axPos="l").
    expect(valAxBlocks[0]).toContain('c:axPos val="b"');
    expect(valAxBlocks[0]).toContain('<c:majorTickMark val="cross"/>');
    expect(valAxBlocks[0]).not.toContain("c:tickLblPos");
    expect(valAxBlocks[1]).toContain('c:axPos val="l"');
    expect(valAxBlocks[1]).toContain('<c:tickLblPos val="high"/>');
    expect(valAxBlocks[1]).not.toContain("c:majorTickMark");
  });

  it("skips tick rendering on pie charts (pie has no axes)", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        axes: {
          y: { majorTickMark: "cross", minorTickMark: "in", tickLblPos: "low" },
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:majorTickMark");
    expect(result.chartXml).not.toContain("c:minorTickMark");
    expect(result.chartXml).not.toContain("c:tickLblPos");
  });

  it("skips tick rendering on doughnut charts (doughnut has no axes either)", () => {
    const result = writeChart(
      makeChart({
        type: "doughnut",
        axes: { y: { majorTickMark: "cross" } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:majorTickMark");
  });

  it("only emits the major element when minor and tickLblPos are unset", () => {
    const result = writeChart(makeChart({ axes: { y: { majorTickMark: "in" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:majorTickMark val="in"/>');
    expect(valAxBlock).not.toContain("c:minorTickMark");
    expect(valAxBlock).not.toContain("c:tickLblPos");
  });

  it("drops invalid tick-mark values silently", () => {
    const result = writeChart(
      makeChart({
        axes: {
          // @ts-expect-error — testing runtime guard against typo'd inputs.
          y: { majorTickMark: "zigzag", minorTickMark: "diagonal" },
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:majorTickMark");
    expect(result.chartXml).not.toContain("c:minorTickMark");
  });

  it("drops invalid tick-label-position values silently", () => {
    const result = writeChart(
      makeChart({
        axes: {
          // @ts-expect-error — testing runtime guard against typo'd inputs.
          y: { tickLblPos: "diagonal" },
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:tickLblPos");
  });

  it("round-trips a non-default majorTickMark / tickLblPos through parseChart", () => {
    const written = writeChart(
      makeChart({
        axes: {
          y: { majorTickMark: "cross", minorTickMark: "in", tickLblPos: "low" },
        },
      }),
      "Sheet1",
    ).chartXml;
    const parsed = parseChart(written);
    expect(parsed?.axes?.y?.majorTickMark).toBe("cross");
    expect(parsed?.axes?.y?.minorTickMark).toBe("in");
    expect(parsed?.axes?.y?.tickLblPos).toBe("low");
  });

  it("emits all four tick-mark presets on the value axis", () => {
    for (const value of ["none", "in", "out", "cross"] as const) {
      const result = writeChart(makeChart({ axes: { y: { majorTickMark: value } } }), "Sheet1");
      const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
      expect(valAxBlock).toContain(`<c:majorTickMark val="${value}"/>`);
    }
  });

  it("emits all four tick-label-position presets on the value axis", () => {
    for (const value of ["nextTo", "low", "high", "none"] as const) {
      const result = writeChart(makeChart({ axes: { y: { tickLblPos: value } } }), "Sheet1");
      const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
      expect(valAxBlock).toContain(`<c:tickLblPos val="${value}"/>`);
    }
  });

  it("co-emits tick rendering with title, gridlines, scale, and number format", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            title: "Revenue",
            gridlines: { major: true },
            scale: { min: 0, max: 100 },
            numberFormat: { formatCode: "$#,##0" },
            majorTickMark: "cross",
            minorTickMark: "in",
            tickLblPos: "low",
          },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    // Spec order: scaling → axPos → majorGridlines → title → numFmt →
    // majorTickMark → minorTickMark → tickLblPos → crossAx → ...
    const scalingIdx = valAxBlock.indexOf("<c:scaling>");
    const gridlinesIdx = valAxBlock.indexOf("<c:majorGridlines");
    const titleIdx = valAxBlock.indexOf("<c:title>");
    const numFmtIdx = valAxBlock.indexOf("<c:numFmt");
    const majorIdx = valAxBlock.indexOf("<c:majorTickMark");
    const minorIdx = valAxBlock.indexOf("<c:minorTickMark");
    const tickLblIdx = valAxBlock.indexOf("<c:tickLblPos");
    const crossAxIdx = valAxBlock.indexOf("c:crossAx");
    expect(scalingIdx).toBeGreaterThan(0);
    expect(gridlinesIdx).toBeGreaterThan(scalingIdx);
    expect(titleIdx).toBeGreaterThan(gridlinesIdx);
    expect(numFmtIdx).toBeGreaterThan(titleIdx);
    expect(majorIdx).toBeGreaterThan(numFmtIdx);
    expect(minorIdx).toBeGreaterThan(majorIdx);
    expect(tickLblIdx).toBeGreaterThan(minorIdx);
    expect(crossAxIdx).toBeGreaterThan(tickLblIdx);
  });
});

// ── Plot Visible Only ────────────────────────────────────────────────

describe("writeChart — plotVisOnly", () => {
  it('emits <c:plotVisOnly val="1"/> when the field is unset (OOXML default)', () => {
    // The writer always emits the element so the rendered intent is
    // explicit on roundtrip — Excel itself includes it in every
    // reference serialization.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain('c:plotVisOnly val="1"');
    expect(result.chartXml).not.toContain('c:plotVisOnly val="0"');
  });

  it("threads plotVisOnly=false through to <c:chart>", () => {
    // false is the non-default — Excel's "Show data in hidden rows
    // and columns" checkbox checked.
    const result = writeChart(makeChart({ plotVisOnly: false }), "Sheet1");
    expect(result.chartXml).toContain('c:plotVisOnly val="0"');
    expect(result.chartXml).not.toContain('c:plotVisOnly val="1"');
  });

  it("threads plotVisOnly=true through to <c:chart>", () => {
    // Setting the OOXML default explicitly produces the same wire
    // shape as omitting the field — the element is always emitted.
    const result = writeChart(makeChart({ plotVisOnly: true }), "Sheet1");
    expect(result.chartXml).toContain('c:plotVisOnly val="1"');
  });

  it("places <c:plotVisOnly> before <c:dispBlanksAs> inside <c:chart> (OOXML order)", () => {
    // CT_Chart sequence: ... plotArea, legend?, plotVisOnly?, dispBlanksAs?, ...
    const result = writeChart(makeChart({ plotVisOnly: false }), "Sheet1");
    expect(result.chartXml.indexOf("c:plotVisOnly")).toBeLessThan(
      result.chartXml.indexOf("c:dispBlanksAs"),
    );
  });

  it("only emits <c:plotVisOnly> once even on a chart that overrides it", () => {
    // Earlier writers emitted a hardcoded `1` even when the chart
    // requested a different value. Guard against any regression that
    // would double-emit the element.
    const result = writeChart(makeChart({ plotVisOnly: false }), "Sheet1");
    const occurrences = result.chartXml.match(/c:plotVisOnly/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads plotVisOnly through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, plotVisOnly: false }), "Sheet1");
      expect(result.chartXml).toContain('c:plotVisOnly val="0"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        plotVisOnly: false,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:plotVisOnly val="0"');
  });

  it("round-trips a non-default plotVisOnly value through parseChart", () => {
    // A chart with plotVisOnly=false should re-parse into a Chart
    // whose `plotVisOnly` field is `false` (not collapsed to undefined,
    // since false is not the OOXML default).
    const written = writeChart(makeChart({ plotVisOnly: false }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.plotVisOnly).toBe(false);
  });

  it("collapses a defaulted plotVisOnly round-trip back to undefined", () => {
    // A fresh chart (plotVisOnly omitted) writes `1` and re-parses to
    // undefined — absence and the OOXML default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.plotVisOnly).toBeUndefined();
  });

  it("collapses an explicit plotVisOnly=true round-trip back to undefined", () => {
    // Pinning the OOXML default also collapses on read, so a template
    // that explicitly emits `<c:plotVisOnly val="1"/>` is treated the
    // same as one that omits the field.
    const written = writeChart(makeChart({ plotVisOnly: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.plotVisOnly).toBeUndefined();
  });
});

// ── writeChart — showDLblsOverMax ────────────────────────────────────

describe("writeChart — showDLblsOverMax", () => {
  it('emits <c:showDLblsOverMax val="1"/> when the field is unset (OOXML default)', () => {
    // The writer always emits the element so the rendered intent is
    // explicit on roundtrip — mirrors the always-emit contract that
    // <c:plotVisOnly> and <c:dispBlanksAs> follow.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain('c:showDLblsOverMax val="1"');
    expect(result.chartXml).not.toContain('c:showDLblsOverMax val="0"');
  });

  it("threads showDLblsOverMax=false through to <c:chart>", () => {
    // false is the non-default — Excel's "Format Axis -> Labels ->
    // Show data labels for values over maximum scale" checkbox unchecked.
    const result = writeChart(makeChart({ showDLblsOverMax: false }), "Sheet1");
    expect(result.chartXml).toContain('c:showDLblsOverMax val="0"');
    expect(result.chartXml).not.toContain('c:showDLblsOverMax val="1"');
  });

  it("threads showDLblsOverMax=true through to <c:chart>", () => {
    // Setting the OOXML default explicitly produces the same wire
    // shape as omitting the field — the element is always emitted.
    const result = writeChart(makeChart({ showDLblsOverMax: true }), "Sheet1");
    expect(result.chartXml).toContain('c:showDLblsOverMax val="1"');
  });

  it("places <c:showDLblsOverMax> after <c:dispBlanksAs> inside <c:chart> (OOXML order)", () => {
    // CT_Chart sequence: ... plotArea, legend?, plotVisOnly?, dispBlanksAs?,
    // showDLblsOverMax?, extLst? — the new toggle sits at the tail.
    const result = writeChart(makeChart({ showDLblsOverMax: false }), "Sheet1");
    expect(result.chartXml.indexOf("c:dispBlanksAs")).toBeLessThan(
      result.chartXml.indexOf("c:showDLblsOverMax"),
    );
  });

  it("only emits <c:showDLblsOverMax> once even on a chart that overrides it", () => {
    const result = writeChart(makeChart({ showDLblsOverMax: false }), "Sheet1");
    const occurrences = result.chartXml.match(/c:showDLblsOverMax/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads showDLblsOverMax through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, showDLblsOverMax: false }), "Sheet1");
      expect(result.chartXml).toContain('c:showDLblsOverMax val="0"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        showDLblsOverMax: false,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:showDLblsOverMax val="0"');
  });

  it("round-trips a non-default showDLblsOverMax value through parseChart", () => {
    // A chart with showDLblsOverMax=false should re-parse into a Chart
    // whose `showDLblsOverMax` field is `false` (not collapsed to undefined,
    // since false is not the OOXML default).
    const written = writeChart(makeChart({ showDLblsOverMax: false }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.showDLblsOverMax).toBe(false);
  });

  it("collapses a defaulted showDLblsOverMax round-trip back to undefined", () => {
    // A fresh chart (showDLblsOverMax omitted) writes `1` and re-parses to
    // undefined — absence and the OOXML default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.showDLblsOverMax).toBeUndefined();
  });

  it("collapses an explicit showDLblsOverMax=true round-trip back to undefined", () => {
    // Pinning the OOXML default also collapses on read, so a template
    // that explicitly emits `<c:showDLblsOverMax val="1"/>` is treated the
    // same as one that omits the field.
    const written = writeChart(makeChart({ showDLblsOverMax: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.showDLblsOverMax).toBeUndefined();
  });
});

// ── writeChart — roundedCorners ──────────────────────────────────────

describe("writeChart — roundedCorners", () => {
  it('emits <c:roundedCorners val="0"/> when the field is unset (OOXML default)', () => {
    // The writer always emits the element so the rendered intent is
    // explicit on roundtrip — Excel itself includes it in every
    // reference serialization.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain('c:roundedCorners val="0"');
    expect(result.chartXml).not.toContain('c:roundedCorners val="1"');
  });

  it("threads roundedCorners=true through to <c:chartSpace>", () => {
    // true is the non-default — Excel's "Format Chart Area → Border →
    // Rounded corners" toggle on.
    const result = writeChart(makeChart({ roundedCorners: true }), "Sheet1");
    expect(result.chartXml).toContain('c:roundedCorners val="1"');
    expect(result.chartXml).not.toContain('c:roundedCorners val="0"');
  });

  it("threads roundedCorners=false through to <c:chartSpace>", () => {
    // Setting the OOXML default explicitly produces the same wire
    // shape as omitting the field — the element is always emitted.
    const result = writeChart(makeChart({ roundedCorners: false }), "Sheet1");
    expect(result.chartXml).toContain('c:roundedCorners val="0"');
  });

  it("places <c:roundedCorners> before <c:chart> inside <c:chartSpace> (OOXML order)", () => {
    // CT_ChartSpace sequence: ... roundedCorners?, style?, ... chart, ...
    // — the toggle must precede the chart element so a strict validator
    // (Excel itself rejects out-of-order children) sees the schema
    // sequence respected.
    const result = writeChart(makeChart({ roundedCorners: true }), "Sheet1");
    expect(result.chartXml.indexOf("c:roundedCorners")).toBeLessThan(
      result.chartXml.indexOf("<c:chart>"),
    );
  });

  it("only emits <c:roundedCorners> once even on a chart that overrides it", () => {
    // Guard against any regression that would double-emit the element
    // (e.g. one hardcoded copy plus a dynamic one).
    const result = writeChart(makeChart({ roundedCorners: true }), "Sheet1");
    const occurrences = result.chartXml.match(/c:roundedCorners/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads roundedCorners through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, roundedCorners: true }), "Sheet1");
      expect(result.chartXml).toContain('c:roundedCorners val="1"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        roundedCorners: true,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:roundedCorners val="1"');
  });

  it("round-trips a non-default roundedCorners value through parseChart", () => {
    // A chart with roundedCorners=true should re-parse into a Chart
    // whose `roundedCorners` field is `true` (not collapsed to undefined,
    // since true is not the OOXML default).
    const written = writeChart(makeChart({ roundedCorners: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.roundedCorners).toBe(true);
  });

  it("collapses a defaulted roundedCorners round-trip back to undefined", () => {
    // A fresh chart (roundedCorners omitted) writes `0` and re-parses to
    // undefined — absence and the OOXML default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.roundedCorners).toBeUndefined();
  });

  it("collapses an explicit roundedCorners=false round-trip back to undefined", () => {
    // Pinning the OOXML default also collapses on read, so a template
    // that explicitly emits `<c:roundedCorners val="0"/>` is treated the
    // same as one that omits the field.
    const written = writeChart(makeChart({ roundedCorners: false }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.roundedCorners).toBeUndefined();
  });
});

// ── writeChart — axis reverse (orientation) ──────────────────────────

describe("writeChart — axis reverse (orientation)", () => {
  it('emits <c:orientation val="maxMin"/> on the value axis when y.reverse is true', () => {
    const result = writeChart(makeChart({ axes: { y: { reverse: true } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:orientation val="maxMin"/>');
    expect(valAxBlock).not.toContain('val="minMax"');
  });

  it('emits <c:orientation val="maxMin"/> on the category axis when x.reverse is true', () => {
    const result = writeChart(makeChart({ axes: { x: { reverse: true } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('<c:orientation val="maxMin"/>');
    // The value axis keeps the forward minMax default.
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:orientation val="minMax"/>');
  });

  it("falls back to minMax when reverse is unset, false, or absent", () => {
    // A fresh chart (no axes block at all) emits the OOXML default on
    // both axes — the writer never omits <c:orientation> because Excel
    // requires it inside <c:scaling>.
    const noAxes = writeChart(makeChart(), "Sheet1").chartXml;
    expect(noAxes.match(/<c:orientation val="minMax"\/>/g)?.length).toBe(2);
    expect(noAxes).not.toContain('val="maxMin"');

    const explicitFalse = writeChart(
      makeChart({ axes: { x: { reverse: false }, y: { reverse: false } } }),
      "Sheet1",
    ).chartXml;
    expect(explicitFalse.match(/<c:orientation val="minMax"\/>/g)?.length).toBe(2);
    expect(explicitFalse).not.toContain('val="maxMin"');
  });

  it("places <c:orientation> in the spec-required slot inside <c:scaling>", () => {
    // CT_Scaling sequence: logBase → orientation → max → min. The
    // writer relies on this order for the OOXML schema validator to
    // accept the chart.
    const result = writeChart(
      makeChart({
        axes: {
          y: { reverse: true, scale: { min: 0, max: 100, logBase: 10 } },
        },
      }),
      "Sheet1",
    );
    const scaling = result.chartXml.match(/<c:scaling>[\s\S]*?<\/c:scaling>/g)!;
    // Two scaling elements (catAx and valAx) — pick the one with logBase
    // / max / min, that's the value axis.
    const valScaling = scaling.find((s) => s.includes("logBase"))!;
    const logIdx = valScaling.indexOf("c:logBase");
    const orientIdx = valScaling.indexOf("c:orientation");
    const maxIdx = valScaling.indexOf("c:max");
    const minIdx = valScaling.indexOf("c:min");
    expect(logIdx).toBeGreaterThan(0);
    expect(orientIdx).toBeGreaterThan(logIdx);
    expect(maxIdx).toBeGreaterThan(orientIdx);
    expect(minIdx).toBeGreaterThan(maxIdx);
  });

  it("works for line and area charts (which share the bar axis builder)", () => {
    for (const type of ["line", "area"] as const) {
      const result = writeChart(makeChart({ type, axes: { y: { reverse: true } } }), "Sheet1");
      const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
      expect(valAxBlock).toContain('<c:orientation val="maxMin"/>');
    }
  });

  it("emits reverse on scatter X (axPos=b) and Y (axPos=l) value axes independently", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { reverse: true }, y: { reverse: false } },
      }),
      "Sheet1",
    );
    const valAxBlocks = [...result.chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map(
      (m) => m[0],
    );
    // First valAx is scatter X axis (axPos="b"), second is Y (axPos="l").
    expect(valAxBlocks[0]).toContain('c:axPos val="b"');
    expect(valAxBlocks[0]).toContain('<c:orientation val="maxMin"/>');
    expect(valAxBlocks[1]).toContain('c:axPos val="l"');
    expect(valAxBlocks[1]).toContain('<c:orientation val="minMax"/>');
  });

  it("skips orientation reverse on pie charts (pie has no axes)", () => {
    const result = writeChart(makeChart({ type: "pie", axes: { y: { reverse: true } } }), "Sheet1");
    // Pie writes no <c:catAx> / <c:valAx> at all, so no <c:scaling>
    // / <c:orientation> elements appear.
    expect(result.chartXml).not.toContain("c:orientation");
    expect(result.chartXml).not.toContain("c:scaling");
  });

  it("skips orientation reverse on doughnut charts (doughnut has no axes either)", () => {
    const result = writeChart(
      makeChart({ type: "doughnut", axes: { y: { reverse: true } } }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:orientation");
  });

  it("only flips the targeted axis — the other stays at the forward default", () => {
    // Reversing X must not propagate to Y (and vice versa) — each axis
    // pulls its own reverse flag off chart.axes.{x,y}.reverse.
    const result = writeChart(makeChart({ axes: { x: { reverse: true } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).toContain('<c:orientation val="maxMin"/>');
    expect(valAxBlock).toContain('<c:orientation val="minMax"/>');
  });

  it("treats truthy non-boolean values as forward (reverse only fires for `=== true`)", () => {
    // A defensively-typed source (e.g. "yes" leaking past the type
    // guard) should not silently flip orientation — only the literal
    // boolean `true` triggers reverse.
    const result = writeChart(
      makeChart({
        axes: {
          // @ts-expect-error — testing runtime guard against typo'd inputs.
          y: { reverse: "yes" },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:orientation val="minMax"/>');
  });

  it("round-trips reverse=true through parseChart", () => {
    const written = writeChart(makeChart({ axes: { y: { reverse: true } } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.y?.reverse).toBe(true);
  });

  it("round-trips reverse=false / absent back to undefined", () => {
    // An unset reverse writes the forward minMax default; on re-parse
    // that default collapses to undefined so absence and the default
    // round-trip identically.
    for (const chart of [makeChart(), makeChart({ axes: { y: { reverse: false } } })]) {
      const written = writeChart(chart, "Sheet1").chartXml;
      const reparsed = parseChart(written);
      expect(reparsed?.axes?.y?.reverse).toBeUndefined();
    }
  });
});

// ── Axis tick label / mark skip ──────────────────────────────────────

describe("writeChart — axis tickLblSkip / tickMarkSkip", () => {
  it("emits <c:tickLblSkip> on the category axis when set", () => {
    const result = writeChart(makeChart({ axes: { x: { tickLblSkip: 3 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:tickLblSkip val="3"');
  });

  it("emits <c:tickMarkSkip> on the category axis when set", () => {
    const result = writeChart(makeChart({ axes: { x: { tickMarkSkip: 5 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:tickMarkSkip val="5"');
  });

  it("emits both skips together in the OOXML-required order", () => {
    // CT_CatAx: lblOffset → tickLblSkip → tickMarkSkip → noMultiLvlLbl.
    const result = writeChart(
      makeChart({ axes: { x: { tickLblSkip: 2, tickMarkSkip: 4 } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const lblOffsetIdx = catAxBlock.indexOf("c:lblOffset");
    const tickLblSkipIdx = catAxBlock.indexOf("c:tickLblSkip");
    const tickMarkSkipIdx = catAxBlock.indexOf("c:tickMarkSkip");
    const noMultiLvlIdx = catAxBlock.indexOf("c:noMultiLvlLbl");
    expect(lblOffsetIdx).toBeGreaterThan(0);
    expect(tickLblSkipIdx).toBeGreaterThan(lblOffsetIdx);
    expect(tickMarkSkipIdx).toBeGreaterThan(tickLblSkipIdx);
    expect(noMultiLvlIdx).toBeGreaterThan(tickMarkSkipIdx);
  });

  it("omits the elements when tickLblSkip / tickMarkSkip are unset (Excel default)", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("c:tickLblSkip");
    expect(result.chartXml).not.toContain("c:tickMarkSkip");
  });

  it("omits the elements when the value is the OOXML default 1", () => {
    // Absence and the default `1` round-trip identically. The writer
    // therefore drops the element when the caller pinned `1` so the
    // emitted XML matches Excel's reference serialization byte-for-byte.
    const result = writeChart(
      makeChart({ axes: { x: { tickLblSkip: 1, tickMarkSkip: 1 } } }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:tickLblSkip");
    expect(result.chartXml).not.toContain("c:tickMarkSkip");
  });

  it("drops out-of-range values without clamping", () => {
    // ST_SkipIntervals restricts the value to 1..32767. Passing 0,
    // -3, or 99999 drops the element silently rather than clamping —
    // a silent clamp would mask the configuration error.
    const result = writeChart(
      makeChart({
        axes: { x: { tickLblSkip: 0, tickMarkSkip: 99999 } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:tickLblSkip");
    expect(result.chartXml).not.toContain("c:tickMarkSkip");
  });

  it("rounds non-integer values to the nearest integer", () => {
    const result = writeChart(makeChart({ axes: { x: { tickLblSkip: 3.7 } } }), "Sheet1");
    expect(result.chartXml).toContain('c:tickLblSkip val="4"');
  });

  it("accepts the schema boundaries 2 and 32767", () => {
    const lo = writeChart(makeChart({ axes: { x: { tickLblSkip: 2 } } }), "Sheet1");
    expect(lo.chartXml).toContain('c:tickLblSkip val="2"');
    const hi = writeChart(makeChart({ axes: { x: { tickLblSkip: 32767 } } }), "Sheet1");
    expect(hi.chartXml).toContain('c:tickLblSkip val="32767"');
  });

  it("emits each element exactly once on the rendered chart", () => {
    const result = writeChart(
      makeChart({ axes: { x: { tickLblSkip: 3, tickMarkSkip: 5 } } }),
      "Sheet1",
    );
    expect((result.chartXml.match(/c:tickLblSkip/g) ?? []).length).toBe(1);
    expect((result.chartXml.match(/c:tickMarkSkip/g) ?? []).length).toBe(1);
  });

  it("threads the skips through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(
        makeChart({ type, axes: { x: { tickLblSkip: 3, tickMarkSkip: 5 } } }),
        "Sheet1",
      );
      expect(result.chartXml).toContain('c:tickLblSkip val="3"');
      expect(result.chartXml).toContain('c:tickMarkSkip val="5"');
    }
  });

  it("ignores the skips on scatter charts (both axes are value axes)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { tickLblSkip: 3, tickMarkSkip: 5 } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:tickLblSkip");
    expect(result.chartXml).not.toContain("c:tickMarkSkip");
  });

  it("ignores the skips on pie / doughnut charts (no axes at all)", () => {
    const pie = writeChart(makeChart({ type: "pie", axes: { x: { tickLblSkip: 3 } } }), "Sheet1");
    expect(pie.chartXml).not.toContain("c:tickLblSkip");
    const dough = writeChart(
      makeChart({ type: "doughnut", axes: { x: { tickMarkSkip: 4 } } }),
      "Sheet1",
    );
    expect(dough.chartXml).not.toContain("c:tickMarkSkip");
  });

  it("does not emit the elements on the value axis even when set on .y", () => {
    // The model surfaces these only on `axes.x`; setting them via
    // `axes.y` is impossible at the type level. This test pins the
    // negative case for the writer: a valAx never carries tick skips.
    const result = writeChart(
      makeChart({ axes: { x: { tickLblSkip: 3, tickMarkSkip: 5 } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).not.toContain("c:tickLblSkip");
    expect(valAxBlock).not.toContain("c:tickMarkSkip");
  });

  it("round-trips a non-default skip pair through parseChart", () => {
    const written = writeChart(
      makeChart({ axes: { x: { tickLblSkip: 3, tickMarkSkip: 5 } } }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.tickLblSkip).toBe(3);
    expect(reparsed?.axes?.x?.tickMarkSkip).toBe(5);
  });

  it("collapses a defaulted skip round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toBeUndefined();
  });

  it("places tick skips inside the catAx without breaking schema-required ordering of other elements", () => {
    // Combine title, gridlines, scale, number format and skips on the
    // X axis to verify the catAx still renders in spec order.
    const result = writeChart(
      makeChart({
        axes: {
          x: {
            title: "Region",
            gridlines: { major: true },
            numberFormat: { formatCode: "@" },
            tickLblSkip: 3,
            tickMarkSkip: 5,
          },
        },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const idx = (needle: string): number => catAxBlock.indexOf(needle);
    expect(idx("c:axId")).toBeLessThan(idx("c:scaling"));
    expect(idx("c:scaling")).toBeLessThan(idx("c:axPos"));
    expect(idx("c:axPos")).toBeLessThan(idx("c:majorGridlines"));
    expect(idx("c:majorGridlines")).toBeLessThan(idx("c:title"));
    expect(idx("c:title")).toBeLessThan(idx("c:numFmt"));
    expect(idx("c:numFmt")).toBeLessThan(idx("c:crossAx"));
    expect(idx("c:lblOffset")).toBeLessThan(idx("c:tickLblSkip"));
    expect(idx("c:tickLblSkip")).toBeLessThan(idx("c:tickMarkSkip"));
    expect(idx("c:tickMarkSkip")).toBeLessThan(idx("c:noMultiLvlLbl"));
  });
});

describe("writeChart — axis lblOffset", () => {
  it("emits the override value on the category axis when set", () => {
    const result = writeChart(makeChart({ axes: { x: { lblOffset: 250 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblOffset val="250"');
  });

  it("emits the OOXML default 100 when the field is unset", () => {
    // Excel's reference serialization always emits `<c:lblOffset val="100"/>`,
    // so the writer keeps that contract on a stock chart even though the
    // parser collapses `100` to undefined on the read side.
    const result = writeChart(makeChart(), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblOffset val="100"');
  });

  it("collapses an explicit 100 (the default) back to the default emit", () => {
    // Pinning the default has the same effect as omitting the field —
    // the OOXML default `100` round-trips identically with absence.
    const result = writeChart(makeChart({ axes: { x: { lblOffset: 100 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblOffset val="100"');
  });

  it("drops out-of-range overrides without clamping (falls back to default 100)", () => {
    // ST_LblOffsetPercent restricts the value to 0..1000. Passing -5 or
    // 9999 collapses to undefined inside `normalizeAxisLblOffset`, so the
    // writer falls back to the default `100` rather than clamping.
    const lo = writeChart(makeChart({ axes: { x: { lblOffset: -5 } } }), "Sheet1");
    const loCatAx = lo.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(loCatAx).toContain('c:lblOffset val="100"');

    const hi = writeChart(makeChart({ axes: { x: { lblOffset: 9999 } } }), "Sheet1");
    const hiCatAx = hi.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(hiCatAx).toContain('c:lblOffset val="100"');
  });

  it("rounds non-integer values to the nearest integer", () => {
    const result = writeChart(makeChart({ axes: { x: { lblOffset: 247.6 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblOffset val="248"');
  });

  it("accepts the schema boundaries 0 and 1000", () => {
    const lo = writeChart(makeChart({ axes: { x: { lblOffset: 0 } } }), "Sheet1");
    expect(lo.chartXml).toContain('c:lblOffset val="0"');
    const hi = writeChart(makeChart({ axes: { x: { lblOffset: 1000 } } }), "Sheet1");
    expect(hi.chartXml).toContain('c:lblOffset val="1000"');
  });

  it("emits exactly one <c:lblOffset> element per category axis", () => {
    const result = writeChart(makeChart({ axes: { x: { lblOffset: 250 } } }), "Sheet1");
    expect((result.chartXml.match(/c:lblOffset/g) ?? []).length).toBe(1);
  });

  it("threads the override through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, axes: { x: { lblOffset: 200 } } }), "Sheet1");
      expect(result.chartXml).toContain('c:lblOffset val="200"');
    }
  });

  it("ignores the override on scatter charts (both axes are value axes)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { lblOffset: 250 } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:lblOffset");
  });

  it("ignores the override on pie / doughnut charts (no axes at all)", () => {
    const pie = writeChart(makeChart({ type: "pie", axes: { x: { lblOffset: 250 } } }), "Sheet1");
    expect(pie.chartXml).not.toContain("c:lblOffset");
    const dough = writeChart(
      makeChart({ type: "doughnut", axes: { x: { lblOffset: 250 } } }),
      "Sheet1",
    );
    expect(dough.chartXml).not.toContain("c:lblOffset");
  });

  it("does not emit lblOffset on the value axis", () => {
    // The model surfaces the offset only on `axes.x`; setting it via
    // `axes.y` is impossible at the type level. This test pins the
    // negative case for the writer: a valAx never carries lblOffset.
    const result = writeChart(makeChart({ axes: { x: { lblOffset: 250 } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).not.toContain("c:lblOffset");
  });

  it("places lblOffset between lblAlgn and tickLblSkip per the OOXML schema", () => {
    // CT_CatAx: ... lblAlgn → lblOffset → tickLblSkip → tickMarkSkip → noMultiLvlLbl.
    const result = writeChart(
      makeChart({ axes: { x: { lblOffset: 250, tickLblSkip: 3 } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const lblAlgnIdx = catAxBlock.indexOf("c:lblAlgn");
    const lblOffsetIdx = catAxBlock.indexOf("c:lblOffset");
    const tickLblSkipIdx = catAxBlock.indexOf("c:tickLblSkip");
    expect(lblAlgnIdx).toBeGreaterThan(0);
    expect(lblOffsetIdx).toBeGreaterThan(lblAlgnIdx);
    expect(tickLblSkipIdx).toBeGreaterThan(lblOffsetIdx);
  });

  it("round-trips a non-default offset through parseChart", () => {
    const written = writeChart(makeChart({ axes: { x: { lblOffset: 200 } } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.lblOffset).toBe(200);
  });

  it("collapses a defaulted offset round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toBeUndefined();
  });
});

// ── writeChart — axis hidden flag (<c:delete>) ──────────────────────

describe("writeChart — axis hidden", () => {
  it('emits <c:delete val="0"/> on both axes by default (Excel reference shape)', () => {
    const result = writeChart(makeChart(), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).toContain('<c:delete val="0"/>');
    expect(valAxBlock).toContain('<c:delete val="0"/>');
  });

  it('emits <c:delete val="1"/> on the category axis when axes.x.hidden=true', () => {
    const result = writeChart(makeChart({ axes: { x: { hidden: true } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).toContain('<c:delete val="1"/>');
    // The Y axis stays visible — the flag is per-axis.
    expect(valAxBlock).toContain('<c:delete val="0"/>');
  });

  it('emits <c:delete val="1"/> on the value axis when axes.y.hidden=true', () => {
    const result = writeChart(makeChart({ axes: { y: { hidden: true } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).toContain('<c:delete val="0"/>');
    expect(valAxBlock).toContain('<c:delete val="1"/>');
  });

  it("hides both axes when axes.x.hidden and axes.y.hidden are both true", () => {
    const result = writeChart(
      makeChart({ axes: { x: { hidden: true }, y: { hidden: true } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).toContain('<c:delete val="1"/>');
    expect(valAxBlock).toContain('<c:delete val="1"/>');
  });

  it("treats axes.x.hidden=false the same as omitting the field", () => {
    const explicit = writeChart(makeChart({ axes: { x: { hidden: false } } }), "Sheet1").chartXml;
    const implicit = writeChart(makeChart(), "Sheet1").chartXml;
    expect(explicit).toEqual(implicit);
  });

  it('collapses non-boolean inputs to the default val="0"', () => {
    // A stray non-boolean leaking past the type guard (e.g. `0` / `1` /
    // `"true"` / `null`) must collapse to the default rather than emit
    // an attribute Excel would reject.
    const result = writeChart(
      makeChart({ axes: { x: { hidden: 1 as unknown as boolean } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('<c:delete val="0"/>');
  });

  it("threads the flag through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, axes: { x: { hidden: true } } }), "Sheet1");
      const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
      expect(catAxBlock).toContain('<c:delete val="1"/>');
    }
  });

  it("emits the flag on scatter X (axPos=b) and Y (axPos=l) value axes", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { hidden: true } },
      }),
      "Sheet1",
    );
    const valAxBlocks = [...result.chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)].map(
      (m) => m[0],
    );
    // First valAx is the X axis (axPos="b"), second is Y (axPos="l").
    expect(valAxBlocks[0]).toContain('c:axPos val="b"');
    expect(valAxBlocks[0]).toContain('<c:delete val="1"/>');
    expect(valAxBlocks[1]).toContain('c:axPos val="l"');
    expect(valAxBlocks[1]).toContain('<c:delete val="0"/>');
  });

  it("ignores the flag on pie charts (no axes at all)", () => {
    const result = writeChart(makeChart({ type: "pie", axes: { x: { hidden: true } } }), "Sheet1");
    // Pie chart emits no <c:catAx> / <c:valAx> at all, so there is no
    // <c:delete> to find. The flag must not leak elsewhere.
    expect(result.chartXml).not.toContain("c:catAx");
    expect(result.chartXml).not.toContain("c:valAx");
  });

  it("ignores the flag on doughnut charts (no axes at all)", () => {
    const result = writeChart(
      makeChart({ type: "doughnut", axes: { y: { hidden: true } } }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:catAx");
    expect(result.chartXml).not.toContain("c:valAx");
  });

  it("places <c:delete> after <c:scaling> and before <c:axPos> (OOXML order)", () => {
    const result = writeChart(makeChart({ axes: { x: { hidden: true } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const scalingIdx = catAxBlock.indexOf("<c:scaling");
    const deleteIdx = catAxBlock.indexOf("<c:delete");
    const axPosIdx = catAxBlock.indexOf("<c:axPos");
    expect(scalingIdx).toBeGreaterThan(0);
    expect(deleteIdx).toBeGreaterThan(scalingIdx);
    expect(axPosIdx).toBeGreaterThan(deleteIdx);
  });

  it("emits exactly one <c:delete> per axis", () => {
    const result = writeChart(
      makeChart({ axes: { x: { hidden: true }, y: { hidden: true } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect((catAxBlock.match(/<c:delete /g) ?? []).length).toBe(1);
    expect((valAxBlock.match(/<c:delete /g) ?? []).length).toBe(1);
  });

  it("composes alongside other axis fields without breaking spec ordering", () => {
    // Combine title, gridlines, scale, number format, tick rendering and
    // hidden on the X axis to verify the catAx still renders in spec order.
    const result = writeChart(
      makeChart({
        axes: {
          x: {
            title: "Region",
            gridlines: { major: true },
            numberFormat: { formatCode: "@" },
            majorTickMark: "cross",
            tickLblPos: "low",
            hidden: true,
          },
        },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const idx = (needle: string): number => catAxBlock.indexOf(needle);
    expect(idx("c:axId")).toBeLessThan(idx("c:scaling"));
    expect(idx("c:scaling")).toBeLessThan(idx("c:delete"));
    expect(idx("c:delete")).toBeLessThan(idx("c:axPos"));
    expect(idx("c:axPos")).toBeLessThan(idx("c:majorGridlines"));
    expect(idx("c:majorGridlines")).toBeLessThan(idx("c:title"));
    expect(idx("c:title")).toBeLessThan(idx("c:numFmt"));
    expect(idx("c:numFmt")).toBeLessThan(idx("c:majorTickMark"));
    expect(idx("c:majorTickMark")).toBeLessThan(idx("c:tickLblPos"));
    expect(idx("c:tickLblPos")).toBeLessThan(idx("c:crossAx"));
    expect(catAxBlock).toContain('<c:delete val="1"/>');
  });

  it("round-trips axes.x.hidden=true through parseChart", () => {
    const written = writeChart(makeChart({ axes: { x: { hidden: true } } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.hidden).toBe(true);
    // The Y axis pinned val="0" so it collapses to undefined.
    expect(reparsed?.axes?.y?.hidden).toBeUndefined();
  });

  it('collapses a default round-trip back to undefined axes (val="0" alone is the default)', () => {
    // No axis fields set at all → the writer still emits <c:delete val="0"/>
    // on every axis but the reader collapses both axes to no info, leaving
    // `axes` undefined.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toBeUndefined();
  });
});

describe("writeChart — axis lblAlgn", () => {
  it("emits the override value on the category axis when set", () => {
    const result = writeChart(makeChart({ axes: { x: { lblAlgn: "l" } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblAlgn val="l"');
  });

  it('emits the OOXML default "ctr" when the field is unset', () => {
    // Excel's reference serialization always emits `<c:lblAlgn val="ctr"/>`,
    // so the writer keeps that contract on a stock chart even though the
    // parser collapses `ctr` to undefined on the read side.
    const result = writeChart(makeChart(), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblAlgn val="ctr"');
  });

  it('collapses an explicit "ctr" (the default) back to the default emit', () => {
    // Pinning the default has the same effect as omitting the field —
    // the OOXML default `"ctr"` round-trips identically with absence.
    const result = writeChart(makeChart({ axes: { x: { lblAlgn: "ctr" } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblAlgn val="ctr"');
  });

  it('emits "r" alignment when the override pins it', () => {
    const result = writeChart(makeChart({ axes: { x: { lblAlgn: "r" } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblAlgn val="r"');
  });

  it('drops unknown overrides without falling through (falls back to default "ctr")', () => {
    // ST_LblAlgn restricts the value to ctr / l / r. Unknown tokens like
    // "left" or "center" collapse to undefined inside `normalizeAxisLblAlgn`,
    // so the writer falls back to the default `"ctr"` rather than fabricating
    // a value Excel would reject.
    const result = writeChart(makeChart({ axes: { x: { lblAlgn: "left" as never } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:lblAlgn val="ctr"');
  });

  it("emits exactly one <c:lblAlgn> element per category axis", () => {
    const result = writeChart(makeChart({ axes: { x: { lblAlgn: "l" } } }), "Sheet1");
    expect((result.chartXml.match(/c:lblAlgn/g) ?? []).length).toBe(1);
  });

  it("threads the override through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, axes: { x: { lblAlgn: "r" } } }), "Sheet1");
      expect(result.chartXml).toContain('c:lblAlgn val="r"');
    }
  });

  it("ignores the override on scatter charts (both axes are value axes)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { lblAlgn: "l" } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:lblAlgn");
  });

  it("ignores the override on pie / doughnut charts (no axes at all)", () => {
    const pie = writeChart(makeChart({ type: "pie", axes: { x: { lblAlgn: "l" } } }), "Sheet1");
    expect(pie.chartXml).not.toContain("c:lblAlgn");
    const dough = writeChart(
      makeChart({ type: "doughnut", axes: { x: { lblAlgn: "l" } } }),
      "Sheet1",
    );
    expect(dough.chartXml).not.toContain("c:lblAlgn");
  });

  it("does not emit lblAlgn on the value axis", () => {
    // The model surfaces the alignment only on `axes.x`; setting it via
    // `axes.y` is impossible at the type level. This test pins the
    // negative case for the writer: a valAx never carries lblAlgn.
    const result = writeChart(makeChart({ axes: { x: { lblAlgn: "l" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).not.toContain("c:lblAlgn");
  });

  it("places lblAlgn between auto and lblOffset per the OOXML schema", () => {
    // CT_CatAx: ... auto -> lblAlgn -> lblOffset -> tickLblSkip -> tickMarkSkip -> noMultiLvlLbl.
    const result = writeChart(
      makeChart({ axes: { x: { lblAlgn: "l", lblOffset: 200 } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const autoIdx = catAxBlock.indexOf("c:auto");
    const lblAlgnIdx = catAxBlock.indexOf("c:lblAlgn");
    const lblOffsetIdx = catAxBlock.indexOf("c:lblOffset");
    expect(autoIdx).toBeGreaterThan(0);
    expect(lblAlgnIdx).toBeGreaterThan(autoIdx);
    expect(lblOffsetIdx).toBeGreaterThan(lblAlgnIdx);
  });

  it("round-trips a non-default alignment through parseChart", () => {
    const written = writeChart(makeChart({ axes: { x: { lblAlgn: "l" } } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.lblAlgn).toBe("l");
  });

  it("collapses a defaulted alignment round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toBeUndefined();
  });

  it("end-to-end: writeXlsx packages the alignment into chart1.xml", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B"],
            [1, 2],
            [3, 4],
          ],
          charts: [
            {
              type: "column",
              series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
              anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
              axes: { x: { lblAlgn: "r" } },
            },
          ],
        },
      ],
    });
    const written = await extractXml(xlsx, "xl/charts/chart1.xml");
    expect(written).toContain('c:lblAlgn val="r"');
  });
});

// ── writeChart — legend overlay ──────────────────────────────────────

describe("writeChart — legendOverlay", () => {
  it('emits <c:overlay val="0"/> when the field is unset (OOXML default)', () => {
    // The writer always emits the element so the rendered intent is
    // explicit on roundtrip — Excel itself includes it in every
    // reference legend serialization.
    const result = writeChart(makeChart(), "Sheet1");
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend).toContain('c:overlay val="0"');
    expect(legend).not.toContain('c:overlay val="1"');
  });

  it("threads legendOverlay=true through to <c:legend>", () => {
    // true is the non-default — Excel's "Show the legend without
    // overlapping the chart" toggle off (the legend is drawn on top of
    // the plot area).
    const result = writeChart(makeChart({ legendOverlay: true }), "Sheet1");
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend).toContain('c:overlay val="1"');
    expect(legend).not.toContain('c:overlay val="0"');
  });

  it("threads legendOverlay=false through to <c:legend>", () => {
    // Setting the OOXML default explicitly produces the same wire shape
    // as omitting the field — the element is always emitted.
    const result = writeChart(makeChart({ legendOverlay: false }), "Sheet1");
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend).toContain('c:overlay val="0"');
  });

  it("places <c:overlay> after <c:legendPos> inside <c:legend> (OOXML order)", () => {
    // CT_Legend sequence: legendPos?, legendEntry*, layout?, overlay?, ...
    const result = writeChart(makeChart({ legendOverlay: true }), "Sheet1");
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend.indexOf("c:legendPos")).toBeLessThan(legend.indexOf("c:overlay"));
  });

  it("only emits <c:overlay> once inside <c:legend> even on a chart that overrides it", () => {
    // Guard against any regression that would double-emit the element
    // (e.g. one hardcoded copy plus a dynamic one). The title also
    // carries its own `<c:overlay>` so we scope the count to the legend.
    const result = writeChart(makeChart({ legendOverlay: true }), "Sheet1");
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    const occurrences = legend.match(/c:overlay/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("does not emit any <c:legend> when legend=false", () => {
    // A hidden legend has no slot for an overlay flag — the writer
    // suppresses the entire legend element rather than emit a stray
    // overlay child Excel would never read.
    const result = writeChart(makeChart({ legend: false, legendOverlay: true }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:legend>");
    // The title still carries its own <c:overlay>; ensure no legend
    // element exists so we know no legend-overlay snuck in.
    expect(result.chartXml.match(/<c:legend\b/g)).toBeNull();
  });

  it("threads legendOverlay through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, legendOverlay: true }), "Sheet1");
      const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
      expect(legend).toContain('c:overlay val="1"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        legendOverlay: true,
      }),
      "Sheet1",
    );
    const legend = scatter.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend).toContain('c:overlay val="1"');
  });

  it("round-trips a non-default legendOverlay value through parseChart", () => {
    // A chart with legendOverlay=true should re-parse into a Chart whose
    // `legendOverlay` field is `true` (not collapsed to undefined since
    // true is not the OOXML default).
    const written = writeChart(makeChart({ legendOverlay: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.legendOverlay).toBe(true);
  });

  it("collapses a defaulted legendOverlay round-trip back to undefined", () => {
    // A fresh chart (legendOverlay omitted) writes `0` and re-parses to
    // undefined — absence and the OOXML default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.legendOverlay).toBeUndefined();
  });

  it("collapses an explicit legendOverlay=false round-trip back to undefined", () => {
    // Pinning the OOXML default also collapses on read, so a template
    // that explicitly emits `<c:overlay val="0"/>` is treated the same
    // as one that omits the field.
    const written = writeChart(makeChart({ legendOverlay: false }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.legendOverlay).toBeUndefined();
  });

  it("ignores non-boolean legendOverlay values", () => {
    // Match how `roundedCorners` / `plotVisOnly` / axis hidden treat
    // their inputs: only literal `true` produces the non-default. A
    // stray non-boolean (e.g. truthy string) collapses to the default.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = writeChart(makeChart({ legendOverlay: "yes" as any }), "Sheet1");
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend).toContain('c:overlay val="0"');
  });

  it("survives a writeXlsx round trip — legendOverlay lands in the packaged chart XML", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            legendOverlay: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    const legend = chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend).toContain('c:overlay val="1"');
  });
});

// ── writeChart — data labels showLegendKey ──────────────────────────

describe("writeChart — data labels showLegendKey", () => {
  function dLblsOf(xml: string): string {
    const m = xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/);
    if (!m) throw new Error("No <c:dLbls> block found in chart XML");
    return m[0];
  }

  it('emits <c:showLegendKey val="0"/> by default when chart-level dataLabels is set', () => {
    const result = writeChart(makeChart({ dataLabels: { showValue: true } }), "Sheet1");
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:showLegendKey val="0"/>');
  });

  it('emits <c:showLegendKey val="1"/> when chart-level showLegendKey=true', () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, showLegendKey: true } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:showLegendKey val="1"/>');
  });

  it('treats showLegendKey=false the same as omitting the field (val="0")', () => {
    const explicit = writeChart(
      makeChart({ dataLabels: { showValue: true, showLegendKey: false } }),
      "Sheet1",
    ).chartXml;
    const implicit = writeChart(makeChart({ dataLabels: { showValue: true } }), "Sheet1").chartXml;
    expect(explicit).toEqual(implicit);
  });

  it('collapses non-boolean inputs to the default val="0"', () => {
    // A stray non-boolean leaking past the type guard (e.g. 1 / "true" /
    // null) must collapse to the default rather than emit something Excel
    // would reject.
    const result = writeChart(
      makeChart({
        dataLabels: { showValue: true, showLegendKey: 1 as unknown as boolean },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:showLegendKey val="0"/>');
  });

  it("emits showLegendKey first among the show* toggles (CT_DLbls order)", () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, showLegendKey: true } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    const idxLk = dLbls.indexOf("<c:showLegendKey");
    const idxVal = dLbls.indexOf("<c:showVal");
    const idxCat = dLbls.indexOf("<c:showCatName");
    const idxSer = dLbls.indexOf("<c:showSerName");
    const idxPct = dLbls.indexOf("<c:showPercent");
    const idxBub = dLbls.indexOf("<c:showBubbleSize");
    expect(idxLk).toBeGreaterThan(0);
    expect(idxLk).toBeLessThan(idxVal);
    expect(idxVal).toBeLessThan(idxCat);
    expect(idxCat).toBeLessThan(idxSer);
    expect(idxSer).toBeLessThan(idxPct);
    expect(idxPct).toBeLessThan(idxBub);
  });

  it("emits exactly one <c:showLegendKey> per <c:dLbls> block", () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, showLegendKey: true } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect((dLbls.match(/<c:showLegendKey /g) ?? []).length).toBe(1);
  });

  it("places <c:showLegendKey> after <c:dLblPos> when the position is set", () => {
    const result = writeChart(
      makeChart({
        dataLabels: { showValue: true, position: "outEnd", showLegendKey: true },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls.indexOf("<c:dLblPos")).toBeLessThan(dLbls.indexOf("<c:showLegendKey"));
  });

  it("threads showLegendKey through a series-level <c:dLbls>", () => {
    const result = writeChart(
      makeChart({
        series: [
          {
            name: "S1",
            values: "B2:B4",
            dataLabels: { showValue: true, showLegendKey: true },
          },
        ],
      }),
      "Sheet1",
    );
    const xml = result.chartXml;
    const serStart = xml.indexOf("<c:ser>");
    const serEnd = xml.indexOf("</c:ser>");
    const inner = xml.slice(serStart, serEnd);
    expect(inner).toContain('<c:showLegendKey val="1"/>');
  });

  it("threads showLegendKey through pie / line / scatter chart families", () => {
    for (const type of ["pie", "line", "scatter"] as const) {
      const result = writeChart(
        makeChart({ type, dataLabels: { showValue: true, showLegendKey: true } }),
        "Sheet1",
      );
      const dLbls = dLblsOf(result.chartXml);
      expect(dLbls).toContain('<c:showLegendKey val="1"/>');
    }
  });

  it("round-trips a chart with showLegendKey=true through parseChart", () => {
    const written = writeChart(
      makeChart({ dataLabels: { showValue: true, showLegendKey: true } }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataLabels?.showLegendKey).toBe(true);
    expect(reparsed?.dataLabels?.showValue).toBe(true);
  });

  it("end-to-end: writeXlsx packages a chart with showLegendKey=true", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            dataLabels: { showValue: true, showLegendKey: true },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    const dLbls = chartXml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/)![0];
    expect(dLbls).toContain('<c:showLegendKey val="1"/>');
  });
});

// ── writeChart — data labels numberFormat ────────────────────────────

describe("writeChart — data labels numberFormat", () => {
  function dLblsOf(xml: string): string {
    const m = xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/);
    if (!m) throw new Error("No <c:dLbls> block found in chart XML");
    return m[0];
  }

  it("emits <c:numFmt> inside <c:dLbls> when chart-level numberFormat is set", () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, numberFormat: { formatCode: "0.00%" } } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:numFmt formatCode="0.00%" sourceLinked="0"/>');
  });

  it('emits sourceLinked="1" when sourceLinked=true is pinned', () => {
    const result = writeChart(
      makeChart({
        dataLabels: {
          showValue: true,
          numberFormat: { formatCode: "General", sourceLinked: true },
        },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:numFmt formatCode="General" sourceLinked="1"/>');
  });

  it('defaults sourceLinked to "0" when the field is omitted', () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, numberFormat: { formatCode: "$#,##0.00" } } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:numFmt formatCode="$#,##0.00" sourceLinked="0"/>');
  });

  it("skips emitting <c:numFmt> when numberFormat is unset", () => {
    const result = writeChart(makeChart({ dataLabels: { showValue: true } }), "Sheet1");
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).not.toContain("<c:numFmt");
  });

  it("skips emitting <c:numFmt> when formatCode is missing or empty", () => {
    const noCode = writeChart(
      makeChart({
        dataLabels: {
          showValue: true,
          numberFormat: { formatCode: undefined as unknown as string },
        },
      }),
      "Sheet1",
    );
    expect(dLblsOf(noCode.chartXml)).not.toContain("<c:numFmt");
    const empty = writeChart(
      makeChart({ dataLabels: { showValue: true, numberFormat: { formatCode: "" } } }),
      "Sheet1",
    );
    expect(dLblsOf(empty.chartXml)).not.toContain("<c:numFmt");
  });

  it("places <c:numFmt> at the head of the CT_DLbls sequence (before <c:dLblPos> and the show* toggles)", () => {
    // CT_DLbls schema sequence: dLbl* -> numFmt? -> spPr? -> txPr? ->
    // dLblPos? -> showLegendKey -> showVal -> ...
    const result = writeChart(
      makeChart({
        dataLabels: {
          showValue: true,
          position: "outEnd",
          showLegendKey: true,
          numberFormat: { formatCode: "0.00%" },
        },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    const numFmtIdx = dLbls.indexOf("<c:numFmt");
    const dLblPosIdx = dLbls.indexOf("<c:dLblPos");
    const showLegKeyIdx = dLbls.indexOf("<c:showLegendKey");
    const showValIdx = dLbls.indexOf("<c:showVal");
    expect(numFmtIdx).toBeGreaterThan(0);
    expect(numFmtIdx).toBeLessThan(dLblPosIdx);
    expect(dLblPosIdx).toBeLessThan(showLegKeyIdx);
    expect(showLegKeyIdx).toBeLessThan(showValIdx);
  });

  it("threads numberFormat through a series-level <c:dLbls>", () => {
    const result = writeChart(
      makeChart({
        series: [
          {
            name: "S1",
            values: "B2:B4",
            dataLabels: { showValue: true, numberFormat: { formatCode: "0.0" } },
          },
        ],
      }),
      "Sheet1",
    );
    const xml = result.chartXml;
    const serStart = xml.indexOf("<c:ser>");
    const serEnd = xml.indexOf("</c:ser>");
    const inner = xml.slice(serStart, serEnd);
    expect(inner).toContain('<c:numFmt formatCode="0.0" sourceLinked="0"/>');
  });

  it("threads numberFormat through pie / line / scatter chart families", () => {
    for (const type of ["pie", "line", "scatter"] as const) {
      const result = writeChart(
        makeChart({
          type,
          dataLabels: { showValue: true, numberFormat: { formatCode: "0.00%" } },
        }),
        "Sheet1",
      );
      const dLbls = dLblsOf(result.chartXml);
      expect(dLbls).toContain('<c:numFmt formatCode="0.00%" sourceLinked="0"/>');
    }
  });

  it("emits exactly one <c:numFmt> per data-labels block", () => {
    const result = writeChart(
      makeChart({ dataLabels: { showValue: true, numberFormat: { formatCode: "0.00" } } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect((dLbls.match(/<c:numFmt/g) ?? []).length).toBe(1);
  });

  it("escapes XML-special characters in the formatCode attribute", () => {
    // Format codes legitimately contain `<`, `>`, `&`, `"` — e.g.
    // accounting formats and conditional brackets like
    // `[<=9999999]###-####;(###) ###-####`. The writer must emit
    // them through the standard XML attribute escape path.
    const result = writeChart(
      makeChart({
        dataLabels: {
          showValue: true,
          numberFormat: { formatCode: '[<=9999999]"$"#,##0' },
        },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain("&lt;");
    expect(dLbls).toContain("&quot;");
  });

  it("collapses non-string formatCode inputs to no emission", () => {
    const result = writeChart(
      makeChart({
        dataLabels: {
          showValue: true,
          numberFormat: { formatCode: 42 as unknown as string },
        },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).not.toContain("<c:numFmt");
  });

  it("collapses non-boolean sourceLinked inputs to the default 0", () => {
    const result = writeChart(
      makeChart({
        dataLabels: {
          showValue: true,
          numberFormat: {
            formatCode: "0.00",
            sourceLinked: "yes" as unknown as boolean,
          },
        },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:numFmt formatCode="0.00" sourceLinked="0"/>');
  });

  it("round-trips a chart with numberFormat through parseChart", () => {
    const written = writeChart(
      makeChart({ dataLabels: { showValue: true, numberFormat: { formatCode: "0.00%" } } }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataLabels?.numberFormat).toEqual({ formatCode: "0.00%" });
    expect(reparsed?.dataLabels?.showValue).toBe(true);
  });

  it("round-trips numberFormat with sourceLinked=true through parseChart", () => {
    const written = writeChart(
      makeChart({
        dataLabels: {
          showValue: true,
          numberFormat: { formatCode: "General", sourceLinked: true },
        },
      }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataLabels?.numberFormat).toEqual({
      formatCode: "General",
      sourceLinked: true,
    });
  });

  it("end-to-end: writeXlsx packages a chart with numberFormat into chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            dataLabels: { showValue: true, numberFormat: { formatCode: "$#,##0" } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    const dLbls = chartXml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/)![0];
    expect(dLbls).toContain('<c:numFmt formatCode="$#,##0" sourceLinked="0"/>');
  });
});

// ── writeChart — axis noMultiLvlLbl ──────────────────────────────────

describe("writeChart — axis noMultiLvlLbl", () => {
  it('emits <c:noMultiLvlLbl val="1"/> on the category axis when the override is true', () => {
    const result = writeChart(makeChart({ axes: { x: { noMultiLvlLbl: true } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:noMultiLvlLbl val="1"');
    expect(catAxBlock).not.toContain('c:noMultiLvlLbl val="0"');
  });

  it('emits the OOXML default <c:noMultiLvlLbl val="0"/> when the field is unset', () => {
    // Excel's reference serialization always emits `<c:noMultiLvlLbl val="0"/>`,
    // so the writer keeps that contract on a stock chart even though the
    // parser collapses `0` to undefined on the read side.
    const result = writeChart(makeChart(), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:noMultiLvlLbl val="0"');
  });

  it("emits the default when the override is explicitly false", () => {
    // Pinning the default has the same wire effect as omitting the
    // field — the OOXML default `false` round-trips identically with
    // absence.
    const result = writeChart(makeChart({ axes: { x: { noMultiLvlLbl: false } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:noMultiLvlLbl val="0"');
  });

  it("emits exactly one <c:noMultiLvlLbl> element per category axis", () => {
    const result = writeChart(makeChart({ axes: { x: { noMultiLvlLbl: true } } }), "Sheet1");
    expect((result.chartXml.match(/c:noMultiLvlLbl/g) ?? []).length).toBe(1);
  });

  it("threads the override through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(
        makeChart({ type, axes: { x: { noMultiLvlLbl: true } } }),
        "Sheet1",
      );
      expect(result.chartXml).toContain('c:noMultiLvlLbl val="1"');
    }
  });

  it("ignores the override on scatter charts (both axes are value axes)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { noMultiLvlLbl: true } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:noMultiLvlLbl");
  });

  it("ignores the override on pie / doughnut charts (no axes at all)", () => {
    const pie = writeChart(
      makeChart({ type: "pie", axes: { x: { noMultiLvlLbl: true } } }),
      "Sheet1",
    );
    expect(pie.chartXml).not.toContain("c:noMultiLvlLbl");
    const dough = writeChart(
      makeChart({ type: "doughnut", axes: { x: { noMultiLvlLbl: true } } }),
      "Sheet1",
    );
    expect(dough.chartXml).not.toContain("c:noMultiLvlLbl");
  });

  it("does not emit noMultiLvlLbl on the value axis", () => {
    // The model surfaces the flag only on `axes.x`; setting it via
    // `axes.y` is impossible at the type level. This test pins the
    // negative case for the writer: a valAx never carries the element.
    const result = writeChart(makeChart({ axes: { x: { noMultiLvlLbl: true } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).not.toContain("c:noMultiLvlLbl");
  });

  it("places noMultiLvlLbl after lblOffset / tickLblSkip / tickMarkSkip per the OOXML schema", () => {
    // CT_CatAx: ... lblOffset -> tickLblSkip? -> tickMarkSkip? -> noMultiLvlLbl.
    const result = writeChart(
      makeChart({
        axes: { x: { tickLblSkip: 3, tickMarkSkip: 5, noMultiLvlLbl: true } },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const lblOffsetIdx = catAxBlock.indexOf("c:lblOffset");
    const tickLblSkipIdx = catAxBlock.indexOf("c:tickLblSkip");
    const tickMarkSkipIdx = catAxBlock.indexOf("c:tickMarkSkip");
    const noMultiLvlLblIdx = catAxBlock.indexOf("c:noMultiLvlLbl");
    expect(lblOffsetIdx).toBeGreaterThan(0);
    expect(tickLblSkipIdx).toBeGreaterThan(lblOffsetIdx);
    expect(tickMarkSkipIdx).toBeGreaterThan(tickLblSkipIdx);
    expect(noMultiLvlLblIdx).toBeGreaterThan(tickMarkSkipIdx);
  });

  it("ignores non-boolean noMultiLvlLbl values (falls back to default 0)", () => {
    // Match how `legendOverlay` / `roundedCorners` / axis `hidden` treat
    // their inputs: only literal `true` produces the non-default. A
    // stray non-boolean collapses to the default.
    const result = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ axes: { x: { noMultiLvlLbl: "yes" as any } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:noMultiLvlLbl val="0"');
  });

  it("round-trips a non-default noMultiLvlLbl through parseChart", () => {
    const written = writeChart(
      makeChart({ axes: { x: { noMultiLvlLbl: true } } }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.noMultiLvlLbl).toBe(true);
  });

  it("collapses a defaulted noMultiLvlLbl round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toBeUndefined();
  });

  it("end-to-end: writeXlsx packages the flag into chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            axes: { x: { noMultiLvlLbl: true } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('c:noMultiLvlLbl val="1"');
  });
});

// ── writeChart — axis auto ───────────────────────────────────────────

describe("writeChart — axis auto", () => {
  it('emits the OOXML default <c:auto val="1"/> on the category axis when the field is unset', () => {
    // Excel's reference serialization always emits `<c:auto val="1"/>` on
    // every category axis, so the writer keeps that contract on a stock
    // chart even though the parser collapses `1` to undefined on the
    // read side.
    const result = writeChart(makeChart(), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:auto val="1"');
    expect(catAxBlock).not.toContain('c:auto val="0"');
  });

  it('emits <c:auto val="0"/> on the category axis when the override is false', () => {
    const result = writeChart(makeChart({ axes: { x: { auto: false } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:auto val="0"');
    expect(catAxBlock).not.toContain('c:auto val="1"');
  });

  it("emits the default when the override is explicitly true", () => {
    // Pinning the default has the same wire effect as omitting the field —
    // the OOXML default `true` round-trips identically with absence.
    const result = writeChart(makeChart({ axes: { x: { auto: true } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:auto val="1"');
  });

  it("emits exactly one <c:auto> element per category axis", () => {
    const result = writeChart(makeChart({ axes: { x: { auto: false } } }), "Sheet1");
    // Use a precise pattern so `<c:autoTitleDeleted>` does not collide
    // with the catAx `<c:auto val=".."/>` element.
    expect((result.chartXml.match(/<c:auto val="/g) ?? []).length).toBe(1);
  });

  it("threads the override through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, axes: { x: { auto: false } } }), "Sheet1");
      expect(result.chartXml).toContain('c:auto val="0"');
    }
  });

  it("ignores the override on scatter charts (both axes are value axes)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { auto: false } },
      }),
      "Sheet1",
    );
    // `<c:autoTitleDeleted>` is always emitted on `<c:chart>`, so check
    // the precise CT_CatAx form to confirm the catAx element never
    // surfaces on a chart whose axes are both `<c:valAx>`.
    expect(result.chartXml).not.toContain("<c:auto val=");
  });

  it("ignores the override on pie / doughnut charts (no axes at all)", () => {
    const pie = writeChart(makeChart({ type: "pie", axes: { x: { auto: false } } }), "Sheet1");
    expect(pie.chartXml).not.toContain("<c:auto val=");
    const dough = writeChart(
      makeChart({ type: "doughnut", axes: { x: { auto: false } } }),
      "Sheet1",
    );
    expect(dough.chartXml).not.toContain("<c:auto val=");
  });

  it("does not emit auto on the value axis", () => {
    // The model surfaces the flag only on `axes.x`; setting it via
    // `axes.y` is impossible at the type level. This test pins the
    // negative case for the writer: a valAx never carries the element.
    const result = writeChart(makeChart({ axes: { x: { auto: false } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).not.toContain("<c:auto");
  });

  it("places auto between crosses and lblAlgn per the OOXML schema", () => {
    // CT_CatAx: ... crossAx -> crosses? -> auto -> lblAlgn -> lblOffset.
    const result = writeChart(
      makeChart({
        axes: { x: { crosses: "max", auto: false, lblAlgn: "l" } },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const crossesIdx = catAxBlock.indexOf("c:crosses");
    const autoIdx = catAxBlock.indexOf("c:auto");
    const lblAlgnIdx = catAxBlock.indexOf("c:lblAlgn");
    expect(crossesIdx).toBeGreaterThan(0);
    expect(autoIdx).toBeGreaterThan(crossesIdx);
    expect(lblAlgnIdx).toBeGreaterThan(autoIdx);
  });

  it("ignores non-boolean auto values (falls back to default 1)", () => {
    // Match how `legendOverlay` / `roundedCorners` / axis `hidden` treat
    // their inputs: only literal `false` produces the non-default `0`. A
    // stray non-boolean collapses to the default.
    const result = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ axes: { x: { auto: "no" as any } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:auto val="1"');
  });

  it("round-trips a non-default auto through parseChart", () => {
    const written = writeChart(makeChart({ axes: { x: { auto: false } } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.auto).toBe(false);
  });

  it("collapses a defaulted auto round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toBeUndefined();
  });

  it("end-to-end: writeXlsx packages the flag into chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            axes: { x: { auto: false } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('c:auto val="0"');
  });
});

// ── writeChart — title overlay ───────────────────────────────────────

describe("writeChart — titleOverlay", () => {
  it('emits <c:overlay val="0"/> inside <c:title> when the field is unset (OOXML default)', () => {
    // The writer always emits the element so the rendered intent is
    // explicit on roundtrip — Excel itself includes it in every
    // reference title serialization.
    const result = writeChart(makeChart(), "Sheet1");
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title).toContain('c:overlay val="0"');
    expect(title).not.toContain('c:overlay val="1"');
  });

  it("threads titleOverlay=true through to <c:title>", () => {
    // true is the non-default — Excel's "Show the title without
    // overlapping the chart" toggle off (the title is drawn on top of
    // the plot area).
    const result = writeChart(makeChart({ titleOverlay: true }), "Sheet1");
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title).toContain('c:overlay val="1"');
    expect(title).not.toContain('c:overlay val="0"');
  });

  it("threads titleOverlay=false through to <c:title>", () => {
    // Setting the OOXML default explicitly produces the same wire shape
    // as omitting the field — the element is always emitted.
    const result = writeChart(makeChart({ titleOverlay: false }), "Sheet1");
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title).toContain('c:overlay val="0"');
  });

  it("places <c:overlay> after <c:tx> inside <c:title> (CT_Title order)", () => {
    // CT_Title sequence: tx?, layout?, overlay?, ...
    const result = writeChart(makeChart({ titleOverlay: true }), "Sheet1");
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title.indexOf("c:tx")).toBeLessThan(title.indexOf("c:overlay"));
  });

  it("only emits <c:overlay> once inside <c:title> even on a chart that overrides it", () => {
    // Guard against any regression that would double-emit the element
    // (e.g. one hardcoded copy plus a dynamic one). Scope the count to
    // the title — the legend also carries its own `<c:overlay>`.
    const result = writeChart(makeChart({ titleOverlay: true }), "Sheet1");
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    const occurrences = title.match(/c:overlay/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("does not emit any <c:title> when the chart has no title", () => {
    // No title means no title block to host the overlay flag — the
    // writer suppresses the entire `<c:title>` element. The chart still
    // emits `<c:autoTitleDeleted val="1"/>` so the picker shows blank.
    const result = writeChart(makeChart({ title: undefined, titleOverlay: true }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:title>");
    // The legend still carries its own `<c:overlay>`; the chart-level
    // title block has none.
    expect(result.chartXml).toContain('c:autoTitleDeleted val="1"');
  });

  it("does not emit any <c:title> when showTitle=false even with titleOverlay", () => {
    // `showTitle: false` suppresses the title block entirely — the
    // writer drops the inherited overlay flag rather than emit a stray
    // overlay child Excel would never read.
    const result = writeChart(makeChart({ showTitle: false, titleOverlay: true }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:title>");
    expect(result.chartXml).toContain('c:autoTitleDeleted val="1"');
  });

  it("threads titleOverlay through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, titleOverlay: true }), "Sheet1");
      const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
      expect(title).toContain('c:overlay val="1"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        titleOverlay: true,
      }),
      "Sheet1",
    );
    const title = scatter.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title).toContain('c:overlay val="1"');
  });

  it("composes independently with legendOverlay", () => {
    // The two flags live on different parents (`<c:title>` vs
    // `<c:legend>`); pinning one must not change the other.
    const result = writeChart(makeChart({ titleOverlay: true, legendOverlay: false }), "Sheet1");
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(title).toContain('c:overlay val="1"');
    expect(legend).toContain('c:overlay val="0"');
  });

  it("round-trips a non-default titleOverlay value through parseChart", () => {
    // A chart with titleOverlay=true should re-parse into a Chart whose
    // `titleOverlay` field is `true` (not collapsed to undefined since
    // true is not the OOXML default).
    const written = writeChart(makeChart({ titleOverlay: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.titleOverlay).toBe(true);
  });

  it("collapses a defaulted titleOverlay round-trip back to undefined", () => {
    // A fresh chart (titleOverlay omitted) writes `0` and re-parses to
    // undefined — absence and the OOXML default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.titleOverlay).toBeUndefined();
  });

  it("collapses an explicit titleOverlay=false round-trip back to undefined", () => {
    // Pinning the OOXML default also collapses on read, so a template
    // that explicitly emits `<c:overlay val="0"/>` is treated the same
    // as one that omits the field.
    const written = writeChart(makeChart({ titleOverlay: false }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.titleOverlay).toBeUndefined();
  });

  it("ignores non-boolean titleOverlay values", () => {
    // Match how `legendOverlay` / `roundedCorners` / axis hidden treat
    // their inputs: only literal `true` produces the non-default. A
    // stray non-boolean (e.g. truthy string) collapses to the default.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = writeChart(makeChart({ titleOverlay: "yes" as any }), "Sheet1");
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title).toContain('c:overlay val="0"');
  });

  it("survives a writeXlsx round trip — titleOverlay lands in the packaged chart XML", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            titleOverlay: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    const title = chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title).toContain('c:overlay val="1"');
  });
});

// ── writeChart — axis crosses / crossesAt ────────────────────────────

describe("writeChart — axis crosses / crossesAt", () => {
  it('emits the OOXML default <c:crosses val="autoZero"/> on every axis when unset', () => {
    // Excel's reference serialization always pins `<c:crosses val="autoZero"/>`
    // on every axis, so the writer keeps that contract on a stock chart even
    // though the parser collapses the default to undefined on read.
    const result = writeChart(makeChart(), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).toContain('c:crosses val="autoZero"');
    expect(valAxBlock).toContain('c:crosses val="autoZero"');
    expect(catAxBlock).not.toContain("c:crossesAt");
    expect(valAxBlock).not.toContain("c:crossesAt");
  });

  it('emits a non-default semantic crosses="min" on the category axis', () => {
    const result = writeChart(makeChart({ axes: { x: { crosses: "min" } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:crosses val="min"');
    expect(catAxBlock).not.toContain('c:crosses val="autoZero"');
  });

  it('emits semantic crosses="max" on the value axis', () => {
    const result = writeChart(makeChart({ axes: { y: { crosses: "max" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('c:crosses val="max"');
  });

  it('falls back to the default when crosses="autoZero" is set explicitly', () => {
    // Pinning the default has the same wire effect as omitting the field.
    const result = writeChart(makeChart({ axes: { x: { crosses: "autoZero" } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('c:crosses val="autoZero"');
  });

  it("emits <c:crossesAt> in place of <c:crosses> when the numeric pin is set", () => {
    const result = writeChart(makeChart({ axes: { y: { crossesAt: 50 } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('c:crossesAt val="50"');
    expect(valAxBlock).not.toContain("c:crosses ");
    expect(valAxBlock).not.toContain("<c:crosses/>");
  });

  it("preserves crossesAt=0 (distinct from the autoZero default)", () => {
    // `crossesAt: 0` pins the crossing point to the numeric value zero,
    // distinct from `crosses: "autoZero"` which defers to Excel's
    // auto-placement. The writer must emit `<c:crossesAt val="0"/>`,
    // not collapse to the semantic default.
    const result = writeChart(makeChart({ axes: { y: { crossesAt: 0 } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('c:crossesAt val="0"');
    expect(valAxBlock).not.toContain("c:crosses ");
  });

  it("emits a negative crossesAt verbatim", () => {
    const result = writeChart(makeChart({ axes: { y: { crossesAt: -25.5 } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('c:crossesAt val="-25.5"');
  });

  it("prefers crossesAt over crosses when both are set (XSD choice)", () => {
    // The OOXML schema places <c:crosses> and <c:crossesAt> in an XSD
    // choice — only one may legally appear. The writer favours the
    // numeric pin, mirroring the reader's preference on malformed input.
    const result = writeChart(
      makeChart({ axes: { y: { crosses: "max", crossesAt: 7 } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('c:crossesAt val="7"');
    expect(valAxBlock).not.toContain("c:crosses ");
    expect(valAxBlock).not.toContain('c:crosses val="max"');
  });

  it("falls back to crosses when crossesAt is non-finite", () => {
    // NaN / Infinity inputs drop through to the semantic crosses
    // (or to the autoZero default when crosses is also unset).
    const result = writeChart(
      makeChart({
        axes: { y: { crosses: "min", crossesAt: Number.NaN } },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('c:crosses val="min"');
    expect(valAxBlock).not.toContain("c:crossesAt");
  });

  it("ignores unknown semantic tokens (falls back to autoZero default)", () => {
    const result = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ axes: { y: { crosses: "middle" as any } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('c:crosses val="autoZero"');
  });

  it("emits exactly one crosses-or-crossesAt element per axis", () => {
    const result = writeChart(
      makeChart({
        axes: { x: { crosses: "min" }, y: { crossesAt: 10 } },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect((catAxBlock.match(/c:crosses\b/g) ?? []).length).toBe(1);
    expect((catAxBlock.match(/c:crossesAt\b/g) ?? []).length).toBe(0);
    expect((valAxBlock.match(/c:crosses\b/g) ?? []).length).toBe(0);
    expect((valAxBlock.match(/c:crossesAt\b/g) ?? []).length).toBe(1);
  });

  it("threads the override through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, axes: { y: { crosses: "max" } } }), "Sheet1");
      expect(result.chartXml).toContain('c:crosses val="max"');
    }
  });

  it("threads the override through scatter charts (both axes are valAx)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { crossesAt: 3.14 }, y: { crosses: "min" } },
      }),
      "Sheet1",
    );
    // Scatter emits two <c:valAx> elements — first is the X axis, second
    // is the Y axis.
    const valAxes = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    expect(valAxes).toHaveLength(2);
    expect(valAxes[0]).toContain('c:crossesAt val="3.14"');
    expect(valAxes[1]).toContain('c:crosses val="min"');
  });

  it("ignores the override on pie / doughnut charts (no axes at all)", () => {
    const pie = writeChart(makeChart({ type: "pie", axes: { y: { crosses: "min" } } }), "Sheet1");
    expect(pie.chartXml).not.toContain("c:crosses");
    expect(pie.chartXml).not.toContain("c:crossesAt");
    const dough = writeChart(
      makeChart({ type: "doughnut", axes: { y: { crossesAt: 5 } } }),
      "Sheet1",
    );
    expect(dough.chartXml).not.toContain("c:crosses");
    expect(dough.chartXml).not.toContain("c:crossesAt");
  });

  it("places crosses after crossAx per the OOXML schema (CT_CatAx / CT_ValAx)", () => {
    // OOXML CT_CatAx / CT_ValAx: ... → tickLblPos → crossAx → (crosses
    // | crossesAt) → ... The writer's emit order pins crossAx first.
    const result = writeChart(makeChart({ axes: { y: { crossesAt: 42 } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const crossAxIdx = valAxBlock.indexOf("c:crossAx");
    const crossesAtIdx = valAxBlock.indexOf("c:crossesAt");
    expect(crossAxIdx).toBeGreaterThan(0);
    expect(crossesAtIdx).toBeGreaterThan(crossAxIdx);
  });

  it("round-trips a non-default semantic crosses through parseChart", () => {
    const written = writeChart(makeChart({ axes: { y: { crosses: "max" } } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.y?.crosses).toBe("max");
    expect(reparsed?.axes?.y?.crossesAt).toBeUndefined();
  });

  it("round-trips a numeric crossesAt through parseChart", () => {
    const written = writeChart(makeChart({ axes: { y: { crossesAt: -3.5 } } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.y?.crossesAt).toBe(-3.5);
    expect(reparsed?.axes?.y?.crosses).toBeUndefined();
  });

  it("collapses a defaulted crosses round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toBeUndefined();
  });

  it("end-to-end: writeXlsx packages the crosses pin into chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            axes: { x: { crosses: "min" }, y: { crossesAt: 0 } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('c:crosses val="min"');
    expect(chartXml).toContain('c:crossesAt val="0"');
  });
});

// ── Drop / hi-low lines ──────────────────────────────────────────────

describe("writeChart — drop lines", () => {
  it("omits <c:dropLines> on a line chart with dropLines unset (default)", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:dropLines");
  });

  it("omits <c:dropLines> on a line chart when dropLines is explicitly false", () => {
    // The writer treats absence and `false` identically — both produce
    // no element, matching Excel's reference serialization.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: false,
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:dropLines");
  });

  it("emits <c:dropLines/> on a line chart when dropLines is true", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:dropLines/>");
  });

  it("emits <c:dropLines/> on an area chart when dropLines is true", () => {
    const result = writeChart(
      makeChart({
        type: "area",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:dropLines/>");
  });

  it("ignores dropLines on chart kinds whose schema rejects the element", () => {
    // CT_BarChart / CT_PieChart / CT_DoughnutChart / CT_ScatterChart
    // all reject `<c:dropLines>` per OOXML. Setting the flag on these
    // families must not leak the element into the output.
    const cases: Array<["column" | "bar" | "pie" | "doughnut" | "scatter"]> = [
      ["column"],
      ["bar"],
      ["pie"],
      ["doughnut"],
      ["scatter"],
    ];
    for (const [type] of cases) {
      const result = writeChart(
        makeChart({
          type,
          series: [{ values: "B2:B4", categories: "A2:A4" }],
          dropLines: true,
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:dropLines");
    }
  });

  it("non-boolean dropLines values collapse to absence (only literal true emits)", () => {
    // Mirrors the title/legend overlay writers — the resolver does not
    // coerce its inputs. Truthy strings, numbers, etc. drop to the
    // default of no element.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        dropLines: 1 as any,
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:dropLines");
  });

  it("places <c:dropLines> after <c:dLbls> and before <c:marker> inside <c:lineChart>", () => {
    // CT_LineChart sequence: grouping, varyColors?, ser*, dLbls?,
    // dropLines?, hiLowLines?, upDownBars?, marker?, axId, axId. We
    // assert the `<c:dropLines>` slot lands after `<c:dLbls>` (when
    // any data labels are emitted) and before `<c:marker>`.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dataLabels: { showValue: true },
        dropLines: true,
      }),
      "Sheet1",
    );
    const lineBlock = result.chartXml.match(/<c:lineChart>[\s\S]*?<\/c:lineChart>/)![0];
    const dLblsIdx = lineBlock.indexOf("<c:dLbls>");
    const dropIdx = lineBlock.indexOf("<c:dropLines/>");
    const markerIdx = lineBlock.indexOf("<c:marker ");
    expect(dLblsIdx).toBeGreaterThan(-1);
    expect(dropIdx).toBeGreaterThan(dLblsIdx);
    expect(markerIdx).toBeGreaterThan(dropIdx);
  });

  it("places <c:dropLines> before <c:axId> inside <c:areaChart>", () => {
    // CT_AreaChart sequence: grouping?, varyColors?, ser*, dLbls?,
    // dropLines?, axId, axId. The `<c:dropLines>` slot lands right
    // before the first `<c:axId>`.
    const result = writeChart(
      makeChart({
        type: "area",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: true,
      }),
      "Sheet1",
    );
    const areaBlock = result.chartXml.match(/<c:areaChart>[\s\S]*?<\/c:areaChart>/)![0];
    const dropIdx = areaBlock.indexOf("<c:dropLines/>");
    const axIdx = areaBlock.indexOf("<c:axId ");
    expect(dropIdx).toBeGreaterThan(-1);
    expect(axIdx).toBeGreaterThan(dropIdx);
  });

  it("round-trips dropLines through parseChart (line)", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: true,
      }),
      "Sheet1",
    );
    expect(parseChart(result.chartXml)?.dropLines).toBe(true);
  });

  it("round-trips dropLines through parseChart (area)", () => {
    const result = writeChart(
      makeChart({
        type: "area",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: true,
      }),
      "Sheet1",
    );
    expect(parseChart(result.chartXml)?.dropLines).toBe(true);
  });

  it("survives a writeXlsx round trip — dropLines lands in the packaged chart XML", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "line",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            dropLines: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:dropLines/>");
  });
});

describe("writeChart — high-low lines", () => {
  it("omits <c:hiLowLines> on a line chart with hiLowLines unset (default)", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:hiLowLines");
  });

  it("emits <c:hiLowLines/> on a line chart when hiLowLines is true", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        hiLowLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:hiLowLines/>");
  });

  it("ignores hiLowLines on an area chart (no slot in the OOXML schema)", () => {
    // CT_AreaChart rejects <c:hiLowLines>. The area writer must not
    // emit the element even when the caller pins the flag.
    const result = writeChart(
      makeChart({
        type: "area",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        hiLowLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:hiLowLines");
  });

  it("ignores hiLowLines on bar / column / pie / doughnut / scatter charts", () => {
    const cases: Array<["column" | "bar" | "pie" | "doughnut" | "scatter"]> = [
      ["column"],
      ["bar"],
      ["pie"],
      ["doughnut"],
      ["scatter"],
    ];
    for (const [type] of cases) {
      const result = writeChart(
        makeChart({
          type,
          series: [{ values: "B2:B4", categories: "A2:A4" }],
          hiLowLines: true,
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:hiLowLines");
    }
  });

  it("non-boolean hiLowLines values collapse to absence (only literal true emits)", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        hiLowLines: 1 as any,
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:hiLowLines");
  });

  it("places <c:hiLowLines> after <c:dropLines> and before <c:marker> inside <c:lineChart>", () => {
    // CT_LineChart sequence places dropLines before hiLowLines; both
    // appear before the chart-level <c:marker> toggle. Verify the slot
    // ordering on a chart that pins both.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: true,
        hiLowLines: true,
      }),
      "Sheet1",
    );
    const lineBlock = result.chartXml.match(/<c:lineChart>[\s\S]*?<\/c:lineChart>/)![0];
    const dropIdx = lineBlock.indexOf("<c:dropLines/>");
    const hiLowIdx = lineBlock.indexOf("<c:hiLowLines/>");
    const markerIdx = lineBlock.indexOf("<c:marker ");
    expect(dropIdx).toBeGreaterThan(-1);
    expect(hiLowIdx).toBeGreaterThan(dropIdx);
    expect(markerIdx).toBeGreaterThan(hiLowIdx);
  });

  it("places <c:hiLowLines> before <c:marker> when <c:dropLines> is absent", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        hiLowLines: true,
      }),
      "Sheet1",
    );
    const lineBlock = result.chartXml.match(/<c:lineChart>[\s\S]*?<\/c:lineChart>/)![0];
    const hiLowIdx = lineBlock.indexOf("<c:hiLowLines/>");
    const markerIdx = lineBlock.indexOf("<c:marker ");
    expect(hiLowIdx).toBeGreaterThan(-1);
    expect(markerIdx).toBeGreaterThan(hiLowIdx);
  });

  it("round-trips hiLowLines through parseChart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        hiLowLines: true,
      }),
      "Sheet1",
    );
    expect(parseChart(result.chartXml)?.hiLowLines).toBe(true);
  });

  it("survives a writeXlsx round trip — hiLowLines lands in the packaged chart XML", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "line",
            series: [
              { name: "High", values: "B2:B3", categories: "A2:A3" },
              { name: "Low", values: "C2:C3", categories: "A2:A3" },
            ],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            hiLowLines: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:hiLowLines/>");
  });

  it("round-trips both dropLines and hiLowLines together via parseChart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        dropLines: true,
        hiLowLines: true,
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    expect(parsed?.dropLines).toBe(true);
    expect(parsed?.hiLowLines).toBe(true);
  });
});

// ── writeChart — series lines ────────────────────────────────────────

describe("writeChart — series lines", () => {
  it("omits <c:serLines> on a column chart with serLines unset (default)", () => {
    const result = writeChart(
      makeChart({
        type: "column",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:serLines");
  });

  it("omits <c:serLines> on a column chart when serLines is explicitly false", () => {
    // The writer treats absence and `false` identically — both produce
    // no element, matching Excel's reference serialization.
    const result = writeChart(
      makeChart({
        type: "column",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: false,
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:serLines");
  });

  it("emits <c:serLines/> on a stacked column chart when serLines is true", () => {
    const result = writeChart(
      makeChart({
        type: "column",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:serLines/>");
  });

  it("emits <c:serLines/> on a stacked bar chart when serLines is true", () => {
    const result = writeChart(
      makeChart({
        type: "bar",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:serLines/>");
  });

  it("emits <c:serLines/> on a percentStacked bar chart when serLines is true", () => {
    const result = writeChart(
      makeChart({
        type: "bar",
        barGrouping: "percentStacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:serLines/>");
  });

  it("emits <c:serLines/> on a clustered bar chart when serLines is true", () => {
    // The OOXML element pins regardless of grouping; Excel only renders
    // the connectors on stacked groupings, but the writer honours the
    // toggle on a clustered chart (matches Excel's own behavior — the
    // element pins, the renderer paints nothing).
    const result = writeChart(
      makeChart({
        type: "column",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:serLines/>");
  });

  it("ignores serLines on chart kinds whose schema rejects the element", () => {
    // CT_LineChart / CT_AreaChart / CT_PieChart / CT_DoughnutChart /
    // CT_ScatterChart all reject `<c:serLines>` per OOXML. Setting the
    // flag on these families must not leak the element into the output.
    const cases: Array<["line" | "area" | "pie" | "doughnut" | "scatter"]> = [
      ["line"],
      ["area"],
      ["pie"],
      ["doughnut"],
      ["scatter"],
    ];
    for (const [type] of cases) {
      const result = writeChart(
        makeChart({
          type,
          series: [{ values: "B2:B4", categories: "A2:A4" }],
          serLines: true,
        }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:serLines");
    }
  });

  it("non-boolean serLines values collapse to absence (only literal true emits)", () => {
    // Mirrors the dropLines / hiLowLines writers — the resolver does
    // not coerce its inputs. Truthy strings, numbers, etc. drop to the
    // default of no element.
    const result = writeChart(
      makeChart({
        type: "column",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        serLines: 1 as any,
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:serLines");
  });

  it("places <c:serLines> after <c:overlap> and before <c:axId> inside <c:barChart>", () => {
    // CT_BarChart sequence: barDir, grouping, varyColors?, ser*,
    // dLbls?, gapWidth?, overlap?, serLines*, axId+. The
    // `<c:serLines>` slot lands after `<c:overlap>` (when emitted on a
    // stacked chart) and before the first `<c:axId>`.
    const result = writeChart(
      makeChart({
        type: "column",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    const barBlock = result.chartXml.match(/<c:barChart>[\s\S]*?<\/c:barChart>/)![0];
    const overlapIdx = barBlock.indexOf("<c:overlap ");
    const serIdx = barBlock.indexOf("<c:serLines/>");
    const axIdx = barBlock.indexOf("<c:axId ");
    expect(overlapIdx).toBeGreaterThan(-1);
    expect(serIdx).toBeGreaterThan(overlapIdx);
    expect(axIdx).toBeGreaterThan(serIdx);
  });

  it("places <c:serLines> after <c:gapWidth> on a clustered bar chart", () => {
    // On a clustered bar / column chart, the writer emits
    // `<c:gapWidth val="150"/>` and skips `<c:overlap>`; `<c:serLines>`
    // must still land after `<c:gapWidth>` and before the first
    // `<c:axId>`.
    const result = writeChart(
      makeChart({
        type: "column",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    const barBlock = result.chartXml.match(/<c:barChart>[\s\S]*?<\/c:barChart>/)![0];
    const gapWidthIdx = barBlock.indexOf("<c:gapWidth ");
    const serIdx = barBlock.indexOf("<c:serLines/>");
    const axIdx = barBlock.indexOf("<c:axId ");
    expect(gapWidthIdx).toBeGreaterThan(-1);
    expect(serIdx).toBeGreaterThan(gapWidthIdx);
    expect(axIdx).toBeGreaterThan(serIdx);
  });

  it("emits <c:serLines/> exactly once even when called repeatedly", () => {
    // Sanity check that the writer does not duplicate the element on a
    // chart with multiple series.
    const result = writeChart(
      makeChart({
        type: "column",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
          { values: "D2:D4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    expect(result.chartXml.match(/<c:serLines\/>/g)?.length).toBe(1);
  });

  it("round-trips serLines through parseChart (column)", () => {
    const result = writeChart(
      makeChart({
        type: "column",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    expect(parseChart(result.chartXml)?.serLines).toBe(true);
  });

  it("round-trips serLines through parseChart (bar)", () => {
    const result = writeChart(
      makeChart({
        type: "bar",
        barGrouping: "stacked",
        series: [
          { values: "B2:B4", categories: "A2:A4" },
          { values: "C2:C4", categories: "A2:A4" },
        ],
        serLines: true,
      }),
      "Sheet1",
    );
    expect(parseChart(result.chartXml)?.serLines).toBe(true);
  });

  it("survives a writeXlsx round trip — serLines lands in the packaged chart XML", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Q1", "Q2"],
          ["North", 100, 120],
          ["South", 200, 180],
        ],
        charts: [
          {
            type: "column",
            barGrouping: "stacked",
            series: [
              { name: "Q1", values: "B2:B3", categories: "A2:A3" },
              { name: "Q2", values: "C2:C3", categories: "A2:A3" },
            ],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            serLines: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:serLines/>");
  });
});

// ── writeChart — upDownBars ──────────────────────────────────────────

describe("writeChart — upDownBars", () => {
  it("omits <c:upDownBars> when the field is unset (OOXML default)", () => {
    // The OOXML default for <c:upDownBars> on CT_LineChart is absence —
    // Excel's reference serialization for a fresh line chart does not
    // emit the element. The writer mirrors that default by only
    // emitting on an explicit `true`.
    const result = writeChart(makeChart({ type: "line" }), "Sheet1");
    expect(result.chartXml).not.toContain("c:upDownBars");
  });

  it('emits <c:upDownBars> with default <c:gapWidth val="150"/> when upDownBars=true', () => {
    // The schema default for CT_UpDownBars/gapWidth is 150 — Excel's
    // reference serialization emits the element with that gap width.
    const result = writeChart(makeChart({ type: "line", upDownBars: true }), "Sheet1");
    expect(result.chartXml).toContain("<c:upDownBars>");
    expect(result.chartXml).toContain("</c:upDownBars>");
    expect(result.chartXml).toContain('c:gapWidth val="150"');
  });

  it("treats upDownBars=false as absence (no element emitted)", () => {
    // The writer only emits on a literal `true`; `false` collapses to
    // the OOXML default (no element) so a stray `false` from clone
    // resolution does not fabricate an empty up/down bars block.
    const result = writeChart(makeChart({ type: "line", upDownBars: false }), "Sheet1");
    expect(result.chartXml).not.toContain("c:upDownBars");
  });

  it("places <c:upDownBars> before <c:marker> inside <c:lineChart> (OOXML order)", () => {
    // CT_LineChart sequence: ... ser*, dLbls?, dropLines?, hiLowLines?,
    // upDownBars?, marker?, axId+. The schema rejects an out-of-order
    // <c:upDownBars> after <c:marker>, so the writer must place it
    // first.
    const result = writeChart(makeChart({ type: "line", upDownBars: true }), "Sheet1");
    const upDownBarsIdx = result.chartXml.indexOf("c:upDownBars");
    const markerIdx = result.chartXml.indexOf('c:marker val="1"');
    expect(upDownBarsIdx).toBeGreaterThan(0);
    expect(markerIdx).toBeGreaterThan(0);
    expect(upDownBarsIdx).toBeLessThan(markerIdx);
  });

  it("places <c:upDownBars> before <c:axId> inside <c:lineChart> (OOXML order)", () => {
    // The axId pair sits at the tail of CT_LineChart — every optional
    // chart-level child must precede them.
    const result = writeChart(makeChart({ type: "line", upDownBars: true }), "Sheet1");
    const upDownBarsIdx = result.chartXml.indexOf("c:upDownBars");
    const firstAxIdIdx = result.chartXml.indexOf("c:axId");
    expect(upDownBarsIdx).toBeLessThan(firstAxIdIdx);
  });

  it("only emits <c:upDownBars> once even on a chart that pins the flag", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(makeChart({ type: "line", upDownBars: true }), "Sheet1");
    const occurrences = result.chartXml.match(/<c:upDownBars\b/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("ignores upDownBars on bar / column / pie / doughnut / area / scatter chart kinds", () => {
    // The OOXML schema places <c:upDownBars> exclusively on CT_LineChart
    // / CT_Line3DChart / CT_StockChart. The writer never authors the
    // 3D / stock variants, so only `type: "line"` should emit. Every
    // other family must drop the flag silently.
    for (const type of ["bar", "column", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, upDownBars: true }), "Sheet1");
      expect(result.chartXml).not.toContain("c:upDownBars");
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        upDownBars: true,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).not.toContain("c:upDownBars");
  });

  it("round-trips upDownBars=true through parseChart", () => {
    // A line chart with upDownBars=true should re-parse into a Chart
    // whose `upDownBars` field is `true`.
    const written = writeChart(makeChart({ type: "line", upDownBars: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.upDownBars).toBe(true);
  });

  it("collapses an unset upDownBars round-trip back to undefined", () => {
    // A fresh line chart (upDownBars omitted) writes no element and
    // re-parses to undefined — absence and the OOXML default round-trip
    // identically.
    const written = writeChart(makeChart({ type: "line" }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.upDownBars).toBeUndefined();
  });

  it("threads upDownBars through writeXlsx end-to-end packaging", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["High", "Low"],
          [10, 5],
          [12, 6],
          [15, 8],
        ],
        charts: [
          {
            type: "line",
            series: [
              { name: "High", values: "A2:A4" },
              { name: "Low", values: "B2:B4" },
            ],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            upDownBars: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:upDownBars>");
    expect(chartXml).toContain('c:gapWidth val="150"');
    // Re-parse the rendered chart to confirm the flag survives the
    // packaging path.
    const reparsed = parseChart(chartXml);
    expect(reparsed?.upDownBars).toBe(true);
  });
});

// ── writeChart — upDownBars gap width ────────────────────────────────

describe("writeChart — upDownBars gap width", () => {
  it("emits a custom gap width inside <c:upDownBars> when explicitly pinned", () => {
    // The OOXML schema (CT_UpDownBars) places <c:gapWidth> inside the
    // parent element. The writer rounds the value to a whole percent
    // and emits the literal value the caller pinned.
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: 200 }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:upDownBars>");
    expect(result.chartXml).toContain('c:gapWidth val="200"');
    expect(result.chartXml).not.toContain('c:gapWidth val="150"');
  });

  it("emits a tight gap width of 0", () => {
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: 0 }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:gapWidth val="0"');
  });

  it("emits the maximum gap width of 500", () => {
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: 500 }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:gapWidth val="500"');
  });

  it("falls back to the OOXML default 150 when the caller omits the field", () => {
    // The writer always emits <c:gapWidth> when the parent element is
    // emitted — Excel's reference serialization includes the child
    // unconditionally. Absence of the field collapses to the default.
    const result = writeChart(makeChart({ type: "line", upDownBars: true }), "Sheet1");
    expect(result.chartXml).toContain('c:gapWidth val="150"');
  });

  it("rounds non-integer values to the nearest whole percent", () => {
    // Excel's UI accepts integer percentages only. Callers passing a
    // fractional value get the rounded value the writer emits.
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: 174.6 }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:gapWidth val="175"');
  });

  it("falls back to 150 when the value is above the schema cap of 500", () => {
    // The OOXML schema (ST_GapAmount) restricts the value to 0..500;
    // out-of-range values fall back to the default rather than clamp
    // to the schema bound — silently rewriting an `800` to `500`
    // would mislead the caller about what Excel ends up rendering.
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: 800 }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:gapWidth val="150"');
    expect(result.chartXml).not.toContain('c:gapWidth val="800"');
  });

  it("falls back to 150 when the value is below the schema floor of 0", () => {
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: -25 }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:gapWidth val="150"');
  });

  it("falls back to 150 when the value is non-finite", () => {
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      const result = writeChart(
        makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: bad }),
        "Sheet1",
      );
      expect(result.chartXml).toContain('c:gapWidth val="150"');
    }
  });

  it("drops the gap width when upDownBars is unset (no parent element to host it)", () => {
    // The OOXML schema scopes <c:gapWidth> exclusively to <c:upDownBars>.
    // Pinning a gap width without the parent toggle never emits the
    // child — there is no slot for it.
    const result = writeChart(makeChart({ type: "line", upDownBarsGapWidth: 200 }), "Sheet1");
    expect(result.chartXml).not.toContain("c:upDownBars");
    expect(result.chartXml).not.toContain('c:gapWidth val="200"');
  });

  it("drops the gap width on a chart that pins upDownBars=false", () => {
    // upDownBars=false collapses to absence of the parent element, so
    // there is no slot for the child.
    const result = writeChart(
      makeChart({ type: "line", upDownBars: false, upDownBarsGapWidth: 200 }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:upDownBars");
  });

  it("only emits <c:gapWidth> once inside the up/down bars block", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: 200 }),
      "Sheet1",
    );
    const occurrences = result.chartXml.match(/<c:gapWidth\b/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("drops the gap width on bar / column / pie / doughnut / area chart kinds", () => {
    // The bar-family <c:gapWidth> writer is independent — it should not
    // surface when only `upDownBarsGapWidth` is set on a bar chart.
    for (const type of ["bar", "column", "pie", "doughnut", "area"] as const) {
      const result = writeChart(
        makeChart({ type, upDownBars: true, upDownBarsGapWidth: 200 }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("c:upDownBars");
    }
  });

  it("drops the gap width on a scatter chart (CT_ScatterChart rejects upDownBars)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        upDownBars: true,
        upDownBarsGapWidth: 200,
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:upDownBars");
  });

  it("does not collide with the bar-chart gap width on a column chart", () => {
    // A column chart with both gapWidth (bar-family) and
    // upDownBarsGapWidth (line-family) emits only the bar-family value;
    // the line-family field is dropped because the parent element is
    // never emitted on a column chart.
    const result = writeChart(
      makeChart({ type: "column", gapWidth: 75, upDownBarsGapWidth: 200 }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('c:gapWidth val="75"');
    expect(result.chartXml).not.toContain('c:gapWidth val="200"');
  });

  it("round-trips a custom gap width through parseChart", () => {
    const written = writeChart(
      makeChart({ type: "line", upDownBars: true, upDownBarsGapWidth: 200 }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.upDownBars).toBe(true);
    expect(reparsed?.upDownBarsGapWidth).toBe(200);
  });

  it("collapses the OOXML default 150 round-trip back to undefined", () => {
    // Writing upDownBars=true (with no explicit gap width) emits the
    // default 150; the reader then collapses that default back to
    // undefined so absence and the default round-trip identically.
    const written = writeChart(makeChart({ type: "line", upDownBars: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.upDownBars).toBe(true);
    expect(reparsed?.upDownBarsGapWidth).toBeUndefined();
  });

  it("threads the custom gap width through writeXlsx end-to-end packaging", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["High", "Low"],
          [10, 5],
          [12, 6],
          [15, 8],
        ],
        charts: [
          {
            type: "line",
            series: [
              { name: "High", values: "A2:A4" },
              { name: "Low", values: "B2:B4" },
            ],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            upDownBars: true,
            upDownBarsGapWidth: 250,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:upDownBars>");
    expect(chartXml).toContain('c:gapWidth val="250"');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.upDownBars).toBe(true);
    expect(reparsed?.upDownBarsGapWidth).toBe(250);
  });
});

// ── writeChart — axis dispUnits ──────────────────────────────────────

describe("writeChart — axis dispUnits", () => {
  it("omits <c:dispUnits> on a stock chart whose axes pin no preset", () => {
    // Excel's reference serialization for a fresh chart does not emit
    // the element at all — absence collapses to Excel's default
    // "no display unit" state.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:dispUnits");
  });

  it('emits <c:dispUnits><c:builtInUnit val="millions"/></c:dispUnits> on the value axis', () => {
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { unit: "millions" } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:builtInUnit val="millions"/>');
    expect(valAxBlock).toContain("<c:dispUnits>");
    // No <c:dispUnitsLbl> on the default (showLabel omitted).
    expect(valAxBlock).not.toContain("c:dispUnitsLbl");
  });

  it("accepts the ChartAxisDispUnit shorthand string", () => {
    const result = writeChart(makeChart({ axes: { y: { dispUnits: "thousands" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:builtInUnit val="thousands"/>');
  });

  it("emits a bare <c:dispUnitsLbl/> when showLabel is true", () => {
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { unit: "billions", showLabel: true } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:builtInUnit val="billions"/>');
    expect(valAxBlock).toContain("<c:dispUnitsLbl/>");
  });

  it("omits <c:dispUnitsLbl> when showLabel is false / undefined", () => {
    const noFlag = writeChart(
      makeChart({ axes: { y: { dispUnits: { unit: "thousands" } } } }),
      "Sheet1",
    );
    expect(noFlag.chartXml).not.toContain("c:dispUnitsLbl");

    const explicitFalse = writeChart(
      makeChart({ axes: { y: { dispUnits: { unit: "thousands", showLabel: false } } } }),
      "Sheet1",
    );
    expect(explicitFalse.chartXml).not.toContain("c:dispUnitsLbl");
  });

  it("drops an unknown ST_BuiltInUnit token rather than fabricating a value", () => {
    const result = writeChart(
      makeChart({
        // Force the unsafe string past the type guard.
        axes: { y: { dispUnits: { unit: "quintillions" as never } } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:dispUnits");
    expect(result.chartXml).not.toContain("c:builtInUnit");
  });

  it("places <c:dispUnits> after <c:minorUnit> inside <c:valAx> (CT_ValAx order)", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            scale: { min: 0, max: 1_000_000, majorUnit: 250_000, minorUnit: 50_000 },
            dispUnits: "millions",
          },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const minorUnitIdx = valAxBlock.indexOf("c:minorUnit");
    const dispUnitsIdx = valAxBlock.indexOf("c:dispUnits");
    expect(minorUnitIdx).toBeGreaterThan(-1);
    expect(dispUnitsIdx).toBeGreaterThan(minorUnitIdx);
  });

  it("does not emit <c:dispUnits> on the X axis of a bar / column chart (catAx rejects it)", () => {
    // The OOXML schema places <c:dispUnits> exclusively on CT_ValAx, so
    // a stale hint on the X axis of a column chart should silently
    // drop at the writer.
    const result = writeChart(
      makeChart({ type: "column", axes: { x: { dispUnits: "millions" } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).not.toContain("c:dispUnits");
  });

  it("emits <c:dispUnits> on both scatter axes (both are valAx)", () => {
    const scatter: SheetChart = {
      type: "scatter",
      series: [{ name: "S1", values: "B2:B5", categories: "A2:A5" }],
      anchor: { from: { row: 0, col: 0 } },
      axes: {
        x: { dispUnits: "thousands" },
        y: { dispUnits: { unit: "millions", showLabel: true } },
      },
    };
    const result = writeChart(scatter, "Sheet1");
    const valAxBlocks = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    expect(valAxBlocks).toHaveLength(2);
    expect(valAxBlocks[0]).toContain('<c:builtInUnit val="thousands"/>');
    expect(valAxBlocks[0]).not.toContain("c:dispUnitsLbl");
    expect(valAxBlocks[1]).toContain('<c:builtInUnit val="millions"/>');
    expect(valAxBlocks[1]).toContain("<c:dispUnitsLbl/>");
  });

  it("survives a parseChart round-trip on the value axis", () => {
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { unit: "millions", showLabel: true } } } }),
      "Sheet1",
    );
    const reparsed = parseChart(result.chartXml);
    expect(reparsed?.axes?.y?.dispUnits).toEqual({ unit: "millions", showLabel: true });
  });

  it("does not emit <c:dispUnits> on a pie chart (no axes at all)", () => {
    // The writer never builds <c:valAx> for pie / doughnut, so even
    // when the caller pins a value the element should not surface.
    const result = writeChart(
      makeChart({
        type: "pie",
        // Pie charts have no axes; the field is simply ignored.
        axes: { y: { dispUnits: "millions" } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:dispUnits");
  });

  it("only emits <c:dispUnits> once on the value axis", () => {
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { unit: "thousands", showLabel: true } } } }),
      "Sheet1",
    );
    const occurrences = result.chartXml.match(/c:dispUnits>/g) ?? [];
    // Two matches: opening + closing tag of <c:dispUnits>.
    expect(occurrences).toHaveLength(2);
  });

  it("packages the chart end-to-end through writeXlsx", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 1_500_000],
          ["Q2", 2_300_000],
          ["Q3", 3_100_000],
        ],
        charts: [
          {
            type: "column",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            axes: { y: { dispUnits: { unit: "millions", showLabel: true } } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('<c:builtInUnit val="millions"/>');
    expect(chartXml).toContain("<c:dispUnitsLbl/>");
    const reparsed = parseChart(chartXml);
    expect(reparsed?.axes?.y?.dispUnits).toEqual({ unit: "millions", showLabel: true });
  });

  it('emits <c:dispUnits><c:custUnit val=".."/></c:dispUnits> when custUnit is pinned', () => {
    // The OOXML schema's xsd:choice between <c:builtInUnit> and
    // <c:custUnit> means the writer must emit exactly one. Pinning
    // `custUnit` selects the numeric divisor path (Excel's "Display
    // units → Other") instead of the named preset.
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { custUnit: 86400 } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:custUnit val="86400"/>');
    expect(valAxBlock).toContain("<c:dispUnits>");
    expect(valAxBlock).not.toContain("c:builtInUnit");
    expect(valAxBlock).not.toContain("c:dispUnitsLbl");
  });

  it("emits a fractional <c:custUnit> per the OOXML CT_Double schema", () => {
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { custUnit: 2.5 } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:custUnit val="2.5"/>');
  });

  it("emits a bare <c:dispUnitsLbl/> alongside custUnit when showLabel is true", () => {
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { custUnit: 500, showLabel: true } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:custUnit val="500"/>');
    expect(valAxBlock).toContain("<c:dispUnitsLbl/>");
  });

  it("prefers <c:custUnit> over <c:builtInUnit> when both fields are pinned", () => {
    // The OOXML schema's xsd:choice forbids both children, so the writer
    // must pick one. `custUnit` is the more specific element — a caller
    // appending a custom divisor to a cloned source need not manually
    // prune the inherited preset. The reader mirrors this preference.
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { unit: "millions", custUnit: 1500 } } } }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:custUnit val="1500"/>');
    expect(valAxBlock).not.toContain("c:builtInUnit");
  });

  it("falls back to <c:builtInUnit> when custUnit is non-positive / non-finite", () => {
    // The normalizer drops invalid custUnit values; if the caller also
    // pinned a valid `unit` preset, that takes over instead of dropping
    // the entire element.
    for (const bad of [0, -100, Number.NaN, Number.POSITIVE_INFINITY]) {
      const result = writeChart(
        makeChart({ axes: { y: { dispUnits: { unit: "millions", custUnit: bad } } } }),
        "Sheet1",
      );
      const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
      expect(valAxBlock).toContain('<c:builtInUnit val="millions"/>');
      expect(valAxBlock).not.toContain("c:custUnit");
    }
  });

  it("drops <c:dispUnits> entirely when neither unit nor custUnit resolves", () => {
    // A bare <c:dispUnits/> shell with no <c:builtInUnit>/<c:custUnit>
    // child fails Excel's strict validator. The writer skips emission.
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            // Both fields invalid — `unit` is an unknown token, `custUnit`
            // is non-finite.
            dispUnits: { unit: "quintillions" as never, custUnit: -5 },
          },
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:dispUnits");
    expect(result.chartXml).not.toContain("c:builtInUnit");
    expect(result.chartXml).not.toContain("c:custUnit");
  });

  it("does not emit <c:custUnit> on the X axis of a column chart (catAx rejects it)", () => {
    // Same scope rule as the built-in preset — `<c:dispUnits>` lives
    // exclusively on `CT_ValAx`, so a stray hint on the column's X axis
    // silently drops at the writer.
    const result = writeChart(
      makeChart({ type: "column", axes: { x: { dispUnits: { custUnit: 1000 } } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).not.toContain("c:dispUnits");
    expect(catAxBlock).not.toContain("c:custUnit");
  });

  it("emits <c:custUnit> on both scatter axes (both are valAx)", () => {
    const scatter: SheetChart = {
      type: "scatter",
      series: [{ name: "S1", values: "B2:B5", categories: "A2:A5" }],
      anchor: { from: { row: 0, col: 0 } },
      axes: {
        x: { dispUnits: { custUnit: 60 } },
        y: { dispUnits: { custUnit: 86400, showLabel: true } },
      },
    };
    const result = writeChart(scatter, "Sheet1");
    const valAxBlocks = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    expect(valAxBlocks).toHaveLength(2);
    expect(valAxBlocks[0]).toContain('<c:custUnit val="60"/>');
    expect(valAxBlocks[0]).not.toContain("c:dispUnitsLbl");
    expect(valAxBlocks[1]).toContain('<c:custUnit val="86400"/>');
    expect(valAxBlocks[1]).toContain("<c:dispUnitsLbl/>");
  });

  it("survives a parseChart round-trip on the value axis with custUnit", () => {
    const result = writeChart(
      makeChart({ axes: { y: { dispUnits: { custUnit: 500, showLabel: true } } } }),
      "Sheet1",
    );
    const reparsed = parseChart(result.chartXml);
    expect(reparsed?.axes?.y?.dispUnits).toEqual({ custUnit: 500, showLabel: true });
  });

  it("packages a custUnit chart end-to-end through writeXlsx", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Hour", "Seconds"],
          ["H1", 360_000],
          ["H2", 720_000],
          ["H3", 1_080_000],
        ],
        charts: [
          {
            type: "column",
            series: [{ name: "Seconds", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            // 86400 seconds per day — convert tick labels to days.
            axes: { y: { dispUnits: { custUnit: 86400, showLabel: true } } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('<c:custUnit val="86400"/>');
    expect(chartXml).toContain("<c:dispUnitsLbl/>");
    expect(chartXml).not.toContain("c:builtInUnit");
    const reparsed = parseChart(chartXml);
    expect(reparsed?.axes?.y?.dispUnits).toEqual({ custUnit: 86400, showLabel: true });
  });
});

// ── writeChart — chart style preset ──────────────────────────────────

describe("writeChart — chart style preset", () => {
  it("skips <c:style> entirely when the field is unset (writer default)", () => {
    // Excel's reference serialization for a fresh chart pins style 2,
    // but the writer skips emission so an unstyled chart stays minimal
    // — Excel falls back to its application default look.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:style ");
    expect(result.chartXml).not.toContain("<c:style/>");
  });

  it('emits <c:style val="N"/> on <c:chartSpace> when the field is pinned', () => {
    const result = writeChart(makeChart({ style: 27 }), "Sheet1");
    expect(result.chartXml).toContain('c:style val="27"');
  });

  it("emits the OOXML range bounds (1 and 48)", () => {
    for (const val of [1, 48]) {
      const result = writeChart(makeChart({ style: val }), "Sheet1");
      expect(result.chartXml).toContain(`c:style val="${val}"`);
    }
  });

  it("places <c:style> after <c:roundedCorners> and before <c:chart>", () => {
    // CT_ChartSpace sequence: ... roundedCorners?, AlternateContent?,
    // clrMapOvr?, style?, ... chart, ... — the preset must follow
    // <c:roundedCorners> and precede <c:chart> so a strict validator
    // (Excel itself rejects out-of-order children) sees the schema
    // sequence respected.
    const result = writeChart(makeChart({ style: 12, roundedCorners: true }), "Sheet1");
    const roundedIdx = result.chartXml.indexOf("c:roundedCorners");
    const styleIdx = result.chartXml.indexOf("c:style ");
    const chartIdx = result.chartXml.indexOf("<c:chart>");
    expect(roundedIdx).toBeGreaterThan(-1);
    expect(styleIdx).toBeGreaterThan(roundedIdx);
    expect(styleIdx).toBeLessThan(chartIdx);
  });

  it("only emits <c:style> once on a chart that pins it", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(makeChart({ style: 27 }), "Sheet1");
    const occurrences = result.chartXml.match(/<c:style /g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("drops out-of-range style values (0 / 49 / 255) rather than emit them", () => {
    // CT_Style declares val as xsd:unsignedByte in the gallery range
    // 1–48. Out-of-range values collapse to absence so the writer
    // never emits a token Excel would reject.
    for (const val of [0, 49, 100, 255, -3]) {
      const result = writeChart(makeChart({ style: val }), "Sheet1");
      expect(result.chartXml).not.toContain("<c:style ");
      expect(result.chartXml).not.toContain("<c:style/>");
    }
  });

  it("drops non-integer style values (3.5 / NaN / Infinity)", () => {
    for (const val of [3.5, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      const result = writeChart(makeChart({ style: val }), "Sheet1");
      expect(result.chartXml).not.toContain("<c:style ");
    }
  });

  it("threads style through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, style: 18 }), "Sheet1");
      expect(result.chartXml).toContain('c:style val="18"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        style: 18,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:style val="18"');
  });

  it("round-trips a pinned style through parseChart", () => {
    const written = writeChart(makeChart({ style: 27 }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.style).toBe(27);
  });

  it("collapses an unset style round-trip back to undefined", () => {
    // A fresh chart writes no element, which re-parses to undefined —
    // absence and the unstyled default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.style).toBeUndefined();
  });

  it("threads style end-to-end through writeXlsx packaging", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Dashboard",
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 10],
          ["Q2", 20],
          ["Q3", 30],
        ],
        charts: [
          {
            type: "column",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            style: 34,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('c:style val="34"');
    // Re-parse the rendered chart to confirm the preset survives the
    // packaging path.
    const reparsed = parseChart(chartXml);
    expect(reparsed?.style).toBe(34);
  });
});

// ── writeChart — chart editing locale (lang) ─────────────────────────

describe("writeChart — chart editing locale", () => {
  it("skips <c:lang> entirely when the field is unset (writer default)", () => {
    // Excel's reference serialization for a fresh chart authored on
    // an English locale pins <c:lang val="en-US"/>, but the writer
    // skips emission when the chart leaves `lang` unset so the file
    // does not silently fabricate a locale Excel falls back to anyway.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:lang ");
    expect(result.chartXml).not.toContain("<c:lang/>");
  });

  it('emits <c:lang val=".."/> on <c:chartSpace> when the field is pinned', () => {
    const result = writeChart(makeChart({ lang: "en-US" }), "Sheet1");
    expect(result.chartXml).toContain('c:lang val="en-US"');
  });

  it("emits a non-English locale verbatim", () => {
    for (const tag of ["tr-TR", "de-DE", "pt-BR", "zh-Hans-CN", "fr"]) {
      const result = writeChart(makeChart({ lang: tag }), "Sheet1");
      expect(result.chartXml).toContain(`c:lang val="${tag}"`);
    }
  });

  it("places <c:lang> before <c:roundedCorners> on <c:chartSpace>", () => {
    // CT_ChartSpace sequence: date1904?, lang?, roundedCorners?, ...
    // — the locale must precede <c:roundedCorners> so a strict
    // validator (Excel itself rejects out-of-order children) sees the
    // schema sequence respected.
    const result = writeChart(makeChart({ lang: "en-US" }), "Sheet1");
    const langIdx = result.chartXml.indexOf("c:lang ");
    const roundedIdx = result.chartXml.indexOf("c:roundedCorners");
    const chartIdx = result.chartXml.indexOf("<c:chart>");
    expect(langIdx).toBeGreaterThan(-1);
    expect(roundedIdx).toBeGreaterThan(langIdx);
    expect(chartIdx).toBeGreaterThan(roundedIdx);
  });

  it("places <c:lang> before <c:style> when both are pinned", () => {
    // <c:lang> precedes <c:roundedCorners> which precedes <c:style>
    // per CT_ChartSpace; the writer threads all three in the right
    // order so a validator never sees them transposed.
    const result = writeChart(
      makeChart({ lang: "tr-TR", style: 27, roundedCorners: true }),
      "Sheet1",
    );
    const langIdx = result.chartXml.indexOf("c:lang ");
    const roundedIdx = result.chartXml.indexOf("c:roundedCorners");
    const styleIdx = result.chartXml.indexOf("c:style ");
    expect(langIdx).toBeGreaterThan(-1);
    expect(roundedIdx).toBeGreaterThan(langIdx);
    expect(styleIdx).toBeGreaterThan(roundedIdx);
  });

  it("only emits <c:lang> once on a chart that pins it", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(makeChart({ lang: "en-US" }), "Sheet1");
    const occurrences = result.chartXml.match(/<c:lang /g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("drops malformed locale tokens rather than emit them", () => {
    // <c:lang> is xsd:language in the OOXML schema (BCP-47 culture
    // names). Tokens that don't match the alphabet / length shape
    // collapse to absence so the writer never emits a value Excel
    // would reject.
    for (const bad of [
      "english",
      "en US",
      "en_US",
      "1234",
      "",
      " ",
      "a-bad-very-long-tag-segment",
      "en-",
      "-US",
    ]) {
      const result = writeChart(makeChart({ lang: bad }), "Sheet1");
      expect(result.chartXml).not.toContain("<c:lang ");
      expect(result.chartXml).not.toContain("<c:lang/>");
    }
  });

  it("drops non-string lang values rather than fabricate one", () => {
    for (const val of [42, true, null, undefined, {}, []]) {
      const result = writeChart(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        makeChart({ lang: val as any }),
        "Sheet1",
      );
      expect(result.chartXml).not.toContain("<c:lang ");
    }
  });

  it("threads lang through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, lang: "tr-TR" }), "Sheet1");
      expect(result.chartXml).toContain('c:lang val="tr-TR"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        lang: "tr-TR",
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:lang val="tr-TR"');
  });

  it("round-trips a pinned lang through parseChart", () => {
    const written = writeChart(makeChart({ lang: "en-US" }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.lang).toBe("en-US");
  });

  it("collapses an unset lang round-trip back to undefined", () => {
    // A fresh chart writes no element, which re-parses to undefined —
    // absence and the unset default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.lang).toBeUndefined();
  });

  it("threads lang end-to-end through writeXlsx packaging", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Dashboard",
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 10],
          ["Q2", 20],
          ["Q3", 30],
        ],
        charts: [
          {
            type: "column",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            lang: "tr-TR",
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('c:lang val="tr-TR"');
    // Re-parse the rendered chart to confirm the locale survives the
    // packaging path.
    const reparsed = parseChart(chartXml);
    expect(reparsed?.lang).toBe("tr-TR");
  });
});

// ── writeChart — chart date system (date1904) ────────────────────────

describe("writeChart — chart date system", () => {
  it("skips <c:date1904> entirely when the field is unset (writer default)", () => {
    // Excel's reference serialization always emits <c:date1904 val="0"/>,
    // but the writer skips emission when the chart leaves `date1904`
    // unset so the file does not silently fabricate a flag Excel
    // falls back to anyway. Absence and the OOXML default round-trip
    // identically through cloneChart.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:date1904 ");
    expect(result.chartXml).not.toContain("<c:date1904/>");
  });

  it("skips <c:date1904> when date1904 is false (matches OOXML default)", () => {
    // `false` and absence both map to the default `val="0"` — the
    // writer skips the element so re-parse collapses back to the
    // same `undefined` that absence would produce.
    const result = writeChart(makeChart({ date1904: false }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:date1904 ");
    expect(result.chartXml).not.toContain("<c:date1904/>");
  });

  it('emits <c:date1904 val="1"/> when the chart pins date1904: true', () => {
    const result = writeChart(makeChart({ date1904: true }), "Sheet1");
    expect(result.chartXml).toContain('c:date1904 val="1"');
  });

  it("places <c:date1904> before <c:roundedCorners> on <c:chartSpace>", () => {
    // CT_ChartSpace sequence: date1904?, lang?, roundedCorners?, ...
    // — the date-system flag must precede <c:roundedCorners> so a
    // strict validator (Excel itself rejects out-of-order children)
    // sees the schema sequence respected.
    const result = writeChart(makeChart({ date1904: true }), "Sheet1");
    const dateIdx = result.chartXml.indexOf("c:date1904 ");
    const roundedIdx = result.chartXml.indexOf("c:roundedCorners");
    const chartIdx = result.chartXml.indexOf("<c:chart>");
    expect(dateIdx).toBeGreaterThan(-1);
    expect(roundedIdx).toBeGreaterThan(dateIdx);
    expect(chartIdx).toBeGreaterThan(roundedIdx);
  });

  it("places <c:date1904> before <c:lang> when both are pinned", () => {
    // <c:date1904> sits at the head of CT_ChartSpace, before <c:lang>
    // which sits before <c:roundedCorners> which sits before
    // <c:style> — the writer threads all four in the right order so
    // a validator never sees them transposed.
    const result = writeChart(
      makeChart({ date1904: true, lang: "tr-TR", style: 27, roundedCorners: true }),
      "Sheet1",
    );
    const dateIdx = result.chartXml.indexOf("c:date1904 ");
    const langIdx = result.chartXml.indexOf("c:lang ");
    const roundedIdx = result.chartXml.indexOf("c:roundedCorners");
    const styleIdx = result.chartXml.indexOf("c:style ");
    expect(dateIdx).toBeGreaterThan(-1);
    expect(langIdx).toBeGreaterThan(dateIdx);
    expect(roundedIdx).toBeGreaterThan(langIdx);
    expect(styleIdx).toBeGreaterThan(roundedIdx);
  });

  it("only emits <c:date1904> once on a chart that pins it", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(makeChart({ date1904: true }), "Sheet1");
    const occurrences = result.chartXml.match(/<c:date1904 /g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads date1904 through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, date1904: true }), "Sheet1");
      expect(result.chartXml).toContain('c:date1904 val="1"');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        date1904: true,
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('c:date1904 val="1"');
  });

  it("round-trips a pinned date1904 through parseChart", () => {
    const written = writeChart(makeChart({ date1904: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.date1904).toBe(true);
  });

  it("collapses an unset date1904 round-trip back to undefined", () => {
    // A fresh chart writes no element, which re-parses to undefined —
    // absence and the unset default round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.date1904).toBeUndefined();
  });

  it("collapses date1904: false round-trip back to undefined", () => {
    // `false` writes nothing (matches OOXML default), so re-parse
    // also returns undefined — there is no asymmetry between the
    // pinned-default and unset states.
    const written = writeChart(makeChart({ date1904: false }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.date1904).toBeUndefined();
  });

  it("threads date1904 end-to-end through writeXlsx packaging", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Dashboard",
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 10],
          ["Q2", 20],
          ["Q3", 30],
        ],
        charts: [
          {
            type: "column",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            date1904: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('c:date1904 val="1"');
    // Re-parse the rendered chart to confirm the date-system flag
    // survives the packaging path.
    const reparsed = parseChart(chartXml);
    expect(reparsed?.date1904).toBe(true);
  });
});

// ── writeChart — axis crossBetween ───────────────────────────────────

describe("writeChart — axis crossBetween", () => {
  it('emits the family default <c:crossBetween val="between"/> on a column chart with no override', () => {
    // The writer always emits `<c:crossBetween>` on the value axis
    // because the OOXML schema requires it. The default for bar /
    // column / line / area is `"between"` — Excel's reference
    // serialization on every freshly-drawn column chart pins exactly
    // that value.
    const result = writeChart(makeChart(), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:crossBetween val="between"/>');
  });

  it("honours a value-axis override on a column chart", () => {
    const result = writeChart(makeChart({ axes: { y: { crossBetween: "midCat" } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:crossBetween val="midCat"/>');
  });

  it('emits the family default <c:crossBetween val="midCat"/> on both scatter axes', () => {
    // Scatter charts route both axes through `<c:valAx>`; the writer
    // pins `"midCat"` on each by default to mirror Excel's reference
    // serialization on a freshly-drawn scatter chart.
    const scatter: SheetChart = {
      type: "scatter",
      series: [{ name: "S1", values: "B2:B5", categories: "A2:A5" }],
      anchor: { from: { row: 0, col: 0 } },
    };
    const result = writeChart(scatter, "Sheet1");
    const valAxBlocks = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    expect(valAxBlocks).toHaveLength(2);
    expect(valAxBlocks[0]).toContain('<c:crossBetween val="midCat"/>');
    expect(valAxBlocks[1]).toContain('<c:crossBetween val="midCat"/>');
  });

  it("honours independent overrides on both scatter axes", () => {
    const scatter: SheetChart = {
      type: "scatter",
      series: [{ name: "S1", values: "B2:B5", categories: "A2:A5" }],
      anchor: { from: { row: 0, col: 0 } },
      axes: {
        x: { crossBetween: "between" },
        y: { crossBetween: "between" },
      },
    };
    const result = writeChart(scatter, "Sheet1");
    const valAxBlocks = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    expect(valAxBlocks).toHaveLength(2);
    expect(valAxBlocks[0]).toContain('<c:crossBetween val="between"/>');
    expect(valAxBlocks[1]).toContain('<c:crossBetween val="between"/>');
  });

  it("drops an unknown ST_CrossBetween token rather than fabricating a value", () => {
    const result = writeChart(
      makeChart({
        axes: { y: { crossBetween: "diagonal" as never } },
      }),
      "Sheet1",
    );
    // Falls back to the family default rather than emitting the bad token.
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain('<c:crossBetween val="between"/>');
    expect(result.chartXml).not.toContain('val="diagonal"');
  });

  it("does not emit <c:crossBetween> on the X axis of a column chart (catAx rejects it)", () => {
    // The OOXML schema places <c:crossBetween> exclusively on CT_ValAx,
    // so even though the user pinned a value on the X axis, the catAx
    // builder should silently drop the field.
    const result = writeChart(
      makeChart({ type: "column", axes: { x: { crossBetween: "midCat" } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).not.toContain("c:crossBetween");
  });

  it("does not emit <c:crossBetween> on a pie chart (no axes at all)", () => {
    // The writer never builds <c:valAx> for pie / doughnut, so even
    // when the caller pins a value the element should not surface.
    const result = writeChart(
      makeChart({
        type: "pie",
        axes: { y: { crossBetween: "midCat" } },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:crossBetween");
  });

  it("places <c:crossBetween> after <c:crosses> inside <c:valAx> (CT_ValAx order)", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            crosses: "max",
            crossBetween: "midCat",
          },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const crossesIdx = valAxBlock.indexOf("c:crosses");
    const crossBetweenIdx = valAxBlock.indexOf("c:crossBetween");
    expect(crossesIdx).toBeGreaterThan(-1);
    expect(crossBetweenIdx).toBeGreaterThan(crossesIdx);
  });

  it("places <c:crossBetween> before <c:majorUnit> inside <c:valAx> (CT_ValAx order)", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: {
            scale: { min: 0, max: 100, majorUnit: 25 },
            crossBetween: "midCat",
          },
        },
      }),
      "Sheet1",
    );
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    const crossBetweenIdx = valAxBlock.indexOf("c:crossBetween");
    const majorUnitIdx = valAxBlock.indexOf("c:majorUnit");
    expect(crossBetweenIdx).toBeGreaterThan(-1);
    expect(majorUnitIdx).toBeGreaterThan(crossBetweenIdx);
  });

  it("only emits <c:crossBetween> once on the value axis", () => {
    const result = writeChart(makeChart({ axes: { y: { crossBetween: "midCat" } } }), "Sheet1");
    const occurrences = result.chartXml.match(/c:crossBetween/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("survives a parseChart round-trip on the value axis", () => {
    const result = writeChart(makeChart({ axes: { y: { crossBetween: "midCat" } } }), "Sheet1");
    const reparsed = parseChart(result.chartXml);
    expect(reparsed?.axes?.y?.crossBetween).toBe("midCat");
  });

  it("survives a parseChart round-trip on a scatter chart with an X-axis override", () => {
    const scatter: SheetChart = {
      type: "scatter",
      series: [{ name: "S1", values: "B2:B5", categories: "A2:A5" }],
      anchor: { from: { row: 0, col: 0 } },
      axes: { x: { crossBetween: "between" } },
    };
    const result = writeChart(scatter, "Sheet1");
    const reparsed = parseChart(result.chartXml);
    expect(reparsed?.axes?.x?.crossBetween).toBe("between");
    // Y axis stayed at the scatter family default — collapses on read.
    expect(reparsed?.axes?.y?.crossBetween).toBeUndefined();
  });

  it("collapses a defaulted crossBetween round-trip back to undefined", () => {
    // A chart that left crossBetween unset emits the family default,
    // and the reader should collapse that default back to undefined.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.y?.crossBetween).toBeUndefined();
  });

  it("packages the chart end-to-end through writeXlsx", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 100],
          ["Q2", 200],
          ["Q3", 150],
        ],
        charts: [
          {
            type: "line",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            axes: { y: { crossBetween: "midCat" } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('<c:crossBetween val="midCat"/>');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.axes?.y?.crossBetween).toBe("midCat");
  });
});

// ── writeChart — data table ──────────────────────────────────────────

describe("writeChart — data table", () => {
  it("skips <c:dTable> entirely when the field is unset (writer default)", () => {
    // Excel renders no data table on a fresh chart — the writer skips
    // the element so the file stays minimal. Absence and the unset
    // default round-trip identically through cloneChart.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:dTable");
  });

  it("skips <c:dTable> when dataTable is false", () => {
    // `false` mirrors the unset default — the writer skips the element.
    const result = writeChart(makeChart({ dataTable: false }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:dTable");
  });

  it("emits a fully-defaulted <c:dTable> when dataTable is true", () => {
    // `true` enables the table with Excel's reference defaults — every
    // border drawn, the outline frame on, and the legend keys shown.
    const result = writeChart(makeChart({ dataTable: true }), "Sheet1");
    expect(result.chartXml).toContain("<c:dTable>");
    expect(result.chartXml).toContain('<c:showHorzBorder val="1"/>');
    expect(result.chartXml).toContain('<c:showVertBorder val="1"/>');
    expect(result.chartXml).toContain('<c:showOutline val="1"/>');
    expect(result.chartXml).toContain('<c:showKeys val="1"/>');
  });

  it("emits a fully-defaulted <c:dTable> when dataTable is an empty object", () => {
    // An empty override accepts every default — equivalent to passing
    // `dataTable: true`. Each unspecified field falls back to the OOXML
    // reference value `true`.
    const result = writeChart(makeChart({ dataTable: {} }), "Sheet1");
    expect(result.chartXml).toContain("<c:dTable>");
    expect(result.chartXml).toContain('<c:showHorzBorder val="1"/>');
    expect(result.chartXml).toContain('<c:showVertBorder val="1"/>');
    expect(result.chartXml).toContain('<c:showOutline val="1"/>');
    expect(result.chartXml).toContain('<c:showKeys val="1"/>');
  });

  it("honours per-field overrides (false for showKeys)", () => {
    // A common pattern — paint the table grid but hide the legend
    // swatches because the chart already carries a separate legend.
    const result = writeChart(makeChart({ dataTable: { showKeys: false } }), "Sheet1");
    expect(result.chartXml).toContain('<c:showHorzBorder val="1"/>');
    expect(result.chartXml).toContain('<c:showVertBorder val="1"/>');
    expect(result.chartXml).toContain('<c:showOutline val="1"/>');
    expect(result.chartXml).toContain('<c:showKeys val="0"/>');
  });

  it("flips every flag false when the caller pins each false", () => {
    const result = writeChart(
      makeChart({
        dataTable: {
          showHorzBorder: false,
          showVertBorder: false,
          showOutline: false,
          showKeys: false,
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<c:showHorzBorder val="0"/>');
    expect(result.chartXml).toContain('<c:showVertBorder val="0"/>');
    expect(result.chartXml).toContain('<c:showOutline val="0"/>');
    expect(result.chartXml).toContain('<c:showKeys val="0"/>');
  });

  it("emits the four <c:dTable> children in CT_DTable schema order", () => {
    // CT_DTable: showHorzBorder?, showVertBorder?, showOutline?, showKeys?,
    // spPr?, txPr?, extLst? — though all four boolean children are
    // always emitted, the order matters for strict validators (Excel
    // itself rejects out-of-order children).
    const result = writeChart(makeChart({ dataTable: true }), "Sheet1");
    const horzIdx = result.chartXml.indexOf("c:showHorzBorder");
    const vertIdx = result.chartXml.indexOf("c:showVertBorder");
    const outlineIdx = result.chartXml.indexOf("c:showOutline");
    const keysIdx = result.chartXml.indexOf("c:showKeys");
    expect(horzIdx).toBeGreaterThan(-1);
    expect(vertIdx).toBeGreaterThan(horzIdx);
    expect(outlineIdx).toBeGreaterThan(vertIdx);
    expect(keysIdx).toBeGreaterThan(outlineIdx);
  });

  it("places <c:dTable> after the axes inside <c:plotArea>", () => {
    // CT_PlotArea sequence: layout?, [chart-types], catAx*/valAx*/dateAx*/
    // serAx*, dTable?, spPr? — the table sits after the axes.
    const result = writeChart(makeChart({ dataTable: true }), "Sheet1");
    const lastValAxClose = result.chartXml.lastIndexOf("</c:valAx>");
    const dTableIdx = result.chartXml.indexOf("<c:dTable>");
    const plotAreaCloseIdx = result.chartXml.indexOf("</c:plotArea>");
    expect(lastValAxClose).toBeGreaterThan(-1);
    expect(dTableIdx).toBeGreaterThan(lastValAxClose);
    expect(plotAreaCloseIdx).toBeGreaterThan(dTableIdx);
  });

  it("only emits <c:dTable> once on a chart that pins it", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(makeChart({ dataTable: true }), "Sheet1");
    const occurrences = result.chartXml.match(/<c:dTable>/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads dataTable through every chart family with axes", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, dataTable: true }), "Sheet1");
      expect(result.chartXml).toContain("<c:dTable>");
    }
    const scatter = writeChart(
      {
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        anchor: { from: { row: 5, col: 0 } },
        dataTable: true,
      },
      "Sheet1",
    );
    expect(scatter.chartXml).toContain("<c:dTable>");
  });

  it("silently drops dataTable on pie charts (no axes, no slot for the element)", () => {
    // The OOXML schema places <c:dTable> inside <c:plotArea> alongside
    // the axes, but pie charts have no axes at all. The writer drops
    // the field rather than emit an element Excel's strict validator
    // would reject.
    const result = writeChart(
      {
        type: "pie",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        anchor: { from: { row: 5, col: 0 } },
        dataTable: true,
      },
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:dTable");
  });

  it("silently drops dataTable on doughnut charts (no axes either)", () => {
    const result = writeChart(
      {
        type: "doughnut",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        anchor: { from: { row: 5, col: 0 } },
        dataTable: true,
      },
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:dTable");
  });

  it("round-trips a pinned dataTable through parseChart", () => {
    const written = writeChart(
      makeChart({
        dataTable: {
          showHorzBorder: true,
          showVertBorder: false,
          showOutline: true,
          showKeys: false,
        },
      }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataTable).toEqual({
      showHorzBorder: true,
      showVertBorder: false,
      showOutline: true,
      showKeys: false,
    });
  });

  it("round-trips dataTable: true as the fully-defaulted shape", () => {
    // The writer emits all four flags as `true` — a re-parse surfaces
    // every one literally because every child is required on
    // CT_DTable.
    const written = writeChart(makeChart({ dataTable: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataTable).toEqual({
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: true,
    });
  });

  it("collapses an unset dataTable round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataTable).toBeUndefined();
  });

  it("collapses dataTable: false round-trip back to undefined", () => {
    // `false` writes nothing (matches the unset default), so re-parse
    // also returns undefined.
    const written = writeChart(makeChart({ dataTable: false }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataTable).toBeUndefined();
  });

  it("threads dataTable end-to-end through writeXlsx packaging", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Dashboard",
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 10],
          ["Q2", 20],
          ["Q3", 30],
        ],
        charts: [
          {
            type: "column",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            dataTable: { showKeys: false },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:dTable>");
    expect(chartXml).toContain('<c:showKeys val="0"/>');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.dataTable).toEqual({
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: false,
    });
  });
});

// ── writeChart — chart-space protection ──────────────────────────────

describe("writeChart — chart-space protection", () => {
  it("skips <c:protection> entirely when the field is unset (writer default)", () => {
    // Excel applies no chart-level protection on a fresh chart — the
    // writer skips the element so the file stays minimal. Absence and
    // the unset default round-trip identically through cloneChart.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:protection");
  });

  it("skips <c:protection> when protection is false", () => {
    // `false` mirrors the unset default — the writer skips the element.
    const result = writeChart(makeChart({ protection: false }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:protection");
  });

  it("emits a fully-defaulted <c:protection> when protection is true", () => {
    // `true` declares the bare <c:protection> shell with every flag at
    // its OOXML default `false` — useful for round-trip parity with
    // templates that pin the bare element without locking any flag.
    const result = writeChart(makeChart({ protection: true }), "Sheet1");
    expect(result.chartXml).toContain("<c:protection>");
    expect(result.chartXml).toContain('<c:chartObject val="0"/>');
    expect(result.chartXml).toContain('<c:data val="0"/>');
    expect(result.chartXml).toContain('<c:formatting val="0"/>');
    expect(result.chartXml).toContain('<c:selection val="0"/>');
    expect(result.chartXml).toContain('<c:userInterface val="0"/>');
  });

  it("emits a fully-defaulted <c:protection> when protection is an empty object", () => {
    // An empty override accepts every default — equivalent to passing
    // `protection: true`. Each unspecified field falls back to the
    // OOXML reference value `false`.
    const result = writeChart(makeChart({ protection: {} }), "Sheet1");
    expect(result.chartXml).toContain("<c:protection>");
    expect(result.chartXml).toContain('<c:chartObject val="0"/>');
    expect(result.chartXml).toContain('<c:data val="0"/>');
    expect(result.chartXml).toContain('<c:formatting val="0"/>');
    expect(result.chartXml).toContain('<c:selection val="0"/>');
    expect(result.chartXml).toContain('<c:userInterface val="0"/>');
  });

  it("honours per-field overrides (true for formatting + selection)", () => {
    // A common pattern — lock the look (formatting + selection) but
    // leave the underlying data references and the on-canvas UI
    // affordances unlocked.
    const result = writeChart(
      makeChart({
        protection: { formatting: true, selection: true },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<c:chartObject val="0"/>');
    expect(result.chartXml).toContain('<c:data val="0"/>');
    expect(result.chartXml).toContain('<c:formatting val="1"/>');
    expect(result.chartXml).toContain('<c:selection val="1"/>');
    expect(result.chartXml).toContain('<c:userInterface val="0"/>');
  });

  it("flips every flag true when the caller pins each true", () => {
    const result = writeChart(
      makeChart({
        protection: {
          chartObject: true,
          data: true,
          formatting: true,
          selection: true,
          userInterface: true,
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<c:chartObject val="1"/>');
    expect(result.chartXml).toContain('<c:data val="1"/>');
    expect(result.chartXml).toContain('<c:formatting val="1"/>');
    expect(result.chartXml).toContain('<c:selection val="1"/>');
    expect(result.chartXml).toContain('<c:userInterface val="1"/>');
  });

  it("emits the five <c:protection> children in CT_Protection schema order", () => {
    // CT_Protection: chartObject?, data?, formatting?, selection?,
    // userInterface? — order matters for strict validators (Excel
    // itself rejects out-of-order children).
    const result = writeChart(makeChart({ protection: true }), "Sheet1");
    const chartObjectIdx = result.chartXml.indexOf("c:chartObject");
    const dataIdx = result.chartXml.indexOf("c:data ");
    const formattingIdx = result.chartXml.indexOf("c:formatting");
    const selectionIdx = result.chartXml.indexOf("c:selection");
    const userInterfaceIdx = result.chartXml.indexOf("c:userInterface");
    expect(chartObjectIdx).toBeGreaterThan(-1);
    expect(dataIdx).toBeGreaterThan(chartObjectIdx);
    expect(formattingIdx).toBeGreaterThan(dataIdx);
    expect(selectionIdx).toBeGreaterThan(formattingIdx);
    expect(userInterfaceIdx).toBeGreaterThan(selectionIdx);
  });

  it("places <c:protection> after <c:style> and before <c:chart> on <c:chartSpace>", () => {
    // CT_ChartSpace sequence: date1904?, lang?, roundedCorners?,
    // AlternateContent?, style?, clrMapOvr?, pivotSource?, protection?,
    // chart, ... — protection sits after style and before the visible
    // <c:chart>.
    const result = writeChart(makeChart({ style: 12, protection: true }), "Sheet1");
    const styleIdx = result.chartXml.indexOf("<c:style");
    const protectionIdx = result.chartXml.indexOf("<c:protection>");
    const chartIdx = result.chartXml.indexOf("<c:chart>");
    expect(styleIdx).toBeGreaterThan(-1);
    expect(protectionIdx).toBeGreaterThan(styleIdx);
    expect(chartIdx).toBeGreaterThan(protectionIdx);
  });

  it("places <c:protection> before <c:chart> even without a <c:style>", () => {
    // The element should still slot directly before <c:chart> when
    // no style preset is set — the writer must not require a sibling
    // to anchor the position.
    const result = writeChart(makeChart({ protection: true }), "Sheet1");
    const protectionIdx = result.chartXml.indexOf("<c:protection>");
    const chartIdx = result.chartXml.indexOf("<c:chart>");
    expect(protectionIdx).toBeGreaterThan(-1);
    expect(chartIdx).toBeGreaterThan(protectionIdx);
  });

  it("only emits <c:protection> once on a chart that pins it", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(makeChart({ protection: true }), "Sheet1");
    const occurrences = result.chartXml.match(/<c:protection>/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads protection through every chart family", () => {
    // <c:protection> lives on <c:chartSpace>, not inside <c:plotArea>
    // — every chart family carries a slot, including pie / doughnut.
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, protection: { formatting: true } }), "Sheet1");
      expect(result.chartXml).toContain("<c:protection>");
      expect(result.chartXml).toContain('<c:formatting val="1"/>');
    }
    for (const type of ["pie", "doughnut"] as const) {
      const result = writeChart(
        {
          type,
          series: [{ values: "B2:B4", categories: "A2:A4" }],
          anchor: { from: { row: 5, col: 0 } },
          protection: { formatting: true },
        },
        "Sheet1",
      );
      expect(result.chartXml).toContain("<c:protection>");
      expect(result.chartXml).toContain('<c:formatting val="1"/>');
    }
    const scatter = writeChart(
      {
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        anchor: { from: { row: 5, col: 0 } },
        protection: { formatting: true },
      },
      "Sheet1",
    );
    expect(scatter.chartXml).toContain("<c:protection>");
  });

  it("ignores stray non-boolean overrides rather than emit invalid tokens", () => {
    // A non-boolean leaking through the type guard collapses to the
    // OOXML default `false` rather than emit a token Excel rejects.
    const result = writeChart(
      makeChart({
        protection: {
          // @ts-expect-error — defensive: stray non-boolean leaking through
          chartObject: 1,
          // @ts-expect-error — defensive: stray non-boolean leaking through
          data: "true",
          formatting: true,
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<c:chartObject val="0"/>');
    expect(result.chartXml).toContain('<c:data val="0"/>');
    expect(result.chartXml).toContain('<c:formatting val="1"/>');
  });

  it("round-trips a pinned protection through parseChart", () => {
    const written = writeChart(
      makeChart({
        protection: {
          chartObject: true,
          data: false,
          formatting: true,
          selection: false,
          userInterface: true,
        },
      }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.protection).toEqual({
      chartObject: true,
      data: false,
      formatting: true,
      selection: false,
      userInterface: true,
    });
  });

  it("round-trips protection: true as the fully-defaulted shape", () => {
    // The writer emits all five flags as `false` — a re-parse surfaces
    // the same shape (every flag literal `false`) so the cloned
    // protection is byte-for-byte stable.
    const written = writeChart(makeChart({ protection: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.protection).toEqual({
      chartObject: false,
      data: false,
      formatting: false,
      selection: false,
      userInterface: false,
    });
  });

  it("threads protection through writeXlsx into xl/charts/chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Month", "Revenue"],
          ["Jan", 100],
          ["Feb", 200],
          ["Mar", 300],
        ],
        charts: [
          {
            type: "column",
            title: "Quarterly",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            protection: { formatting: true, selection: true },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:protection>");
    expect(chartXml).toContain('<c:formatting val="1"/>');
    expect(chartXml).toContain('<c:selection val="1"/>');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.protection).toEqual({
      chartObject: false,
      data: false,
      formatting: true,
      selection: true,
      userInterface: false,
    });
  });
});

// ── writeChart — chart-level line marker visibility ───────────────────

describe("writeChart — showLineMarkers", () => {
  it('emits <c:marker val="1"/> by default on a line chart (Excel reference)', () => {
    // Excel's reference serialization for every authored line chart
    // emits <c:marker val="1"/> — markers shown at each data point.
    // The writer mirrors that on `showLineMarkers` undefined so a
    // back-compat caller that never set the field keeps the same output.
    const result = writeChart(makeChart({ type: "line" }), "Sheet1");
    expect(result.chartXml).toContain('<c:marker val="1"/>');
  });

  it('emits <c:marker val="1"/> when showLineMarkers=true', () => {
    const result = writeChart(makeChart({ type: "line", showLineMarkers: true }), "Sheet1");
    expect(result.chartXml).toContain('<c:marker val="1"/>');
  });

  it('emits <c:marker val="0"/> when showLineMarkers=false', () => {
    // The non-default state — flips the chart-level gate off so
    // per-series marker definitions stop rendering chart-wide.
    const result = writeChart(makeChart({ type: "line", showLineMarkers: false }), "Sheet1");
    expect(result.chartXml).toContain('<c:marker val="0"/>');
  });

  it("emits <c:marker> exactly once on a line chart", () => {
    // Guard against any regression that would double-emit the element.
    // The chart-level <c:marker> sits in CT_LineChart, not CT_LineSer —
    // per-series <c:marker> blocks (line / scatter) live elsewhere.
    const result = writeChart(makeChart({ type: "line", showLineMarkers: false }), "Sheet1");
    // Filter to the chart-level <c:marker val=".."/> shape — per-series
    // <c:marker> blocks are never authored without a child element.
    const occurrences = result.chartXml.match(/<c:marker val="[01]"\s*\/>/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("places <c:marker> after <c:upDownBars> inside <c:lineChart> (OOXML order)", () => {
    // CT_LineChart sequence: ... ser*, dLbls?, dropLines?, hiLowLines?,
    // upDownBars?, marker?, axId+. The schema rejects an out-of-order
    // <c:marker> before <c:upDownBars>, so the writer must place
    // <c:marker> after the up/down bars block when both are present.
    const result = writeChart(
      makeChart({ type: "line", upDownBars: true, showLineMarkers: false }),
      "Sheet1",
    );
    const upDownBarsIdx = result.chartXml.indexOf("c:upDownBars");
    const markerIdx = result.chartXml.indexOf('c:marker val="0"');
    expect(upDownBarsIdx).toBeGreaterThan(0);
    expect(markerIdx).toBeGreaterThan(0);
    expect(upDownBarsIdx).toBeLessThan(markerIdx);
  });

  it("places <c:marker> before <c:axId> inside <c:lineChart> (OOXML order)", () => {
    // The axId pair sits at the tail of CT_LineChart — every optional
    // chart-level child must precede them.
    const result = writeChart(makeChart({ type: "line", showLineMarkers: false }), "Sheet1");
    const markerIdx = result.chartXml.indexOf('c:marker val="0"');
    const firstAxIdIdx = result.chartXml.indexOf("c:axId");
    expect(markerIdx).toBeLessThan(firstAxIdIdx);
  });

  it("places <c:marker> after <c:dropLines> / <c:hiLowLines> inside <c:lineChart>", () => {
    // CT_LineChart: ... dLbls?, dropLines?, hiLowLines?, upDownBars?,
    // marker?, axId+. <c:marker> must sit after every line-only
    // optional block.
    const result = writeChart(
      makeChart({
        type: "line",
        dropLines: true,
        hiLowLines: true,
        showLineMarkers: false,
      }),
      "Sheet1",
    );
    const dropLinesIdx = result.chartXml.indexOf("c:dropLines");
    const hiLowLinesIdx = result.chartXml.indexOf("c:hiLowLines");
    const markerIdx = result.chartXml.indexOf('c:marker val="0"');
    expect(dropLinesIdx).toBeLessThan(markerIdx);
    expect(hiLowLinesIdx).toBeLessThan(markerIdx);
  });

  it("ignores showLineMarkers on bar / column / pie / doughnut / area / scatter chart kinds", () => {
    // The OOXML schema places the chart-level <c:marker> (CT_Boolean)
    // exclusively on CT_LineChart. Every other family must drop the
    // flag silently — the writer never authors a <c:marker> child on
    // <c:barChart> / <c:pieChart> / <c:doughnutChart> / <c:areaChart>
    // / <c:scatterChart> regardless of the flag.
    for (const type of ["bar", "column", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, showLineMarkers: false }), "Sheet1");
      // Only per-series <c:marker> (with style/size children, used on
      // scatter) should ever appear on these families — and column /
      // bar / pie / doughnut / area never emit <c:marker> at all. Confirm
      // the chart-level <c:marker val="0"/> shape never leaks here.
      expect(result.chartXml).not.toContain('<c:marker val="0"/>');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        showLineMarkers: false,
      }),
      "Sheet1",
    );
    // Scatter charts emit per-series <c:marker> blocks (CT_Marker with
    // child elements), but never the chart-level CT_Boolean variant.
    expect(scatter.chartXml).not.toContain('<c:marker val="0"/>');
    expect(scatter.chartXml).not.toContain('<c:marker val="1"/>');
  });

  it("collapses non-boolean showLineMarkers values to the default", () => {
    // A stray non-boolean leaking through the type guard should fall
    // back to the default <c:marker val="1"/> rather than emit a token
    // Excel rejects.
    const result = writeChart(
      makeChart({ type: "line", showLineMarkers: "off" as never }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<c:marker val="1"/>');
    expect(result.chartXml).not.toContain('<c:marker val="0"/>');
  });

  it("round-trips showLineMarkers=false through parseChart", () => {
    const written = writeChart(
      makeChart({ type: "line", showLineMarkers: false }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.showLineMarkers).toBe(false);
  });

  it("collapses an unset showLineMarkers round-trip back to undefined", () => {
    // A fresh line chart writes <c:marker val="1"/> (the Excel default)
    // and re-parses to undefined — absence and the default round-trip
    // identically per the asymmetric collapse contract.
    const written = writeChart(makeChart({ type: "line" }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.showLineMarkers).toBeUndefined();
  });

  it("collapses showLineMarkers=true round-trip back to undefined (default-state)", () => {
    // An explicit `true` writes the same <c:marker val="1"/> the
    // default emits, so the re-parsed value collapses to undefined for
    // symmetry — only the non-default `false` survives the round-trip.
    const written = writeChart(
      makeChart({ type: "line", showLineMarkers: true }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.showLineMarkers).toBeUndefined();
  });

  it("composes with per-series marker blocks (gate vs. styling)", () => {
    // Chart-level <c:marker val="0"/> gates rendering across every
    // series; per-series <c:marker> still picks the symbol / size /
    // fill. Both blocks should coexist in the output — the chart-level
    // gate sits inside <c:lineChart>, per-series markers sit inside
    // <c:ser>.
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "Revenue",
            values: "B2:B4",
            categories: "A2:A4",
            marker: { symbol: "circle", size: 6 },
          },
        ],
        showLineMarkers: false,
      }),
      "Sheet1",
    );
    // Chart-level gate.
    expect(result.chartXml).toContain('<c:marker val="0"/>');
    // Per-series marker block (CT_Marker with children, not CT_Boolean).
    expect(result.chartXml).toContain('<c:symbol val="circle"/>');
    expect(result.chartXml).toContain('<c:size val="6"/>');
  });

  it("threads showLineMarkers=false through writeXlsx end-to-end packaging", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 100],
          ["Q2", 200],
          ["Q3", 150],
        ],
        charts: [
          {
            type: "line",
            series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            showLineMarkers: false,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('<c:marker val="0"/>');
    // Re-parse the rendered chart to confirm the flag survives the
    // packaging path.
    const reparsed = parseChart(chartXml);
    expect(reparsed?.showLineMarkers).toBe(false);
  });

  it("threads showLineMarkers alongside upDownBars / dropLines / hiLowLines", async () => {
    // Compose every line-only optional block on the same chart to
    // guard the schema-order contract end-to-end (writer → re-parse).
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Day", "High", "Low"],
          ["Mon", 12, 5],
          ["Tue", 14, 6],
          ["Wed", 10, 4],
        ],
        charts: [
          {
            type: "line",
            series: [
              { name: "High", values: "B2:B4", categories: "A2:A4" },
              { name: "Low", values: "C2:C4", categories: "A2:A4" },
            ],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            dropLines: true,
            hiLowLines: true,
            upDownBars: true,
            showLineMarkers: false,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    // <c:dropLines> / <c:hiLowLines> are emitted bare (self-closing
    // when XML serialization collapses an empty element); accept both
    // shapes via a regex that does not pin the closing token.
    expect(chartXml).toMatch(/<c:dropLines\s*\/>|<c:dropLines>/);
    expect(chartXml).toMatch(/<c:hiLowLines\s*\/>|<c:hiLowLines>/);
    expect(chartXml).toContain("<c:upDownBars>");
    expect(chartXml).toContain('<c:marker val="0"/>');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.dropLines).toBe(true);
    expect(reparsed?.hiLowLines).toBe(true);
    expect(reparsed?.upDownBars).toBe(true);
    expect(reparsed?.showLineMarkers).toBe(false);
  });
});

// ── writeChart — legend entries ──────────────────────────────────────

describe("writeChart — legend entries", () => {
  it("omits <c:legendEntry> when the field is unset (default)", () => {
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("c:legendEntry");
  });

  it("omits <c:legendEntry> when the field is an empty array", () => {
    // Empty arrays collapse to no element so an opt-out via `[]` produces
    // the same wire shape as omitting the field entirely.
    const result = writeChart(makeChart({ legendEntries: [] }), "Sheet1");
    expect(result.chartXml).not.toContain("c:legendEntry");
  });

  it("emits a single hidden entry with both <c:idx> and <c:delete> children", () => {
    // The writer always emits both children so a re-parse sees the
    // canonical CT_LegendEntry shape (matches Excel's reference
    // serialization on the "Hide legend entry" action).
    const result = writeChart(
      makeChart({
        series: [
          { name: "A", values: "B2:B4", categories: "A2:A4" },
          { name: "B", values: "C2:C4", categories: "A2:A4" },
        ],
        legendEntries: [{ idx: 1, delete: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain(
      '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>',
    );
  });

  it("emits multiple entries in ascending idx order", () => {
    // Excel's reference serialization sorts by `<c:idx>` on write even
    // when the in-memory list arrived unsorted; the writer mirrors that
    // so the rendered file is canonical.
    const result = writeChart(
      makeChart({
        series: [
          { name: "A", values: "B2:B4", categories: "A2:A4" },
          { name: "B", values: "C2:C4", categories: "A2:A4" },
          { name: "C", values: "D2:D4", categories: "A2:A4" },
        ],
        legendEntries: [
          { idx: 2, delete: true },
          { idx: 0, delete: true },
        ],
      }),
      "Sheet1",
    );
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    const first = legend.indexOf('<c:idx val="0"/>');
    const second = legend.indexOf('<c:idx val="2"/>');
    expect(first).toBeGreaterThanOrEqual(0);
    expect(second).toBeGreaterThanOrEqual(0);
    expect(first).toBeLessThan(second);
  });

  it('emits an explicit <c:delete val="0"/> when delete is false (round-trip carries the selector)', () => {
    // A "selector-only" entry with `delete: false` still emits both
    // children so a re-parse sees the same shape — the writer never
    // collapses the <c:delete> to an absent element.
    const result = writeChart(makeChart({ legendEntries: [{ idx: 0, delete: false }] }), "Sheet1");
    expect(result.chartXml).toContain(
      '<c:legendEntry><c:idx val="0"/><c:delete val="0"/></c:legendEntry>',
    );
  });

  it('emits <c:delete val="0"/> when delete is omitted (defaults to false)', () => {
    const result = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ legendEntries: [{ idx: 0 } as any] }),
      "Sheet1",
    );
    expect(result.chartXml).toContain(
      '<c:legendEntry><c:idx val="0"/><c:delete val="0"/></c:legendEntry>',
    );
  });

  it("places <c:legendEntry> between <c:legendPos> and <c:overlay> (CT_Legend order)", () => {
    // CT_Legend sequence: legendPos?, legendEntry*, layout?, overlay?, ...
    const result = writeChart(makeChart({ legendEntries: [{ idx: 0, delete: true }] }), "Sheet1");
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend.indexOf("c:legendPos")).toBeLessThan(legend.indexOf("c:legendEntry"));
    expect(legend.indexOf("c:legendEntry")).toBeLessThan(legend.indexOf("c:overlay"));
  });

  it("does not emit any <c:legend> when legend=false (entries dropped)", () => {
    // A hidden legend has no slot for entries — the writer suppresses
    // the entire legend element rather than emit stray entries Excel
    // would never read.
    const result = writeChart(
      makeChart({ legend: false, legendEntries: [{ idx: 0, delete: true }] }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:legend");
    expect(result.chartXml).not.toContain("c:legendEntry");
  });

  it("drops entries whose idx is non-finite / negative / non-integer", () => {
    const result = writeChart(
      makeChart({
        legendEntries: [
          { idx: Number.NaN, delete: true },
          { idx: Number.POSITIVE_INFINITY, delete: true },
          { idx: -1, delete: true },
          { idx: 1.5, delete: true },
          { idx: 0, delete: true },
        ],
      }),
      "Sheet1",
    );
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    const occurrences = legend.match(/<c:legendEntry/g) ?? [];
    expect(occurrences).toHaveLength(1);
    expect(legend).toContain('<c:idx val="0"/>');
  });

  it("deduplicates duplicate idx values (last wins on write)", () => {
    // The writer's resolve pass dedupes with last-wins semantics so a
    // clone-through that appends an override naturally beats the
    // source's parsed value.
    const result = writeChart(
      makeChart({
        legendEntries: [
          { idx: 0, delete: true },
          { idx: 0, delete: false },
        ],
      }),
      "Sheet1",
    );
    const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend.match(/<c:legendEntry/g) ?? []).toHaveLength(1);
    expect(legend).toContain('<c:delete val="0"/>');
  });

  it("treats non-boolean delete values as false", () => {
    const result = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ legendEntries: [{ idx: 0, delete: "yes" as any }] }),
      "Sheet1",
    );
    expect(result.chartXml).toContain(
      '<c:legendEntry><c:idx val="0"/><c:delete val="0"/></c:legendEntry>',
    );
  });

  it("threads legendEntries through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(
        makeChart({ type, legendEntries: [{ idx: 0, delete: true }] }),
        "Sheet1",
      );
      const legend = result.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
      expect(legend).toContain(
        '<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>',
      );
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        legendEntries: [{ idx: 0, delete: true }],
      }),
      "Sheet1",
    );
    const legend = scatter.chartXml.match(/<c:legend>[\s\S]*?<\/c:legend>/)![0];
    expect(legend).toContain('<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>');
  });

  it("round-trips a non-empty legendEntries list through parseChart", () => {
    const written = writeChart(
      makeChart({
        series: [
          { name: "A", values: "B2:B4", categories: "A2:A4" },
          { name: "B", values: "C2:C4", categories: "A2:A4" },
        ],
        legendEntries: [
          { idx: 1, delete: true },
          { idx: 0, delete: false },
        ],
      }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    // The writer emits in ascending idx order, so the re-parsed list is
    // sorted regardless of the input order.
    expect(reparsed?.legendEntries).toEqual([
      { idx: 0, delete: false },
      { idx: 1, delete: true },
    ]);
  });

  it("collapses an unset legendEntries field to undefined on round-trip", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    expect(parseChart(written)?.legendEntries).toBeUndefined();
  });

  it("survives a writeXlsx round trip — legendEntries lands in the packaged chart XML", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales", "Forecast"],
          ["North", 100, 120],
          ["South", 200, 180],
        ],
        charts: [
          {
            type: "column",
            title: "Sales vs Forecast",
            series: [
              { name: "Sales", values: "B2:B3", categories: "A2:A3" },
              { name: "Forecast", values: "C2:C3", categories: "A2:A3" },
            ],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            // Hide the forecast helper series from the legend without
            // dropping it from the plot.
            legendEntries: [{ idx: 1, delete: true }],
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain(
      '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>',
    );
    const reparsed = parseChart(chartXml);
    expect(reparsed?.legendEntries).toEqual([{ idx: 1, delete: true }]);
  });
});

// ── writeChart — autoTitleDeleted ───────────────────────────────────

describe("writeChart — autoTitleDeleted", () => {
  it('emits <c:autoTitleDeleted val="0"/> on a chart with a literal title (derived default)', () => {
    // Back-compat default: a chart with a literal title pins `val="0"`
    // so Excel keeps the literal visible. The writer always emits the
    // element — Excel itself includes it on every reference
    // serialization.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).toContain('<c:autoTitleDeleted val="0"/>');
    expect(result.chartXml).not.toContain('<c:autoTitleDeleted val="1"/>');
  });

  it('emits <c:autoTitleDeleted val="1"/> on a chart with no literal title (derived default)', () => {
    // Back-compat default: a titleless chart pins `val="1"` so Excel
    // does not silently grow an auto-title from the series name.
    const result = writeChart(makeChart({ title: undefined }), "Sheet1");
    expect(result.chartXml).toContain('<c:autoTitleDeleted val="1"/>');
    expect(result.chartXml).not.toContain('<c:autoTitleDeleted val="0"/>');
  });

  it('emits <c:autoTitleDeleted val="1"/> when showTitle=false suppresses the title', () => {
    // `showTitle: false` suppresses the title block entirely — same
    // wire shape as a titleless chart: pin the flag so Excel does not
    // auto-synthesise one.
    const result = writeChart(makeChart({ showTitle: false }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:title>");
    expect(result.chartXml).toContain('<c:autoTitleDeleted val="1"/>');
  });

  it("threads autoTitleDeleted=true explicitly through to <c:autoTitleDeleted>", () => {
    // Pinning `true` overrides the title-presence-derived default — a
    // chart with a literal title that should also pin
    // `autoTitleDeleted=true` (an unusual but legal state — Excel
    // emits it when a literal title is added but the auto-suppression
    // flag was already on) emits both the title block and the flag.
    const result = writeChart(makeChart({ autoTitleDeleted: true }), "Sheet1");
    expect(result.chartXml).toContain("<c:title>");
    expect(result.chartXml).toContain('<c:autoTitleDeleted val="1"/>');
    expect(result.chartXml).not.toContain('<c:autoTitleDeleted val="0"/>');
  });

  it("threads autoTitleDeleted=false explicitly through to <c:autoTitleDeleted> (titleless chart)", () => {
    // A titleless single-series chart can pin `false` to let Excel
    // synthesise the series-name auto-title. The override beats the
    // title-presence-derived default.
    const result = writeChart(makeChart({ title: undefined, autoTitleDeleted: false }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:title>");
    expect(result.chartXml).toContain('<c:autoTitleDeleted val="0"/>');
    expect(result.chartXml).not.toContain('<c:autoTitleDeleted val="1"/>');
  });

  it("places <c:autoTitleDeleted> after <c:title> per CT_Chart sequence", () => {
    // CT_Chart sequence (ECMA-376 Part 1, §21.2.2.4):
    //   title?, autoTitleDeleted?, pivotFmts?, view3D?, ..., plotArea, ...
    const result = writeChart(makeChart(), "Sheet1");
    const titleIdx = result.chartXml.indexOf("<c:title>");
    const autoIdx = result.chartXml.indexOf("<c:autoTitleDeleted");
    const plotAreaIdx = result.chartXml.indexOf("<c:plotArea>");
    expect(titleIdx).toBeGreaterThanOrEqual(0);
    expect(autoIdx).toBeGreaterThan(titleIdx);
    expect(autoIdx).toBeLessThan(plotAreaIdx);
  });

  it("places <c:autoTitleDeleted> before <c:plotArea> on a titleless chart", () => {
    // No `<c:title>` in the wire shape — the flag still slots before
    // the plot area per CT_Chart.
    const result = writeChart(makeChart({ title: undefined }), "Sheet1");
    const autoIdx = result.chartXml.indexOf("<c:autoTitleDeleted");
    const plotAreaIdx = result.chartXml.indexOf("<c:plotArea>");
    expect(autoIdx).toBeGreaterThan(0);
    expect(autoIdx).toBeLessThan(plotAreaIdx);
  });

  it("emits exactly one <c:autoTitleDeleted> per chart (no double-emit on override)", () => {
    // Guard against a regression that would emit the derived default
    // alongside the override. Only one element should appear.
    const result = writeChart(makeChart({ autoTitleDeleted: true }), "Sheet1");
    const occurrences = result.chartXml.match(/c:autoTitleDeleted/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads autoTitleDeleted through every chart family", () => {
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(
        makeChart({ type, title: undefined, autoTitleDeleted: false }),
        "Sheet1",
      );
      expect(result.chartXml).toContain('<c:autoTitleDeleted val="0"/>');
    }
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        title: undefined,
        autoTitleDeleted: false,
        series: [{ values: "B2:B4", categories: "A2:A4" }],
      }),
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('<c:autoTitleDeleted val="0"/>');
  });

  it("composes independently with titleOverlay (different parents)", () => {
    // `<c:autoTitleDeleted>` lives on `<c:chart>`; `<c:overlay>` lives
    // on `<c:title>`. Pinning one must not change the other.
    const result = writeChart(makeChart({ autoTitleDeleted: true, titleOverlay: true }), "Sheet1");
    expect(result.chartXml).toContain('<c:autoTitleDeleted val="1"/>');
    const title = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(title).toContain('c:overlay val="1"');
  });

  it("ignores non-boolean autoTitleDeleted values (falls back to derived default)", () => {
    // Match how `titleOverlay` / `roundedCorners` / `plotVisOnly` treat
    // their inputs — only literal `true` / `false` honour the override.
    // A stray non-boolean (e.g. truthy string) collapses to the
    // title-presence-derived default.
    const result = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ autoTitleDeleted: "yes" as any }),
      "Sheet1",
    );
    // Title is present on the default makeChart, so derived default is
    // val="0" — even with a truthy string override, the flag stays at
    // the derived default.
    expect(result.chartXml).toContain('<c:autoTitleDeleted val="0"/>');
  });

  it("round-trips a non-default autoTitleDeleted=true through parseChart", () => {
    // A chart that pins `autoTitleDeleted=true` on a titled chart
    // should re-parse into a Chart whose `autoTitleDeleted` field is
    // `true` (not collapsed to undefined since true is not the OOXML
    // default).
    const written = writeChart(makeChart({ autoTitleDeleted: true }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.autoTitleDeleted).toBe(true);
  });

  it("collapses a defaulted autoTitleDeleted=false round-trip back to undefined", () => {
    // A chart with a literal title (default makeChart) writes `val="0"`
    // and re-parses to undefined — absence and the OOXML default
    // round-trip identically.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.autoTitleDeleted).toBeUndefined();
  });

  it("re-parses autoTitleDeleted=true on a titleless chart back to true", () => {
    // The titleless-chart derived default writes `val="1"`; the reader
    // surfaces it because the value differs from the OOXML default.
    const written = writeChart(makeChart({ title: undefined }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.autoTitleDeleted).toBe(true);
  });

  it("survives a writeXlsx round trip — autoTitleDeleted lands in the packaged chart XML", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Quarterly",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            autoTitleDeleted: true,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('<c:autoTitleDeleted val="1"/>');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.autoTitleDeleted).toBe(true);
  });
});

// ── writeChart — view3D ──────────────────────────────────────────────

describe("writeChart — view3D", () => {
  it("skips <c:view3D> entirely when the field is unset (writer default)", () => {
    // Excel falls back to the per-family default rotation / perspective
    // on a fresh chart — the writer skips the element so the file
    // stays minimal. Absence and the unset default round-trip
    // identically through cloneChart.
    const result = writeChart(makeChart(), "Sheet1");
    expect(result.chartXml).not.toContain("<c:view3D");
  });

  it("emits a bare <c:view3D/> shell when view3D is an empty object", () => {
    // An empty override accepts every per-family default — useful for
    // round-trip parity with templates that author the bare element
    // without pinning any field. The writer emits a self-closing
    // shell so the bytes stay minimal.
    const result = writeChart(makeChart({ view3D: {} }), "Sheet1");
    expect(result.chartXml).toContain("<c:view3D/>");
  });

  it("emits every pinned field as a self-closing child", () => {
    const result = writeChart(
      makeChart({
        view3D: {
          rotX: 15,
          hPercent: 100,
          rotY: 20,
          depthPercent: 100,
          rAngAx: true,
          perspective: 30,
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('<c:rotX val="15"/>');
    expect(result.chartXml).toContain('<c:hPercent val="100"/>');
    expect(result.chartXml).toContain('<c:rotY val="20"/>');
    expect(result.chartXml).toContain('<c:depthPercent val="100"/>');
    expect(result.chartXml).toContain('<c:rAngAx val="1"/>');
    expect(result.chartXml).toContain('<c:perspective val="30"/>');
  });

  it("emits signed rotX values literally (ST_RotX accepts -90..90)", () => {
    // Negative tilts must round-trip — ST_RotX is a signed byte so
    // Excel emits the leading `-` and the writer must too.
    const result = writeChart(makeChart({ view3D: { rotX: -45 } }), "Sheet1");
    expect(result.chartXml).toContain('<c:rotX val="-45"/>');
  });

  it("emits the boundary values of every range (min and max)", () => {
    const min = writeChart(
      makeChart({
        view3D: {
          rotX: -90,
          hPercent: 5,
          rotY: 0,
          depthPercent: 20,
          perspective: 0,
        },
      }),
      "Sheet1",
    );
    expect(min.chartXml).toContain('<c:rotX val="-90"/>');
    expect(min.chartXml).toContain('<c:hPercent val="5"/>');
    expect(min.chartXml).toContain('<c:rotY val="0"/>');
    expect(min.chartXml).toContain('<c:depthPercent val="20"/>');
    expect(min.chartXml).toContain('<c:perspective val="0"/>');

    const max = writeChart(
      makeChart({
        view3D: {
          rotX: 90,
          hPercent: 500,
          rotY: 360,
          depthPercent: 2000,
          perspective: 240,
        },
      }),
      "Sheet1",
    );
    expect(max.chartXml).toContain('<c:rotX val="90"/>');
    expect(max.chartXml).toContain('<c:hPercent val="500"/>');
    expect(max.chartXml).toContain('<c:rotY val="360"/>');
    expect(max.chartXml).toContain('<c:depthPercent val="2000"/>');
    expect(max.chartXml).toContain('<c:perspective val="240"/>');
  });

  it("drops out-of-range numeric fields rather than emit a token Excel rejects", () => {
    // Each numeric field is bound by an OOXML simple type. Out-of-
    // range inputs drop silently rather than emit a token Excel's
    // strict validator would reject. The other pinned fields still
    // emit.
    const result = writeChart(
      makeChart({
        view3D: {
          rotX: 180,
          hPercent: 3,
          rotY: 400,
          depthPercent: 10,
          perspective: 500,
          rAngAx: true,
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:rotX");
    expect(result.chartXml).not.toContain("<c:hPercent");
    expect(result.chartXml).not.toContain("<c:rotY");
    expect(result.chartXml).not.toContain("<c:depthPercent");
    expect(result.chartXml).not.toContain("<c:perspective");
    // rAngAx still emits — only the numeric children dropped.
    expect(result.chartXml).toContain('<c:rAngAx val="1"/>');
  });

  it("drops fractional values rather than truncate", () => {
    // Every CT_View3D numeric child is an integer simple type. A
    // fractional input drops rather than silently truncate — `parseInt`
    // would coerce 15.5 → 15 and the round-trip would silently lose
    // the fractional value the caller passed.
    const result = writeChart(makeChart({ view3D: { rotX: 15.5, hPercent: 100 } }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:rotX");
    expect(result.chartXml).toContain('<c:hPercent val="100"/>');
  });

  it("drops non-finite values (NaN, Infinity, -Infinity)", () => {
    const result = writeChart(
      makeChart({
        view3D: {
          rotX: NaN,
          hPercent: Infinity,
          rotY: -Infinity,
          depthPercent: 100,
        },
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:rotX");
    expect(result.chartXml).not.toContain("<c:hPercent");
    expect(result.chartXml).not.toContain("<c:rotY");
    expect(result.chartXml).toContain('<c:depthPercent val="100"/>');
  });

  it('emits rAngAx=false literally (val="0") for round-trip parity', () => {
    // Explicit `false` round-trips as `<c:rAngAx val="0"/>` so the
    // caller can pin the OOXML default literally — useful for parity
    // with templates that author the explicit value.
    const result = writeChart(makeChart({ view3D: { rAngAx: false } }), "Sheet1");
    expect(result.chartXml).toContain('<c:rAngAx val="0"/>');
  });

  it("ignores a non-boolean rAngAx (drops the child rather than emit invalid token)", () => {
    // A non-boolean leaking through the type guard drops the child —
    // mirrors how every other chart-level boolean writer treats its
    // input.
    const result = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ view3D: { rAngAx: 1 as any, rotX: 15 } }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("<c:rAngAx");
    expect(result.chartXml).toContain('<c:rotX val="15"/>');
  });

  it("emits the six CT_View3D children in schema order", () => {
    // CT_View3D sequence per ECMA-376 Part 1, §21.2.2.228:
    //   rotX?, hPercent?, rotY?, depthPercent?, rAngAx?, perspective?
    const result = writeChart(
      makeChart({
        view3D: {
          rotX: 15,
          hPercent: 100,
          rotY: 20,
          depthPercent: 100,
          rAngAx: true,
          perspective: 30,
        },
      }),
      "Sheet1",
    );
    const rotXIdx = result.chartXml.indexOf("<c:rotX");
    const hPercentIdx = result.chartXml.indexOf("<c:hPercent");
    const rotYIdx = result.chartXml.indexOf("<c:rotY");
    const depthPercentIdx = result.chartXml.indexOf("<c:depthPercent");
    const rAngAxIdx = result.chartXml.indexOf("<c:rAngAx");
    const perspectiveIdx = result.chartXml.indexOf("<c:perspective");
    expect(rotXIdx).toBeGreaterThan(-1);
    expect(hPercentIdx).toBeGreaterThan(rotXIdx);
    expect(rotYIdx).toBeGreaterThan(hPercentIdx);
    expect(depthPercentIdx).toBeGreaterThan(rotYIdx);
    expect(rAngAxIdx).toBeGreaterThan(depthPercentIdx);
    expect(perspectiveIdx).toBeGreaterThan(rAngAxIdx);
  });

  it("places <c:view3D> after <c:autoTitleDeleted> and before <c:plotArea> per CT_Chart", () => {
    // CT_Chart sequence (ECMA-376 Part 1, §21.2.2.4):
    //   title?, autoTitleDeleted?, pivotFmts?, view3D?, ..., plotArea, ...
    const result = writeChart(makeChart({ view3D: { rotX: 15 } }), "Sheet1");
    const autoIdx = result.chartXml.indexOf("<c:autoTitleDeleted");
    const view3DIdx = result.chartXml.indexOf("<c:view3D");
    const plotAreaIdx = result.chartXml.indexOf("<c:plotArea>");
    expect(autoIdx).toBeGreaterThan(-1);
    expect(view3DIdx).toBeGreaterThan(autoIdx);
    expect(view3DIdx).toBeLessThan(plotAreaIdx);
  });

  it("only emits <c:view3D> once on a chart that pins it", () => {
    // Guard against any regression that would double-emit the element.
    const result = writeChart(makeChart({ view3D: { rotX: 15 } }), "Sheet1");
    const occurrences = result.chartXml.match(/<c:view3D[\s/>]/g) ?? [];
    expect(occurrences).toHaveLength(1);
  });

  it("threads view3D through every chart family (the schema accepts it on every CT_Chart)", () => {
    // <c:view3D> is only meaningful on 3D families but the OOXML
    // schema accepts it on every CT_Chart. The writer emits it on
    // every family the caller pins it on — Excel silently ignores it
    // on 2D families.
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, view3D: { rotX: 10 } }), "Sheet1");
      expect(result.chartXml).toContain('<c:rotX val="10"/>');
    }
    for (const type of ["pie", "doughnut"] as const) {
      const result = writeChart(
        {
          type,
          series: [{ values: "B2:B4", categories: "A2:A4" }],
          anchor: { from: { row: 5, col: 0 } },
          view3D: { rotX: 10 },
        },
        "Sheet1",
      );
      expect(result.chartXml).toContain('<c:rotX val="10"/>');
    }
    const scatter = writeChart(
      {
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        anchor: { from: { row: 5, col: 0 } },
        view3D: { rotX: 10 },
      },
      "Sheet1",
    );
    expect(scatter.chartXml).toContain('<c:rotX val="10"/>');
  });

  it("round-trips a fully-pinned view3D through parseChart", () => {
    const written = writeChart(
      makeChart({
        view3D: {
          rotX: 15,
          hPercent: 100,
          rotY: 20,
          depthPercent: 100,
          rAngAx: true,
          perspective: 30,
        },
      }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.view3D).toEqual({
      rotX: 15,
      hPercent: 100,
      rotY: 20,
      depthPercent: 100,
      rAngAx: true,
      perspective: 30,
    });
  });

  it("round-trips a partial view3D (only rotation pinned)", () => {
    // Common pattern — pin the rotation only, leave height / depth /
    // perspective at the per-family default. The reader surfaces the
    // exact partial shape the writer emitted.
    const written = writeChart(makeChart({ view3D: { rotX: 20, rotY: 40 } }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.view3D).toEqual({ rotX: 20, rotY: 40 });
  });

  it("round-trips a bare <c:view3D/> shell as an empty object", () => {
    // The empty-object input collapses to the bare self-closing
    // element on emit, and the reader surfaces it as an empty object
    // — closes the round-trip loop on the gating-element shape.
    const written = writeChart(makeChart({ view3D: {} }), "Sheet1").chartXml;
    expect(written).toContain("<c:view3D/>");
    const reparsed = parseChart(written);
    expect(reparsed?.view3D).toEqual({});
  });

  it("threads view3D through writeXlsx into xl/charts/chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "3D Look",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            view3D: { rotX: 15, rotY: 20, perspective: 30 },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain("<c:view3D>");
    expect(chartXml).toContain('<c:rotX val="15"/>');
    expect(chartXml).toContain('<c:rotY val="20"/>');
    expect(chartXml).toContain('<c:perspective val="30"/>');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.view3D).toEqual({ rotX: 15, rotY: 20, perspective: 30 });
  });
});

// ── writeChart — axis label rotation ────────────────────────────────

describe("writeChart — axis labelRotation", () => {
  it("omits <c:txPr> on the category axis when the rotation is unset", () => {
    // Excel's reference serialization omits `<c:txPr>` on every axis
    // whose labels render flat — the writer mirrors that contract so a
    // fresh chart stays byte-clean against Excel's own output.
    const result = writeChart(makeChart(), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).not.toContain("<c:txPr>");
  });

  it('emits <c:txPr><a:bodyPr rot="N"/></c:txPr> when the X axis pins a rotation', () => {
    const result = writeChart(makeChart({ axes: { x: { labelRotation: 45 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain("<c:txPr>");
    // 45 degrees * 60000 = 2,700,000.
    expect(catAxBlock).toContain('<a:bodyPr rot="2700000"/>');
  });

  it("converts negative rotations to negative 60000ths of a degree", () => {
    const result = writeChart(makeChart({ axes: { x: { labelRotation: -45 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('<a:bodyPr rot="-2700000"/>');
  });

  it("emits the rotation on the Y axis (value axis)", () => {
    const result = writeChart(makeChart({ axes: { y: { labelRotation: 30 } } }), "Sheet1");
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(valAxBlock).toContain("<c:txPr>");
    expect(valAxBlock).toContain('<a:bodyPr rot="1800000"/>');
  });

  it("emits independently on both axes when both pin a rotation", () => {
    const result = writeChart(
      makeChart({ axes: { x: { labelRotation: 45 }, y: { labelRotation: -30 } } }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const valAxBlock = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
    expect(catAxBlock).toContain('<a:bodyPr rot="2700000"/>');
    expect(valAxBlock).toContain('<a:bodyPr rot="-1800000"/>');
  });

  it("collapses the OOXML default 0 to absence (no <c:txPr> emitted)", () => {
    // Pinning the default `0` round-trips identically to absence — the
    // writer skips the entire `<c:txPr>` block. Mirrors how every other
    // axis-default-collapse field treats its input.
    const result = writeChart(makeChart({ axes: { x: { labelRotation: 0 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).not.toContain("<c:txPr>");
  });

  it("clamps rotations above 90 to the 90-degree maximum", () => {
    const result = writeChart(makeChart({ axes: { x: { labelRotation: 180 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    // 90 degrees * 60000 = 5,400,000.
    expect(catAxBlock).toContain('<a:bodyPr rot="5400000"/>');
  });

  it("clamps rotations below -90 to the -90-degree minimum", () => {
    const result = writeChart(makeChart({ axes: { x: { labelRotation: -180 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('<a:bodyPr rot="-5400000"/>');
  });

  it("rounds non-integer degree values to the nearest whole degree", () => {
    // The OOXML attribute is an integer in 60000ths of a degree, so a
    // fractional whole-degree input has no meaningful refinement at
    // emit time. Rounding keeps the wire output stable.
    const result = writeChart(makeChart({ axes: { x: { labelRotation: 45.4 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain('<a:bodyPr rot="2700000"/>');
    const result2 = writeChart(makeChart({ axes: { x: { labelRotation: 45.6 } } }), "Sheet1");
    const catAxBlock2 = result2.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock2).toContain('<a:bodyPr rot="2760000"/>');
  });

  it("drops non-finite rotation inputs (NaN, Infinity)", () => {
    const nan = writeChart(makeChart({ axes: { x: { labelRotation: Number.NaN } } }), "Sheet1");
    const inf = writeChart(
      makeChart({ axes: { x: { labelRotation: Number.POSITIVE_INFINITY } } }),
      "Sheet1",
    );
    expect(nan.chartXml).not.toContain("<c:txPr>");
    expect(inf.chartXml).not.toContain("<c:txPr>");
  });

  it("drops non-numeric rotation inputs (string, boolean)", () => {
    // The type guard rejects non-numeric inputs at the normalizer
    // boundary — the writer never emits a value Excel's strict
    // validator would reject.
    const stringy = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ axes: { x: { labelRotation: "45" as any } } }),
      "Sheet1",
    );
    const boolish = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ axes: { x: { labelRotation: true as any } } }),
      "Sheet1",
    );
    expect(stringy.chartXml).not.toContain("<c:txPr>");
    expect(boolish.chartXml).not.toContain("<c:txPr>");
  });

  it("places <c:txPr> between <c:tickLblPos> and <c:crossAx> per the OOXML schema", () => {
    // CT_CatAx: ... tickLblPos -> spPr? -> txPr? -> crossAx -> ...
    const result = writeChart(
      makeChart({
        axes: { x: { tickLblPos: "low", labelRotation: 45 } },
      }),
      "Sheet1",
    );
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    const tickLblPosIdx = catAxBlock.indexOf("c:tickLblPos");
    const txPrIdx = catAxBlock.indexOf("<c:txPr>");
    const crossAxIdx = catAxBlock.indexOf("c:crossAx");
    expect(tickLblPosIdx).toBeGreaterThan(0);
    expect(txPrIdx).toBeGreaterThan(tickLblPosIdx);
    expect(crossAxIdx).toBeGreaterThan(txPrIdx);
  });

  it("threads the rotation through bar, column, line, and area chart families", () => {
    for (const type of ["bar", "column", "line", "area"] as const) {
      const result = writeChart(makeChart({ type, axes: { x: { labelRotation: 30 } } }), "Sheet1");
      expect(result.chartXml).toContain('<a:bodyPr rot="1800000"/>');
    }
  });

  it("threads the rotation through scatter charts (both axes are value axes)", () => {
    const result = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        axes: { x: { labelRotation: 45 }, y: { labelRotation: -45 } },
      }),
      "Sheet1",
    );
    // Scatter has two `<c:valAx>` siblings — confirm both pins survive.
    const valAxes = result.chartXml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    expect(valAxes).toHaveLength(2);
    expect(valAxes[0]).toContain('<a:bodyPr rot="2700000"/>');
    expect(valAxes[1]).toContain('<a:bodyPr rot="-2700000"/>');
  });

  it("ignores the rotation on pie / doughnut charts (no axes at all)", () => {
    const pie = writeChart(
      makeChart({ type: "pie", axes: { x: { labelRotation: 45 } } }),
      "Sheet1",
    );
    const dough = writeChart(
      makeChart({ type: "doughnut", axes: { x: { labelRotation: 45 } } }),
      "Sheet1",
    );
    expect(pie.chartXml).not.toContain("<c:txPr>");
    expect(dough.chartXml).not.toContain("<c:txPr>");
  });

  it("emits exactly one <c:txPr> per axis", () => {
    const result = writeChart(
      makeChart({ axes: { x: { labelRotation: 45 }, y: { labelRotation: 30 } } }),
      "Sheet1",
    );
    expect((result.chartXml.match(/<c:txPr>/g) ?? []).length).toBe(2);
  });

  it("emits the minimal <c:txPr> shape (bodyPr + lstStyle + p)", () => {
    const result = writeChart(makeChart({ axes: { x: { labelRotation: 45 } } }), "Sheet1");
    const catAxBlock = result.chartXml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
    expect(catAxBlock).toContain("<a:lstStyle/>");
    expect(catAxBlock).toMatch(/<a:p>.*<\/a:p>/);
    expect(catAxBlock).toContain("<a:endParaRPr");
  });

  it("round-trips a non-default rotation through parseChart", () => {
    const written = writeChart(
      makeChart({ axes: { x: { labelRotation: 45 } } }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.labelRotation).toBe(45);
  });

  it("collapses a defaulted rotation round-trip back to undefined", () => {
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.x?.labelRotation).toBeUndefined();
  });

  it("end-to-end: writeXlsx packages the rotation into chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            axes: { x: { labelRotation: -45 } },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    expect(chartXml).toContain('<a:bodyPr rot="-2700000"/>');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.axes?.x?.labelRotation).toBe(-45);
  });
});

// ── writeChart — data labels showLeaderLines ─────────────────────────

describe("writeChart — data labels showLeaderLines", () => {
  function dLblsOf(xml: string): string {
    const m = xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/);
    if (!m) throw new Error("No <c:dLbls> block found in chart XML");
    return m[0];
  }

  it("does not emit <c:showLeaderLines> on a pie chart by default (OOXML default true)", () => {
    // Excel's reference serialization omits the element when the user
    // did not flip the toggle off — the OOXML default is `true`.
    // Mirroring that on emit keeps absence and the default
    // round-tripping identically through parseChart.
    const result = writeChart(
      makeChart({ type: "pie", dataLabels: { showValue: true } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).not.toContain("<c:showLeaderLines");
  });

  it('emits <c:showLeaderLines val="0"/> on a pie chart when showLeaderLines=false', () => {
    const result = writeChart(
      makeChart({ type: "pie", dataLabels: { showValue: true, showLeaderLines: false } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:showLeaderLines val="0"/>');
  });

  it('emits <c:showLeaderLines val="0"/> on a doughnut chart when showLeaderLines=false', () => {
    const result = writeChart(
      makeChart({
        type: "doughnut",
        dataLabels: { showValue: true, showLeaderLines: false },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:showLeaderLines val="0"/>');
  });

  it("treats showLeaderLines=true the same as omitting the field (no element emitted)", () => {
    // Pinning the OOXML default has the same wire shape as omitting
    // the field — only an explicit `false` flips the toggle.
    const explicit = writeChart(
      makeChart({ type: "pie", dataLabels: { showValue: true, showLeaderLines: true } }),
      "Sheet1",
    ).chartXml;
    const implicit = writeChart(
      makeChart({ type: "pie", dataLabels: { showValue: true } }),
      "Sheet1",
    ).chartXml;
    expect(explicit).toEqual(implicit);
  });

  it("collapses non-boolean showLeaderLines inputs to the OOXML default (no element)", () => {
    // A stray non-boolean leaking past the type guard (e.g. 0 / "false"
    // / null) must collapse to the default rather than emit something
    // Excel would reject. Mirrors how the other show* toggles treat
    // their inputs — only literal `false` flips the toggle.
    const result = writeChart(
      makeChart({
        type: "pie",
        dataLabels: { showValue: true, showLeaderLines: 0 as unknown as boolean },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).not.toContain("<c:showLeaderLines");
  });

  it("scope-guards showLeaderLines on bar / column / line / area / scatter (drops the override)", () => {
    // The OOXML schema scopes <c:showLeaderLines> to EG_DLbls (pie /
    // doughnut only). Every other family routes through EG_DLblsShared
    // which omits the element — emitting it would break Excel's
    // strict validator.
    const families: Array<{ type: WriteChartKind; setup?: Partial<SheetChart> }> = [
      { type: "bar" },
      { type: "column" },
      { type: "line" },
      { type: "area" },
      {
        type: "scatter",
        setup: { series: [{ values: "B2:B4", categories: "A2:A4" }] },
      },
    ];
    for (const { type, setup } of families) {
      const result = writeChart(
        makeChart({ type, dataLabels: { showValue: true, showLeaderLines: false }, ...setup }),
        "Sheet1",
      );
      const dLbls = dLblsOf(result.chartXml);
      expect(dLbls).not.toContain("<c:showLeaderLines");
    }
  });

  it("places <c:showLeaderLines> after <c:separator> when both are emitted (CT_DLbls order)", () => {
    // CT_DLbls sequence: ... showBubbleSize, separator?, showLeaderLines?
    const result = writeChart(
      makeChart({
        type: "pie",
        dataLabels: { showValue: true, separator: " | ", showLeaderLines: false },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    const idxSep = dLbls.indexOf("<c:separator");
    const idxLead = dLbls.indexOf("<c:showLeaderLines");
    expect(idxSep).toBeGreaterThan(0);
    expect(idxLead).toBeGreaterThan(idxSep);
  });

  it("places <c:showLeaderLines> after <c:showBubbleSize> when no separator is set", () => {
    const result = writeChart(
      makeChart({
        type: "doughnut",
        dataLabels: { showValue: true, showLeaderLines: false },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    const idxBub = dLbls.indexOf("<c:showBubbleSize");
    const idxLead = dLbls.indexOf("<c:showLeaderLines");
    expect(idxBub).toBeGreaterThan(0);
    expect(idxLead).toBeGreaterThan(idxBub);
  });

  it("emits exactly one <c:showLeaderLines> per <c:dLbls> block", () => {
    // Guard against a regression that would double-emit the element.
    const result = writeChart(
      makeChart({ type: "pie", dataLabels: { showValue: true, showLeaderLines: false } }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect((dLbls.match(/<c:showLeaderLines /g) ?? []).length).toBe(1);
  });

  it("threads showLeaderLines through a series-level <c:dLbls> on a pie chart", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [
          {
            name: "S1",
            values: "B2:B4",
            categories: "A2:A4",
            dataLabels: { showValue: true, showLeaderLines: false },
          },
        ],
      }),
      "Sheet1",
    );
    const xml = result.chartXml;
    const serStart = xml.indexOf("<c:ser>");
    const serEnd = xml.indexOf("</c:ser>");
    const inner = xml.slice(serStart, serEnd);
    expect(inner).toContain('<c:showLeaderLines val="0"/>');
  });

  it("scope-guards series-level showLeaderLines on a column chart", () => {
    // The OOXML schema scopes the element to pie / doughnut on the
    // series-level dLbls just as it does on the chart level. The
    // override must be silently dropped on every other family.
    const result = writeChart(
      makeChart({
        type: "column",
        series: [
          {
            name: "S1",
            values: "B2:B4",
            categories: "A2:A4",
            dataLabels: { showValue: true, showLeaderLines: false },
          },
        ],
      }),
      "Sheet1",
    );
    const xml = result.chartXml;
    const serStart = xml.indexOf("<c:ser>");
    const serEnd = xml.indexOf("</c:ser>");
    const inner = xml.slice(serStart, serEnd);
    expect(inner).not.toContain("<c:showLeaderLines");
  });

  it("composes independently with showLegendKey and numberFormat", () => {
    // The three knobs live on different lines of the CT_DLbls
    // sequence; pinning one must not change the others.
    const result = writeChart(
      makeChart({
        type: "pie",
        dataLabels: {
          showValue: true,
          showLegendKey: true,
          showLeaderLines: false,
          numberFormat: { formatCode: "0.00%" },
        },
      }),
      "Sheet1",
    );
    const dLbls = dLblsOf(result.chartXml);
    expect(dLbls).toContain('<c:numFmt formatCode="0.00%"');
    expect(dLbls).toContain('<c:showLegendKey val="1"/>');
    expect(dLbls).toContain('<c:showLeaderLines val="0"/>');
  });

  it("round-trips a pie chart with showLeaderLines=false through parseChart", () => {
    const written = writeChart(
      makeChart({ type: "pie", dataLabels: { showValue: true, showLeaderLines: false } }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataLabels?.showLeaderLines).toBe(false);
    expect(reparsed?.dataLabels?.showValue).toBe(true);
  });

  it("collapses a defaulted pie showLeaderLines round-trip back to undefined", () => {
    // A fresh pie chart (showLeaderLines omitted) emits no element and
    // re-parses to undefined — absence and the OOXML default round-trip
    // identically.
    const written = writeChart(
      makeChart({ type: "pie", dataLabels: { showValue: true } }),
      "Sheet1",
    ).chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.dataLabels?.showLeaderLines).toBeUndefined();
  });

  it("end-to-end: writeXlsx packages a pie chart with showLeaderLines=false", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Slice", "Value"],
          ["A", 30],
          ["B", 70],
        ],
        charts: [
          {
            type: "pie",
            title: "Distribution",
            series: [{ name: "Distribution", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            dataLabels: { showValue: true, showLeaderLines: false },
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    const dLbls = chartXml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/)![0];
    expect(dLbls).toContain('<c:showLeaderLines val="0"/>');
  });
});

// ── writeChart — title rotation ────────────────────────────────────

describe("writeChart — title rotation", () => {
  function titleBodyPrOf(xml: string): string {
    // The title's <a:bodyPr> lives inside <c:title><c:tx><c:rich>.
    const titleBlock = xml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    const bodyPr = titleBlock.match(/<a:bodyPr[^/]*\/>/)![0];
    return bodyPr;
  }

  it('emits <a:bodyPr rot="0"/> by default (matches the OOXML default)', () => {
    // Excel's reference serialization writes `rot="0"` on every visible
    // title even when no rotation is pinned — the writer mirrors that
    // contract so a fresh chart stays byte-clean against Excel's own
    // output.
    const result = writeChart(makeChart(), "Sheet1");
    expect(titleBodyPrOf(result.chartXml)).toContain('rot="0"');
  });

  it('emits <a:bodyPr rot="N"/> when the chart pins a positive rotation', () => {
    const result = writeChart(makeChart({ titleRotation: 45 }), "Sheet1");
    // 45 degrees * 60000 = 2,700,000.
    expect(titleBodyPrOf(result.chartXml)).toContain('rot="2700000"');
  });

  it("converts negative rotations to negative 60000ths of a degree", () => {
    const result = writeChart(makeChart({ titleRotation: -45 }), "Sheet1");
    expect(titleBodyPrOf(result.chartXml)).toContain('rot="-2700000"');
  });

  it("emits a 90-degree rotation as the band endpoint", () => {
    const result = writeChart(makeChart({ titleRotation: 90 }), "Sheet1");
    // 90 degrees * 60000 = 5,400,000.
    expect(titleBodyPrOf(result.chartXml)).toContain('rot="5400000"');
  });

  it('collapses the OOXML default 0 to absence (writer emits rot="0")', () => {
    // Pinning the default `0` round-trips identically to absence — the
    // writer skips the conversion and emits the OOXML default.
    const result = writeChart(makeChart({ titleRotation: 0 }), "Sheet1");
    expect(titleBodyPrOf(result.chartXml)).toContain('rot="0"');
  });

  it("clamps rotations above 90 to the 90-degree maximum", () => {
    const result = writeChart(makeChart({ titleRotation: 180 }), "Sheet1");
    expect(titleBodyPrOf(result.chartXml)).toContain('rot="5400000"');
  });

  it("clamps rotations below -90 to the -90-degree minimum", () => {
    const result = writeChart(makeChart({ titleRotation: -180 }), "Sheet1");
    expect(titleBodyPrOf(result.chartXml)).toContain('rot="-5400000"');
  });

  it("rounds non-integer degree values to the nearest whole degree", () => {
    // 45.4 rounds to 45 (-> 2,700,000); 45.6 rounds to 46 (-> 2,760,000).
    const a = writeChart(makeChart({ titleRotation: 45.4 }), "Sheet1");
    expect(titleBodyPrOf(a.chartXml)).toContain('rot="2700000"');
    const b = writeChart(makeChart({ titleRotation: 45.6 }), "Sheet1");
    expect(titleBodyPrOf(b.chartXml)).toContain('rot="2760000"');
  });

  it("drops non-finite rotation inputs (NaN, Infinity) back to the default", () => {
    const nan = writeChart(makeChart({ titleRotation: Number.NaN }), "Sheet1");
    const inf = writeChart(makeChart({ titleRotation: Number.POSITIVE_INFINITY }), "Sheet1");
    expect(titleBodyPrOf(nan.chartXml)).toContain('rot="0"');
    expect(titleBodyPrOf(inf.chartXml)).toContain('rot="0"');
  });

  it("drops non-numeric rotation inputs (string, boolean) back to the default", () => {
    const stringy = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ titleRotation: "45" as any }),
      "Sheet1",
    );
    const boolish = writeChart(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      makeChart({ titleRotation: true as any }),
      "Sheet1",
    );
    expect(titleBodyPrOf(stringy.chartXml)).toContain('rot="0"');
    expect(titleBodyPrOf(boolish.chartXml)).toContain('rot="0"');
  });

  it("ignores titleRotation when the chart renders no title (showTitle=false)", () => {
    // No `<c:title>` block ever emits — there is no slot for the
    // rotation in either case. The chart still serializes cleanly.
    const result = writeChart(makeChart({ showTitle: false, titleRotation: 45 }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:title>");
  });

  it("ignores titleRotation when the chart has no literal title", () => {
    const result = writeChart(makeChart({ title: undefined, titleRotation: 45 }), "Sheet1");
    expect(result.chartXml).not.toContain("<c:title>");
  });

  it("composes with titleOverlay (both flags survive on the same title)", () => {
    const result = writeChart(makeChart({ titleOverlay: true, titleRotation: -45 }), "Sheet1");
    const titleBlock = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(titleBlock).toContain('rot="-2700000"');
    expect(titleBlock).toContain('<c:overlay val="1"/>');
  });

  it("threads the rotation through every chart family that emits a title", () => {
    // `<c:title>` lives on `<c:chart>` directly — every chart family
    // (bar / column / line / pie / doughnut / area / scatter) that
    // pins a literal title carries the rotation.
    for (const type of ["bar", "column", "line", "pie", "doughnut", "area"] as const) {
      const result = writeChart(makeChart({ type, titleRotation: 30 }), "Sheet1");
      // 30 degrees * 60000 = 1,800,000.
      expect(titleBodyPrOf(result.chartXml)).toContain('rot="1800000"');
    }
    // Scatter requires a numeric category range.
    const scatter = writeChart(
      makeChart({
        type: "scatter",
        series: [{ values: "B2:B4", categories: "A2:A4" }],
        titleRotation: 30,
      }),
      "Sheet1",
    );
    expect(titleBodyPrOf(scatter.chartXml)).toContain('rot="1800000"');
  });

  it("emits exactly one <a:bodyPr> on the chart title", () => {
    const result = writeChart(makeChart({ titleRotation: 45 }), "Sheet1");
    const titleBlock = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect((titleBlock.match(/<a:bodyPr/g) ?? []).length).toBe(1);
  });

  it("does not emit an axis-style <c:txPr> on the title", () => {
    // The chart title carries its rotation on the rich-text body's
    // `<a:bodyPr>` (inside `<c:tx><c:rich>`), not via an axis-style
    // `<c:txPr>` sibling. Confirm the writer keeps the two paths
    // distinct so a parse round-trip does not double-count.
    const result = writeChart(makeChart({ titleRotation: 45 }), "Sheet1");
    const titleBlock = result.chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(titleBlock).not.toContain("<c:txPr>");
  });

  it("round-trips a non-default rotation through parseChart", () => {
    const written = writeChart(makeChart({ titleRotation: 45 }), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.titleRotation).toBe(45);
  });

  it("collapses a defaulted rotation round-trip back to undefined", () => {
    // A fresh chart emits `rot="0"`; the reader collapses that to
    // `undefined` so absence and the default stay symmetric.
    const written = writeChart(makeChart(), "Sheet1").chartXml;
    const reparsed = parseChart(written);
    expect(reparsed?.titleRotation).toBeUndefined();
  });

  it("end-to-end: writeXlsx packages the title rotation into chart1.xml", async () => {
    const sheets: WriteSheet[] = [
      {
        name: "Sheet1",
        rows: [
          ["Region", "Sales"],
          ["North", 100],
          ["South", 200],
        ],
        charts: [
          {
            type: "column",
            title: "Sales",
            series: [{ name: "Sales", values: "B2:B3", categories: "A2:A3" }],
            anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
            titleRotation: -45,
          },
        ],
      },
    ];
    const out = await writeXlsx({ sheets });
    const chartXml = await extractXml(out, "xl/charts/chart1.xml");
    const titleBlock = chartXml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
    expect(titleBlock).toContain('rot="-2700000"');
    const reparsed = parseChart(chartXml);
    expect(reparsed?.titleRotation).toBe(-45);
  });
});
