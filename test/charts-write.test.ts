import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { writeChart, chartKindElement } from "../src/xlsx/chart-writer";
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

  it.each<WriteChartKind>(["bar", "column", "line", "pie", "scatter", "area"])(
    "kind %s parses as well-formed XML",
    (kind) => {
      const result = writeChart(makeChart({ type: kind }), "Sheet1");
      const doc = parseXml(result.chartXml);
      // Document parses without throwing
      expect(doc).toBeTruthy();
    },
  );
});

describe("chartKindElement", () => {
  it("maps each chart kind to the matching DrawingML element", () => {
    expect(chartKindElement("bar")).toBe("c:barChart");
    expect(chartKindElement("column")).toBe("c:barChart");
    expect(chartKindElement("line")).toBe("c:lineChart");
    expect(chartKindElement("pie")).toBe("c:pieChart");
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
});
