import { describe, it, expect } from "vitest";
import { parseChart } from "../src/xlsx/chart-reader";
import { ZipWriter } from "../src/zip/writer";
import { ZipReader } from "../src/zip/reader";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");

// ── parseChart ────────────────────────────────────────────────────

describe("parseChart", () => {
  it("returns undefined for documents that aren't c:chartSpace", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<root/>`;
    expect(parseChart(xml)).toBeUndefined();
  });

  it("returns kinds=[] when chartSpace has no chart child", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`;
    expect(parseChart(xml)).toEqual({ kinds: [], seriesCount: 0 });
  });

  it("parses a single bar chart with two series", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:ser><c:idx val="1"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)).toEqual({
      kinds: ["bar"],
      seriesCount: 2,
      title: "Sales",
      series: [
        { kind: "bar", index: 0 },
        { kind: "bar", index: 1 },
      ],
    });
  });

  it("collects every chart-type element (combo charts)", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser/>
      </c:barChart>
      <c:lineChart>
        <c:ser/>
        <c:ser/>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)).toEqual({
      kinds: ["bar", "line"],
      seriesCount: 3,
      series: [
        { kind: "bar", index: 0 },
        { kind: "line", index: 0 },
        { kind: "line", index: 1 },
      ],
    });
  });

  it("recognizes pie / doughnut / scatter / area / bubble / radar / surface / stock / 3D", () => {
    for (const [tag, expected] of [
      ["pieChart", "pie"],
      ["pie3DChart", "pie3D"],
      ["doughnutChart", "doughnut"],
      ["scatterChart", "scatter"],
      ["areaChart", "area"],
      ["area3DChart", "area3D"],
      ["bubbleChart", "bubble"],
      ["radarChart", "radar"],
      ["surfaceChart", "surface"],
      ["surface3DChart", "surface3D"],
      ["stockChart", "stock"],
      ["bar3DChart", "bar3D"],
      ["line3DChart", "line3D"],
      ["ofPieChart", "ofPie"],
    ] as const) {
      const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:${tag}/></c:plotArea></c:chart></c:chartSpace>`;
      expect(parseChart(xml)?.kinds).toEqual([expected]);
    }
  });

  it("falls back to strRef cached value when title is a formula", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:title>
      <c:tx>
        <c:strRef>
          <c:f>Sheet1!$A$1</c:f>
          <c:strCache>
            <c:ptCount val="1"/>
            <c:pt idx="0"><c:v>Quarterly Revenue</c:v></c:pt>
          </c:strCache>
        </c:strRef>
      </c:tx>
    </c:title>
    <c:plotArea><c:barChart/></c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.title).toBe("Quarterly Revenue");
  });
});

// ── parseChart — series introspection ─────────────────────────────

describe("parseChart — series introspection", () => {
  it("surfaces series name, value range, category range, and color", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Revenue</c:v></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="1F77B4"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:v>Cost</c:v></c:tx>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$10</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series).toEqual([
      {
        kind: "bar",
        index: 0,
        name: "Revenue",
        valuesRef: "Sheet1!$B$2:$B$10",
        categoriesRef: "Sheet1!$A$2:$A$10",
        color: "1F77B4",
      },
      {
        kind: "bar",
        index: 1,
        name: "Cost",
        valuesRef: "Sheet1!$C$2:$C$10",
      },
    ]);
  });

  it("decodes scatter xVal / yVal series wiring", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Trend</c:v></c:tx>
          <c:xVal><c:numRef><c:f>S!$A$2:$A$5</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>S!$B$2:$B$5</c:f></c:numRef></c:yVal>
        </c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.series).toEqual([
      {
        kind: "scatter",
        index: 0,
        name: "Trend",
        valuesRef: "S!$B$2:$B$5",
        categoriesRef: "S!$A$2:$A$5",
      },
    ]);
  });

  it("falls back to strRef cache for the series name", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>From Cache</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.series?.[0].name).toBe("From Cache");
  });

  it("uses the strRef formula text when no cache is present", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx>
            <c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef>
          </c:tx>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.series?.[0].name).toBe("Sheet1!$B$1");
  });

  it("omits valuesRef and categoriesRef for literal numLit series", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val>
            <c:numLit>
              <c:formatCode>General</c:formatCode>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>1</c:v></c:pt>
              <c:pt idx="1"><c:v>2</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const series = parseChart(xml)?.series;
    expect(series).toEqual([{ kind: "bar", index: 0 }]);
  });

  it("ignores malformed srgbClr values", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:spPr><a:solidFill><a:srgbClr val="not-a-color"/></a:solidFill></c:spPr>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.series?.[0].color).toBeUndefined();
  });

  it("strips a leading '#' from srgbClr values and uppercases the result", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:spPr><a:solidFill><a:srgbClr val="#aabbcc"/></a:solidFill></c:spPr>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.series?.[0].color).toBe("AABBCC");
  });

  it("indexes series independently per chart-type element in combo charts", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Bar A</c:v></c:tx>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:v>Bar B</c:v></c:tx>
        </c:ser>
      </c:barChart>
      <c:lineChart>
        <c:ser>
          <c:idx val="2"/>
          <c:tx><c:v>Line A</c:v></c:tx>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.series).toEqual([
      { kind: "bar", index: 0, name: "Bar A" },
      { kind: "bar", index: 1, name: "Bar B" },
      { kind: "line", index: 0, name: "Line A" },
    ]);
  });

  it("does not set series when the chart has no <c:ser> children", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["bar"]);
    expect(chart?.seriesCount).toBe(0);
    expect(chart?.series).toBeUndefined();
  });
});

// ── parseChart — legend & grouping ────────────────────────────────

describe("parseChart — legend", () => {
  function chartWithLegend(legendXml: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea><c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea>
    ${legendXml}
  </c:chart>
</c:chartSpace>`;
  }

  it("maps legendPos val=r → right", () => {
    const xml = chartWithLegend('<c:legend><c:legendPos val="r"/></c:legend>');
    expect(parseChart(xml)?.legend).toBe("right");
  });

  it("maps every legendPos value to the writer-side label", () => {
    for (const [val, expected] of [
      ["t", "top"],
      ["b", "bottom"],
      ["l", "left"],
      ["r", "right"],
      ["tr", "topRight"],
    ] as const) {
      const xml = chartWithLegend(`<c:legend><c:legendPos val="${val}"/></c:legend>`);
      expect(parseChart(xml)?.legend).toBe(expected);
    }
  });

  it('returns false when <c:delete val="1"/> hides the legend', () => {
    const xml = chartWithLegend('<c:legend><c:delete val="1"/></c:legend>');
    expect(parseChart(xml)?.legend).toBe(false);
  });

  it("falls back to right when legend is declared without legendPos", () => {
    // Legend element with no legendPos child is valid OOXML; Excel
    // renders it on the right.
    const xml = chartWithLegend("<c:legend/>");
    expect(parseChart(xml)?.legend).toBe("right");
  });

  it("returns undefined when the chart has no <c:legend>", () => {
    const xml = chartWithLegend("");
    expect(parseChart(xml)?.legend).toBeUndefined();
  });

  it("ignores unknown legendPos values rather than fabricating a default", () => {
    const xml = chartWithLegend('<c:legend><c:legendPos val="bogus"/></c:legend>');
    expect(parseChart(xml)?.legend).toBeUndefined();
  });

  it('ignores <c:delete val="0"/> (visible legend with no position) and falls back to right', () => {
    const xml = chartWithLegend('<c:legend><c:delete val="0"/></c:legend>');
    expect(parseChart(xml)?.legend).toBe("right");
  });
});

describe("parseChart — bar grouping", () => {
  function barChartWithGrouping(groupingXml: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        ${groupingXml}
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it("surfaces stacked grouping", () => {
    const xml = barChartWithGrouping('<c:grouping val="stacked"/>');
    expect(parseChart(xml)?.barGrouping).toBe("stacked");
  });

  it("surfaces percentStacked grouping", () => {
    const xml = barChartWithGrouping('<c:grouping val="percentStacked"/>');
    expect(parseChart(xml)?.barGrouping).toBe("percentStacked");
  });

  it("surfaces explicit clustered grouping", () => {
    const xml = barChartWithGrouping('<c:grouping val="clustered"/>');
    expect(parseChart(xml)?.barGrouping).toBe("clustered");
  });

  it("collapses standard grouping to undefined (writer default)", () => {
    // OOXML's `standard` value renders identical to `clustered` in
    // Excel; we omit it so the cloned chart inherits the writer's
    // default rather than carrying a redundant marker.
    const xml = barChartWithGrouping('<c:grouping val="standard"/>');
    expect(parseChart(xml)?.barGrouping).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:grouping> element", () => {
    const xml = barChartWithGrouping("");
    expect(parseChart(xml)?.barGrouping).toBeUndefined();
  });

  it("does not surface barGrouping for non-bar charts", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="stacked"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.barGrouping).toBeUndefined();
  });

  it("uses the first bar chart's grouping in a combo workbook", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:grouping val="stacked"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
      <c:lineChart>
        <c:ser><c:idx val="1"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.barGrouping).toBe("stacked");
  });
});

describe("parseChart — line grouping", () => {
  function lineChartWithGrouping(groupingXml: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        ${groupingXml}
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it("surfaces stacked grouping", () => {
    const xml = lineChartWithGrouping('<c:grouping val="stacked"/>');
    expect(parseChart(xml)?.lineGrouping).toBe("stacked");
  });

  it("surfaces percentStacked grouping", () => {
    const xml = lineChartWithGrouping('<c:grouping val="percentStacked"/>');
    expect(parseChart(xml)?.lineGrouping).toBe("percentStacked");
  });

  it("collapses standard grouping to undefined (writer default)", () => {
    const xml = lineChartWithGrouping('<c:grouping val="standard"/>');
    expect(parseChart(xml)?.lineGrouping).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:grouping> element", () => {
    const xml = lineChartWithGrouping("");
    expect(parseChart(xml)?.lineGrouping).toBeUndefined();
  });

  it("does not surface lineGrouping for non-line charts", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:grouping val="stacked"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.lineGrouping).toBeUndefined();
  });
});

describe("parseChart — area grouping", () => {
  function areaChartWithGrouping(groupingXml: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        ${groupingXml}
        <c:ser><c:idx val="0"/></c:ser>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it("surfaces stacked grouping", () => {
    const xml = areaChartWithGrouping('<c:grouping val="stacked"/>');
    expect(parseChart(xml)?.areaGrouping).toBe("stacked");
  });

  it("surfaces percentStacked grouping", () => {
    const xml = areaChartWithGrouping('<c:grouping val="percentStacked"/>');
    expect(parseChart(xml)?.areaGrouping).toBe("percentStacked");
  });

  it("collapses standard grouping to undefined (writer default)", () => {
    const xml = areaChartWithGrouping('<c:grouping val="standard"/>');
    expect(parseChart(xml)?.areaGrouping).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:grouping> element", () => {
    const xml = areaChartWithGrouping("");
    expect(parseChart(xml)?.areaGrouping).toBeUndefined();
  });

  it("does not surface areaGrouping for non-area charts", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="percentStacked"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.areaGrouping).toBeUndefined();
  });

  it("surfaces both line and area grouping in a combo workbook", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="percentStacked"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
      <c:areaChart>
        <c:grouping val="stacked"/>
        <c:ser><c:idx val="1"/></c:ser>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const parsed = parseChart(xml);
    expect(parsed?.lineGrouping).toBe("percentStacked");
    expect(parsed?.areaGrouping).toBe("stacked");
  });
});

// ── parseChart — data labels ──────────────────────────────────────

describe("parseChart — data labels", () => {
  it("surfaces chart-level dataLabels with showVal and position", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:dLblPos val="outEnd"/>
          <c:showLegendKey val="0"/>
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toEqual({
      showValue: true,
      position: "outEnd",
    });
  });

  it("collects all show* toggles when set", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:dLblPos val="bestFit"/>
          <c:showLegendKey val="0"/>
          <c:showVal val="1"/>
          <c:showCatName val="1"/>
          <c:showSerName val="1"/>
          <c:showPercent val="1"/>
          <c:showBubbleSize val="0"/>
          <c:separator>; </c:separator>
        </c:dLbls>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toEqual({
      showValue: true,
      showCategoryName: true,
      showSeriesName: true,
      showPercent: true,
      position: "bestFit",
      separator: "; ",
    });
  });

  it("returns undefined dataLabels when <c:dLbls> only has delete=1", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:delete val="1"/>
        </c:dLbls>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toBeUndefined();
  });

  it("returns undefined dataLabels when no toggle is on (all show*=0, no position)", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toBeUndefined();
  });

  it("ignores invalid dLblPos values", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:dLblPos val="moonbase"/>
          <c:showVal val="1"/>
        </c:dLbls>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.position).toBeUndefined();
    expect(chart?.dataLabels?.showValue).toBe(true);
  });

  it("accepts true/false as well as 1/0 in show* attributes", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:showVal val="true"/>
          <c:showCatName val="false"/>
        </c:dLbls>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.showValue).toBe(true);
    expect(chart?.dataLabels?.showCategoryName).toBeUndefined();
  });

  it("surfaces series-level dataLabels independently of chart-level", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Revenue</c:v></c:tx>
          <c:dLbls>
            <c:dLblPos val="ctr"/>
            <c:showVal val="1"/>
          </c:dLbls>
          <c:val><c:numRef><c:f>S!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:v>Cost</c:v></c:tx>
          <c:val><c:numRef><c:f>S!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].dataLabels).toEqual({
      showValue: true,
      position: "ctr",
    });
    expect(chart?.series?.[1].dataLabels).toBeUndefined();
  });

  it("captures only the first chart-type-level <c:dLbls> in combo charts", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:dLblPos val="outEnd"/>
          <c:showVal val="1"/>
        </c:dLbls>
      </c:barChart>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dLbls>
          <c:dLblPos val="t"/>
          <c:showCatName val="1"/>
        </c:dLbls>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    // First chart-type element wins for the chart-level summary.
    expect(chart?.dataLabels).toEqual({
      showValue: true,
      position: "outEnd",
    });
  });
});

// ── parseChart — axis titles ──────────────────────────────────────

describe("parseChart — axis titles", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  it("surfaces x and y axis titles from <c:catAx>/<c:valAx> rich text", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue (USD)</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toEqual({
      x: { title: "Quarter" },
      y: { title: "Revenue (USD)" },
    });
  });

  it("does not surface axes when neither axis carries a title", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("surfaces only the populated axis when one side is titled", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toEqual({ y: { title: "Revenue" } });
    expect(chart?.axes?.x).toBeUndefined();
  });

  it("falls back to a strRef cache when the title is a formula", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:valAx>
        <c:axId val="2"/>
        <c:title>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$A$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Cached Y Label</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
        </c:title>
      </c:valAx>
      <c:catAx><c:axId val="1"/></c:catAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.title).toBe("Cached Y Label");
  });

  it("maps scatter axes to x = first valAx, y = second valAx", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
      <c:valAx>
        <c:axId val="1"/>
        <c:axPos val="b"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Time</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:axPos val="l"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Magnitude</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toEqual({
      x: { title: "Time" },
      y: { title: "Magnitude" },
    });
  });

  it("ignores empty/whitespace-only axis titles", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>   </a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("joins multi-run rich titles into a single string", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:title>
          <c:tx>
            <c:rich>
              <a:p>
                <a:r><a:t>Region </a:t></a:r>
                <a:r><a:t>(2024)</a:t></a:r>
              </a:p>
            </c:rich>
          </c:tx>
        </c:title>
      </c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.title).toBe("Region (2024)");
  });
});

// ── parseChart — axis gridlines ──────────────────────────────────

describe("parseChart — axis gridlines", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  it("surfaces major gridlines on the value axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:majorGridlines/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toEqual({
      y: { gridlines: { major: true } },
    });
  });

  it("surfaces both major and minor gridlines when present", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:majorGridlines/>
        <c:minorGridlines/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.gridlines).toEqual({ major: true, minor: true });
  });

  it("surfaces gridlines on both x and y axes simultaneously", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:majorGridlines/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:minorGridlines/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toEqual({
      x: { gridlines: { major: true } },
      y: { gridlines: { minor: true } },
    });
  });

  it("does not surface axes when neither title nor gridlines are declared", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces gridlines and the axis title when both are present", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:majorGridlines/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y).toEqual({
      title: "Revenue",
      gridlines: { major: true },
    });
  });

  it("ignores nested styling inside the gridline elements", () => {
    // Excel sometimes nests <c:spPr> inside <c:majorGridlines> for line
    // styling. The presence of the outer element is what flips the toggle.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:majorGridlines>
          <c:spPr>
            <a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>
          </c:spPr>
        </c:majorGridlines>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.gridlines).toEqual({ major: true });
  });

  it("maps scatter chart gridlines to x = first valAx, y = second valAx", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
      <c:valAx>
        <c:axId val="1"/>
        <c:axPos val="b"/>
        <c:majorGridlines/>
      </c:valAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:axPos val="l"/>
        <c:minorGridlines/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toEqual({
      x: { gridlines: { major: true } },
      y: { gridlines: { minor: true } },
    });
  });
});

// ── parseChart — doughnut hole size ───────────────────────────────

describe("parseChart — doughnut hole size", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:holeSize val="..."/> off a doughnut chart', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:doughnutChart>
      <c:varyColors val="1"/>
      <c:firstSliceAng val="0"/>
      <c:holeSize val="65"/>
    </c:doughnutChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["doughnut"]);
    expect(chart?.holeSize).toBe(65);
  });

  it("omits holeSize when the doughnut chart does not declare one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:doughnutChart><c:varyColors val="1"/></c:doughnutChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.holeSize).toBeUndefined();
  });

  it("rejects malformed or out-of-range holeSize values", () => {
    const out = (val: string): unknown =>
      parseChart(`<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:doughnutChart><c:holeSize val="${val}"/></c:doughnutChart>
  </c:plotArea></c:chart>
</c:chartSpace>`)?.holeSize;
    expect(out("not-a-number")).toBeUndefined();
    expect(out("0")).toBeUndefined();
    expect(out("100")).toBeUndefined();
    // 1–99 inclusive is what the OOXML schema allows.
    expect(out("1")).toBe(1);
    expect(out("99")).toBe(99);
  });

  it("does not attach holeSize to non-doughnut charts", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart><c:varyColors val="1"/></c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["pie"]);
    expect(chart?.holeSize).toBeUndefined();
  });
});

describe("parseChart — first slice angle", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:firstSliceAng val="..."/> off a pie chart', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:varyColors val="1"/>
      <c:firstSliceAng val="90"/>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["pie"]);
    expect(chart?.firstSliceAng).toBe(90);
  });

  it('surfaces <c:firstSliceAng val="..."/> off a doughnut chart', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:doughnutChart>
      <c:varyColors val="1"/>
      <c:firstSliceAng val="180"/>
      <c:holeSize val="50"/>
    </c:doughnutChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["doughnut"]);
    expect(chart?.firstSliceAng).toBe(180);
  });

  it("collapses the OOXML default 0 to undefined (writer absence)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart><c:firstSliceAng val="0"/></c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.firstSliceAng).toBeUndefined();
  });

  it("collapses the schema-equivalent 360 to undefined (same as 0)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart><c:firstSliceAng val="360"/></c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.firstSliceAng).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:firstSliceAng> element", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart><c:varyColors val="1"/></c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.firstSliceAng).toBeUndefined();
  });

  it("rejects malformed or out-of-range firstSliceAng values", () => {
    const out = (val: string): unknown =>
      parseChart(`<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart><c:firstSliceAng val="${val}"/></c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`)?.firstSliceAng;
    expect(out("not-a-number")).toBeUndefined();
    // Negative values fall outside the CT_FirstSliceAng band.
    expect(out("-1")).toBeUndefined();
    // 361 also falls outside the schema band (0..360 inclusive).
    expect(out("361")).toBeUndefined();
    // 1..359 are accepted verbatim.
    expect(out("1")).toBe(1);
    expect(out("270")).toBe(270);
    expect(out("359")).toBe(359);
  });

  it("does not attach firstSliceAng to non-pie / non-doughnut charts", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:firstSliceAng val="90"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["bar"]);
    expect(chart?.firstSliceAng).toBeUndefined();
  });

  it("ignores firstSliceAng outside of pie/doughnut even in combo charts", () => {
    // A pie sibling in the same plotArea should win over a stray
    // firstSliceAng that happens to sit on a non-pie chart-type
    // element earlier in the document order.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser><c:idx val="0"/></c:ser>
    </c:lineChart>
    <c:pieChart>
      <c:varyColors val="1"/>
      <c:firstSliceAng val="45"/>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["line", "pie"]);
    expect(chart?.firstSliceAng).toBe(45);
  });
});

// ── End-to-end: full XLSX with a chart ────────────────────────────

/**
 * Build a minimal XLSX where Sheet1 references a drawing that anchors
 * one bar chart. Mirrors the part shape Excel writes for a single
 * inserted chart (drawing -> _rels -> chart -> style -> colors).
 */
async function buildXlsxWithChart(): Promise<Uint8Array> {
  const z = new ZipWriter();

  z.add(
    "[Content_Types].xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>
  <Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
</Types>`),
  );

  z.add(
    "_rels/.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/workbook.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
  );

  z.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/worksheets/sheet1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c><c r="B1" t="n"><v>10</v></c></row>
    <row r="2"><c r="A2" t="n"><v>2</v></c><c r="B2" t="n"><v>20</v></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`),
  );

  z.add(
    "xl/worksheets/_rels/sheet1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`),
  );

  // Drawing with one chart anchor
  z.add(
    "xl/drawings/drawing1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`),
  );

  z.add(
    "xl/drawings/_rels/drawing1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/charts/chart1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Quarterly Sales</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
  <c:externalData r:id="rId1"/>
</c:chartSpace>`),
  );

  z.add(
    "xl/charts/_rels/chart1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/charts/style1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>`),
  );

  z.add(
    "xl/charts/colors1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle" id="10"/>`),
  );

  return await z.build();
}

// ── readXlsx — chart integration ─────────────────────────────────

describe("readXlsx — chart integration", () => {
  it("attaches sheet.charts when the drawing references a chart", async () => {
    const buf = await buildXlsxWithChart();
    const wb = await readXlsx(buf);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].charts).toHaveLength(1);
    expect(wb.sheets[0].charts?.[0]).toEqual({
      kinds: ["bar"],
      seriesCount: 1,
      title: "Quarterly Sales",
      series: [{ kind: "bar", index: 0, valuesRef: "Data!$B$1:$B$2" }],
      anchor: {
        from: { row: 1, col: 3 },
        to: { row: 16, col: 10 },
      },
    });
  });

  it("does not set sheet.charts when the workbook has none", async () => {
    const z = new ZipWriter();
    z.add(
      "[Content_Types].xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    );
    z.add(
      "_rels/.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/workbook.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Main" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    );
    z.add(
      "xl/_rels/workbook.xml.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/worksheets/sheet1.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
    );

    const wb = await readXlsx(await z.build());
    expect(wb.sheets[0].charts).toBeUndefined();
  });
});

// ── Roundtrip preservation ───────────────────────────────────────

describe("roundtrip — chart preservation", () => {
  it("preserves chart, style, colors, drawing, and drawing rels", async () => {
    const buf = await buildXlsxWithChart();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);
    const zip = new ZipReader(out);

    expect(zip.has("xl/charts/chart1.xml")).toBe(true);
    expect(zip.has("xl/charts/style1.xml")).toBe(true);
    expect(zip.has("xl/charts/colors1.xml")).toBe(true);
    expect(zip.has("xl/charts/_rels/chart1.xml.rels")).toBe(true);
    expect(zip.has("xl/drawings/drawing1.xml")).toBe(true);
    expect(zip.has("xl/drawings/_rels/drawing1.xml.rels")).toBe(true);

    // Chart body must survive byte-identical (it carries the title).
    const chartXml = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(chartXml).toContain("Quarterly Sales");

    // Drawing body keeps the chart graphicFrame.
    const drawingXml = decoder.decode(await zip.extract("xl/drawings/drawing1.xml"));
    expect(drawingXml).toContain("c:chart");
  });

  it("declares chart parts in [Content_Types].xml", async () => {
    const buf = await buildXlsxWithChart();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);
    const zip = new ZipReader(out);

    const ct = decoder.decode(await zip.extract("[Content_Types].xml"));
    expect(ct).toContain("/xl/charts/chart1.xml");
    expect(ct).toContain("/xl/charts/style1.xml");
    expect(ct).toContain("/xl/charts/colors1.xml");
    expect(ct).toContain("application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    expect(ct).toContain("application/vnd.ms-office.chartstyle+xml");
    expect(ct).toContain("application/vnd.ms-office.chartcolorstyle+xml");
    // The chart-bearing drawing must be declared too — without this
    // the drawing bytes survive but Excel treats them as orphan.
    expect(ct).toContain("/xl/drawings/drawing1.xml");
  });

  it("re-anchors the drawing into the regenerated worksheet body", async () => {
    const buf = await buildXlsxWithChart();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);
    const zip = new ZipReader(out);

    const wsXml = decoder.decode(await zip.extract("xl/worksheets/sheet1.xml"));
    expect(wsXml).toMatch(/<drawing r:id="rId\d+"\/>/);

    const sheetRels = decoder.decode(await zip.extract("xl/worksheets/_rels/sheet1.xml.rels"));
    expect(sheetRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"',
    );
    expect(sheetRels).toContain('Target="../drawings/drawing1.xml"');
  });

  it("re-reading the saved workbook still surfaces the chart", async () => {
    const buf = await buildXlsxWithChart();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);
    const reread = await readXlsx(out);

    expect(reread.sheets[0].charts).toHaveLength(1);
    expect(reread.sheets[0].charts?.[0].title).toBe("Quarterly Sales");
    expect(reread.sheets[0].charts?.[0].kinds).toEqual(["bar"]);
  });

  it("survives a cell modification — does not lose the chart", async () => {
    const buf = await buildXlsxWithChart();
    const wb = await openXlsx(buf);
    // Touch a cell so the worksheet is definitively regenerated.
    wb.sheets[0].rows[0][0] = 99;
    const out = await saveXlsx(wb);
    const zip = new ZipReader(out);

    expect(zip.has("xl/charts/chart1.xml")).toBe(true);
    expect(zip.has("xl/drawings/drawing1.xml")).toBe(true);

    const reread = await readXlsx(out);
    expect(reread.sheets[0].rows[0][0]).toBe(99);
    expect(reread.sheets[0].charts).toHaveLength(1);
  });
});

// ── readXlsx — chart cell anchor ─────────────────────────────────

/**
 * Build a minimal XLSX where Sheet1's drawing anchors a single chart
 * with a custom anchor flavor (`twoCellAnchor`, `oneCellAnchor`, or
 * `absoluteAnchor`). Used to verify {@link Chart.anchor} extraction.
 */
async function buildXlsxWithAnchor(anchorXml: string): Promise<Uint8Array> {
  const z = new ZipWriter();

  z.add(
    "[Content_Types].xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`),
  );

  z.add(
    "_rels/.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/workbook.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
  );

  z.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/worksheets/sheet1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>`),
  );

  z.add(
    "xl/worksheets/_rels/sheet1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/drawings/drawing1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${anchorXml}
</xdr:wsDr>`),
  );

  z.add(
    "xl/drawings/_rels/drawing1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/charts/chart1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart></c:plotArea></c:chart>
</c:chartSpace>`),
  );

  return await z.build();
}

/**
 * Builds a `<xdr:graphicFrame>` payload with a chart reference. Used
 * inside the anchor builders below to keep the test data compact.
 */
const CHART_GRAPHIC_FRAME = `<xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>`;

describe("readXlsx — chart cell anchor", () => {
  it("surfaces from/to from a twoCellAnchor", async () => {
    const buf = await buildXlsxWithAnchor(`<xdr:twoCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    ${CHART_GRAPHIC_FRAME}
    <xdr:clientData/>
  </xdr:twoCellAnchor>`);
    const wb = await readXlsx(buf);
    expect(wb.sheets[0].charts?.[0].anchor).toEqual({
      from: { row: 5, col: 2 },
      to: { row: 20, col: 9 },
    });
  });

  it("surfaces from-only for a oneCellAnchor (intrinsic size lives in <xdr:ext>)", async () => {
    const buf = await buildXlsxWithAnchor(`<xdr:oneCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="6000000" cy="3500000"/>
    ${CHART_GRAPHIC_FRAME}
    <xdr:clientData/>
  </xdr:oneCellAnchor>`);
    const wb = await readXlsx(buf);
    const anchor = wb.sheets[0].charts?.[0].anchor;
    expect(anchor).toEqual({ from: { row: 2, col: 1 } });
    expect(anchor?.to).toBeUndefined();
  });

  it("omits anchor for an absoluteAnchor (EMU-positioned, no cell anchor)", async () => {
    const buf = await buildXlsxWithAnchor(`<xdr:absoluteAnchor>
    <xdr:pos x="914400" y="685800"/>
    <xdr:ext cx="6000000" cy="3500000"/>
    ${CHART_GRAPHIC_FRAME}
    <xdr:clientData/>
  </xdr:absoluteAnchor>`);
    const wb = await readXlsx(buf);
    expect(wb.sheets[0].charts?.[0].anchor).toBeUndefined();
  });

  it("omits anchor when the twoCellAnchor is missing its <xdr:from> block", async () => {
    // Pathological — Excel always writes <xdr:from>, but defensive
    // parsing should not invent a (0,0) anchor.
    const buf = await buildXlsxWithAnchor(`<xdr:twoCellAnchor>
    ${CHART_GRAPHIC_FRAME}
    <xdr:clientData/>
  </xdr:twoCellAnchor>`);
    const wb = await readXlsx(buf);
    expect(wb.sheets[0].charts?.[0].anchor).toBeUndefined();
  });

  it("falls back to from-only when the twoCellAnchor is missing its <xdr:to> block", async () => {
    // Some authoring tools omit <xdr:to> for one-cell-style charts
    // even though the anchor element is twoCellAnchor.
    const buf = await buildXlsxWithAnchor(`<xdr:twoCellAnchor>
    <xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    ${CHART_GRAPHIC_FRAME}
    <xdr:clientData/>
  </xdr:twoCellAnchor>`);
    const wb = await readXlsx(buf);
    expect(wb.sheets[0].charts?.[0].anchor).toEqual({ from: { row: 7, col: 4 } });
  });

  it("attaches the correct anchor to each chart when the drawing carries multiple", async () => {
    // Build a drawing with two anchors, each pointing at its own chart
    // part. Verifies the per-anchor pairing rather than a coarse
    // "any anchor" pickup.
    const z = new ZipWriter();
    z.add(
      "[Content_Types].xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`),
    );
    z.add(
      "_rels/.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/workbook.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    );
    z.add(
      "xl/_rels/workbook.xml.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/worksheets/sheet1.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>`),
    );
    z.add(
      "xl/worksheets/_rels/sheet1.xml.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/drawings/drawing1.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="A"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>30</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="B"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`),
    );
    z.add(
      "xl/drawings/_rels/drawing1.xml.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/charts/chart1.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>First</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea></c:chart>
</c:chartSpace>`),
    );
    z.add(
      "xl/charts/chart2.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Second</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart></c:plotArea></c:chart>
</c:chartSpace>`),
    );

    const wb = await readXlsx(await z.build());
    const charts = wb.sheets[0].charts;
    expect(charts).toHaveLength(2);
    // Order tracks the drawing's anchor sequence — the first
    // graphicFrame becomes charts[0], the second becomes charts[1].
    const byTitle = new Map(charts!.map((c) => [c.title, c]));
    expect(byTitle.get("First")?.anchor).toEqual({
      from: { row: 0, col: 0 },
      to: { row: 10, col: 5 },
    });
    expect(byTitle.get("Second")?.anchor).toEqual({
      from: { row: 12, col: 6 },
      to: { row: 30, col: 13 },
    });
  });

  it("survives roundtrip — re-reading the saved file still reports the anchor", async () => {
    const buf = await buildXlsxWithChart();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);
    const reread = await readXlsx(out);
    expect(reread.sheets[0].charts?.[0].anchor).toEqual({
      from: { row: 1, col: 3 },
      to: { row: 16, col: 10 },
    });
  });
});
