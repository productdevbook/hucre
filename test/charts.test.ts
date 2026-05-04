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

// ── parseChart — axis scale ───────────────────────────────────────

describe("parseChart — axis scale", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces <c:min> / <c:max> / <c:majorUnit> / <c:minorUnit> off the value axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling>
        <c:orientation val="minMax"/>
        <c:max val="100"/>
        <c:min val="0"/>
      </c:scaling>
      <c:majorUnit val="25"/>
      <c:minorUnit val="5"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.scale).toEqual({ min: 0, max: 100, majorUnit: 25, minorUnit: 5 });
  });

  it("surfaces <c:logBase> from inside <c:scaling>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling>
        <c:logBase val="10"/>
        <c:orientation val="minMax"/>
      </c:scaling>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.scale).toEqual({ logBase: 10 });
  });

  it("does not surface a scale when <c:scaling> only carries <c:orientation>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores non-finite, zero, and negative tick spacings", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:majorUnit val="0"/>
      <c:minorUnit val="-2"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("maps scatter axes to x = first valAx, y = second valAx", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:scaling><c:orientation val="minMax"/><c:max val="50"/><c:min val="0"/></c:scaling>
      <c:axPos val="b"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="minMax"/><c:max val="200"/><c:min val="-200"/></c:scaling>
      <c:axPos val="l"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.scale).toEqual({ min: 0, max: 50 });
    expect(chart?.axes?.y?.scale).toEqual({ min: -200, max: 200 });
  });
});

// ── parseChart — axis number format ───────────────────────────────

describe("parseChart — axis number format", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:numFmt formatCode="..."/> off the value axis', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:numFmt formatCode="#,##0" sourceLinked="0"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.numberFormat).toEqual({ formatCode: "#,##0" });
  });

  it("surfaces sourceLinked when set to 1", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:numFmt formatCode="0.00%" sourceLinked="1"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.numberFormat).toEqual({ formatCode: "0.00%", sourceLinked: true });
  });

  it("ignores empty formatCode attributes", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:numFmt formatCode="" sourceLinked="1"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces axis title, gridlines, scale and number format together", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="minMax"/><c:max val="100"/><c:min val="0"/></c:scaling>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:numFmt formatCode="$#,##0" sourceLinked="0"/>
      <c:majorUnit val="25"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y).toEqual({
      title: "Revenue",
      gridlines: { major: true },
      scale: { min: 0, max: 100, majorUnit: 25 },
      numberFormat: { formatCode: "$#,##0" },
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

describe("parseChart — bar gapWidth & overlap", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:gapWidth val="..."/> off a bar chart', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:gapWidth val="75"/>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["bar"]);
    expect(chart?.gapWidth).toBe(75);
  });

  it('surfaces <c:overlap val="..."/> off a bar chart', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:overlap val="-25"/>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["bar"]);
    expect(chart?.overlap).toBe(-25);
  });

  it("surfaces both gapWidth and overlap when both are declared", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:gapWidth val="75"/>
      <c:overlap val="100"/>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.gapWidth).toBe(75);
    expect(chart?.overlap).toBe(100);
  });

  it("collapses the OOXML default gapWidth (150) to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:gapWidth val="150"/></c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.gapWidth).toBeUndefined();
  });

  it("collapses the OOXML default overlap (0) to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:overlap val="0"/></c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.overlap).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:gapWidth> / <c:overlap>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.gapWidth).toBeUndefined();
    expect(chart?.overlap).toBeUndefined();
  });

  it("rejects malformed or out-of-range gapWidth values", () => {
    const out = (val: string): unknown =>
      parseChart(`<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:gapWidth val="${val}"/></c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`)?.gapWidth;
    expect(out("not-a-number")).toBeUndefined();
    // Below schema minimum.
    expect(out("-1")).toBeUndefined();
    // Above schema maximum (ST_GapAmount is 0..500 inclusive).
    expect(out("501")).toBeUndefined();
    // Bounds inclusive.
    expect(out("0")).toBe(0);
    expect(out("500")).toBe(500);
  });

  it("rejects malformed or out-of-range overlap values", () => {
    const out = (val: string): unknown =>
      parseChart(`<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:overlap val="${val}"/></c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`)?.overlap;
    expect(out("not-a-number")).toBeUndefined();
    expect(out("-101")).toBeUndefined();
    expect(out("101")).toBeUndefined();
    // Bounds inclusive (-100..100), 0 collapses to undefined.
    expect(out("-100")).toBe(-100);
    expect(out("100")).toBe(100);
    expect(out("-1")).toBe(-1);
    expect(out("1")).toBe(1);
  });

  it("does not attach gapWidth / overlap to non-bar chart kinds", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:gapWidth val="75"/>
      <c:overlap val="50"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["line"]);
    expect(chart?.gapWidth).toBeUndefined();
    expect(chart?.overlap).toBeUndefined();
  });

  it("surfaces gapWidth / overlap from <c:bar3DChart> as well", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:bar3DChart>
      <c:gapWidth val="50"/>
      <c:overlap val="25"/>
    </c:bar3DChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.kinds).toEqual(["bar3D"]);
    expect(chart?.gapWidth).toBe(50);
    expect(chart?.overlap).toBe(25);
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

// ── parseChart — series smooth flag ───────────────────────────────

describe("parseChart — series smooth flag", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces smooth=true on a <c:lineChart> series with <c:smooth val="1"/>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        <c:smooth val="1"/>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].smooth).toBe(true);
  });

  it("surfaces smooth=true on a <c:scatterChart> series", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser>
        <c:idx val="0"/>
        <c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$5</c:f></c:numRef></c:xVal>
        <c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:yVal>
        <c:smooth val="1"/>
      </c:ser>
    </c:scatterChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].smooth).toBe(true);
  });

  it("collapses the OOXML default smooth=false to undefined", () => {
    // Absence of <c:smooth> and `<c:smooth val="0"/>` round-trip
    // identically through the writer's elision logic, so the parser
    // collapses both to undefined to keep the read-side shape minimal.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        <c:smooth val="0"/>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].smooth).toBeUndefined();
  });

  it("returns smooth undefined when <c:smooth> is absent", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].smooth).toBeUndefined();
  });

  it('also accepts the "true" / "false" boolean spelling', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        <c:smooth val="true"/>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].smooth).toBe(true);
  });

  it("ignores <c:smooth> on chart families whose schema rejects the element", () => {
    // The OOXML schema places <c:smooth> only on CT_LineSer and
    // CT_ScatterSer. A bar/pie/area template carrying a stray smooth
    // element should not surface a flag that the writer would never
    // emit anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        <c:smooth val="1"/>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].smooth).toBeUndefined();
  });

  it("surfaces smooth per-series independently across multi-series line charts", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
        <c:smooth val="1"/>
      </c:ser>
      <c:ser>
        <c:idx val="1"/>
        <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
        <c:smooth val="0"/>
      </c:ser>
      <c:ser>
        <c:idx val="2"/>
        <c:val><c:numRef><c:f>Sheet1!$D$2:$D$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series).toHaveLength(3);
    expect(chart?.series?.[0].smooth).toBe(true);
    expect(chart?.series?.[1].smooth).toBeUndefined();
    expect(chart?.series?.[2].smooth).toBeUndefined();
  });
});

// ── parseChart — series line stroke ───────────────────────────────

describe("parseChart — series line stroke", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  it('surfaces stroke.dash from <a:prstDash val="dash"/>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:ln>
            <a:prstDash val="dash"/>
          </a:ln>
        </c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toEqual({ dash: "dash" });
  });

  it('surfaces stroke.width from <a:ln w="..."/> by converting EMU back to points', () => {
    // 31 750 EMU = 2.5 pt.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:ln w="31750"/>
        </c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toEqual({ width: 2.5 });
  });

  it("surfaces both dash and width when both are present", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:ln w="9525">
            <a:prstDash val="lgDash"/>
          </a:ln>
        </c:spPr>
        <c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$5</c:f></c:numRef></c:xVal>
        <c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:yVal>
      </c:ser>
    </c:scatterChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toEqual({ dash: "lgDash", width: 0.75 });
  });

  it("returns stroke undefined when <a:ln> is absent", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toBeUndefined();
  });

  it("collapses an empty <a:ln/> (no width, no prstDash) to undefined", () => {
    // An empty <a:ln/> carries no meaningful settings; don't surface a
    // record the writer will never re-emit.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr><a:ln/></c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toBeUndefined();
  });

  it("drops an unknown dash value rather than surfacing a malformed token", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:ln>
            <a:prstDash val="wiggle"/>
          </a:ln>
        </c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toBeUndefined();
  });

  it("clamps an absurdly wide <a:ln w=...> back into the 0.25..13.5 pt band", () => {
    // 999 999 EMU ≈ 78.7 pt; clamp to 13.5 pt.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:ln w="999999"/>
        </c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toEqual({ width: 13.5 });
  });

  it("ignores stroke on chart families whose schema does not paint a connecting line", () => {
    // Even if a corrupt template carries <a:ln> on a bar/pie/area
    // series, the read side should not surface the field — it would
    // mislead a clone consumer about what the chart actually renders.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:ln w="31750">
            <a:prstDash val="dash"/>
          </a:ln>
        </c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].stroke).toBeUndefined();
  });

  it("surfaces stroke per-series independently across multi-series line charts", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr><a:ln w="31750"><a:prstDash val="dash"/></a:ln></c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="1"/>
        <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="2"/>
        <c:spPr><a:ln><a:prstDash val="sysDot"/></a:ln></c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$D$2:$D$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series).toHaveLength(3);
    expect(chart?.series?.[0].stroke).toEqual({ dash: "dash", width: 2.5 });
    expect(chart?.series?.[1].stroke).toBeUndefined();
    expect(chart?.series?.[2].stroke).toEqual({ dash: "sysDot" });
  });

  it("does not let stroke shadow the existing series.color (parseSeriesColor still wins)", () => {
    // A series with both a fill color and a stroke should surface
    // `color` and `stroke` independently — the stroke object never
    // duplicates the color (parseSeriesColor already covers it).
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:solidFill><a:srgbClr val="1F77B4"/></a:solidFill>
          <a:ln w="19050">
            <a:solidFill><a:srgbClr val="1F77B4"/></a:solidFill>
            <a:prstDash val="dashDot"/>
          </a:ln>
        </c:spPr>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].color).toBe("1F77B4");
    // 19 050 EMU = 1.5 pt.
    expect(chart?.series?.[0].stroke).toEqual({ dash: "dashDot", width: 1.5 });
  });
});

// ── parseChart — series marker ────────────────────────────────────

describe("parseChart — series marker", () => {
  const NS_C = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;
  const NS_A = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  it("surfaces symbol + size on a <c:lineChart> series", () => {
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:marker>
          <c:symbol val="diamond"/>
          <c:size val="10"/>
        </c:marker>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].marker).toEqual({ symbol: "diamond", size: 10 });
  });

  it("surfaces fill and outline colors from <c:spPr>", () => {
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:marker>
          <c:symbol val="circle"/>
          <c:size val="6"/>
          <c:spPr>
            <a:solidFill><a:srgbClr val="1F77B4"/></a:solidFill>
            <a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
          </c:spPr>
        </c:marker>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].marker).toEqual({
      symbol: "circle",
      size: 6,
      fill: "1F77B4",
      line: "FF0000",
    });
  });

  it("upper-cases hex color values pulled from the marker spPr", () => {
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:marker>
          <c:spPr><a:solidFill><a:srgbClr val="1f77b4"/></a:solidFill></c:spPr>
        </c:marker>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].marker?.fill).toBe("1F77B4");
  });

  it("clamps marker size into the OOXML 2..72 band", () => {
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:marker><c:size val="999"/></c:marker>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="1"/>
        <c:marker><c:size val="0"/></c:marker>
        <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].marker?.size).toBe(72);
    expect(chart?.series?.[1].marker?.size).toBe(2);
  });

  it("collapses an empty <c:marker/> to undefined", () => {
    // No symbol, size, or color — there's nothing meaningful to surface.
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:marker/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].marker).toBeUndefined();
  });

  it("drops unknown marker symbols rather than surface invalid values", () => {
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:marker><c:symbol val="pentagon"/><c:size val="5"/></c:marker>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    // Size still surfaces; the bogus symbol is dropped.
    expect(chart?.series?.[0].marker).toEqual({ size: 5 });
  });

  it("surfaces marker on a <c:scatterChart> series", () => {
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser>
        <c:idx val="0"/>
        <c:marker><c:symbol val="x"/><c:size val="8"/></c:marker>
        <c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$5</c:f></c:numRef></c:xVal>
        <c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:yVal>
      </c:ser>
    </c:scatterChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].marker).toEqual({ symbol: "x", size: 8 });
  });

  it("ignores <c:marker> on chart families whose schema rejects it", () => {
    // A bar / pie / area template carrying a stray <c:marker> on its
    // series should not surface a marker that the writer would never
    // emit on those families anyway.
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:marker><c:symbol val="circle"/></c:marker>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].marker).toBeUndefined();
  });

  it("surfaces marker per-series independently across multi-series line charts", () => {
    const xml = `<c:chartSpace ${NS_C} ${NS_A}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="1"/>
        <c:marker><c:symbol val="square"/></c:marker>
        <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="2"/>
        <c:val><c:numRef><c:f>Sheet1!$D$2:$D$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series).toHaveLength(3);
    expect(chart?.series?.[0].marker).toEqual({ symbol: "circle", size: 6 });
    expect(chart?.series?.[1].marker).toEqual({ symbol: "square" });
    expect(chart?.series?.[2].marker).toBeUndefined();
  });
});

// ── parseChart — dispBlanksAs ─────────────────────────────────────

describe("parseChart — dispBlanksAs", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:dispBlanksAs val="zero"/> off <c:chart>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dispBlanksAs).toBe("zero");
  });

  it('surfaces <c:dispBlanksAs val="span"/> off <c:chart>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:dispBlanksAs val="span"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dispBlanksAs).toBe("span");
  });

  it("collapses the OOXML default 'gap' to undefined (writer absence)", () => {
    // The default carried explicitly by Excel's reference serialization
    // round-trips identically to absence of the field.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dispBlanksAs).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:dispBlanksAs> element", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dispBlanksAs).toBeUndefined();
  });

  it("drops unknown dispBlanksAs values rather than fabricate one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:dispBlanksAs val="bogus"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dispBlanksAs).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:dispBlanksAs>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:dispBlanksAs/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dispBlanksAs).toBeUndefined();
  });
});

// ── parseChart — varyColors ───────────────────────────────────────

describe("parseChart — varyColors", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:varyColors val="1"/> on a column chart (non-default true)', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBe(true);
  });

  it('surfaces <c:varyColors val="0"/> on a doughnut chart (non-default false)', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:doughnutChart>
        <c:varyColors val="0"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="50"/>
      </c:doughnutChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBe(false);
  });

  it("collapses the per-family default to undefined on a column chart (varyColors=0)", () => {
    // Column / bar default is `false` — `<c:varyColors val="0"/>` and
    // absence both round-trip identically.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBeUndefined();
  });

  it("collapses the per-family default to undefined on a pie chart (varyColors=1)", () => {
    // Pie default is `true` — `<c:varyColors val="1"/>` and absence both
    // round-trip identically.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:varyColors> element", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBeUndefined();
  });

  it("accepts the OOXML true / false spellings on the val attribute", () => {
    // The OOXML schema for `xsd:boolean` accepts `"true"` / `"false"`
    // alongside the more common `"1"` / `"0"`. Hucre tolerates both
    // shapes — a hand-edited template using `true` should round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="true"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBe(true);
  });

  it("drops unknown varyColors values rather than fabricate one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:varyColors val="bogus"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:varyColors>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:varyColors/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBeUndefined();
  });

  it("surfaces varyColors from the first chart-type element on combo charts", () => {
    // The reader latches onto the first chart-type element that carries
    // a `<c:varyColors>` value, mirroring how it surfaces grouping /
    // gapWidth on the first matching child.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
      <c:lineChart>
        <c:varyColors val="0"/>
        <c:ser><c:idx val="1"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.varyColors).toBe(true);
  });
});

// ── parseChart — scatterStyle ─────────────────────────────────────

describe("parseChart — scatterStyle", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:scatterStyle val="lineMarker"/> on a scatter chart', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:varyColors val="0"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.scatterStyle).toBe("lineMarker");
  });

  it('surfaces <c:scatterStyle val="smooth"/> on a smooth-line scatter', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="smooth"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.scatterStyle).toBe("smooth");
  });

  it("surfaces every other ST_ScatterStyle preset literally", () => {
    // Walk the remaining four enum tokens — each one round-trips
    // verbatim with no per-family default collapse.
    for (const preset of ["none", "line", "marker", "smoothMarker"] as const) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="${preset}"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.scatterStyle).toBe(preset);
    }
  });

  it("returns undefined when the scatter chart omits <c:scatterStyle>", () => {
    // The OOXML schema lists the element as required, but Excel falls
    // back to the schema default `"marker"` when the file omits it. The
    // reader does not fabricate a value — absence stays absence so the
    // clone layer can decide whether to inherit the writer's own default.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.scatterStyle).toBeUndefined();
  });

  it("ignores a <c:scatterStyle> with no val attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.scatterStyle).toBeUndefined();
  });

  it("drops unknown scatterStyle values rather than fabricate one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="bogus"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.scatterStyle).toBeUndefined();
  });

  it("does not surface scatterStyle on non-scatter charts", () => {
    // The OOXML schema places <c:scatterStyle> exclusively on
    // <c:scatterChart>; even if a hand-edited bar chart somehow carries
    // the element, the reader does not surface it because the parse is
    // gated on the matching kind.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:scatterStyle val="lineMarker"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.scatterStyle).toBeUndefined();
  });

  it("surfaces scatterStyle from the first scatterChart in a combo chart", () => {
    // Combo charts are rare but Excel supports an arbitrary mix of
    // chart-type elements inside one plot area. The reader latches onto
    // the first <c:scatterChart>'s scatterStyle, mirroring how it
    // handles other chart-type-level fields.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="smoothMarker"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:scatterChart>
      <c:scatterChart>
        <c:scatterStyle val="line"/>
        <c:ser><c:idx val="1"/></c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.scatterStyle).toBe("smoothMarker");
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

// ── parseChart — series invertIfNegative flag ─────────────────────

describe("parseChart — series invertIfNegative flag", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces invertIfNegative=true on a <c:barChart> series with <c:invertIfNegative val="1"/>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:invertIfNegative val="1"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].invertIfNegative).toBe(true);
  });

  it("surfaces invertIfNegative=true on a horizontal bar chart series", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="bar"/>
      <c:ser>
        <c:idx val="0"/>
        <c:invertIfNegative val="1"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].invertIfNegative).toBe(true);
  });

  it("collapses the OOXML default invertIfNegative=false to undefined", () => {
    // Absence of <c:invertIfNegative> and `<c:invertIfNegative val="0"/>`
    // round-trip identically through the writer's elision logic, so the
    // parser collapses both to undefined to keep the read-side shape
    // minimal.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:invertIfNegative val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].invertIfNegative).toBeUndefined();
  });

  it("returns invertIfNegative undefined when <c:invertIfNegative> is absent", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].invertIfNegative).toBeUndefined();
  });

  it('also accepts the "true" / "false" boolean spelling', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:invertIfNegative val="true"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].invertIfNegative).toBe(true);
  });

  it("ignores <c:invertIfNegative> on chart families whose schema rejects the element", () => {
    // The OOXML schema places <c:invertIfNegative> only on CT_BarSer
    // and CT_Bar3DSer. A line/pie/area/scatter template carrying a
    // stray invert element should not surface a flag that the writer
    // would never emit anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:ser>
        <c:idx val="0"/>
        <c:invertIfNegative val="1"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].invertIfNegative).toBeUndefined();
  });

  it("surfaces invertIfNegative per-series independently across multi-series bar charts", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:invertIfNegative val="1"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="1"/>
        <c:invertIfNegative val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="2"/>
        <c:val><c:numRef><c:f>Sheet1!$D$2:$D$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series).toHaveLength(3);
    expect(chart?.series?.[0].invertIfNegative).toBe(true);
    expect(chart?.series?.[1].invertIfNegative).toBeUndefined();
    expect(chart?.series?.[2].invertIfNegative).toBeUndefined();
  });

  it("returns invertIfNegative undefined when val attribute is missing", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser>
        <c:idx val="0"/>
        <c:invertIfNegative/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].invertIfNegative).toBeUndefined();
  });
});

// ── parseChart — series explosion (pie / doughnut) ────────────────

describe("parseChart — series explosion", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces explosion=25 on a <c:pieChart> series with <c:explosion val="25"/>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="25"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].explosion).toBe(25);
  });

  it('surfaces explosion on a <c:doughnutChart> series with <c:explosion val="50"/>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:doughnutChart>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="50"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:holeSize val="50"/>
    </c:doughnutChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].explosion).toBe(50);
  });

  it("collapses the OOXML default explosion=0 to undefined", () => {
    // Absence of <c:explosion> and `<c:explosion val="0"/>` round-trip
    // identically through the writer's elision logic, so the parser
    // collapses both to undefined to keep the read-side shape minimal.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].explosion).toBeUndefined();
  });

  it("returns explosion undefined when <c:explosion> is absent", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser>
        <c:idx val="0"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].explosion).toBeUndefined();
  });

  it("rounds non-integer explosion values to the nearest integer", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="33.6"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].explosion).toBe(34);
  });

  it("rejects malformed or negative explosion values", () => {
    const cases = ["bogus", "-50", "NaN", "Infinity", ""];
    for (const val of cases) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="${val}"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.series?.[0].explosion).toBeUndefined();
    }
  });

  it("returns explosion undefined when val attribute is missing", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].explosion).toBeUndefined();
  });

  it("ignores <c:explosion> on chart families whose schema rejects the element", () => {
    // The OOXML schema places <c:explosion> only on CT_PieSer (shared
    // across the pie family). A bar/line/area/scatter template carrying
    // a stray explosion element should not surface a value the writer
    // would never emit anyway.
    const cases = ["barChart", "lineChart", "areaChart", "scatterChart"] as const;
    for (const tag of cases) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:${tag}>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="50"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:${tag}>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.series?.[0].explosion).toBeUndefined();
    }
  });

  it("surfaces explosion per-series independently across multi-series doughnut charts", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:doughnutChart>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="25"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="1"/>
        <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:ser>
        <c:idx val="2"/>
        <c:explosion val="75"/>
        <c:val><c:numRef><c:f>Sheet1!$D$2:$D$5</c:f></c:numRef></c:val>
      </c:ser>
      <c:holeSize val="50"/>
    </c:doughnutChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series).toHaveLength(3);
    expect(chart?.series?.[0].explosion).toBe(25);
    expect(chart?.series?.[1].explosion).toBeUndefined();
    expect(chart?.series?.[2].explosion).toBe(75);
  });

  it("surfaces explosion on <c:pie3DChart> and <c:ofPieChart> series", () => {
    // CT_Pie3DSer / CT_OfPieSer share CT_PieSer through EG_PieSer, so
    // the parser should accept <c:explosion> on both flavors.
    for (const tag of ["pie3DChart", "ofPieChart"] as const) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:${tag}>
      <c:ser>
        <c:idx val="0"/>
        <c:explosion val="40"/>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:${tag}>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.series?.[0].explosion).toBe(40);
    }
  });
});

// ── parseChart — axis tick marks and tick label position ──────────

describe("parseChart — axis tick marks and tick label position", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces non-default <c:majorTickMark val=".."/> off the value axis', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:majorTickMark val="cross"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.majorTickMark).toBe("cross");
  });

  it('surfaces non-default <c:minorTickMark val=".."/> off the value axis', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:minorTickMark val="out"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.minorTickMark).toBe("out");
  });

  it('surfaces non-default <c:tickLblPos val=".."/> off the value axis', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:tickLblPos val="low"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.tickLblPos).toBe("low");
  });

  it("collapses the OOXML default majorTickMark=out to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:majorTickMark val="out"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("collapses the OOXML default minorTickMark=none to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:minorTickMark val="none"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("collapses the OOXML default tickLblPos=nextTo to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:tickLblPos val="nextTo"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores unknown majorTickMark / minorTickMark values", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:majorTickMark val="zigzag"/>
      <c:minorTickMark val="diagonal"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores unknown tickLblPos values", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:tickLblPos val="diagonal"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores tick-mark / tick-lbl-pos elements with no val attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:majorTickMark/>
      <c:minorTickMark/>
      <c:tickLblPos/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("surfaces tick rendering on the category axis (catAx) too", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:majorTickMark val="in"/>
      <c:tickLblPos val="high"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.majorTickMark).toBe("in");
    expect(chart?.axes?.x?.tickLblPos).toBe("high");
  });

  it("surfaces tick rendering on the scatter X axis (first valAx)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:majorTickMark val="cross"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:tickLblPos val="none"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.majorTickMark).toBe("cross");
    expect(chart?.axes?.y?.tickLblPos).toBe("none");
  });

  it("co-surfaces title, gridlines, scale, numberFormat, tick marks and tick label pos together", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="minMax"/><c:max val="100"/><c:min val="0"/></c:scaling>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:numFmt formatCode="$#,##0" sourceLinked="0"/>
      <c:majorTickMark val="cross"/>
      <c:minorTickMark val="in"/>
      <c:tickLblPos val="low"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y).toEqual({
      title: "Revenue",
      gridlines: { major: true },
      scale: { min: 0, max: 100 },
      numberFormat: { formatCode: "$#,##0" },
      majorTickMark: "cross",
      minorTickMark: "in",
      tickLblPos: "low",
    });
  });
});

// ── parseChart — plotVisOnly ──────────────────────────────────────

describe("parseChart — plotVisOnly", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:plotVisOnly val="0"/> on <c:chart> as false (non-default)', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.plotVisOnly).toBe(false);
  });

  it("collapses the OOXML default true to undefined (writer absence)", () => {
    // The default carried explicitly by Excel's reference serialization
    // round-trips identically to absence of the field.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.plotVisOnly).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:plotVisOnly> element", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.plotVisOnly).toBeUndefined();
  });

  it("accepts the OOXML true / false spellings on the val attribute", () => {
    // The OOXML schema for `xsd:boolean` accepts `"true"` / `"false"`
    // alongside the more common `"1"` / `"0"`. Hucre tolerates both
    // shapes — a hand-edited template using `false` should round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="false"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.plotVisOnly).toBe(false);
  });

  it("collapses the 'true' spelling to undefined as well", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="true"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.plotVisOnly).toBeUndefined();
  });

  it("drops unknown plotVisOnly values rather than fabricate one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="bogus"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.plotVisOnly).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:plotVisOnly>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:plotVisOnly/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.plotVisOnly).toBeUndefined();
  });

  it("surfaces plotVisOnly alongside other chart-level toggles", () => {
    // Co-existing with dispBlanksAs / varyColors should not interfere
    // — each toggle parses independently off <c:chart>.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });
});

// ── parseChart — showDLblsOverMax ─────────────────────────────────

describe("parseChart — showDLblsOverMax", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:showDLblsOverMax val="0"/> on <c:chart> as false (non-default)', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showDLblsOverMax).toBe(false);
  });

  it("collapses the OOXML default true to undefined (writer absence)", () => {
    // The default carried explicitly by the writer's reference shape
    // round-trips identically to absence of the field.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:showDLblsOverMax val="1"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showDLblsOverMax).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:showDLblsOverMax> element", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showDLblsOverMax).toBeUndefined();
  });

  it("accepts the OOXML true / false spellings on the val attribute", () => {
    // The OOXML schema for `xsd:boolean` accepts `"true"` / `"false"`
    // alongside the more common `"1"` / `"0"`. Hucre tolerates both
    // shapes — a hand-edited template using `false` should round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:showDLblsOverMax val="false"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showDLblsOverMax).toBe(false);
  });

  it("collapses the 'true' spelling to undefined as well", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:showDLblsOverMax val="true"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showDLblsOverMax).toBeUndefined();
  });

  it("drops unknown showDLblsOverMax values rather than fabricate one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:showDLblsOverMax val="bogus"/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showDLblsOverMax).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:showDLblsOverMax>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:showDLblsOverMax/>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showDLblsOverMax).toBeUndefined();
  });

  it("surfaces showDLblsOverMax alongside other chart-level toggles", () => {
    // Co-existing with plotVisOnly / dispBlanksAs / varyColors should
    // not interfere — each toggle parses independently off <c:chart>.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showDLblsOverMax).toBe(false);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });
});

// ── parseChart — roundedCorners ───────────────────────────────────

describe("parseChart — roundedCorners", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:roundedCorners val="1"/> on <c:chartSpace> as true (non-default)', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.roundedCorners).toBe(true);
  });

  it("collapses the OOXML default false to undefined (writer absence)", () => {
    // The default carried explicitly by Excel's reference serialization
    // round-trips identically to absence of the field.
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.roundedCorners).toBeUndefined();
  });

  it("returns undefined when the chartSpace has no <c:roundedCorners> element", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.roundedCorners).toBeUndefined();
  });

  it("accepts the OOXML true / false spellings on the val attribute", () => {
    // The OOXML schema for `xsd:boolean` accepts `"true"` / `"false"`
    // alongside the more common `"1"` / `"0"`. Hucre tolerates both
    // shapes — a hand-edited template using `true` should round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="true"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.roundedCorners).toBe(true);
  });

  it("collapses the 'false' spelling to undefined as well", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="false"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.roundedCorners).toBeUndefined();
  });

  it("drops unknown roundedCorners values rather than fabricate one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="bogus"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.roundedCorners).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:roundedCorners>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.roundedCorners).toBeUndefined();
  });

  it("surfaces roundedCorners alongside other chart-level toggles", () => {
    // Co-existing with plotVisOnly / dispBlanksAs / varyColors should
    // not interfere — roundedCorners parses off <c:chartSpace> while
    // the others sit on <c:chart>.
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.roundedCorners).toBe(true);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });
});

// ── parseChart — axis reverse (orientation) ──────────────────────────

describe("parseChart — axis reverse (orientation)", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  it('surfaces reverse=true off <c:scaling><c:orientation val="maxMin"/>', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="maxMin"/></c:scaling>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.reverse).toBe(true);
  });

  it('collapses the OOXML default orientation="minMax" to undefined', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    // Neither axis has any other surfaced field, so the whole axes block drops.
    expect(chart?.axes).toBeUndefined();
  });

  it("collapses an axis with no <c:scaling> at all to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores unknown orientation tokens", () => {
    // A typo'd template (e.g. "diagonal", "reverse", empty string) drops
    // to undefined rather than fabricate a reverse flag the writer would
    // pick up.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation val="diagonal"/></c:scaling>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores <c:orientation/> with no val attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling><c:orientation/></c:scaling>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("surfaces reverse on the category axis (catAx) too", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:scaling><c:orientation val="maxMin"/></c:scaling>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.reverse).toBe(true);
    expect(chart?.axes?.y?.reverse).toBeUndefined();
  });

  it("surfaces reverse on both scatter X (axPos=b) and Y (axPos=l) axes", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:scaling><c:orientation val="maxMin"/></c:scaling>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:scaling><c:orientation val="maxMin"/></c:scaling>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.reverse).toBe(true);
    expect(chart?.axes?.y?.reverse).toBe(true);
  });

  it("surfaces reverse alongside other axis fields without interfering", () => {
    // Co-existing with min/max scaling, gridlines, numFmt, and tick rendering
    // exercises the parseAxisInfo merge — reverse pulls from <c:scaling>,
    // the others from sibling elements, so they should slot independently.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:scaling>
        <c:orientation val="maxMin"/>
        <c:max val="100"/>
        <c:min val="0"/>
      </c:scaling>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:numFmt formatCode="$#,##0" sourceLinked="0"/>
      <c:majorTickMark val="cross"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y).toEqual({
      title: "Revenue",
      gridlines: { major: true },
      scale: { min: 0, max: 100 },
      numberFormat: { formatCode: "$#,##0" },
      majorTickMark: "cross",
      reverse: true,
    });
  });
});

// ── parseChart — axis tick label / mark skip ──────────────────────

describe("parseChart — axis tickLblSkip / tickMarkSkip", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces a non-default tickLblSkip on the category axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:tickLblSkip val="3"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.tickLblSkip).toBe(3);
    expect(chart?.axes?.x?.tickMarkSkip).toBeUndefined();
  });

  it("surfaces a non-default tickMarkSkip on the category axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:tickMarkSkip val="5"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.tickMarkSkip).toBe(5);
    expect(chart?.axes?.x?.tickLblSkip).toBeUndefined();
  });

  it("surfaces both skips together when set on the same axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:tickLblSkip val="2"/>
      <c:tickMarkSkip val="4"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({ tickLblSkip: 2, tickMarkSkip: 4 });
  });

  it("collapses the OOXML default tickLblSkip=1 to undefined", () => {
    // Absence of the element and `val="1"` round-trip identically
    // through the writer's elision logic — both mean "show every label".
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:tickLblSkip val="1"/>
      <c:tickMarkSkip val="1"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores out-of-range skip values (drops rather than clamps)", () => {
    // ST_SkipIntervals restricts the value to 1..32767. Out-of-range
    // values like 0, -5, 99999 should drop rather than clamp because a
    // silent clamp would mask a configuration error.
    const out = (val: string): unknown => {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:tickLblSkip val="${val}"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      return parseChart(xml)?.axes?.x?.tickLblSkip;
    };
    expect(out("0")).toBeUndefined();
    expect(out("-5")).toBeUndefined();
    expect(out("99999")).toBeUndefined();
    expect(out("not-a-number")).toBeUndefined();
    // Boundaries 2 and 32767 are accepted.
    expect(out("2")).toBe(2);
    expect(out("32767")).toBe(32767);
  });

  it("returns undefined when tickLblSkip val attribute is missing", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:tickLblSkip/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("does not surface tickLblSkip / tickMarkSkip on a value axis", () => {
    // The OOXML schema places these elements on CT_CatAx / CT_DateAx
    // only — `<c:valAx>` rejects them entirely. A corrupt template
    // carrying a stray skip element on a value axis should not surface
    // a field the writer would never emit anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:tickLblSkip val="3"/>
      <c:tickMarkSkip val="5"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.tickLblSkip).toBeUndefined();
    expect(chart?.axes?.y?.tickMarkSkip).toBeUndefined();
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces tick skips alongside title, gridlines, scale, and number format", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Region</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:numFmt formatCode="@" sourceLinked="0"/>
      <c:tickLblSkip val="3"/>
      <c:tickMarkSkip val="6"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Region",
      gridlines: { major: true },
      numberFormat: { formatCode: "@" },
      tickLblSkip: 3,
      tickMarkSkip: 6,
    });
  });
});

// ── parseChart — axis lblOffset ────────────────────────────────────

describe("parseChart — axis lblOffset", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces a non-default lblOffset on the category axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblOffset val="250"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.lblOffset).toBe(250);
  });

  it("collapses the OOXML default lblOffset=100 to undefined", () => {
    // Excel's reference serialization always emits `<c:lblOffset val="100"/>`,
    // but absence and the default round-trip identically — the writer's
    // elision logic re-emits `100` when the field is undefined, so the
    // parser collapses `100` to undefined to keep the parsed shape minimal.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblOffset val="100"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("ignores out-of-range lblOffset values (drops rather than clamps)", () => {
    // ST_LblOffsetPercent restricts the value to 0..1000. Out-of-range
    // values like -5, 9999 should drop rather than clamp because a silent
    // clamp would mask a configuration error in the source template.
    const out = (val: string): unknown => {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblOffset val="${val}"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      return parseChart(xml)?.axes?.x?.lblOffset;
    };
    expect(out("-5")).toBeUndefined();
    expect(out("9999")).toBeUndefined();
    expect(out("not-a-number")).toBeUndefined();
    // Boundaries 0 and 1000 are accepted.
    expect(out("0")).toBe(0);
    expect(out("1000")).toBe(1000);
  });

  it("returns undefined when lblOffset val attribute is missing", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblOffset/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("does not surface lblOffset on a value axis", () => {
    // The OOXML schema places `<c:lblOffset>` on CT_CatAx / CT_DateAx
    // only — `<c:valAx>` rejects it entirely. A corrupt template carrying
    // a stray offset on a value axis should not surface a field the writer
    // would never emit anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:lblOffset val="250"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.lblOffset).toBeUndefined();
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces lblOffset alongside title, gridlines, and tick skips", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Region</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:lblOffset val="200"/>
      <c:tickLblSkip val="3"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Region",
      gridlines: { major: true },
      lblOffset: 200,
      tickLblSkip: 3,
    });
  });
});

// ── parseChart — axis hidden flag (<c:delete>) ──────────────────────

describe("parseChart — axis hidden", () => {
  const NS_HIDDEN = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces hidden=true on the category axis when val="1"', () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete val="1"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.hidden).toBe(true);
    expect(chart?.axes?.y?.hidden).toBeUndefined();
  });

  it('surfaces hidden=true on the value axis when val="1"', () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:delete val="1"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.hidden).toBeUndefined();
    expect(chart?.axes?.y?.hidden).toBe(true);
  });

  it('surfaces hidden=true on both axes when both pin val="1"', () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete val="1"/>
    </c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:delete val="1"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.hidden).toBe(true);
    expect(chart?.axes?.y?.hidden).toBe(true);
  });

  it('collapses the OOXML default val="0" to undefined', () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete val="0"/>
    </c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:delete val="0"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("collapses absence of <c:delete> to undefined", () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it('accepts the OOXML truthy spelling val="true"', () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete val="true"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.hidden).toBe(true);
  });

  it('accepts the OOXML falsy spelling val="false" and collapses to undefined', () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete val="false"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined when <c:delete> is missing the val attribute", () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined for unknown val tokens", () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete val="yes"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces hidden alongside title, gridlines, and tick rendering", () => {
    const xml = `<c:chartSpace ${NS_HIDDEN}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:delete val="1"/>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Region</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:tickLblPos val="low"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Region",
      gridlines: { major: true },
      tickLblPos: "low",
      hidden: true,
    });
  });

  it("surfaces hidden on a scatter chart's value-axis pair", () => {
    // Scatter has two valAx — the first (axPos="b") is the X axis, the
    // second (axPos="l") is the Y axis. The reader should map them back
    // to axes.x / axes.y the same way it does for the rest of the
    // metadata.
    const xml = `<c:chartSpace ${NS_HIDDEN}>
  <c:chart><c:plotArea>
    <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:delete val="1"/>
      <c:axPos val="b"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.hidden).toBe(true);
    expect(chart?.axes?.y?.hidden).toBeUndefined();
  });
});

// ── parseChart — axis lblAlgn ──────────────────────────────────────

describe("parseChart — axis lblAlgn", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces a non-default lblAlgn on the category axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblAlgn val="l"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.lblAlgn).toBe("l");
  });

  it("collapses the OOXML default lblAlgn=ctr to undefined", () => {
    // Excel's reference serialization always emits `<c:lblAlgn val="ctr"/>`,
    // but absence and the default round-trip identically — the writer's
    // elision logic re-emits `ctr` when the field is undefined, so the
    // parser collapses `ctr` to undefined to keep the parsed shape minimal.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblAlgn val="ctr"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("accepts the ST_LblAlgn tokens 'l' and 'r'", () => {
    const out = (val: string): unknown => {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblAlgn val="${val}"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      return parseChart(xml)?.axes?.x?.lblAlgn;
    };
    expect(out("l")).toBe("l");
    expect(out("r")).toBe("r");
  });

  it("ignores unknown lblAlgn tokens (drops rather than fabricates)", () => {
    // ST_LblAlgn restricts the value to ctr / l / r. Unknown tokens like
    // "left" or empty / whitespace strings should drop rather than fall
    // through to a default — a corrupt template cannot leak a value the
    // writer would never emit.
    const out = (val: string): unknown => {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblAlgn val="${val}"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      return parseChart(xml)?.axes?.x?.lblAlgn;
    };
    expect(out("left")).toBeUndefined();
    expect(out("center")).toBeUndefined();
    expect(out("LEFT")).toBeUndefined();
    expect(out("")).toBeUndefined();
  });

  it("returns undefined when lblAlgn val attribute is missing", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:lblAlgn/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("does not surface lblAlgn on a value axis", () => {
    // The OOXML schema places `<c:lblAlgn>` on CT_CatAx / CT_DateAx
    // only — `<c:valAx>` rejects it entirely. A corrupt template carrying
    // a stray alignment on a value axis should not surface a field the
    // writer would never emit anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:lblAlgn val="r"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.lblAlgn).toBeUndefined();
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces lblAlgn alongside title, gridlines, and lblOffset", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Region</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:lblAlgn val="l"/>
      <c:lblOffset val="200"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Region",
      gridlines: { major: true },
      lblOffset: 200,
      lblAlgn: "l",
    });
  });
});

// ── parseChart — legend overlay ──────────────────────────────────────

describe("parseChart — legendOverlay", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:legend><c:overlay val="1"/></c:legend> as true (non-default)', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="1"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.legend).toBe("right");
    expect(chart?.legendOverlay).toBe(true);
  });

  it("collapses the OOXML default false to undefined (writer absence)", () => {
    // The default carried explicitly by Excel's reference serialization
    // round-trips identically to absence of the field.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.legend).toBe("bottom");
    expect(chart?.legendOverlay).toBeUndefined();
  });

  it("returns undefined when the legend element omits <c:overlay>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="t"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.legend).toBe("top");
    expect(chart?.legendOverlay).toBeUndefined();
  });

  it("accepts the OOXML true / false spellings on the val attribute", () => {
    // The OOXML schema for `xsd:boolean` accepts `"true"` / `"false"`
    // alongside `"1"` / `"0"`. Hucre tolerates both shapes — a hand-
    // edited template using `true` should round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="l"/>
      <c:overlay val="true"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.legendOverlay).toBe(true);
  });

  it("collapses the 'false' spelling to undefined as well", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="false"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.legendOverlay).toBeUndefined();
  });

  it("drops unknown overlay values rather than fabricate one", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="bogus"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.legendOverlay).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:overlay>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.legendOverlay).toBeUndefined();
  });

  it('drops the overlay flag when the legend is hidden via <c:delete val="1"/>', () => {
    // A hidden legend (legend === false) has no <c:overlay> slot in the
    // rendered chart, so the reader does not surface a flag that would
    // carry no on-screen effect through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:delete val="1"/>
      <c:overlay val="1"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.legend).toBe(false);
    expect(chart?.legendOverlay).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:legend> element at all", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.legend).toBeUndefined();
    expect(chart?.legendOverlay).toBeUndefined();
  });

  it("surfaces overlay on every chart family that emits a legend", () => {
    // The element lives on <c:legend>, which is a sibling of
    // <c:plotArea> on every chart-family <c:chart>; the toggle should
    // round-trip identically across families. Pie / doughnut / line /
    // bar all emit legends by default.
    for (const kind of ["lineChart", "barChart", "pieChart", "doughnutChart"]) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:${kind}><c:ser><c:idx val="0"/></c:ser></c:${kind}>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="1"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
      const chart = parseChart(xml);
      expect(chart?.legendOverlay).toBe(true);
    }
  });

  it("co-exists with other chart-level toggles", () => {
    // The legend overlay flag should not interfere with sibling chart-
    // level fields parsed off <c:chart> (plotVisOnly, dispBlanksAs,
    // varyColors).
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="t"/>
      <c:overlay val="1"/>
    </c:legend>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.legend).toBe("top");
    expect(chart?.legendOverlay).toBe(true);
    expect(chart?.roundedCorners).toBe(true);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });
});

// ── parseChart — data labels showLegendKey ──────────────────────────

describe("parseChart — data labels showLegendKey", () => {
  const NS_LK = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces showLegendKey=true on chart-level dLbls when val="1"', () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:dLblPos val="outEnd"/>
        <c:showLegendKey val="1"/>
        <c:showVal val="1"/>
        <c:showCatName val="0"/>
        <c:showSerName val="0"/>
        <c:showPercent val="0"/>
        <c:showBubbleSize val="0"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toEqual({
      position: "outEnd",
      showLegendKey: true,
      showValue: true,
    });
  });

  it('collapses the OOXML default val="0" to undefined', () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showLegendKey val="0"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.showLegendKey).toBeUndefined();
    expect(chart?.dataLabels?.showValue).toBe(true);
  });

  it("collapses absence of <c:showLegendKey> to undefined", () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.showLegendKey).toBeUndefined();
  });

  it('accepts the OOXML truthy spelling val="true"', () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showLegendKey val="true"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.showLegendKey).toBe(true);
  });

  it('accepts the OOXML falsy spelling val="false" and collapses to undefined', () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showLegendKey val="false"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.showLegendKey).toBeUndefined();
  });

  it("ignores unknown val tokens", () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showLegendKey val="yes"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.showLegendKey).toBeUndefined();
  });

  it("returns undefined when <c:showLegendKey> is missing the val attribute", () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showLegendKey/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.showLegendKey).toBeUndefined();
  });

  it("makes showLegendKey alone enough to surface a dataLabels record", () => {
    // Even when no value/category/series/percent toggle is on, a pinned
    // showLegendKey=true is still meaningful — Excel renders the legend
    // swatch beside each (otherwise empty) label slot. The reader must
    // not collapse the block to undefined in that case.
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showLegendKey val="1"/>
        <c:showVal val="0"/>
        <c:showCatName val="0"/>
        <c:showSerName val="0"/>
        <c:showPercent val="0"/>
        <c:showBubbleSize val="0"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toEqual({ showLegendKey: true });
  });

  it("surfaces showLegendKey on a series-level <c:dLbls>", () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser>
        <c:idx val="0"/>
        <c:tx><c:v>Revenue</c:v></c:tx>
        <c:dLbls>
          <c:dLblPos val="ctr"/>
          <c:showLegendKey val="1"/>
          <c:showVal val="1"/>
        </c:dLbls>
        <c:val><c:numRef><c:f>S!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].dataLabels).toEqual({
      position: "ctr",
      showLegendKey: true,
      showValue: true,
    });
  });

  it("co-surfaces showLegendKey alongside other show toggles and separator", () => {
    const xml = `<c:chartSpace ${NS_LK}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:dLblPos val="bestFit"/>
        <c:showLegendKey val="1"/>
        <c:showVal val="1"/>
        <c:showCatName val="1"/>
        <c:showSerName val="0"/>
        <c:showPercent val="1"/>
        <c:showBubbleSize val="0"/>
        <c:separator>; </c:separator>
      </c:dLbls>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toEqual({
      position: "bestFit",
      showLegendKey: true,
      showValue: true,
      showCategoryName: true,
      showPercent: true,
      separator: "; ",
    });
  });
});

// ── parseChart — data labels numberFormat ──────────────────────────

describe("parseChart — data labels numberFormat", () => {
  const NS_NF = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces numberFormat from a chart-level <c:dLbls><c:numFmt>", () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="0.00%" sourceLinked="0"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toEqual({ formatCode: "0.00%" });
  });

  it("surfaces sourceLinked=true when the OOXML attribute is pinned to 1", () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toEqual({
      formatCode: "General",
      sourceLinked: true,
    });
  });

  it('accepts the OOXML truthy spelling sourceLinked="true"', () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="0.0" sourceLinked="true"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toEqual({
      formatCode: "0.0",
      sourceLinked: true,
    });
  });

  it('collapses the OOXML default sourceLinked="0" to undefined', () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="$#,##0.00" sourceLinked="0"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toEqual({ formatCode: "$#,##0.00" });
  });

  it("collapses absence of <c:numFmt> to undefined", () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toBeUndefined();
  });

  it("returns undefined when <c:numFmt> is missing the formatCode attribute", () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt sourceLinked="0"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toBeUndefined();
  });

  it("returns undefined for empty formatCode strings", () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="" sourceLinked="0"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toBeUndefined();
  });

  it("makes numberFormat alone enough to surface a dataLabels record", () => {
    // The number format pin is meaningful even without any show* toggle —
    // Excel still applies it to a per-series label override that turns
    // the labels on. The reader must not collapse the block to undefined
    // when the only pinned field is the numFmt.
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="0.00%" sourceLinked="0"/>
        <c:showLegendKey val="0"/>
        <c:showVal val="0"/>
        <c:showCatName val="0"/>
        <c:showSerName val="0"/>
        <c:showPercent val="0"/>
        <c:showBubbleSize val="0"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toEqual({ numberFormat: { formatCode: "0.00%" } });
  });

  it("surfaces numberFormat on a series-level <c:dLbls>", () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser>
        <c:idx val="0"/>
        <c:tx><c:v>Revenue</c:v></c:tx>
        <c:dLbls>
          <c:numFmt formatCode="$#,##0" sourceLinked="0"/>
          <c:dLblPos val="ctr"/>
          <c:showVal val="1"/>
        </c:dLbls>
        <c:val><c:numRef><c:f>S!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.series?.[0].dataLabels).toEqual({
      position: "ctr",
      showValue: true,
      numberFormat: { formatCode: "$#,##0" },
    });
  });

  it("co-surfaces numberFormat alongside other dataLabels fields", () => {
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="0.00%" sourceLinked="0"/>
        <c:dLblPos val="bestFit"/>
        <c:showLegendKey val="1"/>
        <c:showVal val="0"/>
        <c:showCatName val="1"/>
        <c:showSerName val="0"/>
        <c:showPercent val="1"/>
        <c:showBubbleSize val="0"/>
        <c:separator>; </c:separator>
      </c:dLbls>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels).toEqual({
      position: "bestFit",
      showLegendKey: true,
      showCategoryName: true,
      showPercent: true,
      separator: "; ",
      numberFormat: { formatCode: "0.00%" },
    });
  });

  it("does not leak a per-point <c:dLbl><c:numFmt> into the block-level record", () => {
    // The CT_DLbls schema allows a `<c:numFmt>` to live inside a
    // per-point `<c:dLbl>` as well as at the block level. The reader
    // scopes its lookup to direct `<c:dLbls>` children so only the
    // block-level pin surfaces — a per-point override on a single data
    // point cannot pollute the chart-level record.
    const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:dLbl>
          <c:idx val="0"/>
          <c:numFmt formatCode="0.00" sourceLinked="0"/>
          <c:showVal val="1"/>
        </c:dLbl>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataLabels?.numberFormat).toBeUndefined();
  });

  it("threads numberFormat through line / pie / scatter chart families", () => {
    const families = [
      ["lineChart", "line"],
      ["pieChart", "pie"],
      ["scatterChart", "scatter"],
    ] as const;
    for (const [tag] of families) {
      const xml = `<c:chartSpace ${NS_NF}>
  <c:chart><c:plotArea>
    <c:${tag}>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dLbls>
        <c:numFmt formatCode="0.00%" sourceLinked="0"/>
        <c:showVal val="1"/>
      </c:dLbls>
    </c:${tag}>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      const chart = parseChart(xml);
      expect(chart?.dataLabels?.numberFormat).toEqual({ formatCode: "0.00%" });
    }
  });
});

// ── parseChart — axis noMultiLvlLbl ────────────────────────────────

describe("parseChart — axis noMultiLvlLbl", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces noMultiLvlLbl=true on the category axis when val="1"', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:noMultiLvlLbl val="1"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.noMultiLvlLbl).toBe(true);
    expect(chart?.axes?.y?.noMultiLvlLbl).toBeUndefined();
  });

  it('collapses the OOXML default val="0" to undefined', () => {
    // Excel's reference serialization emits `<c:noMultiLvlLbl val="0"/>`
    // on every category axis even though the schema default is `false`.
    // The parser collapses the default so absence and the default
    // round-trip identically through the writer's elision logic.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:noMultiLvlLbl val="0"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("collapses absence of <c:noMultiLvlLbl> to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it('accepts the OOXML truthy spelling val="true"', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:noMultiLvlLbl val="true"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.noMultiLvlLbl).toBe(true);
  });

  it('accepts the OOXML falsy spelling val="false" and collapses to undefined', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:noMultiLvlLbl val="false"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined when <c:noMultiLvlLbl> is missing the val attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:noMultiLvlLbl/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined for unknown val tokens", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:noMultiLvlLbl val="yes"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("does not surface noMultiLvlLbl on a value axis", () => {
    // The OOXML schema places `<c:noMultiLvlLbl>` on CT_CatAx exclusively
    // — `<c:valAx>` rejects it entirely. A corrupt template carrying a
    // stray flag on a value axis should not surface a field the writer
    // would never emit anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:noMultiLvlLbl val="1"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.noMultiLvlLbl).toBeUndefined();
    expect(chart?.axes).toBeUndefined();
  });

  it("does not surface noMultiLvlLbl on a scatter chart's value axes", () => {
    // Scatter has two valAx — the schema rejects `<c:noMultiLvlLbl>` on
    // both, so a stray flag on either axis must not bleed through into
    // the parsed shape.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:noMultiLvlLbl val="1"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:noMultiLvlLbl val="1"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.noMultiLvlLbl).toBeUndefined();
    expect(chart?.axes?.y?.noMultiLvlLbl).toBeUndefined();
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces noMultiLvlLbl alongside title, gridlines, and other catAx fields", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:majorGridlines/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Region</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:lblOffset val="200"/>
      <c:tickLblSkip val="3"/>
      <c:noMultiLvlLbl val="1"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Region",
      gridlines: { major: true },
      lblOffset: 200,
      tickLblSkip: 3,
      noMultiLvlLbl: true,
    });
  });
});

// ── parseChart — axis auto ─────────────────────────────────────────

describe("parseChart — axis auto", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces auto=false on the category axis when val="0"', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:auto val="0"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.auto).toBe(false);
    expect(chart?.axes?.y?.auto).toBeUndefined();
  });

  it('collapses the OOXML default val="1" to undefined', () => {
    // Excel's reference serialization always emits `<c:auto val="1"/>`
    // on every category axis even though the schema default is `true`.
    // The parser collapses the default so absence and the default
    // round-trip identically through the writer's elision logic.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:auto val="1"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("collapses absence of <c:auto> to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it('accepts the OOXML falsy spelling val="false"', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:auto val="false"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.auto).toBe(false);
  });

  it('accepts the OOXML truthy spelling val="true" and collapses to undefined', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:auto val="true"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined when <c:auto> is missing the val attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:auto/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined for unknown val tokens", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:auto val="maybe"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("does not surface auto on a value axis", () => {
    // The OOXML schema places `<c:auto>` on CT_CatAx exclusively —
    // `<c:valAx>` rejects it entirely. A corrupt template carrying a
    // stray flag on a value axis should not surface a field the writer
    // would never emit anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:auto val="0"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.auto).toBeUndefined();
    expect(chart?.axes).toBeUndefined();
  });

  it("does not surface auto on a scatter chart's value axes", () => {
    // Scatter has two valAx — the schema rejects `<c:auto>` on both, so
    // a stray flag on either axis must not bleed through into the
    // parsed shape.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:auto val="0"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:auto val="0"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.auto).toBeUndefined();
    expect(chart?.axes?.y?.auto).toBeUndefined();
    expect(chart?.axes).toBeUndefined();
  });

  it("co-surfaces auto alongside other catAx fields", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Period</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:auto val="0"/>
      <c:noMultiLvlLbl val="1"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Period",
      auto: false,
      noMultiLvlLbl: true,
    });
  });
});

// ── parseChart — titleOverlay ───────────────────────────────────────

describe("parseChart — titleOverlay", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  function withTitle(overlay?: string): string {
    const overlayElement = overlay === undefined ? "" : `<c:overlay val="${overlay}"/>`;
    return `<c:chartSpace ${NS}>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Sales</a:t></a:r></a:p>
      </c:rich></c:tx>
      ${overlayElement}
    </c:title>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it('surfaces <c:title><c:overlay val="1"/></c:title> as true (non-default)', () => {
    const chart = parseChart(withTitle("1"));
    expect(chart?.title).toBe("Sales");
    expect(chart?.titleOverlay).toBe(true);
  });

  it("collapses the OOXML default false to undefined (writer absence)", () => {
    // The default carried explicitly by Excel's reference serialization
    // round-trips identically to absence of the field.
    const chart = parseChart(withTitle("0"));
    expect(chart?.title).toBe("Sales");
    expect(chart?.titleOverlay).toBeUndefined();
  });

  it("returns undefined when the title element omits <c:overlay>", () => {
    const chart = parseChart(withTitle());
    expect(chart?.title).toBe("Sales");
    expect(chart?.titleOverlay).toBeUndefined();
  });

  it("accepts the OOXML true / false spellings on the val attribute", () => {
    // The OOXML schema for `xsd:boolean` accepts `"true"` / `"false"`
    // alongside `"1"` / `"0"`. Hucre tolerates both shapes — a hand-
    // edited template using `true` should round-trip.
    expect(parseChart(withTitle("true"))?.titleOverlay).toBe(true);
  });

  it("collapses the 'false' spelling to undefined as well", () => {
    expect(parseChart(withTitle("false"))?.titleOverlay).toBeUndefined();
  });

  it("drops unknown overlay values rather than fabricate one", () => {
    expect(parseChart(withTitle("bogus"))?.titleOverlay).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:overlay>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Sales</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay/>
    </c:title>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.titleOverlay).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:title> element at all", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.title).toBeUndefined();
    expect(chart?.titleOverlay).toBeUndefined();
  });

  it("surfaces the overlay flag on every chart family that emits a title", () => {
    // The element lives on <c:title>, a chart-level sibling of
    // <c:plotArea>, so it round-trips identically across families. Pie
    // / doughnut / line / bar all support the title block identically.
    for (const kind of ["lineChart", "barChart", "pieChart", "doughnutChart"]) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Header</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="1"/>
    </c:title>
    <c:plotArea>
      <c:${kind}><c:ser><c:idx val="0"/></c:ser></c:${kind}>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      const chart = parseChart(xml);
      expect(chart?.titleOverlay).toBe(true);
    }
  });

  it("co-exists independently with the legend overlay flag", () => {
    // The chart-title `<c:overlay>` lives on `<c:title>`, while the
    // legend `<c:overlay>` lives on `<c:legend>`; the two flags must
    // surface independently from the same chart even though they share
    // a local element name.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Sales</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="1"/>
    </c:title>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.titleOverlay).toBe(true);
    expect(chart?.legendOverlay).toBeUndefined();
  });

  it("co-exists with other chart-level toggles", () => {
    // The title overlay flag should not interfere with sibling chart-
    // level fields parsed off <c:chart> / <c:chartSpace>.
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Sales</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="1"/>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.title).toBe("Sales");
    expect(chart?.titleOverlay).toBe(true);
    expect(chart?.roundedCorners).toBe(true);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });

  it("does not pull <c:overlay> from a sibling element by mistake", () => {
    // The reader must scope the lookup to direct `<c:title>` children
    // — if the title omits an `<c:overlay>` but the legend has one,
    // the title flag must not pick up the legend's value.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Sales</a:t></a:r></a:p>
      </c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="1"/>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.titleOverlay).toBeUndefined();
    expect(chart?.legendOverlay).toBe(true);
  });
});

// ── parseChart — axis crosses / crossesAt ──────────────────────────

describe("parseChart — axis crosses / crossesAt", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces crosses="min" on the category axis', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:crosses val="min"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crosses).toBe("min");
    expect(chart?.axes?.x?.crossesAt).toBeUndefined();
    expect(chart?.axes?.y?.crosses).toBeUndefined();
  });

  it('surfaces crosses="max" on the value axis', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crosses val="max"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crosses).toBeUndefined();
    expect(chart?.axes?.y?.crosses).toBe("max");
  });

  it('collapses the OOXML default crosses="autoZero" to undefined', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:crosses val="autoZero"/>
    </c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crosses val="autoZero"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("collapses absence of <c:crosses> to undefined", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined for unknown crosses tokens", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:crosses val="middle"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined when <c:crosses> is missing the val attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:crosses/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("surfaces a positive crossesAt on the value axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossesAt val="50"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossesAt).toBe(50);
    expect(chart?.axes?.y?.crosses).toBeUndefined();
  });

  it("surfaces a negative crossesAt", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossesAt val="-25.5"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossesAt).toBe(-25.5);
  });

  it("preserves crossesAt=0 (distinct from autoZero)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossesAt val="0"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossesAt).toBe(0);
    expect(chart?.axes?.y?.crosses).toBeUndefined();
  });

  it("returns undefined when <c:crossesAt> is missing the val attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossesAt/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined when <c:crossesAt val> is non-numeric", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossesAt val="middle"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes).toBeUndefined();
  });

  it("prefers crossesAt over crosses when both are present (XSD choice)", () => {
    // The OOXML schema places <c:crosses> and <c:crossesAt> in an XSD
    // choice — only one may legally appear. The reader handles
    // malformed templates that emit both by keeping the numeric pin
    // (mirroring the writer's emit order).
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:crosses val="max"/>
      <c:crossesAt val="42"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crosses).toBeUndefined();
    expect(chart?.axes?.x?.crossesAt).toBe(42);
  });

  it("falls back to crosses when crossesAt is unparseable", () => {
    // Same malformed-template guard, the other direction: when
    // crossesAt is present but unparseable, the semantic crosses still
    // surfaces.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:crosses val="min"/>
      <c:crossesAt/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crosses).toBe("min");
    expect(chart?.axes?.x?.crossesAt).toBeUndefined();
  });

  it("surfaces crosses on a scatter chart's value-axis pair", () => {
    // Scatter has two valAx — the first (axPos="b") is the X axis, the
    // second (axPos="l") is the Y axis. The reader maps them back to
    // axes.x / axes.y the same way it does for the rest of the
    // metadata.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:crossesAt val="3.14"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:crosses val="max"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crossesAt).toBe(3.14);
    expect(chart?.axes?.y?.crosses).toBe("max");
  });

  it("co-surfaces crosses alongside title and tick rendering", () => {
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Region</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:tickLblPos val="low"/>
      <c:crosses val="min"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Region",
      tickLblPos: "low",
      crosses: "min",
    });
  });
});

// ── parseChart — drop / hi-low lines ──────────────────────────────

describe("parseChart — drop lines", () => {
  function lineChartWithExtras(extras: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser><c:idx val="0"/></c:ser>
        ${extras}
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  function areaChartWithExtras(extras: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:grouping val="standard"/>
        <c:ser><c:idx val="0"/></c:ser>
        ${extras}
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it("surfaces dropLines=true on a line chart that declares <c:dropLines/>", () => {
    const xml = lineChartWithExtras("<c:dropLines/>");
    expect(parseChart(xml)?.dropLines).toBe(true);
  });

  it("surfaces dropLines=true on a line chart that declares <c:dropLines> with a nested <c:spPr>", () => {
    // CT_ChartLines may carry `<c:spPr>` for stroke styling. The
    // reader only surfaces the on/off bit — the shape properties are
    // not modelled in this phase but the presence of the element still
    // surfaces `true` so the clone bridge can carry the intent.
    const xml = lineChartWithExtras(
      `<c:dropLines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr></c:dropLines>`,
    );
    expect(parseChart(xml)?.dropLines).toBe(true);
  });

  it("returns undefined when the line chart omits <c:dropLines>", () => {
    const xml = lineChartWithExtras("");
    expect(parseChart(xml)?.dropLines).toBeUndefined();
  });

  it("surfaces dropLines=true on an area chart that declares <c:dropLines/>", () => {
    const xml = areaChartWithExtras("<c:dropLines/>");
    expect(parseChart(xml)?.dropLines).toBe(true);
  });

  it("returns undefined when the area chart omits <c:dropLines>", () => {
    const xml = areaChartWithExtras("");
    expect(parseChart(xml)?.dropLines).toBeUndefined();
  });

  it("does not surface dropLines for chart kinds that have no <c:dropLines> slot (bar)", () => {
    // The reader only inspects line / line3D / area / area3D children;
    // a stray `<c:dropLines>` on a bar chart (which the OOXML schema
    // rejects) must not surface a value.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dropLines/>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dropLines).toBeUndefined();
  });

  it("surfaces dropLines on the first line/area chart-type element only (combo workbook)", () => {
    // Combo charts (multi-kind plot area) surface the first matching
    // value just like the existing grouping helpers.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:dropLines/>
      </c:lineChart>
      <c:areaChart>
        <c:ser><c:idx val="1"/></c:ser>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dropLines).toBe(true);
  });
});

describe("parseChart — high-low lines", () => {
  function lineChartWithExtras(extras: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser><c:idx val="0"/></c:ser>
        ${extras}
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it("surfaces hiLowLines=true on a line chart that declares <c:hiLowLines/>", () => {
    const xml = lineChartWithExtras("<c:hiLowLines/>");
    expect(parseChart(xml)?.hiLowLines).toBe(true);
  });

  it("surfaces hiLowLines=true on a line chart that declares <c:hiLowLines> with a nested <c:spPr>", () => {
    const xml = lineChartWithExtras(
      `<c:hiLowLines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr></c:hiLowLines>`,
    );
    expect(parseChart(xml)?.hiLowLines).toBe(true);
  });

  it("returns undefined when the line chart omits <c:hiLowLines>", () => {
    const xml = lineChartWithExtras("");
    expect(parseChart(xml)?.hiLowLines).toBeUndefined();
  });

  it("does not surface hiLowLines on an area chart (no slot in the OOXML schema)", () => {
    // The reader's per-kind gate excludes `area` / `area3D` from the
    // hiLowLines lookup — `<c:hiLowLines>` lives only on lineChart /
    // line3DChart / stockChart per CT_AreaChart rejecting the element.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:hiLowLines/>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.hiLowLines).toBeUndefined();
  });

  it("does not surface hiLowLines for chart kinds that have no slot (bar)", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:hiLowLines/>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.hiLowLines).toBeUndefined();
  });

  it("surfaces both dropLines and hiLowLines together on a line chart", () => {
    const xml = lineChartWithExtras("<c:dropLines/><c:hiLowLines/>");
    const parsed = parseChart(xml);
    expect(parsed?.dropLines).toBe(true);
    expect(parsed?.hiLowLines).toBe(true);
  });

  it("surfaces hiLowLines on a stockChart (the third OOXML host for the element)", () => {
    // hucre's writer never authors `<c:stockChart>`, but a parsed
    // stock-chart template should round-trip the flag so a downstream
    // tool that introspects it gets the right answer.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:stockChart>
        <c:ser><c:idx val="0"/></c:ser>
        <c:hiLowLines/>
      </c:stockChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.hiLowLines).toBe(true);
  });
});

// ── parseChart — series lines ──────────────────────────────────────

describe("parseChart — series lines", () => {
  function barChartWithExtras(extras: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="stacked"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:ser><c:idx val="1"/></c:ser>
        ${extras}
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it("surfaces serLines=true on a bar chart that declares <c:serLines/>", () => {
    const xml = barChartWithExtras("<c:serLines/>");
    expect(parseChart(xml)?.serLines).toBe(true);
  });

  it("surfaces serLines=true on a bar chart that declares <c:serLines> with a nested <c:spPr>", () => {
    // CT_ChartLines may carry `<c:spPr>` for stroke styling. The
    // reader only surfaces the on/off bit — the shape properties are
    // not modelled in this phase but the presence of the element still
    // surfaces `true` so the clone bridge can carry the intent.
    const xml = barChartWithExtras(
      `<c:serLines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr></c:serLines>`,
    );
    expect(parseChart(xml)?.serLines).toBe(true);
  });

  it("returns undefined when the bar chart omits <c:serLines>", () => {
    const xml = barChartWithExtras("");
    expect(parseChart(xml)?.serLines).toBeUndefined();
  });

  it("does not surface serLines for chart kinds that have no <c:serLines> slot (line)", () => {
    // The reader only inspects bar / ofPie children; a stray
    // `<c:serLines>` on a line chart (which the OOXML schema rejects)
    // must not surface a value.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:serLines/>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.serLines).toBeUndefined();
  });

  it("does not surface serLines for chart kinds that have no slot (pie)", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:serLines/>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.serLines).toBeUndefined();
  });

  it("does not surface serLines for chart kinds that have no slot (area)", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:grouping val="stacked"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:serLines/>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.serLines).toBeUndefined();
  });

  it("surfaces serLines on an ofPieChart (the second OOXML host for the element)", () => {
    // hucre's writer never authors `<c:ofPieChart>`, but a parsed
    // ofPie template carrying the element should round-trip the flag
    // so a downstream tool that introspects it gets the right answer.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:ofPieChart>
        <c:ofPieType val="pie"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:serLines/>
      </c:ofPieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.serLines).toBe(true);
  });

  it("surfaces serLines on the first bar/ofPie chart-type element only (combo workbook)", () => {
    // Combo charts (multi-kind plot area) surface the first matching
    // value just like the existing connector-line helpers.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="stacked"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:serLines/>
      </c:barChart>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser><c:idx val="1"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.serLines).toBe(true);
  });

  it("surfaces serLines on a clustered bar chart even though Excel paints nothing", () => {
    // The OOXML element pins regardless of the grouping; Excel only
    // renders the connectors on stacked groupings, but the model is a
    // plain presence flag so the reader should not gate on grouping.
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:serLines/>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.serLines).toBe(true);
  });
});

// ── parseChart — upDownBars ────────────────────────────────────────

describe("parseChart — upDownBars", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces upDownBars=true on a line chart with the bare element", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBe(true);
  });

  it("surfaces upDownBars=true when the element carries the optional gapWidth child", () => {
    // Excel's reference serialization includes <c:gapWidth val="150"/>
    // inside <c:upDownBars>. The model is a presence flag at this
    // layer, so the child should not change the surfaced value.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars>
        <c:gapWidth val="150"/>
      </c:upDownBars>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBe(true);
  });

  it("surfaces upDownBars=true when the element carries upBars / downBars children", () => {
    // <c:upBars> and <c:downBars> are CT_UpDownBar — each with an
    // optional <c:spPr>. Their presence does not change the bare
    // toggle exposed at this layer.
    const xml = `<c:chartSpace ${NS}
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars>
        <c:gapWidth val="200"/>
        <c:upBars>
          <c:spPr>
            <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          </c:spPr>
        </c:upBars>
        <c:downBars>
          <c:spPr>
            <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          </c:spPr>
        </c:downBars>
      </c:upDownBars>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBe(true);
  });

  it("collapses absence of <c:upDownBars> to undefined on a line chart", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBeUndefined();
  });

  it("ignores <c:upDownBars> on a bar chart (CT_BarChart rejects the element)", () => {
    // The OOXML schema places <c:upDownBars> on CT_LineChart /
    // CT_Line3DChart / CT_StockChart only. A stray element on a bar
    // / column chart is not surfaced — the writer would never emit
    // it there.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBeUndefined();
  });

  it("ignores <c:upDownBars> on an area chart (CT_AreaChart rejects the element)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:areaChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:areaChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBeUndefined();
  });

  it("ignores <c:upDownBars> on a scatter chart (CT_ScatterChart rejects the element)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:scatterChart>
    <c:valAx><c:axId val="1"/></c:valAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBeUndefined();
  });

  it("surfaces upDownBars on a stock chart (CT_StockChart accepts the element)", () => {
    // CT_StockChart is where Excel typically paints up/down bars in the
    // wild (open / close). The reader accepts the element on any
    // line-flavored chart-type body.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:stockChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:stockChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBe(true);
  });

  it("surfaces upDownBars on a 3D line chart (CT_Line3DChart accepts the element)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:line3DChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:upDownBars/>
      <c:axId val="1"/>
      <c:axId val="2"/>
      <c:axId val="3"/>
    </c:line3DChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:serAx><c:axId val="3"/></c:serAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBe(true);
  });

  it("co-surfaces upDownBars alongside other chart-level fields", () => {
    // upDownBars sits inside the chart-type element (line/stock) and
    // should not interfere with chart-level toggles like dispBlanksAs
    // / plotVisOnly that live on <c:chart> itself.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:upDownBars>
          <c:gapWidth val="150"/>
        </c:upDownBars>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.upDownBars).toBe(true);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
  });
});

// ── parseChart — axis dispUnits ──────────────────────────────────────

describe("parseChart — axis dispUnits", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces a built-in unit preset on the value axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits><c:builtInUnit val="millions"/></c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toEqual({ unit: "millions" });
    expect(chart?.axes?.x?.dispUnits).toBeUndefined();
  });

  it("surfaces showLabel when <c:dispUnitsLbl> is present", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits>
        <c:builtInUnit val="thousands"/>
        <c:dispUnitsLbl/>
      </c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toEqual({ unit: "thousands", showLabel: true });
  });

  it("collapses dispUnits to undefined on a category axis (catAx rejects the element)", () => {
    // The OOXML schema places <c:dispUnits> exclusively on CT_ValAx, so a
    // stray element on <c:catAx> from a corrupt template should never
    // surface — the reader explicitly skips the parse on every non-valAx
    // flavour.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:dispUnits><c:builtInUnit val="millions"/></c:dispUnits>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.dispUnits).toBeUndefined();
  });

  it("drops an unknown ST_BuiltInUnit token rather than fabricating a value", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits><c:builtInUnit val="quintillions"/></c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toBeUndefined();
  });

  it("drops the parsed value when neither <c:builtInUnit> nor <c:custUnit> resolves", () => {
    // <c:dispUnits> has an xsd:choice between <c:builtInUnit> and
    // <c:custUnit>. A bare <c:dispUnits>, or one whose lone child has
    // a missing / malformed `val`, surfaces nothing.
    const bare = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:dispUnits/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(bare)?.axes?.y?.dispUnits).toBeUndefined();

    const noVal = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:dispUnits><c:builtInUnit/></c:dispUnits></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(noVal)?.axes?.y?.dispUnits).toBeUndefined();

    const noCustVal = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:dispUnits><c:custUnit/></c:dispUnits></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(noCustVal)?.axes?.y?.dispUnits).toBeUndefined();
  });

  it("surfaces dispUnits on both scatter axes (both are valAx)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:dispUnits><c:builtInUnit val="hundreds"/></c:dispUnits>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:dispUnits><c:builtInUnit val="billions"/><c:dispUnitsLbl/></c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.dispUnits).toEqual({ unit: "hundreds" });
    expect(chart?.axes?.y?.dispUnits).toEqual({ unit: "billions", showLabel: true });
  });

  it("collapses dispUnits to undefined when the chart has no <c:dispUnits>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toBeUndefined();
  });

  it("surfaces a custom numeric divisor on the value axis", () => {
    // Excel's "Display units → Other" path emits <c:custUnit val=".."/>
    // instead of <c:builtInUnit>. The reader surfaces the numeric divisor
    // as `custUnit` so a templated chart that pins an arbitrary divisor
    // (e.g. 86400 to convert seconds to days) round-trips through clone.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits><c:custUnit val="86400"/></c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toEqual({ custUnit: 86400 });
  });

  it("parses fractional custUnit values per the OOXML CT_Double schema", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits><c:custUnit val="2.5"/></c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toEqual({ custUnit: 2.5 });
  });

  it("surfaces showLabel alongside custUnit when <c:dispUnitsLbl> is present", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits>
        <c:custUnit val="500"/>
        <c:dispUnitsLbl/>
      </c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toEqual({ custUnit: 500, showLabel: true });
  });

  it("drops custUnit values that are non-positive, non-finite, or malformed", () => {
    // OOXML CT_Double accepts any double, but a divisor of 0 or below
    // would silently break the rendered scale. The reader requires a
    // finite positive value rather than fabricate a token Excel rejects.
    for (const raw of ["0", "-100", "NaN", "Infinity", "-Infinity", "abc", ""]) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits><c:custUnit val="${raw}"/></c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.axes?.y?.dispUnits).toBeUndefined();
    }
  });

  it("prefers <c:custUnit> over <c:builtInUnit> when a malformed template declares both", () => {
    // The OOXML schema's xsd:choice forbids both children, but a
    // corrupt template may carry both. The reader picks `custUnit`
    // (the more specific element) and drops the preset; the writer
    // mirrors this preference so the round-trip stays consistent.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits>
        <c:custUnit val="1500"/>
        <c:builtInUnit val="thousands"/>
      </c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toEqual({ custUnit: 1500 });
  });

  it("falls back to <c:builtInUnit> when <c:custUnit> is present but malformed", () => {
    // A `<c:custUnit>` whose `val` fails the positive-finite gate
    // does not poison the parse — the reader falls back to the
    // sibling `<c:builtInUnit>` (also out-of-spec for the schema's
    // choice, but a corrupt template may declare both).
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:dispUnits>
        <c:custUnit val="-5"/>
        <c:builtInUnit val="hundreds"/>
      </c:dispUnits>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.dispUnits).toEqual({ unit: "hundreds" });
  });

  it("collapses custUnit to undefined on a category axis (catAx rejects <c:dispUnits>)", () => {
    // Same scope rule as the built-in preset — `<c:dispUnits>` lives
    // exclusively on `CT_ValAx`, so a stray element on `<c:catAx>`
    // never surfaces.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:dispUnits><c:custUnit val="500"/></c:dispUnits>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.dispUnits).toBeUndefined();
  });
});

// ── parseChart — chart style preset ────────────────────────────────

describe("parseChart — chart style preset", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:style val="2"/> on <c:chartSpace> as the integer 2', () => {
    // Excel's reference serialization for a fresh chart pins style 2 —
    // it surfaces verbatim because the reader does not collapse a
    // default (a chart that omits the element renders identically).
    const xml = `<c:chartSpace ${NS}>
  <c:style val="2"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.style).toBe(2);
  });

  it("surfaces a templated mid-range preset (style 27)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:style val="27"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.style).toBe(27);
  });

  it("surfaces the OOXML range bounds (1 and 48)", () => {
    for (const val of [1, 48]) {
      const xml = `<c:chartSpace ${NS}>
  <c:style val="${val}"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.style).toBe(val);
    }
  });

  it("returns undefined when the chartSpace has no <c:style> element", () => {
    // Absence is the writer's default — the reader surfaces nothing
    // so a fresh chart and a chart that omits the element round-trip
    // identically through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.style).toBeUndefined();
  });

  it("drops out-of-range style values (0 / 49)", () => {
    // CT_Style declares `val` as `xsd:unsignedByte` in the gallery
    // range 1–48; values outside collapse to undefined rather than
    // surface a token Excel would not emit.
    for (const val of ["0", "49", "255"]) {
      const xml = `<c:chartSpace ${NS}>
  <c:style val="${val}"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.style).toBeUndefined();
    }
  });

  it("drops non-integer style values rather than fabricate one", () => {
    // The OOXML schema forbids fractional / negative / alpha values.
    for (const val of ["3.5", "-1", "two", "3px"]) {
      const xml = `<c:chartSpace ${NS}>
  <c:style val="${val}"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.style).toBeUndefined();
    }
  });

  it("ignores a missing val attribute on <c:style>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:style/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.style).toBeUndefined();
  });

  it("surfaces style alongside roundedCorners and other chart-level toggles", () => {
    // <c:style> sits on <c:chartSpace> (after <c:roundedCorners>) per
    // the CT_ChartSpace sequence. Co-existing with chart-level toggles
    // that live on <c:chart> should not interfere.
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="1"/>
  <c:style val="34"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.style).toBe(34);
    expect(chart?.roundedCorners).toBe(true);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });
});

// ── parseChart — chart editing locale (lang) ───────────────────────

describe("parseChart — chart editing locale", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:lang val="en-US"/> on <c:chartSpace> as the string "en-US"', () => {
    // Excel's reference serialization for a fresh chart authored on
    // an English locale pins lang en-US — it surfaces verbatim
    // because the reader does not collapse a default (a chart that
    // omits the element renders identically).
    const xml = `<c:chartSpace ${NS}>
  <c:lang val="en-US"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.lang).toBe("en-US");
  });

  it("surfaces a templated non-English locale (tr-TR)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:lang val="tr-TR"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.lang).toBe("tr-TR");
  });

  it("surfaces locale shapes Excel actually emits", () => {
    // Sample of the BCP-47 forms <c:lang> accepts under xsd:language.
    for (const tag of ["en-US", "tr-TR", "de-DE", "pt-BR", "zh-Hans-CN", "fr"]) {
      const xml = `<c:chartSpace ${NS}>
  <c:lang val="${tag}"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.lang).toBe(tag);
    }
  });

  it("returns undefined when the chartSpace has no <c:lang> element", () => {
    // Absence is the writer's default — the reader surfaces nothing
    // so a fresh chart and a chart that omits the element round-trip
    // identically through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.lang).toBeUndefined();
  });

  it("drops malformed locale tokens rather than surface them", () => {
    // <c:lang> is xsd:language (BCP-47 culture name). Garbage values
    // collapse to undefined so the parsed chart never carries a token
    // Excel itself would not emit.
    for (const bad of ["english", "en US", "en_US", "1234", "en-", "-US", " ", ""]) {
      const xml = `<c:chartSpace ${NS}>
  <c:lang val="${bad}"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.lang).toBeUndefined();
    }
  });

  it("ignores a missing val attribute on <c:lang>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:lang/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.lang).toBeUndefined();
  });

  it("surfaces lang alongside other chart-space toggles", () => {
    // <c:lang> sits before <c:roundedCorners> per CT_ChartSpace and
    // <c:style> sits after — co-existing chart-space children should
    // not interfere with each other's parsing.
    const xml = `<c:chartSpace ${NS}>
  <c:lang val="tr-TR"/>
  <c:roundedCorners val="1"/>
  <c:style val="34"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.lang).toBe("tr-TR");
    expect(chart?.roundedCorners).toBe(true);
    expect(chart?.style).toBe(34);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });
});

// ── parseChart — chart date system (date1904) ──────────────────────

describe("parseChart — chart date system", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces <c:date1904 val="1"/> as true', () => {
    // The non-default state — chart date references use the 1904 base
    // (Excel for Mac's legacy epoch). Surfaces verbatim because a
    // chart that pinned the flag carries the override Excel would
    // otherwise inherit from the host workbook.
    const xml = `<c:chartSpace ${NS}>
  <c:date1904 val="1"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.date1904).toBe(true);
  });

  it('surfaces <c:date1904 val="true"/> as true', () => {
    // OOXML accepts the textual `xsd:boolean` spellings.
    const xml = `<c:chartSpace ${NS}>
  <c:date1904 val="true"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.date1904).toBe(true);
  });

  it('collapses <c:date1904 val="0"/> to undefined (OOXML default)', () => {
    // The OOXML default — chart uses the 1900 base. Absence and the
    // default round-trip identically through cloneChart, so the
    // reader collapses the explicit default to undefined for symmetry
    // with every other chart-space toggle.
    const xml = `<c:chartSpace ${NS}>
  <c:date1904 val="0"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.date1904).toBeUndefined();
  });

  it('collapses <c:date1904 val="false"/> to undefined', () => {
    const xml = `<c:chartSpace ${NS}>
  <c:date1904 val="false"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser><c:idx val="0"/></c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.date1904).toBeUndefined();
  });

  it("returns undefined when the chartSpace has no <c:date1904>", () => {
    // Absence is the writer's default — the reader surfaces nothing
    // so a fresh chart and a chart that omits the element round-trip
    // identically through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.date1904).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:date1904>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:date1904/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.date1904).toBeUndefined();
  });

  it("drops unknown val tokens rather than fabricate a flag", () => {
    // <c:date1904> is xsd:boolean per CT_Boolean. Anything outside
    // `1` / `true` / `0` / `false` collapses to undefined so the
    // parsed chart never carries a token Excel itself would not emit.
    for (const bad of ["yes", "no", "2", "T", "F", " ", ""]) {
      const xml = `<c:chartSpace ${NS}>
  <c:date1904 val="${bad}"/>
  <c:chart>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.date1904).toBeUndefined();
    }
  });

  it("surfaces date1904 alongside lang and other chart-space toggles", () => {
    // <c:date1904> sits at the head of the CT_ChartSpace sequence,
    // before <c:lang> and <c:roundedCorners> — co-existing chart-
    // space children should not interfere with each other's parsing.
    const xml = `<c:chartSpace ${NS}>
  <c:date1904 val="1"/>
  <c:lang val="en-US"/>
  <c:roundedCorners val="1"/>
  <c:style val="34"/>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.date1904).toBe(true);
    expect(chart?.lang).toBe("en-US");
    expect(chart?.roundedCorners).toBe(true);
    expect(chart?.style).toBe(34);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });
});

// ── parseChart — axis crossBetween ───────────────────────────────────

describe("parseChart — axis crossBetween", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces a non-default <c:crossBetween val="midCat"/> on the value axis of a column chart', () => {
    // Bar / column / line / area's family default is `"between"`; pinning
    // `"midCat"` is a non-default override and the reader should surface it.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossBetween val="midCat"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossBetween).toBe("midCat");
    expect(chart?.axes?.x?.crossBetween).toBeUndefined();
  });

  it('collapses the family-default <c:crossBetween val="between"/> on a column chart to undefined', () => {
    // Excel always emits `<c:crossBetween>` on every `<c:valAx>` because
    // the element is required by the schema. The reader must collapse the
    // value when it matches the family default so absence and the default
    // round-trip identically through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossBetween val="between"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossBetween).toBeUndefined();
  });

  it('collapses the family-default <c:crossBetween val="midCat"/> on a scatter chart to undefined', () => {
    // Scatter's family default is `"midCat"` because both axes are value
    // axes and Excel emits `<c:crossBetween val="midCat"/>` on each of
    // them. The reader collapses the default so a clone of an untouched
    // scatter chart stays minimal.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:crossBetween val="midCat"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:crossBetween val="midCat"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crossBetween).toBeUndefined();
    expect(chart?.axes?.y?.crossBetween).toBeUndefined();
  });

  it('surfaces a non-default <c:crossBetween val="between"/> on a scatter chart', () => {
    // Scatter's family default is `"midCat"`; pinning `"between"` is a
    // non-default override and the reader surfaces it on whichever axis
    // it was pinned.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:crossBetween val="between"/>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crossBetween).toBe("between");
    expect(chart?.axes?.y?.crossBetween).toBeUndefined();
  });

  it("collapses crossBetween to undefined on a category axis (catAx rejects the element)", () => {
    // The OOXML schema places <c:crossBetween> exclusively on CT_ValAx,
    // so a stray element on <c:catAx> from a corrupt template should
    // never surface — the reader explicitly skips the parse on every
    // non-valAx flavour.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:crossBetween val="midCat"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.crossBetween).toBeUndefined();
  });

  it("drops an unknown ST_CrossBetween token rather than fabricating a value", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossBetween val="diagonal"/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossBetween).toBeUndefined();
  });

  it("drops a missing val attribute rather than fabricating a value", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:crossBetween/>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossBetween).toBeUndefined();
  });

  it("collapses crossBetween to undefined when the chart has no <c:crossBetween>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.crossBetween).toBeUndefined();
  });
});

// ── parseChart — data table ──────────────────────────────────────────

describe("parseChart — data table", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces a fully-defaulted <c:dTable> with every flag true", () => {
    // Excel's reference serialization for a freshly-enabled data table
    // pins all four boolean children to `1`. The reader surfaces every
    // field literally — `<c:dTable>` is required-children-only on
    // CT_DTable so a clone can replay the exact shape the file carried.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="1"/>
      <c:showVertBorder val="1"/>
      <c:showOutline val="1"/>
      <c:showKeys val="1"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.dataTable).toEqual({
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: true,
    });
  });

  it("surfaces non-default false flags literally", () => {
    // Each boolean child round-trips literally — `false` is just as
    // important as `true` because the writer always emits all four
    // children and a clone must preserve the exact pattern.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="0"/>
      <c:showVertBorder val="0"/>
      <c:showOutline val="0"/>
      <c:showKeys val="0"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({
      showHorzBorder: false,
      showVertBorder: false,
      showOutline: false,
      showKeys: false,
    });
  });

  it("surfaces a mixed shape (keys hidden, borders shown)", () => {
    // A common pattern — paint the table grid but hide the legend
    // swatches because the chart already has a separate legend.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="1"/>
      <c:showVertBorder val="1"/>
      <c:showOutline val="1"/>
      <c:showKeys val="0"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: false,
    });
  });

  it("accepts the OOXML textual <xsd:boolean> spellings", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="true"/>
      <c:showVertBorder val="false"/>
      <c:showOutline val="true"/>
      <c:showKeys val="false"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({
      showHorzBorder: true,
      showVertBorder: false,
      showOutline: true,
      showKeys: false,
    });
  });

  it("returns undefined when the plot area has no <c:dTable> element", () => {
    // Absence is the writer's default — Excel renders no data table.
    // The reader surfaces nothing so a fresh chart and a chart that
    // omits the element round-trip identically through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toBeUndefined();
  });

  it("drops a missing val attribute on a <c:dTable> child rather than fabricate a flag", () => {
    // A child without `val` is malformed per CT_Boolean; the reader
    // drops the field rather than fabricate a value the file did not
    // pin. The other children still round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="1"/>
      <c:showVertBorder/>
      <c:showOutline val="1"/>
      <c:showKeys val="1"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({
      showHorzBorder: true,
      showOutline: true,
      showKeys: true,
    });
  });

  it("drops unknown val tokens rather than fabricate flags", () => {
    // Anything outside the OOXML truthy / falsy spellings collapses
    // to undefined for that field. The other children still surface.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="yes"/>
      <c:showVertBorder val="2"/>
      <c:showOutline val="1"/>
      <c:showKeys val="0"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({
      showOutline: true,
      showKeys: false,
    });
  });

  it("surfaces an empty object when <c:dTable> is present but every child is malformed", () => {
    // The element itself is the gating signal — when it appears, the
    // chart is requesting a data table even if every child carries a
    // malformed `val`. The shape stays minimal (an empty object) so a
    // round-trip through the writer falls back to the OOXML defaults
    // (every flag `true`) which Excel would render anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder/>
      <c:showVertBorder/>
      <c:showOutline/>
      <c:showKeys/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({});
  });

  it("surfaces dataTable on a column chart", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="1"/>
      <c:showVertBorder val="1"/>
      <c:showOutline val="1"/>
      <c:showKeys val="1"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: true,
    });
  });

  it("surfaces dataTable on a scatter chart (both axes are valAx)", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:scatterStyle val="lineMarker"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:scatterChart>
    <c:valAx><c:axId val="1"/></c:valAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:dTable>
      <c:showHorzBorder val="0"/>
      <c:showVertBorder val="1"/>
      <c:showOutline val="0"/>
      <c:showKeys val="1"/>
    </c:dTable>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toEqual({
      showHorzBorder: false,
      showVertBorder: true,
      showOutline: false,
      showKeys: true,
    });
  });

  it("returns undefined when the chart has no plotArea", () => {
    // Defensive — a chart with no plotArea has no slot for <c:dTable>
    // either. Surfaces nothing so the parsed shape stays minimal.
    const xml = `<c:chartSpace ${NS}><c:chart></c:chart></c:chartSpace>`;
    expect(parseChart(xml)?.dataTable).toBeUndefined();
  });
});

// ── parseChart — chart-space protection ──────────────────────────────

describe("parseChart — chart-space protection", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces every flag a fully-pinned <c:protection> declares", () => {
    // CT_Protection lists every child as optional, but a "lock
    // everything" preset emits all five. The reader surfaces every
    // pinned flag literally so a clone can replay the exact shape.
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:chartObject val="1"/>
    <c:data val="1"/>
    <c:formatting val="1"/>
    <c:selection val="1"/>
    <c:userInterface val="1"/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.protection).toEqual({
      chartObject: true,
      data: true,
      formatting: true,
      selection: true,
      userInterface: true,
    });
  });

  it("surfaces non-default false flags literally", () => {
    // Each boolean child round-trips literally — `false` is just as
    // important as `true` because the writer always emits all five
    // children and a clone must preserve the exact pattern.
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:chartObject val="0"/>
    <c:data val="0"/>
    <c:formatting val="0"/>
    <c:selection val="0"/>
    <c:userInterface val="0"/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({
      chartObject: false,
      data: false,
      formatting: false,
      selection: false,
      userInterface: false,
    });
  });

  it("surfaces a partial shape with only the flags the file pinned", () => {
    // Common pattern — lock data and selection but leave the rest
    // unpinned. CT_Protection lists every child as optional so the
    // parser surfaces only the present flags rather than fabricate
    // defaults the file did not declare.
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:data val="1"/>
    <c:selection val="1"/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({
      data: true,
      selection: true,
    });
  });

  it("accepts the OOXML textual <xsd:boolean> spellings", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:chartObject val="true"/>
    <c:data val="false"/>
    <c:formatting val="true"/>
    <c:selection val="false"/>
    <c:userInterface val="true"/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({
      chartObject: true,
      data: false,
      formatting: true,
      selection: false,
      userInterface: true,
    });
  });

  it("returns undefined when the chart has no <c:protection> element", () => {
    // Absence is the writer's default — Excel applies no chart-level
    // protection. The reader surfaces nothing so a fresh chart and a
    // chart that omits the element round-trip identically through
    // cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toBeUndefined();
  });

  it("drops a missing val attribute on a <c:protection> child rather than fabricate a flag", () => {
    // A child without `val` is malformed per CT_Boolean; the reader
    // drops the field rather than fabricate a value the file did not
    // pin. The other children still round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:chartObject val="1"/>
    <c:data/>
    <c:formatting val="1"/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({
      chartObject: true,
      formatting: true,
    });
  });

  it("drops unknown val tokens rather than fabricate flags", () => {
    // Anything outside the OOXML truthy / falsy spellings collapses
    // to undefined for that field. The other children still surface.
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:chartObject val="yes"/>
    <c:data val="2"/>
    <c:selection val="1"/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({
      selection: true,
    });
  });

  it("surfaces an empty object when <c:protection> is present but every child is malformed", () => {
    // The element itself is the gating signal — when it appears, the
    // chart is requesting protection even if every child carries a
    // malformed `val`. The shape stays minimal (an empty object) so a
    // round-trip through the writer falls back to the OOXML defaults
    // (every flag `false`) which Excel would apply anyway.
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:chartObject/>
    <c:data/>
    <c:formatting/>
    <c:selection/>
    <c:userInterface/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({});
  });

  it("surfaces an empty object on a bare <c:protection/> element", () => {
    // A self-closing element with no children — same minimal-shape
    // result as a malformed-children block. Round-trips through the
    // writer which falls back to every-flag `false` for emit.
    const xml = `<c:chartSpace ${NS}>
  <c:protection/>
  <c:chart><c:plotArea>
    <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({});
  });

  it("surfaces protection on a pie chart (no axes, but the element lives on chartSpace)", () => {
    // <c:protection> lives on <c:chartSpace>, not inside <c:plotArea>,
    // so axis-shape has no bearing on whether the slot exists. Pie /
    // doughnut still carry the element when the file pins it.
    const xml = `<c:chartSpace ${NS}>
  <c:protection>
    <c:formatting val="1"/>
  </c:protection>
  <c:chart><c:plotArea>
    <c:pieChart>
      <c:varyColors val="1"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.protection).toEqual({
      formatting: true,
    });
  });
});

// ── parseChart — auto title deleted ────────────────────────────────

describe("parseChart — autoTitleDeleted", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  function withAutoTitleDeleted(val?: string): string {
    const el = val === undefined ? "" : `<c:autoTitleDeleted val="${val}"/>`;
    return `<c:chartSpace ${NS}>
  <c:chart>
    ${el}
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
  }

  it('surfaces <c:autoTitleDeleted val="1"/> as true (non-default)', () => {
    // The non-default state — the user explicitly deleted Excel's
    // auto-generated title. Surfaces verbatim because a chart that
    // pinned the flag carries the override Excel would otherwise
    // synthesise from the series name.
    expect(parseChart(withAutoTitleDeleted("1"))?.autoTitleDeleted).toBe(true);
  });

  it('surfaces <c:autoTitleDeleted val="true"/> as true', () => {
    // OOXML accepts the textual `xsd:boolean` spellings.
    expect(parseChart(withAutoTitleDeleted("true"))?.autoTitleDeleted).toBe(true);
  });

  it('collapses <c:autoTitleDeleted val="0"/> to undefined (OOXML default)', () => {
    // The OOXML default — the auto-title is not suppressed. Absence
    // and the default round-trip identically through cloneChart, so the
    // reader collapses the explicit default to undefined for symmetry
    // with every other chart-level toggle.
    expect(parseChart(withAutoTitleDeleted("0"))?.autoTitleDeleted).toBeUndefined();
  });

  it('collapses <c:autoTitleDeleted val="false"/> to undefined', () => {
    expect(parseChart(withAutoTitleDeleted("false"))?.autoTitleDeleted).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:autoTitleDeleted> element", () => {
    // Absence is identical to the OOXML default; the reader surfaces
    // nothing so a fresh chart and a chart that omits the element
    // round-trip identically through cloneChart.
    expect(parseChart(withAutoTitleDeleted())?.autoTitleDeleted).toBeUndefined();
  });

  it("ignores a missing val attribute on <c:autoTitleDeleted>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:autoTitleDeleted/>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.autoTitleDeleted).toBeUndefined();
  });

  it("drops unknown val tokens rather than fabricate a flag", () => {
    expect(parseChart(withAutoTitleDeleted("bogus"))?.autoTitleDeleted).toBeUndefined();
  });

  it("surfaces autoTitleDeleted independently of the title presence (titleless chart)", () => {
    // The element sits on <c:chart> directly, not nested inside
    // <c:title>, so a chart with no <c:title> can still pin the flag
    // to suppress Excel's series-name auto-title synthesis.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.title).toBeUndefined();
    expect(chart?.autoTitleDeleted).toBe(true);
  });

  it("surfaces autoTitleDeleted alongside a literal title (chart with both)", () => {
    // A chart can emit both a literal <c:title> and pin
    // <c:autoTitleDeleted val="1"/> — the flag suppresses any future
    // auto-synthesis even though the literal already overrides it. The
    // parser surfaces both fields independently.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Sales</a:t></a:r></a:p>
      </c:rich></c:tx>
    </c:title>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.title).toBe("Sales");
    expect(chart?.autoTitleDeleted).toBe(true);
  });

  it("co-exists with other chart-level toggles", () => {
    // The flag should not interfere with sibling chart-level fields
    // parsed off <c:chart> / <c:chartSpace>.
    const xml = `<c:chartSpace ${NS}>
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.autoTitleDeleted).toBe(true);
    expect(chart?.roundedCorners).toBe(true);
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
    expect(chart?.varyColors).toBe(true);
  });

  it("surfaces the flag on every chart family", () => {
    // The element sits on <c:chart>, not inside any chart-type
    // element, so it round-trips identically across families.
    for (const kind of ["lineChart", "barChart", "pieChart", "doughnutChart", "areaChart"]) {
      const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:${kind}><c:ser><c:idx val="0"/></c:ser></c:${kind}>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
      expect(parseChart(xml)?.autoTitleDeleted).toBe(true);
    }
  });
});

// ── parseChart — chart-level line marker visibility ────────────────

describe("parseChart — showLineMarkers", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it('surfaces showLineMarkers=false when the line chart pins <c:marker val="0"/>', () => {
    // The non-default state — flips the chart-level gate off so
    // per-series marker definitions stop rendering chart-wide.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker val="0"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBe(false);
  });

  it('collapses <c:marker val="1"/> (the Excel / OOXML default) to undefined', () => {
    // Excel's reference serialization for every authored line chart
    // emits <c:marker val="1"/>. Surfacing only the non-default value
    // keeps the parsed shape minimal — a fresh chart and a marker-on
    // chart round-trip identically through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker val="1"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBeUndefined();
  });

  it("collapses absence of <c:marker> to undefined on a line chart", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBeUndefined();
  });

  it('accepts the OOXML truthy / falsy spellings ("true" / "false")', () => {
    const xmlFalse = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker val="false"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xmlFalse)?.showLineMarkers).toBe(false);

    const xmlTrue = xmlFalse.replace('val="false"', 'val="true"');
    expect(parseChart(xmlTrue)?.showLineMarkers).toBeUndefined();
  });

  it("drops a missing val attribute (CT_Boolean default would be true)", () => {
    // A bare <c:marker/> carries no `val`; per CT_Boolean the schema
    // default would be true, but the reader collapses ambiguous shapes
    // to undefined rather than fabricate the default state — the
    // writer's always-emit contract means a real chart never emits a
    // bare element.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBeUndefined();
  });

  it("drops an unknown val token rather than fabricating a value", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker val="maybe"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBeUndefined();
  });

  it("ignores chart-level <c:marker> on a 3D line chart (CT_Line3DChart has no slot)", () => {
    // The OOXML schema places the chart-level <c:marker> (CT_Boolean)
    // exclusively on CT_LineChart — CT_Line3DChart has no slot. A
    // stray element on a 3D line chart-type body should not surface
    // through `showLineMarkers`.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:line3DChart>
      <c:grouping val="standard"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker val="0"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
      <c:axId val="3"/>
    </c:line3DChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
    <c:serAx><c:axId val="3"/></c:serAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBeUndefined();
  });

  it("ignores chart-level <c:marker> on a stock chart (CT_StockChart has no slot)", () => {
    // CT_StockChart has hiLowLines / upDownBars but no chart-level
    // marker per the OOXML schema. A stray element should not surface.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:stockChart>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker val="0"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:stockChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBeUndefined();
  });

  it("ignores <c:marker> on bar / column / pie / doughnut / area / scatter charts", () => {
    // The chart-level <c:marker> (CT_Boolean) only lives on
    // CT_LineChart. The reader scopes the lookup to the matching kind
    // — every other family falls through. Note that scatter has its
    // own per-series <c:marker> (CT_Marker) handling, but no
    // chart-level CT_Boolean variant.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:marker val="0"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.showLineMarkers).toBeUndefined();
  });

  it("does not confuse the chart-level <c:marker> with the per-series <c:marker> block", () => {
    // The per-series <c:marker> sits inside <c:ser> and carries
    // CT_Marker children (<c:symbol>, <c:size>, ...). The chart-level
    // gate sits as a sibling of <c:ser> and carries CT_Boolean's `val`
    // attribute. The reader must not surface a per-series marker as
    // `showLineMarkers`.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser>
        <c:idx val="0"/>
        <c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>
      </c:ser>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    // No chart-level <c:marker> is present — only the per-series
    // block — so showLineMarkers must be undefined.
    expect(chart?.showLineMarkers).toBeUndefined();
    // Per-series marker still surfaces on the series side.
    expect(chart?.series?.[0].marker).toMatchObject({ symbol: "circle", size: 6 });
  });

  it("co-surfaces showLineMarkers alongside other line-only chart-level fields", () => {
    // The chart-level <c:marker> sits at the tail of CT_LineChart
    // (after dropLines / hiLowLines / upDownBars, before axId+).
    // Composing every line-only block on the same chart should not
    // disturb the surfaced marker flag.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:varyColors val="0"/>
      <c:ser><c:idx val="0"/></c:ser>
      <c:dropLines/>
      <c:hiLowLines/>
      <c:upDownBars><c:gapWidth val="150"/></c:upDownBars>
      <c:marker val="0"/>
      <c:axId val="1"/>
      <c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.showLineMarkers).toBe(false);
    expect(chart?.dropLines).toBe(true);
    expect(chart?.hiLowLines).toBe(true);
    expect(chart?.upDownBars).toBe(true);
  });
});

// ── parseChart — view3D ────────────────────────────────────────────

describe("parseChart — view3D", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

  it("surfaces every CT_View3D field a fully-pinned <c:view3D> declares", () => {
    // Excel's "3-D Rotation" pane writes every field on a fresh 3D
    // chart. The reader surfaces every pinned value literally so a
    // clone can replay the exact rotation / perspective shape.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="15"/>
      <c:hPercent val="100"/>
      <c:rotY val="20"/>
      <c:depthPercent val="100"/>
      <c:rAngAx val="1"/>
      <c:perspective val="30"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.view3D).toEqual({
      rotX: 15,
      hPercent: 100,
      rotY: 20,
      depthPercent: 100,
      rAngAx: true,
      perspective: 30,
    });
  });

  it("surfaces a partial shape with only the fields the file pinned", () => {
    // Common pattern — pin rotation only, leave height / depth /
    // perspective at the per-family default. CT_View3D lists every
    // child as optional so the parser surfaces only the present
    // fields rather than fabricate defaults the file did not declare.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="20"/>
      <c:rotY val="40"/>
    </c:view3D>
    <c:plotArea>
      <c:line3DChart><c:ser><c:idx val="0"/></c:ser></c:line3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({ rotX: 20, rotY: 40 });
  });

  it("surfaces signed rotX values (the OOXML ST_RotX type accepts -90..90)", () => {
    // ST_RotX is a signed byte — Excel writes negative values for
    // back-tilts. The parser must accept the leading `-` so a tilted
    // template round-trips.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="-30"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({ rotX: -30 });
  });

  it("surfaces the boundary values of every range (min and max)", () => {
    // Verify the inclusive bounds of every child's simple type.
    const xmlMin = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="-90"/>
      <c:hPercent val="5"/>
      <c:rotY val="0"/>
      <c:depthPercent val="20"/>
      <c:perspective val="0"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xmlMin)?.view3D).toEqual({
      rotX: -90,
      hPercent: 5,
      rotY: 0,
      depthPercent: 20,
      perspective: 0,
    });

    const xmlMax = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="90"/>
      <c:hPercent val="500"/>
      <c:rotY val="360"/>
      <c:depthPercent val="2000"/>
      <c:perspective val="240"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xmlMax)?.view3D).toEqual({
      rotX: 90,
      hPercent: 500,
      rotY: 360,
      depthPercent: 2000,
      perspective: 240,
    });
  });

  it("drops out-of-range numeric fields rather than fabricate clamped values", () => {
    // Each numeric field is bound by an OOXML simple type (ST_RotX,
    // ST_HPercent, ...). Out-of-range values drop silently rather
    // than silently clamp — Excel itself would reject the token.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="180"/>
      <c:hPercent val="3"/>
      <c:rotY val="-10"/>
      <c:depthPercent val="10"/>
      <c:perspective val="500"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    // Every field is out of its simple-type range — nothing surfaces,
    // but the empty `<c:view3D>` shell still gates the field as `{}`.
    expect(parseChart(xml)?.view3D).toEqual({});
  });

  it("drops fractional / non-integer values rather than fabricate floats", () => {
    // Every CT_View3D numeric child is an integer simple type.
    // `parseInt` would coerce "15.5" → 15, but Excel never emits the
    // fractional spelling. Drop the field so a hand-edited file with
    // a fractional `val` stays unrecognised rather than silently
    // round-trip a value Excel would not author.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="15.5"/>
      <c:hPercent val="100abc"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({});
  });

  it("accepts the OOXML textual <xsd:boolean> spellings on rAngAx", () => {
    // CT_Boolean accepts "1" / "true" / "0" / "false".
    const xmlTrue = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D><c:rAngAx val="true"/></c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xmlTrue)?.view3D).toEqual({ rAngAx: true });

    const xmlFalse = xmlTrue.replace('val="true"', 'val="false"');
    expect(parseChart(xmlFalse)?.view3D).toEqual({ rAngAx: false });
  });

  it("drops a missing val attribute on a <c:view3D> child rather than fabricate a value", () => {
    // A child without `val` is malformed per its CT type; the reader
    // drops the field rather than fabricate a value the file did not
    // pin. The other children still round-trip.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX/>
      <c:rotY val="20"/>
      <c:rAngAx/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({ rotY: 20 });
  });

  it("drops unknown rAngAx tokens rather than fabricate flags", () => {
    // Anything outside the OOXML truthy / falsy spellings collapses
    // to undefined for the rAngAx field. Other fields still surface.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="15"/>
      <c:rAngAx val="yes"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({ rotX: 15 });
  });

  it("surfaces an empty object when <c:view3D> is present but every child is malformed", () => {
    // The element itself is the gating signal — when it appears, the
    // chart is requesting a 3D view even if every child carries a
    // malformed `val`. The shape stays minimal (an empty object) so
    // a round-trip through the writer falls back to absence.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX/>
      <c:hPercent/>
      <c:rotY/>
      <c:depthPercent/>
      <c:rAngAx/>
      <c:perspective/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({});
  });

  it("surfaces an empty object on a bare <c:view3D/> element", () => {
    // A self-closing element with no children — same minimal-shape
    // result as a malformed-children block.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D/>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({});
  });

  it("returns undefined when the chart has no <c:view3D> element", () => {
    // Absence is the writer's default — Excel falls back to the
    // per-family default rotation / perspective. The reader surfaces
    // nothing so a fresh chart and a chart that omits the element
    // round-trip identically through cloneChart.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:ser><c:idx val="0"/></c:ser>
    </c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toBeUndefined();
  });

  it("surfaces view3D on a 2D chart family (the OOXML schema accepts it on every CT_Chart)", () => {
    // <c:view3D> is only meaningful on 3D families but the schema
    // accepts it on every CT_Chart. A stray element on a 2D chart
    // surfaces here so the round-trip through cloneChart stays
    // lossless even when the template authors a no-op.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotX val="15"/>
    </c:view3D>
    <c:plotArea>
      <c:lineChart><c:ser><c:idx val="0"/></c:ser></c:lineChart>
      <c:catAx><c:axId val="1"/></c:catAx>
      <c:valAx><c:axId val="2"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({ rotX: 15 });
  });

  it("surfaces view3D on a pie chart (no axes, but the element lives on <c:chart>)", () => {
    // <c:view3D> lives on <c:chart>, not inside <c:plotArea>, so
    // axis-shape has no bearing on whether the slot exists. Pie /
    // doughnut still carry the element when the file pins it.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:view3D>
      <c:rotY val="180"/>
    </c:view3D>
    <c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.view3D).toEqual({ rotY: 180 });
  });

  it("co-exists with sibling chart-level toggles", () => {
    // The view3D reader should not interfere with sibling fields
    // parsed off <c:chart> / <c:chartSpace>.
    const xml = `<c:chartSpace ${NS}>
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:view3D>
      <c:rotX val="20"/>
      <c:rotY val="30"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="zero"/>
  </c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.autoTitleDeleted).toBe(true);
    expect(chart?.view3D).toEqual({ rotX: 20, rotY: 30 });
    expect(chart?.plotVisOnly).toBe(false);
    expect(chart?.dispBlanksAs).toBe("zero");
  });
});

// ── parseChart — legend entries ────────────────────────────────────

describe("parseChart — legend entries", () => {
  function chartWithLegend(legendXml: string): string {
    return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/></c:ser>
        <c:ser><c:idx val="1"/></c:ser>
        <c:ser><c:idx val="2"/></c:ser>
      </c:barChart>
    </c:plotArea>
    ${legendXml}
  </c:chart>
</c:chartSpace>`;
  }

  it("surfaces a single hidden entry", () => {
    const xml = chartWithLegend(
      '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry></c:legend>',
    );
    expect(parseChart(xml)?.legendEntries).toEqual([{ idx: 1, delete: true }]);
  });

  it("surfaces multiple entries in <c:idx> declaration order", () => {
    const xml = chartWithLegend(
      `<c:legend>
        <c:legendPos val="r"/>
        <c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry>
        <c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>
      </c:legend>`,
    );
    // The reader preserves the source order — the writer reorders by
    // ascending idx on emit, but parseChart surfaces the file-order list
    // so a roundtrip can be observed without normalization.
    expect(parseChart(xml)?.legendEntries).toEqual([
      { idx: 2, delete: true },
      { idx: 0, delete: true },
    ]);
  });

  it("treats a missing <c:delete> as delete=false (the OOXML default)", () => {
    // CT_LegendEntry's <c:delete> is optional. Some templates (and
    // older Excel versions) emit a bare <c:legendEntry><c:idx/></c:legendEntry>
    // with the entry left visible; the reader still surfaces the index
    // override with `delete: false` so a clone-through carries the
    // selector forward.
    const xml = chartWithLegend(
      '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="1"/></c:legendEntry></c:legend>',
    );
    expect(parseChart(xml)?.legendEntries).toEqual([{ idx: 1, delete: false }]);
  });

  it('parses <c:delete val="0"/> as delete=false', () => {
    const xml = chartWithLegend(
      '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="0"/><c:delete val="0"/></c:legendEntry></c:legend>',
    );
    expect(parseChart(xml)?.legendEntries).toEqual([{ idx: 0, delete: false }]);
  });

  it('accepts the truthy / falsy <c:delete> spellings ("true" / "false")', () => {
    const xml = chartWithLegend(
      `<c:legend>
        <c:legendPos val="r"/>
        <c:legendEntry><c:idx val="0"/><c:delete val="true"/></c:legendEntry>
        <c:legendEntry><c:idx val="1"/><c:delete val="false"/></c:legendEntry>
      </c:legend>`,
    );
    expect(parseChart(xml)?.legendEntries).toEqual([
      { idx: 0, delete: true },
      { idx: 1, delete: false },
    ]);
  });

  it("returns undefined when the chart declares no <c:legendEntry>", () => {
    const xml = chartWithLegend('<c:legend><c:legendPos val="r"/></c:legend>');
    expect(parseChart(xml)?.legendEntries).toBeUndefined();
  });

  it("returns undefined when the chart hides the legend (delete=1)", () => {
    // A hidden legend has no slot for entry overrides — even a stray
    // `<c:legendEntry>` inside it must not surface (the rendered chart
    // would never show those entries anyway).
    const xml = chartWithLegend(
      '<c:legend><c:delete val="1"/><c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry></c:legend>',
    );
    expect(parseChart(xml)?.legendEntries).toBeUndefined();
  });

  it("returns undefined when the chart has no <c:legend> at all", () => {
    const xml = chartWithLegend("");
    expect(parseChart(xml)?.legendEntries).toBeUndefined();
  });

  it("drops entries whose <c:idx> is missing", () => {
    const xml = chartWithLegend(
      `<c:legend>
        <c:legendPos val="r"/>
        <c:legendEntry><c:delete val="1"/></c:legendEntry>
        <c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry>
      </c:legend>`,
    );
    expect(parseChart(xml)?.legendEntries).toEqual([{ idx: 2, delete: true }]);
  });

  it("drops entries whose <c:idx val=..> is malformed", () => {
    const xml = chartWithLegend(
      `<c:legend>
        <c:legendPos val="r"/>
        <c:legendEntry><c:idx val="abc"/><c:delete val="1"/></c:legendEntry>
        <c:legendEntry><c:idx val="-1"/><c:delete val="1"/></c:legendEntry>
        <c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>
      </c:legend>`,
    );
    expect(parseChart(xml)?.legendEntries).toEqual([{ idx: 0, delete: true }]);
  });

  it("deduplicates duplicate <c:idx> entries (first wins)", () => {
    const xml = chartWithLegend(
      `<c:legend>
        <c:legendPos val="r"/>
        <c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>
        <c:legendEntry><c:idx val="1"/><c:delete val="0"/></c:legendEntry>
      </c:legend>`,
    );
    expect(parseChart(xml)?.legendEntries).toEqual([{ idx: 1, delete: true }]);
  });
});

// ── parseChart — axis label rotation ───────────────────────────────

describe("parseChart — axis labelRotation", () => {
  const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

  function withCatAxTxPr(rot: string | undefined): string {
    const txPr =
      rot === undefined
        ? ""
        : `<c:txPr><a:bodyPr rot="${rot}"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
    return `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      ${txPr}
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
  }

  it("surfaces the rotation in whole degrees from the rot attribute (60000ths)", () => {
    // 45 degrees * 60000 = 2,700,000.
    const chart = parseChart(withCatAxTxPr("2700000"));
    expect(chart?.axes?.x?.labelRotation).toBe(45);
  });

  it("surfaces negative rotations literally", () => {
    const chart = parseChart(withCatAxTxPr("-2700000"));
    expect(chart?.axes?.x?.labelRotation).toBe(-45);
  });

  it('collapses the OOXML default rot="0" to undefined', () => {
    // The default `0` round-trips identically to absence — both leave
    // the labels rendering flat.
    const chart = parseChart(withCatAxTxPr("0"));
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined when <c:txPr> is absent", () => {
    const chart = parseChart(withCatAxTxPr(undefined));
    expect(chart?.axes).toBeUndefined();
  });

  it("returns undefined when <a:bodyPr> is absent inside <c:txPr>", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:txPr><a:lstStyle/><a:p/></c:txPr>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.axes).toBeUndefined();
  });

  it("returns undefined when <a:bodyPr> omits the rot attribute", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    expect(parseChart(xml)?.axes).toBeUndefined();
  });

  it("clamps rot values above the 90-degree maximum to 90", () => {
    // Values outside Excel's UI band collapse to the nearest endpoint
    // so a corrupt template cannot surface a rotation the writer would
    // never emit. 180° in 60000ths = 10,800,000.
    const chart = parseChart(withCatAxTxPr("10800000"));
    expect(chart?.axes?.x?.labelRotation).toBe(90);
  });

  it("clamps rot values below the -90-degree minimum to -90", () => {
    const chart = parseChart(withCatAxTxPr("-10800000"));
    expect(chart?.axes?.x?.labelRotation).toBe(-90);
  });

  it("drops non-numeric rot tokens", () => {
    expect(parseChart(withCatAxTxPr("forty-five"))?.axes).toBeUndefined();
  });

  it("rounds non-integer 60000ths to the nearest whole degree", () => {
    // 2,700,030 ≈ 45.0005°, rounds to 45.
    const chart = parseChart(withCatAxTxPr("2700030"));
    expect(chart?.axes?.x?.labelRotation).toBe(45);
  });

  it("surfaces the rotation on the value axis", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx><c:axId val="1"/></c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:txPr><a:bodyPr rot="-1800000"/><a:lstStyle/><a:p/></c:txPr>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.y?.labelRotation).toBe(-30);
  });

  it("surfaces independently on both axes", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:txPr><a:bodyPr rot="2700000"/><a:lstStyle/><a:p/></c:txPr>
    </c:catAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:txPr><a:bodyPr rot="-1800000"/><a:lstStyle/><a:p/></c:txPr>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.labelRotation).toBe(45);
    expect(chart?.axes?.y?.labelRotation).toBe(-30);
  });

  it("surfaces the rotation on a scatter chart's value axes", () => {
    // Scatter has two `<c:valAx>` siblings — the rotation surfaces on
    // both axes for symmetry with the writer-side scatter builder.
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:scatterChart><c:ser><c:idx val="0"/></c:ser></c:scatterChart>
    <c:valAx>
      <c:axId val="1"/>
      <c:axPos val="b"/>
      <c:txPr><a:bodyPr rot="2700000"/><a:lstStyle/><a:p/></c:txPr>
    </c:valAx>
    <c:valAx>
      <c:axId val="2"/>
      <c:axPos val="l"/>
      <c:txPr><a:bodyPr rot="-2700000"/><a:lstStyle/><a:p/></c:txPr>
    </c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x?.labelRotation).toBe(45);
    expect(chart?.axes?.y?.labelRotation).toBe(-45);
  });

  it("co-surfaces alongside other axis fields", () => {
    const xml = `<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:barChart><c:ser><c:idx val="0"/></c:ser></c:barChart>
    <c:catAx>
      <c:axId val="1"/>
      <c:title><c:tx><c:rich><a:p><a:r><a:t>Period</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:tickLblPos val="low"/>
      <c:txPr><a:bodyPr rot="2700000"/><a:lstStyle/><a:p/></c:txPr>
      <c:noMultiLvlLbl val="1"/>
    </c:catAx>
    <c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = parseChart(xml);
    expect(chart?.axes?.x).toEqual({
      title: "Period",
      tickLblPos: "low",
      labelRotation: 45,
      noMultiLvlLbl: true,
    });
  });
});
