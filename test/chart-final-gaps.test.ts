import { describe, expect, it } from "vitest";
import { parseXml } from "../src/xml/parser";
import { writeChart } from "../src/xlsx/chart-writer";
import { parseChart } from "../src/xlsx/chart-reader";
import { cloneChart } from "../src/xlsx/chart-clone";
import { writeXlsx } from "../src/xlsx/writer";
import { ZipReader } from "../src/zip/reader";
import {
  buildDataPoints,
  buildAllErrorBars,
  buildErrorBars,
  buildShape3D,
  buildTrendline,
  buildTrendlines,
  cloneDataPoint,
  cloneErrorBars,
  cloneTrendline,
  normalizeShape3D,
  parseDataPoints,
  parseErrorBars,
  parseShape3D,
  parseTrendlines,
  parseBubbleSizeRef,
  resolveDataPoints,
  resolveErrorBars,
  resolveTrendlines,
  VALID_ERR_DIRECTIONS,
  VALID_ERR_TYPES,
  VALID_ERR_VAL_TYPES,
  VALID_SHAPE_3D,
  VALID_TRENDLINE_TYPES,
} from "../src/xlsx/chart/seriesExtras";
import type {
  Chart,
  ChartDataPoint,
  ChartErrorBars,
  ChartShape3D,
  ChartTrendline,
  SheetChart,
} from "../src/_types";

const decoder = new TextDecoder("utf-8");

// ── helpers ──────────────────────────────────────────────────────────

function makeChart(overrides: Partial<SheetChart> = {}): SheetChart {
  return {
    type: "column",
    title: "Chart",
    series: [{ name: "A", values: "B2:B4", categories: "A2:A4" }],
    anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
    ...overrides,
  };
}

function findChild(el: any, localName: string): any {
  if (!el || !el.children) return undefined;
  return el.children.find((c: any) => typeof c !== "string" && c.local === localName);
}

function findAll(el: any, localName: string): any[] {
  if (!el || !el.children) return [];
  return el.children.filter((c: any) => typeof c !== "string" && c.local === localName);
}

function deepFind(el: any, path: string[]): any {
  let cur = el;
  for (const name of path) {
    cur = findChild(cur, name);
    if (!cur) return undefined;
  }
  return cur;
}

function parseXmlString(s: string) {
  return parseXml(s);
}

// ── 1. Trendlines ────────────────────────────────────────────────────

describe("Trendlines: types & validation", () => {
  it("VALID_TRENDLINE_TYPES exports the OOXML enum", () => {
    expect(VALID_TRENDLINE_TYPES.has("linear")).toBe(true);
    expect(VALID_TRENDLINE_TYPES.has("log")).toBe(true);
    expect(VALID_TRENDLINE_TYPES.has("exp")).toBe(true);
    expect(VALID_TRENDLINE_TYPES.has("power")).toBe(true);
    expect(VALID_TRENDLINE_TYPES.has("poly")).toBe(true);
    expect(VALID_TRENDLINE_TYPES.has("movingAvg")).toBe(true);
    expect(VALID_TRENDLINE_TYPES.size).toBe(6);
  });
});

describe("Trendlines: writer", () => {
  it("emits a minimal linear trendline", () => {
    const xml = buildTrendline({ type: "linear" });
    expect(xml).toBeDefined();
    expect(xml).toContain('c:trendlineType val="linear"');
  });

  it("emits the schema-required <c:trendlineType> child", () => {
    const xml = buildTrendline({ type: "movingAvg", period: 3 });
    expect(xml).toContain('c:trendlineType val="movingAvg"');
    expect(xml).toContain('c:period val="3"');
  });

  it("clamps poly order to 2..6", () => {
    const xml = buildTrendline({ type: "poly", order: 9 });
    expect(xml).toContain('c:order val="6"');
    const xml2 = buildTrendline({ type: "poly", order: -1 });
    expect(xml2).toContain('c:order val="2"');
  });

  it("clamps movingAvg period to 2..100", () => {
    const xml = buildTrendline({ type: "movingAvg", period: 9999 });
    expect(xml).toContain('c:period val="100"');
    const xml2 = buildTrendline({ type: "movingAvg", period: 0 });
    expect(xml2).toContain('c:period val="2"');
  });

  it("only emits order on poly type", () => {
    const xml = buildTrendline({ type: "linear", order: 3 });
    expect(xml).not.toContain("c:order");
  });

  it("only emits period on movingAvg type", () => {
    const xml = buildTrendline({ type: "linear", period: 3 });
    expect(xml).not.toContain("c:period");
  });

  it("emits forecast forward / backward as CT_Double", () => {
    const xml = buildTrendline({ type: "linear", forward: 2.5, backward: 1 });
    expect(xml).toContain('c:forward val="2.5"');
    expect(xml).toContain('c:backward val="1"');
  });

  it("emits intercept", () => {
    const xml = buildTrendline({ type: "linear", intercept: 0 });
    expect(xml).toContain('c:intercept val="0"');
  });

  it("emits dispEq / dispRSqr only when true", () => {
    const xml = buildTrendline({ type: "linear", dispEquation: true, dispRSquared: true });
    expect(xml).toContain('c:dispEq val="1"');
    expect(xml).toContain('c:dispRSqr val="1"');

    const off = buildTrendline({ type: "linear", dispEquation: false });
    expect(off).not.toContain("c:dispEq");
  });

  it("emits stroke color/width/dash inside spPr", () => {
    const xml = buildTrendline({
      type: "linear",
      lineColor: "FF0000",
      lineWidth: 2,
      lineDash: "dash",
    });
    expect(xml).toContain("c:spPr");
    expect(xml).toContain('a:srgbClr val="FF0000"');
    expect(xml).toContain('a:prstDash val="dash"');
    expect(xml).toMatch(/a:ln[^>]*w="\d+/);
  });

  it("emits a trendline name", () => {
    const xml = buildTrendline({ type: "linear", name: "My Trend" });
    expect(xml).toContain("<c:name>My Trend</c:name>");
  });

  it("escapes name xml special characters", () => {
    const xml = buildTrendline({ type: "linear", name: "A & B < C > D" });
    expect(xml).toContain("A &amp; B &lt; C &gt; D");
  });

  it("drops trendlines with invalid type on emit", () => {
    const xml = buildTrendline({ type: "bogus" as any });
    expect(xml).toBeUndefined();
  });

  it("buildTrendlines emits multiple in declaration order", () => {
    const arr = buildTrendlines([{ type: "linear" }, { type: "movingAvg", period: 3 }]);
    expect(arr).toHaveLength(2);
    expect(arr[0]).toContain("linear");
    expect(arr[1]).toContain("movingAvg");
  });

  it("buildTrendlines drops invalid entries silently", () => {
    const arr = buildTrendlines([{ type: "linear" }, { type: "bogus" as any }]);
    expect(arr).toHaveLength(1);
  });

  it("buildTrendlines returns empty for undefined", () => {
    expect(buildTrendlines(undefined)).toEqual([]);
    expect(buildTrendlines([])).toEqual([]);
  });
});

describe("Trendlines: reader", () => {
  it("parses a minimal trendline", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="linear"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res).toEqual([{ type: "linear" }]);
  });

  it("parses moving average period", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="movingAvg"/><c:period val="3"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res?.[0].period).toBe(3);
  });

  it("parses poly order", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="poly"/><c:order val="3"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res?.[0].order).toBe(3);
  });

  it("drops out-of-range order", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="poly"/><c:order val="99"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res?.[0].order).toBeUndefined();
  });

  it("drops trendline with unknown type", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="bogus"/></c:trendline></c:ser>`,
    );
    expect(parseTrendlines(ser)).toBeUndefined();
  });

  it("drops trendline without trendlineType", () => {
    const ser = parseXmlString(`<c:ser xmlns:c="x"><c:trendline/></c:ser>`);
    expect(parseTrendlines(ser)).toBeUndefined();
  });

  it("parses dispEq and dispRSqr", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="linear"/><c:dispEq val="1"/><c:dispRSqr val="1"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res?.[0].dispEquation).toBe(true);
    expect(res?.[0].dispRSquared).toBe(true);
  });

  it("parses forward / backward / intercept", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="linear"/><c:forward val="2.5"/><c:backward val="1"/><c:intercept val="0"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res?.[0].forward).toBe(2.5);
    expect(res?.[0].backward).toBe(1);
    expect(res?.[0].intercept).toBe(0);
  });

  it("parses name", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:name>Custom Name</c:name><c:trendlineType val="linear"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res?.[0].name).toBe("Custom Name");
  });

  it("parses stroke color/width/dash", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:spPr xmlns:a="x"><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr><c:trendlineType val="linear"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res?.[0].lineColor).toBe("FF0000");
    expect(res?.[0].lineWidth).toBe(2);
    expect(res?.[0].lineDash).toBe("dash");
  });

  it("parses multiple trendlines preserving order", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:trendline><c:trendlineType val="linear"/></c:trendline><c:trendline><c:trendlineType val="movingAvg"/><c:period val="5"/></c:trendline></c:ser>`,
    );
    const res = parseTrendlines(ser);
    expect(res).toHaveLength(2);
    expect(res?.[0].type).toBe("linear");
    expect(res?.[1].type).toBe("movingAvg");
  });
});

describe("Trendlines: clone-through", () => {
  it("inherit (undefined) keeps the source's trendlines", () => {
    const r = resolveTrendlines([{ type: "linear" }], undefined);
    expect(r).toEqual([{ type: "linear" }]);
  });

  it("null drops the inherited array", () => {
    const r = resolveTrendlines([{ type: "linear" }], null);
    expect(r).toBeUndefined();
  });

  it("array replaces", () => {
    const r = resolveTrendlines([{ type: "linear" }], [{ type: "exp" }]);
    expect(r).toEqual([{ type: "exp" }]);
  });

  it("empty array collapses to undefined", () => {
    expect(resolveTrendlines(undefined, [])).toBeUndefined();
  });

  it("clone copies all defined fields", () => {
    const c = cloneTrendline({
      type: "poly",
      name: "X",
      order: 4,
      forward: 1,
      backward: 2,
      intercept: 0,
      dispEquation: true,
      dispRSquared: true,
      lineColor: "FF0000",
      lineWidth: 2,
      lineDash: "dot",
    });
    expect(c).toMatchObject({
      type: "poly",
      name: "X",
      order: 4,
      forward: 1,
      backward: 2,
      intercept: 0,
      dispEquation: true,
      dispRSquared: true,
      lineColor: "FF0000",
      lineWidth: 2,
      lineDash: "dot",
    });
  });

  it("clone drops invalid type", () => {
    expect(cloneTrendline({ type: "bogus" as any })).toBeUndefined();
  });
});

describe("Trendlines: end-to-end via writeChart", () => {
  it("emits c:trendline blocks on bar series", () => {
    const result = writeChart(
      makeChart({
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            trendlines: [{ type: "linear" }],
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:trendline");
    expect(result.chartXml).toContain('c:trendlineType val="linear"');
  });

  it("does not emit c:trendline on pie series", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            trendlines: [{ type: "linear" }],
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:trendline");
  });

  it("round-trips trendline through writeChart → parseChart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            trendlines: [{ type: "movingAvg", period: 3 }],
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    expect(parsed?.series?.[0].trendlines).toEqual([{ type: "movingAvg", period: 3 }]);
  });
});

// ── 2. Error bars ────────────────────────────────────────────────────

describe("ErrorBars: types & validation", () => {
  it("VALID_ERR_DIRECTIONS / VALID_ERR_TYPES / VALID_ERR_VAL_TYPES export the OOXML enums", () => {
    expect(VALID_ERR_DIRECTIONS.has("x")).toBe(true);
    expect(VALID_ERR_DIRECTIONS.has("y")).toBe(true);
    expect(VALID_ERR_DIRECTIONS.size).toBe(2);
    expect(VALID_ERR_TYPES.size).toBe(3);
    expect(VALID_ERR_VAL_TYPES.size).toBe(5);
  });
});

describe("ErrorBars: writer", () => {
  it("emits required errDir / errBarType / errValType in schema order", () => {
    const xml = buildErrorBars({ direction: "y", type: "both", valType: "fixedVal", value: 5 });
    expect(xml).toContain('c:errDir val="y"');
    expect(xml).toContain('c:errBarType val="both"');
    expect(xml).toContain('c:errValType val="fixedVal"');
    expect(xml).toContain('c:val val="5"');
    // schema order
    const dirIdx = xml!.indexOf("errDir");
    const typeIdx = xml!.indexOf("errBarType");
    const vtIdx = xml!.indexOf("errValType");
    const valIdx = xml!.indexOf("c:val ");
    expect(dirIdx).toBeLessThan(typeIdx);
    expect(typeIdx).toBeLessThan(vtIdx);
    expect(vtIdx).toBeLessThan(valIdx);
  });

  it("emits noEndCap when true", () => {
    const xml = buildErrorBars({
      direction: "y",
      type: "both",
      valType: "fixedVal",
      value: 1,
      noEndCap: true,
    });
    expect(xml).toContain('c:noEndCap val="1"');
  });

  it("does not emit val for stdErr", () => {
    const xml = buildErrorBars({
      direction: "y",
      type: "both",
      valType: "stdErr",
      value: 5,
    });
    expect(xml).not.toContain("c:val val=");
  });

  it("does not emit val for cust", () => {
    const xml = buildErrorBars({
      direction: "y",
      type: "both",
      valType: "cust",
    });
    expect(xml).not.toContain("c:val val=");
  });

  it("emits stroke spPr block", () => {
    const xml = buildErrorBars({
      direction: "y",
      type: "both",
      valType: "fixedVal",
      value: 1,
      lineColor: "00FF00",
      lineWidth: 1,
      lineDash: "dash",
    });
    expect(xml).toContain("c:spPr");
    expect(xml).toContain("00FF00");
    expect(xml).toContain('a:prstDash val="dash"');
  });

  it("drops error bars with invalid direction", () => {
    expect(
      buildErrorBars({ direction: "z" as any, type: "both", valType: "stdErr" }),
    ).toBeUndefined();
  });

  it("drops error bars with invalid type", () => {
    expect(
      buildErrorBars({ direction: "y", type: "bogus" as any, valType: "stdErr" }),
    ).toBeUndefined();
  });

  it("drops error bars with invalid valType", () => {
    expect(
      buildErrorBars({ direction: "y", type: "both", valType: "bogus" as any }),
    ).toBeUndefined();
  });

  it("buildAllErrorBars emits all valid entries", () => {
    const arr = buildAllErrorBars([
      { direction: "x", type: "both", valType: "fixedVal", value: 1 },
      { direction: "y", type: "both", valType: "fixedVal", value: 2 },
    ]);
    expect(arr).toHaveLength(2);
  });
});

describe("ErrorBars: reader", () => {
  it("parses minimal errBars", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="5"/></c:errBars></c:ser>`,
    );
    const res = parseErrorBars(ser);
    expect(res).toEqual([{ direction: "y", type: "both", valType: "fixedVal", value: 5 }]);
  });

  it("drops errBars without errDir", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:errBars><c:errBarType val="both"/><c:errValType val="fixedVal"/></c:errBars></c:ser>`,
    );
    expect(parseErrorBars(ser)).toBeUndefined();
  });

  it("drops errBars with invalid direction", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:errBars><c:errDir val="z"/><c:errBarType val="both"/><c:errValType val="fixedVal"/></c:errBars></c:ser>`,
    );
    expect(parseErrorBars(ser)).toBeUndefined();
  });

  it("parses noEndCap=true", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="stdErr"/><c:noEndCap val="1"/></c:errBars></c:ser>`,
    );
    expect(parseErrorBars(ser)?.[0].noEndCap).toBe(true);
  });

  it("parses stroke", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="stdErr"/><c:spPr xmlns:a="x"><a:ln w="38100"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:errBars></c:ser>`,
    );
    const res = parseErrorBars(ser);
    expect(res?.[0].lineColor).toBe("00FF00");
    expect(res?.[0].lineWidth).toBe(3);
    expect(res?.[0].lineDash).toBe("dash");
  });

  it("parses both x and y error bars on one series", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:errBars><c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars><c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="2"/></c:errBars></c:ser>`,
    );
    const res = parseErrorBars(ser);
    expect(res).toHaveLength(2);
    expect(res?.[0].direction).toBe("x");
    expect(res?.[1].direction).toBe("y");
  });
});

describe("ErrorBars: clone-through", () => {
  it("inherit / null / replace", () => {
    const src: ChartErrorBars[] = [{ direction: "y", type: "both", valType: "stdErr" }];
    expect(resolveErrorBars(src, undefined)).toEqual(src);
    expect(resolveErrorBars(src, null)).toBeUndefined();
    expect(
      resolveErrorBars(src, [{ direction: "x", type: "minus", valType: "stdDev", value: 2 }]),
    ).toEqual([{ direction: "x", type: "minus", valType: "stdDev", value: 2 }]);
  });

  it("clone drops invalid entries", () => {
    expect(
      cloneErrorBars({ direction: "x", type: "bogus" as any, valType: "stdErr" }),
    ).toBeUndefined();
  });
});

describe("ErrorBars: end-to-end", () => {
  it("emits c:errBars on a bar series", () => {
    const result = writeChart(
      makeChart({
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            errorBars: [{ direction: "y", type: "both", valType: "fixedVal", value: 5 }],
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:errBars");
    expect(result.chartXml).toContain('c:errValType val="fixedVal"');
  });

  it("does not emit c:errBars on a pie series", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            errorBars: [{ direction: "y", type: "both", valType: "stdErr" }],
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("c:errBars");
  });

  it("round-trips error bars through writeChart → parseChart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            errorBars: [
              { direction: "y", type: "both", valType: "stdDev", value: 1, noEndCap: true },
            ],
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    expect(parsed?.series?.[0].errorBars?.[0]).toMatchObject({
      direction: "y",
      type: "both",
      valType: "stdDev",
      value: 1,
      noEndCap: true,
    });
  });
});

// ── 3. Per-data-point overrides ──────────────────────────────────────

describe("DataPoints: writer", () => {
  it("emits a minimal dPt with idx and required bubble3D", () => {
    const arr = buildDataPoints([{ idx: 0 }], "column");
    expect(arr).toHaveLength(1);
    expect(arr[0]).toContain('c:idx val="0"');
    expect(arr[0]).toContain('c:bubble3D val="0"');
  });

  it("emits explosion only on pie family", () => {
    const arr = buildDataPoints([{ idx: 0, explosion: 50 }], "pie");
    expect(arr[0]).toContain('c:explosion val="50"');

    const arr2 = buildDataPoints([{ idx: 0, explosion: 50 }], "column");
    expect(arr2[0]).not.toContain("c:explosion");
  });

  it("emits fill color spPr", () => {
    const arr = buildDataPoints([{ idx: 0, fillColor: "FF0000" }], "column");
    expect(arr[0]).toContain('a:srgbClr val="FF0000"');
  });

  it("emits border color/width/dash", () => {
    const arr = buildDataPoints(
      [{ idx: 0, borderColor: "00FF00", borderWidth: 1, borderDash: "dash" }],
      "column",
    );
    expect(arr[0]).toContain("00FF00");
    expect(arr[0]).toContain('a:prstDash val="dash"');
  });

  it("emits marker block", () => {
    const arr = buildDataPoints([{ idx: 0, marker: { symbol: "circle", size: 5 } }], "line");
    expect(arr[0]).toContain("c:marker");
    expect(arr[0]).toContain('c:symbol val="circle"');
  });

  it("drops dPt with negative idx", () => {
    const arr = buildDataPoints([{ idx: -1 }], "column");
    expect(arr).toHaveLength(0);
  });

  it("clamps explosion to 0..400", () => {
    const arr = buildDataPoints([{ idx: 0, explosion: 999 }], "pie");
    expect(arr[0]).toContain('c:explosion val="400"');
  });
});

describe("DataPoints: reader", () => {
  it("parses minimal dPt", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:dPt><c:idx val="0"/><c:bubble3D val="0"/></c:dPt></c:ser>`,
    );
    const res = parseDataPoints(ser);
    expect(res).toEqual([{ idx: 0 }]);
  });

  it("parses dPt with fill", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr xmlns:a="x"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt></c:ser>`,
    );
    expect(parseDataPoints(ser)?.[0].fillColor).toBe("FF0000");
  });

  it("parses dPt with explosion and bubble3D", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:dPt><c:idx val="0"/><c:bubble3D val="1"/><c:explosion val="50"/></c:dPt></c:ser>`,
    );
    const res = parseDataPoints(ser);
    expect(res?.[0].explosion).toBe(50);
    expect(res?.[0].bubble3D).toBe(true);
  });

  it("drops dPt without idx", () => {
    const ser = parseXmlString(`<c:ser xmlns:c="x"><c:dPt><c:bubble3D val="0"/></c:dPt></c:ser>`);
    expect(parseDataPoints(ser)).toBeUndefined();
  });
});

describe("DataPoints: clone-through", () => {
  it("inherit / null / replace", () => {
    const src: ChartDataPoint[] = [{ idx: 0, fillColor: "FF0000" }];
    expect(resolveDataPoints(src, undefined)).toEqual(src);
    expect(resolveDataPoints(src, null)).toBeUndefined();
    expect(resolveDataPoints(src, [{ idx: 1, fillColor: "00FF00" }])).toEqual([
      { idx: 1, fillColor: "00FF00" },
    ]);
  });

  it("clone preserves all fields", () => {
    const c = cloneDataPoint({
      idx: 2,
      explosion: 25,
      bubble3D: true,
      fillColor: "FF0000",
      borderColor: "00FF00",
      borderWidth: 1.5,
      borderDash: "dash",
      marker: { symbol: "circle" },
    });
    expect(c).toMatchObject({
      idx: 2,
      explosion: 25,
      bubble3D: true,
      fillColor: "FF0000",
      borderColor: "00FF00",
      borderWidth: 1.5,
      borderDash: "dash",
      marker: { symbol: "circle" },
    });
  });
});

describe("DataPoints: end-to-end", () => {
  it("emits dPt blocks on a bar series", () => {
    const result = writeChart(
      makeChart({
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            dataPoints: [
              { idx: 0, fillColor: "FF0000" },
              { idx: 1, fillColor: "00FF00" },
            ],
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:dPt");
    expect(result.chartXml).toContain('a:srgbClr val="FF0000"');
    expect(result.chartXml).toContain('a:srgbClr val="00FF00"');
  });

  it("round-trips dPt through writeChart → parseChart", () => {
    const result = writeChart(
      makeChart({
        type: "pie",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            dataPoints: [{ idx: 1, explosion: 30, fillColor: "FF0000" }],
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    expect(parsed?.series?.[0].dataPoints?.[0]).toMatchObject({
      idx: 1,
      explosion: 30,
      fillColor: "FF0000",
    });
  });
});

// ── 4. Bubble size ──────────────────────────────────────────────────

describe("BubbleSize: reader", () => {
  it("parses bubble size formula", () => {
    const ser = parseXmlString(
      `<c:ser xmlns:c="x"><c:bubbleSize><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f></c:numRef></c:bubbleSize></c:ser>`,
    );
    expect(parseBubbleSizeRef(ser)).toBe("Sheet1!$C$2:$C$4");
  });

  it("returns undefined when bubbleSize is absent", () => {
    const ser = parseXmlString(`<c:ser xmlns:c="x"/>`);
    expect(parseBubbleSizeRef(ser)).toBeUndefined();
  });
});

// ── 5. 3D Shape ─────────────────────────────────────────────────────

describe("Shape3D: types & validation", () => {
  it("VALID_SHAPE_3D exports all 6 presets", () => {
    expect(VALID_SHAPE_3D.size).toBe(6);
    expect(VALID_SHAPE_3D.has("box")).toBe(true);
    expect(VALID_SHAPE_3D.has("cone")).toBe(true);
    expect(VALID_SHAPE_3D.has("coneToMax")).toBe(true);
    expect(VALID_SHAPE_3D.has("cylinder")).toBe(true);
    expect(VALID_SHAPE_3D.has("pyramid")).toBe(true);
    expect(VALID_SHAPE_3D.has("pyramidToMax")).toBe(true);
  });
});

describe("Shape3D: writer", () => {
  it("emits shape", () => {
    expect(buildShape3D("cone")).toContain('val="cone"');
    expect(buildShape3D("box")).toContain('val="box"');
  });

  it("returns undefined for invalid", () => {
    expect(buildShape3D("bogus" as any)).toBeUndefined();
    expect(buildShape3D(undefined)).toBeUndefined();
  });

  it("normalizeShape3D filters", () => {
    expect(normalizeShape3D("cone")).toBe("cone");
    expect(normalizeShape3D("bogus" as any)).toBeUndefined();
    expect(normalizeShape3D(undefined)).toBeUndefined();
  });
});

describe("Shape3D: reader", () => {
  it("parses shape from c:ser", () => {
    const ser = parseXmlString(`<c:ser xmlns:c="x"><c:shape val="cone"/></c:ser>`);
    expect(parseShape3D(ser)).toBe("cone");
  });

  it("drops unknown tokens", () => {
    const ser = parseXmlString(`<c:ser xmlns:c="x"><c:shape val="bogus"/></c:ser>`);
    expect(parseShape3D(ser)).toBeUndefined();
  });

  it("returns undefined when missing", () => {
    const ser = parseXmlString(`<c:ser xmlns:c="x"/>`);
    expect(parseShape3D(ser)).toBeUndefined();
  });
});

// ── 8. Legend entry rich overrides ──────────────────────────────────

describe("LegendEntry: rich overrides", () => {
  it("emits per-entry txPr block when fontSize / bold are set", () => {
    const result = writeChart(
      makeChart({
        legendEntries: [{ idx: 0, fontSize: 14, bold: true, italic: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:legendEntry");
    expect(result.chartXml).toMatch(/<c:legendEntry>[\s\S]*?<c:txPr>/);
    expect(result.chartXml).toContain('sz="1400"');
    expect(result.chartXml).toContain('b="1"');
    expect(result.chartXml).toContain('i="1"');
  });

  it("emits underline / strikethrough / color / fontFamily", () => {
    const result = writeChart(
      makeChart({
        legendEntries: [
          {
            idx: 0,
            underline: true,
            strikethrough: true,
            color: "FF0000",
            fontFamily: "Arial",
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('u="sng"');
    expect(result.chartXml).toContain('strike="sngStrike"');
    expect(result.chartXml).toContain('a:srgbClr val="FF0000"');
    expect(result.chartXml).toContain('a:latin typeface="Arial"');
  });

  it("does not emit txPr when only idx / delete are pinned", () => {
    const result = writeChart(
      makeChart({
        legendEntries: [{ idx: 0, delete: true }],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:legendEntry");
    // Per-entry txPr should NOT appear within the legendEntry; the
    // legend-level txPr may still appear elsewhere depending on pin.
    const entryStart = result.chartXml.indexOf("<c:legendEntry");
    const entryEnd = result.chartXml.indexOf("</c:legendEntry>");
    const entryXml = result.chartXml.slice(entryStart, entryEnd);
    expect(entryXml).not.toContain("c:txPr");
  });

  it("round-trips per-entry typography", () => {
    const result = writeChart(
      makeChart({
        legendEntries: [
          {
            idx: 1,
            delete: false,
            fontSize: 12,
            bold: true,
            italic: false,
            color: "FF00FF",
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    const entry = parsed?.legendEntries?.find((e) => e.idx === 1);
    expect(entry?.fontSize).toBe(12);
    expect(entry?.bold).toBe(true);
    expect(entry?.color).toBe("FF00FF");
  });
});

// ── 9. DispUnits customLabel ────────────────────────────────────────

describe("DispUnits: customLabel", () => {
  it("emits c:dispUnitsLbl with rich text when customLabel is set", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: { dispUnits: { unit: "millions", showLabel: true, customLabel: "$ Millions" } },
        },
      } as any),
      "Sheet1",
    );
    expect(result.chartXml).toContain("c:dispUnitsLbl");
    expect(result.chartXml).toContain("$ Millions");
    expect(result.chartXml).toContain("c:rich");
  });

  it("emits a bare dispUnitsLbl when only showLabel is set", () => {
    const result = writeChart(
      makeChart({
        title: undefined,
        axes: { y: { dispUnits: { unit: "millions", showLabel: true } } },
      } as any),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:dispUnitsLbl/>");
  });

  it("collapses empty / whitespace-only customLabel to absence", () => {
    const result = writeChart(
      makeChart({
        title: undefined,
        axes: { y: { dispUnits: { unit: "millions", showLabel: true, customLabel: "  " } } },
      } as any),
      "Sheet1",
    );
    expect(result.chartXml).toContain("<c:dispUnitsLbl/>");
  });

  it("escapes XML special characters in customLabel", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: { dispUnits: { unit: "millions", showLabel: true, customLabel: "A & B < C" } },
        },
      } as any),
      "Sheet1",
    );
    expect(result.chartXml).toContain("A &amp; B &lt; C");
  });

  it("round-trips customLabel through writeChart → parseChart", () => {
    const result = writeChart(
      makeChart({
        axes: {
          y: { dispUnits: { unit: "millions", showLabel: true, customLabel: "$ Million" } },
        },
      } as any),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    expect(parsed?.axes?.y?.dispUnits?.customLabel).toBe("$ Million");
  });
});

// ── 7. Line cap and compound (series stroke) ───────────────────────

describe("ChartLineStroke: cap and compound", () => {
  it("emits cap and cmpd attributes on a:ln", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { width: 2, cap: "rnd", compound: "dbl" },
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).toContain('cap="rnd"');
    expect(result.chartXml).toContain('cmpd="dbl"');
  });

  it("does not emit cap='flat' (the OOXML default)", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { cap: "flat" },
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain('cap="flat"');
  });

  it("does not emit cmpd='sng' (the OOXML default)", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { compound: "sng" },
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain('cmpd="sng"');
  });

  it("drops invalid cap / compound tokens", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { cap: "bogus" as any, compound: "bogus" as any },
          },
        ],
      }),
      "Sheet1",
    );
    expect(result.chartXml).not.toContain("cap=");
    expect(result.chartXml).not.toContain("cmpd=");
  });

  it("round-trips cap and compound through writeChart → parseChart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { cap: "rnd", compound: "thickThin" },
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    expect(parsed?.series?.[0].stroke?.cap).toBe("rnd");
    expect(parsed?.series?.[0].stroke?.compound).toBe("thickThin");
  });

  it("collapses default flat / sng on parse", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { width: 2 },
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml);
    expect(parsed?.series?.[0].stroke?.cap).toBeUndefined();
    expect(parsed?.series?.[0].stroke?.compound).toBeUndefined();
  });

  it("clone preserves cap and compound", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { cap: "sq", compound: "tri" },
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml)!;
    const cloned = cloneChart(parsed, { anchor: { from: { row: 0, col: 0 } } });
    expect(cloned.series[0].stroke?.cap).toBe("sq");
    expect(cloned.series[0].stroke?.compound).toBe("tri");
  });

  it("clone null override drops the entire stroke including cap/compound", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            stroke: { cap: "sq", compound: "dbl" },
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml)!;
    const cloned = cloneChart(parsed, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ stroke: null }],
    });
    expect(cloned.series[0].stroke).toBeUndefined();
  });
});

// ── End-to-end clone-through tests ──────────────────────────────────

describe("cloneChart: trendlines / errorBars / dataPoints round-trip", () => {
  it("preserves trendlines through writeXlsx round trip", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Cat", "Val"],
            ["A", 1],
            ["B", 2],
            ["C", 3],
          ],
          charts: [
            {
              type: "line",
              title: "T",
              series: [
                {
                  name: "S",
                  values: "B2:B4",
                  categories: "A2:A4",
                  trendlines: [{ type: "linear", dispEquation: true }],
                },
              ],
              anchor: { from: { row: 5, col: 0 } },
            },
          ],
        },
      ],
    });
    const zip = new ZipReader(data);
    expect(zip.has("xl/charts/chart1.xml")).toBe(true);
    const xml = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(xml).toContain("c:trendline");
    expect(xml).toContain('val="linear"');
    const parsed = parseChart(xml);
    expect(parsed?.series?.[0].trendlines?.[0]).toMatchObject({
      type: "linear",
      dispEquation: true,
    });
  });

  it("clones a chart with trendlines via cloneChart", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            trendlines: [{ type: "movingAvg", period: 3 }],
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml)!;
    const cloned = cloneChart(parsed, { anchor: { from: { row: 0, col: 0 } } });
    expect(cloned.series[0].trendlines).toEqual([{ type: "movingAvg", period: 3 }]);
  });

  it("cloneChart supports null override to drop trendlines", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            trendlines: [{ type: "linear" }],
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml)!;
    const cloned = cloneChart(parsed, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ trendlines: null }],
    });
    expect(cloned.series[0].trendlines).toBeUndefined();
  });

  it("cloneChart supports replace override for errorBars", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            errorBars: [{ direction: "y", type: "both", valType: "stdErr" }],
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml)!;
    const cloned = cloneChart(parsed, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [
        { errorBars: [{ direction: "y", type: "minus", valType: "fixedVal", value: 5 }] },
      ],
    });
    expect(cloned.series[0].errorBars?.[0]).toMatchObject({
      direction: "y",
      type: "minus",
      valType: "fixedVal",
      value: 5,
    });
  });

  it("flattens trendlines and errorBars away on a pie clone", () => {
    const result = writeChart(
      makeChart({
        type: "line",
        series: [
          {
            name: "A",
            values: "B2:B4",
            categories: "A2:A4",
            trendlines: [{ type: "linear" }],
            errorBars: [{ direction: "y", type: "both", valType: "stdErr" }],
          },
        ],
      }),
      "Sheet1",
    );
    const parsed = parseChart(result.chartXml)!;
    const cloned = cloneChart(parsed, {
      anchor: { from: { row: 0, col: 0 } },
      type: "pie",
    });
    expect(cloned.series[0].trendlines).toBeUndefined();
    expect(cloned.series[0].errorBars).toBeUndefined();
  });
});
