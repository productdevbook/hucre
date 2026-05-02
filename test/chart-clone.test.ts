import { describe, expect, it } from "vitest";
import { chartKindToWriteKind, cloneChart } from "../src/xlsx/chart-clone";
import { parseChart } from "../src/xlsx/chart-reader";
import { writeChart } from "../src/xlsx/chart-writer";
import { writeXlsx } from "../src/xlsx/writer";
import { ZipReader } from "../src/zip/reader";
import type { Chart, SheetChart } from "../src/_types";

const decoder = new TextDecoder("utf-8");

// ── chartKindToWriteKind ─────────────────────────────────────────────

describe("chartKindToWriteKind", () => {
  it("maps every read-side kind that has a write-side analog", () => {
    expect(chartKindToWriteKind("bar")).toBe("column");
    expect(chartKindToWriteKind("bar3D")).toBe("column");
    expect(chartKindToWriteKind("line")).toBe("line");
    expect(chartKindToWriteKind("line3D")).toBe("line");
    expect(chartKindToWriteKind("pie")).toBe("pie");
    expect(chartKindToWriteKind("pie3D")).toBe("pie");
    expect(chartKindToWriteKind("doughnut")).toBe("pie");
    expect(chartKindToWriteKind("area")).toBe("area");
    expect(chartKindToWriteKind("area3D")).toBe("area");
    expect(chartKindToWriteKind("scatter")).toBe("scatter");
  });

  it("returns undefined for kinds the writer cannot author", () => {
    expect(chartKindToWriteKind("bubble")).toBeUndefined();
    expect(chartKindToWriteKind("radar")).toBeUndefined();
    expect(chartKindToWriteKind("surface")).toBeUndefined();
    expect(chartKindToWriteKind("surface3D")).toBeUndefined();
    expect(chartKindToWriteKind("stock")).toBeUndefined();
    expect(chartKindToWriteKind("ofPie")).toBeUndefined();
  });
});

// ── cloneChart — basics ──────────────────────────────────────────────

describe("cloneChart", () => {
  function source(extra?: Partial<Chart>): Chart {
    return {
      kinds: ["bar"],
      seriesCount: 1,
      title: "Template Revenue",
      series: [
        {
          kind: "bar",
          index: 0,
          name: "Revenue",
          valuesRef: "Sheet1!$B$2:$B$5",
          categoriesRef: "Sheet1!$A$2:$A$5",
          color: "1F77B4",
        },
      ],
      ...extra,
    };
  }

  it("requires options.anchor", () => {
    expect(() =>
      // @ts-expect-error — testing runtime guard for missing required field
      cloneChart(source(), {}),
    ).toThrow(/anchor is required/);
  });

  it("carries source title, name, ranges, and color through to the clone", () => {
    const clone = cloneChart(source(), {
      anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
    });

    expect(clone.type).toBe("column"); // bar → column (vertical default)
    expect(clone.title).toBe("Template Revenue");
    expect(clone.anchor).toEqual({ from: { row: 5, col: 0 }, to: { row: 20, col: 6 } });
    expect(clone.series).toEqual([
      {
        name: "Revenue",
        values: "Sheet1!$B$2:$B$5",
        categories: "Sheet1!$A$2:$A$5",
        color: "1F77B4",
      },
    ]);
  });

  it("honors options.type to coerce kinds the writer cannot author", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"] }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "pie",
    });
    expect(clone.type).toBe("pie");
  });

  it("auto-collapses doughnut to pie when no type override is given", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"] }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("pie");
  });

  it("throws when the source has no writable kind and no override is given", () => {
    expect(() =>
      cloneChart(source({ kinds: ["bubble", "radar"] }), {
        anchor: { from: { row: 0, col: 0 } },
      }),
    ).toThrow(/cannot be authored on the write side/);
  });

  it("throws when the source has no kinds and no override is given", () => {
    expect(() =>
      cloneChart({ kinds: [], seriesCount: 0 }, { anchor: { from: { row: 0, col: 0 } } }),
    ).toThrow(/no kinds/);
  });

  it("falls back to options.type when source has no writable kind", () => {
    const clone = cloneChart(source({ kinds: ["bubble"] }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "scatter",
      seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
    });
    expect(clone.type).toBe("scatter");
  });

  it("drops the source title when title=null is passed", () => {
    const clone = cloneChart(source(), {
      anchor: { from: { row: 0, col: 0 } },
      title: null,
    });
    expect(clone.title).toBeUndefined();
  });

  it("replaces the source title when a string is passed", () => {
    const clone = cloneChart(source(), {
      anchor: { from: { row: 0, col: 0 } },
      title: "Q1 Revenue",
    });
    expect(clone.title).toBe("Q1 Revenue");
  });

  it("forwards legend, barGrouping, showTitle, altText, frameTitle", () => {
    const clone = cloneChart(source(), {
      anchor: { from: { row: 0, col: 0 } },
      legend: "bottom",
      barGrouping: "stacked",
      showTitle: false,
      altText: "Revenue chart",
      frameTitle: "Revenue",
    });
    expect(clone.legend).toBe("bottom");
    expect(clone.barGrouping).toBe("stacked");
    expect(clone.showTitle).toBe(false);
    expect(clone.altText).toBe("Revenue chart");
    expect(clone.frameTitle).toBe("Revenue");
  });

  it("inherits legend from the source chart when no override is given", () => {
    const clone = cloneChart(source({ legend: "bottom" }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.legend).toBe("bottom");
  });

  it("inherits legend=false (hidden) from the source chart", () => {
    const clone = cloneChart(source({ legend: false }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.legend).toBe(false);
  });

  it("override wins over source legend", () => {
    const clone = cloneChart(source({ legend: "bottom" }), {
      anchor: { from: { row: 0, col: 0 } },
      legend: "top",
    });
    expect(clone.legend).toBe("top");
  });

  it("override legend=false hides a legend the source declared", () => {
    const clone = cloneChart(source({ legend: "right" }), {
      anchor: { from: { row: 0, col: 0 } },
      legend: false,
    });
    expect(clone.legend).toBe(false);
  });

  it("inherits barGrouping from the source bar/column chart", () => {
    const clone = cloneChart(source({ barGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("column");
    expect(clone.barGrouping).toBe("stacked");
  });

  it("override barGrouping wins over source barGrouping", () => {
    const clone = cloneChart(source({ barGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
      barGrouping: "percentStacked",
    });
    expect(clone.barGrouping).toBe("percentStacked");
  });

  it("drops inherited barGrouping when the clone target is not bar/column", () => {
    // Source is a bar chart with stacked grouping; override coerces
    // it to a line chart. Stacked grouping is meaningless for line so
    // it should not survive on the clone.
    const clone = cloneChart(source({ kinds: ["bar"], barGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "line",
      seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
    });
    expect(clone.type).toBe("line");
    expect(clone.barGrouping).toBeUndefined();
  });
});

// ── cloneChart — series overrides ────────────────────────────────────

describe("cloneChart — series overrides", () => {
  const twoSeries: Chart = {
    kinds: ["bar"],
    seriesCount: 2,
    series: [
      {
        kind: "bar",
        index: 0,
        name: "Revenue",
        valuesRef: "Tpl!$B$2:$B$5",
        categoriesRef: "Tpl!$A$2:$A$5",
        color: "1070CA",
      },
      {
        kind: "bar",
        index: 1,
        name: "Cost",
        valuesRef: "Tpl!$C$2:$C$5",
        categoriesRef: "Tpl!$A$2:$A$5",
        color: "E54D2E",
      },
    ],
  };

  it("merges per-series overrides on top of the source", () => {
    const clone = cloneChart(twoSeries, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [
        { values: "Dashboard!$B$2:$B$13", color: "00C586" },
        { name: "Total Cost" },
      ],
    });

    expect(clone.series).toEqual([
      {
        name: "Revenue",
        values: "Dashboard!$B$2:$B$13",
        categories: "Tpl!$A$2:$A$5",
        color: "00C586",
      },
      {
        name: "Total Cost",
        values: "Tpl!$C$2:$C$5",
        categories: "Tpl!$A$2:$A$5",
        color: "E54D2E",
      },
    ]);
  });

  it("treats null overrides as 'drop the inherited value'", () => {
    const clone = cloneChart(twoSeries, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ name: null, color: null, categories: null }, undefined],
    });

    expect(clone.series[0].name).toBeUndefined();
    expect(clone.series[0].color).toBeUndefined();
    expect(clone.series[0].categories).toBeUndefined();
    // Untouched series retains its source values.
    expect(clone.series[1].name).toBe("Cost");
  });

  it("appends a new series past the source length when provided", () => {
    const clone = cloneChart(twoSeries, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [undefined, undefined, { name: "Margin", values: "Dashboard!$D$2:$D$13" }],
    });

    expect(clone.series).toHaveLength(3);
    expect(clone.series[2]).toEqual({
      name: "Margin",
      values: "Dashboard!$D$2:$D$13",
    });
  });

  it("replaces the entire series array when options.series is provided", () => {
    const clone = cloneChart(twoSeries, {
      anchor: { from: { row: 0, col: 0 } },
      series: [{ name: "Only", values: "Sheet1!$B$2:$B$10" }],
    });

    expect(clone.series).toEqual([{ name: "Only", values: "Sheet1!$B$2:$B$10" }]);
  });

  it("throws when a series ends up without a values reference", () => {
    const noValues: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, name: "Empty" }],
    };
    expect(() =>
      cloneChart(noValues, {
        anchor: { from: { row: 0, col: 0 } },
      }),
    ).toThrow(/no values reference/);
  });

  it("throws when both source and options produce zero series", () => {
    expect(() =>
      cloneChart({ kinds: ["bar"], seriesCount: 0 }, { anchor: { from: { row: 0, col: 0 } } }),
    ).toThrow(/0 series/);
  });
});

// ── cloneChart — axis titles ────────────────────────────────────────

describe("cloneChart — axis titles", () => {
  const sourceWithAxes: Chart = {
    kinds: ["bar"],
    seriesCount: 1,
    series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    axes: { x: { title: "Quarter" }, y: { title: "Revenue" } },
  };

  it("inherits the source's axes when no override is given", () => {
    const clone = cloneChart(sourceWithAxes, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes).toEqual({
      x: { title: "Quarter" },
      y: { title: "Revenue" },
    });
  });

  it("does not set axes when the source has none", () => {
    const noAxes: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    };
    const clone = cloneChart(noAxes, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes).toBeUndefined();
  });

  it("replaces a single axis title via override", () => {
    const clone = cloneChart(sourceWithAxes, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { x: { title: "Period" } },
    });
    expect(clone.axes).toEqual({
      x: { title: "Period" },
      y: { title: "Revenue" },
    });
  });

  it("drops a source axis title when override is null", () => {
    const clone = cloneChart(sourceWithAxes, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { title: null } },
    });
    expect(clone.axes).toEqual({ x: { title: "Quarter" } });
    expect(clone.axes?.y).toBeUndefined();
  });

  it("adds an axis title that the source did not declare", () => {
    const partial: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { x: { title: "Quarter" } },
    };
    const clone = cloneChart(partial, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { title: "Revenue" } },
    });
    expect(clone.axes).toEqual({
      x: { title: "Quarter" },
      y: { title: "Revenue" },
    });
  });

  it("drops axes silently when the resolved chart type is pie", () => {
    const pieSource: Chart = {
      kinds: ["pie"],
      seriesCount: 1,
      series: [{ kind: "pie", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      // Pie charts shouldn't carry axes, but the parser cannot know
      // ahead of time — make sure cloneChart strips them on output.
      axes: { x: { title: "Spurious" }, y: { title: "Spurious Y" } },
    };
    const clone = cloneChart(pieSource, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("pie");
    expect(clone.axes).toBeUndefined();
  });

  it("returns axes undefined when both x and y resolve to undefined", () => {
    const clone = cloneChart(sourceWithAxes, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { x: { title: null }, y: { title: null } },
    });
    expect(clone.axes).toBeUndefined();
  });
});

// ── cloneChart — axis gridlines ─────────────────────────────────────

describe("cloneChart — axis gridlines", () => {
  const sourceWithGridlines: Chart = {
    kinds: ["bar"],
    seriesCount: 1,
    series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    axes: {
      y: { gridlines: { major: true } },
    },
  };

  it("inherits the source's gridlines when no override is given", () => {
    const clone = cloneChart(sourceWithGridlines, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes).toEqual({
      y: { gridlines: { major: true } },
    });
  });

  it("inherits both major and minor gridlines together", () => {
    const both: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { y: { gridlines: { major: true, minor: true } } },
    };
    const clone = cloneChart(both, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes?.y?.gridlines).toEqual({ major: true, minor: true });
  });

  it("co-inherits the title and gridlines on the same axis", () => {
    const both: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { y: { title: "Revenue", gridlines: { major: true } } },
    };
    const clone = cloneChart(both, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes?.y).toEqual({
      title: "Revenue",
      gridlines: { major: true },
    });
  });

  it("drops inherited gridlines when override is null", () => {
    const clone = cloneChart(sourceWithGridlines, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { gridlines: null } },
    });
    expect(clone.axes).toBeUndefined();
  });

  it("replaces inherited gridlines with the override", () => {
    const clone = cloneChart(sourceWithGridlines, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { gridlines: { major: true, minor: true } } },
    });
    expect(clone.axes?.y?.gridlines).toEqual({ major: true, minor: true });
  });

  it("adds gridlines to an axis the source did not declare them on", () => {
    const noGridlines: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    };
    const clone = cloneChart(noGridlines, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { gridlines: { major: true } } },
    });
    expect(clone.axes?.y?.gridlines).toEqual({ major: true });
  });

  it("strips gridlines silently when the resolved chart type is pie", () => {
    const pieSource: Chart = {
      kinds: ["pie"],
      seriesCount: 1,
      series: [{ kind: "pie", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { y: { gridlines: { major: true } } },
    };
    const clone = cloneChart(pieSource, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("pie");
    expect(clone.axes).toBeUndefined();
  });

  it("treats { major: false, minor: false } overrides as no gridlines", () => {
    const clone = cloneChart(sourceWithGridlines, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { gridlines: { major: false, minor: false } } },
    });
    expect(clone.axes).toBeUndefined();
  });
});

// ── cloneChart — round-trip with parseChart and writeXlsx ────────────

describe("cloneChart — integration", () => {
  it("produces a SheetChart that writeChart accepts and writeXlsx packages", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Template Title</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Series A</c:v></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="1070CA"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const parsed = parseChart(sourceXml);
    expect(parsed).toBeDefined();

    const sheetChart: SheetChart = cloneChart(parsed!, {
      anchor: { from: { row: 14, col: 0 }, to: { row: 28, col: 8 } },
      title: "Revenue",
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$13", color: "00C586" }],
    });

    // writeChart should accept the result and emit the expected fingerprints.
    const result = writeChart(sheetChart, "Dashboard");
    expect(result.chartXml).toContain("<c:lineChart>");
    expect(result.chartXml).toContain("Revenue");
    expect(result.chartXml).toContain("Dashboard!$B$2:$B$13");
    expect(result.chartXml).toContain('val="00C586"');
    // Categories from source should survive.
    expect(result.chartXml).toContain("Tpl!$A$2:$A$5");

    // End-to-end: writeXlsx packages the chart into a valid OOXML file.
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [["Header"], [1], [2], [3], [4]],
          charts: [sheetChart],
        },
      ],
    });

    const zip = new ZipReader(xlsx);
    expect(zip.has("xl/charts/chart1.xml")).toBe(true);
    const writtenChart = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(writtenChart).toContain("<c:lineChart>");
    expect(writtenChart).toContain("Dashboard!$B$2:$B$13");
  });

  it("inherits the source chart's dataLabels block by default", () => {
    const src: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [
        {
          kind: "bar",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
        },
      ],
      dataLabels: { showValue: true, position: "outEnd" },
    };
    const clone = cloneChart(src, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.dataLabels).toEqual({ showValue: true, position: "outEnd" });
  });

  it("replaces the chart-level dataLabels when an override object is given", () => {
    const src: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      dataLabels: { showValue: true },
    };
    const clone = cloneChart(src, {
      anchor: { from: { row: 0, col: 0 } },
      dataLabels: { showCategoryName: true, position: "ctr" },
    });
    expect(clone.dataLabels).toEqual({ showCategoryName: true, position: "ctr" });
  });

  it("drops the chart-level dataLabels when override is null", () => {
    const src: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      dataLabels: { showValue: true },
    };
    const clone = cloneChart(src, {
      anchor: { from: { row: 0, col: 0 } },
      dataLabels: null,
    });
    expect(clone.dataLabels).toBeUndefined();
  });

  it("inherits per-series dataLabels by default", () => {
    const src: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [
        {
          kind: "bar",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          dataLabels: { showValue: true, position: "ctr" },
        },
      ],
    };
    const clone = cloneChart(src, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].dataLabels).toEqual({ showValue: true, position: "ctr" });
  });

  it("replaces per-series dataLabels via seriesOverrides", () => {
    const src: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [
        {
          kind: "bar",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          dataLabels: { showValue: true },
        },
      ],
    };
    const clone = cloneChart(src, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ dataLabels: { showCategoryName: true } }],
    });
    expect(clone.series[0].dataLabels).toEqual({ showCategoryName: true });
  });

  it("drops per-series dataLabels when override is null", () => {
    const src: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [
        {
          kind: "bar",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          dataLabels: { showValue: true },
        },
      ],
    };
    const clone = cloneChart(src, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ dataLabels: null }],
    });
    expect(clone.series[0].dataLabels).toBeUndefined();
  });

  it("suppresses a single series via seriesOverrides[i].dataLabels = false", () => {
    const src: Chart = {
      kinds: ["bar"],
      seriesCount: 2,
      series: [
        { kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" },
        { kind: "bar", index: 1, valuesRef: "Tpl!$C$2:$C$5" },
      ],
      dataLabels: { showValue: true },
    };
    const clone = cloneChart(src, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [undefined, { dataLabels: false }],
    });
    expect(clone.dataLabels).toEqual({ showValue: true });
    expect(clone.series[0].dataLabels).toBeUndefined();
    expect(clone.series[1].dataLabels).toBe(false);
  });

  it("can clone the same template into multiple chart instances", () => {
    const template: Chart = {
      kinds: ["pie"],
      seriesCount: 1,
      title: "Template",
      series: [
        {
          kind: "pie",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          color: "1070CA",
        },
      ],
    };

    const dashboards = [
      { title: "Q1", values: "Dash!$B$2:$B$5", color: "1070CA" },
      { title: "Q2", values: "Dash!$C$2:$C$5", color: "00C586" },
      { title: "Q3", values: "Dash!$D$2:$D$5", color: "F76808" },
    ];

    const clones = dashboards.map((d, i) =>
      cloneChart(template, {
        anchor: { from: { row: i * 16, col: 0 } },
        title: d.title,
        seriesOverrides: [{ values: d.values, color: d.color }],
      }),
    );

    expect(clones).toHaveLength(3);
    expect(clones[0].title).toBe("Q1");
    expect(clones[0].series[0].values).toBe("Dash!$B$2:$B$5");
    expect(clones[2].series[0].color).toBe("F76808");
    // Categories carry through unchanged.
    expect(clones.every((c) => c.series[0].categories === "Tpl!$A$2:$A$5")).toBe(true);
  });

  it("round-trips data labels: parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Revenue</c:v></c:tx>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
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
    const parsed = parseChart(sourceXml)!;
    expect(parsed.dataLabels).toEqual({ showValue: true, position: "outEnd" });

    const sheetChart: SheetChart = cloneChart(parsed, {
      anchor: { from: { row: 5, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.dataLabels).toEqual({ showValue: true, position: "outEnd" });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [["Header"], [10], [20], [30], [40]],
          charts: [sheetChart],
        },
      ],
    });

    const zip = new ZipReader(xlsx);
    const writtenChart = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(writtenChart).toContain('c:dLblPos val="outEnd"');
    expect(writtenChart).toContain('c:showVal val="1"');

    const reparsed = parseChart(writtenChart)!;
    expect(reparsed.dataLabels).toEqual({ showValue: true, position: "outEnd" });
  });

  it("round-trips axis titles through parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Revenue</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="111"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="222"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue (USD)</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml)!;
    expect(source.axes).toEqual({ x: { title: "Quarter" }, y: { title: "Revenue (USD)" } });

    const sheetChart: SheetChart = cloneChart(source, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.axes).toEqual({
      x: { title: "Quarter" },
      y: { title: "Revenue (USD)" },
    });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [["Header"], [1], [2], [3], [4]],
          charts: [sheetChart],
        },
      ],
    });
    const zip = new ZipReader(xlsx);
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    const reparsed = parseChart(written);
    expect(reparsed?.axes).toEqual({
      x: { title: "Quarter" },
      y: { title: "Revenue (USD)" },
    });
  });
});
