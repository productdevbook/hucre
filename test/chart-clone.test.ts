import { describe, expect, it } from "vitest";
import { chartKindToWriteKind, cloneChart } from "../src/xlsx/chart-clone";
import { parseChart } from "../src/xlsx/chart-reader";
import { writeChart } from "../src/xlsx/chart-writer";
import { writeXlsx } from "../src/xlsx/writer";
import { ZipReader } from "../src/zip/reader";
import type { Chart, ChartLineStroke, ChartMarker, SheetChart } from "../src/_types";

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
    expect(chartKindToWriteKind("doughnut")).toBe("doughnut");
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
    const clone = cloneChart(
      source({
        kinds: ["radar"],
        series: [{ kind: "radar", index: 0, valuesRef: "Sheet1!$B$2:$B$5" }],
      }),
      {
        anchor: { from: { row: 0, col: 0 } },
        type: "line",
      },
    );
    expect(clone.type).toBe("line");
  });

  it("preserves doughnut as the writable kind when no type override is given", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"] }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("doughnut");
  });

  it("flattens doughnut to pie when type='pie' is requested", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"], holeSize: 60 }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "pie",
    });
    expect(clone.type).toBe("pie");
    expect(clone.holeSize).toBeUndefined();
  });

  it("inherits the source's holeSize on a doughnut clone", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"], holeSize: 65 }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("doughnut");
    expect(clone.holeSize).toBe(65);
  });

  it("lets options.holeSize override the source's holeSize", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"], holeSize: 65 }), {
      anchor: { from: { row: 0, col: 0 } },
      holeSize: 30,
    });
    expect(clone.holeSize).toBe(30);
  });

  it("drops options.holeSize when the resolved type is not doughnut", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"] }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "pie",
      holeSize: 60,
    });
    expect(clone.holeSize).toBeUndefined();
  });

  // ── gapWidth / overlap (bar / column only) ──────────────────────────

  it("inherits the source's gapWidth on a column clone", () => {
    const clone = cloneChart(source({ kinds: ["bar"], gapWidth: 75 }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("column");
    expect(clone.gapWidth).toBe(75);
  });

  it("inherits the source's overlap on a column clone", () => {
    const clone = cloneChart(source({ kinds: ["bar"], overlap: -25 }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.overlap).toBe(-25);
  });

  it("inherits both gapWidth and overlap together on a bar clone", () => {
    const clone = cloneChart(source({ kinds: ["bar"], gapWidth: 50, overlap: 100 }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "bar",
    });
    expect(clone.type).toBe("bar");
    expect(clone.gapWidth).toBe(50);
    expect(clone.overlap).toBe(100);
  });

  it("lets options.gapWidth override the source's gapWidth", () => {
    const clone = cloneChart(source({ kinds: ["bar"], gapWidth: 75 }), {
      anchor: { from: { row: 0, col: 0 } },
      gapWidth: 200,
    });
    expect(clone.gapWidth).toBe(200);
  });

  it("lets options.overlap override the source's overlap", () => {
    const clone = cloneChart(source({ kinds: ["bar"], overlap: -25 }), {
      anchor: { from: { row: 0, col: 0 } },
      overlap: 50,
    });
    expect(clone.overlap).toBe(50);
  });

  it("drops options.gapWidth / overlap when the resolved type is not bar/column", () => {
    const clone = cloneChart(source({ kinds: ["bar"] }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "line",
      gapWidth: 75,
      overlap: 50,
    });
    expect(clone.gapWidth).toBeUndefined();
    expect(clone.overlap).toBeUndefined();
  });

  it("drops the inherited gapWidth / overlap when the resolved type is not bar/column", () => {
    const clone = cloneChart(source({ kinds: ["bar"], gapWidth: 75, overlap: -25 }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "line",
      seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
    });
    expect(clone.type).toBe("line");
    expect(clone.gapWidth).toBeUndefined();
    expect(clone.overlap).toBeUndefined();
  });

  // ── firstSliceAng (pie / doughnut only) ──────────────────────────

  it("inherits the source's firstSliceAng on a pie clone", () => {
    const clone = cloneChart(source({ kinds: ["pie"], firstSliceAng: 90 }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("pie");
    expect(clone.firstSliceAng).toBe(90);
  });

  it("inherits the source's firstSliceAng on a doughnut clone", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"], firstSliceAng: 180 }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("doughnut");
    expect(clone.firstSliceAng).toBe(180);
  });

  it("carries firstSliceAng through when flattening doughnut to pie", () => {
    // The element lives on both <c:pieChart> and <c:doughnutChart>, so
    // a doughnut template flattened to pie keeps its rotation.
    const clone = cloneChart(source({ kinds: ["doughnut"], firstSliceAng: 270 }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "pie",
    });
    expect(clone.type).toBe("pie");
    expect(clone.firstSliceAng).toBe(270);
  });

  it("lets options.firstSliceAng override the source's firstSliceAng", () => {
    const clone = cloneChart(source({ kinds: ["pie"], firstSliceAng: 45 }), {
      anchor: { from: { row: 0, col: 0 } },
      firstSliceAng: 180,
    });
    expect(clone.firstSliceAng).toBe(180);
  });

  it("drops options.firstSliceAng when the resolved type is neither pie nor doughnut", () => {
    const clone = cloneChart(source({ kinds: ["doughnut"] }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "column",
      firstSliceAng: 90,
      seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
    });
    expect(clone.firstSliceAng).toBeUndefined();
  });

  it("drops the inherited firstSliceAng when the resolved type is not pie/doughnut", () => {
    const clone = cloneChart(source({ kinds: ["pie"], firstSliceAng: 90 }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "line",
      seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
    });
    expect(clone.type).toBe("line");
    expect(clone.firstSliceAng).toBeUndefined();
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

  it("inherits lineGrouping from the source line chart", () => {
    const clone = cloneChart(source({ kinds: ["line"], lineGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("line");
    expect(clone.lineGrouping).toBe("stacked");
  });

  it("override lineGrouping wins over source lineGrouping", () => {
    const clone = cloneChart(source({ kinds: ["line"], lineGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
      lineGrouping: "percentStacked",
    });
    expect(clone.lineGrouping).toBe("percentStacked");
  });

  it("drops inherited lineGrouping when the clone target is not line", () => {
    const clone = cloneChart(source({ kinds: ["line"], lineGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "column",
    });
    expect(clone.type).toBe("column");
    expect(clone.lineGrouping).toBeUndefined();
  });

  it("inherits areaGrouping from the source area chart", () => {
    const clone = cloneChart(source({ kinds: ["area"], areaGrouping: "percentStacked" }), {
      anchor: { from: { row: 0, col: 0 } },
    });
    expect(clone.type).toBe("area");
    expect(clone.areaGrouping).toBe("percentStacked");
  });

  it("override areaGrouping wins over source areaGrouping", () => {
    const clone = cloneChart(source({ kinds: ["area"], areaGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
      areaGrouping: "percentStacked",
    });
    expect(clone.areaGrouping).toBe("percentStacked");
  });

  it("drops inherited areaGrouping when the clone target is not area", () => {
    const clone = cloneChart(source({ kinds: ["area"], areaGrouping: "stacked" }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "line",
      seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
    });
    expect(clone.type).toBe("line");
    expect(clone.areaGrouping).toBeUndefined();
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

// ── cloneChart — axis scale ─────────────────────────────────────────

describe("cloneChart — axis scale", () => {
  const sourceWithScale: Chart = {
    kinds: ["bar"],
    seriesCount: 1,
    series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    axes: {
      y: { scale: { min: 0, max: 100, majorUnit: 25 } },
    },
  };

  it("inherits the source's scale when no override is given", () => {
    const clone = cloneChart(sourceWithScale, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes?.y?.scale).toEqual({ min: 0, max: 100, majorUnit: 25 });
  });

  it("drops inherited scale when override is null", () => {
    const clone = cloneChart(sourceWithScale, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { scale: null } },
    });
    expect(clone.axes).toBeUndefined();
  });

  it("replaces the inherited scale wholesale (does not merge field-by-field)", () => {
    const clone = cloneChart(sourceWithScale, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { scale: { max: 50 } } },
    });
    // No min should leak through from the source — wholesale replace.
    expect(clone.axes?.y?.scale).toEqual({ max: 50 });
  });

  it("adds a scale to an axis the source did not declare it on", () => {
    const noScale: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    };
    const clone = cloneChart(noScale, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { scale: { min: 0, max: 200 } } },
    });
    expect(clone.axes?.y?.scale).toEqual({ min: 0, max: 200 });
  });

  it("strips scale silently when the resolved chart type is pie", () => {
    const pieSource: Chart = {
      kinds: ["pie"],
      seriesCount: 1,
      series: [{ kind: "pie", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { y: { scale: { min: 0, max: 100 } } },
    };
    const clone = cloneChart(pieSource, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("pie");
    expect(clone.axes).toBeUndefined();
  });

  it("filters out non-finite, zero, and negative tick spacings on inherit", () => {
    const dirty: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: {
        y: {
          scale: {
            min: 0,
            max: 100,
            majorUnit: Number.NaN,
            minorUnit: 0,
          } as never,
        },
      },
    };
    const clone = cloneChart(dirty, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes?.y?.scale).toEqual({ min: 0, max: 100 });
  });

  it("co-inherits the title, gridlines and scale on the same axis", () => {
    const all: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: {
        y: { title: "Revenue", gridlines: { major: true }, scale: { min: 0, max: 100 } },
      },
    };
    const clone = cloneChart(all, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes?.y).toEqual({
      title: "Revenue",
      gridlines: { major: true },
      scale: { min: 0, max: 100 },
    });
  });
});

// ── cloneChart — axis number format ─────────────────────────────────

describe("cloneChart — axis number format", () => {
  const sourceWithNumFmt: Chart = {
    kinds: ["bar"],
    seriesCount: 1,
    series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    axes: {
      y: { numberFormat: { formatCode: "$#,##0" } },
    },
  };

  it("inherits the source's number format when no override is given", () => {
    const clone = cloneChart(sourceWithNumFmt, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes?.y?.numberFormat).toEqual({ formatCode: "$#,##0" });
  });

  it("drops inherited number format when override is null", () => {
    const clone = cloneChart(sourceWithNumFmt, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { numberFormat: null } },
    });
    expect(clone.axes).toBeUndefined();
  });

  it("replaces inherited format with the override", () => {
    const clone = cloneChart(sourceWithNumFmt, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { numberFormat: { formatCode: "0.00%" } } },
    });
    expect(clone.axes?.y?.numberFormat).toEqual({ formatCode: "0.00%" });
  });

  it("adds a number format to an axis the source did not declare it on", () => {
    const noFmt: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
    };
    const clone = cloneChart(noFmt, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { numberFormat: { formatCode: "#,##0" } } },
    });
    expect(clone.axes?.y?.numberFormat).toEqual({ formatCode: "#,##0" });
  });

  it("preserves sourceLinked on inherit", () => {
    const linked: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { y: { numberFormat: { formatCode: "0.0", sourceLinked: true } } },
    };
    const clone = cloneChart(linked, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes?.y?.numberFormat).toEqual({ formatCode: "0.0", sourceLinked: true });
  });

  it("strips number format silently when the resolved chart type is pie", () => {
    const pieSource: Chart = {
      kinds: ["pie"],
      seriesCount: 1,
      series: [{ kind: "pie", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { y: { numberFormat: { formatCode: "$#,##0" } } },
    };
    const clone = cloneChart(pieSource, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("pie");
    expect(clone.axes).toBeUndefined();
  });

  it("ignores empty formatCode strings on both inherit and override", () => {
    const empty: Chart = {
      kinds: ["bar"],
      seriesCount: 1,
      series: [{ kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5" }],
      axes: { y: { numberFormat: { formatCode: "" } } },
    };
    const clone = cloneChart(empty, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.axes).toBeUndefined();

    const cloneOverride = cloneChart(sourceWithNumFmt, {
      anchor: { from: { row: 0, col: 0 } },
      axes: { y: { numberFormat: { formatCode: "" } } },
    });
    expect(cloneOverride.axes).toBeUndefined();
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

  it("round-trips axis gridlines through parseChart → cloneChart → writeXlsx → parseChart", async () => {
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
        <c:majorGridlines/>
      </c:catAx>
      <c:valAx>
        <c:axId val="222"/>
        <c:majorGridlines/>
        <c:minorGridlines/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml)!;
    expect(source.axes).toEqual({
      x: { gridlines: { major: true } },
      y: { gridlines: { major: true, minor: true } },
    });

    const sheetChart: SheetChart = cloneChart(source, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.axes).toEqual({
      x: { gridlines: { major: true } },
      y: { gridlines: { major: true, minor: true } },
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
    expect(written).toContain("c:majorGridlines");
    expect(written).toContain("c:minorGridlines");

    const reparsed = parseChart(written);
    expect(reparsed?.axes).toEqual({
      x: { gridlines: { major: true } },
      y: { gridlines: { major: true, minor: true } },
    });
  });

  it("round-trips line grouping through parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="percentStacked"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Revenue</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
      <c:catAx><c:axId val="111"/></c:catAx>
      <c:valAx><c:axId val="222"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml)!;
    expect(source.lineGrouping).toBe("percentStacked");

    const sheetChart: SheetChart = cloneChart(source, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.type).toBe("line");
    expect(sheetChart.lineGrouping).toBe("percentStacked");

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
    expect(written).toContain("c:lineChart");
    expect(written).toContain('c:grouping val="percentStacked"');

    const reparsed = parseChart(written);
    expect(reparsed?.lineGrouping).toBe("percentStacked");
  });

  it("round-trips area grouping through parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:grouping val="stacked"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Cloud</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:v>On-prem</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:areaChart>
      <c:catAx><c:axId val="111"/></c:catAx>
      <c:valAx><c:axId val="222"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml)!;
    expect(source.areaGrouping).toBe("stacked");

    const sheetChart: SheetChart = cloneChart(source, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }, { values: "Dashboard!$C$2:$C$5" }],
    });
    expect(sheetChart.type).toBe("area");
    expect(sheetChart.areaGrouping).toBe("stacked");

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [
            ["A", "B", "C"],
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
            [10, 11, 12],
          ],
          charts: [sheetChart],
        },
      ],
    });
    const zip = new ZipReader(xlsx);
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(written).toContain("c:areaChart");
    expect(written).toContain('c:grouping val="stacked"');

    const reparsed = parseChart(written);
    expect(reparsed?.areaGrouping).toBe("stacked");
  });

  it("clones a doughnut template through writeChart and back through parseChart with hole size intact", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Distribution</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Mix</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="65"/>
      </c:doughnutChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml);
    expect(source?.kinds).toEqual(["doughnut"]);
    expect(source?.holeSize).toBe(65);

    // Default clone keeps the doughnut shape and inherits holeSize from
    // the template.
    const sheetChart: SheetChart = cloneChart(source!, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.type).toBe("doughnut");
    expect(sheetChart.holeSize).toBe(65);

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
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(written).toContain("<c:doughnutChart>");
    expect(written).toContain('c:holeSize val="65"');

    // Re-read the emitted chart and confirm doughnut + holeSize survive.
    const reparsed = parseChart(written);
    expect(reparsed?.kinds).toEqual(["doughnut"]);
    expect(reparsed?.title).toBe("Distribution");
    expect(reparsed?.holeSize).toBe(65);
  });

  it("round-trips axis scale and number format through parseChart -> cloneChart -> writeXlsx", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="111111111"/>
        <c:axId val="222222222"/>
      </c:barChart>
      <c:catAx><c:axId val="111111111"/></c:catAx>
      <c:valAx>
        <c:axId val="222222222"/>
        <c:scaling>
          <c:orientation val="minMax"/>
          <c:max val="100"/>
          <c:min val="0"/>
        </c:scaling>
        <c:numFmt formatCode="$#,##0" sourceLinked="0"/>
        <c:majorUnit val="25"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml);
    expect(source?.axes?.y?.scale).toEqual({ min: 0, max: 100, majorUnit: 25 });
    expect(source?.axes?.y?.numberFormat).toEqual({ formatCode: "$#,##0" });

    // Default clone inherits scale + numberFormat off the template.
    const sheetChart: SheetChart = cloneChart(source!, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.axes?.y?.scale).toEqual({ min: 0, max: 100, majorUnit: 25 });
    expect(sheetChart.axes?.y?.numberFormat).toEqual({ formatCode: "$#,##0" });

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
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(written).toContain('<c:max val="100"/>');
    expect(written).toContain('<c:min val="0"/>');
    expect(written).toContain('<c:majorUnit val="25"/>');
    expect(written).toContain('formatCode="$#,##0"');

    // Re-read the emitted chart and confirm everything survives.
    const reparsed = parseChart(written);
    expect(reparsed?.axes?.y?.scale).toEqual({ min: 0, max: 100, majorUnit: 25 });
    expect(reparsed?.axes?.y?.numberFormat).toEqual({ formatCode: "$#,##0" });
  });

  it("round-trips gapWidth & overlap through parseChart -> cloneChart -> writeXlsx -> parseChart", async () => {
    // A pinned bar template with a tight 50% group gap and a small
    // negative overlap (series slightly separated within each group).
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Revenue</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="50"/>
        <c:overlap val="-25"/>
        <c:axId val="111111111"/>
        <c:axId val="222222222"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="111111111"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="222222222"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="222222222"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="111111111"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml);
    expect(source?.gapWidth).toBe(50);
    expect(source?.overlap).toBe(-25);

    const sheetChart: SheetChart = cloneChart(source!, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.type).toBe("column");
    expect(sheetChart.gapWidth).toBe(50);
    expect(sheetChart.overlap).toBe(-25);

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
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(written).toContain('c:gapWidth val="50"');
    expect(written).toContain('c:overlap val="-25"');

    const reparsed = parseChart(written);
    expect(reparsed?.kinds).toEqual(["bar"]);
    expect(reparsed?.gapWidth).toBe(50);
    expect(reparsed?.overlap).toBe(-25);
  });

  it("round-trips firstSliceAng through parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Mix</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="135"/>
        <c:holeSize val="55"/>
      </c:doughnutChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml);
    expect(source?.firstSliceAng).toBe(135);

    const sheetChart: SheetChart = cloneChart(source!, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }],
    });
    expect(sheetChart.type).toBe("doughnut");
    expect(sheetChart.firstSliceAng).toBe(135);

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
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(written).toContain('c:firstSliceAng val="135"');

    const reparsed = parseChart(written);
    expect(reparsed?.kinds).toEqual(["doughnut"]);
    expect(reparsed?.firstSliceAng).toBe(135);
    expect(reparsed?.holeSize).toBe(55);
  });
});

// ── cloneChart — series smooth flag ─────────────────────────────────

describe("cloneChart — series smooth flag", () => {
  function lineSource(smooth: boolean | undefined): Chart {
    return {
      kinds: ["line"],
      seriesCount: 1,
      series: [
        {
          kind: "line",
          index: 0,
          name: "Revenue",
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          ...(smooth !== undefined ? { smooth } : {}),
        },
      ],
    };
  }

  it("inherits smooth=true from a line series source", () => {
    const clone = cloneChart(lineSource(true), { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("line");
    expect(clone.series[0].smooth).toBe(true);
  });

  it("does not surface smooth when the source series did not declare it", () => {
    const clone = cloneChart(lineSource(undefined), { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].smooth).toBeUndefined();
  });

  it("lets seriesOverrides[i].smooth=true override a source missing the flag", () => {
    const clone = cloneChart(lineSource(undefined), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ smooth: true }],
    });
    expect(clone.series[0].smooth).toBe(true);
  });

  it("lets seriesOverrides[i].smooth=null drop an inherited smooth flag", () => {
    const clone = cloneChart(lineSource(true), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ smooth: null }],
    });
    expect(clone.series[0].smooth).toBeUndefined();
  });

  it("lets seriesOverrides[i].smooth=false drop an inherited smooth flag", () => {
    // `false` collapses to undefined for symmetry with the OOXML
    // default — straight segments and absence round-trip identically.
    const clone = cloneChart(lineSource(true), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ smooth: false }],
    });
    expect(clone.series[0].smooth).toBeUndefined();
  });

  it("carries smooth onto a scatter clone", () => {
    const scatterSource: Chart = {
      kinds: ["scatter"],
      seriesCount: 1,
      series: [
        {
          kind: "scatter",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          smooth: true,
        },
      ],
    };
    const clone = cloneChart(scatterSource, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("scatter");
    expect(clone.series[0].smooth).toBe(true);
  });

  it("drops inherited smooth when the resolved type is not line/scatter", () => {
    // A line template flattened to area / column / pie / doughnut must
    // not leak <c:smooth> — the OOXML schema rejects it on every other
    // chart family.
    for (const type of ["column", "bar", "pie", "doughnut", "area"] as const) {
      const clone = cloneChart(lineSource(true), {
        anchor: { from: { row: 0, col: 0 } },
        type,
        seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
      });
      expect(clone.type).toBe(type);
      expect(clone.series[0].smooth).toBeUndefined();
    }
  });

  it("drops smooth from explicit options.series when the resolved type is not line/scatter", () => {
    // Replacing the entire series array via options.series still goes
    // through the post-build smooth-strip, so a stray smooth flag does
    // not leak into a non-line/scatter target.
    const clone = cloneChart(lineSource(true), {
      anchor: { from: { row: 0, col: 0 } },
      type: "column",
      series: [{ values: "Sheet1!$B$2:$B$5", smooth: true }],
    });
    expect(clone.series[0].smooth).toBeUndefined();
  });

  it("propagates smooth across a multi-series line clone", () => {
    const multi: Chart = {
      kinds: ["line"],
      seriesCount: 3,
      series: [
        { kind: "line", index: 0, valuesRef: "Tpl!$B$2:$B$5", smooth: true },
        { kind: "line", index: 1, valuesRef: "Tpl!$C$2:$C$5" },
        { kind: "line", index: 2, valuesRef: "Tpl!$D$2:$D$5", smooth: true },
      ],
    };
    const clone = cloneChart(multi, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].smooth).toBe(true);
    expect(clone.series[1].smooth).toBeUndefined();
    expect(clone.series[2].smooth).toBe(true);
  });

  it("round-trips smooth through parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Curved</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
          <c:smooth val="1"/>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:v>Straight</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
      <c:catAx><c:axId val="111"/></c:catAx>
      <c:valAx><c:axId val="222"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml)!;
    expect(source.series?.[0].smooth).toBe(true);
    expect(source.series?.[1].smooth).toBeUndefined();

    const sheetChart: SheetChart = cloneChart(source, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }, { values: "Dashboard!$C$2:$C$5" }],
    });
    expect(sheetChart.type).toBe("line");
    expect(sheetChart.series[0].smooth).toBe(true);
    expect(sheetChart.series[1].smooth).toBeUndefined();

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [
            ["A", "B", "C"],
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
            [10, 11, 12],
          ],
          charts: [sheetChart],
        },
      ],
    });
    const zip = new ZipReader(xlsx);
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    // First series surfaces smooth=1, second falls back to the default 0.
    const matches = written.match(/c:smooth val="[01]"/g) ?? [];
    expect(matches).toEqual(['c:smooth val="1"', 'c:smooth val="0"']);

    const reparsed = parseChart(written);
    expect(reparsed?.series?.[0].smooth).toBe(true);
    expect(reparsed?.series?.[1].smooth).toBeUndefined();
  });
});

// ── cloneChart — series line stroke ─────────────────────────────────

describe("cloneChart — series line stroke", () => {
  function lineSource(stroke?: ChartLineStroke): Chart {
    return {
      kinds: ["line"],
      seriesCount: 1,
      series: [
        {
          kind: "line",
          index: 0,
          name: "Revenue",
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          ...(stroke ? { stroke } : {}),
        },
      ],
    };
  }

  it("inherits the stroke block from a line series source", () => {
    const source = lineSource({ dash: "dash", width: 2.5 });
    const clone = cloneChart(source, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("line");
    expect(clone.series[0].stroke).toEqual({ dash: "dash", width: 2.5 });
  });

  it("does not surface stroke when the source series did not declare one", () => {
    const clone = cloneChart(lineSource(), { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].stroke).toBeUndefined();
  });

  it("lets seriesOverrides[i].stroke replace an inherited block wholesale", () => {
    const source = lineSource({ dash: "dash", width: 2.5 });
    const clone = cloneChart(source, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ stroke: { dash: "dot", width: 0.5 } }],
    });
    // Override replaces wholesale; old width does not leak through.
    expect(clone.series[0].stroke).toEqual({ dash: "dot", width: 0.5 });
  });

  it("lets seriesOverrides[i].stroke=null drop an inherited stroke block", () => {
    const source = lineSource({ dash: "dash", width: 2.5 });
    const clone = cloneChart(source, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ stroke: null }],
    });
    expect(clone.series[0].stroke).toBeUndefined();
  });

  it("lets seriesOverrides[i].stroke={} collapse to undefined", () => {
    // An empty stroke carries no meaningful settings; the writer will
    // never emit `<a:ln>` for it, so the resolver collapses it to
    // undefined to keep the round-trip shape minimal.
    const source = lineSource({ dash: "dash" });
    const clone = cloneChart(source, {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ stroke: {} }],
    });
    expect(clone.series[0].stroke).toBeUndefined();
  });

  it("carries stroke onto a scatter clone", () => {
    const source: Chart = {
      kinds: ["scatter"],
      seriesCount: 1,
      series: [
        {
          kind: "scatter",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          stroke: { dash: "lgDashDot", width: 1 },
        },
      ],
    };
    const clone = cloneChart(source, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("scatter");
    expect(clone.series[0].stroke).toEqual({ dash: "lgDashDot", width: 1 });
  });

  it("drops inherited stroke when the resolved type is not line/scatter", () => {
    // A clone that flattens a line template into a column / pie / area
    // chart must not leak <a:ln> styling — the OOXML schema rejects it
    // on every other family that does not paint a connecting line.
    const types: Array<"column" | "bar" | "pie" | "doughnut" | "area"> = [
      "column",
      "bar",
      "pie",
      "doughnut",
      "area",
    ];
    for (const type of types) {
      const clone = cloneChart(lineSource({ dash: "dash" }), {
        anchor: { from: { row: 0, col: 0 } },
        type,
      });
      expect(clone.series[0].stroke).toBeUndefined();
    }
  });

  it("drops stroke from explicit options.series when the resolved type is not line/scatter", () => {
    // Even when the caller bypasses seriesOverrides and passes a fully
    // built `series` array, a stroke field must not leak into a chart
    // family whose schema rejects the element. The post-build sweep
    // strips it after the merge.
    const clone = cloneChart(lineSource(), {
      anchor: { from: { row: 0, col: 0 } },
      type: "column",
      series: [{ values: "Sheet1!$B$2:$B$5", stroke: { dash: "dot" } }],
    });
    expect(clone.series[0].stroke).toBeUndefined();
  });

  it("propagates stroke across a multi-series line clone", () => {
    const source: Chart = {
      kinds: ["line"],
      seriesCount: 3,
      series: [
        { kind: "line", index: 0, valuesRef: "Tpl!$B$2:$B$5", stroke: { dash: "dash" } },
        { kind: "line", index: 1, valuesRef: "Tpl!$C$2:$C$5" },
        { kind: "line", index: 2, valuesRef: "Tpl!$D$2:$D$5", stroke: { width: 2.5 } },
      ],
    };
    const clone = cloneChart(source, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series).toHaveLength(3);
    expect(clone.series[0].stroke).toEqual({ dash: "dash" });
    expect(clone.series[1].stroke).toBeUndefined();
    expect(clone.series[2].stroke).toEqual({ width: 2.5 });
  });

  it("survives a parseChart → cloneChart → writeChart → parseChart round-trip", () => {
    const NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
    const source = parseChart(`<c:chartSpace ${NS}>
  <c:chart><c:plotArea>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:ser>
        <c:idx val="0"/>
        <c:spPr>
          <a:ln w="31750">
            <a:prstDash val="dashDot"/>
          </a:ln>
        </c:spPr>
        <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
      </c:ser>
    </c:lineChart>
  </c:plotArea></c:chart>
</c:chartSpace>`);
    const clone = cloneChart(source!, { anchor: { from: { row: 0, col: 0 } } });
    const written = writeChart(clone, "Sheet1").chartXml;
    expect(written).toContain('<a:prstDash val="dashDot"/>');
    expect(written).toContain('w="31750"');

    const reparsed = parseChart(written);
    expect(reparsed?.series?.[0].stroke).toEqual({ dash: "dashDot", width: 2.5 });
  });
});

// ── cloneChart — series marker ──────────────────────────────────────

describe("cloneChart — series marker", () => {
  function lineSource(marker?: ChartMarker): Chart {
    return {
      kinds: ["line"],
      seriesCount: 1,
      series: [
        {
          kind: "line",
          index: 0,
          name: "Revenue",
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          ...(marker ? { marker } : {}),
        },
      ],
    };
  }

  it("inherits the marker block from a line series source", () => {
    const clone = cloneChart(
      lineSource({ symbol: "diamond", size: 10, fill: "1F77B4", line: "0F3F60" }),
      { anchor: { from: { row: 0, col: 0 } } },
    );
    expect(clone.type).toBe("line");
    expect(clone.series[0].marker).toEqual({
      symbol: "diamond",
      size: 10,
      fill: "1F77B4",
      line: "0F3F60",
    });
  });

  it("does not surface marker when the source series did not declare one", () => {
    const clone = cloneChart(lineSource(), { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].marker).toBeUndefined();
  });

  it("lets seriesOverrides[i].marker replace an inherited block wholesale", () => {
    const clone = cloneChart(lineSource({ symbol: "circle", size: 6, fill: "1F77B4" }), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ marker: { symbol: "square", size: 8 } }],
    });
    // No per-field merge — the override replaces the inherited block,
    // so the inherited fill is dropped along with the inherited symbol.
    expect(clone.series[0].marker).toEqual({ symbol: "square", size: 8 });
  });

  it("lets seriesOverrides[i].marker=null drop an inherited marker block", () => {
    const clone = cloneChart(lineSource({ symbol: "diamond" }), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ marker: null }],
    });
    expect(clone.series[0].marker).toBeUndefined();
  });

  it("lets seriesOverrides[i].marker={} collapse to undefined", () => {
    // An empty marker carries no meaningful settings; the writer will
    // never emit a `<c:marker>` for it, so the resolver collapses it to
    // undefined to keep the materialized SheetChart honest.
    const clone = cloneChart(lineSource({ symbol: "diamond" }), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ marker: {} }],
    });
    expect(clone.series[0].marker).toBeUndefined();
  });

  it("carries marker onto a scatter clone", () => {
    const scatterSource: Chart = {
      kinds: ["scatter"],
      seriesCount: 1,
      series: [
        {
          kind: "scatter",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          marker: { symbol: "x", size: 8 },
        },
      ],
    };
    const clone = cloneChart(scatterSource, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("scatter");
    expect(clone.series[0].marker).toEqual({ symbol: "x", size: 8 });
  });

  it("drops inherited marker when the resolved type is not line/scatter", () => {
    // A line template flattened to area / column / pie / doughnut must
    // not leak <c:marker> — the OOXML schema rejects it on every other
    // chart family's series element.
    for (const type of ["column", "bar", "pie", "doughnut", "area"] as const) {
      const clone = cloneChart(lineSource({ symbol: "diamond", size: 10 }), {
        anchor: { from: { row: 0, col: 0 } },
        type,
        seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
      });
      expect(clone.type).toBe(type);
      expect(clone.series[0].marker).toBeUndefined();
    }
  });

  it("drops marker from explicit options.series when the resolved type is not line/scatter", () => {
    // Replacing the entire series array via options.series still goes
    // through the post-build marker-strip, so a stray marker does not
    // leak into a non-line/scatter target.
    const clone = cloneChart(lineSource({ symbol: "diamond" }), {
      anchor: { from: { row: 0, col: 0 } },
      type: "column",
      series: [{ values: "Sheet1!$B$2:$B$5", marker: { symbol: "circle" } }],
    });
    expect(clone.series[0].marker).toBeUndefined();
  });

  it("propagates marker across a multi-series line clone", () => {
    const multi: Chart = {
      kinds: ["line"],
      seriesCount: 3,
      series: [
        {
          kind: "line",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          marker: { symbol: "circle", size: 6 },
        },
        { kind: "line", index: 1, valuesRef: "Tpl!$C$2:$C$5" },
        { kind: "line", index: 2, valuesRef: "Tpl!$D$2:$D$5", marker: { symbol: "square" } },
      ],
    };
    const clone = cloneChart(multi, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].marker).toEqual({ symbol: "circle", size: 6 });
    expect(clone.series[1].marker).toBeUndefined();
    expect(clone.series[2].marker).toEqual({ symbol: "square" });
  });

  it("returns a fresh marker object so callers cannot mutate the parsed source", () => {
    const sourceMarker = { symbol: "circle" as const, size: 6 };
    const src: Chart = {
      kinds: ["line"],
      seriesCount: 1,
      series: [
        {
          kind: "line",
          index: 0,
          valuesRef: "Tpl!$B$2:$B$5",
          marker: sourceMarker,
        },
      ],
    };
    const clone = cloneChart(src, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].marker).not.toBe(sourceMarker);
    // Mutating the clone does not bleed back into the parsed source.
    if (clone.series[0].marker) clone.series[0].marker.size = 99;
    expect(sourceMarker.size).toBe(6);
  });

  it("round-trips marker through parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Diamonds</c:v></c:tx>
          <c:marker>
            <c:symbol val="diamond"/>
            <c:size val="10"/>
            <c:spPr>
              <a:solidFill><a:srgbClr val="1F77B4"/></a:solidFill>
              <a:ln><a:solidFill><a:srgbClr val="0F3F60"/></a:solidFill></a:ln>
            </c:spPr>
          </c:marker>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:v>Bare</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
      <c:catAx><c:axId val="111"/></c:catAx>
      <c:valAx><c:axId val="222"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml)!;
    expect(source.series?.[0].marker).toEqual({
      symbol: "diamond",
      size: 10,
      fill: "1F77B4",
      line: "0F3F60",
    });
    expect(source.series?.[1].marker).toBeUndefined();

    const sheetChart: SheetChart = cloneChart(source, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }, { values: "Dashboard!$C$2:$C$5" }],
    });
    expect(sheetChart.type).toBe("line");
    expect(sheetChart.series[0].marker).toEqual({
      symbol: "diamond",
      size: 10,
      fill: "1F77B4",
      line: "0F3F60",
    });
    expect(sheetChart.series[1].marker).toBeUndefined();

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [
            ["A", "B", "C"],
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
            [10, 11, 12],
          ],
          charts: [sheetChart],
        },
      ],
    });
    const zip = new ZipReader(xlsx);
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    // First series gets a full marker block; second has none at the
    // series level.
    const markerBlocks = written.match(/<c:marker>[\s\S]*?<\/c:marker>/g) ?? [];
    expect(markerBlocks).toHaveLength(1);
    expect(markerBlocks[0]).toContain('c:symbol val="diamond"');
    expect(markerBlocks[0]).toContain('c:size val="10"');
    expect(markerBlocks[0]).toContain('a:srgbClr val="1F77B4"');
    expect(markerBlocks[0]).toContain('a:srgbClr val="0F3F60"');

    const reparsed = parseChart(written);
    expect(reparsed?.series?.[0].marker).toEqual({
      symbol: "diamond",
      size: 10,
      fill: "1F77B4",
      line: "0F3F60",
    });
    expect(reparsed?.series?.[1].marker).toBeUndefined();
  });
});

// ── cloneChart — series invertIfNegative flag ───────────────────────

describe("cloneChart — series invertIfNegative flag", () => {
  function barSource(invertIfNegative: boolean | undefined): Chart {
    return {
      kinds: ["bar"],
      seriesCount: 1,
      series: [
        {
          kind: "bar",
          index: 0,
          name: "Revenue",
          valuesRef: "Tpl!$B$2:$B$5",
          categoriesRef: "Tpl!$A$2:$A$5",
          ...(invertIfNegative !== undefined ? { invertIfNegative } : {}),
        },
      ],
    };
  }

  it("inherits invertIfNegative=true from a bar series source", () => {
    const clone = cloneChart(barSource(true), { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.type).toBe("column");
    expect(clone.series[0].invertIfNegative).toBe(true);
  });

  it("does not surface invertIfNegative when the source series did not declare it", () => {
    const clone = cloneChart(barSource(undefined), { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].invertIfNegative).toBeUndefined();
  });

  it("lets seriesOverrides[i].invertIfNegative=true override a source missing the flag", () => {
    const clone = cloneChart(barSource(undefined), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ invertIfNegative: true }],
    });
    expect(clone.series[0].invertIfNegative).toBe(true);
  });

  it("lets seriesOverrides[i].invertIfNegative=null drop an inherited flag", () => {
    const clone = cloneChart(barSource(true), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ invertIfNegative: null }],
    });
    expect(clone.series[0].invertIfNegative).toBeUndefined();
  });

  it("lets seriesOverrides[i].invertIfNegative=false drop an inherited flag", () => {
    // `false` collapses to undefined for symmetry with the OOXML
    // default — non-inverted bars and absence round-trip identically.
    const clone = cloneChart(barSource(true), {
      anchor: { from: { row: 0, col: 0 } },
      seriesOverrides: [{ invertIfNegative: false }],
    });
    expect(clone.series[0].invertIfNegative).toBeUndefined();
  });

  it("carries invertIfNegative onto a horizontal bar clone", () => {
    const clone = cloneChart(barSource(true), {
      anchor: { from: { row: 0, col: 0 } },
      type: "bar",
    });
    expect(clone.type).toBe("bar");
    expect(clone.series[0].invertIfNegative).toBe(true);
  });

  it("drops inherited invertIfNegative when the resolved type is not bar/column", () => {
    // A bar template flattened to line / pie / doughnut / area /
    // scatter must not leak <c:invertIfNegative> — the OOXML schema
    // rejects it on every other chart family.
    for (const type of ["line", "pie", "doughnut", "area", "scatter"] as const) {
      const clone = cloneChart(barSource(true), {
        anchor: { from: { row: 0, col: 0 } },
        type,
        seriesOverrides: [{ values: "Sheet1!$B$2:$B$5" }],
      });
      expect(clone.type).toBe(type);
      expect(clone.series[0].invertIfNegative).toBeUndefined();
    }
  });

  it("drops invertIfNegative from explicit options.series when the resolved type is not bar/column", () => {
    // Replacing the entire series array via options.series still goes
    // through the post-build invert-strip, so a stray flag does not
    // leak into a non-bar/column target.
    const clone = cloneChart(barSource(true), {
      anchor: { from: { row: 0, col: 0 } },
      type: "line",
      series: [{ values: "Sheet1!$B$2:$B$5", invertIfNegative: true }],
    });
    expect(clone.series[0].invertIfNegative).toBeUndefined();
  });

  it("propagates invertIfNegative across a multi-series column clone", () => {
    const multi: Chart = {
      kinds: ["bar"],
      seriesCount: 3,
      series: [
        { kind: "bar", index: 0, valuesRef: "Tpl!$B$2:$B$5", invertIfNegative: true },
        { kind: "bar", index: 1, valuesRef: "Tpl!$C$2:$C$5" },
        { kind: "bar", index: 2, valuesRef: "Tpl!$D$2:$D$5", invertIfNegative: true },
      ],
    };
    const clone = cloneChart(multi, { anchor: { from: { row: 0, col: 0 } } });
    expect(clone.series[0].invertIfNegative).toBe(true);
    expect(clone.series[1].invertIfNegative).toBeUndefined();
    expect(clone.series[2].invertIfNegative).toBe(true);
  });

  it("round-trips invertIfNegative through parseChart → cloneChart → writeXlsx → parseChart", async () => {
    const sourceXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Inverted</c:v></c:tx>
          <c:invertIfNegative val="1"/>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:v>Default</c:v></c:tx>
          <c:cat><c:strRef><c:f>Tpl!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Tpl!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/>
        <c:axId val="222"/>
      </c:barChart>
      <c:catAx><c:axId val="111"/></c:catAx>
      <c:valAx><c:axId val="222"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
    const source = parseChart(sourceXml)!;
    expect(source.series?.[0].invertIfNegative).toBe(true);
    expect(source.series?.[1].invertIfNegative).toBeUndefined();

    const sheetChart: SheetChart = cloneChart(source, {
      anchor: { from: { row: 14, col: 0 } },
      seriesOverrides: [{ values: "Dashboard!$B$2:$B$5" }, { values: "Dashboard!$C$2:$C$5" }],
    });
    expect(sheetChart.type).toBe("column");
    expect(sheetChart.series[0].invertIfNegative).toBe(true);
    expect(sheetChart.series[1].invertIfNegative).toBeUndefined();

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Dashboard",
          rows: [
            ["A", "B", "C"],
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
            [10, 11, 12],
          ],
          charts: [sheetChart],
        },
      ],
    });
    const zip = new ZipReader(xlsx);
    const written = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    // Only the inverted series carries <c:invertIfNegative>; the
    // second falls back to the OOXML default (absence of the element).
    const matches = written.match(/c:invertIfNegative val="[01]"/g) ?? [];
    expect(matches).toEqual(['c:invertIfNegative val="1"']);

    const reparsed = parseChart(written);
    expect(reparsed?.series?.[0].invertIfNegative).toBe(true);
    expect(reparsed?.series?.[1].invertIfNegative).toBeUndefined();
  });
});
