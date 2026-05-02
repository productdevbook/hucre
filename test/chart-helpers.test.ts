import { describe, expect, it } from "vitest";
import { addChart, getCharts } from "../src/xlsx/chart-helpers";
import { writeXlsx } from "../src/xlsx/writer";
import { ZipReader } from "../src/zip/reader";
import type { Chart, SheetChart, Workbook, WriteSheet } from "../src/_types";

const decoder = new TextDecoder("utf-8");

// ── getCharts ────────────────────────────────────────────────────────

describe("getCharts", () => {
  function makeChart(title: string): Chart {
    return {
      kinds: ["bar"],
      seriesCount: 1,
      title,
      series: [
        {
          kind: "bar",
          index: 0,
          name: title,
          valuesRef: "Sheet1!$B$2:$B$5",
        },
      ],
    };
  }

  it("returns an empty array when the workbook has no sheets", () => {
    expect(getCharts({ sheets: [] })).toEqual([]);
  });

  it("returns an empty array when no sheet has charts", () => {
    const workbook: Workbook = {
      sheets: [
        { name: "S1", rows: [[1]] },
        { name: "S2", rows: [[2]] },
      ],
    };
    expect(getCharts(workbook)).toEqual([]);
  });

  it("skips sheets whose `charts` is the empty array", () => {
    const workbook: Workbook = {
      sheets: [
        { name: "Empty", rows: [], charts: [] },
        { name: "Has", rows: [], charts: [makeChart("only")] },
      ],
    };
    const found = getCharts(workbook);
    expect(found).toHaveLength(1);
    expect(found[0].sheetName).toBe("Has");
    expect(found[0].sheetIndex).toBe(1);
    expect(found[0].chartIndex).toBe(0);
  });

  it("walks sheets in workbook order and charts in their declared order", () => {
    const a1 = makeChart("A1");
    const a2 = makeChart("A2");
    const b1 = makeChart("B1");
    const workbook: Workbook = {
      sheets: [
        { name: "Alpha", rows: [], charts: [a1, a2] },
        { name: "Beta", rows: [], charts: [b1] },
      ],
    };

    const found = getCharts(workbook);
    expect(found).toHaveLength(3);

    expect(found[0]).toMatchObject({
      sheetName: "Alpha",
      sheetIndex: 0,
      chartIndex: 0,
    });
    expect(found[0].chart).toBe(a1);
    expect(found[0].sheet).toBe(workbook.sheets[0]);

    expect(found[1]).toMatchObject({
      sheetName: "Alpha",
      sheetIndex: 0,
      chartIndex: 1,
    });
    expect(found[1].chart).toBe(a2);

    expect(found[2]).toMatchObject({
      sheetName: "Beta",
      sheetIndex: 1,
      chartIndex: 0,
    });
    expect(found[2].chart).toBe(b1);
  });

  it("preserves identity — the returned `chart` and `sheet` are the same references", () => {
    const chart = makeChart("Identity");
    const sheet = { name: "S", rows: [], charts: [chart] };
    const workbook: Workbook = { sheets: [sheet] };

    const [loc] = getCharts(workbook);
    expect(loc.sheet).toBe(sheet);
    expect(loc.chart).toBe(chart);
    // Mutating the returned reference flows back to the workbook.
    loc.chart.title = "Mutated";
    expect(chart.title).toBe("Mutated");
  });

  it("handles a workbook where only the last sheet carries charts", () => {
    const tail = makeChart("Tail");
    const workbook: Workbook = {
      sheets: [
        { name: "S1", rows: [] },
        { name: "S2", rows: [] },
        { name: "S3", rows: [], charts: [tail] },
      ],
    };
    const found = getCharts(workbook);
    expect(found).toHaveLength(1);
    expect(found[0].sheetIndex).toBe(2);
    expect(found[0].sheetName).toBe("S3");
    expect(found[0].chart).toBe(tail);
  });
});

// ── addChart ─────────────────────────────────────────────────────────

describe("addChart", () => {
  function validChart(): SheetChart {
    return {
      type: "column",
      title: "Revenue",
      series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
      anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
    };
  }

  it("creates the `charts` array on first call", () => {
    const sheet: WriteSheet = { name: "Dashboard" };
    expect(sheet.charts).toBeUndefined();
    const chart = validChart();
    addChart(sheet, chart);
    expect(sheet.charts).toEqual([chart]);
  });

  it("appends to an existing `charts` array in order", () => {
    const c1 = validChart();
    c1.title = "First";
    const c2 = validChart();
    c2.title = "Second";

    const sheet: WriteSheet = { name: "S", charts: [c1] };
    addChart(sheet, c2);

    expect(sheet.charts).toHaveLength(2);
    expect(sheet.charts?.[0].title).toBe("First");
    expect(sheet.charts?.[1].title).toBe("Second");
  });

  it("returns the chart instance for inline use", () => {
    const sheet: WriteSheet = { name: "S" };
    const chart = validChart();
    const returned = addChart(sheet, chart);
    expect(returned).toBe(chart);
  });

  it("rejects a missing chart argument", () => {
    const sheet: WriteSheet = { name: "S" };
    // @ts-expect-error — testing runtime guard
    expect(() => addChart(sheet, undefined)).toThrow(/chart is required/);
    // @ts-expect-error — testing runtime guard for non-object input
    expect(() => addChart(sheet, 42)).toThrow(/chart is required/);
  });

  it("rejects a chart that is missing required fields", () => {
    const sheet: WriteSheet = { name: "S" };
    expect(() =>
      addChart(sheet, {
        // @ts-expect-error — missing type on purpose
        type: undefined,
        series: [{ values: "A1:A2" }],
        anchor: { from: { row: 0, col: 0 } },
      }),
    ).toThrow(/chart\.type/);

    expect(() =>
      addChart(sheet, {
        type: "column",
        // @ts-expect-error — missing series on purpose
        series: undefined,
        anchor: { from: { row: 0, col: 0 } },
      }),
    ).toThrow(/chart\.series/);

    expect(() =>
      addChart(sheet, {
        type: "column",
        series: [],
        anchor: { from: { row: 0, col: 0 } },
      }),
    ).toThrow(/chart\.series/);

    expect(() =>
      addChart(sheet, {
        type: "column",
        series: [{ values: "A1:A2" }],
        // @ts-expect-error — missing anchor on purpose
        anchor: undefined,
      }),
    ).toThrow(/chart\.anchor/);

    expect(() =>
      addChart(sheet, {
        type: "column",
        series: [{ values: "A1:A2" }],
        // @ts-expect-error — anchor.from is required
        anchor: {},
      }),
    ).toThrow(/chart\.anchor/);
  });

  it("end-to-end — addChart → writeXlsx emits a chart part", async () => {
    const sheet: WriteSheet = {
      name: "Dashboard",
      rows: [
        ["Quarter", "Revenue"],
        ["Q1", 12_000],
        ["Q2", 15_500],
        ["Q3", 14_000],
      ],
    };

    addChart(sheet, {
      type: "column",
      title: "Revenue",
      series: [{ name: "Revenue", values: "B2:B4", categories: "A2:A4" }],
      anchor: { from: { row: 5, col: 0 }, to: { row: 20, col: 6 } },
    });

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const zip = new ZipReader(xlsx);
    expect(zip.has("xl/charts/chart1.xml")).toBe(true);
    const chartXml = decoder.decode(await zip.extract("xl/charts/chart1.xml"));
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("Revenue");
  });
});

// ── interop with cloneChart / parseChart ─────────────────────────────

describe("getCharts + addChart compose with the rest of the chart API", () => {
  it("walks a workbook with multiple sheets and surfaces every chart", () => {
    const workbook: Workbook = {
      sheets: [
        {
          name: "Sales",
          rows: [],
          charts: [
            { kinds: ["bar"], seriesCount: 1, title: "Sales-bar" },
            { kinds: ["line"], seriesCount: 2, title: "Sales-line" },
          ],
        },
        {
          name: "Marketing",
          rows: [],
          charts: [{ kinds: ["pie"], seriesCount: 1, title: "Marketing-pie" }],
        },
        // Sheet without charts shouldn't show up.
        { name: "Notes", rows: [["x"]] },
      ],
    };

    const titles = getCharts(workbook).map((loc) => loc.chart.title);
    expect(titles).toEqual(["Sales-bar", "Sales-line", "Marketing-pie"]);
  });
});
