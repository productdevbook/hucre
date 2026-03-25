import { describe, it, expect } from "vitest";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { createStylesCollector } from "../src/xlsx/styles-writer";
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer";
import type { WriteSheet } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector();
  const ss = createSharedStrings();
  const result = writeWorksheetXml(sheet, styles, ss);
  return result.xml;
}

function parseSheet(xml: string) {
  return parseXml(xml);
}

// ── Split Pane Writing Tests ─────────────────────────────────────────

describe("split panes — writing", () => {
  it("writes pane with state='split' for split pane", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      splitPane: { xSplit: 6000, ySplit: 3000 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");

    expect(pane).toBeDefined();
    expect(pane.attrs["state"]).toBe("split");
    expect(pane.attrs["xSplit"]).toBe("6000");
    expect(pane.attrs["ySplit"]).toBe("3000");
    expect(pane.attrs["activePane"]).toBe("bottomRight");
  });

  it("writes split pane with xSplit only", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      splitPane: { xSplit: 4500 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");

    expect(pane).toBeDefined();
    expect(pane.attrs["state"]).toBe("split");
    expect(pane.attrs["xSplit"]).toBe("4500");
    expect(pane.attrs["ySplit"]).toBeUndefined();
    expect(pane.attrs["activePane"]).toBe("topRight");
  });

  it("writes split pane with ySplit only", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      splitPane: { ySplit: 2000 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");

    expect(pane).toBeDefined();
    expect(pane.attrs["state"]).toBe("split");
    expect(pane.attrs["ySplit"]).toBe("2000");
    expect(pane.attrs["xSplit"]).toBeUndefined();
    expect(pane.attrs["activePane"]).toBe("bottomLeft");
  });

  it("writes topLeftCell as A1 for split pane", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      splitPane: { xSplit: 5000, ySplit: 3000 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");

    expect(pane.attrs["topLeftCell"]).toBe("A1");
  });
});

// ── Freeze Pane still works (regression) ─────────────────────────────

describe("split panes — freeze pane regression", () => {
  it("writes pane with state='frozen' for freeze pane", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      freezePane: { rows: 1, columns: 2 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");

    expect(pane).toBeDefined();
    expect(pane.attrs["state"]).toBe("frozen");
    expect(pane.attrs["xSplit"]).toBe("2");
    expect(pane.attrs["ySplit"]).toBe("1");
    expect(pane.attrs["activePane"]).toBe("bottomRight");
  });

  it("freeze pane takes precedence over split pane", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      freezePane: { rows: 1 },
      splitPane: { xSplit: 5000 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");

    expect(pane).toBeDefined();
    expect(pane.attrs["state"]).toBe("frozen");
    expect(pane.attrs["ySplit"]).toBe("1");
    // xSplit should NOT be set (freeze takes precedence)
    expect(pane.attrs["xSplit"]).toBeUndefined();
  });
});

// ── Round-trip Tests ─────────────────────────────────────────────────

describe("split panes — round-trip", () => {
  it("round-trips split pane with both xSplit and ySplit", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          splitPane: { xSplit: 6000, ySplit: 3000 },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].splitPane).toBeDefined();
    expect(workbook.sheets[0].splitPane!.xSplit).toBe(6000);
    expect(workbook.sheets[0].splitPane!.ySplit).toBe(3000);
    // Should NOT have freezePane
    expect(workbook.sheets[0].freezePane).toBeUndefined();
  });

  it("round-trips split pane with xSplit only", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          splitPane: { xSplit: 4500 },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].splitPane).toBeDefined();
    expect(workbook.sheets[0].splitPane!.xSplit).toBe(4500);
    expect(workbook.sheets[0].splitPane!.ySplit).toBeUndefined();
  });

  it("round-trips split pane with ySplit only", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          splitPane: { ySplit: 2000 },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].splitPane).toBeDefined();
    expect(workbook.sheets[0].splitPane!.ySplit).toBe(2000);
    expect(workbook.sheets[0].splitPane!.xSplit).toBeUndefined();
  });

  it("round-trips freeze pane (regression)", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          freezePane: { rows: 2, columns: 1 },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].freezePane).toBeDefined();
    expect(workbook.sheets[0].freezePane!.rows).toBe(2);
    expect(workbook.sheets[0].freezePane!.columns).toBe(1);
    // Should NOT have splitPane
    expect(workbook.sheets[0].splitPane).toBeUndefined();
  });

  it("round-trips freeze pane with rows only", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          freezePane: { rows: 3 },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].freezePane).toBeDefined();
    expect(workbook.sheets[0].freezePane!.rows).toBe(3);
    expect(workbook.sheets[0].freezePane!.columns).toBeUndefined();
  });

  it("round-trips freeze pane with columns only", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          freezePane: { columns: 2 },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].freezePane).toBeDefined();
    expect(workbook.sheets[0].freezePane!.columns).toBe(2);
    expect(workbook.sheets[0].freezePane!.rows).toBeUndefined();
  });

  it("no pane data when neither freeze nor split", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Data"]] }],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].freezePane).toBeUndefined();
    expect(workbook.sheets[0].splitPane).toBeUndefined();
  });
});

// ── Multiple sheets with different pane types ────────────────────────

describe("split panes — multiple sheets", () => {
  it("each sheet can have independent pane settings", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Frozen",
          rows: [["Data"]],
          freezePane: { rows: 1 },
        },
        {
          name: "Split",
          rows: [["Data"]],
          splitPane: { xSplit: 5000, ySplit: 2000 },
        },
        {
          name: "None",
          rows: [["Data"]],
        },
      ],
    });

    const workbook = await readXlsx(data);

    // First sheet: freeze
    expect(workbook.sheets[0].freezePane).toBeDefined();
    expect(workbook.sheets[0].freezePane!.rows).toBe(1);
    expect(workbook.sheets[0].splitPane).toBeUndefined();

    // Second sheet: split
    expect(workbook.sheets[1].splitPane).toBeDefined();
    expect(workbook.sheets[1].splitPane!.xSplit).toBe(5000);
    expect(workbook.sheets[1].splitPane!.ySplit).toBe(2000);
    expect(workbook.sheets[1].freezePane).toBeUndefined();

    // Third sheet: none
    expect(workbook.sheets[2].freezePane).toBeUndefined();
    expect(workbook.sheets[2].splitPane).toBeUndefined();
  });
});
