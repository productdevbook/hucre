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

// ── showGridLines Writing Tests ──────────────────────────────────────

describe("view settings — showGridLines", () => {
  it("writes showGridLines='0' when gridlines hidden", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { showGridLines: false },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["showGridLines"]).toBe("0");
  });

  it("does not emit showGridLines when true (default)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { showGridLines: true },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["showGridLines"]).toBeUndefined();
  });
});

// ── showRowColHeaders Writing Tests ──────────────────────────────────

describe("view settings — showRowColHeaders", () => {
  it("writes showRowColHeaders='0' when headers hidden", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { showRowColHeaders: false },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["showRowColHeaders"]).toBe("0");
  });

  it("does not emit showRowColHeaders when true (default)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["showRowColHeaders"]).toBeUndefined();
  });
});

// ── zoomScale Writing Tests ─────────────────────────────────────────

describe("view settings — zoomScale", () => {
  it("writes zoomScale='75' when zoom is 75%", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { zoomScale: 75 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["zoomScale"]).toBe("75");
  });

  it("writes zoomScale='150' when zoom is 150%", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { zoomScale: 150 },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["zoomScale"]).toBe("150");
  });

  it("does not emit zoomScale when not set", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["zoomScale"]).toBeUndefined();
  });
});

// ── rightToLeft Writing Tests ───────────────────────────────────────

describe("view settings — rightToLeft (RTL)", () => {
  it("writes rightToLeft='1' when RTL enabled", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { rightToLeft: true },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["rightToLeft"]).toBe("1");
  });

  it("does not emit rightToLeft when false", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { rightToLeft: false },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["rightToLeft"]).toBeUndefined();
  });
});

// ── tabColor Writing Tests ──────────────────────────────────────────

describe("view settings — tabColor", () => {
  it("writes sheetPr with tabColor when tab color is set", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { tabColor: { rgb: "0000FF" } },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetPr = findChild(doc, "sheetPr");
    expect(sheetPr).toBeDefined();
    const tabColor = findChild(sheetPr, "tabColor");
    expect(tabColor).toBeDefined();
    expect(tabColor.attrs["rgb"]).toBe("FF0000FF");
  });

  it("writes tabColor with theme color", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { tabColor: { theme: 4, tint: -0.25 } },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetPr = findChild(doc, "sheetPr");
    const tabColor = findChild(sheetPr, "tabColor");
    expect(tabColor.attrs["theme"]).toBe("4");
    expect(tabColor.attrs["tint"]).toBe("-0.25");
  });

  it("sheetPr appears before sheetViews in XML", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { tabColor: { rgb: "FF0000" } },
    };

    const xml = writeXml(sheet);
    const sheetPrPos = xml.indexOf("<sheetPr");
    const sheetViewsPos = xml.indexOf("<sheetViews");
    expect(sheetPrPos).toBeGreaterThan(-1);
    expect(sheetViewsPos).toBeGreaterThan(-1);
    expect(sheetPrPos).toBeLessThan(sheetViewsPos);
  });

  it("does not emit sheetPr when no tabColor", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: { showGridLines: false },
    };

    const xml = writeXml(sheet);
    expect(xml).not.toContain("<sheetPr");
  });
});

// ── Default Values (no view settings) ───────────────────────────────

describe("view settings — defaults", () => {
  it("no extra attributes when no view settings specified", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");

    expect(sheetView.attrs["showGridLines"]).toBeUndefined();
    expect(sheetView.attrs["showRowColHeaders"]).toBeUndefined();
    expect(sheetView.attrs["zoomScale"]).toBeUndefined();
    expect(sheetView.attrs["rightToLeft"]).toBeUndefined();
    expect(sheetView.attrs["workbookViewId"]).toBe("0");
  });

  it("combines multiple view settings", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      view: {
        showGridLines: false,
        showRowColHeaders: false,
        zoomScale: 80,
        rightToLeft: true,
        tabColor: { rgb: "00FF00" },
      },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView.attrs["showGridLines"]).toBe("0");
    expect(sheetView.attrs["showRowColHeaders"]).toBe("0");
    expect(sheetView.attrs["zoomScale"]).toBe("80");
    expect(sheetView.attrs["rightToLeft"]).toBe("1");

    const sheetPr = findChild(doc, "sheetPr");
    expect(sheetPr).toBeDefined();
    const tabColor = findChild(sheetPr, "tabColor");
    expect(tabColor.attrs["rgb"]).toBe("FF00FF00");
  });
});

// ── Round-trip Tests ────────────────────────────────────────────────

describe("view settings — round-trip", () => {
  it("round-trips showGridLines: false", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          view: { showGridLines: false },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].view).toBeDefined();
    expect(workbook.sheets[0].view!.showGridLines).toBe(false);
  });

  it("round-trips showRowColHeaders: false", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          view: { showRowColHeaders: false },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].view).toBeDefined();
    expect(workbook.sheets[0].view!.showRowColHeaders).toBe(false);
  });

  it("round-trips zoomScale", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          view: { zoomScale: 75 },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].view).toBeDefined();
    expect(workbook.sheets[0].view!.zoomScale).toBe(75);
  });

  it("round-trips rightToLeft", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          view: { rightToLeft: true },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].view).toBeDefined();
    expect(workbook.sheets[0].view!.rightToLeft).toBe(true);
  });

  it("round-trips tabColor with RGB", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          view: { tabColor: { rgb: "FF0000" } },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].view).toBeDefined();
    expect(workbook.sheets[0].view!.tabColor).toBeDefined();
    expect(workbook.sheets[0].view!.tabColor!.rgb).toBe("FF0000");
  });

  it("round-trips tabColor with theme", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          view: { tabColor: { theme: 5 } },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].view).toBeDefined();
    expect(workbook.sheets[0].view!.tabColor).toBeDefined();
    expect(workbook.sheets[0].view!.tabColor!.theme).toBe(5);
  });

  it("round-trips all view settings combined", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          view: {
            showGridLines: false,
            showRowColHeaders: false,
            zoomScale: 80,
            rightToLeft: true,
            tabColor: { rgb: "0000FF" },
          },
        },
      ],
    });

    const workbook = await readXlsx(data);
    const view = workbook.sheets[0].view;
    expect(view).toBeDefined();
    expect(view!.showGridLines).toBe(false);
    expect(view!.showRowColHeaders).toBe(false);
    expect(view!.zoomScale).toBe(80);
    expect(view!.rightToLeft).toBe(true);
    expect(view!.tabColor!.rgb).toBe("0000FF");
  });

  it("no view when not specified", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Data"]] }],
    });

    const workbook = await readXlsx(data);
    // view should be undefined since no view settings were set
    expect(workbook.sheets[0].view).toBeUndefined();
  });
});

// ── Multiple Sheets with Different View Settings ─────────────────────

describe("view settings — multiple sheets", () => {
  it("each sheet has independent view settings", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "RTL Sheet",
          rows: [["Arabic"]],
          view: { rightToLeft: true, tabColor: { rgb: "FF0000" } },
        },
        {
          name: "Zoomed Sheet",
          rows: [["Zoomed"]],
          view: { zoomScale: 150, showGridLines: false },
        },
        {
          name: "Default Sheet",
          rows: [["Normal"]],
        },
      ],
    });

    const workbook = await readXlsx(data);

    // RTL Sheet
    const view1 = workbook.sheets[0].view;
    expect(view1).toBeDefined();
    expect(view1!.rightToLeft).toBe(true);
    expect(view1!.tabColor!.rgb).toBe("FF0000");
    expect(view1!.zoomScale).toBeUndefined();

    // Zoomed Sheet
    const view2 = workbook.sheets[1].view;
    expect(view2).toBeDefined();
    expect(view2!.zoomScale).toBe(150);
    expect(view2!.showGridLines).toBe(false);
    expect(view2!.rightToLeft).toBeUndefined();

    // Default Sheet
    expect(workbook.sheets[2].view).toBeUndefined();
  });
});
