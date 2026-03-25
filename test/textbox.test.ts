import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import type { WriteSheet, SheetTextBox } from "../src/_types";

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

/** Create a simple fake PNG-like image */
function fakePng(size = 64): Uint8Array {
  const data = new Uint8Array(size);
  data[0] = 0x89;
  data[1] = 0x50;
  data[2] = 0x4e;
  data[3] = 0x47;
  data[4] = 0x0d;
  data[5] = 0x0a;
  data[6] = 0x1a;
  data[7] = 0x0a;
  for (let i = 8; i < size; i++) {
    data[i] = i % 256;
  }
  return data;
}

// ── Tests ────────────────────────────────────────────────────────────

describe("TextBox", () => {
  it("should write textbox as sp element in drawing XML", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Hello"]],
      textBoxes: [
        {
          text: "Hello World",
          anchor: { from: { row: 0, col: 0 }, to: { row: 3, col: 3 } },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });

    // Drawing file should exist
    expect(zipHas(data, "xl/drawings/drawing1.xml")).toBe(true);

    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    // Verify it contains a shape element (not pic)
    expect(xml).toContain("xdr:sp");
    expect(xml).toContain('txBox="1"');
    expect(xml).toContain("Hello World");
    expect(xml).toContain("xdr:txBody");
    expect(xml).toContain("a:bodyPr");
  });

  it("should write textbox with bold style", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      textBoxes: [
        {
          text: "Bold Text",
          anchor: { from: { row: 0, col: 0 } },
          style: { bold: true },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    expect(xml).toContain('b="1"');
    expect(xml).toContain("Bold Text");
  });

  it("should write textbox with custom font size", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      textBoxes: [
        {
          text: "Large Text",
          anchor: { from: { row: 0, col: 0 } },
          style: { fontSize: 24 },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    // 24pt = 2400 hundredths of a point
    expect(xml).toContain('sz="2400"');
  });

  it("should write textbox with custom text color", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      textBoxes: [
        {
          text: "Red Text",
          anchor: { from: { row: 0, col: 0 } },
          style: { color: "FF0000" },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    expect(xml).toContain('val="FF0000"');
    expect(xml).toContain("Red Text");
  });

  it("should write textbox with fill and border colors", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      textBoxes: [
        {
          text: "Styled Box",
          anchor: { from: { row: 0, col: 0 } },
          style: { fillColor: "FFFF00", borderColor: "0000FF" },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    expect(xml).toContain('val="FFFF00"');
    expect(xml).toContain('val="0000FF"');
  });

  it("should write worksheet drawing reference", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      textBoxes: [
        {
          text: "Box",
          anchor: { from: { row: 0, col: 0 } },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const wsXml = await extractXml(data, "xl/worksheets/sheet1.xml");

    // Worksheet should have a drawing reference
    expect(wsXml).toContain("<drawing");
    expect(wsXml).toContain("r:id=");
  });

  it("should round-trip textbox (write then read)", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      textBoxes: [
        {
          text: "My TextBox",
          anchor: { from: { row: 0, col: 1 }, to: { row: 3, col: 4 } },
          style: {
            fontSize: 14,
            bold: true,
            color: "FF0000",
            fillColor: "FFFFFF",
            borderColor: "000000",
          },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(data);

    expect(workbook.sheets.length).toBe(1);
    const readSheet = workbook.sheets[0];
    expect(readSheet.textBoxes).toBeDefined();
    expect(readSheet.textBoxes!.length).toBe(1);

    const tb = readSheet.textBoxes![0];
    expect(tb.text).toBe("My TextBox");
    expect(tb.anchor.from.row).toBe(0);
    expect(tb.anchor.from.col).toBe(1);
    expect(tb.anchor.to!.row).toBe(3);
    expect(tb.anchor.to!.col).toBe(4);

    // Style round-trip
    expect(tb.style).toBeDefined();
    expect(tb.style!.fontSize).toBe(14);
    expect(tb.style!.bold).toBe(true);
    expect(tb.style!.color).toBe("FF0000");
    expect(tb.style!.fillColor).toBe("FFFFFF");
    expect(tb.style!.borderColor).toBe("000000");
  });

  it("should write textbox alongside images", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      images: [
        {
          data: fakePng(),
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
        },
      ],
      textBoxes: [
        {
          text: "Caption",
          anchor: { from: { row: 6, col: 0 }, to: { row: 8, col: 3 } },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });

    expect(zipHas(data, "xl/drawings/drawing1.xml")).toBe(true);
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    // Both pic and sp should be present
    expect(xml).toContain("xdr:pic");
    expect(xml).toContain("xdr:sp");
    expect(xml).toContain('txBox="1"');
    expect(xml).toContain("Caption");

    // Read back
    const workbook = await readXlsx(data);
    const readSheet = workbook.sheets[0];
    expect(readSheet.images).toBeDefined();
    expect(readSheet.images!.length).toBe(1);
    expect(readSheet.textBoxes).toBeDefined();
    expect(readSheet.textBoxes!.length).toBe(1);
    expect(readSheet.textBoxes![0].text).toBe("Caption");
  });

  it("should handle textbox without explicit to anchor", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      textBoxes: [
        {
          text: "Auto-sized",
          anchor: { from: { row: 2, col: 3 } },
          width: 200,
          height: 50,
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    expect(xml).toContain("Auto-sized");
    // Default to anchor should be from + 3 cols/rows
    expect(xml).toContain("xdr:twoCellAnchor");
  });

  it("should generate unique shape ids when mixing images and textboxes", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      images: [
        {
          data: fakePng(),
          type: "png",
          anchor: { from: { row: 0, col: 0 } },
        },
        {
          data: fakePng(),
          type: "png",
          anchor: { from: { row: 5, col: 0 } },
        },
      ],
      textBoxes: [
        {
          text: "Box 1",
          anchor: { from: { row: 10, col: 0 } },
        },
        {
          text: "Box 2",
          anchor: { from: { row: 15, col: 0 } },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    // Parse XML and check all cNvPr ids are unique
    const idMatches = [...xml.matchAll(/id="(\d+)"/g)];
    const ids = idMatches.map((m) => m[1]);
    const uniqueIds = new Set(ids);
    expect(uniqueIds.size).toBe(ids.length);
  });

  it("should escape special XML characters in text", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Test"]],
      textBoxes: [
        {
          text: "A < B & C > D",
          anchor: { from: { row: 0, col: 0 } },
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/drawings/drawing1.xml");

    // Text should be XML-escaped
    expect(xml).toContain("A &lt; B &amp; C &gt; D");
  });
});
