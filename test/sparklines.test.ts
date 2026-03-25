import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import type { WriteSheet, Sparkline } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

// ── Tests ────────────────────────────────────────────────────────────

describe("Sparklines", () => {
  it("should write sparkline extLst in worksheet XML", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Label", 10, 20, 30, 40, 50],
        ["Data", 5, 15, 25, 35, 45],
      ],
      sparklines: [
        {
          location: "A2",
          dataRange: "Sheet1!B2:F2",
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    // Verify extLst is present
    expect(xml).toContain("<extLst>");
    expect(xml).toContain("x14:sparklineGroups");
    expect(xml).toContain("x14:sparklineGroup");
    expect(xml).toContain("x14:sparkline");
    expect(xml).toContain("<xm:f>Sheet1!B2:F2</xm:f>");
    expect(xml).toContain("<xm:sqref>A2</xm:sqref>");
  });

  it("should write line type sparkline (default)", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["", 1, 2, 3, 4, 5]],
      sparklines: [
        {
          location: "A1",
          dataRange: "Sheet1!B1:F1",
          type: "line",
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    // Line type is default, so type attribute should NOT be present
    expect(xml).toContain("x14:sparklineGroup");
    expect(xml).not.toMatch(/type="line"/);
  });

  it("should write column type sparkline", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["", 1, 2, 3, 4, 5]],
      sparklines: [
        {
          location: "A1",
          dataRange: "Sheet1!B1:F1",
          type: "column",
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    expect(xml).toContain('type="column"');
  });

  it("should write stacked (win/loss) type sparkline", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["", 1, -2, 3, -4, 5]],
      sparklines: [
        {
          location: "A1",
          dataRange: "Sheet1!B1:F1",
          type: "stacked",
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    expect(xml).toContain('type="stacked"');
  });

  it("should write custom color", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["", 1, 2, 3]],
      sparklines: [
        {
          location: "A1",
          dataRange: "Sheet1!B1:D1",
          color: "FF0000",
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    expect(xml).toContain('rgb="FFFF0000"');
  });

  it("should write markers attribute", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["", 1, 2, 3]],
      sparklines: [
        {
          location: "A1",
          dataRange: "Sheet1!B1:D1",
          markers: true,
        },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    expect(xml).toContain('markers="1"');
  });

  it("should write multiple sparklines", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["", 10, 20, 30],
        ["", 5, 15, 25],
        ["", 8, 18, 28],
      ],
      sparklines: [
        { location: "A1", dataRange: "Sheet1!B1:D1" },
        { location: "A2", dataRange: "Sheet1!B2:D2", type: "column" },
        { location: "A3", dataRange: "Sheet1!B3:D3", color: "00FF00" },
      ],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    // Count sparkline elements (each has opening and closing tag)
    const sparklineMatches = xml.match(/<x14:sparkline>/g);
    expect(sparklineMatches).not.toBeNull();
    expect(sparklineMatches!.length).toBe(3);

    // Verify each sparkline's data range
    expect(xml).toContain("Sheet1!B1:D1");
    expect(xml).toContain("Sheet1!B2:D2");
    expect(xml).toContain("Sheet1!B3:D3");
  });

  it("should round-trip sparklines (write then read)", async () => {
    const sparklines: Sparkline[] = [
      { location: "A1", dataRange: "Sheet1!B1:D1" },
      { location: "A2", dataRange: "Sheet1!B2:D2", type: "column", color: "FF0000" },
      { location: "A3", dataRange: "Sheet1!B3:D3", type: "stacked", markers: true },
    ];

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["", 10, 20, 30],
        ["", 5, 15, 25],
        ["", 8, 18, 28],
      ],
      sparklines,
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(data);

    expect(workbook.sheets.length).toBe(1);
    const readSheet = workbook.sheets[0];
    expect(readSheet.sparklines).toBeDefined();
    expect(readSheet.sparklines!.length).toBe(3);

    // Check first sparkline (line type, default color)
    const sp1 = readSheet.sparklines![0];
    expect(sp1.location).toBe("A1");
    expect(sp1.dataRange).toBe("Sheet1!B1:D1");

    // Check second sparkline (column type, red)
    const sp2 = readSheet.sparklines![1];
    expect(sp2.location).toBe("A2");
    expect(sp2.dataRange).toBe("Sheet1!B2:D2");
    expect(sp2.type).toBe("column");
    expect(sp2.color).toBe("FF0000");

    // Check third sparkline (stacked, markers)
    const sp3 = readSheet.sparklines![2];
    expect(sp3.location).toBe("A3");
    expect(sp3.dataRange).toBe("Sheet1!B3:D3");
    expect(sp3.type).toBe("stacked");
    expect(sp3.markers).toBe(true);
  });

  it("should handle sparklines alongside other features", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Name", 10, 20, 30],
        ["Data", 5, 15, 25],
      ],
      sparklines: [{ location: "A1", dataRange: "Sheet1!B1:D1" }],
      freezePane: { rows: 1 },
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    // Both freeze pane and sparklines should be present
    expect(xml).toContain("pane");
    expect(xml).toContain("x14:sparklineGroup");
  });

  it("should write default color when none specified", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["", 1, 2, 3]],
      sparklines: [{ location: "A1", dataRange: "Sheet1!B1:D1" }],
    };

    const data = await writeXlsx({ sheets: [sheet] });
    const xml = await extractXml(data, "xl/worksheets/sheet1.xml");

    // Default color FF376092
    expect(xml).toContain('rgb="FF376092"');
  });
});
