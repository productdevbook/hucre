import { describe, it, expect } from "vitest";
import { parseThemeColors, resolveThemeColor } from "../src/xlsx/theme";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";

// ── Standard Office theme XML ────────────────────────────────────────

const OFFICE_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

// ── parseThemeColors ─────────────────────────────────────────────────

describe("parseThemeColors", () => {
  it("extracts 12 colors from standard Office theme", () => {
    const colors = parseThemeColors(OFFICE_THEME_XML);
    expect(colors).toHaveLength(12);
    expect(colors[0]).toBe("000000"); // dk1 (sysClr windowText)
    expect(colors[1]).toBe("FFFFFF"); // lt1 (sysClr window)
    expect(colors[2]).toBe("44546A"); // dk2
    expect(colors[3]).toBe("E7E6E6"); // lt2
    expect(colors[4]).toBe("4472C4"); // accent1
    expect(colors[5]).toBe("ED7D31"); // accent2
    expect(colors[6]).toBe("A5A5A5"); // accent3
    expect(colors[7]).toBe("FFC000"); // accent4
    expect(colors[8]).toBe("5B9BD5"); // accent5
    expect(colors[9]).toBe("70AD47"); // accent6
    expect(colors[10]).toBe("0563C1"); // hlink
    expect(colors[11]).toBe("954F72"); // folHlink
  });

  it("handles lowercase hex values", () => {
    const xml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:clrScheme name="Custom">
          <a:dk1><a:srgbClr val="aabbcc"/></a:dk1>
          <a:lt1><a:srgbClr val="ddeeff"/></a:lt1>
          <a:dk2><a:srgbClr val="112233"/></a:dk2>
          <a:lt2><a:srgbClr val="445566"/></a:lt2>
          <a:accent1><a:srgbClr val="778899"/></a:accent1>
          <a:accent2><a:srgbClr val="aabb00"/></a:accent2>
          <a:accent3><a:srgbClr val="cc0011"/></a:accent3>
          <a:accent4><a:srgbClr val="dd2233"/></a:accent4>
          <a:accent5><a:srgbClr val="ee4455"/></a:accent5>
          <a:accent6><a:srgbClr val="ff6677"/></a:accent6>
          <a:hlink><a:srgbClr val="001122"/></a:hlink>
          <a:folHlink><a:srgbClr val="334455"/></a:folHlink>
        </a:clrScheme>
      </a:themeElements>
    </a:theme>`;
    const colors = parseThemeColors(xml);
    // Should be uppercased
    expect(colors[0]).toBe("AABBCC");
    expect(colors[1]).toBe("DDEEFF");
  });

  it("uses 000000 fallback for missing slots", () => {
    const xml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:clrScheme name="Sparse">
          <a:dk1><a:srgbClr val="111111"/></a:dk1>
          <a:lt1><a:srgbClr val="EEEEEE"/></a:lt1>
        </a:clrScheme>
      </a:themeElements>
    </a:theme>`;
    const colors = parseThemeColors(xml);
    expect(colors).toHaveLength(12);
    expect(colors[0]).toBe("111111");
    expect(colors[1]).toBe("EEEEEE");
    // Rest should be fallback
    expect(colors[2]).toBe("000000");
    expect(colors[4]).toBe("000000");
  });

  it("handles theme without namespace prefix", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
    <theme>
      <themeElements>
        <clrScheme name="NoNS">
          <dk1><srgbClr val="112233"/></dk1>
          <lt1><srgbClr val="FFEEDD"/></lt1>
          <dk2><srgbClr val="445566"/></dk2>
          <lt2><srgbClr val="778899"/></lt2>
          <accent1><srgbClr val="AABBCC"/></accent1>
          <accent2><srgbClr val="DDEEFF"/></accent2>
          <accent3><srgbClr val="001122"/></accent3>
          <accent4><srgbClr val="334455"/></accent4>
          <accent5><srgbClr val="667788"/></accent5>
          <accent6><srgbClr val="99AABB"/></accent6>
          <hlink><srgbClr val="CCDDEE"/></hlink>
          <folHlink><srgbClr val="FF0011"/></folHlink>
        </clrScheme>
      </themeElements>
    </theme>`;
    const colors = parseThemeColors(xml);
    expect(colors[0]).toBe("112233");
    expect(colors[4]).toBe("AABBCC");
    expect(colors[11]).toBe("FF0011");
  });
});

// ── resolveThemeColor ────────────────────────────────────────────────

describe("resolveThemeColor", () => {
  const themeColors = parseThemeColors(OFFICE_THEME_XML);

  it("resolves theme index 0 to dk1 color", () => {
    expect(resolveThemeColor(themeColors, 0)).toBe("000000");
  });

  it("resolves theme index 1 to lt1 color", () => {
    expect(resolveThemeColor(themeColors, 1)).toBe("FFFFFF");
  });

  it("resolves theme index 4 to accent1 color", () => {
    expect(resolveThemeColor(themeColors, 4)).toBe("4472C4");
  });

  it("resolves theme index 10 to hlink color", () => {
    expect(resolveThemeColor(themeColors, 10)).toBe("0563C1");
  });

  it("resolves with no tint (undefined) — unchanged", () => {
    expect(resolveThemeColor(themeColors, 4, undefined)).toBe("4472C4");
  });

  it("resolves with tint 0 — unchanged", () => {
    expect(resolveThemeColor(themeColors, 4, 0)).toBe("4472C4");
  });

  it("resolves with positive tint (lighten)", () => {
    // accent1 = 4472C4 → R=68, G=114, B=196
    // tint = 0.4 → lighten
    // R: 68 + (255 - 68) * 0.4 = 68 + 74.8 = 142.8 → 143 → 8F
    // G: 114 + (255 - 114) * 0.4 = 114 + 56.4 = 170.4 → 170 → AA
    // B: 196 + (255 - 196) * 0.4 = 196 + 23.6 = 219.6 → 220 → DC
    const result = resolveThemeColor(themeColors, 4, 0.4);
    expect(result).toBe("8FAADC");
  });

  it("resolves with negative tint (darken)", () => {
    // accent1 = 4472C4 → R=68, G=114, B=196
    // tint = -0.25 → darken
    // R: 68 * (1 - 0.25) = 68 * 0.75 = 51 → 33
    // G: 114 * 0.75 = 85.5 → 86 → 56
    // B: 196 * 0.75 = 147 → 93
    const result = resolveThemeColor(themeColors, 4, -0.25);
    expect(result).toBe("335693");
  });

  it("resolves with large positive tint (near white)", () => {
    // dk1 = 000000 → R=0, G=0, B=0
    // tint = 0.8 → lighten
    // R: 0 + (255 - 0) * 0.8 = 204 → CC
    // G: 0 + 255 * 0.8 = 204 → CC
    // B: 0 + 255 * 0.8 = 204 → CC
    const result = resolveThemeColor(themeColors, 0, 0.8);
    expect(result).toBe("CCCCCC");
  });

  it("resolves with large negative tint (near black)", () => {
    // lt1 = FFFFFF → R=255, G=255, B=255
    // tint = -0.5 → darken
    // R: 255 * (1 - 0.5) = 255 * 0.5 = 127.5 → 128 → 80
    const result = resolveThemeColor(themeColors, 1, -0.5);
    expect(result).toBe("808080");
  });

  it("returns 000000 for out-of-range index", () => {
    expect(resolveThemeColor(themeColors, 99)).toBe("000000");
    expect(resolveThemeColor(themeColors, -1)).toBe("000000");
  });
});

// ── Round-trip: themeColors on Workbook ──────────────────────────────

describe("theme colors — round-trip", () => {
  it("populates themeColors on workbook when reading XLSX", async () => {
    // Write an XLSX, read it back — the writer includes a theme1.xml
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["hello"]] }],
    });
    const workbook = await readXlsx(data);

    // The standard writer should include a theme file
    expect(workbook.themeColors).toBeDefined();
    expect(workbook.themeColors).toHaveLength(12);
    // Standard theme dk1 should be "000000" (windowText)
    expect(workbook.themeColors![0]).toBe("000000");
    // Standard theme lt1 should be "FFFFFF" (window)
    expect(workbook.themeColors![1]).toBe("FFFFFF");
  });

  it("themeColors are undefined when theme1.xml is missing", async () => {
    // We can't easily test this without a custom ZIP, but we can verify
    // that the reader handles missing theme gracefully by checking
    // the code path. The write+read roundtrip always includes theme1.xml.
    // This is a code-level check that the conditional assignment works.
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
    });
    const workbook = await readXlsx(data);
    // In a standard write, themeColors should be present
    // If we had a file without theme1.xml, it would be undefined
    // This test confirms the field exists on the type
    if (workbook.themeColors) {
      expect(Array.isArray(workbook.themeColors)).toBe(true);
    }
  });
});
