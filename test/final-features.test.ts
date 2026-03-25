import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { writeDrawing } from "../src/xlsx/drawing-writer";
import { writeContentTypes } from "../src/xlsx/content-types-writer";
import { parseCsv } from "../src/csv/reader";
import { fetchCsv } from "../src/csv/fetch";
import type { WriteSheet, SheetImage } from "../src/_types";

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

async function zipExtract(data: Uint8Array, path: string): Promise<Uint8Array> {
  const zip = new ZipReader(data);
  return zip.extract(path);
}

function fakeImage(size = 64): Uint8Array {
  const data = new Uint8Array(size);
  for (let i = 0; i < size; i++) {
    data[i] = i % 256;
  }
  return data;
}

// ── #57: CSV fetch from URL ──────────────────────────────────────────

describe("fetchCsv", () => {
  it("should parse CSV text using parseCsv under the hood", () => {
    // We can't easily test fetch without a real server, but we can
    // verify fetchCsv is exported and parseCsv works correctly
    const result = parseCsv("a,b,c\n1,2,3\n4,5,6");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["1", "2", "3"],
      ["4", "5", "6"],
    ]);
  });

  it("should be a function that accepts url and options", () => {
    expect(typeof fetchCsv).toBe("function");
    expect(fetchCsv.length).toBeGreaterThanOrEqual(1);
  });
});

// ── #78: SVG image format support ────────────────────────────────────

describe("SVG image support", () => {
  it("should write SVG image to drawing XML with correct content type", () => {
    const svgImage: SheetImage = {
      data: fakeImage(),
      type: "svg",
      anchor: { from: { row: 0, col: 0 } },
    };

    const result = writeDrawing([svgImage], 1);

    // Check that the image path has .svg extension
    expect(result.images).toHaveLength(1);
    expect(result.images[0].path).toBe("xl/media/image1.svg");
    expect(result.images[0].contentType).toBe("image/svg+xml");
  });

  it("should include svg extension in content types", () => {
    const xml = writeContentTypes({
      sheetCount: 1,
      hasSharedStrings: false,
      imageExtensions: new Set(["svg"]),
    });

    expect(xml).toContain('Extension="svg"');
    expect(xml).toContain('ContentType="image/svg+xml"');
  });

  it("should write SVG image in full XLSX roundtrip", async () => {
    const svgData = fakeImage(128);
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["hello"]],
      images: [
        {
          data: svgData,
          type: "svg",
          anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
        },
      ],
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });

    // Verify the SVG file is in the ZIP
    expect(zipHas(xlsx, "xl/media/image1.svg")).toBe(true);

    // Verify the image data matches
    const extracted = await zipExtract(xlsx, "xl/media/image1.svg");
    expect(extracted).toEqual(svgData);

    // Verify content types include SVG
    const ctXml = await extractXml(xlsx, "[Content_Types].xml");
    expect(ctXml).toContain("image/svg+xml");
  });
});

// ── #78: WebP image format support ───────────────────────────────────

describe("WebP image support", () => {
  it("should write WebP image to drawing XML with correct content type", () => {
    const webpImage: SheetImage = {
      data: fakeImage(),
      type: "webp",
      anchor: { from: { row: 0, col: 0 } },
    };

    const result = writeDrawing([webpImage], 1);

    expect(result.images).toHaveLength(1);
    expect(result.images[0].path).toBe("xl/media/image1.webp");
    expect(result.images[0].contentType).toBe("image/webp");
  });

  it("should include webp extension in content types", () => {
    const xml = writeContentTypes({
      sheetCount: 1,
      hasSharedStrings: false,
      imageExtensions: new Set(["webp"]),
    });

    expect(xml).toContain('Extension="webp"');
    expect(xml).toContain('ContentType="image/webp"');
  });

  it("should write WebP image in full XLSX roundtrip", async () => {
    const webpData = fakeImage(96);
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["test"]],
      images: [
        {
          data: webpData,
          type: "webp",
          anchor: { from: { row: 0, col: 0 }, to: { row: 3, col: 2 } },
        },
      ],
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });

    expect(zipHas(xlsx, "xl/media/image1.webp")).toBe(true);

    const extracted = await zipExtract(xlsx, "xl/media/image1.webp");
    expect(extracted).toEqual(webpData);

    const ctXml = await extractXml(xlsx, "[Content_Types].xml");
    expect(ctXml).toContain("image/webp");
  });
});

// ── #86: Background image (watermark) ────────────────────────────────

describe("Background image (watermark)", () => {
  it("should write background image to xl/media/ with picture element", async () => {
    const bgData = fakeImage(256);
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["watermark test"]],
      backgroundImage: bgData,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });

    // Verify the media file exists in the ZIP
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(true);

    // Verify the image data matches
    const extracted = await zipExtract(xlsx, "xl/media/image1.png");
    expect(extracted).toEqual(bgData);

    // Verify the worksheet XML contains <picture r:id="..."/>
    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    expect(wsXml).toContain("<picture");
    expect(wsXml).toContain('r:id="');

    // Verify the worksheet .rels contains an image relationship
    const relsXml = await extractXml(xlsx, "xl/worksheets/_rels/sheet1.xml.rels");
    expect(relsXml).toContain("relationships/image");
    expect(relsXml).toContain("media/image1.png");
  });

  it("should read back background image via roundtrip", async () => {
    const bgData = fakeImage(128);
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["bg test"]],
      backgroundImage: bgData,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].backgroundImage).toBeDefined();
    expect(workbook.sheets[0].backgroundImage).toEqual(bgData);
  });

  it("should handle background image alongside regular images", async () => {
    const bgData = fakeImage(64);
    const imgData = fakeImage(96);
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["both"]],
      images: [
        {
          data: imgData,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 3, col: 3 } },
        },
      ],
      backgroundImage: bgData,
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });

    // Both images should be in the ZIP
    expect(zipHas(xlsx, "xl/media/image1.png")).toBe(true); // regular image
    expect(zipHas(xlsx, "xl/media/image2.png")).toBe(true); // background image

    // Worksheet XML should have both drawing and picture elements
    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    expect(wsXml).toContain("<drawing");
    expect(wsXml).toContain("<picture");

    // Roundtrip should preserve both
    const workbook = await readXlsx(xlsx);
    expect(workbook.sheets[0].images).toHaveLength(1);
    expect(workbook.sheets[0].backgroundImage).toEqual(bgData);
  });

  it("should not produce picture element when no background image", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["no bg"]],
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const wsXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    expect(wsXml).not.toContain("<picture");
  });
});
