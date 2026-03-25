import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { ZipWriter } from "../src/zip/writer";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import type { WriteSheet, SheetImage } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
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

function makeImage(
  type: SheetImage["type"],
  from: { row: number; col: number },
  to?: { row: number; col: number },
  opts?: { width?: number; height?: number },
): SheetImage {
  return {
    data: fakePng(),
    type,
    anchor: { from, to },
    width: opts?.width,
    height: opts?.height,
  };
}

/**
 * Build a minimal XLSX with a oneCellAnchor drawing, by taking a valid XLSX
 * and replacing its drawing XML with a oneCellAnchor variant.
 */
async function buildXlsxWithOneCellAnchor(
  imageData: Uint8Array,
  fromRow: number,
  fromCol: number,
  extCx: number,
  extCy: number,
): Promise<Uint8Array> {
  // First, write a normal XLSX with a twoCellAnchor image
  const sheet: WriteSheet = {
    name: "Sheet1",
    rows: [["Data"]],
    images: [
      {
        data: imageData,
        type: "png",
        anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
      },
    ],
  };

  const baseXlsx = await writeXlsx({ sheets: [sheet] });

  // Now rewrite the drawing XML to use oneCellAnchor instead
  const oneCellDrawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>${fromCol}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>${fromRow}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="${extCx}" cy="${extCy}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${extCx}" cy="${extCy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;

  // Read the base XLSX, replace drawing1.xml, rewrite
  const baseZip = new ZipReader(baseXlsx);
  const newZip = new ZipWriter();

  for (const entry of baseZip.entries()) {
    if (entry === "xl/drawings/drawing1.xml") {
      newZip.add(entry, encoder.encode(oneCellDrawingXml));
    } else {
      const data = await baseZip.extract(entry);
      newZip.add(entry, data);
    }
  }

  return newZip.build();
}

// ── twoCellAnchor still works ─────────────────────────────────────────

describe("twoCellAnchor (existing, verify still works)", () => {
  it("writes and reads back a twoCellAnchor image", async () => {
    const imageData = fakePng(100);
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Hello"]],
      images: [
        {
          data: imageData,
          type: "png",
          anchor: { from: { row: 1, col: 2 }, to: { row: 8, col: 6 } },
        },
      ],
    };

    const xlsx = await writeXlsx({ sheets: [sheet] });
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets[0].images).toHaveLength(1);
    const img = workbook.sheets[0].images![0];
    expect(img.type).toBe("png");
    expect(img.data).toEqual(imageData);
    expect(img.anchor.from).toEqual({ row: 1, col: 2 });
    expect(img.anchor.to).toEqual({ row: 8, col: 6 });
  });
});

// ── oneCellAnchor reading ─────────────────────────────────────────────

describe("oneCellAnchor image reading", () => {
  it("reads image from oneCellAnchor", async () => {
    const imageData = fakePng(128);
    const fromRow = 3;
    const fromCol = 2;
    const extCx = 3000000; // ~315 pixels
    const extCy = 2000000; // ~210 pixels

    const xlsx = await buildXlsxWithOneCellAnchor(imageData, fromRow, fromCol, extCx, extCy);
    const workbook = await readXlsx(xlsx);

    expect(workbook.sheets[0].images).toBeDefined();
    expect(workbook.sheets[0].images).toHaveLength(1);

    const img = workbook.sheets[0].images![0];
    expect(img.type).toBe("png");
    expect(img.data).toEqual(imageData);
  });

  it("parses correct from position", async () => {
    const imageData = fakePng(64);
    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 5, 7, 1000000, 500000);
    const workbook = await readXlsx(xlsx);

    const img = workbook.sheets[0].images![0];
    expect(img.anchor.from.row).toBe(5);
    expect(img.anchor.from.col).toBe(7);
  });

  it("does not have anchor.to for oneCellAnchor", async () => {
    const imageData = fakePng(64);
    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 0, 0, 2000000, 1000000);
    const workbook = await readXlsx(xlsx);

    const img = workbook.sheets[0].images![0];
    // oneCellAnchor has no "to" element
    expect(img.anchor.to).toBeUndefined();
  });

  it("converts ext dimensions from EMU to pixels", async () => {
    const imageData = fakePng(64);
    // 9525 EMU = 1 pixel, so 952500 EMU = 100 pixels
    const extCx = 952500; // 100 pixels
    const extCy = 476250; // 50 pixels

    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 0, 0, extCx, extCy);
    const workbook = await readXlsx(xlsx);

    const img = workbook.sheets[0].images![0];
    expect(img.width).toBe(100);
    expect(img.height).toBe(50);
  });

  it("handles large dimensions correctly", async () => {
    const imageData = fakePng(64);
    // 1920 pixels wide, 1080 pixels tall
    const extCx = 1920 * 9525; // 18288000
    const extCy = 1080 * 9525; // 10287000

    const xlsx = await buildXlsxWithOneCellAnchor(imageData, 0, 0, extCx, extCy);
    const workbook = await readXlsx(xlsx);

    const img = workbook.sheets[0].images![0];
    expect(img.width).toBe(1920);
    expect(img.height).toBe(1080);
  });
});

// ── Mixed anchors ────────────────────────────────────────────────────

describe("mixed anchor types", () => {
  it("reads both twoCellAnchor and oneCellAnchor from same drawing", async () => {
    const imageData1 = fakePng(80);
    const imageData2 = fakePng(90);

    // Build an XLSX with a twoCellAnchor image first
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Data"]],
      images: [
        {
          data: imageData1,
          type: "png",
          anchor: { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } },
        },
      ],
    };

    const baseXlsx = await writeXlsx({ sheets: [sheet] });

    // Replace drawing XML with one that has both twoCellAnchor and oneCellAnchor
    const mixedDrawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="1905000" cy="952500"/>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 2"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;

    const baseZip = new ZipReader(baseXlsx);
    const newZip = new ZipWriter();

    for (const entry of baseZip.entries()) {
      if (entry === "xl/drawings/drawing1.xml") {
        newZip.add(entry, encoder.encode(mixedDrawingXml));
      } else {
        const data = await baseZip.extract(entry);
        newZip.add(entry, data);
      }
    }

    const modifiedXlsx = await newZip.build();
    const workbook = await readXlsx(modifiedXlsx);

    expect(workbook.sheets[0].images).toHaveLength(2);

    // First image: twoCellAnchor
    const img1 = workbook.sheets[0].images![0];
    expect(img1.anchor.from).toEqual({ row: 0, col: 0 });
    expect(img1.anchor.to).toEqual({ row: 5, col: 3 });

    // Second image: oneCellAnchor
    const img2 = workbook.sheets[0].images![1];
    expect(img2.anchor.from).toEqual({ row: 10, col: 5 });
    expect(img2.anchor.to).toBeUndefined();
    // 1905000 / 9525 = 200, 952500 / 9525 = 100
    expect(img2.width).toBe(200);
    expect(img2.height).toBe(100);
  });
});
