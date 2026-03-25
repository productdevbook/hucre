import { describe, expect, it } from "vitest";
import { imageFromBase64 } from "../src/image";
import { fromHtml } from "../src/export/html-import";

// ── imageFromBase64 ─────────────────────────────────────────────────

describe("imageFromBase64", () => {
  it("creates a SheetImage from a base64 PNG string", () => {
    // 1x1 red PNG pixel (minimal valid PNG)
    const base64Png =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    const anchor = { from: { row: 0, col: 0 }, to: { row: 5, col: 3 } };
    const image = imageFromBase64(base64Png, "png", anchor);

    expect(image.type).toBe("png");
    expect(image.anchor).toEqual(anchor);
    expect(image.data).toBeInstanceOf(Uint8Array);
    expect(image.data.length).toBeGreaterThan(0);

    // PNG magic bytes: 137 80 78 71 13 10 26 10
    expect(image.data[0]).toBe(137);
    expect(image.data[1]).toBe(80); // 'P'
    expect(image.data[2]).toBe(78); // 'N'
    expect(image.data[3]).toBe(71); // 'G'
  });

  it("handles data URI prefix", () => {
    const base64Png =
      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";

    const anchor = { from: { row: 0, col: 0 } };
    const image = imageFromBase64(base64Png, "png", anchor);

    expect(image.data).toBeInstanceOf(Uint8Array);
    expect(image.data[0]).toBe(137); // PNG magic byte
  });

  it("creates a JPEG image", () => {
    // Minimal base64 string (not a real JPEG, but tests the plumbing)
    const base64 = btoa("fake-jpeg-data");
    const anchor = { from: { row: 1, col: 2 } };
    const image = imageFromBase64(base64, "jpeg", anchor);

    expect(image.type).toBe("jpeg");
    expect(image.anchor).toEqual(anchor);
    expect(image.data).toBeInstanceOf(Uint8Array);
    expect(new TextDecoder().decode(image.data)).toBe("fake-jpeg-data");
  });
});

// ── fromHtml ────────────────────────────────────────────────────────

describe("fromHtml", () => {
  it("parses a basic table into a Sheet with correct rows", () => {
    const html = `
      <table>
        <tr><td>A</td><td>B</td></tr>
        <tr><td>C</td><td>D</td></tr>
      </table>
    `;

    const sheet = fromHtml(html);
    expect(sheet.name).toBe("Sheet1");
    expect(sheet.rows.length).toBe(2);
    expect(sheet.rows[0]).toEqual(["A", "B"]);
    expect(sheet.rows[1]).toEqual(["C", "D"]);
  });

  it("accepts a custom sheet name", () => {
    const html = "<table><tr><td>X</td></tr></table>";
    const sheet = fromHtml(html, { sheetName: "Data" });
    expect(sheet.name).toBe("Data");
  });

  it("parses a table with thead and tbody", () => {
    const html = `
      <table>
        <thead>
          <tr><th>Name</th><th>Age</th></tr>
        </thead>
        <tbody>
          <tr><td>Alice</td><td>30</td></tr>
          <tr><td>Bob</td><td>25</td></tr>
        </tbody>
      </table>
    `;

    const sheet = fromHtml(html);
    expect(sheet.rows.length).toBe(3);
    expect(sheet.rows[0]).toEqual(["Name", "Age"]);
    expect(sheet.rows[1]).toEqual(["Alice", 30]);
    expect(sheet.rows[2]).toEqual(["Bob", 25]);
  });

  it("handles colspan with merge ranges", () => {
    const html = `
      <table>
        <tr><td colspan="3">Header</td></tr>
        <tr><td>A</td><td>B</td><td>C</td></tr>
      </table>
    `;

    const sheet = fromHtml(html);
    expect(sheet.rows.length).toBe(2);
    // First row: "Header" + 2 null padding for colspan
    expect(sheet.rows[0]).toEqual(["Header", null, null]);
    expect(sheet.rows[1]).toEqual(["A", "B", "C"]);

    // Should have a merge for the colspan
    expect(sheet.merges).toBeDefined();
    expect(sheet.merges!.length).toBe(1);
    expect(sheet.merges![0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
    });
  });

  it("handles rowspan with merge ranges", () => {
    const html = `
      <table>
        <tr><td rowspan="2">Span</td><td>B1</td></tr>
        <tr><td>B2</td></tr>
      </table>
    `;

    const sheet = fromHtml(html);
    expect(sheet.rows.length).toBe(2);
    expect(sheet.rows[0]).toEqual(["Span", "B1"]);
    // Row 2: col 0 is occupied by rowspan, so B2 goes to col 1 with null padding at col 0
    expect(sheet.rows[1]).toEqual([null, "B2"]);

    expect(sheet.merges).toBeDefined();
    expect(sheet.merges!.length).toBe(1);
    expect(sheet.merges![0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 0,
    });
  });

  it("returns empty rows for an empty table", () => {
    const html = "<table></table>";
    const sheet = fromHtml(html);
    expect(sheet.rows).toEqual([]);
    expect(sheet.merges).toBeUndefined();
  });

  it("converts numeric cell text to numbers", () => {
    const html = `
      <table>
        <tr><td>42</td><td>3.14</td><td>hello</td></tr>
      </table>
    `;

    const sheet = fromHtml(html);
    expect(sheet.rows[0]).toEqual([42, 3.14, "hello"]);
  });

  it("treats empty cells as null", () => {
    const html = `
      <table>
        <tr><td></td><td>A</td><td></td></tr>
      </table>
    `;

    const sheet = fromHtml(html);
    expect(sheet.rows[0]).toEqual([null, "A", null]);
  });

  it("handles tfoot section", () => {
    const html = `
      <table>
        <thead><tr><th>Item</th><th>Qty</th></tr></thead>
        <tbody><tr><td>Widget</td><td>5</td></tr></tbody>
        <tfoot><tr><td>Total</td><td>5</td></tr></tfoot>
      </table>
    `;

    const sheet = fromHtml(html);
    expect(sheet.rows.length).toBe(3);
    expect(sheet.rows[2]).toEqual(["Total", 5]);
  });
});
