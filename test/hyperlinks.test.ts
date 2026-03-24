import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { ZipWriter } from "../src/zip/writer";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { collectHyperlinks } from "../src/xlsx/worksheet-writer";
import type { WriteSheet, Cell } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

async function parseXmlFromZip(data: Uint8Array, path: string) {
  const xml = await extractXml(data, path);
  return parseXml(xml);
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function zipHas(data: Uint8Array, path: string): boolean {
  const zip = new ZipReader(data);
  return zip.has(path);
}

// ── collectHyperlinks unit tests ─────────────────────────────────────

describe("collectHyperlinks", () => {
  it("returns empty when no cells", () => {
    const sheet: WriteSheet = { name: "Sheet1", rows: [["Hello"]] };
    const result = collectHyperlinks(sheet);
    expect(result.xml).toBe("");
    expect(result.relationships).toEqual([]);
  });

  it("returns empty when cells have no hyperlinks", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", { value: "Hello" });
    const sheet: WriteSheet = { name: "Sheet1", rows: [["Hello"]], cells };
    const result = collectHyperlinks(sheet);
    expect(result.xml).toBe("");
    expect(result.relationships).toEqual([]);
  });

  it("collects external hyperlink with relationship", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Click me",
      hyperlink: { target: "https://example.com" },
    });
    const sheet: WriteSheet = { name: "Sheet1", rows: [["Click me"]], cells };
    const result = collectHyperlinks(sheet);

    expect(result.relationships).toHaveLength(1);
    expect(result.relationships[0].id).toBe("rId1");
    expect(result.relationships[0].target).toBe("https://example.com");

    // Verify XML contains hyperlink element with r:id
    expect(result.xml).toContain("<hyperlinks>");
    expect(result.xml).toContain('r:id="rId1"');
    expect(result.xml).toContain('ref="A1"');
  });

  it("collects internal hyperlink without relationship", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Go to Sheet2",
      hyperlink: { target: "", location: "Sheet2!A1" },
    });
    const sheet: WriteSheet = { name: "Sheet1", rows: [["Go to Sheet2"]], cells };
    const result = collectHyperlinks(sheet);

    // Internal hyperlinks should NOT generate relationships
    expect(result.relationships).toHaveLength(0);

    // XML should have location attribute
    expect(result.xml).toContain('location="Sheet2!A1"');
    expect(result.xml).toContain('ref="A1"');
    expect(result.xml).not.toContain("r:id");
  });

  it("collects tooltip and display attributes", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Link",
      hyperlink: {
        target: "https://example.com",
        tooltip: "Click here",
        display: "Example Site",
      },
    });
    const sheet: WriteSheet = { name: "Sheet1", rows: [["Link"]], cells };
    const result = collectHyperlinks(sheet);

    expect(result.xml).toContain('tooltip="Click here"');
    expect(result.xml).toContain('display="Example Site"');
  });

  it("collects multiple hyperlinks with sequential rIds", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Link 1",
      hyperlink: { target: "https://example.com" },
    });
    cells.set("1,0", {
      value: "Link 2",
      hyperlink: { target: "https://other.com" },
    });
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Link 1"], ["Link 2"]],
      cells,
    };
    const result = collectHyperlinks(sheet);

    expect(result.relationships).toHaveLength(2);
    expect(result.relationships[0].id).toBe("rId1");
    expect(result.relationships[0].target).toBe("https://example.com");
    expect(result.relationships[1].id).toBe("rId2");
    expect(result.relationships[1].target).toBe("https://other.com");
  });

  it("handles mixed external and internal hyperlinks", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "External",
      hyperlink: { target: "https://example.com" },
    });
    cells.set("1,0", {
      value: "Internal",
      hyperlink: { target: "", location: "Sheet2!B5" },
    });
    cells.set("2,0", {
      value: "Another external",
      hyperlink: { target: "https://other.com" },
    });
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["External"], ["Internal"], ["Another external"]],
      cells,
    };
    const result = collectHyperlinks(sheet);

    // Only external hyperlinks generate relationships
    expect(result.relationships).toHaveLength(2);
    expect(result.relationships[0].target).toBe("https://example.com");
    expect(result.relationships[1].target).toBe("https://other.com");

    // Internal hyperlink should use location
    expect(result.xml).toContain('location="Sheet2!B5"');
  });
});

// ── Writing Tests ────────────────────────────────────────────────────

describe("XLSX hyperlink writing", () => {
  it("writes external URL hyperlink with .rels file", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Click me",
      hyperlink: { target: "https://example.com" },
    });

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Click me"]], cells }],
    });

    // Verify worksheet XML contains <hyperlinks> section
    const wsDoc = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const hyperlinks = findChild(wsDoc, "hyperlinks");
    expect(hyperlinks).toBeDefined();

    const hlElements = findChildren(hyperlinks, "hyperlink");
    expect(hlElements).toHaveLength(1);
    expect(hlElements[0].attrs["ref"]).toBe("A1");
    expect(hlElements[0].attrs["r:id"]).toBe("rId1");

    // Verify .rels file exists and has the hyperlink relationship
    expect(zipHas(data, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(true);
    const relsDoc = await parseXmlFromZip(data, "xl/worksheets/_rels/sheet1.xml.rels");
    const rels = findChildren(relsDoc, "Relationship");
    expect(rels).toHaveLength(1);
    expect(rels[0].attrs["Id"]).toBe("rId1");
    expect(rels[0].attrs["Target"]).toBe("https://example.com");
    expect(rels[0].attrs["TargetMode"]).toBe("External");
    expect(rels[0].attrs["Type"]).toContain("hyperlink");
  });

  it("writes internal hyperlink without .rels entry", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Go to Sheet2",
      hyperlink: { target: "", location: "Sheet2!A1" },
    });

    const data = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["Go to Sheet2"]], cells },
        { name: "Sheet2", rows: [["Target"]] },
      ],
    });

    // Verify worksheet XML contains <hyperlinks> with location attribute
    const wsDoc = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const hyperlinks = findChild(wsDoc, "hyperlinks");
    expect(hyperlinks).toBeDefined();

    const hlElements = findChildren(hyperlinks, "hyperlink");
    expect(hlElements).toHaveLength(1);
    expect(hlElements[0].attrs["ref"]).toBe("A1");
    expect(hlElements[0].attrs["location"]).toBe("Sheet2!A1");
    expect(hlElements[0].attrs["r:id"]).toBeUndefined();

    // No .rels file should be generated (no external links)
    expect(zipHas(data, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(false);
  });

  it("writes hyperlink with tooltip", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Hover me",
      hyperlink: { target: "https://example.com", tooltip: "Click here" },
    });

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hover me"]], cells }],
    });

    const wsDoc = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const hyperlinks = findChild(wsDoc, "hyperlinks");
    const hlElements = findChildren(hyperlinks, "hyperlink");
    expect(hlElements[0].attrs["tooltip"]).toBe("Click here");
  });

  it("writes hyperlink with display text", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Link",
      hyperlink: {
        target: "https://example.com",
        display: "Example Site",
      },
    });

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Link"]], cells }],
    });

    const wsDoc = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const hyperlinks = findChild(wsDoc, "hyperlinks");
    const hlElements = findChildren(hyperlinks, "hyperlink");
    expect(hlElements[0].attrs["display"]).toBe("Example Site");
  });

  it("writes multiple hyperlinks on same sheet", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Link 1",
      hyperlink: { target: "https://example.com" },
    });
    cells.set("1,0", {
      value: "Link 2",
      hyperlink: { target: "https://other.com" },
    });
    cells.set("2,1", {
      value: "Link 3",
      hyperlink: { target: "https://third.com", tooltip: "Third link" },
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Link 1"], ["Link 2"], [null, "Link 3"]],
          cells,
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const hyperlinks = findChild(wsDoc, "hyperlinks");
    const hlElements = findChildren(hyperlinks, "hyperlink");
    expect(hlElements).toHaveLength(3);

    // Verify .rels has 3 relationships
    const relsDoc = await parseXmlFromZip(data, "xl/worksheets/_rels/sheet1.xml.rels");
    const rels = findChildren(relsDoc, "Relationship");
    expect(rels).toHaveLength(3);
  });

  it("writes mixed external and internal hyperlinks", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "External link",
      hyperlink: { target: "https://example.com" },
    });
    cells.set("1,0", {
      value: "Internal link",
      hyperlink: { target: "", location: "Sheet2!A1", display: "Go to Sheet2" },
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["External link"], ["Internal link"]],
          cells,
        },
        { name: "Sheet2", rows: [["Target"]] },
      ],
    });

    const wsDoc = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const hyperlinks = findChild(wsDoc, "hyperlinks");
    const hlElements = findChildren(hyperlinks, "hyperlink");
    expect(hlElements).toHaveLength(2);

    // External should have r:id, internal should have location
    const external = hlElements.find((h: any) => h.attrs["r:id"]);
    const internal = hlElements.find((h: any) => h.attrs["location"]);
    expect(external).toBeDefined();
    expect(internal).toBeDefined();
    expect(internal.attrs["display"]).toBe("Go to Sheet2");

    // Only 1 relationship in .rels (the external one)
    const relsDoc = await parseXmlFromZip(data, "xl/worksheets/_rels/sheet1.xml.rels");
    const rels = findChildren(relsDoc, "Relationship");
    expect(rels).toHaveLength(1);
    expect(rels[0].attrs["Target"]).toBe("https://example.com");
  });

  it("does not generate .rels file when no hyperlinks exist", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Hello", "World"]] }],
    });

    expect(zipHas(data, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(false);
  });
});

// ── Reading Tests ────────────────────────────────────────────────────

describe("XLSX hyperlink reading", () => {
  /**
   * Helper to create an XLSX with hyperlinks in the raw XML
   * for testing the reader path independently.
   */
  async function createXlsxWithHyperlinks(options: {
    hyperlinks: Array<{
      ref: string;
      rId?: string;
      location?: string;
      tooltip?: string;
      display?: string;
    }>;
    rels?: Array<{
      id: string;
      target: string;
    }>;
    sharedStrings?: string[];
    cellData?: Array<{
      ref: string;
      value: string;
      type?: string;
      ssIndex?: number;
    }>;
  }): Promise<Uint8Array> {
    const writer = new ZipWriter();
    const enc = encoder;

    // [Content_Types].xml
    const hasSharedStrings = (options.sharedStrings?.length ?? 0) > 0;
    let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;
    if (hasSharedStrings) {
      contentTypes += `\n  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`;
    }
    contentTypes += `\n</Types>`;
    writer.add("[Content_Types].xml", enc.encode(contentTypes));

    // _rels/.rels
    writer.add(
      "_rels/.rels",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    );

    // xl/workbook.xml
    writer.add(
      "xl/workbook.xml",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    );

    // xl/_rels/workbook.xml.rels
    let wbRels = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
    if (hasSharedStrings) {
      wbRels += `\n  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
    }
    writer.add(
      "xl/_rels/workbook.xml.rels",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${wbRels}
</Relationships>`),
    );

    // xl/styles.xml
    writer.add(
      "xl/styles.xml",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`),
    );

    // xl/sharedStrings.xml
    if (hasSharedStrings) {
      const siElements = options.sharedStrings!.map((s) => `<si><t>${s}</t></si>`).join("");
      writer.add(
        "xl/sharedStrings.xml",
        enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${options.sharedStrings!.length}" uniqueCount="${options.sharedStrings!.length}">
  ${siElements}
</sst>`),
      );
    }

    // Build hyperlinks XML section
    let hyperlinksXml = "";
    if (options.hyperlinks.length > 0) {
      const hlParts = options.hyperlinks.map((hl) => {
        let attrs = `ref="${hl.ref}"`;
        if (hl.rId) attrs += ` r:id="${hl.rId}"`;
        if (hl.location) attrs += ` location="${hl.location}"`;
        if (hl.tooltip) attrs += ` tooltip="${hl.tooltip}"`;
        if (hl.display) attrs += ` display="${hl.display}"`;
        return `<hyperlink ${attrs}/>`;
      });
      hyperlinksXml = `<hyperlinks>${hlParts.join("")}</hyperlinks>`;
    }

    // Build cell data
    let sheetDataXml = "";
    if (options.cellData && options.cellData.length > 0) {
      const cellParts = options.cellData.map((cd) => {
        const typeAttr = cd.type ? ` t="${cd.type}"` : "";
        if (cd.type === "s" && cd.ssIndex !== undefined) {
          return `<c r="${cd.ref}"${typeAttr}><v>${cd.ssIndex}</v></c>`;
        }
        return `<c r="${cd.ref}"${typeAttr}><v>${cd.value}</v></c>`;
      });
      sheetDataXml = `<row r="1">${cellParts.join("")}</row>`;
    }

    // xl/worksheets/sheet1.xml
    writer.add(
      "xl/worksheets/sheet1.xml",
      enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>${sheetDataXml}</sheetData>
  ${hyperlinksXml}
</worksheet>`),
    );

    // xl/worksheets/_rels/sheet1.xml.rels (for external hyperlinks)
    if (options.rels && options.rels.length > 0) {
      const relParts = options.rels.map(
        (r) =>
          `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${r.target}" TargetMode="External"/>`,
      );
      writer.add(
        "xl/worksheets/_rels/sheet1.xml.rels",
        enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relParts.join("\n  ")}
</Relationships>`),
      );
    }

    return writer.build();
  }

  it("reads external hyperlink from worksheet XML + .rels", async () => {
    const xlsxData = await createXlsxWithHyperlinks({
      hyperlinks: [{ ref: "A1", rId: "rId1" }],
      rels: [{ id: "rId1", target: "https://example.com" }],
      sharedStrings: ["Click me"],
      cellData: [{ ref: "A1", value: "Click me", type: "s", ssIndex: 0 }],
    });

    const workbook = await readXlsx(xlsxData);
    const sheet = workbook.sheets[0];

    expect(sheet.cells).toBeDefined();
    const cell = sheet.cells!.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.hyperlink).toBeDefined();
    expect(cell!.hyperlink!.target).toBe("https://example.com");
  });

  it("reads internal hyperlink with location", async () => {
    const xlsxData = await createXlsxWithHyperlinks({
      hyperlinks: [{ ref: "A1", location: "Sheet2!A1", display: "Go to Sheet2" }],
      sharedStrings: ["Go to Sheet2"],
      cellData: [{ ref: "A1", value: "Go to Sheet2", type: "s", ssIndex: 0 }],
    });

    const workbook = await readXlsx(xlsxData);
    const sheet = workbook.sheets[0];

    expect(sheet.cells).toBeDefined();
    const cell = sheet.cells!.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.hyperlink).toBeDefined();
    expect(cell!.hyperlink!.location).toBe("Sheet2!A1");
    expect(cell!.hyperlink!.display).toBe("Go to Sheet2");
  });

  it("reads hyperlink tooltip", async () => {
    const xlsxData = await createXlsxWithHyperlinks({
      hyperlinks: [{ ref: "A1", rId: "rId1", tooltip: "Click here" }],
      rels: [{ id: "rId1", target: "https://example.com" }],
      sharedStrings: ["Link"],
      cellData: [{ ref: "A1", value: "Link", type: "s", ssIndex: 0 }],
    });

    const workbook = await readXlsx(xlsxData);
    const cell = workbook.sheets[0].cells!.get("0,0");
    expect(cell!.hyperlink!.tooltip).toBe("Click here");
    expect(cell!.hyperlink!.target).toBe("https://example.com");
  });

  it("reads multiple hyperlinks on same sheet", async () => {
    const xlsxData = await createXlsxWithHyperlinks({
      hyperlinks: [
        { ref: "A1", rId: "rId1" },
        { ref: "B1", rId: "rId2", tooltip: "Second link" },
      ],
      rels: [
        { id: "rId1", target: "https://example.com" },
        { id: "rId2", target: "https://other.com" },
      ],
      sharedStrings: ["Link 1", "Link 2"],
      cellData: [
        { ref: "A1", value: "Link 1", type: "s", ssIndex: 0 },
        { ref: "B1", value: "Link 2", type: "s", ssIndex: 1 },
      ],
    });

    const workbook = await readXlsx(xlsxData);
    const sheet = workbook.sheets[0];

    const cellA1 = sheet.cells!.get("0,0");
    expect(cellA1!.hyperlink!.target).toBe("https://example.com");

    const cellB1 = sheet.cells!.get("0,1");
    expect(cellB1!.hyperlink!.target).toBe("https://other.com");
    expect(cellB1!.hyperlink!.tooltip).toBe("Second link");
  });

  it("reads hyperlink on cell with no prior detail (creates Cell entry)", async () => {
    // Cell has a numeric value (no shared string), but has a hyperlink
    const xlsxData = await createXlsxWithHyperlinks({
      hyperlinks: [{ ref: "A1", rId: "rId1" }],
      rels: [{ id: "rId1", target: "https://example.com" }],
      cellData: [{ ref: "A1", value: "42" }],
    });

    const workbook = await readXlsx(xlsxData);
    const sheet = workbook.sheets[0];

    const cell = sheet.cells!.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.hyperlink!.target).toBe("https://example.com");
    expect(cell!.value).toBe(42);
  });
});

// ── Round-trip Tests ─────────────────────────────────────────────────

describe("XLSX hyperlink round-trip", () => {
  it("round-trips external hyperlink", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Example",
      hyperlink: {
        target: "https://example.com",
        tooltip: "Visit Example",
      },
    });

    const written = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Example"]], cells }],
    });

    const workbook = await readXlsx(written);
    const sheet = workbook.sheets[0];

    expect(sheet.cells).toBeDefined();
    const cell = sheet.cells!.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.hyperlink).toBeDefined();
    expect(cell!.hyperlink!.target).toBe("https://example.com");
    expect(cell!.hyperlink!.tooltip).toBe("Visit Example");
  });

  it("round-trips internal hyperlink", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Go to Sheet2",
      hyperlink: {
        target: "",
        location: "Sheet2!A1",
        display: "Navigate",
      },
    });

    const written = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["Go to Sheet2"]], cells },
        { name: "Sheet2", rows: [["Target"]] },
      ],
    });

    const workbook = await readXlsx(written);
    const sheet = workbook.sheets[0];

    const cell = sheet.cells!.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.hyperlink).toBeDefined();
    expect(cell!.hyperlink!.location).toBe("Sheet2!A1");
    expect(cell!.hyperlink!.display).toBe("Navigate");
  });

  it("round-trips mixed hyperlinks", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "External link",
      hyperlink: { target: "https://example.com", tooltip: "External" },
    });
    cells.set("1,0", {
      value: "Internal link",
      hyperlink: { target: "", location: "Sheet2!B2", display: "Internal" },
    });
    cells.set("2,0", {
      value: "Another external",
      hyperlink: { target: "https://other.com" },
    });

    const written = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["External link"], ["Internal link"], ["Another external"]],
          cells,
        },
        { name: "Sheet2", rows: [["x"]] },
      ],
    });

    const workbook = await readXlsx(written);
    const sheet = workbook.sheets[0];

    // External
    const cellA1 = sheet.cells!.get("0,0");
    expect(cellA1!.hyperlink!.target).toBe("https://example.com");
    expect(cellA1!.hyperlink!.tooltip).toBe("External");

    // Internal
    const cellA2 = sheet.cells!.get("1,0");
    expect(cellA2!.hyperlink!.location).toBe("Sheet2!B2");
    expect(cellA2!.hyperlink!.display).toBe("Internal");

    // Another external
    const cellA3 = sheet.cells!.get("2,0");
    expect(cellA3!.hyperlink!.target).toBe("https://other.com");
  });

  it("round-trips hyperlinks on multiple sheets", async () => {
    const cells1 = new Map<string, Partial<Cell>>();
    cells1.set("0,0", {
      value: "Sheet1 Link",
      hyperlink: { target: "https://sheet1.com" },
    });

    const cells2 = new Map<string, Partial<Cell>>();
    cells2.set("0,0", {
      value: "Sheet2 Link",
      hyperlink: { target: "https://sheet2.com" },
    });

    const written = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["Sheet1 Link"]], cells: cells1 },
        { name: "Sheet2", rows: [["Sheet2 Link"]], cells: cells2 },
      ],
    });

    // Verify both sheets have their own .rels files
    expect(zipHas(written, "xl/worksheets/_rels/sheet1.xml.rels")).toBe(true);
    expect(zipHas(written, "xl/worksheets/_rels/sheet2.xml.rels")).toBe(true);

    const workbook = await readXlsx(written);

    const cell1 = workbook.sheets[0].cells!.get("0,0");
    expect(cell1!.hyperlink!.target).toBe("https://sheet1.com");

    const cell2 = workbook.sheets[1].cells!.get("0,0");
    expect(cell2!.hyperlink!.target).toBe("https://sheet2.com");
  });

  it("preserves cell value alongside hyperlink", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Visit us",
      hyperlink: { target: "https://example.com" },
    });

    const written = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["Visit us"]], cells }],
    });

    const workbook = await readXlsx(written);
    const sheet = workbook.sheets[0];

    // Value should be preserved
    expect(sheet.rows[0][0]).toBe("Visit us");

    // Hyperlink should also be present
    const cell = sheet.cells!.get("0,0");
    expect(cell!.hyperlink!.target).toBe("https://example.com");
    expect(cell!.value).toBe("Visit us");
  });
});
