import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { writeCoreProperties, writeAppProperties } from "../src/xlsx/doc-props-writer";
import { parseCoreProperties, parseAppProperties } from "../src/xlsx/doc-props-reader";
import type { WriteOptions, WorkbookProperties } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

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

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("");
}

// ── Unit Tests: Writer ──────────────────────────────────────────────

describe("writeCoreProperties", () => {
  it("generates minimal core.xml with modified date when no props", () => {
    const xml = writeCoreProperties();
    const doc = parseXml(xml);

    // Should have cp:coreProperties as root
    expect(doc.local).toBe("coreProperties");

    // Should have a modified element
    const modified = findChild(doc, "modified");
    expect(modified).toBeDefined();
    expect(getElementText(modified)).toMatch(/^\d{4}-\d{2}-\d{2}T/);
  });

  it("includes all properties", () => {
    const props: WorkbookProperties = {
      title: "My Workbook",
      subject: "Test Report",
      creator: "John Doe",
      keywords: "test, report",
      description: "A test workbook",
      lastModifiedBy: "Jane Doe",
      category: "Reports",
      created: new Date("2026-01-15T10:00:00Z"),
      modified: new Date("2026-03-24T12:00:00Z"),
    };

    const xml = writeCoreProperties(props);
    const doc = parseXml(xml);

    expect(getElementText(findChild(doc, "title"))).toBe("My Workbook");
    expect(getElementText(findChild(doc, "subject"))).toBe("Test Report");
    expect(getElementText(findChild(doc, "creator"))).toBe("John Doe");
    expect(getElementText(findChild(doc, "keywords"))).toBe("test, report");
    expect(getElementText(findChild(doc, "description"))).toBe("A test workbook");
    expect(getElementText(findChild(doc, "lastModifiedBy"))).toBe("Jane Doe");
    expect(getElementText(findChild(doc, "category"))).toBe("Reports");

    const created = findChild(doc, "created");
    expect(getElementText(created)).toBe("2026-01-15T10:00:00Z");
    expect(created.attrs["xsi:type"]).toBe("dcterms:W3CDTF");

    const modified = findChild(doc, "modified");
    expect(getElementText(modified)).toBe("2026-03-24T12:00:00Z");
  });

  it("escapes XML special characters in values", () => {
    const xml = writeCoreProperties({
      title: "Report <Q1> & Q2",
      creator: 'John "DJ" Doe',
    });
    const doc = parseXml(xml);
    expect(getElementText(findChild(doc, "title"))).toBe("Report <Q1> & Q2");
    expect(getElementText(findChild(doc, "creator"))).toBe('John "DJ" Doe');
  });
});

describe("writeAppProperties", () => {
  it("always includes Application: defter", () => {
    const xml = writeAppProperties();
    const doc = parseXml(xml);

    const app = findChild(doc, "Application");
    expect(app).toBeDefined();
    expect(getElementText(app)).toBe("defter");
  });

  it("includes company and manager", () => {
    const xml = writeAppProperties({
      company: "Acme Inc",
      manager: "Bob Smith",
    });
    const doc = parseXml(xml);

    expect(getElementText(findChild(doc, "Company"))).toBe("Acme Inc");
    expect(getElementText(findChild(doc, "Manager"))).toBe("Bob Smith");
  });
});

// ── Unit Tests: Reader ──────────────────────────────────────────────

describe("parseCoreProperties", () => {
  it("parses all core properties fields", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>My Workbook</dc:title>
  <dc:subject>Report</dc:subject>
  <dc:creator>John Doe</dc:creator>
  <cp:keywords>finance, report</cp:keywords>
  <dc:description>Monthly report</dc:description>
  <cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-15T10:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-03-24T12:00:00Z</dcterms:modified>
  <cp:category>Reports</cp:category>
</cp:coreProperties>`;

    const props = parseCoreProperties(xml);

    expect(props.title).toBe("My Workbook");
    expect(props.subject).toBe("Report");
    expect(props.creator).toBe("John Doe");
    expect(props.keywords).toBe("finance, report");
    expect(props.description).toBe("Monthly report");
    expect(props.lastModifiedBy).toBe("Jane Doe");
    expect(props.category).toBe("Reports");
    expect(props.created).toEqual(new Date("2026-01-15T10:00:00Z"));
    expect(props.modified).toEqual(new Date("2026-03-24T12:00:00Z"));
  });

  it("handles empty core.xml gracefully", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`;
    const props = parseCoreProperties(xml);
    expect(Object.keys(props)).toHaveLength(0);
  });
});

describe("parseAppProperties", () => {
  it("parses company and manager", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>defter</Application>
  <Company>Acme Inc</Company>
  <Manager>Bob Smith</Manager>
</Properties>`;

    const props = parseAppProperties(xml);
    expect(props.company).toBe("Acme Inc");
    expect(props.manager).toBe("Bob Smith");
  });
});

// ── Integration Tests ───────────────────────────────────────────────

describe("XLSX document properties integration", () => {
  it("writes properties and includes docProps files in ZIP", async () => {
    const options: WriteOptions = {
      sheets: [{ name: "Sheet1", rows: [["Hello"]] }],
      properties: {
        title: "Test Workbook",
        creator: "Test User",
      },
    };

    const data = await writeXlsx(options);
    const zip = new ZipReader(data);

    // Verify docProps files exist
    expect(zip.has("docProps/core.xml")).toBe(true);
    expect(zip.has("docProps/app.xml")).toBe(true);

    // Verify core.xml content
    const coreDoc = await parseXmlFromZip(data, "docProps/core.xml");
    expect(getElementText(findChild(coreDoc, "title"))).toBe("Test Workbook");
    expect(getElementText(findChild(coreDoc, "creator"))).toBe("Test User");

    // Verify app.xml includes Application
    const appDoc = await parseXmlFromZip(data, "docProps/app.xml");
    expect(getElementText(findChild(appDoc, "Application"))).toBe("defter");
  });

  it("writes all property fields", async () => {
    const created = new Date("2026-01-15T10:00:00Z");
    const modified = new Date("2026-03-24T12:00:00Z");

    const options: WriteOptions = {
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
      properties: {
        title: "Full Props Test",
        subject: "Testing",
        creator: "Author",
        keywords: "test, keywords",
        description: "A description",
        lastModifiedBy: "Modifier",
        category: "Category",
        created,
        modified,
        company: "Test Corp",
        manager: "Manager",
      },
    };

    const data = await writeXlsx(options);

    const coreDoc = await parseXmlFromZip(data, "docProps/core.xml");
    expect(getElementText(findChild(coreDoc, "title"))).toBe("Full Props Test");
    expect(getElementText(findChild(coreDoc, "subject"))).toBe("Testing");
    expect(getElementText(findChild(coreDoc, "creator"))).toBe("Author");
    expect(getElementText(findChild(coreDoc, "keywords"))).toBe("test, keywords");
    expect(getElementText(findChild(coreDoc, "description"))).toBe("A description");
    expect(getElementText(findChild(coreDoc, "lastModifiedBy"))).toBe("Modifier");
    expect(getElementText(findChild(coreDoc, "category"))).toBe("Category");
    expect(getElementText(findChild(coreDoc, "created"))).toBe("2026-01-15T10:00:00Z");
    expect(getElementText(findChild(coreDoc, "modified"))).toBe("2026-03-24T12:00:00Z");

    const appDoc = await parseXmlFromZip(data, "docProps/app.xml");
    expect(getElementText(findChild(appDoc, "Company"))).toBe("Test Corp");
    expect(getElementText(findChild(appDoc, "Manager"))).toBe("Manager");
  });

  it("generates minimal core.xml with modified date when no properties provided", async () => {
    const options: WriteOptions = {
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
    };

    const data = await writeXlsx(options);
    const zip = new ZipReader(data);

    // docProps should still be generated
    expect(zip.has("docProps/core.xml")).toBe(true);
    expect(zip.has("docProps/app.xml")).toBe(true);

    // core.xml should have a modified date
    const coreDoc = await parseXmlFromZip(data, "docProps/core.xml");
    const modified = findChild(coreDoc, "modified");
    expect(modified).toBeDefined();

    // app.xml should have Application: defter
    const appDoc = await parseXmlFromZip(data, "docProps/app.xml");
    expect(getElementText(findChild(appDoc, "Application"))).toBe("defter");
  });

  it("includes docProps content type overrides", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
      properties: { title: "Test" },
    });

    const ctDoc = await parseXmlFromZip(data, "[Content_Types].xml");
    const overrides = findChildren(ctDoc, "Override");

    const coreOverride = overrides.find((o: any) => o.attrs["PartName"] === "/docProps/core.xml");
    expect(coreOverride).toBeDefined();
    expect(coreOverride.attrs["ContentType"]).toBe(
      "application/vnd.openxmlformats-package.core-properties+xml",
    );

    const appOverride = overrides.find((o: any) => o.attrs["PartName"] === "/docProps/app.xml");
    expect(appOverride).toBeDefined();
    expect(appOverride.attrs["ContentType"]).toBe(
      "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    );
  });

  it("includes docProps relationships in _rels/.rels", async () => {
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
      properties: { title: "Test" },
    });

    const relsDoc = await parseXmlFromZip(data, "_rels/.rels");
    const rels = findChildren(relsDoc, "Relationship");

    // Should have workbook + core + app relationships
    expect(rels.length).toBe(3);

    const coreRel = rels.find((r: any) => r.attrs["Target"] === "docProps/core.xml");
    expect(coreRel).toBeDefined();
    expect(coreRel.attrs["Type"]).toBe(
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    );

    const appRel = rels.find((r: any) => r.attrs["Target"] === "docProps/app.xml");
    expect(appRel).toBeDefined();
    expect(appRel.attrs["Type"]).toBe(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    );
  });

  it("round-trips properties: write then read", async () => {
    const created = new Date("2026-01-15T10:00:00Z");
    const modified = new Date("2026-03-24T12:00:00Z");

    const properties: WorkbookProperties = {
      title: "Round Trip Test",
      subject: "Testing",
      creator: "Author Name",
      keywords: "round, trip",
      description: "Description here",
      lastModifiedBy: "Last Modifier",
      category: "Testing",
      created,
      modified,
      company: "Round Trip Corp",
      manager: "The Manager",
    };

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
      properties,
    });

    const workbook = await readXlsx(data);

    expect(workbook.properties).toBeDefined();
    expect(workbook.properties!.title).toBe("Round Trip Test");
    expect(workbook.properties!.subject).toBe("Testing");
    expect(workbook.properties!.creator).toBe("Author Name");
    expect(workbook.properties!.keywords).toBe("round, trip");
    expect(workbook.properties!.description).toBe("Description here");
    expect(workbook.properties!.lastModifiedBy).toBe("Last Modifier");
    expect(workbook.properties!.category).toBe("Testing");
    expect(workbook.properties!.created).toEqual(created);
    expect(workbook.properties!.modified).toEqual(modified);
    expect(workbook.properties!.company).toBe("Round Trip Corp");
    expect(workbook.properties!.manager).toBe("The Manager");
  });

  it("reads workbook without properties (backward compatibility)", async () => {
    // Write with no properties, then read
    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["hello"]] }],
    });

    const workbook = await readXlsx(data);
    // properties should be defined since we now always generate docProps
    expect(workbook.properties).toBeDefined();
    // Should at least have a modified date
    expect(workbook.properties!.modified).toBeInstanceOf(Date);
  });

  it("writes properties with dates as Date objects", async () => {
    const created = new Date("2025-06-15T08:30:00Z");
    const modified = new Date("2025-12-25T18:00:00Z");

    const data = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
      properties: { created, modified },
    });

    const coreDoc = await parseXmlFromZip(data, "docProps/core.xml");
    const createdEl = findChild(coreDoc, "created");
    const modifiedEl = findChild(coreDoc, "modified");

    expect(getElementText(createdEl)).toBe("2025-06-15T08:30:00Z");
    expect(getElementText(modifiedEl)).toBe("2025-12-25T18:00:00Z");

    // Verify xsi:type attribute
    expect(createdEl.attrs["xsi:type"]).toBe("dcterms:W3CDTF");
    expect(modifiedEl.attrs["xsi:type"]).toBe("dcterms:W3CDTF");
  });
});
