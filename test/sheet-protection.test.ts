import { describe, it, expect } from "vitest";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { createStylesCollector } from "../src/xlsx/styles-writer";
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer";
import { hashSheetPassword } from "../src/xlsx/password";
import type { WriteSheet, SheetProtection } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector();
  const ss = createSharedStrings();
  const result = writeWorksheetXml(sheet, styles, ss);
  return result.xml;
}

function writeStylesXml(sheet: WriteSheet): string {
  const styles = createStylesCollector();
  const ss = createSharedStrings();
  writeWorksheetXml(sheet, styles, ss);
  return styles.toXml();
}

function parseSheet(xml: string) {
  return parseXml(xml);
}

// ── Password Hash Tests ──────────────────────────────────────────────

describe("hashSheetPassword", () => {
  it("hashes a known password to expected hex value", () => {
    // Known test vector: "password" → "83AF"
    expect(hashSheetPassword("password")).toBe("83AF");
  });

  it("hashes single character", () => {
    // "a" → known hash
    const result = hashSheetPassword("a");
    expect(result).toMatch(/^[0-9A-F]{4}$/);
    expect(result.length).toBe(4);
  });

  it("hashes empty string", () => {
    // Empty string should still produce a 4-char hex
    const result = hashSheetPassword("");
    expect(result).toMatch(/^[0-9A-F]{4}$/);
  });

  it("produces different hashes for different passwords", () => {
    const h1 = hashSheetPassword("abc");
    const h2 = hashSheetPassword("xyz");
    expect(h1).not.toBe(h2);
  });

  it("produces consistent hash for same password", () => {
    const h1 = hashSheetPassword("test123");
    const h2 = hashSheetPassword("test123");
    expect(h1).toBe(h2);
  });

  it("always returns uppercase hex string of length 4", () => {
    const passwords = ["a", "abc", "password", "P@ssw0rd!", "123456", "very long password here"];
    for (const pw of passwords) {
      const hash = hashSheetPassword(pw);
      expect(hash).toMatch(/^[0-9A-F]{4}$/);
    }
  });

  it("hashes '1234' to known value", () => {
    // Another known test vector: "1234" → "CC3D"
    expect(hashSheetPassword("1234")).toBe("CC3D");
  });
});

// ── Sheet Protection Writing Tests ──────────────────────────────────

describe("sheet protection — writing", () => {
  it("writes basic sheet protection (sheet=true)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection: {
        sheet: true,
      },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");
    expect(sp).toBeDefined();
    expect(sp.attrs["sheet"]).toBe("1");
  });

  it("writes protection with password hash", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection: {
        sheet: true,
        password: "password",
      },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");
    expect(sp).toBeDefined();
    expect(sp.attrs["password"]).toBe("83AF");
    expect(sp.attrs["sheet"]).toBe("1");
  });

  it("writes protection with objects and scenarios", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection: {
        sheet: true,
        objects: true,
        scenarios: true,
      },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");
    expect(sp.attrs["sheet"]).toBe("1");
    expect(sp.attrs["objects"]).toBe("1");
    expect(sp.attrs["scenarios"]).toBe("1");
  });

  it("writes granular allow options (sort=true, autoFilter=true)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection: {
        sheet: true,
        sort: true,
        autoFilter: true,
      },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");
    expect(sp.attrs["sheet"]).toBe("1");
    // sort=true (allowed) → "0" in XML (not prohibited)
    expect(sp.attrs["sort"]).toBe("0");
    // autoFilter=true (allowed) → "0" in XML (not prohibited)
    expect(sp.attrs["autoFilter"]).toBe("0");
  });

  it("writes disallowed options correctly (formatCells=false)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection: {
        sheet: true,
        formatCells: false,
        insertRows: false,
        deleteRows: false,
      },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");
    // false (disallowed) → "1" in XML (prohibited)
    expect(sp.attrs["formatCells"]).toBe("1");
    expect(sp.attrs["insertRows"]).toBe("1");
    expect(sp.attrs["deleteRows"]).toBe("1");
  });

  it("does not emit sheetProtection when no protection property", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");
    expect(sp).toBeUndefined();
  });

  it("defaults sheet to protected when protection object exists", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection: {},
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");
    expect(sp).toBeDefined();
    // sheet defaults to "1" when protection object is present
    expect(sp.attrs["sheet"]).toBe("1");
  });

  it("emits sheetProtection after sheetData (OOXML spec order)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection: { sheet: true },
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    // Collect element names in order
    const childNames: string[] = [];
    for (const child of doc.children) {
      if (typeof child !== "string") {
        childNames.push((child as any).local || (child as any).tag);
      }
    }

    const dataIdx = childNames.indexOf("sheetData");
    const protIdx = childNames.indexOf("sheetProtection");

    expect(dataIdx).toBeGreaterThanOrEqual(0);
    expect(protIdx).toBeGreaterThanOrEqual(0);
    // Per ECMA-376: sheetProtection comes after sheetData
    expect(protIdx).toBeGreaterThan(dataIdx);
  });

  it("maps all protection options correctly", () => {
    const protection: SheetProtection = {
      sheet: true,
      objects: true,
      scenarios: true,
      password: "test",
      selectLockedCells: true,
      selectUnlockedCells: false,
      formatCells: true,
      formatColumns: false,
      formatRows: true,
      insertColumns: false,
      insertRows: true,
      insertHyperlinks: false,
      deleteColumns: true,
      deleteRows: false,
      sort: true,
      autoFilter: false,
      pivotTables: true,
    };

    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Data"]],
      protection,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);

    const sp = findChild(doc, "sheetProtection");

    // Direct flags
    expect(sp.attrs["sheet"]).toBe("1");
    expect(sp.attrs["objects"]).toBe("1");
    expect(sp.attrs["scenarios"]).toBe("1");
    expect(sp.attrs["password"]).toBeDefined();

    // Allow = true → "0" (not prohibited)
    expect(sp.attrs["selectLockedCells"]).toBe("0");
    expect(sp.attrs["formatCells"]).toBe("0");
    expect(sp.attrs["formatRows"]).toBe("0");
    expect(sp.attrs["insertRows"]).toBe("0");
    expect(sp.attrs["deleteColumns"]).toBe("0");
    expect(sp.attrs["sort"]).toBe("0");
    expect(sp.attrs["pivotTables"]).toBe("0");

    // Allow = false → "1" (prohibited)
    expect(sp.attrs["selectUnlockedCells"]).toBe("1");
    expect(sp.attrs["formatColumns"]).toBe("1");
    expect(sp.attrs["insertColumns"]).toBe("1");
    expect(sp.attrs["insertHyperlinks"]).toBe("1");
    expect(sp.attrs["deleteRows"]).toBe("1");
    expect(sp.attrs["autoFilter"]).toBe("1");
  });
});

// ── Cell Protection Writing Tests ───────────────────────────────────

describe("cell protection — writing", () => {
  it("writes locked=false on specific cells", () => {
    const cells = new Map<string, any>();
    cells.set("0,0", {
      value: "Editable",
      style: { protection: { locked: false } },
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Editable"]],
      cells,
      protection: { sheet: true },
    };

    const stylesXml = writeStylesXml(sheet);
    const stylesDoc = parseXml(stylesXml);

    // Find cellXfs
    const cellXfs = findChild(stylesDoc, "cellXfs");
    expect(cellXfs).toBeDefined();

    const xfElements = findChildren(cellXfs, "xf");
    // Should have at least 2: default + the unlocked one
    expect(xfElements.length).toBeGreaterThanOrEqual(2);

    // Find xf with applyProtection
    const protectedXf = xfElements.find((xf: any) => xf.attrs["applyProtection"] === "true");
    expect(protectedXf).toBeDefined();

    const protElement = findChild(protectedXf, "protection");
    expect(protElement).toBeDefined();
    expect(protElement.attrs["locked"]).toBe("0");
  });

  it("writes hidden=true on cells", () => {
    const cells = new Map<string, any>();
    cells.set("0,0", {
      value: "=SUM(A1:A10)",
      style: { protection: { hidden: true } },
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Formula"]],
      cells,
    };

    const stylesXml = writeStylesXml(sheet);
    const stylesDoc = parseXml(stylesXml);

    const cellXfs = findChild(stylesDoc, "cellXfs");
    const xfElements = findChildren(cellXfs, "xf");

    const protectedXf = xfElements.find((xf: any) => xf.attrs["applyProtection"] === "true");
    expect(protectedXf).toBeDefined();

    const protElement = findChild(protectedXf, "protection");
    expect(protElement).toBeDefined();
    expect(protElement.attrs["hidden"]).toBe("1");
  });

  it("writes both locked and hidden on cells", () => {
    const cells = new Map<string, any>();
    cells.set("0,0", {
      value: "Secret",
      style: { protection: { locked: true, hidden: true } },
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Secret"]],
      cells,
    };

    const stylesXml = writeStylesXml(sheet);
    const stylesDoc = parseXml(stylesXml);

    const cellXfs = findChild(stylesDoc, "cellXfs");
    const xfElements = findChildren(cellXfs, "xf");

    const protectedXf = xfElements.find((xf: any) => xf.attrs["applyProtection"] === "true");
    expect(protectedXf).toBeDefined();

    const protElement = findChild(protectedXf, "protection");
    expect(protElement).toBeDefined();
    expect(protElement.attrs["locked"]).toBe("1");
    expect(protElement.attrs["hidden"]).toBe("1");
  });

  it("default cells have no explicit protection (all locked by default)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Normal cell"]],
    };

    const stylesXml = writeStylesXml(sheet);
    const stylesDoc = parseXml(stylesXml);

    const cellXfs = findChild(stylesDoc, "cellXfs");
    const xfElements = findChildren(cellXfs, "xf");

    // Default xf (index 0) should NOT have applyProtection
    const defaultXf = xfElements[0];
    expect(defaultXf.attrs["applyProtection"]).toBeUndefined();
  });
});

// ── Round-trip Tests ────────────────────────────────────────────────

describe("sheet protection — round-trip", () => {
  it("round-trips basic sheet protection", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Protected",
          rows: [["Data"]],
          protection: {
            sheet: true,
          },
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets.length).toBe(1);

    const sheet = workbook.sheets[0];
    expect(sheet.protection).toBeDefined();
    expect(sheet.protection!.sheet).toBe(true);
  });

  it("round-trips protection with granular options", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          protection: {
            sheet: true,
            sort: true,
            autoFilter: true,
            formatCells: false,
            insertRows: false,
          },
        },
      ],
    });

    const workbook = await readXlsx(data);
    const prot = workbook.sheets[0].protection!;

    expect(prot.sheet).toBe(true);
    expect(prot.sort).toBe(true);
    expect(prot.autoFilter).toBe(true);
    expect(prot.formatCells).toBe(false);
    expect(prot.insertRows).toBe(false);
  });

  it("round-trips protection with objects and scenarios", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          protection: {
            sheet: true,
            objects: true,
            scenarios: true,
          },
        },
      ],
    });

    const workbook = await readXlsx(data);
    const prot = workbook.sheets[0].protection!;

    expect(prot.sheet).toBe(true);
    expect(prot.objects).toBe(true);
    expect(prot.scenarios).toBe(true);
  });

  it("round-trips all protection options", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
          protection: {
            sheet: true,
            objects: true,
            scenarios: true,
            selectLockedCells: true,
            selectUnlockedCells: false,
            formatCells: true,
            formatColumns: false,
            formatRows: true,
            insertColumns: false,
            insertRows: true,
            insertHyperlinks: false,
            deleteColumns: true,
            deleteRows: false,
            sort: true,
            autoFilter: false,
            pivotTables: true,
          },
        },
      ],
    });

    const workbook = await readXlsx(data);
    const prot = workbook.sheets[0].protection!;

    expect(prot.sheet).toBe(true);
    expect(prot.objects).toBe(true);
    expect(prot.scenarios).toBe(true);
    expect(prot.selectLockedCells).toBe(true);
    expect(prot.selectUnlockedCells).toBe(false);
    expect(prot.formatCells).toBe(true);
    expect(prot.formatColumns).toBe(false);
    expect(prot.formatRows).toBe(true);
    expect(prot.insertColumns).toBe(false);
    expect(prot.insertRows).toBe(true);
    expect(prot.insertHyperlinks).toBe(false);
    expect(prot.deleteColumns).toBe(true);
    expect(prot.deleteRows).toBe(false);
    expect(prot.sort).toBe(true);
    expect(prot.autoFilter).toBe(false);
    expect(prot.pivotTables).toBe(true);
  });

  it("unprotected sheet has no protection property", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Data"]],
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets[0].protection).toBeUndefined();
  });

  it("round-trips cell protection (locked=false) with readStyles", async () => {
    const cells = new Map<string, any>();
    cells.set("0,0", {
      value: "Editable",
      style: { protection: { locked: false } },
    });
    cells.set("0,1", {
      value: "Locked",
      style: { protection: { locked: true } },
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Editable", "Locked"]],
          cells,
          protection: { sheet: true },
        },
      ],
    });

    const workbook = await readXlsx(data, { readStyles: true });
    const sheet = workbook.sheets[0];

    // Sheet protection should be present
    expect(sheet.protection).toBeDefined();
    expect(sheet.protection!.sheet).toBe(true);

    // Cell A1 (0,0) should have locked=false
    const cellA1 = sheet.cells?.get("0,0");
    expect(cellA1).toBeDefined();
    expect(cellA1!.style?.protection?.locked).toBe(false);

    // Cell B1 (0,1) should have locked=true
    const cellB1 = sheet.cells?.get("0,1");
    expect(cellB1).toBeDefined();
    expect(cellB1!.style?.protection?.locked).toBe(true);
  });

  it("round-trips cell hidden protection with readStyles", async () => {
    const cells = new Map<string, any>();
    cells.set("0,0", {
      value: "Hidden formula",
      style: { protection: { hidden: true, locked: true } },
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Hidden formula"]],
          cells,
          protection: { sheet: true },
        },
      ],
    });

    const workbook = await readXlsx(data, { readStyles: true });
    const cell = workbook.sheets[0].cells?.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.style?.protection?.hidden).toBe(true);
    expect(cell!.style?.protection?.locked).toBe(true);
  });

  it("protection on multiple sheets", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Protected",
          rows: [["Data"]],
          protection: { sheet: true, sort: true },
        },
        {
          name: "Unprotected",
          rows: [["Data"]],
        },
        {
          name: "Also Protected",
          rows: [["Data"]],
          protection: { sheet: true, autoFilter: true, formatCells: false },
        },
      ],
    });

    const workbook = await readXlsx(data);

    expect(workbook.sheets[0].protection).toBeDefined();
    expect(workbook.sheets[0].protection!.sheet).toBe(true);
    expect(workbook.sheets[0].protection!.sort).toBe(true);

    expect(workbook.sheets[1].protection).toBeUndefined();

    expect(workbook.sheets[2].protection).toBeDefined();
    expect(workbook.sheets[2].protection!.sheet).toBe(true);
    expect(workbook.sheets[2].protection!.autoFilter).toBe(true);
    expect(workbook.sheets[2].protection!.formatCells).toBe(false);
  });
});
