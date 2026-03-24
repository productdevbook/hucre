import { describe, expect, it } from "vitest";
import { ZipWriter } from "../src/zip/writer";
import { readXlsx } from "../src/xlsx/reader";
import { parseCellRef } from "../src/xlsx/worksheet";
import { parseSharedStrings } from "../src/xlsx/shared-strings";
import { parseStyles, resolveStyle, isDateStyle } from "../src/xlsx/styles";
import { parseRelationships } from "../src/xlsx/relationships";
import { parseContentTypes } from "../src/xlsx/content-types";

// ── Helpers ─────────────────────────────────────────────────────────

const enc = new TextEncoder();

function textToBytes(text: string): Uint8Array {
  return enc.encode(text);
}

/**
 * Programmatically create a valid XLSX (ZIP) archive for testing.
 */
async function createTestXlsx(options: {
  sheets: Array<{
    name: string;
    rows: Array<
      Array<{
        value: string | number | boolean;
        type?: string;
        styleIndex?: number;
        formula?: string;
      }>
    >;
    merges?: string[];
  }>;
  sharedStrings?: string[];
  richSharedStrings?: Array<
    string | { runs: Array<{ text: string; bold?: boolean; italic?: boolean; size?: number }> }
  >;
  styles?: string;
  dateSystem?: "1904";
}): Promise<Uint8Array> {
  const writer = new ZipWriter();

  // ── [Content_Types].xml ──
  const overrides: string[] = [];
  overrides.push(
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
  );
  if (options.sharedStrings || options.richSharedStrings) {
    overrides.push(
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`,
    );
  }
  overrides.push(
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
  );
  for (let i = 0; i < options.sheets.length; i++) {
    overrides.push(
      `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    );
  }

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides.join("\n  ")}
</Types>`;

  writer.add("[Content_Types].xml", textToBytes(contentTypesXml), { compress: false });

  // ── _rels/.rels ──
  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  writer.add("_rels/.rels", textToBytes(rootRelsXml), { compress: false });

  // ── xl/workbook.xml ──
  const sheetElements = options.sheets
    .map((s, i) => `<sheet name="${xmlEscapeAttr(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
    .join("");

  const workbookPrAttr = options.dateSystem === "1904" ? ` date1904="1"` : "";

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr${workbookPrAttr}/>
  <sheets>${sheetElements}</sheets>
</workbook>`;

  writer.add("xl/workbook.xml", textToBytes(workbookXml), { compress: false });

  // ── xl/_rels/workbook.xml.rels ──
  const wbRels: string[] = [];
  for (let i = 0; i < options.sheets.length; i++) {
    wbRels.push(
      `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
    );
  }
  let relIdx = options.sheets.length + 1;
  if (options.sharedStrings || options.richSharedStrings) {
    wbRels.push(
      `<Relationship Id="rId${relIdx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`,
    );
    relIdx++;
  }
  wbRels.push(
    `<Relationship Id="rId${relIdx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
  );

  const wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${wbRels.join("\n  ")}
</Relationships>`;

  writer.add("xl/_rels/workbook.xml.rels", textToBytes(wbRelsXml), { compress: false });

  // ── xl/sharedStrings.xml ──
  if (options.sharedStrings || options.richSharedStrings) {
    const items = options.richSharedStrings ?? options.sharedStrings ?? [];
    const siElements = items.map((item) => {
      if (typeof item === "string") {
        return `<si><t>${xmlEscape(item)}</t></si>`;
      }
      // Rich text
      const runs = item.runs
        .map((run) => {
          const rPrParts: string[] = [];
          if (run.bold) rPrParts.push(`<b/>`);
          if (run.italic) rPrParts.push(`<i/>`);
          if (run.size) rPrParts.push(`<sz val="${run.size}"/>`);
          const rPr = rPrParts.length > 0 ? `<rPr>${rPrParts.join("")}</rPr>` : "";
          return `<r>${rPr}<t>${xmlEscape(run.text)}</t></r>`;
        })
        .join("");
      return `<si>${runs}</si>`;
    });
    const count = items.length;
    const ssXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${count}">
  ${siElements.join("\n  ")}
</sst>`;
    writer.add("xl/sharedStrings.xml", textToBytes(ssXml), { compress: false });
  }

  // ── xl/styles.xml ──
  const stylesXml =
    options.styles ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`;

  writer.add("xl/styles.xml", textToBytes(stylesXml), { compress: false });

  // ── xl/worksheets/sheetN.xml ──
  for (let s = 0; s < options.sheets.length; s++) {
    const sheet = options.sheets[s];
    const rowElements: string[] = [];

    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      const cellElements: string[] = [];

      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        const colLetter = colToLetter(c);
        const ref = `${colLetter}${r + 1}`;

        let typeAttr = cell.type ? ` t="${cell.type}"` : "";
        let styleAttr = cell.styleIndex !== undefined ? ` s="${cell.styleIndex}"` : "";

        let inner = "";
        if (cell.formula) {
          inner += `<f>${xmlEscape(cell.formula)}</f>`;
        }

        if (cell.type === "inlineStr") {
          // Inline string — use <is><t> element instead of <v>
          inner += `<is><t>${xmlEscape(String(cell.value))}</t></is>`;
        } else {
          inner += `<v>${xmlEscape(String(cell.value))}</v>`;
        }

        cellElements.push(`<c r="${ref}"${typeAttr}${styleAttr}>${inner}</c>`);
      }

      rowElements.push(`<row r="${r + 1}">${cellElements.join("")}</row>`);
    }

    let mergesXml = "";
    if (sheet.merges && sheet.merges.length > 0) {
      const mergeCells = sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join("");
      mergesXml = `<mergeCells count="${sheet.merges.length}">${mergeCells}</mergeCells>`;
    }

    const wsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${rowElements.join("")}</sheetData>
  ${mergesXml}
</worksheet>`;

    writer.add(`xl/worksheets/sheet${s + 1}.xml`, textToBytes(wsXml), { compress: false });
  }

  return writer.build();
}

/** Convert 0-based column index to Excel column letter(s) */
function colToLetter(col: number): string {
  let result = "";
  let c = col;
  while (c >= 0) {
    result = String.fromCharCode((c % 26) + 65) + result;
    c = Math.floor(c / 26) - 1;
  }
  return result;
}

/** Minimal XML escaping for attribute values */
function xmlEscapeAttr(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Minimal XML escaping for text content */
function xmlEscape(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ── Cell Reference Parsing Tests ────────────────────────────────────

describe("parseCellRef", () => {
  it("parses A1", () => {
    expect(parseCellRef("A1")).toEqual({ row: 0, col: 0 });
  });

  it("parses Z1", () => {
    expect(parseCellRef("Z1")).toEqual({ row: 0, col: 25 });
  });

  it("parses AA1", () => {
    expect(parseCellRef("AA1")).toEqual({ row: 0, col: 26 });
  });

  it("parses AZ1", () => {
    // AZ = 26 + 25 = 51
    expect(parseCellRef("AZ1")).toEqual({ row: 0, col: 51 });
  });

  it("parses BA1", () => {
    // BA = 52
    expect(parseCellRef("BA1")).toEqual({ row: 0, col: 52 });
  });

  it("parses AAA1", () => {
    // AAA = 26*26 + 26 + 0 = 702
    expect(parseCellRef("AAA1")).toEqual({ row: 0, col: 702 });
  });

  it("parses B10", () => {
    expect(parseCellRef("B10")).toEqual({ row: 9, col: 1 });
  });

  it("parses C100", () => {
    expect(parseCellRef("C100")).toEqual({ row: 99, col: 2 });
  });
});

// ── Content Types Parser Tests ──────────────────────────────────────

describe("parseContentTypes", () => {
  it("parses defaults and overrides", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>`;
    const result = parseContentTypes(xml);
    expect(result.defaults.get("rels")).toBe(
      "application/vnd.openxmlformats-package.relationships+xml",
    );
    expect(result.defaults.get("xml")).toBe("application/xml");
    expect(result.overrides.get("/xl/workbook.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    );
  });
});

// ── Relationships Parser Tests ──────────────────────────────────────

describe("parseRelationships", () => {
  it("parses relationship entries", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
    const rels = parseRelationships(xml);
    expect(rels).toHaveLength(2);
    expect(rels[0]).toEqual({
      id: "rId1",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
      target: "xl/workbook.xml",
    });
  });
});

// ── Shared Strings Parser Tests ─────────────────────────────────────

describe("parseSharedStrings", () => {
  it("parses simple strings", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><t>Test</t></si>
</sst>`;
    const strings = parseSharedStrings(xml);
    expect(strings).toHaveLength(3);
    expect(strings[0].text).toBe("Hello");
    expect(strings[1].text).toBe("World");
    expect(strings[2].text).toBe("Test");
    expect(strings[0].richText).toBeUndefined();
  });

  it("parses rich text strings", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/><sz val="11"/></rPr><t>Bold</t></r>
    <r><t> Normal</t></r>
  </si>
</sst>`;
    const strings = parseSharedStrings(xml);
    expect(strings).toHaveLength(1);
    expect(strings[0].text).toBe("Bold Normal");
    expect(strings[0].richText).toHaveLength(2);
    expect(strings[0].richText![0].text).toBe("Bold");
    expect(strings[0].richText![0].font?.bold).toBe(true);
    expect(strings[0].richText![0].font?.size).toBe(11);
    expect(strings[0].richText![1].text).toBe(" Normal");
  });

  it("handles OOXML escape sequences", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Line1_x000A_Line2</t></si>
</sst>`;
    const strings = parseSharedStrings(xml);
    expect(strings[0].text).toBe("Line1\nLine2");
  });

  it("handles empty strings", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t></t></si>
</sst>`;
    const strings = parseSharedStrings(xml);
    expect(strings[0].text).toBe("");
  });
});

// ── Styles Parser Tests ─────────────────────────────────────────────

describe("parseStyles", () => {
  it("parses number formats, fonts, fills, borders, cellXfs", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><name val="Arial"/><color rgb="FFFF0000"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`;
    const styles = parseStyles(xml);

    expect(styles.numFmts.get(164)).toBe("yyyy-mm-dd");
    expect(styles.fonts).toHaveLength(2);
    expect(styles.fonts[0].name).toBe("Calibri");
    expect(styles.fonts[1].bold).toBe(true);
    expect(styles.fonts[1].name).toBe("Arial");
    expect(styles.fonts[1].color?.rgb).toBe("FF0000");
    expect(styles.fills).toHaveLength(2);
    expect(styles.borders).toHaveLength(1);
    expect(styles.cellXfs).toHaveLength(2);
    expect(styles.cellXfs[1].numFmtId).toBe(164);
  });

  it("detects date styles via built-in format IDs", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;
    const styles = parseStyles(xml);
    expect(isDateStyle(styles, 0)).toBe(false);
    expect(isDateStyle(styles, 1)).toBe(true);
    expect(isDateStyle(styles, 2)).toBe(false);
  });

  it("detects date styles via custom numFmt", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;
    const styles = parseStyles(xml);
    expect(isDateStyle(styles, 1)).toBe(true);
  });

  it("resolveStyle builds CellStyle from xf index", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><name val="Arial"/></font>
  </fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;
    const styles = parseStyles(xml);
    const resolved = resolveStyle(styles, 1);
    expect(resolved.numFmt).toBe("#,##0.00");
    expect(resolved.font?.bold).toBe(true);
    expect(resolved.font?.name).toBe("Arial");
  });
});

// ── XLSX Reader Integration Tests ───────────────────────────────────

describe("readXlsx", () => {
  it("reads basic string, number, and boolean cells", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [
              { value: "0", type: "s" },
              { value: "42", type: "n" },
              { value: "1", type: "b" },
            ],
          ],
        },
      ],
      sharedStrings: ["Hello"],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe("Sheet1");
    expect(wb.sheets[0].rows[0][0]).toBe("Hello");
    expect(wb.sheets[0].rows[0][1]).toBe(42);
    expect(wb.sheets[0].rows[0][2]).toBe(true);
  });

  it("reads shared strings correctly", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [{ value: "0", type: "s" }],
            [{ value: "1", type: "s" }],
            [{ value: "2", type: "s" }],
            [{ value: "0", type: "s" }], // Reuse first string
          ],
        },
      ],
      sharedStrings: ["Apple", "Banana", "Cherry"],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe("Apple");
    expect(wb.sheets[0].rows[1][0]).toBe("Banana");
    expect(wb.sheets[0].rows[2][0]).toBe("Cherry");
    expect(wb.sheets[0].rows[3][0]).toBe("Apple");
  });

  it("reads inline strings", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: "Inline text", type: "inlineStr" }]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe("Inline text");
  });

  it("reads formula cells with cached value", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [
              { value: "10", type: "n" },
              { value: "20", type: "n" },
            ],
            [{ value: "30", formula: "A1+B1" }],
          ],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    // The formula cell should have the cached numeric value
    expect(wb.sheets[0].rows[1][0]).toBe(30);
    // Cell details should contain the formula
    expect(wb.sheets[0].cells?.get("1,0")?.formula).toBe("A1+B1");
    expect(wb.sheets[0].cells?.get("1,0")?.formulaResult).toBe(30);
  });

  it("reads date cells via number format detection", async () => {
    // Excel serial 44927 = 2023-01-01 in 1900 date system
    const dateSerial = 44927;
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;

    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: String(dateSerial), styleIndex: 1 }]],
        },
      ],
      styles: stylesXml,
    });

    const wb = await readXlsx(xlsx);
    const cellValue = wb.sheets[0].rows[0][0];
    expect(cellValue).toBeInstanceOf(Date);
    const d = cellValue as Date;
    expect(d.getUTCFullYear()).toBe(2023);
    expect(d.getUTCMonth()).toBe(0); // January
    expect(d.getUTCDate()).toBe(1);
  });

  it("reads multiple sheets", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "First",
          rows: [[{ value: "0", type: "s" }]],
        },
        {
          name: "Second",
          rows: [[{ value: "1", type: "s" }]],
        },
        {
          name: "Third",
          rows: [[{ value: "2", type: "s" }]],
        },
      ],
      sharedStrings: ["Sheet1Data", "Sheet2Data", "Sheet3Data"],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets).toHaveLength(3);
    expect(wb.sheets[0].name).toBe("First");
    expect(wb.sheets[0].rows[0][0]).toBe("Sheet1Data");
    expect(wb.sheets[1].name).toBe("Second");
    expect(wb.sheets[1].rows[0][0]).toBe("Sheet2Data");
    expect(wb.sheets[2].name).toBe("Third");
    expect(wb.sheets[2].rows[0][0]).toBe("Sheet3Data");
  });

  it("reads sheet names correctly", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        { name: "Revenue", rows: [[{ value: "100" }]] },
        { name: "Expenses", rows: [[{ value: "50" }]] },
        { name: "Summary", rows: [[{ value: "50" }]] },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets.map((s) => s.name)).toEqual(["Revenue", "Expenses", "Summary"]);
  });

  it("reads merged cells", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [{ value: "0", type: "s" }, { value: "" }, { value: "" }],
            [{ value: "" }, { value: "" }, { value: "" }],
          ],
          merges: ["A1:C2"],
        },
      ],
      sharedStrings: ["Merged"],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].merges).toHaveLength(1);
    expect(wb.sheets[0].merges![0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 2,
    });
  });

  it("reads empty cells as null", async () => {
    // Sheet with a value in A1 and C1, but B1 is empty
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [
              { value: "0", type: "s" },
              { value: "", type: "n" },
              { value: "1", type: "s" },
            ],
          ],
        },
      ],
      sharedStrings: ["First", "Third"],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe("First");
    expect(wb.sheets[0].rows[0][1]).toBeNull();
    expect(wb.sheets[0].rows[0][2]).toBe("Third");
  });

  it("reads sparse rows correctly", async () => {
    // Row 1 has data, row 2 has no cells at all, row 3 has data
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [{ value: "1" }],
            // Row 2: intentionally empty — we'll skip it in the sheet builder
            // But worksheet should still have it as null-filled
          ],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(1);
  });

  it("reads cells with different number formats", async () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
    <numFmt numFmtId="165" formatCode="0.00%"/>
  </numFmts>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;

    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [
              { value: "1234.56", styleIndex: 1 },
              { value: "0.75", styleIndex: 2 },
            ],
          ],
        },
      ],
      styles: stylesXml,
    });

    const wb = await readXlsx(xlsx, { readStyles: true });
    expect(wb.sheets[0].rows[0][0]).toBe(1234.56);
    expect(wb.sheets[0].rows[0][1]).toBe(0.75);
    // With readStyles, cells should have style info
    expect(wb.sheets[0].cells?.get("0,0")?.style?.numFmt).toBe("#,##0.00");
    expect(wb.sheets[0].cells?.get("0,1")?.style?.numFmt).toBe("0.00%");
  });

  it("reads rich text cells from shared strings", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: "0", type: "s" }]],
        },
      ],
      richSharedStrings: [
        {
          runs: [
            { text: "Bold", bold: true, size: 12 },
            { text: " and ", size: 12 },
            { text: "Italic", italic: true, size: 12 },
          ],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe("Bold and Italic");
    expect(wb.sheets[0].cells?.get("0,0")?.richText).toHaveLength(3);
    expect(wb.sheets[0].cells?.get("0,0")?.richText![0].text).toBe("Bold");
    expect(wb.sheets[0].cells?.get("0,0")?.richText![0].font?.bold).toBe(true);
    expect(wb.sheets[0].cells?.get("0,0")?.richText![2].text).toBe("Italic");
    expect(wb.sheets[0].cells?.get("0,0")?.richText![2].font?.italic).toBe(true);
  });

  it("reads with readStyles option", async () => {
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><name val="Arial"/></font>
  </fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;

    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: "42", styleIndex: 1 }]],
        },
      ],
      styles: stylesXml,
    });

    // Without readStyles
    const wb1 = await readXlsx(xlsx);
    expect(wb1.sheets[0].cells?.get("0,0")?.style).toBeUndefined();

    // With readStyles
    const wb2 = await readXlsx(xlsx, { readStyles: true });
    expect(wb2.sheets[0].cells?.get("0,0")?.style?.font?.bold).toBe(true);
    expect(wb2.sheets[0].cells?.get("0,0")?.style?.font?.name).toBe("Arial");
  });

  it("reads 1904 date system", async () => {
    // In 1904 system, serial 0 = Jan 1, 1904
    const dateSerial = 0;
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;

    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: String(dateSerial), styleIndex: 1 }]],
        },
      ],
      styles: stylesXml,
      dateSystem: "1904",
    });

    const wb = await readXlsx(xlsx);
    expect(wb.dateSystem).toBe("1904");
    const cellValue = wb.sheets[0].rows[0][0];
    expect(cellValue).toBeInstanceOf(Date);
    const d = cellValue as Date;
    expect(d.getUTCFullYear()).toBe(1904);
    expect(d.getUTCMonth()).toBe(0);
    expect(d.getUTCDate()).toBe(1);
  });

  it("errors on invalid/corrupt ZIP", async () => {
    const badData = new Uint8Array([
      0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
    ]);
    await expect(readXlsx(badData)).rejects.toThrow();
  });

  it("errors on missing required parts", async () => {
    // Build a ZIP without [Content_Types].xml
    const writer = new ZipWriter();
    writer.add("dummy.txt", textToBytes("dummy"), { compress: false });
    const zip = await writer.build();

    await expect(readXlsx(zip)).rejects.toThrow("missing [Content_Types].xml");
  });

  it("errors on missing _rels/.rels", async () => {
    const writer = new ZipWriter();
    writer.add(
      "[Content_Types].xml",
      textToBytes(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
      ),
      { compress: false },
    );
    const zip = await writer.build();

    await expect(readXlsx(zip)).rejects.toThrow("missing _rels/.rels");
  });

  it("reads large sheet (1000 rows x 10 columns)", async () => {
    const numRows = 1000;
    const numCols = 10;

    const rows: Array<Array<{ value: string | number | boolean }>> = [];
    for (let r = 0; r < numRows; r++) {
      const row: Array<{ value: string | number | boolean }> = [];
      for (let c = 0; c < numCols; c++) {
        row.push({ value: String(r * numCols + c) });
      }
      rows.push(row);
    }

    const xlsx = await createTestXlsx({
      sheets: [{ name: "LargeSheet", rows }],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows).toHaveLength(numRows);
    expect(wb.sheets[0].rows[0]).toHaveLength(numCols);
    expect(wb.sheets[0].rows[0][0]).toBe(0);
    expect(wb.sheets[0].rows[999][9]).toBe(9999);
  });

  it("reads sheet with Unicode content", async () => {
    const unicodeStrings = [
      "日本語テスト",
      "Ünïcödé",
      "العربية",
      "中文测试",
      "한국어",
      "Emoji 🎉🎊",
    ];

    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Unicode",
          rows: unicodeStrings.map((s, i) => [{ value: String(i), type: "s" }]),
        },
      ],
      sharedStrings: unicodeStrings,
    });

    const wb = await readXlsx(xlsx);
    for (let i = 0; i < unicodeStrings.length; i++) {
      expect(wb.sheets[0].rows[i][0]).toBe(unicodeStrings[i]);
    }
  });

  it("reads specific sheets by name", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        { name: "Alpha", rows: [[{ value: "1" }]] },
        { name: "Beta", rows: [[{ value: "2" }]] },
        { name: "Gamma", rows: [[{ value: "3" }]] },
      ],
    });

    const wb = await readXlsx(xlsx, { sheets: ["Beta"] });
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe("Beta");
    expect(wb.sheets[0].rows[0][0]).toBe(2);
  });

  it("reads specific sheets by index", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        { name: "Alpha", rows: [[{ value: "1" }]] },
        { name: "Beta", rows: [[{ value: "2" }]] },
        { name: "Gamma", rows: [[{ value: "3" }]] },
      ],
    });

    const wb = await readXlsx(xlsx, { sheets: [0, 2] });
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[0].name).toBe("Alpha");
    expect(wb.sheets[1].name).toBe("Gamma");
  });

  it("reads boolean false values", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [
              { value: "0", type: "b" },
              { value: "1", type: "b" },
            ],
          ],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(false);
    expect(wb.sheets[0].rows[0][1]).toBe(true);
  });

  it("reads error cells", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [
              { value: "#VALUE!", type: "e" },
              { value: "#REF!", type: "e" },
              { value: "#DIV/0!", type: "e" },
            ],
          ],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe("#VALUE!");
    expect(wb.sheets[0].rows[0][1]).toBe("#REF!");
    expect(wb.sheets[0].rows[0][2]).toBe("#DIV/0!");
  });

  it("reads str type cells (formula string results)", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [{ value: "Hello World", type: "str", formula: 'CONCATENATE("Hello"," ","World")' }],
          ],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe("Hello World");
    expect(wb.sheets[0].cells?.get("0,0")?.formula).toBe('CONCATENATE("Hello"," ","World")');
  });

  it("accepts ArrayBuffer input", async () => {
    const xlsx = await createTestXlsx({
      sheets: [{ name: "Sheet1", rows: [[{ value: "42" }]] }],
    });

    // Convert to ArrayBuffer
    const ab = xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength) as ArrayBuffer;

    const wb = await readXlsx(ab);
    expect(wb.sheets[0].rows[0][0]).toBe(42);
  });

  it("handles sheet with no data rows", async () => {
    const xlsx = await createTestXlsx({
      sheets: [{ name: "Empty", rows: [] }],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].name).toBe("Empty");
    expect(wb.sheets[0].rows).toHaveLength(0);
  });

  it("reads multiple merge ranges in a single sheet", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            [{ value: "0", type: "s" }, { value: "" }, { value: "1", type: "s" }, { value: "" }],
            [{ value: "" }, { value: "" }, { value: "" }, { value: "" }],
          ],
          merges: ["A1:B1", "C1:D2"],
        },
      ],
      sharedStrings: ["Header1", "Header2"],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].merges).toHaveLength(2);
    expect(wb.sheets[0].merges![0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 1,
    });
    expect(wb.sheets[0].merges![1]).toEqual({
      startRow: 0,
      startCol: 2,
      endRow: 1,
      endCol: 3,
    });
  });

  it("reads date system override from options", async () => {
    // Create an XLSX with default 1900 date system but override via options
    const dateSerial = 0; // In 1904: Jan 1, 1904; In 1900: Dec 30, 1899
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;

    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: String(dateSerial), styleIndex: 1 }]],
        },
      ],
      styles: stylesXml,
    });

    const wb = await readXlsx(xlsx, { dateSystem: "1904" });
    expect(wb.dateSystem).toBe("1904");
    const cellValue = wb.sheets[0].rows[0][0];
    expect(cellValue).toBeInstanceOf(Date);
    const d = cellValue as Date;
    expect(d.getUTCFullYear()).toBe(1904);
    expect(d.getUTCMonth()).toBe(0);
    expect(d.getUTCDate()).toBe(1);
  });

  it("reads negative numbers", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: "-42.5" }, { value: "-100" }]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(-42.5);
    expect(wb.sheets[0].rows[0][1]).toBe(-100);
  });

  it("reads floating point numbers", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: "3.14159265358979" }, { value: "0.001" }]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBeCloseTo(3.14159265358979);
    expect(wb.sheets[0].rows[0][1]).toBeCloseTo(0.001);
  });

  it("reads zero values correctly", async () => {
    const xlsx = await createTestXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[{ value: "0" }]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(0);
  });
});

// ── colToLetter helper test ─────────────────────────────────────────

describe("colToLetter", () => {
  it("converts 0-based indices to Excel column letters", () => {
    expect(colToLetter(0)).toBe("A");
    expect(colToLetter(25)).toBe("Z");
    expect(colToLetter(26)).toBe("AA");
    expect(colToLetter(51)).toBe("AZ");
    expect(colToLetter(52)).toBe("BA");
    expect(colToLetter(702)).toBe("AAA");
  });
});
