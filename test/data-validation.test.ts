import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { createStylesCollector } from "../src/xlsx/styles-writer";
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer";
import { readXlsx } from "../src/xlsx/reader";
import type { WriteSheet, DataValidation } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

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

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("");
}

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector();
  const ss = createSharedStrings();
  const result = writeWorksheetXml(sheet, styles, ss);
  return result.xml;
}

function parseWorksheetXml(xml: string) {
  return parseXml(xml);
}

// ── Writing Tests ────────────────────────────────────────────────────

describe("data validation — writing", () => {
  it("writes list validation with explicit values", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Status"]],
      dataValidations: [
        {
          type: "list",
          values: ["Active", "Inactive", "Draft"],
          range: "A2:A100",
          allowBlank: true,
          showInputMessage: true,
          showErrorMessage: true,
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    expect(dvs).toBeDefined();
    expect(dvs.attrs["count"]).toBe("1");

    const dv = findChild(dvs, "dataValidation");
    expect(dv).toBeDefined();
    expect(dv.attrs["type"]).toBe("list");
    expect(dv.attrs["sqref"]).toBe("A2:A100");
    expect(dv.attrs["allowBlank"]).toBe("1");
    expect(dv.attrs["showInputMessage"]).toBe("1");
    expect(dv.attrs["showErrorMessage"]).toBe("1");

    const formula1 = findChild(dv, "formula1");
    expect(formula1).toBeDefined();
    const f1Text = formula1.text || getElementText(formula1);
    expect(f1Text).toBe('"Active,Inactive,Draft"');
  });

  it("writes list validation with formula reference", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Category"]],
      dataValidations: [
        {
          type: "list",
          formula1: "Sheet2!$A$1:$A$10",
          range: "B2:B50",
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    const dv = findChild(dvs, "dataValidation");
    expect(dv.attrs["type"]).toBe("list");
    expect(dv.attrs["sqref"]).toBe("B2:B50");

    const formula1 = findChild(dv, "formula1");
    const f1Text = formula1.text || getElementText(formula1);
    expect(f1Text).toBe("Sheet2!$A$1:$A$10");
  });

  it("writes whole number between validation", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Quantity"]],
      dataValidations: [
        {
          type: "whole",
          operator: "between",
          formula1: "0",
          formula2: "1000",
          range: "B2:B100",
          allowBlank: true,
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    const dv = findChild(dvs, "dataValidation");
    expect(dv.attrs["type"]).toBe("whole");
    expect(dv.attrs["operator"]).toBe("between");
    expect(dv.attrs["sqref"]).toBe("B2:B100");
    expect(dv.attrs["allowBlank"]).toBe("1");

    const formula1 = findChild(dv, "formula1");
    expect(formula1.text || getElementText(formula1)).toBe("0");

    const formula2 = findChild(dv, "formula2");
    expect(formula2.text || getElementText(formula2)).toBe("1000");
  });

  it("writes decimal greaterThan validation", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Price"]],
      dataValidations: [
        {
          type: "decimal",
          operator: "greaterThan",
          formula1: "0",
          range: "C2:C100",
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    const dv = findChild(dvs, "dataValidation");
    expect(dv.attrs["type"]).toBe("decimal");
    expect(dv.attrs["operator"]).toBe("greaterThan");

    const formula1 = findChild(dv, "formula1");
    expect(formula1.text || getElementText(formula1)).toBe("0");

    // No formula2 for greaterThan
    const formula2 = findChild(dv, "formula2");
    expect(formula2).toBeUndefined();
  });

  it("writes text length validation", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Code"]],
      dataValidations: [
        {
          type: "textLength",
          operator: "lessThanOrEqual",
          formula1: "50",
          range: "D2:D100",
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    const dv = findChild(dvs, "dataValidation");
    expect(dv.attrs["type"]).toBe("textLength");
    expect(dv.attrs["operator"]).toBe("lessThanOrEqual");

    const formula1 = findChild(dv, "formula1");
    expect(formula1.text || getElementText(formula1)).toBe("50");
  });

  it("writes custom formula validation", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Email"]],
      dataValidations: [
        {
          type: "custom",
          formula1: 'ISNUMBER(FIND("@",A2))',
          range: "A2:A100",
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    const dv = findChild(dvs, "dataValidation");
    expect(dv.attrs["type"]).toBe("custom");

    const formula1 = findChild(dv, "formula1");
    const f1Text = formula1.text || getElementText(formula1);
    expect(f1Text).toBe('ISNUMBER(FIND("@",A2))');
  });

  it("writes multiple validations on same sheet", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Status", "Quantity", "Price"]],
      dataValidations: [
        {
          type: "list",
          values: ["Active", "Inactive"],
          range: "A2:A100",
        },
        {
          type: "whole",
          operator: "between",
          formula1: "1",
          formula2: "9999",
          range: "B2:B100",
        },
        {
          type: "decimal",
          operator: "greaterThan",
          formula1: "0",
          range: "C2:C100",
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    expect(dvs.attrs["count"]).toBe("3");

    const dvList = findChildren(dvs, "dataValidation");
    expect(dvList.length).toBe(3);

    // Check each validation by sqref
    const byRange = new Map(dvList.map((d: any) => [d.attrs["sqref"], d]));
    expect(byRange.get("A2:A100").attrs["type"]).toBe("list");
    expect(byRange.get("B2:B100").attrs["type"]).toBe("whole");
    expect(byRange.get("C2:C100").attrs["type"]).toBe("decimal");
  });

  it("writes validation with input/error messages", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      dataValidations: [
        {
          type: "whole",
          operator: "between",
          formula1: "1",
          formula2: "100",
          range: "A2:A50",
          showInputMessage: true,
          showErrorMessage: true,
          inputTitle: "Enter a number",
          inputMessage: "Please enter a number between 1 and 100",
          errorTitle: "Invalid input",
          errorMessage: "The value must be between 1 and 100",
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    const dv = findChild(dvs, "dataValidation");

    expect(dv.attrs["showInputMessage"]).toBe("1");
    expect(dv.attrs["showErrorMessage"]).toBe("1");
    expect(dv.attrs["promptTitle"]).toBe("Enter a number");
    expect(dv.attrs["prompt"]).toBe("Please enter a number between 1 and 100");
    expect(dv.attrs["errorTitle"]).toBe("Invalid input");
    expect(dv.attrs["error"]).toBe("The value must be between 1 and 100");
  });

  it("writes validation with error style", () => {
    const errorStyles = ["stop", "warning", "information"] as const;

    for (const errorStyle of errorStyles) {
      const sheet: WriteSheet = {
        name: "Test",
        rows: [["Value"]],
        dataValidations: [
          {
            type: "whole",
            operator: "greaterThan",
            formula1: "0",
            range: "A2:A10",
            showErrorMessage: true,
            errorStyle,
          },
        ],
      };

      const xml = writeXml(sheet);
      const doc = parseWorksheetXml(xml);

      const dvs = findChild(doc, "dataValidations");
      const dv = findChild(dvs, "dataValidation");
      expect(dv.attrs["errorStyle"]).toBe(errorStyle);
    }
  });

  it("writes validation with allowBlank", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      dataValidations: [
        {
          type: "list",
          values: ["A", "B", "C"],
          range: "A2:A10",
          allowBlank: true,
        },
      ],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    const dv = findChild(dvs, "dataValidation");
    expect(dv.attrs["allowBlank"]).toBe("1");
  });

  it("does not emit dataValidations when none are provided", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
    };

    const xml = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const dvs = findChild(doc, "dataValidations");
    expect(dvs).toBeUndefined();
  });
});

// ── Reading Tests ────────────────────────────────────────────────────

describe("data validation — reading (round-trip)", () => {
  it("round-trips list validation with values", async () => {
    const validations: DataValidation[] = [
      {
        type: "list",
        values: ["Active", "Inactive", "Draft"],
        range: "A2:A100",
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
      },
    ];

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Status"]],
          dataValidations: validations,
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets.length).toBe(1);

    const sheet = workbook.sheets[0];
    expect(sheet.dataValidations).toBeDefined();
    expect(sheet.dataValidations!.length).toBe(1);

    const dv = sheet.dataValidations![0];
    expect(dv.type).toBe("list");
    expect(dv.values).toEqual(["Active", "Inactive", "Draft"]);
    expect(dv.range).toBe("A2:A100");
    expect(dv.allowBlank).toBe(true);
    expect(dv.showInputMessage).toBe(true);
    expect(dv.showErrorMessage).toBe(true);
    // values parsed into array, formula1 should be undefined
    expect(dv.formula1).toBeUndefined();
  });

  it("round-trips list validation with formula reference", async () => {
    const validations: DataValidation[] = [
      {
        type: "list",
        formula1: "Sheet2!$A$1:$A$10",
        range: "B2:B50",
      },
    ];

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Category"]],
          dataValidations: validations,
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dv = workbook.sheets[0].dataValidations![0];
    expect(dv.type).toBe("list");
    expect(dv.formula1).toBe("Sheet2!$A$1:$A$10");
    expect(dv.values).toBeUndefined();
    expect(dv.range).toBe("B2:B50");
  });

  it("round-trips whole number between validation", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Qty"]],
          dataValidations: [
            {
              type: "whole",
              operator: "between",
              formula1: "0",
              formula2: "1000",
              range: "B2:B100",
              allowBlank: true,
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dv = workbook.sheets[0].dataValidations![0];
    expect(dv.type).toBe("whole");
    expect(dv.operator).toBe("between");
    expect(dv.formula1).toBe("0");
    expect(dv.formula2).toBe("1000");
    expect(dv.range).toBe("B2:B100");
    expect(dv.allowBlank).toBe(true);
  });

  it("round-trips decimal greaterThan validation", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Price"]],
          dataValidations: [
            {
              type: "decimal",
              operator: "greaterThan",
              formula1: "0",
              range: "C2:C100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dv = workbook.sheets[0].dataValidations![0];
    expect(dv.type).toBe("decimal");
    expect(dv.operator).toBe("greaterThan");
    expect(dv.formula1).toBe("0");
    expect(dv.formula2).toBeUndefined();
  });

  it("round-trips text length validation", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Code"]],
          dataValidations: [
            {
              type: "textLength",
              operator: "lessThanOrEqual",
              formula1: "50",
              range: "D2:D100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dv = workbook.sheets[0].dataValidations![0];
    expect(dv.type).toBe("textLength");
    expect(dv.operator).toBe("lessThanOrEqual");
    expect(dv.formula1).toBe("50");
  });

  it("round-trips custom formula validation", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Email"]],
          dataValidations: [
            {
              type: "custom",
              formula1: 'ISNUMBER(FIND("@",A2))',
              range: "A2:A100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dv = workbook.sheets[0].dataValidations![0];
    expect(dv.type).toBe("custom");
    expect(dv.formula1).toBe('ISNUMBER(FIND("@",A2))');
  });

  it("round-trips multiple validations", async () => {
    const validations: DataValidation[] = [
      {
        type: "list",
        values: ["A", "B", "C"],
        range: "A2:A100",
      },
      {
        type: "whole",
        operator: "between",
        formula1: "1",
        formula2: "999",
        range: "B2:B100",
      },
      {
        type: "decimal",
        operator: "greaterThan",
        formula1: "0",
        range: "C2:C100",
      },
    ];

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Status", "Qty", "Price"]],
          dataValidations: validations,
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dvs = workbook.sheets[0].dataValidations!;
    expect(dvs.length).toBe(3);

    const byRange = new Map(dvs.map((d) => [d.range, d]));
    expect(byRange.get("A2:A100")!.type).toBe("list");
    expect(byRange.get("A2:A100")!.values).toEqual(["A", "B", "C"]);
    expect(byRange.get("B2:B100")!.type).toBe("whole");
    expect(byRange.get("B2:B100")!.operator).toBe("between");
    expect(byRange.get("C2:C100")!.type).toBe("decimal");
  });

  it("round-trips validation with input/error messages", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          dataValidations: [
            {
              type: "whole",
              operator: "between",
              formula1: "1",
              formula2: "100",
              range: "A2:A50",
              showInputMessage: true,
              showErrorMessage: true,
              inputTitle: "Enter a number",
              inputMessage: "Between 1 and 100",
              errorTitle: "Invalid",
              errorMessage: "Must be 1-100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dv = workbook.sheets[0].dataValidations![0];
    expect(dv.showInputMessage).toBe(true);
    expect(dv.showErrorMessage).toBe(true);
    expect(dv.inputTitle).toBe("Enter a number");
    expect(dv.inputMessage).toBe("Between 1 and 100");
    expect(dv.errorTitle).toBe("Invalid");
    expect(dv.errorMessage).toBe("Must be 1-100");
  });

  it("round-trips validation with error styles", async () => {
    const errorStyles = ["stop", "warning", "information"] as const;

    for (const errorStyle of errorStyles) {
      const data = await writeXlsx({
        sheets: [
          {
            name: "Sheet1",
            rows: [["Value"]],
            dataValidations: [
              {
                type: "whole",
                operator: "greaterThan",
                formula1: "0",
                range: "A2:A10",
                showErrorMessage: true,
                errorStyle,
              },
            ],
          },
        ],
      });

      const workbook = await readXlsx(data);
      const dv = workbook.sheets[0].dataValidations![0];
      expect(dv.errorStyle).toBe(errorStyle);
    }
  });

  it("preserves all properties in round-trip", async () => {
    const original: DataValidation = {
      type: "whole",
      operator: "between",
      formula1: "10",
      formula2: "200",
      range: "A2:A100",
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      inputTitle: "Input Title",
      inputMessage: "Input Message",
      errorTitle: "Error Title",
      errorMessage: "Error Message",
      errorStyle: "warning",
    };

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          dataValidations: [original],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const dv = workbook.sheets[0].dataValidations![0];

    expect(dv.type).toBe(original.type);
    expect(dv.operator).toBe(original.operator);
    expect(dv.formula1).toBe(original.formula1);
    expect(dv.formula2).toBe(original.formula2);
    expect(dv.range).toBe(original.range);
    expect(dv.allowBlank).toBe(original.allowBlank);
    expect(dv.showInputMessage).toBe(original.showInputMessage);
    expect(dv.showErrorMessage).toBe(original.showErrorMessage);
    expect(dv.inputTitle).toBe(original.inputTitle);
    expect(dv.inputMessage).toBe(original.inputMessage);
    expect(dv.errorTitle).toBe(original.errorTitle);
    expect(dv.errorMessage).toBe(original.errorMessage);
    expect(dv.errorStyle).toBe(original.errorStyle);
  });
});

// ── Integration Tests ────────────────────────────────────────────────

describe("data validation — integration (ZIP verification)", () => {
  it("writes correct dataValidations XML in XLSX file", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Status", "Qty"]],
          dataValidations: [
            {
              type: "list",
              values: ["Active", "Inactive", "Draft"],
              range: "A2:A100",
              allowBlank: true,
              showInputMessage: true,
              showErrorMessage: true,
            },
            {
              type: "whole",
              operator: "between",
              formula1: "0",
              formula2: "1000",
              range: "B2:B100",
              allowBlank: true,
            },
          ],
        },
      ],
    });

    // Extract the worksheet XML from the ZIP
    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml");
    const doc = parseXml(sheetXml);

    // Find dataValidations element
    const dvs = findChild(doc, "dataValidations");
    expect(dvs).toBeDefined();
    expect(dvs.attrs["count"]).toBe("2");

    const dvList = findChildren(dvs, "dataValidation");
    expect(dvList.length).toBe(2);

    // Verify list validation
    const listDv = dvList.find((d: any) => d.attrs["type"] === "list");
    expect(listDv).toBeDefined();
    expect(listDv.attrs["sqref"]).toBe("A2:A100");
    expect(listDv.attrs["allowBlank"]).toBe("1");
    expect(listDv.attrs["showInputMessage"]).toBe("1");
    expect(listDv.attrs["showErrorMessage"]).toBe("1");

    const listF1 = findChild(listDv, "formula1");
    expect(listF1.text || getElementText(listF1)).toBe('"Active,Inactive,Draft"');

    // Verify whole number validation
    const wholeDv = dvList.find((d: any) => d.attrs["type"] === "whole");
    expect(wholeDv).toBeDefined();
    expect(wholeDv.attrs["sqref"]).toBe("B2:B100");
    expect(wholeDv.attrs["operator"]).toBe("between");

    const wholeF1 = findChild(wholeDv, "formula1");
    expect(wholeF1.text || getElementText(wholeF1)).toBe("0");

    const wholeF2 = findChild(wholeDv, "formula2");
    expect(wholeF2.text || getElementText(wholeF2)).toBe("1000");
  });

  it("places dataValidations after autoFilter in XML order", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B"],
            [1, 2],
          ],
          autoFilter: { range: "A1:B2" },
          dataValidations: [
            {
              type: "list",
              values: ["X", "Y"],
              range: "A3:A10",
            },
          ],
        },
      ],
    });

    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml");

    // Verify XML ordering: autoFilter appears before dataValidations
    const autoFilterPos = sheetXml.indexOf("<autoFilter");
    const dataValidationsPos = sheetXml.indexOf("<dataValidations");
    expect(autoFilterPos).toBeGreaterThan(-1);
    expect(dataValidationsPos).toBeGreaterThan(-1);
    expect(autoFilterPos).toBeLessThan(dataValidationsPos);
  });

  it("does not include dataValidations element when none exist", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Hello", "World"]],
        },
      ],
    });

    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml");
    expect(sheetXml).not.toContain("<dataValidations");
    expect(sheetXml).not.toContain("<dataValidation");
  });
});
