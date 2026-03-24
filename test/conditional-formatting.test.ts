import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { createStylesCollector } from "../src/xlsx/styles-writer";
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer";
import { readXlsx } from "../src/xlsx/reader";
import type { WriteSheet, ConditionalRule } from "../src/_types";

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

function writeXml(sheet: WriteSheet): { xml: string; stylesXml: string } {
  const styles = createStylesCollector();
  const ss = createSharedStrings();
  const result = writeWorksheetXml(sheet, styles, ss);
  return { xml: result.xml, stylesXml: styles.toXml() };
}

function parseWorksheetXml(xml: string) {
  return parseXml(xml);
}

// ── Writing Tests ────────────────────────────────────────────────────

describe("conditional formatting — writing", () => {
  it("writes cellIs rule with operator and formula", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "cellIs",
          priority: 1,
          operator: "greaterThan",
          formula: "100",
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    expect(cf).toBeDefined();
    expect(cf.attrs["sqref"]).toBe("A1:A100");

    const cfRule = findChild(cf, "cfRule");
    expect(cfRule).toBeDefined();
    expect(cfRule.attrs["type"]).toBe("cellIs");
    expect(cfRule.attrs["priority"]).toBe("1");
    expect(cfRule.attrs["operator"]).toBe("greaterThan");

    const formula = findChild(cfRule, "formula");
    expect(formula).toBeDefined();
    const fText = formula.text || getElementText(formula);
    expect(fText).toBe("100");
  });

  it("writes cellIs rule with style → dxfId in cfRule and dxf in styles.xml", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "cellIs",
          priority: 1,
          operator: "greaterThan",
          formula: "100",
          range: "A1:A100",
          style: {
            font: { bold: true, color: { rgb: "9C0006" } },
            fill: {
              type: "pattern",
              pattern: "solid",
              bgColor: { rgb: "FFC7CE" },
            },
          },
        },
      ],
    };

    const { xml, stylesXml } = writeXml(sheet);

    // Verify cfRule has dxfId
    const doc = parseWorksheetXml(xml);
    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["dxfId"]).toBe("0");

    // Verify dxf in styles.xml
    const stylesDoc = parseXml(stylesXml);
    const dxfs = findChild(stylesDoc, "dxfs");
    expect(dxfs).toBeDefined();
    expect(dxfs.attrs["count"]).toBe("1");

    const dxf = findChild(dxfs, "dxf");
    expect(dxf).toBeDefined();

    const font = findChild(dxf, "font");
    expect(font).toBeDefined();
    const bold = findChild(font, "b");
    expect(bold).toBeDefined();

    const fill = findChild(dxf, "fill");
    expect(fill).toBeDefined();
  });

  it("writes expression rule with custom formula", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "expression",
          priority: 1,
          formula: "MOD(ROW(),2)=0",
          range: "A1:D100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("expression");

    const formula = findChild(cfRule, "formula");
    const fText = formula.text || getElementText(formula);
    expect(fText).toBe("MOD(ROW(),2)=0");
  });

  it("writes 2-color colorScale", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "colorScale",
          priority: 1,
          range: "A1:A100",
          colorScale: {
            cfvo: [{ type: "min" }, { type: "max" }],
            colors: ["FF63BE7B", "FFF8696B"],
          },
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("colorScale");

    const colorScale = findChild(cfRule, "colorScale");
    expect(colorScale).toBeDefined();

    const cfvos = findChildren(colorScale, "cfvo");
    expect(cfvos.length).toBe(2);
    expect(cfvos[0].attrs["type"]).toBe("min");
    expect(cfvos[1].attrs["type"]).toBe("max");

    const colors = findChildren(colorScale, "color");
    expect(colors.length).toBe(2);
    expect(colors[0].attrs["rgb"]).toBe("FF63BE7B");
    expect(colors[1].attrs["rgb"]).toBe("FFF8696B");
  });

  it("writes 3-color colorScale", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "colorScale",
          priority: 1,
          range: "B1:B100",
          colorScale: {
            cfvo: [{ type: "min" }, { type: "percentile", value: "50" }, { type: "max" }],
            colors: ["FFF8696B", "FFFFEB84", "FF63BE7B"],
          },
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    const colorScale = findChild(cfRule, "colorScale");

    const cfvos = findChildren(colorScale, "cfvo");
    expect(cfvos.length).toBe(3);
    expect(cfvos[1].attrs["type"]).toBe("percentile");
    expect(cfvos[1].attrs["val"]).toBe("50");

    const colors = findChildren(colorScale, "color");
    expect(colors.length).toBe(3);
  });

  it("writes dataBar", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "dataBar",
          priority: 1,
          range: "C1:C100",
          dataBar: {
            cfvo: [{ type: "min" }, { type: "max" }],
            color: "FF638EC6",
          },
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("dataBar");

    const dataBar = findChild(cfRule, "dataBar");
    expect(dataBar).toBeDefined();

    const cfvos = findChildren(dataBar, "cfvo");
    expect(cfvos.length).toBe(2);

    const color = findChild(dataBar, "color");
    expect(color.attrs["rgb"]).toBe("FF638EC6");
  });

  it("writes iconSet", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "iconSet",
          priority: 1,
          range: "D1:D100",
          iconSet: {
            iconSet: "3Arrows",
            cfvo: [
              { type: "percent", value: "0" },
              { type: "percent", value: "33" },
              { type: "percent", value: "67" },
            ],
          },
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("iconSet");

    const iconSet = findChild(cfRule, "iconSet");
    expect(iconSet).toBeDefined();
    expect(iconSet.attrs["iconSet"]).toBe("3Arrows");

    const cfvos = findChildren(iconSet, "cfvo");
    expect(cfvos.length).toBe(3);
    expect(cfvos[0].attrs["type"]).toBe("percent");
    expect(cfvos[0].attrs["val"]).toBe("0");
    expect(cfvos[1].attrs["val"]).toBe("33");
    expect(cfvos[2].attrs["val"]).toBe("67");
  });

  it("writes containsText rule", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "containsText",
          priority: 1,
          text: "error",
          formula: 'NOT(ISERROR(SEARCH("error",A1)))',
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("containsText");
    expect(cfRule.attrs["text"]).toBe("error");

    const formula = findChild(cfRule, "formula");
    const fText = formula.text || getElementText(formula);
    expect(fText).toContain("SEARCH");
  });

  it("writes beginsWith rule", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "beginsWith",
          priority: 1,
          text: "OK",
          formula: 'LEFT(A1,LEN("OK"))="OK"',
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("beginsWith");
    expect(cfRule.attrs["text"]).toBe("OK");
  });

  it("writes endsWith rule", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "endsWith",
          priority: 1,
          text: ".com",
          formula: 'RIGHT(A1,LEN(".com"))=".com"',
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("endsWith");
    expect(cfRule.attrs["text"]).toBe(".com");
  });

  it("writes duplicateValues rule", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "duplicateValues",
          priority: 1,
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("duplicateValues");
  });

  it("writes uniqueValues rule", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "uniqueValues",
          priority: 1,
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("uniqueValues");
  });

  it("writes top10 rule", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "top10",
          priority: 1,
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("top10");
  });

  it("writes aboveAverage rule", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "aboveAverage",
          priority: 1,
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["type"]).toBe("aboveAverage");
  });

  it("writes multiple rules on same range (priorities)", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "cellIs",
          priority: 1,
          operator: "greaterThan",
          formula: "100",
          range: "A1:A100",
        },
        {
          type: "cellIs",
          priority: 2,
          operator: "lessThan",
          formula: "0",
          range: "A1:A100",
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    // Same range → single conditionalFormatting element with 2 cfRules
    const cfs = findChildren(doc, "conditionalFormatting");
    expect(cfs.length).toBe(1);
    expect(cfs[0].attrs["sqref"]).toBe("A1:A100");

    const cfRules = findChildren(cfs[0], "cfRule");
    expect(cfRules.length).toBe(2);
    expect(cfRules[0].attrs["priority"]).toBe("1");
    expect(cfRules[0].attrs["operator"]).toBe("greaterThan");
    expect(cfRules[1].attrs["priority"]).toBe("2");
    expect(cfRules[1].attrs["operator"]).toBe("lessThan");
  });

  it("writes multiple ranges", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A", "B"]],
      conditionalRules: [
        {
          type: "cellIs",
          priority: 1,
          operator: "greaterThan",
          formula: "100",
          range: "A1:A100",
        },
        {
          type: "colorScale",
          priority: 2,
          range: "B1:B100",
          colorScale: {
            cfvo: [{ type: "min" }, { type: "max" }],
            colors: ["FF63BE7B", "FFF8696B"],
          },
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cfs = findChildren(doc, "conditionalFormatting");
    expect(cfs.length).toBe(2);

    const byRange = new Map(cfs.map((cf: any) => [cf.attrs["sqref"], cf]));
    expect(byRange.has("A1:A100")).toBe(true);
    expect(byRange.has("B1:B100")).toBe(true);

    const cfRuleA = findChild(byRange.get("A1:A100"), "cfRule");
    expect(cfRuleA.attrs["type"]).toBe("cellIs");

    const cfRuleB = findChild(byRange.get("B1:B100"), "cfRule");
    expect(cfRuleB.attrs["type"]).toBe("colorScale");
  });

  it("writes stopIfTrue attribute", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
      conditionalRules: [
        {
          type: "cellIs",
          priority: 1,
          operator: "greaterThan",
          formula: "100",
          range: "A1:A100",
          stopIfTrue: true,
        },
      ],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    const cfRule = findChild(cf, "cfRule");
    expect(cfRule.attrs["stopIfTrue"]).toBe("true");
  });

  it("does not emit conditionalFormatting when none are provided", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Value"]],
    };

    const { xml } = writeXml(sheet);
    const doc = parseWorksheetXml(xml);

    const cf = findChild(doc, "conditionalFormatting");
    expect(cf).toBeUndefined();
  });

  it("places conditionalFormatting after mergeCells and before dataValidations", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B"],
            [1, 2],
          ],
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              operator: "greaterThan",
              formula: "0",
              range: "A2:B2",
            },
          ],
          dataValidations: [
            {
              type: "whole",
              operator: "greaterThan",
              formula1: "0",
              range: "A3:A10",
            },
          ],
        },
      ],
    });

    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml");

    const mergeCellsPos = sheetXml.indexOf("<mergeCells");
    const condFmtPos = sheetXml.indexOf("<conditionalFormatting");
    const dataValidationsPos = sheetXml.indexOf("<dataValidations");

    expect(mergeCellsPos).toBeGreaterThan(-1);
    expect(condFmtPos).toBeGreaterThan(-1);
    expect(dataValidationsPos).toBeGreaterThan(-1);
    expect(mergeCellsPos).toBeLessThan(condFmtPos);
    expect(condFmtPos).toBeLessThan(dataValidationsPos);
  });
});

// ── Reading Tests (Round-Trip) ──────────────────────────────────────

describe("conditional formatting — reading (round-trip)", () => {
  it("round-trips cellIs rule", async () => {
    const rules: ConditionalRule[] = [
      {
        type: "cellIs",
        priority: 1,
        operator: "greaterThan",
        formula: "100",
        range: "A1:A100",
      },
    ];

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: rules,
        },
      ],
    });

    const workbook = await readXlsx(data);
    expect(workbook.sheets.length).toBe(1);

    const sheet = workbook.sheets[0];
    expect(sheet.conditionalRules).toBeDefined();
    expect(sheet.conditionalRules!.length).toBe(1);

    const rule = sheet.conditionalRules![0];
    expect(rule.type).toBe("cellIs");
    expect(rule.priority).toBe(1);
    expect(rule.operator).toBe("greaterThan");
    expect(rule.formula).toBe("100");
    expect(rule.range).toBe("A1:A100");
  });

  it("round-trips cellIs with stopIfTrue", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              operator: "equal",
              formula: "0",
              range: "A1:A50",
              stopIfTrue: true,
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.stopIfTrue).toBe(true);
  });

  it("round-trips colorScale (2-color)", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "colorScale",
              priority: 1,
              range: "A1:A100",
              colorScale: {
                cfvo: [{ type: "min" }, { type: "max" }],
                colors: ["FF63BE7B", "FFF8696B"],
              },
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.type).toBe("colorScale");
    expect(rule.colorScale).toBeDefined();
    expect(rule.colorScale!.cfvo.length).toBe(2);
    expect(rule.colorScale!.cfvo[0].type).toBe("min");
    expect(rule.colorScale!.cfvo[1].type).toBe("max");
    expect(rule.colorScale!.colors.length).toBe(2);
    expect(rule.colorScale!.colors[0]).toBe("FF63BE7B");
    expect(rule.colorScale!.colors[1]).toBe("FFF8696B");
  });

  it("round-trips colorScale (3-color)", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "colorScale",
              priority: 1,
              range: "A1:A100",
              colorScale: {
                cfvo: [{ type: "min" }, { type: "percentile", value: "50" }, { type: "max" }],
                colors: ["FFF8696B", "FFFFEB84", "FF63BE7B"],
              },
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.colorScale!.cfvo.length).toBe(3);
    expect(rule.colorScale!.cfvo[1].type).toBe("percentile");
    expect(rule.colorScale!.cfvo[1].value).toBe("50");
    expect(rule.colorScale!.colors.length).toBe(3);
  });

  it("round-trips dataBar", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "dataBar",
              priority: 1,
              range: "A1:A100",
              dataBar: {
                cfvo: [{ type: "min" }, { type: "max" }],
                color: "FF638EC6",
              },
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.type).toBe("dataBar");
    expect(rule.dataBar).toBeDefined();
    expect(rule.dataBar!.cfvo.length).toBe(2);
    expect(rule.dataBar!.color).toBe("FF638EC6");
  });

  it("round-trips iconSet", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "iconSet",
              priority: 1,
              range: "A1:A100",
              iconSet: {
                iconSet: "3Arrows",
                cfvo: [
                  { type: "percent", value: "0" },
                  { type: "percent", value: "33" },
                  { type: "percent", value: "67" },
                ],
              },
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.type).toBe("iconSet");
    expect(rule.iconSet).toBeDefined();
    expect(rule.iconSet!.iconSet).toBe("3Arrows");
    expect(rule.iconSet!.cfvo.length).toBe(3);
    expect(rule.iconSet!.cfvo[1].value).toBe("33");
  });

  it("round-trips expression rule", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "expression",
              priority: 1,
              formula: "MOD(ROW(),2)=0",
              range: "A1:D100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.type).toBe("expression");
    expect(rule.formula).toBe("MOD(ROW(),2)=0");
  });

  it("round-trips containsText rule", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "containsText",
              priority: 1,
              text: "error",
              formula: 'NOT(ISERROR(SEARCH("error",A1)))',
              range: "A1:A100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.type).toBe("containsText");
    expect(rule.text).toBe("error");
    expect(rule.formula).toContain("SEARCH");
  });

  it("round-trips duplicateValues rule", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "duplicateValues",
              priority: 1,
              range: "A1:A100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rule = workbook.sheets[0].conditionalRules![0];
    expect(rule.type).toBe("duplicateValues");
    expect(rule.range).toBe("A1:A100");
  });

  it("round-trips multiple rules on same range", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value"]],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              operator: "greaterThan",
              formula: "100",
              range: "A1:A100",
            },
            {
              type: "cellIs",
              priority: 2,
              operator: "lessThan",
              formula: "0",
              range: "A1:A100",
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rules = workbook.sheets[0].conditionalRules!;
    expect(rules.length).toBe(2);
    expect(rules[0].operator).toBe("greaterThan");
    expect(rules[0].priority).toBe(1);
    expect(rules[1].operator).toBe("lessThan");
    expect(rules[1].priority).toBe(2);
  });

  it("round-trips multiple ranges", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A", "B"]],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              operator: "greaterThan",
              formula: "100",
              range: "A1:A100",
            },
            {
              type: "colorScale",
              priority: 2,
              range: "B1:B100",
              colorScale: {
                cfvo: [{ type: "min" }, { type: "max" }],
                colors: ["FF63BE7B", "FFF8696B"],
              },
            },
          ],
        },
      ],
    });

    const workbook = await readXlsx(data);
    const rules = workbook.sheets[0].conditionalRules!;
    expect(rules.length).toBe(2);

    const byRange = new Map(rules.map((r) => [r.range, r]));
    expect(byRange.get("A1:A100")!.type).toBe("cellIs");
    expect(byRange.get("B1:B100")!.type).toBe("colorScale");
  });
});

// ── Integration Tests ────────────────────────────────────────────────

describe("conditional formatting — integration (ZIP verification)", () => {
  it("writes correct conditionalFormatting XML in XLSX file", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Value", "Score"]],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              operator: "greaterThan",
              formula: "100",
              range: "A2:A100",
              style: {
                font: { bold: true, color: { rgb: "006100" } },
                fill: {
                  type: "pattern",
                  pattern: "solid",
                  bgColor: { rgb: "C6EFCE" },
                },
              },
            },
            {
              type: "colorScale",
              priority: 2,
              range: "B2:B100",
              colorScale: {
                cfvo: [{ type: "min" }, { type: "max" }],
                colors: ["FF63BE7B", "FFF8696B"],
              },
            },
          ],
        },
      ],
    });

    // Extract and verify the worksheet XML
    const sheetXml = await extractXml(data, "xl/worksheets/sheet1.xml");
    const doc = parseXml(sheetXml);

    const cfs = findChildren(doc, "conditionalFormatting");
    expect(cfs.length).toBe(2);

    // Verify the styles.xml has dxf entries
    const stylesXml = await extractXml(data, "xl/styles.xml");
    const stylesDoc = parseXml(stylesXml);
    const dxfs = findChild(stylesDoc, "dxfs");
    expect(dxfs).toBeDefined();
    expect(dxfs.attrs["count"]).toBe("1");
  });
});
