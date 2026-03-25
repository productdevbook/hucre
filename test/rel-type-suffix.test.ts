import { describe, expect, it } from "vitest";
import { readXlsx, writeXlsx } from "../src/index";

describe("matchesRelType suffix matching (#130)", () => {
  it("should read a standard XLSX file (transitional namespace)", async () => {
    // This tests that the existing transitional namespace matching still works
    // after adding suffix-based matching.
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B"],
            [1, 2],
          ],
        },
      ],
    });
    const workbook = await readXlsx(data);
    expect(workbook.sheets.length).toBe(1);
    expect(workbook.sheets[0]!.rows[0]).toEqual(["A", "B"]);
    expect(workbook.sheets[0]!.rows[1]).toEqual([1, 2]);
  });

  it("should successfully roundtrip proving suffix matching does not break existing files", async () => {
    const original = {
      sheets: [
        {
          name: "TestSheet",
          rows: [
            ["Name", "Value"],
            ["Widget", 9.99],
            ["Gadget", 24.5],
          ],
        },
      ],
    };
    const data = await writeXlsx(original);
    const workbook = await readXlsx(data);
    expect(workbook.sheets[0]!.name).toBe("TestSheet");
    expect(workbook.sheets[0]!.rows.length).toBe(3);
  });
});
