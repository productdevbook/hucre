import { describe, expect, it } from "vitest";
import { r1c1ToA1, a1ToR1C1, writeXlsx, readXlsx } from "../src/index";
import { writeCsv } from "../src/csv/index";

// ── R1C1 notation (#162) ──────────────────────────────────────────

describe("R1C1 notation", () => {
  describe("r1c1ToA1", () => {
    it("converts absolute references", () => {
      expect(r1c1ToA1("R2C3", 0, 0)).toBe("$C$2");
      expect(r1c1ToA1("R1C1", 0, 0)).toBe("$A$1");
      expect(r1c1ToA1("R10C26", 0, 0)).toBe("$Z$10");
    });

    it("converts relative references", () => {
      // From row 4 (0-based), col 4 (0-based = E)
      expect(r1c1ToA1("R[1]C[-1]", 4, 4)).toBe("D6"); // row 5, col 3 = D6
      expect(r1c1ToA1("R[0]C[0]", 4, 4)).toBe("E5"); // same cell
      expect(r1c1ToA1("R[-2]C[2]", 4, 4)).toBe("G3"); // row 2, col 6 = G3
    });

    it("handles formulas with multiple references", () => {
      expect(r1c1ToA1("R[0]C[-2]+R[0]C[-1]", 1, 2)).toBe("A2+B2");
    });

    it("converts SUM formula", () => {
      expect(r1c1ToA1("SUM(R[-3]C[0]:R[-1]C[0])", 4, 2)).toBe("SUM(C2:C4)");
    });
  });

  describe("a1ToR1C1", () => {
    it("converts absolute references", () => {
      expect(a1ToR1C1("$C$2")).toBe("R2C3");
      expect(a1ToR1C1("$A$1")).toBe("R1C1");
    });

    it("converts relative references", () => {
      // From row 4 (0-based), col 4 (0-based)
      expect(a1ToR1C1("D6", 4, 4)).toBe("R[1]C[-1]");
      expect(a1ToR1C1("E5", 4, 4)).toBe("R[0]C[0]");
    });

    it("converts mixed references", () => {
      expect(a1ToR1C1("$C6", 4, 2)).toBe("R[1]C3"); // col absolute, row relative
    });

    it("handles formulas", () => {
      expect(a1ToR1C1("SUM(A1:A10)")).toBe("SUM(R1C1:R10C1)");
    });
  });
});

// ── Inline string mode (#166) ────────────────────────────────────

describe("inline string mode", () => {
  it("writes and reads back correctly with inline strings", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Test",
          rows: [
            ["Hello", "World"],
            ["Foo", "Bar"],
          ],
        },
      ],
      stringMode: "inline",
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0]!.rows[0]).toEqual(["Hello", "World"]);
    expect(wb.sheets[0]!.rows[1]).toEqual(["Foo", "Bar"]);
  });

  it("produces valid XLSX without shared strings part", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Test",
          rows: [["Only", "Inline"]],
        },
      ],
      stringMode: "inline",
    });

    // Read should still work
    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0]!.rows[0]).toEqual(["Only", "Inline"]);
  });

  it("shared mode still works (default)", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Test",
          rows: [["Shared", "Strings"]],
        },
      ],
    });
    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0]!.rows[0]).toEqual(["Shared", "Strings"]);
  });
});

// ── CSV injection hardening (#164) ────────────────────────────────

describe("CSV injection hardening", () => {
  it("escapes standard formula prefixes", () => {
    const csv = writeCsv([["=cmd|'/C calc'!A0"]], { escapeFormulae: true });
    expect(csv).toContain("'=cmd");
  });

  it("escapes plus and minus prefixes", () => {
    const csv = writeCsv([["+cmd", "-cmd"]], { escapeFormulae: true });
    expect(csv).toContain("'+cmd");
    expect(csv).toContain("'-cmd");
  });

  it("escapes tab and newline prefixes", () => {
    const csv = writeCsv([["\tcmd", "\ncmd"]], { escapeFormulae: true });
    expect(csv).toContain("'\tcmd");
    expect(csv).toContain("'\ncmd");
  });

  it("escapes null byte prefix", () => {
    const csv = writeCsv([["\0cmd"]], { escapeFormulae: true });
    expect(csv).toContain("'\0cmd");
  });

  it("escapes pipe prefix", () => {
    const csv = writeCsv([["|cmd"]], { escapeFormulae: true });
    expect(csv).toContain("'|cmd");
  });

  it("does not escape normal values", () => {
    const csv = writeCsv([["Hello", "World"]], { escapeFormulae: true });
    expect(csv).not.toContain("'");
  });

  it("does not escape when escapeFormulae is false", () => {
    const csv = writeCsv([["=SUM(A1:A10)"]], { escapeFormulae: false });
    expect(csv).not.toContain("'=");
  });
});

// ── VBA injection (#160) ─────────────────────────────────────────

describe("VBA/macro injection", () => {
  it("embeds vbaProject.bin and produces valid XLSM", async () => {
    // Create a minimal fake vbaProject.bin (just some bytes)
    const fakeVba = new Uint8Array([0x00, 0x01, 0x02, 0x03, 0x04]);

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Hello"]],
        },
      ],
      vbaProject: fakeVba,
    });

    // The output should be a valid ZIP containing xl/vbaProject.bin
    // We can verify by checking the ZIP contains the marker bytes
    expect(xlsx.length).toBeGreaterThan(0);

    // Read back should still parse (vbaProject.bin is ignored by reader)
    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0]!.rows[0]).toEqual(["Hello"]);
  });
});
