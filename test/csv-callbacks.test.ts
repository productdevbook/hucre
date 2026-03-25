import { describe, expect, it } from "vitest";
import { parseCsv, parseCsvObjects } from "../src/csv/index";
import type { CellValue } from "../src/_types";

describe("onRow callback (#58)", () => {
  it("should call onRow for each parsed row", () => {
    const collected: { row: CellValue[]; index: number }[] = [];
    parseCsv("a,b\n1,2\n3,4", {
      onRow: (row, index) => {
        collected.push({ row: [...row], index });
      },
    });
    expect(collected).toEqual([
      { row: ["a", "b"], index: 0 },
      { row: ["1", "2"], index: 1 },
      { row: ["3", "4"], index: 2 },
    ]);
  });

  it("should call onRow with type-inferred values", () => {
    const collected: CellValue[][] = [];
    parseCsv("42,true,hello", {
      typeInference: true,
      onRow: (row) => {
        collected.push([...row]);
      },
    });
    expect(collected[0]).toEqual([42, true, "hello"]);
  });

  it("should not call onRow for empty input", () => {
    const collected: CellValue[][] = [];
    parseCsv("", {
      onRow: (row) => {
        collected.push([...row]);
      },
    });
    expect(collected).toEqual([]);
  });

  it("should call onRow after skip/filter operations", () => {
    const collected: CellValue[][] = [];
    parseCsv("# comment\na,b\n\nc,d", {
      comment: "#",
      skipEmptyRows: true,
      onRow: (row) => {
        collected.push([...row]);
      },
    });
    expect(collected).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("should allow progressive sum without buffering", () => {
    let sum = 0;
    parseCsv("10\n20\n30", {
      typeInference: true,
      onRow: (row) => {
        if (typeof row[0] === "number") sum += row[0];
      },
    });
    expect(sum).toBe(60);
  });
});

describe("transformHeader callback (#60)", () => {
  it("should transform headers in parseCsvObjects", () => {
    const { headers, data } = parseCsvObjects("Name,Age\nAlice,30", {
      header: true,
      transformHeader: (h) => h.toLowerCase(),
    });
    expect(headers).toEqual(["name", "age"]);
    expect(data[0]).toEqual({ name: "Alice", age: "30" });
  });

  it("should receive header index", () => {
    const indices: number[] = [];
    parseCsvObjects("A,B,C\n1,2,3", {
      header: true,
      transformHeader: (h, i) => {
        indices.push(i);
        return h;
      },
    });
    expect(indices).toEqual([0, 1, 2]);
  });

  it("should work with custom prefix", () => {
    const { headers } = parseCsvObjects("x,y\n1,2", {
      header: true,
      transformHeader: (h, i) => `col_${i}_${h}`,
    });
    expect(headers).toEqual(["col_0_x", "col_1_y"]);
  });
});

describe("transformValue callback (#60)", () => {
  it("should transform values in parseCsv", () => {
    const result = parseCsv("hello,world", {
      transformValue: (val) => (typeof val === "string" ? val.toUpperCase() : val),
    });
    expect(result).toEqual([["HELLO", "WORLD"]]);
  });

  it("should transform values after type inference", () => {
    const result = parseCsv("42,text", {
      typeInference: true,
      transformValue: (val) => {
        if (typeof val === "number") return val * 2;
        return val;
      },
    });
    expect(result).toEqual([[84, "text"]]);
  });

  it("should receive row and col indices in parseCsv", () => {
    const positions: [number, number][] = [];
    parseCsv("a,b\nc,d", {
      transformValue: (_val, _header, row, col) => {
        positions.push([row, col]);
        return _val;
      },
    });
    expect(positions).toEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]);
  });

  it("should transform values in parseCsvObjects", () => {
    const { data } = parseCsvObjects("price,qty\n10,5\n20,3", {
      header: true,
      typeInference: true,
      transformValue: (val, header) => {
        if (header === "price" && typeof val === "number") return val * 100;
        return val;
      },
    });
    expect(data).toEqual([
      { price: 1000, qty: 5 },
      { price: 2000, qty: 3 },
    ]);
  });

  it("should work with transformHeader and transformValue together", () => {
    const { headers, data } = parseCsvObjects("Name,Score\nAlice,95", {
      header: true,
      typeInference: true,
      transformHeader: (h) => h.toLowerCase(),
      transformValue: (val, header) => {
        if (header === "name" && typeof val === "string") return val.toUpperCase();
        return val;
      },
    });
    expect(headers).toEqual(["name", "score"]);
    // transformHeader uses transformed header name in transformValue for parseCsvObjects
    expect(data[0]).toEqual({ name: "ALICE", score: 95 });
  });
});
