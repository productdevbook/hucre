import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { addFuturePrefixes } from "../src/xlsx/formula-prefix"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { ZipReader } from "../src/zip/reader"
import type { Cell, WriteOptions, WriteSheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #597 — functions added after Excel 2007 are stored in the file with a
// prefix: `_xlfn.XLOOKUP`, `_xlfn.IFS`, and `_xlfn._xlws.FILTER` /
// `_xlfn._xlws.SORT` ([MS-XLSX] §2.2.2, "future functions"). Excel adds
// the prefix when it saves and hides it in the formula bar, so callers
// write the bare name. hucre wrote that bare name into `<f>`, and a file
// that says `XLOOKUP(` names no function any reader knows: LibreOffice
// shows `#NAME?`.
//
// The writer now adds the prefix wherever formula text goes into the
// file — cells (all three `<f>` shapes and the streaming writers),
// conditional formatting, data validation, table totals and defined
// names. It has to be idempotent, because the reader hands back the
// prefixed text from Excel's files and that has to survive a rewrite.
// ═══════════════════════════════════════════════════════════════════════

const decoder = new TextDecoder("utf-8")

async function part(buf: Uint8Array, path: string): Promise<string> {
  return decoder.decode(await new ZipReader(buf).extract(path))
}

function unescape(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
}

/** Every element's text, in document order, unescaped. */
function texts(xml: string, tag: string): string[] {
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([^<]*)</${tag}>`, "g")
  return [...xml.matchAll(re)].map((m) => unescape(m[1]))
}

async function sheetXml(sheet: WriteSheet, options?: Omit<WriteOptions, "sheets">) {
  return part(await writeXlsx({ sheets: [sheet], ...options }), "xl/worksheets/sheet1.xml")
}

/** Write one formula into A1 and hand back the `<f>` text as stored. */
async function stored(formula: string): Promise<string> {
  const xml = await sheetXml({ name: "Sheet1", cells: new Map([["0,0", { formula }]]) })
  const [f] = texts(xml, "f")
  return f
}

describe("future functions are written with their stored prefix (#597)", () => {
  it("prefixes the five functions from the issue", async () => {
    const formulas = [
      'XLOOKUP("b",A1:A3,B1:B3)',
      'IFS(A1>1,"big",TRUE,"small")',
      'TEXTJOIN(",",TRUE,A1:A3)',
      "FILTER(A1:A3,B1:B3>1)",
      "SORT(A1:A3)",
    ]
    const xml = await sheetXml({
      name: "Sheet1",
      rows: [
        ["a", 1],
        ["b", 2],
        ["c", 3],
      ],
      cells: new Map(formulas.map((formula, i) => [`${i},3`, { formula }])),
    })
    expect(texts(xml, "f")).toEqual([
      '_xlfn.XLOOKUP("b",A1:A3,B1:B3)',
      '_xlfn.IFS(A1>1,"big",TRUE,"small")',
      '_xlfn.TEXTJOIN(",",TRUE,A1:A3)',
      "_xlfn._xlws.FILTER(A1:A3,B1:B3>1)",
      "_xlfn._xlws.SORT(A1:A3)",
    ])
  })

  it("prefixes dotted names and the dynamic-array family", async () => {
    expect(await stored("STDEV.S(A1:A9)+CEILING.MATH(B1)")).toBe(
      "_xlfn.STDEV.S(A1:A9)+_xlfn.CEILING.MATH(B1)",
    )
    expect(await stored("UNIQUE(SORTBY(A1:A9,B1:B9))")).toBe(
      "_xlfn.UNIQUE(_xlfn.SORTBY(A1:A9,B1:B9))",
    )
    expect(await stored("XMATCH(1,SEQUENCE(5))")).toBe("_xlfn.XMATCH(1,_xlfn.SEQUENCE(5))")
  })

  it("prefixes the newer functions Excel was seen to store with _xlfn.", async () => {
    // Neither [MS-XLSX] nor XlsxWriter lists these yet. Excel 16.114 for
    // Mac saved each with `_xlfn.` when the bare formula was typed in.
    // GROUPBY and PIVOTBY take a LAMBDA here: given a bare `SUM`, Excel
    // stores `_xleta.SUM` (#617); a LAMBDA keeps these cases about #608.
    const seen: Array<[string, string]> = [
      ["PERCENTOF(B2,B2:B3)", "_xlfn.PERCENTOF(B2,B2:B3)"],
      ["TRIMRANGE(B1:B10)", "_xlfn.TRIMRANGE(B1:B10)"],
      ['REGEXTEST(A1,"a")', '_xlfn.REGEXTEST(A1,"a")'],
      ['REGEXEXTRACT(A1,"b")', '_xlfn.REGEXEXTRACT(A1,"b")'],
      ['REGEXREPLACE(A1,"a","z")', '_xlfn.REGEXREPLACE(A1,"a","z")'],
      ['TRANSLATE(A1,"en","fi")', '_xlfn.TRANSLATE(A1,"en","fi")'],
      ["DETECTLANGUAGE(A1)", "_xlfn.DETECTLANGUAGE(A1)"],
      ['STOCKHISTORY("MSFT",DATE(2024,1,2))', '_xlfn.STOCKHISTORY("MSFT",DATE(2024,1,2))'],
      [
        "GROUPBY(A2:A3,B2:B3,LAMBDA(x,SUM(x)))",
        "_xlfn.GROUPBY(A2:A3,B2:B3,_xlfn.LAMBDA(_xlpm.x,SUM(_xlpm.x)))",
      ],
      [
        "PIVOTBY(A2:A3,A2:A3,B2:B3,LAMBDA(x,SUM(x)))",
        "_xlfn.PIVOTBY(A2:A3,A2:A3,B2:B3,_xlfn.LAMBDA(_xlpm.x,SUM(_xlpm.x)))",
      ],
    ]
    for (const [typed, saved] of seen) expect(await stored(typed)).toBe(saved)
  })

  it("does not prefix ENCODEURL, which no source shows prefixed", async () => {
    expect(await stored("ENCODEURL(A1)")).toBe("ENCODEURL(A1)")
  })

  it("leaves functions Excel 2007 already had alone", async () => {
    expect(await stored('IFERROR(XLOOKUP(A1,B:B,C:C),"")')).toBe(
      'IFERROR(_xlfn.XLOOKUP(A1,B:B,C:C),"")',
    )
    expect(await stored("SUM(A1:A3)*VLOOKUP(1,B:C,2,FALSE)")).toBe(
      "SUM(A1:A3)*VLOOKUP(1,B:C,2,FALSE)",
    )
  })

  // The spec lists these four as future functions with no prefix.
  it("does not prefix the future functions the spec stores bare", async () => {
    const f = "NETWORKDAYS.INTL(A1,B1)+WORKDAY.INTL(A1,1)+ISO.CEILING(C1)+ECMA.CEILING(C1,1)"
    expect(await stored(f)).toBe(f)
  })

  it("matches the name in any case and keeps the caller's spelling", async () => {
    expect(await stored("xlookup(1,A:A,B:B)")).toBe("_xlfn.xlookup(1,A:A,B:B)")
    expect(await stored("Filter(A1:A3,B1:B3)")).toBe("_xlfn._xlws.Filter(A1:A3,B1:B3)")
  })

  it("does not prefix an already-prefixed name again", async () => {
    for (const f of [
      '_xlfn.XLOOKUP("b",A1:A3,B1:B3)',
      "_xlfn._xlws.FILTER(A1:A3,B1:B3>1)",
      "_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))",
      "_XLFN.IFS(TRUE,1)",
    ]) {
      expect(await stored(f)).toBe(f)
    }
  })

  it("does not read SORTBY as SORT", async () => {
    expect(await stored("SORTBY(A1:A3,B1:B3)")).toBe("_xlfn.SORTBY(A1:A3,B1:B3)")
  })

  it("leaves string literals alone", async () => {
    expect(await stored('IF(A1="XLOOKUP(",IFS(TRUE,1),"say ""SORT("" twice")')).toBe(
      'IF(A1="XLOOKUP(",_xlfn.IFS(TRUE,1),"say ""SORT("" twice")',
    )
  })

  it("leaves quoted sheet names alone", async () => {
    expect(await stored("SUM('IFS(x'!A1,'it''s SORT('!B2)+SORT(C1:C3)")).toBe(
      "SUM('IFS(x'!A1,'it''s SORT('!B2)+_xlfn._xlws.SORT(C1:C3)",
    )
  })

  it("leaves structured-reference brackets alone", async () => {
    expect(await stored("SUM(T[[#This Row],[IFS(a)]])+T['[SORT(]")).toBe(
      "SUM(T[[#This Row],[IFS(a)]])+T['[SORT(]",
    )
    expect(await stored('TEXTJOIN(",",TRUE,T[Name])')).toBe('_xlfn.TEXTJOIN(",",TRUE,T[Name])')
  })

  it("leaves names that are not a call alone", async () => {
    // A sheet called IFS, a defined name called SORT, a cell reference
    // inside a longer name, and a sheet-qualified name call.
    expect(await stored("IFS!A1+SORT+MYSORT(1)+Book!SORT(1)+SORT.X(1)")).toBe(
      "IFS!A1+SORT+MYSORT(1)+Book!SORT(1)+SORT.X(1)",
    )
  })

  it("prefixes shared and array formulas, and leaves a shared slave empty", async () => {
    const cells = new Map<string, Partial<Cell>>([
      [
        "0,0",
        {
          formula: "XLOOKUP(A2,B:B,C:C)",
          formulaType: "shared",
          formulaSharedIndex: 0,
          formulaRef: "A1:A2",
        },
      ],
      ["1,0", { formula: "", formulaType: "shared", formulaSharedIndex: 0 }],
      ["0,1", { formula: "SORT(C1:C3)", formulaType: "array", formulaRef: "B1:B3" }],
    ])
    const xml = await sheetXml({ name: "Sheet1", cells })
    expect(xml).toContain('<f t="shared" si="0" ref="A1:A2">_xlfn.XLOOKUP(A2,B:B,C:C)</f>')
    expect(xml).toContain('<f t="shared" si="0"/>')
    expect(xml).toContain('<f t="array" ref="B1:B3">_xlfn._xlws.SORT(C1:C3)</f>')
  })

  it("prefixes formulas from the streaming writer", async () => {
    const stream = writeXlsxStream([[{ formula: 'TEXTJOIN(",",TRUE,B1:B3)' }]], { name: "S" })
    const chunks: Uint8Array[] = []
    for await (const chunk of stream) chunks.push(chunk)
    const buf = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
    let offset = 0
    for (const c of chunks) {
      buf.set(c, offset)
      offset += c.length
    }
    const xml = await part(buf, "xl/worksheets/sheet1.xml")
    expect(texts(xml, "f")).toEqual(['_xlfn.TEXTJOIN(",",TRUE,B1:B3)'])
  })

  it("prefixes conditional formatting and data validation formulas", async () => {
    const xml = await sheetXml({
      name: "Sheet1",
      rows: [[1]],
      conditionalRules: [
        {
          type: "expression",
          priority: 1,
          range: "A1:A3",
          formula: "MAXIFS(B:B,C:C,A1)>1",
          style: { font: { bold: true } },
        },
      ],
      dataValidations: [
        { type: "custom", range: "A1:A3", formula1: "ISNUMBER(XMATCH(A1,UNIQUE(B:B)))" },
        {
          type: "whole",
          operator: "between",
          range: "B1",
          formula1: "MINIFS(C:C,D:D,1)",
          formula2: "MAXIFS(C:C,D:D,1)",
        },
      ],
    })
    expect(texts(xml, "formula")).toEqual(["_xlfn.MAXIFS(B:B,C:C,A1)>1"])
    expect(texts(xml, "formula1")).toEqual([
      "ISNUMBER(_xlfn.XMATCH(A1,_xlfn.UNIQUE(B:B)))",
      "_xlfn.MINIFS(C:C,D:D,1)",
    ])
    expect(texts(xml, "formula2")).toEqual(["_xlfn.MAXIFS(C:C,D:D,1)"])
  })

  it("prefixes table total-row formulas and defined names", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Name"], ["a"], ["b"]],
          tables: [
            {
              name: "T",
              range: "A1:A4",
              showTotalRow: true,
              columns: [
                {
                  name: "Name",
                  totalFunction: "custom",
                  totalFormula: 'TEXTJOIN(",",TRUE,[Name])',
                },
              ],
            },
          ],
        },
      ],
      namedRanges: [
        { name: "Joined", range: 'TEXTJOIN(",",TRUE,Sheet1!$A$2:$A$3)' },
        { name: "Data", range: "Sheet1!$A$1:$A$3" },
      ],
    })
    expect(texts(await part(buf, "xl/tables/table1.xml"), "totalsRowFormula")).toEqual([
      '_xlfn.TEXTJOIN(",",TRUE,[Name])',
    ])
    expect(texts(await part(buf, "xl/workbook.xml"), "definedName")).toEqual([
      '_xlfn.TEXTJOIN(",",TRUE,Sheet1!$A$2:$A$3)',
      "Sheet1!$A$1:$A$3",
    ])
  })

  it("reads back the stored form and does not prefix it twice on rewrite", async () => {
    const first = await writeXlsx({
      sheets: [{ name: "Sheet1", cells: new Map([["0,0", { formula: "IFS(A2>1,SORT(B1:B3))" }]]) }],
    })
    const wb = await readXlsx(first)
    const formula = wb.sheets[0].cells?.get("0,0")?.formula
    expect(formula).toBe("_xlfn.IFS(A2>1,_xlfn._xlws.SORT(B1:B3))")

    const second = await writeXlsx({
      sheets: [{ name: "Sheet1", cells: new Map([["0,0", { formula: formula! }]]) }],
    })
    expect(texts(await part(second, "xl/worksheets/sheet1.xml"), "f")).toEqual([
      "_xlfn.IFS(A2>1,_xlfn._xlws.SORT(B1:B3))",
    ])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// #597, continued — the names LET and LAMBDA bind are stored with a
// prefix too: `_xlpm.x` where bound and where used, `_xlop.b` for an
// optional `[b]` where it is bound. With only `_xlfn.` added, Excel for
// Mac asked to repair the file and deleted both the formula and a named
// LAMBDA, leaving `=Double(B2)` at `#NAME?`.
//
// Every expected string below is what Excel 16.114 for Mac itself wrote
// when the bare formula was typed into a cell and the book saved.
// ═══════════════════════════════════════════════════════════════════════

describe("LET and LAMBDA names are written with their stored prefix (#597)", () => {
  // [bare as typed, stored as Excel saved it]
  const EXCEL: [string, string][] = [
    ["LET(x,2,x*3)", "_xlfn.LET(_xlpm.x,2,_xlpm.x*3)"],
    ["LAMBDA(x,x*2)(3)", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(3)"],
    // Nested, and an inner LET binding a name the outer one already has.
    [
      "LET(x,1,y,x+1,LET(x,10,x+y))",
      "_xlfn.LET(_xlpm.x,1,_xlpm.y,_xlpm.x+1,_xlfn.LET(_xlpm.x,10,_xlpm.x+_xlpm.y))",
    ],
    // A LET-bound LAMBDA called by name.
    ["LET(f,LAMBDA(n,n+1),f(2))", "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n+1),_xlpm.f(2))"],
    [
      "LET(p,LAMBDA(x,x*2),q,p,q(4))",
      "_xlfn.LET(_xlpm.p,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.q,_xlpm.p,_xlpm.q(4))",
    ],
    // An optional parameter is bound as `_xlop.` and used as `_xlpm.`.
    [
      "LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1)",
      "_xlfn.LAMBDA(_xlpm.a,_xlop.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a,_xlpm.a+_xlpm.b))(1)",
    ],
    ["LET(x,1,x)+LET(y,2,y)", "_xlfn.LET(_xlpm.x,1,_xlpm.x)+_xlfn.LET(_xlpm.y,2,_xlpm.y)"],
    // A LET name seen from a LAMBDA inside it.
    [
      "LET(k,3,SUM(MAP(A1:A3,LAMBDA(v,v*k))))",
      "_xlfn.LET(_xlpm.k,3,SUM(_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*_xlpm.k))))",
    ],
    // The arguments of a LAMBDA invoked on the spot are outside it: the
    // last `x` is the outer LET's.
    ["LET(x,5,LAMBDA(x,x+1)(x))", "_xlfn.LET(_xlpm.x,5,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(_xlpm.x))"],
    [
      "LAMBDA(x,LET(y,x*2,y+x))(4)",
      "_xlfn.LAMBDA(_xlpm.x,_xlfn.LET(_xlpm.y,_xlpm.x*2,_xlpm.y+_xlpm.x))(4)",
    ],
    ["LET(t,A1:A3,SUM(t))", "_xlfn.LET(_xlpm.t,A1:A3,SUM(_xlpm.t))"],
    ['LET(x,"x",x&"x")', '_xlfn.LET(_xlpm.x,"x",_xlpm.x&"x")'],
    ["LET(x,{1,2,3},SUM(x))", "_xlfn.LET(_xlpm.x,{1,2,3},SUM(_xlpm.x))"],
    // Excel rewrites a use to its binding's spelling.
    ["LET(Ab,2,aB*3)", "_xlfn.LET(_xlpm.Ab,2,_xlpm.Ab*3)"],
    // Spaces stay where they are (Excel also drops the ones before a
    // comma, which hucre does not need to copy).
    ["LET( x, 2, x * 3 )", "_xlfn.LET( _xlpm.x, 2, _xlpm.x * 3 )"],
    // Intersection of two bound ranges.
    [
      "LET(a,A1:A3,b,A1:A2,SUM(a b))",
      "_xlfn.LET(_xlpm.a,A1:A3,_xlpm.b,A1:A2,SUM(_xlpm.a _xlpm.b))",
    ],
    // With `a` bound, `A:A` is the name twice; `$` makes a cell.
    ["LET(a,2,a+SUM(A:A))", "_xlfn.LET(_xlpm.a,2,_xlpm.a+SUM(_xlpm.a:_xlpm.a))"],
    ["LET(a,2,a+$A$1+A$1+$A1)", "_xlfn.LET(_xlpm.a,2,_xlpm.a+$A$1+A$1+$A1)"],
    ["LET(c,2,c*SUM(C1:C1))", "_xlfn.LET(_xlpm.c,2,_xlpm.c*SUM(C1:C1))"],
  ]

  it("writes the forms Excel writes", async () => {
    for (const [bare, excel] of EXCEL) expect(await stored(bare)).toBe(excel)
  })

  it("does not prefix Excel's own form again", async () => {
    for (const [, excel] of EXCEL) expect(await stored(excel)).toBe(excel)
  })

  // These two ran with a workbook-level name `x` defined as 5, which is
  // what the bare `x` outside each scope means.
  it("leaves a name outside the scope that binds it alone", async () => {
    // A LET value is resolved before its own name is bound.
    expect(await stored("LET(y,x,y)+x")).toBe("_xlfn.LET(_xlpm.y,x,_xlpm.y)+x")
    expect(await stored("LET(y,x,x,y+1,x)+x")).toBe(
      "_xlfn.LET(_xlpm.y,x,_xlpm.x,_xlpm.y+1,_xlpm.x)+x",
    )
    expect(await stored("LAMBDA(x,x*2)(x)")).toBe("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(x)")
  })

  it("leaves a bound name in a string, a sheet name or a qualified name alone", async () => {
    expect(await stored(`LET(x,1,x+Sheet1!x+x!A1+'x'!A1+"x"&T[x])`)).toBe(
      `_xlfn.LET(_xlpm.x,1,_xlpm.x+Sheet1!x+x!A1+'x'!A1+"x"&T[x])`,
    )
  })

  it("prefixes a named LAMBDA, and leaves the cell that calls it alone", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[null, 1]],
          cells: new Map([["0,0", { formula: "Double(B2)" }]]),
        },
      ],
      namedRanges: [{ name: "Double", range: "LAMBDA(x,x*2)" }],
    })
    expect(texts(await part(buf, "xl/workbook.xml"), "definedName")).toEqual([
      "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)",
    ])
    expect(texts(await part(buf, "xl/worksheets/sheet1.xml"), "f")).toEqual(["Double(B2)"])
  })

  it("reads back the stored form and does not prefix it twice on rewrite", async () => {
    const first = await writeXlsx({
      sheets: [{ name: "Sheet1", cells: new Map([["0,0", { formula: "LET(x,2,x*3)" }]]) }],
    })
    const formula = (await readXlsx(first)).sheets[0].cells?.get("0,0")?.formula
    expect(formula).toBe("_xlfn.LET(_xlpm.x,2,_xlpm.x*3)")
    expect(await stored(formula!)).toBe(formula)
  })

  // A formula naming LET or LAMBDA goes through the scoped parse rather
  // than the plain scan; it has to treat everything else the same way.
  it("prefixes the rest of a formula the same way inside a LET", async () => {
    for (const f of [
      'XLOOKUP("b",A1:A3,B1:B3)',
      "UNIQUE(SORTBY(A1:A9,B1:B9))",
      "Filter(A1:A3,B1:B3)",
      "_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))",
      'IF(A1="XLOOKUP(",IFS(TRUE,1),"say ""SORT("" twice")',
      "SUM('IFS(x'!A1,'it''s SORT('!B2)+SORT(C1:C3)",
      "SUM(T[[#This Row],[IFS(a)]])+T['[SORT(]",
      "IFS!A1+SORT+MYSORT(1)+Book!SORT(1)+SORT.X(1)",
      "NETWORKDAYS.INTL(A1,B1)+ISO.CEILING(C1)",
    ]) {
      expect(await stored(`LET(z,1,${f})`)).toBe(`_xlfn.LET(_xlpm.z,1,${await stored(f)})`)
    }
    // A name that only contains LET still gets the functions after it.
    expect(await stored("OUTLET(1)+IFS(TRUE,1)")).toBe("OUTLET(1)+_xlfn.IFS(TRUE,1)")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// #617 — a built-in function passed as an argument to a LAMBDA helper is
// stored `_xleta.SUM`. Written bare, the cell computes `#NAME?`.
// ═══════════════════════════════════════════════════════════════════════

describe("a function passed as a value gets _xleta. (#617)", () => {
  it("writes what Excel 16.114 for Mac saved for each example in the issue", async () => {
    const cases: Array<[string, string]> = [
      ["GROUPBY(A2:A3,B2:B3,SUM)", "_xlfn.GROUPBY(A2:A3,B2:B3,_xleta.SUM)"],
      ["PIVOTBY(A2:A3,A2:A3,B2:B3,SUM)", "_xlfn.PIVOTBY(A2:A3,A2:A3,B2:B3,_xleta.SUM)"],
      ["BYROW(B2:B3,SUM)", "_xlfn.BYROW(B2:B3,_xleta.SUM)"],
      ["BYCOL(B2:B3,MAX)", "_xlfn.BYCOL(B2:B3,_xleta.MAX)"],
      ["MAP(B2:B3,ABS)", "_xlfn.MAP(B2:B3,_xleta.ABS)"],
      ["REDUCE(0,B2:B3,SUM)", "_xlfn.REDUCE(0,B2:B3,_xleta.SUM)"],
      ["SCAN(0,B2:B3,SUM)", "_xlfn.SCAN(0,B2:B3,_xleta.SUM)"],
      ["MAKEARRAY(2,2,PRODUCT)", "_xlfn.MAKEARRAY(2,2,_xleta.PRODUCT)"],
      // A future function gets only the _xleta., not _xlfn. as well.
      ["GROUPBY(B2:B3,B2:B3,PERCENTOF)", "_xlfn.GROUPBY(B2:B3,B2:B3,_xleta.PERCENTOF)"],
    ]
    for (const [bare, expected] of cases) {
      expect(addFuturePrefixes(bare)).toBe(expected)
      expect(await stored(bare)).toBe(expected)
    }
  })

  it("takes MAP's function from its last argument, after several arrays", () => {
    expect(addFuturePrefixes("MAP(A1:A2,B1:B2,SUM)")).toBe("_xlfn.MAP(A1:A2,B1:B2,_xleta.SUM)")
  })

  it("is idempotent, including on Excel's own form", () => {
    for (const f of [
      "_xlfn.GROUPBY(A2:A3,B2:B3,_xleta.SUM)",
      "_xlfn.GROUPBY(B2:B3,B2:B3,_xleta.PERCENTOF)",
      "_xlfn.MAP(B2:B3,_xleta.ABS)",
    ]) {
      expect(addFuturePrefixes(f)).toBe(f)
    }
    const once = addFuturePrefixes("BYROW(B2:B3,SUM)")
    expect(addFuturePrefixes(once)).toBe(once)
  })

  it("leaves a LET-bound name that spells a function alone", () => {
    expect(addFuturePrefixes("LET(sum,LAMBDA(x,x*2),MAP(B2:B3,sum))")).toBe(
      "_xlfn.LET(_xlpm.sum,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlfn.MAP(B2:B3,_xlpm.sum))",
    )
  })

  it("leaves a LAMBDA parameter that spells a function alone", () => {
    expect(addFuturePrefixes("LAMBDA(sum,BYROW(A1:A3,sum))")).toBe(
      "_xlfn.LAMBDA(_xlpm.sum,_xlfn.BYROW(A1:A3,_xlpm.sum))",
    )
  })

  it("leaves a defined name, such as a named LAMBDA, alone", () => {
    expect(addFuturePrefixes("MAP(B2:B3,MyLambda)")).toBe("_xlfn.MAP(B2:B3,MyLambda)")
    expect(addFuturePrefixes("BYROW(B2:B3,Sheet1!SUM)")).toBe("_xlfn.BYROW(B2:B3,Sheet1!SUM)")
  })

  it("leaves a cell reference that looks like a function alone", () => {
    expect(addFuturePrefixes("MAP(B2:B3,SUM1)")).toBe("_xlfn.MAP(B2:B3,SUM1)")
    expect(addFuturePrefixes("MAP(B2:B3,ATAN2)")).toBe("_xlfn.MAP(B2:B3,ATAN2)")
  })

  it("leaves a function name in a string literal or a sheet name alone", () => {
    expect(addFuturePrefixes('MAP(B2:B3,"SUM")')).toBe('_xlfn.MAP(B2:B3,"SUM")')
    expect(addFuturePrefixes('MAP(B2:B3,LAMBDA(x,x&"MAX"))')).toBe(
      '_xlfn.MAP(B2:B3,_xlfn.LAMBDA(_xlpm.x,_xlpm.x&"MAX"))',
    )
    expect(addFuturePrefixes("BYROW(SUM!B2:B3,SUM)")).toBe("_xlfn.BYROW(SUM!B2:B3,_xleta.SUM)")
    expect(addFuturePrefixes("BYROW('SUM'!B2:B3,SUM)")).toBe("_xlfn.BYROW('SUM'!B2:B3,_xleta.SUM)")
  })

  it("only prefixes the function slot, and only a lone name", () => {
    expect(addFuturePrefixes("MAP(SUM,ABS)")).toBe("_xlfn.MAP(SUM,_xleta.ABS)")
    expect(addFuturePrefixes("SUM(A1:A3)")).toBe("SUM(A1:A3)")
    expect(addFuturePrefixes("IF(A1,SUM,MAX)")).toBe("IF(A1,SUM,MAX)")
    expect(addFuturePrefixes("MAP(B2:B3,SUM+1)")).toBe("_xlfn.MAP(B2:B3,SUM+1)")
  })

  it("prefixes a nested helper's function and the outer one's", () => {
    expect(addFuturePrefixes("BYROW(MAP(B2:C3,ABS),SUM)")).toBe(
      "_xlfn.BYROW(_xlfn.MAP(B2:C3,_xleta.ABS),_xleta.SUM)",
    )
    expect(addFuturePrefixes("MAP(B2:B3,LAMBDA(x,REDUCE(0,x,SUM)))")).toBe(
      "_xlfn.MAP(B2:B3,_xlfn.LAMBDA(_xlpm.x,_xlfn.REDUCE(0,_xlpm.x,_xleta.SUM)))",
    )
  })

  it("matches in any case and keeps the caller's spelling, as _xlfn. does", () => {
    expect(addFuturePrefixes("map(B2:B3,abs)")).toBe("_xlfn.map(B2:B3,_xleta.abs)")
  })

  it("returns a formula without a call untouched", () => {
    expect(addFuturePrefixes("SUM")).toBe("SUM")
  })
})

describe("conditional-formatting thresholds (#608 follow-up)", () => {
  it("prefixes a formula cfvo in a colour scale, data bar and icon set", async () => {
    // `type: "formula"` is what the reader hands back from Excel's files; the
    // declared union does not list it, so the tests cast.
    const formula = (value: string) => ({ type: "formula" as unknown as "num", value })
    const xml = await sheetXml({
      name: "Sheet1",
      rows: [[1]],
      conditionalRules: [
        {
          type: "colorScale",
          priority: 1,
          range: "A1:A3",
          colorScale: {
            cfvo: [formula("MINIFS(B:B,C:C,1)"), { type: "max" }],
            colors: ["FFFF0000", "FF00FF00"],
          },
        },
        {
          type: "dataBar",
          priority: 2,
          range: "B1:B3",
          dataBar: { cfvo: [{ type: "min" }, formula("XLOOKUP(1,D:D,E:E)")], color: "FF638EC6" },
        },
        {
          type: "iconSet",
          priority: 3,
          range: "C1:C3",
          iconSet: {
            iconSet: "3Arrows",
            cfvo: [{ type: "percent", value: "0" }, formula("IFS(A1,1,TRUE,2)")],
          },
        },
      ],
    })
    const vals = [...xml.matchAll(/<cfvo type="formula" val="([^"]*)"/g)].map((m) => m[1])
    expect(vals).toEqual([
      "_xlfn.MINIFS(B:B,C:C,1)",
      "_xlfn.XLOOKUP(1,D:D,E:E)",
      "_xlfn.IFS(A1,1,TRUE,2)",
    ])
  })

  it("leaves a non-formula cfvo value alone", async () => {
    const xml = await sheetXml({
      name: "Sheet1",
      rows: [[1]],
      conditionalRules: [
        {
          type: "colorScale",
          priority: 1,
          range: "A1:A3",
          colorScale: {
            cfvo: [{ type: "num", value: "XLOOKUP(" }, { type: "max" }],
            colors: ["FFFF0000", "FF00FF00"],
          },
        },
      ],
    })
    expect(xml).toContain('<cfvo type="num" val="XLOOKUP("')
  })
})

describe("pathologically deep nesting (#608 follow-up)", () => {
  it("returns the formula unchanged past the depth cap instead of throwing", async () => {
    const deep = "(".repeat(100000) + "1" + ")".repeat(100000) + "+LET(x,1,x)"
    expect(addFuturePrefixes(deep)).toBe(deep)
    // And through the writer, which is where the RangeError surfaced.
    await expect(stored(deep)).resolves.toBe(deep)
  })

  it("still prefixes a formula nested as deep as Excel allows", () => {
    const nested = "(".repeat(64) + "1" + ")".repeat(64) + "+LET(x,1,x)"
    expect(addFuturePrefixes(nested)).toBe(
      "(".repeat(64) + "1" + ")".repeat(64) + "+_xlfn.LET(_xlpm.x,1,_xlpm.x)",
    )
  })
})
