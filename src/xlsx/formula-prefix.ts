// ── Future-function prefixes ────────────────────────────────────────
//
// Functions added to Excel after the 2007 file format are stored in the
// file under a prefixed name — `_xlfn.XLOOKUP`, `_xlfn.IFS`, and
// `_xlfn._xlws.FILTER` / `_xlfn._xlws.SORT` for the two worksheet-only
// ones. Excel adds the prefix when it saves and hides it in the formula
// bar, so callers write the bare name. Written bare, the file names a
// function no reader knows and the cell computes `#NAME?` (#597).
//
// The names are the "future functions" of [MS-XLSX] §2.2.2 (Functions,
// and the `future-function-list` / `worksheet-only-function-list` rules
// of the Formulas grammar):
// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf
// Added to that: XMATCH, ARRAYTOTEXT, VALUETOTEXT and IMAGE, which the
// spec does not list yet but Excel stores with `_xlfn.` too; XlsxWriter's
// future-function table (`_prepare_formula` in xlsxwriter/worksheet.py)
// has them. And DETECTLANGUAGE, GROUPBY, PERCENTOF, PIVOTBY, REGEXEXTRACT,
// REGEXREPLACE, REGEXTEST, STOCKHISTORY, TRANSLATE and TRIMRANGE, which
// neither list has yet: Excel 16.114 for Mac saved each with `_xlfn.`
// when it was typed in bare. Every name is in one of those lists or was
// seen that way: prefixing a name Excel stores bare would break a formula
// that works today. ENCODEURL is not here for that reason — neither list
// has it, and Excel for Mac, which lacks the function, stores it bare.
//
// Left out on purpose:
// - ECMA.CEILING, ISO.CEILING, NETWORKDAYS.INTL, WORKDAY.INTL: the spec
//   lists them as future functions with no prefix.
// - `_xlfn._xlws.PY`: a Python cell is a solo function with metadata of
//   its own, not something a prefix makes work.
// - `_xlfn.SINGLE` / `_xlfn.ANCHORARRAY`: the stored forms of the `@` and
//   `#` operators, not names a caller types.
//
// ── Parameter prefixes ──────────────────────────────────────────────
//
// A name that LET or LAMBDA binds is stored as `_xlpm.x`, both where it
// is bound and everywhere it is used: `_xlfn.LET(_xlpm.x,2,_xlpm.x*3)`.
// An optional LAMBDA parameter, typed `[b]`, is stored `_xlop.b` where it
// is bound and `_xlpm.b` where it is used. Written bare, Excel for Mac
// asks to repair the file and removes the formula; a named LAMBDA goes
// with it, and every cell that calls it computes `#NAME?` (#597).
//
// That needs scope, which a scan for `NAME(` cannot see, so formulas that
// mention LET or LAMBDA take a second path that parses the argument lists.
// What it does is what Excel 16.114 for Mac wrote when given the bare
// form (each case is a test in test/formula-future-prefix.test.ts):
// - LET(n1, v1, n2, v2, …, calc): n1 is in scope from v2 onwards, not in
//   v1, where `x` still means whatever it meant outside.
// - LAMBDA(p1, …, body): the parameters are in scope in the body only.
//   The arguments of a LAMBDA invoked on the spot, `LAMBDA(x,x*2)(x)`,
//   are outside it, so that last `x` stays bare.
// - A bound name called as a function, `f(2)`, is `_xlpm.f(2)`.
// - A use is rewritten to the spelling of its binding: `LET(Ab,2,aB)` is
//   stored `_xlfn.LET(_xlpm.Ab,2,_xlpm.Ab)`. Names ignore case.
// - Inside a scope that binds `a`, `A:A` is `_xlpm.a:_xlpm.a`, but `$A$1`,
//   `A$1` and `$A1` are still cells.

const XLFN = "_xlfn."
const XLWS = "_xlfn._xlws."

const XLFN_NAMES = [
  "ACOT",
  "ACOTH",
  "AGGREGATE",
  "ARABIC",
  "ARRAYTOTEXT",
  "BASE",
  "BETA.DIST",
  "BETA.INV",
  "BINOM.DIST",
  "BINOM.DIST.RANGE",
  "BINOM.INV",
  "BITAND",
  "BITLSHIFT",
  "BITOR",
  "BITRSHIFT",
  "BITXOR",
  "BYCOL",
  "BYROW",
  "CEILING.MATH",
  "CEILING.PRECISE",
  "CHISQ.DIST",
  "CHISQ.DIST.RT",
  "CHISQ.INV",
  "CHISQ.INV.RT",
  "CHISQ.TEST",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "COMBINA",
  "CONCAT",
  "CONFIDENCE.NORM",
  "CONFIDENCE.T",
  "COT",
  "COTH",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "CSC",
  "CSCH",
  "DAYS",
  "DECIMAL",
  "DETECTLANGUAGE",
  "DROP",
  "ERF.PRECISE",
  "ERFC.PRECISE",
  "EXPAND",
  "EXPON.DIST",
  "F.DIST",
  "F.DIST.RT",
  "F.INV",
  "F.INV.RT",
  "F.TEST",
  "FIELDVALUE",
  "FILTERXML",
  "FLOOR.MATH",
  "FLOOR.PRECISE",
  "FORECAST.ETS",
  "FORECAST.ETS.CONFINT",
  "FORECAST.ETS.SEASONALITY",
  "FORECAST.ETS.STAT",
  "FORECAST.LINEAR",
  "FORMULATEXT",
  "GAMMA",
  "GAMMA.DIST",
  "GAMMA.INV",
  "GAMMALN.PRECISE",
  "GAUSS",
  "GROUPBY",
  "HSTACK",
  "HYPGEOM.DIST",
  "IFNA",
  "IFS",
  "IMAGE",
  "IMCOSH",
  "IMCOT",
  "IMCSC",
  "IMCSCH",
  "IMSEC",
  "IMSECH",
  "IMSINH",
  "IMTAN",
  "ISFORMULA",
  "ISOMITTED",
  "ISOWEEKNUM",
  "LAMBDA",
  "LET",
  "LOGNORM.DIST",
  "LOGNORM.INV",
  "MAKEARRAY",
  "MAP",
  "MAXIFS",
  "MINIFS",
  "MODE.MULT",
  "MODE.SNGL",
  "MUNIT",
  "NEGBINOM.DIST",
  "NORM.DIST",
  "NORM.INV",
  "NORM.S.DIST",
  "NORM.S.INV",
  "NUMBERVALUE",
  "PDURATION",
  "PERCENTILE.EXC",
  "PERCENTILE.INC",
  "PERCENTOF",
  "PERCENTRANK.EXC",
  "PERCENTRANK.INC",
  "PERMUTATIONA",
  "PHI",
  "PIVOTBY",
  "POISSON.DIST",
  "PQSOURCE",
  "PYTHON_STR",
  "PYTHON_TYPE",
  "PYTHON_TYPENAME",
  "QUARTILE.EXC",
  "QUARTILE.INC",
  "QUERYSTRING",
  "RANDARRAY",
  "RANK.AVG",
  "RANK.EQ",
  "REDUCE",
  "REGEXEXTRACT",
  "REGEXREPLACE",
  "REGEXTEST",
  "RRI",
  "SCAN",
  "SEC",
  "SECH",
  "SEQUENCE",
  "SHEET",
  "SHEETS",
  "SKEW.P",
  "SORTBY",
  "STDEV.P",
  "STDEV.S",
  "STOCKHISTORY",
  "SWITCH",
  "T.DIST",
  "T.DIST.2T",
  "T.DIST.RT",
  "T.INV",
  "T.INV.2T",
  "T.TEST",
  "TAKE",
  "TEXTAFTER",
  "TEXTBEFORE",
  "TEXTJOIN",
  "TEXTSPLIT",
  "TOCOL",
  "TOROW",
  "TRANSLATE",
  "TRIMRANGE",
  "UNICHAR",
  "UNICODE",
  "UNIQUE",
  "VALUETOTEXT",
  "VAR.P",
  "VAR.S",
  "VSTACK",
  "WEBSERVICE",
  "WEIBULL.DIST",
  "WRAPCOLS",
  "WRAPROWS",
  "XLOOKUP",
  "XMATCH",
  "XOR",
  "Z.TEST",
]

/** Upper-case function name → the prefix it is stored under. */
const FUTURE_FUNCTION_PREFIXES: ReadonlyMap<string, string> = new Map([
  ...XLFN_NAMES.map((name): [string, string] => [name, XLFN]),
  ["FILTER", XLWS],
  ["SORT", XLWS],
])

// ── Functions passed as values: _xleta. ─────────────────────────────
//
// The LAMBDA helpers take a function where they would take a LAMBDA, and
// a built-in function can stand in for it: `GROUPBY(A2:A3,B2:B3,SUM)`.
// Excel stores that bare name as `_xleta.SUM`, and a file with `SUM`
// there computes `#NAME?` (#617). What Excel 16.114 for Mac saved when
// the bare forms were typed in:
// `_xlfn.GROUPBY(A2:A3,B2:B3,_xleta.SUM)`, `_xlfn.MAP(B2:B3,_xleta.ABS)`,
// `_xlfn.MAKEARRAY(2,2,_xleta.PRODUCT)`. A future function gets only the
// `_xleta.`, never `_xlfn.` as well: `GROUPBY(B2:B3,B2:B3,_xleta.PERCENTOF)`.
//
// `SUM` cannot be told from a defined name by its spelling, so three
// things must all hold before a name is prefixed:
// 1. Position. It is the whole argument, in the slot where the helper
//    expects a function (ETA_ARGUMENT below). `MAP(B2:B3,Total)` with
//    Total a named LAMBDA stays as it is, and a defined name anywhere
//    else is never looked at: only a helper's own arguments are.
// 2. Not bound. A LET or LAMBDA name that spells a function,
//    `LET(sum,LAMBDA(x,x),MAP(B2:B3,sum))`, is `_xlpm.sum`, not eta.
// 3. Spelling. It is in ETA_FUNCTIONS: the classic Excel function names
//    plus every future function above. Some list is unavoidable, since
//    `MyLambda` and `SUM` look alike; the position rule keeps it from
//    having to be perfect. A name missing from it costs one unprefixed
//    argument, and the list is not consulted for names outside a
//    function slot, so it cannot prefix a defined name elsewhere.
// Excel is not known to refuse a defined name that spells a function
// (`SUM` may be accepted; only cell-like spellings and R/C are refused),
// and this was not checked against Excel, so rules 1 and 2 are not
// optional. A defined name that is in the list and sits in a function
// slot is the one case this gets wrong.
// Left out of the list on purpose: names that read as cells (`ATAN2` is
// column ATA, row 2), `TRUE`/`FALSE`/`NA`, which are values bare, and the
// single letters `N` and `T`, likelier a defined name than a function.
const ETA = "_xleta."

const CLASSIC_NAMES = [
  "ABS",
  "ACCRINT",
  "ACCRINTM",
  "ACOS",
  "ACOSH",
  "ADDRESS",
  "AMORDEGRC",
  "AMORLINC",
  "AND",
  "AREAS",
  "ASC",
  "ASIN",
  "ASINH",
  "ATAN",
  "ATANH",
  "AVEDEV",
  "AVERAGE",
  "AVERAGEA",
  "AVERAGEIF",
  "AVERAGEIFS",
  "BAHTTEXT",
  "BESSELI",
  "BESSELJ",
  "BESSELK",
  "BESSELY",
  "BETADIST",
  "BETAINV",
  "BIN2DEC",
  "BIN2HEX",
  "BIN2OCT",
  "BINOMDIST",
  "CEILING",
  "CELL",
  "CHAR",
  "CHIDIST",
  "CHIINV",
  "CHITEST",
  "CHOOSE",
  "CLEAN",
  "CODE",
  "COLUMN",
  "COLUMNS",
  "COMBIN",
  "COMPLEX",
  "CONCATENATE",
  "CONFIDENCE",
  "CONVERT",
  "CORREL",
  "COS",
  "COSH",
  "COUNT",
  "COUNTA",
  "COUNTBLANK",
  "COUNTIF",
  "COUNTIFS",
  "COUPDAYBS",
  "COUPDAYS",
  "COUPDAYSNC",
  "COUPNCD",
  "COUPNUM",
  "COUPPCD",
  "COVAR",
  "CRITBINOM",
  "CUBEKPIMEMBER",
  "CUBEMEMBER",
  "CUBEMEMBERPROPERTY",
  "CUBERANKEDMEMBER",
  "CUBESET",
  "CUBESETCOUNT",
  "CUBEVALUE",
  "CUMIPMT",
  "CUMPRINC",
  "DATE",
  "DATEDIF",
  "DATEVALUE",
  "DAVERAGE",
  "DAY",
  "DAYS360",
  "DB",
  "DBCS",
  "DCOUNT",
  "DCOUNTA",
  "DDB",
  "DEC2BIN",
  "DEC2HEX",
  "DEC2OCT",
  "DEGREES",
  "DELTA",
  "DEVSQ",
  "DGET",
  "DISC",
  "DMAX",
  "DMIN",
  "DOLLAR",
  "DOLLARDE",
  "DOLLARFR",
  "DPRODUCT",
  "DSTDEV",
  "DSTDEVP",
  "DSUM",
  "DURATION",
  "DVAR",
  "DVARP",
  "ECMA.CEILING",
  "EDATE",
  "EFFECT",
  "ENCODEURL",
  "EOMONTH",
  "ERF",
  "ERFC",
  "ERROR.TYPE",
  "EUROCONVERT",
  "EVEN",
  "EXACT",
  "EXP",
  "EXPONDIST",
  "FACT",
  "FACTDOUBLE",
  "FDIST",
  "FILTER",
  "FIND",
  "FINDB",
  "FINV",
  "FISHER",
  "FISHERINV",
  "FIXED",
  "FLOOR",
  "FORECAST",
  "FREQUENCY",
  "FTEST",
  "FV",
  "FVSCHEDULE",
  "GAMMADIST",
  "GAMMAINV",
  "GAMMALN",
  "GCD",
  "GEOMEAN",
  "GESTEP",
  "GETPIVOTDATA",
  "GROWTH",
  "HARMEAN",
  "HEX2BIN",
  "HEX2DEC",
  "HEX2OCT",
  "HLOOKUP",
  "HOUR",
  "HYPERLINK",
  "HYPGEOMDIST",
  "IF",
  "IFERROR",
  "IMABS",
  "IMAGINARY",
  "IMARGUMENT",
  "IMCONJUGATE",
  "IMCOS",
  "IMDIV",
  "IMEXP",
  "IMLN",
  "IMLOG10",
  "IMLOG2",
  "IMPOWER",
  "IMPRODUCT",
  "IMREAL",
  "IMSIN",
  "IMSQRT",
  "IMSUB",
  "IMSUM",
  "INDEX",
  "INDIRECT",
  "INFO",
  "INT",
  "INTERCEPT",
  "INTRATE",
  "IPMT",
  "IRR",
  "ISBLANK",
  "ISERR",
  "ISERROR",
  "ISEVEN",
  "ISLOGICAL",
  "ISNA",
  "ISNONTEXT",
  "ISNUMBER",
  "ISO.CEILING",
  "ISODD",
  "ISPMT",
  "ISREF",
  "ISTEXT",
  "JIS",
  "KURT",
  "LARGE",
  "LCM",
  "LEFT",
  "LEFTB",
  "LEN",
  "LENB",
  "LINEST",
  "LN",
  "LOG",
  "LOGEST",
  "LOGINV",
  "LOGNORMDIST",
  "LOOKUP",
  "LOWER",
  "MATCH",
  "MAX",
  "MAXA",
  "MDETERM",
  "MDURATION",
  "MEDIAN",
  "MID",
  "MIDB",
  "MIN",
  "MINA",
  "MINUTE",
  "MINVERSE",
  "MIRR",
  "MMULT",
  "MOD",
  "MODE",
  "MONTH",
  "MROUND",
  "MULTINOMIAL",
  "NEGBINOMDIST",
  "NETWORKDAYS",
  "NETWORKDAYS.INTL",
  "NOMINAL",
  "NORMDIST",
  "NORMINV",
  "NORMSDIST",
  "NORMSINV",
  "NOT",
  "NOW",
  "NPER",
  "NPV",
  "OCT2BIN",
  "OCT2DEC",
  "OCT2HEX",
  "ODD",
  "ODDFPRICE",
  "ODDFYIELD",
  "ODDLPRICE",
  "ODDLYIELD",
  "OFFSET",
  "OR",
  "PEARSON",
  "PERCENTILE",
  "PERCENTRANK",
  "PERMUT",
  "PHONETIC",
  "PI",
  "PMT",
  "POISSON",
  "POWER",
  "PPMT",
  "PRICE",
  "PRICEDISC",
  "PRICEMAT",
  "PROB",
  "PRODUCT",
  "PROPER",
  "PV",
  "QUARTILE",
  "QUOTIENT",
  "RADIANS",
  "RAND",
  "RANDBETWEEN",
  "RANK",
  "RATE",
  "RECEIVED",
  "REPLACE",
  "REPLACEB",
  "REPT",
  "RIGHT",
  "RIGHTB",
  "ROMAN",
  "ROUND",
  "ROUNDDOWN",
  "ROUNDUP",
  "ROW",
  "ROWS",
  "RSQ",
  "RTD",
  "SEARCH",
  "SEARCHB",
  "SECOND",
  "SERIESSUM",
  "SIGN",
  "SIN",
  "SINH",
  "SKEW",
  "SLN",
  "SLOPE",
  "SMALL",
  "SORT",
  "SQRT",
  "SQRTPI",
  "STANDARDIZE",
  "STDEV",
  "STDEVA",
  "STDEVP",
  "STDEVPA",
  "STEYX",
  "SUBSTITUTE",
  "SUBTOTAL",
  "SUM",
  "SUMIF",
  "SUMIFS",
  "SUMPRODUCT",
  "SUMSQ",
  "SUMX2MY2",
  "SUMX2PY2",
  "SUMXMY2",
  "SYD",
  "TAN",
  "TANH",
  "TBILLEQ",
  "TBILLPRICE",
  "TBILLYIELD",
  "TDIST",
  "TEXT",
  "TIME",
  "TIMEVALUE",
  "TINV",
  "TOCOL",
  "TODAY",
  "TRANSPOSE",
  "TREND",
  "TRIM",
  "TRIMMEAN",
  "TRUNC",
  "TTEST",
  "TYPE",
  "UPPER",
  "VALUE",
  "VAR",
  "VARA",
  "VARP",
  "VARPA",
  "VDB",
  "VLOOKUP",
  "WEEKDAY",
  "WEEKNUM",
  "WEIBULL",
  "WORKDAY",
  "WORKDAY.INTL",
  "XIRR",
  "XNPV",
  "YEAR",
  "YEARFRAC",
  "YIELD",
  "YIELDDISC",
  "YIELDMAT",
  "ZTEST",
]

/** Upper-case function names that can be passed as a value. */
const ETA_FUNCTIONS: ReadonlySet<string> = new Set([
  ...CLASSIC_NAMES,
  ...XLFN_NAMES,
  "FILTER",
  "SORT",
])

/**
 * Upper-case helper → the index of the argument that takes a function.
 * `-1` is the last one: MAP takes any number of arrays before it.
 */
const ETA_ARGUMENT: ReadonlyMap<string, number> = new Map([
  ["GROUPBY", 2],
  ["PIVOTBY", 3],
  ["BYROW", 1],
  ["BYCOL", 1],
  ["REDUCE", 2],
  ["SCAN", 2],
  ["MAKEARRAY", 2],
  ["MAP", -1],
])

/**
 * A character that can be part of a function or defined name. Dots are in
 * (`STDEV.S`, `_xlfn.IFS`) so a prefixed name scans as one token, and
 * anything outside ASCII is in because defined names may use it.
 */
function isNameChar(code: number): boolean {
  return (
    (code >= 0x41 && code <= 0x5a) || // A-Z
    (code >= 0x61 && code <= 0x7a) || // a-z
    (code >= 0x30 && code <= 0x39) || // 0-9
    code === 0x5f || // _
    code === 0x2e || // .
    code === 0x5c || // \
    code > 0x7f
  )
}

/**
 * Skip a string literal or quoted sheet name starting at `i`; a doubled
 * quote is an escape. Returns the index after the closing quote.
 */
function skipQuoted(formula: string, i: number): number {
  const quote = formula.charCodeAt(i)
  const n = formula.length
  i++
  while (i < n) {
    if (formula.charCodeAt(i) === quote) {
      if (formula.charCodeAt(i + 1) === quote) i += 2
      else break
    } else i++
  }
  return i + 1
}

/**
 * Skip a structured reference (`T[[#This Row],[Col]]`) or external book
 * index (`[1]Sheet1!A1`) starting at `i`. Brackets nest, and `'` escapes
 * the next character inside them. Returns the index after the last `]`.
 */
function skipBracket(formula: string, i: number): number {
  const n = formula.length
  let depth = 0
  while (i < n) {
    const c = formula.charCodeAt(i)
    if (c === 0x27) i++
    else if (c === 0x5b) depth++
    else if (c === 0x5d && --depth === 0) break
    i++
  }
  return i + 1
}

/**
 * `LET(`, `LAMBDA(` or a LAMBDA helper (which may take `_xleta.`), prefixed
 * or not; a false hit only costs the slow path.
 */
const BINDS_NAMES = /(?:LET|LAMBDA|GROUPBY|PIVOTBY|BYROW|BYCOL|MAP|REDUCE|SCAN|MAKEARRAY)\(/i

/**
 * Add the stored prefixes to `formula`: `_xlfn.` (or `_xlfn._xlws.`) on
 * every call of a future function, so `XLOOKUP(` is written as
 * `_xlfn.XLOOKUP(` (#597), and `_xlpm.` / `_xlop.` on the names LET and
 * LAMBDA bind.
 *
 * Only a name immediately followed by `(` is a call. String literals,
 * quoted sheet names and `[…]` structured references are skipped whole,
 * and a name after `!` is a sheet- or workbook-scoped name, not a
 * built-in. Already-prefixed text scans as a different name and is left
 * as it is, so the reader's output can be written back unchanged.
 */
export function addFuturePrefixes(formula: string): string {
  // No call, nothing to prefix; this is on the per-cell write path.
  if (!formula.includes("(")) return formula
  if (BINDS_NAMES.test(formula)) return addScopedPrefixes(formula)
  let out = ""
  let last = 0
  let i = 0
  const n = formula.length
  while (i < n) {
    const ch = formula.charCodeAt(i)
    if (ch === 0x22 /* " */ || ch === 0x27 /* ' */) {
      i = skipQuoted(formula, i)
    } else if (ch === 0x5b /* [ */) {
      i = skipBracket(formula, i)
    } else if (isNameChar(ch)) {
      const start = i
      while (i < n && isNameChar(formula.charCodeAt(i))) i++
      if (formula.charCodeAt(i) !== 0x28 /* ( */) continue
      if (start > 0 && formula.charCodeAt(start - 1) === 0x21 /* ! */) continue
      const prefix = FUTURE_FUNCTION_PREFIXES.get(formula.slice(start, i).toUpperCase())
      if (prefix === undefined) continue
      out += formula.slice(last, start) + prefix
      last = start
    } else {
      i++
    }
  }
  return out === "" ? formula : out + formula.slice(last)
}

// ── The scoped path ─────────────────────────────────────────────────

const XLPM = "_xlpm."
const XLOP = "_xlop."

interface Span {
  start: number
  end: number
}

/**
 * What the parse keeps of a formula: names, calls with their argument
 * lists, and `[…]` groups (an optional LAMBDA parameter is one). Anything
 * else is `other`, kept only so an argument of `x+1` is not mistaken for
 * the bare name `x`.
 */
type Item =
  | { kind: "name"; span: Span }
  | { kind: "bracket"; span: Span }
  | { kind: "call"; callee: Span | undefined; args: Item[][] }
  | { kind: "other" }

const OTHER: Item = { kind: "other" }

/** Replace `formula.slice(start, end)` with `text`; `start === end` inserts. */
interface Edit extends Span {
  text: string
}

/** Upper-case name → the spelling it was bound with. */
type Scope = ReadonlyMap<string, string>

/**
 * How deep the scoped path will nest before it gives up. Excel allows 64
 * levels of nesting, so a real formula is nowhere near this; the cap is for
 * text like `"(".repeat(100000) + "1" + ")".repeat(100000) + "+LET(x,1,x)"`,
 * where `parseList`, `resolve` and `call` recurse once per level and the
 * stack runs out first. That used to surface as a RangeError out of
 * `writeXlsx`.
 */
const MAX_DEPTH = 1024

/** Thrown past `MAX_DEPTH`; never leaves this module. */
const TOO_DEEP = Symbol("formula too deep")

function addScopedPrefixes(formula: string): string {
  try {
    return scopedPrefixes(formula)
  } catch (e) {
    // Past the cap the formula is not one Excel could have made, so it is
    // written as given rather than failing the whole workbook over a
    // prefix. Anything that is not our own signal is not ours to swallow.
    if (e === TOO_DEEP) return formula
    throw e
  }
}

function scopedPrefixes(formula: string): string {
  const edits: Edit[] = []
  const [top] = parseList(formula, 0, false, 0)
  resolve(formula, top[0], new Map(), edits)
  if (edits.length === 0) return formula
  // Each edit is at a different token, so in source order they never overlap.
  edits.sort((a, b) => a.start - b.start)
  let out = ""
  let last = 0
  for (const e of edits) {
    out += formula.slice(last, e.start) + e.text
    last = e.end
  }
  return out + formula.slice(last)
}

/**
 * Parse from `i` to the `)` that closes the call (`inCall`), or to the end
 * of the formula. Returns the argument lists — one when not in a call —
 * and the index after the `)`. A stray `)` at the top level, or a missing
 * one at the end, is tolerated: this is a writer, not a validator.
 */
function parseList(formula: string, i: number, inCall: boolean, depth: number): [Item[][], number] {
  if (depth > MAX_DEPTH) throw TOO_DEEP
  const n = formula.length
  const args: Item[][] = [[]]
  let current = args[0]
  // Commas inside an array constant `{1,2;3,4}` separate its columns,
  // not arguments: `LET(x,{1,2},x)` still binds one name.
  let braces = 0
  while (i < n) {
    const ch = formula.charCodeAt(i)
    if (ch === 0x22 /* " */ || ch === 0x27 /* ' */) {
      i = skipQuoted(formula, i)
      current.push(OTHER)
    } else if (ch === 0x5b /* [ */) {
      const start = i
      i = skipBracket(formula, i)
      current.push({ kind: "bracket", span: { start, end: Math.min(i, n) } })
    } else if (isNameChar(ch)) {
      const start = i
      while (i < n && isNameChar(formula.charCodeAt(i))) i++
      current.push({ kind: "name", span: { start, end: i } })
    } else if (ch === 0x28 /* ( */) {
      // A name right before the `(` is the function being called. After
      // a `)` it is not — `LAMBDA(x,x*2)(3)` calls the LAMBDA's result —
      // and neither is a bare `(` grouping an expression.
      const prev = current[current.length - 1]
      let callee: Span | undefined
      if (prev?.kind === "name" && prev.span.end === i) {
        current.pop()
        callee = prev.span
      }
      const [sub, next] = parseList(formula, i + 1, true, depth + 1)
      current.push({ kind: "call", callee, args: sub })
      i = next
    } else if (ch === 0x29 /* ) */ && inCall) {
      return [args, i + 1]
    } else if (ch === 0x2c /* , */ && inCall && braces === 0) {
      current = []
      args.push(current)
      i++
    } else {
      if (ch === 0x7b /* { */) braces++
      else if (ch === 0x7d /* } */ && braces > 0) braces--
      if (ch !== 0x20 && ch !== 0x0a && ch !== 0x0d && ch !== 0x09) current.push(OTHER)
      i++
    }
  }
  return [args, n]
}

function resolve(formula: string, items: Item[], scope: Scope, edits: Edit[]): void {
  for (const item of items) {
    if (item.kind === "name") reference(formula, item.span, scope, edits)
    else if (item.kind === "call") call(formula, item, scope, edits)
  }
}

/**
 * Whether the name at `span` stands on its own as a reference, rather than
 * being a sheet (`x!A1`), a table (`T[Col]`), a name qualified by its sheet
 * (`Sheet1!x`) or part of an absolute cell reference (`$A$1`, `A$1`).
 */
function standsAlone(formula: string, span: Span): boolean {
  const before = span.start > 0 ? formula.charCodeAt(span.start - 1) : 0
  const after = formula.charCodeAt(span.end)
  return before !== 0x21 && before !== 0x24 && after !== 0x21 && after !== 0x5b && after !== 0x24
}

/** A use of a name: `_xlpm.` if a LET or LAMBDA around it binds it. */
function reference(formula: string, span: Span, scope: Scope, edits: Edit[]): void {
  if (scope.size === 0 || !standsAlone(formula, span)) return
  const bound = scope.get(formula.slice(span.start, span.end).toUpperCase())
  if (bound !== undefined) edits.push({ ...span, text: XLPM + bound })
}

function call(
  formula: string,
  item: Extract<Item, { kind: "call" }>,
  scope: Scope,
  edits: Edit[],
): void {
  const { callee, args } = item
  let fn = ""
  if (callee !== undefined && standsAlone(formula, callee)) {
    const upper = formula.slice(callee.start, callee.end).toUpperCase()
    const bound = scope.get(upper)
    if (bound !== undefined) {
      // A LET-bound LAMBDA called by its name: `_xlpm.f(2)`.
      edits.push({ ...callee, text: XLPM + bound })
    } else {
      const prefix = FUTURE_FUNCTION_PREFIXES.get(upper)
      if (prefix !== undefined) edits.push({ start: callee.start, end: callee.start, text: prefix })
      fn = upper.startsWith("_XLFN.") ? upper.slice(XLFN.length) : upper
    }
  }
  if (fn === "LET") {
    const inner = new Map(scope)
    let k = 0
    for (; k + 1 < args.length; k += 2) {
      const name = binding(formula, args[k], false, edits)
      if (name === undefined) resolve(formula, args[k], inner, edits)
      // The value is resolved before its name is bound: in
      // `LET(x,x+1,…)` the second `x` is the one from outside.
      resolve(formula, args[k + 1], inner, edits)
      if (name !== undefined) inner.set(name.toUpperCase(), name)
    }
    if (k < args.length) resolve(formula, args[k], inner, edits)
  } else if (fn === "LAMBDA") {
    const inner = new Map(scope)
    const body = args.length - 1
    for (let k = 0; k < body; k++) {
      const name = binding(formula, args[k], true, edits)
      if (name === undefined) resolve(formula, args[k], scope, edits)
      else inner.set(name.toUpperCase(), name)
    }
    resolve(formula, args[body], inner, edits)
  } else {
    const slot = ETA_ARGUMENT.get(fn)
    const eta = slot === -1 ? args.length - 1 : slot
    for (let k = 0; k < args.length; k++) {
      if (k === eta) etaArgument(formula, args[k], scope, edits)
      resolve(formula, args[k], scope, edits)
    }
  }
}

/**
 * `_xleta.` on `arg` when it is a lone built-in function name that no LET
 * or LAMBDA around it binds; see the rules above ETA_FUNCTIONS.
 */
function etaArgument(formula: string, arg: Item[], scope: Scope, edits: Edit[]): void {
  if (arg.length !== 1 || arg[0].kind !== "name") return
  const { span } = arg[0]
  if (!standsAlone(formula, span)) return
  const upper = formula.slice(span.start, span.end).toUpperCase()
  if (scope.has(upper) || !ETA_FUNCTIONS.has(upper)) return
  edits.push({ start: span.start, end: span.start, text: ETA })
}

/**
 * The name an argument in a binding position binds, after queueing the
 * edit that prefixes it — or `undefined` when the argument is not a lone
 * name, which leaves a malformed LET or LAMBDA as it was. `optional` is
 * for LAMBDA, whose `[b]` is stored `_xlop.b`.
 */
function binding(
  formula: string,
  arg: Item[],
  optional: boolean,
  edits: Edit[],
): string | undefined {
  if (arg.length !== 1) return undefined
  const item = arg[0]
  if (item.kind === "name") {
    const text = formula.slice(item.span.start, item.span.end)
    const upper = text.toUpperCase()
    // Already in the stored form, as the reader hands it back.
    if (upper.startsWith("_XLPM.") || (optional && upper.startsWith("_XLOP."))) {
      return text.slice(XLPM.length)
    }
    if (!isParameterName(text)) return undefined
    edits.push({ start: item.span.start, end: item.span.start, text: XLPM })
    return text
  }
  if (item.kind === "bracket" && optional && formula.charCodeAt(item.span.end - 1) === 0x5d) {
    const name = formula.slice(item.span.start + 1, item.span.end - 1).trim()
    if (!isParameterName(name)) return undefined
    edits.push({ ...item.span, text: XLOP + name })
    return name
  }
  return undefined
}

/**
 * Whether `text` can be a LET or LAMBDA name: name characters, not
 * starting with a digit, and no dot — Excel refuses one in a parameter,
 * which is what keeps `_xlpm.x` from being read as a name of its own.
 */
function isParameterName(text: string): boolean {
  if (text === "") return false
  const first = text.charCodeAt(0)
  if (first >= 0x30 && first <= 0x39) return false
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i)
    if (c === 0x2e || !isNameChar(c)) return false
  }
  return true
}
