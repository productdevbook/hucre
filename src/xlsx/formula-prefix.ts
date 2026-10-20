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
// has them. Every name is in one of those two lists: prefixing a name
// Excel stores bare would break a formula that works today.
//
// Left out on purpose:
// - ECMA.CEILING, ISO.CEILING, NETWORKDAYS.INTL, WORKDAY.INTL: the spec
//   lists them as future functions with no prefix.
// - `_xlfn._xlws.PY`: a Python cell is a solo function with metadata of
//   its own, not something a prefix makes work.
// - `_xlfn.SINGLE` / `_xlfn.ANCHORARRAY`: the stored forms of the `@` and
//   `#` operators, not names a caller types.
// - `_xlpm.` on LET / LAMBDA parameter names: that needs to know which
//   names a LET binds, which a scan for `NAME(` cannot.

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
  "PERCENTRANK.EXC",
  "PERCENTRANK.INC",
  "PERMUTATIONA",
  "PHI",
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
 * Add the stored prefix to every call of a future function in `formula`,
 * so `XLOOKUP(` is written as `_xlfn.XLOOKUP(` (#597).
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
  let out = ""
  let last = 0
  let i = 0
  const n = formula.length
  while (i < n) {
    const ch = formula.charCodeAt(i)
    if (ch === 0x22 /* " */ || ch === 0x27 /* ' */) {
      // String literal or quoted sheet name; a doubled quote is an escape.
      i++
      while (i < n) {
        if (formula.charCodeAt(i) === ch) {
          if (formula.charCodeAt(i + 1) === ch) i += 2
          else break
        } else i++
      }
      i++
    } else if (ch === 0x5b /* [ */) {
      // Structured reference (`T[[#This Row],[Col]]`) or external book
      // index (`[1]Sheet1!A1`). Brackets nest, and `'` escapes the next
      // character inside them.
      let depth = 0
      while (i < n) {
        const c = formula.charCodeAt(i)
        if (c === 0x27) i++
        else if (c === 0x5b) depth++
        else if (c === 0x5d && --depth === 0) break
        i++
      }
      i++
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
