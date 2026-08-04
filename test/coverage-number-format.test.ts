import { describe, expect, it } from "vitest"
import { formatValue } from "../src/_format"
import {
  dateToSerial,
  formatDate,
  isDateFormat,
  parseDate,
  serialToDate,
  serialToTime,
  timeToSerial,
} from "../src/_date"
import { bufferReadableStream, readInputToUint8Array } from "../src/_input"

// ── Helpers ──────────────────────────────────────────────────────────

/** A fixed instant with a non-trivial time part: Fri 5 Mar 2021, 14:07:09.250 UTC. */
const MOMENT = new Date(Date.UTC(2021, 2, 5, 14, 7, 9, 250))

/** Serial for {@link MOMENT} in the 1900 system. */
const MOMENT_SERIAL = dateToSerial(MOMENT)

/** Build a byte stream that yields the given chunks, one `read()` at a time. */
function streamOf(...chunks: Uint8Array[]): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(chunk)
      controller.close()
    },
  })
}

// ═══════════════════════════════════════════════════════════════════════
// formatValue — value kinds the format engine has to normalize first
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — input coercion", () => {
  // A `Date` reaching the General branch has no format to follow, so the
  // engine falls back to the unambiguous ISO rendering rather than the
  // host locale's `toString()`.
  it("renders a Date under General as an ISO 8601 timestamp", () => {
    expect(formatValue(new Date(Date.UTC(2021, 0, 15)), "General")).toBe("2021-01-15T00:00:00.000Z")
  })

  // Cells read from a workbook carry Date objects; applying a numeric
  // format to one must go through the serial conversion, not String().
  it("converts a Date to its serial before applying a numeric format", () => {
    expect(formatValue(new Date(Date.UTC(2021, 0, 15)), "0.00")).toBe("44211.00")
  })

  // The 1904 workbook flag shifts the epoch by 1462 days.
  it("honours the 1904 date system when serializing a Date", () => {
    const jan15 = new Date(Date.UTC(2021, 0, 15))
    expect(formatValue(jan15, "0", { is1904: true })).toBe("42749")
    expect(formatValue(jan15, "0")).toBe("44211")
  })

  // Anything that is neither string, number, boolean nor Date has no
  // meaningful numeric form — the engine stringifies it instead of
  // producing NaN.
  it("stringifies values that are not numbers, strings, booleans or Dates", () => {
    expect(formatValue(10n, "0.00")).toBe("10")
    expect(formatValue({ toString: () => "obj" }, "0.00")).toBe("obj")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Section selection — positive; negative; zero; text
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — section selection", () => {
  // The Excel accounting-style format: the third section renders zero and
  // the second renders the magnitude of negatives (sign supplied by the
  // parentheses, not by the number).
  const ACCOUNTING = '$#,##0.00;[Red]($#,##0.00);"-"'

  it("uses the first section for positive values", () => {
    expect(formatValue(1234.5, ACCOUNTING)).toBe("$1,234.50")
  })

  it("uses the second section — on the absolute value — for negatives", () => {
    expect(formatValue(-1234.5, ACCOUNTING)).toBe("($1,234.50)")
  })

  it("uses the third section for exactly zero", () => {
    expect(formatValue(0, ACCOUNTING)).toBe("-")
  })

  // An empty section suppresses the value entirely — that is how Excel's
  // "hide positives"/"hide negatives" formats work.
  it("renders nothing when the matching section is empty", () => {
    expect(formatValue(0, '#,##0.00;;"zero"')).toBe("zero")
    expect(formatValue(-5, "0.0;;")).toBe("")
  })

  // Sections are split on *unquoted* semicolons only; a semicolon that is
  // escaped or quoted belongs to the literal text.
  it("does not split on a backslash-escaped semicolon", () => {
    expect(formatValue(5, "0.00\\;x")).toBe("5.00;x")
  })

  it("does not split on a quoted semicolon", () => {
    expect(formatValue(5, '0.00";"')).toBe("5.00;")
  })

  // A format ending in a lone backslash has nothing to escape — the
  // scanner must stop rather than read past the end of the string.
  it("tolerates a trailing backslash with nothing to escape", () => {
    expect(formatValue(5, "0.00\\")).toBe("5.00")
  })

  // Strings take the fourth section when one exists, and are passed
  // through untouched when the format has fewer sections.
  it("routes strings through the fourth (text) section", () => {
    expect(formatValue("hi", '0.00;-0.00;"z";"["@"]"')).toBe("[hi]")
  })

  it("returns a string unchanged when the format has no text section", () => {
    expect(formatValue("hi", "0.00")).toBe("hi")
  })

  // Backslash escapes inside a text section are literal characters, and a
  // section with no @ placeholder discards its literals entirely.
  it("expands backslash escapes around the @ placeholder", () => {
    expect(formatValue("hi", "\\@@\\!")).toBe("hihi!")
  })

  it("tolerates a text section ending in a bare backslash", () => {
    expect(formatValue("hi", "@\\")).toBe("hi")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Conditional sections — [>100], [<=0], …
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — conditions", () => {
  it("applies the first section when its condition matches", () => {
    expect(formatValue(150, '[>100]"big";[<=100]"small"')).toBe("big")
  })

  it("falls through to the second section when the condition fails", () => {
    expect(formatValue(50, '[>100]"big";"small"')).toBe("small")
  })

  // A single conditional section always applies: Excel has no other
  // section to fall back to, so the condition only strips itself.
  it("strips a condition from a lone section regardless of the outcome", () => {
    expect(formatValue(5, "[>0]0.00")).toBe("5.00")
    expect(formatValue(5, "[<0]0.00")).toBe("5.00")
  })

  it.each([
    [">", '[>5]"gt";"le"', "le"],
    ["<", '[<5]"lt";"ge"', "ge"],
    [">=", '[>=5]"ge";"lt"', "ge"],
    ["<=", '[<=5]"le";"gt"', "le"],
    ["=", '[=5]"eq";"ne"', "eq"],
    ["<>", '[<>5]"ne";"eq"', "eq"],
  ])("evaluates the %s comparison operator", (_op, fmt, expected) => {
    expect(formatValue(5, fmt)).toBe(expected)
  })

  // `==` and `!=` are not Excel syntax, but the condition scanner accepts
  // any run of comparison characters, so they behave as their Excel
  // spellings rather than silently matching everything.
  it("treats == and != as aliases of = and <>", () => {
    expect(formatValue(5, '[==5]"eq";"ne"')).toBe("eq")
    expect(formatValue(5, '[!=5]"ne";"eq"')).toBe("eq")
  })

  // An unrecognised operator must not throw or drop the value — it
  // degrades to "always true" so the first section still renders.
  it("treats an unrecognised operator as always matching", () => {
    expect(formatValue(5, '[!5]"first";"second"')).toBe("first")
  })

  // Conditions compare against the numeric value, so a Date cell is
  // compared on its serial.
  it("compares a Date against the condition using its serial number", () => {
    const jan15 = new Date(Date.UTC(2021, 0, 15)) // serial 44211
    expect(formatValue(jan15, '[>44000]"future";"past"')).toBe("future")
    expect(formatValue(jan15, '[>50000]"future";"past"')).toBe("past")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Integer / decimal placeholders
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — digit placeholders", () => {
  // `0` demands a digit, `#` does not — so an all-# format renders zero
  // as an empty cell, which is the classic "hide zeros" trick.
  it("renders zero as empty under an all-# format", () => {
    expect(formatValue(0, "#")).toBe("")
  })

  // …but the decimal point of "#.##" survives, exactly as Excel shows it.
  it("keeps the decimal point of #.## when the value is zero", () => {
    expect(formatValue(0, "#.##")).toBe(".")
  })

  it("pads with leading zeros up to the count of 0 placeholders", () => {
    expect(formatValue(123, "000000")).toBe("000123")
  })

  // A format with no `0` at all still has to group correctly.
  it("groups a format built only from # placeholders", () => {
    expect(formatValue(1234, "#,###")).toBe("1,234")
  })

  // `#` after the decimal point suppresses insignificant trailing digits;
  // `0` keeps them.
  it("suppresses trailing zeros behind a # decimal placeholder", () => {
    expect(formatValue(5, "0.0#")).toBe("5.0")
    expect(formatValue(1.05, "0.0#")).toBe("1.05")
    expect(formatValue(1.5, "0.00#")).toBe("1.50")
  })

  it("adds group separators to a bare negative number", () => {
    expect(formatValue(-1234.5, "#,##0")).toBe("-1,235")
  })

  // A format made only of literal text has no placeholders at all; the
  // engine returns the literals and drops the number.
  it("returns only the literal text when the section has no placeholders", () => {
    expect(formatValue(42, '"N/A"')).toBe("N/A")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Literals, currency, colours and locale prefixes
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — literals and bracketed directives", () => {
  it("keeps a backslash-escaped currency prefix", () => {
    expect(formatValue(42, "\\$0.00")).toBe("$42.00")
  })

  it("keeps quoted text before and after the number", () => {
    expect(formatValue(42, '"USD "0.00')).toBe("USD 42.00")
    expect(formatValue(42, '0.00" units"')).toBe("42.00 units")
    expect(formatValue(1234.5, '#,##0.00 "TL"')).toBe("1,234.50 TL")
  })

  it("keeps a backslash-escaped suffix character", () => {
    expect(formatValue(42, "0.00\\!")).toBe("42.00!")
  })

  // Colour directives, locale/currency prefixes and `_x` alignment
  // padding are display hints — they must not leak into the text.
  it.each([
    ["[Red]#,##0.00", "a named colour"],
    ["[Color 3]#,##0.00", "an indexed colour"],
    ["[$$-409]#,##0.00", "a locale-qualified currency prefix"],
    ["_(#,##0.00_)", "alignment padding"],
  ])("strips %s", (fmt) => {
    expect(formatValue(1234.5, fmt)).toBe("1,234.50")
  })

  // Excel's built-in accounting format (numFmtId 39/40) wraps negatives
  // in parentheses and pads positives to match.
  it("renders the built-in accounting format's parenthesised negative", () => {
    expect(formatValue(-1234.5, "#,##0.00_);(#,##0.00)")).toBe("(1,234.50)")
    expect(formatValue(1234.5, "#,##0.00_);(#,##0.00)")).toBe("1,234.50")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Percentages, scaling and scientific notation
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — percentage", () => {
  it("multiplies by 100 and appends the sign", () => {
    expect(formatValue(0.1234, "0.00%")).toBe("12.34%")
    expect(formatValue(0.1234, "0%")).toBe("12%")
  })
})

describe("formatValue — thousands scaling", () => {
  // Two trailing commas divide by a million — the "in millions" format
  // used on financial summaries.
  it("divides by a million for a double trailing comma", () => {
    expect(formatValue(123456789, '#,##0.00,," M"')).toBe("123.46 M")
    expect(formatValue(1500, '0.0,,"M"')).toBe("0.0M")
  })
})

describe("formatValue — scientific notation", () => {
  it("renders the mantissa with the declared decimals and a signed exponent", () => {
    expect(formatValue(12345.6789, "0.000E+00")).toBe("1.235E+04")
    expect(formatValue(12345678, "0.0E+00")).toBe("1.2E+07")
    expect(formatValue(1, "0.0E+00")).toBe("1.0E+00")
  })

  it("renders a negative exponent for values below one", () => {
    expect(formatValue(0.000123, "0.0E+00")).toBe("1.2E-04")
    expect(formatValue(0.00012, "0.00E-00")).toBe("1.20E-04")
  })

  // `E-00` only shows the sign when the exponent is negative, unlike `E+00`.
  it("hides the + sign when the format asks for E-", () => {
    expect(formatValue(12345.6789, "0.00E-00")).toBe("1.23E04")
  })

  it("preserves the case of the exponent marker", () => {
    expect(formatValue(12345.6789, "0.00e+00")).toBe("1.23e+04")
  })

  // A mantissa with no decimal point renders the exponent alone.
  it("renders a mantissa with no decimals", () => {
    expect(formatValue(12345.6789, "0E+00")).toBe("1E+04")
  })

  // A prefix in front of the mantissa defeats the strict pattern match,
  // so the engine falls back to deriving the decimals from the mantissa.
  it("falls back to mantissa-derived decimals when the pattern is unusual", () => {
    expect(formatValue(12345.6789, "$0.00E+00")).toBe("1.23E+04")
    expect(formatValue(12345.6789, "0.00E00")).toBe("1.23E04")
  })

  // …and to two decimals when the fallback cannot find a mantissa either.
  it("defaults to two decimals when the fallback finds no mantissa", () => {
    expect(formatValue(12345.6789, "$0E+00")).toBe("1.23E+04")
  })

  it("keeps the exponent marker case and sign rules on the fallback path", () => {
    expect(formatValue(12345.6789, "$0.00e+00")).toBe("1.23e+04")
    expect(formatValue(0.00012, "$0.00E+00")).toBe("1.20E-04")
    expect(formatValue(0.00012, "$0.00E00")).toBe("1.20E-04")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Fractions
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — fractions", () => {
  // Built-in numFmtId 12 ("# ?/?") searches for the best denominator with
  // a single digit; numFmtId 13 ("# ??/??") allows two.
  it("finds the closest single-digit denominator", () => {
    expect(formatValue(0.5, "# ?/?")).toBe("1/2")
  })

  it("keeps the whole part beside a two-digit fraction", () => {
    expect(formatValue(2.25, "# ??/??")).toBe("2  1/ 4")
  })

  // With no fractional remainder Excel shows the integer alone.
  it("shows only the integer when the value is whole", () => {
    expect(formatValue(3, "# ?/?")).toBe("3")
  })

  // Zero has no integer to show either, so the fraction area is blanked.
  it("pads zero out to the width of the fraction area", () => {
    expect(formatValue(0, "# ?/?")).toBe("0      ")
  })

  it("carries the minus sign of a negative mixed number", () => {
    expect(formatValue(-2.5, "# ?/?")).toBe("-2 1/2")
  })

  // Three `?` in the denominator widen the search to 999.
  it("widens the denominator search as the format grows", () => {
    expect(formatValue(0.125, "# ?/???")).toBe("1/  8")
  })

  // `?` is an integer placeholder too, so a whole number still renders
  // without a fraction under "? ?/?".
  it("shows a whole number under a ?-only integer placeholder", () => {
    expect(formatValue(3, "? ?/?")).toBe("3")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Excel parity — the defects reported in #388
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — Excel parity", () => {
  // A negative section is handed the absolute value, since the section
  // itself supplies the presentation of the sign — so the "-" written in
  // the format has to survive as literal text. `extractLiterals` therefore
  // keeps "+" and "-" out of `core`, where no placeholder would render them.
  it("keeps the explicit minus sign of a negative section", () => {
    expect(formatValue(-42, "0.00;-0.00")).toBe("-42.00")
    expect(formatValue(-1234.5, "#,##0.00;-#,##0.00")).toBe("-1,234.50")
  })

  // A comma after the last digit placeholder scales the value by 1000; a
  // comma between placeholders turns on group separators. The two are
  // independent, so "#,##0," does both while "0," only scales.
  it("divides by a thousand for a single trailing comma", () => {
    expect(formatValue(1234567, "#,##0,")).toBe("1,235")
    expect(formatValue(1234567, "0,")).toBe("1235")
  })

  it("keeps group separators on a scaled-down value", () => {
    expect(formatValue(1234567890, "#,##0,,")).toBe("1,235")
  })

  // Excel's built-in fixed-denominator fractions: "As halves" (`# ?/2`),
  // "As eighths" (`# ?/8`), "As sixteenths" (`# ??/16`). The numerator is
  // rounded onto the denominator the format names rather than searched for.
  it("honours a fixed denominator written in the format", () => {
    expect(formatValue(3.5, "# ?/2")).toBe("3 1/2")
    expect(formatValue(3.5, "# ?/8")).toBe("3 4/8")
  })

  // Math.trunc(-0.5) is -0, so the sign cannot be read off the whole part
  // — it has to come from the value itself.
  it("keeps the sign of a negative value smaller than one", () => {
    expect(formatValue(-0.5, "# ?/?")).toBe("-1/2")
  })

  // A "?" in front of the fraction is a whole-number placeholder just like
  // "#" and "0", so "? ?/?" is a mixed number.
  it("keeps the whole part in front of a ?-placeholder fraction", () => {
    expect(formatValue(2.5, "? ?/?")).toBe("2 1/2")
  })

  // ── #397 ──────────────────────────────────────────────────────────
  //
  // With no placeholder at all ahead of the slash there is nowhere to put
  // the whole part, so Excel folds it into the numerator: the fraction is
  // improper. Formatting the remainder alone dropped it — 2.5 read back as
  // one half.
  it("folds the whole part into the numerator when the format has no integer slot", () => {
    expect(formatValue(2.5, "??/??")).toBe(" 5/ 2")
    expect(formatValue(2.5, "??/2")).toBe(" 5/2")
  })

  it("keeps the sign of an improper fraction", () => {
    expect(formatValue(-2.5, "??/??")).toBe("- 5/ 2")
  })

  // A whole number under an improper format still has to render as a
  // fraction — "??/??" has no other place to show it.
  it("gives a whole number a denominator of one under an improper format", () => {
    expect(formatValue(3, "??/??")).toBe(" 3/ 1")
  })

  // A denominator the format fixes can round the remainder away: 0.1 over
  // halves is 0. Excel renders the whole part rather than a zero numerator
  // — it never shows "3 0/2". The whole-number path is shared, so 3.1 and
  // 3 print alike under "# ?/2".
  it("drops the fraction area when the remainder rounds to a zero numerator", () => {
    expect(formatValue(3.1, "# ?/2")).toBe("3")
    expect(formatValue(3.1, "# ?/2")).toBe(formatValue(3, "# ?/2"))
    expect(formatValue(-3.1, "# ?/2")).toBe("-3")
  })

  // The remainder only rounds away against a coarse denominator; a search
  // that can reach the value still renders a fraction.
  it("still renders a fraction when the denominator can represent the remainder", () => {
    expect(formatValue(3.1, "# ??/??")).toBe("3  1/10")
  })

  // ── #402 ──────────────────────────────────────────────────────────
  //
  // One placeholder either side of the slash is a fraction. Recognition
  // used to demand two before it, so "?/?" was formatted as a plain number
  // — which renders the "/" as a literal and puts the digits in the one
  // placeholder it can see, losing the numerator entirely: 2.5 came out
  // " /3".
  it("reads a single placeholder either side of the slash as a fraction", () => {
    expect(formatValue(2.5, "?/?")).toBe("5/2")
    expect(formatValue(0.5, "?/?")).toBe("1/2")
    expect(formatValue(0.125, "?/?")).toBe("1/8")
  })

  // Padding follows the width of the format, so the one-character form has
  // nothing to pad — "?/?" gives "5/2" where "??/??" gives " 5/ 2".
  it("pads a one-character fraction to its own width, not the two-character one", () => {
    expect(formatValue(2.5, "?/?")).toBe("5/2")
    expect(formatValue(2.5, "??/??")).toBe(" 5/ 2")
  })

  // No placeholder ahead of the slash, so the whole part folds into the
  // numerator exactly as it does under "??/??".
  it("folds the whole part into a single-placeholder numerator", () => {
    expect(formatValue(3, "?/?")).toBe("3/1")
    expect(formatValue(-2.5, "?/?")).toBe("-5/2")
    expect(formatValue(-0.5, "?/?")).toBe("-1/2")
  })

  // A one-character denominator holds a one-digit search, the same rule the
  // wider formats follow: "?/?" searches to 9, "?/???" to 999.
  it("keeps the denominator search at one digit for a one-character denominator", () => {
    expect(formatValue(3.1, "?/?")).toBe("28/9")
    expect(formatValue(3.1, "??/??")).toBe("31/10")
  })

  // A literal denominator still reads as literal when the numerator is one
  // character: "0/2" is halves, not a one-digit search.
  it("honours a fixed denominator behind a single-placeholder numerator", () => {
    expect(formatValue(0.5, "0/2")).toBe("1/2")
    expect(formatValue(2.5, "0/2")).toBe("5/2")
    expect(formatValue(0.125, "?/8")).toBe("1/8")
    expect(formatValue(2.5, "?/8")).toBe("20/8")
  })

  // "#", "0" and "?" are all digit placeholders, so these are the same
  // format as "?/?" wearing different padding rules — none of which show
  // for a numerator that fills its width anyway.
  it("treats #, 0 and ? alike on either side of the slash", () => {
    expect(formatValue(2.5, "#/#")).toBe("5/2")
    expect(formatValue(2.5, "0/0")).toBe("5/2")
    expect(formatValue(2.5, "0/0")).toBe(formatValue(2.5, "?/?"))
  })

  // The whole-part detection from #397 reads the text in front of the
  // fraction, so a one-character fraction group changes nothing about it.
  it("still finds the whole part in front of a one-character fraction", () => {
    expect(formatValue(2.5, "# ?/?")).toBe("2 1/2")
    expect(formatValue(3, "# ?/?")).toBe("3")
    expect(formatValue(3.1, "# ?/2")).toBe("3")
  })

  // The widened pattern must not drag non-fractions in with it. A date
  // format is claimed before fractions are considered, and neither it nor a
  // plain number puts a digit placeholder against a slash.
  it("leaves formats that are not fractions alone", () => {
    expect(formatValue(45000, "m/d/yyyy")).toBe("3/15/2023")
    expect(formatValue(45000, "yyyy/mm/dd")).toBe("2023/03/15")
    expect(formatValue(3.7, "0")).toBe("4")
    expect(formatValue(3.7, "0.00")).toBe("3.70")
    expect(formatValue(1234.5, "#,##0.00")).toBe("1,234.50")
    expect(formatValue(0.5, "0%")).toBe("50%")
  })

  // ── #426 ──────────────────────────────────────────────────────────
  //
  // The fraction path built its output from the digits alone, so literal
  // text was dropped where every other path keeps it — "$?/?" rendered as
  // "5/2". Literals are read with the same `extractLiterals` the plain
  // number path uses, so a bare "$", a quoted run and a "\$" escape mean
  // the same thing under a fraction as under "0.00".
  it("keeps literal text around a fraction", () => {
    expect(formatValue(2.5, "$?/?")).toBe("$5/2")
    expect(formatValue(2.5, '"USD "?/?')).toBe("USD 5/2")
    expect(formatValue(2.5, '?/?" kg"')).toBe("5/2 kg")
    expect(formatValue(2.5, "\\$# ?/?")).toBe("$2 1/2")
    expect(formatValue(2.5, "# ?/?\\!")).toBe("2 1/2!")
  })

  // The padding belongs to the fraction, not to the field, so a prefix
  // lands outside it: "$ 5/ 2", never " $5/ 2".
  it("puts a literal prefix outside the placeholder padding", () => {
    expect(formatValue(2.5, "$??/??")).toBe("$ 5/ 2")
    expect(formatValue(2.5, '# ??/??" in"')).toBe("2  1/ 2 in")
  })

  // The prefix leads the sign, as it does for "$#,##0.00" — the two paths
  // agree on where the minus goes.
  it("writes a literal prefix ahead of the sign", () => {
    expect(formatValue(-2.5, "$# ?/?")).toBe("$-2 1/2")
    expect(formatValue(-2.5, "$#,##0.00")).toBe("$-2.50")
  })

  // Whole numbers and zero take an earlier exit out of the formatter; the
  // literals have to survive that route too.
  it("keeps literals on the paths that print no fraction", () => {
    expect(formatValue(3, "$# ?/?")).toBe("$3")
    expect(formatValue(3.1, "$# ?/2")).toBe("$3")
    expect(formatValue(0, "$# ?/?")).toBe("$0      ")
  })

  // A negative section is handed the absolute value, so its "-" is literal
  // text in the format — dropping literals dropped the sign with them, and
  // -2.5 read back as positive.
  it("keeps the sign a negative fraction section writes for itself", () => {
    expect(formatValue(-2.5, "# ?/?;-# ?/?")).toBe("-2 1/2")
    expect(formatValue(-2.5, "$# ?/?;($# ?/?)")).toBe("($2 1/2)")
  })

  // Sections are split before any of this, so one section's literals cannot
  // reach another's output.
  it("keeps each section's literals to itself", () => {
    expect(formatValue(2.5, '"a"# ?/?;"b"# ?/?')).toBe("a2 1/2")
    expect(formatValue(-2.5, '"a"# ?/?;"b"# ?/?')).toBe("b2 1/2")
  })

  // What separates the whole part from the numerator is the format's own
  // literal text — the space in "# ?/?" is not punctuation the formatter
  // owns, and a quoted placeholder character is text rather than an
  // integer slot, so '"#"?/?' is an improper fraction.
  it("takes the whole-part separator from the format", () => {
    expect(formatValue(2.5, '#" and "?/?')).toBe("2 and 1/2")
    expect(formatValue(2.5, '"#"?/?')).toBe("#5/2")
  })

  // ── #429 ──────────────────────────────────────────────────────────
  //
  // Detection scanned the raw format for placeholders around a slash, so a
  // *literal* containing both was read as a fraction spec and the whole
  // format was reinterpreted. A regression from #402: unifying detection
  // with the parse regex inherited the parse regex's blind spot, where the
  // old pattern had excluded this case by accident.

  it("does not read a quoted literal as a fraction spec", () => {
    expect(formatValue(3.5, '0.00" 0/2"')).toBe("3.50 0/2")
    expect(formatValue(3.5, '#,##0.00" 0/2"')).toBe("3.50 0/2")
  })

  it("does not read an escaped slash as a fraction bar", () => {
    expect(formatValue(3.5, "0.00\\/2")).toBe("3.50/2")
  })

  it("still leaves the formats that never tripped it alone", () => {
    // "1" is not a placeholder, so these were unaffected either way —
    // pinned so the mask cannot start eating them.
    expect(formatValue(3.5, '0.00" 1/2"')).toBe("3.50 1/2")
    expect(formatValue(3.5, '0" m/s"')).toBe("4 m/s")
    expect(formatValue(3.5, '#,##0.00" km/h"')).toBe("3.50 km/h")
  })

  it("still recognises a real fraction that also carries literals", () => {
    // The mask must not blind detection to the fraction itself.
    expect(formatValue(2.5, '?/?" kg"')).toBe("5/2 kg")
    expect(formatValue(2.5, '"about "# ?/?')).toBe("about 2 1/2")
  })

  // maskLiterals and extractLiterals are both private, so their agreement
  // is checked through behaviour: each literal shape has to mean the same
  // thing under a number format and under a fraction format. Two
  // implementations of one concept is what produced #429.
  it("means the same thing by 'literal' on both paths", () => {
    const shapes: Array<[string, string, string]> = [
      // literal, number format, fraction format
      ['"x"', '"x"0.00', '"x"# ?/?'],
      ["\\x", "\\x0.00", "\\x# ?/?"],
      ['" 0/2"', '0.00" 0/2"', '# ?/?" 0/2"'],
    ]
    for (const [, numFmt, fracFmt] of shapes) {
      const asNumber = formatValue(2.5, numFmt)
      const asFraction = formatValue(2.5, fracFmt)
      // Same leading literal, same trailing literal — only the middle,
      // which is the actual number, differs.
      expect(asNumber.startsWith("x") ? asFraction.startsWith("x") : true).toBe(true)
      expect(asNumber.endsWith(" 0/2") ? asFraction.endsWith(" 0/2") : true).toBe(true)
    }
  })

  it("survives an unterminated quote without swallowing the fraction", () => {
    // extractLiterals consumes an unterminated quote to the end of the
    // format; the mask has to agree, or the two disagree about where the
    // literal stops.
    expect(formatValue(2.5, '# ?/?"trailing')).toBe("2 1/2trailing")
  })

  // "?" renders insignificant decimals as spaces so decimal points line up
  // down a column, where "0" would render them as zeros.
  it("pads insignificant decimals with spaces for a ? placeholder", () => {
    expect(formatValue(0.5, "0.???")).toBe("0.5  ")
  })

  // Excel's Special formats (SSN, phone number) put literal characters
  // *between* digit placeholders. Placeholders are filled right-to-left and
  // the literals stay where the format put them.
  it("keeps separators written between digit groups", () => {
    expect(formatValue(123456789, "000-00-0000")).toBe("123-45-6789")
    expect(formatValue(5551234567, "(000) 000-0000")).toBe("(555) 123-4567")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// isDateFormat — telling dates from numbers
// ═══════════════════════════════════════════════════════════════════════

describe("isDateFormat", () => {
  it("returns false for an empty format", () => {
    expect(isDateFormat("")).toBe(false)
  })

  // Built-in numeric IDs from ECMA-376 §18.8.30 arrive as strings from
  // the styles part.
  it.each([
    ["14", true], // m/d/yyyy
    ["45", true], // mm:ss
    ["0", false], // integer
    ["9", false], // 0%
  ])("classifies the built-in format id %s", (id, expected) => {
    expect(isDateFormat(id)).toBe(expected)
  })

  it.each([
    ["General", false],
    ["@", false],
    ["yyyy", true],
    ["d/m", true],
    ["hh:mm", true],
    ["mmm", true],
    ["mmmm", true],
    ["AM/PM", true],
    ["a/p", true],
    ["[h]", true],
    ["[s]", true],
  ])("classifies the format string %s", (fmt, expected) => {
    expect(isDateFormat(fmt)).toBe(expected)
  })

  // A lone "m"/"mm" is genuinely ambiguous (month or minute) and Excel
  // only reads it as a date when a d/y/h/s token is present, so the
  // reader treats it as a number.
  it("does not treat a lone m or mm as a date", () => {
    expect(isDateFormat("m")).toBe(false)
    expect(isDateFormat("mm")).toBe(false)
  })

  // Literal text must not be mistaken for time tokens: the "s" of
  // "shares" and an escaped "\s" would both otherwise look like seconds.
  it("ignores letters that live inside quoted or escaped literals", () => {
    expect(isDateFormat('0.00 "shares"')).toBe(false)
    expect(isDateFormat('#,##0" units"')).toBe(false)
    expect(isDateFormat("0.0 \\s")).toBe(false)
  })

  // Unquoted trailing words are common in hand-written formats. "hrs"
  // contains both an h and an s but neither stands alone as a token, so
  // the format is still a number.
  it("does not treat a word containing h or s as a time token", () => {
    expect(isDateFormat("0.0 hrs")).toBe(false)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// formatDate — token rendering
// ═══════════════════════════════════════════════════════════════════════

describe("formatDate — tokens", () => {
  it("renders single-letter day, month and second tokens unpadded", () => {
    expect(formatDate(MOMENT, "m")).toBe("3")
    expect(formatDate(MOMENT, "d")).toBe("5")
    expect(formatDate(MOMENT, "s")).toBe("9")
  })

  it("renders day and month names", () => {
    expect(formatDate(MOMENT, "dddd")).toBe("Friday")
    expect(formatDate(MOMENT, "ddd")).toBe("Fri")
    expect(formatDate(MOMENT, "mmm d, yyyy")).toBe("Mar 5, 2021")
    expect(formatDate(MOMENT, "yy")).toBe("21")
  })

  // Fractional-second tokens read the millisecond field, truncated to the
  // number of zeros written.
  it("renders fractional seconds at the requested precision", () => {
    expect(formatDate(MOMENT, "yyyy-mm-dd hh:mm:ss.000")).toBe("2021-03-05 14:07:09.250")
    expect(formatDate(MOMENT, "h:mm:ss.00")).toBe("14:07:09.25")
    expect(formatDate(MOMENT, "mm:ss.0")).toBe("07:09.2")
  })

  // A "." that is not followed by a zero is an ordinary separator — this
  // is what makes the German "d.m.yyyy" date work.
  it("treats a dot that is not a fractional-second token as a separator", () => {
    expect(formatDate(MOMENT, "d.m.yyyy")).toBe("5.3.2021")
    expect(formatDate(MOMENT, "hh:mm:ss.")).toBe("14:07:09.")
  })

  it("switches to a 12-hour clock when AM/PM is present", () => {
    expect(formatDate(MOMENT, "h:mm:ss.0 AM/PM")).toBe("2:07:09.2 PM")
    expect(formatDate(MOMENT, "h:mm A/P")).toBe("2:07 P")
  })

  it("stays on the 24-hour clock without an AM/PM token", () => {
    expect(formatDate(MOMENT, "hh:mm:ss")).toBe("14:07:09")
  })

  // Elapsed-time brackets accumulate past their usual range and are
  // computed from the serial, not from the Date's clock fields.
  it("accumulates elapsed time from the serial number", () => {
    expect(formatDate(MOMENT, "[h]:mm:ss", 1.5)).toBe("36:07:09")
    expect(formatDate(MOMENT, "[mm]:ss", 1.5)).toBe("2160:09")
    expect(formatDate(MOMENT, "[ss]", 1.5)).toBe("129600")
  })

  // Without a serial (and for negative serials) elapsed totals clamp to 0
  // rather than rendering NaN.
  it("clamps elapsed totals to zero when no serial is supplied", () => {
    expect(formatDate(MOMENT, "[h]:mm")).toBe("0:07")
    expect(formatDate(MOMENT, "[h]:mm", -3)).toBe("0:07")
  })

  it("emits quoted and backslash-escaped text literally", () => {
    expect(formatDate(MOMENT, '"Year "yyyy')).toBe("Year 2021")
    expect(formatDate(MOMENT, "\\Qyyyy")).toBe("Q2021")
  })

  // Colour and locale directives are display hints; the tokenizer drops
  // them so they cannot end up in the rendered text.
  it("drops colour and locale directives", () => {
    expect(formatDate(MOMENT, "[Red]hh:mm")).toBe("14:07")
    expect(formatDate(MOMENT, "[$-409]d mmmm yyyy")).toBe("5 March 2021")
  })

  // "mm" is minutes after an hour token and a month everywhere else.
  it("disambiguates mm between month and minute by its neighbours", () => {
    expect(formatDate(MOMENT, "mm/dd")).toBe("03/05")
    expect(formatDate(MOMENT, "hh:mm")).toBe("14:07")
    expect(formatDate(MOMENT, "mm:ss")).toBe("07:09")
  })

  it("passes unknown letters through as literal text", () => {
    expect(formatDate(MOMENT, "nn")).toBe("nn")
  })

  // A truncated file can leave a bracket directive unterminated; the
  // tokenizer must fall back to treating it as literal text instead of
  // consuming the rest of the format.
  it("treats an unterminated bracket directive as literal text", () => {
    expect(formatDate(MOMENT, "[$-409")).toBe("[$-409")
    expect(formatDate(MOMENT, "[h")).toBe("[14")
  })

  it("tolerates a format ending in a bare backslash", () => {
    expect(formatDate(MOMENT, "yyyy\\")).toBe("2021")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Serial ↔ Date conversion
// ═══════════════════════════════════════════════════════════════════════

describe("serial conversion", () => {
  // Serial 60 is the phantom "29 Feb 1900" the 1900 system inherited from
  // Lotus 1-2-3; it has no real timestamp, so it maps to 28 Feb.
  it("maps the Lotus phantom serial 60 onto 28 Feb 1900", () => {
    expect(serialToDate(60).toISOString()).toBe("1900-02-28T00:00:00.000Z")
  })

  it("maps serial 0 to the Excel 'Jan 0, 1900' placeholder", () => {
    expect(serialToDate(0).toISOString()).toBe("1899-12-31T00:00:00.000Z")
  })

  it("skips the phantom day for serials above 60", () => {
    expect(serialToDate(61).toISOString()).toBe("1900-03-01T00:00:00.000Z")
  })

  it("uses 1 Jan 1904 as day zero in the 1904 system", () => {
    expect(serialToDate(0, true).toISOString()).toBe("1904-01-01T00:00:00.000Z")
  })

  it("round-trips a moment through the serial and back", () => {
    expect(serialToDate(MOMENT_SERIAL).getTime()).toBe(MOMENT.getTime())
  })

  it("round-trips through the 1904 system too", () => {
    const s = dateToSerial(MOMENT, true)
    expect(serialToDate(s, true).getTime()).toBe(MOMENT.getTime())
  })

  it("splits a serial fraction into clock components", () => {
    expect(serialToTime(0.5)).toEqual({ hours: 12, minutes: 0, seconds: 0, milliseconds: 0 })
    expect(serialToTime(-1.75).hours).toBe(18)
  })

  it("builds a serial fraction from clock components", () => {
    expect(timeToSerial(12, 0)).toBe(0.5)
    expect(timeToSerial(0, 0, 0, 86_400_000 / 2)).toBe(0.5)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// parseDate
// ═══════════════════════════════════════════════════════════════════════

describe("parseDate", () => {
  it("returns null for blank or unparseable input", () => {
    expect(parseDate("")).toBeNull()
    expect(parseDate("   ")).toBeNull()
    expect(parseDate("not a date")).toBeNull()
  })

  it("applies an ISO 8601 UTC offset", () => {
    expect(parseDate("2021-01-15T14:30:00+05:00")?.toISOString()).toBe("2021-01-15T09:30:00.000Z")
    expect(parseDate("2021-01-15T14:30:00-05:00")?.toISOString()).toBe("2021-01-15T19:30:00.000Z")
  })

  // RFC 3339 allows the offset without its colon; Excel exports written
  // by non-Microsoft tools use that spelling.
  it("accepts an offset written without a colon", () => {
    expect(parseDate("2021-01-15T14:30:00-0530")?.toISOString()).toBe("2021-01-15T20:00:00.000Z")
  })

  it("pads a short fractional-second field to milliseconds", () => {
    expect(parseDate("2021-01-15T14:30:00.5Z")?.toISOString()).toBe("2021-01-15T14:30:00.500Z")
  })

  it("parses the US, EU and dashed day-first spellings", () => {
    expect(parseDate("1/15/2021")?.toISOString()).toBe("2021-01-15T00:00:00.000Z")
    expect(parseDate("15.01.2021")?.toISOString()).toBe("2021-01-15T00:00:00.000Z")
    expect(parseDate("15-01-2021")?.toISOString()).toBe("2021-01-15T00:00:00.000Z")
  })

  // Out-of-range components are rejected rather than silently rolled over
  // by the Date constructor.
  it("rejects out-of-range month and day components", () => {
    expect(parseDate("13/45/2021")).toBeNull()
    expect(parseDate("45.13.2021")).toBeNull()
    expect(parseDate("31-13-2021")).toBeNull()
  })
})

// ═══════════════════════════════════════════════════════════════════════
// ReadInput normalization
// ═══════════════════════════════════════════════════════════════════════

describe("readInputToUint8Array", () => {
  it("returns a Uint8Array untouched", async () => {
    const bytes = new Uint8Array([1, 2, 3])
    expect(await readInputToUint8Array(bytes)).toBe(bytes)
  })

  it("wraps an ArrayBuffer without copying its contents", async () => {
    const buf = new Uint8Array([4, 5, 6]).buffer
    expect(Array.from(await readInputToUint8Array(buf))).toEqual([4, 5, 6])
  })

  it("rejects input shapes that are neither bytes nor a stream", async () => {
    await expect(readInputToUint8Array("nope" as never)).rejects.toThrow(/Unsupported input type/)
  })
})

describe("bufferReadableStream", () => {
  // A closed-without-data stream is what an empty file yields; it has to
  // produce an empty buffer rather than throwing on the concat path.
  it("returns an empty buffer for a stream that yields nothing", async () => {
    expect((await bufferReadableStream(streamOf())).length).toBe(0)
  })

  it("returns the single chunk as-is when the stream yields one", async () => {
    const only = new Uint8Array([7, 8])
    expect(await bufferReadableStream(streamOf(only))).toBe(only)
  })

  // Some stream sources (and polyfills) hand back `{ done: false,
  // value: undefined }` for an empty read; that must not append an
  // undefined chunk.
  it("ignores a read that yields no value", async () => {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(undefined as unknown as Uint8Array)
        controller.enqueue(new Uint8Array([1]))
        controller.close()
      },
    })
    expect(Array.from(await bufferReadableStream(stream))).toEqual([1])
  })

  it("concatenates multiple chunks in order", async () => {
    const merged = await bufferReadableStream(
      streamOf(new Uint8Array([1, 2]), new Uint8Array([3]), new Uint8Array([4, 5])),
    )
    expect(Array.from(merged)).toEqual([1, 2, 3, 4, 5])
  })

  it("aborts once the running total passes the cap", async () => {
    await expect(
      bufferReadableStream(streamOf(new Uint8Array(10), new Uint8Array(10)), 15),
    ).rejects.toThrow(/exceeds the maximum of 15 bytes/)
  })

  // A caller passing 0, a negative number or Infinity has not asked for
  // "no limit" — the documented default applies instead.
  it.each([0, -1, Number.POSITIVE_INFINITY, Number.NaN])(
    "falls back to the default cap for a nonsensical limit (%s)",
    async (cap) => {
      const bytes = await bufferReadableStream(streamOf(new Uint8Array([9])), cap)
      expect(Array.from(bytes)).toEqual([9])
    },
  )
})
