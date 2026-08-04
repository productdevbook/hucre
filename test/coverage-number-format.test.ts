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
// Known deviations from Excel — see the report accompanying this file
// ═══════════════════════════════════════════════════════════════════════

describe("formatValue — Excel parity gaps", () => {
  // BUG: src/_format.ts:541 decides whether to prepend "-" from
  // `value < 0`, but `formatValue` already replaced the value with its
  // absolute value when it picked the negative section (line 108), so
  // `isNegative` is always false there. The explicit "-" written in the
  // format is swallowed into `core` by `extractLiterals` (line 598 puts
  // "-" in the placeholder set) and `formatIntegerPart` never emits it,
  // so the sign disappears entirely.
  // Excel renders -42 as "-42.00" under "0.00;-0.00".
  it.skip("keeps the explicit minus sign of a negative section", () => {
    expect(formatValue(-42, "0.00;-0.00")).toBe("-42.00")
    expect(formatValue(-1234.5, "#,##0.00;-#,##0.00")).toBe("-1,234.50")
  })

  // BUG: src/_format.ts:485-493. A single trailing comma is Excel's
  // "scale by 1000" directive, but `useThousandSep` is true for any
  // format containing `[#0?],` — which a trailing comma also satisfies —
  // and the scaling block is gated on `!useThousandSep`, so it never
  // runs. Excel renders 1234567 as "1,235" under "#,##0," and as "1235"
  // under "0,".
  it.skip("divides by a thousand for a single trailing comma", () => {
    expect(formatValue(1234567, "#,##0,")).toBe("1,235")
    expect(formatValue(1234567, "0,")).toBe("1235")
  })

  // BUG: src/_format.ts:485. When the scaling *does* apply (",,") the
  // same flag turns group separators off, so the scaled result loses its
  // commas. Excel renders 1234567890 as "1,235" under "#,##0,,".
  it.skip("keeps group separators on a scaled-down value", () => {
    expect(formatValue(1234567890, "#,##0,,")).toBe("1,235")
  })

  // BUG: src/_format.ts:404. `isFractionFormat` requires the denominator
  // to be made of `?`, `#` or `0`, so Excel's fixed-denominator fraction
  // formats ("As halves" = `# ?/2`, "As eighths" = `# ?/8`, "As
  // sixteenths" = `# ??/16`) are not recognised as fractions at all and
  // fall through to plain number formatting: 3.5 renders as "4/2".
  // `formatFraction`'s own `fixedDenom` branch (lines 429-443) is
  // therefore dead code — it can never be reached.
  it.skip("honours a fixed denominator written in the format", () => {
    expect(formatValue(3.5, "# ?/2")).toBe("3 1/2")
    expect(formatValue(3.5, "# ?/8")).toBe("3 4/8")
  })

  // BUG: src/_format.ts:447. For -0.5 the integer part is `-0`, and both
  // `intPart !== 0` and `intPart < 0` are false for negative zero, so no
  // sign is emitted. Excel renders -0.5 as "-1/2" under "# ?/?".
  it.skip("keeps the sign of a negative value smaller than one", () => {
    expect(formatValue(-0.5, "# ?/?")).toBe("-1/2")
  })

  // BUG: src/_format.ts:446. `hasIntPart` only looks for `#` and `0`, so
  // a `?` integer placeholder is not recognised and the whole part is
  // dropped from a mixed number. Excel renders 2.5 as "2 1/2" under
  // "? ?/?".
  it.skip("keeps the whole part in front of a ?-placeholder fraction", () => {
    expect(formatValue(2.5, "? ?/?")).toBe("2 1/2")
  })

  // BUG: src/_format.ts:695-702. The `?` decimal placeholder is meant to
  // render insignificant digits as spaces so decimal points line up, but
  // `decStr` is produced by `toFixed(decimalPlaces)` and therefore always
  // has exactly as many digits as the format has placeholders — the
  // `else result += " "` arm can never run and `?` behaves like `0`.
  // Excel renders 0.5 as "0.5  " under "0.???".
  it.skip("pads insignificant decimals with spaces for a ? placeholder", () => {
    expect(formatValue(0.5, "0.???")).toBe("0.5  ")
  })

  // BUG: src/_format.ts:598. Literal characters that sit *between* digit
  // placeholders (rather than before or after the number) are folded into
  // `core` and then dropped, because `formatIntegerPart` only understands
  // `0` and `#`. This breaks Excel's fixed-width identifier formats.
  it.skip("keeps separators written between digit groups", () => {
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
