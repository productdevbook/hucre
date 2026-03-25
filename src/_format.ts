// ── Number Format Renderer ─────────────────────────────────────────────
//
// Apply an Excel number format string to a value and return formatted text.
// Handles: General, numbers, currency, percentage, scientific, date/time,
// fractions, accounting, multi-section formats, color codes, conditions,
// locale prefixes.
// ─────────────────────────────────────────────────────────────────────

import { isDateFormat, formatDate, serialToDate, dateToSerial } from "./_date";

// ── Locale Definitions ──────────────────────────────────────────────

export interface LocaleFormat {
  /** Decimal separator character */
  decimal: string;
  /** Thousands grouping separator character */
  thousands: string;
  /** Currency symbol */
  currency: string;
}

const LOCALE_MAP: Record<string, LocaleFormat> = {
  "en-US": { decimal: ".", thousands: ",", currency: "$" },
  "de-DE": { decimal: ",", thousands: ".", currency: "\u20AC" },
  "fr-FR": { decimal: ",", thousands: "\u00A0", currency: "\u20AC" },
  "tr-TR": { decimal: ",", thousands: ".", currency: "\u20BA" },
};

/** Resolve a locale string to its format definition, or undefined if unsupported. */
function resolveLocale(locale?: string): LocaleFormat | undefined {
  if (!locale) return undefined;
  return LOCALE_MAP[locale];
}

export interface FormatOptions {
  /** BCP 47 locale tag for number formatting (e.g. "de-DE", "tr-TR"). */
  locale?: string;
}

/**
 * Apply an Excel number format string to a value and return formatted text.
 *
 * @param value - The raw cell value (number, string, boolean, Date)
 * @param numFmt - Excel number format string (e.g., "#,##0.00", "0%", "yyyy-mm-dd")
 * @param options - Optional formatting options (locale, etc.)
 * @returns Formatted string
 */
export function formatValue(value: unknown, numFmt: string, options?: FormatOptions): string {
  // Null/undefined → ""
  if (value === null || value === undefined) {
    return "";
  }

  // Boolean → "TRUE"/"FALSE"
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  // No format or "General"
  if (!numFmt || /^General$/i.test(numFmt.trim())) {
    if (value instanceof Date) {
      return value.toISOString();
    }
    return String(value);
  }

  // Parse sections: positive;negative;zero;text
  const sections = splitSections(numFmt);

  // If value is a string
  if (typeof value === "string") {
    // Use text section (4th) if available, otherwise return as-is
    const textSection = sections.length >= 4 ? sections[3] : sections[0];
    return applyTextSection(value, textSection);
  }

  // Convert Date to serial for numeric formatting
  let numValue: number;
  if (value instanceof Date) {
    numValue = dateToSerial(value);
  } else if (typeof value === "number") {
    numValue = value;
  } else {
    return String(value);
  }

  // Select the right section based on value sign
  let section: string;
  if (sections.length >= 3) {
    if (numValue > 0) {
      section = sections[0];
    } else if (numValue < 0) {
      section = sections[1];
      numValue = Math.abs(numValue); // negative section handles sign display
    } else {
      section = sections[2];
    }
  } else if (sections.length === 2) {
    if (numValue >= 0) {
      section = sections[0];
    } else {
      section = sections[1];
      numValue = Math.abs(numValue);
    }
  } else {
    section = sections[0];
    // For single-section, keep sign handling in the formatting
  }

  // Check for conditions in the section like [>100]
  const condResult = extractCondition(section);
  if (condResult.condition) {
    // With conditions, we use all sections but match against condition
    if (sections.length >= 2) {
      if (evaluateCondition(typeof value === "number" ? value : numValue, condResult.condition)) {
        section = condResult.rest;
      } else {
        section = sections.length >= 2 ? stripCondition(sections[1]) : condResult.rest;
      }
    } else {
      section = condResult.rest;
    }
  }

  const localeInfo = resolveLocale(options?.locale);
  return applyNumberSection(numValue, section, localeInfo);
}

// ── Section Parsing ─────────────────────────────────────────────────

/**
 * Split format string by unquoted semicolons.
 * Respects quoted strings and escaped characters.
 */
function splitSections(fmt: string): string[] {
  const sections: string[] = [];
  let current = "";
  let inQuote = false;
  let i = 0;

  while (i < fmt.length) {
    const ch = fmt[i];

    if (ch === "\\") {
      current += ch;
      i++;
      if (i < fmt.length) {
        current += fmt[i];
        i++;
      }
      continue;
    }

    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
      i++;
      continue;
    }

    if (ch === ";" && !inQuote) {
      sections.push(current);
      current = "";
      i++;
      continue;
    }

    current += ch;
    i++;
  }

  sections.push(current);
  return sections;
}

// ── Color & Locale Stripping ────────────────────────────────────────

/** Strip color codes like [Red], [Blue], [Color 3] etc. */
function stripColorCodes(fmt: string): string {
  return fmt.replace(/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\s*\d+)\]/gi, "");
}

/** Strip locale prefixes like [$-409], [$€-407], [$-F800] */
function stripLocalePrefix(fmt: string): string {
  return fmt.replace(/\[\$[^\]]*\]/g, "");
}

/** Strip fill/padding characters like _( and *  */
function stripFillPadding(fmt: string): string {
  return fmt.replace(/[_*]./g, "");
}

/** Clean a format section: remove color, locale, padding */
function cleanSection(fmt: string): string {
  let cleaned = stripColorCodes(fmt);
  cleaned = stripLocalePrefix(cleaned);
  cleaned = stripFillPadding(cleaned);
  return cleaned;
}

// ── Condition Handling ──────────────────────────────────────────────

interface Condition {
  operator: string;
  value: number;
}

function extractCondition(fmt: string): { condition: Condition | null; rest: string } {
  const match = fmt.match(/\[([<>=!]+)(-?\d+(?:\.\d+)?)\]/);
  if (!match) {
    return { condition: null, rest: fmt };
  }

  return {
    condition: { operator: match[1], value: Number(match[2]) },
    rest: fmt.replace(match[0], ""),
  };
}

function stripCondition(fmt: string): string {
  return fmt.replace(/\[([<>=!]+)(-?\d+(?:\.\d+)?)\]/, "");
}

function evaluateCondition(value: number, cond: Condition): boolean {
  switch (cond.operator) {
    case ">":
      return value > cond.value;
    case "<":
      return value < cond.value;
    case ">=":
      return value >= cond.value;
    case "<=":
      return value <= cond.value;
    case "=":
    case "==":
      return value === cond.value;
    case "<>":
    case "!=":
      return value !== cond.value;
    default:
      return true;
  }
}

// ── Text Section ────────────────────────────────────────────────────

function applyTextSection(value: string, section: string): string {
  const cleaned = cleanSection(section);

  // Expand quoted strings and backslash-escaped chars, then replace @ with value
  const expanded = expandLiterals(cleaned);

  if (expanded.includes("@")) {
    return expanded.replace(/@/g, value);
  }

  // If no @ placeholder, return value as-is
  return value;
}

/** Expand quoted strings ("text") and backslash escapes (\c) into literal text */
function expandLiterals(fmt: string): string {
  let result = "";
  let i = 0;
  while (i < fmt.length) {
    if (fmt[i] === '"') {
      i++;
      while (i < fmt.length && fmt[i] !== '"') {
        result += fmt[i];
        i++;
      }
      i++; // skip closing quote
    } else if (fmt[i] === "\\") {
      i++;
      if (i < fmt.length) {
        result += fmt[i];
        i++;
      }
    } else {
      result += fmt[i];
      i++;
    }
  }
  return result;
}

// ── Number Section ──────────────────────────────────────────────────

function applyNumberSection(value: number, section: string, locale?: LocaleFormat): string {
  const cleaned = cleanSection(section);

  // Text format: @ — return as string
  if (cleaned.trim() === "@") {
    return String(value);
  }

  // Check if it's a date format — delegate to formatDate
  if (isDateFormat(cleaned)) {
    const date = serialToDate(value);
    return formatDate(date, section); // Pass original section, formatDate handles [Red] etc
  }

  // Percentage: multiply by 100
  if (cleaned.includes("%")) {
    return formatPercentage(value, cleaned, locale);
  }

  // Scientific notation
  if (/[eE][+-]/.test(cleaned) || /[eE]\d/.test(cleaned)) {
    return formatScientific(value, cleaned, locale);
  }

  // Fractions
  if (isFractionFormat(cleaned)) {
    return formatFraction(value, cleaned);
  }

  // Regular number format
  return formatNumber(value, cleaned, locale);
}

// ── Percentage ──────────────────────────────────────────────────────

function formatPercentage(value: number, fmt: string, locale?: LocaleFormat): string {
  const percentValue = value * 100;
  // Remove the % sign, format the number, then add % back
  const numFmt = fmt.replace(/%/g, "");
  const formatted = formatNumber(percentValue, numFmt, locale);
  return formatted + "%";
}

// ── Scientific Notation ─────────────────────────────────────────────

function formatScientific(value: number, fmt: string, locale?: LocaleFormat): string {
  // Parse the format: e.g., "0.00E+00"
  const match = fmt.match(/^([#0?.,]*?)([eE])([+-])(\d+)$/);
  if (!match) {
    // Fallback: determine decimal places from the mantissa part
    const decMatch = fmt.match(/\.([0#?]+)[eE]/);
    const decPlaces = decMatch ? decMatch[1].length : 2;
    const expStr = value.toExponential(decPlaces);
    return formatExponentialString(expStr, fmt);
  }

  const mantissaFmt = match[1];
  const eChar = match[2];
  const signChar = match[3];
  const expDigits = match[4].length;

  // Count decimal places in mantissa
  const dotIdx = mantissaFmt.indexOf(".");
  const decPlaces = dotIdx >= 0 ? mantissaFmt.length - dotIdx - 1 : 0;

  const expStr = value.toExponential(decPlaces);
  const parts = expStr.split(/[eE]/);
  let mantissa = parts[0];
  let exp = Number.parseInt(parts[1], 10);

  // Apply locale decimal separator
  if (locale && locale.decimal !== ".") {
    mantissa = mantissa.replace(".", locale.decimal);
  }

  const expSign = exp >= 0 ? "+" : "-";
  const absExp = Math.abs(exp).toString().padStart(expDigits, "0");

  const displaySign = signChar === "+" ? expSign : exp < 0 ? "-" : "";

  return mantissa + eChar + displaySign + absExp;
}

function formatExponentialString(expStr: string, fmt: string): string {
  const parts = expStr.split(/[eE]/);
  const mantissa = parts[0];
  let exp = Number.parseInt(parts[1], 10);

  // Determine E character case
  const eChar = fmt.includes("E") ? "E" : "e";
  const hasPlus = fmt.includes("E+") || fmt.includes("e+");

  const expSign = exp >= 0 ? "+" : "-";
  const absExp = Math.abs(exp).toString().padStart(2, "0");

  const displaySign = hasPlus ? expSign : exp < 0 ? "-" : "";

  return mantissa + eChar + displaySign + absExp;
}

// ── Fraction Format ─────────────────────────────────────────────────

function isFractionFormat(fmt: string): boolean {
  // Matches patterns like "# ?/?", "# ??/??", "# ?/8", etc.
  // But not date formats or paths
  return /[#0?]\s*[?#0]+\/[?#0]+/.test(fmt);
}

function formatFraction(value: number, fmt: string): string {
  const intPart = Math.trunc(value);
  let frac = Math.abs(value - intPart);

  if (frac === 0) {
    // Show integer only for whole numbers
    const showInt = fmt.includes("#") || /^[0?]/.test(fmt.trim());
    if (showInt && intPart !== 0) {
      return String(intPart);
    }
    return String(intPart) + "      "; // padded like Excel
  }

  // Determine denominator precision from format
  const fracMatch = fmt.match(/([?#0]+)\/([?#0]+)/);
  if (!fracMatch) {
    return String(value);
  }

  const denomLen = fracMatch[2].length;

  // Check for fixed denominator (all digits)
  const fixedDenom = /^\d+$/.test(fracMatch[2]) ? Number.parseInt(fracMatch[2], 10) : 0;

  let bestNum: number;
  let bestDen: number;

  if (fixedDenom > 0) {
    bestDen = fixedDenom;
    bestNum = Math.round(frac * fixedDenom);
  } else {
    // Find best fraction with denominator up to 10^denomLen
    const maxDen = Math.pow(10, denomLen) - 1;
    const result = findBestFraction(frac, maxDen);
    bestNum = result.num;
    bestDen = result.den;
  }

  // Build the formatted string
  const hasIntPart = fmt.includes("#") || fmt.includes("0");
  const prefix = intPart !== 0 && hasIntPart ? String(intPart) + " " : intPart < 0 ? "-" : "";

  const numStr = String(bestNum).padStart(fracMatch[1].length, " ");
  const denStr = String(bestDen).padStart(fracMatch[2].length, " ");

  return prefix + numStr + "/" + denStr;
}

function findBestFraction(value: number, maxDen: number): { num: number; den: number } {
  let bestNum = 0;
  let bestDen = 1;
  let bestError = Math.abs(value);

  for (let den = 1; den <= maxDen; den++) {
    const num = Math.round(value * den);
    const error = Math.abs(value - num / den);
    if (error < bestError) {
      bestError = error;
      bestNum = num;
      bestDen = den;
      if (error === 0) break;
    }
  }

  return { num: bestNum, den: bestDen };
}

// ── Number Formatting ───────────────────────────────────────────────

function formatNumber(value: number, fmt: string, locale?: LocaleFormat): string {
  // Extract literal strings and escaped chars first
  const literals: { index: number; text: string }[] = [];
  let stripped = "";
  let i = 0;
  let outputIndex = 0;

  // We need to track prefix/suffix text vs number placeholders
  // Strategy: extract all quoted text and escaped chars, then parse the number format core

  let prefixText = "";
  let suffixText = "";

  // Extract currency symbol and literal text from the format
  const { prefix, suffix, core } = extractLiterals(fmt);

  if (!core.trim()) {
    // No number placeholders at all — return just the literal text
    return prefix + suffix;
  }

  const hasThousands = core.includes(",") && /[#0?],/.test(core);
  const useThousandSep = hasThousands && !core.match(/,{2,}/); // ,, means scale down

  // Count trailing commas (each divides by 1000)
  let scaleDown = 0;
  const scaleMatch = core.match(/(,+)(?=[^#0?]*$)/);
  if (scaleMatch && !useThousandSep) {
    scaleDown = scaleMatch[1].length;
  }

  let scaledValue = value;
  for (let s = 0; s < scaleDown; s++) {
    scaledValue /= 1000;
  }

  // Determine decimal places
  const dotIndex = core.indexOf(".");
  let decimalPlaces = 0;
  if (dotIndex >= 0) {
    const afterDot = core.slice(dotIndex + 1).replace(/[^0#?]/g, "");
    decimalPlaces = afterDot.length;
  }

  // Round the value
  const roundedValue = roundToDecimal(Math.abs(scaledValue), decimalPlaces);
  const isNegative = value < 0;

  // Split into integer and decimal parts
  const [intStr, decStr] = splitNumber(roundedValue, decimalPlaces);

  // Format integer part
  const intFmt = dotIndex >= 0 ? core.slice(0, dotIndex) : core;
  const formattedInt = formatIntegerPart(intStr, intFmt.replace(/,/g, ""), useThousandSep);

  // Format decimal part
  let formattedDec = "";
  if (dotIndex >= 0) {
    const decFmt = core.slice(dotIndex + 1);
    formattedDec = "." + formatDecimalPart(decStr, decFmt);
  }

  // Apply locale-specific separators if requested
  let localizedInt = formattedInt;
  let localizedDec = formattedDec;
  if (locale) {
    if (locale.thousands !== "," && useThousandSep) {
      localizedInt = localizedInt.replace(/,/g, locale.thousands);
    }
    if (locale.decimal !== "." && localizedDec.length > 0) {
      // Replace the leading "." with locale decimal
      localizedDec = locale.decimal + localizedDec.slice(1);
    }
  }

  // Combine
  let result = prefix;
  if (isNegative && fmt.indexOf("-") === -1) {
    // Only add minus if the format doesn't explicitly have one
    result += "-";
  }
  result += localizedInt + localizedDec + suffix;

  return result;
}

/**
 * Extract literal prefix/suffix text and the core number format.
 */
function extractLiterals(fmt: string): { prefix: string; suffix: string; core: string } {
  let prefix = "";
  let suffix = "";
  let core = "";
  let i = 0;
  let foundDigitPlaceholder = false;
  let afterDigits = false;

  while (i < fmt.length) {
    const ch = fmt[i];

    // Quoted string
    if (ch === '"') {
      let literal = "";
      i++;
      while (i < fmt.length && fmt[i] !== '"') {
        literal += fmt[i];
        i++;
      }
      i++; // skip closing quote
      if (!foundDigitPlaceholder) {
        prefix += literal;
      } else {
        afterDigits = true;
        suffix += literal;
      }
      continue;
    }

    // Escaped character
    if (ch === "\\") {
      i++;
      if (i < fmt.length) {
        if (!foundDigitPlaceholder) {
          prefix += fmt[i];
        } else {
          afterDigits = true;
          suffix += fmt[i];
        }
        i++;
      }
      continue;
    }

    // Digit placeholders or format chars
    if ("#0?.,%Ee+-".includes(ch)) {
      if (afterDigits && "#0?".includes(ch)) {
        // More digit placeholders after suffix text — unusual but handle it
        core += suffix + ch;
        suffix = "";
        afterDigits = false;
      } else {
        core += ch;
      }
      if ("#0?".includes(ch)) {
        foundDigitPlaceholder = true;
      }
      i++;
      continue;
    }

    // Comma within number section
    if (ch === ",") {
      if (foundDigitPlaceholder) {
        core += ch;
      }
      i++;
      continue;
    }

    // Currency symbols and other characters
    if (!foundDigitPlaceholder) {
      prefix += ch;
    } else {
      afterDigits = true;
      suffix += ch;
    }
    i++;
  }

  return { prefix, suffix, core };
}

function roundToDecimal(value: number, decimals: number): number {
  const factor = Math.pow(10, decimals);
  return Math.round(value * factor) / factor;
}

function splitNumber(value: number, decimalPlaces: number): [string, string] {
  const fixed = value.toFixed(decimalPlaces);
  const dotIdx = fixed.indexOf(".");
  if (dotIdx < 0) {
    return [fixed, ""];
  }
  return [fixed.slice(0, dotIdx), fixed.slice(dotIdx + 1)];
}

function formatIntegerPart(intStr: string, fmt: string, useThousandSep: boolean): string {
  // Count minimum digits from format (0s require digits, # are optional)
  const minDigits = (fmt.match(/0/g) || []).length;
  const hasHash = fmt.includes("#");

  // Pad with leading zeros if needed
  let padded = intStr;
  if (padded.length < minDigits) {
    padded = padded.padStart(minDigits, "0");
  }

  // If all # and value is 0, show nothing (or just 0 if minDigits > 0)
  if (padded === "0" && minDigits === 0 && hasHash) {
    padded = "";
  }

  // Add thousand separators
  if (useThousandSep && padded.length > 0) {
    padded = addThousandSeparators(padded);
  }

  return padded;
}

function formatDecimalPart(decStr: string, fmt: string): string {
  // The format contains 0, #, ? placeholders
  let result = "";
  const cleanFmt = fmt.replace(/[^0#?]/g, "");

  for (let i = 0; i < cleanFmt.length; i++) {
    const placeholder = cleanFmt[i];
    const digit = i < decStr.length ? decStr[i] : "0";

    switch (placeholder) {
      case "0":
        // Always show digit
        result += digit;
        break;
      case "#":
        // Show digit only if significant (trailing zeros suppressed)
        // Check if there are any non-zero digits from this position onwards
        if (hasSignificantDigits(decStr, i)) {
          result += digit;
        }
        break;
      case "?":
        // Show digit or space
        if (i < decStr.length) {
          result += digit;
        } else {
          result += " ";
        }
        break;
    }
  }

  return result;
}

function hasSignificantDigits(str: string, fromIndex: number): boolean {
  for (let i = fromIndex; i < str.length; i++) {
    if (str[i] !== "0") return true;
  }
  return false;
}

function addThousandSeparators(intStr: string): string {
  // Handle negative sign
  const negative = intStr.startsWith("-");
  const digits = negative ? intStr.slice(1) : intStr;

  let result = "";
  const len = digits.length;
  for (let i = 0; i < len; i++) {
    if (i > 0 && (len - i) % 3 === 0) {
      result += ",";
    }
    result += digits[i];
  }

  return negative ? "-" + result : result;
}
