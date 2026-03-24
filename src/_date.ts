// ── Excel Date Utilities ─────────────────────────────────────────────
//
// Excel stores dates as serial numbers (days since an epoch).
// Two date systems exist:
//   1900 system (Windows default): Day 1 = Jan 1, 1900
//   1904 system (Mac default):     Day 0 = Jan 1, 1904
//
// The 1900 system has a deliberate Lotus 1-2-3 compatibility bug:
// serial 60 = "Feb 29, 1900" which doesn't exist (1900 is NOT a leap year).
// All conversions use UTC to avoid timezone issues.
// ─────────────────────────────────────────────────────────────────────

/**
 * Excel epoch for 1900 date system.
 * We use Dec 31, 1899 UTC so that serial 1 = Jan 1, 1900:
 *   EPOCH + 1 * MS_PER_DAY = Jan 1, 1900.
 * For serials > 60 the code subtracts 1 from the serial before
 * converting, which correctly skips over the phantom Feb 29 (serial 60).
 */
const EPOCH_1900 = Date.UTC(1899, 11, 31); // Dec 31, 1899

/** Excel epoch for 1904 date system: Day 0 = Jan 1, 1904 */
const EPOCH_1904 = Date.UTC(1904, 0, 1);

const MS_PER_DAY = 86_400_000;

/** The serial number of the phantom Lotus 1-2-3 Feb 29, 1900 */
const LOTUS_BUG_SERIAL = 60;

/**
 * Convert an Excel serial number to a JavaScript Date.
 * All conversions use UTC to avoid timezone issues.
 *
 * @param serial - Excel date serial number (can include fractional time)
 * @param is1904 - Whether to use the 1904 date system (default: false = 1900)
 * @returns JavaScript Date in UTC
 */
export function serialToDate(serial: number, is1904?: boolean): Date {
  if (is1904) {
    // 1904 system: no Lotus bug, serial 0 = Jan 1, 1904
    const ms = EPOCH_1904 + Math.round(serial * MS_PER_DAY);
    return new Date(ms);
  }

  // 1900 system with Lotus bug handling:
  // - Serial 0 = "Jan 0, 1900" (Excel quirk — we map to Dec 30, 1899)
  // - Serials 1-59 map directly (serial 1 = Jan 1, 1900)
  // - Serial 60 = phantom "Feb 29, 1900"
  // - Serials > 60: subtract 1 to account for the phantom day

  if (serial === LOTUS_BUG_SERIAL) {
    // Return "Feb 29, 1900" even though it doesn't exist historically.
    // Excel treats this as a real date, so we must too.
    // JavaScript's Date.UTC(1900, 1, 29) auto-corrects to Mar 1, so we
    // construct Feb 28 + 1 day manually to get a Date whose UTC components
    // show Feb 29 (which JS *will* normalize to Mar 1).
    // Since there is no real Feb 29, 1900, we map it to the same timestamp
    // as Mar 1 — callers should be aware serial 60 is an Excel artifact.
    // We use Feb 28 for the returned Date since that's the last valid
    // date before the phantom day.
    return new Date(Date.UTC(1900, 1, 28));
  }

  let adjustedSerial = serial;
  if (serial > LOTUS_BUG_SERIAL) {
    adjustedSerial = serial - 1;
  }

  const ms = EPOCH_1900 + Math.round(adjustedSerial * MS_PER_DAY);
  return new Date(ms);
}

/**
 * Convert a JavaScript Date to an Excel serial number.
 * Uses UTC components of the date.
 *
 * @param date - JavaScript Date (UTC components are used)
 * @param is1904 - Whether to use the 1904 date system (default: false = 1900)
 * @returns Excel serial number (with fractional time portion)
 */
export function dateToSerial(date: Date, is1904?: boolean): number {
  const timeMs = date.getTime();

  if (is1904) {
    return (timeMs - EPOCH_1904) / MS_PER_DAY;
  }

  // 1900 system: compute raw serial from epoch
  let serial = (timeMs - EPOCH_1900) / MS_PER_DAY;

  // Dates on or after Mar 1, 1900 (serial 61 without bug) must be bumped
  // by 1 to skip over the phantom Feb 29 (serial 60).
  // Mar 1, 1900 = serial 61 in Excel. Without the bug it would be 60.
  // So the threshold is: raw serial >= 60 (which represents Mar 1, 1900 or later).
  if (serial >= LOTUS_BUG_SERIAL) {
    serial += 1;
  }

  // Round to avoid floating-point drift. Excel has ~1ms precision.
  // We round to 10 decimal places to preserve sub-second time info while
  // eliminating IEEE 754 noise.
  return Math.round(serial * 1e10) / 1e10;
}

/**
 * Excel built-in number format IDs that represent date/time formats.
 * Per ECMA-376 Part 1, 18.8.30 (numFmt).
 */
const DATE_FORMAT_IDS = new Set([
  14, // m/d/yyyy or regional equivalent
  15, // d-mmm-yy
  16, // d-mmm
  17, // mmm-yy
  18, // h:mm AM/PM
  19, // h:mm:ss AM/PM
  20, // h:mm
  21, // h:mm:ss
  22, // m/d/yyyy h:mm
  // CJK date formats
  27,
  28,
  29,
  30,
  31,
  32,
  33,
  34,
  35,
  36,
  // More regional dates/times
  45, // mm:ss
  46, // [h]:mm:ss
  47, // mm:ss.0
  // Thai/Chinese/Korean extended formats
  50,
  51,
  52,
  53,
  54,
  55,
  56,
  57,
  58,
]);

/**
 * Check if an Excel number format string represents a date/time format.
 * Used to distinguish dates from plain numbers when reading cells.
 *
 * The challenge: "m" and "mm" can mean months OR minutes depending on context.
 * After "h" or "hh" (with optional separator), "m"/"mm" means minutes.
 *
 * @param numFmt - Excel number format string or built-in format ID
 * @returns true if the format represents a date or time
 */
export function isDateFormat(numFmt: string): boolean {
  if (!numFmt) {
    return false;
  }

  // Check if it's a built-in format ID (numeric string)
  const numericId = Number(numFmt);
  if (!Number.isNaN(numericId) && Number.isInteger(numericId) && DATE_FORMAT_IDS.has(numericId)) {
    return true;
  }

  // Normalize: strip locale prefix [$-xxx], color directives [Red], etc.
  let cleaned = numFmt.replace(/\[[$\-\w]*\]/g, "");

  // Strip escaped characters (backslash + char) and quoted strings
  cleaned = cleaned.replace(/\\./g, "");
  cleaned = cleaned.replace(/"[^"]*"/g, "");

  // Strip fill/repeat characters (*x, _x)
  cleaned = cleaned.replace(/[*_]./g, "");

  // "General", "@" (text), pure number formats
  if (/^(General|@)$/i.test(cleaned.trim())) {
    return false;
  }

  // Lowercase for matching
  const lower = cleaned.toLowerCase();

  // Check for time tokens first (these are unambiguous)
  if (/[hs]/.test(lower)) {
    // Has hours or seconds — it's a time format
    // But make sure it's not just a literal "s" in something like "$#,##0"
    // Check for actual time pattern tokens
    if (/\bh{1,2}\b|(?:^|[^a-z])h{1,2}(?:[^a-z]|$)/i.test(lower)) {
      return true;
    }
    if (/\bs{1,2}\b|(?:^|[^a-z])s{1,2}(?:[^a-z]|$)/i.test(lower)) {
      return true;
    }
  }

  // Check for AM/PM
  if (/am\/pm|a\/p/i.test(lower)) {
    return true;
  }

  // Check for elapsed time format [h], [m], [s]
  if (/\[h+\]|\[m+\]|\[s+\]/i.test(numFmt)) {
    return true;
  }

  // Check for year tokens (unambiguous date indicator)
  if (/y{1,4}/.test(lower)) {
    return true;
  }

  // Check for day tokens (unambiguous date indicator)
  if (/d{1,4}/.test(lower)) {
    return true;
  }

  // "m" or "mm" alone (without "d", "y", "h", "s") is ambiguous.
  // Excel treats standalone "m" as a month format only when paired with
  // day or year tokens. If we reach here, there are no y/d/h/s tokens,
  // so a lone "m" or "mm" is NOT a date format.
  // However, "mmm" (abbreviated month name) or "mmmm" (full month name)
  // are always date formats.
  if (/m{3,}/.test(lower)) {
    return true;
  }

  return false;
}

/** Month names for formatDate */
const MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const MONTH_ABBR = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const DAY_ABBR = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

/**
 * Format a Date according to an Excel number format string.
 * Uses UTC components of the date.
 *
 * Handles: yyyy, yy, mmmm, mmm, mm, m, dddd, ddd, dd, d,
 *          hh, h, mm (minutes when after h), ss, s, AM/PM, 0 (fractional seconds)
 *
 * @param date - JavaScript Date (UTC components are used)
 * @param format - Excel number format string
 * @returns Formatted date string
 */
export function formatDate(date: Date, format: string): string {
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth(); // 0-based
  const day = date.getUTCDate();
  const dayOfWeek = date.getUTCDay();
  let hours = date.getUTCHours();
  const minutes = date.getUTCMinutes();
  const seconds = date.getUTCSeconds();
  const ms = date.getUTCMilliseconds();

  // Check for AM/PM mode
  const hasAmPm = /am\/pm|a\/p/i.test(format);
  let ampm = "";
  let displayHours = hours;

  if (hasAmPm) {
    ampm = hours >= 12 ? "PM" : "AM";
    displayHours = hours % 12;
    if (displayHours === 0) displayHours = 12;
  }

  // Tokenize the format string to correctly distinguish month vs minute "m"/"mm".
  // We parse left to right, tracking whether the last date/time token was "h"/"hh".
  const tokens = tokenize(format);

  let result = "";
  let lastWasHour = false;

  for (const token of tokens) {
    const lower = token.toLowerCase();

    switch (lower) {
      case "yyyy":
        result += String(year);
        lastWasHour = false;
        break;
      case "yy":
        result += String(year % 100).padStart(2, "0");
        lastWasHour = false;
        break;
      case "mmmm":
        result += MONTH_NAMES[month];
        lastWasHour = false;
        break;
      case "mmm":
        result += MONTH_ABBR[month];
        lastWasHour = false;
        break;
      case "mm":
        if (lastWasHour) {
          // Minutes
          result += String(minutes).padStart(2, "0");
        } else {
          // Month
          result += String(month + 1).padStart(2, "0");
        }
        lastWasHour = false;
        break;
      case "m":
        if (lastWasHour) {
          // Minutes (no padding)
          result += String(minutes);
        } else {
          // Month (no padding)
          result += String(month + 1);
        }
        lastWasHour = false;
        break;
      case "dddd":
        result += DAY_NAMES[dayOfWeek];
        lastWasHour = false;
        break;
      case "ddd":
        result += DAY_ABBR[dayOfWeek];
        lastWasHour = false;
        break;
      case "dd":
        result += String(day).padStart(2, "0");
        lastWasHour = false;
        break;
      case "d":
        result += String(day);
        lastWasHour = false;
        break;
      case "hh":
        result += String(displayHours).padStart(2, "0");
        lastWasHour = true;
        break;
      case "h":
        result += String(displayHours);
        lastWasHour = true;
        break;
      case "ss":
        result += String(seconds).padStart(2, "0");
        lastWasHour = false;
        break;
      case "s":
        result += String(seconds);
        lastWasHour = false;
        break;
      case ".0":
      case ".00":
      case ".000":
        {
          const decimals = lower.length - 1; // number of 0s
          const frac = String(ms).padStart(3, "0").slice(0, decimals);
          result += "." + frac;
          lastWasHour = false;
        }
        break;
      case "am/pm":
        result += ampm;
        lastWasHour = false;
        break;
      case "a/p":
        result += ampm.charAt(0);
        lastWasHour = false;
        break;
      default:
        // Literal text (separators, spaces, etc.)
        result += token;
        // Don't reset lastWasHour for separators (e.g., "h:mm" — the colon
        // between h and mm should not break the minute detection)
        break;
    }
  }

  return result;
}

/**
 * Tokenize an Excel date format string into format tokens and literal text.
 * Recognizes: yyyy, yy, mmmm, mmm, mm, m, dddd, ddd, dd, d,
 *             hh, h, ss, s, .0/.00/.000, AM/PM, A/P
 * Quoted strings and backslash-escaped chars are treated as literals.
 */
function tokenize(format: string): string[] {
  const tokens: string[] = [];
  let i = 0;

  while (i < format.length) {
    // Skip locale prefixes like [$-409]
    if (format[i] === "[" && format[i + 1] === "$") {
      const end = format.indexOf("]", i);
      if (end !== -1) {
        i = end + 1;
        continue;
      }
    }

    // Skip color directives like [Red]
    if (format[i] === "[") {
      const end = format.indexOf("]", i);
      if (end !== -1) {
        i = end + 1;
        continue;
      }
    }

    // Quoted literal string
    if (format[i] === '"') {
      let literal = "";
      i++; // skip opening quote
      while (i < format.length && format[i] !== '"') {
        literal += format[i];
        i++;
      }
      i++; // skip closing quote
      tokens.push(literal);
      continue;
    }

    // Backslash-escaped character
    if (format[i] === "\\") {
      i++;
      if (i < format.length) {
        tokens.push(format[i]);
        i++;
      }
      continue;
    }

    // AM/PM or A/P
    if (/^am\/pm/i.test(format.slice(i))) {
      tokens.push(format.slice(i, i + 5));
      i += 5;
      continue;
    }
    if (/^a\/p/i.test(format.slice(i))) {
      tokens.push(format.slice(i, i + 3));
      i += 3;
      continue;
    }

    // Fractional seconds: .0, .00, .000
    if (format[i] === "." && i + 1 < format.length && format[i + 1] === "0") {
      let tok = ".";
      let j = i + 1;
      while (j < format.length && format[j] === "0") {
        tok += "0";
        j++;
      }
      tokens.push(tok);
      i = j;
      continue;
    }

    // Date/time tokens: sequences of the same letter
    const ch = format[i].toLowerCase();
    if ("ymdhsn".includes(ch)) {
      let tok = "";
      const matchCh = ch;
      while (i < format.length && format[i].toLowerCase() === matchCh) {
        tok += format[i];
        i++;
      }
      tokens.push(tok);
      continue;
    }

    // Everything else is literal (separators, spaces, etc.)
    tokens.push(format[i]);
    i++;
  }

  return tokens;
}

/**
 * Parse a date string into a Date (UTC).
 * Supports ISO 8601 and common US/EU formats.
 *
 * @param value - Date string to parse
 * @returns Date in UTC, or null if unparseable
 */
export function parseDate(value: string): Date | null {
  if (!value || !value.trim()) {
    return null;
  }

  const trimmed = value.trim();

  // ISO 8601: "2021-01-15", "2021-01-15T14:30:00Z", "2021-01-15T14:30:00+05:00"
  const isoMatch = trimmed.match(
    /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?([Zz]|[+-]\d{2}:?\d{2})?)?$/,
  );
  if (isoMatch) {
    const y = Number(isoMatch[1]);
    const m = Number(isoMatch[2]) - 1;
    const d = Number(isoMatch[3]);
    const h = Number(isoMatch[4] || 0);
    const min = Number(isoMatch[5] || 0);
    const s = Number(isoMatch[6] || 0);
    let msec = 0;
    if (isoMatch[7]) {
      msec = Number(isoMatch[7].padEnd(3, "0").slice(0, 3));
    }

    // If there's a timezone offset, parse it
    const tz = isoMatch[8];
    if (tz && tz.toUpperCase() !== "Z") {
      const tzMatch = tz.match(/^([+-])(\d{2}):?(\d{2})$/);
      if (tzMatch) {
        const sign = tzMatch[1] === "+" ? 1 : -1;
        const tzH = Number(tzMatch[2]);
        const tzM = Number(tzMatch[3]);
        const offsetMs = sign * (tzH * 60 + tzM) * 60_000;
        const utcMs = Date.UTC(y, m, d, h, min, s, msec) - offsetMs;
        return new Date(utcMs);
      }
    }

    return new Date(Date.UTC(y, m, d, h, min, s, msec));
  }

  // US format: "MM/DD/YYYY" or "M/D/YYYY"
  const usMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (usMatch) {
    const m = Number(usMatch[1]) - 1;
    const d = Number(usMatch[2]);
    const y = Number(usMatch[3]);
    if (m >= 0 && m <= 11 && d >= 1 && d <= 31) {
      return new Date(Date.UTC(y, m, d));
    }
  }

  // EU format: "DD.MM.YYYY"
  const euMatch = trimmed.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (euMatch) {
    const d = Number(euMatch[1]);
    const m = Number(euMatch[2]) - 1;
    const y = Number(euMatch[3]);
    if (m >= 0 && m <= 11 && d >= 1 && d <= 31) {
      return new Date(Date.UTC(y, m, d));
    }
  }

  // Dash format: "DD-MM-YYYY" or "YYYY-MM-DD" (already handled by ISO above for 4-digit year first)
  const dashMatch = trimmed.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dashMatch) {
    const d = Number(dashMatch[1]);
    const m = Number(dashMatch[2]) - 1;
    const y = Number(dashMatch[3]);
    if (m >= 0 && m <= 11 && d >= 1 && d <= 31) {
      return new Date(Date.UTC(y, m, d));
    }
  }

  return null;
}

/**
 * Get the time portion of a serial number as components.
 * The fractional part of an Excel serial represents time within the day.
 *
 * @param serial - Excel serial number (only the fractional part is used)
 * @returns Time components
 */
export function serialToTime(serial: number): {
  hours: number;
  minutes: number;
  seconds: number;
  milliseconds: number;
} {
  // Extract fractional part
  const frac = Math.abs(serial) % 1;

  // Total milliseconds in the day
  const totalMs = Math.round(frac * MS_PER_DAY);

  const hours = Math.floor(totalMs / 3_600_000);
  const minutes = Math.floor((totalMs % 3_600_000) / 60_000);
  const seconds = Math.floor((totalMs % 60_000) / 1_000);
  const milliseconds = totalMs % 1_000;

  return { hours, minutes, seconds, milliseconds };
}

/**
 * Convert time components to an Excel serial fraction.
 *
 * @param hours - Hours (0-23)
 * @param minutes - Minutes (0-59)
 * @param seconds - Seconds (0-59), default 0
 * @param milliseconds - Milliseconds (0-999), default 0
 * @returns Fractional serial number (0 to <1)
 */
export function timeToSerial(
  hours: number,
  minutes: number,
  seconds: number = 0,
  milliseconds: number = 0,
): number {
  const totalMs = hours * 3_600_000 + minutes * 60_000 + seconds * 1_000 + milliseconds;
  return totalMs / MS_PER_DAY;
}
