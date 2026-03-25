// ── Theme Color Parser ─────────────────────────────────────────────
// Parses xl/theme/theme1.xml and resolves theme colors to RGB hex.

import { parseSax } from "../xml/parser";

/**
 * The 12 standard theme color slots in Office Open XML order:
 * [dk1, lt1, dk2, lt2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink]
 */
const COLOR_ELEMENT_ORDER = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

/**
 * Parse xl/theme/theme1.xml and extract the color scheme.
 * Returns an array of 12 colors (6-char hex RGB, uppercase) in order:
 * [dk1, lt1, dk2, lt2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink]
 */
export function parseThemeColors(xml: string): string[] {
  const colorMap = new Map<string, string>();

  // SAX state
  let inClrScheme = false;
  let currentSlot = "";

  parseSax(xml, {
    onOpenTag(tag, attrs) {
      if (tag === "clrScheme" || tag === "a:clrScheme") {
        inClrScheme = true;
        return;
      }

      if (!inClrScheme) return;

      // Strip namespace prefix
      const local = tag.includes(":") ? tag.split(":").pop()! : tag;

      // Check if this is a color slot element (dk1, lt1, accent1, etc.)
      if (COLOR_ELEMENT_ORDER.includes(local)) {
        currentSlot = local;
        return;
      }

      // Inside a color slot, look for srgbClr or sysClr
      if (currentSlot) {
        if (local === "srgbClr" && attrs["val"]) {
          colorMap.set(currentSlot, attrs["val"].toUpperCase());
        } else if (local === "sysClr" && attrs["lastClr"]) {
          colorMap.set(currentSlot, attrs["lastClr"].toUpperCase());
        }
      }
    },
    onCloseTag(tag) {
      const local = tag.includes(":") ? tag.split(":").pop()! : tag;
      if (local === "clrScheme") {
        inClrScheme = false;
      }
      if (currentSlot && COLOR_ELEMENT_ORDER.includes(local)) {
        currentSlot = "";
      }
    },
  });

  // Build the ordered array
  return COLOR_ELEMENT_ORDER.map((slot) => colorMap.get(slot) ?? "000000");
}

/**
 * Resolve a theme color index + optional tint to a 6-char RGB hex string.
 *
 * Theme index mapping (per OOXML spec):
 *   0 → dk1, 1 → lt1, 2 → dk2, 3 → lt2,
 *   4 → accent1, 5 → accent2, 6 → accent3, 7 → accent4,
 *   8 → accent5, 9 → accent6, 10 → hlink, 11 → folHlink
 *
 * Tint algorithm (per OOXML spec):
 *   - tint < 0: darken each channel: newVal = val * (1 + tint)
 *   - tint > 0: lighten each channel: newVal = val + (255 - val) * tint
 *   - tint == 0 or undefined: no change
 */
export function resolveThemeColor(
  themeColors: string[],
  themeIndex: number,
  tint?: number,
): string {
  if (themeIndex < 0 || themeIndex >= themeColors.length) {
    return "000000";
  }

  const hex = themeColors[themeIndex];

  if (tint === undefined || tint === 0) {
    return hex;
  }

  // Parse RGB components
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);

  // Apply tint
  const applyTint = (channel: number): number => {
    let result: number;
    if (tint < 0) {
      // Darken
      result = channel * (1 + tint);
    } else {
      // Lighten
      result = channel + (255 - channel) * tint;
    }
    return Math.round(Math.min(255, Math.max(0, result)));
  };

  const nr = applyTint(r);
  const ng = applyTint(g);
  const nb = applyTint(b);

  return (
    nr.toString(16).padStart(2, "0").toUpperCase() +
    ng.toString(16).padStart(2, "0").toUpperCase() +
    nb.toString(16).padStart(2, "0").toUpperCase()
  );
}
