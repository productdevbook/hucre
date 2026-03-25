import type { SheetImage } from "./_types";

/**
 * Decode a base64 string to a Uint8Array.
 * Works in both Node.js (via Buffer) and browsers (via atob).
 */
function base64ToUint8Array(base64: string): Uint8Array {
  // Strip data URI prefix if present (e.g. "data:image/png;base64,...")
  const clean = base64.includes(",") ? base64.slice(base64.indexOf(",") + 1) : base64;

  if (typeof globalThis !== "undefined" && "Buffer" in globalThis) {
    // Node.js
    return new Uint8Array((globalThis as any).Buffer.from(clean, "base64"));
  }

  // Browser
  const binary = atob(clean);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/** Create a SheetImage from a base64 string */
export function imageFromBase64(
  base64: string,
  type: "png" | "jpeg" | "gif",
  anchor: SheetImage["anchor"],
): SheetImage {
  const data = base64ToUint8Array(base64);
  return { data, type, anchor };
}
