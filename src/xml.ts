// ── hucre/xml entry point ─────────────────────────────────────────────
// Read & write tabular XML (product feeds, ERP exports, GS1, etc.).

export { readXml } from "./xml/data-reader";
export type { XmlReadOptions, XmlReadResult } from "./xml/data-reader";

export { writeXml } from "./xml/data-writer";
export type { XmlWriteOptions } from "./xml/data-writer";
