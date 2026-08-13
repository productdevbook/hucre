#!/usr/bin/env node
// ── Spec coverage: what the format defines against what hucre handles ─
//
// Finding gaps by probing finds one at a time, and only the ones someone
// thought to probe for. The formats are specified, and the specifications
// are machine-readable — ECMA-376 ships XSDs for SpreadsheetML, OASIS
// ships a RELAX NG grammar for ODF. So the gap list can be *derived*
// rather than discovered.
//
// This reads a schema, reads hucre's source, reads the fixture corpus,
// and crosses the three:
//
//   in spec + in corpus + in code   supported, and a real file exercises it
//   in spec + in corpus + NOT code  **a gap that occurs in files people have**
//   in spec + no corpus + in code   supported, nothing here exercises it
//   in spec + no corpus + NOT code  not supported, and nothing here needs it
//
// The second row is the point. It is the difference between "hucre does
// not implement all of ECMA-376" — true of every implementation, Excel
// included — and "hucre drops something that is in the files you have".
//
// What this is not: a conformance test. Finding the name `pivotField` in
// `src/` does not prove hucre reads it correctly, only that something
// there knows the word. It answers "what have we never heard of", which
// is the question that cannot be answered by testing.
//
// The schemas are not vendored — ECMA-376 is a 42 MB download and the
// ODF grammar is 583 KB, and neither is hucre's to redistribute:
//
//   curl -sSLO https://www.ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip
//   unzip -o ECMA-376-1_5th_edition_december_2016.zip OfficeOpenXML-XMLSchema-Strict.zip
//   unzip -o OfficeOpenXML-XMLSchema-Strict.zip -d xsd
//   curl -sSLO https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/OpenDocument-v1.3-schema.rng
//
//   node scripts/spec-coverage.mjs --sml xsd/sml.xsd --odf OpenDocument-v1.3-schema.rng
//
// Either may be omitted; the report covers what it was given.

import { readdirSync, readFileSync, statSync, writeFileSync } from "node:fs"
import { join, extname } from "node:path"

// ── Arguments ────────────────────────────────────────────────────────

const args = new Map()
for (let i = 2; i < process.argv.length; i += 2) {
  args.set(process.argv[i].replace(/^--/, ""), process.argv[i + 1])
}
const smlPath = args.get("sml")
const odfPath = args.get("odf")
const outPath = args.get("out") ?? "docs/SPEC-COVERAGE.md"
const srcDir = args.get("src") ?? "src"
const fixturesDir = args.get("fixtures") ?? "test/fixtures"

if (!smlPath && !odfPath) {
  console.error("usage: node scripts/spec-coverage.mjs --sml <sml.xsd> --odf <odf.rng>")
  console.error("       (see the header of this file for where to get them)")
  process.exit(1)
}

// ── The names hucre's source knows ───────────────────────────────────

/**
 * Every string literal in `src/`.
 *
 * A reader switches on a local name — `case "sheetPr"` — and a writer
 * emits one, so a name the source has never heard of cannot appear as a
 * literal anywhere in it. The converse does not hold, which is why the
 * report says "mentioned" rather than "supported".
 */
function sourceLiterals(dir) {
  /** The whole literal is the name: `case "sheetPr"`, `attr["spans"]`. */
  const exact = new Set()
  /** The name occurs anywhere, including inside a longer template. */
  let corpusText = ""

  const walk = (d) => {
    for (const entry of readdirSync(d)) {
      const p = join(d, entry)
      if (statSync(p).isDirectory()) walk(p)
      else if (extname(p) === ".ts") {
        const text = readFileSync(p, "utf8")
        corpusText += text + "\n"
        // A quoted literal — `case "sheetPr"` — is how a reader switches.
        for (const m of text.matchAll(/["'`]([A-Za-z_][\w:.\-]*)["'`]/g)) exact.add(m[1])
        // An object key — `xmlSelfClose("calcPr", { calcId: 0 })` — is how
        // a writer emits an attribute. Missing this form reported a dozen
        // attributes hucre writes on every workbook as unknown to it.
        for (const m of text.matchAll(/(?:^|[{,\s])([A-Za-z_]\w*)\s*:/gm)) exact.add(m[1])
      }
    }
  }
  walk(dir)

  // Two signals, because one is not enough. A writer that emits
  // `<calcPr fullCalcOnLoad="1"/>` as one template has the name nowhere
  // as a literal of its own, so an exact-match-only report calls it a
  // gap — and hucre has written that attribute since #474. A
  // word-anywhere report has the opposite fault: `row`, `count`, `name`
  // and `value` occur in any codebase, so everything looks handled.
  //
  // Reporting both separates "the source switches on this" from "the
  // word appears somewhere" from "never heard of it".
  return {
    exact,
    mentions(name) {
      return new RegExp(`\\b${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`).test(corpusText)
    },
  }
}

// ── The names the corpus actually contains ───────────────────────────

/** Minimal ZIP entry reader — enough to pull the XML parts out. */
async function zipEntries(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let i = bytes.length - 22; i >= 0 && i > bytes.length - 66000; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i
      break
    }
  }
  if (eocd < 0) return []

  const count = dv.getUint16(eocd + 10, true)
  let at = dv.getUint32(eocd + 16, true)
  const out = []

  for (let n = 0; n < count && at + 46 <= bytes.length; n++) {
    if (dv.getUint32(at, true) !== 0x02014b50) break
    const method = dv.getUint16(at + 10, true)
    const compSize = dv.getUint32(at + 20, true)
    const nameLen = dv.getUint16(at + 28, true)
    const extraLen = dv.getUint16(at + 30, true)
    const commentLen = dv.getUint16(at + 32, true)
    const local = dv.getUint32(at + 42, true)
    const name = new TextDecoder().decode(bytes.subarray(at + 46, at + 46 + nameLen))
    at += 46 + nameLen + extraLen + commentLen

    if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue

    const lnameLen = dv.getUint16(local + 26, true)
    const lextraLen = dv.getUint16(local + 28, true)
    const start = local + 30 + lnameLen + lextraLen
    const raw = bytes.subarray(start, start + compSize)

    if (method === 0) {
      out.push(new TextDecoder().decode(raw))
    } else if (method === 8) {
      const ds = new DecompressionStream("deflate-raw")
      const stream = new Blob([raw]).stream().pipeThrough(ds)
      out.push(new TextDecoder().decode(new Uint8Array(await new Response(stream).arrayBuffer())))
    }
  }
  return out
}

/** Element and attribute local names present in the fixture corpus. */
async function corpusNames(dir) {
  const elements = new Set()
  const attributes = new Set()

  const files = []
  const walk = (d) => {
    for (const entry of readdirSync(d)) {
      const p = join(d, entry)
      if (statSync(p).isDirectory()) walk(p)
      else if (p.endsWith(".xlsx") || p.endsWith(".xlsb")) files.push(p)
    }
  }
  walk(dir)

  for (const file of files) {
    for (const xml of await zipEntries(new Uint8Array(readFileSync(file)))) {
      for (const m of xml.matchAll(/<\/?([A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)/g))
        elements.add(m[2])
      for (const m of xml.matchAll(/[\s]([A-Za-z_][\w.-]*)=["']/g)) attributes.add(m[1])
    }
  }
  return { elements, attributes, fileCount: files.length }
}

// ── SpreadsheetML, from the XSD ──────────────────────────────────────

/**
 * Every element and attribute the schema declares, grouped by the
 * complexType that declares it — which is the unit a reader works in.
 */
function parseSml(path) {
  const xsd = readFileSync(path, "utf8")
  const types = []

  for (const block of xsd.split(/<xsd:complexType\s/).slice(1)) {
    const nameMatch = block.match(/^name="(CT_[\w]+)"/)
    if (!nameMatch) continue
    const body = block.split("</xsd:complexType>")[0]

    const elements = [...body.matchAll(/<xsd:element\s+name="([\w]+)"/g)].map((m) => m[1])
    const attributes = [...body.matchAll(/<xsd:attribute\s+name="([\w]+)"/g)].map((m) => m[1])
    if (elements.length === 0 && attributes.length === 0) continue

    types.push({ name: nameMatch[1], elements, attributes })
  }
  return types
}

// ── ODF, from the RELAX NG grammar ───────────────────────────────────

/** Elements and attributes in the namespaces a spreadsheet uses. */
function parseOdf(path) {
  const rng = readFileSync(path, "utf8")
  // `table:` and `office:` are the spreadsheet's own vocabulary, and
  // `number:` is the data-style one. `style:`, `text:`, `draw:` and `fo:`
  // are shared with the rest of ODF — most of a text document's grammar
  // is reachable from a cell, and listing all of it would bury the part
  // that is about spreadsheets in a thousand lines that are not.
  const wanted = /^(table|office|number):/

  const elements = new Map()
  for (const m of rng.matchAll(/<(?:\w+:)?element\s+name="([\w:.-]+)"/g)) {
    if (!wanted.test(m[1])) continue
    elements.set(m[1], (elements.get(m[1]) ?? 0) + 1)
  }

  const attributes = new Map()
  for (const m of rng.matchAll(/<(?:\w+:)?attribute\s+name="([\w:.-]+)"/g)) {
    if (!wanted.test(m[1])) continue
    attributes.set(m[1], (attributes.get(m[1]) ?? 0) + 1)
  }
  return { elements: [...elements.keys()].sort(), attributes: [...attributes.keys()].sort() }
}

// ── Triage ───────────────────────────────────────────────────────────

/**
 * Gaps that have been looked at and deliberately left.
 *
 * The point of writing them down is the ratchet: a regeneration lists
 * only what is *not* here, so a new gap stands out instead of being lost
 * in thirty already-judged ones. Removing a name from this map is how you
 * reopen the question.
 */
const REVIEWED = new Map(
  Object.entries({
    // Window geometry and view state. Where a window sat on someone's
    // screen is not data, and hucre does not model a window.
    autoFilterDateGrouping: "view state — not modelled",
    firstSheet: "view state — not modelled",
    minimized: "window geometry — not modelled",
    showHorizontalScroll: "window geometry — not modelled",
    showSheetTabs: "window geometry — not modelled",
    showVerticalScroll: "window geometry — not modelled",
    tabRatio: "window geometry — not modelled",
    tabSelected: "view state — not modelled",
    zoomToFit: "view state — not modelled",
    activeCell: "view state — not modelled",

    // Provenance of the writing application.
    appName: "writer provenance — not modelled",
    lastEdited: "writer provenance — not modelled",
    lowestEdited: "writer provenance — not modelled",
    rupBuild: "writer provenance — not modelled",
    fileVersion: "writer provenance — not modelled",
    defaultThemeVersion: "writer provenance — not modelled",

    // Out of scope by design.
    filterPrivacy: "privacy flag, no data — not modelled",
    pivotButton: "pivot UI affordance — pivots are round-trip only",
    customFormat: "restates that the row has a style, which the style says",
    outlineLevelCol: "summary of the per-column outline levels hucre reads",
    outlineLevelRow: "summary of the per-row outline levels hucre reads",

    visibility: "window visibility — not modelled",
    shapeId:
      "the VML shape a comment or control is drawn as. hucre generates its own on write and does not need the file's on read",
    spans:
      "a hint at which columns a row uses. hucre derives that from the cells themselves, which is the authority — Excel treats a wrong `spans` as advisory too",

    // Genuinely worth attention — kept here with the reason, so the
    // report shows them as judged rather than unseen.
    quotePrefix:
      "**open** — marks a cell forced to text by a leading apostrophe. The value is already a string in the file, so nothing is lost from `rows`; the flag itself is not carried",
    baseColWidth:
      "**open** — the base width column widths are relative to. hucre assumes the 8.43 default; a file that sets another makes every width wrong",
    indexedColors:
      "**open** — the legacy palette. `ColorSpec.indexed` carries the *index* and hucre has no palette to resolve it against, so a file that overrides the palette is dropped silently. Not in PARITY.md",
    rgbColor: "**open** — an entry of the `indexedColors` palette above",
  }),
)

// ── Report ───────────────────────────────────────────────────────────

function classify(name, literals, corpus) {
  const inCorpus = corpus.has(name)
  const handled = literals.exact.has(name)
  const mentioned = handled || literals.mentions(name)

  if (handled && inCorpus) return "both"
  if (handled) return "code"
  if (inCorpus) return mentioned ? "weak" : "gap"
  return "neither"
}

const literals = sourceLiterals(srcDir)
const corpus = await corpusNames(fixturesDir)

const lines = []
lines.push("# Spec coverage")
lines.push("")
lines.push(
  "Generated by `scripts/spec-coverage.mjs`. **Not** a conformance result:",
  "a name found in `src/` proves the source knows the word, not that hucre",
  "reads it correctly. What it answers is the question testing cannot —",
  "*what has this library never heard of* — and, by crossing the schema",
  "with the fixture corpus, which of those actually turn up in real files.",
  "",
)
lines.push(
  `Corpus: ${corpus.fileCount} workbooks under \`${fixturesDir}\`, ` +
    `${corpus.elements.size} distinct element names, ` +
    `${corpus.attributes.size} distinct attribute names.`,
  "",
)

if (smlPath) {
  const types = parseSml(smlPath)
  const rows = []
  let totals = { both: 0, code: 0, gap: 0, weak: 0, neither: 0 }

  for (const type of types) {
    for (const [kind, names] of [
      ["element", type.elements],
      ["attribute", type.attributes],
    ]) {
      for (const name of names) {
        const bucket = kind === "element" ? corpus.elements : corpus.attributes
        const verdict = classify(name, literals, bucket)
        totals[verdict]++
        if (verdict === "gap" || verdict === "weak")
          rows.push({ type: type.name, kind, name, verdict })
      }
    }
  }

  lines.push("## SpreadsheetML (ECMA-376 Part 1)", "")
  lines.push(
    `${types.length} complex types. ` +
      `**${totals.gap}** names appear in the corpus and are *nowhere* in \`src/\`; ` +
      `${totals.weak} appear in the corpus and occur in \`src/\` only inside a longer ` +
      `string, which usually means a writer emits them in a template; ` +
      `${totals.both} the source switches on directly; ${totals.code} it knows but ` +
      `the corpus does not use; ${totals.neither} are in neither.`,
    "",
  )

  if (rows.length > 0) {
    const fresh = rows.filter((r) => !REVIEWED.has(r.name))
    const judged = rows.filter((r) => REVIEWED.has(r.name))
    const byName = (a, b) => a.type.localeCompare(b.type) || a.name.localeCompare(b.name)

    lines.push("### Not yet looked at", "")
    if (fresh.length === 0) {
      lines.push(
        "Nothing. Every name the corpus uses and the source does not switch on",
        "has been judged — see the table below. A new one appearing here is the",
        "signal this report exists for.",
        "",
      )
    } else {
      lines.push("| complex type | kind | name | in src at all? |", "| --- | --- | --- | --- |")
      for (const r of fresh.sort(byName)) {
        lines.push(
          `| \`${r.type}\` | ${r.kind} | \`${r.name}\` | ${r.verdict === "gap" ? "**no**" : "in a template"} |`,
        )
      }
      lines.push("")
    }

    lines.push("### Looked at, and left", "")
    lines.push("| name | kind | why |", "| --- | --- | --- |")
    const seen = new Set()
    for (const r of judged.sort(byName)) {
      if (seen.has(r.name)) continue
      seen.add(r.name)
      lines.push(`| \`${r.name}\` | ${r.kind} | ${REVIEWED.get(r.name)} |`)
    }
    lines.push("")
  } else {
    lines.push("Nothing in the corpus is unknown to the source.", "")
  }
}

if (odfPath) {
  const { elements, attributes } = parseOdf(odfPath)
  const missing = { element: [], attribute: [] }
  let known = 0

  for (const [kind, names] of [
    ["element", elements],
    ["attribute", attributes],
  ]) {
    for (const name of names) {
      const local = name.includes(":") ? name.split(":")[1] : name
      // Prefixed or bare: the ODS reader switches on the local name, the
      // writer emits the prefixed one.
      if (literals.exact.has(name) || literals.exact.has(local)) known++
      else missing[kind].push(name)
    }
  }

  lines.push("## OpenDocument (OASIS ODF 1.3)", "")
  lines.push(
    `${elements.length} elements and ${attributes.length} attributes in the ` +
      "spreadsheet-relevant namespaces. " +
      `${known} are named somewhere in \`src/\`; ` +
      `${missing.element.length + missing.attribute.length} are not.`,
    "",
    "ODF is not crossed with the corpus: there are no third-party `.ods`",
    "fixtures beyond the SheetJS ones, and SheetJS writes a narrow subset.",
    "",
  )
  lines.push("### Not named in the source", "")
  lines.push(
    "`table:`, `office:` and `number:` only — the spreadsheet's own",
    "vocabulary and its data styles. `style:`, `text:`, `draw:` and `fo:`",
    "are shared with the rest of ODF; most of a text document's grammar is",
    "reachable from a cell, and listing it would bury this in a thousand",
    "lines that are not about spreadsheets.",
    "",
  )
  for (const [kind, names] of Object.entries(missing)) {
    if (names.length === 0) continue
    lines.push(`<details><summary>${names.length} ${kind}s</summary>`, "")
    lines.push("```")
    lines.push(...names)
    lines.push("```", "</details>", "")
  }
}

writeFileSync(outPath, lines.join("\n"))
console.log(`wrote ${outPath} (${lines.length} lines)`)
