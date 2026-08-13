#!/usr/bin/env node
// ── Does hucre write valid ODF? ──────────────────────────────────────
//
// Every ODS test in this repository reads the file back with hucre, so
// they answer "does hucre agree with itself". They cannot answer "is this
// a valid OpenDocument", and LibreOffice — lenient by design — cannot
// either: it opened both of the defects this script found on its first
// run.
//
//   * `<style:style style:family="table-cell">` has an *ordered* content
//     model. hucre wrote `style:text-properties` before
//     `style:table-cell-properties`, so every cell with both a font and a
//     fill produced a document the schema rejects.
//   * `<office:settings/>` was written empty. The element is optional and,
//     when present, requires one or more `config:config-item-set`.
//
// And one the schema caught the day after it was introduced:
// `number:min-decimal-places` is an ODF 1.3 attribute, and hucre declared
// `office:version="1.2"`.
//
// What this is: the published RELAX NG grammar, run over documents hucre
// wrote, by `jing` — the reference validator. Neither the schemas nor the
// validator are vendored; they are large and none of them is hucre's to
// redistribute.
//
//   curl -sSLO https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/OpenDocument-v1.3-schema.rng
//   curl -sSLO https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/OpenDocument-v1.3-manifest-schema.rng
//   curl -sSLO https://repo1.maven.org/maven2/org/relaxng/jing/20220510/jing-20220510.jar
//
//   node scripts/validate-odf.mjs \
//     --schema OpenDocument-v1.3-schema.rng \
//     --manifest-schema OpenDocument-v1.3-manifest-schema.rng \
//     --jing jing-20220510.jar
//
// `-i` is passed to jing throughout: the OASIS grammar does not compile
// under its ID/IDREF checking, which is a property of the schema rather
// than of any document.

import { mkdtempSync, readFileSync, writeFileSync, mkdirSync } from "node:fs"
import { execFileSync } from "node:child_process"
import { tmpdir } from "node:os"
import { join, dirname } from "node:path"

const args = new Map()
for (let i = 2; i < process.argv.length; i += 2) {
  args.set(process.argv[i].replace(/^--/, ""), process.argv[i + 1])
}
const schema = args.get("schema")
const manifestSchema = args.get("manifest-schema")
const jing = args.get("jing")
const dist = args.get("dist") ?? new URL("../dist/index.mjs", import.meta.url).href

if (!schema || !jing) {
  console.error("usage: node scripts/validate-odf.mjs --schema <odf.rng> --jing <jing.jar>")
  console.error("       [--manifest-schema <manifest.rng>]  (see the header for where to get them)")
  process.exit(1)
}

const hucre = await import(dist)

// ── The documents to check ───────────────────────────────────────────
//
// Chosen to reach the parts of the writer that produce structure rather
// than values: a style with several facets at once (which is what found
// the child-order bug), every kind of number format, and each of the
// three writers, because they build the document differently.

const STYLE = {
  font: { bold: true, italic: true, size: 14, color: { rgb: "FF0000" } },
  fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
  alignment: { horizontal: "center" },
}

const NUMBER_FORMATS = [
  "0.00",
  "#.##",
  "#,##0.00",
  "0%",
  "0.0#%",
  '"$"#,##0.00',
  "yyyy-mm-dd",
  "hh:mm:ss",
  "0.00E+00",
  "##0.0E+0",
  "@",
]

async function drain(stream) {
  const chunks = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}

async function documents() {
  const out = []

  const cells = new Map([["0,0", { value: "Name", style: STYLE }]])
  NUMBER_FORMATS.forEach((numFmt, i) => {
    cells.set(`${i + 1},0`, { value: 1234.5, style: { numFmt } })
  })
  cells.set(`${NUMBER_FORMATS.length + 1},0`, { value: 1, formula: "A2+1" })
  cells.set(`${NUMBER_FORMATS.length + 2},0`, {
    value: "link",
    hyperlink: { target: "https://example.test", tooltip: "t" },
  })

  const rows = [
    ["Name", "Qty"],
    ...NUMBER_FORMATS.map(() => [1234.5, null]),
    [1, null],
    ["link", null],
  ]

  out.push([
    "writeOds (styles, formats, formula, hyperlink, merge)",
    await hucre.writeOds({
      sheets: [
        {
          name: "Data",
          rows,
          cells,
          columns: [{ width: 20 }, { width: 10 }],
          merges: ["A1:B1"],
        },
        { name: "Second", rows: [["a", "b"]] },
      ],
      properties: { title: "T", creator: "C", description: "D" },
      // Named ranges live in the epilogue and spell a sheet name with a
      // space as `$'My Sheet'.$A$1`. The round trip could not tell the
      // quoted form from the bare one; this could.
      namedRanges: [
        { name: "Region", range: "Data!$A$1:$B$2" },
        { name: "Quoted", range: "'Second'!$A$1:$A$1" },
      ],
    }),
  ])

  out.push([
    "writeOdsStream (values only)",
    await drain(hucre.writeOdsStream(rows, { name: "Stream", columns: [{ width: 12 }] })),
  ])

  const incremental = new hucre.OdsStreamWriter({
    name: "Incremental",
    columns: [{ header: "A", width: 15 }, { header: "B" }],
  })
  incremental.addRow(["x", { value: 1, style: STYLE }])
  incremental.addRow([{ value: 2, style: { numFmt: "#,##0.00" } }, null])
  out.push(["OdsStreamWriter (buffered, styled)", await incremental.finish()])

  return out
}

// ── Validation ───────────────────────────────────────────────────────

const PARTS = ["content.xml", "styles.xml", "meta.xml", "settings.xml"]

/** Pull one entry out of an ODS with hucre's own ZIP reader. */
async function extract(bytes, path) {
  const { ZipReader } = await import(new URL("../dist/zip/reader.mjs", import.meta.url).href).catch(
    () => ({}),
  )
  if (ZipReader) return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
  // The bundle does not export it; read the workbook back through the
  // public API is not enough, so fall back to `unzip -p`.
  const dir = mkdtempSync(join(tmpdir(), "hucre-odf-"))
  const file = join(dir, "doc.ods")
  writeFileSync(file, bytes)
  return execFileSync("unzip", ["-p", file, path], {
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  })
}

function validate(file, rng) {
  try {
    execFileSync("java", ["-jar", jing, "-i", rng, file], { encoding: "utf8", stdio: "pipe" })
    return null
  } catch (error) {
    return `${error.stdout ?? ""}${error.stderr ?? ""}`.trim() || String(error)
  }
}

const work = mkdtempSync(join(tmpdir(), "hucre-odf-val-"))
let failures = 0

for (const [label, bytes] of await documents()) {
  console.log(`\n${label}  (${bytes.length} bytes)`)

  const checks = [...PARTS.map((p) => [p, schema])]
  if (manifestSchema) checks.push(["META-INF/manifest.xml", manifestSchema])

  for (const [part, rng] of checks) {
    let xml
    try {
      xml = await extract(bytes, part)
    } catch {
      console.log(`  ${part.padEnd(24)} (absent)`)
      continue
    }

    const file = join(work, part.replace(/\//g, "_"))
    mkdirSync(dirname(file), { recursive: true })
    writeFileSync(file, xml)

    const errors = validate(file, rng)
    if (errors) {
      failures++
      console.log(`  ${part.padEnd(24)} INVALID`)
      for (const line of errors.split("\n").slice(0, 4)) console.log(`    ${line}`)
    } else {
      console.log(`  ${part.padEnd(24)} valid`)
    }
  }
}

console.log(failures === 0 ? "\nAll parts valid." : `\n${failures} invalid part(s).`)
process.exit(failures === 0 ? 0 : 1)
