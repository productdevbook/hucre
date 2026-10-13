import { mkdirSync, writeFileSync } from "node:fs"
import { dirname, join } from "node:path"

export async function validateOdfParts({
  documents,
  parts,
  schema,
  manifestSchema,
  work,
  extract,
  validate,
  log,
}) {
  let failures = 0
  let skipped = 0

  for (const [label, bytes] of documents) {
    log(`\n${label}  (${bytes.length} bytes)`)

    const checks = parts.map((part) => [part, schema])
    if (manifestSchema) {
      checks.push(["META-INF/manifest.xml", manifestSchema])
    } else {
      skipped++
      log("  META-INF/manifest.xml   SKIPPED (no manifest schema supplied)")
    }

    for (const [part, rng] of checks) {
      let xml
      try {
        xml = await extract(bytes, part)
      } catch {
        // A required part that could not be extracted is not a successful check.
        failures++
        log(`  ${part.padEnd(24)} MISSING/UNREADABLE`)
        continue
      }

      const file = join(work, part.replace(/\//g, "_"))
      mkdirSync(dirname(file), { recursive: true })
      writeFileSync(file, xml)

      const errors = validate(file, rng)
      if (errors) {
        failures++
        log(`  ${part.padEnd(24)} INVALID`)
        for (const line of errors.split("\n").slice(0, 4)) log(`    ${line}`)
      } else {
        log(`  ${part.padEnd(24)} valid`)
      }
    }
  }

  return { failures, skipped }
}
