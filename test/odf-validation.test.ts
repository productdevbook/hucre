import { afterEach, describe, expect, it } from "vitest"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { delimiter, join } from "node:path"
import { fileURLToPath } from "node:url"
import { spawnSync } from "node:child_process"
import type { SpawnSyncReturns } from "node:child_process"

// The validator is a developer script, not library code: it needs `jing`,
// a 40 MB OASIS grammar and a JVM, so the suite cannot run it for real.
// What it can run is the script's own bookkeeping — which parts counted
// as checked — against stub `unzip` and `java` on PATH and a stub bundle
// behind `--dist`. That is the part #566 got wrong.

let work: string | undefined

afterEach(() => {
  if (work) rmSync(work, { recursive: true, force: true })
  work = undefined
})

/** The line the script prints for one part, in its own column layout. */
function line(part: string, status: string): string {
  return `  ${part.padEnd(24)} ${status}`
}

function runValidator(unzipMode: "missing-settings" | "unreadable"): SpawnSyncReturns<string> {
  work = mkdtempSync(join(tmpdir(), "hucre-odf-validation-test-"))
  const bin = join(work, "bin")
  mkdirSync(bin)

  // A stub for every export `documents()` reaches for. Add to the script
  // and this list grows with it — the alternative is building dist to
  // test the counting.
  const dist = join(work, "fake-dist.mjs")
  writeFileSync(
    dist,
    `export async function writeOds() { return new Uint8Array([1]) }
export function writeOdsStream() {
  let read = false
  return { getReader: () => ({ async read() {
    if (read) return { done: true }
    read = true
    return { done: false, value: new Uint8Array([1]) }
  } }) }
}
export class OdsStreamWriter {
  addRow() {}
  async finish() { return new Uint8Array([1]) }
}
export function createCellStore() { return {} }
export function setCell() {}
`,
  )

  // `unzip -p <file> <part>`: one part refuses to come out, or the whole
  // package does.
  writeFileSync(
    join(bin, "unzip"),
    `#!/bin/sh
if [ "$FAKE_UNZIP_MODE" = "missing-settings" ] && [ "$3" != "settings.xml" ]; then
  printf '<valid/>'
  exit 0
fi
exit 1
`,
    { mode: 0o755 },
  )
  // jing, accepting everything it is handed.
  writeFileSync(join(bin, "java"), "#!/bin/sh\nexit 0\n", { mode: 0o755 })

  return spawnSync(
    process.execPath,
    [
      fileURLToPath(new URL("../scripts/validate-odf.mjs", import.meta.url)),
      "--schema",
      "unused.rng",
      "--jing",
      "unused.jar",
      "--dist",
      dist,
    ],
    {
      cwd: work,
      encoding: "utf8",
      env: {
        ...process.env,
        PATH: `${bin}${delimiter}${process.env["PATH"]}`,
        FAKE_UNZIP_MODE: unzipMode,
      },
    },
  )
}

describe("ODF validator CLI", () => {
  it("fails when a required package part is missing (#566)", () => {
    const result = runValidator("missing-settings")

    expect(result.status).toBe(1)
    expect(result.stdout).toContain(line("settings.xml", "MISSING/UNREADABLE"))
    // The parts that did come out are still reported as checked, so the
    // failure is the missing one rather than a package-wide collapse.
    expect(result.stdout).toContain(line("content.xml", "valid"))
    expect(result.stdout).toContain(line("styles.xml", "valid"))
    expect(result.stdout).toContain(line("meta.xml", "valid"))
    // One part × the three documents the script writes.
    expect(result.stdout).toContain("3 invalid or unreadable part(s).")
    expect(result.stdout).not.toContain("All parts valid.")
  })

  it("fails when the package cannot be read at all (#566)", () => {
    const result = runValidator("unreadable")

    expect(result.status).toBe(1)
    expect(result.stdout).toContain(line("content.xml", "MISSING/UNREADABLE"))
    expect(result.stdout).not.toContain(line("content.xml", "valid"))
    // Four parts × three documents, none of them checked.
    expect(result.stdout).toContain("12 invalid or unreadable part(s).")
  })

  it("names the manifest check as skipped when no schema is supplied (#566)", () => {
    const result = runValidator("unreadable")

    expect(result.stdout).toContain(
      line("META-INF/manifest.xml", "SKIPPED (no manifest schema supplied)"),
    )
  })
})
