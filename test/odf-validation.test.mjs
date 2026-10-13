import { afterEach, describe, expect, it } from "vitest"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { delimiter, join } from "node:path"
import { fileURLToPath } from "node:url"
import { spawnSync } from "node:child_process"

let work

afterEach(() => {
  if (work) rmSync(work, { recursive: true, force: true })
  work = undefined
})

function runValidator(unzipMode) {
  work = mkdtempSync(join(tmpdir(), "hucre-odf-validation-test-"))
  const bin = join(work, "bin")
  mkdirSync(bin)

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
`,
  )

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
        PATH: `${bin}${delimiter}${process.env.PATH}`,
        FAKE_UNZIP_MODE: unzipMode,
      },
    },
  )
}

describe("ODF validator CLI", () => {
  it("fails when a required package part is missing", () => {
    const result = runValidator("missing-settings")

    expect(result.status).toBe(1)
    expect(result.stdout).toContain("settings.xml             MISSING/UNREADABLE")
    expect(result.stdout).toContain("invalid or unreadable part(s)")
    expect(result.stdout).not.toContain("All parts valid.")
  })

  it("fails when the package cannot be read", () => {
    const result = runValidator("unreadable")

    expect(result.status).toBe(1)
    expect(result.stdout).toContain("MISSING/UNREADABLE")
    expect(result.stdout).not.toContain("All parts valid.")
  })

  it("labels manifest validation as skipped when no schema is supplied", () => {
    const result = runValidator("unreadable")

    expect(result.stdout).toContain("SKIPPED (no manifest schema supplied)")
  })
})
