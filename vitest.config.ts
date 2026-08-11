import { defineConfig } from "vitest/config"

export default defineConfig({
  test: {
    // Vitest's default excludes cover node_modules and dist, but not
    // `.claude/worktrees` — git worktrees created by agent tooling live
    // there, each with a full copy of `test/`. Without this, a local run
    // silently collects every worktree's suite as well, and the reported
    // test count is whatever happens to be checked out beside you.
    exclude: ["**/node_modules/**", "**/dist/**", "**/.claude/**"],
    coverage: {
      // 98.8% coverage enforced by nothing is a number, not a floor —
      // it can slide a point a release and nobody notices until it has
      // slid ten. See #474.
      //
      // Set just under the measurement that produced them, so an
      // ordinary edit does not fail while a subsystem landing untested
      // does. Raise them when the real figure moves up; that is the
      // point of a ratchet.
      thresholds: {
        statements: 98.5,
        branches: 96,
        functions: 98,
        lines: 99,
      },
      // The CLI is verified by running the packaged binary
      // (scripts/verify-package.mjs), not by the unit suite, so counting
      // it here would report a floor the suite never defended.
      exclude: ["src/cli.ts", "src/cli/**"],
    },
  },
})
