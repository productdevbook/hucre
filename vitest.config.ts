import { defineConfig } from "vitest/config"

export default defineConfig({
  test: {
    // Vitest's default excludes cover node_modules and dist, but not
    // `.claude/worktrees` — git worktrees created by agent tooling live
    // there, each with a full copy of `test/`. Without this, a local run
    // silently collects every worktree's suite as well, and the reported
    // test count is whatever happens to be checked out beside you.
    exclude: ["**/node_modules/**", "**/dist/**", "**/.claude/**"],
  },
})
