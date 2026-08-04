#!/usr/bin/env node

// ── CLI entry point ─────────────────────────────────────────────────
// hucre convert input.xlsx output.csv
// hucre convert input.csv output.xlsx
// hucre convert input.xlsx output.ods
// hucre inspect file.xlsx
// hucre inspect file.xlsx --sheet 0
// hucre validate data.xlsx --schema schema.json
//
// Only the bin lives here. The commands are in ./cli/commands so a test
// can import and run them without this file's side effect of parsing
// argv and taking over the process — which is why the CLI used to be the
// one module in the tree at 0% coverage. See #399.
// ─────────────────────────────────────────────────────────────────────

import { runMain } from "citty"
import { consola } from "consola"
import {
  CliError,
  convertCommand,
  inspectCommand,
  mainCommand,
  validateCommand,
} from "./cli/commands"

// The commands throw CliError rather than calling process.exit, so they
// stay testable. Translating that back into an exit code is this file's
// job: citty's runMain would otherwise print a raw stack trace for it,
// and a CliError's message is already written for a user to read.
for (const sub of [convertCommand, inspectCommand, validateCommand]) {
  const command = sub as { run?: (ctx: never) => unknown }
  const run = command.run
  if (!run) continue
  command.run = async (ctx: never) => {
    try {
      return await run(ctx)
    } catch (error) {
      if (!(error instanceof CliError)) throw error
      consola.error(error.message)
      process.exit(1)
    }
  }
}

runMain(mainCommand)
