# CLAUDE.md

Notes for anyone — human or agent — writing code in this repo.

## Node 24 is the floor

**Do not write code that has to work on Node below 24.** No polyfills, no
capability shims for older runtimes, no "this exists from 20 onward"
branches. If a Web API is in Node 24, use it.

`package.json` says `"engines": { "node": ">=24" }`, CI runs 24 and
`lts/*`, and the release workflow publishes from 24. Those four have to
stay in step — if you raise one, raise them all.

The repo used to claim Node 18+. Nothing tested it, and when a matrix was
finally added it turned out `writeXlsxStream` did not work there at all:
Node 18's `CompressionStream` has no `deflate-raw`. A floor nobody
verifies is not a floor.

## Still no Node APIs in the library

Separate rule, and it has not changed. `src/` outside the CLI uses **Web
APIs only** — no `node:` imports, no `process`, no `Buffer`.
`tsconfig.json` sets `"types": []` so reaching for one is a compile error
rather than something that quietly works on your machine.

That is what makes the Deno, Bun, browser and Workers claims true. The
CLI is the exception and is checked by `tsconfig.cli.json`; tests that
read from disk belong in that project's include list.

Raising the Node floor does not licence `node:fs` in a reader.

## The rest

`CONTRIBUTING.md` has the working conventions — the gate (`pnpm test`),
the registers that break when a field is added and a copy forgets it, and
the two habits this codebase holds to: a test that fails _before_ the fix,
and a reason written into the code rather than only into the PR.

`docs/PARITY.md` is the statement of what hucre reads versus what it
writes, and it ends with "Anything else is a bug". If a change adds a
loss, that page has to say so in the same PR.

`bench/` has reproducible measurements. If a change is about speed or
memory, run it and put the numbers in the PR — one process per
measurement, because `maxRSS` is a high-water mark for the whole process.
