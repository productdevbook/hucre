# Contributing

Thanks for looking. This is a short page — most of what you need is in the
code, and the rest is here.

## Getting set up

```bash
pnpm install
pnpm test        # lint + typecheck + vitest
```

`pnpm test` is the whole gate, and it is what CI runs. `pnpm dev` starts
vitest in watch mode.

## What a change looks like here

**A test that fails before your fix.** Not "a test that passes after" —
the two are different, and only the first proves the change does
something. Say so in the PR: _"three of the five fail against `main`"_ is
the most useful sentence you can write.

**A reason in the code, not just in the PR.** The comments in this
codebase explain _why_ a thing is the way it is, usually with the issue
number that caused it. A PR moves on; the comment stays with the line
someone will be reading in a year.

**No silently accepted options.** A typed field that is discarded is worse
than no field at all — that is why `WriteSheet.threadedComments` and
`CsvReadOptions.schema` were removed rather than left in place. If the
code cannot honour something, do not accept it.

## The registers

Several tests exist to fail when someone adds a field and forgets a place
that has to carry it. If one of these breaks, it is doing its job — fix
the thing it points at rather than the test:

| test                                  | guards                                                                                                  |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| `test/xlsx-write-read-parity.test.ts` | every `WriteSheet` / `WriteOptions` field either round-trips or is registered one-way **with a reason** |
| `test/parity-statement.test.ts`       | `docs/PARITY.md` and the README still describe the types as they are                                    |
| `test/clone-sheet-coverage.test.ts`   | `cloneSheet`, `cloneCell` and the worker serializer carry every field of `Sheet`, `Cell` and `Workbook` |
| `test/write-model.test.ts`            | `toWriteOptions` names every read-model field that has no write counterpart                             |
| `test/exports.test.ts`                | the public surface of each entry point                                                                  |

## Read/write parity

`docs/PARITY.md` is the statement of what hucre reads versus what it
writes. It ends with "Anything else is a bug", which is a commitment: if
your change adds a loss, the page has to say so in the same PR.

## Node 24 is the floor

Do not write code that has to work below it — see CLAUDE.md. `engines`,
the CI matrix and the release workflow all say 24, and they have to stay
in step.

## Platform neutrality

The library core uses Web APIs only — no `node:` imports, no `process`, no
`Buffer`. `tsconfig.json` sets `"types": []` so reaching for one is a
compile error rather than something that quietly works on your machine.
The CLI is the exception and is checked by `tsconfig.cli.json`.

Tests that read from disk belong in `tsconfig.cli.json`'s include list for
the same reason.

## Commits and PRs

Conventional commits — `fix(xlsx):`, `feat:`, `docs:`, `perf(csv):`. The
scope is the format or the module.

For the PR body, the shape that works is: what was wrong, shown with real
input and output; why the tests did not catch it; what landed; and what
you did **not** verify. That last one matters more than it sounds — most
of this library's output cannot be checked without Excel, and saying "I
did not open this in Excel" is more useful than implying you did.

## Performance claims

If a change is about speed or memory, measure it and put the numbers in
the PR, with the shape of the input and how many runs. One process per
measurement if you are quoting peak RSS — `maxRSS` is a high-water mark
for the whole process, so a second measurement in the same run inherits
the first one's peak.
