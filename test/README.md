# Tests

`bun run test` is the complete gate: lint, formatting, all three TypeScript
projects and Vitest. Every test file is collected explicitly from `test/`;
worktrees, generated outputs and examples cannot become accidental suites.

| Location            | Purpose                                                                |
| ------------------- | ---------------------------------------------------------------------- |
| `test/*.test.ts`    | Focused unit tests, regressions, API and model contracts               |
| `test/integration/` | Committed files from independent spreadsheet producers                 |
| `test/support/`     | Fixture loading, readable model projection and small raw-XML scaffolds |
| `test/fixtures/`    | Microsoft Excel, openpyxl, ExcelJS, SheetJS and LibreOffice corpus     |
| `examples/`         | Downloadable business scenarios, authored expectations and checksums   |

```sh
bun run test:unit
bun run test:integration
bun run coverage
bun run verify:package
bun run size
```

Unit fixtures should expose the bytes relevant to a regression. Use
`support/xlsx.ts` for raw cell XML or a one-part mutation instead of
copying a ZIP rebuild loop. Independent producer files exercise complete
imports and both save paths, so a reader and writer cannot validate the
same mistaken interpretation against each other.

Use `support/ods.ts` for ODF namespaces, content/meta wrappers and minimal
ZIP packages. Keep cell/repeat/formula XML visible in the regression;
the support code supplies the envelope, not the expected spreadsheet data.

Expected values must be authored separately from hucre's output. Change
goldens only after checking the source workbook or authored scenario.
Adding an example requires a scenario and a checksum; unregistered files,
missing files and modified bytes fail the integration suite.

The library is checked without Node types in `tsconfig.json`; the CLI has
`tsconfig.cli.json`; **all** tests use `tsconfig.test.json`. A new disk-backed
test needs no synchronized include/exclude lists. Model parity registers
remain exhaustive and clone tests still exercise `structuredClone`.
