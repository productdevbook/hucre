# V2 architecture audit

The target is one model vocabulary, shared operations, fewer duplicated
paths, and independent spreadsheet fixtures in a maintainable test layout.
Passing one subsystem's tests does not finish this work. This audit records
the evidence for the whole target and the concrete gaps found in code.

| Requirement                                   | Evidence                                                                                                                                                                 | State                                                                                                         |
| --------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------- |
| One readable/writable workbook model          | `WorkbookInput` and `SheetInput` derive from read types; writers normalize in `_write-model.ts`; model, parity and clone registers cover fields                          | Implemented                                                                                                   |
| One authoring state                           | Builders assemble one state; `set()` and named methods share it; writer options are separate from workbook data                                                          | Implemented                                                                                                   |
| One sparse metadata representation            | Numeric `CellStore` is used by readers, writers, builders, templates, exports and editing; bounds/count/clone/capacity regressions exercise it                           | Implemented                                                                                                   |
| Shared structural editing                     | Insert/delete operations share `editAxis` and the A1 rewriter; invalid geometry is checked before mutation; independent invoice edits verify both save paths             | Implemented for insert/delete; other geometry APIs still need an audit                                        |
| Dense reads enforce limits before expansion   | XLSX/XLS/XLSB guard growing boxes; ODS, HTML and text normalization share `_grid` checks; ODS repeats keep metadata and respect `maxRows` before expansion               | Implemented; small regressions and a fresh-process default-bound benchmark verify the previously missed paths |
| Consistent values, errors and streaming       | `CellInput`, `CellError`, reader/stream options are shared; search/replacement/templates use `_sheet-values`; partial metadata merges preserve inline fields             | Implemented for search/replacement/templates; value/type/cache, sparse and both save paths are covered        |
| Independent workbook examples                 | Twelve authored business XLSX files, SHA-256 manifest, literal expectations and native/third-party producer corpus; integration suites cover read, stream and save paths | Implemented; generated examples are not claimed to be Microsoft Excel-authored                                |
| Shared test infrastructure                    | Integration suites use shared loading/model/XML helpers; ODS package scaffolding is shared in four existing suites and new regressions; every test is typechecked        | Partly consolidated; remaining ZIP/XML helpers still need an inventory                                        |
| Published API and runtime verification        | Export/type registers, installed tarball/CLI, all ten entry points, Node 24/LTS, Bun, Deno and non-UTC CI                                                                | Implemented for those runtimes; browser/Worker execution needs direct evidence                                |
| Truthful migration and fidelity documentation | `MIGRATION.md`, `docs/PARITY.md`, producer provenance and `onDrop` capability decisions                                                                                  | Maintained with each change; recheck after remaining work                                                     |

## Remaining work found by inspection

1. Review remaining geometry helpers (`moveRows`, `copyRange`, hiding,
   grouping, sorting) for index validation and mutation atomicity. Inventory
   position-bearing sheet fields (chart/pivot locations, filter columns,
   print areas and drawing anchors) and record maintenance decisions. Preserve documented limits on
   workbook-level references. Value-edit consistency is verified by typed/cache/sparse regressions and independent invoice saves.
2. Audit object/export projection against the same sparse/value model.
   A current-source probe with `Name`/`Ada` returns the record through
   `readObjects(bytes)`, but empty headers/data through
   `readObjects(bytes, { sparse: true })`. The object readers still pass
   only `sheet.rows` to projection, discarding populated sparse metadata.
   Verify object options, row indexes, limits and display/export consumers.
3. Consolidate repeated ODS/ZIP fixture scaffolds into `test/support/`;
   retain raw regression inputs and independently authored expectations.
4. Run the built package in a browser and Worker, then perform the final
   whole-architecture audit against current code, corpus and CI evidence.

## Completion gate

The remaining items require regression evidence before fixes, the full
`bun run test` gate, unchanged coverage/size thresholds and installed-package
verification. Verify independent corpus behavior after model/reader/edit
changes. Keep the V2 branch current with main before each V2-only PR.
Merge reviewed, passing work into V2; the parent V2-to-main PR remains a
separate release decision. Completion requires every item above to be
verified against the current V2 tree, not an earlier green run.
