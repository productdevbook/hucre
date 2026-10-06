# V2 architecture audit

The target is one model vocabulary, shared operations, fewer duplicated
paths, and independent spreadsheet fixtures in a maintainable test layout.
Passing one subsystem's tests does not finish this work. This audit records
the evidence for the whole target and the concrete gaps found in code.

| Requirement                                   | Evidence                                                                                                                                                                                                | State                                                                                                                                                                                        |
| --------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| One readable/writable workbook model          | `WorkbookInput` and `SheetInput` derive from read types; writers normalize in `_write-model.ts`; model, parity and clone registers cover fields                                                         | Implemented                                                                                                                                                                                  |
| One authoring state                           | Builders assemble one state; `set()` and named methods share it; writer options are separate from workbook data                                                                                         | Implemented                                                                                                                                                                                  |
| One sparse metadata representation            | Numeric `CellStore` is used by readers, writers, builders, templates, exports and editing; bounds/count/clone/capacity regressions exercise it                                                          | Implemented                                                                                                                                                                                  |
| Shared structural editing                     | Insert/delete operations share `editAxis` and the A1 rewriter; invalid geometry is checked before mutation; independent invoice edits verify both save paths                                            | Implemented for insert/delete; other geometry APIs still need an audit                                                                                                                       |
| Dense reads enforce limits before expansion   | XLSX/XLS/XLSB guard growing boxes; ODS, HTML and text normalization share `_grid` checks; ODS repeats keep metadata and respect `maxRows` before expansion                                              | Implemented; small regressions and a fresh-process default-bound benchmark verify the previously missed paths                                                                                |
| Consistent values, errors and streaming       | `CellInput`, `CellError`, reader/stream options are shared; search/replacement/templates use `_sheet-values`; partial metadata merges preserve inline fields                                            | Implemented for value edits and inline cells across seven writer paths; streaming headers, object projection, XLSX columns and physical bounds are verified; ODS layout still needs an audit |
| Shared object/export projection               | `_sheet-grid` and `_objects` provide effective values, physical sparse row order, output bounds, header disambiguation and own-property records; independent invoice exports and regressions cover them | Implemented for read sheets and buffered authoring `rows`/`data`/`cells`                                                                                                                     |
| Shared authoring row-source boundary          | `_sheet-input` resolves object fields/headers/links once; buffered XLSX, ODS, text and saving share it; ODS applies column defaults; pivots use effective worksheet headers/caches                      | Implemented; ambiguity rejects before output and text object-data expansion is bounded before value lookup                                                                                   |
| Independent workbook examples                 | Twelve authored business XLSX files, SHA-256 manifest, literal expectations and native/third-party producer corpus; integration suites cover read, stream and save paths                                | Implemented; generated examples are not claimed to be Microsoft Excel-authored                                                                                                               |
| Shared test infrastructure                    | Integration suites use shared loading/model/XML helpers; ODS package scaffolding is shared in four existing suites and new regressions; every test is typechecked                                       | Partly consolidated; remaining ZIP/XML helpers still need an inventory                                                                                                                       |
| Published API and runtime verification        | Export/type registers, installed tarball/CLI, all ten entry points, Node 24/LTS, Bun, Deno and non-UTC CI                                                                                               | Implemented for those runtimes and Chromium 154/module Web Worker; the committed browser harness verifies ESM, crypto and native model/buffer transfer                                       |
| Truthful migration and fidelity documentation | `MIGRATION.md`, `docs/PARITY.md`, producer provenance and `onDrop` capability decisions                                                                                                                 | Maintained with each change; recheck after remaining work                                                                                                                                    |

## Inline spreadsheet cell audit evidence

`resolveCellInput` carries values/caches and full metadata through buffered,
incremental, single-stream and multi-stream XLSX plus all three ODS paths.
XLSX links/comments share sheet-tail collectors and package writers;
shared/array/dynamic formulas and checkboxes emit their required parts.
Cell and run/comment fonts now share property serialization. ODS uses the
buffered cell serializer and an exhaustive top-level field-drop register;
true streaming remains unstyled and reports omitted run fonts/styles.
The 13-field output register and reader limitations are in `PARITY.md`.

The 102 unit cases include explicit caches of six types, null precedence,
sparse Date caches, links, rich text, related ZIP parts, rollover header
snapshots, physical sheet relationships and unsupported ODS fields. Of
these, 78 fail on V2 `adcf5e5` before the fix. A shared test adapter inventory
also verifies the independently exported invoice through all seven paths,
with literal values and the cached total 210. Existing native/third-party
producer fixtures remain in the full gate. Node 24 throughput/RSS samples
for ordinary streamed/incremental output are recorded in `bench/README.md`;
they do not establish a speedup or measure metadata-heavy sheets.

## Remaining work found by inspection

1. Review remaining geometry helpers (`moveRows`, `copyRange`, hiding,
   grouping, sorting) for index validation and mutation atomicity. Inventory
   position-bearing sheet fields (chart/pivot locations, filter columns,
   print areas and drawing anchors) and record maintenance decisions. Preserve documented limits on
   workbook-level references. Value-edit consistency is verified by typed/cache/sparse regressions and independent invoice saves.
2. Inline spreadsheet cells now share resolution and serialization,
   including explicit caches, rich text and XLSX metadata parts. Typed
   regressions and independent invoice saves cover all seven writer paths.
   Streaming headers/object projection, integer rollover caps, physical
   row/column limits, frozen panes, XLSX column defaults/collapsed groups and
   buffered/incremental auto-width now share rules. The 83 new regressions
   include 65 failures against `e2f64a0`; this removes 20 library source
   lines. Six independent invoice object-save
   paths preserve literal values and the cached total 210. True XLSX streaming
   explicitly rejects automatic sizing without a fixed width. Audit ODS
   column layout (buffered widths, hiding/grouping and automatic sizing),
   nested styles and auto-width's approximate number-format scanning next.
   Keep true ODS streaming's style limit and current-sheet XLSX link/comment
   retention explicit; nested ODF style and hyperlink capabilities still
   need an inventory.
3. Consolidate repeated ODS/ZIP fixture scaffolds into `test/support/`;
   retain raw regression inputs and independently authored expectations.
4. Recheck published runtime claims and perform the final whole-architecture
   audit after the remaining work. The built package now has direct
   Chromium 154/module Web Worker evidence from `scripts/browser-smoke`,
   including crypto and native workbook cloning/transfer. Other browser
   engines and deployed edge environments are not claimed verified.

## Completion gate

The remaining items require regression evidence before fixes, the full
`bun run test` gate, unchanged coverage/size thresholds and installed-package
verification. Verify independent corpus behavior after model/reader/edit
changes. Keep the V2 branch current with main before each V2-only PR.
Merge reviewed, passing work into V2; the parent V2-to-main PR remains a
separate release decision. Completion requires every item above to be
verified against the current V2 tree, not an earlier green run.
