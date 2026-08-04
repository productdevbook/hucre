# Migrating to v1

v1 is where hucre's public API becomes a stability commitment. Getting there meant fixing things that were wrong, inconsistent, or documented-but-inert — several of which could not be fixed afterwards without a major bump.

This guide lists every change that can affect existing code. Most projects will need to touch nothing; the ones that do are marked.

If a change here breaks something and the reason isn't clear, please open an issue — a migration that needs a guess is a migration that needs better docs.

## At a glance

| Change                                                                                               | Affects you if…                                                                                   |
| ---------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| [`writeXlsxStream` argument order](#writexlsxstream-takes-rows-first)                                | you call `writeXlsxStream`                                                                        |
| [`headerRow` is 0-based everywhere](#headerrow-means-one-thing-now)                                  | you pass `headerRow` to `validateWithSchema`                                                      |
| [`hasHeaderRow` on the exporters](#hasheaderrow-on-tohtml-and-tomarkdown)                            | you pass `headerRow` to `toHtml` / `toMarkdown`                                                   |
| [`streamCsvRows` header handling](#streamcsvrows-matches-parsecsv)                                   | you call `streamCsvRows({ header: true })`                                                        |
| [`readObjects` / `sheetToObjects` return shape](#readobjects-and-sheettoobjects-return-data-headers) | you call either                                                                                   |
| [`HucreError`](#deftererror-is-now-hucreerror)                                                       | nothing — the old name still works                                                                |
| [`readNdjsonStream`](#readndjsonstream-is-now-streamndjsonrows)                                      | nothing — the old name still works                                                                |
| [Sheet names are validated](#sheet-names-are-validated-on-write)                                     | you write sheet names with `: * ? / \ [ ]`, over 31 chars, or duplicates                          |
| [`parseJson` on a multi-sheet document](#parsejson-rejects-a-workbook-instead-of-mangling-it)        | you pass `workbookToJson` output of ≠1 sheet back into `parseJson`                                |
| [Removed dead API](#removed-api-that-never-did-anything)                                             | you reference `ReadResult`, `WORKER_SAFE_FUNCTIONS`, `isoDates`, `WriteSheet.threadedComments`, … |
| [Files that were silently corrupt](#files-that-were-silently-wrong)                                  | you round-trip workbooks, or print them                                                           |

---

## `writeXlsxStream` takes rows first

Every other writer in the library is `write*(data, options)`. This one was inverted.

```diff
- writeXlsxStream({ name: "Export", columns }, rows)
+ writeXlsxStream(rows, { name: "Export", columns })
```

TypeScript flags every call site.

## `headerRow` means one thing now

It used to mean four different things. `validateWithSchema` was the last 1-based holdout:

```diff
- validateWithSchema(rows, schema, { headerRow: 2 })   // 1-based: the second row
+ validateWithSchema(rows, schema, { headerRow: 1 })   // 0-based: the second row
```

**The subtle one.** `headerRow: 0` used to mean _"there is no header row"_ — 1-based numbering had no other way to say it, so two concepts shared one value. They are separate now:

```diff
- validateWithSchema(rows, schema, { headerRow: 0 })   // meant "no header row"
+ validateWithSchema(rows, schema, { headerRow: -1 })  // says it explicitly
```

`headerRow: 0` now means _the first row is the header_, like everywhere else in the library.

If you rely on the **default**, nothing changes: the old default of `1` and the new default of `0` both mean the first row.

The options type is now exported as `SchemaValidateOptions`, so you can build one in a typed variable.

## `hasHeaderRow` on `toHtml` and `toMarkdown`

`headerRow` was a **boolean** on these two, while it is a row index everywhere else. It is renamed:

```diff
- toHtml(sheet, { headerRow: true })
+ toHtml(sheet, { hasHeaderRow: true })
```

The old spelling still works for one major and is marked `@deprecated`.

## `streamCsvRows` matches `parseCsv`

The two took the same `CsvReadOptions` and behaved differently in four places. `header: true` was the one that changed the data:

```diff
- streamCsvRows(input, { header: true })                        // dropped the header row
+ streamCsvRows(input, { header: true, skipHeaderRow: true })   // same output
```

`header: true` now only _marks_ the header row — it is still yielded, and used to name columns for `transformValue` — which is what `parseCsv` has always done. `skipHeaderRow` is the explicit way to consume it.

Also newly honoured, having previously been silently ignored: `onRow`, `transformValue`, and `fastMode`. If you passed any of those to `streamCsvRows` before, they now actually run. `fastMode` in particular changes the parsed fields, because it skips quote handling by design.

## `readObjects` and `sheetToObjects` return `{ data, headers }`

They returned a bare array while every other `*Objects` reader returned `{ data, headers }`.

```diff
- const rows = await readObjects(buffer)
+ const { data: rows } = await readObjects(buffer)
```

Three smaller behaviour changes come with `readObjects` joining the family — in each case it now matches its siblings rather than diverging:

- empty-string header keys are **kept** (only `readObjects` dropped them)
- fully empty rows are **skipped** by default
- a missing sheet **throws** `ParseError` instead of returning `[]`

## `DefterError` is now `HucreError`

`instanceof DefterError` was the documented catch-all for every error the library throws — in a package called `hucre`.

```diff
- catch (e) { if (e instanceof DefterError) … }
+ catch (e) { if (e instanceof HucreError) … }
```

**No action required.** `DefterError` is still exported and is the _same class object_, so `instanceof` behaves identically. It is marked `@deprecated`.

One visible difference: `error.name` now reports `"HucreError"` rather than `"DefterError"`. If you match on that string, update it.

Separately, the `ValidationError` **interface** — one row/column schema failure — is renamed `SchemaValidationIssue`. The `ValidationError` **class** keeps its name. The `ValidationErrorType` alias is gone.

## `readNdjsonStream` is now `streamNdjsonRows`

So every streaming reader reads `stream*Rows`. **No action required** — the old name is still exported as a deprecated alias of the same function.

## Sheet names are validated on write

Previously anything was written verbatim, producing files Excel opens with "unreadable content" and no warning. `writeXlsx`, `writeOds`, `XlsxStreamWriter`, and `writeXlsxStream` now throw `InvalidArgumentError` before producing any bytes for:

- empty names, or names longer than 31 characters
- `[ ] : * ? / \`
- a leading or trailing apostrophe (it breaks quoted range references)
- the reserved name `History`
- duplicates — compared **case-insensitively**, because Excel does

It throws rather than sanitizing on purpose: truncating a name or stripping its colons produces a workbook whose sheets are not the ones you asked for, and any range reference built against the original names would then dangle.

If you generate sheet names from user data — report titles, date ranges, file names — sanitize before calling.

## `parseJson` rejects a workbook instead of mangling it

`workbookToJson` emits a bare array for a one-sheet workbook and
`{ "Sheet1": [...], "Sheet2": [...] }` for anything else. `parseJson` only
ever unwrapped a _single_ array-valued property, so the multi-sheet shape fell
through to "treat the object as one row" and came back as one row whose cells
were JSON-stringified sheets — silently, and only once a workbook grew a second
sheet.

It now throws `ParseError`, and there are three ways forward:

```diff
- parseJson(workbookToJson(wb))                    // one row of nonsense
+ jsonToWorkbook(workbookToJson(wb))               // every sheet, as a Workbook
+ parseJson(workbookToJson(wb), { rowsAt: "S1" })  // one named sheet as a table
+ parseJson(json, { rowsAt: "" })                  // the old single-row reading
```

The guard is deliberately narrow: it only fires when **every** property is an
array of plain objects and there are at least two of them. `{ a: [1, 2], b: [3, 4] }`
is a row with two list-valued columns and still reads as one row.

Two additions come with it, neither of them breaking:

- **`jsonToWorkbook`** reads either shape back into a `Workbook`, so the round
  trip no longer depends on the sheet count.
- **`workbookToJson(wb, { shape: "sheets" })`** always emits the keyed object.
  The default `"auto"` keeps today's count-dependent shape.

## Nested JSON can be rebuilt: `unflattenRow`

`parseJson` flattens `{user: {name}}` to a `"user.name"` column by default, and
nothing reversed it, so `parseJson` → `writeJson` permanently destroyed the
nesting. There is now an inverse, and the writers take it as an **opt-in**:

```diff
- writeJson(parseJson(text).data)                    // { "user.name": "Ada" }
+ writeJson(parseJson(text).data, { unflatten: true }) // { user: { name: "Ada" } }
```

Opt-in rather than on by default because `writeJson` takes any flat row set,
most of which never went through `flatten` — and spreadsheet headers contain
dots routinely (`Q1.2024`, `v1.2`). Nesting those by default would be a new
silent mangling in the fix for one. `writeNdjson` and `NdjsonStreamWriter` take
the same option; `unflattenRow` / `unflattenRows` are exported directly.

Two things it does not undo, because they are not recoverable from the flat
form: a primitive array joined into `"1, 2"` stays a string, and a key that
contained a literal dot comes back nested.

## Dates come back from JSON

The CSV reader has always inferred ISO 8601 dates under `typeInference`. The
JSON reader had no equivalent, so a `Date` written by `writeJson` came back a
string. The same option name now does the same job on both:

```diff
- parseJson(writeJson([{ at: new Date() }])).data[0].at  // string
+ parseJson(writeJson([{ at: new Date() }]), { typeInference: true }).data[0].at // Date
```

Off by default in both readers, and it accepts exactly the instants CSV
accepts — the rule is one function now, shared by `parseCsv`, `streamCsvRows`,
`parseJson`, `parseNdjson` and `streamNdjsonRows`. It infers **only** dates for
JSON: numbers and booleans are already typed there, so `"007"` stays a string.

## Removed API that never did anything

Each of these was exported or declared and had no effect. v1 would have frozen them permanently.

| Removed                                       | Why                                                          |
| --------------------------------------------- | ------------------------------------------------------------ |
| `ReadResult<T>`                               | no function ever returned it                                 |
| `StreamReadOptions`, `StreamWriteOptions`     | zero references anywhere                                     |
| `WORKER_SAFE_FUNCTIONS`                       | 40 entries against ~125 exports; every export is worker-safe |
| `ReadOptions.headerRow`, `ReadOptions.schema` | honoured by no reader                                        |
| `CsvReadOptions.lineSeparator`, `.encoding`   | never read (`CsvWriteOptions.lineSeparator` is fine)         |
| `CsvReadOptions.schema`                       | zero references under `src/csv/` — no CSV reader validated   |
| `isoDates` on the JSON writers                | see below                                                    |
| `WriteSheet.threadedComments`                 | typed and accepted; no writer ever produced the part         |

`CsvReadOptions.schema` never validated anything: `parseCsv` returned every row exactly as parsed whatever you passed. Validate the parsed rows with `validateWithSchema`, which is what the option looked like it was doing.

`isoDates` is worth explaining because it _looked_ like it worked. `JSON.stringify` calls `Date.prototype.toJSON` **before** consulting the replacer, so a replacer testing `value instanceof Date` is never reached — `isoDates: false` produced byte-identical output. Dates still serialize as ISO strings, exactly as before.

`RoundtripWorkbook` no longer exposes `_rawEntries`, `_modifiedParts`, `_contentTypes`, or `_rootRels`. Pass the object from `openXlsx` straight to `saveXlsx`, as documented. One consequence: `saveXlsx({ ...workbook })` no longer works, because spreading drops the internal state.

## Files that were silently wrong

No action needed — but if you have workbooks produced by 0.6.x, they may be affected.

- **`openXlsx` → `saveXlsx` dropped state it understood**: split panes, page breaks, sparklines, text boxes, background images, Excel 2024 checkboxes, and **workbook protection** — a structurally locked workbook came back unlocked.
- **Setting any `pageSetup` turned off printed gridlines and row/column headings.** `showGridLines` and `showRowColHeaders` were inert in both directions and now work. `horizontalCentered` / `verticalCentered` moved to `<printOptions>` where ECMA-376 puts them; they were written where Excel ignores them. Files from earlier versions are still read correctly.
- **`fillTemplate` could put a function into a cell.** A placeholder named `toString`, `constructor`, `valueOf`, or `__proto__` resolved against `Object.prototype`.
- **Rich-text colours parsed differently** in `sharedStrings.xml` than in an inline `<is>` — `"FF0000"` became `"0000"`.
- **A workbook redefining a built-in number-format id** could be read as a date and then formatted numerically.

## Reading untrusted files

Several inputs could previously hang or kill the process. If you accept uploads, these are now bounded and throw typed errors instead:

- an unbounded `<col>` range could hang `readXlsx` **forever** from a 1.4 KB file
- two entirely legal cells at opposite corners caused a **fatal, uncatchable OOM** — readers return a dense grid, so a sheet costs its bounding box, not its cell count
- ODS `<text:s text:c>` and `number:decimal-places` could allocate gigabytes
- `read(response.body)` had **no size ceiling** at all; there is now `ReadOptions.maxInputBytes`, defaulting to 1 GiB
- the zip-bomb cap covered only 3 of 4 decompression paths
- `fromHtml` could hang on a hostile `rowspan` × `colspan`

Very large legitimate files may now hit a limit that used to be absent. Every one is a named constant in `src/limits.ts`, and the errors say which.

## Also worth knowing

- **The CLI works now.** `dist/cli.mjs` imported packages that were not runtime dependencies, so `npx hucre` failed on every clean install. It is bundled, and CI verifies the packaged tarball.
- **`hucre/xlsx` and `hucre/ods` export what the README documents.** `readXlsxObjects`, `readOdsObjects`, `streamOdsRows`, and others were root-only despite documented subpath imports.
- **Format entry points export their own types**, so `import type { WriteSheet } from "hucre/xlsx"` works without a second import from the root.
- **`streamOdsRows` takes options and a `ReadableStream`.** It previously accepted neither.
- **`toJson`'s `"arrays"` and `"columns"` formats are write-only**, and now say so. They are handoffs to a charting library or a dataframe; only `"objects"` has a reader. Export in `"objects"` if the JSON has to come back into hucre.
- **`writeCsvStream` exists** — constant-memory CSV writing, the counterpart to `writeXlsxStream`.
- **A CSV write option now has a way back in.** `escapeFormulae: true` prefixed `= + - @ | \t \r \n \0` values with `'` and nothing removed it, so a round trip through hucre turned `-5` into `'-5` permanently. `parseCsv` and `streamCsvRows` take `unescapeFormulae: true`, which drops that `'` — and only where the writer would have added one, so `'quoted'` is left alone. `nullValue` and a custom `dateFormat` remain one-way by decision; both are documented as such on the option.
- **The streaming CSV writers escape formulae too.** `escapeFormulae` was honoured by `writeCsv` alone — `CsvStreamWriter` and `writeCsvStream` ignored it silently, which for an injection escape meant the protection you asked for was simply absent. If you passed it to either, their output changes.
- **`parseCsv` honours `skipHeaderRow`.** It was implemented in `streamCsvRows` and ignored here, so the same option on the same options type behaved two ways. `parseCsv(input, { header: true, skipHeaderRow: true })` now drops the header row, and `maxRows` counts data rows in both readers.
- **`CsvWriteOptions.comment` quotes values that would read as comments.** A value starting with `#`, written bare, is deleted by a reader configured with `comment: "#"` — the whole row, silently. Pass the same character to the writer and those values are quoted instead. Off by default; the output is unchanged unless you set it.
- **ZIP64 archives are readable**, and writable via `zip64: true`.
- **`hucre/ooxml` exists.** The low-level OOXML part parsers — `parseChart`, `parsePivotTable`, `parseSlicers`, `parseThemeColors` and friends — have a home of their own, explicitly outside the v1 stability commitment. They are still exported from the root, marked deprecated, so nothing breaks.
