# bench

Reproducible measurements for the claims in the README, and for anyone
changing a hot path.

```bash
pnpm build     # bench runs against dist/, which is what users get
pnpm bench     # every scenario, one process each
```

Or one at a time:

```bash
node bench/write.mjs writeXlsxStream 300000
node bench/read.mjs streamXlsxRows high-cardinality
```

## Why one process per measurement

`process.resourceUsage().maxRSS` is a high-water mark for the **whole
process**. Measure two things in one run and the second inherits the
first's peak — which is how a streaming writer can appear to use 800 MB.
`pnpm bench` forks a child per scenario for that reason, and the numbers
below were taken that way.

## The scenarios

`write.mjs` — 12 columns of mixed text, number and date, through each of
the three XLSX write paths.

`read.mjs` — the same sheet read back four ways, against **two** fixtures
that differ in one line:

```js
c % 3 === 0 ? `text ${i}-${c}`        // high cardinality: ~400k distinct strings
c % 3 === 0 ? WORDS[(i + c) % 10]     // low cardinality: 10 distinct strings
```

That one line is the difference between 90 MB and 556 MB in
`streamXlsxRows`, which is the thing worth knowing about the streaming
reader: its peak tracks the number of **distinct strings**, not the number
of rows, because `xl/sharedStrings.xml` has to be read up front.

## Numbers to compare against

Taken on Linux, Node 24, 100,000 rows × 12 columns. Yours will differ;
the ratios are the point.

| write              |     time |  peak RSS |
| ------------------ | -------: | --------: |
| `writeXlsxStream`  | 2,230 ms | **92 MB** |
| `XlsxStreamWriter` | 2,877 ms |    604 MB |
| `writeXlsx`        | 3,186 ms |    758 MB |

At 300,000 rows `writeXlsxStream` grows to 126 MB and `XlsxStreamWriter`
to 1,515 MB — flat against linear, which is the promise the README makes.

| read                         |      high cardinality |      low cardinality |
| ---------------------------- | --------------------: | -------------------: |
| `readXlsx`                   |     2,721 ms / 662 MB |    1,896 ms / 392 MB |
| `readXlsx` + `maxRows: 1000` |     1,953 ms / 656 MB |    1,150 ms / 221 MB |
| `streamXlsxRows`             | 2,446 ms / **556 MB** | 1,592 ms / **90 MB** |

Two things fall out of that table and are worth keeping in view when
changing the readers:

- `streamXlsxRows` is constant-memory in rows and **linear in distinct
  strings**.
- `maxRows` bounds the output, not the work: on the high-cardinality
  fixture it saved 28% of the time and 1% of the memory.
