# @bilig/benchmarks

Benchmark scenarios, metrics helpers, and workbook generators for bilig.

## Install

```bash
npm install @bilig/benchmarks
```

## Package entrypoints

- ESM: `./dist/index.js`
- Types: `./dist/index.d.ts`
- Corpus: `./dist/workbook-corpus.js`

## WorkPaper baseline

The repo tracks a checked-in WorkPaper benchmark artifact at
`packages/benchmarks/baselines/workpaper-baseline.json`.

Refresh or validate it with:

```bash
pnpm workpaper:bench:generate
pnpm workpaper:bench:check
```

## WorkPaper vs HyperFormula benchmark

The repo tracks one checked-in competitive benchmark artifact at
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

This is the competitive matrix. It includes the original control workloads, the broader expanded
workload set against HyperFormula `3.2.0`, and leadership workloads that must be labeled
unsupported instead of silently omitted.

Refresh or validate it with:

```bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

For exploratory local refreshes on larger machines, shard independent workloads across worker
processes:

```bash
pnpm workpaper:bench:competitive:generate -- --jobs 8
```

Each workload still runs its warmup and measured samples sequentially; `--jobs` only parallelizes
different workloads and the checked artifact remains ordered canonically.

## Other headless competitor lanes

The headless leadership scorecard also consumes scoped competitor artifacts for public headless
spreadsheet engines:

```bash
pnpm workpaper:bench:truecalc:generate
pnpm workpaper:bench:univer:generate
pnpm workpaper:bench:xlsx-calc:generate
pnpm headless:performance:generate
```

`workpaper-vs-univer.json` uses Univer's documented `@univerjs/preset-sheets-node-core` Node preset
and times public Facade API edits through formula recalculation completion. The leadership
scorecard stays `active-not-achieved` until every direct workbook-wide competitor lane covers the
broader comparable workload set.

## Named giant-workbook corpus

The package now ships deterministic named workbook cases for giant-data restore and warm-start
contracts:

- `dense-mixed-100k`
- `dense-mixed-250k`
- `wide-mixed-250k`
- `analysis-multisheet-100k`
- `analysis-multisheet-250k`

Use `buildWorkbookBenchmarkCorpus(...)` to materialize an exact-size workbook snapshot and stable
viewport metadata for CI and perf harnesses.

This package is part of the [bilig](https://github.com/proompteng/bilig) monorepo.
