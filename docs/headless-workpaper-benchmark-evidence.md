# Headless WorkPaper Benchmark Evidence

Status: public evidence note for `@bilig/headless`

This note keeps the public performance claim auditable from checked-in repo
artifacts instead of README copy alone.

## Current Artifact

The primary workbook-wide decision artifact is
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

The additional scalar formula-engine comparison artifact is
`packages/benchmarks/baselines/workpaper-vs-truecalc.json`.

The additional limited workbook-wide comparison artifact is
`packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json`.

The goal-tracking scorecard for broad headless-engine performance leadership is
`packages/benchmarks/baselines/headless-performance-leadership-scorecard.json`.
It intentionally stays `active-not-achieved` until the checked-in evidence covers
at least two workbook-wide direct headless spreadsheet engines across the full
eligible workload set and every comparable workload wins both mean and p95
latency. Scalar formula-engine lanes and partial workbook-wide lanes are tracked
as useful evidence, but they do not satisfy broad coverage alone.

Current checked-in metadata:

- generated at `2026-05-22T03:35:00.425Z`
- benchmark sampling: `200` measured samples after `2` warmup samples
- WorkPaper package: `@bilig/headless` `0.41.0`
- comparison engine: HyperFormula `3.2.0`, local checkout commit
  `9a510a2acb97c3d3490f9e3b9e961a1c4a98b9ad`, GPL-v3 license key
- scalar formula comparison engine: TrueCalc `0.6.4`, `7` comparable scalar
  workloads via `@truecalc/core`
- limited workbook-wide comparison engine: xlsx-calc `0.9.2`, `16` comparable
  recalculation workloads covering aggregate, aggregate-2d, overlapping-range,
  exact lookup, approximate lookup, formula-chain, fanout, range-stat, and
  cross-sheet families

## What The Claim Is

The current scorecard is not a blanket performance-leadership claim. A fresh
checked-in run shows WorkPaper leading HyperFormula in aggregate across the
directly comparable workbook-wide headless spreadsheet-engine workloads, with
visible holdouts. The current checked-in artifact records `99/100` mean-latency wins:

| Lane    | Comparable Workloads | WorkPaper Mean Wins | HyperFormula Mean Wins |
| ------- | -------------------: | ------------------: | ---------------------: |
| Overall |                `100` |                `99` |                    `1` |
| Public  |                 `73` |                `73` |                    `0` |
| Holdout |                 `27` |                `26` |                    `1` |

The overall directional mean-ratio geomean is `0.29358029192125573`. The overall
directional p95-ratio geomean is `0.30456268903585865`. Ratios below `1.0` mean
WorkPaper is faster for that metric.

The current worst mean row is `structural-append-formula-rows-large`, with a
mean ratio of `1.0047527772334053`. The current worst p95 row is
`structural-append-formula-rows-large`, with a p95 ratio of
`1.19563861123766`. The
headless leadership scorecard
currently records `96/100` workloads winning both
mean and p95 against HyperFormula.

It is also not a blanket "fastest against every formula evaluator" claim. The
TrueCalc scalar lane currently reports `7/7` WorkPaper mean+p95 wins, with a
directional mean-ratio geomean of `0.1836987971377073`. That lane stays in the
leadership scorecard as limited scalar coverage rather than proof of workbook
dependency-graph, range, or structural-edit leadership.

The xlsx-calc lane is a direct workbook-wide recalculation comparison for the
formula families both engines can evaluate equivalently. It currently reports
`16/16` WorkPaper mean+p95 wins with a directional mean-ratio geomean of
`0.09196510101241956`, but it covers only the formula families xlsx-calc
supports equivalently, so the scorecard treats it as workbook-wide-limited
coverage rather than proof of blanket leadership.

## How To Read The p95 Evidence

The `99/100` count is about mean latency: for each comparable workload row,
WorkPaper's average measured time is lower than HyperFormula's average measured
time. Mean wins are useful because they summarize the normal cost of each
workload, and the current scorecard keeps p95 holdouts visible.

Each p95 row asks a different question: "near the slow end of this workload's
sample set, which engine was faster?" A single row can lose on p95 even when its
mean wins, because a small number of slower samples can move the tail without
moving the average enough to flip the mean result.

The p95 geomean is an aggregate across the per-workload p95 ratios. Read the
current result as: WorkPaper leads the overall mean and p95 aggregate while
retaining four comparable p95 holdouts that still need implementation work.

## What Is Measured

Scorecard-eligible families cover:

- workbook build and rebuild paths
- runtime restore from snapshot
- sheet lifecycle and named-expression changes
- cross-sheet scalar and aggregate recalculation
- dirty execution after single edits, chains, fanout, mixed frontiers, and
  formula edits
- batch edits, suspended batches, and undo-including batches
- structural row and column inserts, deletes, and moves
- dense and sparse range reads
- 2D, overlapping, sliding-window, and conditional aggregation
- exact lookup, INDEX/MATCH, INDEX reference, approximate lookup, after-write
  lookup, and text lookup

The scorecard excludes the `config-toggle` control family and `dynamic-array`
leadership-only family from the directly comparable win count.

## How To Verify

Check that the committed artifact still has the expected workload coverage and
shape:

```bash
pnpm workpaper:bench:competitive:check
pnpm workpaper:bench:truecalc:check
pnpm workpaper:bench:xlsx-calc:check
pnpm headless:performance:check
```

Regenerate timing evidence only when intentionally refreshing the benchmark
artifact:

```bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
pnpm workpaper:bench:truecalc:generate
pnpm workpaper:bench:truecalc:check
pnpm workpaper:bench:xlsx-calc:generate
pnpm workpaper:bench:xlsx-calc:check
```

Do not change workload sizes, sampling, scoring, or definitions to preserve a
claim. If a rerun moves a row red, update the artifact, update this note, and
fix the production engine path rather than hiding the loss.

If a workload family is missing, a row looks too synthetic, or the p95 wording
is still too broad, use the public benchmark critique thread:
<https://github.com/proompteng/bilig/discussions/340>.
