# Headless WorkPaper Benchmark Evidence

Status: public evidence note for `@bilig/headless`

This note keeps the public performance claim auditable from checked-in repo
artifacts instead of README copy alone.

## Current Source Of Truth

The goal-tracking scorecard for broad headless-engine performance leadership is
`packages/benchmarks/baselines/headless-performance-leadership-scorecard.json`.
The short artifact name is `headless-performance-leadership-scorecard.json`.
It currently reports `achieved`: `100/100` comparable workloads win both mean
and p95 latency across `5` comparison engines and `2` workbook-wide engines.

Comparison engines: HyperFormula, TrueCalc, Univer, xlsx-calc, IronCalc Rust.

| Provider      | Coverage tier         | Mean+p95 wins | Mean wins |  p95 wins | Mean geomean ratio | p95 geomean ratio |      Unsupported |
| ------------- | --------------------- | ------------: | --------: | --------: | -----------------: | ----------------: | ---------------: |
| HyperFormula  | workbook-wide         |    `100/100` | `100/100` | `100/100` |         `0.2586x` |        `0.2807x` |              `0` |
| Univer        | workbook-wide         |    `100/100` | `100/100` | `100/100` |         `0.0028x` |        `0.0034x` |              `0` |
| IronCalc Rust | workbook-wide-limited |      `90/90` |   `90/90` |   `90/90` |         `0.1224x` |        `0.1658x` | `10` unsupported |
| xlsx-calc     | workbook-wide-limited |      `16/16` |   `16/16` |   `16/16` |         `0.0839x` |        `0.0786x` |              `0` |
| TrueCalc      | scalar-formula        |        `7/7` |     `7/7` |     `7/7` |         `0.1837x` |        `0.2359x` |              `0` |

Coverage tiers matter:

- `workbook-wide`: direct headless workbook engine comparison over the broad
  eligible workload suite.
- `workbook-wide-limited`: direct workbook comparison for the subset the other
  engine can represent fairly.
- `scalar-formula`: scalar formula API comparison only; it does not claim graph,
  range, structural-edit, or workbook lifecycle coverage.

## Artifact Inventory

- Primary workbook-wide artifact:
  `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.
- Additional workbook-wide artifact:
  `packages/benchmarks/baselines/workpaper-vs-univer.json`.
- Limited workbook-wide artifacts:
  `packages/benchmarks/baselines/workpaper-vs-ironcalc-rust.json` and
  `packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json`.
- Scalar formula-engine artifact:
  `packages/benchmarks/baselines/workpaper-vs-truecalc.json`.
- Public generated evidence:
  `docs/public-evidence.json`.

Current checked-in metadata:

- benchmark sampling: `200` measured samples after `2` warmup samples for the
  broad WorkPaper-vs-HyperFormula artifact
- comparison engine: HyperFormula `3.2.0`, local checkout commit
  `9a510a2acb97c3d3490f9e3b9e961a1c4a98b9ad`
- comparison engine: Univer `0.23.0`
- comparison engine: IronCalc Rust `0.7.1`, pinned through
  `ironcalc_base = "=0.7.1"` in a release-mode Rust sidecar
- comparison engine: xlsx-calc `0.9.2`
- comparison engine: TrueCalc `0.6.4`

## Primary Workbook-Wide Lane

The current checked-in WorkPaper-vs-HyperFormula artifact records WorkPaper
`100/100` mean-latency wins:

| Lane    | Comparable Workloads | WorkPaper Mean Wins | HyperFormula Mean Wins |
| ------- | -------------------: | ------------------: | ---------------------: |
| Overall |                `100` |               `100` |                    `0` |
| Public  |                 `73` |                `73` |                    `0` |
| Holdout |                 `27` |                `27` |                    `0` |

The artifact was generated at `2026-05-23T17:51:04.599Z`.

The overall directional mean-ratio geomean is `0.2586071973976171`. The overall
directional p95-ratio geomean is `0.2806672128213908`. Ratios below `1.0` mean
WorkPaper is faster for that metric.

The current worst mean row is `sheet-rename-dependencies`, with a mean ratio of
`0.8056914279903578`. The current worst p95 row is
`sheet-rename-dependencies`, with a p95 ratio of `0.7917355369127405`. The
headless leadership scorecard records `100/100` workloads winning both mean and
p95 against HyperFormula.

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

## What Is Not Claimed

This is not a blanket "fastest at every spreadsheet task" claim.

IronCalc Rust has `10` unsupported workload adapters in the current artifact.
Those rows are recorded explicitly and are not counted as wins.

TrueCalc is scalar-only coverage. xlsx-calc and IronCalc Rust are
workbook-wide-limited lanes. HyperFormula and Univer are the direct workbook-wide
lanes that satisfy the broad leadership coverage criterion.

Browser-grid rendering, import/export fidelity, collaborative sync, and every
possible user workbook are outside this headless runtime scorecard.

## How To Verify

Check that the committed artifacts still have the expected workload coverage and
shape:

```bash
pnpm workpaper:bench:competitive:check
pnpm workpaper:bench:univer:check
pnpm workpaper:bench:ironcalc-rust:check
pnpm workpaper:bench:xlsx-calc:check
pnpm workpaper:bench:truecalc:check
pnpm headless:performance:check
pnpm public:evidence:check
```

Regenerate timing evidence only when intentionally refreshing benchmark
artifacts:

```bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:univer:generate
pnpm workpaper:bench:ironcalc-rust:generate
pnpm workpaper:bench:xlsx-calc:generate
pnpm workpaper:bench:truecalc:generate
pnpm headless:performance:generate
pnpm public:evidence:generate
```

Do not change workload sizes, sampling, scoring, or definitions to preserve a
claim. If a rerun moves a row red, update the artifact, update this note, and
fix the production engine path rather than hiding the loss.

If a workload family is missing, a row looks too synthetic, or the p95 wording
is still too broad, use the public benchmark critique thread:
<https://github.com/proompteng/bilig/discussions/340>.
