import { WorkPaper } from '../../headless/src/work-paper.js'
import {
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT,
  type ComparativeBenchmarkSuiteOptions,
  type ComparativeMeasuredEngineResult,
  type ComparativeMemorySummary,
} from './benchmark-workpaper-vs-hyperformula.js'
import {
  measureMutationSample,
  measureWorkPaperBuildFromSheets,
  normalizeWorkPaperValue,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { normalizeBenchmarkValue } from './benchmark-workpaper-vs-univer-fixtures.js'
import {
  WORKPAPER_UNIVER_WORKLOADS,
  type WorkPaperUniverFixture,
  type WorkPaperUniverScenario,
  type WorkPaperUniverWorkload,
} from './benchmark-workpaper-vs-univer.js'
import { univerScenario } from './benchmark-workpaper-vs-univer-workload-resolver.js'
import type { MemoryMeasurement } from './metrics.js'
import { summarizeNumbers } from './stats.js'
import type {
  IronCalcRustApiPath,
  IronCalcRustMeasuredEngineResult,
  WorkPaperIronCalcRustBenchmarkReport,
  WorkPaperIronCalcRustBenchmarkResult,
  WorkPaperIronCalcRustFixture,
  WorkPaperIronCalcRustScorecard,
  WorkPaperIronCalcRustUnsupportedWorkload,
  WorkPaperIronCalcRustWorkload,
} from './benchmark-workpaper-vs-ironcalc-rust-types.js'

export type {
  IronCalcRustApiPath,
  IronCalcRustMeasuredEngineResult,
  WorkPaperIronCalcRustBenchmarkReport,
  WorkPaperIronCalcRustBenchmarkResult,
  WorkPaperIronCalcRustComparison,
  WorkPaperIronCalcRustFixture,
  WorkPaperIronCalcRustScorecard,
  WorkPaperIronCalcRustUnsupportedWorkload,
  WorkPaperIronCalcRustWorkload,
  WorkPaperIronCalcRustWorkloadFamily,
} from './benchmark-workpaper-vs-ironcalc-rust-types.js'

export interface IronCalcRustRunnerInput {
  readonly benchmark: {
    readonly sampleCount: number
    readonly warmupCount: number
  }
  readonly workloads: readonly IronCalcRustRunnerWorkloadInput[]
}

export interface IronCalcRustRunnerWorkloadInput {
  readonly observations: readonly IronCalcRustRunnerObservationInput[]
  readonly operation: IronCalcRustRunnerOperationInput
  readonly sheets: readonly IronCalcRustRunnerSheetInput[]
  readonly workload: WorkPaperIronCalcRustWorkload
}

export interface IronCalcRustRunnerObservationInput {
  readonly col: number
  readonly key: string
  readonly row: number
  readonly sheetName: string
}

export type IronCalcRustRunnerOperationInput =
  | {
      readonly kind: 'build'
    }
  | {
      readonly edits: readonly {
        readonly col: number
        readonly row: number
        readonly sheetName: string
        readonly value: boolean | number | string | null
      }[]
      readonly evaluation: 'auto' | 'paused'
      readonly kind: 'cell-edits'
    }
  | {
      readonly edit: {
        readonly col: number
        readonly row: number
        readonly sheetName: string
        readonly value: boolean | number | string | null
      }
      readonly kind: 'single-cell-edit'
    }
  | {
      readonly index: number
      readonly kind: 'insert-rows' | 'delete-rows'
      readonly sheetName: string
    }
  | {
      readonly delta: number
      readonly index: number
      readonly kind: 'move-rows' | 'move-columns'
      readonly sheetName: string
    }
  | {
      readonly index: number
      readonly kind: 'insert-columns' | 'delete-columns'
      readonly sheetName: string
    }
  | {
      readonly kind: 'rename-sheet'
      readonly newName: string
      readonly oldName: string
    }
  | {
      readonly endCol: number
      readonly endRow: number
      readonly kind: 'range-read'
      readonly sheetName: string
      readonly startCol: number
      readonly startRow: number
      readonly summary: 'dense' | 'formula-grid' | 'sparse-wide'
      readonly middleCol?: number
    }

export interface IronCalcRustRunnerSheetInput {
  readonly cells: readonly (readonly (boolean | number | string | null)[])[]
  readonly name: string
}

export interface IronCalcRustRunnerOutput {
  readonly engine: {
    readonly crate: 'ironcalc_base'
    readonly version: string
  }
  readonly results: readonly IronCalcRustRunnerWorkloadOutput[]
}

export interface IronCalcRustRunnerWorkloadOutput {
  readonly apiPath: IronCalcRustApiPath
  readonly elapsedMs: readonly number[]
  readonly verification: Record<string, unknown>
  readonly workload: WorkPaperIronCalcRustWorkload
}

interface IronCalcRustComparableScenario {
  readonly fixture: WorkPaperIronCalcRustFixture
  readonly observations: readonly IronCalcRustRunnerObservationInput[]
  readonly operation: IronCalcRustRunnerOperationInput
  readonly scenario: WorkPaperUniverScenario
  readonly workload: WorkPaperIronCalcRustWorkload
}

interface ResolvedBenchmarkSuiteOptions {
  readonly sampleCount: number
  readonly warmupCount: number
}

interface WorkPaperIronCalcRustSuiteOptions extends ComparativeBenchmarkSuiteOptions {
  readonly onWorkloadStart?: (workload: WorkPaperIronCalcRustWorkload, index: number, total: number) => void
  readonly workloads?: readonly WorkPaperIronCalcRustWorkload[]
}

export const IRONCALC_RUST_CRATE_NAME = 'ironcalc_base'
export const IRONCALC_RUST_CRATE_VERSION = '0.7.1'
export const WORKPAPER_IRONCALC_RUST_CANDIDATE_WORKLOADS = WORKPAPER_UNIVER_WORKLOADS
export const WORKPAPER_IRONCALC_RUST_COMPARABLE_SCENARIOS = buildIronCalcRustComparableScenarios()
export const WORKPAPER_IRONCALC_RUST_WORKLOADS = WORKPAPER_IRONCALC_RUST_COMPARABLE_SCENARIOS.map((scenario) => scenario.workload)
export const WORKPAPER_IRONCALC_RUST_UNSUPPORTED_WORKLOADS = buildIronCalcRustUnsupportedWorkloads()

export function buildIronCalcRustRunnerInput(options: WorkPaperIronCalcRustSuiteOptions = {}): IronCalcRustRunnerInput {
  const resolvedOptions = resolveSuiteOptions(options)
  const scenarios = selectIronCalcRustScenarios(options.workloads)
  return {
    benchmark: resolvedOptions,
    workloads: scenarios.map((scenario) => ({
      observations: scenario.observations,
      operation: scenario.operation,
      sheets: Object.entries(scenario.scenario.buildWorkPaperSheets()).map(([name, cells]) => ({ cells, name })),
      workload: scenario.workload,
    })),
  }
}

export function runWorkPaperVsIronCalcRustBenchmarkSuite(
  runnerOutput: IronCalcRustRunnerOutput,
  options: WorkPaperIronCalcRustSuiteOptions = {},
): WorkPaperIronCalcRustBenchmarkResult[] {
  const resolvedOptions = resolveSuiteOptions(options)
  const ironCalcResults = new Map(runnerOutput.results.map((result) => [result.workload, result]))
  const scenarios = selectIronCalcRustScenarios(options.workloads)
  return scenarios.map((scenario, index) => {
    options.onWorkloadStart?.(scenario.workload, index, scenarios.length)
    const ironCalcResult = ironCalcResults.get(scenario.workload)
    if (ironCalcResult === undefined) {
      throw new Error(`IronCalc Rust runner did not return workload ${scenario.workload}`)
    }
    return runIronCalcRustScenario(scenario, ironCalcResult, resolvedOptions)
  })
}

export function buildWorkPaperVsIronCalcRustBenchmarkReport(
  results: readonly WorkPaperIronCalcRustBenchmarkResult[],
): WorkPaperIronCalcRustBenchmarkReport {
  return {
    suite: 'workpaper-vs-ironcalc-rust',
    scorecard: deriveWorkPaperIronCalcRustScorecard(results, WORKPAPER_IRONCALC_RUST_UNSUPPORTED_WORKLOADS),
    results,
  }
}

export function deriveWorkPaperIronCalcRustScorecard(
  results: readonly WorkPaperIronCalcRustBenchmarkResult[],
  unsupportedWorkloads: readonly WorkPaperIronCalcRustUnsupportedWorkload[],
): WorkPaperIronCalcRustScorecard {
  if (results.length === 0) {
    throw new Error('Cannot build a WorkPaper vs IronCalc Rust scorecard without benchmark results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToIronCalcRustMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToIronCalcRustP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToIronCalcRustMeanRatio < 1 && result.comparison.workpaperToIronCalcRustP95Ratio < 1,
  ).length
  return {
    comparableWorkloadCount: results.length,
    coverageNote:
      'IronCalc Rust is covered through a generated local sidecar that pins crates.io `ironcalc_base = "=0.7.1"` and runs all samples in one release-mode Rust process. This workbook-wide-limited lane uses canonical headless benchmark workload names, times `Model` for workbook construction rows and `UserModel` for single-cell user-operation rows, and records unsupported rows explicitly until Rust operation adapters prove parity.',
    coverageTier: 'workbook-wide-limited',
    directionalMeanRatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToIronCalcRustMeanRatio)),
    directionalP95RatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToIronCalcRustP95Ratio)),
    ironCalcRustMeanWinCount: results.length - meanWinCount,
    ironCalcRustP95WinCount: results.length - p95WinCount,
    meanAndP95WinCount,
    meanWinCount,
    p95WinCount,
    unsupportedWorkloads,
    workloadFamilies: orderedUnique(results.map((result) => result.fixture.family)),
    worstMeanRatioWorkload: maxComparableRatioWorkload(results, 'workpaperToIronCalcRustMeanRatio'),
    worstP95RatioWorkload: maxComparableRatioWorkload(results, 'workpaperToIronCalcRustP95Ratio'),
    worstWorkpaperToIronCalcRustMeanRatio: maxComparableRatio(results, 'workpaperToIronCalcRustMeanRatio'),
    worstWorkpaperToIronCalcRustP95Ratio: maxComparableRatio(results, 'workpaperToIronCalcRustP95Ratio'),
  }
}

function buildIronCalcRustComparableScenarios(): readonly IronCalcRustComparableScenario[] {
  return WORKPAPER_IRONCALC_RUST_CANDIDATE_WORKLOADS.flatMap((workload) => {
    const comparable = resolveIronCalcRustComparableScenario(workload)
    return comparable === undefined ? [] : [comparable]
  })
}

function buildIronCalcRustUnsupportedWorkloads(): readonly WorkPaperIronCalcRustUnsupportedWorkload[] {
  const comparableWorkloads = new Set(WORKPAPER_IRONCALC_RUST_WORKLOADS)
  return WORKPAPER_IRONCALC_RUST_CANDIDATE_WORKLOADS.filter((workload) => !comparableWorkloads.has(workload)).map((workload) => {
    const scenario = univerScenario(workload)
    const operationShape =
      scenario.kind === 'build'
        ? 'build'
        : scenario.fixture.edit === undefined
          ? 'custom mutation without a single-cell edit fixture'
          : 'custom mutation requiring a dedicated Rust operation adapter'
    return {
      evidence: [
        `canonical workload family: ${scenario.fixture.family}`,
        `scenario shape: ${operationShape}`,
        'IronCalc Rust sidecar adapters currently cover build, single-cell edit, batch edit, supported structural, sheet rename, and range-read rows; unsupported rows are not counted as wins.',
      ],
      reason: 'Rust-side operation adapter and typed parity check are not implemented for this canonical workload yet.',
      workload,
    }
  })
}

function resolveIronCalcRustComparableScenario(workload: WorkPaperUniverWorkload): IronCalcRustComparableScenario | undefined {
  const scenario = univerScenario(workload)
  const result = scenario.fixture.result
  if (result === undefined) {
    return undefined
  }
  if (scenario.kind === 'build') {
    return {
      fixture: toIronCalcRustFixture(scenario.fixture),
      observations: [{ ...result, key: 'value' }],
      operation: { kind: 'build' },
      scenario,
      workload,
    }
  }
  const customOperation = customIronCalcRustOperation(workload, scenario.fixture)
  if (customOperation !== undefined) {
    return {
      fixture: toIronCalcRustFixture(scenario.fixture),
      observations: [{ ...result, key: 'value' }],
      operation: customOperation,
      scenario,
      workload,
    }
  }
  const edit = scenario.fixture.edit
  if (edit === undefined) {
    return undefined
  }
  return {
    fixture: toIronCalcRustFixture(scenario.fixture),
    observations: [{ ...result, key: 'value' }],
    operation: {
      edit: {
        col: edit.col,
        row: edit.row,
        sheetName: edit.sheetName,
        value: edit.value,
      },
      kind: 'single-cell-edit',
    },
    scenario,
    workload,
  }
}

function selectIronCalcRustScenarios(
  workloads: readonly WorkPaperIronCalcRustWorkload[] | undefined,
): readonly IronCalcRustComparableScenario[] {
  if (workloads === undefined) {
    return WORKPAPER_IRONCALC_RUST_COMPARABLE_SCENARIOS
  }
  const selected = new Set(workloads)
  return WORKPAPER_IRONCALC_RUST_COMPARABLE_SCENARIOS.filter((scenario) => selected.has(scenario.workload))
}

function customIronCalcRustOperation(
  workload: WorkPaperUniverWorkload,
  fixture: WorkPaperUniverFixture,
): IronCalcRustRunnerOperationInput | undefined {
  const customWorkload: string = workload
  switch (customWorkload) {
    case 'batch-edit-recalc':
    case 'batch-edit-single-column':
    case 'batch-edit-single-column-small':
    case 'batch-edit-single-column-large':
      return {
        edits: Array.from({ length: fixture.rowCount }, (_value, row) => ({
          col: 0,
          row,
          sheetName: 'Bench',
          value: row * 3,
        })),
        evaluation: 'paused',
        kind: 'cell-edits',
      }
    case 'batch-suspended-single-column':
      return {
        edits: Array.from({ length: fixture.rowCount }, (_value, row) => ({
          col: 0,
          row,
          sheetName: 'Bench',
          value: row * 7,
        })),
        evaluation: 'paused',
        kind: 'cell-edits',
      }
    case 'batch-edit-multi-column-small':
    case 'batch-edit-multi-column':
    case 'batch-edit-multi-column-large':
    case 'batch-suspended-multi-column':
      return {
        edits: Array.from({ length: fixture.rowCount }, (_value, row) => [
          { col: 0, row, sheetName: 'Bench', value: row * 3 },
          { col: 1, row, sheetName: 'Bench', value: row * 5 },
        ]).flat(),
        evaluation: 'paused',
        kind: 'cell-edits',
      }
    case 'batch-edit-rectangular-block':
    case 'batch-edit-rectangular-block-wide': {
      const inputCols = fixture.result?.col ?? 0
      return {
        edits: Array.from({ length: fixture.rowCount }, (_rowValue, row) =>
          Array.from({ length: inputCols }, (_colValue, col) => ({
            col,
            row,
            sheetName: 'Bench',
            value: (row + 1) * (col + 2),
          })),
        ).flat(),
        evaluation: 'paused',
        kind: 'cell-edits',
      }
    }
    case 'batch-clear-rectangular-block':
    case 'batch-clear-rectangular-block-wide': {
      const inputCols = fixture.result?.col ?? 0
      return {
        edits: Array.from({ length: fixture.rowCount }, (_rowValue, row) =>
          Array.from({ length: inputCols }, (_colValue, col) => ({
            col,
            row,
            sheetName: 'Bench',
            value: null,
          })),
        ).flat(),
        evaluation: 'paused',
        kind: 'cell-edits',
      }
    }
    case 'lookup-with-column-index-after-batch-write':
    case 'lookup-with-column-index-after-batch-write-large': {
      const rowCount = fixture.rowCount - 1
      const editCount = workload === 'lookup-with-column-index-after-batch-write-large' ? 512 : 256
      return {
        edits: Array.from({ length: editCount }, (_value, index) => {
          const row = rowCount - index
          return { col: 0, row, sheetName: 'Bench', value: row + 10_000 }
        }),
        evaluation: 'paused',
        kind: 'cell-edits',
      }
    }
    case 'structural-insert-rows':
    case 'structural-insert-rows-small':
    case 'structural-insert-rows-large':
      return { index: Math.floor(fixture.rowCount / 2), kind: 'insert-rows', sheetName: 'Bench' }
    case 'structural-delete-rows':
      return { index: Math.floor(fixture.rowCount / 2), kind: 'delete-rows', sheetName: 'Bench' }
    case 'structural-insert-columns':
    case 'structural-insert-columns-small':
    case 'structural-insert-columns-large':
      return { index: 1, kind: 'insert-columns', sheetName: 'Bench' }
    case 'structural-delete-columns':
    case 'structural-delete-columns-large':
      return { index: 1, kind: 'delete-columns', sheetName: 'Bench' }
    case 'sheet-rename-dependencies':
      return { kind: 'rename-sheet', newName: 'Source', oldName: 'Data' }
    case 'range-read':
    case 'range-read-dense':
      return {
        endCol: 23,
        endRow: 239,
        kind: 'range-read',
        sheetName: 'Bench',
        startCol: 0,
        startRow: 0,
        summary: 'dense',
      }
    case 'range-read-wide':
      return {
        endCol: 95,
        endRow: 127,
        kind: 'range-read',
        sheetName: 'Bench',
        startCol: 0,
        startRow: 0,
        summary: 'dense',
      }
    case 'range-read-sparse-wide':
      return {
        endCol: 95,
        endRow: 127,
        kind: 'range-read',
        middleCol: 48,
        sheetName: 'Bench',
        startCol: 0,
        startRow: 0,
        summary: 'sparse-wide',
      }
    case 'range-read-formula-grid':
      return {
        endCol: 11,
        endRow: 255,
        kind: 'range-read',
        sheetName: 'Bench',
        startCol: 4,
        startRow: 0,
        summary: 'formula-grid',
      }
    case 'range-read-formula-grid-wide':
      return {
        endCol: 23,
        endRow: 127,
        kind: 'range-read',
        sheetName: 'Bench',
        startCol: 8,
        startRow: 0,
        summary: 'formula-grid',
      }
    default:
      return undefined
  }
}

function toIronCalcRustFixture(fixture: WorkPaperUniverFixture): WorkPaperIronCalcRustFixture {
  const result = fixture.result
  if (result === undefined) {
    throw new Error('IronCalc Rust comparable workloads require a result cell')
  }
  return {
    ...(fixture.edit ? { edit: fixture.edit } : {}),
    family: fixture.family,
    formula: fixture.formula,
    result,
    rowCount: fixture.rowCount,
  }
}

function runIronCalcRustScenario(
  scenario: IronCalcRustComparableScenario,
  ironCalcRunnerResult: IronCalcRustRunnerWorkloadOutput,
  options: ResolvedBenchmarkSuiteOptions,
): WorkPaperIronCalcRustBenchmarkResult {
  const workpaper = benchmarkSupportedEngine(() => measureWorkPaperScenarioSample(scenario), options)
  const ironCalcRust = summarizeIronCalcRustResult(ironCalcRunnerResult)
  const workPaperVerification = JSON.stringify(workpaper.verification)
  const ironCalcVerification = JSON.stringify(ironCalcRust.verification)
  if (workPaperVerification !== ironCalcVerification) {
    throw new Error(
      `Verification mismatch for ${scenario.workload}: WorkPaper ${workPaperVerification} !== IronCalc Rust ${ironCalcVerification}`,
    )
  }

  const fasterEngine = workpaper.elapsedMs.mean <= ironCalcRust.elapsedMs.mean ? 'workpaper' : 'ironcalc-rust'
  const fasterMean = fasterEngine === 'workpaper' ? workpaper.elapsedMs.mean : ironCalcRust.elapsedMs.mean
  const slowerMean = fasterEngine === 'workpaper' ? ironCalcRust.elapsedMs.mean : workpaper.elapsedMs.mean

  return {
    workload: scenario.workload,
    category: 'workbook-wide-limited',
    comparable: true,
    fixture: scenario.fixture,
    comparison: {
      confidenceIntervalOverlaps:
        workpaper.elapsedMs.confidence95.low <= ironCalcRust.elapsedMs.confidence95.high &&
        ironCalcRust.elapsedMs.confidence95.low <= workpaper.elapsedMs.confidence95.high,
      fasterEngine,
      maxRelativeNoise: Math.max(workpaper.elapsedMs.relativeStandardDeviation, ironCalcRust.elapsedMs.relativeStandardDeviation),
      meanSpeedup: slowerMean / fasterMean,
      verificationEquivalent: true,
      workpaperToIronCalcRustMeanRatio: workpaper.elapsedMs.mean / ironCalcRust.elapsedMs.mean,
      workpaperToIronCalcRustMedianRatio: workpaper.elapsedMs.median / ironCalcRust.elapsedMs.median,
      workpaperToIronCalcRustP95Ratio: workpaper.elapsedMs.p95 / ironCalcRust.elapsedMs.p95,
    },
    engines: {
      ironCalcRust,
      workpaper,
    },
  }
}

function measureWorkPaperScenarioSample(scenario: IronCalcRustComparableScenario): BenchmarkSample {
  if (scenario.operation.kind === 'build') {
    return measureWorkPaperBuildFromSheets(
      scenario.scenario.buildWorkPaperSheets(),
      (workbook) => verifyWorkPaperObservations(workbook, scenario.observations),
      scenario.scenario.workpaperOptions,
      scenario.scenario.workpaperNamedExpressions,
    )
  }
  const workbook = WorkPaper.buildFromSheets(
    scenario.scenario.buildWorkPaperSheets(),
    scenario.scenario.workpaperOptions,
    scenario.scenario.workpaperNamedExpressions,
  )
  if (scenario.operation.kind !== 'single-cell-edit') {
    if (scenario.operation.kind === 'range-read') {
      return measureMutationSample(
        workbook,
        () => {
          if (scenario.scenario.executeWorkPaperMutation === undefined) {
            throw new Error(`IronCalc Rust workload ${scenario.workload} is missing a WorkPaper operation`)
          }
          return scenario.scenario.executeWorkPaperMutation(workbook)
        },
        (result) => scenario.scenario.verifyWorkPaper(workbook, result),
      )
    }
    return measureMutationSample(
      workbook,
      () => {
        if (scenario.scenario.executeWorkPaperMutation === undefined) {
          throw new Error(`IronCalc Rust workload ${scenario.workload} is missing a WorkPaper operation`)
        }
        return scenario.scenario.executeWorkPaperMutation(workbook)
      },
      () => verifyWorkPaperObservations(workbook, scenario.observations),
    )
  }
  const edit = scenario.operation.edit
  const editSheetId = workbook.getSheetId(edit.sheetName)
  if (editSheetId === undefined) {
    workbook.dispose()
    throw new Error(`WorkPaper IronCalc Rust benchmark fixture did not create ${edit.sheetName}`)
  }
  return measureMutationSample(
    workbook,
    () =>
      scenario.scenario.executeWorkPaperMutation?.(workbook) ??
      workbook.setCellContents({ sheet: editSheetId, row: edit.row, col: edit.col }, edit.value),
    () => verifyWorkPaperObservations(workbook, scenario.observations),
  )
}

function verifyWorkPaperObservations(
  workbook: WorkPaper,
  observations: readonly IronCalcRustRunnerObservationInput[],
): Record<string, unknown> {
  return Object.fromEntries(
    observations.map((observation) => {
      const sheetId = workbook.getSheetId(observation.sheetName)
      if (sheetId === undefined) {
        throw new Error(`WorkPaper IronCalc Rust benchmark fixture did not create ${observation.sheetName}`)
      }
      return [
        observation.key,
        normalizeBenchmarkValue(
          normalizeWorkPaperValue(workbook.getCellValue({ sheet: sheetId, row: observation.row, col: observation.col })),
        ),
      ]
    }),
  )
}

function benchmarkSupportedEngine(
  runSample: () => BenchmarkSample,
  options: ResolvedBenchmarkSuiteOptions,
): ComparativeMeasuredEngineResult {
  for (let warmup = 0; warmup < options.warmupCount; warmup += 1) {
    runSample()
  }

  const samples: BenchmarkSample[] = []
  for (let sample = 0; sample < options.sampleCount; sample += 1) {
    samples.push(runSample())
  }

  const verificationStrings = new Set(samples.map((sample) => JSON.stringify(sample.verification)))
  if (verificationStrings.size !== 1) {
    throw new Error('Benchmark verification drifted across WorkPaper samples')
  }

  return {
    status: 'supported',
    elapsedMs: summarizeNumbers(samples.map((sample) => sample.elapsedMs)),
    memoryDeltaBytes: summarizeMemory(samples.map((sample) => sample.memory)),
    verification: samples[0]?.verification ?? {},
  }
}

function summarizeIronCalcRustResult(result: IronCalcRustRunnerWorkloadOutput): IronCalcRustMeasuredEngineResult {
  if (result.elapsedMs.length === 0) {
    throw new Error(`IronCalc Rust runner returned no samples for ${result.workload}`)
  }
  return {
    status: 'supported',
    apiPath: result.apiPath,
    elapsedMs: summarizeNumbers([...result.elapsedMs]),
    verification: Object.fromEntries(Object.entries(result.verification).map(([key, value]) => [key, normalizeBenchmarkValue(value)])),
  }
}

function summarizeMemory(samples: readonly MemoryMeasurement[]): ComparativeMemorySummary {
  return {
    rssBytes: summarizeNumbers(samples.map((sample) => sample.delta.rssBytes)),
    heapUsedBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapUsedBytes)),
    heapTotalBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapTotalBytes)),
    externalBytes: summarizeNumbers(samples.map((sample) => sample.delta.externalBytes)),
    arrayBuffersBytes: summarizeNumbers(samples.map((sample) => sample.delta.arrayBuffersBytes)),
  }
}

function geometricMean(values: readonly number[]): number {
  const totalLog = values.reduce((sum, value) => {
    if (value <= 0) {
      throw new Error(`Cannot compute geomean for non-positive value: ${String(value)}`)
    }
    return sum + Math.log(value)
  }, 0)
  return Math.exp(totalLog / values.length)
}

function maxComparableRatio(
  results: readonly WorkPaperIronCalcRustBenchmarkResult[],
  ratioKey: 'workpaperToIronCalcRustMeanRatio' | 'workpaperToIronCalcRustP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly WorkPaperIronCalcRustBenchmarkResult[],
  ratioKey: 'workpaperToIronCalcRustMeanRatio' | 'workpaperToIronCalcRustP95Ratio',
): WorkPaperIronCalcRustWorkload {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}

function orderedUnique<T extends string>(values: readonly T[]): T[] {
  return [...new Set(values)]
}

function resolveSuiteOptions(options: ComparativeBenchmarkSuiteOptions): ResolvedBenchmarkSuiteOptions {
  return {
    sampleCount: options.sampleCount ?? DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: options.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  }
}
