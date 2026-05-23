import { performance } from 'node:perf_hooks'
import * as UniverPresetModule from '@univerjs/presets'
import * as UniverSheetsNodeCorePresetModule from '@univerjs/preset-sheets-node-core'
import sheetsNodeCoreEnUS from '@univerjs/preset-sheets-node-core/locales/en-US'
import { WorkPaper, type WorkPaperSheet } from '../../headless/src/work-paper.js'
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  type ComparativeBenchmarkSuiteOptions,
  type ComparativeMeasuredEngineResult,
  type ComparativeMemorySummary,
} from './benchmark-workpaper-vs-hyperformula.js'
import {
  measureWorkPaperBuildFromSheets,
  measureMutationSample,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import {
  crossSheetAggregateRecalcScenario,
  crossSheetDashboardBuildScenario,
  crossSheetDashboardRecalcScenario,
  crossSheetScalarFanoutRecalcScenario,
  manySheetsBuildScenario,
} from './benchmark-workpaper-vs-univer-multisheet-workloads.js'
import {
  aggregate2dCanonicalScenario,
  approximateDuplicateLookupCanonicalScenario,
  approximateLookupCanonicalScenario,
  denseLiteralBuildScenario,
  exactLookupCanonicalScenario,
  formulaChainRowScenario,
  formulaFanoutRowScenario,
  indexMatchExactCanonicalScenario,
  indexReferenceCanonicalScenario,
  mixedContentBuildScenario,
  overlappingAggregateCanonicalScenario,
  parserCacheMixedTemplateBuildScenario,
  parserCacheRowTemplateBuildScenario,
  parserCacheUniqueFormulaBuildScenario,
  slidingAggregateCanonicalScenario,
  textLookupCanonicalScenario,
} from './benchmark-workpaper-vs-univer-single-sheet-workloads.js'
import { waitForUniverVerification } from './benchmark-workpaper-vs-univer-sync.js'
import { measureMemory, sampleMemory, type MemoryMeasurement } from './metrics.js'
import { summarizeNumbers } from './stats.js'

const { createUniver, LocaleType, mergeLocales } = resolveModuleExports<{
  readonly createUniver: typeof UniverPresetModule.createUniver
  readonly LocaleType: typeof UniverPresetModule.LocaleType
  readonly mergeLocales: typeof UniverPresetModule.mergeLocales
}>(UniverPresetModule, 'createUniver')
const { UniverSheetsNodeCorePreset } = resolveModuleExports<{
  readonly UniverSheetsNodeCorePreset: typeof UniverSheetsNodeCorePresetModule.UniverSheetsNodeCorePreset
}>(UniverSheetsNodeCorePresetModule, 'UniverSheetsNodeCorePreset')

export type WorkPaperUniverWorkload =
  | 'build-from-sheets'
  | 'build-dense-literals'
  | 'build-dense-literals-wide'
  | 'build-dense-literals-tall'
  | 'build-mixed-content'
  | 'build-mixed-content-small'
  | 'build-mixed-content-large'
  | 'build-parser-cache-row-templates'
  | 'build-parser-cache-mixed-templates'
  | 'build-parser-cache-unique-formulas'
  | 'build-many-sheets'
  | 'build-many-sheets-wide'
  | 'build-many-sheets-narrow'
  | 'build-cross-sheet-dashboard'
  | 'build-cross-sheet-dashboard-small'
  | 'build-cross-sheet-dashboard-large'
  | 'cross-sheet-scalar-recalc'
  | 'cross-sheet-aggregate-recalc'
  | 'cross-sheet-dashboard-recalc'
  | 'single-edit-chain'
  | 'single-edit-chain-small'
  | 'single-edit-chain-large'
  | 'single-edit-fanout'
  | 'single-edit-fanout-small'
  | 'single-edit-fanout-large'
  | 'aggregate-2d-ranges'
  | 'aggregate-2d-ranges-small'
  | 'aggregate-2d-ranges-large'
  | 'aggregate-overlapping-ranges'
  | 'aggregate-overlapping-ranges-small'
  | 'aggregate-overlapping-sliding-window'
  | 'aggregate-overlapping-sliding-window-wide'
  | 'lookup-no-column-index'
  | 'lookup-no-column-index-small'
  | 'lookup-with-column-index'
  | 'lookup-with-column-index-large'
  | 'lookup-index-match-exact'
  | 'lookup-index-match-exact-large'
  | 'lookup-index-reference'
  | 'lookup-index-reference-large'
  | 'lookup-approximate-sorted'
  | 'lookup-approximate-sorted-large'
  | 'lookup-approximate-duplicates'
  | 'lookup-text-exact'
  | 'lookup-text-exact-large'

export type WorkPaperUniverWorkloadFamily =
  | 'aggregate'
  | 'aggregate-2d'
  | 'build'
  | 'formula-chain'
  | 'formula-fanout'
  | 'cross-sheet'
  | 'lookup-approximate'
  | 'lookup-exact'
  | 'overlapping-aggregate'
  | 'range-stats'

export type UniverEditableValue = boolean | number | string
type UniverCellValue = UniverEditableValue

export interface WorkPaperUniverFixture {
  readonly edit?: {
    readonly address: string
    readonly col: number
    readonly row: number
    readonly sheetName: string
    readonly value: UniverEditableValue
  }
  readonly family: WorkPaperUniverWorkloadFamily
  readonly formula: string
  readonly result?: {
    readonly address: string
    readonly col: number
    readonly row: number
    readonly sheetName: string
  }
  readonly columnCount?: number
  readonly rowCount: number
}

export interface WorkPaperUniverComparison {
  readonly confidenceIntervalOverlaps: boolean
  readonly fasterEngine: 'univer' | 'workpaper'
  readonly maxRelativeNoise: number
  readonly meanSpeedup: number
  readonly verificationEquivalent: true
  readonly workpaperToUniverMeanRatio: number
  readonly workpaperToUniverMedianRatio: number
  readonly workpaperToUniverP95Ratio: number
}

export interface WorkPaperUniverBenchmarkResult {
  readonly workload: WorkPaperUniverWorkload
  readonly category: 'workbook-wide'
  readonly comparable: true
  readonly fixture: WorkPaperUniverFixture
  readonly comparison: WorkPaperUniverComparison
  readonly engines: {
    readonly univer: ComparativeMeasuredEngineResult
    readonly workpaper: ComparativeMeasuredEngineResult
  }
}

export interface WorkPaperUniverScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: 'workbook-wide'
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly univerMeanWinCount: number
  readonly univerP95WinCount: number
  readonly workloadFamilies: readonly WorkPaperUniverWorkloadFamily[]
  readonly worstMeanRatioWorkload: WorkPaperUniverWorkload
  readonly worstP95RatioWorkload: WorkPaperUniverWorkload
  readonly worstWorkpaperToUniverMeanRatio: number
  readonly worstWorkpaperToUniverP95Ratio: number
}

export interface WorkPaperUniverBenchmarkReport {
  readonly suite: 'workpaper-vs-univer'
  readonly scorecard: WorkPaperUniverScorecard
  readonly results: readonly WorkPaperUniverBenchmarkResult[]
}

export interface WorkPaperUniverScenario {
  readonly kind: 'build' | 'mutation'
  readonly fixture: WorkPaperUniverFixture
  readonly buildWorkPaperSheets: () => Record<string, WorkPaperSheet>
  readonly workpaperOptions?: Parameters<typeof WorkPaper.buildFromSheets>[1]
  readonly setupUniver: (runtime: UniverRuntime) => Promise<void>
  readonly verifyUniver: (runtime: UniverRuntime) => Record<string, unknown>
  readonly verifyWorkPaper: (workbook: WorkPaper) => Record<string, unknown>
}

interface ResolvedBenchmarkSuiteOptions {
  readonly sampleCount: number
  readonly warmupCount: number
}

export interface UniverRuntime {
  readonly formula: UniverFormulaFacade
  readonly sheet: UniverWorksheetFacade
  readonly univer: { dispose(): void }
  readonly workbook: UniverWorkbookFacade
}

export interface UniverFormulaFacade {
  onCalculationEnd(): Promise<void>
}

interface UniverRangeFacade {
  getValue(): unknown
  getValues(): unknown[][]
  setFormula(formula: string): UniverRangeFacade
  setFormulas(formulas: string[][]): UniverRangeFacade
  setValue(value: UniverEditableValue): UniverRangeFacade
  setValues(values: UniverCellValue[][]): UniverRangeFacade
}

export interface UniverWorksheetFacade {
  getRange(row: number, column: number): UniverRangeFacade
  getRange(row: number, column: number, numRows: number): UniverRangeFacade
  getRange(row: number, column: number, numRows: number, numColumns: number): UniverRangeFacade
  getRange(a1Notation: string): UniverRangeFacade
}

export interface UniverWorkbookFacade {
  create(name: string, rows: number, columns: number): UniverWorksheetFacade
  dispose(): void
  getActiveSheet(): UniverWorksheetFacade
  getSheetByName(name: string): UniverWorksheetFacade | null
}

export const WORKPAPER_UNIVER_WORKLOADS = [
  'build-from-sheets',
  'build-dense-literals',
  'build-dense-literals-wide',
  'build-dense-literals-tall',
  'build-mixed-content',
  'build-mixed-content-small',
  'build-mixed-content-large',
  'build-parser-cache-row-templates',
  'build-parser-cache-mixed-templates',
  'build-parser-cache-unique-formulas',
  'single-edit-chain',
  'single-edit-chain-small',
  'single-edit-chain-large',
  'single-edit-fanout',
  'single-edit-fanout-small',
  'single-edit-fanout-large',
  'aggregate-2d-ranges',
  'aggregate-2d-ranges-small',
  'aggregate-2d-ranges-large',
  'aggregate-overlapping-ranges',
  'aggregate-overlapping-ranges-small',
  'aggregate-overlapping-sliding-window',
  'aggregate-overlapping-sliding-window-wide',
  'lookup-no-column-index',
  'lookup-no-column-index-small',
  'lookup-with-column-index',
  'lookup-with-column-index-large',
  'lookup-index-match-exact',
  'lookup-index-match-exact-large',
  'lookup-index-reference',
  'lookup-index-reference-large',
  'lookup-approximate-sorted',
  'lookup-approximate-sorted-large',
  'lookup-approximate-duplicates',
  'lookup-text-exact',
  'lookup-text-exact-large',
  'build-many-sheets',
  'build-many-sheets-wide',
  'build-many-sheets-narrow',
  'build-cross-sheet-dashboard',
  'build-cross-sheet-dashboard-small',
  'build-cross-sheet-dashboard-large',
  'cross-sheet-scalar-recalc',
  'cross-sheet-aggregate-recalc',
  'cross-sheet-dashboard-recalc',
] as const satisfies readonly WorkPaperUniverWorkload[]

const univerCalculationTimeoutMs = 10_000

export async function runWorkPaperVsUniverBenchmarkSuite(
  options: ComparativeBenchmarkSuiteOptions = {},
): Promise<WorkPaperUniverBenchmarkResult[]> {
  const resolvedOptions = resolveSuiteOptions(options)
  const results: WorkPaperUniverBenchmarkResult[] = []
  for (const workload of WORKPAPER_UNIVER_WORKLOADS) {
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Competitive benchmark workloads run sequentially to keep timing samples isolated.
    results.push(await runUniverScenario(workload, univerScenario(workload), resolvedOptions))
  }
  return results
}

export function buildWorkPaperVsUniverBenchmarkReport(results: readonly WorkPaperUniverBenchmarkResult[]): WorkPaperUniverBenchmarkReport {
  if (results.length === 0) {
    throw new Error('Cannot build a WorkPaper vs Univer scorecard without benchmark results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToUniverMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToUniverP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToUniverMeanRatio < 1 && result.comparison.workpaperToUniverP95Ratio < 1,
  ).length

  return {
    suite: 'workpaper-vs-univer',
    scorecard: {
      comparableWorkloadCount: results.length,
      coverageNote:
        'Univer is covered through its documented @univerjs/preset-sheets-node-core headless Node preset. This workbook-wide lane times public Facade API workbook construction, cell edits, and formula recalculation completion on equivalent build, aggregate, chain, fanout, and lookup workloads.',
      coverageTier: 'workbook-wide',
      directionalMeanRatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToUniverMeanRatio)),
      directionalP95RatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToUniverP95Ratio)),
      meanAndP95WinCount,
      meanWinCount,
      p95WinCount,
      univerMeanWinCount: results.length - meanWinCount,
      univerP95WinCount: results.length - p95WinCount,
      workloadFamilies: orderedUnique(results.map((result) => result.fixture.family)),
      worstMeanRatioWorkload: maxComparableRatioWorkload(results, 'workpaperToUniverMeanRatio'),
      worstP95RatioWorkload: maxComparableRatioWorkload(results, 'workpaperToUniverP95Ratio'),
      worstWorkpaperToUniverMeanRatio: maxComparableRatio(results, 'workpaperToUniverMeanRatio'),
      worstWorkpaperToUniverP95Ratio: maxComparableRatio(results, 'workpaperToUniverP95Ratio'),
    },
    results,
  }
}

async function runUniverScenario(
  workload: WorkPaperUniverWorkload,
  scenario: WorkPaperUniverScenario,
  options: ResolvedBenchmarkSuiteOptions,
): Promise<WorkPaperUniverBenchmarkResult> {
  const workpaper = benchmarkSupportedEngine(() => measureWorkPaperSample(scenario), options)
  const univer = await benchmarkSupportedEngineAsync(() => measureUniverSample(scenario, workpaper.verification), options)
  const workPaperVerification = JSON.stringify(workpaper.verification)
  const univerVerification = JSON.stringify(univer.verification)
  if (workPaperVerification !== univerVerification) {
    throw new Error(`Verification mismatch for ${workload}: WorkPaper ${workPaperVerification} !== Univer ${univerVerification}`)
  }

  const fasterEngine = workpaper.elapsedMs.mean <= univer.elapsedMs.mean ? 'workpaper' : 'univer'
  const fasterMean = fasterEngine === 'workpaper' ? workpaper.elapsedMs.mean : univer.elapsedMs.mean
  const slowerMean = fasterEngine === 'workpaper' ? univer.elapsedMs.mean : workpaper.elapsedMs.mean

  return {
    workload,
    category: 'workbook-wide',
    comparable: true,
    fixture: scenario.fixture,
    comparison: {
      confidenceIntervalOverlaps:
        workpaper.elapsedMs.confidence95.low <= univer.elapsedMs.confidence95.high &&
        univer.elapsedMs.confidence95.low <= workpaper.elapsedMs.confidence95.high,
      fasterEngine,
      maxRelativeNoise: Math.max(workpaper.elapsedMs.relativeStandardDeviation, univer.elapsedMs.relativeStandardDeviation),
      meanSpeedup: slowerMean / fasterMean,
      verificationEquivalent: true,
      workpaperToUniverMeanRatio: workpaper.elapsedMs.mean / univer.elapsedMs.mean,
      workpaperToUniverMedianRatio: workpaper.elapsedMs.median / univer.elapsedMs.median,
      workpaperToUniverP95Ratio: workpaper.elapsedMs.p95 / univer.elapsedMs.p95,
    },
    engines: {
      univer,
      workpaper,
    },
  }
}

function measureWorkPaperSample(scenario: WorkPaperUniverScenario): BenchmarkSample {
  if (scenario.kind === 'build') {
    return measureWorkPaperBuildFromSheets(scenario.buildWorkPaperSheets(), scenario.verifyWorkPaper, scenario.workpaperOptions)
  }
  const workbook = WorkPaper.buildFromSheets(scenario.buildWorkPaperSheets(), scenario.workpaperOptions)
  const edit = requireMutationEdit(scenario.fixture)
  const editSheetId = workbook.getSheetId(edit.sheetName)
  if (editSheetId === undefined) {
    workbook.dispose()
    throw new Error(`WorkPaper Univer benchmark fixture did not create ${edit.sheetName}`)
  }
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents({ sheet: editSheetId, row: edit.row, col: edit.col }, edit.value),
    () => scenario.verifyWorkPaper(workbook),
  )
}

async function measureUniverSample(
  scenario: WorkPaperUniverScenario,
  expectedVerification: Record<string, unknown>,
): Promise<BenchmarkSample> {
  if (scenario.kind === 'build') {
    return measureUniverBuildSample(scenario, expectedVerification)
  }
  return measureUniverRecalcSample(scenario, expectedVerification)
}

async function measureUniverBuildSample(
  scenario: WorkPaperUniverScenario,
  expectedVerification: Record<string, unknown>,
): Promise<BenchmarkSample> {
  const sheetName = scenario.fixture.result?.sheetName ?? 'Bench'
  const columnCount = scenario.fixture.columnCount ?? 5
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const runtime = createUniverRuntime(scenario.fixture.rowCount, columnCount, sheetName)
  try {
    await scenario.setupUniver(runtime)
    const verification = await waitForUniverVerification(runtime, scenario, expectedVerification, univerCalculationTimeoutMs)
    const elapsedMs = performance.now() - started
    const memoryAfter = sampleMemory()

    return {
      elapsedMs,
      memory: measureMemory(memoryBefore, memoryAfter),
      verification,
    }
  } finally {
    runtime.workbook.dispose()
    runtime.univer.dispose()
  }
}

async function measureUniverRecalcSample(
  scenario: WorkPaperUniverScenario,
  expectedVerification: Record<string, unknown>,
): Promise<BenchmarkSample> {
  const edit = requireMutationEdit(scenario.fixture)
  const runtime = createUniverRuntime(
    scenario.fixture.rowCount,
    scenario.fixture.columnCount ?? Math.max((scenario.fixture.result?.col ?? 0) + 1, edit.col + 1, 5),
    edit.sheetName,
  )
  try {
    await scenario.setupUniver(runtime)
    const memoryBefore = sampleMemory()
    const started = performance.now()
    runtime.sheet.getRange(edit.address).setValue(edit.value)
    const verification = await waitForUniverVerification(runtime, scenario, expectedVerification, univerCalculationTimeoutMs)
    const elapsedMs = performance.now() - started
    const memoryAfter = sampleMemory()

    return {
      elapsedMs,
      memory: measureMemory(memoryBefore, memoryAfter),
      verification,
    }
  } finally {
    runtime.workbook.dispose()
    runtime.univer.dispose()
  }
}

function requireMutationEdit(fixture: WorkPaperUniverFixture): NonNullable<WorkPaperUniverFixture['edit']> {
  if (fixture.edit === undefined) {
    throw new Error('Expected Univer mutation workload fixture to include an edit')
  }
  return fixture.edit
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

  return summarizeBenchmarkSamples(samples)
}

async function benchmarkSupportedEngineAsync(
  runSample: () => Promise<BenchmarkSample>,
  options: ResolvedBenchmarkSuiteOptions,
): Promise<ComparativeMeasuredEngineResult> {
  for (let warmup = 0; warmup < options.warmupCount; warmup += 1) {
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Warmups must finish before measured samples begin.
    await runSample()
  }

  const samples: BenchmarkSample[] = []
  for (let sample = 0; sample < options.sampleCount; sample += 1) {
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential samples avoid overlapping Univer async recalculation work.
    samples.push(await runSample())
  }

  return summarizeBenchmarkSamples(samples)
}

function summarizeBenchmarkSamples(samples: readonly BenchmarkSample[]): ComparativeMeasuredEngineResult {
  const verificationStrings = new Set(samples.map((sample) => JSON.stringify(sample.verification)))
  if (verificationStrings.size !== 1) {
    throw new Error('Benchmark verification drifted across samples')
  }

  return {
    status: 'supported',
    elapsedMs: summarizeNumbers(samples.map((sample) => sample.elapsedMs)),
    memoryDeltaBytes: summarizeMemory(samples.map((sample) => sample.memory)),
    verification: samples[0]?.verification ?? {},
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

function univerScenario(workload: WorkPaperUniverWorkload): WorkPaperUniverScenario {
  switch (workload) {
    case 'build-from-sheets':
      return denseLiteralBuildScenario(workload, 160, 24)
    case 'build-dense-literals':
      return denseLiteralBuildScenario(workload, 160, 24)
    case 'build-dense-literals-wide':
      return denseLiteralBuildScenario(workload, 96, 96)
    case 'build-dense-literals-tall':
      return denseLiteralBuildScenario(workload, 768, 12)
    case 'build-mixed-content':
      return mixedContentBuildScenario(workload, 750)
    case 'build-mixed-content-small':
      return mixedContentBuildScenario(workload, 250)
    case 'build-mixed-content-large':
      return mixedContentBuildScenario(workload, 1_500)
    case 'build-parser-cache-row-templates':
      return parserCacheRowTemplateBuildScenario(workload, 1_500)
    case 'build-parser-cache-mixed-templates':
      return parserCacheMixedTemplateBuildScenario(workload, 1_500)
    case 'build-parser-cache-unique-formulas':
      return parserCacheUniqueFormulaBuildScenario(workload, 1_500)
    case 'build-many-sheets':
      return manySheetsBuildScenario(workload, 6, 96, 16)
    case 'build-many-sheets-wide':
      return manySheetsBuildScenario(workload, 4, 64, 48)
    case 'build-many-sheets-narrow':
      return manySheetsBuildScenario(workload, 12, 128, 8)
    case 'build-cross-sheet-dashboard':
      return crossSheetDashboardBuildScenario(workload, 4, 500)
    case 'build-cross-sheet-dashboard-small':
      return crossSheetDashboardBuildScenario(workload, 2, 250)
    case 'build-cross-sheet-dashboard-large':
      return crossSheetDashboardBuildScenario(workload, 6, 750)
    case 'cross-sheet-scalar-recalc':
      return crossSheetScalarFanoutRecalcScenario(workload, 1_500)
    case 'cross-sheet-aggregate-recalc':
      return crossSheetAggregateRecalcScenario(workload, 1_500)
    case 'cross-sheet-dashboard-recalc':
      return crossSheetDashboardRecalcScenario(workload, 4, 1_000)
    case 'single-edit-chain':
      return formulaChainRowScenario(workload, 2_000)
    case 'single-edit-chain-small':
      return formulaChainRowScenario(workload, 500)
    case 'single-edit-chain-large':
      return formulaChainRowScenario(workload, 3_000)
    case 'single-edit-fanout':
      return formulaFanoutRowScenario(workload, 2_000)
    case 'single-edit-fanout-small':
      return formulaFanoutRowScenario(workload, 500)
    case 'single-edit-fanout-large':
      return formulaFanoutRowScenario(workload, 3_000)
    case 'aggregate-2d-ranges':
      return aggregate2dCanonicalScenario(workload, 1_500)
    case 'aggregate-2d-ranges-small':
      return aggregate2dCanonicalScenario(workload, 500)
    case 'aggregate-2d-ranges-large':
      return aggregate2dCanonicalScenario(workload, 3_000)
    case 'aggregate-overlapping-ranges':
      return overlappingAggregateCanonicalScenario(workload, 1_500)
    case 'aggregate-overlapping-ranges-small':
      return overlappingAggregateCanonicalScenario(workload, 500)
    case 'aggregate-overlapping-sliding-window':
      return slidingAggregateCanonicalScenario(workload, 1_500, 32)
    case 'aggregate-overlapping-sliding-window-wide':
      return slidingAggregateCanonicalScenario(workload, 1_500, 128)
    case 'lookup-no-column-index':
      return exactLookupCanonicalScenario(workload, 5_000, false)
    case 'lookup-no-column-index-small':
      return exactLookupCanonicalScenario(workload, 1_000, false)
    case 'lookup-with-column-index':
      return exactLookupCanonicalScenario(workload, 5_000, true)
    case 'lookup-with-column-index-large':
      return exactLookupCanonicalScenario(workload, 10_000, true)
    case 'lookup-index-match-exact':
      return indexMatchExactCanonicalScenario(workload, 5_000)
    case 'lookup-index-match-exact-large':
      return indexMatchExactCanonicalScenario(workload, 10_000)
    case 'lookup-index-reference':
      return indexReferenceCanonicalScenario(workload, 5_000)
    case 'lookup-index-reference-large':
      return indexReferenceCanonicalScenario(workload, 10_000)
    case 'lookup-approximate-sorted':
      return approximateLookupCanonicalScenario(workload, 5_000)
    case 'lookup-approximate-sorted-large':
      return approximateLookupCanonicalScenario(workload, 10_000)
    case 'lookup-approximate-duplicates':
      return approximateDuplicateLookupCanonicalScenario(workload, 5_000)
    case 'lookup-text-exact':
      return textLookupCanonicalScenario(workload, 5_000)
    case 'lookup-text-exact-large':
      return textLookupCanonicalScenario(workload, 10_000)
  }
}

function createUniverRuntime(rowCount: number, columnCount: number, sheetName = 'Sheet1'): UniverRuntime {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: mergeLocales(sheetsNodeCoreEnUS),
    },
    presets: [UniverSheetsNodeCorePreset()],
  })
  const workbook = univerAPI.createWorkbook({
    id: 'bench',
    name: 'bench',
    sheetOrder: ['sheet1'],
    sheets: {
      sheet1: {
        id: 'sheet1',
        name: sheetName,
        rowCount,
        columnCount,
        cellData: {},
      },
    },
  })
  return {
    formula: univerAPI.getFormula(),
    sheet: workbook.getActiveSheet(),
    univer,
    workbook,
  }
}

function resolveModuleExports<T extends object>(moduleExports: object, requiredKey: keyof T & string): T {
  if (Reflect.has(moduleExports, requiredKey)) {
    // oxlint-disable-next-line typescript-eslint(no-unsafe-type-assertion) -- Univer publishes mixed CJS/ESM shapes; this narrows after checking the requested export key.
    return moduleExports as T
  }
  const defaultExport = Reflect.get(moduleExports, 'default')
  // oxlint-disable-next-line typescript-eslint(no-unsafe-type-assertion) -- Fall back to Univer's default export object when direct namespace exports are absent.
  return (isRecord(defaultExport) ? defaultExport : moduleExports) as T
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function geometricMean(values: readonly number[]): number {
  if (values.length === 0) {
    return Number.NaN
  }
  const totalLog = values.reduce((sum, value) => {
    if (value <= 0) {
      throw new Error(`Cannot compute geomean for non-positive value: ${String(value)}`)
    }
    return sum + Math.log(value)
  }, 0)
  return Math.exp(totalLog / values.length)
}

function maxComparableRatio(
  results: readonly WorkPaperUniverBenchmarkResult[],
  ratioKey: 'workpaperToUniverMeanRatio' | 'workpaperToUniverP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly WorkPaperUniverBenchmarkResult[],
  ratioKey: 'workpaperToUniverMeanRatio' | 'workpaperToUniverP95Ratio',
): WorkPaperUniverWorkload {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}

function orderedUnique(values: readonly WorkPaperUniverWorkloadFamily[]): WorkPaperUniverWorkloadFamily[] {
  return [...new Set(values)]
}

function resolveSuiteOptions(options: ComparativeBenchmarkSuiteOptions): ResolvedBenchmarkSuiteOptions {
  return {
    sampleCount: options.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: options.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  }
}
