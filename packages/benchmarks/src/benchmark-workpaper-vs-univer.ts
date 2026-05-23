import { performance } from 'node:perf_hooks'
import * as UniverPresetModule from '@univerjs/presets'
import * as UniverSheetsNodeCorePresetModule from '@univerjs/preset-sheets-node-core'
import sheetsNodeCoreEnUS from '@univerjs/preset-sheets-node-core/locales/en-US'
import { WorkPaper, type WorkPaperSheet } from '../../headless/src/work-paper.js'
import {
  build2dAggregateSheet,
  buildApproxLookupDuplicateSheet,
  buildApproxLookupSheet,
  buildDenseLiteralSheet,
  buildFormulaChainRow,
  buildFormulaFanoutRow,
  buildLookupSheet,
  buildMixedContentSheet,
  buildOverlappingAggregateSheet,
  buildParserCacheMixedTemplateSheet,
  buildParserCacheTemplateSheet,
  buildParserCacheUniqueFormulaSheet,
  buildSlidingAggregateSheet,
  buildTextLookupSheet,
  columnLabel,
  textLookupKey,
} from './workpaper-benchmark-fixtures.js'
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
  normalizeWorkPaperValue,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { normalizeBenchmarkValue } from './benchmark-workpaper-vs-univer-fixtures.js'
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

interface WorkPaperUniverScenario {
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

interface UniverRuntime {
  readonly formula: UniverFormulaFacade
  readonly sheet: UniverWorksheetFacade
  readonly univer: { dispose(): void }
  readonly workbook: UniverWorkbookFacade
}

interface UniverFormulaFacade {
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

interface UniverWorksheetFacade {
  getRange(row: number, column: number): UniverRangeFacade
  getRange(row: number, column: number, numRows: number): UniverRangeFacade
  getRange(row: number, column: number, numRows: number, numColumns: number): UniverRangeFacade
  getRange(a1Notation: string): UniverRangeFacade
}

interface UniverWorkbookFacade {
  dispose(): void
  getActiveSheet(): UniverWorksheetFacade
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
  'lookup-approximate-sorted',
  'lookup-approximate-sorted-large',
  'lookup-approximate-duplicates',
  'lookup-text-exact',
  'lookup-text-exact-large',
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
  const univer = await benchmarkSupportedEngineAsync(() => measureUniverSample(scenario), options)
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

async function measureUniverSample(scenario: WorkPaperUniverScenario): Promise<BenchmarkSample> {
  if (scenario.kind === 'build') {
    return measureUniverBuildSample(scenario)
  }
  return measureUniverRecalcSample(scenario)
}

async function measureUniverBuildSample(scenario: WorkPaperUniverScenario): Promise<BenchmarkSample> {
  const sheetName = scenario.fixture.result?.sheetName ?? 'Bench'
  const columnCount = scenario.fixture.columnCount ?? 5
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const runtime = createUniverRuntime(scenario.fixture.rowCount, columnCount, sheetName)
  try {
    await scenario.setupUniver(runtime)
    const elapsedMs = performance.now() - started
    const memoryAfter = sampleMemory()

    return {
      elapsedMs,
      memory: measureMemory(memoryBefore, memoryAfter),
      verification: scenario.verifyUniver(runtime),
    }
  } finally {
    runtime.workbook.dispose()
    runtime.univer.dispose()
  }
}

async function measureUniverRecalcSample(scenario: WorkPaperUniverScenario): Promise<BenchmarkSample> {
  const edit = requireMutationEdit(scenario.fixture)
  const runtime = createUniverRuntime(
    scenario.fixture.rowCount,
    scenario.fixture.columnCount ?? Math.max((scenario.fixture.result?.col ?? 0) + 1, edit.col + 1, 5),
    edit.sheetName,
  )
  try {
    await scenario.setupUniver(runtime)
    const completion = waitForUniverCalculation(runtime.formula, scenario.fixture.formula)
    const memoryBefore = sampleMemory()
    const started = performance.now()
    runtime.sheet.getRange(edit.address).setValue(edit.value)
    await completion
    const elapsedMs = performance.now() - started
    const memoryAfter = sampleMemory()

    return {
      elapsedMs,
      memory: measureMemory(memoryBefore, memoryAfter),
      verification: scenario.verifyUniver(runtime),
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
      return parserCacheTemplateBuildScenario(workload, buildParserCacheTemplateSheet(1_500), 1_500)
    case 'build-parser-cache-mixed-templates':
      return parserCacheTemplateBuildScenario(workload, buildParserCacheMixedTemplateSheet(1_500), 1_500)
    case 'build-parser-cache-unique-formulas':
      return parserCacheTemplateBuildScenario(workload, buildParserCacheUniqueFormulaSheet(1_500), 1_500)
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

function denseLiteralBuildScenario(workload: WorkPaperUniverWorkload, rows: number, cols: number): WorkPaperUniverScenario {
  return canonicalSingleSheetBuildScenario({
    family: 'build',
    observedCells: [{ col: cols - 1, key: 'terminalValue', row: rows - 1 }],
    rowCount: rows,
    sheet: buildDenseLiteralSheet(rows, cols),
    workload,
  })
}

function mixedContentBuildScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetBuildScenario({
    family: 'build',
    observedCells: [{ col: 5, key: 'terminalFormulaValue', row: rowCount - 1 }],
    rowCount,
    sheet: buildMixedContentSheet(rowCount),
    workload,
  })
}

function parserCacheTemplateBuildScenario(
  workload: WorkPaperUniverWorkload,
  sheet: WorkPaperSheet,
  rowCount: number,
): WorkPaperUniverScenario {
  return canonicalSingleSheetBuildScenario({
    family: 'build',
    observedCells: [{ col: 5, key: 'terminalValue', row: rowCount - 1 }],
    rowCount,
    sheet,
    workload,
  })
}

function formulaChainRowScenario(workload: WorkPaperUniverWorkload, downstreamCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'formula-chain',
    observedCells: [{ col: downstreamCount, key: 'terminalValue', row: 0 }],
    rowCount: 1,
    sheet: [buildFormulaChainRow(downstreamCount)],
    workload,
  })
}

function formulaFanoutRowScenario(workload: WorkPaperUniverWorkload, fanoutCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'formula-fanout',
    observedCells: [
      { col: fanoutCount, key: 'terminalValue', row: 0 },
      { col: fanoutCount, key: 'width', row: 0, value: fanoutCount + 1 },
    ],
    rowCount: 1,
    sheet: [buildFormulaFanoutRow(fanoutCount)],
    workload,
  })
}

function aggregate2dCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'aggregate-2d',
    observedCells: [
      { col: 2, key: 'terminalSum', row: rowCount - 1 },
      { col: 2, key: 'leadingSum', row: 0 },
    ],
    rowCount,
    sheet: build2dAggregateSheet(rowCount),
    workload,
  })
}

function overlappingAggregateCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'overlapping-aggregate',
    observedCells: [{ col: 1, key: 'terminalSum', row: rowCount - 1 }],
    rowCount,
    sheet: buildOverlappingAggregateSheet(rowCount),
    workload,
  })
}

function slidingAggregateCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number, window: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'overlapping-aggregate',
    observedCells: [
      { col: 1, key: 'terminalSum', row: rowCount - 1 },
      { col: 1, key: 'leadingSum', row: 0 },
    ],
    rowCount,
    sheet: buildSlidingAggregateSheet(rowCount, window),
    workload,
  })
}

function exactLookupCanonicalScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  useColumnIndex: boolean,
): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: rowCount },
    family: 'lookup-exact',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildLookupSheet(rowCount),
    workbookOptions: { useColumnIndex },
    workload,
  })
}

function approximateLookupCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: rowCount - 0.5 },
    family: 'lookup-approximate',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildApproxLookupSheet(rowCount),
    workload,
  })
}

function approximateDuplicateLookupCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: Math.floor(rowCount / 5) + 0.5 },
    family: 'lookup-approximate',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildApproxLookupDuplicateSheet(rowCount),
    workload,
  })
}

function textLookupCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: textLookupKey(rowCount - 1) },
    family: 'lookup-exact',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildTextLookupSheet(rowCount),
    workload,
  })
}

function canonicalSingleSheetBuildScenario(args: {
  readonly family: WorkPaperUniverWorkloadFamily
  readonly observedCells: readonly {
    readonly col: number
    readonly key: string
    readonly row: number
    readonly value?: number
  }[]
  readonly rowCount: number
  readonly sheet: WorkPaperSheet
  readonly workbookOptions?: Parameters<typeof WorkPaper.buildFromSheets>[1]
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const sheetName = 'Bench'
  const columnCount = Math.max(...args.observedCells.map((cell) => cell.col + 1), ...args.sheet.map((row) => row.length))
  const formula = firstFormula(args.sheet) ?? args.workload
  const fixture = {
    family: args.family,
    formula,
    result: {
      address: formatA1(args.observedCells[0]!.row, args.observedCells[0]!.col),
      col: args.observedCells[0]!.col,
      row: args.observedCells[0]!.row,
      sheetName,
    },
    columnCount,
    rowCount: args.rowCount,
  } as const satisfies WorkPaperUniverFixture
  const observeWorkPaper = (workbook: WorkPaper): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ??
          normalizeBenchmarkValue(
            normalizeWorkPaperValue(workbook.getCellValue({ sheet: workbook.getSheetId(sheetName)!, row: cell.row, col: cell.col })),
          ),
      ]),
    )
  const observeUniver = (runtime: UniverRuntime): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ?? normalizeBenchmarkValue(runtime.sheet.getRange(formatA1(cell.row, cell.col)).getValue()),
      ]),
    )
  return {
    kind: 'build',
    fixture,
    buildWorkPaperSheets: () => ({ [sheetName]: args.sheet }),
    ...(args.workbookOptions ? { workpaperOptions: args.workbookOptions } : {}),
    setupUniver: (runtime) => setupUniverSheet(runtime, args.sheet, formula),
    verifyUniver: observeUniver,
    verifyWorkPaper: observeWorkPaper,
  }
}

function canonicalSingleSheetScenario(args: {
  readonly edit: { readonly col: number; readonly row: number; readonly value: UniverEditableValue }
  readonly family: WorkPaperUniverWorkloadFamily
  readonly observedCells: readonly {
    readonly col: number
    readonly key: string
    readonly row: number
    readonly value?: number
  }[]
  readonly rowCount: number
  readonly sheet: WorkPaperSheet
  readonly workbookOptions?: Parameters<typeof WorkPaper.buildFromSheets>[1]
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const sheetName = 'Bench'
  const columnCount = Math.max(args.edit.col + 1, ...args.observedCells.map((cell) => cell.col + 1), ...args.sheet.map((row) => row.length))
  const formula = firstFormula(args.sheet) ?? args.workload
  const fixture = {
    edit: {
      address: formatA1(args.edit.row, args.edit.col),
      col: args.edit.col,
      row: args.edit.row,
      sheetName,
      value: args.edit.value,
    },
    family: args.family,
    formula,
    result: {
      address: formatA1(args.observedCells[0]!.row, args.observedCells[0]!.col),
      col: args.observedCells[0]!.col,
      row: args.observedCells[0]!.row,
      sheetName,
    },
    columnCount,
    rowCount: args.rowCount,
  } as const satisfies WorkPaperUniverFixture
  const observeWorkPaper = (workbook: WorkPaper): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ??
          normalizeBenchmarkValue(
            normalizeWorkPaperValue(workbook.getCellValue({ sheet: workbook.getSheetId(sheetName)!, row: cell.row, col: cell.col })),
          ),
      ]),
    )
  const observeUniver = (runtime: UniverRuntime): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ?? normalizeBenchmarkValue(runtime.sheet.getRange(formatA1(cell.row, cell.col)).getValue()),
      ]),
    )
  return {
    kind: 'mutation',
    fixture,
    buildWorkPaperSheets: () => ({ [sheetName]: args.sheet }),
    ...(args.workbookOptions ? { workpaperOptions: args.workbookOptions } : {}),
    setupUniver: (runtime) => setupUniverSheet(runtime, args.sheet, formula),
    verifyUniver: observeUniver,
    verifyWorkPaper: observeWorkPaper,
  }
}

async function setupUniverSheet(runtime: UniverRuntime, sheet: WorkPaperSheet, label: string): Promise<void> {
  const rowCount = sheet.length
  const columnCount = Math.max(1, ...sheet.map((row) => row.length))
  const formulaRuns: { readonly formulas: string[]; readonly row: number; readonly startCol: number }[] = []
  const values: UniverCellValue[][] = Array.from({ length: rowCount }, (_rowValue, row) =>
    Array.from({ length: columnCount }, (_colValue, col) => {
      const value = sheet[row]?.[col] ?? ''
      return typeof value === 'string' && value.startsWith('=') ? '' : value
    }),
  )
  for (let row = 0; row < sheet.length; row += 1) {
    const cells = sheet[row] ?? []
    let col = 0
    while (col < cells.length) {
      const value = cells[col]
      if (typeof value !== 'string' || !value.startsWith('=')) {
        col += 1
        continue
      }
      const startCol = col
      const formulas: string[] = []
      while (col < cells.length) {
        const formula = cells[col]
        if (typeof formula !== 'string' || !formula.startsWith('=')) {
          break
        }
        formulas.push(formula)
        col += 1
      }
      formulaRuns.push({ formulas, row, startCol })
    }
  }
  runtime.sheet.getRange(0, 0, rowCount, columnCount).setValues(values)

  if (formulaRuns.length === 0) {
    return
  }

  const completion = waitForUniverCalculation(runtime.formula, label)
  for (const run of formulaRuns) {
    runtime.sheet.getRange(run.row, run.startCol, 1, run.formulas.length).setFormulas([run.formulas])
  }
  await completion
}

function firstFormula(sheet: WorkPaperSheet): string | undefined {
  for (const row of sheet) {
    for (const value of row) {
      if (typeof value === 'string' && value.startsWith('=')) {
        return value
      }
    }
  }
  return undefined
}

function formatA1(row: number, col: number): string {
  return `${columnLabel(col)}${String(row + 1)}`
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

async function waitForUniverCalculation(formula: UniverFormulaFacade, label: string): Promise<void> {
  let timeout: ReturnType<typeof setTimeout> | undefined
  try {
    await Promise.race([
      formula.onCalculationEnd(),
      new Promise<never>((_, reject) => {
        timeout = setTimeout(
          () => reject(new Error(`Timed out waiting for Univer formula recalculation for ${label}`)),
          univerCalculationTimeoutMs,
        )
      }),
    ])
  } finally {
    if (timeout) {
      clearTimeout(timeout)
    }
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
