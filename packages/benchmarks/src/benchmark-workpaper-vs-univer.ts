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
  measureMutationSample,
  normalizeWorkPaperValue,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import {
  columnName,
  deepChainFormulas,
  duplicateApproximateLookupTableValues,
  evenLookupTableValues,
  formulaChainFormulas,
  hlookupTableValues,
  lookupTableValues,
  normalizeBenchmarkValue,
  numberColumnSheet,
  numberColumnValues,
  overlappingAggregateFormulas,
  rangeStatsValues,
  scalarFanoutFormulas,
  textLookupKey,
  textLookupTableValues,
  twoDimensionalValues,
} from './benchmark-workpaper-vs-univer-fixtures.js'
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
  | 'univer-large-sum-recalc'
  | 'univer-2d-sum-recalc'
  | 'univer-overlapping-sum-recalc'
  | 'univer-formula-chain-recalc'
  | 'univer-scalar-fanout-recalc'
  | 'univer-deep-chain-recalc'
  | 'univer-vlookup-exact-recalc'
  | 'univer-vlookup-approximate-recalc'
  | 'univer-vlookup-text-exact-recalc'
  | 'univer-vlookup-approximate-duplicates-recalc'
  | 'univer-hlookup-exact-numeric-recalc'
  | 'univer-range-stats-recalc'

export type WorkPaperUniverWorkloadFamily =
  | 'aggregate'
  | 'aggregate-2d'
  | 'formula-chain'
  | 'formula-fanout'
  | 'lookup-approximate'
  | 'lookup-exact'
  | 'overlapping-aggregate'
  | 'range-stats'

export type UniverEditableValue = boolean | number | string

export interface WorkPaperUniverFixture {
  readonly edit: {
    readonly address: string
    readonly col: number
    readonly row: number
    readonly sheetName: string
    readonly value: UniverEditableValue
  }
  readonly family: WorkPaperUniverWorkloadFamily
  readonly formula: string
  readonly result: {
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
  readonly fixture: WorkPaperUniverFixture
  readonly buildWorkPaperSheets: () => Record<string, WorkPaperSheet>
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
  setValues(values: UniverEditableValue[][]): UniverRangeFacade
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
  'univer-large-sum-recalc',
  'univer-2d-sum-recalc',
  'univer-overlapping-sum-recalc',
  'univer-formula-chain-recalc',
  'univer-scalar-fanout-recalc',
  'univer-deep-chain-recalc',
  'univer-vlookup-exact-recalc',
  'univer-vlookup-approximate-recalc',
  'univer-vlookup-text-exact-recalc',
  'univer-vlookup-approximate-duplicates-recalc',
  'univer-hlookup-exact-numeric-recalc',
  'univer-range-stats-recalc',
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
        'Univer is covered through its documented @univerjs/preset-sheets-node-core headless Node preset. This workbook-wide lane times public Facade API cell edits plus formula recalculation completion on equivalent aggregate, chain, fanout, and lookup workloads.',
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
  const workpaper = benchmarkSupportedEngine(() => measureWorkPaperRecalcSample(scenario), options)
  const univer = await benchmarkSupportedEngineAsync(() => measureUniverRecalcSample(scenario), options)
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

function measureWorkPaperRecalcSample(scenario: WorkPaperUniverScenario): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(scenario.buildWorkPaperSheets())
  const editSheetId = workbook.getSheetId(scenario.fixture.edit.sheetName)
  if (editSheetId === undefined) {
    workbook.dispose()
    throw new Error(`WorkPaper Univer benchmark fixture did not create ${scenario.fixture.edit.sheetName}`)
  }
  return measureMutationSample(
    workbook,
    () =>
      workbook.setCellContents(
        { sheet: editSheetId, row: scenario.fixture.edit.row, col: scenario.fixture.edit.col },
        scenario.fixture.edit.value,
      ),
    () => scenario.verifyWorkPaper(workbook),
  )
}

async function measureUniverRecalcSample(scenario: WorkPaperUniverScenario): Promise<BenchmarkSample> {
  const runtime = createUniverRuntime(
    scenario.fixture.rowCount,
    scenario.fixture.columnCount ?? Math.max(scenario.fixture.result.col + 1, scenario.fixture.edit.col + 1, 5),
  )
  try {
    await scenario.setupUniver(runtime)
    const completion = waitForUniverCalculation(runtime.formula, scenario.fixture.formula)
    const memoryBefore = sampleMemory()
    const started = performance.now()
    runtime.sheet.getRange(scenario.fixture.edit.address).setValue(scenario.fixture.edit.value)
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
    case 'univer-large-sum-recalc':
      return aggregateScenario(5_000)
    case 'univer-2d-sum-recalc':
      return twoDimensionalAggregateScenario(1_000, 8)
    case 'univer-overlapping-sum-recalc':
      return overlappingAggregateScenario(1_500)
    case 'univer-formula-chain-recalc':
      return formulaChainScenario(1_000)
    case 'univer-scalar-fanout-recalc':
      return scalarFanoutScenario(1_000)
    case 'univer-deep-chain-recalc':
      return deepChainScenario(1_000)
    case 'univer-vlookup-exact-recalc':
      return vlookupExactScenario(5_000)
    case 'univer-vlookup-approximate-recalc':
      return vlookupApproximateScenario(5_000)
    case 'univer-vlookup-text-exact-recalc':
      return vlookupTextExactScenario(5_000)
    case 'univer-vlookup-approximate-duplicates-recalc':
      return vlookupApproximateDuplicatesScenario(5_000)
    case 'univer-hlookup-exact-numeric-recalc':
      return hlookupExactNumericScenario(3_000)
    case 'univer-range-stats-recalc':
      return rangeStatsScenario(1_000)
  }
}

function aggregateScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=SUM(A1:A${String(rowCount)})`
  const fixture = {
    edit: { address: 'A2500', col: 0, row: 2_499, sheetName: 'Sheet1', value: 10_000 },
    family: 'aggregate',
    formula,
    result: { address: 'C1', col: 2, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet = numberColumnSheet(rowCount)
      sheet[0]![2] = formula
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, rowCount, 1).setValues(numberColumnValues(rowCount))
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(fixture.result.address).setFormula(formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function twoDimensionalAggregateScenario(rowCount: number, colCount: number): WorkPaperUniverScenario {
  const lastCol = columnName(colCount - 1)
  const formula = `=SUM(A1:${lastCol}${String(rowCount)})`
  const resultAddress = `${columnName(colCount + 1)}1`
  const fixture = {
    edit: { address: 'D500', col: 3, row: 499, sheetName: 'Sheet1', value: 10_000 },
    family: 'aggregate-2d',
    formula,
    result: { address: resultAddress, col: colCount + 1, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  const values = twoDimensionalValues(rowCount, colCount)
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = values.map((row) => [...row])
      sheet[0]![colCount + 1] = formula
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, rowCount, colCount).setValues(values)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(fixture.result.address).setFormula(formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function overlappingAggregateScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=B${String(rowCount)}`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Sheet1', value: 99 },
    family: 'overlapping-aggregate',
    formula,
    result: { address: 'C1', col: 2, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  const formulas = overlappingAggregateFormulas(rowCount)
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet = numberColumnSheet(rowCount)
      for (let row = 0; row < rowCount; row += 1) {
        sheet[row]![1] = formulas[row]![0]!
      }
      sheet[0]![2] = formula
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, rowCount, 1).setValues(numberColumnValues(rowCount))
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(0, 1, rowCount, 1).setFormulas(formulas)
      runtime.sheet.getRange(fixture.result.address).setFormula(formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function formulaChainScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=B${String(rowCount - 1)}+1`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Sheet1', value: 10_000 },
    family: 'formula-chain',
    formula,
    result: { address: `B${String(rowCount)}`, col: 1, row: rowCount - 1, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  const formulas = formulaChainFormulas(rowCount)
  return {
    fixture,
    buildWorkPaperSheets: () => ({
      Sheet1: [[1, '=A1+1'], ...formulas.slice(1).map((row) => [null, row[0]!])],
    }),
    setupUniver: async (runtime) => {
      runtime.sheet.getRange('A1').setValue(1)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(0, 1, rowCount, 1).setFormulas(formulas)
      await completion
    },
    verifyUniver: (runtime) => ({
      checksum: normalizeBenchmarkValue(runtime.sheet.getRange('B1').getValue()),
      value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()),
    }),
    verifyWorkPaper: (workbook) => ({
      checksum: normalizeBenchmarkValue(
        normalizeWorkPaperValue(workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: 0, col: 1 })),
      ),
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function scalarFanoutScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=A1+${String(rowCount)}`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Sheet1', value: 5_000 },
    family: 'formula-fanout',
    formula,
    result: { address: `B${String(rowCount)}`, col: 1, row: rowCount - 1, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  const formulas = scalarFanoutFormulas(rowCount)
  return {
    fixture,
    buildWorkPaperSheets: () => ({
      Sheet1: [[1, '=A1+1'], ...formulas.slice(1).map((row) => [null, row[0]!])],
    }),
    setupUniver: async (runtime) => {
      runtime.sheet.getRange('A1').setValue(1)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(0, 1, rowCount, 1).setFormulas(formulas)
      await completion
    },
    verifyUniver: (runtime) => ({
      checksum: normalizeBenchmarkValue(runtime.sheet.getRange('B1').getValue()),
      value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()),
    }),
    verifyWorkPaper: (workbook) => ({
      checksum: normalizeBenchmarkValue(
        normalizeWorkPaperValue(workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: 0, col: 1 })),
      ),
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function deepChainScenario(chainLength: number): WorkPaperUniverScenario {
  const resultCol = chainLength
  const resultAddress = `${columnName(resultCol)}1`
  const formula = `=${columnName(resultCol - 1)}1+1`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Sheet1', value: 5_000 },
    family: 'formula-chain',
    formula,
    result: { address: resultAddress, col: resultCol, row: 0, sheetName: 'Sheet1' },
    rowCount: chainLength,
  } as const satisfies WorkPaperUniverFixture
  const formulas = deepChainFormulas(chainLength)
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const row: Array<number | string | null> = [1]
      for (let index = 1; index <= chainLength; index += 1) {
        row[index] = formulas[0]![index - 1]!
      }
      return { Sheet1: [row] }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange('A1').setValue(1)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(0, 1, 1, chainLength).setFormulas(formulas)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function vlookupExactScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=VLOOKUP(D1,A1:B${String(rowCount)},2,FALSE)`
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: 3_100 },
    family: 'lookup-exact',
    formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  const values = lookupTableValues(rowCount)
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = values.map((row) => [...row])
      sheet[0]![3] = 2_400
      sheet[0]![4] = formula
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, rowCount, 2).setValues(values)
      runtime.sheet.getRange('D1').setValue(2_400)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(fixture.result.address).setFormula(formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function vlookupApproximateScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=VLOOKUP(D1,A1:B${String(rowCount)},2,TRUE)`
  return numericVlookupScenario({
    editValue: 6_201,
    family: 'lookup-approximate',
    formula,
    initialLookupValue: 2_401,
    rowCount,
  })
}

function numericVlookupScenario(args: {
  readonly editValue: number
  readonly family: 'lookup-approximate'
  readonly formula: string
  readonly initialLookupValue: number
  readonly rowCount: number
}): WorkPaperUniverScenario {
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: args.editValue },
    family: args.family,
    formula: args.formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount: args.rowCount,
  } as const satisfies WorkPaperUniverFixture
  const values = evenLookupTableValues(args.rowCount)
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = values.map((row) => [...row])
      sheet[0] = [2, 10, null, args.initialLookupValue, args.formula]
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, args.rowCount, 2).setValues(values)
      runtime.sheet.getRange('D1').setValue(args.initialLookupValue)
      const completion = waitForUniverCalculation(runtime.formula, args.formula)
      runtime.sheet.getRange(fixture.result.address).setFormula(args.formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function vlookupTextExactScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=VLOOKUP(D1,A1:B${String(rowCount)},2,FALSE)`
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: textLookupKey(3_100) },
    family: 'lookup-exact',
    formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  return textLookupScenario({ fixture, formula, initialLookupValue: textLookupKey(2_400), rowCount })
}

function textLookupScenario(args: {
  readonly fixture: WorkPaperUniverFixture
  readonly formula: string
  readonly initialLookupValue: string
  readonly rowCount: number
}): WorkPaperUniverScenario {
  const values = textLookupTableValues(args.rowCount)
  return {
    fixture: args.fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = values.map((row) => [...row])
      sheet[0] = [textLookupKey(1), 10, null, args.initialLookupValue, args.formula]
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, args.rowCount, 2).setValues(values)
      runtime.sheet.getRange('D1').setValue(args.initialLookupValue)
      const completion = waitForUniverCalculation(runtime.formula, args.formula)
      runtime.sheet.getRange(args.fixture.result.address).setFormula(args.formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(args.fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: args.fixture.result.row, col: args.fixture.result.col }),
        ),
      ),
    }),
  }
}

function vlookupApproximateDuplicatesScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=VLOOKUP(D1,A1:B${String(rowCount)},2,TRUE)`
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: 1_550.5 },
    family: 'lookup-approximate',
    formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  const values = duplicateApproximateLookupTableValues(rowCount)
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = values.map((row) => [...row])
      sheet[0] = [1, 1, null, 1_200.5, formula]
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, rowCount, 2).setValues(values)
      runtime.sheet.getRange('D1').setValue(1_200.5)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(fixture.result.address).setFormula(formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function hlookupExactNumericScenario(colCount: number): WorkPaperUniverScenario {
  const lastCol = columnName(colCount - 1)
  const formula = `=HLOOKUP(A4,A1:${lastCol}2,2,FALSE)`
  const fixture = {
    edit: { address: 'A4', col: 0, row: 3, sheetName: 'Sheet1', value: 620 },
    family: 'lookup-exact',
    formula,
    result: { address: 'B4', col: 1, row: 3, sheetName: 'Sheet1' },
    columnCount: colCount,
    rowCount: colCount,
  } as const satisfies WorkPaperUniverFixture
  const values = hlookupTableValues(colCount)
  return {
    fixture,
    buildWorkPaperSheets: () => ({ Sheet1: [values[0]!, values[1]!, [], [200, formula]] }),
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, 2, colCount).setValues(values)
      runtime.sheet.getRange('A4').setValue(200)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(fixture.result.address).setFormula(formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function rangeStatsScenario(rowCount: number): WorkPaperUniverScenario {
  const formula = `=AVERAGE(A1:A${String(rowCount)})+MAX(B1:B${String(rowCount)})-MIN(C1:C${String(rowCount)})`
  const fixture = {
    edit: { address: 'B1000', col: 1, row: 999, sheetName: 'Sheet1', value: 9_999 },
    family: 'range-stats',
    formula,
    result: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperUniverFixture
  const values = rangeStatsValues(rowCount)
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = values.map((row) => [...row])
      sheet[0]![3] = formula
      return { Sheet1: sheet }
    },
    setupUniver: async (runtime) => {
      runtime.sheet.getRange(0, 0, rowCount, 3).setValues(values)
      const completion = waitForUniverCalculation(runtime.formula, formula)
      runtime.sheet.getRange(fixture.result.address).setFormula(formula)
      await completion
    },
    verifyUniver: (runtime) => ({ value: normalizeBenchmarkValue(runtime.sheet.getRange(fixture.result.address).getValue()) }),
    verifyWorkPaper: (workbook) => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: workbook.getSheetId('Sheet1')!, row: fixture.result.row, col: fixture.result.col }),
        ),
      ),
    }),
  }
}

function createUniverRuntime(rowCount: number, columnCount: number): UniverRuntime {
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
        name: 'Sheet1',
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
