import { describe, expect, it } from 'vitest'

import {
  WORKPAPER_UNIVER_WORKLOADS,
  buildWorkPaperVsUniverBenchmarkReport,
  runWorkPaperVsUniverBenchmarkSuite,
  type WorkPaperUniverWorkload,
} from '../benchmark-workpaper-vs-univer.js'

const expectedVerificationByWorkload = {
  'build-from-sheets': { terminalValue: 3_840 },
  'build-dense-literals': { terminalValue: 3_840 },
  'build-dense-literals-wide': { terminalValue: 9_216 },
  'build-dense-literals-tall': { terminalValue: 9_216 },
  'build-mixed-content': { terminalFormulaValue: 16_500 },
  'build-mixed-content-small': { terminalFormulaValue: 5_500 },
  'build-mixed-content-large': { terminalFormulaValue: 33_000 },
  'build-parser-cache-row-templates': { terminalValue: 1_134_750 },
  'build-parser-cache-mixed-templates': { terminalValue: 6_000 },
  'build-parser-cache-unique-formulas': { terminalValue: 1_142_250 },
  'build-many-sheets': { terminalValue: 9_216 },
  'build-many-sheets-wide': { terminalValue: 12_288 },
  'build-many-sheets-narrow': { terminalValue: 12_288 },
  'build-cross-sheet-dashboard': { leadingDataTotal: 125_250, terminalDataTotal: 2_500 },
  'single-edit-chain': { terminalValue: 2_099 },
  'single-edit-chain-small': { terminalValue: 599 },
  'single-edit-chain-large': { terminalValue: 3_099 },
  'single-edit-fanout': { terminalValue: 2_099, width: 2_001 },
  'single-edit-fanout-small': { terminalValue: 599, width: 501 },
  'single-edit-fanout-large': { terminalValue: 3_099, width: 3_001 },
  'aggregate-2d-ranges': { leadingSum: 101, terminalSum: 3_377_348 },
  'aggregate-2d-ranges-small': { leadingSum: 101, terminalSum: 375_848 },
  'aggregate-2d-ranges-large': { leadingSum: 101, terminalSum: 13_504_598 },
  'aggregate-overlapping-ranges': { terminalSum: 1_125_848 },
  'aggregate-overlapping-ranges-small': { terminalSum: 125_348 },
  'aggregate-overlapping-sliding-window': { leadingSum: 626, terminalSum: 1_500 },
  'aggregate-overlapping-sliding-window-wide': { leadingSum: 8_354, terminalSum: 1_500 },
  'lookup-no-column-index': { formulaValue: 5_000 },
  'lookup-no-column-index-small': { formulaValue: 1_000 },
  'lookup-with-column-index': { formulaValue: 5_000 },
  'lookup-with-column-index-large': { formulaValue: 10_000 },
  'lookup-approximate-sorted': { formulaValue: 4_999 },
  'lookup-approximate-sorted-large': { formulaValue: 9_999 },
  'lookup-approximate-duplicates': { formulaValue: 2_000 },
  'lookup-text-exact': { formulaValue: 4_999 },
  'lookup-text-exact-large': { formulaValue: 9_999 },
} as const satisfies Record<WorkPaperUniverWorkload, Record<string, unknown>>

const expectedWorkloads = [
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
  'build-many-sheets',
  'build-many-sheets-wide',
  'build-many-sheets-narrow',
  'build-cross-sheet-dashboard',
] as const satisfies readonly WorkPaperUniverWorkload[]

const benchmarkTimeoutMs = 60_000
let benchmarkResultsPromise: ReturnType<typeof runWorkPaperVsUniverBenchmarkSuite> | undefined

function benchmarkResults(): ReturnType<typeof runWorkPaperVsUniverBenchmarkSuite> {
  benchmarkResultsPromise ??= runWorkPaperVsUniverBenchmarkSuite({ sampleCount: 1, warmupCount: 0 })
  return benchmarkResultsPromise
}

describe('WorkPaper vs Univer benchmark', () => {
  it(
    'covers the documented Univer Node preset with workbook recalculation workloads',
    async () => {
      expect(WORKPAPER_UNIVER_WORKLOADS).toEqual(expectedWorkloads)

      const results = await benchmarkResults()

      expect(results).toHaveLength(WORKPAPER_UNIVER_WORKLOADS.length)
      for (const result of results) {
        const expectedVerification = expectedVerificationByWorkload[result.workload]
        expect(result.category).toBe('workbook-wide')
        expect(result.comparable).toBe(true)
        expect(result.comparison.verificationEquivalent).toBe(true)
        expect(result.engines.workpaper.verification).toEqual(expectedVerification)
        expect(result.engines.univer.verification).toEqual(expectedVerification)
        expect(result.fixture.rowCount).toBeGreaterThanOrEqual(1)
        expect(Number.isFinite(result.comparison.workpaperToUniverMeanRatio)).toBe(true)
        expect(Number.isFinite(result.comparison.workpaperToUniverP95Ratio)).toBe(true)
      }
    },
    benchmarkTimeoutMs,
  )

  it(
    'derives scorecard totals from measured Univer results',
    async () => {
      const results = await benchmarkResults()
      const report = buildWorkPaperVsUniverBenchmarkReport(results)

      expect(report.suite).toBe('workpaper-vs-univer')
      expect(report.scorecard.coverageTier).toBe('workbook-wide')
      expect(report.scorecard.comparableWorkloadCount).toBe(WORKPAPER_UNIVER_WORKLOADS.length)
      expect(report.scorecard.meanWinCount + report.scorecard.univerMeanWinCount).toBe(WORKPAPER_UNIVER_WORKLOADS.length)
      expect(report.scorecard.p95WinCount + report.scorecard.univerP95WinCount).toBe(WORKPAPER_UNIVER_WORKLOADS.length)
      expect(report.scorecard.meanAndP95WinCount).toBeLessThanOrEqual(WORKPAPER_UNIVER_WORKLOADS.length)
      expect(report.scorecard.workloadFamilies).toEqual([
        'build',
        'formula-chain',
        'formula-fanout',
        'aggregate-2d',
        'overlapping-aggregate',
        'lookup-exact',
        'lookup-approximate',
        'cross-sheet',
      ])
    },
    benchmarkTimeoutMs,
  )
})
