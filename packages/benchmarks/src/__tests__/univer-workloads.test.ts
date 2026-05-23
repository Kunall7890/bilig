import { describe, expect, it } from 'vitest'

import {
  WORKPAPER_UNIVER_WORKLOADS,
  buildWorkPaperVsUniverBenchmarkReport,
  runWorkPaperVsUniverBenchmarkSuite,
  type WorkPaperUniverWorkload,
} from '../benchmark-workpaper-vs-univer.js'

const expectedVerificationByWorkload = {
  'univer-large-sum-recalc': { value: 12_510_000 },
  'univer-2d-sum-recalc': { value: 18_026_000 },
  'univer-overlapping-sum-recalc': { value: 1_125_848 },
  'univer-formula-chain-recalc': { checksum: 10_001, value: 11_000 },
  'univer-scalar-fanout-recalc': { checksum: 5_001, value: 6_000 },
  'univer-deep-chain-recalc': { value: 6_000 },
  'univer-vlookup-exact-recalc': { value: 31_000 },
  'univer-vlookup-approximate-recalc': { value: 31_000 },
  'univer-vlookup-text-exact-recalc': { value: 31_000 },
  'univer-vlookup-approximate-duplicates-recalc': { value: 3_100 },
  'univer-hlookup-exact-numeric-recalc': { value: 3_100 },
  'univer-range-stats-recalc': { value: 10_498.5 },
} as const satisfies Record<WorkPaperUniverWorkload, Record<string, number>>

describe('WorkPaper vs Univer benchmark', () => {
  it('covers the documented Univer Node preset with workbook recalculation workloads', async () => {
    expect(WORKPAPER_UNIVER_WORKLOADS).toEqual([
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
    ])

    const results = await runWorkPaperVsUniverBenchmarkSuite({ sampleCount: 1, warmupCount: 0 })

    expect(results).toHaveLength(WORKPAPER_UNIVER_WORKLOADS.length)
    for (const result of results) {
      const expectedVerification = expectedVerificationByWorkload[result.workload]
      expect(result.category).toBe('workbook-wide')
      expect(result.comparable).toBe(true)
      expect(result.comparison.verificationEquivalent).toBe(true)
      expect(result.engines.workpaper.verification).toEqual(expectedVerification)
      expect(result.engines.univer.verification).toEqual(expectedVerification)
      expect(result.fixture.rowCount).toBeGreaterThanOrEqual(1_000)
      expect(Number.isFinite(result.comparison.workpaperToUniverMeanRatio)).toBe(true)
      expect(Number.isFinite(result.comparison.workpaperToUniverP95Ratio)).toBe(true)
    }
  })

  it('derives scorecard totals from measured Univer results', async () => {
    const results = await runWorkPaperVsUniverBenchmarkSuite({ sampleCount: 1, warmupCount: 0 })
    const report = buildWorkPaperVsUniverBenchmarkReport(results)

    expect(report.suite).toBe('workpaper-vs-univer')
    expect(report.scorecard.coverageTier).toBe('workbook-wide')
    expect(report.scorecard.comparableWorkloadCount).toBe(WORKPAPER_UNIVER_WORKLOADS.length)
    expect(report.scorecard.meanWinCount + report.scorecard.univerMeanWinCount).toBe(WORKPAPER_UNIVER_WORKLOADS.length)
    expect(report.scorecard.p95WinCount + report.scorecard.univerP95WinCount).toBe(WORKPAPER_UNIVER_WORKLOADS.length)
    expect(report.scorecard.meanAndP95WinCount).toBeLessThanOrEqual(WORKPAPER_UNIVER_WORKLOADS.length)
    expect(report.scorecard.workloadFamilies).toEqual([
      'aggregate',
      'aggregate-2d',
      'overlapping-aggregate',
      'formula-chain',
      'formula-fanout',
      'lookup-exact',
      'lookup-approximate',
      'range-stats',
    ])
  })
})
