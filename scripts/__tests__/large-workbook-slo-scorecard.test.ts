import { describe, expect, it } from 'vitest'

import { buildLargeWorkbookSloScorecard } from '../gen-large-workbook-slo-scorecard.ts'

describe('large workbook SLO scorecard', () => {
  it('maps benchmark-contract output into checked large-workbook and worker-runtime SLOs', () => {
    const scorecard = buildLargeWorkbookSloScorecard(buildReportFixture())

    expect(scorecard.summary.coveredLargeWorkbookRows).toEqual([100_000, 250_000])
    expect(scorecard.summary.allSloBudgetsPassed).toBe(true)
    expect(scorecard.summary.headedBrowserFrameP95Evidence).toBe('playwright-contracts')
    expect(scorecard.summary.headedBrowserFrameP95ContractsPassed).toBe(true)
    expect(scorecard.measurements.map((measurement) => measurement.id)).toEqual([
      'load100k',
      'load250k',
      'workerWarmStart100k',
      'workerWarmStart250k',
      'workerVisibleEdit10k',
      'workerReconnectCatchUp100Pending',
    ])
    expect(scorecard.measurements.find((measurement) => measurement.id === 'workerVisibleEdit10k')).toMatchObject({
      category: 'ui-responsiveness',
      actualP95: 4,
      budgetP95: 16,
      passed: true,
    })
    expect(scorecard.headedBrowserFrameP95Contracts.map((contract) => contract.id)).toEqual([
      'headedDense100kDiagonalBrowse',
      'headedWide250kMainBodyBrowse',
      'headedWide250kVisibleEditCommit',
    ])
    expect(scorecard.headedBrowserFrameP95Contracts.every((contract) => contract.passed)).toBe(true)
    expect(scorecard.headedBrowserFrameP95Contracts[0]).toMatchObject({
      category: 'large-workbook-scale',
      corpusCaseId: 'dense-mixed-100k',
      materializedCells: 100_000,
      metric: 'frameMs.p95',
      budgetP95: 20,
    })
  })

  it('rejects reports that do not cover both 100k and 250k workbook sessions', () => {
    const report = buildReportFixture()
    delete report.results.load250k

    expect(() => buildLargeWorkbookSloScorecard(report)).toThrow('Missing large workbook SLO benchmark result: load250k')
  })
})

function buildReportFixture() {
  return {
    baseBudgets: {
      load100kP95Ms: 1500,
      load250kP95Ms: 1500,
      workerWarmStart100kP95Ms: 500,
      workerWarmStart250kP95Ms: 700,
      workerVisibleEdit10kP95Ms: 16,
      workerReconnectCatchUp100PendingP95Ms: 2000,
    },
    budgets: {
      load100kP95Ms: 1500,
      load250kP95Ms: 1500,
      workerWarmStart100kP95Ms: 500,
      workerWarmStart250kP95Ms: 700,
      workerVisibleEdit10kP95Ms: 16,
      workerReconnectCatchUp100PendingP95Ms: 2000,
    },
    toleranceMultiplier: 1,
    sampleCounts: {
      load100k: 5,
      load250k: 3,
      workerWarmStart100k: 3,
      workerWarmStart250k: 3,
      workerVisibleEdit10k: 5,
      workerReconnectCatchUp100Pending: 3,
    },
    results: {
      load100k: benchmarkResult('dense-mixed-100k', 100_000, 230),
      load250k: benchmarkResult('dense-mixed-250k', 250_000, 600),
      workerWarmStart100k: benchmarkResult('dense-mixed-100k', 100_000, 12),
      workerWarmStart250k: benchmarkResult('dense-mixed-250k', 250_000, 17),
      workerVisibleEdit10k: {
        materializedCells: 10_000,
        visiblePatchMs: numericSummary(4),
      },
      workerReconnectCatchUp100Pending: {
        materializedCells: 10_000,
        pendingMutationCount: 100,
        catchUpMs: numericSummary(270),
      },
    },
    headedBrowserTestSource: `
      function expectSmoothBrowse(report) {
        expect(report.samples.frameMs.length).toBeGreaterThan(120)
        expect(frameSummary.p95).toBeLessThan(20)
      }

      function expectBoundedVisibleMutation(report) {
        expect(report.samples.mutationToVisibleMs.length).toBeGreaterThan(0)
      }

      test('keeps dense 100k browse inside headed frame budgets', async ({ page }, testInfo) => {
        await gotoWorkbookShell(page, '/?benchmarkCorpus=dense-mixed-100k')
        const benchmarkState = await waitForBenchmarkCorpus(page)
        expect(benchmarkState.fixture?.id).toBe('dense-mixed-100k')
        await warmStartWorkbookScrollPerf(page, 'dense-100k-diagonal-main-body')
        const report = await stopWorkbookScrollPerf(page)
        await writeFile(testInfo.outputPath('scroll-perf-dense-100k-diagonal.json'), JSON.stringify(report, null, 2), 'utf8')
        expect(report.fixture?.id).toBe('dense-mixed-100k')
        expectSmoothBrowse(report, { longTaskMax: 60 })
      })

      test('keeps horizontal browse inside one resident window smooth and free of data-canvas redraw churn', async ({ page }, testInfo) => {
        await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
        const benchmarkState = await waitForBenchmarkCorpus(page)
        expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')
        await warmStartWorkbookScrollPerf(page, 'wide-250k-main-body')
        const report = await stopWorkbookScrollPerf(page)
        await writeFile(testInfo.outputPath('scroll-perf-wide-250k-main-body.json'), JSON.stringify(report, null, 2), 'utf8')
        expect(report.fixture?.id).toBe('wide-mixed-250k')
        expectSmoothBrowse(report, { longTaskMax: 60 })
      })

      test('keeps visible edit commits bounded to dirty V3 tiles', async ({ page }, testInfo) => {
        await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
        const benchmarkState = await waitForBenchmarkCorpus(page)
        expect(benchmarkState.fixture?.id).toBe('wide-mixed-250k')
        await warmStartWorkbookScrollPerf(page, 'wide-250k-visible-edit-commit')
        const report = await stopWorkbookScrollPerf(page)
        await writeFile(testInfo.outputPath('scroll-perf-wide-250k-visible-edit.json'), JSON.stringify(report, null, 2), 'utf8')
        expect(report.fixture?.id).toBe('wide-mixed-250k')
        expectBoundedVisibleMutation(report, { mutationToVisibleP95Max: 50 })
      })
    `,
  }
}

function benchmarkResult(corpusCaseId: string, materializedCells: number, p95: number) {
  return {
    corpusCaseId,
    materializedCells,
    elapsedMs: numericSummary(p95),
  }
}

function numericSummary(p95: number) {
  return {
    samples: [p95],
    min: p95,
    median: p95,
    p95,
    max: p95,
    mean: p95,
  }
}
