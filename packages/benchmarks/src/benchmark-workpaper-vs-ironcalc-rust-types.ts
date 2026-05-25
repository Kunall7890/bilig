import type { ComparativeMeasuredEngineResult } from './benchmark-workpaper-vs-hyperformula.js'
import type { WorkPaperUniverFixture, WorkPaperUniverWorkload, WorkPaperUniverWorkloadFamily } from './benchmark-workpaper-vs-univer.js'
import type { NumericSummary } from './stats.js'

export type WorkPaperIronCalcRustWorkload = WorkPaperUniverWorkload
export type WorkPaperIronCalcRustWorkloadFamily = WorkPaperUniverWorkloadFamily
export type IronCalcRustApiPath = 'Model' | 'UserModel'
export type IronCalcRustEditableValue = boolean | number | string | null

export interface WorkPaperIronCalcRustFixture {
  readonly edit?: WorkPaperUniverFixture['edit']
  readonly family: WorkPaperIronCalcRustWorkloadFamily
  readonly formula: string
  readonly result: NonNullable<WorkPaperUniverFixture['result']>
  readonly rowCount: number
}

export interface IronCalcRustMeasuredEngineResult {
  readonly status: 'supported'
  readonly apiPath: IronCalcRustApiPath
  readonly elapsedMs: NumericSummary
  readonly verification: Record<string, unknown>
}

export interface WorkPaperIronCalcRustComparison {
  readonly confidenceIntervalOverlaps: boolean
  readonly fasterEngine: 'workpaper' | 'ironcalc-rust'
  readonly maxRelativeNoise: number
  readonly meanSpeedup: number
  readonly verificationEquivalent: true
  readonly workpaperToIronCalcRustMeanRatio: number
  readonly workpaperToIronCalcRustMedianRatio: number
  readonly workpaperToIronCalcRustP95Ratio: number
}

export interface WorkPaperIronCalcRustBenchmarkResult {
  readonly workload: WorkPaperIronCalcRustWorkload
  readonly category: 'workbook-wide-limited'
  readonly comparable: true
  readonly fixture: WorkPaperIronCalcRustFixture
  readonly comparison: WorkPaperIronCalcRustComparison
  readonly engines: {
    readonly ironCalcRust: IronCalcRustMeasuredEngineResult
    readonly workpaper: ComparativeMeasuredEngineResult
  }
}

export interface WorkPaperIronCalcRustScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: 'workbook-wide-limited'
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly ironCalcRustMeanWinCount: number
  readonly ironCalcRustP95WinCount: number
  readonly unsupportedWorkloads: readonly WorkPaperIronCalcRustUnsupportedWorkload[]
  readonly workloadFamilies: readonly WorkPaperIronCalcRustWorkloadFamily[]
  readonly worstMeanRatioWorkload: WorkPaperIronCalcRustWorkload
  readonly worstP95RatioWorkload: WorkPaperIronCalcRustWorkload
  readonly worstWorkpaperToIronCalcRustMeanRatio: number
  readonly worstWorkpaperToIronCalcRustP95Ratio: number
}

export interface WorkPaperIronCalcRustUnsupportedWorkload {
  readonly evidence: readonly string[]
  readonly reason: string
  readonly workload: string
}

export interface WorkPaperIronCalcRustBenchmarkReport {
  readonly suite: 'workpaper-vs-ironcalc-rust'
  readonly scorecard: WorkPaperIronCalcRustScorecard
  readonly results: readonly WorkPaperIronCalcRustBenchmarkResult[]
}
