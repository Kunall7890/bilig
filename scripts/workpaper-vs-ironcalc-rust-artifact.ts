import type { ExtraHeadlessComparisonEngineSummary, HeadlessComparisonCoverageTier } from './headless-comparison-engine-summary.ts'
import { arrayField, asObject, literalField, numberField, objectField, stringArrayField, stringField } from './json-scorecard-helpers.ts'
import type {
  IronCalcRustApiPath,
  WorkPaperIronCalcRustUnsupportedWorkload,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-ironcalc-rust-types.js'

export interface ParsedWorkPaperIronCalcRustArtifact {
  readonly generatedAt: string
  readonly engines: {
    readonly ironCalcRust: {
      readonly version: string
    }
  }
  readonly results: readonly ParsedWorkPaperIronCalcRustResult[]
  readonly scorecard: ParsedWorkPaperIronCalcRustScorecard
  readonly suite: 'workpaper-vs-ironcalc-rust'
}

export interface ParsedWorkPaperIronCalcRustResult {
  readonly category: 'workbook-wide-limited'
  readonly comparable: true
  readonly comparison: {
    readonly workpaperToIronCalcRustMeanRatio: number
    readonly workpaperToIronCalcRustP95Ratio: number
  }
  readonly engines: {
    readonly ironCalcRust: {
      readonly apiPath: IronCalcRustApiPath
    }
  }
  readonly fixture: {
    readonly family: string
    readonly rowCount: number
  }
  readonly workload: string
}

export interface ParsedWorkPaperIronCalcRustScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: HeadlessComparisonCoverageTier
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly ironCalcRustMeanWinCount: number
  readonly ironCalcRustP95WinCount: number
  readonly unsupportedWorkloads: readonly WorkPaperIronCalcRustUnsupportedWorkload[]
  readonly workloadFamilies: readonly string[]
  readonly worstMeanRatioWorkload: string
  readonly worstP95RatioWorkload: string
  readonly worstWorkpaperToIronCalcRustMeanRatio: number
  readonly worstWorkpaperToIronCalcRustP95Ratio: number
}

export function parseWorkPaperIronCalcRustArtifact(value: Record<string, unknown>): ParsedWorkPaperIronCalcRustArtifact {
  const engines = objectField(value, 'engines')
  const ironCalcRust = objectField(engines, 'ironCalcRust')
  return {
    generatedAt: stringField(value, 'generatedAt'),
    engines: {
      ironCalcRust: {
        version: stringField(ironCalcRust, 'version'),
      },
    },
    results: arrayField(value, 'results').map(parseWorkPaperIronCalcRustResult),
    scorecard: parseWorkPaperIronCalcRustScorecard(objectField(value, 'scorecard')),
    suite: literalField(value, 'suite', 'workpaper-vs-ironcalc-rust'),
  }
}

export function parseWorkPaperIronCalcRustExtraComparisonEngineSummary(
  value: Record<string, unknown>,
  artifactPath: string,
): ExtraHeadlessComparisonEngineSummary {
  const artifact = parseWorkPaperIronCalcRustArtifact(value)
  return {
    artifactPath,
    comparableWorkloadCount: artifact.scorecard.comparableWorkloadCount,
    coverageNote: artifact.scorecard.coverageNote,
    coverageTier: artifact.scorecard.coverageTier,
    engineName: 'IronCalc Rust',
    generatedAt: artifact.generatedAt,
    meanAndP95WinCount: artifact.scorecard.meanAndP95WinCount,
    meanWinCount: artifact.scorecard.meanWinCount,
    p95WinCount: artifact.scorecard.p95WinCount,
    version: artifact.engines.ironCalcRust.version,
    workloadFamilies: artifact.scorecard.workloadFamilies,
    workloads: artifact.results.map((result) => result.workload),
  }
}

export function deriveWorkPaperIronCalcRustScorecard(
  results: readonly ParsedWorkPaperIronCalcRustResult[],
  coverageNote: string,
  unsupportedWorkloads: readonly WorkPaperIronCalcRustUnsupportedWorkload[],
): ParsedWorkPaperIronCalcRustScorecard {
  if (results.length === 0) {
    throw new Error('Cannot derive a WorkPaper vs IronCalc Rust scorecard without results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToIronCalcRustMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToIronCalcRustP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToIronCalcRustMeanRatio < 1 && result.comparison.workpaperToIronCalcRustP95Ratio < 1,
  ).length
  return {
    comparableWorkloadCount: results.length,
    coverageNote,
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

function parseWorkPaperIronCalcRustScorecard(value: Record<string, unknown>): ParsedWorkPaperIronCalcRustScorecard {
  return {
    comparableWorkloadCount: numberField(value, 'comparableWorkloadCount'),
    coverageNote: stringField(value, 'coverageNote'),
    coverageTier: literalField(value, 'coverageTier', 'workbook-wide-limited'),
    directionalMeanRatioGeomean: numberField(value, 'directionalMeanRatioGeomean'),
    directionalP95RatioGeomean: numberField(value, 'directionalP95RatioGeomean'),
    ironCalcRustMeanWinCount: numberField(value, 'ironCalcRustMeanWinCount'),
    ironCalcRustP95WinCount: numberField(value, 'ironCalcRustP95WinCount'),
    meanAndP95WinCount: numberField(value, 'meanAndP95WinCount'),
    meanWinCount: numberField(value, 'meanWinCount'),
    p95WinCount: numberField(value, 'p95WinCount'),
    unsupportedWorkloads: arrayField(value, 'unsupportedWorkloads').map(parseUnsupportedWorkload),
    workloadFamilies: stringArrayField(value, 'workloadFamilies'),
    worstMeanRatioWorkload: stringField(value, 'worstMeanRatioWorkload'),
    worstP95RatioWorkload: stringField(value, 'worstP95RatioWorkload'),
    worstWorkpaperToIronCalcRustMeanRatio: numberField(value, 'worstWorkpaperToIronCalcRustMeanRatio'),
    worstWorkpaperToIronCalcRustP95Ratio: numberField(value, 'worstWorkpaperToIronCalcRustP95Ratio'),
  }
}

function parseWorkPaperIronCalcRustResult(value: unknown): ParsedWorkPaperIronCalcRustResult {
  const result = asObject(value, 'WorkPaper IronCalc Rust result')
  const comparison = objectField(result, 'comparison')
  const fixture = objectField(result, 'fixture')
  return {
    category: literalField(result, 'category', 'workbook-wide-limited'),
    comparable: literalField(result, 'comparable', true),
    comparison: {
      workpaperToIronCalcRustMeanRatio: numberField(comparison, 'workpaperToIronCalcRustMeanRatio'),
      workpaperToIronCalcRustP95Ratio: numberField(comparison, 'workpaperToIronCalcRustP95Ratio'),
    },
    engines: {
      ironCalcRust: {
        apiPath: parseIronCalcRustApiPath(objectField(objectField(result, 'engines'), 'ironCalcRust')),
      },
    },
    fixture: {
      family: stringField(fixture, 'family'),
      rowCount: numberField(fixture, 'rowCount'),
    },
    workload: stringField(result, 'workload'),
  }
}

function parseUnsupportedWorkload(value: unknown): WorkPaperIronCalcRustUnsupportedWorkload {
  const unsupported = asObject(value, 'unsupported IronCalc Rust workload')
  return {
    evidence: stringArrayField(unsupported, 'evidence'),
    reason: stringField(unsupported, 'reason'),
    workload: stringField(unsupported, 'workload'),
  }
}

function parseIronCalcRustApiPath(value: Record<string, unknown>): IronCalcRustApiPath {
  const apiPath = stringField(value, 'apiPath')
  if (apiPath !== 'Model' && apiPath !== 'UserModel') {
    throw new Error(`Expected IronCalc Rust apiPath to be Model or UserModel, got ${apiPath}`)
  }
  return apiPath
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
  results: readonly ParsedWorkPaperIronCalcRustResult[],
  ratioKey: 'workpaperToIronCalcRustMeanRatio' | 'workpaperToIronCalcRustP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly ParsedWorkPaperIronCalcRustResult[],
  ratioKey: 'workpaperToIronCalcRustMeanRatio' | 'workpaperToIronCalcRustP95Ratio',
): string {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}

function orderedUnique(values: readonly string[]): string[] {
  return [...new Set(values)]
}
