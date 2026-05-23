import type { ExtraHeadlessComparisonEngineSummary, HeadlessComparisonCoverageTier } from './headless-comparison-engine-summary.ts'
import { arrayField, asObject, literalField, numberField, objectField, stringArrayField, stringField } from './json-scorecard-helpers.ts'

export interface ParsedWorkPaperUniverArtifact {
  readonly generatedAt: string
  readonly engines: {
    readonly univer: {
      readonly version: string
    }
  }
  readonly results: readonly ParsedWorkPaperUniverResult[]
  readonly scorecard: ParsedWorkPaperUniverScorecard
  readonly suite: 'workpaper-vs-univer'
}

export interface ParsedWorkPaperUniverResult {
  readonly category: 'workbook-wide'
  readonly comparable: true
  readonly comparison: {
    readonly workpaperToUniverMeanRatio: number
    readonly workpaperToUniverP95Ratio: number
  }
  readonly fixture: {
    readonly family: string
    readonly rowCount: number
  }
  readonly workload: string
}

export interface ParsedWorkPaperUniverScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: HeadlessComparisonCoverageTier
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly univerMeanWinCount: number
  readonly univerP95WinCount: number
  readonly workloadFamilies: readonly string[]
  readonly worstMeanRatioWorkload: string
  readonly worstP95RatioWorkload: string
  readonly worstWorkpaperToUniverMeanRatio: number
  readonly worstWorkpaperToUniverP95Ratio: number
}

export function parseWorkPaperUniverArtifact(value: Record<string, unknown>): ParsedWorkPaperUniverArtifact {
  const engines = objectField(value, 'engines')
  const univer = objectField(engines, 'univer')
  return {
    generatedAt: stringField(value, 'generatedAt'),
    engines: {
      univer: {
        version: stringField(univer, 'version'),
      },
    },
    results: arrayField(value, 'results').map(parseWorkPaperUniverResult),
    scorecard: parseWorkPaperUniverScorecard(objectField(value, 'scorecard')),
    suite: literalField(value, 'suite', 'workpaper-vs-univer'),
  }
}

export function parseWorkPaperUniverExtraComparisonEngineSummary(
  value: Record<string, unknown>,
  artifactPath: string,
): ExtraHeadlessComparisonEngineSummary {
  const artifact = parseWorkPaperUniverArtifact(value)
  return {
    artifactPath,
    comparableWorkloadCount: artifact.scorecard.comparableWorkloadCount,
    coverageNote: artifact.scorecard.coverageNote,
    coverageTier: artifact.scorecard.coverageTier,
    engineName: 'Univer',
    generatedAt: artifact.generatedAt,
    meanAndP95WinCount: artifact.scorecard.meanAndP95WinCount,
    meanWinCount: artifact.scorecard.meanWinCount,
    p95WinCount: artifact.scorecard.p95WinCount,
    version: artifact.engines.univer.version,
    workloadFamilies: artifact.scorecard.workloadFamilies,
  }
}

export function deriveWorkPaperUniverScorecard(
  results: readonly ParsedWorkPaperUniverResult[],
  coverageNote: string,
): ParsedWorkPaperUniverScorecard {
  if (results.length === 0) {
    throw new Error('Cannot derive a WorkPaper vs Univer scorecard without results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToUniverMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToUniverP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToUniverMeanRatio < 1 && result.comparison.workpaperToUniverP95Ratio < 1,
  ).length
  return {
    comparableWorkloadCount: results.length,
    coverageNote,
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
  }
}

function parseWorkPaperUniverScorecard(value: Record<string, unknown>): ParsedWorkPaperUniverScorecard {
  return {
    comparableWorkloadCount: numberField(value, 'comparableWorkloadCount'),
    coverageNote: stringField(value, 'coverageNote'),
    coverageTier: literalField(value, 'coverageTier', 'workbook-wide'),
    directionalMeanRatioGeomean: numberField(value, 'directionalMeanRatioGeomean'),
    directionalP95RatioGeomean: numberField(value, 'directionalP95RatioGeomean'),
    meanAndP95WinCount: numberField(value, 'meanAndP95WinCount'),
    meanWinCount: numberField(value, 'meanWinCount'),
    p95WinCount: numberField(value, 'p95WinCount'),
    univerMeanWinCount: numberField(value, 'univerMeanWinCount'),
    univerP95WinCount: numberField(value, 'univerP95WinCount'),
    workloadFamilies: stringArrayField(value, 'workloadFamilies'),
    worstMeanRatioWorkload: stringField(value, 'worstMeanRatioWorkload'),
    worstP95RatioWorkload: stringField(value, 'worstP95RatioWorkload'),
    worstWorkpaperToUniverMeanRatio: numberField(value, 'worstWorkpaperToUniverMeanRatio'),
    worstWorkpaperToUniverP95Ratio: numberField(value, 'worstWorkpaperToUniverP95Ratio'),
  }
}

function parseWorkPaperUniverResult(value: unknown): ParsedWorkPaperUniverResult {
  const result = asObject(value, 'WorkPaper Univer result')
  const comparison = objectField(result, 'comparison')
  const fixture = objectField(result, 'fixture')
  return {
    category: literalField(result, 'category', 'workbook-wide'),
    comparable: literalField(result, 'comparable', true),
    comparison: {
      workpaperToUniverMeanRatio: numberField(comparison, 'workpaperToUniverMeanRatio'),
      workpaperToUniverP95Ratio: numberField(comparison, 'workpaperToUniverP95Ratio'),
    },
    fixture: {
      family: stringField(fixture, 'family'),
      rowCount: numberField(fixture, 'rowCount'),
    },
    workload: stringField(result, 'workload'),
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
  results: readonly ParsedWorkPaperUniverResult[],
  ratioKey: 'workpaperToUniverMeanRatio' | 'workpaperToUniverP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly ParsedWorkPaperUniverResult[],
  ratioKey: 'workpaperToUniverMeanRatio' | 'workpaperToUniverP95Ratio',
): string {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}

function orderedUnique(values: readonly string[]): string[] {
  return [...new Set(values)]
}
