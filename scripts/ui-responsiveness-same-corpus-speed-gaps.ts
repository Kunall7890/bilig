import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

const requiredGoogleSheetsSpeedup = 10
export type SameCorpusUiSpeedMetric = 'operationResponseMs' | 'scrollEventResponseMs'

export interface SameCorpusUiSpeedSummary {
  readonly mean: number
  readonly p95: number
}

export interface SameCorpusUiSpeedMeasurement {
  readonly operationResponseMs: SameCorpusUiSpeedSummary
  readonly scrollEventResponseMs?: SameCorpusUiSpeedSummary
}

export interface SameCorpusUiSpeedGapCase {
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly tenXMeanAndP95Metric?: SameCorpusUiSpeedMetric
  readonly bilig: SameCorpusUiSpeedMeasurement
  readonly googleSheets: SameCorpusUiSpeedMeasurement
}

export interface SameCorpusUiSpeedGap {
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly metric: SameCorpusUiSpeedMetric
  readonly biligMeanMs: number
  readonly biligP95Ms: number
  readonly googleMeanMs: number
  readonly googleP95Ms: number
  /** Google Sheets time divided by Bilig time. Values >= 10 mean Bilig is at least 10x faster. */
  readonly meanRatio: number
  /** Google Sheets p95 divided by Bilig p95. Values >= 10 mean Bilig is at least 10x faster at p95. */
  readonly p95Ratio: number
  readonly requiredMeanRatio: typeof requiredGoogleSheetsSpeedup
  readonly requiredP95Ratio: typeof requiredGoogleSheetsSpeedup
  readonly meanAdditionalSpeedupTo10x: number
  readonly p95AdditionalSpeedupTo10x: number
  readonly limitingAdditionalSpeedupTo10x: number
}

export function sameCorpusUiSpeedGaps(proof: { readonly cases: readonly SameCorpusUiSpeedGapCase[] }): readonly SameCorpusUiSpeedGap[] {
  return proof.cases
    .map(sameCorpusCaseSpeedGap)
    .filter((entry) => entry.meanRatio < requiredGoogleSheetsSpeedup || entry.p95Ratio < requiredGoogleSheetsSpeedup)
    .toSorted((left, right) => right.limitingAdditionalSpeedupTo10x - left.limitingAdditionalSpeedupTo10x)
}

export function formatSameCorpusUiSpeedGap(gap: SameCorpusUiSpeedGap): string {
  return [
    `${gap.workload} (${gap.metric})`,
    `current Google/Bilig mean ${formatSpeedup(gap.meanRatio)}x`,
    `p95 ${formatSpeedup(gap.p95Ratio)}x`,
    `needs ${formatSpeedup(gap.limitingAdditionalSpeedupTo10x)}x additional Bilig speedup to satisfy 10x mean+p95`,
  ].join(': ')
}

function sameCorpusCaseSpeedGap(entry: SameCorpusUiSpeedGapCase): SameCorpusUiSpeedGap {
  const metric = entry.tenXMeanAndP95Metric ?? 'operationResponseMs'
  const biligMeanMs = metric === 'scrollEventResponseMs' ? entry.bilig.scrollEventResponseMs?.mean : entry.bilig.operationResponseMs.mean
  const biligP95Ms = metric === 'scrollEventResponseMs' ? entry.bilig.scrollEventResponseMs?.p95 : entry.bilig.operationResponseMs.p95
  const googleMeanMs =
    metric === 'scrollEventResponseMs' ? entry.googleSheets.scrollEventResponseMs?.mean : entry.googleSheets.operationResponseMs.mean
  const googleP95Ms =
    metric === 'scrollEventResponseMs' ? entry.googleSheets.scrollEventResponseMs?.p95 : entry.googleSheets.operationResponseMs.p95
  const meanRatio = speedupRatio(googleMeanMs, biligMeanMs)
  const p95Ratio = speedupRatio(googleP95Ms, biligP95Ms)
  const meanAdditionalSpeedupTo10x = additionalSpeedupTo10x(meanRatio)
  const p95AdditionalSpeedupTo10x = additionalSpeedupTo10x(p95Ratio)
  return {
    workload: entry.workload,
    metric,
    biligMeanMs: biligMeanMs ?? Number.POSITIVE_INFINITY,
    biligP95Ms: biligP95Ms ?? Number.POSITIVE_INFINITY,
    googleMeanMs: googleMeanMs ?? 0,
    googleP95Ms: googleP95Ms ?? 0,
    meanRatio,
    p95Ratio,
    requiredMeanRatio: requiredGoogleSheetsSpeedup,
    requiredP95Ratio: requiredGoogleSheetsSpeedup,
    meanAdditionalSpeedupTo10x,
    p95AdditionalSpeedupTo10x,
    limitingAdditionalSpeedupTo10x: Math.max(meanAdditionalSpeedupTo10x, p95AdditionalSpeedupTo10x),
  }
}

function speedupRatio(googleMs: number | undefined, biligMs: number | undefined): number {
  if (googleMs === undefined || biligMs === undefined || googleMs <= 0 || biligMs <= 0) {
    return 0
  }
  return googleMs / biligMs
}

function additionalSpeedupTo10x(currentSpeedupRatio: number): number {
  if (currentSpeedupRatio >= requiredGoogleSheetsSpeedup) {
    return 1
  }
  if (currentSpeedupRatio <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return requiredGoogleSheetsSpeedup / currentSpeedupRatio
}

function formatSpeedup(value: number): string {
  if (!Number.isFinite(value)) {
    return 'infinite'
  }
  if (value >= 100) {
    return value.toFixed(0)
  }
  if (value >= 10) {
    return value.toFixed(1)
  }
  return value.toFixed(2)
}
