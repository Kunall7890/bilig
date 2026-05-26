import type {
  SameCorpusCaptureCase,
  SameCorpusCaptureMeasurement,
  SameCorpusOperationResponseProof,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'
import {
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

export const sameCorpusSampleCount = 3

export function sameCorpusCaseOperationResponseProofGuardrailPassed(entry: UiResponsivenessSameCorpusCase): boolean {
  return sameCorpusMeasurementOperationResponseProofsPassed(entry.workload, entry.sampleCount, [
    entry.bilig,
    entry.googleSheets,
    ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : []),
  ])
}

export function sameCorpusCaptureCaseOperationResponseProofGuardrailPassed(entry: SameCorpusCaptureCase, sampleCount: number): boolean {
  return sameCorpusMeasurementOperationResponseProofsPassed(entry.workload, sampleCount, [
    entry.bilig,
    entry.googleSheets,
    ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : []),
  ])
}

export function sameCorpusMeasurementOperationResponseProofsPassed(
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
  measurements: readonly Pick<SameCorpusCaptureMeasurement | UiResponsivenessSameCorpusMeasurement, 'operationResponseProofs'>[],
): boolean {
  const expected = expectedSameCorpusOperationResponseProof(workload)
  return measurements.every(
    (measurement) =>
      measurement.operationResponseProofs.length >= sampleCount && measurement.operationResponseProofs.every((proof) => proof === expected),
  )
}

function expectedSameCorpusOperationResponseProof(workload: UiResponsivenessSameCorpusWorkload): SameCorpusOperationResponseProof {
  if (workload === 'open-workbook') {
    return 'load-to-ready'
  }
  if (uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
    return 'visible-scroll-movement'
  }
  return 'visible-non-scroll-response'
}

export function hasBiligAuthoritativeRenderProofTiming(
  measurement: Pick<
    SameCorpusCaptureMeasurement | UiResponsivenessSameCorpusMeasurement,
    'product' | 'authoritativeRenderProofMsSamples' | 'authoritativeRenderProofMs'
  >,
  sampleCount: number,
): boolean {
  if (measurement.product !== 'bilig') {
    return false
  }
  if ('authoritativeRenderProofMs' in measurement && measurement.authoritativeRenderProofMs) {
    return (
      measurement.authoritativeRenderProofMs.samples.length >= sampleCount &&
      measurement.authoritativeRenderProofMs.samples.every((value) => Number.isFinite(value) && value >= 0)
    )
  }
  const samples = 'authoritativeRenderProofMsSamples' in measurement ? measurement.authoritativeRenderProofMsSamples : undefined
  return Boolean(samples && samples.length >= sampleCount && samples.every((value) => Number.isFinite(value) && value >= 0))
}

export function hasCommittedTargetProofTiming(
  measurement: UiResponsivenessSameCorpusMeasurement | SameCorpusCaptureMeasurement,
  sampleCount: number,
): boolean {
  if ('committedTargetProofMs' in measurement && measurement.committedTargetProofMs) {
    return (
      measurement.committedTargetProofMs.samples.length >= sampleCount &&
      measurement.committedTargetProofMs.samples.every((value) => Number.isFinite(value) && value >= 0)
    )
  }
  const samples = 'committedTargetProofMsSamples' in measurement ? measurement.committedTargetProofMsSamples : undefined
  return Boolean(samples && samples.length >= sampleCount && samples.every((value) => Number.isFinite(value) && value >= 0))
}

export function sameCorpusCommittedTargetProofMetrics(
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): {
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio: number
  readonly biligToMicrosoftExcelWebP95Ratio: number
} | null {
  if (
    !bilig.committedTargetProofMs ||
    !googleSheets.committedTargetProofMs ||
    (microsoftExcelWeb && !microsoftExcelWeb.committedTargetProofMs)
  ) {
    return null
  }
  return {
    biligToGoogleSheetsMeanRatio: sameCorpusTimingRatio(bilig.committedTargetProofMs.mean, googleSheets.committedTargetProofMs.mean),
    biligToGoogleSheetsP95Ratio: sameCorpusTimingRatio(bilig.committedTargetProofMs.p95, googleSheets.committedTargetProofMs.p95),
    biligToMicrosoftExcelWebMeanRatio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.committedTargetProofMs.mean, microsoftExcelWeb.committedTargetProofMs!.mean)
      : Number.POSITIVE_INFINITY,
    biligToMicrosoftExcelWebP95Ratio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.committedTargetProofMs.p95, microsoftExcelWeb.committedTargetProofMs!.p95)
      : Number.POSITIVE_INFINITY,
  }
}

export function hasBiligProductionRuntimeProof(
  measurement: Pick<SameCorpusCaptureMeasurement | UiResponsivenessSameCorpusMeasurement, 'product' | 'biligRuntimeProof'>,
): boolean {
  const proof = measurement.biligRuntimeProof
  return (
    measurement.product === 'bilig' &&
    proof !== undefined &&
    proof.product === 'bilig' &&
    proof.requiredBuildKind === 'production' &&
    proof.actualBuildKind === 'production' &&
    proof.prod &&
    !proof.dev &&
    proof.verified &&
    proof.sampleCount >= sameCorpusSampleCount &&
    proof.samples.length >= sameCorpusSampleCount &&
    proof.samples.every((sample) => sample.present && sample.buildKind === 'production' && sample.prod && !sample.dev)
  )
}

function sameCorpusTimingRatio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}
