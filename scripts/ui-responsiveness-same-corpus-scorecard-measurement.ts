import { summarizeNumbers } from '../packages/benchmarks/src/stats.js'
import type {
  SameCorpusBiligRuntimeProof,
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusMeasurement,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { cloneSameCorpusVerification } from './ui-responsiveness-same-corpus-validation-helpers.ts'

export function buildSameCorpusMeasurement(capture: SameCorpusCaptureMeasurement): UiResponsivenessSameCorpusMeasurement {
  return {
    product: capture.product,
    source: capture.source,
    operationResponseMs: summarizeNumbers(capture.operationResponseMsSamples),
    operationResponseProofs: [...capture.operationResponseProofs],
    ...(capture.authoritativeRenderProofMsSamples
      ? { authoritativeRenderProofMs: summarizeNumbers(capture.authoritativeRenderProofMsSamples) }
      : {}),
    ...(capture.committedTargetProofMsSamples
      ? {
          committedTargetProofMs: summarizeNumbers(capture.committedTargetProofMsSamples),
          committedTargetProofMsSamples: [...capture.committedTargetProofMsSamples],
        }
      : {}),
    ...(capture.visibleTargetRenderMsSamples
      ? {
          visibleTargetRenderMs: summarizeNumbers(capture.visibleTargetRenderMsSamples),
          visibleTargetRenderMsSamples: [...capture.visibleTargetRenderMsSamples],
        }
      : {}),
    ...(capture.committedStateValidationMsSamples
      ? {
          committedStateValidationMs: summarizeNumbers(capture.committedStateValidationMsSamples),
          committedStateValidationMsSamples: [...capture.committedStateValidationMsSamples],
        }
      : {}),
    ...(capture.restoreValidationMsSamples
      ? {
          restoreValidationMs: summarizeNumbers(capture.restoreValidationMsSamples),
          restoreValidationMsSamples: [...capture.restoreValidationMsSamples],
        }
      : {}),
    postOperationFrameMs: summarizeNumbers(capture.postOperationFrameMsSamples),
    ...(capture.scrollEventResponseMsSamples ? { scrollEventResponseMs: summarizeNumbers(capture.scrollEventResponseMsSamples) } : {}),
    ...(capture.scrollMovementPxSamples ? { scrollMovementPx: summarizeNumbers(capture.scrollMovementPxSamples) } : {}),
    ...(capture.biligRuntimeProof ? { biligRuntimeProof: cloneBiligRuntimeProof(capture.biligRuntimeProof) } : {}),
    corpusVerification: cloneSameCorpusVerification(capture.corpusVerification),
    limitations: [...capture.limitations],
  }
}

export function sameCorpusScrollEventMetrics(
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): SameCorpusTimingMetrics | null {
  if (
    !bilig.scrollEventResponseMs ||
    !googleSheets.scrollEventResponseMs ||
    (microsoftExcelWeb && !microsoftExcelWeb.scrollEventResponseMs)
  ) {
    return null
  }
  return {
    biligToGoogleSheetsMeanRatio: sameCorpusTimingRatio(bilig.scrollEventResponseMs.mean, googleSheets.scrollEventResponseMs.mean),
    biligToGoogleSheetsP95Ratio: sameCorpusTimingRatio(bilig.scrollEventResponseMs.p95, googleSheets.scrollEventResponseMs.p95),
    biligToMicrosoftExcelWebMeanRatio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.scrollEventResponseMs.mean, microsoftExcelWeb.scrollEventResponseMs!.mean)
      : Number.POSITIVE_INFINITY,
    biligToMicrosoftExcelWebP95Ratio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.scrollEventResponseMs.p95, microsoftExcelWeb.scrollEventResponseMs!.p95)
      : Number.POSITIVE_INFINITY,
  }
}

export function sameCorpusVisibleTargetRenderMetrics(
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): SameCorpusTimingMetrics | null {
  if (
    !bilig.visibleTargetRenderMs ||
    !googleSheets.visibleTargetRenderMs ||
    (microsoftExcelWeb && !microsoftExcelWeb.visibleTargetRenderMs)
  ) {
    return null
  }
  return {
    biligToGoogleSheetsMeanRatio: sameCorpusTimingRatio(bilig.visibleTargetRenderMs.mean, googleSheets.visibleTargetRenderMs.mean),
    biligToGoogleSheetsP95Ratio: sameCorpusTimingRatio(bilig.visibleTargetRenderMs.p95, googleSheets.visibleTargetRenderMs.p95),
    biligToMicrosoftExcelWebMeanRatio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.visibleTargetRenderMs.mean, microsoftExcelWeb.visibleTargetRenderMs!.mean)
      : Number.POSITIVE_INFINITY,
    biligToMicrosoftExcelWebP95Ratio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.visibleTargetRenderMs.p95, microsoftExcelWeb.visibleTargetRenderMs!.p95)
      : Number.POSITIVE_INFINITY,
  }
}

export function sameCorpusTimingRatio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}

interface SameCorpusTimingMetrics {
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio: number
  readonly biligToMicrosoftExcelWebP95Ratio: number
}

function cloneBiligRuntimeProof(proof: SameCorpusBiligRuntimeProof): SameCorpusBiligRuntimeProof {
  return {
    ...proof,
    samples: proof.samples.map((sample) => ({ ...sample })),
  }
}
