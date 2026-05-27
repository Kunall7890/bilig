import type {
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { sameCorpusSampleCount } from './ui-responsiveness-same-corpus-guardrails.ts'
import { validateOrderedSameCorpusTimingSamples } from './ui-responsiveness-same-corpus-ordered-timing-samples.ts'
import { validateBiligRuntimeProof } from './ui-responsiveness-same-corpus-bilig-runtime-proof.ts'
import { validateSameCorpusCaptureVerification, validateSummary } from './ui-responsiveness-same-corpus-validation-helpers.ts'
import { uiSameCorpusWorkloadMutatesWorkbook, type UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function validateSameCorpusScorecardMeasurement(
  measurement: UiResponsivenessSameCorpusMeasurement,
  product: UiResponsivenessSameCorpusProduct,
  caseId: string,
  corpusCaseId: string,
  materializedCells: number,
  workload: UiResponsivenessSameCorpusWorkload,
): void {
  if (measurement.product !== product) {
    throw new Error(`UI responsiveness same-corpus product mismatch for ${caseId}`)
  }
  if (measurement.source.length === 0) {
    throw new Error(`UI responsiveness same-corpus source is missing for ${caseId}`)
  }
  if (product === 'bilig' && measurement.biligRuntimeProof) {
    validateBiligRuntimeProof(measurement.biligRuntimeProof, measurement.source, caseId)
  }
  validateSummary(measurement.operationResponseMs, `${caseId} ${product} operationResponseMs`, sameCorpusSampleCount)
  if (measurement.operationResponseProofs.length < sameCorpusSampleCount) {
    throw new Error(`UI responsiveness same-corpus proof is missing operation-response proof for ${caseId}`)
  }
  if (measurement.authoritativeRenderProofMs) {
    validateSummary(measurement.authoritativeRenderProofMs, `${caseId} ${product} authoritativeRenderProofMs`, sameCorpusSampleCount)
  }
  if (measurement.committedTargetProofMs) {
    validateSummary(measurement.committedTargetProofMs, `${caseId} ${product} committedTargetProofMs`, sameCorpusSampleCount)
  }
  if (measurement.visibleTargetRenderMs) {
    validateSummary(measurement.visibleTargetRenderMs, `${caseId} ${product} visibleTargetRenderMs`, sameCorpusSampleCount)
  }
  if (measurement.committedStateValidationMs) {
    validateSummary(measurement.committedStateValidationMs, `${caseId} ${product} committedStateValidationMs`, sameCorpusSampleCount)
  }
  if (measurement.restoreValidationMs) {
    validateSummary(measurement.restoreValidationMs, `${caseId} ${product} restoreValidationMs`, sameCorpusSampleCount)
  }
  if (uiSameCorpusWorkloadMutatesWorkbook(workload)) {
    validateMutationTimingSamples(measurement, caseId, product)
  }
  validateSummary(measurement.postOperationFrameMs, `${caseId} ${product} postOperationFrameMs`, sameCorpusSampleCount)
  if (measurement.scrollEventResponseMs) {
    validateSummary(measurement.scrollEventResponseMs, `${caseId} ${product} scrollEventResponseMs`, sameCorpusSampleCount)
  }
  if (measurement.scrollMovementPx) {
    validateSummary(measurement.scrollMovementPx, `${caseId} ${product} scrollMovementPx`, sameCorpusSampleCount)
  }
  validateSameCorpusCaptureVerification(measurement.corpusVerification, product, materializedCells, corpusCaseId, caseId)
}

function validateMutationTimingSamples(
  measurement: UiResponsivenessSameCorpusMeasurement,
  caseId: string,
  product: UiResponsivenessSameCorpusProduct,
): void {
  validateOrderedSameCorpusTimingSamples(
    measurement.committedTargetProofMsSamples,
    measurement.committedTargetProofMs,
    `${caseId} ${product} committedTargetProofMs`,
  )
  validateOrderedSameCorpusTimingSamples(
    measurement.visibleTargetRenderMsSamples,
    measurement.visibleTargetRenderMs,
    `${caseId} ${product} visibleTargetRenderMs`,
  )
  validateOrderedSameCorpusTimingSamples(
    measurement.committedStateValidationMsSamples,
    measurement.committedStateValidationMs,
    `${caseId} ${product} committedStateValidationMs`,
  )
  validateOrderedSameCorpusTimingSamples(
    measurement.restoreValidationMsSamples,
    measurement.restoreValidationMs,
    `${caseId} ${product} restoreValidationMs`,
  )
}
