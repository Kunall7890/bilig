import type { SameCorpusCaptureMeasurement, UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { validateSameCorpusMutationTargetTimingSamplesMatchArrays } from './ui-responsiveness-same-corpus-mutation-timing-samples.ts'
import {
  uiSameCorpusWorkloadMutatesWorkbook,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

export function assertSameCorpusProductMeasurement(
  product: UiResponsivenessSameCorpusProduct,
  source: string,
  measurement: SameCorpusCaptureMeasurement,
  workload: UiResponsivenessSameCorpusWorkload,
  options: { readonly allowIncompleteEvidence?: boolean } = {},
): void {
  if (measurement.product !== product) {
    throw new Error(`same-corpus UI measurement expected ${product} but received ${measurement.product}`)
  }
  if (measurement.source !== source) {
    throw new Error(`same-corpus UI measurement for ${product} used an unexpected source URL`)
  }
  assertSameCorpusSampleArray(product, 'operation response', measurement.operationResponseMsSamples)
  if (product === 'bilig') {
    assertSameCorpusSampleArray(
      product,
      'authoritative render proof',
      measurement.authoritativeRenderProofMsSamples,
      measurement.operationResponseMsSamples.length,
    )
  }
  assertSameCorpusSampleArray(
    product,
    'post-operation frame',
    measurement.postOperationFrameMsSamples,
    measurement.operationResponseMsSamples.length,
  )
  if (uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
    assertSameCorpusSampleArray(
      product,
      'scroll-event response',
      measurement.scrollEventResponseMsSamples,
      measurement.operationResponseMsSamples.length,
    )
    assertSameCorpusSampleArray(
      product,
      'scroll movement',
      measurement.scrollMovementPxSamples,
      measurement.operationResponseMsSamples.length,
    )
  }
  if (uiSameCorpusWorkloadMutatesWorkbook(workload)) {
    if (options.allowIncompleteEvidence === true && measurement.committedTargetProofMsSamples === undefined) {
      return
    }
    assertSameCorpusSampleArray(
      product,
      'committed target proof',
      measurement.committedTargetProofMsSamples,
      measurement.operationResponseMsSamples.length,
    )
    assertSameCorpusSampleArray(
      product,
      'visible target render',
      measurement.visibleTargetRenderMsSamples,
      measurement.operationResponseMsSamples.length,
    )
    assertSameCorpusSampleArray(
      product,
      'committed-state validation',
      measurement.committedStateValidationMsSamples,
      measurement.operationResponseMsSamples.length,
    )
    assertSameCorpusSampleArray(
      product,
      'restore validation',
      measurement.restoreValidationMsSamples,
      measurement.operationResponseMsSamples.length,
    )
    validateSameCorpusMutationTargetTimingSamplesMatchArrays(measurement.committedTargetProofTimingSamples, {
      committedStateValidationMsSamples: measurement.committedStateValidationMsSamples,
      committedTargetProofMsSamples: measurement.committedTargetProofMsSamples,
      expectedLength: measurement.operationResponseMsSamples.length,
      expectedProduct: product,
      label: `same-corpus UI measurement for ${product}`,
      restoreValidationMsSamples: measurement.restoreValidationMsSamples,
      visibleTargetRenderMsSamples: measurement.visibleTargetRenderMsSamples,
    })
  }
}

function assertSameCorpusSampleArray(
  product: UiResponsivenessSameCorpusProduct,
  label: string,
  samples: readonly number[] | undefined,
  expectedLength?: number,
): void {
  if (!samples || samples.length === 0) {
    throw new Error(`same-corpus UI measurement for ${product} is missing ${label} samples`)
  }
  if (expectedLength !== undefined && samples.length !== expectedLength) {
    throw new Error(
      `same-corpus UI measurement for ${product} has ${String(samples.length)} ${label} samples but expected ${String(expectedLength)}`,
    )
  }
  for (const sample of samples) {
    if (!Number.isFinite(sample) || sample < 0) {
      throw new Error(`same-corpus UI measurement for ${product} has an invalid ${label} sample`)
    }
  }
}
