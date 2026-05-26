import type { SameCorpusMutationTargetProof } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function sameCorpusMutationTargetTimingInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  const timings = [
    sample.operationStartedAtMs,
    sample.visibleTargetRenderCapturedAtMs,
    sample.postMutationProofCapturedAtMs,
    sample.restoreProofCapturedAtMs,
  ]
  if (timings.some((value) => !Number.isFinite(value) || value < 0)) {
    return [`semantic UI mutation target proof for ${workload} is missing operation proof timing bounds`]
  }
  const invalidReasons: string[] = []
  if (
    sample.visibleTargetRenderCapturedAtMs < sample.operationStartedAtMs ||
    sample.postMutationProofCapturedAtMs < sample.visibleTargetRenderCapturedAtMs ||
    sample.postMutationProofCapturedAtMs < sample.operationStartedAtMs ||
    sample.restoreProofCapturedAtMs < sample.postMutationProofCapturedAtMs
  ) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} has non-monotonic operation proof timing bounds`)
  }
  invalidReasons.push(
    ...sameCorpusMutationTargetDurationInvalidReasons(
      workload,
      'visible target render',
      sample.visibleTargetRenderMs,
      sample.visibleTargetRenderCapturedAtMs - sample.operationStartedAtMs,
    ),
    ...sameCorpusMutationTargetDurationInvalidReasons(
      workload,
      'committed-state validation',
      sample.committedStateValidationMs,
      sample.postMutationProofCapturedAtMs - sample.visibleTargetRenderCapturedAtMs,
    ),
    ...sameCorpusMutationTargetDurationInvalidReasons(
      workload,
      'restore validation',
      sample.restoreValidationMs,
      sample.restoreProofCapturedAtMs - sample.postMutationProofCapturedAtMs,
    ),
    ...sameCorpusMutationTargetDurationInvalidReasons(
      workload,
      'committed target',
      sample.committedTargetProofMs,
      sample.postMutationProofCapturedAtMs - sample.operationStartedAtMs,
    ),
  )
  return invalidReasons
}

function sameCorpusMutationTargetDurationInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  label: string,
  actual: number,
  expected: number,
): string[] {
  if (!Number.isFinite(actual) || actual < 0 || Math.abs(actual - expected) <= 1) {
    return []
  }
  return [`semantic UI mutation target proof for ${workload} ${label} timing does not match proof window`]
}
