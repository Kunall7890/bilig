import type { SameCorpusMutationTargetProof } from './ui-responsiveness-same-corpus-proof.ts'
import type {
  SameCorpusMutationTargetTimingSample,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'

export function sameCorpusMutationTargetTimingSample(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetTimingSample {
  return {
    sampleIndex: proof.sampleIndex,
    product: proof.product,
    sheetName: proof.sheetName,
    sheetId: proof.sheetId,
    targetRange: proof.targetRange,
    targetProofSignature: proof.targetProofSignature ?? '',
    committedTargetProofMs: proof.committedTargetProofMs,
    visibleTargetRenderMs: proof.visibleTargetRenderMs,
    committedStateValidationMs: proof.committedStateValidationMs,
    restoreValidationMs: proof.restoreValidationMs,
  }
}

export function cloneSameCorpusMutationTargetTimingSamples(
  samples: readonly SameCorpusMutationTargetTimingSample[],
): SameCorpusMutationTargetTimingSample[] {
  return samples.map((sample) => ({ ...sample }))
}

export function validateSameCorpusMutationTargetTimingSamples(
  samples: readonly SameCorpusMutationTargetTimingSample[] | undefined,
  args: {
    readonly expectedProduct: UiResponsivenessSameCorpusProduct
    readonly expectedLength: number
    readonly label: string
  },
): void {
  if (!samples || samples.length === 0) {
    throw new Error(`${args.label} is missing mutation target timing samples`)
  }
  if (samples.length !== args.expectedLength) {
    throw new Error(
      `${args.label} has ${String(samples.length)} mutation target timing samples but expected ${String(args.expectedLength)}`,
    )
  }
  const sampleIndexes = new Set<number>()
  for (const sample of samples) {
    validateSameCorpusMutationTargetTimingSample(sample, args.expectedProduct, args.label)
    if (sampleIndexes.has(sample.sampleIndex)) {
      throw new Error(`${args.label} has duplicate mutation target timing sample index ${String(sample.sampleIndex)}`)
    }
    sampleIndexes.add(sample.sampleIndex)
  }
  for (let sampleIndex = 0; sampleIndex < args.expectedLength; sampleIndex += 1) {
    if (!sampleIndexes.has(sampleIndex)) {
      throw new Error(`${args.label} is missing mutation target timing sample index ${String(sampleIndex)}`)
    }
  }
}

export function validateSameCorpusMutationTargetTimingSamplesMatchArrays(
  samples: readonly SameCorpusMutationTargetTimingSample[] | undefined,
  args: {
    readonly expectedProduct: UiResponsivenessSameCorpusProduct
    readonly expectedLength: number
    readonly label: string
    readonly committedTargetProofMsSamples: readonly number[] | undefined
    readonly visibleTargetRenderMsSamples: readonly number[] | undefined
    readonly committedStateValidationMsSamples: readonly number[] | undefined
    readonly restoreValidationMsSamples: readonly number[] | undefined
  },
): void {
  validateSameCorpusMutationTargetTimingSamples(samples, args)
  const checkedSamples = samples ?? []
  for (const sample of checkedSamples) {
    const sampleIndex = sample.sampleIndex
    assertSampleTimingMatchesArray(
      sample.committedTargetProofMs,
      args.committedTargetProofMsSamples,
      sampleIndex,
      args.label,
      'committed target proof',
    )
    assertSampleTimingMatchesArray(
      sample.visibleTargetRenderMs,
      args.visibleTargetRenderMsSamples,
      sampleIndex,
      args.label,
      'visible target render',
    )
    assertSampleTimingMatchesArray(
      sample.committedStateValidationMs,
      args.committedStateValidationMsSamples,
      sampleIndex,
      args.label,
      'committed-state validation',
    )
    assertSampleTimingMatchesArray(
      sample.restoreValidationMs,
      args.restoreValidationMsSamples,
      sampleIndex,
      args.label,
      'restore validation',
    )
  }
}

export function sameCorpusMutationTargetTimingSampleMatchesProof(
  sample: SameCorpusMutationTargetTimingSample,
  proof: SameCorpusMutationTargetProof,
): boolean {
  return (
    sample.sampleIndex === proof.sampleIndex &&
    sample.product === proof.product &&
    sample.sheetName === proof.sheetName &&
    sample.sheetId === proof.sheetId &&
    sample.targetRange === proof.targetRange &&
    sample.targetProofSignature === proof.targetProofSignature &&
    sameCorpusTimingValuesMatch(sample.committedTargetProofMs, proof.committedTargetProofMs) &&
    sameCorpusTimingValuesMatch(sample.visibleTargetRenderMs, proof.visibleTargetRenderMs) &&
    sameCorpusTimingValuesMatch(sample.committedStateValidationMs, proof.committedStateValidationMs) &&
    sameCorpusTimingValuesMatch(sample.restoreValidationMs, proof.restoreValidationMs)
  )
}

export function sameCorpusTimingValuesMatch(left: number, right: number): boolean {
  return Number.isFinite(left) && left >= 0 && Number.isFinite(right) && right >= 0 && Math.abs(left - right) <= 1
}

function validateSameCorpusMutationTargetTimingSample(
  sample: SameCorpusMutationTargetTimingSample,
  expectedProduct: UiResponsivenessSameCorpusProduct,
  label: string,
): void {
  if (sample.product !== expectedProduct) {
    throw new Error(`${label} mutation target timing sample has product ${sample.product}, expected ${expectedProduct}`)
  }
  if (!Number.isInteger(sample.sampleIndex) || sample.sampleIndex < 0) {
    throw new Error(`${label} mutation target timing sample has invalid sample index`)
  }
  if (sample.sheetName.trim().length === 0) {
    throw new Error(`${label} mutation target timing sample is missing sheet name`)
  }
  if (sample.targetRange.trim().length === 0) {
    throw new Error(`${label} mutation target timing sample is missing target range`)
  }
  if (sample.targetProofSignature.trim().length === 0) {
    throw new Error(`${label} mutation target timing sample is missing target proof signature`)
  }
  validateTiming(sample.committedTargetProofMs, label, 'committed target proof')
  validateTiming(sample.visibleTargetRenderMs, label, 'visible target render')
  validateTiming(sample.committedStateValidationMs, label, 'committed-state validation')
  validateTiming(sample.restoreValidationMs, label, 'restore validation')
}

function validateTiming(value: number, label: string, timingLabel: string): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`${label} mutation target timing sample has invalid ${timingLabel} timing`)
  }
}

function assertSampleTimingMatchesArray(
  value: number,
  samples: readonly number[] | undefined,
  sampleIndex: number,
  label: string,
  timingLabel: string,
): void {
  const expectedValue = samples?.[sampleIndex]
  if (expectedValue === undefined || !sameCorpusTimingValuesMatch(value, expectedValue)) {
    throw new Error(`${label} mutation target timing sample ${String(sampleIndex)} does not match ${timingLabel} samples`)
  }
}
