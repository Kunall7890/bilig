import type { SameCorpusMutationTargetProofProductSummary } from './ui-responsiveness-same-corpus-scorecard-types.ts'

export function sameCorpusMutationTargetProofProductEvidenceLines(
  summaries: readonly SameCorpusMutationTargetProofProductSummary[],
): readonly string[] {
  return actionableMutationTargetProofProductSummaries(summaries).map(formatMutationTargetProofProductSummary)
}

export function sameCorpusMutationTargetProofProductGapLines(
  summaries: readonly SameCorpusMutationTargetProofProductSummary[],
): readonly string[] {
  return actionableMutationTargetProofProductSummaries(summaries)
    .filter((summary) => !summary.accepted)
    .map((summary) => `same-corpus mutation target proof gap: ${formatMutationTargetProofProductSummary(summary)}`)
}

function actionableMutationTargetProofProductSummaries(
  summaries: readonly SameCorpusMutationTargetProofProductSummary[],
): readonly SameCorpusMutationTargetProofProductSummary[] {
  return summaries.filter((summary) => summary.requiredSampleCount > 0 || summary.rawSampleCount > 0)
}

function formatMutationTargetProofProductSummary(summary: SameCorpusMutationTargetProofProductSummary): string {
  const missingSampleIndexes = summary.samples.filter((sample) => !sample.present).map((sample) => String(sample.sampleIndex))
  const rejectedSampleIndexes = summary.samples
    .filter((sample) => sample.present && !sample.accepted)
    .map((sample) => String(sample.sampleIndex))
  const sampleInvalidReasons = summary.samples.flatMap((sample) =>
    sample.present && !sample.accepted ? sample.invalidReasons.map((reason) => `sample ${String(sample.sampleIndex)}: ${reason}`) : [],
  )
  return [
    `${summary.workload}/${summary.product} accepted ${String(summary.acceptedSampleCount)}/${String(summary.requiredSampleCount)} samples (raw ${String(summary.rawSampleCount)})`,
    `missing samples: ${missingSampleIndexes.join(', ') || 'none'}`,
    `rejected samples: ${rejectedSampleIndexes.join(', ') || 'none'}`,
    `invalid reasons: ${formatInvalidReasons([...sampleInvalidReasons, ...summary.invalidReasons])}`,
  ].join('; ')
}

function formatInvalidReasons(reasons: readonly string[]): string {
  if (reasons.length === 0) {
    return 'none'
  }
  const visibleReasons = reasons.slice(0, 3)
  const hiddenReasonCount = reasons.length - visibleReasons.length
  return `${visibleReasons.join('; ')}${hiddenReasonCount > 0 ? `; +${String(hiddenReasonCount)} more` : ''}`
}
