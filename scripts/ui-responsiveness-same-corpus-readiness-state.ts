export type SameCorpusClaimReadinessState =
  | 'not-captured'
  | 'diagnostic-capture-incomplete'
  | 'claim-grade-capture-speed-failed'
  | 'claim-grade-10x-passed'

export const sameCorpusTenXSpeedInvalidReason = 'not every required workload is 10x against Google Sheets'

export function sameCorpusClaimReadinessState(args: {
  readonly captured: boolean
  readonly currentContractEvidenceComplete: boolean
  readonly googleSheetsTenXRequirementSatisfied: boolean
  readonly invalidReasons: readonly string[]
}): SameCorpusClaimReadinessState {
  if (!args.captured) {
    return 'not-captured'
  }
  if (args.googleSheetsTenXRequirementSatisfied && args.invalidReasons.length === 0) {
    return 'claim-grade-10x-passed'
  }
  if (args.currentContractEvidenceComplete && args.invalidReasons.every((reason) => reason === sameCorpusTenXSpeedInvalidReason)) {
    return 'claim-grade-capture-speed-failed'
  }
  return 'diagnostic-capture-incomplete'
}

export function parseSameCorpusClaimReadinessState(value: string): SameCorpusClaimReadinessState {
  if (
    value === 'not-captured' ||
    value === 'diagnostic-capture-incomplete' ||
    value === 'claim-grade-capture-speed-failed' ||
    value === 'claim-grade-10x-passed'
  ) {
    return value
  }
  throw new Error(`Unexpected same-corpus claim readiness state: ${value}`)
}
