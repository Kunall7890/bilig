import type { SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function sameCorpusBiligVisibleSceneProofInvalidReasons(args: {
  readonly actual: SameCorpusMutationTargetReadback
  readonly expected: SameCorpusMutationTargetReadback
  readonly label: string
  readonly workload: UiResponsivenessSameCorpusWorkload
}): string[] {
  const expectedSceneProof = normalizeVisibleSceneProofSha256(args.expected.visibleSceneProofSha256)
  const actualSceneProof = normalizeVisibleSceneProofSha256(args.actual.visibleSceneProofSha256)
  if (!expectedSceneProof) {
    return [`semantic UI mutation target proof for ${args.workload} ${args.label} is missing target readback visible scene proof`]
  }
  if (!actualSceneProof) {
    return [`semantic UI mutation target proof for ${args.workload} ${args.label} is missing Bilig visible scene proof`]
  }
  return actualSceneProof === expectedSceneProof
    ? []
    : [`semantic UI mutation target proof for ${args.workload} ${args.label} visible scene proof does not match target readback`]
}

function normalizeVisibleSceneProofSha256(value: string | null | undefined): string | null {
  const normalized = value?.trim().toLowerCase() ?? ''
  return /^[a-f0-9]{64}$/u.test(normalized) ? normalized : null
}
