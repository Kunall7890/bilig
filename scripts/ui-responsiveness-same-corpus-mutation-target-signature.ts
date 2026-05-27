import { createHash } from 'node:crypto'

import type { SameCorpusMutationTargetProof } from './ui-responsiveness-same-corpus-semantic-proof.ts'

export const sameCorpusMutationTargetProofSignaturePrefix = 'same-corpus-mutation-target-proof-sha256:'

export function sameCorpusMutationTargetProofSignature(
  proof: SameCorpusMutationTargetProof | Omit<SameCorpusMutationTargetProof, 'targetProofSignature'>,
): string {
  const { targetProofSignature: _targetProofSignature, ...signedFields } = proof as SameCorpusMutationTargetProof & {
    readonly targetProofSignature?: string | null
  }
  return `${sameCorpusMutationTargetProofSignaturePrefix}${sha256Hex(stableJsonValue(signedFields))}`
}

export function isSameCorpusMutationTargetProofSignature(value: string | null | undefined): value is string {
  return typeof value === 'string' && new RegExp(`^${sameCorpusMutationTargetProofSignaturePrefix}[a-f0-9]{64}$`, 'u').test(value)
}

function sha256Hex(value: unknown): string {
  return createHash('sha256').update(JSON.stringify(value)).digest('hex')
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter((entry) => entry[1] !== undefined)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, stableJsonValue(entry)]),
    )
  }
  return value
}
