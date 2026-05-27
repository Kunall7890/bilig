import { hasBiligProductionRuntimeProof } from './ui-responsiveness-same-corpus-guardrails.ts'
import type { SameCorpusBiligRuntimeProof } from './ui-responsiveness-same-corpus-scorecard-types.ts'

export function validateBiligRuntimeProof(proof: SameCorpusBiligRuntimeProof, source: string, caseId: string): void {
  if (proof.product !== 'bilig' || proof.source !== source || proof.verificationMethod !== 'window.__biligRuntimeBuild') {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof identity mismatch for ${caseId}`)
  }
  if (proof.requiredBuildKind !== 'production') {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof has stale required build kind for ${caseId}`)
  }
  if (!['development', 'production', 'unknown'].includes(proof.actualBuildKind)) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof has invalid build kind for ${caseId}`)
  }
  if (proof.sampleCount !== proof.samples.length || proof.sampleCount <= 0) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample count is stale for ${caseId}`)
  }
  for (const sample of proof.samples) {
    if (!Number.isInteger(sample.sampleIndex) || sample.sampleIndex < 0) {
      throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample index is invalid for ${caseId}`)
    }
    if (!['development', 'production', 'unknown'].includes(sample.buildKind)) {
      throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample build kind is invalid for ${caseId}`)
    }
  }
  const verified = hasBiligProductionRuntimeProof({ product: 'bilig', biligRuntimeProof: proof })
  if (proof.verified !== verified) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof verified flag is stale for ${caseId}`)
  }
}
