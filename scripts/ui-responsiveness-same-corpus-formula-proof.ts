import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type { SameCorpusMutationTargetProof } from './ui-responsiveness-same-corpus-semantic-proof.ts'

export function sameCorpusFormulaEditRenderedResultProven(
  product: UiResponsivenessSameCorpusProduct,
  sample: SameCorpusMutationTargetProof,
): boolean {
  if (sample.intendedPayload.kind !== 'formula') {
    return false
  }
  const expectedRenderedValue = String(sample.sampleIndex + 2)
  if (sample.visibleAfter.value === expectedRenderedValue || sample.visibleAfter.visibleText === expectedRenderedValue) {
    return true
  }
  return false
}
