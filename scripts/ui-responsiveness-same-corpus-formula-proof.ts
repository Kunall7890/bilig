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
  if (product !== 'bilig' && sample.visibleAfter.source === 'visible-grid-target-screenshot') {
    const screenshots = sample.targetScreenshots
    const beforeHash = normalizeScreenshotHash(screenshots?.before.screenshotSha256)
    const afterHash = normalizeScreenshotHash(screenshots?.after.screenshotSha256)
    const restoredHash = normalizeScreenshotHash(screenshots?.restored.screenshotSha256)
    return Boolean(
      beforeHash &&
      afterHash &&
      restoredHash &&
      beforeHash !== afterHash &&
      afterHash !== restoredHash &&
      sample.committedStateProof?.after.readback.value === expectedRenderedValue,
    )
  }
  return false
}

function normalizeScreenshotHash(value: string | null | undefined): string | null {
  const hash = value?.trim().toLowerCase() ?? ''
  return /^[a-f0-9]{64}$/u.test(hash) ? hash : null
}
