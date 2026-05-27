import type { SameCorpusMutationTargetProof, SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'
import { sameCorpusFillColorsMatch } from './ui-responsiveness-same-corpus-fill-proof.ts'
import { sameCorpusBiligVisibleSceneProofInvalidReasons } from './ui-responsiveness-same-corpus-bilig-visible-proof.ts'

type SameCorpusMutationTargetScreenshotPhase = 'before' | 'after' | 'restored'

export function sameCorpusMutationTargetScreenshotSemanticInvalidReasons(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  const screenshots = sample.targetScreenshots
  if (!screenshots) {
    return []
  }
  return (['before', 'after', 'restored'] as const).flatMap((phase) =>
    sameCorpusMutationTargetScreenshotPhaseSemanticInvalidReasons(product, workload, sample, phase),
  )
}

function sameCorpusMutationTargetScreenshotPhaseSemanticInvalidReasons(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  phase: SameCorpusMutationTargetScreenshotPhase,
): string[] {
  const screenshots = sample.targetScreenshots
  const screenshot = screenshots?.[phase]
  const expectedReadback = phase === 'after' ? sample.after : phase === 'before' ? sample.before : sample.restored
  const readback = screenshot?.semanticReadback
  if (!readback) {
    return [`semantic UI mutation target proof for ${workload} ${phase} screenshot is missing semantic target readback`]
  }
  if (!sameCorpusScreenshotSemanticSourceAccepted(product, readback.source)) {
    return [
      `semantic UI mutation target proof for ${workload} ${phase} screenshot semantic readback did not come from an accepted browser-visible source`,
    ]
  }
  if (product === 'bilig') {
    const sceneProofInvalidReasons = sameCorpusBiligVisibleSceneProofInvalidReasons({
      actual: readback,
      expected: expectedReadback,
      label: `${phase} screenshot semantic readback`,
      workload,
    })
    if (sceneProofInvalidReasons.length > 0) {
      return sceneProofInvalidReasons
    }
  }
  if (!sameCorpusScreenshotReadbackMatches(product, workload, expectedReadback, readback)) {
    return [`semantic UI mutation target proof for ${workload} ${phase} screenshot semantic readback does not match target readback`]
  }
  return []
}

function sameCorpusScreenshotSemanticSourceAccepted(
  _product: UiResponsivenessSameCorpusProduct,
  source: SameCorpusMutationTargetReadback['source'],
): boolean {
  return source === 'visible-grid-cell'
}

function sameCorpusScreenshotReadbackMatches(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  expected: SameCorpusMutationTargetReadback,
  actual: SameCorpusMutationTargetReadback,
): boolean {
  if (workload === 'fill-format-change') {
    return sameCorpusFillColorMatches(actual.fillColor, expected.fillColor)
  }
  if (workload === 'formula-edit' && actual.source === 'visible-grid-cell') {
    return sameCorpusTextMatches(actual.value, expected.value) || sameCorpusTextMatches(actual.visibleText, expected.visibleText)
  }
  if (workload === 'formula-edit') {
    if (actual.formula !== null || expected.formula !== null) {
      return sameCorpusTextMatches(actual.formula, expected.formula)
    }
    return sameCorpusTextMatches(actual.value, expected.value) || sameCorpusTextMatches(actual.visibleText, expected.visibleText)
  }
  return sameCorpusTextMatches(actual.value, expected.value) || sameCorpusTextMatches(actual.visibleText, expected.visibleText)
}

function sameCorpusTextMatches(actual: string | null, expected: string | null): boolean {
  return actual !== null && expected !== null && actual === expected
}

function sameCorpusFillColorMatches(actual: string | null, expected: string | null): boolean {
  if (actual === null || expected === null) {
    return actual === expected
  }
  return sameCorpusFillColorsMatch(actual, expected)
}
