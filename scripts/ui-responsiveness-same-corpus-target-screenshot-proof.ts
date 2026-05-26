import type { SameCorpusMutationTargetProof, SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

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
  if (!sameCorpusScreenshotReadbackMatches(product, workload, expectedReadback, readback)) {
    return [`semantic UI mutation target proof for ${workload} ${phase} screenshot semantic readback does not match target readback`]
  }
  return []
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
  if (workload === 'formula-edit' && product === 'bilig') {
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
  const actualColor = sameCorpusFillColorValue(actual)
  const expectedColor = sameCorpusFillColorValue(expected)
  return actualColor !== null && expectedColor !== null && actualColor === expectedColor
}

function sameCorpusFillColorValue(value: string | null): string | null {
  const trimmed = value?.trim().toLowerCase() ?? ''
  if (/^#[0-9a-f]{6}$/u.test(trimmed)) {
    return trimmed
  }
  const rgbMatch = trimmed.match(/^rgba?\((\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})(?:,\s*(?:0|0?\.\d+|1(?:\.0)?))?\)$/u)
  if (!rgbMatch) {
    return null
  }
  const channels = rgbMatch.slice(1, 4).map((channel) => Number(channel))
  if (channels.some((channel) => !Number.isInteger(channel) || channel < 0 || channel > 255)) {
    return null
  }
  return `#${channels.map((channel) => channel.toString(16).padStart(2, '0')).join('')}`
}
