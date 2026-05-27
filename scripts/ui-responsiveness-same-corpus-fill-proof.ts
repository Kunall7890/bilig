import type { SameCorpusMutationTargetProof } from './ui-responsiveness-same-corpus-proof.ts'

export function sameCorpusFillFormatTargetProofInvalidReasons(sample: SameCorpusMutationTargetProof): string[] {
  if (sample.intendedPayload.kind !== 'fill-color') {
    return []
  }
  const expectedFillColor = sample.intendedPayload.expectedFillColor
  const invalidReasons: string[] = []
  if (!sameCorpusFillColorsMatch(sample.after.fillColor, expectedFillColor)) {
    invalidReasons.push('semantic UI mutation target proof for fill-format-change target readback does not match intended fill color')
  }
  if (!sameCorpusFillColorsMatch(sample.visibleAfter.fillColor, expectedFillColor)) {
    invalidReasons.push('semantic UI mutation target proof for fill-format-change rendered target cell does not show intended fill color')
  }
  const committedAfterFill = sample.committedStateProof?.after.readback.fillColor ?? null
  if (sample.product === 'google-sheets' && !sameCorpusFillColorsMatch(committedAfterFill, expectedFillColor)) {
    invalidReasons.push(
      'semantic UI mutation target proof for fill-format-change committed-state after export does not contain intended fill color',
    )
  }
  if (
    sample.after.fillColor !== null &&
    sample.visibleAfter.fillColor !== null &&
    !sameCorpusFillColorsMatch(sample.visibleAfter.fillColor, sample.after.fillColor)
  ) {
    invalidReasons.push('semantic UI mutation target proof for fill-format-change rendered target fill does not match target readback')
  }
  if (
    committedAfterFill !== null &&
    sample.after.fillColor !== null &&
    !sameCorpusFillColorsMatch(committedAfterFill, sample.after.fillColor)
  ) {
    invalidReasons.push('semantic UI mutation target proof for fill-format-change committed-state fill does not match target readback')
  }
  if (sameCorpusFillColorsMatch(sample.before.fillColor, expectedFillColor)) {
    invalidReasons.push('semantic UI mutation target proof for fill-format-change pre-mutation target already had intended fill color')
  }
  if (sameCorpusFillColorsMatch(sample.restored.fillColor, expectedFillColor)) {
    invalidReasons.push('semantic UI mutation target proof for fill-format-change restored target still has intended fill color')
  }
  if (sample.product === 'google-sheets') {
    invalidReasons.push(...sameCorpusGoogleSheetsFillRestoreInvalidReasons(sample, expectedFillColor))
  }
  return invalidReasons
}

function sameCorpusGoogleSheetsFillRestoreInvalidReasons(sample: SameCorpusMutationTargetProof, expectedFillColor: string): string[] {
  const proof = sample.committedStateProof
  if (!proof) {
    return []
  }
  const invalidReasons: string[] = []
  if (sameCorpusFillColorsMatch(proof.before.readback.fillColor, expectedFillColor)) {
    invalidReasons.push(
      'semantic UI mutation target proof for fill-format-change committed-state before export already had intended fill color',
    )
  }
  if (sameCorpusFillColorsMatch(proof.restored.readback.fillColor, expectedFillColor)) {
    invalidReasons.push(
      'semantic UI mutation target proof for fill-format-change committed-state restored export still has intended fill color',
    )
  }
  if (!sameCorpusFillReadbacksMatch(sample.before.fillColor, proof.before.readback.fillColor)) {
    invalidReasons.push(
      'semantic UI mutation target proof for fill-format-change committed-state before fill does not match target readback',
    )
  }
  if (!sameCorpusFillReadbacksMatch(sample.restored.fillColor, proof.restored.readback.fillColor)) {
    invalidReasons.push(
      'semantic UI mutation target proof for fill-format-change committed-state restored fill does not match target readback',
    )
  }
  return invalidReasons
}

function sameCorpusFillReadbacksMatch(left: string | null, right: string | null): boolean {
  const leftColor = sameCorpusFillColorValue(left)
  const rightColor = sameCorpusFillColorValue(right)
  return leftColor === rightColor
}

export function sameCorpusFillColorsMatch(actual: string | null | undefined, expected: string | null | undefined): boolean {
  const actualColor = sameCorpusFillColorValue(actual)
  const expectedColor = sameCorpusFillColorValue(expected)
  return actualColor !== null && expectedColor !== null && actualColor === expectedColor
}

export function sameCorpusFillColorValue(value: string | null | undefined): string | null {
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
