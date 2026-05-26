import type { SameCorpusMutationTargetProof } from './ui-responsiveness-same-corpus-proof.ts'

export function parseSameCorpusMutationTargetScreenshotProofSet(value: unknown): SameCorpusMutationTargetProof['targetScreenshots'] {
  const record = asObject(value, 'UI responsiveness same-corpus mutation target screenshot proof set')
  return {
    before: parseSameCorpusMutationTargetScreenshotProof(objectField(record, 'before')),
    after: parseSameCorpusMutationTargetScreenshotProof(objectField(record, 'after')),
    restored: parseSameCorpusMutationTargetScreenshotProof(objectField(record, 'restored')),
  }
}

function parseSameCorpusMutationTargetScreenshotProof(
  value: Record<string, unknown>,
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before'] {
  return {
    phase: parseSameCorpusMutationTargetScreenshotPhase(stringField(value, 'phase')),
    scope: parseSameCorpusMutationTargetScreenshotScope(stringField(value, 'scope')),
    targetRange: stringField(value, 'targetRange'),
    screenshotPath: nullableStringField(value, 'screenshotPath'),
    screenshotSha256: nullableStringField(value, 'screenshotSha256'),
  }
}

function parseSameCorpusMutationTargetScreenshotPhase(
  value: string,
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before']['phase'] {
  if (value === 'before' || value === 'after' || value === 'restored') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation target screenshot phase: ${value}`)
}

function parseSameCorpusMutationTargetScreenshotScope(
  value: string,
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before']['scope'] {
  if (value === 'target-cell' || value === 'visible-grid-fallback') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation target screenshot scope: ${value}`)
}

function objectField(value: Record<string, unknown>, key: string): Record<string, unknown> {
  return asObject(value[key], key)
}

function asObject(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${label} to be an object`)
  }
  return Object.fromEntries(Object.entries(value))
}

function stringField(value: Record<string, unknown>, key: string): string {
  const fieldValue = value[key]
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${key} to be a string`)
  }
  return fieldValue
}

function nullableStringField(value: Record<string, unknown>, key: string): string | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${key} to be a string or null`)
  }
  return fieldValue
}
