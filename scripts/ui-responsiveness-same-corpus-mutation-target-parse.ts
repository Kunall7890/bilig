import { arrayField, asObject, booleanField, numberField, objectField, stringArrayField, stringField } from './json-scorecard-helpers.ts'
import type {
  SameCorpusMutationTargetProofProductSummary,
  SameCorpusMutationTargetProofSampleSummary,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type { SameCorpusMutationTargetProof, SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-proof.ts'
import { isUiResponsivenessSameCorpusWorkload, type UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function parseSameCorpusMutationTargetProof(value: unknown): SameCorpusMutationTargetProof {
  const record = asObject(value, 'UI responsiveness same-corpus mutation target proof')
  return {
    product: parseSameCorpusProduct(stringField(record, 'product')),
    sampleIndex: numberField(record, 'sampleIndex'),
    committedTargetProofMs: optionalNumberField(record, 'committedTargetProofMs') ?? Number.NaN,
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    intendedOperation: parseSameCorpusMutatingWorkload(stringField(record, 'intendedOperation')),
    intendedPayload: parseSameCorpusMutationTargetIntendedPayload(objectField(record, 'intendedPayload')),
    sheetName: stringField(record, 'sheetName'),
    sheetId: Object.hasOwn(record, 'sheetId') ? nullableStringField(record, 'sheetId') : null,
    targetRange: stringField(record, 'targetRange'),
    before: parseSameCorpusMutationTargetReadback(objectField(record, 'before')),
    after: parseSameCorpusMutationTargetReadback(objectField(record, 'after')),
    restored: parseSameCorpusMutationTargetReadback(objectField(record, 'restored')),
    visibleAfter: Object.hasOwn(record, 'visibleAfter')
      ? parseSameCorpusMutationTargetReadback(objectField(record, 'visibleAfter'))
      : missingSameCorpusMutationTargetReadback(),
    visibleRestored: Object.hasOwn(record, 'visibleRestored')
      ? parseSameCorpusMutationTargetReadback(objectField(record, 'visibleRestored'))
      : missingSameCorpusMutationTargetReadback(),
    visibleAfterSelectedRange: Object.hasOwn(record, 'visibleAfterSelectedRange')
      ? nullableStringField(record, 'visibleAfterSelectedRange')
      : null,
    visibleRestoredSelectedRange: Object.hasOwn(record, 'visibleRestoredSelectedRange')
      ? nullableStringField(record, 'visibleRestoredSelectedRange')
      : null,
    authoritativeReadbackRevision: nullableStringField(record, 'authoritativeReadbackRevision'),
    visibleRenderRevision: nullableStringField(record, 'visibleRenderRevision'),
    targetScreenshots: Object.hasOwn(record, 'targetScreenshots')
      ? parseSameCorpusMutationTargetScreenshotProofSet(objectField(record, 'targetScreenshots'))
      : null,
    screenshotPath: nullableStringField(record, 'screenshotPath'),
    screenshotSha256: nullableStringField(record, 'screenshotSha256'),
    undoRestoreStatus: parseSameCorpusMutationUndoRestoreStatus(stringField(record, 'undoRestoreStatus')),
  }
}

export function parseSameCorpusMutationTargetProofProductSummary(value: unknown): SameCorpusMutationTargetProofProductSummary {
  const record = asObject(value, 'UI responsiveness same-corpus mutation target proof product summary')
  return {
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    product: parseSameCorpusProduct(stringField(record, 'product')),
    requiredSampleCount: numberField(record, 'requiredSampleCount'),
    rawSampleCount: numberField(record, 'rawSampleCount'),
    acceptedSampleCount: numberField(record, 'acceptedSampleCount'),
    accepted: booleanField(record, 'accepted'),
    samples: arrayField(record, 'samples').map(parseSameCorpusMutationTargetProofSampleSummary),
    invalidReasons: stringArrayField(record, 'invalidReasons'),
  }
}

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
    product: parseSameCorpusProduct(stringField(value, 'product')),
    scope: parseSameCorpusMutationTargetScreenshotScope(stringField(value, 'scope')),
    sampleIndex: numberField(value, 'sampleIndex'),
    sheetId: Object.hasOwn(value, 'sheetId') ? nullableStringField(value, 'sheetId') : null,
    sheetName: stringField(value, 'sheetName'),
    targetRange: stringField(value, 'targetRange'),
    workload: parseSameCorpusWorkload(stringField(value, 'workload')),
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

function optionalNumberField(value: Record<string, unknown>, key: string): number | undefined {
  if (!Object.hasOwn(value, key)) {
    return undefined
  }
  const fieldValue = value[key]
  if (typeof fieldValue !== 'number' || !Number.isFinite(fieldValue)) {
    throw new Error(`Expected ${key} to be a finite number`)
  }
  return fieldValue
}

function parseSameCorpusMutationTargetProofSampleSummary(value: unknown): SameCorpusMutationTargetProofSampleSummary {
  const record = asObject(value, 'UI responsiveness same-corpus mutation target proof sample summary')
  return {
    sampleIndex: numberField(record, 'sampleIndex'),
    present: booleanField(record, 'present'),
    accepted: booleanField(record, 'accepted'),
    product: Object.hasOwn(record, 'product') ? nullableSameCorpusProduct(record, 'product') : null,
    committedTargetProofMs: Object.hasOwn(record, 'committedTargetProofMs') ? nullableNumberField(record, 'committedTargetProofMs') : null,
    sheetName: nullableStringField(record, 'sheetName'),
    sheetId: Object.hasOwn(record, 'sheetId') ? nullableStringField(record, 'sheetId') : null,
    targetRange: nullableStringField(record, 'targetRange'),
    intendedOperation: nullableSameCorpusMutatingWorkload(record, 'intendedOperation'),
    intendedPayload: nullableSameCorpusMutationTargetIntendedPayload(record, 'intendedPayload'),
    before: nullableSameCorpusMutationTargetReadback(record, 'before'),
    after: nullableSameCorpusMutationTargetReadback(record, 'after'),
    restored: nullableSameCorpusMutationTargetReadback(record, 'restored'),
    visibleAfter: nullableSameCorpusMutationTargetReadback(record, 'visibleAfter'),
    visibleRestored: nullableSameCorpusMutationTargetReadback(record, 'visibleRestored'),
    visibleAfterSelectedRange: Object.hasOwn(record, 'visibleAfterSelectedRange')
      ? nullableStringField(record, 'visibleAfterSelectedRange')
      : null,
    visibleRestoredSelectedRange: Object.hasOwn(record, 'visibleRestoredSelectedRange')
      ? nullableStringField(record, 'visibleRestoredSelectedRange')
      : null,
    authoritativeReadbackRevision: nullableStringField(record, 'authoritativeReadbackRevision'),
    visibleRenderRevision: nullableStringField(record, 'visibleRenderRevision'),
    targetScreenshots: Object.hasOwn(record, 'targetScreenshots')
      ? record.targetScreenshots === null
        ? null
        : parseSameCorpusMutationTargetScreenshotProofSet(objectField(record, 'targetScreenshots'))
      : null,
    screenshotPath: nullableStringField(record, 'screenshotPath'),
    screenshotSha256: nullableStringField(record, 'screenshotSha256'),
    undoRestoreStatus: nullableSameCorpusMutationUndoRestoreStatus(record, 'undoRestoreStatus'),
    invalidReasons: Object.hasOwn(record, 'invalidReasons') ? stringArrayField(record, 'invalidReasons') : [],
  }
}

function nullableSameCorpusProduct(record: Record<string, unknown>, key: string): UiResponsivenessSameCorpusProduct | null {
  return record[key] === null ? null : parseSameCorpusProduct(stringField(record, key))
}

function parseSameCorpusMutationTargetIntendedPayload(value: Record<string, unknown>): SameCorpusMutationTargetProof['intendedPayload'] {
  const kind = stringField(value, 'kind')
  if (kind === 'cell-value') {
    return { kind, value: stringField(value, 'value') }
  }
  if (kind === 'formula') {
    return { kind, formula: stringField(value, 'formula') }
  }
  if (kind === 'fill-color') {
    return {
      kind,
      expectedFillColor: Object.hasOwn(value, 'expectedFillColor') ? stringField(value, 'expectedFillColor') : '',
      swatchLabel: stringField(value, 'swatchLabel'),
    }
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation target payload kind: ${kind}`)
}

function missingSameCorpusMutationTargetReadback(): SameCorpusMutationTargetReadback {
  return {
    fillColor: null,
    formula: null,
    source: 'unknown',
    value: null,
    visibleText: null,
  }
}

function parseSameCorpusMutationTargetReadback(value: Record<string, unknown>): SameCorpusMutationTargetReadback {
  const batchId = optionalNumberField(value, 'batchId')
  const capturedRevision = Object.hasOwn(value, 'capturedRevision') ? nullableStringField(value, 'capturedRevision') : undefined
  const visibleSceneProofSha256 = Object.hasOwn(value, 'visibleSceneProofSha256')
    ? nullableStringField(value, 'visibleSceneProofSha256')
    : undefined
  return {
    value: nullableStringField(value, 'value'),
    formula: nullableStringField(value, 'formula'),
    fillColor: nullableStringField(value, 'fillColor'),
    visibleText: nullableStringField(value, 'visibleText'),
    source: Object.hasOwn(value, 'source') ? parseSameCorpusMutationTargetReadbackSource(stringField(value, 'source')) : 'unknown',
    ...(batchId !== undefined ? { batchId } : {}),
    ...(capturedRevision !== undefined ? { capturedRevision } : {}),
    ...(visibleSceneProofSha256 !== undefined ? { visibleSceneProofSha256 } : {}),
  }
}

function parseSameCorpusMutationTargetReadbackSource(value: string): SameCorpusMutationTargetReadback['source'] {
  if (value === 'bilig-authoritative-range' || value === 'visible-formula-bar' || value === 'visible-grid-cell' || value === 'unknown') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation readback source: ${value}`)
}

function parseSameCorpusMutatingWorkload(value: string): SameCorpusMutationTargetProof['intendedOperation'] {
  if (value === 'edit-visible-cell' || value === 'formula-edit' || value === 'fill-format-change') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutating workload: ${value}`)
}

function parseSameCorpusMutationUndoRestoreStatus(value: string): SameCorpusMutationTargetProof['undoRestoreStatus'] {
  if (value === 'verified' || value === 'missing' || value === 'failed') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation undo restore status: ${value}`)
}

function nullableSameCorpusMutatingWorkload(
  value: Record<string, unknown>,
  key: string,
): SameCorpusMutationTargetProof['intendedOperation'] | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${key} to be a same-corpus mutating workload or null`)
  }
  return parseSameCorpusMutatingWorkload(fieldValue)
}

function nullableSameCorpusMutationTargetIntendedPayload(
  value: Record<string, unknown>,
  key: string,
): SameCorpusMutationTargetProof['intendedPayload'] | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  return parseSameCorpusMutationTargetIntendedPayload(asObject(fieldValue, key))
}

function nullableSameCorpusMutationTargetReadback(value: Record<string, unknown>, key: string): SameCorpusMutationTargetReadback | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  return parseSameCorpusMutationTargetReadback(asObject(fieldValue, key))
}

function nullableSameCorpusMutationUndoRestoreStatus(
  value: Record<string, unknown>,
  key: string,
): SameCorpusMutationTargetProof['undoRestoreStatus'] | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${key} to be a same-corpus mutation undo status or null`)
  }
  return parseSameCorpusMutationUndoRestoreStatus(fieldValue)
}

function nullableNumberField(value: Record<string, unknown>, key: string): number | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'number' || !Number.isFinite(fieldValue)) {
    throw new Error(`Expected ${key} to be a finite number or null`)
  }
  return fieldValue
}

function parseSameCorpusProduct(value: string): UiResponsivenessSameCorpusProduct {
  if (value === 'bilig' || value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected same-corpus product: ${value}`)
}

function parseSameCorpusWorkload(value: string): UiResponsivenessSameCorpusWorkload {
  if (isUiResponsivenessSameCorpusWorkload(value)) {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus workload: ${value}`)
}
