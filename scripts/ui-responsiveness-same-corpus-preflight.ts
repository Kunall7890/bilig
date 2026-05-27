import type {
  SameCorpusCaptureCorpusVerification,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { SameCorpusMutationTargetCommittedStateProof } from './ui-responsiveness-same-corpus-committed-state-proof.ts'
import type { SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'

export interface PreflightEditableMutationProof {
  readonly product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>
  readonly captured: boolean
  readonly method: 'sentinel-write-readback-restore'
  readonly sampleIndex: number
  readonly sheetName: string
  readonly sheetId: string | null
  readonly targetRange: string
  readonly intendedOperation: 'edit-visible-cell'
  readonly intendedValue: string
  readonly before: SameCorpusMutationTargetReadback
  readonly after: SameCorpusMutationTargetReadback
  readonly restored: SameCorpusMutationTargetReadback
  readonly authoritativeReadbackRevision: string
  readonly visibleReadbackRevision: string
  readonly undoRestoreStatus: 'verified' | 'failed'
  readonly evidence: readonly string[]
}

export interface PreflightFillFormatMutationProof {
  readonly product: 'google-sheets'
  readonly captured: boolean
  readonly method: 'fill-color-commit-readback-restore'
  readonly sampleIndex: number
  readonly sheetName: string
  readonly sheetId: string | null
  readonly targetRange: string
  readonly intendedOperation: 'fill-format-change'
  readonly intendedFillColor: string
  readonly swatchLabel: string
  readonly before: SameCorpusMutationTargetReadback
  readonly after: SameCorpusMutationTargetReadback
  readonly restored: SameCorpusMutationTargetReadback
  readonly committedStateProof: SameCorpusMutationTargetCommittedStateProof | null
  readonly undoRestoreStatus: 'verified' | 'failed'
  readonly evidence: readonly string[]
}

export interface PreflightProductResult {
  readonly product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>
  readonly source: string
  readonly finalUrl: string
  readonly title: string
  readonly status: 'ready' | 'blocked'
  readonly blocker: string | null
  readonly corpusVerification: SameCorpusCaptureCorpusVerification | null
  readonly editableMutationProof: PreflightEditableMutationProof | null
  readonly fillFormatMutationProof: PreflightFillFormatMutationProof | null
  readonly limitations: string[]
}

export interface SameCorpusPreflight {
  readonly mode: 'preflight'
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly requiredProductCount: 2
  readonly checkedProductCount: number
  readonly readyProductCount: number
  readonly blockedProductCount: number
  readonly allCheckedProductsReady: boolean
  readonly products: readonly PreflightProductResult[]
}

export function sameCorpusPreflightProductReady(product: PreflightProductResult): boolean {
  return sameCorpusPreflightProductInvalidReasons(product).length === 0
}

export function sameCorpusPreflightProductInvalidReasons(product: PreflightProductResult): readonly string[] {
  if (product.status === 'blocked') {
    return [product.blocker ?? 'blocked without diagnostic']
  }
  const invalidReasons: string[] = []
  if (!product.corpusVerification?.verified) {
    invalidReasons.push(`${product.product} corpus verification is missing or not verified`)
  }
  const proof = product.editableMutationProof
  if (!proof) {
    invalidReasons.push(`${product.product} is missing editable sentinel write/readback/restore proof`)
    return invalidReasons
  }
  if (!proof.captured) {
    invalidReasons.push(`${product.product} editable sentinel proof is not marked captured`)
  }
  if (product.corpusVerification && proof.sheetName !== product.corpusVerification.sheetName) {
    invalidReasons.push(`${product.product} editable sentinel proof sheet does not match corpus verification`)
  }
  if (proof.product !== product.product) {
    invalidReasons.push(`${product.product} editable sentinel proof has mismatched product`)
  }
  if (proof.method !== 'sentinel-write-readback-restore') {
    invalidReasons.push(`${product.product} editable sentinel proof method is stale`)
  }
  if (proof.targetRange.trim().length === 0) {
    invalidReasons.push(`${product.product} editable sentinel proof is missing target range`)
  }
  if (!sameCorpusPreflightReadbackProvesValue(proof.after, proof.intendedValue)) {
    invalidReasons.push(`${product.product} editable sentinel proof did not prove the written value`)
  }
  if (proof.undoRestoreStatus !== 'verified') {
    invalidReasons.push(`${product.product} editable sentinel proof did not verify undo/restore`)
  }
  if (!sameCorpusPreflightReadbacksEqual(proof.before, proof.restored)) {
    invalidReasons.push(`${product.product} editable sentinel proof restored readback does not match before readback`)
  }
  if (proof.authoritativeReadbackRevision.trim().length === 0 || proof.visibleReadbackRevision.trim().length === 0) {
    invalidReasons.push(`${product.product} editable sentinel proof is missing readback revisions`)
  }
  if (product.product === 'google-sheets') {
    invalidReasons.push(...sameCorpusPreflightFillFormatProofInvalidReasons(product))
  }
  return invalidReasons
}

export function sameCorpusPreflightReadbackProvesValue(readback: SameCorpusMutationTargetReadback, value: string): boolean {
  return readback.value === value || readback.visibleText === value || readback.formula === value
}

function sameCorpusPreflightFillFormatProofInvalidReasons(product: PreflightProductResult): readonly string[] {
  const proof = product.fillFormatMutationProof
  if (!proof) {
    return ['google-sheets is missing fill-format commit/readback/restore preflight proof']
  }
  const invalidReasons: string[] = []
  if (proof.product !== 'google-sheets') {
    invalidReasons.push('google-sheets fill-format preflight proof has mismatched product')
  }
  if (!proof.captured) {
    invalidReasons.push('google-sheets fill-format preflight proof is not marked captured')
  }
  if (proof.method !== 'fill-color-commit-readback-restore') {
    invalidReasons.push('google-sheets fill-format preflight proof method is stale')
  }
  if (product.corpusVerification && proof.sheetName !== product.corpusVerification.sheetName) {
    invalidReasons.push('google-sheets fill-format preflight proof sheet does not match corpus verification')
  }
  if (proof.targetRange.trim().length === 0) {
    invalidReasons.push('google-sheets fill-format preflight proof is missing target range')
  }
  if (normalizeColor(proof.after.fillColor) !== normalizeColor(proof.intendedFillColor)) {
    invalidReasons.push('google-sheets fill-format preflight did not prove the intended rendered fill color')
  }
  if (normalizeColor(proof.committedStateProof?.after.readback.fillColor ?? null) !== normalizeColor(proof.intendedFillColor)) {
    invalidReasons.push('google-sheets fill-format preflight did not prove committed workbook fill color')
  }
  if (proof.undoRestoreStatus !== 'verified') {
    invalidReasons.push('google-sheets fill-format preflight did not verify undo/restore')
  }
  if (normalizeColor(proof.before.fillColor) !== normalizeColor(proof.restored.fillColor)) {
    invalidReasons.push('google-sheets fill-format preflight restored fill does not match before fill')
  }
  return invalidReasons
}

function sameCorpusPreflightReadbacksEqual(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return (
    left.value === right.value &&
    left.formula === right.formula &&
    left.fillColor === right.fillColor &&
    left.visibleText === right.visibleText &&
    left.source === right.source
  )
}

function normalizeColor(value: string | null): string | null {
  return value ? value.trim().toLowerCase() : null
}
