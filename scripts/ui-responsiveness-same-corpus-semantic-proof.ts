import { createHash } from 'node:crypto'

import type { Page } from '@playwright/test'

import type {
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureVerifiedCell,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import type { SameCorpusProductPixelGridProof } from './ui-responsiveness-same-corpus-proof.ts'
import { uiSameCorpusWorkloadMutatesWorkbook, type UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export interface SameCorpusSemanticUiProof {
  readonly captured: boolean
  readonly requiredProducts: readonly UiResponsivenessSameCorpusProduct[]
  readonly products: readonly SameCorpusProductSemanticUiProof[]
  readonly productVerdicts: readonly SameCorpusProductSemanticUiProofVerdict[]
  readonly missingProducts: readonly UiResponsivenessSameCorpusProduct[]
}

export interface SameCorpusProductSemanticUiProof {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly captured: boolean
  readonly method: 'bilig-visible-semantic-readback' | 'google-sheets-visible-semantic-readback' | 'excel-web-visible-semantic-readback'
  readonly sheetName: string
  readonly sheetId: string | null
  readonly selectedRange: string | null
  readonly checkedCells: readonly SameCorpusCaptureVerifiedCell[]
  readonly authoritativeRenderRevision: string | null
  readonly visibleRenderRevision: string | null
  readonly screenshotSha256: string | null
  readonly mutationTargetProofs: readonly SameCorpusMutationTargetProof[]
  readonly evidence: readonly string[]
}

export interface SameCorpusMutationTargetProof {
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly intendedOperation: 'edit-visible-cell' | 'formula-edit' | 'fill-format-change'
  readonly intendedPayload: SameCorpusMutationTargetIntendedPayload
  readonly sheetName: string
  readonly targetRange: string
  readonly before: SameCorpusMutationTargetReadback
  readonly after: SameCorpusMutationTargetReadback
  readonly restored: SameCorpusMutationTargetReadback
  readonly visibleAfter: SameCorpusMutationTargetReadback
  readonly visibleRestored: SameCorpusMutationTargetReadback
  readonly authoritativeReadbackRevision: string | null
  readonly visibleRenderRevision: string | null
  readonly screenshotSha256: string | null
  readonly screenshotPath: string | null
  readonly undoRestoreStatus: 'verified' | 'missing' | 'failed'
}

export type SameCorpusMutationTargetIntendedPayload =
  | { readonly kind: 'cell-value'; readonly value: string }
  | { readonly kind: 'formula'; readonly formula: string }
  | { readonly kind: 'fill-color'; readonly expectedFillColor: string; readonly swatchLabel: string }

export interface SameCorpusMutationTargetReadback {
  readonly value: string | null
  readonly formula: string | null
  readonly fillColor: string | null
  readonly visibleText: string | null
  readonly source: SameCorpusMutationTargetReadbackSource
  readonly batchId?: number | null
  readonly capturedRevision?: string | null
  readonly visibleSceneProofSha256?: string | null
}

export type SameCorpusMutationTargetReadbackSource = 'bilig-authoritative-range' | 'visible-formula-bar' | 'visible-grid-cell' | 'unknown'

export type SameCorpusProductSemanticUiProofEvidenceStatus = 'current-contract' | 'missing' | 'invalid'

export interface SameCorpusProductSemanticUiProofVerdict {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly evidenceStatus: SameCorpusProductSemanticUiProofEvidenceStatus
  readonly acceptedForCurrentScorecard: boolean
  readonly invalidReasons: readonly string[]
}

export interface SameCorpusMutationTargetProofSampleVerdict {
  readonly acceptedForCurrentScorecard: boolean
  readonly invalidReasons: readonly string[]
}

export interface SameCorpusScreenshotBuffer {
  toString(encoding: 'base64'): string
}

export interface SameCorpusCapturedProductScreenshot {
  readonly buffer: SameCorpusScreenshotBuffer | null
  readonly captured: boolean
}

export function validateSameCorpusProductSemanticUiProof(
  proof: SameCorpusProductSemanticUiProof,
  options: { readonly workload?: UiResponsivenessSameCorpusWorkload; readonly sampleCount?: number } = {},
): SameCorpusProductSemanticUiProofVerdict {
  const invalidReasons = sameCorpusProductSemanticUiInvalidReasons(proof, options)
  const acceptedForCurrentScorecard = invalidReasons.length === 0
  return {
    product: proof.product,
    evidenceStatus: acceptedForCurrentScorecard ? 'current-contract' : proof.captured ? 'invalid' : 'missing',
    acceptedForCurrentScorecard,
    invalidReasons,
  }
}

export function validateSameCorpusMutationTargetProofSample(
  proof: SameCorpusProductSemanticUiProof,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  options: {
    readonly duplicateSampleIndex?: boolean
    readonly duplicateScreenshotPath?: boolean
    readonly sampleCount: number
  },
): SameCorpusMutationTargetProofSampleVerdict {
  const invalidReasons = sameCorpusMutationTargetProofSampleInvalidReasons(proof, workload, sample, options)
  return {
    acceptedForCurrentScorecard: invalidReasons.length === 0,
    invalidReasons,
  }
}

export function missingSemanticUiProof(product: UiResponsivenessSameCorpusProduct): SameCorpusProductSemanticUiProof {
  return {
    product,
    captured: false,
    method: semanticUiProofMethod(product),
    sheetName: '',
    sheetId: null,
    selectedRange: null,
    checkedCells: [],
    authoritativeRenderRevision: null,
    visibleRenderRevision: null,
    screenshotSha256: null,
    mutationTargetProofs: [],
    evidence: ['semantic UI proof missing'],
  }
}

export async function readProductSemanticUiProof(args: {
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly screenshot: SameCorpusCapturedProductScreenshot
  readonly pixelGridProof: SameCorpusProductPixelGridProof
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleCount: number
  readonly mutationTargetProofs?: readonly SameCorpusMutationTargetProof[]
}): Promise<SameCorpusProductSemanticUiProof> {
  const selectedRange = await readSameCorpusVisibleSelectedRange(args.page, args.product)
  const evidence = sameCorpusSemanticUiEvidence(args.product, args.pixelGridProof, args.corpusVerification, selectedRange)
  const screenshotSha256 = args.screenshot.buffer ? screenshotBufferSha256(args.screenshot.buffer) : null
  const proof: SameCorpusProductSemanticUiProof = {
    product: args.product,
    captured:
      args.screenshot.captured &&
      args.corpusVerification.verified &&
      selectedRange !== null &&
      screenshotSha256 !== null &&
      args.corpusVerification.checkedCells.length >= 3 &&
      args.corpusVerification.checkedCells.every((cell) => cell.expected === cell.actual),
    method: semanticUiProofMethod(args.product),
    sheetName: args.corpusVerification.sheetName,
    sheetId: null,
    selectedRange,
    checkedCells: args.corpusVerification.checkedCells.map((cell) => ({ ...cell })),
    authoritativeRenderRevision: sameCorpusEvidenceMap(args.pixelGridProof.evidence).get('gridAuthoritativeRevision') ?? null,
    visibleRenderRevision: sameCorpusEvidenceMap(args.pixelGridProof.evidence).get('visibleRenderRevision') ?? null,
    screenshotSha256,
    mutationTargetProofs:
      args.mutationTargetProofs?.map((sample) => ({
        ...sample,
        before: { ...sample.before },
        after: { ...sample.after },
        restored: { ...sample.restored },
        visibleAfter: { ...sample.visibleAfter },
        visibleRestored: { ...sample.visibleRestored },
      })) ?? [],
    evidence,
  }
  return proof
}

function sameCorpusProductSemanticUiInvalidReasons(
  proof: SameCorpusProductSemanticUiProof,
  options: { readonly workload?: UiResponsivenessSameCorpusWorkload; readonly sampleCount?: number },
): string[] {
  const invalidReasons: string[] = []
  if (!proof.captured) {
    invalidReasons.push('semantic UI proof is not marked captured')
  }
  if (proof.sheetName.trim().length === 0) {
    invalidReasons.push('semantic UI proof is missing sheet name')
  }
  if (proof.selectedRange === null || proof.selectedRange.trim().length === 0) {
    invalidReasons.push('semantic UI proof is missing selected range')
  }
  if (proof.checkedCells.length < 3) {
    invalidReasons.push('semantic UI proof checks fewer than 3 cells')
  }
  for (const cell of proof.checkedCells) {
    if (cell.address.trim().length === 0 || cell.expected !== cell.actual) {
      invalidReasons.push(`semantic UI checked cell mismatch at ${cell.address || 'unknown'}`)
    }
  }
  if (proof.screenshotSha256 === null || !/^[a-f0-9]{64}$/u.test(proof.screenshotSha256)) {
    invalidReasons.push('semantic UI proof is missing screenshot SHA256')
  }
  if (proof.product === 'bilig') {
    if (proof.method !== 'bilig-visible-semantic-readback') {
      invalidReasons.push('Bilig semantic UI proof method is not bilig-visible-semantic-readback')
    }
    if (proof.authoritativeRenderRevision === null || proof.authoritativeRenderRevision.trim().length === 0) {
      invalidReasons.push('Bilig semantic UI proof is missing authoritative render revision')
    }
    if (proof.visibleRenderRevision === null || proof.visibleRenderRevision.trim().length === 0) {
      invalidReasons.push('Bilig semantic UI proof is missing visible render revision')
    }
  } else if (proof.product === 'google-sheets') {
    if (proof.method !== 'google-sheets-visible-semantic-readback') {
      invalidReasons.push('Google Sheets semantic UI proof method is not google-sheets-visible-semantic-readback')
    }
  } else if (proof.method !== 'excel-web-visible-semantic-readback') {
    invalidReasons.push('Excel Web semantic UI proof method is not excel-web-visible-semantic-readback')
  }
  if (invalidReasons.length === 0 && options.workload && uiSameCorpusWorkloadMutatesWorkbook(options.workload)) {
    invalidReasons.push(...sameCorpusMutationTargetProofInvalidReasons(proof, options.workload, options.sampleCount ?? 0))
  }
  return [...new Set(invalidReasons)]
}

function sameCorpusMutationTargetProofInvalidReasons(
  proof: SameCorpusProductSemanticUiProof,
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
): string[] {
  const invalidReasons: string[] = []
  const mutationProofs = proof.mutationTargetProofs
  if (sampleCount <= 0) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing sample count`)
  }
  if (mutationProofs.length < sampleCount) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} covers ${String(mutationProofs.length)}/${String(sampleCount)} samples`,
    )
  }
  const sampleIndexCounts = sameCorpusMutationSampleIndexCounts(mutationProofs)
  const screenshotPathCounts = sameCorpusMutationScreenshotPathCounts(mutationProofs)
  for (const sample of mutationProofs) {
    const screenshotPath = sample.screenshotPath ? normalizeSameCorpusMutationTargetScreenshotPath(sample.screenshotPath) : null
    invalidReasons.push(
      ...sameCorpusMutationTargetProofSampleInvalidReasons(proof, workload, sample, {
        duplicateSampleIndex: (sampleIndexCounts.get(sample.sampleIndex) ?? 0) > 1,
        duplicateScreenshotPath: screenshotPath !== null && (screenshotPathCounts.get(screenshotPath) ?? 0) > 1,
        sampleCount,
      }),
    )
  }
  return invalidReasons
}

function sameCorpusMutationTargetProofSampleInvalidReasons(
  proof: SameCorpusProductSemanticUiProof,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  options: {
    readonly duplicateSampleIndex?: boolean
    readonly duplicateScreenshotPath?: boolean
    readonly sampleCount: number
  },
): string[] {
  const invalidReasons: string[] = []
  if (!Number.isInteger(sample.sampleIndex) || sample.sampleIndex < 0 || sample.sampleIndex >= options.sampleCount) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} has invalid sample index`)
  } else if (options.duplicateSampleIndex === true) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} duplicates sample ${String(sample.sampleIndex + 1)}`)
  }
  if (sample.workload !== workload || sample.intendedOperation !== workload) {
    invalidReasons.push(`semantic UI mutation target proof operation does not match ${workload}`)
  }
  invalidReasons.push(...sameCorpusMutationTargetPayloadInvalidReasons(workload, sample))
  if (sample.sheetName.trim().length === 0 || sample.sheetName !== proof.sheetName) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing the target sheet`)
  }
  if (sample.targetRange.trim().length === 0) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing the target range`)
  }
  if (sample.targetRange.trim().length > 0 && !sameCorpusSelectedRangeMatchesTarget(proof.selectedRange, sample.targetRange)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} target range does not match the rendered selection`)
  }
  if (!sameCorpusReadbacksDiffer(sample.before, sample.after)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} did not prove a before/after target change`)
  }
  if (!sameCorpusReadbacksEqual(sample.before, sample.restored)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} did not prove undo restore`)
  }
  if (sample.undoRestoreStatus !== 'verified') {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} has unverified undo restore status`)
  }
  if (sample.authoritativeReadbackRevision === null || sample.authoritativeReadbackRevision.trim().length === 0) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing authoritative readback revision`)
  }
  if (sample.visibleRenderRevision === null || sample.visibleRenderRevision.trim().length === 0) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing visible render revision`)
  }
  if (sample.screenshotSha256 === null || !/^[a-f0-9]{64}$/u.test(sample.screenshotSha256)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing screenshot SHA256`)
  }
  if (sample.screenshotPath === null || sample.screenshotPath === undefined || sample.screenshotPath.trim().length === 0) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing screenshot artifact path`)
  } else {
    const screenshotPath = normalizeSameCorpusMutationTargetScreenshotPath(sample.screenshotPath)
    if (options.duplicateScreenshotPath === true) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} duplicates screenshot artifact path`)
    }
    invalidReasons.push(...sameCorpusMutationTargetScreenshotPathInvalidReasons(proof.product, workload, sample, screenshotPath))
  }
  invalidReasons.push(...sameCorpusMutationTargetReadbackSourceInvalidReasons(proof.product, workload, sample))
  invalidReasons.push(...sameCorpusMutationTargetVisibleReadbackInvalidReasons(proof.product, workload, sample))
  invalidReasons.push(...sameCorpusMutationTargetRevisionInvalidReasons(proof.product, workload, sample))
  invalidReasons.push(...sameCorpusMutationTargetExpectedReadbackInvalidReasons(workload, sample))
  if (workload === 'fill-format-change' && sample.before.fillColor === sample.after.fillColor) {
    invalidReasons.push('semantic UI mutation target proof for fill-format-change did not prove a fill color change')
  }
  if (workload === 'fill-format-change' && (sample.after.fillColor === null || sample.visibleAfter.fillColor === null)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing rendered post-mutation fill color`)
  }
  return invalidReasons
}

function sameCorpusMutationSampleIndexCounts(samples: readonly SameCorpusMutationTargetProof[]): Map<number, number> {
  const counts = new Map<number, number>()
  for (const sample of samples) {
    counts.set(sample.sampleIndex, (counts.get(sample.sampleIndex) ?? 0) + 1)
  }
  return counts
}

function sameCorpusMutationScreenshotPathCounts(samples: readonly SameCorpusMutationTargetProof[]): Map<string, number> {
  const counts = new Map<string, number>()
  for (const sample of samples) {
    if (sample.screenshotPath === null || sample.screenshotPath === undefined || sample.screenshotPath.trim().length === 0) {
      continue
    }
    const path = normalizeSameCorpusMutationTargetScreenshotPath(sample.screenshotPath)
    counts.set(path, (counts.get(path) ?? 0) + 1)
  }
  return counts
}

function sameCorpusMutationTargetPayloadInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  const payload = sample.intendedPayload
  if (!payload) {
    return [`semantic UI mutation target proof for ${workload} is missing intended operation payload`]
  }
  if (workload === 'edit-visible-cell') {
    return payload.kind === 'cell-value' && payload.value.trim().length > 0
      ? []
      : ['semantic UI mutation target proof for edit-visible-cell is missing intended value payload']
  }
  if (workload === 'formula-edit') {
    return payload.kind === 'formula' && payload.formula.trim().startsWith('=')
      ? []
      : ['semantic UI mutation target proof for formula-edit is missing intended formula payload']
  }
  if (workload === 'fill-format-change') {
    return payload.kind === 'fill-color' &&
      payload.swatchLabel.trim().length > 0 &&
      sameCorpusFillColorValue(payload.expectedFillColor) !== null
      ? []
      : ['semantic UI mutation target proof for fill-format-change is missing intended fill payload']
  }
  return []
}

function sameCorpusMutationTargetExpectedReadbackInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  const payload = sample.intendedPayload
  if (!payload) {
    return []
  }
  if (workload === 'edit-visible-cell') {
    if (payload.kind !== 'cell-value' || sample.after.value !== payload.value || sample.visibleAfter.value !== payload.value) {
      return ['semantic UI mutation target proof for edit-visible-cell did not prove the intended committed target value']
    }
    return []
  }
  if (workload === 'formula-edit') {
    if (payload.kind !== 'formula' || sample.after.formula !== payload.formula || sample.visibleAfter.formula !== payload.formula) {
      return ['semantic UI mutation target proof for formula-edit did not prove the intended edited formula']
    }
    return []
  }
  if (workload === 'fill-format-change') {
    if (payload.kind !== 'fill-color') {
      return ['semantic UI mutation target proof for fill-format-change did not prove the intended fill color']
    }
    if (
      !sameCorpusColorsMatch(sample.after.fillColor, payload.expectedFillColor) ||
      !sameCorpusColorsMatch(sample.visibleAfter.fillColor, payload.expectedFillColor)
    ) {
      return ['semantic UI mutation target proof for fill-format-change did not prove the intended fill color']
    }
  }
  return []
}

function sameCorpusMutationTargetScreenshotPathInvalidReasons(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  screenshotPath: string,
): string[] {
  const invalidReasons: string[] = []
  const segments = screenshotPath.split('/').filter((segment) => segment.length > 0)
  const mutationTargetIndex = segments.lastIndexOf('mutation-target')
  const caseSegment = mutationTargetIndex > 0 ? segments[mutationTargetIndex - 1] : ''
  const fileName = mutationTargetIndex >= 0 ? (segments[mutationTargetIndex + 1] ?? '') : (segments.at(-1) ?? '')
  const expectedFileName = `${product}-sample-${String(sample.sampleIndex + 1)}-after.png`
  if (mutationTargetIndex < 1 || (caseSegment !== workload && !caseSegment.endsWith(`-${workload}`))) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} screenshot artifact path is not tied to the workload`)
  }
  if (fileName !== expectedFileName || mutationTargetIndex !== segments.length - 2) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} screenshot artifact path is not tied to the mutation target sample`,
    )
  }
  return invalidReasons
}

function normalizeSameCorpusMutationTargetScreenshotPath(path: string): string {
  return path.trim().replaceAll('\\', '/')
}

function sameCorpusMutationTargetReadbackSourceInvalidReasons(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  const readbacks = [sample.before, sample.after, sample.restored]
  if (readbacks.some((readback) => readback.source === 'unknown')) {
    return [`semantic UI mutation target proof for ${workload} is missing readback source`]
  }
  if (product === 'bilig' && readbacks.some((readback) => readback.source !== 'bilig-authoritative-range')) {
    return [`semantic UI mutation target proof for ${workload} used visible editor text instead of Bilig authoritative range readback`]
  }
  return []
}

function sameCorpusMutationTargetVisibleReadbackInvalidReasons(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  const readbacks = [sample.visibleAfter, sample.visibleRestored]
  if (readbacks.some((readback) => readback.source === 'unknown')) {
    return [`semantic UI mutation target proof for ${workload} is missing visible render readback source`]
  }
  const expectedSource = product === 'bilig' && workload === 'fill-format-change' ? 'visible-grid-cell' : 'visible-formula-bar'
  if (readbacks.some((readback) => readback.source !== expectedSource)) {
    if (expectedSource === 'visible-grid-cell') {
      return [`semantic UI mutation target proof for ${workload} visible render readback did not come from rendered grid cell pixels`]
    }
    return [`semantic UI mutation target proof for ${workload} visible render readback did not come from the browser-visible surface`]
  }
  if (product !== 'bilig') {
    return []
  }
  if (!sameCorpusVisibleReadbackMatchesAuthoritative(workload, sample.after, sample.visibleAfter)) {
    return [`semantic UI mutation target proof for ${workload} rendered readback does not match Bilig authoritative range readback`]
  }
  if (!sameCorpusVisibleReadbackMatchesAuthoritative(workload, sample.restored, sample.visibleRestored)) {
    return [`semantic UI mutation target proof for ${workload} rendered restore readback does not match Bilig authoritative range readback`]
  }
  return []
}

function sameCorpusMutationTargetRevisionInvalidReasons(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  if (product !== 'bilig') {
    return []
  }
  const invalidReasons: string[] = []
  const readbacks = [sample.before, sample.after, sample.restored]
  if (readbacks.some((readback) => !hasSameCorpusText(readback.capturedRevision))) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing Bilig authoritative readback revision`)
  }
  if (hasSameCorpusText(sample.after.capturedRevision) && sample.authoritativeReadbackRevision !== sample.after.capturedRevision) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} authoritative revision does not match target readback`)
  }
  if (!hasSameCorpusText(sample.after.visibleSceneProofSha256)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing Bilig visible scene proof`)
  } else if (sample.visibleRenderRevision !== `bilig-visible-scene-sha256:${sample.after.visibleSceneProofSha256}`) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} visible render revision does not match target scene proof`)
  }
  return invalidReasons
}

function sameCorpusVisibleReadbackMatchesAuthoritative(
  workload: UiResponsivenessSameCorpusWorkload,
  authoritative: SameCorpusMutationTargetReadback,
  visible: SameCorpusMutationTargetReadback,
): boolean {
  if (workload === 'formula-edit') {
    return authoritative.formula === visible.formula
  }
  if (workload === 'fill-format-change') {
    if (authoritative.fillColor === null || visible.fillColor === null) {
      return authoritative.fillColor === visible.fillColor
    }
    return sameCorpusColorsMatch(visible.fillColor, authoritative.fillColor)
  }
  return authoritative.value === visible.value || authoritative.visibleText === visible.visibleText
}

function sameCorpusReadbacksEqual(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return (
    left.value === right.value &&
    left.formula === right.formula &&
    left.fillColor === right.fillColor &&
    left.visibleText === right.visibleText &&
    left.source === right.source
  )
}

function sameCorpusReadbacksDiffer(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return !sameCorpusReadbacksEqual(left, right)
}

function sameCorpusSelectedRangeMatchesTarget(selectedRange: string | null, targetRange: string): boolean {
  return normalizeSameCorpusRenderedSelectionRange(selectedRange) === normalizeSameCorpusRenderedSelectionRange(targetRange)
}

function normalizeSameCorpusRenderedSelectionRange(range: string | null): string {
  const rawRange = range?.split('!').at(-1)?.replace(/\$/gu, '').trim().toUpperCase() ?? ''
  const match = rawRange.match(/^[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?$/u)
  return match?.[0] ?? ''
}

function hasSameCorpusText(value: string | null | undefined): value is string {
  return value !== null && value !== undefined && value.trim().length > 0
}

function sameCorpusColorsMatch(actual: string | null, expected: string): boolean {
  const actualColor = sameCorpusFillColorValue(actual)
  const expectedColor = sameCorpusFillColorValue(expected)
  return actualColor !== null && expectedColor !== null && actualColor === expectedColor
}

function sameCorpusFillColorValue(value: string | null | undefined): string | null {
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

export async function readSameCorpusVisibleSelectedRange(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<string | null> {
  if (product === 'bilig') {
    const status = await page
      .locator('[data-testid="status-selection"]')
      .first()
      .textContent({ timeout: 1_000 })
      .catch(() => null)
    if (status && status.trim().length > 0) {
      return status.trim()
    }
    const nameBox = await page
      .locator('[data-testid="name-box"]')
      .first()
      .inputValue({ timeout: 1_000 })
      .catch(() => null)
    return nameBox && nameBox.trim().length > 0 ? nameBox.trim() : null
  }
  if (product === 'google-sheets') {
    const selectedRange = await readGoogleSheetsNameBoxSelection(page)
    if (selectedRange) {
      return selectedRange
    }
  }
  return await page.evaluate(() => {
    const selectors = [
      '#t-name-box input',
      'input[aria-label="Name box"]',
      '[aria-label="Name box"] input',
      '[aria-label="Name box"]',
      '[aria-label*="selected cell" i]',
      '[aria-label*="selection" i]',
    ]
    for (const selector of selectors) {
      const element = document.querySelector<HTMLInputElement | HTMLElement>(selector)
      const value =
        element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement
          ? element.value
          : (element?.getAttribute('aria-label') ?? element?.textContent ?? '')
      const trimmed = value.trim()
      if (trimmed.length > 0 && trimmed.length <= 64) {
        return trimmed
      }
    }
    const activeElement = document.activeElement
    const activeLabel = activeElement?.getAttribute('aria-label') ?? ''
    const activeLabelMatch = activeLabel.match(/(?:cell|range)\s+([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)/iu)
    return activeLabelMatch?.[1] ?? null
  })
}

export interface SameCorpusNameBoxReaderPage {
  readonly keyboard: {
    press(key: string): Promise<void>
  }
  locator(selector: string): SameCorpusReadableLocator
}

interface SameCorpusReadableLocator {
  first(): SameCorpusReadableLocator
  inputValue(options?: { readonly timeout?: number }): Promise<string>
}

export async function readGoogleSheetsNameBoxSelection(page: SameCorpusNameBoxReaderPage): Promise<string | null> {
  await page.keyboard.press(primaryShortcut('J')).catch(() => null)
  const value = await page
    .locator('#t-name-box')
    .first()
    .inputValue({ timeout: 1_500 })
    .catch(() => null)
  await page.keyboard.press('Escape').catch(() => null)
  return value && value.trim().length > 0 ? value.trim() : null
}

function primaryShortcut(key: string, platform: NodeJS.Platform = process.platform): string {
  return platform === 'darwin' ? `Meta+${key}` : `Control+${key}`
}

function semanticUiProofMethod(product: UiResponsivenessSameCorpusProduct): SameCorpusProductSemanticUiProof['method'] {
  if (product === 'bilig') {
    return 'bilig-visible-semantic-readback'
  }
  if (product === 'google-sheets') {
    return 'google-sheets-visible-semantic-readback'
  }
  return 'excel-web-visible-semantic-readback'
}

function sameCorpusSemanticUiEvidence(
  product: UiResponsivenessSameCorpusProduct,
  pixelGridProof: SameCorpusProductPixelGridProof,
  corpusVerification: SameCorpusCaptureCorpusVerification,
  selectedRange: string | null,
): string[] {
  const evidence = sameCorpusEvidenceMap(pixelGridProof.evidence)
  return [
    'semanticUiProofVersion=semantic-ui-v1',
    `product=${product}`,
    `sheetName=${corpusVerification.sheetName}`,
    `selectedRange=${selectedRange ?? ''}`,
    `checkedCellCount=${String(corpusVerification.checkedCells.length)}`,
    `corpusVerified=${String(corpusVerification.verified)}`,
    `gridAuthoritativeRevision=${evidence.get('gridAuthoritativeRevision') ?? ''}`,
    `visibleRenderRevision=${evidence.get('visibleRenderRevision') ?? ''}`,
  ]
}

function screenshotBufferSha256(screenshotBuffer: SameCorpusScreenshotBuffer): string {
  return createHash('sha256')
    .update(Buffer.from(screenshotBuffer.toString('base64'), 'base64'))
    .digest('hex')
}

function sameCorpusEvidenceMap(evidence: readonly string[]): ReadonlyMap<string, string> {
  return new Map(
    evidence.map((entry) => {
      const [key, ...valueParts] = entry.split('=')
      return [key ?? '', valueParts.join('=')] as const
    }),
  )
}
