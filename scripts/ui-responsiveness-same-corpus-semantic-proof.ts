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
  readonly sheetName: string
  readonly targetRange: string
  readonly before: SameCorpusMutationTargetReadback
  readonly after: SameCorpusMutationTargetReadback
  readonly restored: SameCorpusMutationTargetReadback
  readonly authoritativeReadbackRevision: string | null
  readonly visibleRenderRevision: string | null
  readonly screenshotSha256: string | null
  readonly undoRestoreStatus: 'verified' | 'missing' | 'failed'
}

export interface SameCorpusMutationTargetReadback {
  readonly value: string | null
  readonly formula: string | null
  readonly fillColor: string | null
  readonly visibleText: string | null
}

export type SameCorpusProductSemanticUiProofEvidenceStatus = 'current-contract' | 'missing' | 'invalid'

export interface SameCorpusProductSemanticUiProofVerdict {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly evidenceStatus: SameCorpusProductSemanticUiProofEvidenceStatus
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
      })) ?? [],
    evidence,
  }
  const verdict = validateSameCorpusProductSemanticUiProof(proof, { workload: args.workload, sampleCount: args.sampleCount })
  return verdict.acceptedForCurrentScorecard ? proof : { ...proof, captured: false }
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
  const seenSamples = new Set<number>()
  for (const sample of mutationProofs) {
    if (!Number.isInteger(sample.sampleIndex) || sample.sampleIndex < 0 || sample.sampleIndex >= sampleCount) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} has invalid sample index`)
    } else if (seenSamples.has(sample.sampleIndex)) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} duplicates sample ${String(sample.sampleIndex + 1)}`)
    } else {
      seenSamples.add(sample.sampleIndex)
    }
    if (sample.workload !== workload || sample.intendedOperation !== workload) {
      invalidReasons.push(`semantic UI mutation target proof operation does not match ${workload}`)
    }
    if (sample.sheetName.trim().length === 0 || sample.sheetName !== proof.sheetName) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing the target sheet`)
    }
    if (sample.targetRange.trim().length === 0) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} is missing the target range`)
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
    if (workload === 'formula-edit' && (sample.after.formula === null || !sample.after.formula.trim().startsWith('='))) {
      invalidReasons.push('semantic UI mutation target proof for formula-edit is missing the edited formula')
    }
    if (workload === 'edit-visible-cell' && (sample.after.value === null || sample.after.value === sample.before.value)) {
      invalidReasons.push('semantic UI mutation target proof for edit-visible-cell did not prove a committed target value change')
    }
    if (workload === 'fill-format-change' && sample.before.fillColor === sample.after.fillColor) {
      invalidReasons.push('semantic UI mutation target proof for fill-format-change did not prove a fill color change')
    }
  }
  return invalidReasons
}

function sameCorpusReadbacksEqual(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return (
    left.value === right.value &&
    left.formula === right.formula &&
    left.fillColor === right.fillColor &&
    left.visibleText === right.visibleText
  )
}

function sameCorpusReadbacksDiffer(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return !sameCorpusReadbacksEqual(left, right)
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
