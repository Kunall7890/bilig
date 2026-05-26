import { createHash } from 'node:crypto'

import type { Page } from '@playwright/test'

import type {
  SameCorpusCaptureCorpusVerification,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { sameCorpusPreflightReadbackProvesValue, type PreflightEditableMutationProof } from './ui-responsiveness-same-corpus-preflight.ts'
import {
  readSameCorpusDeclaredMutationTargetSelection,
  readSameCorpusMutationTargetReadback,
  selectSameCorpusMutationTargetRange,
  type SameCorpusMutationTargetSelection,
} from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import { settleFrames } from './ui-responsiveness-same-corpus-page-utils.ts'
import type { SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import { restoreProductWorkbookMutation } from './ui-responsiveness-same-corpus-workload-runner.ts'

export async function captureSameCorpusPreflightEditableMutationProof(args: {
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly page: Page
  readonly product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>
}): Promise<PreflightEditableMutationProof> {
  const target = await readSameCorpusDeclaredMutationTargetSelection({
    page: args.page,
    product: args.product,
    sampleIndex: 0,
    sheetName: args.corpusVerification.sheetName,
    workload: 'edit-visible-cell',
  })
  await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target })
  const before = await readSameCorpusMutationTargetReadback({ page: args.page, product: args.product, target })
  const intendedValue = preflightSentinelValue(args.product, target)
  let restoredAfterMutation = false
  try {
    await args.page.keyboard.type(intendedValue)
    await args.page.keyboard.press('Enter')
    await settleFrames(args.page, 12)
    await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target })
    const after = await readSameCorpusMutationTargetReadback({ page: args.page, product: args.product, target })
    await restoreProductWorkbookMutation(args.page, 'edit-visible-cell')
    restoredAfterMutation = true
    await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target })
    const restored = await readSameCorpusMutationTargetReadback({ page: args.page, product: args.product, target })
    return preflightEditableMutationProof({
      after,
      before,
      intendedValue,
      product: args.product,
      restored,
      target,
    })
  } finally {
    if (!restoredAfterMutation) {
      await restoreProductWorkbookMutation(args.page, 'edit-visible-cell').catch(() => undefined)
      await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target }).catch(() => undefined)
    }
  }
}

function preflightEditableMutationProof(args: {
  readonly after: SameCorpusMutationTargetReadback
  readonly before: SameCorpusMutationTargetReadback
  readonly intendedValue: string
  readonly product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>
  readonly restored: SameCorpusMutationTargetReadback
  readonly target: SameCorpusMutationTargetSelection
}): PreflightEditableMutationProof {
  const undoRestoreStatus = sameCorpusPreflightReadbacksEqual(args.before, args.restored) ? 'verified' : 'failed'
  const captured = sameCorpusPreflightReadbackProvesValue(args.after, args.intendedValue) && undoRestoreStatus === 'verified'
  return {
    after: args.after,
    authoritativeReadbackRevision: preflightReadbackRevision('authoritative', args.product, args.target, args.after),
    before: args.before,
    captured,
    evidence: [
      'method=sentinel-write-readback-restore',
      `targetRange=${args.target.targetRange}`,
      `writtenValueVerified=${String(sameCorpusPreflightReadbackProvesValue(args.after, args.intendedValue))}`,
      `undoRestoreStatus=${undoRestoreStatus}`,
    ],
    intendedOperation: 'edit-visible-cell',
    intendedValue: args.intendedValue,
    method: 'sentinel-write-readback-restore',
    product: args.product,
    restored: args.restored,
    sampleIndex: 0,
    sheetId: args.target.sheetId,
    sheetName: args.target.sheetName,
    targetRange: args.target.targetRange,
    undoRestoreStatus,
    visibleReadbackRevision: preflightReadbackRevision('visible', args.product, args.target, args.after),
  }
}

function preflightSentinelValue(
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  target: SameCorpusMutationTargetSelection,
): string {
  return `same-corpus-preflight-${product}-${target.targetRange.toLowerCase()}-${Date.now().toString(36)}`
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

function preflightReadbackRevision(
  phase: 'authoritative' | 'visible',
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  target: SameCorpusMutationTargetSelection,
  readback: SameCorpusMutationTargetReadback,
): string {
  return `${product}-preflight-${phase}-readback-sha256:${sha256Hex(
    stableJsonBytes({ product, readback, sheetId: target.sheetId, sheetName: target.sheetName, targetRange: target.targetRange }),
  )}`
}

function stableJsonBytes(value: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(stableJsonValue(value)))
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entryValue]) => entryValue !== undefined)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entryValue]) => [key, stableJsonValue(entryValue)]),
    )
  }
  return value
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}
