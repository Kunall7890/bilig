import type { Page } from '@playwright/test'

import {
  buildSameCorpusCommittedStateProof,
  captureSameCorpusCommittedStatePhaseProof,
} from './ui-responsiveness-same-corpus-committed-state-proof.ts'
import type { SameCorpusCaptureCorpusVerification } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import {
  readSameCorpusDeclaredMutationTargetSelection,
  readSameCorpusVisibleMutationTargetReadback,
  selectSameCorpusMutationTargetRange,
} from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { PreflightFillFormatMutationProof } from './ui-responsiveness-same-corpus-preflight.ts'
import { settleFrames } from './ui-responsiveness-same-corpus-page-utils.ts'
import {
  performSameCorpusFillColorOperation,
  restoreProductWorkbookMutation,
  sameCorpusFillColorExpectedColor,
  sameCorpusFillColorSwatchLabel,
} from './ui-responsiveness-same-corpus-workload-runner.ts'

export async function captureSameCorpusPreflightFillFormatMutationProof(args: {
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly page: Page
}): Promise<PreflightFillFormatMutationProof> {
  const sampleIndex = 0
  const target = await readSameCorpusDeclaredMutationTargetSelection({
    page: args.page,
    product: 'google-sheets',
    sampleIndex,
    sheetName: args.corpusVerification.sheetName,
    workload: 'fill-format-change',
  })
  await selectSameCorpusMutationTargetRange({ page: args.page, product: 'google-sheets', target })
  const before = await readSameCorpusVisibleMutationTargetReadback({
    page: args.page,
    product: 'google-sheets',
    target,
    workload: 'fill-format-change',
  })
  const beforeCommittedStateProof = await captureSameCorpusCommittedStatePhaseProof({
    expectedReadback: before,
    page: args.page,
    phase: 'before',
    product: 'google-sheets',
    sampleIndex,
    target,
    workload: 'fill-format-change',
  })
  const intendedFillColor = sameCorpusFillColorExpectedColor(sampleIndex)
  let restoredAfterMutation = false
  try {
    await performSameCorpusFillColorOperation(args.page, 'google-sheets', sampleIndex, { exactSwatchOnly: true })
    await settleFrames(args.page, 12)
    await selectSameCorpusMutationTargetRange({ page: args.page, product: 'google-sheets', target })
    const after = await readSameCorpusVisibleMutationTargetReadback({
      page: args.page,
      product: 'google-sheets',
      target,
      workload: 'fill-format-change',
    })
    const afterCommittedStateProof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: { ...after, fillColor: intendedFillColor },
      page: args.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex,
      target,
      workload: 'fill-format-change',
    })
    await restoreProductWorkbookMutation(args.page, 'fill-format-change')
    restoredAfterMutation = true
    await selectSameCorpusMutationTargetRange({ page: args.page, product: 'google-sheets', target })
    const restored = await readSameCorpusVisibleMutationTargetReadback({
      page: args.page,
      product: 'google-sheets',
      target,
      workload: 'fill-format-change',
    })
    const restoredCommittedStateProof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: before,
      page: args.page,
      phase: 'restored',
      product: 'google-sheets',
      sampleIndex,
      target,
      workload: 'fill-format-change',
    })
    const committedStateProof = buildSameCorpusCommittedStateProof({
      after: afterCommittedStateProof,
      before: beforeCommittedStateProof,
      product: 'google-sheets',
      restored: restoredCommittedStateProof,
      sampleIndex,
      target,
      workload: 'fill-format-change',
    })
    const undoRestoreStatus = normalizeColor(before.fillColor) === normalizeColor(restored.fillColor) ? 'verified' : 'failed'
    const captured =
      normalizeColor(after.fillColor) === normalizeColor(intendedFillColor) &&
      normalizeColor(committedStateProof?.after.readback.fillColor ?? null) === normalizeColor(intendedFillColor) &&
      undoRestoreStatus === 'verified'
    return {
      after,
      before,
      captured,
      committedStateProof,
      evidence: [
        'method=fill-color-commit-readback-restore',
        `targetRange=${target.targetRange}`,
        `intendedFillColor=${intendedFillColor}`,
        `renderedFillVerified=${String(normalizeColor(after.fillColor) === normalizeColor(intendedFillColor))}`,
        `committedFillVerified=${String(
          normalizeColor(committedStateProof?.after.readback.fillColor ?? null) === normalizeColor(intendedFillColor),
        )}`,
        `undoRestoreStatus=${undoRestoreStatus}`,
      ],
      intendedFillColor,
      intendedOperation: 'fill-format-change',
      method: 'fill-color-commit-readback-restore',
      product: 'google-sheets',
      restored,
      sampleIndex,
      sheetId: target.sheetId,
      sheetName: target.sheetName,
      swatchLabel: sameCorpusFillColorSwatchLabel(sampleIndex),
      targetRange: target.targetRange,
      undoRestoreStatus,
    }
  } finally {
    if (!restoredAfterMutation) {
      await restoreProductWorkbookMutation(args.page, 'fill-format-change').catch(() => undefined)
      await selectSameCorpusMutationTargetRange({ page: args.page, product: 'google-sheets', target }).catch(() => undefined)
    }
  }
}

function normalizeColor(value: string | null): string | null {
  return value ? value.trim().toLowerCase() : null
}
